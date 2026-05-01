# What Python's xlsx libraries get right

*Architectural takeaways from openpyxl and xlsxwriter, written for future engine-related decisions in xlsx-for-ai. Reading time: ~10 minutes.*

The Python xlsx ecosystem (mostly **openpyxl** for read+write, **xlsxwriter** for write-only) is more architecturally mature than the Node ecosystem. They've been actively maintained for over a decade and have absorbed lessons from the entire long tail of real-world workbooks. This document captures what they do well so the patterns can inform JS-side decisions — whether we eventually fork ExcelJS, replace it, or build a thin engine abstraction in xlsx-for-ai itself.

---

## 1. The fundamental insight: xlsx is XML in a zip

Both libraries treat an `.xlsx` file the way it actually is: a ZIP archive (per the Open Packaging Conventions / ECMA-376) containing a directory tree of XML files. The library's job decomposes neatly into three layers:

1. **ZIP container** — read/write the archive
2. **XML parse/emit** — turn `<c r="A1" t="n"><v>42</v></c>` into objects (or vice versa)
3. **OpenXML schema knowledge** — knowing that `<c>` is "cell", `r` is the address, `t="n"` is the value type, etc.

Layers 1 and 2 are commodity — every modern language has good zip and XML libraries. **Layer 3 is where every xlsx library spends 95% of its code, and where every library's bugs and limitations live.** Both Python libraries are organized to make that schema-knowledge layer maintainable.

ExcelJS organizes the same problem less cleanly. The schema knowledge is spread across dozens of files with overlapping responsibilities, which is part of why the library has accumulated 786 open issues without addressing them.

---

## 2. xlsxwriter: streaming writer pattern

[`xlsxwriter`](https://github.com/jmcnamara/XlsxWriter) is **write-only by design** — it deliberately doesn't try to read existing files. That single constraint produces a much cleaner architecture.

### File-per-OpenXML-part

Looking at `xlsxwriter/`'s package layout:

```
chart.py              chart_area.py    chart_bar.py     chart_line.py
chart_pie.py          chartsheet.py    comments.py      contenttypes.py
core.py               custom.py        drawing.py       format.py
image.py              metadata.py      packager.py      relationships.py
sharedstrings.py      shape.py         styles.py        table.py
theme.py              url.py           vml.py           workbook.py
worksheet.py          xmlwriter.py     ...
```

**Every OpenXML part has its own file.** `styles.py` knows the styles XML schema. `chart_pie.py` knows the pie-chart schema. `sharedstrings.py` knows the shared-strings table. `comments.py` knows comments. Etc.

Each file is bounded — you can read `comments.py` end-to-end and understand exactly how comments work. Schema bugs are localized: a chart bug doesn't accidentally break sharedstrings.

### `xmlwriter.py` as a tiny shared base class

Every part-emitter inherits from a common base (`XMLwriter`) that provides the XML primitives — `_xml_start_tag()`, `_xml_data_element()`, `_xml_end_tag()`. The base class is intentionally simple: just enough to emit well-formed XML with proper escaping. It doesn't know anything about xlsx.

This separation matters: schema knowledge lives in the part files; XML mechanics lives in the base; neither concerns the other.

### `packager.py` orchestrates the assembly

When you call `workbook.close()`, `packager.py` is what runs. It walks every part, asks each one to write its XML to a buffer, then stuffs the buffers into a zip with the correct filenames and `[Content_Types].xml` manifest. **One file owns "how an xlsx is assembled."** No part file has to know it lives inside a zip.

### Streaming mode for memory bounds

xlsxwriter has a "constant memory" mode (`Workbook(filename, {'constant_memory': True})`) that flushes each row to disk immediately instead of holding the workbook in memory. Lets you write multi-gigabyte xlsx files in a few hundred MB of RAM. The architecture supports this because each part can serialize independently — you don't need the whole workbook resolved before you start writing.

### What this implies for JS

- **One file per OpenXML part** is a cleaner organizing principle than ExcelJS's mix of monolithic-and-fragmented files.
- A **shared XML-primitives base class** is worth ~50 lines and pays dividends.
- A **single packager/orchestrator** keeps the zip-assembly logic from leaking into part files.
- **Streaming write** is achievable if you commit to the "each part serializes independently" rule from day one. ExcelJS has streaming write support but it's grafted on, with sharp edges.

---

## 3. openpyxl: object model + typed descriptors

[`openpyxl`](https://foss.heptapod.net/openpyxl/openpyxl) does both read and write. That's harder, and its architecture reflects the additional complexity, but it gets several things right.

### Schema-as-code via descriptors

openpyxl models every OpenXML element as a Python class with typed attributes declared as **descriptors**:

```python
class Cell(Serialisable):
    tagname = "c"
    r = String()                # cell reference, e.g. "A1"
    s = Integer()               # style index
    t = NoneSet(values=("b","n","e","s","str","inlineStr"))  # type
    v = Typed(expected_type=ValueDescriptor, allow_none=True)  # value
```

Each descriptor (`String`, `Integer`, `NoneSet`, `Typed`) carries:
- The XML attribute name
- The expected Python type
- Validation rules
- Serialization/deserialization logic

This means **reading and writing share the same source of truth**. The descriptors describe the schema once; the framework derives both the parser and the serializer from them. You can't accidentally have a writer that emits something the reader can't understand.

This is the architectural pattern openpyxl pioneered for xlsx, borrowed loosely from XML-Schema bindings. ExcelJS doesn't have it; its read and write paths are independent code, which is why round-trip fidelity bugs creep in (the read code captures something the write code doesn't recognize, or vice versa).

### Read-only and write-only modes for memory bounds

openpyxl has two streaming modes:

- **`load_workbook(path, read_only=True)`** — uses lxml's iterparse to stream sheet rows. The whole workbook isn't in memory; you iterate `sheet.rows` and each row is yielded once.
- **`Workbook(write_only=True)`** — like xlsxwriter's constant-memory mode. You append rows; openpyxl flushes them to disk.

Both modes preserve **most** functionality but drop a few features that need full-workbook context (formulas referencing cells you haven't read yet, sheet-cross-references in defined names, etc.). The library is honest about what each mode preserves and what it drops.

ExcelJS has streaming counterparts but with worse documentation and more inconsistencies between streaming and non-streaming output.

### Strict typing at attribute-set time

Because descriptors validate on assignment, you get errors early:

```python
cell.t = "invalid"  # raises immediately: NoneSet validation failed
```

vs ExcelJS, which mostly accepts whatever you give it and produces silently-broken output later. The typed-descriptor approach is more verbose but produces dramatically more reliable round-trip behavior.

### Extension-list preservation

OpenXML has an "extension list" mechanism — vendor-specific XML elements that the spec allows but doesn't define. openpyxl preserves these blindly: it parses them as opaque XML, holds them on the in-memory object, and writes them back unchanged. ExcelJS strips them (anything it doesn't recognize). That's a real fidelity difference for round-trip workflows on workbooks produced by Excel itself, which uses the extension mechanism for some newer features.

---

## 4. Architectural patterns worth borrowing in JS

Concrete patterns that would improve any JS-side engine work, whether forking ExcelJS, replacing it, or building a wrapper:

### A. One file per OpenXML part

Borrow xlsxwriter's organizing principle. `lib/parts/styles.js`, `lib/parts/sharedstrings.js`, `lib/parts/charts/pie.js`, etc. Bounded. Inspectable. Bug-localized. ExcelJS today has files like `xlsx/xlsx.js` (4000+ lines mixing zip handling, xml parsing, and several part schemas) — exactly the wrong organization.

### B. Schema descriptors as the single source of truth

A descriptor system in JS is straightforward — a small library of `string()`, `int()`, `enum(...)`, `typed(...)` factories that produce parser+serializer pairs. ~200 lines. Each part class declares its attributes once; reading and writing both derive from that declaration. Eliminates the entire class of "round-trip drift because read sees X and write doesn't" bugs we've been hitting.

### C. Tiny XML-primitives base, owned by the engine

A 30-line base that handles tag start/end/empty, attribute escaping, text-content escaping. Every part-emitter uses it. Don't reinvent escaping in each file (that's how you ship the xlsx-for-ai equivalent of `\r\n`-vs-`\n` bugs).

### D. One packager/orchestrator

A single file knows: "to produce a complete xlsx, walk these parts in this order, emit `[Content_Types].xml` + `_rels/.rels` + the rest, zip them with these compression rules." No part file has to know about the zip layer.

### E. Streaming as a first-class mode, not an afterthought

If we ever build our own engine, design streaming in from day one. Each part serializes independently → packager assembles → zip flushes incrementally. xlsxwriter does this; ExcelJS retrofits it.

### F. Honest preservation of unknown XML

Capture and re-emit anything we don't understand (the "extension list" pattern). Means our round-trip works on Excel-produced files using newer features without us having to chase the spec. Tradeoff: slightly bigger output buffer per cell, but worth it.

### G. Validation at object-construction time, not at write time

If a spec passes validation, writing it should never fail. Discover bad input as early as possible. JS's lack of types makes this harder than Python's descriptor system, but TypeScript or Zod-style validators get most of the way there.

---

## 5. What we should *not* borrow

- **Python-specific machinery.** Descriptors, metaclasses, ABCs — these are Python idioms; the JS equivalent is small validator factories or TypeScript types.
- **The full schema surface.** openpyxl tries to model nearly everything in OpenXML. xlsx-for-ai's wedge is AI-native I/O, not Excel-feature-completeness. We can support a tight subset — values, formulas, formatting, named ranges, merges, frozen panes, hidden rows/cols — and explicitly punt on charts, pivots, conditional formatting, etc., relying on Python or SheetJS Pro server-side for those when needed.
- **Bidirectional everything.** It's tempting to make every feature work both ways. Sometimes it's fine to read a feature you can't write (we already do this for charts and pivots — flag their existence on read, don't try to round-trip). The honesty in the report tab is more valuable than the asymmetric-but-fully-functional alternative.

---

## 6. Practical implications for xlsx-for-ai

Today, xlsx-for-ai is a thin layer on ExcelJS. ExcelJS's architecture is the architectural ceiling we're inheriting, and its slow maintenance means that ceiling won't rise.

When the time comes to either fork or replace:

**Cheap path — patches via patch-package:** apply targeted patches to ExcelJS source addressing specific bugs we hit. Doesn't fix the architectural mess but solves concrete user issues. Already wired in xlsx-for-ai (devDep) for when we need it.

**Medium path — engine wrapper inside xlsx-for-ai:** define an interface (`engine.readWorkbook(path)`, `engine.writeWorkbook(spec, path)`) that abstracts over whichever engine implementation is current. Lets us swap engines per-installation or per-feature without rewriting xlsx-for-ai's logic. The wrapper is the seam; what's behind it can be ExcelJS, a fork, or a from-scratch implementation. **This is the highest-value pre-engine move** because it makes any future engine choice cheap.

**Expensive path — minimal from-scratch JS engine:** apply patterns A through G above. Build only what xlsx-for-ai needs (the AI-native subset, not Excel-feature-complete). Realistically 4-8 weeks of focused work for a working v1, plus ongoing maintenance of the schema surface. Worth it only if (a) ExcelJS becomes unmaintainable, (b) the supervisor product has revenue funding the work, and (c) the existing test corpus is robust enough to catch regressions.

**Don't fork ExcelJS for general modernization.** The cost-benefit is bad: you inherit the architectural mess plus all the open issues plus ongoing maintenance, with no community uplift to share the load. Patches via patch-package give you 80% of the value at 5% of the cost.

---

## Sources

- [xlsxwriter on PyPI](https://pypi.org/project/XlsxWriter/) — actively maintained, last release Sept 2025
- [xlsxwriter source](https://github.com/jmcnamara/XlsxWriter) — file layout cited above
- [openpyxl on PyPI](https://pypi.org/project/openpyxl/) — actively maintained, last release June 2024
- [openpyxl docs](https://openpyxl.readthedocs.io/) — when the docs site is up
- ECMA-376 OpenXML specification — the underlying schema both libraries implement
