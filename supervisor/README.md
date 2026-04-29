# xlsx-supervisor

> **Status:** WIP MVP. Currently lives on `feat/supervisor-mvp` branch in the
> `xlsx-for-ai` repo for convenience. Will move to its own repo
> (`xlsx-supervisor`) once the structure stabilizes.

Server-side spreadsheet review. Upload an `.xlsx`, get back a copy with a
`_xlsx-for-ai` review tab explaining what's wrong and how to fix it.

Same product DNA as the OSS [`xlsx-for-ai`](../) CLI's write mode + review
tab — repackaged as a web app for non-programmers who don't want to touch
a terminal.

## What's written (not yet smoke-tested in this environment)

- FastAPI backend (`/upload`, `/download/{id}`, `/health`)
- HTMX-driven web UI for drag-drop upload → download
- Deterministic structural review (no LLM yet — see "stubbed" below):
  - Formula errors (`#REF!`, `#NAME?`, `#DIV/0!`, etc.)
  - Broken cross-sheet formula references
  - Hidden rows/columns containing data
  - External workbook links
- `_xlsx-for-ai` review tab embedded in output (matches the CLI's format
  — What happened / What we did / Risk / Tradeoff / Alternative per issue)
- In-memory file storage with auto-purge after 1 hour
- ~25 tests across reviewer, processor, and API layers

> **Smoke-test caveat:** `pip install` was stalling in the build session
> when these files were committed (network was fine, multiple backgrounded
> install processes appeared to deadlock). Code is written using standard
> FastAPI / openpyxl / xlsxwriter patterns; expect it to run cleanly once
> the venv is built fresh. First thing to do on pickup is verify the venv
> install + run the tests.

## What's stubbed

These are deliberately deferred until a strategic decision is made:

- **LLM integration**: the "review" step uses deterministic checks only.
  No Claude/OpenAI API calls. The skeleton is real; the AI step is a
  placeholder. Decision needed: **bring-your-own-AI** (user provides an
  API key, we just do the wiring) vs **we-use-our-key** (lower friction,
  we pay per call). Affects pricing and architecture.
- **Auth, accounts, billing**: not implemented. Anyone can upload.
- **Persistent storage**: files held in process memory only. A restart
  loses everything. Fine for local dev; not for production.
- **Production deployment**: no Dockerfile, no infrastructure config.
- **Pretty UI**: functional ugly. HTMX + minimal CSS. No design pass.
- **Concurrency / scale**: single-process, in-memory. Won't horizontally
  scale as written.

## Run locally

```bash
cd supervisor
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

Then open <http://localhost:8000>.

## Run tests

```bash
cd supervisor
source venv/bin/activate
pip install pytest httpx
PYTHONPATH=. pytest tests/ -v
```

## Architecture

```
supervisor/
├── app/
│   ├── main.py        - FastAPI routes + HTMX integration
│   ├── reviewer.py    - structural checks (deterministic, no LLM)
│   ├── processor.py   - read xlsx (openpyxl) → review → write (xlsxwriter)
│   └── storage.py     - in-memory file store with TTL
├── templates/         - Jinja2 / HTMX templates
├── static/            - CSS only
└── tests/             - pytest unit + integration + E2E
```

**Engine choice:** openpyxl for reading (best-in-class Python xlsx reader),
xlsxwriter for writing (cleanest output + chart/pivot/conditional-formatting
support when we eventually need it). Both are BSD/MIT licensed — zero
license fees, no commercial gotchas.

**Why not preserve source byte-for-byte:** We rebuild the output rather
than modifying the input in-place. Loses some advanced formatting we don't
round-trip in v1; gains clean output and reliable review-tab insertion. If
a real user need surfaces for full round-trip, swap the write path to
openpyxl-only (load → modify in-place → save) — same pipeline shape,
different tool.

## Pickup queue (for when this becomes its own repo)

1. Decide LLM integration strategy and wire it up
2. Add auth (probably oauth2 + users; or magic-link email)
3. Add billing (probably Stripe; $10/month flat per the supervisor strategy)
4. Production deployment (Hetzner / Railway / Fly)
5. Real frontend pass (or stay HTMX-minimal — depends on positioning)
6. Move to its own repo (`xlsx-supervisor`)
