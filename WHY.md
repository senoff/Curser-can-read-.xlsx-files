# Why xlsx-for-ai exists

*A plain-English version. For the technical reference, see [README.md](README.md).*

## The problem you've probably hit

You have a spreadsheet — a budget, a financial model, a tax estimate, a list of customers. You ask Claude (or ChatGPT, or Cursor) for help with it.

So you copy and paste a section into the chat. The AI gives you advice that sounds reasonable but feels generic. It misses the broken formula in row 47. It doesn't notice that one tab's totals don't match another tab's source. It can't tell you why the gross margin number changes when you add a new column. It treats your spreadsheet as a blob of numbers — because that's all it can see.

You're not going crazy. The AI literally cannot read the file. It can read text, code, even images of your spreadsheet — but the actual `.xlsx` binary is invisible to it. Formulas, formatting, named ranges, links between sheets — all of that disappears the moment you hit copy-paste.

## What changes when you install this

Once `xlsx-for-ai` is on your machine, your AI tools (Claude, Cursor, Copilot, ChatGPT desktop apps with code execution) can finally **read your spreadsheet the way they read everything else** — every formula, every colored cell, every hidden row, every formula reference between sheets.

Now when you ask for help, you get a real review:

- *"Cell B47 has `#REF!` — it's pointing at a sheet you renamed last week."*
- *"Your gross margin formula in row 12 references the wrong column on the COGS tab — it's pulling Q3 numbers into the Q4 totals."*
- *"This 'Total' cell on the Summary tab shows $312k, but if I add up the source rows on the Detail tab I get $327k. Something's off."*

That's the difference between a friend skimming the printed numbers and an analyst who actually opens the file.

## Things that become possible

A few examples people find useful:

- **Have your AI find errors in a financial model** before you send it to your accountant or your board.
- **Have your AI hand you back a corrected version** — not just *say* what should change, but actually produce the fixed `.xlsx` with the changes applied. The corrected file even includes a built-in review note explaining what the AI changed, why, and how to override anything you don't agree with. Same shape as having a careful editor mark up your draft.
- **Compare two versions of the same spreadsheet** ("what changed between V11 and V14?") and get a list of every cell that moved.
- **Turn a CSV export from QuickBooks into a clean SQL database table** in one command, with the column types figured out automatically.
- **Walk through a 50-tab model someone else built** and have the AI explain how the sheets reference each other.
- **Process a folder of legacy `.xls` files** that won't even open in modern Excel without complaint.

## How to actually use it

It's a small command-line tool. Once a programmer sets it up (one line: `npm install -g xlsx-for-ai`), you don't have to think about it again — your AI tools pick it up automatically and start using it whenever they encounter a spreadsheet.

If you're the programmer doing the install, the [README](README.md) has the full reference. If you're handing this to a programmer to set up for you, that link is what they'll need.

## How it works in plain terms

Today's AI is great at reasoning about text but blind to spreadsheet binaries. `xlsx-for-ai` is the translator in both directions:

- **Reading:** turns your spreadsheet into a format the AI can fully see — every formula, every formatting cue, every relationship between sheets.
- **Writing:** turns the AI's response back into a real `.xlsx` file you can open, edit, and share. The AI can now hand you a corrected workbook, not just words about it.

When the AI delivers you a corrected file, the file itself contains a small **review tab** explaining what the AI changed about the structure, why it made each call, what the risks are, and what your alternatives are if you'd prefer a different approach. The tool's design follows the *supervisor* model — it surfaces decisions for you to review, rather than silently making changes you'd discover later.

## Why this didn't exist before

Spreadsheet libraries are designed for developers building software *on top of* spreadsheets. They output JavaScript objects, database rows, raw bytes — formats other programs consume. None of them were designed for the case where the consumer is a language model and the goal is a text format the model can actually understand.

`xlsx-for-ai` is the first one built specifically for that. The output is shaped for an LLM's context window — markdown tables when the model just needs to read, structured JSON when it needs to reason, token-aware truncation when the spreadsheet is too big to fit, and a real `.xlsx` writer that produces a file you can hand back to a human along with a built-in note explaining everything that changed.

It's a small tool. It just happens to fix the one thing standing between AI assistants and the file format most knowledge work actually lives in.
