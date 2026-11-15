# Draftsman

Fill Word templates from a form or a spreadsheet. A template is an ordinary
Word document with `{{tags}}` in it; Draftsman reads the tags, shows a form
with a box for each one, and writes a finished `.docx` with your formatting,
headers, footers and tables untouched. Reusable wording lives in a **clause
library** and is ticked in as needed. One row of a spreadsheet makes one
document, so fifty letters take the same effort as one. It comes with fifty
starter templates for data protection, information security, quality and audit,
governance, complaints and everyday letters.

Everything is plain files in a folder you choose - a synced SharePoint or
OneDrive folder is fine, and then the whole team shares one set of templates.
Nothing is uploaded anywhere. Windows and macOS, the same Phosphor glass theme
as the rest of the family (Ctrl+D cycles the four themes).

---

## Which file do I use?

| File | When |
|---|---|
| `build-exe.bat` | **Windows.** Double-click once. Installs Python if it is missing, installs every dependency, and produces `dist\Draftsman.exe` - one standalone file. |
| `build-app.command` | **macOS.** Double-click once. Installs Homebrew and Python if they are missing, installs every dependency, and produces `dist/Draftsman.app`. |
| `run.bat` / `run.command` | Run from source without building. Sets up a virtual environment the first time. |
| `python draftsman.py` | Any platform, if you already have Python 3.9+ with tkinter and `pip install -r requirements.txt`. |
| `Draftsman selftest` | Makes the starter library in a temporary folder, fills every template, checks the awkward cases (tags Word has split up, clauses, conditions, drop-downs, batch, library upkeep), then drives the window. The build scripts run it on the finished app and only say DONE if it reads `SELF-TEST PASSED`. |

The repository is deliberately flat - no Windows/Mac subfolders.

On macOS, if a freshly written `.command` will not run:
`chmod +x *.command && xattr -d com.apple.quarantine *.command`.

If a build fails, look at `build-win-log.txt` / `build-mac-log.txt`; if it
builds but the self-test reported problems, `dist/draftsman-selftest.txt`; if
the app did not start at all, `dist/draftsman-crash.log`.

---

## Using it

The first run makes a starter library in `Documents/Draftsman Library`
(**Change...** points it somewhere else):

```
Draftsman Library/
  Templates/
    Subject access requests/   acknowledgement, extension, response, proof of identity,
                               third-party consent, no information held, refusal
    Individual rights/         erasure, rectification, objection or restriction
    Data protection/           breach record, breach letters (to individuals; to a controller
                               or prime contractor), DPIA and screening, legitimate interests
                               assessment, data sharing agreement, supplier questionnaire,
                               record of processing entry, privacy notice, complaint response,
                               record of destruction
    Information security/      incident report, starter-mover-leaver access request, asset
                               disposal, business continuity test, supplier security review
    Quality and audit/         audit plan and report, nonconformity and corrective action,
                               management review minutes, risk assessment, change request,
                               supplier evaluation
    Governance/                board or SMT paper, agenda, minutes, briefing note, terms of reference
    Policies/                  policy skeleton, procedure skeleton
    Complaints/                acknowledgement, outcome
    Letters and notes/         general letter, memo, file note, handover note, policy acknowledgement
    Participants/              consent form, consent withdrawal confirmation
  Clauses/
    Complaint rights.txt, Rights summary.txt, Confidentiality notice.txt, Supplementary information.docx
    DSAR exemptions/  Breach advice/  Erasure exceptions/      a folder = a tick-list of clauses
  standing values.txt              filled in on every template
  READ ME.txt                      the tag list
```

A newer Draftsman adds the starters a library has not had before the next time
it opens it (**Add starter templates** does the same for a library that did not
start as one of ours). It never overwrites a file, and a starter you delete
stays deleted.

1. Pick a template - **Find** narrows the list as you type. The form on the
   right is built from its tags; names shown dimmed are fine to leave empty.
2. Fill it in. Values carry over when you switch template, so the
   acknowledgement, the extension notice and the response for one request are
   typed once. **Save values...** keeps them in a file for another day.
3. **Preview** shows the words the document will hold - which clauses went in,
   which sections were switched off, what is still empty - without making it.
4. **Generate** (Ctrl+Enter). The document is saved (never over an existing
   file - a second one gets ` (2)`) and opened. Tick **PDF copy too** for a PDF
   beside it.

**Standing values** are the things that never change - your organisation, its
address, your name and title. Set them once; any field with the same name is
filled in on every template.

**Load values...** takes the first row of a `.csv` or `.xlsx`, a `.json`
object (what Save values writes), or a `Name = value` text file.

**Batch from a spreadsheet...** makes one document per row. The first row is
the field names; `Requester Name`, `requester_name` and `REQUESTER-NAME` are
all the same column. Whatever is typed on the form fills the fields the
spreadsheet does not have. A column headed `clauses:DSAR exemptions` picks
that row's clauses (`information about other people; negotiations`, or `all`).

**The issue log.** Every document made is added to `Draftsman log.csv` in the
folder it was saved to: when, which template, the file, its reference and
version, and who made it.

### Tags

| Tag | What it does |
|---|---|
| `{{Name}}` | A field. |
| `{{Name\|default}}` | A field with a starting value. |
| `{{Due\|today+1m}}` | A date worked out for you: `today`, `today+14`, `today+2w`, `today+1m`, `today+1y`, or `-` for the past. Shown on the form, so it can be changed. |
| `{{Outcome=Upheld\|Partly upheld\|Not upheld}}` | A drop-down of choices (anything else can still be typed). |
| `{{Address+}}` | A field with room for several lines. |
| `{{Middle Name?}}` | A field that is fine to leave empty - no warning. |
| `{{TODAY}}` `{{TODAY_SHORT}}` `{{YEAR}}` | 15 November 2026 / 15/11/2026 / 2026. |
| `{{DOC_REF}}` `{{DOC_TITLE}}` `{{VERSION}}` `{{OWNER}}` `{{ISSUE_DATE}}` `{{REVIEW_DATE}}` | The document-control block. Also written into the file's properties, and optionally added as a table at the end. |
| `{{clause:Complaint rights}}` | Drops in that clause. Alone in its own paragraph. |
| `{{clauses:DSAR exemptions}}` | A tick-list of every clause in that folder; the ticked ones go in, in file-name order (number them). Alone in its own paragraph. |
| `{{#if Extended}}` ... `{{/if}}` | Keeps the paragraphs between them only when that field is filled in and is not "no". If no field has that name it becomes a tick-box. Each alone in its own paragraph. |
| `{{#unless Extended}}` ... `{{/if}}` | The opposite. |
| `{{#else}}` | Between the two: the other case. |

Tags work in the body, tables, headers and footers. Clauses can be `.docx`
(formatting kept) or `.txt` (takes the formatting of the paragraph the tag was
in), and can hold fields of their own.

The file name box takes the same tags, plus `{{TEMPLATE}}`.

---

## Things to know

- **Word splits text into runs.** `{{Requester Name}}` can be three separate
  pieces inside the file. Draftsman matches a tag against the whole paragraph
  and writes the value back with the formatting the tag *starts* in - so make
  the whole tag bold, not half of it.
- **`#if`, `#else`, `/if` and clause tags only work in the body of the
  document**, each alone in a paragraph - not inside a table cell, a header or
  a text box. The form says so when a template breaks this.
- **Text boxes and shapes are not searched.** Put tags in ordinary text.
- **Pictures and links inside a `.docx` clause are dropped** when it is pulled
  in (they point at the clause file's own innards). Put logos in the template.
- **A template open in Word is locked** on Windows. Close it, then Refresh.
- **A PDF copy needs Microsoft Word or LibreOffice on the computer.** Draftsman
  asks whichever is there to do the conversion, because only a word processor
  lays a Word file out faithfully. With neither, you get the `.docx` and a
  message.
- The starter templates are starting points. Check the wording against your
  own procedures before first use.
