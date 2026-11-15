"""Self-test: makes the starter library in a temporary folder, fills every
template, and checks the awkward cases. Run by the build scripts on the
finished app; also `python draftsman.py selftest [--no-gui]`."""
from __future__ import annotations

import csv
import datetime as dt
import os
import shutil
import sys
import tempfile
import traceback

from . import APP_NAME, __version__, batch, merge, paths, starter
from .library import Library, file_name, match_keys, parse_pairs, read_rows, unique_path

REPORT = "draftsman-selftest.txt"


def _text(path) -> str:
    import docx
    return "\n".join(p.text for p in merge.all_paragraphs(docx.Document(path)))


def run_checks(work, gui=True):
    import docx
    results = []

    def check(label, fn):
        try:
            detail = fn()
            results.append((True, label, detail if isinstance(detail, str) else ""))
        except Exception as exc:  # noqa: BLE001
            results.append((False, label, f"{type(exc).__name__}: {exc}\n" + traceback.format_exc(limit=4)))

    root = os.path.join(work, "Library")
    lib = Library(root)
    out = os.path.join(work, "out")
    os.makedirs(out)

    def t_starter():
        made = starter.build(root)
        assert len(lib.templates()) == len(starter.T) >= 45, len(lib.templates())
        assert starter.build(root) == [], "building twice overwrote something"
        assert [n for n, _p in lib.clauses_in("DSAR exemptions")][0] == "1 Information about other people"
        assert lib.clause_path("Complaint rights") and lib.clause_path("Supplementary information")
        assert lib.clause_path("../standing values") is None
        problems = []
        for _folder, name, path in lib.templates():
            scan = merge.scan_template(path, lib)
            problems += [f"{name}: {p}" for p in scan.problems]
        assert not problems, problems
        return f"{len(made)} files written; every template reads cleanly"
    check("Starter library", t_starter)

    def t_tags():
        assert merge.parse_tag("Name|Sir or Madam") == ("field", "Name", "Sir or Madam", False, [], False)
        assert merge.parse_tag(" Address+ ") == ("field", "Address", "", True, [], False)
        assert merge.parse_tag("Note?+") == ("field", "Note", "", True, [], True) == merge.parse_tag("Note+?")
        assert merge.parse_tag("Outcome=Upheld|Partly upheld| Not upheld ") == ("field", "Outcome", "", False, ["Upheld", "Partly upheld", "Not upheld"], False)
        assert merge.parse_tag("clauses: DSAR exemptions") == ("clauses", "DSAR exemptions")
        assert merge.parse_tag("#if Extended") == ("if", "Extended", False) and merge.parse_tag("/if") == ("endif",)
        assert merge.parse_tag("#unless Extended") == ("if", "Extended", True) and merge.parse_tag("#else") == ("else",)
        day = dt.date(2026, 1, 31)
        assert [merge.resolve_default(x, day) for x in ("today", "today+1m", "today + 2w", "today-1y", "today+30", "Today is fine")] == \
            ["31 January 2026", "28 February 2026", "14 February 2026", "31 January 2025", "2 March 2026", "Today is fine"]
        assert merge.parse_tag("VERSION") == ("control", "VERSION") and merge.parse_tag("TODAY") == ("builtin", "TODAY")
        assert merge.builtins(dt.date(2026, 11, 5)) == {"TODAY": "5 November 2026", "TODAY_SHORT": "05/11/2026", "YEAR": "2026"}
        assert parse_pairs("# c\nOrganisation = Example Org\nAddress = 1 High St\n  Town\n\nX=1") == \
            {"Organisation": "Example Org", "Address": "1 High St\nTown", "X": "1"}
        assert not merge.truthy("No") and not merge.truthy("  ") and merge.truthy("Yes")
        return "fields, defaults, choices, optional, relative dates, clauses, if/unless/else, standing values"
    check("Reading tags", t_tags)

    def t_split_runs():
        d = docx.Document()
        p = d.add_paragraph()
        for chunk, bold in (("Dear {", False), ("{Na", True), ("me}", True), ("}, your ref is {{", False),
                            ("Ref|R-1}} and {{Name}} again. {{ TODAY }}", False)):
            p.add_run(chunk).bold = bold
        d.sections[0].header.paragraphs[0].add_run("Header for {{Name}}")
        cell = d.add_table(rows=1, cols=1).cell(0, 0)
        cell.text = "In a table: {{Name}} / {{Lines+}}"
        src, dst = os.path.join(work, "split.docx"), os.path.join(out, "split.docx")
        d.save(src)
        scan = merge.scan_template(src)
        assert scan.names() == ["Name", "Ref", "Lines"] and scan.fields[1].default == "R-1" and scan.fields[2].multiline, scan
        rep = merge.render(src, dst, {"Name": "Dr {{Odd}} Jones", "Lines": "one\ntwo"}, today=dt.date(2026, 11, 5))
        done = docx.Document(dst)
        first = done.paragraphs[0]
        assert first.text == "Dear Dr {{Odd}} Jones, your ref is R-1 and Dr {{Odd}} Jones again. 5 November 2026", first.text
        assert first.runs[0].text == "Dear Dr {{Odd}} Jones" and not first.runs[0].bold, [r.text for r in first.runs]
        assert done.sections[0].header.paragraphs[0].text == "Header for Dr {{Odd}} Jones"
        cell = done.tables[0].cell(0, 0)
        assert cell.text == "In a table: Dr {{Odd}} Jones / one\ntwo", repr(cell.text)
        assert rep["replaced"] == 7 and rep["empty"] == [], rep
        return "a tag split across runs is filled with the formatting it starts in; header, table and line breaks filled; a value containing {{ }} is left alone"
    check("Filling tags that Word has split up", t_split_runs)

    values = {"Organisation": "Example Org", "Organisation Address": "1 High Street\nMerthyr Tydfil", "Requester Name": "Alex Morgan",
              "Requester Address": "2 Low Road", "Reference": "SAR-2026-014", "Date Received": "6 March 2026",
              "Due Date": "7 April 2026", "Handler Name": "Sam Rees", "Handler Title": "Data Protection Officer",
              "Contact Email": "dpo@example.org", "Bundle Pages": "42", "Redactions Applied": "Yes", "Information Withheld": "No"}
    response = os.path.join(root, "Templates", "Subject access requests", "3 Response.docx")

    def t_response():
        scan = merge.scan_template(response, lib)
        assert scan.groups == ["DSAR exemptions"] and "Retention Period" in scan.names(), scan     # a field from inside a clause
        dst = os.path.join(out, "response.docx")
        picks = {"DSAR exemptions": ["1 Information about other people", "3 Confidential references"]}
        rep = merge.render(response, dst, values, lib, picks)
        text = _text(dst)
        assert not merge.leftover_tags(dst), merge.leftover_tags(dst)
        assert "would identify them" in text and "reference given in confidence" in text and "legal advisers" not in text
        assert text.index("would identify them") < text.index("reference given in confidence")
        assert "withheld in full" not in text and "Each place is marked REDACTED" in text       # the two conditions
        assert "How long we keep it: In line with our retention schedule" in text                 # docx clause + default
        assert "contact Sam Rees at dpo@example.org" in text and "Dear Sir or Madam" in text
        assert "SAR-2026-014 - response" in text                                                  # the footer
        assert rep["empty"] == [] and rep["missing_clauses"] == [], rep
        rep = merge.render(response, os.path.join(out, "response2.docx"), dict(values, **{"Redactions Applied": "No"}), lib, picks)
        text = _text(os.path.join(out, "response2.docx"))
        assert "REDACTED" not in text and "would identify them" not in text
        return "clauses picked and dropped in order, conditions applied, clause fields and footer filled, nothing left over"
    check("A full letter: clauses and conditions", t_response)

    def t_every_template():
        n = 0
        for _folder, name, path in lib.templates():
            scan = merge.scan_template(path, lib)
            vals = {f: f"<{f}>" for f in scan.names()}
            vals.update({s: "yes" for s in scan.switches})
            ctl = {k: f"<{k}>" for k in merge.CONTROL}
            dst = os.path.join(out, f"all - {name}.docx")
            merge.render(path, dst, vals, lib, {g: [c for c, _p in lib.clauses_in(g)] for g in scan.groups}, ctl, control_table=True)
            left = merge.leftover_tags(dst)
            assert not left, (name, left)
            docx.Document(dst)
            n += 1
        policy = os.path.join(out, "all - Policy skeleton.docx")
        props = docx.Document(policy).core_properties
        assert props.title == "<DOC_TITLE>" and props.author == "<OWNER>" and props.version == "<VERSION>", (props.title, props.author)
        assert "<DOC_REF>  |  <DOC_TITLE>  |  Version <VERSION>" in _text(policy)
        assert docx.Document(policy).tables[-1].cell(0, 0).text == "Document reference"
        return f"{n} templates filled with every clause; document-control block and properties stamped"
    check("Every starter template, everything switched on", t_every_template)

    def t_batch():
        import openpyxl
        ack = os.path.join(root, "Templates", "Subject access requests", "1 Acknowledgement.docx")
        scan = merge.scan_template(ack, lib)
        assert set(scan.switches) == set() and "Clarification Needed" in scan.names()
        sheet = os.path.join(work, "rows.csv")
        with open(sheet, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["reference", "Requester_Name", "Date Received", "Due Date", "Clarification Needed"])
            w.writerow(["SAR/1", "Alex Morgan", "6 March 2026", "7 April 2026", ""])
            w.writerow(["SAR/2", "Bo Li", "9 March 2026", "9 April 2026", "which service you used"])
            w.writerow(["SAR/2", "Bo Li", "9 March 2026", "9 April 2026", ""])
        rows = read_rows(sheet)
        assert len(rows) == 3 and match_keys(rows[0], ["Requester Name", "Reference"]) == {"Requester Name": "Alex Morgan", "Reference": "SAR/1"}
        reports = batch.run(ack, rows, os.path.join(out, "batch"), "{{Reference}} - {{Requester Name}}", scan, lib,
                            base={"Organisation": "Example Org", "Handler Name": "Sam Rees"})
        names = [os.path.basename(r["path"]) for r in reports]
        assert names == ["SAR-1 - Alex Morgan.docx", "SAR-2 - Bo Li.docx", "SAR-2 - Bo Li (2).docx"], names
        one, two = _text(reports[0]["path"]), _text(reports[1]["path"])
        assert "paused" not in one and "which service you used" in two and "paused" in two
        assert "Example Org" in one and "Requester Address" in reports[0]["empty"]
        wb = openpyxl.Workbook()
        wb.active.append(["Reference", "Received", "clauses:DSAR exemptions"])
        wb.active.append(["SAR/9", dt.datetime(2026, 3, 6), "information about other people; 5 Negotiations"])
        xl = os.path.join(work, "rows.xlsx")
        wb.save(xl)
        row = read_rows(xl)[0]
        assert row["Received"] == "6 March 2026", row
        rscan = merge.scan_template(response, lib)
        picks = batch.row_choices(row, rscan, lib, {})
        assert picks == {"DSAR exemptions": ["1 Information about other people", "5 Negotiations"]}, picks
        assert file_name("{{A}}: {{B}}?", {"A": "x/y", "B": ""}, "fallback") == "x-y" and file_name("", {}, "fallback") == "fallback"
        assert unique_path(out, "nothing here").endswith("nothing here.docx")
        return "CSV and Excel rows, loose column names, safe unique file names, per-row clauses and conditions"
    check("Batch from a spreadsheet", t_batch)

    def t_conditions():
        d = docx.Document()
        for line in ("Start", "{{#if Extended}}", "Extended to {{New Date|today+1m}}.", "{{#else}}", "On time.", "{{/if}}",
                     "{{#unless Fee?}}", "No fee.", "{{/if}}", "Outcome: {{Outcome=Upheld|Not upheld}} {{Note?}}", "End"):
            d.add_paragraph(line)
        src = os.path.join(work, "cond.docx")
        d.save(src)
        scan = merge.scan_template(src)
        assert not scan.problems, scan.problems
        assert scan.switches == ["Extended", "Fee"] and scan.names() == ["New Date", "Outcome", "Note"], scan
        by = {f.name: f for f in scan.fields}
        assert by["Outcome"].choices == ["Upheld", "Not upheld"] and by["Note"].optional and not by["Outcome"].optional
        day = dt.date(2026, 1, 31)
        one, two = os.path.join(out, "cond1.docx"), os.path.join(out, "cond2.docx")
        rep = merge.render(src, one, {"Extended": "yes", "Outcome": "Upheld"}, today=day)
        assert _text(one).split("\n") == ["Start", "Extended to 28 February 2026.", "No fee.", "Outcome: Upheld ", "End"], _text(one)
        assert rep["empty"] == ["Note"], rep
        merge.render(src, two, {"Fee": "yes"}, today=day)
        assert _text(two).split("\n") == ["Start", "On time.", "Outcome:  ", "End"], _text(two)
        assert "Outcome: Upheld" in merge.document_text(one)
        bad = docx.Document()
        bad.add_paragraph("{{#else}}")
        bad.add_paragraph("Text {{#if X}} inline")
        bad.add_table(rows=1, cols=1).cell(0, 0).text = "{{clause:Complaint rights}}"
        bad.save(os.path.join(work, "bad.docx"))
        assert len(merge.scan_template(os.path.join(work, "bad.docx")).problems) == 3
        return "if / else / unless keep the right half; optional and drop-down fields; relative dates; bad tags reported"
    check("Conditions, choices and optional fields", t_conditions)

    def t_upkeep():
        from .library import log_issue, save_values
        # a starter that was deleted stays deleted; one the library never had arrives
        gone = os.path.join(root, "Templates", "Letters and notes", "Memo.docx")
        os.remove(gone)
        assert starter.build(root) == [] and not os.path.exists(gone)
        with open(os.path.join(root, starter.MANIFEST), encoding="utf-8") as fh:
            listed = [ln.strip() for ln in fh if ln.strip()]
        with open(os.path.join(root, starter.MANIFEST), "w", encoding="utf-8") as fh:
            fh.write("\n".join(x for x in listed if not x.endswith("File note.docx")) + "\n")
        note = os.path.join(root, "Templates", "Letters and notes", "File note.docx")
        with open(note, "wb") as fh:
            fh.write(b"mine now")
        assert starter.build(root) == [] and open(note, "rb").read() == b"mine now"        # never overwrites
        os.remove(note)
        starter.build(root)                 # it is on the list now, so it does not come back either
        assert not os.path.exists(note)
        # a version 1.0 library (no list, the letter in its old folder) is topped up without a second letter
        old = os.path.join(work, "Old library")
        os.makedirs(os.path.join(old, "Templates", "Letters"))
        with open(os.path.join(old, "READ ME.txt"), "w", encoding="utf-8") as fh:
            fh.write("DRAFTSMAN LIBRARY\n")
        with open(os.path.join(old, "Templates", "Letters", "General letter.docx"), "wb") as fh:
            fh.write(b"x")
        assert starter.is_ours(old) and not starter.is_ours(work)
        made = starter.build(old)
        assert len(made) == len(starter.starter_files()) - 1, len(made)
        assert not os.path.exists(os.path.join(old, "Templates", "Letters and notes", "General letter.docx"))
        log_issue(out, "Memo", "a.docx", "REF-1", "2.0")
        log_issue(out, "Memo", "b.docx")
        with open(os.path.join(out, "Draftsman log.csv"), encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        assert rows[0][:3] == ["When", "Template", "Document"] and rows[-2][2:5] == ["a.docx", "REF-1", "2.0"] and len(rows) >= 3, rows
        kept = os.path.join(work, "kept.json")
        save_values(kept, {"Reference": "R-9", "Blank": " ", "clauses:DSAR exemptions": "5 Negotiations"})
        assert read_rows(kept) == [{"Reference": "R-9", "clauses:DSAR exemptions": "5 Negotiations"}]
        return "deleted starters stay deleted, new ones arrive, nothing overwritten; issue log; saved values load back"
    check("Library upkeep, issue log and saved values", t_upkeep)

    if gui:
        def t_gui():
            from . import ui
            return ui.selftest(root, out)
        check("The window: every template on the form, generate, preview, find, all four themes", t_gui)
    return results


def main(argv=None) -> int:
    argv = list(argv or [])
    gui = "--no-gui" not in argv
    work = tempfile.mkdtemp(prefix="draftsman-selftest-")
    try:
        results = run_checks(work, gui=gui)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    bad = [r for r in results if not r[0]]
    lines = [f"{APP_NAME} {__version__} self-test - Python {sys.version.split()[0]} on {sys.platform}", ""]
    for ok, label, detail in results:
        lines.append(f"  {'ok  ' if ok else 'FAIL'}  {label}")
        if detail:
            lines += ["          " + ln for ln in detail.strip().splitlines()]
    if not gui:
        lines.append("  --    The window was not tested (--no-gui)")
    lines += ["", "SELF-TEST PASSED" if not bad else f"PROBLEMS FOUND: {len(bad)}", ""]
    text = "\n".join(lines)
    try:
        with open(os.path.join(paths.data_dir(), REPORT), "w", encoding="utf-8") as fh:
            fh.write(text)
    except OSError:
        pass
    try:
        print(text)
    except Exception:
        pass
    return 0 if not bad else 1
