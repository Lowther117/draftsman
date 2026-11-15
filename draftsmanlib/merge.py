"""Reads the {{tags}} in a Word template and fills them in.

What a template can hold, anywhere text can go (body, tables, headers, footers):

    {{Name}}                a field: becomes a box on the form
    {{Name|default}}        a field with a starting value
    {{Due|today+1m}}        a date worked out for you: today, today+14, today+2w, today+1m, today+1y
    {{Outcome=Upheld|Partly upheld|Not upheld}}
                            a field with a drop-down of choices
    {{Address+}}            a field with room for several lines
    {{Middle Name?}}        a field that is fine to leave empty
    {{TODAY}}               today's date, 15 November 2026   ({{TODAY_SHORT}} 15/11/2026, {{YEAR}})
    {{DOC_REF}} {{VERSION}} {{OWNER}} {{ISSUE_DATE}} {{REVIEW_DATE}} {{DOC_TITLE}}
                            the document-control block

and, each alone in its own paragraph in the body of the document:

    {{clause:Complaint rights}}     drops in that one clause from the library
    {{clauses:DSAR exemptions}}     a tick-list of every clause in that folder
    {{#if Extended}} ... {{/if}}    keeps the paragraphs in between only when
                                    that field is filled in (and is not "no")
    {{#unless Extended}} ... {{/if}}    the opposite
    {{#else}}                       between the two, switches to the other case

Word often splits what looks like one piece of text into several runs, so a
tag is matched against the paragraph's whole text and then written back
across the runs it covers - the formatting of the tag's first character is
what the value gets.
"""
from __future__ import annotations

import calendar
import copy
import datetime as dt
import os
import re
from dataclasses import dataclass, field

TAG = re.compile(r"\{\{\s*([^{}]*?)\s*\}\}")
CONTROL = ("DOC_REF", "DOC_TITLE", "VERSION", "OWNER", "ISSUE_DATE", "REVIEW_DATE")
CONTROL_LABELS = {"DOC_REF": "Document reference", "DOC_TITLE": "Document title", "VERSION": "Version",
                  "OWNER": "Owner", "ISSUE_DATE": "Issue date", "REVIEW_DATE": "Review date"}
FALSE_WORDS = {"", "no", "n", "false", "0", "none", "n/a", "na"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def builtins(today=None) -> dict:
    d = today or dt.date.today()
    return {"TODAY": long_date(d), "TODAY_SHORT": d.strftime("%d/%m/%Y"), "YEAR": str(d.year)}


def long_date(d: dt.date) -> str:
    return f"{d.day} {d.strftime('%B %Y')}"


RELATIVE = re.compile(r"today(?:\s*([+-])\s*(\d{1,4})\s*(d|w|m|y|days?|weeks?|months?|years?)?)?", re.I)


def resolve_default(default: str, today=None) -> str:
    """A default of "today", "today+14", "today+2w", "today+1m" or "today+1y"
    becomes that date, written out; any other default is returned as it is."""
    m = RELATIVE.fullmatch((default or "").strip())
    if not m:
        return default or ""
    d = today or dt.date.today()
    if m.group(2):
        n = int(m.group(2)) * (-1 if m.group(1) == "-" else 1)
        unit = (m.group(3) or "d").lower()[0]
        if unit in "dw":
            d += dt.timedelta(days=n * (7 if unit == "w" else 1))
        else:
            months = d.month - 1 + n * (12 if unit == "y" else 1)
            year, month = d.year + months // 12, months % 12 + 1
            d = dt.date(year, month, min(d.day, calendar.monthrange(year, month)[1]))
    return long_date(d)


@dataclass
class Field:
    name: str
    default: str = ""
    multiline: bool = False
    choices: list = field(default_factory=list)
    optional: bool = False              # fine to leave empty: marked ?, or used as an #if condition


@dataclass
class Scan:
    fields: list = field(default_factory=list)          # Field, in the order first met
    switches: list = field(default_factory=list)        # {{#if X}} names that are not also fields
    clauses: list = field(default_factory=list)         # single clause names
    groups: list = field(default_factory=list)          # clause group (folder) names
    control: list = field(default_factory=list)         # which CONTROL fields the template uses
    problems: list = field(default_factory=list)

    def names(self):
        return [f.name for f in self.fields]


def parse_tag(inner: str):
    """('field', name, default, multiline, choices, optional) / ('clause', name) / ('clauses', group) /
    ('if', name, negated) / ('else',) / ('endif',) / ('builtin', name) / ('control', name)"""
    inner = inner.strip()
    low = inner.lower()
    if low.startswith("clause:"):
        return ("clause", inner[7:].strip())
    if low.startswith("clauses:"):
        return ("clauses", inner[8:].strip())
    if low.startswith("#if"):
        return ("if", inner[3:].strip().rstrip("+?"), False)
    if low.startswith("#unless"):
        return ("if", inner[7:].strip().rstrip("+?"), True)
    if low in ("#else", "else"):
        return ("else",)
    if low in ("/if", "#endif", "/endif", "#end", "/unless"):
        return ("endif",)
    name, _, default = inner.partition("|")
    choices = []
    if "=" in name:                       # Name=first|second|third: a drop-down
        name, _, first = name.partition("=")
        choices = [c.strip() for c in [first] + (default.split("|") if default else []) if c.strip()]
        default = ""
    name = name.strip()
    multiline = "+" in name[-2:]
    optional = "?" in name[-2:]
    name = name.rstrip("+?").strip()
    if name in ("TODAY", "TODAY_SHORT", "YEAR"):
        return ("builtin", name)
    if name in CONTROL:
        return ("control", name)
    return ("field", name, default.strip(), multiline, choices, optional)


# ------------------------------------------------------------ walking a doc
def _paragraphs_in(container):
    """Every paragraph in a body, header, footer or cell, tables included."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    element = container._element
    if element.tag == W + "document":          # a Document's own element wraps the body
        element = element.body
    for child in element.iterchildren():
        if child.tag == W + "p":
            yield Paragraph(child, container)
        elif child.tag == W + "tbl":
            table = Table(child, container)
            seen = set()
            for row in table.rows:
                for cell in row.cells:
                    if cell._tc in seen:            # merged cells repeat. (The element itself, not
                        continue                    # its id(): lxml reuses ids once a proxy is freed.)
                    seen.add(cell._tc)
                    yield from _paragraphs_in(cell)
        elif child.tag == W + "sdt":                      # a content control wraps its paragraphs
            for p in child.iter(W + "p"):
                yield Paragraph(p, container)


def all_paragraphs(document):
    yield from _paragraphs_in(document)
    seen = set()
    for section in document.sections:
        for part in (section.header, section.first_page_header, section.even_page_header,
                     section.footer, section.first_page_footer, section.even_page_footer):
            try:
                if part.is_linked_to_previous or part._element in seen:
                    continue
                seen.add(part._element)
                yield from _paragraphs_in(part)
            except Exception:
                continue


def para_text(p_element) -> str:
    return "".join(t.text or "" for t in p_element.iter(W + "t"))


def sole_tag(text: str):
    """The parsed tag when a paragraph holds one tag and nothing else."""
    m = TAG.fullmatch(text.strip())
    return parse_tag(m.group(1)) if m else None


# ------------------------------------------------------------------- scan
def scan_text(text: str, scan: Scan, seen: dict, conditions: list):
    for m in TAG.finditer(text):
        tag = parse_tag(m.group(1))
        kind = tag[0]
        if kind == "field":
            if not tag[1]:
                continue
            f = seen.get(tag[1])
            if f is None:
                f = seen[tag[1]] = Field(tag[1], tag[2], tag[3], list(tag[4]), tag[5])
                scan.fields.append(f)
            else:
                f.default = f.default or tag[2]
                f.multiline = f.multiline or tag[3]
                f.choices = f.choices or list(tag[4])
                f.optional = f.optional or tag[5]
        elif kind == "control" and tag[1] not in scan.control:
            scan.control.append(tag[1])
        elif kind == "clause" and tag[1] not in scan.clauses:
            scan.clauses.append(tag[1])
        elif kind == "clauses" and tag[1] not in scan.groups:
            scan.groups.append(tag[1])
        elif kind == "if" and tag[1] and tag[1] not in conditions:
            conditions.append(tag[1])


def scan_template(path, library=None) -> Scan:
    """What a template asks for. With a library, the fields inside the clauses
    it can pull in are included too."""
    import docx
    scan, seen, conditions = Scan(), {}, []
    document = docx.Document(path)
    depth = 0
    top_level = {child for child in document.element.body.iterchildren() if child.tag == W + "p"}
    for p in all_paragraphs(document):
        text = p.text
        if "{{" not in text:
            continue
        scan_text(text, scan, seen, conditions)
        tag = sole_tag(text)
        if tag and tag[0] in ("if", "else", "endif", "clause", "clauses") and p._p not in top_level:
            scan.problems.append(f"\"{text.strip()[:60]}\" - this tag only works in the body of the document, "
                                 "not inside a table, header or footer.")
            continue
        if tag and tag[0] == "if":
            depth += 1
        elif tag and tag[0] == "else":
            if depth < 1:
                scan.problems.append("There is an {{#else}} outside any {{#if ...}}.")
        elif tag and tag[0] == "endif":
            depth -= 1
            if depth < 0:
                scan.problems.append("There is an {{/if}} with no {{#if ...}} before it.")
                depth = 0
        elif tag and tag[0] in ("clause", "clauses"):
            pass
        elif any(parse_tag(m.group(1))[0] in ("if", "else", "endif", "clause", "clauses") for m in TAG.finditer(text)):
            scan.problems.append(f"\"{text.strip()[:60]}\" - an #if, #else, /if or clause tag only works alone in its own paragraph.")
    if depth:
        scan.problems.append("An {{#if ...}} is never closed with {{/if}}.")
    if library is not None:
        for name in scan.clauses:
            if library.clause_path(name) is None:
                scan.problems.append(f"Clause \"{name}\" is not in the library's clauses folder.")
        for group in scan.groups:
            if not library.clauses_in(group):
                scan.problems.append(f"Clause folder \"{group}\" is missing or empty.")
        for cpath in ([library.clause_path(n) for n in scan.clauses]
                      + [c for g in scan.groups for _n, c in library.clauses_in(g)]):
            if cpath:
                try:
                    for line in clause_lines(cpath):
                        scan_text(line, scan, seen, conditions)
                except Exception as exc:  # noqa: BLE001
                    scan.problems.append(f"Clause {os.path.basename(cpath)} could not be read: {exc}")
    scan.switches = [c for c in conditions if c not in seen and c not in CONTROL]
    for name in conditions:                 # a field that switches a section on is allowed to be empty
        if name in seen:
            seen[name].optional = True
    return scan


def clause_lines(path) -> list:
    """A clause's text, a line per paragraph (for finding the fields in it)."""
    if path.lower().endswith(".docx"):
        import docx
        return [p.text for p in _paragraphs_in(docx.Document(path))]
    with open(path, encoding="utf-8-sig", errors="replace") as fh:
        return [ln.rstrip("\n") for ln in fh]


# ------------------------------------------------------------------- fill
def replace_in_paragraph(paragraph, lookup) -> int:
    """Replace every {{tag}} in a paragraph, in place. lookup(inner) returns
    the text, or None to leave that tag as it is."""
    done, pos = 0, 0
    while True:
        runs = paragraph.runs
        texts = [r.text for r in runs]
        full = "".join(texts)
        m = TAG.search(full, pos)
        if not m:
            return done
        value = lookup(m.group(1))
        if value is None:
            pos = m.end()
            continue
        start, end = m.start(), m.end()
        # which run holds the first and the last character of the tag
        offs, first, last = 0, None, None
        for i, txt in enumerate(texts):
            if first is None and start < offs + len(txt):
                first = (i, start - offs)
            if end <= offs + len(txt):
                last = (i, end - offs)
                break
            offs += len(txt)
        if first is None or last is None:       # text that is not in plain runs (a hyperlink, a field)
            pos = m.end()
            continue
        (fi, fo), (li, lo) = first, last
        if fi == li:
            runs[fi].text = texts[fi][:fo] + value + texts[fi][lo:]
        else:
            runs[fi].text = texts[fi][:fo] + value
            for i in range(fi + 1, li):
                runs[i].text = ""
            runs[li].text = texts[li][lo:]
        pos = start + len(value)
        done += 1


def truthy(value) -> bool:
    return str(value or "").strip().lower() not in FALSE_WORDS


def _strip_relationships(element):
    """A clause comes from another file, and anything in it that points at
    that file's own parts (a picture, a hyperlink target) would point at
    nothing here and make Word call the result corrupt."""
    for bad in list(element.iter(W + "drawing", W + "pict", W + "object")):
        bad.getparent().remove(bad)
    for el in element.iter():
        for attr in [a for a in el.attrib if a.startswith(R_NS)]:
            del el.attrib[attr]
    for bad in list(element.iter(W + "sectPr")):
        bad.getparent().remove(bad)


def insert_clause(marker, path) -> int:
    """Put a clause's paragraphs where the marker paragraph is. Returns how
    many blocks went in. The marker itself is removed by the caller."""
    n = 0
    if path.lower().endswith(".docx"):
        import docx
        source = docx.Document(path)
        for child in source.element.body.iterchildren():
            if child.tag not in (W + "p", W + "tbl"):
                continue
            new = copy.deepcopy(child)
            _strip_relationships(new)
            marker._p.addprevious(new)
            n += 1
    else:
        for line in clause_lines(path):
            if line.strip():
                marker.insert_paragraph_before(line.strip(), style=marker.style)
                n += 1
    return n


def _expand_clauses(document, library, choices, report):
    """Swap each {{clause:..}} / {{clauses:..}} paragraph for its clauses."""
    for p in list(_paragraphs_in(document)):
        tag = sole_tag(p.text) if "{{" in p.text else None
        if not tag or tag[0] not in ("clause", "clauses"):
            continue
        if tag[0] == "clause":
            cpath = library.clause_path(tag[1]) if library else None
            if cpath:
                insert_clause(p, cpath)
            else:
                report["missing_clauses"].append(tag[1])
        else:
            available = dict(library.clauses_in(tag[1])) if library else {}
            for name in choices.get(tag[1], []):
                if name in available:
                    insert_clause(p, available[name])
                else:
                    report["missing_clauses"].append(f"{tag[1]} / {name}")
        p._p.getparent().remove(p._p)


def _apply_conditions(document, values):
    """Remove {{#if X}}..{{/if}} blocks whose X is empty or 'no' (and
    {{#unless X}} blocks whose X is not); {{#else}} swaps which half is kept.
    The marker paragraphs go either way."""
    body = document.element.body
    stack = []
    for child in list(body.iterchildren()):
        if child.tag == W + "sectPr":
            continue
        tag = sole_tag(para_text(child)) if child.tag == W + "p" else None
        if tag and tag[0] == "if":
            stack.append(truthy(values.get(tag[1], "")) != tag[2])
            body.remove(child)
        elif tag and tag[0] == "else":
            if stack:
                stack[-1] = not stack[-1]
            body.remove(child)
        elif tag and tag[0] == "endif":
            if stack:
                stack.pop()
            body.remove(child)
        elif stack and not all(stack):
            body.remove(child)


def add_control_table(document, control):
    document.add_paragraph()
    rows = [(CONTROL_LABELS[k], control.get(k, "")) for k in CONTROL if control.get(k, "")]
    if not rows:
        return
    table = document.add_table(rows=len(rows), cols=2)
    try:
        table.style = "Table Grid"
    except Exception:
        pass
    for i, (label, value) in enumerate(rows):
        table.cell(i, 0).text = label
        table.cell(i, 1).text = value
        for run in table.cell(i, 0).paragraphs[0].runs:
            run.bold = True


def render(template, out_path, values, library=None, choices=None, control=None,
           control_table=False, today=None) -> dict:
    """Fill a template and save it. Returns a report: fields left empty,
    clauses that could not be found, tags that were not understood."""
    import docx
    values = {k: ("" if v is None else str(v)) for k, v in (values or {}).items()}
    control = {k: str(v or "") for k, v in (control or {}).items()}
    choices = choices or {}
    report = dict(path=out_path, empty=[], missing_clauses=[], replaced=0)
    document = docx.Document(template)
    _expand_clauses(document, library, choices, report)
    _apply_conditions(document, {**control, **values})
    fixed = builtins(today)

    def lookup(inner):
        tag = parse_tag(inner)
        if tag[0] == "builtin":
            return fixed[tag[1]]
        if tag[0] == "control":
            return control.get(tag[1], "")
        if tag[0] == "field":
            if not tag[1]:
                return None
            value = values.get(tag[1], "")
            if not value.strip():
                value = resolve_default(tag[2], today)
            if not value.strip() and tag[1] not in report["empty"]:
                report["empty"].append(tag[1])
            return value
        return None                    # a stray #if or clause tag: left visible so it gets noticed

    for p in all_paragraphs(document):
        if "{{" in p.text:
            report["replaced"] += replace_in_paragraph(p, lookup)
    props = document.core_properties
    try:
        if control.get("DOC_TITLE"):
            props.title = control["DOC_TITLE"][:255]
        if control.get("OWNER"):
            props.author = control["OWNER"][:255]
        if control.get("VERSION"):
            props.version = control["VERSION"][:255]
        if control.get("DOC_REF"):
            props.identifier = control["DOC_REF"][:255]
        if control.get("REVIEW_DATE"):
            props.comments = f"Review date: {control['REVIEW_DATE']}"[:255]
        props.last_modified_by = (control.get("OWNER") or props.last_modified_by or "")[:255]
        props.modified = dt.datetime.now()
    except Exception:
        pass
    if control_table:
        add_control_table(document, control)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    document.save(out_path)
    return report


def leftover_tags(path) -> list:
    """Any {{...}} still in a finished document."""
    import docx
    out = []
    for p in all_paragraphs(docx.Document(path)):
        out.extend(m.group(0) for m in TAG.finditer(p.text))
    return out


def document_text(path) -> str:
    """A finished document as plain text, in reading order, for the preview:
    headers, then the body with each table row on a line, then footers."""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    document = docx.Document(path)
    lines = []

    def block(container, element):
        for child in element.iterchildren():
            if child.tag == W + "p":
                lines.append(Paragraph(child, container).text)
            elif child.tag == W + "tbl":
                for row in Table(child, container).rows:
                    cells, seen = [], set()
                    for cell in row.cells:
                        if cell._tc in seen:
                            continue
                        seen.add(cell._tc)
                        cells.append(cell.text.strip().replace("\n", " / "))
                    lines.append("   |   ".join(cells))
                lines.append("")

    section = document.sections[0]
    head = [p.text for p in section.header.paragraphs if p.text.strip()]
    if head:
        lines += ["[header] " + h for h in head] + [""]
    block(document, document.element.body)
    foot = [p.text for p in section.footer.paragraphs if p.text.strip()]
    if foot:
        lines += [""] + ["[footer] " + f for f in foot]
    return "\n".join(lines)
