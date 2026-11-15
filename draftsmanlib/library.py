"""The library: a plain folder of templates and clauses that can live anywhere,
a synced SharePoint or OneDrive folder included, so a team shares one set.

    <library>/Templates/<any folders>/<name>.docx
    <library>/Clauses/<name>.docx or .txt                 single clauses
    <library>/Clauses/<group>/<name>.docx or .txt         a tick-list group
    <library>/standing values.txt                         Name = value, one per line
"""
from __future__ import annotations

import csv
import json
import os
import re

TEMPLATES, CLAUSES, STANDING = "Templates", "Clauses", "standing values.txt"
CLAUSE_EXT = (".docx", ".txt")


def _visible(name: str) -> bool:
    return not name.startswith((".", "~$", "_"))


def natural(text: str):
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", text)]


class Library:
    def __init__(self, root: str):
        self.root = os.path.abspath(root)

    @property
    def templates_dir(self):
        return os.path.join(self.root, TEMPLATES)

    @property
    def clauses_dir(self):
        return os.path.join(self.root, CLAUSES)

    def ensure(self):
        os.makedirs(self.templates_dir, exist_ok=True)
        os.makedirs(self.clauses_dir, exist_ok=True)

    def is_empty(self) -> bool:
        return not self.templates()

    def templates(self) -> list:
        """(folder shown, name, path) for every template, folders first by name."""
        out = []
        base = self.templates_dir
        for folder, dirs, files in os.walk(base):
            dirs[:] = sorted((d for d in dirs if _visible(d)), key=natural)
            rel = os.path.relpath(folder, base)
            for f in sorted(files, key=natural):
                if f.lower().endswith((".docx", ".dotx")) and _visible(f):
                    out.append(("" if rel == "." else rel.replace(os.sep, " / "), os.path.splitext(f)[0],
                                os.path.join(folder, f)))
        return out

    def clause_path(self, name: str):
        """A single clause by name, with or without its extension; it may sit
        in a sub-folder, written "Folder/Name"."""
        name = name.strip().replace("\\", "/")
        parts = [p for p in name.split("/") if p and p not in (".", "..")]
        if not parts:
            return None
        base = os.path.join(self.clauses_dir, *parts)
        for candidate in [base] + [base + ext for ext in CLAUSE_EXT]:
            if os.path.isfile(candidate) and candidate.lower().endswith(CLAUSE_EXT):
                return candidate
        return None

    def clauses_in(self, group: str) -> list:
        """(name, path) for each clause in a group folder, in natural order -
        number the files (1 ..., 2 ...) to fix the order they go in."""
        parts = [p for p in group.strip().replace("\\", "/").split("/") if p and p not in (".", "..")]
        folder = os.path.join(self.clauses_dir, *parts) if parts else ""
        if not parts or not os.path.isdir(folder):
            return []
        return [(os.path.splitext(f)[0], os.path.join(folder, f))
                for f in sorted(os.listdir(folder), key=natural)
                if f.lower().endswith(CLAUSE_EXT) and _visible(f) and os.path.isfile(os.path.join(folder, f))]

    # ----------------------------------------------------- standing values
    @property
    def standing_path(self):
        return os.path.join(self.root, STANDING)

    def standing_text(self) -> str:
        try:
            with open(self.standing_path, encoding="utf-8-sig") as fh:
                return fh.read()
        except OSError:
            return ""

    def standing(self) -> dict:
        return parse_pairs(self.standing_text())

    def save_standing(self, text: str):
        os.makedirs(self.root, exist_ok=True)
        with open(self.standing_path, "w", encoding="utf-8") as fh:
            fh.write(text.rstrip() + "\n")


def parse_pairs(text: str) -> dict:
    """Name = value lines. A value runs on over following lines that start
    with a space or tab (an address), and # starts a comment line."""
    out, last = {}, None
    for line in (text or "").splitlines():
        if line.strip().startswith("#") or not line.strip():
            last = None if not line.strip() else last
            continue
        if line[:1] in " \t" and last:
            out[last] += "\n" + line.strip()
        elif "=" in line:
            name, _, value = line.partition("=")
            last = name.strip()
            if last:
                out[last] = value.strip()
        else:
            last = None
    return out


# ------------------------------------------------------- values from files
def read_rows(path) -> list:
    """Rows of a .csv or .xlsx as dicts keyed by the header row; or the one
    'row' in a .json object or a Name = value text file."""
    low = path.lower()
    if low.endswith((".xlsx", ".xlsm")):
        import openpyxl
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            rows = list(book.worksheets[0].iter_rows(values_only=True))
        finally:
            book.close()
        rows = [r for r in rows if r and any(c not in (None, "") for c in r)]
        if not rows:
            return []
        header = [str(h).strip() if h is not None else "" for h in rows[0]]
        return [{h: _cell(v) for h, v in zip(header, r) if h} for r in rows[1:]]
    if low.endswith(".json"):
        with open(path, encoding="utf-8-sig") as fh:
            data = json.load(fh)
        data = data if isinstance(data, list) else [data]
        return [{str(k): "" if v is None else str(v) for k, v in d.items()} for d in data if isinstance(d, dict)]
    if low.endswith(".txt"):
        with open(path, encoding="utf-8-sig") as fh:
            pairs = parse_pairs(fh.read())
        return [pairs] if pairs else []
    with open(path, encoding="utf-8-sig", newline="") as fh:
        sample = fh.read(4096)
        fh.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(fh, dialect=dialect)
        return [{(k or "").strip(): (v or "").strip() for k, v in row.items() if k} for row in reader
                if any((v or "").strip() for v in row.values() if isinstance(v, str))]


def _cell(value) -> str:
    import datetime as dt
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        if value.hour or value.minute:
            return f"{value.day} {value.strftime('%B %Y %H:%M')}"
        return f"{value.day} {value.strftime('%B %Y')}"
    if isinstance(value, dt.date):
        return f"{value.day} {value.strftime('%B %Y')}"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def match_keys(row: dict, wanted) -> dict:
    """A row's values under the template's own field names: 'requester name',
    'Requester_Name' and 'Requester Name' are the same column."""
    def norm(s):
        return re.sub(r"[\s_\-]+", " ", str(s)).strip().lower()
    by = {norm(k): v for k, v in row.items()}
    out = {}
    for name in wanted:
        if norm(name) in by:
            out[name] = by[norm(name)]
    return out


ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def file_name(pattern: str, values: dict, fallback: str) -> str:
    """A safe file name from a pattern such as "{{Reference}} - {{Requester Name}}"."""
    from .merge import TAG, builtins, parse_tag
    fixed = builtins()

    def sub(m):
        tag = parse_tag(m.group(1))
        if tag[0] == "builtin":
            return fixed[tag[1]].replace("/", "-")
        if tag[0] in ("field", "control"):
            return str(values.get(tag[1], "")).splitlines()[0] if str(values.get(tag[1], "")).strip() else ""
        return ""
    name = TAG.sub(sub, pattern or "")
    name = ILLEGAL.sub("-", name)
    name = re.sub(r"\s+", " ", name).strip(" .-")
    return (name or fallback)[:150]


def unique_path(folder: str, name: str, ext: str = ".docx") -> str:
    path = os.path.join(folder, name + ext)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{name} ({n}){ext}")
        n += 1
    return path


LOG_NAME = "Draftsman log.csv"


def log_issue(folder: str, template: str, document: str, reference: str = "", version: str = "") -> None:
    """Add a line to the issue log kept in the output folder: what was made,
    from which template, when and by whom. Never stops a document being made."""
    import datetime as dt
    import getpass
    path = os.path.join(folder, LOG_NAME)
    try:
        new = not os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8-sig" if new else "utf-8") as fh:
            w = csv.writer(fh)
            if new:
                w.writerow(["When", "Template", "Document", "Reference", "Version", "Made by"])
            try:
                who = getpass.getuser()
            except Exception:
                who = ""
            w.writerow([dt.datetime.now().strftime("%d/%m/%Y %H:%M"), template, document, reference, version, who])
    except OSError:
        pass


def save_values(path: str, values: dict) -> None:
    """Keep a form's values for another day, as a .json that Load values reads."""
    with open(path, "w", encoding="utf-8") as fh:
        json.dump({k: v for k, v in values.items() if str(v).strip()}, fh, indent=2, ensure_ascii=False)
