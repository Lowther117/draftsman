"""One document per row of a spreadsheet."""
from __future__ import annotations

import os
import re

from . import merge
from .library import file_name, match_keys, unique_path


def row_choices(row: dict, scan, library, default: dict) -> dict:
    """Which clauses a row asks for. A column headed "clauses:Group" (or just
    the group's name) lists them, separated by semicolons; "all" takes every
    clause in the folder. Without such a column the ticks on the form apply."""
    def norm(s):
        return re.sub(r"[\s_\-]+", " ", str(s)).strip().lower()
    by = {norm(k): v for k, v in row.items()}
    out = dict(default)
    for group in scan.groups:
        raw = by.get(norm("clauses:" + group), by.get(norm(group)))
        if raw is None:
            continue
        available = [n for n, _p in library.clauses_in(group)] if library else []
        if str(raw).strip().lower() in ("all", "*"):
            out[group] = available
            continue
        wanted = [norm(x) for x in str(raw).split(";") if x.strip()]
        # "1 Information about other people" can be asked for with or without its number
        out[group] = [n for n in available if norm(n) in wanted or norm(re.sub(r"^\d+[\s.)-]*", "", n)) in wanted]
    return out


def run(template, rows, out_dir, pattern, scan, library=None, base=None, choices=None, control=None,
        control_table=False, progress=None) -> list:
    """Fill the template once per row. `base` holds the values typed on the
    form and the standing values; a row's own columns win over them."""
    os.makedirs(out_dir, exist_ok=True)
    wanted = scan.names() + list(scan.switches) + list(merge.CONTROL)
    stem = os.path.splitext(os.path.basename(template))[0]
    reports = []
    for n, row in enumerate(rows, 1):
        got = match_keys(row, wanted)
        values = {**(base or {}), **{k: v for k, v in got.items() if k not in merge.CONTROL and str(v).strip()}}
        ctl = {**(control or {}), **{k: v for k, v in got.items() if k in merge.CONTROL and str(v).strip()}}
        name = file_name(pattern, {**ctl, **values}, f"{stem} {n}")
        path = unique_path(out_dir, name)
        if progress:
            progress(f"Row {n} of {len(rows)}: {os.path.basename(path)}")
        report = merge.render(template, path, values, library, row_choices(row, scan, library, choices or {}),
                              ctl, control_table)
        report["row"] = n
        reports.append(report)
    return reports
