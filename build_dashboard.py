#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dealer Performance Dashboard — monthly generator.

Reads the refreshed workbook (MACRO Monyhly macro.xlsm), rebuilds the whole
dealer database from the raw rows, and injects it into the dashboard HTML.

Design notes (verified against the existing dashboard to the unit):
  * contracts        = rows with application_status == 'Contract', bucketed by contract_date
  * App In (ai)      = every application row, bucketed by application_date (NOT deduped,
                        see dedupe rule below — App In always reflects every raw application)
  * Decline (dc)     = application_status == 'Decline/Cancel', bucketed by application_date
  * segment MO       = sub_product_type_name in {มือถือใหม่, มือถือมือสอง}
    segment EA       = everything else (incl. used appliances / furniture / e-motorcycle)
  * flat_rate        = raw value / 100  (raw 84 -> 0.84%)
  * gender           = Thai title prefix of customer_name (100% coverage: นาย/นาง/นางสาว)
  * dealer identity  = dealer_name_th
  * profiles are whole-period aggregates; the dashboard scales them by the
    selected-period contract ratio at render time.

Repeat-customer dedupe rule (added 2026-08, per Puis):
  Same customer_name can appear more than once (re-applying after a decline, or a
  genuine repeat purchase). If a customer_name has more than one row AND at least
  one of those rows is Decline/Cancel, and the group contains at least one actual
  Contract (i.e. a contract_no exists to anchor to), then ONLY the row with the
  latest contract_no counts toward Contract/Decline totals (and its profile
  dimensions) — every other row in that name-group (earlier contracts, all
  declines) is dropped from Contract/Decline counting. Groups with no Contract at
  all (only repeated declines, never approved) are left untouched — there's no
  contract_no to anchor "latest" to.
  App In (ai) is explicitly NOT affected — every application row is still counted,
  exactly as before. This is a customer-identity dedupe (by full name, across all
  dealers — there's no unique customer ID in the source data), not a per-dealer one.

Safety: refuses to publish if the workbook was not refreshed, or if the rebuilt
numbers drift from the previous build beyond tolerance.
"""

import io
import json
import os
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from datetime import date

# ── paths ──────────────────────────────────────────────────────────────────
BASE = os.path.dirname(os.path.abspath(__file__))
XLSM = os.path.join(BASE, "MACRO Monyhly macro.xlsm")
HTML = os.path.join(BASE, "Dealer Performance Dashboard.html")
INDEX = os.path.join(BASE, "index.html")
TEMPLATE = os.path.join(BASE, "dashboard_template.html")
STATE = os.path.join(BASE, ".dashboard_build_state.json")

# ── domain constants ───────────────────────────────────────────────────────
MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
SUB_MO = {"มือถือใหม่", "มือถือมือสอง"}          # everything else -> EA
MALE_PREFIX = {"นาย"}
FEMALE_PREFIX = {"นาง", "นางสาว"}

# list caps observed in the original dashboard
CAPS = {"sw": None, "md": 15, "pv": 15, "oc": 12, "brand": 40}

# columns we actually read (letter -> short name)
NEED = {
    "A": "appno", "B": "app_date", "C": "conno", "D": "cust",
    "E": "con_date", "H": "status",
    "AP": "dlr", "AS": "sw", "AW": "brand", "AZ": "md",
    "BG": "cd", "BJ": "car", "BP": "fin", "BV": "tm",
    "CA": "flat", "CD": "nc", "CF": "age", "CK": "pv",
    "CS": "oc", "DX": "subpt", "EB": "ct",
}

DICT_FIELDS = ["ct", "nc", "tk", "gn", "ag", "cd", "tm", "fr"]
LIST_FIELDS = ["sw", "md", "pv", "oc"]
SEGS = ("", "MO", "EA")


def log(msg):
    sys.stdout.write(msg + "\n")
    sys.stdout.flush()


# ── bucket helpers ─────────────────────────────────────────────────────────
def bucket_flat(x):
    if x < 0.5:
        return "<0.5%"
    if x < 0.75:
        return "0.5–0.74%"
    if x < 1.0:
        return "0.75–0.99%"
    if x < 1.5:
        return "1.0–1.49%"
    return "≥1.5%"          # overflow guard — never silently drop a contract


def bucket_ticket(v):
    if v < 10000:
        return "<10K"
    if v < 20000:
        return "10-20K"
    if v < 30000:
        return "20-30K"
    if v < 50000:
        return "30-50K"
    return "≥50K"           # overflow guard


def bucket_age(a):
    if a < 25:
        return "<25"
    if a >= 60:
        return "60+"
    lo = (a // 5) * 5
    return "%d-%d" % (lo, lo + 4)


def month_key(ds):
    """'DD/MM/YYYY' -> ("Sep'22", (year, month))"""
    if not ds or len(ds) < 10:
        return None, None
    try:
        dd, mm, yy = ds.split("/")
        m, y = int(mm), int(yy)
    except ValueError:
        return None, None
    if not (1 <= m <= 12):
        return None, None
    return "%s'%s" % (MON[m - 1], yy[2:]), (y, m)


# ── workbook access ────────────────────────────────────────────────────────
def read_params(z):
    """Pull B1..B3 (flag / from / to) out of the first rows."""
    f = z.open("xl/sharedStrings.xml")
    head = f.read(2_000_000).decode("utf-8", "ignore")
    ss = re.findall(r"<si>(?:<t[^>]*>|.*?<t[^>]*>)(.*?)</t>", head, re.S)[:400]
    g = z.open("xl/worksheets/sheet1.xml")
    chunk = g.read(300_000).decode("utf-8", "ignore")
    out = {}
    for rn, body in re.findall(r'<row[^>]*r="(\d)"[^>]*>(.*?)</row>', chunk, re.S)[:3]:
        for full, col in re.findall(
                r'(<c[^>]*r="([A-Z]+)\d+"[^>]*(?:/>|>.*?</c>))', body, re.S):
            if col != "B":
                continue
            t = re.search(r't="([^"]*)"', full)
            v = re.search(r"<v>(.*?)</v>", full)
            val = v.group(1) if v else ""
            if t and t.group(1) == "s" and val.isdigit():
                i = int(val)
                val = ss[i] if i < len(ss) else ""
            out[{"1": "flag", "2": "from", "3": "to"}[rn]] = val
    return out


def load_shared_strings(z):
    ss = []
    dec = io.TextIOWrapper(z.open("xl/sharedStrings.xml"),
                           encoding="utf-8", errors="ignore")
    tre = re.compile(r"<t[^>]*>(.*?)</t>", re.S)
    carry = ""
    while True:
        chunk = dec.read(8_000_000)
        if not chunk:
            break
        carry += chunk
        parts = carry.split("</si>")
        carry = parts.pop()
        for p in parts:
            m = tre.findall(p)
            ss.append("".join(m) if m else "")
    return ss


def iter_rows(z, ss):
    """Stream the sheet, yielding dicts of just the columns in NEED."""
    cre = re.compile(r'<c r="([A-Z]+)\d+"([^>]*)(?:/>|>(.*?)</c>)', re.S)
    vre = re.compile(r"<v>(.*?)</v>", re.S)
    dec = io.TextIOWrapper(z.open("xl/worksheets/sheet1.xml"),
                           encoding="utf-8", errors="ignore")
    carry = ""
    while True:
        chunk = dec.read(8_000_000)
        if not chunk:
            break
        carry += chunk
        parts = carry.split("</row>")
        carry = parts.pop()
        for body in parts:
            if '<c r="A' not in body:
                continue
            rec = {}
            for col, attrs, inner in cre.findall(body):
                key = NEED.get(col)
                if not key or inner is None:
                    continue
                m = vre.search(inner or "")
                if not m:
                    continue
                val = m.group(1)
                if 't="s"' in attrs:
                    i = int(val)
                    val = ss[i] if i < len(ss) else ""
                rec[key] = val
            if rec:
                yield rec


# ── per-dealer accumulator ─────────────────────────────────────────────────
def new_dealer():
    d = {"tot": {s: 0 for s in SEGS}}
    d["mo"] = {s: defaultdict(lambda: [0, 0.0, 0.0]) for s in SEGS}
    d["ai"] = {s: defaultdict(int) for s in SEGS}
    d["dc"] = {s: defaultdict(int) for s in SEGS}
    for f in DICT_FIELDS + LIST_FIELDS + ["brand"]:
        d[f] = {s: defaultdict(int) for s in SEGS}
    return d


def compute_dedupe_exclusions(rows):
    """Repeat-customer dedupe (see module docstring).

    For each customer_name appearing on more than one relevant row, where the
    group contains at least one Decline/Cancel AND at least one actual Contract
    (i.e. a contract_no to anchor to), every row except the one with the
    highest contract_no is excluded from Contract/Decline counting. Groups
    with no Contract at all are left alone. Returns a set of appno's to skip
    when tallying Contract/Decline (App In stays unaffected — caller must not
    apply this set there).
    """
    byname = defaultdict(list)
    for rec in rows:
        name = rec.get("cust")
        if name:
            byname[name].append(rec)

    def conno_key(r):
        c = r.get("conno") or ""
        try:
            return int(c)
        except ValueError:
            return -1

    excluded = set()
    for name, grp in byname.items():
        if len(grp) < 2:
            continue
        if not any(r.get("status") == "Decline/Cancel" for r in grp):
            continue
        contract_rows = [r for r in grp
                         if r.get("status") == "Contract" and r.get("conno")]
        if not contract_rows:
            continue                      # nothing to anchor "latest" to
        latest = max(contract_rows, key=conno_key)
        for r in grp:
            if r is not latest:
                excluded.add(r.get("appno"))
    return excluded


def build(z, ss, cutoff):
    """cutoff = (year, month) of the first INCOMPLETE month; drop it and later."""
    D = defaultdict(new_dealer)
    stats = defaultdict(int)
    subtypes = defaultdict(int)
    unknown_prefix = defaultdict(int)

    rows = [rec for rec in iter_rows(z, ss)
            if rec.get("status") in ("Contract", "Decline/Cancel", "Approve")
            and rec.get("dlr")]
    dedupe_excluded = compute_dedupe_exclusions(rows)
    stats["dedupe_excluded_rows"] = len(dedupe_excluded)

    for rec in rows:
        status = rec.get("status")
        name = rec.get("dlr")
        sub = rec.get("subpt", "")
        seg = "MO" if sub in SUB_MO else "EA"
        subtypes[sub] += 1
        d = D[name]
        deduped_out = rec.get("appno") in dedupe_excluded

        # ---- App In, keyed by application month — NEVER deduped ----
        amk, amt = month_key(rec.get("app_date", ""))
        if amk and amt < cutoff:
            d["ai"][""][amk] += 1
            d["ai"][seg][amk] += 1
            # ---- Decline, keyed by application month — deduped ----
            if status == "Decline/Cancel" and not deduped_out:
                d["dc"][""][amk] += 1
                d["dc"][seg][amk] += 1
                stats["decline"] += 1

        if status != "Contract":
            continue
        if deduped_out:
            stats["contract_deduped_out"] += 1
            continue

        # ---- contracts, keyed by contract month ----
        cmk, cmt = month_key(rec.get("con_date", ""))
        if not cmk or cmt >= cutoff:
            continue
        stats["contract"] += 1

        def num(k):
            try:
                return float(rec.get(k) or 0)
            except ValueError:
                return 0.0

        car, fin = num("car"), num("fin")
        for s in ("", seg):
            d["tot"][s] += 1
            row = d["mo"][s][cmk]
            row[0] += 1
            row[1] += car
            row[2] += fin

        # ---- profile dimensions ----
        vals = {
            "sw": rec.get("sw"), "md": rec.get("md"), "pv": rec.get("pv"),
            "oc": rec.get("oc"), "brand": rec.get("brand"),
            "ct": rec.get("ct"), "nc": rec.get("nc"), "cd": rec.get("cd"),
        }
        prefix = (rec.get("cust") or "").split(" ")[0]
        if prefix in MALE_PREFIX:
            vals["gn"] = "Male"
        elif prefix in FEMALE_PREFIX:
            vals["gn"] = "Female"
        else:
            unknown_prefix[prefix] += 1
            vals["gn"] = None

        vals["tk"] = bucket_ticket(fin)
        try:
            vals["fr"] = bucket_flat(num("flat") / 100.0)
        except Exception:
            vals["fr"] = None
        try:
            vals["ag"] = bucket_age(int(num("age")))
        except Exception:
            vals["ag"] = None
        t = rec.get("tm")
        try:
            vals["tm"] = str(int(float(t))) if t else None
        except ValueError:
            vals["tm"] = None

        for field, v in vals.items():
            if not v:
                continue
            for s in ("", seg):
                d[field][s][v] += 1

    return D, stats, subtypes, unknown_prefix


def to_db(D):
    """Materialise the accumulator into the dashboard's exact DB schema."""
    out = []
    for name, d in D.items():
        if d["tot"][""] <= 0:
            continue                      # dealers with no contracts are omitted
        rec = {"n": name,
               "tot": d["tot"][""], "totMO": d["tot"]["MO"], "totEA": d["tot"]["EA"]}

        def mname(base, s):
            return base + s

        for s in SEGS:
            rec[mname("mo", s)] = {
                k: [v[0], int(round(v[1])), int(round(v[2]))]
                for k, v in sorted(d["mo"][s].items(), key=lambda kv: mo_sort(kv[0]))
            }
        for base in ("ai", "dc"):
            for s in SEGS:
                rec[mname(base, s)] = {
                    k: v for k, v in sorted(d[base][s].items(),
                                            key=lambda kv: mo_sort(kv[0]))
                }
        for base in LIST_FIELDS:
            cap = CAPS.get(base)
            for s in SEGS:
                items = sorted(d[base][s].items(), key=lambda kv: -kv[1])
                if cap:
                    items = items[:cap]
                rec[mname(base, s)] = [[k, v] for k, v in items]
        for base in DICT_FIELDS:
            for s in SEGS:
                rec[mname(base, s)] = dict(d[base][s])
        # brand comes last, and in MO/EA/all order
        cap = CAPS["brand"]
        for s in ("MO", "EA", ""):
            items = sorted(d["brand"][s].items(), key=lambda kv: -kv[1])[:cap]
            rec["brand" + s] = [[k, v] for k, v in items]
        out.append(rec)
    out.sort(key=lambda r: -r["tot"])
    return out


def mo_sort(m):
    mo, yr = m.split("'")
    return (int(yr), MON.index(mo))


def month_range(db):
    seen = set()
    for r in db:
        seen.update(r["mo"].keys())
    if not seen:
        return [], {}, []
    ordered = sorted(seen, key=mo_sort)
    (y0, m0), (y1, m1) = mo_sort(ordered[0]), mo_sort(ordered[-1])
    full = []
    y, m = y0, m0 + 1
    while (y, m) <= (y1, m1 + 1):
        full.append("%s'%02d" % (MON[m - 1], y))
        m += 1
        if m > 12:
            m = 1
            y += 1
    years = sorted({m.split("'")[1] for m in full}, key=int)
    yr_mo = {yy: [m for m in full if m.split("'")[1] == yy] for yy in years}
    return full, yr_mo, years


# ── html injection ─────────────────────────────────────────────────────────
def render(db, mo_all, yr_mo, years):
    """Prefer the placeholder template; fall back to re-templating the live file.

    Using the template keeps each build independent — a bad build can never
    poison the next one, which self-templating would allow.
    """
    j = lambda o: json.dumps(o, ensure_ascii=False, separators=(", ", ": "))
    if os.path.exists(TEMPLATE):
        html = open(TEMPLATE, encoding="utf-8").read()
        for token, payload in (("__DB__", j(db)),
                               ("__MO_ALL__", j(mo_all)),
                               ("__YR_MO__", j(yr_mo)),
                               ("__NDEALERS__", "{:,}".format(len(db)))):
            if token not in html:
                raise RuntimeError("template is missing %s" % token)
            html = html.replace(token, payload, 1)
        # YEARS has no placeholder in the template, but must stay in sync
        # or a new calendar year would be missing from the year filter.
        html = re.sub(r"var YEARS=\[.*?\];", "var YEARS=%s;" % j(years),
                      html, count=1)
        return html, "dashboard_template.html"
    return inject(open(HTML, encoding="utf-8").read(),
                  db, mo_all, yr_mo, years), "live HTML (no template found)"


def inject(html, db, mo_all, yr_mo, years):
    def repl_var(text, var, payload):
        pat = re.compile(r"^var %s=.*?;$" % re.escape(var), re.M)
        new = "var %s=%s;" % (var, payload)
        text, n = pat.subn(lambda _: new, text, count=1)
        if n != 1:
            raise RuntimeError("could not replace var %s (matched %d)" % (var, n))
        return text

    j = lambda o: json.dumps(o, ensure_ascii=False, separators=(", ", ": "))
    html = repl_var(html, "MO_ALL", j(mo_all))
    html = repl_var(html, "YEARS", j(years))
    html = repl_var(html, "YR_MO", j(yr_mo))

    # DB is ~9MB on one line — bracket-match instead of regex (O(n), no backtracking)
    start = html.index("DB=[") + 3
    depth, k = 0, start
    while k < len(html):
        ch = html[k]
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                k += 1
                break
        k += 1
    else:
        raise RuntimeError("could not locate end of DB= payload")
    html = html[:start] + j(db) + html[k:]

    # header dealer count must not stay hardcoded
    html = re.sub(r"(<b>Dealers</b>)[^<]*(</div>)",
                  r"\g<1>{:,} ราย\g<2>".format(len(db)), html, count=1)
    return html


# ── main ───────────────────────────────────────────────────────────────────
def main(force=False):
    if not os.path.exists(XLSM):
        log("ERROR: workbook not found: %s" % XLSM)
        return 2

    st = os.stat(XLSM)
    z = zipfile.ZipFile(XLSM)
    params = read_params(z)
    fingerprint = {"mtime": int(st.st_mtime), "size": st.st_size,
                   "to": params.get("to"), "from": params.get("from")}
    log("workbook: to=%s  from=%s  size=%.1fMB  mtime=%d"
        % (params.get("to"), params.get("from"), st.st_size / 1e6, st.st_mtime))

    prev = {}
    if os.path.exists(STATE):
        try:
            prev = json.load(open(STATE, encoding="utf-8"))
        except Exception:
            prev = {}

    # ---- GATE 1: staleness ----
    if not force and prev.get("fingerprint") == fingerprint:
        log("STOP: workbook unchanged since the last build "
            "(same mtime/size/date range). Refresh it in Excel first.")
        return 3

    today = date.today()
    cutoff = (today.year, today.month)     # current month is incomplete -> drop
    log("including complete months up to %s'%02d"
        % (MON[(cutoff[1] - 2) % 12], (cutoff[0] - 1) if cutoff[1] == 1 else cutoff[0] % 100))

    log("loading shared strings...")
    ss = load_shared_strings(z)
    log("  %d strings" % len(ss))
    log("parsing rows...")
    D, stats, subtypes, unknown = build(z, ss, cutoff)
    db = to_db(D)
    mo_all, yr_mo, years = month_range(db)

    tot = sum(r["tot"] for r in db)
    tmo = sum(r["totMO"] for r in db)
    tea = sum(r["totEA"] for r in db)
    retail = sum(v[1] for r in db for v in r["mo"].values())
    finance = sum(v[2] for r in db for v in r["mo"].values())
    log("")
    log("dealers   : %d" % len(db))
    log("contracts : %d   (MO %d + EA %d = %d)" % (tot, tmo, tea, tmo + tea))
    log("retail    : %d" % retail)
    log("finance   : %d" % finance)
    log("months    : %d  (%s .. %s)" % (len(mo_all), mo_all[0], mo_all[-1]))
    log("sub_product_type seen: %s" % dict(sorted(subtypes.items(), key=lambda kv: -kv[1])))
    if unknown:
        log("WARNING unmapped name prefixes: %s" % dict(unknown))
    if stats.get("dedupe_excluded_rows"):
        log("repeat-customer dedupe: %d rows excluded from Contract/Decline "
            "(%d superseded contracts, App In unaffected)"
            % (stats["dedupe_excluded_rows"], stats.get("contract_deduped_out", 0)))

    # ---- GATE 2: internal consistency ----
    problems = []
    if tmo + tea != tot:
        problems.append("MO+EA (%d) != tot (%d)" % (tmo + tea, tot))
    if len(db) < 2000:
        problems.append("dealer count implausibly low: %d" % len(db))
    if tot < 100000:
        problems.append("contract count implausibly low: %d" % tot)
    for r in db[:50]:
        if sum(v[0] for v in r["mo"].values()) != r["tot"]:
            problems.append("dealer %r: mo sum != tot" % r["n"])
            break

    # ---- GATE 3: drift vs previous build ----
    # DOWN_TOLERANCE allows small decreases (source corrections, or the
    # repeat-customer dedupe rule) without tripping the gate; a real data
    # problem (wrong file, broken query) will still show up as a much bigger
    # drop than this.
    DOWN_TOLERANCE = 0.03
    if prev.get("totals"):
        p = prev["totals"]
        for label, old, new in (("contracts", p.get("contracts"), tot),
                                ("dealers", p.get("dealers"), len(db))):
            if not old:
                continue
            if new < old:
                drop = (old - new) / old
                if drop > DOWN_TOLERANCE:
                    problems.append("%s went DOWN >%.0f%%: %d -> %d"
                                    % (label, DOWN_TOLERANCE * 100, old, new))
                else:
                    log("note: %s decreased %d -> %d (%.1f%%, within tolerance)"
                        % (label, old, new, drop * 100))
            elif old and (new - old) / old > 0.25:
                problems.append("%s jumped >25%%: %d -> %d" % (label, old, new))

    if problems:
        log("")
        log("VALIDATION FAILED — nothing was written:")
        for p in problems:
            log("  * " + p)
        return 4

    # ---- write ----
    out, src = render(db, mo_all, yr_mo, years)
    log("rendered from: %s" % src)
    shutil.copyfile(HTML, HTML + ".bak")
    open(HTML, "w", encoding="utf-8").write(out)
    open(INDEX, "w", encoding="utf-8").write(out)
    json.dump({"fingerprint": fingerprint,
               "totals": {"contracts": tot, "dealers": len(db),
                          "retail": retail, "finance": finance,
                          "last_month": mo_all[-1]},
               "built": today.isoformat()},
              open(STATE, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    log("")
    log("OK wrote both HTML files (through %s). Previous kept as .bak" % mo_all[-1])
    return 0


if __name__ == "__main__":
    sys.exit(main(force="--force" in sys.argv))
