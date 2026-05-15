#!/usr/bin/env python3
"""
sync_data.py — reads "Golf_Tee_Time_Tracker small.xlsx" and updates the
FALLBACK array in index.html with the latest course data.

Run from the repo root:
    python sync_data.py
"""
import re
import datetime
import openpyxl

EXCEL_FILE = "Golf_Tee_Time_Tracker small.xlsx"
HTML_FILE  = "index.html"
SHEET_NAME = "Tee Time Rules"


def fmt_time(val):
    """Convert an Excel cell value to an opensAt string like '8:00 AM'."""
    if val is None:
        return "Unknown"
    if isinstance(val, datetime.time):
        h, m = val.hour, val.minute
        ampm = "AM" if h < 12 else "PM"
        hour = h % 12 or 12          # 0 → 12, 13 → 1, etc.
        return f"{hour}:{m:02d} {ampm}"
    s = str(val).strip()
    s = re.sub(r'\s*\(.*?\)', '', s).strip()   # strip "(Midnight)", "(Noon)", etc.
    return s if s else "Unknown"


def parse_dist_mi(val):
    """Extract the numeric miles from a string like '~2 mi' → 2."""
    m = re.search(r'(\d+)', str(val or ''))
    return int(m.group(1)) if m else 0


def js_str(val):
    """Escape a value for embedding inside a JS double-quoted string."""
    return str(val or '').replace('\\', '\\\\').replace('"', '\\"')


def build_row(c):
    name    = js_str(c['course'])
    pad     = max(0, 33 - len(name))   # align columns for readability
    return (
        f'  {{course:"{name}",{" " * pad}'
        f'tier:"{js_str(c["tier"])}",holes:{c["holes"]},'
        f'distance:"{js_str(c["distance"])}",distMi:{c["distMi"]},'
        f'window:"{js_str(c["window"])}",opensAt:"{js_str(c["opensAt"])}",'
        f'dataStatus:"{js_str(c["dataStatus"])}",'
        f'phone:"{js_str(c["phone"])}",website:"{js_str(c["website"])}",'
        f'notes:"{js_str(c["notes"])}"}}'
    )


# ── Read Excel ───────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb[SHEET_NAME]

all_rows = list(ws.iter_rows(values_only=True))
# Normalise header names: lowercase, newlines → spaces
headers = [
    str(c).strip().lower().replace('\n', ' ') if c else ''
    for c in all_rows[0]
]

def col(frag):
    """Return the index of the first header containing frag."""
    for i, h in enumerate(headers):
        if frag in h:
            return i
    return None

ci_course     = col('course name')
ci_tier       = col('tier')
ci_holes      = col('holes')
ci_dist       = col('distance')
ci_window     = col('booking window')
ci_opens      = col('opens at')
ci_status     = col('data status')
ci_phone      = col('phone')
ci_website    = col('website')
ci_notes      = col('notes')
ci_cond_flag  = col('condition flag')
ci_cond       = col('conditions')

courses = []
for row in all_rows[1:]:
    if not row[ci_course]:
        continue   # skip blank rows

    # Booking window
    win_raw = row[ci_window]
    if isinstance(win_raw, (int, float)):
        window_str = f"{int(win_raw)} days"
    elif win_raw:
        window_str = str(win_raw).strip()
    else:
        window_str = "Unknown"

    # Notes — optionally append Conditions if the flag is set
    notes_val = str(row[ci_notes]).strip() if row[ci_notes] else ''
    if ci_cond_flag is not None and ci_cond is not None:
        cond_flag = str(row[ci_cond_flag]).strip() if row[ci_cond_flag] else ''
        cond_val  = str(row[ci_cond]).strip()      if row[ci_cond]      else ''
        if cond_flag and cond_val:
            notes_val = (notes_val + ' · ' + cond_val).strip(' · ')

    # Website — ensure https://
    website_raw = str(row[ci_website]).strip() if row[ci_website] else ''
    if website_raw and not website_raw.startswith('http'):
        website_raw = 'https://' + website_raw

    courses.append({
        'course':     str(row[ci_course]).strip(),
        'tier':       str(row[ci_tier]).strip()   if row[ci_tier]   else '',
        'holes':      int(row[ci_holes])           if row[ci_holes]  else 18,
        'distance':   str(row[ci_dist]).strip()   if row[ci_dist]   else '',
        'distMi':     parse_dist_mi(row[ci_dist]),
        'window':     window_str,
        'opensAt':    fmt_time(row[ci_opens]),
        'dataStatus': str(row[ci_status]).strip() if row[ci_status] else '',
        'phone':      str(row[ci_phone]).strip()  if row[ci_phone]  else '',
        'website':    website_raw,
        'notes':      notes_val,
    })

# ── Build replacement FALLBACK block ─────────────────────────────────────────
rows_js = [build_row(c) for c in courses]
new_fallback = (
    f'// FALLBACK DATA — all {len(courses)} courses from {EXCEL_FILE}\n'
    '// opensAt can be a string ("8:00 AM") or numeric fraction (e.g. 0.875 = 9PM)\n'
    '// ============================================================\n'
    'const FALLBACK = [\n'
    + ',\n'.join(rows_js)
    + '\n];'
)

# ── Patch index.html ──────────────────────────────────────────────────────────
with open(HTML_FILE, encoding='utf-8') as f:
    html = f.read()

pattern = (
    r'// FALLBACK DATA.*?'
    r'const FALLBACK = \[.*?\];'
)
new_html, count = re.subn(pattern, new_fallback, html, flags=re.DOTALL)

if count != 1:
    raise RuntimeError(
        f"Expected exactly 1 FALLBACK block to replace, found {count}. Aborting."
    )

with open(HTML_FILE, 'w', encoding='utf-8') as f:
    f.write(new_html)

print(f"✅  Synced {len(courses)} courses from '{EXCEL_FILE}' → '{HTML_FILE}'")
