"""
CSUN Catcher Report
-------------------
Builds a one-page report per catcher from TrackMan game files plus our manual
tracking (blocks, steal attempts, passed balls):

  * Framing:  called strikes vs. an expected-strike model built from our own
              historical pitch locations
  * Blocking: success rate by side and pitch group, plus run-value credit
  * Throwing: pop times, throw velocity and run-value credit for caught stealings

Run value comes from a published D1 run-expectancy table. Every report also
writes a review log (markdown) listing each credited play and anything that
looks off in the data, so the numbers can be checked against video.

Usage:
    python catcher_report.py <game file> [--manual tracking.xlsx]

A bare file name is looked up in ../Data/single_game_csv. The PDF and review
log go to ../Reports. Requires: reportlab, openpyxl.
"""

import argparse
import base64
import csv
import io
import json
import math
import zlib
from collections import Counter
from datetime import date, datetime
from pathlib import Path


# ============================================================
# Settings
# ============================================================

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent

# CSUN Baseball logo in the top-right corner of each page (points; page is 612 x 792)
LOGO_WIDTH = 90
LOGO_RIGHT = 576          # lines up with the right margin
LOGO_TOP = 766            # just below the red band
INPUT_EXTENSIONS = {".csv", ".tsv", ".xlsx", ".xlsm"}

PITCH_GROUPS = ["Fastball", "RH breaking", "RH offspeed", "LH breaking", "LH offspeed"]

# Columns that come from the manual tracking sheet
MANUAL_FIELDS = ["Catcher", "Pitcher", "Hitter", "R1B", "R2B", "R3B", "StealAtt",
                 "DubSteal", "StealBase", "AdvThrw", "Safe/Out", "BlockAtt",
                 "BlockSide", "Success", "PassedB"]

# Fields a corrections file is allowed to change
OVERRIDABLE_FIELDS = set(MANUAL_FIELDS) | {
    "Outs", "ManualCatcher", "ManualPitcher", "PitchCall", "Balls", "Strikes",
    "ThrowBase", "OutBase", "ThrowVerified", "PopTime", "ThrowSpeed"}

# Saved correction files for specific scrimmages (applied unless --no-known-corrections)
KNOWN_CORRECTION_FILES = ["scrimmage_corrections.json", "scrimmage_nov4_corrections.json"]

# Strike zone in feet: 17 in wide, 1.5 to 3.5 ft tall
ZONE_HALF_WIDTH = 17 / 24
ZONE_BOTTOM, ZONE_TOP = 1.5, 3.5
SHADOW_BAND = 0.2        # the shadow zone straddles the zone edge by this much on each side


# ============================================================
# Small helpers
# ============================================================

def number(value):
    """Float, or None if blank / not a number / infinite."""
    try:
        n = float(value)
        return n if math.isfinite(n) else None
    except (ValueError, TypeError):
        return None


def yes(value):
    return str(value).strip().lower() in {"yes", "y", "true", "1", "x"}


def cell_text(value):
    """Turn an Excel cell into the same kind of text a CSV would give."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return "yes" if value else "no"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def add_issue(issues, row, message):
    issues.append({"game": row.get("GameUID", ""), "pitch": row.get("PitchNo", ""),
                   "catcher": row.get("Catcher", "") or "Unknown", "message": message})


def rate(successes, attempts):
    return "--" if attempts == 0 else f"{successes / attempts:.0%}"


# ============================================================
# Reading TrackMan and manual-tracking files
# ============================================================

def read_table(path, required, sheet=None):
    """Find the header row (the first row containing every `required` column) in a
    CSV/TSV/XLSX file and return (header, rows). Rows are padded to the header width."""
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in {".csv", ".tsv"}:
        with path.open(encoding="utf-8-sig", newline="") as f:
            delimiter = "\t" if suffix == ".tsv" else ","
            sheets = [(path.name, list(csv.reader(f, delimiter=delimiter)))]
    elif suffix in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            if sheet and sheet not in book.sheetnames:
                raise ValueError(f"{path}: sheet {sheet!r} not found")
            sheets = []
            for name in ([sheet] if sheet else book.sheetnames):
                values = [[cell_text(v) for v in row] for row in book[name].iter_rows(values_only=True)]
                sheets.append((name, values))
        finally:
            book.close()
    else:
        raise ValueError(f"{path}: use .csv, .tsv, .xlsx or .xlsm; save older .xls files as .xlsx")

    matches = []
    for name, values in sheets:
        start = next((i for i, row in enumerate(values) if required <= {v.strip() for v in row}), None)
        if start is None:
            continue
        header = [v.strip() for v in values[start]]
        while header and not header[-1]:
            header.pop()
        body = []
        for row in values[start + 1:]:
            if not any(row):
                continue
            if len(row) > len(header) and any(row[len(header):]):
                raise ValueError(f"{path} ({name}): data extends beyond the header")
            body.append((row + [""] * len(header))[:len(header)])
        matches.append((header, body))

    if len(matches) != 1:
        raise ValueError(f"{path}: expected one data sheet with {sorted(required)}; found {len(matches)}. "
                         "Supply --sheet or --manual-sheet if needed.")
    return matches[0]


def parse_date(text, path):
    """Accept the date formats TrackMan and Excel exports use; return YYYY-MM-DD."""
    try:
        return date.fromisoformat(text).isoformat()
    except ValueError:
        pass
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y%m%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
                "%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%m/%d/%y %H:%M"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"{path}: invalid Date {text!r}; use YYYY-MM-DD")


def read_trackman(path, skipped=None, sheet=None):
    """Read one TrackMan export into a list of row dicts."""
    columns, body = read_table(path, {"PitchNo", "Pitcher", "PitchCall"}, sheet)

    # Some exports have a second Catcher/Pitcher column from manual entry;
    # rename those to ManualCatcher/ManualPitcher, and number any other repeats.
    seen = Counter()
    header = []
    for field in columns:
        field = field.strip()
        seen[field] += 1
        if field in {"Catcher", "Pitcher"} and seen[field] == 2:
            header.append("Manual" + field)
        elif seen[field] == 1:
            header.append(field)
        else:
            header.append(f"{field}.{seen[field]}")

    required = {"Date", "GameUID", "PitchUID", "Catcher", "Outs", "PitchCall",
                "PlateLocSide", "PlateLocHeight", "TaggedPitchType", "PitcherThrows"}
    missing = required - set(header)
    if missing:
        raise ValueError(f"{path}: missing required columns: {sorted(missing)}")

    rows = []
    for line, values in enumerate(body, 2):
        if not any(s.strip() for s in values):
            continue
        row = {k: v.strip() for k, v in zip(header, values) if k}

        pitch = number(row.get("PitchNo"))
        if pitch is None or pitch < 1 or not pitch.is_integer():
            raise ValueError(f"{path}, line {line}: invalid PitchNo")
        row["PitchNo"] = int(pitch)
        if not row["GameUID"] or not row["PitchUID"]:
            raise ValueError(f"{path}, line {line}: missing game/pitch identifier")

        # TrackMan sometimes logs empty placeholder pitches; drop them but note it
        is_empty = (not row["Date"] and not row.get("Time") and not row["PlateLocHeight"]
                    and not row["PlateLocSide"] and row["PitchCall"] in {"", "Undefined"})
        if is_empty:
            if skipped is not None:
                add_issue(skipped, row, "Excluded empty TrackMan record (no date, time, location, or call)")
            continue

        row["Date"] = parse_date(row["Date"], path)
        row["_source"] = str(Path(path).resolve())
        rows.append(row)
    return rows


def load_rows(paths, skipped=None, sheet=None):
    """Read every game file and refuse duplicate games or pitches."""
    rows = [row for path in paths for row in read_trackman(path, skipped, sheet)]
    for key_fields in [("GameUID", "PitchNo"), ("PitchUID",)]:
        keys = [tuple(r[k] for k in key_fields) for r in rows]
        if len(keys) != len(set(keys)):
            raise ValueError(f"Duplicate {key_fields}; overlapping exports must be resolved before reporting")
    return sorted(rows, key=lambda r: (r["Date"], r["GameUID"], r["PitchNo"]))


def expand_files(paths):
    """Turn a mix of files and folders into a list of game files."""
    result = []
    for value in paths:
        path = Path(value).expanduser().resolve()
        if path.is_dir():
            result += sorted(p for p in path.iterdir()
                             if p.suffix.lower() in INPUT_EXTENSIONS and not p.name.startswith("~$"))
        elif path.is_file():
            result.append(path)
        else:
            raise ValueError(f"Input not found: {path}")
    if not result:
        raise ValueError("No input files found")
    if len(result) != len(set(result)):
        raise ValueError("The same file was selected more than once")
    return result


def resolve_input(name, data_dir):
    """A path as given, otherwise a file name inside the Data folder."""
    path = Path(name).expanduser()
    if path.exists():
        return path
    if data_dir and (Path(data_dir) / name).exists():
        return Path(data_dir) / name
    raise ValueError(f"File not found: {name}" + (f" (also looked in {data_dir})" if data_dir else ""))


# ============================================================
# Manual tracking and corrections
# ============================================================

def join_manual(rows, path, game, pitch_map=None, measured=False, sheet=None):
    """Copy manual catcher events onto the matching TrackMan rows by pitch number.

    measured=True means the manual sheet numbers only the pitches TrackMan actually
    measured (so it skips empty records); otherwise PitchNo is matched directly,
    after any saved pitch_map renumbering."""
    header, body = read_table(path, {"PitchNo", "Catcher", "BlockAtt", "StealAtt"}, sheet)
    if len(header) != len(set(header)):
        raise ValueError(f"{path}: duplicate manual headers")

    game_rows = sorted([r for r in rows if r["GameUID"] == game], key=lambda r: r["PitchNo"])
    by_pitch = {r["PitchNo"]: r for r in game_rows}
    matched = set()
    changes = []

    for values in body:
        event = dict(zip(header, values))
        if not event.get("PitchNo"):
            if any(event.get(k) for k in ("BlockAtt", "StealAtt", "PassedB")):
                raise ValueError(f"{path}: manual event without PitchNo")
            continue

        value = number(event["PitchNo"])
        if value is None or not value.is_integer() or value < 1:
            raise ValueError(f'{path}: invalid manual PitchNo {event["PitchNo"]!r}')
        manual_pitch = int(value)

        if measured:
            if manual_pitch > len(game_rows):
                raise ValueError(f"{path}: measured pitch ordinal out of range")
            target = game_rows[manual_pitch - 1]["PitchNo"]
        else:
            target = int((pitch_map or {}).get(str(manual_pitch), manual_pitch))

        if target not in by_pitch:
            raise ValueError(f"{path}: manual pitch {manual_pitch} matches no measured pitch; check --manual-index")
        if target in matched:
            raise ValueError(f"{path}: duplicate manual event for pitch {target}")
        matched.add(target)

        row = by_pitch[target]
        for field in MANUAL_FIELDS:
            incoming = event.get(field, "")
            if incoming == "":
                continue
            dest = "Manual" + field if field in {"Catcher", "Pitcher"} else field
            previous = row.get(dest, "")
            if previous != incoming:
                changes.append({"game": game, "pitch": target, "field": dest, "old": previous, "new": incoming})
                row[dest] = incoming
    return changes, len(matched)


def normalize_manual(row):
    """Standardize the many ways the sheet gets filled in (Y/N, Mid/M/L/R, Safe/Out,
    2nd/3rd, ...). Blank stays blank. If block side wasn't logged, fall back to the
    pitch location (catcher's view: negative PlateLocSide = right), which is how the
    existing logs were scored."""
    for field in ("BlockAtt", "StealAtt", "DubSteal", "AdvThrw", "PassedB", "Success"):
        v = str(row.get(field, "")).strip().lower()
        if v:
            if v in {"yes", "y", "true", "1", "x"}:
                row[field] = "yes"
            elif v in {"no", "n", "false", "0"}:
                row[field] = "no"
            else:
                row[field] = v

    v = str(row.get("Safe/Out", "")).strip().lower()
    if v:
        row["Safe/Out"] = "safe" if v.startswith("s") else "out" if v.startswith("o") else v

    v = str(row.get("StealBase", "")).strip().lower().replace("nd", "").replace("rd", "")
    if v:
        row["StealBase"] = v

    v = str(row.get("BlockSide", "")).strip().lower()
    if v:
        row["BlockSide"] = {"l": "left", "r": "right", "m": "middle"}.get(v[0], v)
    elif yes(row.get("BlockAtt")) and number(row.get("PlateLocSide")) is not None:
        x = number(row["PlateLocSide"])
        row["BlockSide"] = "middle" if abs(x) <= .3 else "left" if x > 0 else "right"


def apply_overrides(rows, path=None, specs=None):
    """Apply pitch-level corrections (from a JSON file or a list of dicts).
    Returns a list of what changed for the review log."""
    if path is None and specs is None:
        return []
    if specs is None:
        specs = json.loads(Path(path).read_text(encoding="utf-8"))["overrides"]

    by_key = {(r["GameUID"], r["PitchNo"]): r for r in rows}
    used = set()
    changes = []
    for spec in specs:
        key = (spec["GameUID"], int(spec["PitchNo"]))
        if key in used:
            raise ValueError(f"Duplicate override for {key}")
        used.add(key)
        if key not in by_key:
            raise ValueError(f"Override targets absent game/pitch: {key}")

        for field, value in spec.items():
            if field in {"GameUID", "PitchNo"}:
                continue
            if field not in OVERRIDABLE_FIELDS:
                raise ValueError(f"Override field is not supported: {field}")
            if field == "PitchCall" and value not in {"BallCalled", "StrikeCalled"}:
                raise ValueError("PitchCall correction must be BallCalled or StrikeCalled")
            if field in {"Balls", "Strikes"} and number(value) not in range(4 if field == "Balls" else 3):
                raise ValueError(f"Invalid corrected {field} count")
            changes.append({"game": key[0], "pitch": key[1], "field": field,
                            "old": by_key[key].get(field, ""), "new": str(value)})
            by_key[key][field] = str(value)
    return changes


# ============================================================
# Expected-strike model and run-expectancy table
# ============================================================

def embedded_strike_model():
    model = dict(STRIKE_MODEL_INFO)
    model["probabilities"] = json.loads(zlib.decompress(base64.b64decode(STRIKE_MODEL_GRID)))
    return model


def load_strike_model(source):
    """Load and sanity-check the expected-strike grid (a dict or a JSON file path)."""
    model = source if isinstance(source, dict) else json.loads(Path(source).read_text(encoding="utf-8"))
    grid = model["probabilities"]
    if len(grid) < 2 or len(grid[0]) < 2 or any(len(row) != len(grid[0]) for row in grid):
        raise ValueError("Invalid expected-strike probability grid")
    if not math.isfinite(model["step"]) or model["step"] <= 0:
        raise ValueError("Invalid expected-strike grid step")
    if any(number(model[k]) is None for k in ("x_min", "y_min")):
        raise ValueError("Invalid expected-strike grid origin")
    if any(not 0 <= p <= 1 for row in grid for p in row):
        raise ValueError("Expected-strike grid values must be probabilities")
    return model


def expected_strike(model, side, height):
    """Probability a pitch at this location is called a strike (bilinear interpolation
    on the model grid). None if the location is off the grid."""
    grid = model["probabilities"]
    step, x_min, y_min = model["step"], model["x_min"], model["y_min"]
    n_cols, n_rows = len(grid[0]), len(grid)
    if not (x_min <= side <= x_min + step * (n_cols - 1) and y_min <= height <= y_min + step * (n_rows - 1)):
        return None

    x = min(max((side - x_min) / step, 0), n_cols - 1)
    y = min(max((height - y_min) / step, 0), n_rows - 1)
    i = min(int(x), n_cols - 2)
    j = min(int(y), n_rows - 2)
    dx, dy = x - i, y - j
    lower = (1 - dx) * grid[j][i] + dx * grid[j][i + 1]
    upper = (1 - dx) * grid[j + 1][i] + dx * grid[j + 1][i + 1]
    return (1 - dy) * lower + dy * upper


def load_run_expectancy(source):
    data = source if isinstance(source, dict) else json.loads(Path(source).read_text(encoding="utf-8"))
    all_states = {f"{a}{b}{c}" for a in (0, 1) for b in (0, 1) for c in (0, 1)}
    if set(data["states"]) != all_states:
        raise ValueError("Run expectancy must contain all eight base states")
    for state, values in data["states"].items():
        if len(values) != 3 or any(number(v) is None or float(v) < 0 for v in values):
            raise ValueError(f"Invalid run expectancy for {state}")
    return data


def expectancy(table, bases, outs):
    if outs >= 3:
        return 0.0
    return float(table["states"]["".join(map(str, bases))][outs])


# ============================================================
# Run-value credit for blocks and throws
# ============================================================
# Bases are a tuple (first, second, third), 1 = occupied.

def bases_of(row):
    values = [number(row.get(f"R{i}B")) for i in (1, 2, 3)]
    if any(v not in (0, 1) for v in values):
        raise ValueError("Starting bases must be explicit 0/1 values; blanks are unknown")
    return tuple(int(v) for v in values)


def outs_of(row):
    value = number(row.get("Outs"))
    if value not in (0, 1, 2):
        raise ValueError("Outs must be 0, 1, or 2")
    return int(value)


def block_credit(table, bases, outs, success):
    """Credit for a successful block: the run value of every runner NOT moving up one
    base (a runner on third would have scored). Uses only the pre-pitch base/out
    state; count, batter result and advancement odds aren't modeled. A failed block
    is worth zero (we only credit positive plays)."""
    if not success:
        return 0.0
    if not any(bases):
        raise ValueError("A block with empty bases is outside the tracked population")
    everyone_advances = (0, bases[0], bases[1])
    raw = bases[2] + expectancy(table, everyone_advances, outs) - expectancy(table, bases, outs)
    return max(0.0, raw)


def steal_targets(value):
    """Which base(s) the runner(s) tried to steal: [2], [3] or [2, 3]."""
    text = str(value).strip().lower().replace(" ", "")
    if text in {"2", "2.0"}:
        return [2]
    if text in {"3", "3.0"}:
        return [3]
    if text in {"2&3", "2,3"}:
        return [2, 3]
    raise ValueError(f"Unsupported steal destination {value!r}; expected 2, 3, or 2&3")


def steal_result(bases, targets, advance=False, out_target=None):
    """Bases and runs after the stealing runner(s) move. `advance` = each runner took
    an extra base on the throw; `out_target` = the base where a runner was caught."""
    if len(targets) != len(set(targets)):
        raise ValueError("Repeated steal destination")
    after = list(bases)
    for target in targets:
        if not bases[target - 2]:
            raise ValueError(f"No starting runner for steal of base {target}")
        after[target - 2] = 0

    runs = 0
    for target in targets:
        if target == out_target:
            continue
        end = target + int(advance)
        if end >= 4:
            runs += 1
        elif after[end - 1]:
            raise ValueError("Runner movement would put two runners on one base")
        else:
            after[end - 1] = 1
    return tuple(after), runs


def throwing_credit(table, bases, outs, targets, outcome, advance=False, out_target=None):
    """Credit for a caught stealing: run value if the runner had been safe minus the
    run value after the out. A safe steal is worth zero.
    Returns (credit, bases after, runs scored on the play)."""
    if outcome == "safe":
        after, runs = steal_result(bases, targets, advance)
        return 0.0, after, runs
    if outcome != "out":
        raise ValueError("Steal outcome must be safe or out")
    if advance:
        raise ValueError("Caught stealing plus advance on throw needs individual runner outcomes")
    if len(targets) > 1 and out_target not in targets:
        raise ValueError("A double steal with an out needs OutBase (2 or 3)")

    out_target = out_target or targets[0]
    if_safe, safe_runs = steal_result(bases, targets)
    actual, actual_runs = steal_result(bases, targets, out_target=out_target)
    credit = (safe_runs + expectancy(table, if_safe, outs)) - (actual_runs + expectancy(table, actual, outs + 1))
    return max(0.0, credit), actual, actual_runs


# ============================================================
# Pitch-level analysis
# ============================================================

def pitch_group(row):
    pitch = row.get("TaggedPitchType", "").lower().replace("-", "").replace(" ", "")
    if pitch in {"fastball", "fourseamfastball", "twoseamfastball", "sinker", "cutter"}:
        return "Fastball"
    hand = {"Right": "RH", "Left": "LH"}.get(row.get("PitcherThrows"))
    if hand and pitch in {"slider", "curveball", "sweeper", "knucklecurve"}:
        return hand + " breaking"
    if hand and pitch in {"changeup", "splitter", "forkball"}:
        return hand + " offspeed"
    return "Other / unknown"


def abbreviated(full_name):
    """'Smith, John' -> 'j. smith' (how names are written on the manual sheet)."""
    if "," not in full_name:
        return full_name.strip().lower()
    last, first = full_name.split(",", 1)
    return (first.strip()[:1] + ". " + last.strip()).lower()


def called_pitches(pitches, model=None):
    """Called balls and strikes with location and expected-strike probability."""
    located = []
    for row in pitches:
        if row["PitchCall"] not in {"BallCalled", "StrikeCalled"}:
            continue
        side, height = number(row.get("PlateLocSide")), number(row.get("PlateLocHeight"))
        if side is None or height is None:
            continue
        located.append({"game": row["GameUID"], "pitch": row["PitchNo"], "catcher": row["Catcher"],
                        "side": side, "height": height, "call": row["PitchCall"],
                        "expected_strike": expected_strike(model, side, height) if model else None})
    return located


def throw_measurements(row):
    """(throw velocity, base, pop time) for a steal throw.

    Velocity counts on any logged steal (or a manually verified throw). A pop time
    only counts when we can tell which base the throw went to: either it was
    manually verified, or the ball's landing spot matches 2B/3B, the exchange + flight
    times add up, TrackMan's confidence isn't low, and that base was being stolen."""
    manual = yes(row.get("ThrowVerified")) and row.get("ThrowBase") in {"2", "3"}
    logged_steal = yes(row.get("StealAtt"))

    speed = number(row.get("ThrowSpeed"))
    if speed is None or speed <= 0 or not (logged_steal or manual):
        speed = None

    pop = number(row.get("PopTime"))
    if pop is None or pop <= 0:
        return speed, None, None
    if manual:
        return speed, int(row["ThrowBase"]), pop
    if not logged_steal:
        return speed, None, None

    # The export measures forward distance (X) and lateral position (Z) in feet.
    x, z = number(row.get("BasePositionX")), number(row.get("BasePositionZ"))
    if x is None or z is None:
        return speed, None, None
    if abs(x - 127.28) <= 15 and abs(z) <= 20:
        base = 2
    elif abs(x - 63.64) <= 15 and abs(z + 63.64) <= 20:
        base = 3
    else:
        return speed, None, None

    exchange, flight = number(row.get("ExchangeTime")), number(row.get("TimeToBase"))
    if exchange is None or flight is None or min(exchange, flight) <= 0 or abs(exchange + flight - pop) > .02:
        return speed, None, None

    confidence_fields = ("CatcherThrowCatchConfidence", "CatcherThrowReleaseConfidence",
                         "CatcherThrowLocationConfidence")
    if any(row.get(f, "").strip().lower() == "low" for f in confidence_fields):
        return speed, None, None

    try:
        targets = steal_targets(row.get("StealBase"))
    except ValueError:
        return speed, None, None
    return (speed, base, pop) if base in targets else (speed, None, None)


def analyze(rows, table):
    """Turn pitches into block and steal events with run credit, and collect
    anything worth double-checking for the review log."""
    issues, events = [], []
    last_in_half_inning = {}

    for row in rows:
        normalize_manual(row)
        catcher = row.get("Catcher", "")
        if not catcher:
            add_issue(issues, row, "Missing native catcher; pitch excluded from catcher pages")

        # Flag disagreements between TrackMan names and the manual sheet
        for native, manual in [("Catcher", "ManualCatcher"), ("Pitcher", "ManualPitcher"), ("Batter", "Hitter")]:
            if row.get(manual) and abbreviated(row.get(native, "")) != row[manual].lower():
                add_issue(issues, row, f"{manual}={row[manual]!r} differs from {native}={row.get(native)!r}; "
                                       "native identity retained")

        half_inning = (row["GameUID"], row.get("Inning"), row.get("Top/Bottom"))
        if half_inning in last_in_half_inning and last_in_half_inning[half_inning].get("Catcher") != catcher:
            before = last_in_half_inning[half_inning].get("Catcher") or "Unknown"
            add_issue(issues, row, f"Catcher changed within half-inning: {before} -> {catcher or 'Unknown'}; "
                                   "verify assignment")
        last_in_half_inning[half_inning] = row

        block, steal, passed = yes(row.get("BlockAtt")), yes(row.get("StealAtt")), yes(row.get("PassedB"))

        if row.get("PitchCall") == "BallinDirt" and not block:
            add_issue(issues, row, "Dirt tag without logged block; not an error by itself because "
                                   "empty-base blocks are excluded")

        if number(row.get("ThrowSpeed")) is not None or number(row.get("PopTime")) is not None:
            speed, base, pop = throw_measurements(row)
            if speed is None and pop is None:
                add_issue(issues, row, "Raw throw measurements excluded: no eligible matched steal throw")
            elif number(row.get("PopTime")) is not None and pop is None:
                add_issue(issues, row, "Velocity retained for logged steal; pop excluded because "
                                       "destination/timing checks did not pass")

        if not (block or steal or passed):
            continue

        event = {"game": row["GameUID"], "pitch": row["PitchNo"], "catcher": catcher or "Unknown",
                 "pitch_uid": row["PitchUID"], "block": block, "steal": steal, "passed_ball": passed,
                 "block_side": row.get("BlockSide", ""), "pitch_group": pitch_group(row),
                 "success": row.get("Success", "").lower(), "block_credit": 0.0,
                 "throw_credit": 0.0, "state": None, "outs": None, "after_steal": None,
                 "steal_runs": None, "steal_targets": [], "outcome": row.get("Safe/Out", "").lower()}

        # Without a valid base/out state we can't value the play; mark credit for review
        try:
            bases, outs = bases_of(row), outs_of(row)
            event["state"], event["outs"] = bases, outs
        except ValueError as exc:
            add_issue(issues, row, str(exc))
            if block and event["success"] != "no":
                event["block_credit"] = None
            if steal and event["outcome"] != "safe":
                event["throw_credit"] = None
            events.append(event)
            continue

        if block:
            try:
                if not any(bases):
                    raise ValueError("Logged block has empty bases; excluded from blocking opportunities")
                if event["success"] not in {"yes", "no"}:
                    raise ValueError("Block success must be yes or no")
                if event["block_side"].lower() not in {"left", "middle", "right"}:
                    add_issue(issues, row, "Unrecognized block side; kept as Other in side table")
                event["block_credit"] = block_credit(table, bases, outs, event["success"] == "yes")
                height = number(row.get("PlateLocHeight"))
                if height is not None and height >= 1.5:
                    add_issue(issues, row, "Logged block has plate height >=1.5 ft; review location/video, "
                                           "do not automatically relocate")
                if row.get("KorBB") in {"Walk", "Strikeout"}:
                    add_issue(issues, row, f'Block coincides with {row["KorBB"]}; credit uses pre-pitch '
                                           "base/out context, not the batter result")
            except ValueError as exc:
                event["block_credit"] = None
                add_issue(issues, row, str(exc))

        if steal:
            try:
                targets = steal_targets(row.get("StealBase"))
                event["steal_targets"] = targets
                if yes(row.get("DubSteal")) != (len(targets) == 2):
                    add_issue(issues, row, "DubSteal flag disagrees with StealBase")
                credit, after, runs = throwing_credit(table, bases, outs, targets, event["outcome"],
                                                      yes(row.get("AdvThrw")), number(row.get("OutBase")))
                event["throw_credit"], event["after_steal"], event["steal_runs"] = credit, after, runs
                if row.get("KorBB") in {"Walk", "Strikeout"}:
                    add_issue(issues, row, f'Steal coincides with {row["KorBB"]}; review combined play '
                                           "if valuing a caught stealing")
            except ValueError as exc:
                event["throw_credit"] = 0.0 if event["outcome"] == "safe" else None
                add_issue(issues, row, str(exc))

        # Plays that touch both blocking and throwing: make sure nothing is credited twice
        if block and steal and event["success"] == "yes":
            if event["outcome"] == "out":
                add_issue(issues, row, "Block with caught stealing: counted as a successful block; block credit 0 "
                                       "so the out is credited once, under throwing")
                event["block_credit"] = 0.0
            else:
                add_issue(issues, row, "Successful block and safe steal on same pitch conflict with "
                                       "no-runner-advanced success definition")
                event["block_credit"] = None
        if block and passed and event["success"] == "yes":
            add_issue(issues, row, "Successful block also marked passed ball; block credit withheld pending review")
            event["block_credit"] = None

        events.append(event)
    return events, issues


def eligible_blocks(events):
    """Blocks with runners on (or with unknown bases, so they still show up)."""
    return [e for e in events if e["block"] and (e["state"] is None or any(e["state"]))]


def credit_total(events, field):
    """Sum of a credit field, or None if any play is still waiting on review."""
    values = [e[field] for e in events]
    return None if any(v is None for v in values) else sum(values)


def summarize(rows, events, model=None):
    """Per-catcher totals: framing, blocking, throwing."""
    summaries = []
    for catcher in sorted({r["Catcher"] for r in rows if r["Catcher"]}):
        pitches = [r for r in rows if r["Catcher"] == catcher]
        catcher_events = [e for e in events if e["catcher"] == catcher]
        blocks = eligible_blocks(catcher_events)
        steals = [e for e in catcher_events if e["steal"]]

        modeled = [p for p in called_pitches(pitches, model) if p["expected_strike"] is not None]
        n = len(modeled)
        expected = sum(p["expected_strike"] for p in modeled)
        actual = sum(p["call"] == "StrikeCalled" for p in modeled)

        throws = [throw_measurements(r) for r in pitches]
        velocities = [speed for speed, _, _ in throws if speed is not None]
        pop_times = {}
        for base in (2, 3):
            pops = [pop for _, dest, pop in throws if dest == base and pop is not None]
            pop_times[str(base)] = {"n": len(pops), "average_seconds": sum(pops) / len(pops) if pops else None}

        summaries.append({
            "catcher": catcher,
            "pitches": len(pitches),
            "called_strikes": sum(r["PitchCall"] == "StrikeCalled" for r in pitches),
            "called_balls": sum(r["PitchCall"] == "BallCalled" for r in pitches),
            "block_attempts": len(blocks),
            "block_successes": sum(e["success"] == "yes" for e in blocks),
            "blocking_credit": credit_total(catcher_events, "block_credit"),
            "steal_plays": len(steals),
            "stealing_runners": sum(len(e["steal_targets"]) for e in steals),
            "caught_stealing_plays": sum(e["outcome"] == "out" for e in steals),
            "throwing_credit": credit_total(catcher_events, "throw_credit"),
            "passed_balls": sum(e["passed_ball"] for e in catcher_events),
            "modeled_called_pitches": n,
            "modeled_called_strike_rate": actual / n if n else None,
            "expected_called_strike_rate": expected / n if n else None,
            "called_strikes_above_expected_pp": 100 * (actual - expected) / n if n else None,
            "extra_called_strikes": actual - expected if n else None,
            "throw_velocity_n": len(velocities),
            "throw_velocity_mph": sum(velocities) / len(velocities) if velocities else None,
            "pop_times": pop_times,
        })
    return summaries


# ============================================================
# PDF report
# ============================================================

def write_pdf(path, rows, events, summaries, model=None):
    """One page per catcher: framing charts, blocking and throwing."""
    from PIL import Image
    from reportlab.lib.colors import HexColor, white
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    c = canvas.Canvas(str(path), pagesize=(612, 792))
    c.setTitle("CSUN Catcher Report")
    c.setAuthor("CSUN Baseball")
    red, ink, gray = HexColor("#be1732"), HexColor("#18212b"), HexColor("#66717e")
    pale, rule, green = HexColor("#f1f3f5"), HexColor("#d9dde2"), HexColor("#4d9847")

    dates = sorted({r["Date"] for r in rows})
    period = dates[0] if len(dates) == 1 else dates[0] + " to " + dates[-1]

    logo_image = Image.open(io.BytesIO(base64.b64decode(LOGO_B64))).convert("RGBA")
    logo = ImageReader(logo_image)
    logo_height = LOGO_WIDTH * logo_image.height / logo_image.width

    # ---- drawing helpers ----
    def text(x, y, value, size=9, bold=False, color=ink):
        c.setFillColor(color)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
        c.drawString(x, y, str(value))

    def section(x, y, label, width=540):
        text(x, y, label, 12, True)
        c.setStrokeColor(red)
        c.setLineWidth(1)
        c.line(x, y - 7, x + width, y - 7)

    def table(x, y, widths, headers, records, row_height=19):
        total_width = sum(widths)
        c.setFillColor(pale)
        c.rect(x, y - row_height, total_width, row_height, fill=1, stroke=0)
        for i, record in enumerate([headers] + records):
            xx = x
            yy = y - i * row_height
            for value, width in zip(record, widths):
                text(xx + 6, yy - 13, value, 8.2, i == 0)
                xx += width
            c.setStrokeColor(rule)
            c.setLineWidth(.35)
            c.line(x, yy - row_height, x + total_width, yy - row_height)

    def credit(value):
        return "--" if value is None else f"{value:.3f}"

    def percent(value):
        return "--" if value is None else f"{value:.1%}"

    def in_shadow_zone(x, y):
        def inside(margin):
            return abs(x) <= ZONE_HALF_WIDTH + margin and ZONE_BOTTOM - margin <= y <= ZONE_TOP + margin
        return inside(SHADOW_BAND) and not inside(-SHADOW_BAND)

    panel_w, panel_h, panel_bottom = 260, 176, 426

    def zone_panel(x0, title, x_extent, lo, hi, dots, show_probabilities):
        """Strike-zone chart from the catcher's view. Dots are (side, height, call)."""
        scale = min(panel_w / (2 * x_extent), panel_h / (hi - lo))
        left = x0 + (panel_w - 2 * x_extent * scale) / 2

        def xy(x, y):
            return left + (x + x_extent) * scale, panel_bottom + (y - lo) * scale

        text(x0, 612, title, 8.5, True)

        # Height gridlines
        for height in range(math.ceil(lo), math.ceil(hi)):
            if lo <= height <= hi:
                xx, yy = xy(-x_extent, height)
                c.setStrokeColor(HexColor("#e9ecef"))
                c.setLineWidth(.4)
                c.line(xx, yy, xy(x_extent, height)[0], yy)
                text(xx - 10, yy - 3, height, 7, color=gray)

        # Strike zone
        zx, zy = xy(-ZONE_HALF_WIDTH, ZONE_BOTTOM)
        c.setStrokeColor(ink)
        c.setLineWidth(1.3)
        c.rect(zx, zy, 2 * ZONE_HALF_WIDTH * scale, (ZONE_TOP - ZONE_BOTTOM) * scale, fill=0, stroke=1)

        # Shadow-zone edges (dashed)
        c.setStrokeColor(HexColor("#c4c9cf"))
        c.setLineWidth(.8)
        c.setDash(3, 2)
        for m in (SHADOW_BAND, -SHADOW_BAND):
            ax, ay = xy(-ZONE_HALF_WIDTH - m, ZONE_BOTTOM - m)
            c.rect(ax, ay, (2 * ZONE_HALF_WIDTH + 2 * m) * scale, (ZONE_TOP - ZONE_BOTTOM + 2 * m) * scale,
                   fill=0, stroke=1)
        c.setDash()

        # 3x3 grid inside the zone
        c.setStrokeColor(HexColor("#b9c0c7"))
        c.setLineWidth(.5)
        for x in [-ZONE_HALF_WIDTH / 3, ZONE_HALF_WIDTH / 3]:
            c.line(*xy(x, ZONE_BOTTOM), *xy(x, ZONE_TOP))
        for y in [ZONE_BOTTOM + 2 / 3, ZONE_BOTTOM + 4 / 3]:
            c.line(*xy(-ZONE_HALF_WIDTH, y), *xy(ZONE_HALF_WIDTH, y))

        # Home plate, if the chart goes low enough to show it
        if lo < .3:
            plate = c.beginPath()
            plate.moveTo(*xy(-ZONE_HALF_WIDTH, .1))
            for x, y in [(ZONE_HALF_WIDTH, .1), (ZONE_HALF_WIDTH, -.08), (0, -.32), (-ZONE_HALF_WIDTH, -.08)]:
                plate.lineTo(*xy(x, y))
            plate.close()
            c.setStrokeColor(ink)
            c.drawPath(plate, stroke=1, fill=0)

        # Pitches, drawn top to bottom so low pitches sit on top
        radius = 5.6 if show_probabilities else 3.3
        for x, y, call in sorted(dots, key=lambda d: -d[1]):
            px, py = xy(x, y)
            c.setFillColor(red if call == "StrikeCalled" else green)
            c.setStrokeColor(ink)
            c.setLineWidth(.55)
            c.circle(px, py, radius, stroke=1, fill=1)
            if show_probabilities and model:
                probability = expected_strike(model, x, y)
                if probability is not None:
                    c.setFillColor(white)
                    c.setFont("Helvetica-Bold", 5.6)
                    c.drawCentredString(px, py - 2, f"{probability * 100:.0f}")
        text(x0 + panel_w / 2 - 36, panel_bottom - 11, "Catcher's view (ft)", 7, color=gray)

    # ---- one page per catcher ----
    for page, summary in enumerate(summaries, 1):
        catcher = summary["catcher"]
        pitches = [r for r in rows if r["Catcher"] == catcher]
        blocks = eligible_blocks([e for e in events if e["catcher"] == catcher])

        # Header
        c.setFillColor(red)
        c.rect(0, 774, 612, 18, stroke=0, fill=1)
        c.drawImage(logo, LOGO_RIGHT - LOGO_WIDTH, LOGO_TOP - logo_height,
                    width=LOGO_WIDTH, height=logo_height, mask="auto")
        text(36, 734, catcher, 23, True)
        text(36, 715, "CATCHER REPORT  /  " + period, 10, color=gray)

        kpis = [("TOTAL PITCHES CAUGHT", summary["pitches"]),
                ("BLOCKING RUNS CREDITED", credit(summary["blocking_credit"])),
                ("THROWING RUNS CREDITED", credit(summary["throwing_credit"]))]
        for i, (label, value) in enumerate(kpis):
            x = 36 + i * 184
            c.setFillColor(pale)
            c.roundRect(x, 650, 172, 43, 5, stroke=0, fill=1)
            text(x + 10, 680, label, 7.5, True, gray)
            text(x + 10, 658, value, 20, True)

        # Framing
        section(36, 630, "FRAMING")
        located = [(number(r.get("PlateLocSide")), number(r.get("PlateLocHeight")), r["PitchCall"])
                   for r in pitches if r["PitchCall"] in {"BallCalled", "StrikeCalled"}]
        located = [(x, y, call) for x, y, call in located if x is not None and y is not None]

        x_extent = max([1.9] + [abs(x) + .2 for x, _, _ in located])
        lo = min([-.45] + [y - .2 for _, y, _ in located])
        hi = max([4.3] + [y + .2 for _, y, _ in located])
        zone_panel(36, "All called pitches", x_extent, lo, hi, located, False)
        zone_panel(316, "Shadow zone pitches: expected strike %", 1.3, 1.0, 4.0,
                   [d for d in located if in_shadow_zone(d[0], d[1])], True)
        for dx, label, color in [(150, "Called strike", red), (222, "Called ball", green)]:
            c.setFillColor(color)
            c.circle(dx, 615, 3, fill=1, stroke=0)
            text(dx + 6, 612, label, 7.5)

        delta = summary.get("called_strikes_above_expected_pp")
        stats = [("CALLED PITCHES", len(located)),
                 ("CALLED STRIKES", summary["called_strikes"]),
                 ("CALLED BALLS", summary["called_balls"]),
                 ("CSTR%", percent(summary.get("modeled_called_strike_rate"))),
                 ("XCSTR%", percent(summary.get("expected_called_strike_rate"))),
                 ("CSTR ABOVE EXPECTED", "--" if delta is None else f"{delta:+.1f} pp")]
        for i, (label, value) in enumerate(stats):
            x = 36 + i * 90
            text(x, 394, label, 7, True, gray)
            text(x, 376, value, 14, True)

        # Blocking
        section(36, 346, "BLOCKING")
        text(36, 326, f"With runners on: {summary['block_successes']}/{summary['block_attempts']} successful "
                      f"({rate(summary['block_successes'], summary['block_attempts'])})", 9)

        sides = ["left", "middle", "right"]
        if any(e["block_side"].lower() not in sides for e in blocks):
            sides.append("other")
        side_rows = []
        for side in sides:
            subset = [e for e in blocks
                      if (e["block_side"].lower() if e["block_side"].lower() in sides[:3] else "other") == side]
            made = sum(e["success"] == "yes" for e in subset)
            side_rows.append([side.title(), f"{made}/{len(subset)}", rate(made, len(subset))])
        table(36, 314, [94, 72, 70], ["Block side", "Made / Att.", "Success"], side_rows, 18)

        groups = PITCH_GROUPS + (["Other / unknown"] if any(e["pitch_group"] not in PITCH_GROUPS for e in blocks) else [])
        group_rows = []
        for group in groups:
            subset = [e for e in blocks if e["pitch_group"] == group]
            made = sum(e["success"] == "yes" for e in subset)
            group_rows.append([group, f"{made}/{len(subset)}", rate(made, len(subset))])
        table(302, 314, [120, 80, 74], ["Pitch group", "Made / Att.", "Success"], group_rows, 13.5)

        # Throwing
        section(36, 200, "THROWING")
        text(36, 180, f"Steal plays: {summary['steal_plays']}    Stealing runners: {summary['stealing_runners']}    "
                      f"CS plays: {summary['caught_stealing_plays']}    Passed balls: {summary['passed_balls']}", 8.6)
        throw_rows = []
        for base in (2, 3):
            pops = summary["pop_times"][str(base)]
            throw_rows.append([f"To {base}B", pops["n"], f"{pops['average_seconds']:.3f}" if pops["n"] else "--"])
        table(36, 168, [96, 70, 104], ["Throw to", "Pop n", "Avg pop (s)"], throw_rows, 18)

        velocity = summary["throw_velocity_mph"]
        text(330, 153, "STEAL-THROW VELOCITY", 8, True, gray)
        text(330, 133, "--" if velocity is None else f"{velocity:.1f} mph", 16, True)
        text(330, 115, f"Recorded throws: {summary['throw_velocity_n']}", 8, color=gray)
        text(36, 94, "Recorded TrackMan measurements", 8, color=gray)

        # Footer
        c.setStrokeColor(rule)
        c.line(36, 46, 576, 46)
        text(36, 31, f"CSUN / {period}", 7, color=gray)
        text(532, 31, f"{page} / {len(summaries)}", 7, color=gray)
        c.showPage()
    c.save()


# ============================================================
# Review log (markdown)
# ============================================================

def write_review_log(path, rows, events, summaries, issues, corrections, table, notes=None, model=None):
    """Totals, corrections, every credited play, and items to check against video."""
    def safe(v):
        return str(v).replace("|", "/").replace("\n", " ")

    def bases_text(v):
        return "unknown" if v is None else "".join(map(str, v))

    def value(v):
        return "REVIEW" if v is None else f"{v:.3f}"

    lines = [
        "# Catcher report review", "",
        f'{len(rows)} pitches; {sum(not r.get("Catcher") for r in rows)} without a catcher assignment.', "",
        "## Scoring convention", "",
        "All credits are nonnegative. Failed blocks, passed balls, and safe steals add zero. "
        "The report measures cumulative credit for logged successful plays, not net runs saved or runs above average. "
        "Blocking opportunities exclude empty-base pitches. An unlogged dirt pitch is not automatically a missed block.", "",
        "Credits use the pre-pitch base/out state. A successful block is compared with every existing runner "
        "advancing exactly one base; a runner on third scores in that hypothetical. Walks, strikeouts, count, "
        "pitch difficulty, and advancement probability are not included in this simple context score. "
        "Simultaneous batter outcomes are flagged below. A caught stealing is compared with that runner "
        "being safe at the attempted base, including the extra out in the caught-stealing scenario.", "",
        f'Baseline: [{table["name"]}]({table["source"]}). {table["note"]}', "",
        "## Catcher totals", "",
        "| Catcher | Pitches | Blocks made/attempted | Blocking credit | Throwing credit |",
        "|---|---:|---:|---:|---:|",
    ]
    for s in summaries:
        lines.append(f'| {safe(s["catcher"])} | {s["pitches"]} | {s["block_successes"]}/{s["block_attempts"]} | '
                     f'{value(s["blocking_credit"])} | {value(s["throwing_credit"])} |')

    lines += ["", "## Applied corrections", "", "Corrections are applied in memory; the source CSV is unchanged.", "",
              "| Game | Pitch | Field | Original | Corrected |", "|---|---:|---|---|---|"]
    for change in corrections:
        lines.append("| " + " | ".join(safe(change[k]) for k in ["game", "pitch", "field", "old", "new"]) + " |")

    lines += ["", "## Event ledger", "", "Base-state digits are first, second, third (1 = occupied).", "",
              "| Pitch | Catcher | Before | Outs | Block | Throw | After steal | Runs on steal/throw |",
              "|---:|---|---|---:|---:|---:|---|---:|"]
    for e in events:
        runs = e["steal_runs"] if e["steal_runs"] is not None else "--"
        lines.append(f'| {e["pitch"]} | {safe(e["catcher"])} | {bases_text(e["state"])} | {e["outs"]} | '
                     f'{value(e["block_credit"])} | {value(e["throw_credit"])} | {bases_text(e["after_steal"])} | '
                     f"{runs} |")

    lines += ["", "## Raw throw measurements (not validated)", "",
              "Eligible logged-steal velocities and destination/timing-matched pop times appear in the PDF. "
              "These are recorded TrackMan measurements, not independently video-validated values. "
              "Do not infer destination from a 1.8-second threshold.", "",
              "| Pitch | Catcher | Manual steal | Pop (s) | Exchange (s) | Flight (s) | Velo (mph) | Destination X / Y / Z (ft) |",
              "|---:|---|---|---:|---:|---:|---:|---|"]
    for r in rows:
        if number(r.get("ThrowSpeed")) is None and number(r.get("PopTime")) is None:
            continue
        cells = [str(r["PitchNo"]), r["Catcher"], r.get("StealBase", "") or "--"]
        cells += [r.get(k, "") or "--" for k in ["PopTime", "ExchangeTime", "TimeToBase", "ThrowSpeed"]]
        cells.append(" / ".join(r.get("BasePosition" + axis, "") or "--" for axis in "XYZ"))
        lines.append("| " + " | ".join(map(safe, cells)) + " |")

    lines += ["", "## Items to review", ""]
    lines += [f'- Pitch {i["pitch"]} ({safe(i["catcher"])}): {i["message"]}' for i in issues]
    lines += ["", "## Reconciliation notes", ""] + ["- " + note for note in (notes or [])]

    if model:
        lines += ["", "## Expected called-strike model", "",
                  f'Model: {model.get("name", "custom")}. See MODEL_NOTES.md and model_validation.json for training and holdout results.',
                  "xCStr% is the mean predicted called-strike probability. CStr above expected is actual minus expected in percentage points, using exactly the same modeled pitches. Negative differences are retained; the positive-only convention applies to blocking and throwing run credits.",
                  "Location-only comparison does not isolate catcher skill from umpire, count, batter height or other effects. Missing/out-of-grid locations are excluded from both modeled rates."]

    lines += ["", "## Sources", ""]
    lines += ["- " + p for p in sorted({r["_source"] for r in rows})]
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def save_report(rows, output_dir, table, model, corrections=None, notes=None, skipped=None):
    """Write the PDF, review log, JSON data and called-pitch CSV."""
    if not rows:
        raise ValueError("No pitches remain after filtering")
    if {r["GameUID"] for r in rows} & set(model.get("training_game_ids", [])):
        raise ValueError("Report games overlap model training games; use an independently trained model")

    events, issues = analyze(rows, table)
    issues = list(skipped or []) + issues
    summaries = summarize(rows, events, model)
    if not summaries:
        raise ValueError("No catcher assignments available for report pages")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    dates = sorted({r["Date"] for r in rows})
    label = dates[0] if len(dates) == 1 else dates[0] + "_to_" + dates[-1]

    pdf_path = output_dir / f"Catcher_Report_{label}.pdf"
    review_path = output_dir / f"Catcher_Review_{label}.md"
    write_pdf(pdf_path, rows, events, summaries, model)
    write_review_log(review_path, rows, events, summaries, issues, corrections or [], table, notes, model)

    data = {"baseline": table, "pitch_count": len(rows), "summaries": summaries,
            "events": events, "issues": issues, "corrections": corrections or [],
            "called_pitches": called_pitches(rows, model), "reconciliation_notes": notes or [],
            "model": {k: v for k, v in model.items() if k not in ("probabilities", "training_game_ids")}}
    (output_dir / f"Catcher_Data_{label}.json").write_text(json.dumps(data, indent=2), encoding="utf-8")

    with (output_dir / f"Catcher_Pitches_{label}.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["game", "pitch", "catcher", "side", "height", "call", "expected_strike"])
        writer.writeheader()
        writer.writerows(data["called_pitches"])

    print(f"Created {pdf_path}\nCreated {review_path}")
    return data


# ============================================================
# Putting it together
# ============================================================

def prepare(paths, manual_pairs=(), overrides=(), known=False, sheet=None,
            manual_sheet=None, manual_index="trackman", start=None, end=None, team="CAL_MAT"):
    """Load the games, merge manual tracking, apply corrections and filter.
    Returns (rows, changes, notes, skipped)."""
    files = expand_files(paths)
    skipped = []
    rows = load_rows(files, skipped, sheet)
    games = {r["GameUID"] for r in rows}

    # Saved correction bundles that apply to any of these games
    bundles = []
    if known:
        for name in KNOWN_CORRECTION_FILES:
            if not (BASE / name).exists():
                continue
            bundle = json.loads((BASE / name).read_text(encoding="utf-8"))
            if any(s["GameUID"] in games for s in bundle["overrides"]):
                bundles.append(bundle)

    changes, notes, manual_games = [], [], set()

    # Manual tracking sheets, each paired with exactly one game file
    for trackman_path, manual_path in manual_pairs:
        source = str(Path(trackman_path).expanduser().resolve())
        candidates = {r["GameUID"] for r in rows if r["_source"] == source}
        if len(candidates) != 1:
            raise ValueError("Each manual workbook must be paired with a selected TrackMan file "
                             "containing exactly one game")
        game = next(iter(candidates))
        if game in manual_games:
            raise ValueError("More than one manual workbook was supplied for the same game")
        manual_games.add(game)

        pitch_map = {}
        for bundle in bundles:
            if any(s["GameUID"] == game for s in bundle["overrides"]):
                pitch_map.update(bundle.get("manual_pitch_map", {}))

        imported, n = join_manual(rows, manual_path, game, pitch_map, manual_index == "measured", manual_sheet)
        changes += imported
        if manual_index == "measured":
            matching = "measured-pitch ordinal"
        elif pitch_map:
            matching = "saved game map"
        else:
            matching = "exact TrackMan PitchNo"
        notes.append(f"Imported {n} manual rows from {Path(manual_path).resolve()} for game {game}. "
                     f"Matching: {matching}.")

    for bundle in bundles:
        specs = [s for s in bundle["overrides"] if s["GameUID"] in games]
        changes += apply_overrides(rows, specs=specs)
        notes += bundle.get("notes", [bundle.get("note", "Applied bundled game corrections.")])

    for path in overrides:
        bundle = json.loads(Path(path).read_text(encoding="utf-8"))
        changes += apply_overrides(rows, path)
        notes += bundle.get("notes", [])

    # Team and date filters
    for value in (start, end):
        if value:
            date.fromisoformat(value)
    if start and end and start > end:
        raise ValueError("Start date must be on or before end date")
    rows = [r for r in rows
            if (team == "ALL" or r.get("CatcherTeam") == team)
            and (not start or r["Date"] >= start)
            and (not end or r["Date"] <= end)]
    if not rows:
        raise ValueError("No pitches match the team/date filters")

    # TrackMan alone can't tell us whether a block worked or a runner was safe
    selected_games = {r["GameUID"] for r in rows}
    for game in selected_games:
        game_rows = [r for r in rows if r["GameUID"] == game]
        has_manual = game in manual_games or any("BlockAtt" in r or "StealAtt" in r for r in game_rows)
        if not has_manual:
            raise ValueError(f"Game {game} has no manual tracking fields. Supply its manual workbook; "
                             "TrackMan alone cannot determine blocking success or steal outcomes.")

    notes.append("Manual names are retained for review. Future-game catcher substitutions are not inferred from "
                 "isolated manual events; correct the TrackMan catcher assignments or provide explicit "
                 "game-scoped overrides.")

    kept = {(r["GameUID"], r["PitchNo"]) for r in rows}
    changes = [c for c in changes if (c["game"], c["pitch"]) in kept]
    skipped = [s for s in skipped if s["game"] in selected_games]
    return rows, changes, notes, skipped


def run(paths, output_dir=None, combined=False, daily_too=False, model_path=None, re_path=None, **kwargs):
    """Build one report from one game (daily) or several (combined)."""
    rows, changes, notes, skipped = prepare(paths, **kwargs)
    games = {r["GameUID"] for r in rows}
    if not combined and len(games) != 1:
        raise ValueError("Daily reports require one game. Use Combined_Catcher_Reports.py for multiple games.")

    model = load_strike_model(model_path or embedded_strike_model())
    table = load_run_expectancy(re_path or RUN_EXPECTANCY)
    output_dir = Path(output_dir or BASE / "reports")
    result = save_report(rows, output_dir, table, model, changes, notes, skipped)

    if combined and daily_too:
        for game in sorted(games):
            subset = [r for r in rows if r["GameUID"] == game]
            target = output_dir / "individual_games" / (subset[0]["Date"] + "_" + game[:8])
            save_report(subset, target, table, model,
                        [c for c in changes if c["game"] == game], notes,
                        [i for i in skipped if i["game"] == game])

    print(f"Done: {len(games)} game(s), {len(rows)} pitches. Reports: {output_dir.resolve()}")
    return result


def cli(combined=False, argv=None, data_dir=None, output_dir=None):
    """Command-line entry point shared by the daily and combined scripts."""
    if data_dir and not Path(data_dir).exists():
        alternate = Path(data_dir).with_name(Path(data_dir).name.replace("_", " "))
        if alternate.exists():
            data_dir = alternate

    parser = argparse.ArgumentParser(description="CSUN combined catcher report" if combined
                                     else "CSUN daily catcher report")
    parser.add_argument("files", nargs="*" if combined else 1,
                        help="Game file name or path (a name is looked up in the Data folder). "
                             "Combined with no files uses every file in the Data folder.")
    if combined:
        parser.add_argument("--manual-for", nargs=2, action="append", default=[],
                            metavar=("TRACKMAN_FILE", "MANUAL_WORKBOOK"))
        parser.add_argument("--daily-too", action="store_true",
                            help="Also create individual-game PDFs with boxed pitch labels")
    else:
        parser.add_argument("--manual", type=Path,
                            help="Separate manual tracking XLSX/CSV (optional when event fields are already in TrackMan)")
    parser.add_argument("--output-dir", type=Path, default=output_dir or BASE / "reports")
    parser.add_argument("--start-date", help="Inclusive YYYY-MM-DD")
    parser.add_argument("--end-date", help="Inclusive YYYY-MM-DD")
    parser.add_argument("--sheet", help="TrackMan worksheet when an Excel file has multiple data sheets")
    parser.add_argument("--manual-sheet", help="Manual worksheet, usually Input")
    parser.add_argument("--manual-index", choices=["trackman", "measured"], default="trackman",
                        help="Use measured only when manual numbering skips empty TrackMan records")
    parser.add_argument("--overrides", action="append", type=Path, default=[])
    parser.add_argument("--no-known-corrections", action="store_true")
    parser.add_argument("--team", default="CAL_MAT")
    parser.add_argument("--strike-model", type=Path)
    parser.add_argument("--re-table", type=Path)
    args = parser.parse_args(argv)

    try:
        if combined and not args.files:
            if not data_dir:
                parser.error("give at least one game file")
            args.files = [str(data_dir)]
        else:
            args.files = [str(resolve_input(f, data_dir)) for f in args.files]
    except ValueError as error:
        parser.exit(1, f"Report not created: {error}\n")

    if combined:
        pairs = args.manual_for
    else:
        pairs = [(args.files[0], args.manual)] if args.manual else []

    try:
        return run(args.files, args.output_dir, combined=combined, daily_too=getattr(args, "daily_too", False),
                   manual_pairs=pairs, overrides=args.overrides, known=not args.no_known_corrections,
                   sheet=args.sheet, manual_sheet=args.manual_sheet, manual_index=args.manual_index,
                   start=args.start_date, end=args.end_date, team=args.team,
                   model_path=args.strike_model, re_path=args.re_table)
    except (ValueError, OSError, KeyError, ImportError) as error:
        parser.exit(1, f"Report not created: {error}\n")


# ============================================================
# Embedded data
# ============================================================
# Expected-strike model: called-strike probability on a 0.05 ft grid, trained on
# our historical called pitches (location only). Stored compressed so the script
# runs without extra files.

STRIKE_MODEL_INFO = {'name': 'CSUN historical location model v1 (embedded, 3-decimal)',
 'x_min': -5.0,
 'y_min': -2.0,
 'step': 0.05,
 'training_pitches': 91280,
 'training_games': 741,
 'bandwidth_ft': 0.08,
 'test_metrics': {'pitches': 15546,
                  'brier': 0.06627706186276851,
                  'log_loss': 0.21476652077245567,
                  'actual_rate': 0.3252283545606587,
                  'predicted_rate': 0.3301067148236391}}

STRIKE_MODEL_GRID = (
    "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"
    "w8BGVb1Mu7KuWkO/a0uSVMCSvnacyPCiVLFQu3ejZi6M2BHEvFECDeOwEwdFUNpOvMylB0Uwlt1rxGnWyoCxoWQlsUSgQxqHl8MP"
    "Bv7ryFZrABRJLJ7B94R/qJMqzgEBqkKTTK9MOJ+tNNbhoKAl/e6ZCgKYx11jc5HHHV9uzU57jurXiMYJtGLCXfzYieBjjS9hj5Cw"
    "AiIRa5C3wVMvhZEqlsQizju2MIK2s8ZqK9gtIEBsYXxfYmLJEbVoqHlth4Bt26SsXgbhSvpMo6scUkxCF28Za7FLvZKF+Ix560rb"
    "EiudRgIenaAywif1Qr0USZs/TwGe44d1xvle4/VcM9iKayu0xSPx5CUFVuOudg0TilrSp8sV60lSNDaCzXCaz2TMlWmzeyYzPA/T"
    "aMYRB5MAHGOFeThIK72EDO4LbZUt5BRa4iDULRfmTUPoFYr4aHDdsmynlWpaYpb0KYxMr27jejf2ovLPKzqIEtR2n8gt3bk2KVOj"
    "prErJibK48bJZ63j+ODqhXWbqtk7QPvtkNM+ivMMvMo6DRC1pFc9D3Nr896rOmsbsqNtDVcjePyafDPskolNDJhyMwqzTyAvutL5"
    "6rKYdZe9ZcuEz77X05SvNleO/DSrqMUr6XdHWfOEV+N2zpo2EPONxNRYG2P6lmVuBgmvfDYTs59UyesJm+4b+NYTXq/Nn+G74hn6"
    "1Xq7L+8fJMIr4BK0ft8sTz05yz/JnnZ05Xtx0CsRz7iJj3bjqXx30Ls9PDKTT75gvfvO9QnywBGZMY4TepxHVDr4/NXWN2KWOh8e"
    "Q8bGry0l6yeA2Vk4HjiQOwNPfuO+x74h9DmbRumfAIRYLQZBR3SaZE7u9bdYzJL+W1tEf5gDOw3OUrbHryH/0Zbas8OG/JS/7Um7"
    "kkoYcDQ6bk1dsZa4Jf3DrtRnnsnLr+e6V8nrc5a06ZDUPTiimyNxNq6ylrAl/X1npc8cdPWtlsgpuc4QxS3p7zgz/VmzOVmT2+KW"
    "VPR65SRRDBC3pNKfwzD9HsQtqVQqFbekUqlU3JJKn63BV78TcUsqfe5hqvFDujDELalUxLrsn0xhU5ipK8s7MkQvcUsqfdaBnRxj"
    "0WiFT1kDO+KWVPo83EpkxQoD6ZHCglqDhuKWVPqE3IqwCr76AFarMXOYqrAlbEmlT6Sd5qMn2f6lMKD2Er3KW+KWVPqsA9KNSloa"
    "kBa3pNKf42CmBghxSyqVSsUtqVQqbkmlUqm4JZVKpeKWVCoVt6RSqVTckkqlUnFLKpWKW1KpVCpuSaVSqbgllUrFLalUKhW3pFKp"
    "VNySSqXillQqlYpbUqlUKm5JpVJxS9ySSqXillQqlYpbUqlUKm5JpVJxSyqVSsUtqVQqFbekUqm4JZVKpeKWVCqViltSqVTckkql"
    "UnFLKpVKxS2pVCpuSaVSqbgllUql4pZUKhW3pFKpVNySSqVScUsqlYpbUqlUKm5JpVKpuCWVSqXillQqFbekUqlU3JJKpVJxSyqV"
    "iltSqVQqbkmlUqm4JZVKxS2pVCoVt6RSqVTckkql4pZUKpWKW1KpVCpuSaVScUsqlUrFLalUKhW3pFKpuCWVSqXillQqlYpbUqlU"
    "3JJKpVJxSyqVSsUtqVQqFbekUulP5Nb//h9Lcodr"
)

# Run expectancy by base state (first/second/third occupied) and outs (0, 1, 2)
RUN_EXPECTANCY = {'name': '64Analytics published Division I reference',
 'source': 'https://www.64analytics.com/articles/bringing-mlb-precision-to-the-college-game',
 'table_image': 'https://www.64analytics.com/images/re-matrix.png',
 'publication_date': '2025-01-30',
 'accessed': '2026-10-30',
 'note': 'Published D1 reference, not CSUN-specific. The article does not specify the seasons used '
         'for this displayed matrix. State digits mean occupied first, second, third; columns mean '
         '0, 1, 2 outs.',
 'states': {'000': [0.793, 0.417, 0.16],
            '100': [1.303, 0.762, 0.327],
            '010': [1.586, 0.937, 0.433],
            '110': [2.094, 1.274, 0.6],
            '001': [1.825, 1.234, 0.488],
            '101': [2.312, 1.516, 0.678],
            '011': [2.53, 1.71, 0.744],
            '111': [3.006, 2.03, 0.997]}}

# CSUN Baseball logo (PNG, base64)
LOGO_B64 = (
    "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"
    "6/8Aze9Ey4+6r2sAAAAASUVORK5CYII="
)


if __name__ == "__main__":
    cli(combined=False, data_dir=ROOT / "Data" / "single_game_csv", output_dir=ROOT / "Reports")
