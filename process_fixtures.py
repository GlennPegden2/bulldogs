#!/usr/bin/env python3
"""
Process Junior Ice Hockey Fixtures spreadsheet and generate text lists
and iCalendar (.ics) files organised by Team/Age-Division, Team, and Rink.

Usage:
    python process_fixtures.py [path/to/MasterFixtures Official.xlsx]

If no path is supplied the script looks for 'MasterFixtures Official.xlsx'
in the same directory as this script.
"""

import os
import re
import sys
from datetime import datetime, timedelta

import openpyxl
from icalendar import Calendar, Event
from uuid import uuid4


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

RENAME_MAP = {
    "Blue": "Leeds Blue",
    "White": "Leeds White",
}


def normalise_team_name(raw: str) -> str:
    """
    Extract the team display name from the composite 'Rink-Division-TeamName'
    field and apply any renaming rules.

    Some entries contain multiple teams separated by ':' (e.g. multi-team
    fixtures).  In that case each component is normalised and the results are
    joined with ' & '.

    Returns the display name (or names joined with ' & ').
    """
    parts_multi = str(raw).split(":")
    names = []
    for part in parts_multi:
        segments = part.strip().split("-", 2)
        if len(segments) == 3:
            team = segments[2].strip()
        else:
            team = part.strip()
        names.append(RENAME_MAP.get(team, team))
    return " & ".join(names)


def team_names_list(raw: str) -> list:
    """
    Return a list of individual team display names from a (possibly
    colon-separated) composite team field.
    """
    parts_multi = str(raw).split(":")
    names = []
    for part in parts_multi:
        segments = part.strip().split("-", 2)
        if len(segments) == 3:
            team = segments[2].strip()
        else:
            team = part.strip()
        names.append(RENAME_MAP.get(team, team))
    return names


def safe_filename(name: str) -> str:
    """Return a filesystem-safe version of *name*."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def make_calendar(name: str) -> Calendar:
    cal = Calendar()
    cal.add("prodid", f"-//Ice Hockey Fixtures//{name}//EN")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    cal.add("x-wr-calname", name)
    return cal


def add_event(cal: Calendar, fixture: dict) -> None:
    event = Event()
    event.add("uid", str(uuid4()))
    dt_start = datetime.combine(fixture["date"], fixture["time"])
    dt_end = dt_start + timedelta(hours=2)
    event.add("dtstart", dt_start)
    event.add("dtend", dt_end)
    event.add(
        "summary",
        f"{fixture['home']} vs {fixture['away']}",
    )
    event.add(
        "description",
        (
            f"Home: {fixture['home']}\n"
            f"Away: {fixture['away']}\n"
            f"Rink: {fixture['rink']}\n"
            f"Division: {fixture['division']}"
        ),
    )
    event.add("location", fixture["rink"])
    cal.add_component(event)


def write_txt(path: str, fixtures: list) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        for f in fixtures:
            fh.write(
                f"{f['date'].strftime('%Y-%m-%d')}  {f['time'].strftime('%H:%M')}  "
                f"{f['home']} vs {f['away']}  [{f['rink']}]  ({f['division']})\n"
            )


def write_ics(path: str, name: str, fixtures: list) -> None:
    cal = make_calendar(name)
    for f in fixtures:
        add_event(cal, f)
    with open(path, "wb") as fh:
        fh.write(cal.to_ical())


# ---------------------------------------------------------------------------
# Load spreadsheet
# ---------------------------------------------------------------------------

def load_fixtures(xlsx_path: str) -> list:
    """Return a list of fixture dicts from the spreadsheet."""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active

    fixtures = []
    for row in ws.iter_rows(values_only=True):
        date_val, time_val, home_raw, away_raw, rink, division = (
            row[0], row[1], row[2], row[3], row[4], row[5]
        )

        # Skip rows missing essential fields
        if not all([date_val, time_val, home_raw, away_raw, rink, division]):
            continue

        # date_val may already be a datetime object (openpyxl behaviour)
        if isinstance(date_val, datetime):
            date = date_val.date()
        else:
            continue

        # time_val is a datetime.time object from openpyxl
        import datetime as dt_mod
        if isinstance(time_val, dt_mod.time):
            time = time_val
        else:
            continue

        home = normalise_team_name(str(home_raw))
        away = normalise_team_name(str(away_raw))
        home_list = team_names_list(str(home_raw))
        away_list = team_names_list(str(away_raw))
        rink = str(rink).strip()
        division = str(division).strip()

        fixtures.append(
            {
                "date": date,
                "time": time,
                "home": home,
                "away": away,
                "home_list": home_list,
                "away_list": away_list,
                "rink": rink,
                "division": division,
            }
        )

    return fixtures


# ---------------------------------------------------------------------------
# Grouping helpers
# ---------------------------------------------------------------------------

def group_by_team_division(fixtures: list) -> dict:
    """Key: (team_name, division)."""
    groups: dict = {}
    for f in fixtures:
        for team in f["home_list"] + f["away_list"]:
            key = (team, f["division"])
            groups.setdefault(key, []).append(f)
    return groups


def group_by_team(fixtures: list) -> dict:
    """Key: team_name (all divisions combined)."""
    groups: dict = {}
    for f in fixtures:
        for team in f["home_list"] + f["away_list"]:
            groups.setdefault(team, []).append(f)
    return groups


def group_by_rink(fixtures: list) -> dict:
    """Key: rink name."""
    groups: dict = {}
    for f in fixtures:
        groups.setdefault(f["rink"], []).append(f)
    return groups


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

def write_group(base_dir: str, label: str, fixtures: list) -> None:
    """Write both .txt and .ics for a single group into *base_dir*."""
    name = safe_filename(label)
    write_txt(os.path.join(base_dir, f"{name}.txt"), fixtures)
    write_ics(os.path.join(base_dir, f"{name}.ics"), label, fixtures)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(xlsx_path: str, output_dir: str = "output") -> None:
    print(f"Reading fixtures from: {xlsx_path}")
    fixtures = load_fixtures(xlsx_path)
    print(f"Loaded {len(fixtures)} fixtures")

    # Create output directories
    dir_team_div = os.path.join(output_dir, "by_team_division")
    dir_team = os.path.join(output_dir, "by_team")
    dir_rink = os.path.join(output_dir, "by_rink")
    for d in (dir_team_div, dir_team, dir_rink):
        os.makedirs(d, exist_ok=True)

    # 1. By team/age-division
    for (team, division), group_fixtures in group_by_team_division(fixtures).items():
        label = f"{team} - {division}"
        write_group(dir_team_div, label, sorted(group_fixtures, key=lambda x: (x["date"], x["time"])))

    # 2. By team (all divisions)
    for team, group_fixtures in group_by_team(fixtures).items():
        write_group(dir_team, team, sorted(group_fixtures, key=lambda x: (x["date"], x["time"])))

    # 3. By rink
    for rink, group_fixtures in group_by_rink(fixtures).items():
        write_group(dir_rink, rink, sorted(group_fixtures, key=lambda x: (x["date"], x["time"])))

    print(f"\nOutput written to: {output_dir}/")
    print(f"  by_team_division/ : {len(group_by_team_division(fixtures))} groups")
    print(f"  by_team/          : {len(group_by_team(fixtures))} groups")
    print(f"  by_rink/          : {len(group_by_rink(fixtures))} groups")


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_xlsx = os.path.join(script_dir, "MasterFixtures Official.xlsx")

    if len(sys.argv) >= 2:
        xlsx_path = sys.argv[1]
    else:
        xlsx_path = default_xlsx

    if not os.path.exists(xlsx_path):
        print(f"Error: spreadsheet not found at {xlsx_path}", file=sys.stderr)
        sys.exit(1)

    output_dir = sys.argv[2] if len(sys.argv) >= 3 else os.path.join(script_dir, "output")
    main(xlsx_path, output_dir)
