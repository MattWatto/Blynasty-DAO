"""
Pulls rosters, league users, and traded picks from the Sleeper API, assigns every
draft pick (traded or not) to whichever team currently owns it, joins player names
from a locally-saved players file, and writes one Excel sheet per export run with
each team's roster listed vertically (players first, then picks).

Run with: python build_roster_export.py
Requires: pip install requests pandas openpyxl

Sleeper API docs: https://docs.sleeper.com/
"""

from pathlib import Path

import requests
import pandas as pd

# ---- config ----------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent   # Input/Output are resolved relative to this script's own location, not the working directory
LEAGUE_ID = "1314308900652056576"
PLAYERS_FILE = BASE_DIR / "Input" / "sleeper_players.xlsx"   # produced by the separate player-pull script
PLAYER_ID_COL = "player_id"             # column names in PLAYERS_FILE -- adjust if different
PLAYER_NAME_COL = "full_name"
OUTPUT_FILE = BASE_DIR / "Output" / "roster_export.xlsx"
FUTURE_DRAFT_YEARS = 2                  # how many draft classes ahead of the current season to include
DEFAULT_ROUNDS = 4                      # fallback if the draft's round count can't be found
BASE_URL = "https://api.sleeper.app/v1"
# -----------------------------------------------------------------------------


def get_json(path):
    resp = requests.get(f"{BASE_URL}{path}", timeout=15)
    resp.raise_for_status()
    return resp.json()


def get_draft_rounds(league_id):
    drafts = get_json(f"/league/{league_id}/drafts")
    if drafts and drafts[0].get("settings", {}).get("rounds"):
        return drafts[0]["settings"]["rounds"]
    return DEFAULT_ROUNDS


def main():
    league = get_json(f"/league/{LEAGUE_ID}")
    rosters = get_json(f"/league/{LEAGUE_ID}/rosters")
    users = get_json(f"/league/{LEAGUE_ID}/users")
    traded_picks = get_json(f"/league/{LEAGUE_ID}/traded_picks")
    rounds = get_draft_rounds(LEAGUE_ID)
    current_season = int(league["season"])

    # team display name per roster_id (prefer the team nickname, fall back to display name)
    user_by_id = {u["user_id"]: u for u in users}
    team_name_by_roster = {}
    for r in rosters:
        user = user_by_id.get(r["owner_id"], {})
        team_name_by_roster[r["roster_id"]] = (
            (user.get("metadata") or {}).get("team_name") or user.get("display_name") or f"Roster {r['roster_id']}"
        )

    # current owner roster_id for every (season, round, original_roster_id) pick,
    # defaulting to the original owner unless a trade moved it
    traded_owner = {
        (p["season"], p["round"], p["roster_id"]): p["owner_id"] for p in traded_picks
    }
    pick_ownership = {}
    for season in range(current_season + 1, current_season + 1 + FUTURE_DRAFT_YEARS):
        for rnd in range(1, rounds + 1):
            for r in rosters:
                orig_roster_id = r["roster_id"]
                key = (str(season), rnd, orig_roster_id)
                pick_ownership[key] = traded_owner.get(key, orig_roster_id)

    # player_id -> full name
    players = pd.read_excel(PLAYERS_FILE, dtype={PLAYER_ID_COL: str})
    player_name = dict(zip(players[PLAYER_ID_COL], players[PLAYER_NAME_COL]))

    team_rows = {team: [] for team in team_name_by_roster.values()}

    for r in rosters:
        team = team_name_by_roster[r["roster_id"]]
        for pid in r.get("players") or []:
            team_rows[team].append(player_name.get(str(pid), str(pid)))

    for (season, rnd, orig_roster_id), owner_roster_id in pick_ownership.items():
        team = team_name_by_roster[owner_roster_id]
        if owner_roster_id == orig_roster_id:
            label = f"{season} Rd {rnd}"
        else:
            label = f"{season} Rd {rnd} (from {team_name_by_roster[orig_roster_id]})"
        team_rows[team].append(label)

    export = pd.DataFrame({team: pd.Series(rows) for team, rows in team_rows.items()})
    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    export.to_excel(OUTPUT_FILE, index=False)
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
