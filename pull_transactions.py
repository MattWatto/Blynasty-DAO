"""
Walks a dynasty league's full history (Sleeper gives each season its own league_id,
chained together via previous_league_id) and pulls every transaction from every
week of every season. Player names come from a locally-saved players file; team
names come from each season's own rosters/users, since roster ownership can change
season to season. Writes one readable row per transaction to Excel.

Run with: python pull_transactions.py
Requires: pip install requests pandas openpyxl

Sleeper API docs: https://docs.sleeper.com/
"""

from pathlib import Path

import requests
import pandas as pd

# ---- config ----------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent   # Input/Output are resolved relative to this script's own location
LEAGUE_ID = "1314308900652056576"            # the most recent season's league_id; earlier seasons are discovered automatically
PLAYERS_FILE = BASE_DIR / "Input" / "sleeper_players.xlsx"
PLAYER_ID_COL = "player_id"
PLAYER_NAME_COL = "full_name"
OUTPUT_FILE = BASE_DIR / "Output" / "transactions_export.xlsx"
WEEKS = range(1, 19)                         # weeks checked for transactions in each season
BASE_URL = "https://api.sleeper.app/v1"
# -----------------------------------------------------------------------------


def get_json(path):
    resp = requests.get(f"{BASE_URL}{path}", timeout=15)
    resp.raise_for_status()
    return resp.json()


def get_season_chain(league_id):
    """Follow previous_league_id back to the first season, returning [(season, league_id), ...] oldest first."""
    seasons = []
    while league_id and league_id != "0":
        league = get_json(f"/league/{league_id}")
        seasons.append((league["season"], league_id))
        league_id = league.get("previous_league_id")
    seasons.reverse()
    return seasons


def get_team_names(league_id):
    """roster_id -> team name for one season (team nickname, falling back to the owner's display name)."""
    users_by_id = {u["user_id"]: u for u in get_json(f"/league/{league_id}/users")}
    team_names = {}
    for r in get_json(f"/league/{league_id}/rosters"):
        user = users_by_id.get(r["owner_id"], {})
        team_names[r["roster_id"]] = (
            (user.get("metadata") or {}).get("team_name") or user.get("display_name") or f"Roster {r['roster_id']}"
        )
    return team_names


def team_name(roster_id, team_names):
    return team_names.get(roster_id, f"Roster {roster_id}")


def describe_moves(mapping, verb, team_names, player_names):
    if not mapping:
        return ""
    return "; ".join(
        f"{player_names.get(pid, pid)} ({verb} by {team_name(rid, team_names)})" for pid, rid in mapping.items()
    )


def describe_draft_picks(picks, team_names):
    if not picks:
        return ""
    return "; ".join(f"{p['season']} Rd {p['round']} -> {team_name(p['owner_id'], team_names)}" for p in picks)


def describe_waiver_budget(entries, team_names):
    if not entries:
        return ""
    return "; ".join(
        f"${e['amount']} {team_name(e['sender'], team_names)} -> {team_name(e['receiver'], team_names)}"
        for e in entries
    )


def main():
    players = pd.read_excel(PLAYERS_FILE, dtype={PLAYER_ID_COL: str})
    player_names = dict(zip(players[PLAYER_ID_COL], players[PLAYER_NAME_COL]))

    seasons = get_season_chain(LEAGUE_ID)

    rows = []
    for season, league_id in seasons:
        team_names = get_team_names(league_id)
        for week in WEEKS:
            for tx in get_json(f"/league/{league_id}/transactions/{week}"):
                rows.append({
                    "season": season,
                    "week": week,
                    "transaction_id": tx.get("transaction_id"),
                    "type": tx.get("type"),
                    "status": tx.get("status"),
                    "created": tx.get("created"),
                    "teams_involved": "; ".join(team_name(rid, team_names) for rid in tx.get("roster_ids") or []),
                    "adds": describe_moves(tx.get("adds"), "added", team_names, player_names),
                    "drops": describe_moves(tx.get("drops"), "dropped", team_names, player_names),
                    "draft_picks": describe_draft_picks(tx.get("draft_picks"), team_names),
                    "waiver_budget": describe_waiver_budget(tx.get("waiver_budget"), team_names),
                })

    export = pd.DataFrame(rows)
    if not export.empty:
        export["created"] = pd.to_datetime(export["created"], unit="ms")

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    export.to_excel(OUTPUT_FILE, index=False)
    print(f"Wrote {OUTPUT_FILE} ({len(rows)} transactions across {len(seasons)} seasons)")


if __name__ == "__main__":
    main()
