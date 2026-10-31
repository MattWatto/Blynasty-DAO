"""
Walks a dynasty league's full history (Sleeper gives each season its own league_id,
chained together via previous_league_id) and pulls every transaction from every
week of every season, writing them all to one Excel file.

Run with: python pull_transactions.py
Requires: pip install requests pandas openpyxl

Sleeper API docs: https://docs.sleeper.com/
"""

import json
from pathlib import Path

import requests
import pandas as pd

# ---- config ----------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent   # Output is resolved relative to this script's own location
LEAGUE_ID = "1314308900652056576"            # the most recent season's league_id; earlier seasons are discovered automatically
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


def main():
    seasons = get_season_chain(LEAGUE_ID)

    rows = []
    for season, league_id in seasons:
        for week in WEEKS:
            for tx in get_json(f"/league/{league_id}/transactions/{week}"):
                rows.append({
                    "season": season,
                    "week": week,
                    "transaction_id": tx.get("transaction_id"),
                    "type": tx.get("type"),
                    "status": tx.get("status"),
                    "created": tx.get("created"),
                    "roster_ids": json.dumps(tx.get("roster_ids")),
                    "adds": json.dumps(tx.get("adds")),
                    "drops": json.dumps(tx.get("drops")),
                    "draft_picks": json.dumps(tx.get("draft_picks")),
                    "waiver_budget": json.dumps(tx.get("waiver_budget")),
                })

    export = pd.DataFrame(rows)
    if not export.empty:
        export["created"] = pd.to_datetime(export["created"], unit="ms")

    OUTPUT_FILE.parent.mkdir(exist_ok=True)
    export.to_excel(OUTPUT_FILE, index=False)
    print(f"Wrote {OUTPUT_FILE} ({len(rows)} transactions across {len(seasons)} seasons)")


if __name__ == "__main__":
    main()
