#!/usr/bin/env python3
"""
Generate documentation from Domino governance bundles
"""
import os
import sys
from datetime import datetime
import requests

# Configuration
DOMINO_DOMAIN = os.environ.get("DOMINO_DOMAIN", "se-demo.domino.tech")
DOMINO_API_KEY = os.environ.get("DOMINO_USER_API_KEY", "")
DOMINO_PROJECT_ID = os.environ.get("DOMINO_PROJECT_ID", "")


def parse_timestamp(timestamp_str):
    """Parse ISO timestamp with robust error handling"""
    try:

        # Handle Z suffix
        ts = timestamp_str.replace("Z", "+00:00")
        # Ensure microseconds are exactly 6 digits
        if "." in ts and "+" in ts:
            before_dot, after_dot = ts.split(".")
            microseconds, tz = after_dot.split("+")
            microseconds = microseconds[:6].ljust(6, "0")
            ts = f"{before_dot}.{microseconds}+{tz}"
        return datetime.fromisoformat(ts)
    except Exception as e:
        print(f"Warning: Failed to parse timestamp '{timestamp_str}': {e}")
        return None


def get_bundles(project_id):
    """Fetch bundles for a given project ID"""
    if not DOMINO_API_KEY:
        raise ValueError("DOMINO_USER_API_KEY environment variable is not set")

    domain = DOMINO_DOMAIN.removeprefix("https://").removeprefix("http://")
    url = f"https://{domain}/api/governance/v1/bundles"
    headers = {
        "accept": "application/json",
        "X-Domino-Api-Key": DOMINO_API_KEY
    }
    params = {"project_id": project_id}

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    return response.json()["data"]


def get_results(bundle_id):
    """Fetch results for a given bundle ID"""
    domain = DOMINO_DOMAIN.removeprefix("https://").removeprefix("http://")
    url = f"https://{domain}/api/governance/v1/results"
    headers = {
        "accept": "application/json",
        "X-Domino-Api-Key": DOMINO_API_KEY
    }
    params = {"bundleID": bundle_id}

    try:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        return response.json()["data"]
    except requests.RequestException as e:
        print(f"Warning: Failed to fetch results for bundle {bundle_id}: {e}")
        return []


def main():
    if not DOMINO_PROJECT_ID:
        print("Error: DOMINO_PROJECT_ID environment variable is not set", file=sys.stderr)
        sys.exit(1)

    bundles = get_bundles(DOMINO_PROJECT_ID)
    print(f"Retrieved {len(bundles)} bundles\n")

    bundle_latest_updates = {}
    bundle_info = {}

    for bundle in bundles:
        bundle_id = bundle["id"]
        bundle_name = bundle["name"]
        policy_id = bundle.get("policyId")

        bundle_info[bundle_id] = {
            "name": bundle_name,
            "policy_id": policy_id
        }

        results = get_results(bundle_id)

        if results:
            parsed_timestamps = [
                parse_timestamp(r["createdAt"])
                for r in results
            ]
            valid_timestamps = [ts for ts in parsed_timestamps if ts is not None]

            if valid_timestamps:
                latest_created_at = max(valid_timestamps)
                bundle_latest_updates[bundle_id] = {
                    "name": bundle_name,
                    "policy_id": policy_id,
                    "latest_update": latest_created_at
                }
                print(f"Bundle: {bundle_name}")
                print(f"  Latest update: {latest_created_at}")
            else:
                print(f"Bundle: {bundle_name}")
                print(f"  No valid timestamps found")
        else:
            print(f"Bundle: {bundle_name}")
            print(f"  No results found")
        print()

    if bundle_latest_updates:
        most_recent_bundle_id = max(
            bundle_latest_updates.items(),
            key=lambda x: x[1]["latest_update"]
        )

        print("=" * 80)
        print("MOST RECENTLY UPDATED BUNDLE:")
        print(f"  ID: {most_recent_bundle_id[0]}")
        print(f"  Name: {most_recent_bundle_id[1]['name']}")
        print(f"  Policy ID: {most_recent_bundle_id[1]['policy_id']}")
        print(f"  Latest Update: {most_recent_bundle_id[1]['latest_update']}")
        print("=" * 80)


if __name__ == "__main__":
    main()
