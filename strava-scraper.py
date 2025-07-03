#!/usr/bin/env python3

import asyncio
import datetime
import json
import os
import requests
import time

from typing import Dict, List, Optional


# On Windows
if os.name == "nt":
    STRAVA_OAUTH_SECRET_FILE = os.path.join(os.environ.get("HOMEPATH"), ".strava_oauth")
else:
    STRAVA_OAUTH_SECRET_FILE = os.path.join(os.environ.get("HOME"), ".strava_oauth")

STRAVA_CLIENT_ACCESS_TOKEN = os.environ["STRAVA_CLIENT_ACCESS_TOKEN"]
STRAVA_CLIENT_SECRET = os.environ["STRAVA_CLIENT_SECRET"]
STRAVA_CLIENT_ID = os.environ["STRAVA_CLIENT_ID"]

ACCESS_TOKEN_URL = "https://www.strava.com/oauth/token"
ACTIVITIES_URL = "https://www.strava.com/api/v3/athlete/activities"


def meters_to_miles(meters):
    """Converts meters to miles."""
    return meters * 0.000621371


def seconds_to_hms(seconds):
    """Converts seconds to HH:MM:SS format."""
    return str(datetime.timedelta(seconds=seconds))


async def get_strava_access_token() -> Optional[str]:
    """
    Completes the Strava OAuth flow to get an access token.
    Args:
        client_id (str): Your Strava client ID.
        client_secret (str): Your Strava client secret.
        code (str): The authorization code received from Strava.
    Returns:
        str: Access token if successful, None otherwise.
    """

    oauth_tok_data = None
    payload = {
        "client_id": STRAVA_CLIENT_ID,
        "client_secret": STRAVA_CLIENT_SECRET,
        "grant_type": "authorization_code",
    }
    if os.path.exists(STRAVA_OAUTH_SECRET_FILE):
        # Try to load it from disk
        with open(STRAVA_OAUTH_SECRET_FILE, "r") as fin:
            oauth_tok_data = json.loads(fin.read())
            if datetime.datetime.now().timestamp() < oauth_tok_data.get(
                "expires_at", None
            ):
                print(
                    "Using existing Strava access token, it's valid until "
                    f"{datetime.datetime.fromtimestamp(oauth_tok_data.get('expires_at', 0)).strftime('%Y-%m-%d %H:%M:%S')}"
                )
                return oauth_tok_data.get("access_token", None)
            else:
                print("Existing Strava access token has expired. Re-authenticating...")
                payload["refresh_token"] = oauth_tok_data.get("access_token", None)

    else:
        # If we don't already have a valid token, we need to re-oauth and get a
        # new one.
        print("No pre-existing Strava access token found, starting OAuth flow.")
        print("Please visit the following URL to authorize the application:")
        print(
            f"http://www.strava.com/oauth/authorize?client_id={STRAVA_CLIENT_ID}&redirect_uri=https://localhost&response_type=code&approval_prompt=auto&scope=activity:read_all"
        )
        strava_oath_code = input(
            "\nAfter authorizing, paste the code from the URL here: "
        )
        payload["code"] = strava_oath_code

    data = None
    try:
        response = requests.post(ACCESS_TOKEN_URL, data=payload)
        response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)
        data = response.json()
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
        print(f"Response content: {response.text}")
        return None
    except requests.exceptions.RequestException as req_err:
        print(f"An error occurred: {req_err}")
        return None

    if not data:
        print("No data received from Strava.")
        return None

    # Write out the oauth token to disk for later use
    with open(STRAVA_OAUTH_SECRET_FILE, "w") as fout:
        fout.write(json.dumps(data))
        print(
            f"Saved Strava OAuth token to {os.path.join(STRAVA_OAUTH_SECRET_FILE, 'strava_oauth.json')}"
        )

    print("Successfully fetched Strava access token! ")
    print(
        f"It's valid until {datetime.datetime.fromtimestamp(data.get('expires_in', 0)).strftime('%Y-%m-%d %H:%M:%S')}."
    )
    return data.get("access_token", None)


async def get_strava_activities(access_token, page=1, per_page=30):
    """
    Fetches activities from the Strava API.
    Args:
        access_token (str): Your Strava access token.
        page (int): The page number of results to fetch.
        per_page (int): The number of activities per page (max 200).
    Returns:
        list: A list of activity dictionaries.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"page": page, "per_page": per_page}
    try:
        response = requests.get(ACTIVITIES_URL, headers=headers, params=params)
        response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
        print(f"Response content: {response.text}")
        return []
    except requests.exceptions.ConnectionError as conn_err:
        print(f"Connection error occurred: {conn_err}")
        return []
    except requests.exceptions.Timeout as timeout_err:
        print(f"Timeout error occurred: {timeout_err}")
        return []
    except requests.exceptions.RequestException as req_err:
        print(f"An error occurred: {req_err}")
        return []


async def main() -> None:

    tok = await get_strava_access_token()
    if not tok:
        print(
            "Failed to obtain Strava access token. Please check your credentials and try again."
        )
        return

    print("Fetching your Strava run data...")

    all_runs = []
    page = 1
    while True:
        # Fetch activities from Strava API
        activities = await get_strava_activities(tok, page=page, per_page=100)

        if not activities:
            # No more activities or an error occurred
            break

        if not os.path.exists("out"):
            os.makedirs("out")

        # Filter for 'Run' activities and extract relevant data
        with open(
            os.path.join("out", f"strava_activities_{int(time.time())}.json"), "w"
        ) as fout:
            for activity in activities:
                fout.write(json.dumps(activity) + "\n")

        # Verbosity
        for activity in activities:
            if activity.get("type") == "Run":
                distance_meters = activity.get("distance", 0)
                elapsed_time_seconds = activity.get("elapsed_time", 0)
                name = activity.get("name", "Unnamed Run")
                start_date = activity.get("start_date_local", "N/A")

                miles_run = meters_to_miles(distance_meters)
                time_elapsed_hms = seconds_to_hms(elapsed_time_seconds)

                all_runs.append(
                    {
                        "name": name,
                        "date": start_date,
                        "miles_run": f"{miles_run:.2f}",  # Format to 2 decimal places
                        "time_elapsed": time_elapsed_hms,
                        "average_cadence": activity.get("average_cadence", "N/A"),
                    }
                )

        # If the number of activities fetched is less than per_page, it means we've reached the end
        if len(activities) < 100:
            break

        page += 1  # Move to the next page

    if all_runs:

        # TODO: Dump this to CSV and/or Google Sheets.
        print("\n--- Your Strava Runs ---")
        for run in all_runs:
            print(f"Run Name: {run['name']}")
            print(f"Date: {run['date']}")
            print(f"Miles Run: {run['miles_run']}")
            print(f"Time Elapsed: {run['time_elapsed']}")
            print(f"Average Cadence: {run['average_cadence']}")
            print("-" * 30)

    else:
        print(
            "No run activities found or unable to fetch data. Please check your token and internet connection."
        )


if __name__ == "__main__":

    asyncio.run(main())
