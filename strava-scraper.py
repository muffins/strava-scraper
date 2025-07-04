#!/usr/bin/env python3

import argparse
import asyncio
import datetime
import json
import os
import requests
import time

from typing import Any, Dict, List, Optional

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError


# On Windows
if os.name == "nt":
    STRAVA_OAUTH_SECRET_FILE = os.path.join(os.environ.get("HOMEPATH"), ".strava_oauth")
    GOOGLE_OAUTH_SECRET_TOKEN_FILE = os.path.join(os.environ.get("HOMEPATH"), ".google_oauth_token")
    GOOGLE_OAUTH_CREDENTIALS_FILE = os.path.join(os.environ.get("HOMEPATH"), ".google_oauth_credentials.json")
else:
    STRAVA_OAUTH_SECRET_FILE = os.path.join(os.environ.get("HOME"), ".strava_oauth")
    GOOGLE_OAUTH_SECRET_TOKEN_FILE = os.path.join(os.environ.get("HOME"), ".google_oauth_token")
    GOOGLE_OAUTH_CREDENTIALS_FILE = os.path.join(os.environ.get("HOME"), ".google_oauth_credentials.json")


STRAVA_CLIENT_ACCESS_TOKEN = os.environ["STRAVA_CLIENT_ACCESS_TOKEN"]
STRAVA_CLIENT_SECRET = os.environ["STRAVA_CLIENT_SECRET"]
STRAVA_CLIENT_ID = os.environ["STRAVA_CLIENT_ID"]

ACCESS_TOKEN_URL = "https://www.strava.com/oauth/token"
ACTIVITIES_URL = "https://www.strava.com/api/v3/athlete/activities"

# Read/Write scopes
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Cut a new Google sheet and paste the ID here
GSHEET_ID = '1zb7GAxkNPRd3dRctgdcPQgvhHv7I3KWWlndIsZry9EQ'
GSHEET_NAME = "Raw Run Data"


def meters_to_miles(meters):
    """Converts meters to miles."""
    return meters * 0.000621371


def seconds_to_hms(seconds):
    """Converts seconds to HH:MM:SS format."""
    return str(datetime.timedelta(seconds=seconds))


async def google_auth() -> Credentials:
    """
    Authenticates with Google and returns credentials.
    Returns:
        Credentials: Authenticated Google credentials.
    """
    creds = None
    if os.path.exists(GOOGLE_OAUTH_SECRET_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(GOOGLE_OAUTH_SECRET_TOKEN_FILE, SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(GOOGLE_OAUTH_CREDENTIALS_FILE):
                print(
                    "Google OAuth credentials file not found. Follow the guide\n" \
                    "at https://developers.google.com/workspace/guides/create-credentials\n" \
                    "to create a credentials file and save it as "
                    f"{GOOGLE_OAUTH_CREDENTIALS_FILE}. Press enter when you have done this."
                )
                input()
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_OAUTH_CREDENTIALS_FILE, SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(GOOGLE_OAUTH_SECRET_TOKEN_FILE, "w") as token:
            token.write(creds.to_json())

    return creds

async def upload_runs_to_gsheets(runs: List[Dict[str, str]]) -> None:
    """
    Uploads run data to Google Sheets.
    Args:
        runs (List[Dict[str, str]]): List of runs with details.
    Returns:
        None
    """

    # First we auth
    creds = await google_auth()
    if not creds:
        print("Failed to authenticate with Google. Please check your credentials.")
        return
    
    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=GSHEET_ID, range=GSHEET_NAME)
            .execute()
        )
        existing_values = result.get("values", [])
    except HttpError as err:
        print(err)

    # De-dupe runs, so we're not uploading the same stuff
    existing_runs = set()
    if not existing_values:
        print("No data found, sheet looks new.")
    
    else:
        # Skip the header row
        for row in existing_values[1:]:
            try:
                existing_runs.add(int(row[0]))
            except (TypeError, IndexError, ValueError) as e:
                # If the row is empty or the first column is not an ID, skip it
                print(f"Skipping row {row[0]} due to TypeError, IndexError, or ValueError - {e}")
                continue

    # Add a header if this is our first run
    if len(existing_runs) == 0:
        print("No existing runs found in Google Sheets, will add all new runs.")
        values = [
            ["ID", "Name", "Date", "Distance Run (Meters)", "Time Elapsed", "Average Cadence", "Elevation Gain", "Average Speed", "Average Heart Rate"]
        ]
    else:
        print(f"Found {len(existing_runs)} existing runs in Google Sheets, will skip these.")
        values = []

    # Now add the new runs
    for run in runs:
        
        if run.get("id") in existing_runs:
            continue

        values.append(
            [
                run.get("id", 0),
                run.get("name", "Unnamed Run"),
                run.get("start_date_local", "N/A"),
                run.get("distance", 0),
                run.get("elapsed_time", "00:00:00"),
                run.get("average_cadence", "N/A"),
                run.get("total_elevation_gain", "N/A"),
                run.get("average_speed", "N/A"),
                run.get("average_heartrate", "N/A"),
            ]
        )

    if not values:
        print("No new runs to upload to Google Sheets, nothing to do :)")
        return
    body = {
        "values": values,
    }
    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=GSHEET_ID,
            range=GSHEET_NAME,
            valueInputOption="RAW",
            body=body,
        ).execute()
        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as err:
        print(f"An error occurred: {err}")
        return

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
                payload["refresh_token"] = oauth_tok_data.get("refresh_token", None)
                payload["grant_type"] = "refresh_token"

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


async def get_strava_activities(access_token, page=1, per_page=30) -> Any:
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


async def main(upload: bool = False) -> None:

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

        # All data is dumped out to a file for debugging or later use, but
        # we ship a subset to Google Sheets.
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

    if upload:
        print("Uploading run data to Google Sheets...")
        await upload_runs_to_gsheets(activities)

    elif all_runs:
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

    ap = argparse.ArgumentParser(
        description="Strava Scraper: Fetch and upload your Strava run data to Google Sheets."
    )
    ap.add_argument(
        "--upload",
        action="store_true",
        help="Upload the fetched run data to Google Sheets, see the README for configuration.",
    )
    args = ap.parse_args()

    asyncio.run(main(args.upload))
