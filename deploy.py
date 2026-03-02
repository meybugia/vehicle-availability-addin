#!/usr/bin/env python3
"""
Vehicle Availability Tracker - Deploy Script

Deploys (or removes) the add-in to a client's MyGeotab database.

Usage:
    Deploy:
        python deploy.py --server my.geotab.com --database CLIENT_DB --username admin@example.com --password SECRET

    Remove:
        python deploy.py --server my.geotab.com --database CLIENT_DB --username admin@example.com --password SECRET --remove

    Dry run (show config without deploying):
        python deploy.py --dry-run
"""

import argparse
import getpass
import json
import os
import sys
import urllib.request
import urllib.error

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ADDIN_NAME = "Vehicle Availability Tracker"
ADDIN_VERSION = "1.0.4"


def load_file(filename):
    """Load a file from the dist/ directory."""
    path = os.path.join(SCRIPT_DIR, "dist", filename)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def api_call(server, method, params):
    """Make a MyGeotab API call."""
    payload = json.dumps({"method": method, "params": params}).encode("utf-8")
    req = urllib.request.Request(
        f"https://{server}/apiv1",
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    try:
        resp = json.loads(urllib.request.urlopen(req).read())
    except urllib.error.HTTPError as e:
        print(f"  HTTP Error {e.code}: {e.reason}", file=sys.stderr)
        sys.exit(1)

    if "error" in resp:
        print(f"  API Error: {resp['error']['message']}", file=sys.stderr)
        sys.exit(1)

    return resp.get("result")


def build_addin_config():
    """Build the full add-in config with separate file entries.

    MyGeotab strips inline <style> and <script> tags from embedded HTML,
    so CSS and JS must be provided as separate entries in the files object
    with flat keys (no subdirectory prefixes). The HTML references them
    via <link> and <script src> tags.
    """
    css = load_file(os.path.join("css", "vehicle-availability.css"))
    js = load_file(os.path.join("js", "vehicle-availability.js"))

    html = (
        '<!DOCTYPE html>\n'
        '<html lang="en">\n'
        '<head>\n'
        '    <meta charset="UTF-8">\n'
        '    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        '    <title>Vehicle Availability Tracker</title>\n'
        '    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" crossorigin="" />\n'
        '    <link rel="stylesheet" href="vehicle-availability.css" />\n'
        '</head>\n'
        '<body>\n'
        '    <div id="vehicle-availability-addin">\n'
        '        <div class="va-header">\n'
        '            <h2>Vehicle Availability</h2>\n'
        '            <div class="va-controls">\n'
        '                <span id="va-last-updated" class="va-last-updated"></span>\n'
        '                <button id="va-refresh-btn" class="va-refresh-btn">Refresh</button>\n'
        '            </div>\n'
        '        </div>\n'
        '        <div id="va-loading" class="va-loading">\n'
        '            <div class="va-spinner"></div>\n'
        '            Loading vehicle data...\n'
        '        </div>\n'
        '        <div class="va-summary">\n'
        '            <div class="va-card available">\n'
        '                <div id="va-count-available" class="va-card-count">-</div>\n'
        '                <div class="va-card-label">Available</div>\n'
        '            </div>\n'
        '            <div class="va-card dispatched">\n'
        '                <div id="va-count-dispatched" class="va-card-count">-</div>\n'
        '                <div class="va-card-label">Dispatched</div>\n'
        '            </div>\n'
        '            <div class="va-card total">\n'
        '                <div id="va-count-total" class="va-card-count">-</div>\n'
        '                <div class="va-card-label">Total Vehicles</div>\n'
        '            </div>\n'
        '        </div>\n'
        '        <div class="va-map-container">\n'
        '            <div id="va-map"></div>\n'
        '        </div>\n'
        '        <div class="va-filter-bar">\n'
        '            <label for="va-filter-status">Filter:</label>\n'
        '            <select id="va-filter-status" class="va-filter-select">\n'
        '                <option value="all">All</option>\n'
        '                <option value="available">Available</option>\n'
        '                <option value="dispatched">Dispatched</option>\n'
        '            </select>\n'
        '            <input id="va-search" type="text" class="va-search-input" placeholder="Search by name, serial, zone..." />\n'
        '        </div>\n'
        '        <div class="va-table-container">\n'
        '            <table class="va-table">\n'
        '                <thead>\n'
        '                    <tr>\n'
        '                        <th data-sort="name">Vehicle Name <span class="sort-arrow active">&#9650;</span></th>\n'
        '                        <th data-sort="serialNumber">Serial Number <span class="sort-arrow">&#9650;</span></th>\n'
        '                        <th data-sort="status">Status <span class="sort-arrow">&#9650;</span></th>\n'
        '                        <th data-sort="location">Location <span class="sort-arrow">&#9650;</span></th>\n'
        '                        <th data-sort="currentZones">Current Zone <span class="sort-arrow">&#9650;</span></th>\n'
        '                        <th data-sort="lastUpdated">Last Updated <span class="sort-arrow">&#9650;</span></th>\n'
        '                    </tr>\n'
        '                </thead>\n'
        '                <tbody id="va-table-body">\n'
        '                    <tr><td colspan="6" class="va-empty">Loading...</td></tr>\n'
        '                </tbody>\n'
        '            </table>\n'
        '        </div>\n'
        '    </div>\n'
        '    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js" crossorigin=""></script>\n'
        '    <script src="vehicle-availability.js"></script>\n'
        '</body>\n'
        '</html>'
    )

    return {
        "name": ADDIN_NAME,
        "supportEmail": "support@example.com",
        "version": ADDIN_VERSION,
        "items": [
            {
                "url": "vehicle-availability.html",
                "path": "ActivityLink/",
                "menuName": {"en": "Vehicle Availability"},
                "icon": "https://unpkg.com/heroicons@2.0.18/24/outline/truck.svg",
            }
        ],
        "files": {
            "vehicle-availability.html": html,
            "vehicle-availability.css": css,
            "vehicle-availability.js": js,
        },
    }


def deploy(server, database, username, password):
    """Deploy the add-in to a MyGeotab database."""
    print(f"\n  Authenticating to {server}/{database}...")
    result = api_call(server, "Authenticate", {
        "database": database,
        "userName": username,
        "password": password,
    })
    creds = result["credentials"]
    path = result.get("path", "")

    # Use the correct server if redirected
    if path and path != "ThisServer":
        server = path
        print(f"  Redirected to server: {server}")

    print(f"  Authenticated as {creds['userName']}")

    api_creds = {
        "database": creds["database"],
        "userName": creds["userName"],
        "sessionId": creds["sessionId"],
    }

    print("  Fetching current SystemSettings...")
    settings_list = api_call(server, "Get", {
        "typeName": "SystemSettings",
        "credentials": api_creds,
    })
    ss = settings_list[0]
    addins = ss.get("customerPages", [])

    # Remove existing version if present
    new_addins = []
    replaced = False
    for a in addins:
        content = a if isinstance(a, str) else json.dumps(a)
        if ADDIN_NAME in content:
            replaced = True
        else:
            new_addins.append(a)

    # Add new version
    addin_config = build_addin_config()
    new_addins.append(json.dumps(addin_config))
    ss["customerPages"] = new_addins

    action = "Updating" if replaced else "Installing"
    print(f"  {action} '{ADDIN_NAME}' v{ADDIN_VERSION}...")

    api_call(server, "Set", {
        "typeName": "SystemSettings",
        "entity": ss,
        "credentials": api_creds,
    })

    print(f"\n  SUCCESS! '{ADDIN_NAME}' deployed to {database}")
    print(f"  Refresh MyGeotab and go to: Activity > Vehicle Availability\n")


def remove(server, database, username, password):
    """Remove the add-in from a MyGeotab database."""
    print(f"\n  Authenticating to {server}/{database}...")
    result = api_call(server, "Authenticate", {
        "database": database,
        "userName": username,
        "password": password,
    })
    creds = result["credentials"]
    path = result.get("path", "")

    if path and path != "ThisServer":
        server = path

    api_creds = {
        "database": creds["database"],
        "userName": creds["userName"],
        "sessionId": creds["sessionId"],
    }

    print("  Fetching current SystemSettings...")
    settings_list = api_call(server, "Get", {
        "typeName": "SystemSettings",
        "credentials": api_creds,
    })
    ss = settings_list[0]
    addins = ss.get("customerPages", [])

    new_addins = []
    found = False
    for a in addins:
        content = a if isinstance(a, str) else json.dumps(a)
        if ADDIN_NAME in content:
            found = True
        else:
            new_addins.append(a)

    if not found:
        print(f"\n  '{ADDIN_NAME}' not found in {database}. Nothing to remove.\n")
        return

    ss["customerPages"] = new_addins

    print(f"  Removing '{ADDIN_NAME}'...")
    api_call(server, "Set", {
        "typeName": "SystemSettings",
        "entity": ss,
        "credentials": api_creds,
    })

    print(f"\n  SUCCESS! '{ADDIN_NAME}' removed from {database}\n")


def main():
    parser = argparse.ArgumentParser(
        description="Deploy Vehicle Availability Tracker to a MyGeotab database"
    )
    parser.add_argument("--server", default="my.geotab.com",
                        help="MyGeotab server (default: my.geotab.com)")
    parser.add_argument("--database", "-d", help="MyGeotab database name")
    parser.add_argument("--username", "-u", help="MyGeotab admin username")
    parser.add_argument("--password", "-p", help="MyGeotab password (prompts if omitted)")
    parser.add_argument("--remove", action="store_true",
                        help="Remove the add-in instead of deploying")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the config JSON without deploying")

    args = parser.parse_args()

    if args.dry_run:
        config = build_addin_config()
        # Print without embedded files for readability
        display = {k: v for k, v in config.items() if k != "files"}
        display["files"] = {k: f"<{len(v)} chars>" for k, v in config["files"].items()}
        print(json.dumps(display, indent=2))
        return

    if not args.database:
        args.database = input("Database name: ").strip()
    if not args.username:
        args.username = input("Username: ").strip()
    if not args.password:
        args.password = getpass.getpass("Password: ")

    if args.remove:
        remove(args.server, args.database, args.username, args.password)
    else:
        deploy(args.server, args.database, args.username, args.password)


if __name__ == "__main__":
    main()
