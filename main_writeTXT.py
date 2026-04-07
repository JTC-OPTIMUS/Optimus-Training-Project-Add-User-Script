import os   
import json
import sys
from excel_to_json import excel_to_json
from get_token import get_token
from elevate_token import elevate_token
from send_request import send_request

class Tee:
    """Writes output to both console and a file simultaneously."""
    def __init__(self, filepath):
        self.console = sys.stdout
        self.file = open(filepath, "w", encoding="utf-8")

    def write(self, message):
        self.console.write(message)
        self.file.write(message)

    def flush(self):
        self.console.flush()
        self.file.flush()

    def close(self):
        self.file.close()

# Redirect stdout to Tee (console + file)
tee = Tee("output_log.txt")
sys.stdout = tee

try:
    # Read Excel input, prepare JSON payloads
    try:
        excel_to_json()
    except:
        os._exit(1)

    # Print out access to be granted for each user, for verification
    sys.stdout = tee.console                    # Temporarily restore console only for input()
    cont_input = input("Proceed with request (y/n)? ")
    sys.stdout = tee                            # Resume dual output

    if cont_input.lower() != "y":
        print("Stopped")
        os._exit(0)
    print()

    # Get access token
    try:
        access_token = get_token()
    except Exception as err:
        print("get_token responded with error: " + str(err))
        os._exit(1)
    print()

    # Elevate token permissions
    try:
        access_token = elevate_token()
    except Exception as err:
        print("\nelevate_token responded with error: " + str(err))
        os._exit(1)
    print()

    # Send JSON payloads sequentially
    with open("json_requests.json", "r") as file:
        data = json.load(file)
    for payload in data:
        print(payload["user"]["userName"])
        response = send_request(payload)
        if response.status_code == 200:
            print("Success")
        else:
            print(response.json()["error"]["message"])
        print()

    print("Completed")

finally:
    sys.stdout = tee.console    # Always restore original stdout
    tee.close()                 # Close the log file