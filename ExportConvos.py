import requests
import csv
import os
import json
from datetime import datetime

# User-configurable variables
AUTH_TOKEN = "[INSERT VF API KEY]"
PROJECT_ID = "[INSERT VF PROJECT ID]"
START_DATE = "2024-07-14"  # Format: YYYY-MM-DD
#START_DATE = "2024-08-06"  # Format: YYYY-MM-DD
END_DATE = "2024-08-07"    # Format: YYYY-MM-DD
OUTPUT_DIRECTORY = "2024-07-14 to 2024-08-07"  # Name of the directory to create
#OUTPUT_DIRECTORY = "2024-08-06 to 2024-08-07"  # Name of the directory to create

# Constants
BASE_URL = "https://api.voiceflow.com/v2/transcripts"
HEADERS = {
    "Authorization": AUTH_TOKEN,
    "accept": "application/json"
}

def create_output_directory():
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.makedirs(OUTPUT_DIRECTORY)
        print(f"Created directory: {OUTPUT_DIRECTORY}")
    else:
        print(f"Directory already exists: {OUTPUT_DIRECTORY}")

def get_transcript_ids():
    url = f"{BASE_URL}/{PROJECT_ID}"
    params = {
        "startDate": START_DATE,
        "endDate": END_DATE
    }
    response = requests.get(url, headers=HEADERS, params=params)
    response.raise_for_status()
    return [transcript["_id"] for transcript in response.json()]

def get_transcript_dialog(transcript_id):
    url = f"{BASE_URL}/{PROJECT_ID}/{transcript_id}"
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    return response.json()

def extract_messages(dialog):
    messages = []
    for turn in dialog:
        if turn["type"] == "request":
            if "payload" in turn and "query" in turn["payload"].get("payload", {}):
                messages.append({
                    "role": "HUMAN",
                    "content": turn["payload"]["payload"]["query"],
                    "timestamp": turn.get("startTime", "")
                })
        elif turn["type"] == "text" and "payload" in turn and "message" in turn["payload"].get("payload", {}):
            messages.append({
                "role": "BOT",
                "content": turn["payload"]["payload"]["message"],
                "timestamp": turn.get("startTime", "")
            })
        elif turn["type"] == "debug" and "payload" in turn and "payload" in turn["payload"]:
            debug_payload = turn["payload"]["payload"]
            if debug_payload.get("type") == "code" and "{Tags}" in debug_payload.get("message", ""):
                messages.append({
                    "role": "DEBUG",
                    "content": debug_payload["message"],
                    "timestamp": turn.get("startTime", "")
                })
    return messages

def save_transcript_to_txt(transcript_id, messages):
    filename = os.path.join(OUTPUT_DIRECTORY, f"transcript_{transcript_id}.txt")
    with open(filename, 'w', encoding='utf-8') as file:
        for message in messages:
            if message['role'] == 'DEBUG':
                file.write(f"DEBUG (Tags): {message['content']}\n")
            else:
                file.write(f"{message['role']}: {message['content']}\n")
            file.write("----------\n")  # Add separator line
    print(f"Saved transcript to {filename}")

def export_to_csv(all_messages):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.join(OUTPUT_DIRECTORY, f"voiceflow_transcripts_{START_DATE}_to_{END_DATE}_{timestamp}.csv")
    
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Timestamp', 'Role', 'Message']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for message in all_messages:
            writer.writerow({
                'Timestamp': message['timestamp'],
                'Role': message['role'],
                'Message': message['content']
            })
    
    print(f"Exported all transcripts to {filename}")

def main():
    create_output_directory()
    transcript_ids = get_transcript_ids()
    all_messages = []
    
    for transcript_id in transcript_ids:
        dialog = get_transcript_dialog(transcript_id)
        messages = extract_messages(dialog)
        save_transcript_to_txt(transcript_id, messages)
        all_messages.extend(messages)
    
    export_to_csv(all_messages)

if __name__ == "__main__":
    main()