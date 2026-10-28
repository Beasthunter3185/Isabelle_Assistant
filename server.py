from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import msal
import os
import datetime
import calendar
import json
import requests
import uvicorn
import logging
import ollama

# --- Robust Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("IsabelleServer")

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

CACHE_FILE = "token_cache.bin"
CLIENT_ID = "d6abb8a2-3a3c-436d-90ca-1e10cd99b3d9"
TENANT_ID = "consumers" 
SCOPES = ["Calendars.ReadWrite"]

def get_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        logger.info(f"Loading token cache from {CACHE_FILE}")
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    else:
        logger.warning(f"Token cache file '{CACHE_FILE}' not found.")
    return cache

def save_token_cache(cache):
    if cache.has_state_changed:
        logger.info("Token state changed. Saving cache to disk.")
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def check_authenticated() -> bool:
    """Non-blocking check to verify if valid/refreshable credentials exist."""
    cache = get_token_cache()
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app_msal = msal.PublicClientApplication(CLIENT_ID, authority=authority, token_cache=cache)
    accounts = app_msal.get_accounts()
    if accounts:
        result = app_msal.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_token_cache(cache)
            logger.info("Authentication check: SUCCESS (Valid token found/refreshed).")
            return True
    logger.warning("Authentication check: FAILED (No valid accounts or tokens).")
    return False

def get_access_token():
    """Retrieves access token silently; returns None if re-authentication is required."""
    cache = get_token_cache()
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app_msal = msal.PublicClientApplication(CLIENT_ID, authority=authority, token_cache=cache)
    
    accounts = app_msal.get_accounts()
    if accounts:
        result = app_msal.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_token_cache(cache)
            return result["access_token"]
    logger.error("Failed to retrieve silent access token.")
    return None

# --- AUTH ENDPOINTS ---

@app.get("/auth/status")
async def get_auth_status():
    logger.info("GET /auth/status called by client.")
    is_auth = check_authenticated()
    return {"authenticated": is_auth}

@app.post("/auth/login")
async def trigger_auth_login():
    logger.info("POST /auth/login triggered.")
    if check_authenticated():
        return {"status": "already_authenticated", "message": "User is already logged in."}

    cache = get_token_cache()
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app_msal = msal.PublicClientApplication(CLIENT_ID, authority=authority, token_cache=cache)
    
    # Opens host OS browser for interactive OAuth2 flow
    result = app_msal.acquire_token_interactive(scopes=SCOPES)
    if "access_token" in result:
        save_token_cache(cache)
        logger.info("Interactive login successful.")
        return {"status": "success", "message": "Successfully authenticated with Microsoft."}
    
    logger.error("Interactive authentication failed or was canceled.")
    raise HTTPException(status_code=400, detail="Authentication failed or was canceled.")

class UserQuery(BaseModel):
    message: str

conversation_history = []
MAX_HISTORY_MESSAGES = 10

# Isabelle's standard help manual & opening message
COMMANDS_GUIDE = """Hello Mayor! Let me know if you need any help managing your island schedule! You can type !help anytime to see this guide.

Here are the commands you can use:

1️⃣ CREATE EVENT (!create)
• Fields: Title in quotes, Date, Time (or All-Day)
• Example: !create "Haircut" tomorrow at 2:00 PM

2️⃣ DELETE EVENT (!delete)
• Fields: Date OR Title in quotes OR Number choice
• Example: !delete tomorrow or !delete "Haircut"

3️⃣ VIEW SCHEDULE (!view)
• Fields: Optional day, week, or month
• Example: !view or !view this week"""

def get_tz_offset_string():
    """Generates dynamic ISO timezone offset (e.g., '-04:00' or '-05:00')."""
    tz = datetime.datetime.now().astimezone().strftime("%z")
    return f"{tz[:3]}:{tz[3:]}"

def sanitize_iso_for_graph(iso_str: str) -> str:
    """Strips timezone offsets and 'Z' suffix so MS Graph receives a clean local ISO string."""
    if not iso_str:
        return iso_str
    iso_str = iso_str.replace("Z", "")
    if len(iso_str) > 19 and ("+" in iso_str[10:] or "-" in iso_str[10:]):
        iso_str = iso_str[:19]
    return iso_str

# --- Microsoft Graph API Functions ---

def check_calendar_conflict(start_iso: str, end_iso: str):
    token = get_access_token()
    if not token or not start_iso or not end_iso:
        return False, []

    tz_offset = get_tz_offset_string()
    query_start = start_iso if ("+" in start_iso or "-" in start_iso[10:]) else f"{start_iso}{tz_offset}"
    query_end = end_iso if ("+" in end_iso or "-" in end_iso[10:]) else f"{end_iso}{tz_offset}"

    url = "https://graph.microsoft.com/v1.0/me/calendar/calendarView"
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.timezone="Eastern Standard Time"'
    }
    params = {"startDateTime": query_start, "endDateTime": query_end}

    logger.info(f"Checking calendar conflicts between {query_start} and {query_end}")
    response = requests.get(url, headers=headers, params=params)
    if response.status_code == 200:
        events = response.json().get("value", [])
        return len(events) > 0, events

    logger.error(f"Failed to check calendar conflict. Status: {response.status_code}")
    return False, []

def create_outlook_event(event_data: dict):
    token = get_access_token()
    if not token:
        return False

    url = "https://graph.microsoft.com/v1.0/me/events"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Prefer": 'outlook.timezone="Eastern Standard Time"'
    }

    start_clean = sanitize_iso_for_graph(event_data.get("start_iso"))
    end_clean = sanitize_iso_for_graph(event_data.get("end_iso"))

    payload = {
        "subject": event_data.get("subject", "New Event"),
        "isAllDay": event_data.get("is_all_day", False),
        "start": {"dateTime": start_clean, "timeZone": "Eastern Standard Time"},
        "end": {"dateTime": end_clean, "timeZone": "Eastern Standard Time"}
    }

    logger.info(f"Creating Outlook event: '{payload['subject']}'")
    response = requests.post(url, headers=headers, json=payload)
    success = response.status_code == 201
    if success:
        logger.info("Event successfully created.")
    else:
        logger.error(f"Failed to create event. Response: {response.text}")
    return success

def fetch_events_in_range(start_iso: str, end_iso: str):
    token = get_access_token()
    if not token:
        logger.warning("Skipping event fetch: No access token available.")
        return []

    tz_offset = get_tz_offset_string()
    
    if len(start_iso) == 10:
        start_iso = f"{start_iso}T00:00:00"
    if len(end_iso) == 10:
        end_iso = f"{end_iso}T23:59:59"

    start_iso = start_iso.replace("Z", "")
    end_iso = end_iso.replace("Z", "")

    query_start = start_iso if ("+" in start_iso or "-" in start_iso[10:]) else f"{start_iso}{tz_offset}"
    query_end = end_iso if ("+" in end_iso or "-" in end_iso[10:]) else f"{end_iso}{tz_offset}"

    url = "https://graph.microsoft.com/v1.0/me/calendar/calendarView"
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": 'outlook.timezone="Eastern Standard Time"'
    }
    # Request up to 100 per page to reduce round trips
    params = {
        "startDateTime": query_start, 
        "endDateTime": query_end,
        "$top": 100,
        "$select": "subject,start,end,type,isAllDay"
    }

    all_events = []
    logger.info(f"Fetching calendar events from Graph API between {query_start} and {query_end}")

    while url:
        response = requests.get(url, headers=headers, params=params if "?" not in url else None)
        if response.status_code == 200:
            data = response.json()
            events = data.get("value", [])
            all_events.extend(events)
            
            # Check if there's another page of events
            url = data.get("@odata.nextLink")
            params = None # Parameters are already embedded in the nextLink URL
        else:
            logger.error(f"Graph API error {response.status_code}: {response.text}")
            break

    logger.info(f"Successfully fetched a total of {len(all_events)} events.")
    return all_events

def delete_event_by_id(event_id: str):
    token = get_access_token()
    if not token:
        return False
    url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"
    headers = {"Authorization": f"Bearer {token}"}
    logger.info(f"Deleting event ID: {event_id}")
    response = requests.delete(url, headers=headers)
    return response.status_code == 204

def smart_delete_event(subject_query: str = "", start_iso: str = None, end_iso: str = None, target_index: int = None):
    now = datetime.datetime.now()
    
    if not start_iso or not end_iso:
        _, last_day = calendar.monthrange(now.year, now.month)
        start_iso = now.strftime("%Y-%m-01T00:00:00")
        end_iso = f"{now.strftime('%Y-%m')}-{last_day:02d}T23:59:59"

    events = fetch_events_in_range(start_iso, end_iso)

    if subject_query and subject_query.strip().lower() not in ["none", "null", "it", "that", "event", "meeting", "appointment", ""]:
        clean_query = subject_query.replace('"', '').replace("'", '').lower().strip()
        events = [e for e in events if clean_query in e.get("subject", "").lower()]

    if not events:
        if subject_query:
            return f"Mayor, I couldn't find any event matching '{subject_query}'!"
        return "Mayor, I couldn't find any events scheduled during that timeframe."

    if target_index is not None and target_index > 0:
        if target_index <= len(events):
            target_event = events[target_index - 1]
            if delete_event_by_id(target_event["id"]):
                return f"I've successfully removed '{target_event['subject']}' from your calendar!"
            return "I ran into a problem while trying to delete that event."
        else:
            return f"I couldn't find option #{target_index}. Please pick a number between 1 and {len(events)}."

    if len(events) == 1:
        target_event = events[0]
        if delete_event_by_id(target_event["id"]):
            return f"I've removed '{target_event['subject']}' from your calendar!"
        return "I ran into a problem while trying to delete that event."

    event_lines = []
    for idx, e in enumerate(events, 1):
        time_part = "ALL DAY" if e.get("isAllDay") else e['start']['dateTime'][11:16]
        date_part = e['start']['dateTime'][:10]
        event_lines.append(f"{idx}) '{e['subject']}' on {date_part} ({time_part})")

    list_str = "\n".join(event_lines)
    return (
        f"Mayor, I found multiple events on that date:\n\n{list_str}\n\n"
        f"Which numbered event would you like me to remove?"
    )

def read_outlook_events(start_iso: str = None, end_iso: str = None):
    token = get_access_token()
    if not token:
        return "I couldn't connect to your calendar right now."

    now = datetime.datetime.now()
    if not start_iso or not end_iso:
        _, last_day = calendar.monthrange(now.year, now.month)
        start_iso = now.strftime("%Y-%m-01T00:00:00")
        end_iso = f"{now.strftime('%Y-%m')}-{last_day:02d}T23:59:59"

    events = fetch_events_in_range(start_iso, end_iso)
    if not events:
        return "Your calendar is completely clear for that period!"
    
    event_list = []
    for e in events:
        start_str = e['start']['dateTime']
        date_part = start_str[:10]
        time_part = "ALL DAY" if e.get("isAllDay") else start_str[11:16]
        event_list.append(f"- {date_part} ({time_part}): '{e['subject']}'")

    return "Here is what you have scheduled:\n" + "\n".join(event_list)

# --- LLM Parsing ---

def parse_with_ollama(user_message: str):
    now = datetime.datetime.now()
    today_str = now.strftime("%A, %B %d, %Y at %I:%M %p")

    system_prompt = """
You are Isabelle from Animal Crossing acting as a calendar assistant.
Current local date and time: {today_str}.

TASK:
Extract the user's intent and details into structured JSON.

RULES:
1. "CREATE": Use when the user asks to add, schedule, or create an event/appointment.
   - Extract the event name as "subject". If missing, set "missing_field": "subject".
   - Convert relative dates into local naive ISO 8601 strings for "start_iso" and "end_iso" (e.g. "2026-07-28T06:00:00"). DO NOT include timezone offsets (+04:00) or 'Z'. Assume 1-hour duration unless specified.
2. "DELETE": Use when removing or cancelling an event.
3. "READ": Use when checking or viewing the schedule.
4. "HELP": Use if the user asks for commands or help.
5. "NEEDS_INFO": Request is vague or missing key details.

Return STRICT JSON format ONLY:
{{
    "actions": [
        {{
            "intent": "CREATE" | "DELETE" | "READ" | "HELP" | "NEEDS_INFO" | "OTHER",
            "is_all_day": false,
            "subject": "",
            "target_index": null,
            "start_iso": null,
            "end_iso": null,
            "missing_field": null,
            "isabelle_response": ""
        }}
    ]
}}
""".format(today_str=today_str)

    messages_payload = [{"role": "system", "content": system_prompt}]
    messages_payload.extend(conversation_history)
    messages_payload.append({"role": "user", "content": user_message})

    logger.info(f"Sending prompt to Ollama for message: '{user_message}'")
    response = ollama.chat(model="llama3.1", messages=messages_payload, format="json")
    return json.loads(response["message"]["content"])

# --- Main Route ---
@app.get("/v1/calendar/events")
async def get_godot_calendar_events(startDateTime: str, endDateTime: str):
    logger.info(f"GET /v1/calendar/events requested with range: {startDateTime} to {endDateTime}")
    events = fetch_events_in_range(startDateTime, endDateTime)
    return {"value": events}

@app.post("/chat")
async def chat_endpoint(query: UserQuery):
    msg = query.message.strip().lower()
    logger.info(f"POST /chat received message: '{query.message}'")

    # ⚡ DIRECT OVERRIDE: Return commands guide immediately!
    if msg in ["!welcome", "!help", "help", "!commands"]:
        return {"reply": COMMANDS_GUIDE}

    # Otherwise, pass through Ollama for AI parsing
    parsed = parse_with_ollama(query.message)
    actions = parsed.get("actions", [])
    combined_responses = []

    for action in actions:
        intent = action.get("intent")

        if intent == "HELP":
            combined_responses.append(COMMANDS_GUIDE)

        elif intent == "NEEDS_INFO":
            missing = action.get("missing_field")
            if missing == "subject":
                combined_responses.append("Mayor, what specific title in quotes (e.g., \"Haircut\") should I give this event?")
            elif missing == "time":
                combined_responses.append("Mayor, what time should I schedule that for? (e.g., \"at 2:00 PM\")")
            elif missing == "date_or_title":
                combined_responses.append("Mayor, which day or title should I look up to remove? (e.g., \"tomorrow\" or \"Haircut\")")
            else:
                combined_responses.append(action.get("isabelle_response", "Could you specify a few more details?"))

        elif intent == "CREATE":
            start_iso = action.get("start_iso")
            end_iso = action.get("end_iso")
            is_all_day = action.get("is_all_day", False)
            subject = action.get("subject", "").strip()

            if not subject or subject.lower() in ["none", "null", "new event", "event", "something"]:
                combined_responses.append("Mayor, what specific title in quotation marks (e.g., \"Haircut\") should I give this event?")
                continue

            action["subject"] = subject.title()
            isabelle_reply = action.get("isabelle_response") or f"I've added '{action['subject']}' to your calendar!"

            has_conflict, conflicting_events = check_calendar_conflict(start_iso, end_iso)
            if has_conflict:
                conflict_title = conflicting_events[0]["subject"]
                combined_responses.append(
                    f"Oh! You already have '{conflict_title}' scheduled during that time slot. "
                    f"Would you like me to schedule '{action['subject']}' anyway, or pick a different time?"
                )
            else:
                created = create_outlook_event(action)
                combined_responses.append(isabelle_reply if created else "I ran into a hiccup adding that event!")

        elif intent == "DELETE":
            subject = action.get("subject", "")
            start_iso = action.get("start_iso")
            end_iso = action.get("end_iso")
            target_index = action.get("target_index")

            delete_msg = smart_delete_event(
                subject_query=subject,
                start_iso=start_iso,
                end_iso=end_iso,
                target_index=target_index
            )
            combined_responses.append(delete_msg)

        elif intent == "READ":
            start_iso = action.get("start_iso")
            end_iso = action.get("end_iso")
            schedule_msg = read_outlook_events(start_iso, end_iso)
            combined_responses.append(f"Mayor, here is your schedule overview!\n{schedule_msg}")

        else:
            combined_responses.append(action.get("isabelle_response", "I'm always here to help with your island tasks!"))

    final_reply = "\n\n".join(combined_responses)

    conversation_history.append({"role": "user", "content": query.message})
    conversation_history.append({"role": "assistant", "content": final_reply})

    if len(conversation_history) > MAX_HISTORY_MESSAGES:
        conversation_history[:] = conversation_history[-MAX_HISTORY_MESSAGES:]

    return {"reply": final_reply}

if __name__ == "__main__":
    logger.info("--- Starting Isabelle Backend Server ---")
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=True)