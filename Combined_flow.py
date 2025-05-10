from msal import PublicClientApplication, SerializableTokenCache
import requests
import re
import os
from typing import List, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
from google import genai

# === CONFIGURATION ===
CLIENT_ID = '8ff06bc0-610e-4f85-b58c-01071a7b39c6'
AUTHORITY = 'https://login.microsoftonline.com/consumers'
SCOPES = ['https://graph.microsoft.com/Mail.Read']
CACHE_FILE = 'token_cache.bin'
API_KEY = 'MY API KEY'  # Replace with your real Gemini API key
genai_client = genai.Client(api_key=API_KEY)

# === SIGNATURE REMOVAL FUNCTION ===
def remove_signature(text: str) -> str:
    signature_cues = [
        r"\bBest regards\b", r"\bRegards\b", r"\bThanks\b", r"\bSincerely\b",
        r"\bWarm regards\b", r"\bKind regards\b", r"\bCheers\b", r"Sent from my iPhone",
        r"--", r"—", r"___", r"^Sent on behalf of", r"^From:"
    ]
    pattern = re.compile("|".join(signature_cues), re.IGNORECASE)
    match = pattern.search(text)
    if match:
        return text[:match.start()].strip()
    return text.strip()

# === FETCH EMAILS FROM OUTLOOK ===
def fetch_emails_from_outlook(max_emails=5) -> List[str]:
    cache = SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    app = PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if 'message' not in flow:
            raise Exception("Failed to initiate device flow. Details:", flow)
        print(flow['message'])
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    if 'access_token' not in result:
        raise Exception("Token error: " + result.get("error_description", "No token"))

    headers = {
        'Authorization': f"Bearer {result['access_token']}",
        'Prefer': 'outlook.body-content-type="text"'
    }

    url = f'https://graph.microsoft.com/v1.0/me/messages?$top={max_emails}&$select=subject,from,body'
    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        raise Exception("Error fetching emails: " + response.text)

    emails_data = response.json().get('value', [])
    cleaned_emails = []

    for mail in emails_data:
        raw_body = mail.get('body', {}).get('content', '')

        # Clean unwanted content
        cleaned = re.sub(r'http\S+|www\.\S+', '', raw_body)  # Remove URLs
        cleaned = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', '', cleaned)  # Remove emails
        cleaned = re.sub(r'(?i)from:.*|sent by.*', '', cleaned)  # Remove common sender lines
        cleaned = remove_signature(cleaned)  # Remove signatures

        cleaned_emails.append(cleaned)

    return cleaned_emails

# === GEMINI CLASSIFICATION ===
def classify_email(index: int, email_text: str) -> Tuple[int, str, str]:
    prompt = f"""
    You are analyzing the emails received by an investment bank. Analyze this email response regarding an investment opportunity.
    Classify it strictly into one of these three categories:

    1. Interested - The sender expresses any enthusiasm about investing or asks for next steps/follow-up.  

    2. Not Interested - The sender explicitly declines or shows zero interest in the opportunity and requests no further dialogue. **Only** use this when you are certain they are saying “no.” Note that if the sender expresses any interest in continuing the dialogue in any way or shows any signs that they might want to work with us, an investment bank, avoid classifying this email as "Not Interested".

    3. Other - Anything else (off-topic, general chit-chat, ambiguous or requests unrelated to this deal). If in doubt, pick “Other” (to avoid false negatives).
    
    Email to classify:
    "{email_text}"

    Respond with ONLY one choice: "Interested", "Not Interested", or "Other".
    """
    try:
        response = genai_client.models.generate_content(
            model="gemini-2.0-flash",
            contents=prompt
        )
        category = response.text.strip()
        if category not in ["Interested", "Not Interested", "Other"]:
            category = f"Invalid response: {category}"
        return index, email_text, category
    except Exception as e:
        return index, email_text, f"API Error: {str(e)}"

# === PROCESS RESULTS IN PARALLEL ===
def process_and_print_results(emails: List[str], max_threads: int = 5):
    unordered_results = []

    with ThreadPoolExecutor(max_threads) as executor:
        futures = {
            executor.submit(classify_email, i, email): i
            for i, email in enumerate(emails)
        }

        for future in as_completed(futures):
            unordered_results.append(future.result())

    ordered_results = sorted(unordered_results, key=lambda x: x[0])

    print("\n--- Email Body with Classification ---\n")
    for _, email, category in ordered_results:
        print(email)
        print(f"\n🔎 Classification: {category}")
        print("=" * 80)

# === MAIN FUNCTION ===
def main():
    emails = fetch_emails_from_outlook(max_emails=5)
    if not emails:
        print("No emails found.")
        return
    process_and_print_results(emails, max_threads=5)

if __name__ == "__main__":
    main()
