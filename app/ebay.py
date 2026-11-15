
import os
import base64
import requests

from dotenv import load_dotenv

load_dotenv()

BASE_URL = "https://api.ebay.com"
MARKETPLACE_ID = os.getenv("EBAY_MARKETPLACE_ID", "EBAY_SG")


def get_access_token():
    client_id = os.getenv("EBAY_CLIENT_ID")
    client_secret = os.getenv("EBAY_CLIENT_SECRET")

    if not client_id or not client_secret:
        raise ValueError("Add eBay Client ID and Secret to .env")

    raw = f"{client_id}:{client_secret}"
    encoded = base64.b64encode(raw.encode()).decode()

    response = requests.post(
        f"{BASE_URL}/identity/v1/oauth2/token",
        headers={
            "Authorization": f"Basic {encoded}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        data={
            "grant_type": "client_credentials",
            "scope": "https://api.ebay.com/oauth/api_scope",
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()["access_token"]


def get_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "X-EBAY-C-MARKETPLACE-ID": MARKETPLACE_ID,
    }


def get_books_category_ids(token):
    headers = get_headers(token)

    tree_response = requests.get(
        f"{BASE_URL}/commerce/taxonomy/v1/"
        "get_default_category_tree_id",
        headers=headers,
        params={"marketplace_id": MARKETPLACE_ID},
        timeout=30,
    )
    tree_response.raise_for_status()
    tree_id = tree_response.json()["categoryTreeId"]

    suggestions_response = requests.get(
        f"{BASE_URL}/commerce/taxonomy/v1/"
        f"category_tree/{tree_id}/get_category_suggestions",
        headers=headers,
        params={"q": "Books"},
        timeout=30,
    )
    suggestions_response.raise_for_status()

    suggestions = suggestions_response.json().get(
        "categorySuggestions", []
    )

    category_ids = []

    for suggestion in suggestions:
        category = suggestion.get("category", {})
        category_id = category.get("categoryId")

        if category_id:
            category_ids.append(str(category_id))

    if not category_ids:
        raise RuntimeError("No Books categories found for EBAY_SG")

    return list(dict.fromkeys(category_ids))


def search_latest_books(max_results=600):
    token = get_access_token()
    headers = get_headers(token)
    category_ids = get_books_category_ids(token)

    listings = {}

    for category_id in category_ids:
        for offset in range(0, max_results, 200):
            page_size = min(200, max_results - offset)

            response = requests.get(
                f"{BASE_URL}/buy/browse/v1/item_summary/search",
                headers=headers,
                params={
                    "category_ids": category_id,
                    "limit": page_size,
                    "offset": offset,
                    "sort": "newlyListed",
                },
                timeout=30,
            )
            response.raise_for_status()

            page = response.json().get("itemSummaries", [])

            for item in page:
                item_id = item.get("itemId")
                if item_id:
                    listings[item_id] = item

            if len(page) < page_size:
                break

    newest = sorted(
        listings.values(),
        key=lambda item: item.get("itemOriginDate", ""),
        reverse=True,
    )

    return newest[:max_results]


def fetch_latest_books_as_records(max_results=600):
    records = []

    for item in search_latest_books(max_results):
        seller = item.get("seller") or {}
        item_id = item.get("itemId", "")
        username = seller.get("username") or item_id

        records.append({
            "platform": "ebay_sg",
            "seller_id": username,
            "seller_name": username,
            "store_name": "",
            "business_email": "",
            "business_phone": "",
            "registered_at": "",
            "latest_upload_id": item_id,
            "latest_upload_name": item.get("title", ""),
            "latest_upload_at": item.get("itemOriginDate", ""),
            "latest_upload_url": item.get("itemWebUrl", ""),
            "source_url": item.get("itemWebUrl", ""),
        })

    return records
