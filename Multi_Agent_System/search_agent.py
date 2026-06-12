"""
search_agent.py

Given a natural-language task description (e.g. "find painters to paint the office"),
this agent:
  1. Uses DuckDuckGo to search for local/relevant service providers
  2. Scrapes each result page for contact details (email, phone, address)
  3. Returns structured provider cards
  4. Can draft a formal outreach email to any selected provider

Dependencies:
    pip install duckduckgo-search requests beautifulsoup4 lxml
"""

import re
import time
import json
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from ollama_client import ask_llm_chat, ask_llm
from utils import safe_parse_json

# ── Constants ─────────────────────────────────────────────────────────────────

REQUEST_TIMEOUT  = 8      # seconds per HTTP request
MAX_SCRAPE_PAGES = 2      # how many internal pages to follow (e.g. /contact)
DEFAULT_N        = 5      # default number of providers to find

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# Patterns for contact detail extraction
EMAIL_RE   = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE   = re.compile(
    r"(\+?1[\s\-.]?)?"
    r"(\(?\d{3}\)?[\s\-.]?)"
    r"\d{3}[\s\-.]?\d{4}"
)

# Pages that are likely to have contact info
CONTACT_PAGE_HINTS = [
    "contact", "contact-us", "contactus", "get-in-touch",
    "about", "about-us", "reach-us", "enquiry", "enquiries",
]


# ── Query generation ──────────────────────────────────────────────────────────

QUERY_SYSTEM = """You are a search query generator for finding business service providers.
Output JSON only. No explanation.

Given a task description, generate:
1. A focused web search query to find service providers (businesses, freelancers, agencies)
2. A short label describing the type of provider being sought

Output: {"query": "...", "provider_type": "..."}

Rules:
- The query should find actual business listings or directories
- Include words like "services", "company", "hire", "professional" as appropriate
- Keep queries specific enough to return relevant businesses"""


def _generate_search_query(task: str) -> dict:
    response = ask_llm_chat(QUERY_SYSTEM, f'Task: "{task}"\nJSON:')
    result   = safe_parse_json(response)
    if not result or "query" not in result:
        # Fallback: use task directly
        return {"query": task + " services hire", "provider_type": "service providers"}
    return result


# ── Web search ────────────────────────────────────────────────────────────────

def _ddg_search(query: str, max_results: int) -> list[dict]:
    """
    Search DuckDuckGo and return a list of {title, url, snippet} dicts.
    Supports both the old duckduckgo_search package and the renamed ddgs package.
    """
    try:
        # Try new package name first (pip install ddgs)
        
        from ddgs import DDGS
        

        results = []
        with DDGS() as ddg:
            for r in ddg.text(query, max_results=max_results * 2):
                results.append({
                    "title":   r.get("title", ""),
                    "url":     r.get("href",  ""),
                    "snippet": r.get("body",  ""),
                })
        return results
    except ImportError:
        print("[search] Search package not installed. Run: pip install ddgs")
        return []
    except Exception as e:
        print(f"[search] DDG search error: {e}")
        return []


# ── Scraper ───────────────────────────────────────────────────────────────────

def _fetch_html(url: str) -> str | None:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT, allow_redirects=True)
        if resp.status_code == 200 and "text/html" in resp.headers.get("Content-Type", ""):
            return resp.text
    except Exception:
        pass
    return None


def _extract_emails(html: str, base_url: str) -> set[str]:
    """Extract email addresses from HTML text and mailto links."""
    soup   = BeautifulSoup(html, "lxml")
    emails = set()

    # mailto: hrefs
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if href.lower().startswith("mailto:"):
            addr = href[7:].split("?")[0].strip()
            if EMAIL_RE.match(addr):
                emails.add(addr.lower())

    # Raw text scan
    for match in EMAIL_RE.findall(soup.get_text()):
        emails.add(match.lower())

    # Filter out obvious non-emails (image filenames, etc.)
    emails = {e for e in emails if "." in e.split("@")[-1] and len(e) < 80}
    return emails


def _extract_phones(html: str) -> set[str]:
    soup   = BeautifulSoup(html, "lxml")
    phones = set()

    # tel: hrefs
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if href.lower().startswith("tel:"):
            phones.add(href[4:].strip())

    # Raw text
    for match in PHONE_RE.findall(soup.get_text()):
        raw = "".join(match).strip()
        if len(re.sub(r"\D", "", raw)) >= 10:
            phones.add(raw)

    return phones


def _find_contact_page_url(html: str, base_url: str) -> str | None:
    """Try to find the URL of a /contact or /about page."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup.find_all("a", href=True):
        href = tag["href"].lower().rstrip("/")
        slug = href.split("/")[-1]
        # Strip .html extension so both /contact and /contact.html match
        slug = re.sub(r"\.html?$", "", slug)
        if slug in CONTACT_PAGE_HINTS:
            return urljoin(base_url, tag["href"])
    return None


def _extract_business_name(html: str, fallback_title: str) -> str:
    soup = BeautifulSoup(html, "lxml")
    # Try og:site_name first
    og = soup.find("meta", property="og:site_name")
    if og and og.get("content", "").strip():
        return og["content"].strip()
    # Then <title>
    if soup.title and soup.title.string:
        raw = soup.title.string.strip()
        # Strip common suffixes like "| Home", "- Official Site"
        raw = re.split(r"[|\-–—]", raw)[0].strip()
        if raw:
            return raw
    return fallback_title


def scrape_provider(url: str, title: str) -> dict:
    """
    Scrape a single URL for contact details.
    Returns a provider dict with name, url, emails, phones, scraped_from.
    """
    provider = {
        "name":         title,
        "url":          url,
        "emails":       [],
        "phones":       [],
        "scraped_from": [],
        "snippet":      "",
    }

    # Fetch homepage
    html = _fetch_html(url)
    if not html:
        return provider

    provider["name"] = _extract_business_name(html, title)
    emails  = _extract_emails(html, url)
    phones  = _extract_phones(html)
    provider["scraped_from"].append(url)

    # Always try the contact page to collect additional emails/phones —
    # not just as a fallback when the homepage yields nothing.
    contact_url = _find_contact_page_url(html, url)
    if contact_url and contact_url != url:
        contact_html = _fetch_html(contact_url)
        if contact_html:
            emails |= _extract_emails(contact_html, url)
            phones |= _extract_phones(contact_html)
            provider["scraped_from"].append(contact_url)

    provider["emails"] = sorted(emails)
    provider["phones"] = sorted(phones)
    return provider


# ── Result filtering ──────────────────────────────────────────────────────────

# Domains to skip (aggregators, social media, job boards — not actual providers)
_SKIP_DOMAINS = {
    "facebook.com", "instagram.com", "twitter.com", "x.com",
    "linkedin.com", "youtube.com", "yelp.com", "tripadvisor.com",
    "indeed.com", "glassdoor.com", "reddit.com", "wikipedia.org",
    "google.com", "bing.com", "yahoo.com", "amazon.com",
    "thumbtack.com", "angi.com", "houzz.com", "bark.com",
    "trustpilot.com", "bbb.org",
}

def _is_skippable(url: str) -> bool:
    try:
        domain = urlparse(url).netloc.lower().lstrip("www.")
        return any(domain == skip or domain.endswith("." + skip) for skip in _SKIP_DOMAINS)
    except Exception:
        return False


# ── Outreach email drafting ───────────────────────────────────────────────────

OUTREACH_SYSTEM = """You are a professional business email writer.
Write a formal outreach email on behalf of a business seeking a service provider.
Output JSON only. No explanation.
Output: {"subject": "...", "body": "..."}

Rules:
- Tone: professional, courteous, concise
- Mention what the job is and that you are requesting a quote
- Leave a placeholder [YOUR NAME] and [YOUR COMPANY] for personalisation
- Do not invent specific details not provided"""


def draft_outreach_email(provider: dict, task_description: str, extra_context: str = "") -> dict:
    """
    Draft a formal outreach email to a provider for a given task.
    Returns {subject, body, to} or {error}.
    """
    to_email = provider["emails"][0] if provider.get("emails") else ""

    user_msg = (
        f"Service provider: {provider['name']} ({provider['url']})\n"
        f"Task we need done: {task_description}\n"
        f"Additional context: {extra_context or 'None'}\n\n"
        f"Write a professional outreach email requesting their services and a quote.\n"
        f"JSON:"
    )

    # RAG WEEK 12: index outreach contact before LLM call so it always runs
    try:
        from rag_store import index_outreach_contact
        index_outreach_contact(provider, task_description)
    except Exception as e:
        print(f"[search] RAG index skipped: {e}")

    response = ask_llm_chat(OUTREACH_SYSTEM, user_msg)
    result   = safe_parse_json(response)

    if not result:
        # Fallback template - LLM failed but outreach can still proceed
        print("[search] LLM draft failed - using fallback template.")
        result = {
            "subject": f"Enquiry - {task_description[:60].title()}",
            "body": (
                f"Hi,\n\n"
                f"We are interested in your services regarding: {task_description}.\n\n"
                f"Could you please provide more information and a quote?\n\n"
                f"Best regards,\n[YOUR NAME]\n[YOUR COMPANY]"
            ),
        }

    result.setdefault("subject", f"Enquiry - {task_description[:60]}")
    result.setdefault("body",    "Please provide a quote for the work described.")
    result["to"]       = to_email
    result["provider"] = provider["name"]

    return result


# ── Main entry point ──────────────────────────────────────────────────────────

def search_agent(task: str, n: int = DEFAULT_N) -> dict:
    """
    Full pipeline: generate query → search → scrape → return provider list.

    Returns:
        {
            "task":          str,
            "provider_type": str,
            "query":         str,
            "providers":     list[dict],   # up to n providers with contact info
            "total_found":   int,
        }
    """
    print(f"[search] Task: {task!r}")

    # 1. Generate a focused search query
    query_info    = _generate_search_query(task)
    query         = query_info["query"]
    provider_type = query_info.get("provider_type", "service providers")
    print(f"[search] Query: {query!r}  |  Type: {provider_type}")

    # 2. Search DuckDuckGo
    raw_results = _ddg_search(query, max_results=n * 3)
    print(f"[search] Raw results: {len(raw_results)}")

    # 3. Filter out aggregator/social sites
    filtered = [r for r in raw_results if r["url"] and not _is_skippable(r["url"])]
    print(f"[search] After filtering: {len(filtered)}")

    # 4. Scrape each result — stop once we have n providers
    providers = []
    seen_domains = set()

    for result in filtered:
        if len(providers) >= n:
            break

        url = result["url"]
        try:
            domain = urlparse(url).netloc.lower().lstrip("www.")
        except Exception:
            continue

        # Skip duplicate domains
        if domain in seen_domains:
            continue
        seen_domains.add(domain)

        print(f"[search] Scraping: {url}")
        provider          = scrape_provider(url, result["title"])
        provider["snippet"] = result.get("snippet", "")

        # Only include if we got at least a name and URL
        if provider["name"] or provider["url"]:
            providers.append(provider)

        time.sleep(0.4)  # Polite delay between requests

    print(f"[search] Done — {len(providers)} provider(s) found.")

    return {
        "task":          task,
        "provider_type": provider_type,
        "query":         query,
        "providers":     providers,
        "total_found":   len(providers),
    }