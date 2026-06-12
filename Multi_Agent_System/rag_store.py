"""
rag_store.py  -  Week 12: Contextual Memory / RAG Knowledge Base

Indexes session outcomes, negotiation histories, and email interactions
into a per-account ChromaDB collection using mxbai-embed-large embeddings.

Before any new request is routed, the planner and monitor query this store
to retrieve relevant prior context about the sender or topic.

Key design rules:
- Account resolved at call time via _get_account() - never at import time.
- Per-account ChromaDB collection: rag_<account_slug>
- Separate from the contacts store in vector.py (./chroma_rag_db vs ./chroma_langchain_db)
- All public functions are safe to call even if ChromaDB is unavailable
  (they log a warning and return gracefully).
"""

import os
import json
from datetime import datetime, timedelta

# ── Account helper ─────────────────────────────────────────────────────────────

def _get_account() -> str:
    return os.environ.get("AGENT_ACCOUNT", "zoomertron@outlook.com")


def _get_slug() -> str:
    return _get_account().replace("@", "_").replace(".", "_")


# ── ChromaDB collection ────────────────────────────────────────────────────────

_chroma_client = None
_collections: dict = {}


def _get_collection():
    """
    Return (or create) the per-account ChromaDB collection.
    Lazy-initialised so import never fails if langchain_chroma is missing.
    """
    global _chroma_client, _collections
    slug = _get_slug()

    if slug in _collections:
        return _collections[slug]

    try:
        import chromadb
        from langchain_ollama import OllamaEmbeddings
        from langchain_chroma import Chroma

        embeddings = OllamaEmbeddings(model="mxbai-embed-large")

        if _chroma_client is None:
            _chroma_client = chromadb.PersistentClient(path="./chroma_rag_db")

        collection = Chroma(
            client              = _chroma_client,
            collection_name     = f"rag_{slug}",
            embedding_function  = embeddings,
        )
        _collections[slug] = collection
        print(f"[rag] Collection 'rag_{slug}' ready.")
        return collection

    except Exception as e:
        print(f"[rag] Could not initialise ChromaDB collection: {e}")
        return None


# ── Document builders ──────────────────────────────────────────────────────────

def _build_negotiation_docs(negotiations: dict, account: str) -> list[dict]:
    """
    Convert negotiation state records into indexable documents.
    Only indexes confirmed, rejected, and failed threads (not active ones).
    """
    docs = []
    for thread_id, rec in negotiations.items():
        status = rec.get("status", "")
        if status not in ("confirmed", "rejected", "failed"):
            continue

        title      = rec.get("title", "Meeting")
        initiator  = rec.get("initiator", "")
        rounds     = rec.get("rounds", 0)
        created    = rec.get("created_at", "")
        closed     = rec.get("closed_at", "")
        is_human   = rec.get("human_thread", False)

        # Extract the final confirmed slot if any
        slot_str = ""
        for entry in reversed(rec.get("history", [])):
            if entry.get("action") == "confirmed" and entry.get("slot"):
                s = entry["slot"]
                slot_str = f"{s.get('date','')} {s.get('time','')}".strip()
                break
            if entry.get("slot") and status == "confirmed":
                s = entry["slot"]
                slot_str = f"{s.get('date','')} {s.get('time','')}".strip()
                break

        thread_type = "human" if is_human else "agent"
        text = (
            f"Negotiation: {title}. "
            f"Counterparty: {initiator}. "
            f"Type: {thread_type}. "
            f"Outcome: {status}. "
            f"Rounds: {rounds}. "
            f"{'Confirmed slot: ' + slot_str + '.' if slot_str else ''} "
            f"Started: {created}. Closed: {closed}."
        ).strip()

        docs.append({
            "id":       f"neg_{_get_slug()}_{thread_id}",
            "text":     text,
            "metadata": {
                "type":        "negotiation",
                "account":     account,
                "thread_id":   thread_id,
                "title":       title,
                "counterparty": initiator,
                "status":      status,
                "rounds":      rounds,
                "slot":        slot_str,
                "created_at":  created,
            },
        })
    return docs


def _build_session_docs(session_log: list[dict], account: str) -> list[dict]:
    """
    Convert session log entries into indexable documents.
    Each entry represents a processed email with its outcome.
    """
    docs = []
    for i, entry in enumerate(session_log):
        email   = entry.get("email", {})
        summary = entry.get("summary", {})
        status  = entry.get("status", "")

        sender  = email.get("email", "")
        subject = email.get("subject", "")
        sumtext = summary.get("summary", "") if summary else ""
        cat     = summary.get("category", "") if summary else ""
        slot    = entry.get("slot", {}) or {}
        slot_str = f"{slot.get('date','')} {slot.get('time','')}".strip() if slot else ""

        if not sender and not subject:
            continue

        text = (
            f"Email from {sender}. "
            f"Subject: {subject}. "
            f"Category: {cat}. "
            f"Outcome: {status}. "
            f"{sumtext} "
            f"{'Confirmed slot: ' + slot_str + '.' if slot_str else ''}"
        ).strip()

        docs.append({
            "id":       f"sess_{_get_slug()}_{datetime.now().strftime('%Y%m%d')}_{i}",
            "text":     text,
            "metadata": {
                "type":        "session_email",
                "account":     account,
                "sender":      sender,
                "subject":     subject,
                "status":      status,
                "category":    cat,
                "slot":        slot_str,
                "indexed_at":  datetime.now().isoformat(),
            },
        })
    return docs


# ── Public API ─────────────────────────────────────────────────────────────────

def index_session(session_log: list[dict], account: str | None = None) -> int:
    """
    Index all session outcomes and negotiation histories into ChromaDB.
    Called by email_monitor.stop_monitor() at the end of each session.

    Returns the number of documents indexed, or 0 on failure.
    """
    if account is None:
        account = _get_account()

    collection = _get_collection()
    if collection is None:
        print("[rag] Skipping index — collection unavailable.")
        return 0

    # Load negotiation state for this account
    slug       = account.replace("@", "_").replace(".", "_")
    state_file = f"negotiations_{slug}.json"
    negotiations = {}
    try:
        with open(state_file, "r", encoding="utf-8") as f:
            negotiations = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        pass

    docs = []
    docs.extend(_build_negotiation_docs(negotiations, account))
    docs.extend(_build_session_docs(session_log, account))

    if not docs:
        print("[rag] Nothing to index this session.")
        return 0

    try:
        collection.add_texts(
            texts     = [d["text"]     for d in docs],
            metadatas = [d["metadata"] for d in docs],
            ids       = [d["id"]       for d in docs],
        )
        print(f"[rag] Indexed {len(docs)} document(s) for account '{account}'.")
        return len(docs)
    except Exception as e:
        print(f"[rag] Index error: {e}")
        return 0


def query_context(
    query: str,
    account: str | None = None,
    top_k: int = 3,
    sender_email: str | None = None,
) -> str:
    """
    Query the RAG store for context relevant to the query or sender.
    Returns a formatted string suitable for prepending to an LLM prompt.
    Returns empty string if nothing relevant found or store unavailable.

    Called by:
    - email_monitor.process_email() before negotiation routing
    - planner_agent.planner_agent() before routing user requests
    """
    if account is None:
        account = _get_account()

    collection = _get_collection()
    if collection is None:
        return ""

    try:
        # If we have a sender email, bias the query toward that contact
        search_query = f"{sender_email} {query}".strip() if sender_email else query

        results = collection.similarity_search_with_score(search_query, k=top_k)
        if not results:
            return ""

        # Filter out low-relevance results (distance > 0.8 in cosine space)
        relevant = [(doc, score) for doc, score in results if score < 0.8]
        if not relevant:
            return ""

        lines = ["Prior context from memory:"]
        for doc, score in relevant:
            meta = doc.metadata
            lines.append(f"- {doc.page_content} [relevance: {1 - score:.2f}]")

        context = "\n".join(lines)
        print(f"[rag] Retrieved {len(relevant)} context doc(s) for query: '{query[:60]}'")
        return context

    except Exception as e:
        print(f"[rag] Query error: {e}")
        return ""


def query_sender_history(sender_email: str, account: str | None = None) -> str:
    """
    Query prior interactions with a specific sender using exact metadata filtering.
    Uses ChromaDB where filter on sender field rather than pure semantic similarity
    to prevent false positives from unrelated documents in a small store.
    Falls back to semantic search if the metadata filter returns nothing.
    Called by email_monitor before passing to negotiation agent.
    """
    if account is None:
        account = _get_account()

    if not sender_email:
        return ""

    collection = _get_collection()
    if collection is None:
        return ""

    try:
        # Pass 1 - exact metadata match on sender email
        results = collection.get(
            where   = {"sender": sender_email},
            include = ["documents", "metadatas"],
        )
        docs  = results.get("documents", [])
        metas = results.get("metadatas", [])

        if docs:
            lines_out = ["Prior context from memory:"]
            for text, meta in zip(docs, metas):
                lines_out.append(f"- {text}")
            print(f"[rag] Retrieved {len(docs)} context doc(s) for sender: '{sender_email}'")
            return "\n".join(lines_out)

        # Pass 2 - fallback to semantic search if no exact match
        return query_context(
            query        = f"interactions with {sender_email}",
            account      = account,
            top_k        = 3,
            sender_email = sender_email,
        )

    except Exception as e:
        print(f"[rag] query_sender_history error: {e}")
        return ""

def get_all_documents(account: str | None = None) -> list[dict]:
    """
    Return all indexed documents for the current account.
    Used by the GUI /api/rag/query endpoint for inspection.
    """
    if account is None:
        account = _get_account()

    collection = _get_collection()
    if collection is None:
        return []

    try:
        results = collection.get(include=["documents", "metadatas"])
        docs = []
        for text, meta in zip(
            results.get("documents", []),
            results.get("metadatas", []),
        ):
            docs.append({"text": text, "metadata": meta})
        return docs
    except Exception as e:
        print(f"[rag] get_all_documents error: {e}")
        return []


def clear_collection(account: str | None = None) -> bool:
    """
    Delete all documents from the current account's RAG collection.
    The collection itself is preserved — only its contents are cleared.
    Called by POST /api/rag/clear from the GUI.
    Returns True on success, False on failure.
    """
    if account is None:
        account = _get_account()

    slug = account.replace("@", "_").replace(".", "_")

    # Remove from cache so next call re-initialises cleanly
    _collections.pop(slug, None)

    try:
        import chromadb
        client = chromadb.PersistentClient(path="./chroma_rag_db")
        # Delete and recreate the collection — fastest way to clear all docs
        try:
            client.delete_collection(f"rag_{slug}")
            print(f"[rag] Collection 'rag_{slug}' cleared.")
        except Exception:
            pass  # Collection may not exist yet — that's fine
        return True
    except Exception as e:
        print(f"[rag] Clear failed: {e}")
        return False


def index_outreach_contact(provider: dict, task: str, account: str | None = None) -> None:
    """
    Index a provider contact when an outreach email is sent.
    Called by search_agent.draft_outreach_email() so the monitor
    has full context when the provider replies - including what
    service was being sourced and why this provider was contacted.
    """
    if account is None:
        account = _get_account()

    collection = _get_collection()
    if collection is None:
        return

    provider_email = (provider.get("emails") or [""])[0]
    provider_name  = provider.get("name", "Unknown Provider")
    provider_url   = provider.get("url", "")

    text = (
        f"Outreach sent to {provider_name} ({provider_url}) "
        f"regarding task: {task}. "
        f"Contact email: {provider_email}. "
        f"This provider was sourced via the search agent. "
        f"Awaiting reply to arrange a meeting or service agreement."
    )

    # Use provider email as part of ID so it can be found by sender query
    safe_name = provider_name[:30].replace(" ", "_").replace("/", "_")
    doc_id    = f"outreach_{_get_slug()}_{safe_name}_{datetime.now().strftime('%Y%m%d%H%M%S')}"

    try:
        collection.add_texts(
            texts     = [text],
            metadatas = [{
                "type":        "outreach_contact",
                "account":     account,
                "provider":    provider_name,
                "task":        task,
                "sender":      provider_email,
                "url":         provider_url,
                "indexed_at":  datetime.now().isoformat(),
            }],
            ids = [doc_id],
        )
        print(f"[rag] Outreach contact indexed: {provider_name} <{provider_email}>")
    except Exception as e:
        print(f"[rag] Could not index outreach contact: {e}")