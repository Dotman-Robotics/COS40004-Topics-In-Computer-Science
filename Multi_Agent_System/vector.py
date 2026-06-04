import threading
import uuid
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from langchain_ollama import OllamaEmbeddings
from langchain_chroma import Chroma
from langchain_core.documents import Document

CSV_PATH        = "sample-data.csv"
DB_LOCATION     = "./chroma_langchain_db"
COLLECTION_NAME = "email_addresses"
CANDIDATE_K     = 5

embeddings = OllamaEmbeddings(model="mxbai-embed-large")

vector_store = Chroma(
    collection_name=COLLECTION_NAME,
    persist_directory=DB_LOCATION,
    embedding_function=embeddings,
)


def _seed_database():
    try:
        count = vector_store._collection.count()
        if count > 0:
            print(f"[vector] DB has {count} contacts — skipping seed.")
            return
    except Exception as e:
        print(f"[vector] Count check failed ({e}), proceeding with seed.")

    print("[vector] DB is empty — seeding from CSV...")
    try:
        df = pd.read_csv(CSV_PATH)
    except FileNotFoundError:
        raise FileNotFoundError(f"CSV not found at '{CSV_PATH}'.")
    except pd.errors.EmptyDataError:
        raise ValueError(f"CSV '{CSV_PATH}' is empty or malformed.")

    documents, ids = [], []
    for i, row in df.iterrows():
        first = str(row.get("first name", "")).strip()
        last  = str(row.get("last name",  "")).strip()
        email = str(row.get("email",      "")).strip()
        if not email or email == "nan":
            continue
        documents.append(Document(
            page_content=f"{first} {last} {email}",
            metadata={
                "state":      str(row.get("state",     "")),
                "birthdate":  str(row.get("birthdate", "")),
                "first_name": first,
                "last_name":  last,
            },
            id=str(i),
        ))
        ids.append(str(i))

    if not documents:
        print("[vector] WARNING: No valid contacts found.")
        return

    vector_store.add_documents(documents=documents, ids=ids)
    print(f"[vector] Seeded {len(documents)} contacts.")


_seed_database()

retriever = vector_store.as_retriever(search_kwargs={"k": CANDIDATE_K})


def _cosine_rank_candidates(query: str, candidates: list) -> list:
    if not candidates:
        return []

    names = []
    for doc in candidates:
        parts = doc.page_content.split()
        name  = " ".join(parts[:-1]) if len(parts) >= 2 else doc.page_content
        names.append(name)

    corpus = [query] + names
    try:
        vectorizer   = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
        tfidf_matrix = vectorizer.fit_transform(corpus)
        scores       = cosine_similarity(tfidf_matrix[0], tfidf_matrix[1:])[0]
    except ValueError:
        scores = [1.0] * len(candidates)

    return sorted(
        zip(scores, candidates, names),
        key=lambda x: x[0],
        reverse=True,
    )


# ── GUI confirmation queue ────────────────────────────────────────────────────
# Replaces the blocking input() call. When the GUI sends a command that needs
# contact confirmation, find_email() creates a pending request here and waits.
# The Flask API resolves it when the user clicks confirm/reject in the modal.

_pending_lock         = threading.Lock()
_pending_confirmations: dict[str, dict] = {}
# { token: { candidates:[{name,email,score}], event: threading.Event,
#            answer: int|None } }


def create_confirmation_request(candidates: list) -> str:
    """
    Register a confirmation request and return a token the GUI can poll.
    candidates: list of (score, doc, name) tuples from _cosine_rank_candidates
    """
    token = str(uuid.uuid4())
    items = []
    for score, doc, name in candidates:
        parts = doc.page_content.split()
        if len(parts) >= 2:
            items.append({
                "name":  name,
                "email": parts[-1],
                "score": round(float(score), 3),
            })

    event = threading.Event()
    with _pending_lock:
        _pending_confirmations[token] = {
            "candidates": items,
            "event":      event,
            "answer":     None,   # index into candidates, or -1 for none
        }
    return token


def get_pending_confirmation(token: str) -> dict | None:
    """Return the pending request data for a token (for the GUI to display)."""
    with _pending_lock:
        req = _pending_confirmations.get(token)
        if not req:
            return None
        return {"token": token, "candidates": req["candidates"]}


def resolve_confirmation(token: str, index: int) -> bool:
    """
    Resolve a pending confirmation.
    index = 0-based index into candidates list, or -1 for 'none of these'.
    Returns False if the token doesn't exist.
    """
    with _pending_lock:
        req = _pending_confirmations.get(token)
        if not req:
            return False
        req["answer"] = index
        req["event"].set()
    return True


def get_all_pending() -> list[dict]:
    """Return all unresolved confirmation requests (for GUI polling)."""
    with _pending_lock:
        return [
            {"token": t, "candidates": v["candidates"]}
            for t, v in _pending_confirmations.items()
            if not v["event"].is_set()
        ]


def find_email(query: str, interactive: bool = True) -> dict | None:
    if not query or not query.strip():
        return None

    candidates = retriever.invoke(query.strip())
    if not candidates:
        print(f"[vector] No candidates found for '{query}'.")
        return None

    ranked = _cosine_rank_candidates(query.strip(), candidates)

    # Non-interactive: return top result immediately (used by planner, search)
    if not interactive:
        score, doc, name = ranked[0]
        parts = doc.page_content.split()
        if len(parts) < 2:
            return None
        return {
            "name":     " ".join(parts[:-1]),
            "email":    parts[-1],
            "score":    round(float(score), 3),
            "metadata": doc.metadata,
        }

    # Interactive: use GUI confirmation queue instead of input()
    token = create_confirmation_request(ranked)
    print(f"[vector] Contact confirmation required — token: {token}")
    print(f"[vector] Waiting for GUI response…")

    with _pending_lock:
        event = _pending_confirmations[token]["event"]

    # Block until the GUI resolves it (timeout 120s)
    resolved = event.wait(timeout=120)

    with _pending_lock:
        req = _pending_confirmations.pop(token, {})

    if not resolved or req.get("answer") is None or req["answer"] == -1:
        print(f"[vector] Confirmation timed out or rejected.")
        return None

    idx = req["answer"]
    if idx < 0 or idx >= len(ranked):
        return None

    score, doc, name = ranked[idx]
    parts = doc.page_content.split()
    if len(parts) < 2:
        return None

    return {
        "name":     name,
        "email":    parts[-1],
        "score":    round(float(score), 3),
        "metadata": doc.metadata,
    }