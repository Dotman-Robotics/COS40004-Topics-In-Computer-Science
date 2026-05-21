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
                "state":      str(row.get("state",      "")),
                "birthdate":  str(row.get("birthdate",  "")),
                "first_name": first,
                "last_name":  last,
            },
            id=str(i),
        ))
        ids.append(str(i))

    if not documents:
        print("[vector] WARNING: No valid contacts found. Check CSV column names.")
        return

    vector_store.add_documents(documents=documents, ids=ids)
    print(f"[vector] Seeded {len(documents)} contacts. DB count: {vector_store._collection.count()}")


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

    ranked = sorted(
        zip(scores, candidates, names),
        key=lambda x: x[0],
        reverse=True,
    )
    return ranked


def _confirm_contact(name: str, email: str) -> bool:
    while True:
        answer = input(f"  Did you mean {name} <{email}>? (y/n): ").strip().lower()
        if answer in ("y", "yes"):
            return True
        if answer in ("n", "no"):
            return False
        print("  Please enter y or n.")


def find_email(query: str, interactive: bool = True) -> dict | None:
    if not query or not query.strip():
        return None

    candidates = retriever.invoke(query.strip())
    if not candidates:
        print(f"[vector] No candidates found for '{query}'.")
        return None

    ranked = _cosine_rank_candidates(query.strip(), candidates)

    if not interactive:
        score, doc, name = ranked[0]
        parts = doc.page_content.split()
        if len(parts) < 2:
            return None
        return {
            "name":     " ".join(parts[:-1]),
            "email":    parts[-1],
            "metadata": doc.metadata,
        }

    print(f"\n[vector] Found {len(ranked)} candidate(s) for '{query}':")
    for i, (score, doc, name) in enumerate(ranked):
        parts = doc.page_content.split()
        if len(parts) < 2:
            continue
        email = parts[-1]

        confirmed = _confirm_contact(name, email)
        if confirmed:
            print(f"  Confirmed: {name} <{email}>")
            return {
                "name":     name,
                "email":    email,
                "metadata": doc.metadata,
            }
        else:
            if i < len(ranked) - 1:
                print("  Trying next candidate...")
            else:
                print("  No more candidates. Contact not found.")

    return None