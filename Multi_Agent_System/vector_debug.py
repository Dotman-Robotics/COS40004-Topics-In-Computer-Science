# debug_vector.py
import os
import pandas as pd
from langchain_ollama import OllamaEmbeddings
from langchain_chroma import Chroma
from langchain_core.documents import Document

CSV_PATH = "sample-data.csv"
DB_LOCATION = "./chroma_langchain_db_test"  # temp test DB

print("=== STEP 1: CSV CHECK ===")
try:
    df = pd.read_csv(CSV_PATH)
    print(f"Rows: {df.shape[0]}, Columns: {df.columns.tolist()}")
    print(df.head(3))
except Exception as e:
    print(f"FAILED to load CSV: {e}")
    exit()

print("\n=== STEP 2: BUILD DOCUMENTS ===")
documents, ids = [], []
for i, row in df.iterrows():
    first = str(row.get("first name", row.get("First Name", ""))).strip()
    last  = str(row.get("last name",  row.get("Last Name",  ""))).strip()
    email = str(row.get("email",      row.get("Email",      ""))).strip()
    print(f"  Row {i}: '{first}' '{last}' '{email}'")
    if not email or email == "nan":
        print(f"  SKIPPING row {i} — no email")
        continue
    documents.append(Document(
        page_content=f"{first} {last} {email}",
        metadata={"state": str(row.get("state", "")), "birthdate": str(row.get("birthdate", ""))},
        id=str(i),
    ))
    ids.append(str(i))

print(f"\nDocuments built: {len(documents)}")
if not documents:
    print("PROBLEM: No documents were built — check column names above")
    exit()

print("\n=== STEP 3: EMBEDDINGS ===")
try:
    embeddings = OllamaEmbeddings(model="mxbai-embed-large")
    test = embeddings.embed_query("test")
    print(f"Embeddings OK — vector length: {len(test)}")
except Exception as e:
    print(f"FAILED to generate embeddings: {e}")
    exit()

print("\n=== STEP 4: CHROMA INSERT ===")
try:
    vs = Chroma(
        collection_name="email_test",
        persist_directory=DB_LOCATION,
        embedding_function=embeddings,
    )
    vs.add_documents(documents=documents, ids=ids)
    count = vs._collection.count()
    print(f"Inserted. Doc count: {count}")
except Exception as e:
    print(f"FAILED to insert into Chroma: {e}")
    exit()

print("\n=== STEP 5: SEARCH TEST ===")
try:
    retriever = vs.as_retriever(search_kwargs={"k": 1})
    results = retriever.invoke("Johnson")
    print(f"Search results: {results}")
except Exception as e:
    print(f"FAILED to search: {e}")

# Cleanup test DB
import shutil
shutil.rmtree(DB_LOCATION, ignore_errors=True)
print("\n=== DONE — paste full output above ===")