from langchain_ollama import OllamaEmbeddings
from langchain_chroma import Chroma
from langchain_core.documents import Document
import os
import pandas as pd

df = pd.read_csv("sample-data.csv")
embeddings = OllamaEmbeddings(model="mxbai-embed-large")

db_location = "./chrome_langchain_db"
add_documents = not os.path.exists(db_location)

if add_documents:
    documents = []
    ids = []

    for i, row in df.iterrows():
        document = Document(
            page_content=row["first name"] + " " + row["last name"] + " " + row["email"],
            metadata={"state": row["state"], "birthdate": row["birthdate"]},
            id=str(i)
        )
        
        ids.append(str(i))
        documents.append(document)

vector_store = Chroma(
    collection_name = "email_addresses",
    persist_directory=db_location,
    embedding_function=embeddings
)

if add_documents:
    vector_store.add_documents(documents=documents, ids=ids)

retriver = vector_store.as_retriever(
    search_kwargs={"k": 1}
)

def find_email(query):
    results = retriver.invoke(query)

    if not results:
        return None

    doc = results[0]

    parts = doc.page_content.split()
    email = parts[-1]
    full_name = " ".join(parts[:-1])

    return {
        "name": full_name,
        "email": email,
        "metadata": doc.metadata
    }