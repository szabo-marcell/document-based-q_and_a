import os
from dotenv import load_dotenv
from langchain_huggingface import HuggingFaceEmbeddings
from sklearn.metrics.pairwise import cosine_similarity
from typing import List
import numpy as np

load_dotenv()

# A HuggingFace Hub az HF_TOKEN nevet keresi; ha csak HF_API_KEY van, mappeljük.
if not os.getenv("HF_TOKEN") and os.getenv("HF_API_KEY"):
    os.environ["HF_TOKEN"] = os.environ["HF_API_KEY"]

EMBEDDING_MODEL = "sentence-transformers/all-MiniLM-L6-v2"

_embeddings_instance = None

def _get_embeddings() -> HuggingFaceEmbeddings:
    global _embeddings_instance
    if _embeddings_instance is None:
        _embeddings_instance = HuggingFaceEmbeddings(model_name=EMBEDDING_MODEL)
    return _embeddings_instance


def find_most_similar(query: str, docs: List[str], top_k: int = 3) -> List[tuple]:
    """
    Visszaadja a query-hez legjobban illeszkedő top_k dokumentumot
    hasonlósági pontszámmal.

    HuggingFace API hívás: sentence-transformers/all-MiniLM-L6-v2 (embedding)

    Args:
        query:  A keresési kérdés.
        docs:   A dokumentumok listája.
        top_k:  Hány legjobb találatot adjon vissza.

    Returns:
        Lista (dokumentum, hasonlóság_pontszám) párokat top_k-ig.
    """
    if not docs:
        return []

    embeddings = _get_embeddings()
    doc_emb = embeddings.embed_documents(docs)
    query_emb = embeddings.embed_query(query)

    scores = cosine_similarity([query_emb], doc_emb)[0]
    top_indices = np.argsort(scores)[::-1][:top_k]

    return [(docs[i], round(float(scores[i]), 4)) for i in top_indices]
