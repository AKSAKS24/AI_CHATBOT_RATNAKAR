import os
import re
from functools import lru_cache
from typing import Optional, Tuple, List, Dict

import spacy
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import ChatOpenAI, OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain.chains import RetrievalQA
from langchain.prompts import PromptTemplate

# ----------------------- CONFIG -----------------------

DEFAULT_MODEL = "gpt-4o"      # adjust if needed
CHUNK_SIZE = 1000
CHUNK_OVERLAP = 100

# ----------------------- NLP Loader -----------------------

@lru_cache(maxsize=1)
def get_nlp():
    try:
        return spacy.load("en_core_web_sm", disable=["parser", "tagger"])
    except OSError:
        from spacy.cli import download
        download("en_core_web_sm")
        return spacy.load("en_core_web_sm", disable=["parser", "tagger"])

# ----------------------- TEXT SANITIZATION -----------------------

def _mask_patterns(text: str) -> str:
    # Emails
    text = re.sub(
        r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b',
        "[EMAIL]",
        text,
        flags=re.IGNORECASE
    )
    # URLs/domains
    text = re.sub(
        r'\b(?:https?://|http://|www\.)?[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b',
        "[DOMAIN]",
        text,
        flags=re.IGNORECASE
    )
    # Phone numbers (simple)
    text = re.sub(
        r'(?:(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{2,4}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{4,})',
        "[PHONE]",
        text
    )
    # IPv4
    text = re.sub(r'\b(?:\d{1,3}\.){3}\d{1,3}\b', "[IP]", text)
    # Project-like
    text = re.sub(
        r'\b(Project|Programme|Program|Engagement|SOW|RFP|Contract|Agreement)\s*(?:Name|ID|Code|Title)?\s*[:\-–=]\s*([A-Za-z][\w &().\-]{2,})',
        lambda m: f"{m.group(1)}: [PROJECT]",
        text,
        flags=re.IGNORECASE
    )
    # Client-like
    text = re.sub(
        r'\b(Client|Customer|Account|Partner)\s*(?:Name)?\s*[:\-–=]\s*([A-Za-z][\w &().\-]{2,})',
        lambda m: f"{m.group(1)}: [CLIENT]",
        text,
        flags=re.IGNORECASE
    )
    # "for/with <Proper Noun Phrase>" after relevant keywords
    text = re.sub(
        r'\b(?:project|programme|program|client|customer|engagement|contract)\s+(?:for|with)\s+([A-Z][\w&().\-]+(?:\s+[A-Z][\w&().\-]+){0,4})',
        lambda m: re.sub(re.escape(m.group(1)), "[ENTITY]", m.group(0)),
        text,
        flags=re.IGNORECASE
    )
    return text


def _apply_span_masks(text: str, spans: List[Dict[str, str]]) -> str:
    # spans = list of dicts: {"start": int, "end": int, "label": str}
    spans = sorted(spans, key=lambda s: (s["start"], s["end"]))
    merged = []
    prev = None
    for s in spans:
        if prev is None:
            prev = s
        else:
            if s["start"] <= prev["end"]:
                prev["end"] = max(prev["end"], s["end"])
                label_priority = {"PERSON": 4, "ORG": 3, "GPE": 2, "LOC": 2, "NORP": 1, "PRODUCT": 1}
                if label_priority.get(s["label"], 0) > label_priority.get(prev["label"], 0):
                    prev["label"] = s["label"]
            else:
                merged.append(prev)
                prev = s
    if prev:
        merged.append(prev)

    out = []
    last_idx = 0
    label_map = {
        "PERSON": "[PERSON]",
        "ORG": "[ORG]",
        "GPE": "[LOCATION]",
        "LOC": "[LOCATION]",
        "NORP": "[AFFILIATION]",
        "PRODUCT": "[PRODUCT]"
    }
    for s in merged:
        if last_idx < s["start"]:
            out.append(text[last_idx:s["start"]])
        out.append(label_map.get(s["label"], "[ENTITY]"))
        last_idx = s["end"]
    out.append(text[last_idx:])
    return "".join(out)


def sanitize_text(text: str) -> str:
    """
    Cleans sensitive or client-specific data before vectorization.
    Masks: emails, domains, phone numbers, IPs, orgs, persons, locations,
           affiliations, project names, clients, etc.
    """
    # 1) Regex-based quick masks
    text = _mask_patterns(text)

    # 2) Named entity masking via spaCy (offset-safe)
    nlp = get_nlp()
    doc = nlp(text)
    spans = []
    for ent in doc.ents:
        if ent.label_ in ["ORG", "PERSON", "GPE", "LOC", "NORP", "PRODUCT"]:
            spans.append({"start": ent.start_char, "end": ent.end_char, "label": ent.label_})
    text = _apply_span_masks(text, spans)

    return text

# ----------------------- EMBEDDINGS -----------------------

def _get_embeddings(api_key: str, api_base: Optional[str] = None):
    """Helper to get OpenAI embedding model."""
    kwargs = {"openai_api_key": api_key}
    if api_base:
        kwargs["openai_api_base"] = api_base
    return OpenAIEmbeddings(**kwargs)

# ----------------------- QA ENGINE BUILDER -----------------------

def build_qa_engine(
    raw_text: Optional[str],
    openai_api_key: str,
    model_name: Optional[str] = DEFAULT_MODEL,
    chunk_size: int = CHUNK_SIZE,
    chunk_overlap: int = CHUNK_OVERLAP,
    cache_name: Optional[str] = None,
    load_vectorstore_obj: Optional[FAISS] = None,
    openai_api_base: Optional[str] = None
) -> Tuple[RetrievalQA, Optional[FAISS]]:
    """
    Builds or reloads a QA engine.
    - If `load_vectorstore_obj` is provided → uses existing FAISS.
    - Else → sanitizes, chunks, embeds, and creates new FAISS index.
    """
    if load_vectorstore_obj:
        vectorstore = load_vectorstore_obj
    else:
        if not raw_text:
            raise ValueError("Cannot build vectorstore — no raw_text provided.")

        # 1) Sanitize text before embeddings
        clean_text = sanitize_text(raw_text)

        # 2) Split text into chunks
        splitter = RecursiveCharacterTextSplitter(chunk_size=chunk_size, chunk_overlap=chunk_overlap)
        chunks = splitter.split_text(clean_text)

        # 3) Embed & build FAISS index
        embeddings = _get_embeddings(openai_api_key, api_base=openai_api_base)
        vectorstore = FAISS.from_texts(chunks, embeddings)

    # 4) Retriever configuration — controls recall level
    retriever = vectorstore.as_retriever(search_kwargs={"k": 6})

    # 5) LLM setup
    llm_kwargs = {"model": model_name, "temperature": 0.3, "openai_api_key": openai_api_key}
    if openai_api_base:
        llm_kwargs["openai_api_base"] = openai_api_base
    llm = ChatOpenAI(**llm_kwargs)

    # 6) Context-based prompt
    prompt_template = """Use the context below to answer accurately and completely.

- Always use the provided context.
- Infer missing details logically if partially available.
- If the question is generic, answer in the context domain.
- Be concise and clear.

Context:
{context}

Question:
{question}

Answer:
"""
    prompt = PromptTemplate(template=prompt_template, input_variables=["context", "question"])

    # 7) QA Chain
    qa_chain = RetrievalQA.from_chain_type(
        llm=llm,
        retriever=retriever,
        chain_type="stuff",
        chain_type_kwargs={"prompt": prompt},
        return_source_documents=True,
    )

    return qa_chain, vectorstore

# ----------------------- SAVE / LOAD HANDLERS -----------------------

def save_vectorstore(vectorstore: FAISS, persist_dir: str, cache_name: Optional[str] = None):
    """
    Persists FAISS vectorstore locally.
    """
    cache_name = cache_name or "default"
    target_dir = os.path.join(persist_dir, cache_name)
    os.makedirs(target_dir, exist_ok=True)
    vectorstore.save_local(target_dir)


def load_vectorstore(
    openai_api_key: str,
    persist_dir: str,
    cache_name: Optional[str] = None,
    openai_api_base: Optional[str] = None
) -> Optional[FAISS]:
    """
    Loads existing FAISS vectorstore if available.
    Returns None if not found.
    """
    cache_name = cache_name or "default"
    target_dir = os.path.join(persist_dir, cache_name)

    if not os.path.isdir(target_dir):
        return None

    embeddings = _get_embeddings(openai_api_key, api_base=openai_api_base)
    return FAISS.load_local(target_dir, embeddings, allow_dangerous_deserialization=True)