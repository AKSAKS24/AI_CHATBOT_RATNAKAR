import streamlit as st
import os
import io
import json
import time
import base64
import hashlib
import requests
from urllib.parse import urlparse
from dotenv import load_dotenv
from typing import Dict, Any, List, Optional

# local modules
from file_loader import get_raw_text
from qa_engine import build_qa_engine, save_vectorstore, load_vectorstore

# ---------------- Load Environment ----------------
dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path=dotenv_path)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_API_BASE = os.getenv("OPENAI_API_BASE")  # Optional
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin")  # Simple settings guard

# ---------------- Persistence ----------------
BASE_DIR = os.path.dirname(__file__)
PERSIST_DIR = os.path.join(BASE_DIR, "persisted_data")
os.makedirs(PERSIST_DIR, exist_ok=True)

# ---------------- Streamlit UI Init ----------------
st.set_page_config(page_title="Doc Chatbot", layout="wide")

# ---------------- Global CSS ----------------
GLOBAL_CSS = """
<style>
...
</style>
"""

st.markdown(GLOBAL_CSS, unsafe_allow_html=True)

# ---------------- Session State Defaults ----------------
defaults = {
    "page_initialized": False,
    "chat_history": [],
    "qa": None,
    "current_cache_name": None,
    "page": "chat",  # Start with chat
    "authorized_settings": False,
    "autosync_enabled": False,
    "cache_configs": {},  # local read cache for configs
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

if not st.session_state.page_initialized:
    with st.spinner("Initializing..."):
        time.sleep(0.7)
    st.session_state.page_initialized = True

# ---------------- Utility: Config per cache ----------------
def cache_dir(cache_name: str) -> str:
    return os.path.join(PERSIST_DIR, cache_name)


def config_path(cache_name: str) -> str:
    return os.path.join(cache_dir(cache_name), "config.json")


def read_cache_config(cache_name: str) -> Dict[str, Any]:
    try:
        with open(config_path(cache_name), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def write_cache_config(cache_name: str, config: Dict[str, Any]):
    os.makedirs(cache_dir(cache_name), exist_ok=True)
    with open(config_path(cache_name), "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


def list_caches() -> List[str]:
    return [d for d in os.listdir(PERSIST_DIR) if os.path.isdir(os.path.join(PERSIST_DIR, d))]

# ---------------- SharePoint Helpers ----------------
def get_graph_token() -> str:
    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise RuntimeError("SharePoint credentials not set in .env.")
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    token_response = requests.post(token_url, data=token_data)
    token_response.raise_for_status()
    return token_response.json()["access_token"]


def share_link_to_drive_item_meta(share_link: str, access_token: str) -> Dict[str, Any]:
    encoded_url = base64.urlsafe_b64encode(share_link.strip().encode("utf-8")).decode("utf-8").rstrip("=")
    meta_url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem"
    meta_res = requests.get(meta_url, headers={"Authorization": f"Bearer {access_token}"})
    meta_res.raise_for_status()
    return meta_res.json()


def list_children_for_item(item_id: str, access_token: str) -> List[Dict[str, Any]]:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id_from_item(item_id, access_token)}/items/{item_id}/children"
    res = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})
    res.raise_for_status()
    return res.json().get("value", [])


def drive_id_from_item(item_id: str, access_token: str) -> str:
    # Retrieve driveId for given itemId
    meta_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{item_id}"
    url = f"https://graph.microsoft.com/v1.0/drive/items/{item_id}"
    res = requests.get(url, headers={"Authorization": f"Bearer {access_token}"})
    res.raise_for_status()
    parent = res.json().get("parentReference", {})
    drive_id = parent.get("driveId")
    if not drive_id:
        # fallback: root drive
        drive_id = res.json().get("parentReference", {}).get("driveId", "")
    return drive_id


def collect_files_recursively_from_item(item_json: Dict[str, Any], access_token: str) -> List[Dict[str, Any]]:
    results = []

    def _walk(item):
        if "file" in item:
            results.append({
                "id": item.get("id"),
                "name": item.get("name"),
                "etag": item.get("eTag") or item.get("@microsoft.graph.downloadUrl", "")[:32],
                "size": item.get("size"),
                "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                "downloadUrl": item.get("@microsoft.graph.downloadUrl"),
            })
        elif "folder" in item:
            # fetch children of this folder
            drive_id = item.get("parentReference", {}).get("driveId")
            current_id = item.get("id")
            if not drive_id:
                drive_id = drive_id_from_item(current_id, access_token)
            children_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{current_id}/children"
            res = requests.get(children_url, headers={"Authorization": f"Bearer {access_token}"})
            res.raise_for_status()
            for child in res.json().get("value", []):
                _walk(child)

    _walk(item_json)
    return results


def build_manifest(files: List[Dict[str, Any]]) -> Dict[str, Any]:
    # Map id -> etag for delta detection; also keep readable list
    return {
        "files": [
            {
                "id": f.get("id"),
                "name": f.get("name"),
                "etag": f.get("etag"),
                "size": f.get("size"),
                "lastModifiedDateTime": f.get("lastModifiedDateTime"),
            } for f in files
        ],
        "map": {f.get("id"): f.get("etag") for f in files},
        "count": len(files)
    }


def manifests_equal(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
    return (a.get("map") == b.get("map")) and (a.get("count") == b.get("count"))


def download_and_extract_text(files: List[Dict[str, Any]]) -> str:
    all_text = ""
    for f in files:
        url = f.get("downloadUrl")
        if not url:
            continue
        fres = requests.get(url)
        fres.raise_for_status()
        all_text += get_raw_text(fres.content, f.get("name", "file")) + "\n\n"
    return all_text

# ---------------- QA Loader/Builder ----------------
def load_cache_into_memory(name: str):
    try:
        vectorstore = load_vectorstore(OPENAI_API_KEY, PERSIST_DIR, cache_name=name, openai_api_base=OPENAI_API_BASE)
        if vectorstore:
            st.session_state.qa, _ = build_qa_engine(
                raw_text="",
                openai_api_key=OPENAI_API_KEY,
                openai_api_base=OPENAI_API_BASE,
                load_vectorstore_obj=vectorstore
            )
            st.session_state.current_cache_name = name
            return True
    except Exception as e:
        st.error(f"Failed to load cache '{name}': {e}")
    return False


def rebuild_vectorstore_and_save(cache_name: str, raw_text: str):
    qa, vectorstore = build_qa_engine(
        raw_text=raw_text,
        openai_api_key=OPENAI_API_KEY,
        openai_api_base=OPENAI_API_BASE,
        cache_name=cache_name
    )
    save_vectorstore(vectorstore, PERSIST_DIR, cache_name=cache_name)
    st.session_state.qa = qa
    st.session_state.current_cache_name = cache_name

# ---------------- Auto-sync for SharePoint ----------------
def maybe_autosync_current_cache():
    cache = st.session_state.current_cache_name
    if not cache:
        return
    cfg = read_cache_config(cache)
    if cfg.get("type") != "sharepoint":
        return
    if not cfg.get("autosync", False):
        return
    share_link = cfg.get("sharepoint", {}).get("link")
    if not share_link:
        return

    try:
        token = get_graph_token()
        item_json = share_link_to_drive_item_meta(share_link, token)
        files = collect_files_recursively_from_item(item_json, token)
        new_manifest = build_manifest(files)
        old_manifest = cfg.get("manifest", {})
        if not manifests_equal(new_manifest, old_manifest):
            # Rebuild
            with st.spinner("Detected SharePoint changes. Syncing knowledge base..."):
                text = download_and_extract_text(files)
                rebuild_vectorstore_and_save(cache, text)
                cfg["manifest"] = new_manifest
                write_cache_config(cache, cfg)
                st.toast(f"Auto-synced changes for cache '{cache}'.", icon="✅")
    except Exception as e:
        st.warning(f"Auto-sync error: {e}")

# ---------------- Header (Chat-first UI) ----------------
def render_header():
    st.markdown('<div class="header-bar">', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([4, 3, 2])

    with c1:
        st.markdown('<div class="header-title">Doc Chatbot</div>', unsafe_allow_html=True)

    with c2:
        caches = ["-- select --"] + list_caches()
        selected = st.selectbox(
            "Cache",
            options=caches,
            index=caches.index(st.session_state.current_cache_name) if st.session_state.current_cache_name in caches else 0,
            key="header_cache_select"
        )
        if selected and selected != "-- select --":
            if selected != st.session_state.current_cache_name:
                if load_cache_into_memory(selected):
                    st.session_state.chat_history = []
                    st.rerun()
    with c3:
        settings_clicked = st.button("⚙️ Settings", key="go_settings", help="Open settings", use_container_width=True)
        if settings_clicked:
            st.session_state.page = "settings"
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Pages ----------------
def page_chat():
    render_header()

    # Auto-sync check for SharePoint caches
    maybe_autosync_current_cache()

    st.header("Chat")
    if not st.session_state.qa:
        st.info("No knowledge base loaded. Choose a cache from the top-right selector or add one in Settings.")
        return

    # Show a badge row with current cache and autosync state
    cfg = read_cache_config(st.session_state.current_cache_name)
    badge = f"Using cache: {st.session_state.current_cache_name}"
    if cfg.get("type") == "sharepoint":
        badge += "  |  Source: SharePoint"
        if cfg.get("autosync"):
            badge += "  |  Auto-sync: ON"
        else:
            badge += "  |  Auto-sync: OFF"
    st.markdown(f"<span class='badge'>{badge}</span>", unsafe_allow_html=True)
    st.markdown("<hr/>", unsafe_allow_html=True)

    chat_container = st.container()
    user_query = st.chat_input("Ask something about the document...")

    if user_query:
        with st.spinner("Thinking..."):
            try:
                result = st.session_state.qa({"query": user_query})
                st.session_state.chat_history.append({
                    "question": user_query,
                    "answer": result.get("result", ""),
                    "context": result.get("source_documents", [])
                })
            except Exception as e:
                st.error(f"Query error: {e}")

    with chat_container:
        for chat in st.session_state.chat_history:
            with st.chat_message("user"):
                st.markdown(chat["question"])
            with st.chat_message("assistant"):
                st.markdown(chat["answer"])


def page_settings():
    render_header()
    st.header("Settings")

    if not st.session_state.authorized_settings:
        st.warning("Authorization required to access settings.")
        pwd = st.text_input("Admin password", type="password")
        if st.button("Unlock"):
            if pwd == ADMIN_PASSWORD:
                st.session_state.authorized_settings = True
                st.success("Authorized")
                st.rerun()
            else:
                st.error("Invalid password")
        return

    tabs = st.tabs(["Upload from File", "Load from SharePoint", "Memory Management"])

    # ------------- Tab: Upload -------------
    with tabs[0]:
        st.subheader("Upload a File")
        uploaded_file = st.file_uploader(
            "Upload PDF, DOCX, XLS, XLSX, or ZIP",
            type=["pdf", "docx", "xls", "xlsx", "zip"]
        )
        cache_name = st.text_input("Cache name (unique)")

        if st.button("Process & Save to Memory (File)"):
            if not uploaded_file or not cache_name:
                st.warning("Please upload a file and provide a cache name.")
            else:
                try:
                    with st.spinner("Extracting text and building QA engine..."):
                        raw_bytes = uploaded_file.read()
                        raw_text = get_raw_text(raw_bytes, uploaded_file.name)
                        if not raw_text.strip():
                            st.error("No text extracted from the file.")
                        else:
                            rebuild_vectorstore_and_save(cache_name, raw_text)
                            cfg = {
                                "type": "file",
                                "source": {
                                    "filename": uploaded_file.name,
                                },
                                "autosync": False,
                                "manifest": {}
                            }
                            write_cache_config(cache_name, cfg)
                            st.success(f"Saved knowledge base as '{cache_name}'.")
                except Exception as e:
                    st.error(f"Failed to process file: {e}")

    # ------------- Tab: SharePoint -------------
    with tabs[1]:
        st.subheader("Load from SharePoint (Graph API)")
        st.caption("Provide a SharePoint sharing link to a file or folder. App uses app-only Graph access.")

        sp_link = st.text_input("SharePoint File/Folder Sharing Link")
        cache_name_sp = st.text_input("Cache name for this SharePoint source (unique)")
        autosync = st.toggle(
            "Enable auto-sync (checks for changes and refreshes automatically)",
            value=True
        )

        if st.button("Load and Save (SharePoint)"):
            if not sp_link or not cache_name_sp:
                st.warning("Please provide both SharePoint link and cache name.")
            elif not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
                st.error("SharePoint credentials missing. Set TENANT_ID, CLIENT_ID, CLIENT_SECRET in .env.")
            else:
                try:
                    with st.spinner("Loading from SharePoint..."):
                        token = get_graph_token()
                        item_json = share_link_to_drive_item_meta(sp_link, token)
                        files = collect_files_recursively_from_item(item_json, token)
                        text = download_and_extract_text(files)
                        rebuild_vectorstore_and_save(cache_name_sp, text)
                        manifest = build_manifest(files)
                        cfg = {
                            "type": "sharepoint",
                            "sharepoint": {
                                "link": sp_link
                            },
                            "autosync": autosync,
                            "manifest": manifest
                        }
                        write_cache_config(cache_name_sp, cfg)
                        st.success(f"Loaded and saved SharePoint source as '{cache_name_sp}'.")
                except Exception as e:
                    st.error(f"Error loading SharePoint: {e}")

    # ------------- Tab: Memory Management -------------
    with tabs[2]:
        st.subheader("Persistent Memory")
        caches = list_caches()
        st.write("Available caches:", caches if caches else "No saved caches yet.")

        selected_cache = st.selectbox(
            "Select cache to load",
            options=["-- select --"] + caches,
            key="settings_cache_select"
        )

        if st.button("Load selected cache into memory"):
            if selected_cache and selected_cache != "-- select --":
                if load_cache_into_memory(selected_cache):
                    st.success(f"Loaded '{selected_cache}' into memory.")

        if st.button("Clear in-memory selection"):
            st.session_state.qa = None
            st.session_state.current_cache_name = None
            st.session_state.chat_history = []
            st.success("Cleared in-memory QA engine.")

        st.markdown("<hr/>", unsafe_allow_html=True)

        # Manage autosync per cache
        if caches:
            autosync_cache = st.selectbox(
                "Select cache to toggle auto-sync (SharePoint only)",
                options=["-- select --"] + caches,
                key="autosync_cache_sel"
            )
            if autosync_cache and autosync_cache != "-- select --":
                cfg = read_cache_config(autosync_cache)
                if cfg.get("type") == "sharepoint":
                    current = cfg.get("autosync", False)
                    st.write(f"Current auto-sync status: {'ON' if current else 'OFF'}")
                    new_val = st.toggle("Auto-sync for this cache", value=current, key="autosync_toggle_val")
                    if st.button("Save auto-sync setting"):
                        cfg["autosync"] = new_val
                        write_cache_config(autosync_cache, cfg)
                        st.success("Auto-sync setting updated.")
                else:
                    st.info("Selected cache is not a SharePoint source.")

    if st.button("⬅️ Back to Chat"):
        st.session_state.page = "chat"
        st.rerun()

# ---------------- Router ----------------
if st.session_state.page == "chat":
    page_chat()
else:
    page_settings()