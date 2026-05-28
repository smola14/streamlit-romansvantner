from __future__ import annotations

import hmac
import hashlib
import io
import json
import os
import unicodedata
from collections.abc import Mapping
from pathlib import Path
from datetime import date, datetime, time, timedelta, timezone
from typing import Any

import matplotlib.pyplot as plt
from matplotlib import font_manager
from openpyxl import load_workbook
import requests
import streamlit as st
from fpdf import FPDF


API_BASE_URL = os.getenv("API1080_BASE_URL", "https://publicapi.1080motion.com").rstrip("/")
APP_API_KEY = st.secrets.get("api_1080_key", os.getenv("API1080_KEY", "")).strip()
API_TIMEOUT_SECONDS = 20
BLUE_RGB = (48, 54, 116)
BLACK_RGB = (37, 36, 35)
BLUE_HEX = "#303674"
FV_NORMS_PATH = Path(__file__).resolve().parent / "data" / "fv_norms.xlsx"
FV_NORM_SCATTER_PATH = Path(__file__).resolve().parent / "data" / "fv_norm_scatter.json"
UPLOADED_LOGOS_DIR = Path(__file__).resolve().parent / "uploaded_logos"
UPLOADED_PLAYER_PHOTOS_DIR = Path(__file__).resolve().parent / "uploaded_player_photos"
MAX_LOGO_BYTES = 2 * 1024 * 1024
MAX_PLAYER_PHOTO_BYTES = 5 * 1024 * 1024


PDF_TEXT = {
    "English": {
        "fv_title": "Force-Velocity Profile",
        "recommendation": "Recommendation",
        "notes": "Notes",
        "f0_status": "F0 vs norm",
        "v0_status": "V0 vs norm",
        "within": "Within",
        "below": "Below",
        "above": "Above",
        "no_norm": "No norm",
        "q1_result": "Faster acceleration / higher speed",
        "q2_result": "Faster acceleration / lower speed",
        "q3_result": "Slower acceleration / lower speed",
        "q4_result": "Slower acceleration / higher speed",
        "q1_recs": [
            "Acceleration and speed work (<7 s)",
            "Resisted sprint training (25-50% speed drop; 10-20 m)",
            "Flying sprints",
            "Assisted sprint training",
            "Speed bounding",
            "Improve hamstring isometric strength",
        ],
        "q2_recs": [
            "Acceleration and speed work (<7 s)",
            "Resisted sprint training (25-50% speed drop; 10-20 m)",
            "Flying sprints",
            "Assisted sprint training",
            "Improve stretch-shortening cycle (SSC)",
            "Improve reactive strength",
            "Improve connective tissue strength",
        ],
        "q3_recs": [
            "Resisted sprint training (50-75% speed drop; 10 m)",
            "Flying sprints",
            "Improve stretch-shortening cycle (SSC)",
            "Improve hip extensor strength",
            "Improve soleus and gastrocnemius strength",
            "Improve absolute and relative strength",
            "Improve reactive strength",
            "Improve rate of force development (RFD)",
        ],
        "q4_recs": [
            "Acceleration and speed work (<7 s)",
            "Resisted sprint training (50-75% speed drop; 10 m)",
            "Improve hip extensor strength",
            "Improve absolute and relative strength",
            "Improve rate of force development (RFD)",
            "Improve connective tissue strength",
        ],
    },
    "Slovak": {
        "fv_title": "Silovo-rychlostny profil",
        "recommendation": "Odporucanie",
        "notes": "Poznamka",
        "f0_status": "F0 vs norma",
        "v0_status": "V0 vs norma",
        "within": "V norme",
        "below": "Pod normou",
        "above": "Nad normou",
        "no_norm": "Bez normy",
        "q1_result": "Rychlejsia akceleracia / vyssia rychlost",
        "q2_result": "Rychlejsia akceleracia / nizsia rychlost",
        "q3_result": "Pomalsia akceleracia / nizsia rychlost",
        "q4_result": "Pomalsia akceleracia / vyssia rychlost",
        "q1_recs": [
            "Akceleracia/praca na rychlosti (<7 s)",
            "Sprintersky trening s odporom (25-50% pokles rychlosti; 10-20 m)",
            "Letme sprinty",
            "Sprintersky trening s asistenciou",
            "Speed bounding",
            "Zlepsenie izometrickej sily hamstringov",
        ],
        "q2_recs": [
            "Akceleracia/praca na rychlosti (<7 s)",
            "Sprintersky trening s odporom (25-50% pokles rychlosti; 10-20 m)",
            "Letme sprinty",
            "Sprintersky trening s asistenciou",
            "Zlepsenie cyklu natiahnutie-skratenie (SSC)",
            "Zlepsenie reaktivnej sily",
            "Zlepsenie sily spojivovych tkaniv",
        ],
        "q3_recs": [
            "Sprintersky trening s odporom (50-75% pokles rychlosti; 10 m)",
            "Letme sprinty",
            "Zlepsenie cyklu natiahnutia-skratenia (SSC)",
            "Zlepsenie sily extenzorov bedroveho klbu",
            "Zlepsenie sily soleusu a gastrocnemiusu",
            "Zlepsenie absolutnej/relativnej sily",
            "Zlepsenie reaktivnej sily",
            "Zlepsenie rychlosti produkcie sily (RFD)",
        ],
        "q4_recs": [
            "Akceleracia/praca na rychlosti (<7 s)",
            "Sprintersky trening s odporom (50-75% pokles rychlosti; 10 m)",
            "Zlepsenie sily extenzorov bedroveho klbu",
            "Zlepsenie absolutnej/relativnej sily",
            "Zlepsenie rychlosti produkcie sily (RFD)",
            "Zlepsenie sily spojivovych tkaniv",
        ],
    },
}

CLIENT_STORAGE_COMPONENT = st.components.v2.component(
    "client_storage",
    html="""
    <div id="client-storage-root"></div>
    """,
    js="""
    export default function(component) {
      const { data, setStateValue } = component;
      const storageKey = data?.storageKey ?? "1080_clients";
      const syncedKey = `${storageKey}:last_synced`;
      const currentCommandId = data?.commandId ?? "";
      const command = data?.command ?? "read";

      const emitState = () => {
        setStateValue("clients_json", localStorage.getItem(storageKey) ?? "");
        setStateValue("last_synced", localStorage.getItem(syncedKey) ?? "");
      };

      const commandMarkerKey = `__last_command__:${storageKey}`;
      const lastCommandId = window[commandMarkerKey];

      if (currentCommandId && currentCommandId !== lastCommandId) {
        if (command === "write") {
          localStorage.setItem(storageKey, data?.clientsJson ?? "");
          localStorage.setItem(syncedKey, data?.lastSynced ?? "");
        } else if (command === "clear") {
          localStorage.removeItem(storageKey);
          localStorage.removeItem(syncedKey);
        }

        window[commandMarkerKey] = currentCommandId;
      }

      emitState();
    }
    """,
)


def build_headers(api_key: str) -> dict[str, str]:
    return {
        "X-1080-API-Key": api_key,
        "Accept": "application/json",
    }


def validate_api_key(api_key: str) -> tuple[bool, str]:
    try:
        response = requests.get(
            f"{API_BASE_URL}/Session",
            headers=build_headers(api_key),
            params={"maxAgeDays": 0},
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return False, f"Communication error while connecting to the 1080 API: {exc}"

    if response.status_code == 200:
        return True, "API key is valid."

    if response.status_code in (401, 403):
        return False, "API key is invalid or does not have access to the requested data."

    return False, f"1080 API returned an unexpected response: {response.status_code}"


def fetch_clients(api_key: str) -> tuple[list[dict[str, Any]] | None, str | None]:
    try:
        response = requests.get(
            f"{API_BASE_URL}/Client",
            headers=build_headers(api_key),
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading clients: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, list):
            return payload, None
        return None, "The client response did not return the expected list format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load clients."

    return None, f"1080 API returned an unexpected response while loading clients: {response.status_code}"


def fetch_client_sessions(
    api_key: str,
    client_id: str,
    from_date: date,
    to_date: date,
) -> tuple[list[dict[str, Any]] | None, str | None]:
    from_dt = datetime.combine(from_date, time.min, tzinfo=timezone.utc)
    to_dt = datetime.combine(to_date, time.max, tzinfo=timezone.utc)

    try:
        response = requests.get(
            f"{API_BASE_URL}/Session/Search",
            headers=build_headers(api_key),
            params={
                "client": client_id,
                "fromDate": from_dt.isoformat(),
                "toDate": to_dt.isoformat(),
            },
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading sessions: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, list):
            return payload, None
        return None, "The session response did not return the expected list format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load sessions."

    return None, f"1080 API returned an unexpected response while loading sessions: {response.status_code}"


def fetch_recent_sessions(
    api_key: str,
    from_date: date,
    to_date: date,
) -> tuple[list[dict[str, Any]] | None, str | None]:
    from_dt = datetime.combine(from_date, time.min, tzinfo=timezone.utc)
    to_dt = datetime.combine(to_date, time.max, tzinfo=timezone.utc)

    try:
        response = requests.get(
            f"{API_BASE_URL}/Session/Search",
            headers=build_headers(api_key),
            params={
                "fromDate": from_dt.isoformat(),
                "toDate": to_dt.isoformat(),
            },
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading recent sessions: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, list):
            return payload, None
        return None, "The recent sessions response did not return the expected list format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load recent sessions."

    return None, f"1080 API returned an unexpected response while loading recent sessions: {response.status_code}"


def fetch_session_detail(api_key: str, session_id: str) -> tuple[dict[str, Any] | None, str | None]:
    try:
        response = requests.get(
            f"{API_BASE_URL}/Session/{session_id}",
            headers=build_headers(api_key),
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading session detail: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, dict):
            return payload, None
        return None, "The session detail response did not return the expected object format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load session detail."

    return None, f"1080 API returned an unexpected response while loading session detail: {response.status_code}"


def fetch_force_velocity_exercise(
    api_key: str, exercise_id: str
) -> tuple[dict[str, Any] | None, str | None]:
    try:
        response = requests.get(
            f"{API_BASE_URL}/ForceVelocity/Exercise/{exercise_id}",
            headers=build_headers(api_key),
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading FV profile: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, dict):
            return payload, None
        return None, "The FV response did not return the expected object format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load FV profile data."

    return None, f"1080 API returned an unexpected response while loading FV profile: {response.status_code}"


def fetch_split_exercise(
    api_key: str, exercise_id: str
) -> tuple[dict[str, Any] | None, str | None]:
    try:
        response = requests.get(
            f"{API_BASE_URL}/Split/Exercise/{exercise_id}",
            headers=build_headers(api_key),
            params={
                "splitLength": 5,
                "useYards": False,
                "includeRawPeaksAndAverages": True,
            },
            timeout=API_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        return None, f"Communication error while loading split data: {exc}"

    if response.status_code == 200:
        payload = response.json()
        if isinstance(payload, dict):
            return payload, None
        return None, "The split response did not return the expected object format."

    if response.status_code in (401, 403):
        return None, "The API key is no longer authorized to load split data."

    return None, f"1080 API returned an unexpected response while loading split data: {response.status_code}"


def storage_namespace(api_key: str) -> str:
    digest = hashlib.sha256(api_key.encode("utf-8")).hexdigest()[:12]
    return f"1080_clients:{digest}"


def normalize_text(text: str) -> str:
    return unicodedata.normalize("NFD", str(text)).encode("ascii", "ignore").decode("utf-8").lower().strip()


def safe_filename(value: str) -> str:
    value = normalize_text(value).replace(" ", "_")
    value = "".join(ch for ch in value if ch.isalnum() or ch in {"_", "-"})
    return value[:120] if value else "export"


def ensure_uploaded_logos_dir() -> None:
    UPLOADED_LOGOS_DIR.mkdir(exist_ok=True)


def list_uploaded_logos() -> list[str]:
    ensure_uploaded_logos_dir()
    valid_suffixes = {".png", ".jpg", ".jpeg", ".webp"}
    return sorted(
        [
            file.name
            for file in UPLOADED_LOGOS_DIR.iterdir()
            if file.is_file() and file.suffix.lower() in valid_suffixes
        ]
    )


def save_uploaded_logo(uploaded_file: Any) -> str:
    ensure_uploaded_logos_dir()
    original_name = getattr(uploaded_file, "name", "logo.png")
    stem = Path(original_name).stem
    suffix = Path(original_name).suffix.lower() or ".png"
    base_name = safe_filename(stem) or "logo"
    candidate = f"{base_name}{suffix}"
    counter = 1

    while (UPLOADED_LOGOS_DIR / candidate).exists():
        candidate = f"{base_name}_{counter}{suffix}"
        counter += 1

    target_path = UPLOADED_LOGOS_DIR / candidate
    target_path.write_bytes(uploaded_file.getvalue())
    return candidate


def load_saved_logo_bytes(file_name: str) -> bytes | None:
    ensure_uploaded_logos_dir()
    if not file_name:
        return None

    path = UPLOADED_LOGOS_DIR / file_name
    if not path.is_file():
        return None

    return path.read_bytes()


def ensure_uploaded_player_photos_dir() -> None:
    UPLOADED_PLAYER_PHOTOS_DIR.mkdir(exist_ok=True)


def save_player_photo(uploaded_file: Any, client_id: str) -> str:
    ensure_uploaded_player_photos_dir()
    suffix = Path(getattr(uploaded_file, "name", "photo.png")).suffix.lower() or ".png"
    base_name = safe_filename(client_id) or "client"

    for existing in UPLOADED_PLAYER_PHOTOS_DIR.glob(f"{base_name}.*"):
        if existing.is_file():
            existing.unlink()

    file_name = f"{base_name}{suffix}"
    target_path = UPLOADED_PLAYER_PHOTOS_DIR / file_name
    target_path.write_bytes(uploaded_file.getvalue())
    return file_name


def get_saved_player_photo_name(client_id: str) -> str | None:
    ensure_uploaded_player_photos_dir()
    base_name = safe_filename(client_id) or "client"
    for file in sorted(UPLOADED_PLAYER_PHOTOS_DIR.glob(f"{base_name}.*")):
        if file.is_file():
            return file.name
    return None


def load_player_photo_bytes(client_id: str) -> bytes | None:
    saved_name = get_saved_player_photo_name(client_id)
    if not saved_name:
        return None
    path = UPLOADED_PLAYER_PHOTOS_DIR / saved_name
    if not path.is_file():
        return None
    return path.read_bytes()


def validate_uploaded_file_size(uploaded_file: Any, max_bytes: int, label: str) -> bool:
    if uploaded_file is None:
        return True

    file_size = len(uploaded_file.getvalue())
    if file_size > max_bytes:
        max_mb = max_bytes / (1024 * 1024)
        st.error(f"{label} is too large. Maximum allowed size is {max_mb:.0f} MB.")
        return False

    return True


def render_logo_library_selector(key_prefix: str) -> bytes | None:
    saved_logos = list_uploaded_logos()
    selected_logo_state_key = f"{key_prefix}_selected_logo_name"
    pending_logo_state_key = f"{key_prefix}_pending_saved_logo"

    if selected_logo_state_key not in st.session_state:
        st.session_state[selected_logo_state_key] = "No saved logo"

    pending_logo_name = st.session_state.pop(pending_logo_state_key, None)
    if pending_logo_name:
        st.session_state[selected_logo_state_key] = pending_logo_name

    options = ["No saved logo", *saved_logos]
    if st.session_state[selected_logo_state_key] not in options:
        st.session_state[selected_logo_state_key] = "No saved logo"

    select_col, upload_col = st.columns([1, 1])
    selected_logo = select_col.selectbox(
        "Choose logo",
        options=options,
        key=selected_logo_state_key,
    )
    uploaded_logo = upload_col.file_uploader(
        "Upload new logo",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{key_prefix}_logo_upload",
        help="Maximum file size: 2 MB",
    )

    logo_is_valid = validate_uploaded_file_size(uploaded_logo, MAX_LOGO_BYTES, "Logo")

    if uploaded_logo is not None and logo_is_valid:
        if st.button("Save uploaded logo", key=f"{key_prefix}_save_logo", use_container_width=True):
            saved_name = save_uploaded_logo(uploaded_logo)
            st.session_state[pending_logo_state_key] = saved_name
            st.success(f"Saved logo: {saved_name}")
            st.rerun()

    if uploaded_logo is not None and logo_is_valid:
        return uploaded_logo.getvalue()

    if selected_logo != "No saved logo":
        return load_saved_logo_bytes(selected_logo)

    return None


def render_player_photo_selector(key_prefix: str, client_id: str) -> bytes | None:
    saved_photo = get_saved_player_photo_name(client_id)
    current_bytes = load_player_photo_bytes(client_id)

    if current_bytes:
        st.caption(f"Saved player photo: {saved_photo}")
        st.image(current_bytes, width=120)
        return current_bytes

    st.caption("No saved player photo")
    uploaded_photo = st.file_uploader(
        "Upload player photo",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=False,
        key=f"{key_prefix}_player_photo_upload",
        help="Maximum file size: 5 MB",
    )

    photo_is_valid = validate_uploaded_file_size(uploaded_photo, MAX_PLAYER_PHOTO_BYTES, "Player photo")
    if uploaded_photo is not None and photo_is_valid:
        saved_name = save_player_photo(uploaded_photo, client_id)
        st.success(f"Saved player photo: {saved_name}")
        st.rerun()

    return current_bytes


def iso_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def format_sync_label(value: str) -> str:
    if not value:
        return "Never"

    try:
        timestamp = datetime.fromisoformat(value)
    except ValueError:
        return value

    return timestamp.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")


def ensure_session_defaults() -> None:
    st.session_state.setdefault("auth_verified", False)
    st.session_state.setdefault("auth_user_email", "")
    st.session_state.setdefault("api_key", "")
    st.session_state.setdefault("api_valid", False)
    st.session_state.setdefault("clients_cache", [])
    st.session_state.setdefault("clients_last_synced", "")
    st.session_state.setdefault("client_storage_command", "read")
    st.session_state.setdefault("client_storage_command_id", "")
    st.session_state.setdefault("client_storage_payload", "")
    st.session_state.setdefault("client_storage_autoload_complete", False)
    st.session_state.setdefault("client_storage_error", "")
    st.session_state.setdefault("selected_client_id", "")
    st.session_state.setdefault("selected_client_last_id", "")
    st.session_state.setdefault("client_sessions", [])
    st.session_state.setdefault("client_sessions_error", "")
    st.session_state.setdefault("client_sessions_loaded_for", "")
    st.session_state.setdefault("session_filter_from", date.today() - timedelta(days=7))
    st.session_state.setdefault("session_filter_to", date.today())
    st.session_state.setdefault("selected_session_id", "")
    st.session_state.setdefault("session_detail", None)
    st.session_state.setdefault("session_detail_error", "")
    st.session_state.setdefault("exercise_report_cache", {})
    st.session_state.setdefault("exercise_report_errors", {})


def sync_clients_to_storage(clients: list[dict[str, Any]], last_synced: str) -> None:
    st.session_state["clients_cache"] = clients
    st.session_state["clients_last_synced"] = last_synced
    st.session_state["client_storage_command"] = "write"
    st.session_state["client_storage_command_id"] = f"write:{last_synced}"
    st.session_state["client_storage_payload"] = json.dumps(clients)
    st.session_state["client_storage_error"] = ""


def reset_client_cache_state() -> None:
    st.session_state["clients_cache"] = []
    st.session_state["clients_last_synced"] = ""
    st.session_state["client_storage_command"] = "clear"
    st.session_state["client_storage_command_id"] = f"clear:{iso_now()}"
    st.session_state["client_storage_payload"] = ""
    st.session_state["client_storage_autoload_complete"] = False
    st.session_state["client_storage_error"] = ""
    st.session_state["selected_client_id"] = ""
    st.session_state["selected_client_last_id"] = ""
    st.session_state["client_sessions"] = []
    st.session_state["client_sessions_error"] = ""
    st.session_state["client_sessions_loaded_for"] = ""
    st.session_state["session_filter_from"] = date.today() - timedelta(days=7)
    st.session_state["session_filter_to"] = date.today()
    st.session_state["selected_session_id"] = ""
    st.session_state["session_detail"] = None
    st.session_state["session_detail_error"] = ""
    st.session_state["exercise_report_cache"] = {}
    st.session_state["exercise_report_errors"] = {}


def logout() -> None:
    st.session_state["auth_verified"] = False
    st.session_state["auth_user_email"] = ""
    st.session_state["api_key"] = ""
    st.session_state["api_valid"] = False
    reset_client_cache_state()


def get_auth_users() -> dict[str, str]:
    users = st.secrets.get("auth_users", {})
    return dict(users) if isinstance(users, Mapping) else {}


def verify_password(password: str, stored_hash: str) -> bool:
    try:
        algorithm, iterations_str, salt, expected_hex = stored_hash.split("$", 3)
        if algorithm != "pbkdf2_sha256":
            return False
        iterations = int(iterations_str)
    except (ValueError, TypeError):
        return False

    derived = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        iterations,
    ).hex()
    return hmac.compare_digest(derived, expected_hex)


def render_login() -> None:
    _, center_col, _ = st.columns([1.2, 1, 1.2])
    with center_col:
        st.title("1080 Reports")
        st.write("Sign in with your email and password to access reporting tools.")

        auth_users = get_auth_users()
        if not auth_users:
            st.error("No app users are configured. Add `auth_users` to Streamlit secrets.")
            return
        if not APP_API_KEY:
            st.error("1080 API key is not configured. Set `api_1080_key` in Streamlit secrets or `API1080_KEY` as an environment variable.")
            return

        with st.form("auth-login-form"):
            email = st.text_input("Email").strip().lower()
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Sign in", use_container_width=True)

        if submitted:
            stored_hash = auth_users.get(email)
            if not stored_hash or not verify_password(password, str(stored_hash)):
                st.error("Incorrect email or password.")
                return

            with st.spinner("Validating 1080 API access..."):
                is_valid, message = validate_api_key(APP_API_KEY)

            if is_valid:
                st.session_state["auth_verified"] = True
                st.session_state["auth_user_email"] = email
                st.session_state["api_key"] = APP_API_KEY
                st.session_state["api_valid"] = True
                st.session_state["client_storage_autoload_complete"] = False
                st.rerun()

            st.session_state["api_valid"] = False
            st.error(message)


def load_clients_from_api(api_key: str) -> bool:
    clients, error = fetch_clients(api_key)
    if error:
        st.session_state["client_storage_error"] = error
        return False

    sync_clients_to_storage(clients or [], iso_now())
    return True


def format_optional_value(value: Any) -> str:
    if value in (None, "", []):
        return "-"
    if isinstance(value, list):
        return ", ".join(str(item) for item in value) if value else "-"
    return str(value)


def format_session_timestamp(value: Any) -> str:
    if not value:
        return "-"

    try:
        normalized = str(value).replace("Z", "+00:00")
        timestamp = datetime.fromisoformat(normalized)
    except ValueError:
        return str(value)

    return timestamp.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")


def get_client_lookup(clients: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    return {str(client.get("id") or ""): client for client in clients}


def get_client_display_name(client_id: Any, client_lookup: dict[str, dict[str, Any]]) -> str:
    client = client_lookup.get(str(client_id or ""))
    if not client:
        return "-"
    return format_optional_value(client.get("displayName"))


def has_missing_client_names(sessions: list[dict[str, Any]], client_lookup: dict[str, dict[str, Any]]) -> bool:
    for session in sessions:
        client_id = str(session.get("clientId") or "")
        if client_id and client_id not in client_lookup:
            return True
    return False


def ensure_session_scope_defaults(scope: str) -> None:
    st.session_state.setdefault(f"selected_session_id_{scope}", "")
    st.session_state.setdefault(f"session_detail_{scope}", None)
    st.session_state.setdefault(f"session_detail_error_{scope}", "")


def clear_session_scope(scope: str) -> None:
    st.session_state[f"selected_session_id_{scope}"] = ""
    st.session_state[f"session_detail_{scope}"] = None
    st.session_state[f"session_detail_error_{scope}"] = ""
    st.session_state["exercise_report_cache"] = {}
    st.session_state["exercise_report_errors"] = {}


def load_sessions_for_selected_client() -> bool:
    selected_client_id = st.session_state.get("selected_client_id", "")
    api_key = st.session_state.get("api_key", "")
    from_date = st.session_state.get("session_filter_from")
    to_date = st.session_state.get("session_filter_to")

    if not selected_client_id or not api_key:
        return False

    sessions, error = fetch_client_sessions(api_key, selected_client_id, from_date, to_date)
    if error:
        st.session_state["client_sessions"] = []
        st.session_state["client_sessions_error"] = error
        st.session_state["client_sessions_loaded_for"] = selected_client_id
        return False

    st.session_state["client_sessions"] = sessions or []
    st.session_state["client_sessions_error"] = ""
    st.session_state["client_sessions_loaded_for"] = selected_client_id
    clear_session_scope("client")
    return True


def load_selected_session_detail(scope: str) -> bool:
    api_key = st.session_state.get("api_key", "")
    session_id = st.session_state.get(f"selected_session_id_{scope}", "")

    if not api_key or not session_id:
        return False

    session_detail, error = fetch_session_detail(api_key, session_id)
    if error:
        st.session_state[f"session_detail_{scope}"] = None
        st.session_state[f"session_detail_error_{scope}"] = error
        return False

    st.session_state[f"session_detail_{scope}"] = session_detail
    st.session_state[f"session_detail_error_{scope}"] = ""
    st.session_state["exercise_report_cache"] = {}
    st.session_state["exercise_report_errors"] = {}
    return True


def format_decimal(value: Any, digits: int = 2) -> str:
    try:
        return f"{float(value):.{digits}f}"
    except (TypeError, ValueError):
        return "-"


def rounded_corner_cell(pdf: FPDF, x: float, y: float, w: float, h: float, text: str) -> None:
    pdf.set_xy(x, y)
    pdf.rect(x, y, w, h, round_corners=True, style="F")
    pdf.cell(w, h, text, 0, 0, "C")


def configure_pdf_font(pdf: FPDF) -> str:
    font_path = font_manager.findfont("DejaVu Sans")
    if font_path and os.path.isfile(font_path):
        pdf.add_font("DejaVuSans", "", font_path)
        return "DejaVuSans"
    return "Helvetica"


def make_fv_profile_player_only(report: dict[str, Any]) -> io.BytesIO:
    v0 = float(report["v0"])
    f0 = float(report["f0"])

    x_player = [0, v0]
    y_player = [f0, 0]

    bbox = dict(boxstyle="round", edgecolor="none", facecolor=BLUE_HEX)

    fig, ax = plt.subplots()
    ax.plot(x_player, y_player, label="Player", color=BLUE_HEX, linewidth=2.5)

    ax.annotate(
        str(round(v0, 2)),
        (v0, 0),
        xytext=(v0, -0.6),
        textcoords="data",
        ha="center",
        va="center",
        color="white",
        bbox=bbox,
    )
    ax.annotate(
        str(round(f0, 2)),
        (0, f0),
        xytext=(-0.6, f0),
        textcoords="data",
        ha="center",
        va="center",
        color="white",
        bbox=bbox,
    )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_xlabel("V0 [m/s]")
    ax.set_ylabel("F0 [N/kg]")
    ax.legend(loc="upper right", frameon=False)
    ax.margins(x=0.05, y=0.05)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def make_norm_scatter_plot(report: dict[str, Any], norm_row: dict[str, Any], scatter_entry: dict[str, Any]) -> io.BytesIO:
    points = scatter_entry.get("points") or []
    x_values = [float(point["v0"]) for point in points]
    y_values = [float(point["f0"]) for point in points]
    player_v0 = float(report["v0"])
    player_f0 = float(report["f0"])
    v0_median = float(norm_row["v0_median"])
    f0_median = float(norm_row["f0_median"])

    all_x = [*x_values, player_v0]
    all_y = [*y_values, player_f0]
    x_span = max(all_x) - min(all_x) if all_x else 1.0
    y_span = max(all_y) - min(all_y) if all_y else 1.0
    x_pad = max(x_span * 0.08, 0.15)
    y_pad = max(y_span * 0.08, 0.15)

    fig, ax = plt.subplots()
    if x_values and y_values:
        ax.scatter(x_values, y_values, color="#d0d4db", s=32, alpha=0.9, edgecolors="none")

    ax.scatter([player_v0], [player_f0], color="#FB3331", s=80, zorder=3)
    ax.axhline(f0_median, color="#252423", linestyle="--", linewidth=1.2)
    ax.axvline(v0_median, color="#252423", linestyle="--", linewidth=1.2)

    ax.text(v0_median + x_pad * 0.15, max(all_y) + y_pad * 0.1, "Q1", fontsize=10, color="#252423")
    ax.text(min(all_x) - x_pad * 0.1, max(all_y) + y_pad * 0.1, "Q2", fontsize=10, color="#252423")
    ax.text(min(all_x) - x_pad * 0.1, min(all_y) - y_pad * 0.35, "Q3", fontsize=10, color="#252423")
    ax.text(v0_median + x_pad * 0.15, min(all_y) - y_pad * 0.35, "Q4", fontsize=10, color="#252423")

    ax.set_xlim(min(all_x) - x_pad, max(all_x) + x_pad)
    ax.set_ylim(min(all_y) - y_pad, max(all_y) + y_pad)
    ax.set_xlabel("V0 [m/s]")
    ax.set_ylabel("F0 [N/kg]")
    ax.set_title(f"Quadrant reference | {norm_row.get('category')}")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.18)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def build_non_normative_fv_pdf(
    report: dict[str, Any],
    player_name: str,
    logo_bytes: bytes | None,
    player_photo_bytes: bytes | None,
    language: str,
) -> bytes:
    texts = PDF_TEXT[language]
    fv_buf = make_fv_profile_player_only(report)
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)

    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=10, y=12, w=25)

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 32)
    pdf.set_xy(45, 18)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")

    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 14)
    pdf.set_xy(45, 34)
    pdf.cell(0, 10, texts["fv_title"], new_x="LMARGIN", new_y="NEXT")

    pdf.image(fv_buf, x=20, y=55, w=140)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=194, y=18, w=45, h=45, keep_aspect_ratio=True)

    data_x = 170
    data_y = 70
    y_second_row = 18
    cell_w = 28
    cell_h = 12
    cell_h_sub = 7

    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)

    f0 = float(report["f0"])
    v0 = float(report["v0"])
    pmax = float(report["pMax"])
    drf = float(report.get("ratioOfForceDecrease") or 0)
    rfmax = float(report.get("ratioOfForceMax") or 0)

    rounded_corner_cell(pdf, data_x, data_y, cell_w, cell_h, str(round(f0, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y, cell_w, cell_h, str(round(v0, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y, cell_w, cell_h, str(round(pmax, 2)))
    rounded_corner_cell(pdf, data_x, data_y + y_second_row, cell_w, cell_h, str(round(v0 * 3.6, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y + y_second_row, cell_w, cell_h, str(round(drf, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y + y_second_row, cell_w, cell_h, str(round(rfmax, 2)))

    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, data_y + 9, cell_w, cell_h_sub, "F0 [N/kg]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9, cell_w, cell_h_sub, "V0 [m/s]")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9, cell_w, cell_h_sub, "PMax [W]")
    rounded_corner_cell(pdf, data_x, data_y + 9 + y_second_row, cell_w, cell_h_sub, "V0 [km/h]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9 + y_second_row, cell_w, cell_h_sub, "DRF")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9 + y_second_row, cell_w, cell_h_sub, "RFmax")

    return bytes(pdf.output(dest="S"))


def load_fv_norms() -> tuple[list[dict[str, Any]], str | None]:
    if not FV_NORMS_PATH.is_file():
        return [], f"Norms file not found: {FV_NORMS_PATH}"

    try:
        workbook = load_workbook(FV_NORMS_PATH, data_only=True)
    except Exception as exc:
        return [], f"Could not read norms file: {exc}"

    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], "The norms workbook is empty."

    headers = [str(cell).strip() if cell is not None else "" for cell in rows[0]]
    required = ["category", "f0_median", "v0_median"]
    missing = [name for name in required if name not in headers]
    if missing:
        return [], f"The norms workbook is missing columns: {', '.join(missing)}"

    items: list[dict[str, Any]] = []
    for row in rows[1:]:
        item = {headers[idx]: row[idx] if idx < len(row) else None for idx in range(len(headers))}
        if item.get("category"):
            items.append(item)

    return items, None


def load_fv_norm_scatter() -> tuple[dict[str, dict[str, Any]], str | None]:
    if not FV_NORM_SCATTER_PATH.is_file():
        return {}, f"Norm scatter file not found: {FV_NORM_SCATTER_PATH}"

    try:
        payload = json.loads(FV_NORM_SCATTER_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        return {}, f"Could not read norm scatter file: {exc}"

    if not isinstance(payload, dict):
        return {}, "The norm scatter file did not return the expected object format."

    return payload, None


def get_norm_quadrant(report: dict[str, Any], norm_row: dict[str, Any]) -> str:
    f0 = float(report["f0"])
    v0 = float(report["v0"])
    f0_mid = float(norm_row["f0_median"])
    v0_mid = float(norm_row["v0_median"])

    if v0 > v0_mid and f0 > f0_mid:
        return "Q1"
    if v0 <= v0_mid and f0 > f0_mid:
        return "Q2"
    if v0 <= v0_mid and f0 <= f0_mid:
        return "Q3"
    return "Q4"


def get_quadrant_result(quadrant: str, language: str) -> str:
    texts = PDF_TEXT[language]
    return texts[f"{quadrant.lower()}_result"]


def get_quadrant_recommendations(quadrant: str, language: str) -> list[str]:
    texts = PDF_TEXT[language]
    return texts[f"{quadrant.lower()}_recs"]


def build_normative_fv_pdf(
    report: dict[str, Any],
    player_name: str,
    logo_bytes: bytes | None,
    player_photo_bytes: bytes | None,
    norm_row: dict[str, Any],
    scatter_entry: dict[str, Any] | None,
    language: str,
) -> bytes:
    texts = PDF_TEXT[language]
    fv_buf = (
        make_norm_scatter_plot(report, norm_row, scatter_entry)
        if scatter_entry
        else make_fv_profile_player_only(report)
    )
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)

    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=10, y=12, w=25)

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 30)
    pdf.set_xy(45, 18)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")

    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 14)
    pdf.set_xy(45, 34)
    pdf.cell(0, 10, f"{texts['fv_title']} | Norm: {norm_row.get('category')}", new_x="LMARGIN", new_y="NEXT")

    quadrant = get_norm_quadrant(report, norm_row)
    quadrant_result = get_quadrant_result(quadrant, language)
    recommendations = get_quadrant_recommendations(quadrant, language)

    pdf.set_xy(45, 46)
    pdf.set_font(font_family, "", 16)
    pdf.set_text_color(*BLACK_RGB)
    pdf.cell(0, 10, quadrant_result, new_x="LMARGIN", new_y="NEXT")

    pdf.image(fv_buf, x=20, y=55, w=125)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=194, y=18, w=45, h=45, keep_aspect_ratio=True)

    f0 = float(report["f0"])
    v0 = float(report["v0"])
    pmax = float(report["pMax"])
    drf = float(report.get("ratioOfForceDecrease") or 0)
    rfmax = float(report.get("ratioOfForceMax") or 0)

    data_x = 170
    data_y = 70
    y_second_row = 18
    cell_w = 28
    cell_h = 12
    cell_h_sub = 7

    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)

    rounded_corner_cell(pdf, data_x, data_y, cell_w, cell_h, str(round(f0, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y, cell_w, cell_h, str(round(v0, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y, cell_w, cell_h, str(round(pmax, 2)))
    rounded_corner_cell(pdf, data_x, data_y + y_second_row, cell_w, cell_h, str(round(v0 * 3.6, 2)))
    rounded_corner_cell(pdf, data_x + 32, data_y + y_second_row, cell_w, cell_h, str(round(drf, 2)))
    rounded_corner_cell(pdf, data_x + 64, data_y + y_second_row, cell_w, cell_h, str(round(rfmax, 2)))

    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, data_y + 9, cell_w, cell_h_sub, "F0 [N/kg]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9, cell_w, cell_h_sub, "V0 [m/s]")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9, cell_w, cell_h_sub, "PMax [W]")
    rounded_corner_cell(pdf, data_x, data_y + 9 + y_second_row, cell_w, cell_h_sub, "V0 [km/h]")
    rounded_corner_cell(pdf, data_x + 32, data_y + 9 + y_second_row, cell_w, cell_h_sub, "DRF")
    rounded_corner_cell(pdf, data_x + 64, data_y + 9 + y_second_row, cell_w, cell_h_sub, "RFmax")

    pdf.set_font(font_family, "", 16)
    pdf.set_text_color(*BLACK_RGB)
    pdf.text(170, 135, texts["recommendation"])
    pdf.set_font(font_family, "", 10)
    y_coordinate = 142
    for item in recommendations:
        pdf.text(170, y_coordinate, "* " + item)
        y_coordinate += 8

    return bytes(pdf.output(dest="S"))


def load_exercise_report(api_key: str, exercise_id: str, report_type: str) -> bool:
    cache_key = f"{report_type}:{exercise_id}"

    if report_type == "fv":
        payload, error = fetch_force_velocity_exercise(api_key, exercise_id)
    else:
        payload, error = fetch_split_exercise(api_key, exercise_id)

    if error:
        st.session_state["exercise_report_errors"][cache_key] = error
        st.session_state["exercise_report_cache"].pop(cache_key, None)
        return False

    st.session_state["exercise_report_cache"][cache_key] = payload
    st.session_state["exercise_report_errors"].pop(cache_key, None)
    return True


@st.dialog("FV PDF export")
def render_fv_export_dialog(
    exercise: dict[str, Any],
    reports: list[dict[str, Any]],
    runner_info: dict[str, Any] | None,
    client: dict[str, Any],
) -> None:
    run_options = {
        f"{report.get('motionGroupId')} | F0 {format_decimal(report.get('f0'))} | V0 {format_decimal(report.get('v0'))}": report
        for report in reports
    }

    selected_run_label = st.selectbox(
        "Select run for PDF export",
        options=list(run_options.keys()),
        key=f"fv_run_select_{exercise.get('id')}",
    )
    selected_report = run_options[selected_run_label]

    export_col1, export_col2 = st.columns([1, 1])
    export_mode = export_col1.radio(
        "PDF mode",
        options=["Non-normative", "Normative"],
        horizontal=True,
        key=f"fv_export_mode_{exercise.get('id')}",
    )
    export_language = st.selectbox(
        "PDF language",
        options=["English", "Slovak"],
        key=f"fv_export_language_{exercise.get('id')}",
    )
    logo_bytes = render_logo_library_selector(f"fv_logo_{exercise.get('id')}")

    selected_runner = selected_report.get("runnerInfo") or runner_info or {}
    export_name = (
        format_optional_value(selected_runner.get("displayName"))
        if format_optional_value(selected_runner.get("displayName")) != "-"
        else format_optional_value(client.get("displayName"))
    )
    player_client_id = str(
        selected_runner.get("clientId")
        or client.get("id")
        or ""
    )

    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
    metric_col1.metric("Selected F0", format_decimal(selected_report.get("f0")))
    metric_col2.metric("Selected V0", format_decimal(selected_report.get("v0")))
    metric_col3.metric("Selected PMax", format_decimal(selected_report.get("pMax")))
    metric_col4.metric("Selected Confidence", format_decimal(selected_report.get("confidence"), 3))

    player_photo_bytes = render_player_photo_selector(
        f"fv_player_photo_{exercise.get('id')}",
        player_client_id,
    ) if player_client_id else None

    if export_mode == "Normative":
        norms, norms_error = load_fv_norms()
        scatter_map, scatter_error = load_fv_norm_scatter()
        if norms_error:
            st.error(norms_error)
            return
        if scatter_error:
            st.warning(scatter_error)
            scatter_map = {}

        norm_options = {str(item.get("category")): item for item in norms}
        selected_norm_category = st.selectbox(
            "Select quadrant reference category",
            options=list(norm_options.keys()),
            key=f"fv_norm_category_{exercise.get('id')}",
        )
        selected_norm = norm_options[selected_norm_category]

        reference_col1, reference_col2, reference_col3 = st.columns(3)
        reference_col1.metric(
            "Reference F0 median",
            format_decimal(selected_norm.get("f0_median")),
        )
        reference_col2.metric(
            "Reference V0 median",
            format_decimal(selected_norm.get("v0_median")),
        )
        reference_col3.metric(
            "Reference sample",
            format_optional_value(selected_norm.get("used_n") or selected_norm.get("raw_n")),
        )
        scatter_entry = scatter_map.get(selected_norm_category)
        quadrant = get_norm_quadrant(selected_report, selected_norm)
        st.write(f"**Quadrant result:** {get_quadrant_result(quadrant, export_language)}")
        st.write("**Recommendation:**")
        for item in get_quadrant_recommendations(quadrant, export_language):
            st.write(f"- {item}")

        missing_values = [
            key
            for key in ("f0_median", "v0_median")
            if selected_norm.get(key) in (None, "")
        ]
        if missing_values:
            st.warning(
                "Selected norm category is missing values in the workbook: "
                + ", ".join(missing_values)
            )
            return

            pdf_bytes = build_normative_fv_pdf(
                selected_report,
                export_name,
                logo_bytes,
                player_photo_bytes,
                selected_norm,
                scatter_entry,
                export_language,
            )
        file_name = (
            f"{safe_filename(export_name)}_{safe_filename(selected_norm_category)}_"
            f"{safe_filename(str(exercise.get('id')))}_fv_normative.pdf"
        )
        st.download_button(
            "Download normative FV PDF",
            data=pdf_bytes,
            file_name=file_name,
            mime="application/pdf",
            key=f"fv_pdf_normative_{exercise.get('id')}",
            use_container_width=True,
        )
    else:
        pdf_bytes = build_non_normative_fv_pdf(
            selected_report,
            export_name,
            logo_bytes,
            player_photo_bytes,
            export_language,
        )
        file_name = f"{safe_filename(export_name)}_{safe_filename(str(exercise.get('id')))}_fv_profile.pdf"
        st.download_button(
            "Download FV PDF",
            data=pdf_bytes,
            file_name=file_name,
            mime="application/pdf",
            key=f"fv_pdf_download_{exercise.get('id')}",
            use_container_width=True,
        )


def render_fv_profile(exercise: dict[str, Any], payload: dict[str, Any], client: dict[str, Any]) -> None:
    reports = payload.get("reports") or []
    failed_reports = payload.get("failedReports") or []
    runner_info = reports[0].get("runnerInfo") if reports else None

    st.markdown(f"**Running (LR) FV profile**  `{format_optional_value(exercise.get('id'))}`")

    top_col1, top_col2, top_col3 = st.columns(3)
    top_col1.metric("Valid runs", len(reports))
    top_col2.metric("Failed runs", len(failed_reports))
    top_col3.metric("Runner", format_optional_value((runner_info or {}).get("displayName")))

    if reports:
        summary_rows = [
            {
                "motionGroupId": report.get("motionGroupId"),
                "f0": format_decimal(report.get("f0")),
                "v0": format_decimal(report.get("v0")),
                "pMax": format_decimal(report.get("pMax")),
                "tau": format_decimal(report.get("tau")),
                "confidence": format_decimal(report.get("confidence"), 3),
                "estimatedUnloadedMaxSpeed": format_decimal(report.get("estimatedUnloadedMaxSpeed")),
                "ratioOfForceMax": format_decimal(report.get("ratioOfForceMax")),
                "ratioOfForceDecrease": format_decimal(report.get("ratioOfForceDecrease")),
            }
            for report in reports
        ]
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)
        st.caption("PDF export options are now grouped in a separate popup.")
        if st.button("Open FV PDF export", key=f"fv_export_open_{exercise.get('id')}", use_container_width=True):
            render_fv_export_dialog(exercise, reports, runner_info, client)
    else:
        st.info("No FV runs were returned for this exercise.")

    if failed_reports:
        st.caption("Failed FV reports")
        st.dataframe(failed_reports, use_container_width=True, hide_index=True)

    with st.expander("Raw FV payload"):
        st.json(payload)


def render_split_profile(exercise: dict[str, Any], payload: dict[str, Any]) -> None:
    reports = payload.get("reports") or []

    st.markdown(f"**15-0-5 split profile**  `{format_optional_value(exercise.get('id'))}`")

    top_col1, top_col2, top_col3 = st.columns(3)
    top_col1.metric("Runs", len(reports))
    top_col2.metric("Split length", format_decimal(reports[0].get("splitLength")) if reports else "5.00")
    top_col3.metric("Units", "meters" if reports and not reports[0].get("isYards") else "yards")

    split_rows: list[dict[str, Any]] = []
    for report in reports:
        for split in report.get("splits") or []:
            split_rows.append(
                {
                    "motionGroupId": report.get("motionGroupId"),
                    "start": format_decimal(split.get("start")),
                    "end": format_decimal(split.get("end")),
                    "time": format_decimal(split.get("time"), 3),
                    "topSpeed": format_decimal(split.get("topSpeed")),
                    "maxForce": format_decimal(split.get("maxForce")),
                    "maxPower": format_decimal(split.get("maxPower")),
                }
            )

    if split_rows:
        st.dataframe(split_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No split rows were returned for this exercise.")

    with st.expander("Raw split payload"):
        st.json(payload)


def render_session_detail_content(
    session_detail: dict[str, Any],
    client_lookup: dict[str, dict[str, Any]],
) -> None:
    exercises = session_detail.get("exercises") or []
    client = client_lookup.get(str(session_detail.get("clientId") or ""), {})

    st.subheader("Session detail")

    detail_col1, detail_col2, detail_col3, detail_col4 = st.columns(4)
    detail_col1.metric("Session ID", format_optional_value(session_detail.get("id")))
    detail_col2.metric("Exercises", len(exercises))
    detail_col3.metric("Client ID", format_optional_value(session_detail.get("clientId")))
    detail_col4.metric("Client name", get_client_display_name(session_detail.get("clientId"), client_lookup))

    if exercises:
        exercise_rows = [
            {
                "exerciseId": exercise.get("id"),
                "exerciseTypeName": exercise.get("exerciseTypeName"),
                "sets": len(exercise.get("sets") or []),
            }
            for exercise in exercises
        ]
        st.dataframe(exercise_rows, use_container_width=True, hide_index=True)

        api_key = st.session_state.get("api_key", "")
        fv_exercises = [
            exercise
            for exercise in exercises
            if str(exercise.get("exerciseTypeName") or "").strip() == "Running (LR)"
        ]
        split_exercises = [
            exercise
            for exercise in exercises
            if str(exercise.get("exerciseTypeName") or "").strip() == "15-0-5"
        ]

        for exercise in fv_exercises:
            cache_key = f"fv:{exercise.get('id')}"
            if cache_key not in st.session_state["exercise_report_cache"] and api_key:
                with st.spinner("Loading FV profile data..."):
                    load_exercise_report(api_key, str(exercise.get("id") or ""), "fv")

            error = st.session_state["exercise_report_errors"].get(cache_key)
            if error:
                st.error(error)
            else:
                payload = st.session_state["exercise_report_cache"].get(cache_key)
                if payload:
                    render_fv_profile(exercise, payload, client)

        for exercise in split_exercises:
            cache_key = f"split:{exercise.get('id')}"
            if cache_key not in st.session_state["exercise_report_cache"] and api_key:
                with st.spinner("Loading deceleration split data..."):
                    load_exercise_report(api_key, str(exercise.get("id") or ""), "split")

            error = st.session_state["exercise_report_errors"].get(cache_key)
            if error:
                st.error(error)
            else:
                payload = st.session_state["exercise_report_cache"].get(cache_key)
                if payload:
                    render_split_profile(exercise, payload)
    else:
        st.info("No exercises were returned for this session.")

    with st.expander("Raw session payload"):
        st.json(session_detail)


def render_session_selection_block(
    sessions: list[dict[str, Any]],
    client_lookup: dict[str, dict[str, Any]],
    scope: str,
    table_key: str,
) -> None:
    ensure_session_scope_defaults(scope)

    session_rows = [
        {
            "timestamp": format_session_timestamp(session.get("timestamp")),
            "clientName": get_client_display_name(session.get("clientId"), client_lookup),
            "sessionId": session.get("id"),
            "clientId": session.get("clientId"),
        }
        for session in sessions
    ]

    session_event = st.dataframe(
        session_rows,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key=table_key,
    )

    selected_rows = session_event.selection.rows
    if selected_rows:
        selected_session = sessions[selected_rows[0]]
        selected_session_id = str(selected_session.get("id") or "")
        if selected_session_id != st.session_state.get(f"selected_session_id_{scope}", ""):
            st.session_state[f"selected_session_id_{scope}"] = selected_session_id
            with st.spinner("Loading session detail..."):
                load_selected_session_detail(scope)

    session_detail_error = st.session_state.get(f"session_detail_error_{scope}", "")
    if session_detail_error:
        st.error(session_detail_error)

    session_detail = st.session_state.get(f"session_detail_{scope}")
    if session_detail:
        render_session_detail_content(session_detail, client_lookup)


def render_client_detail(client: dict[str, Any]) -> None:
    st.subheader("Client detail")

    top_col1, top_col2, top_col3 = st.columns(3)
    top_col1.metric("Name", format_optional_value(client.get("displayName")))
    top_col2.metric("Group", format_optional_value(client.get("group")))
    top_col3.metric("External ID", format_optional_value(client.get("externalId")))

    info_col1, info_col2 = st.columns(2)
    with info_col1:
        st.write(f"**Client ID:** {format_optional_value(client.get('id'))}")
        st.write(f"**Date of birth:** {format_optional_value(client.get('dateOfBirth'))}")
        st.write(f"**Height:** {format_optional_value(client.get('height'))}")
        st.write(f"**Weight:** {format_optional_value(client.get('weight'))}")
    with info_col2:
        st.write(f"**Created:** {format_optional_value(client.get('created'))}")
        st.write(f"**Edited:** {format_optional_value(client.get('edited'))}")
        st.write(f"**Tags:** {format_optional_value(client.get('tags'))}")

    with st.expander("Raw client payload"):
        st.json(client)

    st.subheader("Client sessions")

    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 1])
    filter_col1.date_input("From", key="session_filter_from")
    filter_col2.date_input("To", key="session_filter_to")
    reload_sessions = filter_col3.button("Load sessions", use_container_width=True)

    if st.session_state["session_filter_from"] > st.session_state["session_filter_to"]:
        st.error("The start date must be earlier than or equal to the end date.")
        return

    if reload_sessions:
        with st.spinner("Loading sessions for the selected client..."):
            load_sessions_for_selected_client()

    if st.session_state["client_sessions_error"]:
        st.error(st.session_state["client_sessions_error"])

    sessions = st.session_state.get("client_sessions", [])
    st.caption(
        f"Showing {len(sessions)} sessions from "
        f"{st.session_state['session_filter_from'].isoformat()} to {st.session_state['session_filter_to'].isoformat()}"
    )

    if sessions:
        render_session_selection_block(
            sessions=sessions,
            client_lookup=get_client_lookup(st.session_state.get("clients_cache", [])),
            scope="client",
            table_key="client_sessions_table",
        )
    else:
        st.info("No sessions found for the selected range.")

def render_dashboard() -> None:
    api_key = st.session_state["api_key"]
    storage_key = storage_namespace(api_key)
    auth_user_email = st.session_state.get("auth_user_email", "")

    storage_result = CLIENT_STORAGE_COMPONENT(
        data={
            "storageKey": storage_key,
            "command": st.session_state["client_storage_command"],
            "commandId": st.session_state["client_storage_command_id"],
            "clientsJson": st.session_state["client_storage_payload"],
            "lastSynced": st.session_state["clients_last_synced"],
        },
        default={"clients_json": "", "last_synced": ""},
        on_clients_json_change=lambda: None,
        on_last_synced_change=lambda: None,
        key="client_storage_bridge",
        height=0,
    )

    cached_clients_json = storage_result.clients_json or ""
    cached_last_synced = storage_result.last_synced or ""

    if (
        not st.session_state["client_storage_autoload_complete"]
        and cached_clients_json
        and not st.session_state["clients_cache"]
    ):
        try:
            st.session_state["clients_cache"] = json.loads(cached_clients_json)
            st.session_state["clients_last_synced"] = cached_last_synced
            st.session_state["client_storage_command"] = "read"
            st.session_state["client_storage_payload"] = cached_clients_json
            st.session_state["client_storage_autoload_complete"] = True
            st.rerun()
        except json.JSONDecodeError:
            st.session_state["client_storage_error"] = "Client cache in local storage is corrupted."
            st.session_state["client_storage_autoload_complete"] = True

    if (
        not st.session_state["client_storage_autoload_complete"]
        and not cached_clients_json
        and not st.session_state["clients_cache"]
    ):
        with st.spinner("Loading clients from the 1080 API..."):
            loaded = load_clients_from_api(api_key)
        st.session_state["client_storage_autoload_complete"] = True
        if loaded:
            st.rerun()

    st.title("1080 Reports")
    st.success("Signed in successfully via the 1080 API.")
    if auth_user_email:
        st.caption(f"Signed in as {auth_user_email}")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Status", "Connected")
    col2.metric("API", "1080 Motion")
    col3.metric("Clients", len(st.session_state["clients_cache"]))
    col4.metric("Last synced", format_sync_label(st.session_state["clients_last_synced"]))

    action_col1, action_col2 = st.columns([1, 1])
    if action_col1.button("Reload clients", use_container_width=True):
        with st.spinner("Reloading clients from the 1080 API..."):
            loaded = load_clients_from_api(api_key)
        st.session_state["client_storage_autoload_complete"] = True
        if loaded:
            st.rerun()

    action_col2.button("Sign out", on_click=logout, use_container_width=True)

    if st.session_state["client_storage_error"]:
        st.error(st.session_state["client_storage_error"])

    st.subheader("Next step")
    st.write(
        "Clients are cached in browser local storage. If no cache is present, the app fetches them automatically."
    )

    st.subheader("Recent sessions")
    recent_from = date.today() - timedelta(days=7)
    recent_to = date.today()
    with st.spinner("Loading sessions from the last 7 days..."):
        recent_sessions, recent_sessions_error = fetch_recent_sessions(api_key, recent_from, recent_to)

    if recent_sessions_error:
        st.error(recent_sessions_error)
    else:
        client_lookup = get_client_lookup(st.session_state.get("clients_cache", []))
        if recent_sessions and has_missing_client_names(recent_sessions, client_lookup):
            with st.spinner("Refreshing clients to resolve missing session names..."):
                loaded = load_clients_from_api(api_key)
            if loaded:
                st.rerun()

        recent_rows = [
            {
                "timestamp": format_session_timestamp(session.get("timestamp")),
                "sessionId": session.get("id"),
                "clientId": session.get("clientId"),
            }
            for session in (recent_sessions or [])
        ]
        st.caption(
            f"Showing {len(recent_rows)} sessions from {recent_from.isoformat()} to {recent_to.isoformat()}"
        )
        if recent_rows:
            render_session_selection_block(
                sessions=recent_sessions or [],
                client_lookup=get_client_lookup(st.session_state.get("clients_cache", [])),
                scope="recent",
                table_key="recent_sessions_table",
            )
        else:
            st.info("No sessions were found in the last 7 days.")

    clients = st.session_state["clients_cache"]
    if clients:
        group_options = sorted(
            {
                str(client.get("group")).strip()
                for client in clients
                if client.get("group") not in (None, "")
            }
        )

        filter_col1, filter_col2 = st.columns([2, 1])
        search_query = filter_col1.text_input(
            "Search clients",
            placeholder="Search by name or external ID",
        ).strip()
        selected_group = filter_col2.selectbox(
            "Group",
            options=["All groups", *group_options],
            index=0,
        )

        normalized_query = search_query.lower()
        filtered_clients = []
        for client in clients:
            name = str(client.get("displayName") or "")
            external_id = str(client.get("externalId") or "")
            group = str(client.get("group") or "")

            matches_group = selected_group == "All groups" or group == selected_group
            matches_query = (
                not normalized_query
                or normalized_query in name.lower()
                or normalized_query in external_id.lower()
            )

            if matches_group and matches_query:
                filtered_clients.append(client)

        preview = [
            {
                "name": client.get("displayName"),
                "group": client.get("group"),
                "externalId": client.get("externalId"),
            }
            for client in filtered_clients
        ]
        st.caption(f"Showing {len(filtered_clients)} of {len(clients)} clients")

        table_event = st.dataframe(
            preview,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="clients_table",
        )

        selected_rows = table_event.selection.rows
        if selected_rows:
            selected_client = filtered_clients[selected_rows[0]]
            st.session_state["selected_client_id"] = str(selected_client.get("id") or "")

        selected_client_id = st.session_state.get("selected_client_id", "")
        selected_client = next(
            (
                client
                for client in clients
                if str(client.get("id") or "") == selected_client_id
            ),
            None,
        )

        if selected_client:
            if st.session_state["selected_client_last_id"] != selected_client_id:
                st.session_state["selected_client_last_id"] = selected_client_id
                st.session_state["session_filter_from"] = date.today() - timedelta(days=7)
                st.session_state["session_filter_to"] = date.today()
                st.session_state["client_sessions"] = []
                st.session_state["client_sessions_error"] = ""
                with st.spinner("Loading sessions for the selected client..."):
                    load_sessions_for_selected_client()
            render_client_detail(selected_client)
    else:
        st.info("No clients are currently cached.")


def main() -> None:
    st.set_page_config(
        page_title="1080 Reports",
        page_icon="🚀",
        layout="wide",
    )

    ensure_session_defaults()

    if st.session_state["auth_verified"] and st.session_state["api_valid"]:
        render_dashboard()
    else:
        render_login()


if __name__ == "__main__":
    main()
