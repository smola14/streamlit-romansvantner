from __future__ import annotations

import hmac
import hashlib
import base64
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
AUTH_SESSION_SECRET = st.secrets.get("auth_session_secret", os.getenv("AUTH_SESSION_SECRET", "")).strip()
API_TIMEOUT_SECONDS = 20
BLUE_RGB = (48, 54, 116)
BLACK_RGB = (37, 36, 35)
GREEN_RGB = (0, 192, 96)
ORANGE_RGB = (254, 148, 65)
RED_RGB = (251, 51, 49)
BLUE_HEX = "#303674"
FV_NORMS_PATH = Path(__file__).resolve().parent / "data" / "fv_norms.xlsx"
FV_NORM_SCATTER_PATH = Path(__file__).resolve().parent / "data" / "fv_norm_scatter.json"
RS_LOGO_PATH = Path(__file__).resolve().parent / "rs-logo.png"
UPLOADED_LOGOS_DIR = Path(__file__).resolve().parent / "uploaded_logos"
UPLOADED_PLAYER_PHOTOS_DIR = Path(__file__).resolve().parent / "uploaded_player_photos"
MAX_LOGO_BYTES = 2 * 1024 * 1024
MAX_PLAYER_PHOTO_BYTES = 5 * 1024 * 1024
AUTH_SESSION_DURATION = timedelta(days=1)
AUTH_STORAGE_KEY = "1080_auth_session"


PDF_TEXT = {
    "English": {
        "fv_title": "Force-Velocity Profile",
        "recommendation": "Recommendation",
        "player_fv_profile": "Player FV profile",
        "quadrant_reference": "Quadrant reference",
        "upper_reference": "Upper reference",
        "player_label": "Player",
        "lower_reference": "Lower reference",
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
        "fv_title": "Silovo-rýchlostný profil",
        "recommendation": "Odporúčanie",
        "player_fv_profile": "Hráčsky FV profil",
        "quadrant_reference": "Kvadrantová referencia",
        "upper_reference": "Horná referencia",
        "player_label": "Hráč",
        "lower_reference": "Dolná referencia",
        "notes": "Poznámka",
        "f0_status": "F0 vs norma",
        "v0_status": "V0 vs norma",
        "within": "V norme",
        "below": "Pod normou",
        "above": "Nad normou",
        "no_norm": "Bez normy",
        "q1_result": "Rýchlejšia akcelerácia / vyššia rýchlosť",
        "q2_result": "Rýchlejšia akcelerácia / nižšia rýchlosť",
        "q3_result": "Pomalšia akcelerácia / nižšia rýchlosť",
        "q4_result": "Pomalšia akcelerácia / vyššia rýchlosť",
        "q1_recs": [
            "Akcelerácia/práca na rýchlosti (<7 s)",
            "Šprintérsky tréning s odporom (25-50% pokles rýchlosti; 10-20 m)",
            "Letmé šprinty",
            "Šprintérsky tréning s asistenciou",
            "Speed bounding",
            "Zlepšenie izometrickej sily hamstringov",
        ],
        "q2_recs": [
            "Akcelerácia/práca na rýchlosti (<7 s)",
            "Šprintérsky tréning s odporom (25-50% pokles rýchlosti; 10-20 m)",
            "Letmé šprinty",
            "Šprintérsky tréning s asistenciou",
            "Zlepšenie cyklu natiahnutie-skrátenie (SSC)",
            "Zlepšenie reaktívnej sily",
            "Zlepšenie sily spojivových tkanív",
        ],
        "q3_recs": [
            "Šprintérsky tréning s odporom (50-75% pokles rýchlosti; 10 m)",
            "Letmé šprinty",
            "Zlepšenie cyklu natiahnutia-skrátenia (SSC)",
            "Zlepšenie sily extenzorov bedrového kĺbu",
            "Zlepšenie sily soleusu a gastrocnemiusu",
            "Zlepšenie absolútnej/relatívnej sily",
            "Zlepšenie reaktívnej sily",
            "Zlepšenie rýchlosti produkcie sily (RFD)",
        ],
        "q4_recs": [
            "Akcelerácia/práca na rýchlosti (<7 s)",
            "Šprintérsky tréning s odporom (50-75% pokles rýchlosti; 10 m)",
            "Zlepšenie sily extenzorov bedrového kĺbu",
            "Zlepšenie absolútnej/relatívnej sily",
            "Zlepšenie rýchlosti produkcie sily (RFD)",
            "Zlepšenie sily spojivových tkanív",
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
      const authStorageKey = data?.authStorageKey ?? "1080_auth_session";
      const authCommand = data?.authCommand ?? "read";
      const authCommandId = data?.authCommandId ?? "";

      const emitState = () => {
        setStateValue("clients_json", localStorage.getItem(storageKey) ?? "");
        setStateValue("last_synced", localStorage.getItem(syncedKey) ?? "");
        setStateValue("auth_session_token", localStorage.getItem(authStorageKey) ?? "");
        setStateValue("auth_storage_ready", true);
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

      const authCommandMarkerKey = `__last_auth_command__:${authStorageKey}`;
      const lastAuthCommandId = window[authCommandMarkerKey];

      if (authCommandId && authCommandId !== lastAuthCommandId) {
        if (authCommand === "write") {
          localStorage.setItem(authStorageKey, data?.authSessionToken ?? "");
        } else if (authCommand === "clear") {
          localStorage.removeItem(authStorageKey);
        }

        window[authCommandMarkerKey] = authCommandId;
      }

      emitState();
    }
    """,
)


def auth_session_signature(payload: str) -> str:
    if not AUTH_SESSION_SECRET:
        return ""
    return hmac.new(
        AUTH_SESSION_SECRET.encode("utf-8"),
        payload.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def create_auth_session_token(email: str) -> str:
    expires_at = int((datetime.now(timezone.utc) + AUTH_SESSION_DURATION).timestamp())
    payload = json.dumps({"email": email, "exp": expires_at}, separators=(",", ":"))
    payload_b64 = base64.urlsafe_b64encode(payload.encode("utf-8")).decode("ascii")
    signature = auth_session_signature(payload_b64)
    return f"{payload_b64}.{signature}"


def parse_auth_session_token(token: str) -> str | None:
    if not token or not AUTH_SESSION_SECRET or "." not in token:
        return None

    payload_b64, provided_signature = token.rsplit(".", 1)
    expected_signature = auth_session_signature(payload_b64)
    if not expected_signature or not hmac.compare_digest(expected_signature, provided_signature):
        return None

    try:
        payload_json = base64.urlsafe_b64decode(payload_b64.encode("ascii")).decode("utf-8")
        payload = json.loads(payload_json)
        email = str(payload["email"]).strip().lower()
        expires_at = int(payload["exp"])
    except (ValueError, KeyError, TypeError, json.JSONDecodeError):
        return None

    if not email or datetime.now(timezone.utc).timestamp() > expires_at:
        return None

    return email


def queue_auth_storage_command(command: str, token: str = "") -> None:
    st.session_state["auth_storage_command"] = command
    st.session_state["auth_storage_command_id"] = f"{command}:{iso_now()}"
    st.session_state["auth_session_token"] = token


def restore_auth_session_from_token(token: str) -> bool:
    auth_users = get_auth_users()
    email = parse_auth_session_token(token)
    if not email or email not in auth_users or not APP_API_KEY:
        return False

    st.session_state["auth_verified"] = True
    st.session_state["auth_user_email"] = email
    st.session_state["api_key"] = APP_API_KEY
    st.session_state["api_valid"] = True
    st.session_state["client_storage_autoload_complete"] = False
    st.session_state["auth_storage_autoload_complete"] = True
    return True


def render_app_styles() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: radial-gradient(circle at top, rgba(48,54,116,0.08), transparent 28%), #0f1117;
        }
        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1240px;
        }
        .hero-card {
            padding: 1.25rem 1.25rem 0.25rem;
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            background: linear-gradient(180deg, rgba(255,255,255,0.03), rgba(255,255,255,0.015));
            margin-bottom: 1rem;
        }
        .section-intro {
            color: rgba(255,255,255,0.72);
            margin-top: -0.35rem;
            margin-bottom: 0;
            font-size: 0.95rem;
        }
        div[data-testid="stMetric"] {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 0.9rem 1rem;
            background: rgba(255,255,255,0.02);
        }
        div[data-testid="stDataFrame"] {
            border-radius: 14px;
            overflow: hidden;
        }
        div[data-testid="stForm"] {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            padding: 1rem 1rem 0.25rem;
            background: rgba(255,255,255,0.02);
        }
        .thumb-card {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 0.75rem;
            background: rgba(255,255,255,0.02);
            width: fit-content;
        }
        .session-row {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 0.85rem 1rem;
            background: rgba(255,255,255,0.02);
            margin-bottom: 0.65rem;
        }
        .session-row-title {
            font-size: 1rem;
            font-weight: 600;
            margin: 0;
        }
        .session-row-subtitle {
            color: rgba(255,255,255,0.68);
            font-size: 0.92rem;
            margin: 0.18rem 0 0;
        }
        .session-row-meta {
            color: rgba(255,255,255,0.68);
            font-size: 0.92rem;
            margin: 0.24rem 0 0;
        }
        .session-row-inline {
            display: flex;
            align-items: baseline;
            gap: 0.65rem;
            flex-wrap: wrap;
        }
        .session-row-shell {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
        }
        div[data-testid="stButton"] > button,
        div[data-testid="stDownloadButton"] > button,
        div[data-testid="stFormSubmitButton"] > button {
            background: #2f6df6;
            color: white;
            border: 1px solid #2f6df6;
        }
        div[data-testid="stButton"] > button:hover,
        div[data-testid="stDownloadButton"] > button:hover,
        div[data-testid="stFormSubmitButton"] > button:hover {
            background: #2559c9;
            border-color: #2559c9;
            color: white;
        }
        </style>
        """,
        unsafe_allow_html=True,
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
    target_path = UPLOADED_LOGOS_DIR / candidate

    if target_path.exists():
        raise FileExistsError(candidate)

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
    processed_logo_upload_key = f"{key_prefix}_processed_logo_upload"
    default_option = "No saved logo" if not saved_logos else "Choose logo"

    if selected_logo_state_key not in st.session_state:
        st.session_state[selected_logo_state_key] = default_option

    pending_logo_name = st.session_state.pop(pending_logo_state_key, None)
    if pending_logo_name:
        st.session_state[selected_logo_state_key] = pending_logo_name

    options = ["Choose logo", *saved_logos] if saved_logos else ["No saved logo"]
    if st.session_state[selected_logo_state_key] not in options:
        st.session_state[selected_logo_state_key] = default_option

    st.markdown("##### Branding")
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

    if uploaded_logo is None:
        st.session_state.pop(processed_logo_upload_key, None)
    elif logo_is_valid:
        upload_token = f"{getattr(uploaded_logo, 'name', 'logo')}:{len(uploaded_logo.getvalue())}"
        if st.session_state.get(processed_logo_upload_key) != upload_token:
            st.session_state[processed_logo_upload_key] = upload_token
            try:
                saved_name = save_uploaded_logo(uploaded_logo)
            except FileExistsError as exc:
                st.error(f"Logo `{exc.args[0]}` has already been uploaded.")
            else:
                st.session_state[pending_logo_state_key] = saved_name
                st.success(f"Saved logo: {saved_name}")
                st.rerun()

    if uploaded_logo is not None and logo_is_valid:
        return uploaded_logo.getvalue()

    if selected_logo not in {"No saved logo", "Choose logo"}:
        return load_saved_logo_bytes(selected_logo)

    return None


def render_player_photo_selector(key_prefix: str, client_id: str) -> bytes | None:
    saved_photo = get_saved_player_photo_name(client_id)
    current_bytes = load_player_photo_bytes(client_id)

    if current_bytes:
        st.caption("Player photo")
        thumb_col, _ = st.columns([0.34, 0.66])
        thumb_col.image(current_bytes, width=140)
        return current_bytes

    st.caption("Add a player photo once and it will be reused for future exports.")
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


def handle_auth_storage_bridge() -> None:
    storage_result = CLIENT_STORAGE_COMPONENT(
        data={
            "storageKey": "1080_auth_bridge_clients",
            "command": "read",
            "commandId": "",
            "clientsJson": "",
            "lastSynced": "",
            "authStorageKey": AUTH_STORAGE_KEY,
            "authCommand": st.session_state["auth_storage_command"],
            "authCommandId": st.session_state["auth_storage_command_id"],
            "authSessionToken": st.session_state["auth_session_token"],
        },
        default={"clients_json": "", "last_synced": "", "auth_session_token": "", "auth_storage_ready": False},
        on_clients_json_change=lambda: None,
        on_last_synced_change=lambda: None,
        on_auth_session_token_change=lambda: None,
        on_auth_storage_ready_change=lambda: None,
        key="auth_storage_bridge",
        height=0,
    )

    auth_token = getattr(storage_result, "auth_session_token", "") or ""
    auth_storage_ready = bool(getattr(storage_result, "auth_storage_ready", False))

    if (
        not st.session_state["auth_verified"]
        and not st.session_state["auth_storage_autoload_complete"]
    ):
        if not auth_storage_ready:
            return

        if auth_token and restore_auth_session_from_token(auth_token):
            st.rerun()

        st.session_state["auth_storage_autoload_complete"] = True
        if auth_token:
            queue_auth_storage_command("clear")


def ensure_session_defaults() -> None:
    st.session_state.setdefault("auth_verified", False)
    st.session_state.setdefault("auth_user_email", "")
    st.session_state.setdefault("api_key", "")
    st.session_state.setdefault("api_valid", False)
    st.session_state.setdefault("auth_storage_command", "read")
    st.session_state.setdefault("auth_storage_command_id", "")
    st.session_state.setdefault("auth_session_token", "")
    st.session_state.setdefault("auth_storage_autoload_complete", False)
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
    st.session_state["auth_storage_autoload_complete"] = True
    queue_auth_storage_command("clear")
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
        st.markdown(
            """
            <div class="hero-card">
              <h1 style="margin:0;">1080 Reports</h1>
              <p class="section-intro">Sign in with your email and password to access reporting tools.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        auth_users = get_auth_users()
        if not auth_users:
            st.error("No app users are configured. Add `auth_users` to Streamlit secrets.")
            return
        if not APP_API_KEY:
            st.error("1080 API key is not configured. Set `api_1080_key` in Streamlit secrets or `API1080_KEY` as an environment variable.")
            return

        with st.form("auth-login-form"):
            email = st.text_input("Email", placeholder="you@example.com").strip().lower()
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
                if AUTH_SESSION_SECRET:
                    queue_auth_storage_command("write", create_auth_session_token(email))
                st.session_state["client_storage_autoload_complete"] = False
                st.session_state["auth_storage_autoload_complete"] = True
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


def get_quadrant_badge_fill(quadrant: str) -> tuple[int, int, int]:
    if quadrant == "Q1":
        return (14, 108, 79)
    if quadrant == "Q2":
        return ORANGE_RGB
    if quadrant == "Q3":
        return RED_RGB
    return ORANGE_RGB


def get_quadrant_result_fill(quadrant: str) -> tuple[int, int, int]:
    if quadrant == "Q1":
        return GREEN_RGB
    if quadrant == "Q2":
        return (255, 188, 89)
    if quadrant == "Q3":
        return (240, 131, 133)
    return (255, 188, 89)


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


def make_norm_scatter_plot(
    report: dict[str, Any],
    norm_row: dict[str, Any],
    scatter_entry: dict[str, Any],
    language: str,
) -> io.BytesIO:
    texts = PDF_TEXT[language]
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
    ax.set_title(f"{texts['quadrant_reference']} | {norm_row.get('category')}")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(True, linestyle="--", alpha=0.18)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    plt.close(fig)
    return buf


def make_normative_fv_profile(
    report: dict[str, Any],
    norm_row: dict[str, Any],
    language: str,
) -> io.BytesIO:
    texts = PDF_TEXT[language]
    player_v0 = float(report["v0"])
    player_f0 = float(report["f0"])
    bbox = dict(boxstyle="round", edgecolor="none", facecolor=BLUE_HEX)
    upper_x = [0, float(norm_row["v0_max"])]
    upper_y = [float(norm_row["f0_max"]), 0]
    lower_x = [0, float(norm_row["v0_min"])]
    lower_y = [float(norm_row["f0_min"]), 0]
    x_span = max(upper_x[1], player_v0) - min(lower_x[1], 0)
    y_span = max(upper_y[0], player_f0) - min(lower_y[0], 0)
    x_pad = max(x_span * 0.08, 0.15)
    y_pad = max(y_span * 0.08, 0.15)

    x_player = [0, player_v0]
    y_player = [player_f0, 0]

    fig, ax_line = plt.subplots(figsize=(5.8, 3.8))
    ax_line.plot(
        upper_x,
        upper_y,
        label=texts["upper_reference"],
        linestyle="--",
        color="#00C060",
        linewidth=2.0,
    )
    ax_line.plot(x_player, y_player, label=texts["player_label"], color=BLUE_HEX, linewidth=2.5)
    ax_line.plot(
        lower_x,
        lower_y,
        label=texts["lower_reference"],
        linestyle="--",
        color="#FB3331",
        linewidth=2.0,
    )
    ax_line.annotate(
        str(round(player_v0, 2)),
        (player_v0, 0),
        xytext=(player_v0, -0.6),
        textcoords="data",
        ha="center",
        va="center",
        color="white",
        bbox=bbox,
    )
    ax_line.annotate(
        str(round(player_f0, 2)),
        (0, player_f0),
        xytext=(-0.6, player_f0),
        textcoords="data",
        ha="center",
        va="center",
        color="white",
        bbox=bbox,
    )
    ax_line.spines["top"].set_visible(False)
    ax_line.spines["right"].set_visible(False)
    ax_line.set_xlabel("V0 [m/s]")
    ax_line.set_ylabel("F0 [N/kg]")
    ax_line.set_title(texts["player_fv_profile"], fontsize=10)
    ax_line.set_xlim(-x_pad * 0.2, max(upper_x[1], player_v0) + x_pad)
    ax_line.set_ylim(-y_pad * 0.2, max(upper_y[0], player_f0) + y_pad)
    ax_line.legend(loc="upper right", frameon=False)
    ax_line.margins(x=0.05, y=0.05)

    buf = io.BytesIO()
    fig.tight_layout()
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
    required = ["category", "f0_min", "f0_max", "v0_min", "v0_max", "f0_median", "v0_median"]
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
    fv_buf = make_normative_fv_profile(report, norm_row, language)
    scatter_buf = (
        make_norm_scatter_plot(report, norm_row, scatter_entry, language)
        if scatter_entry
        else None
    )
    pdf = FPDF("L", "mm", "A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    font_family = configure_pdf_font(pdf)

    left_x = 14
    left_w = 126
    right_x = 154
    right_w = 129
    logo_w = 32
    top_y = 18
    subtitle_y = 29
    badge_y = 39
    header_bottom_y = badge_y + 11
    logo_x = pdf.w - logo_w - 12
    logo_y = top_y + ((header_bottom_y - top_y) - logo_w) / 2
    has_logo = bool(logo_bytes)
    identity_x = left_x
    fv_chart_y = 56
    fv_chart_w = left_w
    metrics_y = 148
    rec_title_y = 154
    rec_text_y = 161
    photo_w = 46
    photo_x = right_x + (right_w - photo_w) / 2
    photo_y = 14
    scatter_y = 66
    scatter_w = 96
    scatter_x = right_x + (right_w - scatter_w) / 2
    rs_logo_w = 36
    rs_logo_x = pdf.w - rs_logo_w - 10
    rs_logo_y = pdf.h - 15

    if logo_bytes:
        pdf.image(io.BytesIO(logo_bytes), x=logo_x, y=logo_y, w=logo_w)

    pdf.set_text_color(*BLACK_RGB)
    pdf.set_font(font_family, "", 24)
    pdf.set_xy(identity_x, top_y)
    pdf.cell(0, 10, player_name, new_x="LMARGIN", new_y="NEXT")

    pdf.set_text_color(128, 128, 128)
    pdf.set_font(font_family, "", 12)
    pdf.set_xy(identity_x, subtitle_y)
    pdf.cell(0, 10, texts["fv_title"], new_x="LMARGIN", new_y="NEXT")

    quadrant = get_norm_quadrant(report, norm_row)
    quadrant_result = get_quadrant_result(quadrant, language)
    recommendations = get_quadrant_recommendations(quadrant, language)

    badge_x = identity_x
    pdf.set_font(font_family, "", 12)
    pdf.set_fill_color(*get_quadrant_badge_fill(quadrant))
    pdf.set_text_color(255, 255, 255)
    rounded_corner_cell(pdf, badge_x, badge_y, 14, 11, quadrant)

    pdf.set_fill_color(*get_quadrant_result_fill(quadrant))
    rounded_corner_cell(pdf, badge_x + 16, badge_y, min(100, left_w - 16), 11, quadrant_result)

    pdf.image(fv_buf, x=left_x, y=fv_chart_y, w=fv_chart_w)
    if player_photo_bytes:
        pdf.image(io.BytesIO(player_photo_bytes), x=photo_x, y=photo_y, w=photo_w, h=photo_w, keep_aspect_ratio=True)
    if scatter_buf:
        pdf.image(scatter_buf, x=scatter_x, y=scatter_y, w=scatter_w)

    f0 = float(report["f0"])
    v0 = float(report["v0"])
    pmax = float(report["pMax"])
    drf = float(report.get("ratioOfForceDecrease") or 0)
    rfmax = float(report.get("ratioOfForceMax") or 0)

    metrics_block_w = 74
    data_x = left_x + (left_w - metrics_block_w) / 2
    data_y = metrics_y
    y_second_row = 18
    cell_w = 22
    cell_h = 12
    cell_h_sub = 7

    pdf.set_fill_color(*BLUE_RGB)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(font_family, "", 12)

    rounded_corner_cell(pdf, data_x, data_y, cell_w, cell_h, str(round(f0, 2)))
    rounded_corner_cell(pdf, data_x + 26, data_y, cell_w, cell_h, str(round(v0, 2)))
    rounded_corner_cell(pdf, data_x + 52, data_y, cell_w, cell_h, str(round(pmax, 2)))
    rounded_corner_cell(pdf, data_x, data_y + y_second_row, cell_w, cell_h, str(round(v0 * 3.6, 2)))
    rounded_corner_cell(pdf, data_x + 26, data_y + y_second_row, cell_w, cell_h, str(round(drf, 2)))
    rounded_corner_cell(pdf, data_x + 52, data_y + y_second_row, cell_w, cell_h, str(round(rfmax, 2)))

    pdf.set_font(font_family, "", 8)
    pdf.set_text_color(220, 220, 220)
    rounded_corner_cell(pdf, data_x, data_y + 9, cell_w, cell_h_sub, "F0 [N/kg]")
    rounded_corner_cell(pdf, data_x + 26, data_y + 9, cell_w, cell_h_sub, "V0 [m/s]")
    rounded_corner_cell(pdf, data_x + 52, data_y + 9, cell_w, cell_h_sub, "PMax [W]")
    rounded_corner_cell(pdf, data_x, data_y + 9 + y_second_row, cell_w, cell_h_sub, "V0 [km/h]")
    rounded_corner_cell(pdf, data_x + 26, data_y + 9 + y_second_row, cell_w, cell_h_sub, "DRF")
    rounded_corner_cell(pdf, data_x + 52, data_y + 9 + y_second_row, cell_w, cell_h_sub, "RFmax")

    pdf.set_font(font_family, "", 14)
    pdf.set_text_color(*BLACK_RGB)
    pdf.text(right_x, rec_title_y, texts["recommendation"])
    pdf.set_font(font_family, "", 8)
    y_coordinate = rec_text_y
    for item in recommendations:
        pdf.text(right_x, y_coordinate, "* " + item)
        y_coordinate += 7

    if RS_LOGO_PATH.is_file():
        pdf.image(str(RS_LOGO_PATH), x=rs_logo_x, y=rs_logo_y, w=rs_logo_w)

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
        f"Run {index + 1} | F0 {format_decimal(report.get('f0'))} | V0 {format_decimal(report.get('v0'))}": report
        for index, report in enumerate(reports)
    }

    st.caption("Choose the run, report style, and optional assets for the export.")
    selected_run_label = st.selectbox(
        "Run",
        options=list(run_options.keys()),
        key=f"fv_run_select_{exercise.get('id')}",
    )
    selected_report = run_options[selected_run_label]

    export_col1, export_col2 = st.columns([1, 1])
    export_mode = export_col1.radio(
        "Report type",
        options=["Non-normative", "Normative"],
        horizontal=True,
        key=f"fv_export_mode_{exercise.get('id')}",
    )
    export_language = st.selectbox(
        "Language",
        options=["English", "Slovak"],
        key=f"fv_export_language_{exercise.get('id')}",
    )
    selected_norm = None
    selected_norm_category = ""
    scatter_entry = None
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
            "Reference category",
            options=list(norm_options.keys()),
            key=f"fv_norm_category_{exercise.get('id')}",
        )
        selected_norm = norm_options[selected_norm_category]
        scatter_entry = scatter_map.get(selected_norm_category)

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

    player_photo_bytes = render_player_photo_selector(
        f"fv_player_photo_{exercise.get('id')}",
        player_client_id,
    ) if player_client_id else None

    if export_mode == "Normative" and selected_norm is not None:
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

    st.markdown("### Running (LR) force-velocity profile")
    st.caption(f"Valid runs: {len(reports)} | Failed runs: {len(failed_reports)}")

    if reports:
        if st.button(
            "Open FV PDF export",
            key=f"fv_export_open_top_{exercise.get('id')}",
            use_container_width=True,
        ):
            render_fv_export_dialog(exercise, reports, runner_info, client)

        summary_rows = [
            {
                "Run": index + 1,
                "F0": format_decimal(report.get("f0")),
                "V0": format_decimal(report.get("v0")),
                "PMax": format_decimal(report.get("pMax")),
                "Tau": format_decimal(report.get("tau")),
                "Confidence": format_decimal(report.get("confidence"), 3),
                "Est. unloaded max speed": format_decimal(report.get("estimatedUnloadedMaxSpeed")),
                "RF max": format_decimal(report.get("ratioOfForceMax")),
                "Force decrease": format_decimal(report.get("ratioOfForceDecrease")),
            }
            for index, report in enumerate(reports)
        ]
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)
        st.caption("Review the runs, then use the export button above when you are ready to create a PDF.")
    else:
        st.info("No FV runs were returned for this exercise.")


def render_split_profile(exercise: dict[str, Any], payload: dict[str, Any]) -> None:
    reports = payload.get("reports") or []

    st.markdown("### 15-0-5 split profile")

    top_col1, top_col2, top_col3 = st.columns(3)
    top_col1.metric("Runs", len(reports))
    top_col2.metric("Split length", format_decimal(reports[0].get("splitLength")) if reports else "5.00")
    top_col3.metric("Units", "meters" if reports and not reports[0].get("isYards") else "yards")

    split_rows: list[dict[str, Any]] = []
    for report_index, report in enumerate(reports):
        for split in report.get("splits") or []:
            split_rows.append(
                {
                    "Run": report_index + 1,
                    "Start": format_decimal(split.get("start")),
                    "End": format_decimal(split.get("end")),
                    "Time": format_decimal(split.get("time"), 3),
                    "Top speed": format_decimal(split.get("topSpeed")),
                    "Max force": format_decimal(split.get("maxForce")),
                    "Max power": format_decimal(split.get("maxPower")),
                }
            )

    if split_rows:
        st.dataframe(split_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No split rows were returned for this exercise.")


def render_session_detail_content(
    session_detail: dict[str, Any],
    client_lookup: dict[str, dict[str, Any]],
) -> None:
    exercises = session_detail.get("exercises") or []
    client = client_lookup.get(str(session_detail.get("clientId") or ""), {})
    session_time = format_session_timestamp(session_detail.get("timestamp"))

    st.subheader("Selected session")
    st.caption("Session overview and exercise-specific reports.")

    if exercises:
        exercise_rows = [
            {
                "Exercise": exercise.get("exerciseTypeName"),
                "Sets": len(exercise.get("sets") or []),
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


def render_session_selection_block(
    sessions: list[dict[str, Any]],
    client_lookup: dict[str, dict[str, Any]],
    scope: str,
    table_key: str,
) -> None:
    ensure_session_scope_defaults(scope)
    selected_session_id = st.session_state.get(f"selected_session_id_{scope}", "")

    for index, session in enumerate(sessions):
        session_id = str(session.get("id") or "")
        session_time = format_session_timestamp(session.get("timestamp"))
        athlete_name = get_client_display_name(session.get("clientId"), client_lookup)
        is_selected = session_id == selected_session_id

        st.markdown('<div class="session-row">', unsafe_allow_html=True)
        row_col1, row_col2 = st.columns([5, 1.1], vertical_alignment="center")
        with row_col1:
            st.markdown(
                f"""
                <div class="session-row-inline">
                  <p class="session-row-title">{session_time}</p>
                  <p class="session-row-subtitle">{athlete_name}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
        button_label = "Selected" if is_selected else "Open session"
        button_type = "secondary" if is_selected else "primary"
        if row_col2.button(
            button_label,
            key=f"{table_key}_open_{index}",
            use_container_width=True,
            type=button_type,
        ) and not is_selected:
            st.session_state[f"selected_session_id_{scope}"] = session_id
            with st.spinner("Loading session detail..."):
                load_selected_session_detail(scope)
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    session_detail_error = st.session_state.get(f"session_detail_error_{scope}", "")
    if session_detail_error:
        st.error(session_detail_error)

    session_detail = st.session_state.get(f"session_detail_{scope}")
    if session_detail:
        render_session_detail_content(session_detail, client_lookup)


def render_client_detail(client: dict[str, Any]) -> None:
    st.subheader("Athlete overview")
    st.caption("Profile details and recent sessions for the selected athlete.")

    top_col1, top_col2, top_col3 = st.columns(3)
    top_col1.metric("Name", format_optional_value(client.get("displayName")))
    top_col2.metric("Group", format_optional_value(client.get("group")))
    top_col3.metric("Date of birth", format_optional_value(client.get("dateOfBirth")))

    info_col1, info_col2 = st.columns(2)
    with info_col1:
        st.write(f"**Height:** {format_optional_value(client.get('height'))}")
        st.write(f"**Weight:** {format_optional_value(client.get('weight'))}")
    with info_col2:
        st.write(f"**Created:** {format_optional_value(client.get('created'))}")
        st.write(f"**Edited:** {format_optional_value(client.get('edited'))}")
        st.write(f"**Tags:** {format_optional_value(client.get('tags'))}")

    st.subheader("Sessions")

    filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 0.8])
    filter_col1.date_input("From", key="session_filter_from")
    filter_col2.date_input("To", key="session_filter_to")
    reload_sessions = filter_col3.button("Refresh sessions", use_container_width=True)

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

    st.markdown(
        """
        <div class="hero-card">
          <h1 style="margin:0;">1080 Reports</h1>
          <p class="section-intro">Browse athletes, review sessions, and export clean performance reports.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
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

    overview_tab, athletes_tab = st.tabs(["Recent sessions", "Athletes"])

    with overview_tab:
        st.markdown("### Recent sessions")
        st.caption("A quick view of the latest activity from the last 7 days.")
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

            st.caption(
                f"Showing {len(recent_sessions or [])} sessions from {recent_from.isoformat()} to {recent_to.isoformat()}"
            )
            if recent_sessions:
                render_session_selection_block(
                    sessions=recent_sessions,
                    client_lookup=get_client_lookup(st.session_state.get("clients_cache", [])),
                    scope="recent",
                    table_key="recent_sessions_table",
                )
            else:
                st.info("No sessions were found in the last 7 days.")

    with athletes_tab:
        st.markdown("### Athletes")
        st.caption("Search and filter the athlete list, then open a profile to review session history.")

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
                "Search athletes",
                placeholder="Search by athlete name or external reference",
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

            st.caption(f"Showing {len(filtered_clients)} of {len(clients)} athletes")

            selected_client_id = st.session_state.get("selected_client_id", "")

            for index, filtered_client in enumerate(filtered_clients):
                client_id = str(filtered_client.get("id") or "")
                is_selected = client_id == selected_client_id
                group_label = format_optional_value(filtered_client.get("group"))
                tags_label = format_optional_value(filtered_client.get("tags"))

                st.markdown('<div class="session-row">', unsafe_allow_html=True)
                row_col1, row_col2 = st.columns([5, 1.1], vertical_alignment="center")
                with row_col1:
                    st.markdown(
                        f"""
                        <p class="session-row-title">{format_optional_value(filtered_client.get("displayName"))}</p>
                        <p class="session-row-meta">Group: {group_label} | Tags: {tags_label}</p>
                        """,
                        unsafe_allow_html=True,
                    )

                button_label = "Selected" if is_selected else "Open athlete"
                button_type = "secondary" if is_selected else "primary"
                if row_col2.button(
                    button_label,
                    key=f"clients_table_open_{index}",
                    use_container_width=True,
                    type=button_type,
                ) and not is_selected:
                    st.session_state["selected_client_id"] = client_id
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

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

    render_app_styles()
    ensure_session_defaults()
    handle_auth_storage_bridge()

    if not st.session_state["auth_verified"] and not st.session_state["auth_storage_autoload_complete"]:
        st.caption("Checking saved session...")
        return

    if st.session_state["auth_verified"] and st.session_state["api_valid"]:
        render_dashboard()
    else:
        render_login()


if __name__ == "__main__":
    main()
