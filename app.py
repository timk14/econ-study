from __future__ import annotations

import hashlib
import hmac
import io
import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from dotenv import load_dotenv

try:
    from sqlalchemy import bindparam, create_engine, text
except ImportError:
    create_engine = None
    text = None

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    from google import genai
except ImportError:
    genai = None


load_dotenv()
RAG_DATA_DIR = Path(__file__).parent / "rag_data"
DOCUMENT_CACHE_DIR = Path(__file__).parent / ".document_cache"


st.set_page_config(
    page_title="EconMaster AI",
    page_icon="◒",
    layout="wide",
    initial_sidebar_state="expanded",
)


APP_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Space+Grotesk:wght@400;500;600;700&display=swap');
:root { --ink:#f2f5f4; --muted:#9aa9ad; --paper:#101719; --panel:#182226; --panel-strong:#202d32; --line:#314148; --coral:#ff8060; --teal:#53c5bb; --yellow:#f4c95d; }
html, body, [class*="css"] { font-family: 'Space Grotesk', sans-serif; color: var(--ink); }
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] { background: var(--paper); }
[data-testid="stHeader"] { background: rgba(16,23,25,.9); }
[data-testid="stMain"] { background: var(--paper); }
[data-testid="stSidebar"] { background: #17212b; border-right: 0; }
[data-testid="stSidebar"] * { color: #f6f3ed !important; }
[data-testid="stSidebar"] .stCaption { color: #aeb8bd !important; }
[data-testid="stSidebar"] hr { border-color: #3d4a53; }
[data-testid="stSidebar"] input { background:#26343e; border:1px solid #52616a; }
.hero { padding: 2rem 0 1rem; border-bottom: 1px solid var(--line); margin-bottom: 1.75rem; }
.kicker { color: var(--coral); font-family:'DM Mono', monospace; font-size:.76rem; letter-spacing:.08em; text-transform:uppercase; }
.hero h1 { font-size: clamp(2.6rem, 6vw, 5.6rem); letter-spacing:-.07em; line-height:.88; margin:.5rem 0 1rem; max-width: 800px; }
.hero p { color: var(--muted); font-size:1.05rem; max-width:650px; line-height:1.55; }
.section-label { color:var(--muted); font-family:'DM Mono',monospace; font-size:.72rem; letter-spacing:.08em; text-transform:uppercase; margin: 1.6rem 0 .65rem; }
.metric-strip { display:grid; grid-template-columns:repeat(3,1fr); gap:1px; background:var(--line); border:1px solid var(--line); margin:1.5rem 0; }
.metric { background:var(--panel); padding:1rem 1.1rem; }
.metric strong { display:block; font-size:1.8rem; letter-spacing:-.05em; }
.metric span { color:var(--muted); font-size:.78rem; }
.question-panel { border:1px solid var(--line); background:var(--panel); padding:1.7rem; margin-top:1rem; }
.question-number { color:var(--coral); font-family:'DM Mono',monospace; font-size:.78rem; }
.question-text { font-size:1.32rem; line-height:1.45; margin:.6rem 0 1.25rem; }
.source-chip { display:inline-block; background:#183b3a; color:var(--teal); padding:.25rem .55rem; font-size:.72rem; font-family:'DM Mono',monospace; }
.explanation { border-left:3px solid var(--yellow); background:#3a3420; color:#f7edc1; padding:1rem 1.1rem; line-height:1.55; }
.stButton button { border-radius:2px; font-weight:600; min-height:2.7rem; }
.stButton button[kind="primary"] { background:var(--coral); border-color:var(--coral); }
button, [role="button"] { color:var(--ink); }
input, textarea, [data-baseweb="select"] > div, [data-testid="stFileUploaderDropzone"] { background:var(--panel-strong) !important; color:var(--ink) !important; border-color:var(--line) !important; }
div[data-testid="stFileUploader"] { border:1px dashed #718089; padding:.4rem; }
[data-testid="stAlert"] { background:var(--panel-strong); color:var(--ink); border-color:var(--line); }
[data-testid="stDataFrame"] { border:1px solid var(--line); }
.stProgress > div > div { background:var(--coral); }
.small-mono { font-family:'DM Mono',monospace; font-size:.72rem; color:var(--muted); }
@media (max-width: 700px) { .metric-strip { grid-template-columns:1fr; } .hero h1 { font-size:3.4rem; } }
</style>
"""
st.markdown(APP_CSS, unsafe_allow_html=True)


QUESTION_SCHEMA = {
    "type": "object",
    "properties": {
        "question": {"type": "string"},
        "options": {"type": "array", "items": {"type": "string"}},
        "correct_answer": {"type": "string", "enum": ["A", "B", "C", "D"]},
        "topic_tag": {"type": "string"},
        "explanation": {"type": "string"},
        "source_reference": {"type": "string"},
        "data_table": {
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "columns": {"type": "array", "items": {"type": "string"}},
                "rows": {
                    "type": "array",
                    "items": {"type": "array", "items": {"type": "string"}},
                },
            },
            "required": ["title", "columns", "rows"],
        },
        "learning": {
            "type": "object",
            "properties": {
                "objective": {"type": "string"},
                "hint": {"type": "string"},
                "worked_steps": {"type": "array", "items": {"type": "string"}},
                "misconception": {"type": "string"},
            },
            "required": ["objective", "hint", "worked_steps", "misconception"],
        },
    },
    "required": [
        "question", "options", "correct_answer", "topic_tag",
        "explanation", "source_reference",
    ],
}

DEMO_QUESTION = {
    "question": "A country's nominal GDP rises from $500 billion to $525 billion while its GDP deflator rises from 100 to 105. What happened to real GDP?",
    "options": ["A) It increased by 5%", "B) It decreased by 5%", "C) It stayed approximately constant", "D) It increased by 10%"],
    "correct_answer": "C",
    "topic_tag": "Real vs. nominal GDP",
    "explanation": "Real GDP = nominal GDP / (GDP deflator / 100). Initially it is $500 billion. In the later year, $525 / 1.05 = $500 billion. Nominal growth was fully explained by the price-level increase, so real GDP stayed approximately constant.",
    "source_reference": "Demo question shown before course materials and a Gemini API key are configured.",
}

QUESTION_POOL_TARGET = 12
QUESTION_POOL_REFILL_THRESHOLD = 4
GAME_ROUNDS = 10
GAME_STARTING_INDICATORS = {
    "growth": 60,
    "inflation": 60,
    "employment": 60,
    "stability": 60,
}
GAME_TRACKS = ["Mixed Final", "Macroeconomic Crisis", "Business Strategy"]

GAME_ROUND_SCHEMA = {
    "type": "object",
    "properties": {
        **QUESTION_SCHEMA["properties"],
        "scenario_title": {"type": "string"},
        "scenario_context": {"type": "string"},
        "impact_profile": {
            "type": "object",
            "properties": {
                "growth": {"type": "integer", "minimum": -12, "maximum": 12},
                "inflation": {"type": "integer", "minimum": -12, "maximum": 12},
                "employment": {"type": "integer", "minimum": -12, "maximum": 12},
                "stability": {"type": "integer", "minimum": -12, "maximum": 12},
            },
            "required": ["growth", "inflation", "employment", "stability"],
        },
    },
    "required": [
        *QUESTION_SCHEMA["required"], "scenario_title", "scenario_context", "impact_profile",
    ],
}


def init_state() -> None:
    defaults: dict[str, Any] = {
        "course_context": "",
        "document_names": [],
        "document_hash": "",
        "local_context": "",
        "local_document_names": [],
        "uploaded_context": "",
        "uploaded_document_names": [],
        "rag_errors": [],
        "history": [],
        "history_loaded": False,
        "recorded_attempt_questions": set(),
        "authenticated": False,
        "active_question": None,
        "active_mode": "Practice",
        "exam_questions": [],
        "exam_index": 0,
        "exam_answers": {},
        "exam_submitted": False,
        "practice_answer": None,
        "practice_revealed": False,
        "learning_question": None,
        "learning_answer": None,
        "learning_revealed": False,
        "learning_hint_visible": False,
        "difficulty": "Intermediate",
        "question_loading": False,
        "question_pool": [],
        "question_pool_signature": "",
        "served_question_ids": set(),
        "recent_question_fingerprints": [],
        "recent_topic_tags": [],
        "recent_concept_families": [],
        "recent_graph_families": [],
        "question_style": "auto",
        "question_style_cycle_index": 0,
        "document_cache_summary": {"cached_files": 0, "total_bytes": 0},
        "game_active": False,
        "game_track": "Mixed Final",
        "game_round": 0,
        "game_questions": [],
        "game_indicators": GAME_STARTING_INDICATORS.copy(),
        "game_round_results": [],
        "game_answer": None,
        "game_revealed": False,
        "game_over": False,
        "game_session_id": None,
        "game_result_saved": False,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def configured_value(name: str, default: str = "") -> str:
    """Read a deployment secret without ever rendering its value."""
    try:
        secret_value = st.secrets.get(name, "")
    except Exception:
        secret_value = ""
    return str(secret_value or os.getenv(name, default))


def database_url() -> str:
    default_path = Path(__file__).parent / "history.db"
    url = configured_value("DATABASE_URL")
    if not url:
        return f"sqlite:///{default_path}"
    if url.startswith("postgres://"):
        url = "postgresql+psycopg://" + url.removeprefix("postgres://")
    elif url.startswith("postgresql://") and "+psycopg" not in url:
        url = "postgresql+psycopg://" + url.removeprefix("postgresql://")
    parsed = urlparse(url)
    placeholders = {"host", "user", "password", "database", "dbname"}
    if parsed.scheme.startswith("postgresql") and (
        not parsed.hostname
        or parsed.hostname.lower() in placeholders
        or (parsed.username and parsed.username.lower() in placeholders)
        or parsed.path.lstrip("/").lower() in placeholders
    ):
        return f"sqlite:///{default_path}"
    return url


@st.cache_resource
def get_database_engine(url: str):
    if create_engine is None:
        raise RuntimeError("Install SQLAlchemy and psycopg to enable history storage.")
    return create_engine(url, pool_pre_ping=True)


def initialize_database() -> None:
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            CREATE TABLE IF NOT EXISTS attempts (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                mode TEXT NOT NULL,
                topic TEXT NOT NULL,
                difficulty TEXT NOT NULL,
                question TEXT NOT NULL,
                selected TEXT,
                correct TEXT NOT NULL,
                is_correct INTEGER NOT NULL,
                verified INTEGER NOT NULL DEFAULT 0
            )
        """))
        connection.execute(text("""
            CREATE TABLE IF NOT EXISTS question_bank (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                context_signature TEXT NOT NULL,
                difficulty TEXT NOT NULL,
                question_json TEXT NOT NULL,
                times_served INTEGER NOT NULL DEFAULT 0
            )
        """))
        connection.execute(text("""
            CREATE TABLE IF NOT EXISTS game_round_bank (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                context_signature TEXT NOT NULL,
                difficulty TEXT NOT NULL,
                track TEXT NOT NULL,
                round_json TEXT NOT NULL,
                times_served INTEGER NOT NULL DEFAULT 0
            )
        """))
        connection.execute(text("""
            CREATE TABLE IF NOT EXISTS game_sessions (
                id TEXT PRIMARY KEY,
                started_at TEXT NOT NULL,
                completed_at TEXT,
                track TEXT NOT NULL,
                difficulty TEXT NOT NULL,
                final_growth INTEGER,
                final_inflation INTEGER,
                final_employment INTEGER,
                final_stability INTEGER,
                final_score INTEGER,
                outcome TEXT
            )
        """))
        connection.execute(text("""
            CREATE TABLE IF NOT EXISTS game_rounds (
                id TEXT PRIMARY KEY,
                game_session_id TEXT NOT NULL,
                round_number INTEGER NOT NULL,
                topic TEXT NOT NULL,
                question TEXT NOT NULL,
                selected TEXT,
                correct TEXT NOT NULL,
                is_correct INTEGER NOT NULL,
                impact_json TEXT NOT NULL
            )
        """))


def load_history() -> None:
    if st.session_state.history_loaded:
        return
    engine = get_database_engine(database_url())
    with engine.connect() as connection:
        rows = connection.execute(text("""
            SELECT created_at, mode, topic, difficulty, question,
                   selected, correct, is_correct, verified
            FROM attempts ORDER BY id
        """)).mappings().all()
    st.session_state.history = [
        {
            **dict(row),
            "timestamp": row["created_at"],
            "is_correct": bool(row["is_correct"]),
            "verified": bool(row["verified"]),
        }
        for row in rows
    ]
    st.session_state.history_loaded = True


def save_attempt(attempt: dict[str, Any]) -> None:
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            INSERT INTO attempts
            (id, created_at, mode, topic, difficulty, question, selected,
             correct, is_correct, verified)
            VALUES (:id, :created_at, :mode, :topic, :difficulty, :question,
                    :selected, :correct, :is_correct, :verified)
        """), {
            "id": str(uuid.uuid4()),
            "created_at": attempt["timestamp"],
            **attempt,
            "is_correct": int(attempt["is_correct"]),
            "verified": int(attempt.get("verified", False)),
        })


def _question_fingerprint(question: dict[str, Any]) -> str:
    text = f"{question.get('question', '')}|{question.get('source_reference', '')}|{question.get('topic_tag', '')}"
    return hashlib.sha256(text.encode()).hexdigest()


def question_concept_family(question: dict[str, Any]) -> str:
    text = " ".join(str(question.get(key, "")) for key in ("topic_tag", "source_reference", "question")).lower()
    families = {
        "scarcity_ppf": ("ppf", "production possibil", "opportunity cost", "scarcity"),
        "trade": ("comparative advantage", "absolute advantage", "tariff", "trade"),
        "supply_demand": ("supply", "demand", "equilibrium", "consumer surplus", "price ceiling", "price floor"),
        "elasticity": ("elasticity", "total revenue", "tax incidence"),
        "production_costs": ("total cost", "fixed cost", "variable cost", "marginal cost", "average cost", "cost curve"),
        "labor": ("labor", "labour", "wage", "employment", "unemployment", "marginal revenue product"),
        "gdp_growth": ("gdp", "gross domestic", "real vs", "nominal", "economic growth"),
        "inflation": ("inflation", "cpi", "deflator", "price level"),
        "macro_policy": ("aggregate demand", "aggregate supply", "ad-as", "fiscal", "monetary", "interest rate"),
        "market_structure": ("monopoly", "oligopoly", "perfect competition", "market structure"),
    }
    for family, keywords in families.items():
        if any(keyword in text for keyword in keywords):
            return family
    return "general_economics"


def question_graph_family(question: dict[str, Any]) -> str:
    graph_type = infer_graph_type(question.get("question", ""))
    return graph_type if graph_type != "generic" else "nonvisual"


def _question_diversity_sort_key(question: dict[str, Any]) -> tuple[int, int, int, int, int, str]:
    recent_hashes = set(st.session_state.get("recent_question_fingerprints", [])[-8:])
    recent_topics = set(st.session_state.get("recent_topic_tags", [])[-3:])
    recent_concepts = set(st.session_state.get("recent_concept_families", [])[-4:])
    recent_graphs = set(st.session_state.get("recent_graph_families", [])[-3:])
    topic = question.get("topic_tag", "General economics")
    return (
        1 if _question_fingerprint(question) in recent_hashes else 0,
        1 if topic in recent_topics else 0,
        1 if question_concept_family(question) in recent_concepts else 0,
        1 if question_graph_family(question) in recent_graphs else 0,
        int(question.get("times_served", 0)),
        hashlib.sha256((question.get("bank_id") or question.get("question", "")).encode()).hexdigest(),
    )


def order_question_pool(questions: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Create a varied serving sequence instead of merely ranking independent rows."""
    remaining = list(questions)
    ordered: list[dict[str, Any]] = []
    selected_topics: set[str] = set()
    selected_concepts: set[str] = set()
    selected_graphs: set[str] = set()
    while remaining:
        def pool_key(question: dict[str, Any]) -> tuple[int, int, int, int, int, int, str]:
            base_key = _question_diversity_sort_key(question)
            topic = question.get("topic_tag", "General economics")
            concept = question_concept_family(question)
            graph = question_graph_family(question)
            return (
                base_key[0], base_key[1], base_key[2], base_key[3],
                1 if topic in selected_topics else 0,
                1 if concept in selected_concepts or graph in selected_graphs else 0,
                base_key[5],
            )
        next_question = min(remaining, key=pool_key)
        remaining.remove(next_question)
        ordered.append(next_question)
        selected_topics.add(next_question.get("topic_tag", "General economics"))
        selected_concepts.add(question_concept_family(next_question))
        selected_graphs.add(question_graph_family(next_question))
    return ordered


def _note_question_seen(question: dict[str, Any]) -> None:
    recent_hashes = st.session_state.get("recent_question_fingerprints", [])
    recent_hashes.append(_question_fingerprint(question))
    if len(recent_hashes) > 8:
        recent_hashes = recent_hashes[-8:]
    st.session_state.recent_question_fingerprints = recent_hashes
    topic = question.get("topic_tag", "General economics")
    recent_topics = st.session_state.get("recent_topic_tags", [])
    recent_topics.append(topic)
    if len(recent_topics) > 3:
        recent_topics = recent_topics[-3:]
    st.session_state.recent_topic_tags = recent_topics
    recent_concepts = st.session_state.get("recent_concept_families", [])
    recent_concepts.append(question_concept_family(question))
    if len(recent_concepts) > 4:
        recent_concepts = recent_concepts[-4:]
    st.session_state.recent_concept_families = recent_concepts
    recent_graphs = st.session_state.get("recent_graph_families", [])
    recent_graphs.append(question_graph_family(question))
    if len(recent_graphs) > 3:
        recent_graphs = recent_graphs[-3:]
    st.session_state.recent_graph_families = recent_graphs


def load_saved_questions(context_signature: str, difficulty: str) -> list[dict[str, Any]]:
    """Load the least-used verified questions for this exact course library."""
    engine = get_database_engine(database_url())
    with engine.connect() as connection:
        rows = connection.execute(text("""
            SELECT id, question_json, times_served FROM question_bank
            WHERE context_signature IN :context_signatures AND difficulty = :difficulty
            ORDER BY times_served ASC, created_at ASC
            LIMIT 100
        """).bindparams(bindparam("context_signatures", expanding=True)), {
            "context_signatures": [
                context_signature,
                *[
                    hashlib.sha256(
                        f"{st.session_state.course_context}\n{difficulty}\n{style}".encode()
                    ).hexdigest()
                    for style in ("mixed", "graph_heavy", "table_heavy")
                ],
            ],
            "difficulty": difficulty,
        }).mappings().all()
    saved_questions = []
    for row in rows:
        if row["id"] in st.session_state.served_question_ids:
            continue
        question = json.loads(row["question_json"])
        question["bank_id"] = row["id"]
        question["times_served"] = int(row["times_served"])
        saved_questions.append(question)
    return saved_questions


def save_questions_to_bank(
    context_signature: str, difficulty: str, questions: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Store audited questions once so later sessions can reuse them without Gemini."""
    engine = get_database_engine(database_url())
    created_at = datetime.now().isoformat(timespec="seconds")
    saved_questions = []
    with engine.begin() as connection:
        for question in questions:
            question_id = str(uuid.uuid4())
            stored_question = {key: value for key, value in question.items() if key != "bank_id"}
            connection.execute(text("""
                INSERT INTO question_bank
                (id, created_at, context_signature, difficulty, question_json)
                VALUES (:id, :created_at, :context_signature, :difficulty, :question_json)
            """), {
                "id": question_id,
                "created_at": created_at,
                "context_signature": context_signature,
                "difficulty": difficulty,
                "question_json": json.dumps(stored_question),
            })
            stored_question["bank_id"] = question_id
            saved_questions.append(stored_question)
    return saved_questions


def mark_question_served(question: dict[str, Any]) -> None:
    question_id = question.get("bank_id")
    if not question_id:
        return
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            UPDATE question_bank SET times_served = times_served + 1 WHERE id = :id
        """), {"id": question_id})
    st.session_state.served_question_ids.add(question_id)


def render_login() -> None:
    configured_password = configured_value("APP_PASSWORD")
    if not configured_password:
        st.error("This app is locked until APP_PASSWORD is configured in Streamlit secrets or .env.")
        st.code('APP_PASSWORD = "choose-a-password"', language="toml")
        st.stop()
    st.markdown(
        '<div class="hero"><div class="kicker">PRIVATE STUDY STUDIO</div>'
        '<h1>Welcome back.</h1><p>Enter the study-room password to continue.</p></div>',
        unsafe_allow_html=True,
    )
    password = st.text_input("Password", type="password")
    if st.button("Enter study room", type="primary"):
        if hmac.compare_digest(password, configured_password):
            st.session_state.authenticated = True
            st.rerun()
        st.error("That password did not match.")


def extract_text(file_name: str, file_bytes: bytes) -> str:
    """Extract readable text from an uploaded PDF or DOCX."""
    suffix = file_name.lower().rsplit(".", 1)[-1]
    if suffix == "pdf":
        if PdfReader is None:
            raise RuntimeError("PyPDF2 is not installed.")
        reader = PdfReader(io.BytesIO(file_bytes))
        pages = [(page.extract_text() or "") for page in reader.pages]
        return "\n\n".join(pages)
    if suffix == "docx":
        if Document is None:
            raise RuntimeError("python-docx is not installed.")
        document = Document(io.BytesIO(file_bytes))
        paragraphs = [p.text for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                paragraphs.append(" | ".join(cell.text for cell in row.cells))
        return "\n".join(paragraphs)
    raise ValueError(f"Unsupported file type: {suffix}")


def cached_extract_text(file_name: str, file_bytes: bytes, cache_label: str = "document") -> str:
    """Parse a PDF or DOCX once and persist the normalized text for reuse."""
    DOCUMENT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    digest = hashlib.sha256(file_bytes).hexdigest()
    safe_label = re.sub(r"[^a-zA-Z0-9_-]+", "_", cache_label).strip("_") or "document"
    cache_path = DOCUMENT_CACHE_DIR / f"{safe_label}-{digest}.json"
    if cache_path.exists():
        try:
            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            if isinstance(payload, dict) and payload.get("text"):
                return payload["text"]
        except (OSError, ValueError, TypeError):
            pass
    text = extract_text(file_name, file_bytes).strip()
    payload = {
        "file_name": file_name,
        "cache_label": cache_label,
        "sha256": digest,
        "text": text,
        "bytes": len(file_bytes),
    }
    cache_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    st.session_state.document_cache_summary = get_document_cache_summary()
    return text


def get_document_cache_summary() -> dict[str, int]:
    """Return a lightweight summary of cached extracted document text."""
    DOCUMENT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_files = [path for path in DOCUMENT_CACHE_DIR.glob("*.json") if path.is_file()]
    total_bytes = sum(path.stat().st_size for path in cache_files)
    return {
        "cached_files": len(cache_files),
        "total_bytes": total_bytes,
    }


def clear_document_cache() -> None:
    """Remove all cached extracted document text so uploads are reparsed."""
    DOCUMENT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    for path in list(DOCUMENT_CACHE_DIR.glob("*.json")):
        if path.is_file():
            path.unlink(missing_ok=True)
    st.session_state.document_cache_summary = {"cached_files": 0, "total_bytes": 0}


def load_rag_folder() -> None:
    """Index bundled course materials once per app session."""
    if st.session_state.local_context or not RAG_DATA_DIR.exists():
        return
    texts: list[str] = []
    names: list[str] = []
    errors: list[str] = []
    for source_path in sorted(RAG_DATA_DIR.iterdir()):
        if not source_path.is_file() or source_path.suffix.lower() not in {".pdf", ".docx"}:
            continue
        try:
            raw_bytes = source_path.read_bytes()
            text = cached_extract_text(source_path.name, raw_bytes, cache_label="rag").strip()
            if text:
                texts.append(f"SOURCE FILE: {source_path.name}\n{text}")
                names.append(source_path.name)
            else:
                errors.append(f"{source_path.name}: no selectable text found")
        except Exception as exc:
            errors.append(f"{source_path.name}: {exc}")
    st.session_state.local_context = "\n\n---\n\n".join(texts)
    st.session_state.local_document_names = names
    st.session_state.rag_errors = errors


def refresh_course_context() -> None:
    contexts = [value for value in [
        st.session_state.local_context,
        st.session_state.uploaded_context,
    ] if value]
    st.session_state.course_context = "\n\n---\n\n".join(contexts)
    st.session_state.document_names = (
        st.session_state.local_document_names + st.session_state.uploaded_document_names
    )


def ingest_uploads(uploaded_files: list[Any]) -> None:
    signature = hashlib.sha256(
        b"".join(file.name.encode() + file.getvalue() for file in uploaded_files)
    ).hexdigest()
    if signature == st.session_state.document_hash:
        return
    texts: list[str] = []
    names: list[str] = []
    errors: list[str] = []
    for uploaded_file in uploaded_files:
        try:
            raw_bytes = uploaded_file.getvalue()
            text = cached_extract_text(uploaded_file.name, raw_bytes, cache_label="upload").strip()
            if text:
                texts.append(f"SOURCE FILE: {uploaded_file.name}\n{text}")
                names.append(uploaded_file.name)
            else:
                errors.append(f"{uploaded_file.name}: no selectable text found")
        except Exception as exc:
            errors.append(f"{uploaded_file.name}: {exc}")
    st.session_state.uploaded_context = "\n\n---\n\n".join(texts)
    st.session_state.uploaded_document_names = names
    st.session_state.document_hash = signature
    refresh_course_context()
    if errors:
        st.sidebar.warning("\n".join(errors))
    if texts:
        st.sidebar.success(f"Indexed {len(texts)} document(s)")


def build_prompt(
    context_text: str,
    difficulty: str,
    question_count: int = 1,
    question_style: str = "mixed",
) -> str:
    limited_context = context_text[:120000]
    quantity = "one" if question_count == 1 else f"exactly {question_count}"
    if question_style == "graph_heavy":
        style_block = """This batch is graph-heavy. Make at least 60% of the questions explicitly about graphs, curves, or charts. The graph topics must be diversified across the core beginning-economics sequence used in accelerated MBA/ABE courses, not all the same kind of chart. Use a rotating mix of these graph families: PPF/opportunity cost, supply and demand equilibrium, labor market wage/employment, AD-AS macro equilibrium, elasticity/revenue, and GDP or circular-flow relationships. Each graph question should clearly state the axes, curve(s), and the economic interpretation. Do not write a batch where every graph is just a demand curve or every question is a PPF. Force variation by concept and by graph shape.

For each graph-type question, describe the visual setup in a realistic way: axes, curves, movement, equilibrium point, or shifts. Make each graph correspond to a different textbook concept from the course material.
"""
    elif question_style == "table_heavy":
        style_block = """This batch is data-heavy. Make a visible share of questions use tables, schedules, or numeric scenarios. Use genuine economics data structures: a total-cost schedule for TFC/TVC/AVC/MC; demand and supply schedules for equilibrium or surplus/shortage; elasticity or total-revenue comparisons; GDP expenditure components; inflation/unemployment time series; or comparative-advantage production tables. The question should require students to inspect, compare, or calculate from the displayed values. Do not invent a table when a verbal question is clearer.
"""
    else:
        style_block = """Vary the question formats across the batch. Use a mix of numerical calculation, policy/trade-off analysis, graph or curve interpretation, market equilibrium or elasticity comparison, and if-then scenarios using table or chart data. Do not make the whole batch only short conceptual questions or only repeated verbal scenarios.

For graph/chart-based questions, explicitly describe the graph or table in the question text, including the relevant axes, shifts, equilibrium change, or data comparison. Rotate among PPF, supply-demand, labor market, AD-AS, elasticity/revenue, and GDP-style charts rather than repeating one figure pattern.
"""
    return f"""You are an economics professor and assessment designer for an accelerated MBA/ABE beginning-economics course.

Create {quantity} brand-new multiple-choice question(s) based ONLY on the course material below. Mimic its style, difficulty, vocabulary, topic distribution, and solution depth. Do not copy wording or simply reproduce a question. Change the scenario, values, or framing while testing the same economic principles. The requested difficulty is {difficulty}.

{style_block}

Make the questions distinct from one another. Rotate among core beginning-econ topics: scarcity, opportunity cost, PPF, comparative advantage, supply and demand, elasticity, consumer surplus, production costs, labor market, unemployment, GDP, inflation, real vs nominal, AD-AS, fiscal policy, and monetary policy. Use realistic MBA-style business applications where relevant, but keep the concepts at the introductory level. Avoid reusing the same exact scenario pattern, graph family, data-table family, or wording structure more than once. Each question must feel fresh, not templated.

Return only valid JSON matching the supplied schema. Return an array with exactly {question_count} objects. Each object's options must be exactly four strings beginning with A), B), C), and D). correct_answer must be one letter. Make each explanation step-by-step and include formulas or calculations where relevant. Before returning, independently solve every question and check that the marked answer, numerical values, and explanation agree.

When a question depends on a schedule, dataset, cost table, demand/supply schedule, GDP-component breakdown, or other values students must inspect, include a data_table object with a short title, concise column headers, and rows of string values. Keep the question prose focused on the task; do not repeat the full table in prose. Do not attach a data_table to a question that does not require one.

Every question must include a learning object. objective names the single skill being practiced. hint gives one productive next step without revealing the answer letter or final result. worked_steps contains 2 to 4 short, ordered reasoning steps that lead to the answer. misconception explains why a tempting wrong approach fails. The learning material must be specific to the actual values, graph, or causal mechanism in that question, never generic study advice.

source_reference should briefly name the concept or pattern from the source material that this new question mirrors.

COURSE MATERIAL:
{limited_context}
"""


def generate_questions(
    context_text: str,
    difficulty: str,
    api_key: str,
    question_count: int,
    question_style: str = "mixed",
) -> list[dict[str, Any]]:
    if not context_text.strip():
        raise ValueError("Upload at least one course PDF or DOCX before generating questions.")
    if not api_key.strip():
        raise ValueError("Configure a Gemini API key in .env or Streamlit secrets before generating questions.")
    if genai is None:
        raise RuntimeError("Install google-genai to enable Gemini generation.")
    client = genai.Client(api_key=api_key.strip())
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=build_prompt(context_text, difficulty, question_count, question_style),
        config={
            "response_mime_type": "application/json",
            "response_schema": {
                "type": "array",
                "items": QUESTION_SCHEMA,
            },
            "temperature": 0.8,
        },
    )
    raw = response.text or ""
    questions = json.loads(raw)
    if not isinstance(questions, list) or len(questions) != question_count:
        raise ValueError("Gemini returned an invalid question batch. Please try again.")
    if any(
        not isinstance(question, dict)
        or len(question.get("options", [])) != 4
        or question.get("correct_answer") not in {"A", "B", "C", "D"}
        or any(not isinstance(option, str) or not option[:2] in {"A)", "B)", "C)", "D)"} for option in question.get("options", []))
        or len(set(question.get("options", []))) != 4
        for question in questions
    ):
        raise ValueError("Gemini returned an invalid option set. Please try again.")
    fingerprints = [_question_fingerprint(question) for question in questions]
    if len(fingerprints) != len(set(fingerprints)):
        raise ValueError("Gemini returned duplicate questions in the same batch. Please try again.")
    quality_issues = [
        issue for question in questions for issue in question_quality_issues(question)
    ]
    if quality_issues:
        raise ValueError(f"Gemini returned a low-quality question batch: {quality_issues[0]}.")
    return audit_questions(difficulty, api_key, questions)


def audit_questions(
    difficulty: str, api_key: str, questions: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Have Gemini independently recalculate answers before they enter the pool."""
    client = genai.Client(api_key=api_key.strip())
    audit_prompt = f"""You are the answer-key reviewer for an economics exam.

Independently solve every question in the candidate batch below. Check all arithmetic, economic definitions, graph or curve logic, the marked correct answer, and whether the explanation actually supports it. Use ONLY the course material as the style and topic reference. Return exactly one final object per candidate in the same order. Preserve a sound question, but correct any wrong option, answer letter, calculation, or explanation you find. Do not mention this review process in the returned fields. The difficulty is {difficulty}.

Also enforce diversity: do not keep multiple questions that are near-duplicates of one another. Reject repetitive verbal patterns, repeated graph setups, and repeated table-only questions. Prefer a mix of graph interpretation, data comparison, policy reasoning, and calculation. Every final question should still be valid, but the set should feel varied and fresh. Require a specific learning object: one measurable objective, a non-revealing hint, 2 to 4 correct worked steps, and a misconception correction tied to a plausible distractor.

Return only a JSON array matching the supplied question schema.

CANDIDATE QUESTIONS:
{json.dumps(questions)}
"""
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=audit_prompt,
        config={
            "response_mime_type": "application/json",
            "response_schema": {"type": "array", "items": QUESTION_SCHEMA},
            "temperature": 0.1,
        },
    )
    audited = json.loads(response.text or "")
    if not isinstance(audited, list) or len(audited) != len(questions):
        raise ValueError("The answer review returned an incomplete question batch. Please try again.")
    for question in audited:
        if (
            not isinstance(question, dict)
            or len(question.get("options", [])) != 4
            or question.get("correct_answer") not in {"A", "B", "C", "D"}
        ):
            raise ValueError("The answer review returned an invalid question. Please try again.")
        quality_issues = question_quality_issues(question)
        if quality_issues:
            raise ValueError(f"The answer review returned a low-quality question: {quality_issues[0]}.")
        question["verified"] = True
    return audited


def build_game_prompt(context_text: str, difficulty: str, track: str) -> str:
    return f"""You are designing Market Shock, a 10-round economics strategy game.

Create exactly {GAME_ROUNDS} distinct, connected multiple-choice economics scenarios for the {track} track at {difficulty} difficulty. Base the concepts, vocabulary, calculations, and difficulty ONLY on the course material below. Each scenario should read like a consequential event in a fictional economy or business. Do not copy source questions.

For each round, include a concise scenario_title and scenario_context, then a question with exactly four options. The correct choice is the economically sound decision or analysis. impact_profile describes changes to four health indicators when the answer is correct: growth, inflation (price stability), employment, and stability. Use integer values from -12 to 12 and include realistic trade-offs when appropriate. Ensure every correct answer, calculation, and explanation is internally consistent.

Return only a JSON array matching the supplied schema.

COURSE MATERIAL:
{context_text[:120000]}
"""


def validate_game_rounds(rounds: list[dict[str, Any]], expected_count: int) -> None:
    if not isinstance(rounds, list) or len(rounds) != expected_count:
        raise ValueError("Gemini returned an incomplete Market Shock game. Please try again.")
    for game_round in rounds:
        impact = game_round.get("impact_profile", {})
        if (
            not isinstance(game_round, dict)
            or len(game_round.get("options", [])) != 4
            or game_round.get("correct_answer") not in {"A", "B", "C", "D"}
            or any(not isinstance(option, str) or option[:2] not in {"A)", "B)", "C)", "D)"} for option in game_round.get("options", []))
            or set(impact) != set(GAME_STARTING_INDICATORS)
            or any(not isinstance(value, int) or not -12 <= value <= 12 for value in impact.values())
        ):
            raise ValueError("Gemini returned an invalid Market Shock round. Please try again.")


def audit_game_rounds(
    difficulty: str, track: str, api_key: str, rounds: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    client = genai.Client(api_key=api_key.strip())
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=f"""You are the answer-key reviewer for the {track} Market Shock game.

Independently solve each candidate round below. Correct any wrong answer letter, arithmetic, economic explanation, scenario logic, or impact_profile that contradicts the economically sound decision. Impact values represent indicator health, so higher inflation means better price stability. Preserve valid rounds and return exactly {len(rounds)} final rounds in the same order. Use ONLY the course material as a topic and style reference.

Return only a JSON array matching the supplied game-round schema.

CANDIDATE ROUNDS:
{json.dumps(rounds)}
""",
        config={
            "response_mime_type": "application/json",
            "response_schema": {"type": "array", "items": GAME_ROUND_SCHEMA},
            "temperature": 0.1,
        },
    )
    audited = json.loads(response.text or "")
    validate_game_rounds(audited, len(rounds))
    for game_round in audited:
        game_round["verified"] = True
    return audited


def generate_game_rounds(
    context_text: str, difficulty: str, track: str, api_key: str
) -> list[dict[str, Any]]:
    if not api_key.strip():
        raise ValueError("A Gemini API key must be configured before starting a game.")
    client = genai.Client(api_key=api_key.strip())
    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=build_game_prompt(context_text, difficulty, track),
        config={
            "response_mime_type": "application/json",
            "response_schema": {"type": "array", "items": GAME_ROUND_SCHEMA},
            "temperature": 0.75,
        },
    )
    rounds = json.loads(response.text or "")
    validate_game_rounds(rounds, GAME_ROUNDS)
    return audit_game_rounds(difficulty, track, api_key, rounds)


def load_saved_game_rounds(
    context_signature: str, difficulty: str, track: str
) -> list[dict[str, Any]]:
    engine = get_database_engine(database_url())
    with engine.connect() as connection:
        rows = connection.execute(text("""
            SELECT id, round_json FROM game_round_bank
            WHERE context_signature = :context_signature AND difficulty = :difficulty AND track = :track
            ORDER BY times_served ASC, created_at ASC
            LIMIT :limit
        """), {
            "context_signature": context_signature,
            "difficulty": difficulty,
            "track": track,
            "limit": GAME_ROUNDS,
        }).mappings().all()
    rounds = []
    for row in rows:
        game_round = json.loads(row["round_json"])
        game_round["bank_id"] = row["id"]
        rounds.append(game_round)
    return rounds


def save_game_rounds(
    context_signature: str, difficulty: str, track: str, rounds: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    engine = get_database_engine(database_url())
    created_at = datetime.now().isoformat(timespec="seconds")
    saved_rounds = []
    with engine.begin() as connection:
        for game_round in rounds:
            round_id = str(uuid.uuid4())
            stored_round = {key: value for key, value in game_round.items() if key != "bank_id"}
            connection.execute(text("""
                INSERT INTO game_round_bank
                (id, created_at, context_signature, difficulty, track, round_json)
                VALUES (:id, :created_at, :context_signature, :difficulty, :track, :round_json)
            """), {
                "id": round_id,
                "created_at": created_at,
                "context_signature": context_signature,
                "difficulty": difficulty,
                "track": track,
                "round_json": json.dumps(stored_round),
            })
            stored_round["bank_id"] = round_id
            saved_rounds.append(stored_round)
    return saved_rounds


def apply_game_impact(impact: dict[str, int], is_correct: bool) -> dict[str, int]:
    multiplier = 1 if is_correct else -0.6
    return {key: int(round(value * multiplier)) for key, value in impact.items()}


def start_game_session(track: str, difficulty: str) -> str:
    session_id = str(uuid.uuid4())
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            INSERT INTO game_sessions (id, started_at, track, difficulty)
            VALUES (:id, :started_at, :track, :difficulty)
        """), {
            "id": session_id,
            "started_at": datetime.now().isoformat(timespec="seconds"),
            "track": track,
            "difficulty": difficulty,
        })
    return session_id


def save_game_round_result(result: dict[str, Any]) -> None:
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            INSERT INTO game_rounds
            (id, game_session_id, round_number, topic, question, selected,
             correct, is_correct, impact_json)
            VALUES (:id, :game_session_id, :round_number, :topic, :question,
                    :selected, :correct, :is_correct, :impact_json)
        """), {
            "id": str(uuid.uuid4()),
            **result,
            "is_correct": int(result["is_correct"]),
            "impact_json": json.dumps(result["impact"]),
        })


def complete_game_session(outcome: str) -> None:
    if st.session_state.game_result_saved or not st.session_state.game_session_id:
        return
    indicators = st.session_state.game_indicators
    final_score = round(sum(indicators.values()) / len(indicators))
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        connection.execute(text("""
            UPDATE game_sessions
            SET completed_at = :completed_at, final_growth = :growth,
                final_inflation = :inflation, final_employment = :employment,
                final_stability = :stability, final_score = :final_score, outcome = :outcome
            WHERE id = :id
        """), {
            "id": st.session_state.game_session_id,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
            "final_score": final_score,
            "outcome": outcome,
            **indicators,
        })
    st.session_state.game_result_saved = True


def start_market_shock(api_key: str, track: str) -> None:
    signature = question_pool_signature()
    cached_rounds = load_saved_game_rounds(signature, st.session_state.difficulty, track)
    if len(cached_rounds) < GAME_ROUNDS:
        with st.spinner("Building and checking your Market Shock scenario..."):
            cached_rounds = save_game_rounds(
                signature,
                st.session_state.difficulty,
                track,
                generate_game_rounds(
                    st.session_state.course_context,
                    st.session_state.difficulty,
                    track,
                    api_key,
                ),
            )
    st.session_state.game_active = True
    st.session_state.game_track = track
    st.session_state.game_questions = cached_rounds[:GAME_ROUNDS]
    st.session_state.game_round = 0
    st.session_state.game_indicators = GAME_STARTING_INDICATORS.copy()
    st.session_state.game_round_results = []
    st.session_state.game_answer = None
    st.session_state.game_revealed = False
    st.session_state.game_over = False
    st.session_state.game_result_saved = False
    st.session_state.game_session_id = start_game_session(track, st.session_state.difficulty)
    engine = get_database_engine(database_url())
    with engine.begin() as connection:
        for game_round in st.session_state.game_questions:
            if game_round.get("bank_id"):
                connection.execute(text("""
                    UPDATE game_round_bank SET times_served = times_served + 1 WHERE id = :id
                """), {"id": game_round["bank_id"]})


def reset_market_shock() -> None:
    st.session_state.game_active = False
    st.session_state.game_questions = []
    st.session_state.game_round_results = []
    st.session_state.game_round = 0
    st.session_state.game_indicators = GAME_STARTING_INDICATORS.copy()
    st.session_state.game_answer = None
    st.session_state.game_revealed = False
    st.session_state.game_over = False
    st.session_state.game_session_id = None
    st.session_state.game_result_saved = False


def generate_question(context_text: str, difficulty: str, api_key: str) -> dict[str, Any]:
    """Keep the single-question API available for callers outside the pool."""
    return generate_questions(context_text, difficulty, api_key, 1)[0]


def record_result(question: dict[str, Any], selected: str | None, mode: str) -> None:
    question_key = f"{mode}:{question.get('question', '')}"
    if question_key in st.session_state.recorded_attempt_questions:
        return
    attempt = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "topic": question.get("topic_tag", "General economics"),
        "difficulty": st.session_state.difficulty,
        "question": question.get("question", ""),
        "selected": selected or "",
        "correct": question.get("correct_answer", ""),
        "is_correct": selected == question.get("correct_answer"),
        "verified": question.get("verified", False),
    }
    save_attempt(attempt)
    st.session_state.history.append(attempt)
    st.session_state.recorded_attempt_questions.add(question_key)


def next_question_style() -> str:
    styles = ["mixed", "graph_heavy", "table_heavy"]
    index = st.session_state.get("question_style_cycle_index", 0) % len(styles)
    style = styles[index]
    st.session_state.question_style_cycle_index = (index + 1) % len(styles)
    return style


def question_pool_signature() -> str:
    return hashlib.sha256(
        f"question-bank-v2\n{st.session_state.course_context}\n{st.session_state.difficulty}".encode()
    ).hexdigest()


def fill_question_pool(api_key: str, minimum: int = QUESTION_POOL_TARGET) -> None:
    signature = question_pool_signature()
    if st.session_state.question_pool_signature != signature:
        st.session_state.question_pool = []
        st.session_state.served_question_ids = set()
        st.session_state.recent_question_fingerprints = []
        st.session_state.recent_topic_tags = []
        st.session_state.recent_concept_families = []
        st.session_state.recent_graph_families = []
        st.session_state.question_pool_signature = signature
    missing = minimum - len(st.session_state.question_pool)
    if missing <= 0:
        return
    saved_questions = load_saved_questions(signature, st.session_state.difficulty)
    saved_question_ids = {
        question.get("bank_id") for question in st.session_state.question_pool
    }
    reusable_questions = [
        question for question in saved_questions
        if question.get("bank_id") not in saved_question_ids
    ]
    st.session_state.question_pool.extend(order_question_pool(reusable_questions)[:missing])
    missing = minimum - len(st.session_state.question_pool)
    if missing <= 0:
        st.session_state.question_pool = order_question_pool(st.session_state.question_pool)
        return
    with st.spinner(f"Creating and checking {missing} new questions for your bank..."):
        question_style = st.session_state.question_style
        if question_style == "auto":
            question_style = next_question_style()
        elif question_style == "mixed" and len(st.session_state.question_pool) % 2 == 0:
            question_style = "graph_heavy"
        generated_questions = generate_questions(
            st.session_state.course_context,
            st.session_state.difficulty,
            api_key,
            missing,
            question_style=question_style,
        )
        st.session_state.question_pool.extend(save_questions_to_bank(
            signature,
            st.session_state.difficulty,
            generated_questions,
        ))
    st.session_state.question_pool = order_question_pool(st.session_state.question_pool)


def take_pooled_question(api_key: str) -> dict[str, Any]:
    if not st.session_state.question_pool:
        fill_question_pool(api_key, minimum=QUESTION_POOL_TARGET)
    elif len(st.session_state.question_pool) <= QUESTION_POOL_REFILL_THRESHOLD:
        fill_question_pool(api_key, minimum=QUESTION_POOL_TARGET)
    question = st.session_state.question_pool.pop(0)
    mark_question_served(question)
    _note_question_seen(question)
    st.session_state.active_question = question
    st.session_state.practice_answer = None
    st.session_state.practice_revealed = False
    return question


def configured_api_key() -> str:
    """Resolve deployment-safe secrets without displaying them."""
    try:
        secret_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        secret_key = ""
    return str(
        secret_key
        or os.getenv("GEMINI_API_KEY", "")
        or os.getenv("GEMINI_KEY", "")
        or os.getenv("$GEMINI_KEY", "")
    )


def render_sidebar() -> tuple[str, str]:
    with st.sidebar:
        st.markdown("## EconMaster AI")
        st.caption("Your course material, turned into a sharper study loop.")
        if st.button("Lock study room"):
            st.session_state.authenticated = False
            st.rerun()
        api_key = configured_api_key()
        if api_key:
            st.caption("Gemini connection: configured")
        else:
            st.warning("Gemini connection: missing API key")
        st.divider()
        uploaded_files = st.file_uploader(
            "Course materials", type=["pdf", "docx"], accept_multiple_files=True,
            help="Upload practice exams, solution sets, or lecture handouts.",
        )
        if uploaded_files:
            ingest_uploads(uploaded_files)
        elif st.session_state.document_names:
            st.caption(f"{len(st.session_state.document_names)} document(s) indexed")
        cache_summary = get_document_cache_summary()
        st.session_state.document_cache_summary = cache_summary
        if cache_summary["cached_files"]:
            st.caption(
                f"Cached extracts: {cache_summary['cached_files']} files · "
                f"{cache_summary['total_bytes'] / 1024:.1f} KB"
            )
            if st.button("Clear cache", key="clear_document_cache_button"):
                clear_document_cache()
                st.rerun()
        elif st.button("Refresh cache", key="refresh_document_cache_button"):
            st.session_state.document_cache_summary = get_document_cache_summary()
            st.rerun()
        if st.session_state.rag_errors:
            st.warning("\n".join(st.session_state.rag_errors))
        st.divider()
        workspace_options = ["Practice", "Learning", "Exam", "Market Shock", "Analytics"]
        mode = st.radio(
            "Workspace", workspace_options,
            index=workspace_options.index(st.session_state.active_mode) if st.session_state.active_mode in workspace_options else 0,
        )
        st.session_state.active_mode = mode
        st.session_state.difficulty = st.select_slider(
            "Question difficulty", options=["Foundational", "Intermediate", "Challenging"],
            value=st.session_state.difficulty,
        )
        style_options = ["auto", "mixed", "graph_heavy", "table_heavy"]
        style_labels = {
            "auto": "Auto rotation",
            "mixed": "Mixed",
            "graph_heavy": "Graph-heavy",
            "table_heavy": "Table-heavy",
        }
        current_style = st.session_state.question_style if st.session_state.question_style in style_options else "auto"
        st.session_state.question_style = st.selectbox(
            "Question style",
            style_options,
            index=style_options.index(current_style),
            format_func=lambda value: style_labels.get(value, value),
        )
        if st.session_state.question_pool:
            st.caption(f"Question bank: {len(st.session_state.question_pool)} verified and ready")
        if st.session_state.course_context:
            st.divider()
            st.markdown("**Indexed sources**")
            for name in st.session_state.document_names:
                st.caption(f"· {name}")
    return api_key, mode


def render_header(mode: str) -> None:
    st.markdown(
        f'<div class="hero"><div class="kicker">ECONOMICS STUDY STUDIO / {mode.upper()}</div>'
        # '<h1>Think like the<br>answer key.</h1>'
        '<p>Turn your own practice exams into an adaptive question set that keeps the concepts familiar and the questions fresh.</p></div>',
        unsafe_allow_html=True,
    )
    total = len(st.session_state.history)
    correct = sum(item["is_correct"] for item in st.session_state.history)
    accuracy = f"{round(correct / total * 100)}%" if total else "—"
    st.markdown(
        f'<div class="metric-strip"><div class="metric"><strong>{len(st.session_state.document_names)}</strong><span>source documents</span></div>'
        f'<div class="metric"><strong>{total}</strong><span>questions attempted</span></div>'
        f'<div class="metric"><strong>{accuracy}</strong><span>overall accuracy</span></div></div>',
        unsafe_allow_html=True,
    )


def infer_graph_type(question_text: str) -> str:
    lower_text = question_text.lower()
    if "ppf" in lower_text or "production possibilities frontier" in lower_text or "opportunity cost" in lower_text:
        return "ppf"
    if any(token in lower_text for token in ("ad-as", "aggregate demand", "aggregate supply", "inflation", "output gap", "real gdp")):
        return "ad_as"
    if any(token in lower_text for token in ("wage", "employment", "labor market", "labor supply", "marginal revenue product", "unemployment")):
        return "labor_market"
    if any(token in lower_text for token in ("elasticity", "total revenue", "price elasticity", "demand curve", "consumer surplus")):
        return "elasticity"
    if any(token in lower_text for token in ("demand", "supply", "equilibrium", "price", "quantity")):
        return "supply_demand"
    if any(token in lower_text for token in ("gdp", "national income", "consumption", "investment", "exports", "imports")):
        return "gdp"
    return "generic"


def build_ppf_plot(question_text: str) -> go.Figure | None:
    coordinate_pairs = re.findall(
        r"\(\s*(\d+(?:\.\d+)?)\s*(?:[A-Za-z][^,()]*)?,\s*(\d+(?:\.\d+)?)\s*(?:[A-Za-z][^,()]*)?\)",
        question_text,
    )
    if len(coordinate_pairs) < 2:
        coordinate_pairs = [(100, 80), (150, 50)]
    xs = [float(x) for x, _ in coordinate_pairs]
    ys = [float(y) for _, y in coordinate_pairs]
    x_max = max(xs) * 1.25 if xs else 200
    y_max = max(ys) * 1.25 if ys else 120
    curve_x = [0.0, x_max * 0.23, x_max * 0.52, x_max * 0.77, x_max]
    curve_y = [y_max * 1.04, y_max * 0.78, y_max * 0.58, y_max * 0.3, 0.0]
    labels = ["Point A", "Point B", "Point C", "Point D", "Point E"]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=curve_x,
        y=curve_y,
        mode="lines",
        line=dict(color="#53c5bb", width=3),
        name="PPF",
    ))
    for idx, (x, y) in enumerate(zip(xs[:5], ys[:5])):
        fig.add_trace(go.Scatter(
            x=[x],
            y=[y],
            mode="markers+text",
            text=[labels[idx] if idx < len(labels) else f"P{idx + 1}"],
            textposition="top center",
            marker=dict(size=11, color="#ff8060"),
            name=labels[idx] if idx < len(labels) else f"P{idx + 1}",
            showlegend=False,
        ))

    x_axis_label = "Solar Panels" if "solar" in question_text.lower() else "Good X"
    y_axis_label = "Wind Turbines" if "wind" in question_text.lower() else "Good Y"
    fig.update_layout(
        title="Production Possibilities Frontier",
        xaxis_title=x_axis_label,
        yaxis_title=y_axis_label,
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    fig.update_xaxes(range=[0, x_max * 1.08], zeroline=False)
    fig.update_yaxes(range=[0, y_max * 1.12], zeroline=False)
    return fig


def build_supply_demand_plot(question_text: str) -> go.Figure | None:
    fig = go.Figure()
    x_values = [0, 12, 24, 36, 48, 60]
    if "inelastic" in question_text.lower() or "elastic" in question_text.lower():
        demand_y = [48, 40, 33, 27, 22, 18]
        supply_y = [10, 15, 22, 29, 36, 43]
        title = "Elasticity and Market Equilibrium"
    else:
        demand_y = [52, 44, 35, 26, 18, 11]
        supply_y = [10, 16, 22, 31, 39, 49]
        title = "Supply and Demand"
    fig.add_trace(go.Scatter(x=x_values, y=demand_y, mode="lines", name="Demand", line=dict(color="#ff8060", width=3)))
    fig.add_trace(go.Scatter(x=x_values, y=supply_y, mode="lines", name="Supply", line=dict(color="#53c5bb", width=3)))
    fig.add_trace(go.Scatter(x=[30], y=[26], mode="markers+text", text=["Equilibrium"], textposition="top center", name="Equilibrium", marker=dict(size=10, color="#f4c95d"), showlegend=False))
    fig.update_layout(
        title=title,
        xaxis_title="Quantity",
        yaxis_title="Price",
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    return fig


def build_labor_market_plot(question_text: str) -> go.Figure | None:
    fig = go.Figure()
    x_values = [0, 20, 40, 60, 80, 100]
    labor_demand = [95, 80, 65, 50, 35, 20]
    labor_supply = [10, 25, 40, 55, 70, 90]
    fig.add_trace(go.Scatter(x=x_values, y=labor_demand, mode="lines", name="Labor demand", line=dict(color="#ff8060", width=3)))
    fig.add_trace(go.Scatter(x=x_values, y=labor_supply, mode="lines", name="Labor supply", line=dict(color="#53c5bb", width=3)))
    fig.add_trace(go.Scatter(x=[52], y=[50], mode="markers+text", text=["Equilibrium wage"], textposition="top center", name="Equilibrium", marker=dict(size=10, color="#f4c95d"), showlegend=False))
    fig.update_layout(
        title="Labor Market",
        xaxis_title="Employment / Labor",
        yaxis_title="Wage",
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    return fig


def build_ad_as_plot(question_text: str) -> go.Figure | None:
    fig = go.Figure()
    real_output = [0, 25, 50, 75, 100]
    ad = [95, 90, 85, 80, 75]
    sras = [60, 70, 80, 90, 100]
    lras = [75, 75, 75, 75, 75]
    fig.add_trace(go.Scatter(x=real_output, y=ad, mode="lines", name="AD", line=dict(color="#ff8060", width=3)))
    fig.add_trace(go.Scatter(x=real_output, y=sras, mode="lines", name="SRAS", line=dict(color="#53c5bb", width=3)))
    fig.add_trace(go.Scatter(x=real_output, y=lras, mode="lines", name="LRAS", line=dict(color="#f4c95d", width=3, dash="dash")))
    fig.add_trace(go.Scatter(x=[75], y=[75], mode="markers+text", text=["Macro equilibrium"], textposition="top center", name="Equilibrium", marker=dict(size=10, color="#9cc7ff"), showlegend=False))
    fig.update_layout(
        title="AD-AS Model",
        xaxis_title="Real GDP",
        yaxis_title="Price Level",
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    return fig


def build_elasticity_plot(question_text: str) -> go.Figure | None:
    fig = go.Figure()
    price = [0, 2, 4, 6, 8, 10]
    quantity = [14, 12, 9, 7, 5, 2]
    fig.add_trace(go.Scatter(x=price, y=quantity, mode="lines", name="Demand", line=dict(color="#ff8060", width=3)))
    fig.add_trace(go.Scatter(x=[6], y=[7], mode="markers+text", text=["Elasticity point"], textposition="top center", marker=dict(size=10, color="#f4c95d"), showlegend=False))
    fig.update_layout(
        title="Elasticity of Demand",
        xaxis_title="Price",
        yaxis_title="Quantity Demanded",
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    return fig


def build_gdp_plot(question_text: str) -> go.Figure | None:
    categories = ["Consumption", "Investment", "Government", "Net Exports"]
    values = [55, 20, 18, 7]
    fig = go.Figure(data=[go.Bar(x=categories, y=values, marker_color=["#ff8060", "#53c5bb", "#f4c95d", "#9cc7ff"])])
    fig.update_layout(
        title="GDP Expenditure Components",
        xaxis_title="Component",
        yaxis_title="Share of GDP (%)",
        template="plotly_dark",
        margin=dict(l=45, r=20, t=45, b=45),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=360,
    )
    return fig


def generate_question_plot(question: dict[str, Any]) -> go.Figure | None:
    question_text = question.get("question", "")
    graph_type = infer_graph_type(question_text)
    if graph_type == "ppf":
        return build_ppf_plot(question_text)
    if graph_type == "ad_as":
        return build_ad_as_plot(question_text)
    if graph_type == "labor_market":
        return build_labor_market_plot(question_text)
    if graph_type == "elasticity":
        return build_elasticity_plot(question_text)
    if graph_type == "supply_demand":
        return build_supply_demand_plot(question_text)
    if graph_type == "gdp":
        return build_gdp_plot(question_text)
    return None


def normalized_question_table(question: dict[str, Any]) -> tuple[str, list[str], list[list[str]]] | None:
    data_table = question.get("data_table")
    if isinstance(data_table, dict):
        title = str(data_table.get("title", "Data table")).strip()
        columns = data_table.get("columns")
        rows = data_table.get("rows")
        if (
            isinstance(columns, list)
            and len(columns) >= 2
            and isinstance(rows, list)
            and rows
            and all(isinstance(row, list) and len(row) == len(columns) for row in rows)
        ):
            return title, [str(column) for column in columns], [
                [str(value) for value in row] for row in rows
            ]

    question_text = question.get("question", "")
    cost_language = ("total cost", "marginal cost", "average variable cost", "fixed cost")
    if not any(phrase in question_text.lower() for phrase in cost_language):
        return None
    pairs = re.findall(
        r"\bQ\s*=\s*(\d+(?:\.\d+)?)\s*(?:units?)?\s*,?\s*TC\s*=\s*\$?\s*(\d+(?:\.\d+)?)",
        question_text,
        flags=re.IGNORECASE,
    )
    if len(pairs) < 2:
        return None
    rows = [[quantity, f"${cost}"] for quantity, cost in pairs]
    return "Total cost schedule", ["Output (Q)", "Total Cost (TC)"], rows


def question_display_text(question: dict[str, Any]) -> str:
    question_text = question.get("question", "")
    if not normalized_question_table(question) or question.get("data_table"):
        return question_text
    return re.sub(
        r"\s+where\s+.*?\bQ\s*=.*?(?=\s+(?:At\s+an\s+output|What\s+are|Which\s+of|Calculate)\b|$)",
        "",
        question_text,
        flags=re.IGNORECASE,
    )


def learning_material(question: dict[str, Any]) -> dict[str, Any]:
    supplied = question.get("learning")
    if isinstance(supplied, dict) and all(supplied.get(key) for key in ("objective", "hint", "worked_steps", "misconception")):
        return supplied
    concept = question_concept_family(question).replace("_", " ")
    table_data = normalized_question_table(question)
    if table_data is not None:
        title, columns, rows = table_data
        return {
            "objective": f"Read a {title.lower()} and identify the relevant economic measure.",
            "hint": f"Start with the row or rows named in the question, then write the formula before substituting values from {columns[0]} and {columns[1]}.",
            "worked_steps": [
                "Identify the requested measure and the observations it requires.",
                "Use the displayed schedule rather than treating every total as a per-unit value.",
                question.get("explanation", "Check the correct option against the definition and arithmetic."),
            ],
            "misconception": "A common error is to use a total where the question requires a change, average, or fixed component. Match the formula to the measure first.",
        }
    return {
        "objective": f"Apply the core logic of {concept} to a new scenario.",
        "hint": "Name the economic relationship that changes first, then trace its consequence before comparing the answer choices.",
        "worked_steps": [
            "Identify the economic concept and the relevant change described in the scenario.",
            "Eliminate options that reverse the direction of the relationship or ignore the stated condition.",
            question.get("explanation", "Check the remaining option against the economic definition."),
        ],
        "misconception": "A tempting wrong answer often describes a related concept but reverses the causal direction or overlooks the question's stated assumption.",
    }


def question_quality_issues(question: dict[str, Any]) -> list[str]:
    issues: list[str] = []
    if len(question.get("question", "").strip()) < 40:
        issues.append("question stem is too short")
    if len(question.get("explanation", "").strip()) < 60:
        issues.append("explanation is too brief")
    data_table = question.get("data_table")
    if data_table is not None and normalized_question_table({"data_table": data_table}) is None:
        issues.append("data table is malformed")
    learning = question.get("learning")
    if learning is not None:
        if not isinstance(learning, dict):
            issues.append("learning material is malformed")
        elif not all(str(learning.get(key, "")).strip() for key in ("objective", "hint", "misconception")):
            issues.append("learning material is incomplete")
        elif not isinstance(learning.get("worked_steps"), list) or not 2 <= len(learning["worked_steps"]) <= 4:
            issues.append("worked steps must contain 2 to 4 items")
    return issues


def learning_priority(question: dict[str, Any]) -> tuple[int, int, int, int, str]:
    concept = question_concept_family(question)
    relevant_attempts = [
        attempt for attempt in st.session_state.history
        if (
            question_concept_family({"topic_tag": attempt.get("topic", "")}) == concept
            or attempt.get("topic", "").strip().lower() == question.get("topic_tag", "").strip().lower()
        )
    ]
    missed = sum(not attempt["is_correct"] for attempt in relevant_attempts)
    attempts = len(relevant_attempts)
    recent_concepts = set(st.session_state.get("recent_concept_families", [])[-4:])
    return (
        0 if missed else 1,
        -missed,
        attempts,
        1 if concept in recent_concepts else 0,
        _question_fingerprint(question),
    )


def learning_focus_summary() -> tuple[str, int, int] | None:
    attempts_by_concept: dict[str, list[dict[str, Any]]] = {}
    for attempt in st.session_state.history:
        concept = question_concept_family({"topic_tag": attempt.get("topic", "")})
        attempts_by_concept.setdefault(concept, []).append(attempt)
    candidates = [
        (concept, attempts)
        for concept, attempts in attempts_by_concept.items()
        if attempts
    ]
    if not candidates:
        return None
    concept, attempts = min(
        candidates,
        key=lambda item: (sum(attempt["is_correct"] for attempt in item[1]) / len(item[1]), -len(item[1])),
    )
    correct = sum(attempt["is_correct"] for attempt in attempts)
    return concept.replace("_", " ").title(), correct, len(attempts)


def take_learning_question(api_key: str) -> dict[str, Any]:
    if not st.session_state.question_pool or len(st.session_state.question_pool) <= QUESTION_POOL_REFILL_THRESHOLD:
        fill_question_pool(api_key, minimum=QUESTION_POOL_TARGET)
    question = min(st.session_state.question_pool, key=learning_priority)
    st.session_state.question_pool.remove(question)
    mark_question_served(question)
    _note_question_seen(question)
    st.session_state.learning_question = question
    st.session_state.learning_answer = None
    st.session_state.learning_revealed = False
    st.session_state.learning_hint_visible = False
    return question


def render_question_visuals(question: dict[str, Any]) -> None:
    table_data = normalized_question_table(question)
    if table_data is not None:
        title, columns, rows = table_data
        st.caption(title)
        st.dataframe(pd.DataFrame(rows, columns=columns), hide_index=True, use_container_width=True)
    plot_figure = generate_question_plot(question)
    if plot_figure is not None:
        st.caption("Visual model")
        st.plotly_chart(plot_figure, use_container_width=True)


def render_question(question: dict[str, Any], api_key: str, mode: str) -> None:
    st.markdown('<div class="question-panel">', unsafe_allow_html=True)
    verification = " · VERIFIED" if question.get("verified") else ""
    st.markdown(f'<div class="question-number">QUESTION {len(st.session_state.history) + 1:02d} · {question.get("topic_tag", "ECONOMICS").upper()}{verification}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="question-text">{question_display_text(question)}</div>', unsafe_allow_html=True)
    render_question_visuals(question)
    options = question.get("options", [])
    selected = st.radio("Choose an answer", options, key=f"answer_{id(question)}", label_visibility="collapsed")
    selected_letter = selected[:1] if selected else ""
    if mode == "Practice":
        if st.button("Check answer", type="primary", key=f"check_{id(question)}"):
            st.session_state.practice_answer = selected_letter
            st.session_state.practice_revealed = True
            record_result(question, selected_letter, mode)
            st.rerun()
        if st.session_state.practice_revealed:
            is_correct = st.session_state.practice_answer == question["correct_answer"]
            st.success("Correct. The economic logic checks out." if is_correct else f"Not quite. The correct answer is {question['correct_answer']}.")
            st.markdown(f'<div class="explanation"><strong>Reasoning</strong><br>{question["explanation"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<p class="small-mono">MIMICS: {question.get("source_reference", "Course material")}</p>', unsafe_allow_html=True)
    else:
        st.session_state.exam_answers[st.session_state.exam_index] = selected_letter
    st.markdown('</div>', unsafe_allow_html=True)


def render_learning(api_key: str) -> None:
    st.markdown('<div class="section-label">Guided practice / retrieve, reason, correct</div>', unsafe_allow_html=True)
    if not st.session_state.course_context:
        st.info("Upload course material before starting a learning path. Learning mode will then prioritize concepts you miss.")
        return
    focus = learning_focus_summary()
    if focus is None:
        st.caption("First pass: the learning path will sample the course broadly, then adapt to your results.")
    else:
        concept, correct, attempts = focus
        st.caption(f"Current review focus: {concept} ({correct} of {attempts} correct).")
    if st.session_state.learning_question is None:
        st.markdown("## Learning mode")
        st.write("Work one verified question at a time. Ask for a hint before committing, then use the worked solution to correct the exact misconception behind a miss.")
        if st.button("Start a learning card", type="primary"):
            try:
                take_learning_question(api_key)
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        return

    question = st.session_state.learning_question
    material = learning_material(question)
    st.markdown('<div class="question-panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="question-number">LEARNING CARD · {question.get("topic_tag", "ECONOMICS").upper()}</div>', unsafe_allow_html=True)
    st.caption(f"Skill: {material['objective']}")
    st.markdown(f'<div class="question-text">{question_display_text(question)}</div>', unsafe_allow_html=True)
    render_question_visuals(question)
    if not st.session_state.learning_revealed and not st.session_state.learning_hint_visible:
        if st.button("Show a hint", key=f"hint_{id(question)}"):
            st.session_state.learning_hint_visible = True
            st.rerun()
    if st.session_state.learning_hint_visible and not st.session_state.learning_revealed:
        st.info(material["hint"])
    selected = st.radio("Choose an answer", question.get("options", []), key=f"learning_answer_{id(question)}", label_visibility="collapsed")
    selected_letter = selected[:1] if selected else ""
    if not st.session_state.learning_revealed:
        if st.button("Check my reasoning", type="primary", key=f"learning_check_{id(question)}"):
            st.session_state.learning_answer = selected_letter
            st.session_state.learning_revealed = True
            record_result(question, selected_letter, "Learning")
            st.rerun()
    else:
        is_correct = st.session_state.learning_answer == question["correct_answer"]
        if is_correct:
            st.success("Correct. Now connect the answer to its underlying rule.")
        else:
            st.error(
                f"The correct answer is {question['correct_answer']}. "
                "Use the reasoning below to repair the model, not just memorize the letter."
            )
        st.markdown("#### Worked reasoning")
        for step_number, step in enumerate(material["worked_steps"], start=1):
            st.markdown(f"{step_number}. {step}")
        st.markdown(f'<div class="explanation"><strong>Watch for this</strong><br>{material["misconception"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<p class="small-mono">MIMICS: {question.get("source_reference", "Course material")}</p>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    if st.session_state.learning_revealed and st.button("Next learning card", type="primary"):
        try:
            take_learning_question(api_key)
            st.rerun()
        except Exception as exc:
            st.error(str(exc))
    if st.session_state.learning_revealed and st.button("Retry this question"):
        st.session_state.learning_answer = None
        st.session_state.learning_revealed = False
        st.session_state.learning_hint_visible = True
        st.rerun()


def render_practice(api_key: str) -> None:
    st.markdown('<div class="section-label">Immediate feedback / one concept at a time</div>', unsafe_allow_html=True)
    if not st.session_state.course_context:
        st.info("Upload a PDF practice exam and a DOCX solution set in the sidebar to begin. A small demo question is available below so you can preview the study loop.")
        question = DEMO_QUESTION
    elif st.session_state.active_question is None:
        st.info("Your source library is ready. Load a small question bank so the next questions arrive instantly.")
        if st.button("Load question bank", type="primary"):
            try:
                take_pooled_question(api_key)
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        return
    else:
        question = st.session_state.active_question
    render_question(question, api_key, "Practice")
    if st.session_state.practice_revealed or question is DEMO_QUESTION:
        if st.button("Next question", type="primary"):
            try:
                if question is DEMO_QUESTION:
                    st.session_state.active_question = None
                else:
                    take_pooled_question(api_key)
                st.rerun()
            except Exception as exc:
                st.error(str(exc))


def render_exam(api_key: str) -> None:
    st.markdown('<div class="section-label">10 questions / results held until submission</div>', unsafe_allow_html=True)
    if st.session_state.exam_submitted:
        render_exam_results()
        if st.button("Start a new exam", type="primary"):
            st.session_state.exam_questions = []
            st.session_state.exam_answers = {}
            st.session_state.exam_index = 0
            st.session_state.exam_submitted = False
            st.session_state.active_question = None
            st.rerun()
        return
    if not st.session_state.course_context:
        st.info("Upload your course material first. Exam mode intentionally hides feedback until all 10 questions are complete.")
        return
    if not st.session_state.exam_questions:
        if st.button("Build 10-question exam", type="primary"):
            try:
                fill_question_pool(api_key, minimum=10)
                st.session_state.exam_questions = [
                    st.session_state.question_pool.pop(0) for _ in range(10)
                ]
                for question in st.session_state.exam_questions:
                    mark_question_served(question)
                st.session_state.active_question = st.session_state.exam_questions[0]
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        return
    index = st.session_state.exam_index
    question = st.session_state.exam_questions[index]
    st.progress(index / 10, text=f"Question {index + 1} of 10")
    render_question(question, api_key, "Exam")
    if index < 9:
        if st.button("Save and continue", type="primary"):
            st.session_state.exam_index += 1
            st.session_state.active_question = st.session_state.exam_questions[st.session_state.exam_index]
            st.rerun()
    else:
        if st.button("Submit exam", type="primary"):
            for exam_index, exam_question in enumerate(st.session_state.exam_questions):
                record_result(exam_question, st.session_state.exam_answers.get(exam_index), "Exam")
            st.session_state.exam_submitted = True
            st.rerun()


def render_exam_results() -> None:
    rows = [item for item in st.session_state.history if item["mode"] == "Exam"][-10:]
    score = sum(item["is_correct"] for item in rows)
    st.markdown(f"## {score} / 10 correct")
    st.caption("Your explanations and topic patterns are now available in Analytics.")
    for index, row in enumerate(rows, 1):
        icon = "✓" if row["is_correct"] else "×"
        st.markdown(f"**{icon} {index}. {row['topic']}** · chosen `{row['selected'] or '—'}` · answer `{row['correct']}`")


def render_indicator_board(indicators: dict[str, int]) -> None:
    labels = {
        "growth": "Growth",
        "inflation": "Price stability",
        "employment": "Employment",
        "stability": "Confidence",
    }
    columns = st.columns(4)
    for column, key in zip(columns, GAME_STARTING_INDICATORS):
        with column:
            st.metric(labels[key], f"{indicators[key]} / 100")
            st.progress(indicators[key] / 100)


def render_market_shock(api_key: str) -> None:
    st.markdown('<div class="section-label">10 rounds / decisions under pressure</div>', unsafe_allow_html=True)
    if not st.session_state.course_context:
        st.info("Market Shock needs your course materials before it can build a scenario.")
        return
    if not st.session_state.game_active:
        st.markdown("## Market Shock")
        st.write("Guide a fictional economy through ten shocks. Strong analysis protects growth, price stability, employment, and confidence.")
        track = st.selectbox("Scenario track", GAME_TRACKS, key="game_track_picker")
        if st.button("Start Market Shock", type="primary"):
            try:
                start_market_shock(api_key, track)
                st.rerun()
            except Exception as exc:
                st.error(str(exc))
        return

    if st.session_state.game_over:
        render_market_shock_results()
        return

    round_index = st.session_state.game_round
    game_round = st.session_state.game_questions[round_index]
    st.progress(round_index / GAME_ROUNDS, text=f"Round {round_index + 1} of {GAME_ROUNDS} · {st.session_state.game_track}")
    render_indicator_board(st.session_state.game_indicators)
    st.markdown('<div class="question-panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="question-number">MARKET SHOCK · {game_round["scenario_title"].upper()} · VERIFIED</div>', unsafe_allow_html=True)
    st.markdown(f"#### {game_round['scenario_context']}")
    st.markdown(f'<div class="question-text">{game_round["question"]}</div>', unsafe_allow_html=True)
    selected = st.radio(
        "Choose your decision",
        game_round["options"],
        key=f"game_answer_{round_index}",
        label_visibility="collapsed",
    )
    selected_letter = selected[:1] if selected else ""
    if not st.session_state.game_revealed:
        if st.button("Commit decision", type="primary"):
            is_correct = selected_letter == game_round["correct_answer"]
            changes = apply_game_impact(game_round["impact_profile"], is_correct)
            st.session_state.game_indicators = {
                key: max(0, min(100, value + changes[key]))
                for key, value in st.session_state.game_indicators.items()
            }
            result = {
                "game_session_id": st.session_state.game_session_id,
                "round_number": round_index + 1,
                "topic": game_round["topic_tag"],
                "question": game_round["question"],
                "selected": selected_letter,
                "correct": game_round["correct_answer"],
                "is_correct": is_correct,
                "impact": changes,
            }
            save_game_round_result(result)
            record_result(game_round, selected_letter, "Market Shock")
            st.session_state.game_round_results.append(result)
            st.session_state.game_answer = selected_letter
            st.session_state.game_revealed = True
            st.rerun()
    else:
        result = st.session_state.game_round_results[-1]
        if result["is_correct"]:
            st.success("Sound decision. Your economy absorbs the shock.")
        else:
            st.error(f"The economy takes a hit. The sound decision was {game_round['correct_answer']}.")
        impact_columns = st.columns(4)
        for column, key in zip(impact_columns, GAME_STARTING_INDICATORS):
            change = result["impact"][key]
            with column:
                st.metric(key.title() if key != "inflation" else "Price stability", f"{change:+d}")
        st.markdown(f'<div class="explanation"><strong>Economic reasoning</strong><br>{game_round["explanation"]}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        crisis = any(value == 0 for value in st.session_state.game_indicators.values())
        final_round = round_index == GAME_ROUNDS - 1
        if crisis or final_round:
            outcome = "Crisis" if crisis else "Stabilized"
            complete_game_session(outcome)
            st.session_state.game_over = True
            if st.button("View final economic report", type="primary"):
                st.rerun()
        elif st.button("Advance to next shock", type="primary"):
            st.session_state.game_round += 1
            st.session_state.game_answer = None
            st.session_state.game_revealed = False
            st.rerun()
        return
    st.markdown('</div>', unsafe_allow_html=True)


def render_market_shock_results() -> None:
    indicators = st.session_state.game_indicators
    score = round(sum(indicators.values()) / len(indicators))
    outcome = "Economic stability secured" if all(value > 0 for value in indicators.values()) else "Economic crisis declared"
    st.markdown(f"## {outcome}")
    st.caption(f"Final national resilience score: {score} / 100")
    render_indicator_board(indicators)
    missed_topics = [
        result["topic"] for result in st.session_state.game_round_results
        if not result["is_correct"]
    ]
    if missed_topics:
        review_topic = pd.Series(missed_topics).value_counts().index[0]
        st.warning(f"Priority review: {review_topic}")
    else:
        st.success("Clean run. You handled every shock correctly.")
    correct_count = sum(result["is_correct"] for result in st.session_state.game_round_results)
    st.write(f"**{correct_count} / {len(st.session_state.game_round_results)}** decisions were economically sound.")
    if st.button("Start another scenario", type="primary"):
        reset_market_shock()
        st.rerun()


def render_analytics() -> None:
    st.markdown('<div class="section-label">Performance / pattern recognition</div>', unsafe_allow_html=True)
    if not st.session_state.history:
        st.info("Complete a practice question or an exam to see your performance patterns here.")
        return
    dataframe = pd.DataFrame(st.session_state.history)
    game_attempts = dataframe[dataframe["mode"] == "Market Shock"]
    if not game_attempts.empty:
        game_accuracy = round(game_attempts["is_correct"].mean() * 100)
        game_topics = game_attempts.groupby("topic", as_index=False).agg(
            attempts=("is_correct", "size"), accuracy=("is_correct", "mean")
        ).sort_values("accuracy")
        weakest_game_topic = game_topics.iloc[0]["topic"]
        game_columns = st.columns(3)
        game_columns[0].metric("Market Shock decisions", len(game_attempts))
        game_columns[1].metric("Game decision accuracy", f"{game_accuracy}%")
        game_columns[2].metric("Game review priority", weakest_game_topic)
        st.divider()
    topic_summary = dataframe.groupby("topic", as_index=False).agg(
        attempts=("is_correct", "size"), accuracy=("is_correct", "mean")
    )
    topic_summary["accuracy"] = (topic_summary["accuracy"] * 100).round(0)
    left, right = st.columns([1.1, 1])
    with left:
        st.markdown("#### Topic mastery")
        chart = px.bar(topic_summary.sort_values("accuracy"), x="accuracy", y="topic", orientation="h", text="accuracy", color="accuracy", color_continuous_scale=["#e9694a", "#227c79"])
        chart.update_layout(
            template="plotly_dark",
            height=360,
            margin=dict(l=0, r=0, t=10, b=0),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#f2f5f4", family="Space Grotesk"),
            xaxis_title="Accuracy %",
            yaxis_title="",
            coloraxis_showscale=False,
        )
        st.plotly_chart(chart, use_container_width=True)
    with right:
        st.markdown("#### Attempts")
        st.dataframe(topic_summary.rename(columns={"topic": "Topic", "attempts": "Attempts", "accuracy": "Accuracy %"}), hide_index=True, use_container_width=True)
    st.markdown("#### Recent work")
    recent = dataframe.tail(12).copy()
    recent["result"] = recent["is_correct"].map({True: "Correct", False: "Review"})
    st.dataframe(recent[["timestamp", "mode", "topic", "difficulty", "result"]].rename(columns={"timestamp": "When", "mode": "Mode", "topic": "Topic", "difficulty": "Level", "result": "Result"}), hide_index=True, use_container_width=True)


def main() -> None:
    init_state()
    if not st.session_state.authenticated:
        render_login()
        return
    try:
        initialize_database()
        load_history()
    except Exception as exc:
        st.error(f"Could not connect to the history database: {exc}")
        st.stop()
    load_rag_folder()
    refresh_course_context()
    api_key, mode = render_sidebar()
    render_header(mode)
    if mode == "Practice":
        render_practice(api_key)
    elif mode == "Learning":
        render_learning(api_key)
    elif mode == "Exam":
        render_exam(api_key)
    elif mode == "Market Shock":
        render_market_shock(api_key)
    else:
        render_analytics()


if __name__ == "__main__":
    main()
