from __future__ import annotations

import hashlib
import hmac
import io
import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv

try:
    from sqlalchemy import create_engine, text
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
        "difficulty": "Intermediate",
        "question_loading": False,
        "question_pool": [],
        "question_pool_signature": "",
        "served_question_ids": set(),
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


def load_saved_questions(context_signature: str, difficulty: str) -> list[dict[str, Any]]:
    """Load the least-used verified questions for this exact course library."""
    engine = get_database_engine(database_url())
    with engine.connect() as connection:
        rows = connection.execute(text("""
            SELECT id, question_json FROM question_bank
            WHERE context_signature = :context_signature AND difficulty = :difficulty
            ORDER BY times_served ASC, created_at ASC
            LIMIT 100
        """), {
            "context_signature": context_signature,
            "difficulty": difficulty,
        }).mappings().all()
    saved_questions = []
    for row in rows:
        if row["id"] in st.session_state.served_question_ids:
            continue
        question = json.loads(row["question_json"])
        question["bank_id"] = row["id"]
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
            text = extract_text(source_path.name, source_path.read_bytes()).strip()
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
            text = extract_text(uploaded_file.name, uploaded_file.getvalue()).strip()
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


def build_prompt(context_text: str, difficulty: str, question_count: int = 1) -> str:
    limited_context = context_text[:120000]
    quantity = "one" if question_count == 1 else f"exactly {question_count}"
    return f"""You are an economics professor and assessment designer.

Create {quantity} brand-new multiple-choice question(s) based ONLY on the course material below. Mimic its style, difficulty, vocabulary, topic distribution, and solution depth. Do not copy wording or simply reproduce a question. Change the scenario, values, or framing while testing the same economic principles. The requested difficulty is {difficulty}.

Return only valid JSON matching the supplied schema. Return an array with exactly {question_count} objects. Each object's options must be exactly four strings beginning with A), B), C), and D). correct_answer must be one letter. Make each explanation step-by-step and include formulas or calculations where relevant. Before returning, independently solve every question and check that the marked answer, numerical values, and explanation agree. Make the questions distinct from one another. source_reference should briefly name the concept or pattern from the source material that this new question mirrors.

COURSE MATERIAL:
{limited_context}
"""


def generate_questions(
    context_text: str, difficulty: str, api_key: str, question_count: int
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
        contents=build_prompt(context_text, difficulty, question_count),
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
    return audit_questions(difficulty, api_key, questions)


def audit_questions(
    difficulty: str, api_key: str, questions: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    """Have Gemini independently recalculate answers before they enter the pool."""
    client = genai.Client(api_key=api_key.strip())
    audit_prompt = f"""You are the answer-key reviewer for an economics exam.

Independently solve every question in the candidate batch below. Check all arithmetic, economic definitions, graph or curve logic, the marked correct answer, and whether the explanation actually supports it. Use ONLY the course material as the style and topic reference. Return exactly one final object per candidate in the same order. Preserve a sound question, but correct any wrong option, answer letter, calculation, or explanation you find. Do not mention this review process in the returned fields. The difficulty is {difficulty}.

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


def question_pool_signature() -> str:
    return hashlib.sha256(
        f"{st.session_state.course_context}\n{st.session_state.difficulty}".encode()
    ).hexdigest()


def fill_question_pool(api_key: str, minimum: int = QUESTION_POOL_TARGET) -> None:
    signature = question_pool_signature()
    if st.session_state.question_pool_signature != signature:
        st.session_state.question_pool = []
        st.session_state.served_question_ids = set()
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
    st.session_state.question_pool.extend(reusable_questions[:missing])
    missing = minimum - len(st.session_state.question_pool)
    if missing <= 0:
        return
    with st.spinner(f"Creating and checking {missing} new questions for your bank..."):
        generated_questions = generate_questions(
            st.session_state.course_context,
            st.session_state.difficulty,
            api_key,
            missing,
        )
        st.session_state.question_pool.extend(save_questions_to_bank(
            signature,
            st.session_state.difficulty,
            generated_questions,
        ))


def take_pooled_question(api_key: str) -> dict[str, Any]:
    if not st.session_state.question_pool:
        fill_question_pool(api_key, minimum=QUESTION_POOL_TARGET)
    elif len(st.session_state.question_pool) <= QUESTION_POOL_REFILL_THRESHOLD:
        fill_question_pool(api_key, minimum=QUESTION_POOL_TARGET)
    question = st.session_state.question_pool.pop(0)
    mark_question_served(question)
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
        if st.session_state.rag_errors:
            st.warning("\n".join(st.session_state.rag_errors))
        st.divider()
        mode = st.radio(
            "Workspace", ["Practice", "Exam", "Market Shock", "Analytics"],
            index=["Practice", "Exam", "Market Shock", "Analytics"].index(st.session_state.active_mode),
        )
        st.session_state.active_mode = mode
        st.session_state.difficulty = st.select_slider(
            "Question difficulty", options=["Foundational", "Intermediate", "Challenging"],
            value=st.session_state.difficulty,
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


def render_question(question: dict[str, Any], api_key: str, mode: str) -> None:
    st.markdown('<div class="question-panel">', unsafe_allow_html=True)
    verification = " · VERIFIED" if question.get("verified") else ""
    st.markdown(f'<div class="question-number">QUESTION {len(st.session_state.history) + 1:02d} · {question.get("topic_tag", "ECONOMICS").upper()}{verification}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="question-text">{question["question"]}</div>', unsafe_allow_html=True)
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
    elif mode == "Exam":
        render_exam(api_key)
    elif mode == "Market Shock":
        render_market_shock(api_key)
    else:
        render_analytics()


if __name__ == "__main__":
    main()
