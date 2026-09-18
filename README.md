# EconMaster AI

A Streamlit economics study studio that turns your own practice exams and solution sets into fresh, course-grounded questions with Google Gemini.

## Study Workspaces

- **Practice:** immediate answer feedback using the verified question bank.
- **Learning:** guided cards with a learning objective, optional hint, worked reasoning, misconception correction, and adaptive priority for concepts you have missed.
- **Exam:** a ten-question assessment with answers held until submission.
- **Market Shock:** a scenario-based economics strategy game.
- **Analytics:** topic-level performance patterns across study activity.

Question generation supports structured tables for total-cost schedules, demand and supply schedules, GDP components, macroeconomic data, and comparative-advantage comparisons. Graph questions use concept-specific visual models rather than a single generic chart.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

The app automatically indexes every PDF and DOCX in `rag_data/` when the session starts. You can upload additional selectable-text PDFs and DOCX files from the sidebar; they are added to the bundled course library.

## Gemini API key

For local use, keep the key in the ignored `.env` file at the project root:

```dotenv
GEMINI_API_KEY=your_key_here
```

The app loads this with `python-dotenv`. The key is never displayed, stored in the course context, or entered by end users.

The bundled `rag_data/ABE Accelerated Economics Reference.md` provides a compact concept map for the accelerated ABE/MBA sequence: scarcity, PPFs, trade, supply and demand, elasticity, production costs, labor markets, GDP, inflation, aggregate demand/supply, policy, and market structures. Uploaded exams and solutions remain the primary style source. Gemini receives the full compact reference plus balanced excerpts from every other source, rather than only the first part of the first long PDF.

For Streamlit Community Cloud or another hosted deployment, do not deploy `.env`. Add this under the app's **Secrets** settings instead:

```toml
GEMINI_API_KEY = "your_key_here"
```

The app checks Streamlit secrets first, then the process environment, then `.env`. The `rag_data/` folder must be committed or otherwise included in the deployment because it is the source library used for generation.

## Password and historical data

The app has a lightweight shared-password gate. Configure it locally in `.env`:

```dotenv
APP_PASSWORD=choose-a-password
```

For Streamlit Community Cloud, add the following under the app's **Secrets** settings:

```toml
APP_PASSWORD = "choose-a-password"
DATABASE_URL = "postgresql+psycopg://user:password@host:5432/database"
```

`APP_PASSWORD` protects the study UI for casual/private use. It is a shared password rather than a full user-management system, so do not use it for sensitive data.

The app uses SQLite in local development and stores the file as `history.db`. On a hosted Streamlit deployment, use an external Postgres database through `DATABASE_URL`; the local filesystem may be reset when the app restarts. Tables are created automatically. Completed attempts, verified Practice/Exam questions, verified Market Shock rounds, game sessions, and game outcomes are stored in the database. The app reuses cached questions for the same course-material fingerprint, difficulty, and game track before making another Gemini request.

Install dependencies with `python -m pip install -r requirements.txt`, not `pip install requirements.txt`.

## Modes

- **Practice:** immediate answer checks and step-by-step explanations.
- **Exam:** a 10-question run with feedback held until submission.
- **Market Shock:** guide a fictional economy through 10 course-grounded decisions; audited game rounds are cached and reused to reduce Gemini calls.
- **Analytics:** topic-level mastery, accuracy, and recent attempts.

Scanned/image-only PDFs need OCR before upload because `PyPDF2` extracts selectable text rather than pixels.
