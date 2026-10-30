# Sayraa AI PPT Generator

Sayraa creates editable PowerPoint presentations from a topic using Google Gemini and Flask.

## Features

- Four visual templates and seven output languages
- Configurable outline of 3–12 content slides
- One structured Gemini request produces the slide plan, bullets, and colors
- Four to five concise topic points per content slide
- In-memory PPTX generation; generated files are not persisted by the app
- Responsive web interface and PWA shell

## Requirements

- Python 3.12 (see `.python-version` and `runtime.txt`)
- A Gemini API key from [Google AI Studio](https://aistudio.google.com/apikey)

## Run locally

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Set `GEMINI_API_KEY` in `.env`, then start the server:

```powershell
python app.py
```

The app runs at `http://localhost:5000`. `GEMINI_MODEL` can override the default `gemini-3.5-flash-lite`. Keep `.env` private; it is excluded from Git. Use hosting-provider secrets for deployments and leave Flask debug mode disabled in production.

## Deployment

The included `Procfile` uses Gunicorn, and `vercel.json` configures a Vercel Python function. Set `GEMINI_API_KEY` in the deployment environment. Generation is designed around one Gemini request per presentation; response time still depends on Gemini and network availability.

## Project map

- `app.py` — Flask routes, Gemini request, and PPTX generation
- `templates/index.html` — generator form
- `static/` — styles, browser interactions, and PWA assets
- `AGENTS.md` — repository engineering and Git approval rules
- `.codex/rules/git-approval.md` — Git command policy
- `.env.example` — environment variable template with no secrets
