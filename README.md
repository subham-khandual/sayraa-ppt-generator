# Sayraa AI PPT Generator — AI Presentation Generator for PowerPoint Slides

An AI-powered PowerPoint generator that creates presentation slides automatically using advanced language models. This project helps users convert ideas into professional slide decks with minimal effort and faster turnaround time.

**Developer:** [@subham-khandual](https://github.com/subham-khandual)

**Keywords:** AI presentation generator, PowerPoint generator, AI slide maker, automatic PPT generator, LLM presentation tool, AI content generation, presentation automation, pptx generator, AI slides

## Overview

**Sayraa AI PPT Generator** is a Python-based AI presentation tool designed to generate presentation content automatically from user topics. By combining large language models with PowerPoint automation, it helps users create polished slides for business, academic, and professional presentations.

The project is ideal for:
- Students
- Professionals
- Business teams
- Research presenters
- Content creators

## Why this project matters

Creating professional presentations can be time-consuming and repetitive. This project solves that by automating:
- Content generation from a topic or idea
- Structured slide formation
- PowerPoint file generation
- Faster presentation creation workflows

## Features

- ✅ AI-powered content generation for slides
- ✅ Automatic PPTX file creation
- ✅ Professional presentation structure
- ✅ LLM-based content generation using Google Generative AI and Groq
- ✅ Responsive web interface
- ✅ PWA support for offline-friendly usage
- ✅ Easy deployment with Flask

## Tech Stack

- Python
- Flask
- OpenAI/AI model integration via Google Generative AI and Groq
- python-pptx
- HTML
- CSS
- JavaScript
- PWA assets and service worker support

## Project Structure

```bash
ppt-generator/
├── app.py
├── requirements.txt
├── .env.example
├── Procfile
├── runtime.txt
├── static/
│   ├── logo.png
│   ├── manifest.json
│   ├── script.js
│   └── styles.css
├── templates/
│   └── index.html
├── README.md
└── .gitignore
```

## Requirements

- Python 3.8+
- Flask
- python-pptx
- google-generativeai
- groq
- Pillow
- lxml

## Installation

### Step 1: Clone the repository

```bash
git clone https://github.com/subham-khandual/sayraa-ppt-generator.git
cd sayraa-ppt-generator
```

### Step 2: Create a virtual environment

```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install dependencies

```bash
pip install -r requirements.txt
```

### Step 4: Add environment variables

Create a `.env` file:

```env
GEMINI_API_KEY=your_google_generative_ai_key
GROQ_API_KEY=your_groq_api_key
FLASK_ENV=development
```

### Step 5: Run the app

```bash
python app.py
```

The application will run at `http://localhost:5000`

## How it works

1. User enters a presentation topic
2. AI generates slides and content
3. Presentation content is structured into sections
4. PowerPoint file is generated automatically
5. User downloads the final deck

## Deployment

### Heroku

```bash
heroku create your-app-name
git push heroku main
```

### Vercel

```bash
vercel
```

## Developer

Created by **@subham-khandual**  
GitHub: https://github.com/subham-khandual

## SEO Summary

Sayraa AI PPT Generator is an AI presentation generator that creates professional slides automatically from user prompts. It uses advanced AI models to generate slide content, structure presentations, and produce PowerPoint files for faster and smarter presentation creation.

---

**Project Name:** Sayraa AI PPT Generator  
**Category:** AI Presentation Generator, PowerPoint Automation, AI Content Creation  
**Audience:** Students, professionals, business teams, educators, research presenters
