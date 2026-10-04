# Sayraa AI PPT Generator

An intelligent PowerPoint presentation generator powered by AI. This application leverages advanced language models (Google Generative AI and Groq) to automatically generate professional presentations from text input.

## 🌟 Features

- **AI-Powered Content Generation**: Uses Google Generative AI and Groq to create presentation content
- **Automatic Slide Creation**: Converts text input into structured slides
- **Web-Based Interface**: User-friendly Flask web application
- **Progressive Web App (PWA)**: Works offline with service worker support
- **Professional Templates**: Generates well-formatted PowerPoint presentations
- **Responsive Design**: Works seamlessly on desktop and mobile devices

## 🗂️ Project Structure

```
ppt-generator/
├── .gitignore
├── .python-version
├── Procfile
├── README.md
├── app.py
├── requirements.txt
├── runtime.txt
├── static/
│   ├── logo.png
│   ├── manifest.json
│   ├── script.js
│   └── styles.css
└── templates/
    └── index.html
```

### Directory Details

**Root Level Files:**
- **`.gitignore`** - Git ignore rules for version control
- **`.python-version`** - Python version specification (for pyenv)
- **`Procfile`** - Heroku deployment configuration
- **`README.md`** - Project documentation
- **`app.py`** - Main Flask application entry point (Backend logic)
- **`requirements.txt`** - Python dependencies list
- **`runtime.txt`** - Runtime version for deployment

**`static/` Directory - Frontend Assets:**
- **`logo.png`** - Application logo (~175 KB, optimized)
- **`manifest.json`** - PWA (Progressive Web App) manifest configuration
- **`script.js`** - Client-side JavaScript logic and interactions
- **`styles.css`** - Application styling and responsive design

**`templates/` Directory - HTML Templates:**
- **`index.html`** - Main application interface (Jinja2 template)

## 📋 Requirements

- Python 3.8+
- Flask
- python-pptx
- google-generativeai
- groq
- Pillow
- lxml

## 🚀 Installation

### Step 1: Clone the repository
```bash
git clone https://github.com/SubhamKhandual007/ppt-generator.git
cd ppt-generator
```

### Step 2: Create a virtual environment
```bash
# On macOS/Linux
python3 -m venv venv
source venv/bin/activate

# On Windows
python -m venv venv
venv\Scripts\activate
```

### Step 3: Install dependencies
```bash
pip install -r requirements.txt
```

### Step 4: Set up environment variables
Create a `.env` file in the root directory:
```env
GEMINI_API_KEY=your_google_generative_ai_key
GROQ_API_KEY=your_groq_api_key
FLASK_ENV=development
```

### Step 5: Run the application
```bash
python app.py
```

The application will be available at `http://localhost:5000`

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| Flask | Web framework for building the application |
| python-pptx | PowerPoint file generation and manipulation |
| google-generativeai | Google's Generative AI API integration |
| groq | Groq API for alternative LLM provider |
| Pillow | Image processing and manipulation |
| python-dotenv | Environment variable management from .env files |
| gunicorn | Production WSGI HTTP server |
| lxml | XML and HTML processing |
| setuptools | Package installation and distribution tools |

## 🎯 How It Works

1. **User Input**: Users provide presentation topic via web interface
2. **AI Processing**: Application queries Groq or Google Generative AI
3. **Content Generation**: AI generates structured slide content
4. **PPTX Creation**: python-pptx creates professional PowerPoint file
5. **Download**: Users download the generated presentation

## 🌐 Deployment

### Heroku Deployment
```bash
heroku create your-app-name
git push heroku main
```

### Vercel Deployment
```bash
vercel
```

## 🔐 API Keys

This project requires:
- **Google Generative AI (Gemini)**: [Get Key](https://makersuite.google.com/app/apikey)
- **Groq**: [Get Key](https://console.groq.com/)

Store in `.env` file (never commit to version control).

## 📱 Progressive Web App

The application includes PWA support:
- `manifest.json` - App configuration
- `sw.js` - Service worker for offline functionality

## 🎨 Frontend

- **HTML** - `templates/index.html`
- **CSS** - `static/styles.css`
- **JavaScript** - `static/script.js`
- **Logo** - `static/logo.png`

## 📝 License

MIT License

## 👤 Author

**Subham Khandual**
- GitHub: [@SubhamKhandual](https://github.com/subham-khandual)

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

---

**Last Updated**: May 22, 2026
