# 🤖 Auto Apply Agent

An AI-powered career agent that scrapes job postings, analyses your fit, generates tailored LaTeX CVs and cover letters, 
and auto-fills application forms — all powered by the Claude API.

## Features

- 🔍 Scrapes any job URL automatically
- 🎯 Analyses job fit with a match score and explanation
- 📄 Generates a tailored LaTeX CV ready for Overleaf
- ✉️ Writes a cover letter matching your tone and style
- 🌐 Auto-fills application forms using Playwright
- 🐙 Fetches your live GitHub profile (cached daily)
- 📊 Logs every application with status and score
- ✅ You always have final say before anything is submitted

## Tech Stack

- Python 3.14
- Anthropic Claude API (Haiku)
- Playwright (browser automation)
- pdfplumber / python-docx (CV parsing)
- BeautifulSoup (web scraping)


## How It Works

```
Paste a job URL
      ↓
Agent scrapes the job page
      ↓
Claude analyses your fit (0-100 score)
      ↓
You decide → yes / no / ask why
      ↓
Claude generates tailored LaTeX CV + cover letter
      ↓
You review and request changes if needed
      ↓
Browser opens and fills the form automatically
      ↓
You do a final check → SUBMIT or CANCEL
      ↓
Everything is logged and saved
```

---

## Project Structure

```
auto-apply-agent/
├── career_agent.py       # Main agent
├── requirements.txt
├── .env                  # Your API key (not committed)
└── data/
    ├── my_skills.json    # Your skills and profile
    ├── links.json        # LinkedIn, GitHub, Portfolio
    ├── raw/              # Your CV files
    ├── chats/            # Writing style exports
    ├── processed/        # Generated master profile
    └── past_applications/# Saved CVs and cover letters
```

---

## Cost

Approximately **$3/month** applying to 10 jobs per day using Claude Haiku — less than a coffee.
