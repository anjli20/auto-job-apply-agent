import os
import json
import requests
import time
from pathlib import Path
from dotenv import load_dotenv
import anthropic
import pdfplumber
from docx import Document
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

load_dotenv()
client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

DATA_DIR = Path("data")
RAW_DIR = DATA_DIR / "raw"
PROCESSED_DIR = DATA_DIR / "processed"
PAST_APPS_DIR = DATA_DIR / "past_applications"
CHATS_DIR = DATA_DIR / "chats"
LOG_FILE = DATA_DIR / "applications_log.json"

for d in [RAW_DIR, PROCESSED_DIR, PAST_APPS_DIR, CHATS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ─────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────

def load_log():
    if LOG_FILE.exists():
        return json.loads(LOG_FILE.read_text())
    return []

def save_log(log):
    LOG_FILE.write_text(json.dumps(log, indent=2))

def log_application(url, company, role, score, status):
    log = load_log()
    log.append({
        "url": url,
        "company": company,
        "role": role,
        "match_score": score,
        "status": status,
        "date": time.strftime("%Y-%m-%d %H:%M")
    })
    save_log(log)

# ─────────────────────────────────────────
# GITHUB — SMART CACHED FETCH
# ─────────────────────────────────────────

def fetch_url_text(url, label):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(resp.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer"]):
            tag.decompose()
        return soup.get_text(separator="\n", strip=True)[:5000]
    except Exception as e:
        print(f"  ⚠️ Could not fetch {label}: {e}")
        return ""

def fetch_github_live(username, force_refresh=False):
    """
    Smart GitHub fetch:
    - Uses cached version if less than 24 hours old
    - Fetches live if cache is stale or force_refresh=True
    - Saves cache to data/processed/github_cache.json
    This avoids fetching GitHub on every single API call
    which would waste tokens and slow things down.
    """
    cache_path = PROCESSED_DIR / "github_cache.json"

    # Check cache first
    if not force_refresh and cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text())
            cached_time = cache.get("fetched_at", 0)
            age_hours = (time.time() - cached_time) / 3600

            if age_hours < 24:
                print(
                    f"  🐙 GitHub: using cache "
                    f"({age_hours:.1f}hrs old — "
                    f"refreshes after 24hrs)"
                )
                return cache.get("data", "")
            else:
                print(
                    f"  🐙 GitHub cache is "
                    f"{age_hours:.0f}hrs old — refreshing..."
                )
        except:
            pass

    # Fetch live from GitHub API
    print(f"  🐙 Fetching LIVE GitHub for @{username}...")
    try:
        # Basic profile info
        profile_resp = requests.get(
            f"https://api.github.com/users/{username}",
            headers={"Accept": "application/vnd.github.v3+json"},
            timeout=10
        ).json()

        bio = profile_resp.get("bio", "")
        public_repos = profile_resp.get("public_repos", 0)

        # All repos sorted by last updated
        repos_resp = requests.get(
            f"https://api.github.com/users/{username}"
            f"/repos?sort=updated&per_page=30",
            headers={"Accept": "application/vnd.github.v3+json"},
            timeout=10
        ).json()

        repos_summary = []
        languages_used = set()

        if isinstance(repos_resp, list):
            for repo in repos_resp:
                # Skip forks — focus on original work
                if repo.get("fork"):
                    continue
                name = repo.get("name", "")
                desc = repo.get("description") or "No description"
                lang = repo.get("language") or "Unknown"
                stars = repo.get("stargazers_count", 0)
                updated = repo.get("updated_at", "")[:10]
                repo_url = repo.get("html_url", "")
                languages_used.add(lang)
                repos_summary.append(
                    f"- [{name}]({repo_url})\n"
                    f"  Language: {lang} | "
                    f"Stars: {stars} | "
                    f"Updated: {updated}\n"
                    f"  {desc}"
                )

        output = f"""GITHUB PROFILE: @{username}
Bio: {bio}
Public Repos: {public_repos}
Languages Used: {', '.join(languages_used - {'Unknown'})}

REPOSITORIES (original only, recently updated):
{chr(10).join(repos_summary[:20])}
"""
        # Save to cache with timestamp
        cache_data = {
            "fetched_at": time.time(),
            "fetched_date": time.strftime("%Y-%m-%d %H:%M"),
            "username": username,
            "data": output
        }
        cache_path.write_text(
            json.dumps(cache_data, indent=2),
            encoding="utf-8"
        )

        print(
            f"  ✅ GitHub: {len(repos_summary)} original repos"
            f" — cached for 24hrs"
        )
        return output

    except Exception as e:
        print(f"  ⚠️ GitHub live fetch failed: {e}")
        # Return stale cache if available rather than nothing
        if cache_path.exists():
            print("  ℹ️ Using stale cache as fallback")
            try:
                cache = json.loads(cache_path.read_text())
                return cache.get("data", "")
            except:
                pass
        return ""

# ─────────────────────────────────────────
# PROFILE BUILDER
# ─────────────────────────────────────────

def extract_pdf(filepath):
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
    except Exception as e:
        print(f"  ⚠️ Could not read {filepath.name}: {e}")
    return text

def extract_docx(filepath):
    try:
        doc = Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"  ⚠️ Could not read {filepath.name}: {e}")
        return ""

def extract_all_cvs():
    """
    Reads all PDF and Word CVs from data/raw/
    Keep only 2-3 real CVs in that folder.
    Too many files = too many tokens = higher cost.
    """
    all_text = ""
    cv_files = (
        list(RAW_DIR.glob("*.pdf")) +
        list(RAW_DIR.glob("*.docx"))
    )
    if not cv_files:
        print("❌ No CV files found in data/raw/")
        print("   Add your CV PDF or Word file there.")
        return ""
    for cv in cv_files:
        print(f"  📄 Reading {cv.name}...")
        if cv.suffix == ".pdf":
            all_text += (
                f"\n\n=== CV: {cv.name} ===\n"
                + extract_pdf(cv)
            )
        elif cv.suffix == ".docx":
            all_text += (
                f"\n\n=== CV: {cv.name} ===\n"
                + extract_docx(cv)
            )
    return all_text

def extract_chats():
    """Read Claude chat exports from data/chats/"""
    all_text = ""
    chat_files = (
        list(CHATS_DIR.glob("*.txt")) +
        list(CHATS_DIR.glob("*.md"))
    )
    for chat in chat_files:
        print(f"  💬 Reading {chat.name}...")
        all_text += (
            f"\n\n=== CHAT: {chat.name} ===\n" +
            chat.read_text(encoding="utf-8", errors="ignore")
        )
    return all_text

def build_profile():
    print("\n🔨 Building your master profile...\n")
    profile = {}

    # Count CV files and warn if too many
    cv_files = (
        list(RAW_DIR.glob("*.pdf")) +
        list(RAW_DIR.glob("*.docx"))
    )
    if len(cv_files) > 4:
        print(
            f"  ⚠️ WARNING: {len(cv_files)} CV files found."
        )
        print(
            "  Keep only 2-3 best CVs to save tokens and cost."
        )
        print(
            "  Delete random/duplicate files from data/raw/\n"
        )

    print("📄 Loading CVs...")
    profile["cvs"] = extract_all_cvs()

    print("💬 Loading chat exports...")
    profile["chats"] = extract_chats()

    links_path = DATA_DIR / "links.json"
    if links_path.exists():
        links = json.loads(links_path.read_text())
        if links.get("linkedin"):
            print("  🔗 Fetching LinkedIn...")
            profile["linkedin"] = fetch_url_text(
                links["linkedin"], "LinkedIn"
            )
        if links.get("portfolio"):
            print("  🌐 Fetching Portfolio...")
            profile["portfolio"] = fetch_url_text(
                links["portfolio"], "Portfolio"
            )
        print(
            "  🐙 GitHub uses smart cache "
            "(fetched once per day per application)"
        )
    else:
        print("  ℹ️ No links.json — run option 2 first")

    with open(
        PROCESSED_DIR / "master_profile.json",
        "w", encoding="utf-8"
    ) as f:
        json.dump(profile, f, indent=2, ensure_ascii=False)

    total = sum(len(v) for v in profile.values() if v)
    tokens_estimate = total // 4
    print(f"\n✅ Profile saved!")
    print(f"   Size: ~{tokens_estimate:,} tokens")

    if tokens_estimate > 20000:
        print(
            f"\n  ⚠️ Profile is large ({tokens_estimate:,} tokens)."
        )
        print(
            "  Consider reducing CV files in data/raw/"
        )
    else:
        print("   ✅ Good size — efficient for API calls")

def load_profile():
    path = PROCESSED_DIR / "master_profile.json"
    if not path.exists():
        print("❌ No profile found. Run option 1 first.")
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)

def load_skills():
    skills_path = DATA_DIR / "my_skills.json"
    if skills_path.exists():
        return json.loads(skills_path.read_text())
    print("  ⚠️ No my_skills.json found in data/")
    return {}

# ─────────────────────────────────────────
# CONTEXT BUILDER
# ─────────────────────────────────────────

def build_context(profile, skills, github_username=None):
    """
    Combines all knowledge sources into one context string.
    GitHub uses smart cache — fetched once per day max.
    CVs trimmed to 3000 chars to keep tokens low.
    """
    context = ""

    # ── Skills file ──────────────────────────
    if skills:
        personal = skills.get("personal", {})
        context += "CANDIDATE SKILLS AND BACKGROUND:\n"
        context += f"Name: {personal.get('name', '')}\n"
        context += f"Title: {personal.get('title', '')}\n"
        context += f"Location: {personal.get('location', '')}\n"
        context += f"Email: {personal.get('email', '')}\n"
        context += f"Phone: {personal.get('phone', '')}\n"
        context += (
            f"Experience: "
            f"{skills.get('experience_years', '?')} years\n"
        )
        context += (
            f"Domains: "
            f"{', '.join(skills.get('domains', []))}\n"
        )

        edu = skills.get("education", {})
        if isinstance(edu, dict):
            pg = edu.get("postgraduate", {})
            if isinstance(pg, dict):
                context += (
                    f"Education: {pg.get('degree', '')} — "
                    f"{pg.get('university', '')} "
                    f"({pg.get('year', '')})\n"
                )
            ug = edu.get("undergraduate", {})
            if isinstance(ug, dict):
                context += (
                    f"Undergraduate: {ug.get('degree', '')} — "
                    f"{ug.get('university', '')} "
                    f"({ug.get('grade', '')})\n"
                )
            certs = edu.get("certifications", [])
            if certs:
                context += (
                    f"Certifications: {', '.join(certs)}\n"
                )

        tech = skills.get("technical_skills", {})
        if tech:
            context += "\nTECHNICAL SKILLS:\n"
            for category, items in tech.items():
                if items:
                    context += (
                        f"  {category.replace('_',' ').title()}: "
                        f"{', '.join(items)}\n"
                    )

        soft = skills.get("soft_skills", [])
        if soft:
            context += f"\nSoft Skills: {', '.join(soft)}\n"

        metrics = skills.get("key_achievements", {})
        if metrics:
            context += "\nKEY ACHIEVEMENTS (use these numbers):\n"
            for k, v in metrics.items():
                context += (
                    f"  - {k.replace('_',' ').title()}: {v}\n"
                )

        projects = skills.get("github_projects", [])
        if projects:
            context += "\nGITHUB PROJECTS:\n"
            for p in projects:
                if isinstance(p, dict):
                    context += (
                        f"  - {p.get('name', '')}: "
                        f"{', '.join(p.get('skills', []))}\n"
                        f"    URL: {p.get('url', '')}\n"
                    )

        similar = skills.get("similar_skills", {})
        if similar:
            context += "\nSIMILAR/EQUIVALENT SKILLS:\n"
            context += json.dumps(similar, indent=2) + "\n"

        visa = skills.get("visa", {})
        if visa:
            context += (
                f"\nVISA STATUS: {visa.get('status', '')}\n"
                f"Right to work UK: "
                f"{visa.get('right_to_work_uk', False)}\n"
                f"Employer sponsorship required: "
                f"{visa.get('employer_sponsorship_required', False)}\n"
            )

        targets = skills.get("role_targets", [])
        if targets:
            context += (
                f"\nTarget Roles: {', '.join(targets)}\n"
            )

        context += "\n\n"

    # ── CVs (trimmed to save tokens) ─────────
    if profile.get("cvs"):
        context += f"CV CONTENT:\n{profile['cvs'][:3000]}\n\n"

    # ── LinkedIn ──────────────────────────────
    if profile.get("linkedin"):
        context += (
            f"LINKEDIN:\n{profile['linkedin'][:1500]}\n\n"
        )

    # ── GitHub (smart cache — once per day) ───
    if github_username:
        github_data = fetch_github_live(github_username)
        if github_data:
            context += f"GITHUB:\n{github_data[:2500]}\n\n"

    # ── Portfolio ─────────────────────────────
    if profile.get("portfolio"):
        context += (
            f"PORTFOLIO:\n{profile['portfolio'][:1000]}\n\n"
        )

    # ── Writing style from chats ──────────────
    if profile.get("chats"):
        context += (
            f"WRITING STYLE (from past chats):\n"
            f"{profile['chats'][:1500]}\n\n"
        )

    return context

# ─────────────────────────────────────────
# JOB SCRAPER
# ─────────────────────────────────────────

def scrape_job(url):
    print(f"\n🔍 Scraping job from:\n   {url}\n")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(
                url,
                wait_until="networkidle",
                timeout=30000
            )
            time.sleep(2)
            html = page.content()
            browser.close()

        soup = BeautifulSoup(html, "html.parser")
        for tag in soup([
            "script", "style", "nav", "header", "footer"
        ]):
            tag.decompose()

        text = soup.get_text(separator="\n", strip=True)
        title_tag = soup.find("title")
        title_text = (
            title_tag.get_text()
            if title_tag else "Unknown Role"
        )

        print(f"  ✅ Scraped: {title_text[:80]}")
        return {
            "url": url,
            "title": title_text,
            "full_text": text[:6000]
        }
    except Exception as e:
        print(f"  ❌ Scrape failed: {e}")
        return None

# ─────────────────────────────────────────
# MATCH ANALYSER
# ─────────────────────────────────────────

def analyse_match(job, profile, skills, github_username):
    print("\n⏳ Analysing your fit for this role...\n")

    context = build_context(profile, skills, github_username)

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1200,
        messages=[{
            "role": "user",
            "content": f"""You are an honest experienced career advisor.

IMPORTANT RULES:
1. Similar or equivalent skills COUNT as matches
   (e.g. PowerBI counts for Tableau, Azure for AWS)
2. Clearly separate exact matches from similar matches
3. Only list a GAP if candidate has nothing similar
4. Be honest — do not oversell or undersell
5. Consider transferable domain experience
6. Note visa and right to work status if relevant

{context}

JOB POSTING:
{job['full_text']}

Reply in this EXACT format — keep labels exactly as shown:

ROLE: (job title)
COMPANY: (company name)
SCORE: (number 0-100)
RECOMMENDATION: (APPLY or SKIP)

EXACT MATCHES:
- (requirement the candidate meets exactly)

SIMILAR MATCHES:
- (job wants X → candidate has Y — why it counts)

GAPS:
- (genuine gap with nothing transferable)

VERDICT: (2-3 honest sentences explaining the score)
"""
        }]
    )
    return response.content[0].text

def parse_field(text, field):
    for line in text.split("\n"):
        if line.strip().startswith(f"{field}:"):
            return line.split(":", 1)[-1].strip()
    return "Unknown"

def parse_score(text):
    for line in text.split("\n"):
        if line.strip().startswith("SCORE:"):
            try:
                return int(
                    line.split(":", 1)[-1].strip().split()[0]
                )
            except:
                return 50
    return 50

# ─────────────────────────────────────────
# CV TAILOR — LATEX + COVER LETTER
# ─────────────────────────────────────────

def tailor_cv(job, profile, skills, github_username):
    """
    Generates a full LaTeX CV ready for Overleaf
    plus a plain text cover letter.
    Includes a review and edit loop before saving.
    """
    context = build_context(profile, skills, github_username)
    feedback_history = ""

    while True:
        print(
            "\n⏳ Generating tailored LaTeX CV "
            "+ cover letter...\n"
        )

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4000,
            messages=[{
                "role": "user",
                "content": f"""You are an expert career coach 
and LaTeX document specialist.

Using everything you know about this candidate, create 
tailored application documents for this specific job.

{context}
{feedback_history}

JOB POSTING:
{job['full_text']}

Create TWO documents:

1. A FULL LaTeX CV ready to paste into Overleaf:
   - Use moderncv package, classic style, blue colour
   - Sections in this order:
     * Personal info (name, email, phone, LinkedIn, GitHub)
     * Professional Summary (3-4 lines for THIS role)
     * Work Experience (relevant first, real metrics as bullets)
     * Education (with grades)
     * Technical Skills (grouped, most relevant first)
     * Key Projects (from GitHub, most relevant to role)
     * Certifications
   - Use real achievement numbers from the profile
   - Use similar skills where exact ones are missing
   - Must compile in Overleaf without errors
   - Escape all special LaTeX characters properly
     (& % $ # _ {{ }} ~ ^ \\)

2. A plain text cover letter:
   - Match candidate's natural tone from chat history
   - Address the specific role and company
   - Mention relevant projects and numbers
   - 3-4 paragraphs, professional but personal

Format your response EXACTLY like this:

--- LATEX CV ---
(complete LaTeX starting with \\documentclass)

--- COVER LETTER ---
(plain text cover letter)
"""
            }]
        )

        result = response.content[0].text

        # Show LaTeX preview (first 50 lines)
        print("\n" + "="*55)
        print("📄 GENERATED DOCUMENTS PREVIEW")
        print("="*55)

        if "--- LATEX CV ---" in result:
            latex_part = result.split("--- LATEX CV ---")[1]
            if "--- COVER LETTER ---" in latex_part:
                latex_part = latex_part.split(
                    "--- COVER LETTER ---"
                )[0]
            lines = latex_part.strip().split("\n")
            print("\n📄 LATEX CV (first 50 lines):")
            print("\n".join(lines[:50]))
            if len(lines) > 50:
                print(
                    f"\n  ... ({len(lines) - 50} more lines)"
                )

        if "--- COVER LETTER ---" in result:
            cl = result.split(
                "--- COVER LETTER ---"
            )[-1].strip()
            print("\n📝 COVER LETTER:")
            print(cl)

        print("="*55)
        print("\n✏️  Options:")
        print("  ok       → save and proceed to form")
        print("  change   → request a specific change")
        print("  again    → regenerate from scratch")
        print("  show     → print full LaTeX code")
        print("  skip     → abandon this application")

        choice = input("\nYour choice: ").strip().lower()

        if choice == "ok":
            return result
        elif choice == "skip":
            return None
        elif choice == "again":
            feedback_history = ""
            print("\n🔄 Regenerating from scratch...")
        elif choice == "show":
            print("\n" + "="*55)
            print("FULL LaTeX CODE:")
            print("="*55)
            print(result)
            print("="*55)
        elif choice == "change":
            feedback = input(
                "What should be changed? "
            ).strip()
            feedback_history += (
                f"\nREVISION: {feedback}\n"
                f"Please incorporate this change.\n"
            )
        else:
            print("Type ok / change / again / show / skip")

# ─────────────────────────────────────────
# DOCUMENT SAVER
# ─────────────────────────────────────────

def save_documents(tailored, url, role, company, score):
    """
    Saves three files per application:
    - cv.tex        (paste into Overleaf)
    - cover_letter.txt
    - full_output.txt
    """
    safe_company = (
        company.replace(" ", "_")
               .replace("/", "_")
               .replace("\\", "_")[:30]
    )
    app_num = len(list(PAST_APPS_DIR.iterdir())) + 1
    app_folder = (
        PAST_APPS_DIR / f"app_{app_num}_{safe_company}"
    )
    app_folder.mkdir(exist_ok=True)

    # Full raw output
    (app_folder / "full_output.txt").write_text(
        f"URL: {url}\n"
        f"Role: {role}\n"
        f"Company: {company}\n"
        f"Score: {score}/100\n"
        f"Date: {time.strftime('%Y-%m-%d %H:%M')}\n"
        f"\n{'='*50}\n\n{tailored}",
        encoding="utf-8"
    )

    # LaTeX CV
    if "--- LATEX CV ---" in tailored:
        latex_part = tailored.split("--- LATEX CV ---")[1]
        if "--- COVER LETTER ---" in latex_part:
            latex_part = latex_part.split(
                "--- COVER LETTER ---"
            )[0]
        latex_path = app_folder / "cv.tex"
        latex_path.write_text(
            latex_part.strip(), encoding="utf-8"
        )
        print(f"\n📄 LaTeX CV saved:     {latex_path}")
        print(
            "   → Open Overleaf → New Project → Blank\n"
            "   → Delete default text → Paste cv.tex\n"
            "   → Click Recompile → Download PDF"
        )

    # Cover letter
    if "--- COVER LETTER ---" in tailored:
        cl = tailored.split("--- COVER LETTER ---")[-1].strip()
        cl_path = app_folder / "cover_letter.txt"
        cl_path.write_text(cl, encoding="utf-8")
        print(f"📝 Cover letter saved: {cl_path}")

    print(f"📁 All files in:       {app_folder}")
    return app_folder

# ─────────────────────────────────────────
# FORM FILLER
# ─────────────────────────────────────────

def fill_form(url, tailored_docs, profile, skills):
    print("\n🤖 Opening application form...")
    print("   Agent fills fields automatically.")
    print("   Stops before submitting for your review.\n")

    # Get contact from skills file
    contact = skills.get("personal", {})
    if not contact.get("email"):
        cv_text = profile.get("cvs", "")
        try:
            info_resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=200,
                messages=[{
                    "role": "user",
                    "content": (
                        f"Extract from this CV:\n"
                        f"- Full name\n"
                        f"- Email\n"
                        f"- Phone\n"
                        f"- City\n\n"
                        f"CV: {cv_text[:1000]}\n\n"
                        f"Reply JSON only, no other text:\n"
                        f'{{"name":"","email":"",'
                        f'"phone":"","location":""}}'
                    )
                }]
            )
            contact = json.loads(
                info_resp.content[0].text.strip()
            )
        except:
            contact = {}

    # Save latest CV text
    cv_latest = PROCESSED_DIR / "tailored_cv_latest.txt"
    cv_latest.write_text(tailored_docs, encoding="utf-8")

    # Extract cover letter
    cover_letter = ""
    if "--- COVER LETTER ---" in tailored_docs:
        cover_letter = tailored_docs.split(
            "--- COVER LETTER ---"
        )[-1].strip()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto(
            url, wait_until="networkidle", timeout=30000
        )
        time.sleep(3)

        filled = []

        def try_fill(selector, value, label):
            try:
                els = page.locator(selector)
                if els.count() > 0:
                    els.first.fill(str(value))
                    filled.append(label)
            except:
                pass

        # Name
        if contact.get("name"):
            try_fill('input[name*="name" i]',
                     contact["name"], "name")
            try_fill('input[placeholder*="name" i]',
                     contact["name"], "name")
            try_fill('input[autocomplete="name"]',
                     contact["name"], "name")

        # Email
        if contact.get("email"):
            try_fill('input[type="email"]',
                     contact["email"], "email")
            try_fill('input[name*="email" i]',
                     contact["email"], "email")

        # Phone
        if contact.get("phone"):
            try_fill('input[type="tel"]',
                     contact["phone"], "phone")
            try_fill('input[name*="phone" i]',
                     contact["phone"], "phone")

        # Location
        if contact.get("location"):
            try_fill('input[name*="location" i]',
                     contact["location"], "location")
            try_fill('input[name*="city" i]',
                     contact["location"], "city")
            try_fill('input[name*="address" i]',
                     contact["location"], "address")

        # Cover letter
        if cover_letter:
            try_fill('textarea[name*="cover" i]',
                     cover_letter, "cover letter")
            try_fill('textarea[name*="letter" i]',
                     cover_letter, "cover letter")
            try_fill('textarea[placeholder*="cover" i]',
                     cover_letter, "cover letter")
            try_fill('textarea[placeholder*="motivation" i]',
                     cover_letter, "motivation")

        print(
            f"\n  ✅ Auto-filled: "
            f"{', '.join(filled) if filled else 'basic fields'}"
        )
        print("\n" + "="*55)
        print("👀 BROWSER IS OPEN — REVIEW THE FORM")
        print("="*55)
        print("• Check all fields are correct")
        print("• Fill anything the agent missed")
        print("• Upload CV PDF if there is a file upload field")
        print(f"  (CV text saved at: {cv_latest})")
        print()

        decision = input(
            "Type SUBMIT to submit / CANCEL to abort: "
        ).strip().upper()

        result = "cancelled"
        if decision == "SUBMIT":
            try:
                btn = page.locator(
                    'button[type="submit"]'
                ).first
                if btn.count() > 0:
                    btn.click()
                    time.sleep(3)
                    print("\n✅ Submitted!")
                    result = "applied"
                else:
                    print("\n⚠️ Submit button not found")
                    print("Please click Submit manually.")
                    input("Press Enter when submitted...")
                    result = "applied"
            except Exception as e:
                print(f"\n⚠️ Auto-click failed: {e}")
                print("Please click Submit manually.")
                input("Press Enter when submitted...")
                result = "applied"
        else:
            print("\n❌ Cancelled.")

        input("\nPress Enter to close browser...")
        browser.close()
        return result

# ─────────────────────────────────────────
# MAIN APPLY FLOW
# ─────────────────────────────────────────

def apply_for_job():
    profile = load_profile()
    if not profile:
        return

    skills = load_skills()

    github_username = None
    links_path = DATA_DIR / "links.json"
    if links_path.exists():
        links = json.loads(links_path.read_text())
        github_username = links.get("github_username")

    url = input("\n🔗 Paste the job URL: ").strip()
    if not url.startswith("http"):
        print("❌ Please paste a full URL starting with http")
        return

    # ── 1. Scrape ────────────────────────────
    job = scrape_job(url)
    if not job:
        print("❌ Could not scrape. Try a direct job page URL.")
        return

    # ── 2. Analyse match ─────────────────────
    analysis = analyse_match(
        job, profile, skills, github_username
    )
    score = parse_score(analysis)
    role = parse_field(analysis, "ROLE")
    company = parse_field(analysis, "COMPANY")
    recommendation = parse_field(analysis, "RECOMMENDATION")

    # ── 3. Show match report ─────────────────
    print("\n" + "━"*55)
    print("🎯 JOB MATCH ANALYSIS")
    print("━"*55)
    print(analysis)
    print("━"*55)

    if recommendation == "APPLY":
        print(
            f"\n✅ Score: {score}/100 — "
            f"Claude recommends APPLY"
        )
    else:
        print(
            f"\n⚠️  Score: {score}/100 — "
            f"Claude recommends SKIP"
        )
        print("   Final call is always yours.\n")

    print("\nOptions:")
    print("  yes    → proceed with application")
    print("  no     → skip this job")
    print("  why    → ask Claude about this role")

    # ── 4. Decision loop ─────────────────────
    while True:
        decision = input("\nFinal call → ").strip().lower()

        if decision == "yes":
            if recommendation == "SKIP":
                print(
                    "\n💪 Overriding skip — proceeding.\n"
                    "   Claude will address gaps in your CV.\n"
                )
            break

        elif decision == "no":
            print("\n" + "━"*55)
            reason = input(
                "Quick reason (Enter to skip): "
            ).strip()
            log_application(
                url, company, role, score,
                f"skipped — "
                f"{reason if reason else 'user choice'}"
            )
            print(f"📝 Logged: {role} @ {company} — skipped")
            return

        elif decision == "why":
            q = input("What do you want to know? ").strip()
            resp = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=400,
                messages=[{
                    "role": "user",
                    "content": (
                        f"Based on this job analysis:\n"
                        f"{analysis}\n\n"
                        f"Answer honestly in 3-5 sentences:\n"
                        f"{q}"
                    )
                }]
            )
            print("\n💬 " + resp.content[0].text)
            print("\n(Type yes / no when ready)")

        else:
            print("Please type yes / no / why")

    # ── 5. Generate LaTeX CV + cover letter ──
    tailored = tailor_cv(
        job, profile, skills, github_username
    )
    if not tailored:
        log_application(url, company, role, score, "abandoned")
        return

    # ── 6. Save all documents ────────────────
    app_folder = save_documents(
        tailored, url, role, company, score
    )

    # ── 7. Form filler (optional) ────────────
    print("\n" + "━"*55)
    fill_choice = input(
        "Open form filler now? (yes/no): "
    ).strip().lower()

    if fill_choice == "yes":
        fill_result = fill_form(
            url, tailored, profile, skills
        )
    else:
        fill_result = "saved — form not filled yet"
        print(f"\n📁 Documents saved in:\n   {app_folder}")
        print(
            "   Open cv.tex in Overleaf when ready."
        )

    # ── 8. Log result ────────────────────────
    log_application(url, company, role, score, fill_result)
    print(f"\n✅ Done! Logged as: {fill_result}")

# ─────────────────────────────────────────
# APPLICATION HISTORY
# ─────────────────────────────────────────

def show_log():
    log = load_log()
    if not log:
        print("\n📭 No applications logged yet.")
        return

    applied = [x for x in log if x["status"] == "applied"]
    skipped = [
        x for x in log
        if "skipped" in x.get("status", "")
    ]
    abandoned = [
        x for x in log
        if x.get("status") == "abandoned"
    ]

    print("\n📊 APPLICATION HISTORY")
    print(
        f"   Total: {len(log)} | "
        f"Applied: {len(applied)} | "
        f"Skipped: {len(skipped)} | "
        f"Abandoned: {len(abandoned)}"
    )
    print("─" * 65)

    for entry in log[-15:]:
        if entry["status"] == "applied":
            icon = "✅"
        elif "skipped" in entry.get("status", ""):
            icon = "⏭️ "
        else:
            icon = "❌"
        print(
            f"{icon} {entry['date']} | "
            f"{entry['role'][:25]:<25} | "
            f"{entry['company'][:18]:<18} | "
            f"Score: {entry['match_score']:>3} | "
            f"{entry['status']}"
        )

# ─────────────────────────────────────────
# MAIN MENU
# ─────────────────────────────────────────

def main():
    while True:
        print("\n" + "━"*45)
        print("🤖  CAREER AGENT — Powered by Claude")
        print("━"*45)
        print("1.  Build / Rebuild my profile")
        print("2.  Add / update profile links")
        print("3.  Apply for a job (paste URL)")
        print("4.  View application history")
        print("5.  Refresh GitHub cache now")
        print("6.  Exit")
        print("━"*45)

        choice = input("Choose option (1-6): ").strip()

        if choice == "1":
            build_profile()

        elif choice == "2":
            print("\n🔗 Update your links:\n")
            links = {
                "linkedin": input(
                    "LinkedIn URL: "
                ).strip(),
                "github_username": input(
                    "GitHub username (e.g. anjali20): "
                ).strip(),
                "portfolio": input(
                    "Portfolio URL: "
                ).strip()
            }
            with open(DATA_DIR / "links.json", "w") as f:
                json.dump(links, f, indent=2)
            print("✅ Links saved!")

        elif choice == "3":
            apply_for_job()

        elif choice == "4":
            show_log()

        elif choice == "5":
            links_path = DATA_DIR / "links.json"
            if links_path.exists():
                links = json.loads(links_path.read_text())
                username = links.get("github_username")
                if username:
                    fetch_github_live(
                        username, force_refresh=True
                    )
                    print("✅ GitHub cache refreshed!")
                else:
                    print(
                        "❌ No GitHub username found.\n"
                        "   Run option 2 to add it."
                    )
            else:
                print("❌ No links.json — run option 2 first")

        elif choice == "6":
            print("\nGoodbye! 👋")
            break

        else:
            print("Please choose 1-6")

        if choice != "6":
            input("\nPress Enter to return to menu...")

if __name__ == "__main__":
    main()