# 📖 Novel Writing Studio

AI-powered 6-stage novel writing workflow built with Streamlit + Claude.

## Stages
1. **Outline** — Generate chapter-by-chapter structure from your idea
2. **Draft** — Expand chapters into full prose (one-by-one or all at once)
3. **Logic check** — Flag plot holes, timeline errors, inconsistencies
4. **Format** — Review pacing and apply manuscript formatting standards
5. **Polish** — Humanize AI prose, remove clichés, sharpen voice
6. **Export** — Download as `.txt` or formatted `.docx`

---

## Run locally

### 1. Clone / download the project
```bash
git clone https://github.com/YOUR_USERNAME/novel-studio.git
cd novel-studio
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Add your API key
Edit `.streamlit/secrets.toml`:
```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
```
Get your key at [console.anthropic.com](https://console.anthropic.com)

### 4. Run the app
```bash
streamlit run app.py
```
Open http://localhost:8501 in your browser.

---

## Deploy to Streamlit Cloud (free)

1. Push this folder to a GitHub repo (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in
3. Click **New app** → select your repo → set main file to `app.py`
4. Go to **Settings → Secrets** and add:
   ```
   ANTHROPIC_API_KEY = "sk-ant-your-key-here"
   ```
5. Click **Deploy** — your app is live in ~2 minutes!

---

## Project structure
```
novel-studio/
├── app.py                  # Main Streamlit app
├── requirements.txt        # Python dependencies
├── README.md
└── .streamlit/
    ├── config.toml         # Theme & server config
    └── secrets.toml        # API key (local only, never commit this!)
```

> ⚠️ Add `.streamlit/secrets.toml` to your `.gitignore` before pushing to GitHub.
