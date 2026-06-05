# Canva Automation — Streamlit App

Deploy in ~5 minutes to [Streamlit Cloud](https://streamlit.io/cloud) (free tier).

## Files

```
streamlit_app/
├── app.py                          ← main app
├── requirements.txt
└── .streamlit/
    └── secrets.toml.example        ← copy → secrets.toml and fill in
```

---

## Deploy to Streamlit Cloud

### 1 · Push to GitHub

```bash
git init
git add .
git commit -m "initial"
gh repo create canva-automation --public --push
```

### 2 · Connect Streamlit Cloud

1. Go to **https://share.streamlit.io** → **New app**
2. Connect your GitHub repo
3. Set **Main file path** → `app.py`
4. Click **Deploy**

The app works immediately without Google credentials — Step 1 outputs a downloadable `.txt` file and Step 2 accepts that upload.

---

## Optional: Enable Google Docs push (~10 min, free)

Without this, the app still works fully — Step 1 downloads an outline `.txt`, Step 2 uploads it to generate the CSV.

**If you want the "Push to Google Doc" button to work:**

### A · Create a Google Cloud service account

1. Go to **https://console.cloud.google.com** → New project (name it anything)
2. **APIs & Services → Library** → search and enable:
   - **Google Docs API**
   - **Google Drive API**
3. **APIs & Services → Credentials → + Create Credentials → Service account**
   - Name: `canva-automation` → Create and continue → Done
4. Click the service account row → **Keys** tab → **Add Key → Create new key → JSON** → Download

### B · Add to Streamlit secrets

1. Open the downloaded JSON file
2. In Streamlit Cloud, go to your app → **⋮ menu → Settings → Secrets**
3. Paste this (replacing values from your JSON):

```toml
[gcp_service_account]
type                        = "service_account"
project_id                  = "your-project-id"
private_key_id              = "abc123"
private_key                 = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email                = "canva-automation@your-project.iam.gserviceaccount.com"
client_id                   = "123456789"
auth_uri                    = "https://accounts.google.com/o/oauth2/auth"
token_uri                   = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url        = "https://www.googleapis.com/robot/v1/metadata/x509/canva-automation%40your-project.iam.gserviceaccount.com"
```

> ⚠️ Keep secrets private. Never commit `secrets.toml` to git — it's already in `.gitignore`.

4. Click **Save** → app restarts → Google Docs push is now enabled.

---

## How the app works

### Tab 1 — PPTX → Outline

| Action | What happens |
|--------|-------------|
| Upload `.pptx` | Extracts text from all slides |
| Preview | Shows colour-coded outline hierarchy |
| Download `.txt` | Plain-text outline for Tab 2 upload |
| Push to Google Doc | Creates a formatted Google Doc, shares with your email |

### Outline hierarchy
```
I.   Slide title              ← Roman numeral
    A. Main point             ← Capital letter
        1. Sub-point          ← Number
            a) Detail         ← Lowercase + )
                (1) Fine note ← Digit in ()
```

### Tab 2 — Outline → Canva CSV

| Source option | When to use |
|--------------|-------------|
| From Step 1 (this session) | Ran Step 1 in the same browser session |
| Upload outline .txt | Downloaded the file from Step 1 |
| Google Doc ID | Pushed to Google Docs and want to re-process |

Generates `canva_bulk_create.csv` with columns: `slide_title`, `point_1`, `point_2`, `subpoint_1`, …

### Canva Bulk Create (free, built-in)
1. Open your Canva template
2. Click a textbox → **Connect data** → type the field name (e.g. `slide_title`)
3. Repeat for each textbox
4. **Apps → Bulk Create → Upload CSV** → select the downloaded file
5. Canva generates one page per row ✅

---

## Run locally

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# edit secrets.toml with your credentials (or leave blank for no-Google mode)
streamlit run app.py
```
