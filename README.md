# tmOsphere — Trademark Document Generator

Fill one form → download two Word documents instantly:
- **Affidavit** (sworn statement for trademark registration)
- **General Power of Attorney** (under Section 145, Trade Marks Act 1999)

---

## Project Structure

```
tmosphere/
├── server.js          ← Express backend (API + serves frontend)
├── package.json
├── render.yaml        ← Render.com deployment config
├── .gitignore
└── public/
    └── index.html     ← Frontend (no build step needed)
```

---

## Step 1 — Push to GitHub

```bash
# 1. Create a new repo on github.com named "tmosphere" (keep it empty)

# 2. On your computer, open a terminal and run:
git init tmosphere
cd tmosphere

# 3. Copy all project files into this folder, then:
git add .
git commit -m "tmOsphere v1 — Trademark document generator"

# 4. Connect and push
git remote add origin https://github.com/YOUR_USERNAME/tmosphere.git
git branch -M main
git push -u origin main
```

---

## Step 2 — Deploy on Render (free)

1. Go to **https://render.com** → Sign up / Log in with GitHub
2. Click **"New +"** → **"Web Service"**
3. Select your **tmosphere** GitHub repository
4. Render will auto-detect `render.yaml` — settings will be filled in automatically:
   - **Name:** tmosphere
   - **Runtime:** Node
   - **Build Command:** `npm install`
   - **Start Command:** `node server.js`
   - **Plan:** Free
5. Click **"Create Web Service"**
6. Wait ~2 minutes for the first deploy
7. Your live URL will be: `https://tmosphere.onrender.com` (or similar)

That's it! Every time you push to GitHub, Render auto-deploys.

---

## Run Locally

```bash
npm install
npm start
# Open http://localhost:3001
```

---

## API

### POST /api/generate

**Request JSON:**

| Field | Description |
|---|---|
| `applicantName` | Full name of trademark applicant |
| `brandName` | Trademark / brand name |
| `registeredAddress` | Registered office address |
| `residentialAddress` | Residential address |
| `businessClass` | TM class number (1–45) |
| `businessType` | e.g. BUSINESS / SERVICES |
| `affidavitDate` | Date (YYYY-MM-DD) |
| `place` | Place of signing |
| `agentName` | Attorney full name |
| `agentCode` | Trade Marks Agent Code |
| `agentAddress` | Attorney office address |
| `poaDate` | POA signing date (YYYY-MM-DD) |
| `poaExecutionDate` | e.g. "3rd day of September, 2025" |
| `tmOffice` | TM Registry office(s) |

**Response:** `application/zip` with `Affidavit_NAME.docx` + `POA_NAME.docx`

### GET /health

Returns `{ status: "ok", app: "tmOsphere" }`

---

## Tech Stack

| | |
|---|---|
| Runtime | Node.js 18+ |
| Server | Express |
| Documents | `docx` npm package |
| Packaging | `jszip` |
| Frontend | Vanilla HTML/CSS/JS |
| Hosting | Render.com (free tier) |

---

*For professional use. Always review documents with a qualified trademark attorney before submission.*
