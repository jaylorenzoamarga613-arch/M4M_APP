M4M: SIDLAKANG BALINGOAN — Two-Role Registration App
A production-grade event registration system with a public registration page
and a private admin dashboard — just like veent.net.
---
Two URLs, Two Experiences
URL	Who sees it	What they get
`https://yourapp.streamlit.app`	Everyone	Hero, registration form, public announcements
`https://yourapp.streamlit.app/?admin=true`	Organizer only	Full dashboard (password protected)
---
Admin Dashboard Features
📊 Live stats: total registered, revenue, today's count, 7K/15K/21K split
📈 Charts: registrations by category, timeline, shirt size distribution
👥 Registrants table with search + filter by category/shirt
🧾 Payment proof viewer — see each runner's screenshot
📣 Announcements — post updates that appear on the public page
🖼 Materials manager — upload shirts, medal, QR codes, banner, route map
⬇ Export: Excel (3 sheets), CSV (Google Sheets), JSON backup
Public Page Features
Beautiful Sidlakang-branded hero
4-step registration form (Info → Payment → Waiver → Confirm)
GCash QR codes shown automatically once uploaded by admin
Payment proof upload
Full waiver with digital signature
Updates tab showing admin announcements
Pinned announcements shown at top of page
---
Deploy to Streamlit Cloud (Free, ~5 minutes)
Step 1 — Push to GitHub
Create a free account at https://github.com
New repository → name it `sidlakang-app`
Upload `app.py` and `requirements.txt`
Step 2 — Deploy
Go to https://share.streamlit.io → sign in with GitHub
Click New app → select your repo → Main file: `app.py`
Click Deploy
Your app goes live at `https://sidlakang-app.streamlit.app`
Step 3 — Access admin
Go to: `https://sidlakang-app.streamlit.app/?admin=true`
Default password: `bakyard2026` (change it in app.py line 14!)
---
Change the Admin Password
Open `app.py`, line 14:
```python
ADMIN\_PASSWORD = "bakyard2026"   # ← change this to something secret
```
Change Event Details
Lines 15–22 in `app.py`:
```python
EVENT\_NAME      = "M4M: SIDLAKANG BALINGOAN"
EVENT\_DATE      = "May 24, 2026"
EVENT\_LOCATION  = "Brgy. Dalihig, Balingoan, Misamis Oriental"
EVENT\_ORGANIZER = "BAKYARD Events"
GCASH\_NAME      = "Jorgil Amarga"
GCASH\_NUMBER    = "0916 481 3822"
```
---
Data Persistence
Registrations and announcements are saved to `registrations.json`
Uploaded images (QR codes, proofs, shirts) are saved in the `assets/` folder
On Streamlit Cloud, files reset on redeploy — use the JSON export to back up regularly
For permanent storage, upgrade to a database (Supabase free tier works great)
---
Files
`app.py` — full application (1006 lines)
`requirements.txt` — Python dependencies
