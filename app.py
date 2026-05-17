import streamlit as st
import datetime, io, uuid, base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

# ── PAGE CONFIG ────────────────────────────────────────────────
st.set_page_config(
    page_title="M4M: SIDLAKANG BALINGOAN",
    page_icon="🏔️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CONSTANTS ──────────────────────────────────────────────────
CATEGORIES  = {"7KM": 800, "15KM": 1200, "21KM": 1400}
SHIRT_SIZES = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
GENDERS     = ["Male", "Female", "Other / Prefer not to say"]
WAIVERS = [
    ("Assumption of Risk",
     "I acknowledge that my participation in M4M: BALINGOAN 'SIDLAKANG' involves inherent risks, "
     "hazards, and dangers including slips, falls, uneven terrain, physical exertion, weather "
     "conditions, and encounters with vehicles or animals. Injuries may range from minor to severe "
     "or fatal. I voluntarily assume full responsibility for all risks.",
     "I have read, understand, and agree to assume all risks."),
    ("Fitness to Participate",
     "I represent that I am in good physical condition and have no medical history that would "
     "prevent me from safely participating. I agree it is my sole responsibility to determine "
     "whether I am sufficiently fit.",
     "I confirm I am medically and physically fit to participate."),
    ("Waiver and Release of Liability",
     "I, on behalf of myself, my heirs, and assigns, hereby fully release, waive, and discharge "
     "the event organizers, directors, employees, volunteers, and sponsors from any and all claims "
     "arising out of any loss, damage, or injury sustained while participating.",
     "I agree to the waiver and release of liability."),
    ("Indemnification",
     "I agree to indemnify, defend, and hold harmless the event organizers from any costs, damages, "
     "lawsuits, or liabilities arising from my bodily injury, death, property damage, or other loss "
     "from my participation.",
     "I agree to the indemnification terms."),
    ("Medical Consent",
     "In the event of injury or medical emergency, I authorize the event organizers and volunteers "
     "to secure emergency medical care or transportation. I agree to assume all costs associated "
     "with such treatment.",
     "I consent to emergency medical treatment and assume all associated costs."),
    ("Media Release",
     "I grant the event organizers the irrevocable right to photograph, record, and use my name, "
     "image, voice, and likeness in any media for marketing and promotional purposes, without "
     "compensation.",
     "I grant permission for the media release."),
    ("Digital Signature",
     "By checking this box, I agree that this digital submission acts as my legally binding "
     "signature for the M4M: SIDLAKANG | BALINGOAN event.",
     "I agree and submit my digital signature."),
]

# ── SESSION STATE INIT ─────────────────────────────────────────
def init_state():
    defaults = {
        "registrations": [],
        "step": 1,
        "form": {},
        "assets": {
            "shirt_7km":  {"img": None, "label": "7KM Finisher Shirt"},
            "shirt_15km": {"img": None, "label": "15KM Finisher Shirt"},
            "shirt_21km": {"img": None, "label": "21KM Finisher Shirt"},
            "medal":      {"img": None, "label": "Finisher Medal"},
            "qr_gcash":   {"img": None, "label": "GCash QR Code"},
            "qr_maya":    {"img": None, "label": "Maya QR Code"},
            "banner":     {"img": None, "label": "Event Banner"},
            "route_map":  {"img": None, "label": "Route Map"},
        },
        "proof_bytes": None,
        "proof_name":  None,
        "success_id":  None,
        "active_tab":  "register",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ── GLOBAL CSS ─────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@700;800;900&family=Barlow:wght@400;500;600&display=swap');

/* Hide Streamlit chrome */
#MainMenu, footer, header {visibility: hidden;}
.block-container {padding-top: 0 !important; padding-bottom: 2rem;}
[data-testid="stSidebar"] {display: none;}

/* Root vars */
:root {
  --gold: #c9a84c; --gold2: #e8c96a; --gold3: #8a6820;
  --dark: #0a0f0d; --dark2: #111a14; --dark3: #1c2a20;
  --green2: #2a5c3a;
  --cream: #f5f0e8; --text: #e8e2d4; --muted: #8a9a8a;
  --red: #e74c3c; --border: rgba(201,168,76,.22);
}

/* Hero */
.hero {
  background: linear-gradient(135deg,#050a07,#0f1f12,#0c1a10);
  padding: 36px 40px 28px; margin: -1rem -1rem 0;
  border-bottom: 2px solid var(--gold3);
  font-family: 'Barlow Condensed', sans-serif;
}
.hero-eye {font-size:11px;font-weight:700;letter-spacing:5px;text-transform:uppercase;color:var(--gold);opacity:.8;margin-bottom:6px;}
.hero-title {font-size:clamp(48px,8vw,80px);font-weight:900;line-height:.9;text-transform:uppercase;letter-spacing:-2px;
  background:linear-gradient(180deg,#e8d88a,#c9a84c,#8a6820);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;margin-bottom:4px;}
.hero-sub {font-size:20px;font-weight:700;letter-spacing:6px;text-transform:uppercase;color:var(--muted);margin-bottom:16px;}
.hero-pill {display:inline-block;font-size:11px;font-weight:700;letter-spacing:2px;text-transform:uppercase;
  border:1px solid var(--border);border-radius:2px;padding:5px 14px;color:var(--muted);margin-right:8px;margin-bottom:6px;}
.hstat-n {font-size:36px;font-weight:900;color:var(--gold);line-height:1;}
.hstat-l {font-size:10px;letter-spacing:2px;text-transform:uppercase;color:var(--muted);}

/* Step bar */
.stepbar {display:flex;gap:0;margin-bottom:32px;position:relative;}
.stepbar::before {content:'';position:absolute;top:16px;left:0;right:0;height:1px;background:var(--border);z-index:0;}
.sitem {flex:1;display:flex;flex-direction:column;align-items:center;gap:7px;z-index:1;}
.snum {width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;
  font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;background:#1c2a20;border:1px solid var(--border);color:var(--muted);}
.snum.active {background:var(--gold);border-color:var(--gold);color:#000;}
.snum.done {background:#2a5c3a;border-color:#2a5c3a;color:#fff;}
.slbl {font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:var(--muted);}
.slbl.active {color:var(--gold);}
.slbl.done {color:#2a5c3a;}

/* Cards / boxes */
.pay-box {background:rgba(201,168,76,.07);border:1px solid rgba(201,168,76,.3);border-radius:4px;padding:20px 24px;margin-bottom:20px;}
.pay-title {font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--gold);margin-bottom:12px;}
.fee-grid {display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:14px;}
.fee-card {background:rgba(0,0,0,.2);border:1px solid var(--border);border-radius:3px;padding:12px;text-align:center;}
.fee-dist {font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:900;color:var(--gold);}
.fee-amt  {font-size:15px;font-weight:600;color:var(--text);margin-top:2px;}

/* Waiver box */
.wbox {background:rgba(0,0,0,.2);border:1px solid var(--border);border-radius:3px;padding:16px 20px;margin-bottom:12px;}
.wbox-title {font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--gold);margin-bottom:8px;}
.wbox-text {font-size:12px;color:var(--muted);line-height:1.7;}

/* Summary row */
.srow {display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid rgba(255,255,255,.06);font-size:13px;}
.srow:last-child{border:none;}
.srow span:first-child{color:var(--muted);}
.srow span:last-child{color:var(--text);font-weight:500;text-align:right;}

/* Success */
.success-card {background:rgba(255,255,255,.04);border:1px solid var(--border);border-radius:4px;padding:20px 24px;margin:16px 0;}

/* Stat card */
.scard {background:#1c2a20;border:1px solid var(--border);border-radius:4px;padding:14px;text-align:center;}
.scard-n {font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:900;color:var(--gold);line-height:1;}
.scard-l {font-size:10px;letter-spacing:2px;text-transform:uppercase;color:var(--muted);margin-top:3px;}

/* Override Streamlit elements for dark theme */
.stApp {background:#0a0f0d !important;}
section[data-testid="stMain"] {background:#0a0f0d;}
.stTabs [data-baseweb="tab-list"] {background:#111a14;border-bottom:2px solid #8a6820;gap:0;}
.stTabs [data-baseweb="tab"] {
  font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:700;
  letter-spacing:2px;text-transform:uppercase;color:#8a9a8a;
  border-bottom:3px solid transparent;padding:13px 20px;
}
.stTabs [aria-selected="true"] {color:#c9a84c !important;border-bottom:3px solid #c9a84c !important;background:rgba(201,168,76,.07);}
.stTabs [data-baseweb="tab-panel"] {padding-top:24px;}
div[data-testid="stForm"] {border:none;padding:0;}
.stTextInput input, .stTextArea textarea, .stSelectbox select,
[data-baseweb="input"] input, [data-baseweb="textarea"] textarea {
  background:rgba(255,255,255,.04) !important;
  border:1px solid rgba(201,168,76,.22) !important;
  border-radius:3px !important;
  color:#e8e2d4 !important;
  font-family:'Barlow',sans-serif !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {border-color:rgba(201,168,76,.6) !important;}
label {font-family:'Barlow Condensed',sans-serif !important;font-size:11px !important;
  font-weight:700 !important;letter-spacing:2px !important;text-transform:uppercase !important;
  color:rgba(201,168,76,.85) !important;}
.stSelectbox [data-baseweb="select"] {background:rgba(255,255,255,.04);border:1px solid rgba(201,168,76,.22);}
.stSelectbox [data-baseweb="select"] span {color:#e8e2d4;}
.stRadio label {font-family:'Barlow',sans-serif !important;font-size:14px !important;
  font-weight:400 !important;letter-spacing:0 !important;text-transform:none !important;color:#e8e2d4 !important;}
.stRadio [data-testid="stWidgetLabel"] p {font-family:'Barlow Condensed',sans-serif !important;
  font-size:11px !important;font-weight:700 !important;letter-spacing:2px !important;
  text-transform:uppercase !important;color:rgba(201,168,76,.85) !important;}
.stCheckbox label {font-family:'Barlow',sans-serif !important;font-size:13px !important;
  font-weight:400 !important;letter-spacing:0 !important;text-transform:none !important;color:#e8e2d4 !important;}
.stCheckbox [data-testid="stWidgetLabel"] {color:#e8e2d4 !important;}
.stButton > button {
  font-family:'Barlow Condensed',sans-serif !important;font-size:14px !important;
  font-weight:700 !important;letter-spacing:3px !important;text-transform:uppercase !important;
  background:#c9a84c !important;color:#000 !important;border:none !important;
  border-radius:3px !important;padding:12px 32px !important;
}
.stButton > button:hover {background:#e8c96a !important;}
.stButton.back-btn > button {background:transparent !important;color:#8a9a8a !important;
  border:1px solid rgba(201,168,76,.22) !important;}
div[data-testid="stFileUploader"] {border:1px dashed rgba(201,168,76,.35);border-radius:3px;
  background:rgba(255,255,255,.02);padding:8px;}
div[data-testid="stFileUploader"]:hover {border-color:#c9a84c;}
div[data-testid="stFileUploader"] label {color:#8a9a8a !important;}
.stSuccess {background:rgba(46,200,100,.12) !important;border:1px solid rgba(46,200,100,.3) !important;}
.stError {background:rgba(231,76,60,.12) !important;border:1px solid rgba(231,76,60,.3) !important;}
.stInfo {background:rgba(201,168,76,.08) !important;border:1px solid rgba(201,168,76,.25) !important;}
p, .stMarkdown p {color:#e8e2d4 !important;font-family:'Barlow',sans-serif !important;}
h1,h2,h3 {font-family:'Barlow Condensed',sans-serif !important;text-transform:uppercase !important;letter-spacing:-1px !important;color:#f5f0e8 !important;}
div[data-testid="stDataFrame"] {border:1px solid rgba(201,168,76,.22);border-radius:3px;}
.stDataFrame th {background:#111a14 !important;color:#c9a84c !important;
  font-family:'Barlow Condensed',sans-serif !important;letter-spacing:1px !important;text-transform:uppercase !important;}
.stDataFrame td {color:#e8e2d4 !important;background:#0a0f0d !important;}
.stDownloadButton > button {
  font-family:'Barlow Condensed',sans-serif !important;font-size:13px !important;
  font-weight:700 !important;letter-spacing:2px !important;text-transform:uppercase !important;
  background:#2a5c3a !important;color:#c9a84c !important;
  border:1px solid rgba(201,168,76,.3) !important;border-radius:3px !important;
}
.stDownloadButton > button:hover {background:#1e3a28 !important;}
div[data-testid="stImage"] img {border-radius:4px;border:1px solid rgba(201,168,76,.2);}
</style>
""", unsafe_allow_html=True)

# ── HELPERS ────────────────────────────────────────────────────
def step_bar(current):
    labels = ["Personal Info", "Payment", "Waiver", "Confirm"]
    nums = ""
    for i, lbl in enumerate(labels, 1):
        cls_n = "active" if i == current else ("done" if i < current else "")
        cls_l = cls_n
        nums += f"<div class='sitem'><div class='snum {cls_n}'>{i}</div><div class='slbl {cls_l}'>{lbl}</div></div>"
    st.markdown(f"<div class='stepbar'>{nums}</div>", unsafe_allow_html=True)

def generate_excel(registrations):
    wb = openpyxl.Workbook()
    hf    = Font(name="Arial", bold=True, color="C9A84C", size=11)
    hfill = PatternFill("solid", fgColor="0A0F0D")
    ctr   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bdr   = Border(
        left=Side(style="thin", color="1C2A20"),
        right=Side(style="thin", color="1C2A20"),
        bottom=Side(style="thin", color="1C2A20"),
    )
    ef  = PatternFill("solid", fgColor="111A14")
    df  = Font(name="Arial", size=10, color="E8E2D4")

    ws = wb.active
    ws.title = "Registrations"
    ws.sheet_properties.tabColor = "C9A84C"

    headers = [
        "#", "Registered At", "First Name", "Last Name", "Email", "Phone",
        "Date of Birth", "Age", "Gender", "Address", "Category", "Fee (PHP)",
        "Shirt Size", "Team/Club", "Emergency Contact", "Emergency Phone",
        "Medical Notes", "GCash Account", "GCash Number", "Payment Proof", "Age Verification",
    ]
    widths = [4,20,14,14,26,16,13,6,10,28,10,10,9,18,20,16,22,20,14,14,18]

    for c, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = hf; cell.fill = hfill; cell.alignment = ctr
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 30

    for i, r in enumerate(registrations, 2):
        fill = ef if i % 2 == 0 else None
        vals = [
            i-1, r.get("registered_at",""), r.get("first_name",""), r.get("last_name",""),
            r.get("email",""), r.get("contact",""), r.get("dob",""), r.get("age",""),
            r.get("gender",""), r.get("address",""), r.get("category",""), r.get("fee",0),
            r.get("shirt",""), r.get("team",""), r.get("emergency_name",""),
            r.get("emergency_phone",""), r.get("medical",""),
            r.get("pay_account_name",""), r.get("pay_account_number",""),
            "Yes" if r.get("has_proof") else "No", r.get("age_verify",""),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = df; cell.border = bdr
            if fill: cell.fill = fill
            if c == 1: cell.alignment = ctr
            if c == 12: cell.number_format = '"₱"#,##0'

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:U{max(len(registrations)+1, 2)}"

    # Sheet 2 – Category Summary
    ws2 = wb.create_sheet("Category Summary")
    ws2.sheet_properties.tabColor = "C9A84C"
    for c, h in enumerate(["Category", "Fee", "Registered", "Revenue"], 1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = hf; cell.fill = hfill; cell.alignment = ctr
        ws2.column_dimensions[get_column_letter(c)].width = [14,12,12,16][c-1]

    for i, (cat, fee) in enumerate(CATEGORIES.items(), 2):
        cnt = sum(1 for r in registrations if r.get("category") == cat)
        rev = cnt * fee
        fill = ef if i % 2 == 0 else None
        for c, v in enumerate([cat, fee, cnt, rev], 1):
            cell = ws2.cell(row=i, column=c, value=v)
            cell.font = df; cell.border = bdr
            if fill: cell.fill = fill
            if c in (2, 4): cell.number_format = '"₱"#,##0'

    tr = len(CATEGORIES) + 2
    tfill = PatternFill("solid", fgColor="1E3A28")
    tf = Font(name="Arial", bold=True, size=10, color="C9A84C")
    for c, v in enumerate(["TOTALS", "", f"=SUM(C2:C{tr-1})", f"=SUM(D2:D{tr-1})"], 1):
        cell = ws2.cell(row=tr, column=c, value=v)
        cell.font = tf; cell.fill = tfill
        if c == 4: cell.number_format = '"₱"#,##0'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── HERO ───────────────────────────────────────────────────────
total   = len(st.session_state.registrations)
revenue = sum(r["fee"] for r in st.session_state.registrations)
today   = sum(1 for r in st.session_state.registrations
              if r["registered_at"].startswith(str(datetime.date.today())))

st.markdown(f"""
<div class='hero'>
  <div style='display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px;'>
    <div>
      <div class='hero-eye'>Miles for Minds &bull; Inaugural Edition</div>
      <div class='hero-title'>SIDLAKANG</div>
      <div class='hero-sub'>Balingoan</div>
      <div>
        <span class='hero-pill'>May 24, 2026</span>
        <span class='hero-pill'>Brgy. Dalihig</span>
        <span class='hero-pill'>BAKYARD Events</span>
        <span class='hero-pill'>Balingoan, MisOr</span>
      </div>
    </div>
    <div style='display:flex;gap:28px;flex-wrap:wrap;'>
      <div><div class='hstat-n'>{total}</div><div class='hstat-l'>Registered</div></div>
      <div><div class='hstat-n'>&#8369;{revenue:,}</div><div class='hstat-l'>Revenue</div></div>
      <div><div class='hstat-n'>{today}</div><div class='hstat-l'>Today</div></div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── TABS ───────────────────────────────────────────────────────
tab_register, tab_registrants, tab_materials = st.tabs([
    "📝  Register", "📋  Registrants", "🖼  Materials & QR"
])

# ════════════════════════════════════════════════════════════════
#  TAB 1 — REGISTER
# ════════════════════════════════════════════════════════════════
with tab_register:

    # SUCCESS SCREEN
    if st.session_state.success_id:
        rid = st.session_state.success_id
        r   = next((x for x in st.session_state.registrations if x["id"] == rid), None)
        st.markdown("## 🏔️ You're In!")
        st.success(f"**{r['first_name']} {r['last_name']}** is registered for **M4M: SIDLAKANG BALINGOAN**!")
        if r:
            rows_html = "".join(f"<div class='srow'><span>{a}</span><span>{b}</span></div>" for a,b in [
                ("Name",      f"{r['first_name']} {r['last_name']}"),
                ("Category",  r["category"]),
                ("Fee",       f"₱{r['fee']:,}"),
                ("T-Shirt",   r["shirt"]),
                ("Race Day",  "May 24, 2026"),
                ("Location",  "Brgy. Dalihig, Balingoan"),
            ])
            st.markdown(f"<div class='success-card'>{rows_html}</div>", unsafe_allow_html=True)
        st.info("📣 Follow **Bakyard** on Facebook for route maps, training tips, and race-day updates.")
        if st.button("Register Another Runner"):
            st.session_state.success_id = None
            st.session_state.step = 1
            st.session_state.form = {}
            st.session_state.proof_bytes = None
            st.session_state.proof_name  = None
            st.rerun()
        st.stop()

    step = st.session_state.step
    step_bar(step)

    # ── STEP 1: PERSONAL INFO ──────────────────────────────────
    if step == 1:
        st.markdown("### Personal Information")
        st.caption("Complete all required fields to continue.")

        f = st.session_state.form

        c1, c2 = st.columns(2)
        with c1:
            fullname = st.text_input("Full Name (First and Last) *", value=f.get("fullname",""), placeholder="Juan Dela Cruz")
        with c2:
            email = st.text_input("Email Address *", value=f.get("email",""), placeholder="juan@email.com")

        c3, c4 = st.columns(2)
        with c3:
            phone = st.text_input("Phone Number *", value=f.get("phone",""), placeholder="+63 912 345 6789")
        with c4:
            dob_val = f.get("dob", None)
            if isinstance(dob_val, str) and dob_val:
                try:    dob_val = datetime.date.fromisoformat(dob_val)
                except: dob_val = None
            dob = st.date_input("Date of Birth *", value=dob_val or datetime.date(2000,1,1),
                                min_value=datetime.date(1930,1,1), max_value=datetime.date.today())

        c5, c6 = st.columns(2)
        with c5:
            today_d = datetime.date.today()
            auto_age = today_d.year - dob.year - ((today_d.month, today_d.day) < (dob.month, dob.day))
            age = st.number_input("Age *", min_value=1, max_value=100, value=int(f.get("age", auto_age) or auto_age))
        with c6:
            gender = st.radio("Gender *", GENDERS, index=GENDERS.index(f.get("gender", GENDERS[0])) if f.get("gender") in GENDERS else 0, horizontal=True)

        address = st.text_input("Home Address *", value=f.get("address",""), placeholder="Street, Barangay, City, Province")

        c7, c8 = st.columns(2)
        with c7:
            cat_opts = list(CATEGORIES.keys())
            cat_idx  = cat_opts.index(f.get("category")) if f.get("category") in cat_opts else 0
            category = st.selectbox("Category *", cat_opts,
                                    index=cat_idx,
                                    format_func=lambda k: f"{k} — ₱{CATEGORIES[k]:,}")
        with c8:
            shirt_idx = SHIRT_SIZES.index(f.get("shirt")) if f.get("shirt") in SHIRT_SIZES else 2
            shirt = st.selectbox("Event T-Shirt Size (Unisex) *", SHIRT_SIZES, index=shirt_idx)

        team = st.text_input("Team / Running Club Name", value=f.get("team",""), placeholder="Optional")

        c9, c10 = st.columns(2)
        with c9:
            ecname  = st.text_input("Emergency Contact Name *", value=f.get("ecname",""), placeholder="Maria Dela Cruz")
        with c10:
            ecphone = st.text_input("Emergency Contact Phone *", value=f.get("ecphone",""), placeholder="+63 912 345 6789")

        medical = st.text_area("Medical Conditions or Allergies *", value=f.get("medical",""),
                               placeholder="Type 'None' if you have no conditions our medics need to be aware of.",
                               height=80)
        st.caption("Type 'None' if not applicable.")

        st.divider()
        _, col_btn = st.columns([3, 1])
        with col_btn:
            next1 = st.button("Continue →", key="next1", use_container_width=True)

        if next1:
            errs = []
            if not fullname.strip():          errs.append("Full name is required.")
            if not email.strip() or "@" not in email: errs.append("Valid email is required.")
            if not phone.strip():             errs.append("Phone number is required.")
            if not address.strip():           errs.append("Home address is required.")
            if not ecname.strip():            errs.append("Emergency contact name is required.")
            if not ecphone.strip():           errs.append("Emergency contact phone is required.")
            if not medical.strip():           errs.append("Medical field is required (type 'None' if none).")
            if errs:
                for e in errs: st.error(e)
            else:
                st.session_state.form.update({
                    "fullname": fullname.strip(), "email": email.strip(),
                    "phone": phone.strip(), "dob": str(dob), "age": age,
                    "gender": gender, "address": address.strip(),
                    "category": category, "shirt": shirt, "team": team.strip(),
                    "ecname": ecname.strip(), "ecphone": ecphone.strip(),
                    "medical": medical.strip(),
                })
                st.session_state.step = 2
                st.rerun()

    # ── STEP 2: PAYMENT ───────────────────────────────────────
    elif step == 2:
        st.markdown("### Payment")
        st.caption("Send your fee via GCash, then upload your screenshot proof.")

        cat = st.session_state.form.get("category","—")
        fee = CATEGORIES.get(cat, 0)

        st.markdown(f"""
        <div class='pay-box'>
          <div class='pay-title'>GCash Payment Details</div>
          <div style='font-size:14px;color:#e8e2d4;margin-bottom:4px;'><span style='color:#8a9a8a'>Account Name:</span> &nbsp;<strong>Jorgil Amarga</strong></div>
          <div style='font-size:14px;color:#e8e2d4;margin-bottom:14px;'><span style='color:#8a9a8a'>GCash Number:</span> &nbsp;<strong>0916 481 3822</strong></div>
          <div class='fee-grid'>
            <div class='fee-card'><div class='fee-dist'>7KM</div><div class='fee-amt'>&#8369;800</div></div>
            <div class='fee-card'><div class='fee-dist'>15KM</div><div class='fee-amt'>&#8369;1,200</div></div>
            <div class='fee-card'><div class='fee-dist'>21KM</div><div class='fee-amt'>&#8369;1,400</div></div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Show QR codes if uploaded
        qr_gcash = st.session_state.assets["qr_gcash"]["img"]
        qr_maya  = st.session_state.assets["qr_maya"]["img"]
        if qr_gcash or qr_maya:
            qc1, qc2 = st.columns(2)
            with qc1:
                if qr_gcash:
                    st.image(qr_gcash, caption="GCash QR Code", use_container_width=True)
            with qc2:
                if qr_maya:
                    st.image(qr_maya, caption="Maya QR Code", use_container_width=True)

        st.info(f"Your selected category: **{cat} — ₱{fee:,}**")

        f = st.session_state.form
        p1, p2 = st.columns(2)
        with p1:
            payname   = st.text_input("GCash / Maya Account Name *", value=f.get("payname",""), placeholder="Full name on your account")
        with p2:
            paynumber = st.text_input("GCash Mobile Number (11 digits) *", value=f.get("paynumber",""), placeholder="09XXXXXXXXX", max_chars=11)

        proof_file = st.file_uploader(
            "Payment Screenshot / Proof of Transaction *",
            type=["jpg","jpeg","png","pdf"],
            help="Upload a clear screenshot of your successful GCash transaction."
        )
        if proof_file:
            st.session_state.proof_bytes = proof_file.read()
            st.session_state.proof_name  = proof_file.name
            st.success(f"✓ {proof_file.name} uploaded")
        elif st.session_state.proof_name:
            st.success(f"✓ {st.session_state.proof_name} (already uploaded)")

        st.divider()
        bc, nc = st.columns([1, 3])
        with bc:
            if st.button("← Back", key="back2"):
                st.session_state.step = 1
                st.rerun()
        with nc:
            _, btn_col = st.columns([2,1])
            with btn_col:
                next2 = st.button("Continue →", key="next2", use_container_width=True)

        if next2:
            errs = []
            if not payname.strip():               errs.append("GCash/Maya account name is required.")
            if not paynumber.strip() or len(paynumber) != 11: errs.append("Valid 11-digit GCash number is required.")
            if not st.session_state.proof_bytes:  errs.append("Please upload your payment proof screenshot.")
            if errs:
                for e in errs: st.error(e)
            else:
                st.session_state.form.update({"payname": payname.strip(), "paynumber": paynumber.strip()})
                st.session_state.step = 3
                st.rerun()

    # ── STEP 3: WAIVER ────────────────────────────────────────
    elif step == 3:
        st.markdown("### Waiver & Release of Liability")
        st.caption("Please read each section carefully and check all boxes to proceed. This is legally binding.")

        all_checked = True
        for wid, (title, text, confirm) in enumerate(WAIVERS, 1):
            st.markdown(f"""
            <div class='wbox'>
              <div class='wbox-title'>{wid}. {title}</div>
              <div class='wbox-text'>{text}</div>
            </div>
            """, unsafe_allow_html=True)
            checked = st.checkbox(confirm, key=f"waiver_{wid}")
            if not checked:
                all_checked = False

        st.markdown("---")
        st.markdown("**Age Verification**")
        age_verify = st.radio(
            "Select the option that applies to you *",
            options=["adult", "guardian"],
            format_func=lambda x: (
                "I am 18 years of age or older. I have read this waiver and voluntarily agree, understanding I am giving up substantial legal rights."
                if x == "adult" else
                "I am the Parent / Legal Guardian of a participant under 18. I agree to the waiver terms on their behalf."
            ),
            index=0,
            key="age_verify_radio"
        )

        st.divider()
        bc2, nc2 = st.columns([1, 3])
        with bc2:
            if st.button("← Back", key="back3"):
                st.session_state.step = 2
                st.rerun()
        with nc2:
            _, btn_col2 = st.columns([2,1])
            with btn_col2:
                next3 = st.button("Continue →", key="next3", use_container_width=True)

        if next3:
            if not all_checked:
                st.error("Please check all waiver boxes before continuing.")
            else:
                st.session_state.form["ageVerify"] = age_verify
                st.session_state.step = 4
                st.rerun()

    # ── STEP 4: CONFIRM ───────────────────────────────────────
    elif step == 4:
        st.markdown("### Confirm Registration")
        st.caption("Review your information before final submission.")

        f   = st.session_state.form
        fee = CATEGORIES.get(f.get("category",""), 0)

        rows = [
            ("Full Name",         f.get("fullname","—")),
            ("Email",             f.get("email","—")),
            ("Phone",             f.get("phone","—")),
            ("Date of Birth",     str(f.get("dob","—"))),
            ("Age",               str(f.get("age","—"))),
            ("Gender",            f.get("gender","—")),
            ("Address",           f.get("address","—")),
            ("Category",          f.get("category","—")),
            ("Registration Fee",  f"₱{fee:,}"),
            ("T-Shirt Size",      f.get("shirt","—")),
            ("Team / Club",       f.get("team","—") or "—"),
            ("Emergency Contact", f"{f.get('ecname','—')} · {f.get('ecphone','—')}"),
            ("Medical",           f.get("medical","—")),
            ("GCash Account",     f"{f.get('payname','—')} · {f.get('paynumber','—')}"),
            ("Age Verification",  "18+ (Self)" if f.get("ageVerify")=="adult" else "Parent/Guardian"),
            ("Payment Proof",     f"✓ {st.session_state.proof_name}" if st.session_state.proof_name else "None"),
        ]
        rows_html = "".join(f"<div class='srow'><span>{a}</span><span>{b}</span></div>" for a,b in rows)
        st.markdown(f"""
        <div style='background:rgba(255,255,255,.04);border:1px solid rgba(201,168,76,.22);border-radius:4px;padding:20px 24px;margin-bottom:20px'>
          <div style='font-family:Barlow Condensed,sans-serif;font-size:11px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#c9a84c;margin-bottom:14px'>Registration Summary</div>
          {rows_html}
        </div>
        """, unsafe_allow_html=True)

        st.info("By clicking **Submit Registration**, you confirm all information is accurate and you have agreed to all waiver terms.")

        bc3, nc3 = st.columns([1, 3])
        with bc3:
            if st.button("← Back", key="back4"):
                st.session_state.step = 3
                st.rerun()
        with nc3:
            _, btn_col3 = st.columns([2,1])
            with btn_col3:
                submit = st.button("Submit Registration", key="submit", use_container_width=True)

        if submit:
            name_parts = f.get("fullname","").strip().split()
            first = name_parts[0] if name_parts else ""
            last  = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            rid   = str(uuid.uuid4())[:8].upper()
            reg   = {
                "id":                 rid,
                "registered_at":      datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "first_name":         first,
                "last_name":          last,
                "email":              f.get("email",""),
                "contact":            f.get("phone",""),
                "dob":                str(f.get("dob","")),
                "age":                f.get("age",""),
                "gender":             f.get("gender",""),
                "address":            f.get("address",""),
                "category":           f.get("category",""),
                "fee":                CATEGORIES.get(f.get("category",""), 0),
                "shirt":              f.get("shirt",""),
                "team":               f.get("team",""),
                "emergency_name":     f.get("ecname",""),
                "emergency_phone":    f.get("ecphone",""),
                "medical":            f.get("medical",""),
                "pay_account_name":   f.get("payname",""),
                "pay_account_number": f.get("paynumber",""),
                "has_proof":          bool(st.session_state.proof_bytes),
                "age_verify":         f.get("ageVerify",""),
            }
            st.session_state.registrations.append(reg)
            st.session_state.success_id = rid
            st.session_state.form       = {}
            st.session_state.step       = 1
            st.session_state.proof_bytes = None
            st.session_state.proof_name  = None
            st.rerun()

# ════════════════════════════════════════════════════════════════
#  TAB 2 — REGISTRANTS
# ════════════════════════════════════════════════════════════════
with tab_registrants:
    regs  = st.session_state.registrations
    total = len(regs)
    rev   = sum(r["fee"] for r in regs)
    today_c = sum(1 for r in regs if r["registered_at"].startswith(str(datetime.date.today())))
    by7   = sum(1 for r in regs if r.get("category") == "7KM")
    by15  = sum(1 for r in regs if r.get("category") == "15KM")
    by21  = sum(1 for r in regs if r.get("category") == "21KM")

    s1,s2,s3,s4 = st.columns(4)
    with s1:
        st.markdown(f"<div class='scard'><div class='scard-n'>{total}</div><div class='scard-l'>Total</div></div>", unsafe_allow_html=True)
    with s2:
        st.markdown(f"<div class='scard'><div class='scard-n'>&#8369;{rev:,}</div><div class='scard-l'>Revenue</div></div>", unsafe_allow_html=True)
    with s3:
        st.markdown(f"<div class='scard'><div class='scard-n'>{today_c}</div><div class='scard-l'>Today</div></div>", unsafe_allow_html=True)
    with s4:
        st.markdown(f"<div class='scard'><div class='scard-n'>{by7}/{by15}/{by21}</div><div class='scard-l'>7K / 15K / 21K</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if regs:
        excel_bytes = generate_excel(regs)
        st.download_button(
            label=f"⬇  Download Excel Report  ({total} registrants)",
            data=excel_bytes,
            file_name=f"sidlakang_registrations_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
        )

        import pandas as pd
        df_data = [{
            "#":           i+1,
            "Name":        f"{r['first_name']} {r['last_name']}",
            "Category":    r.get("category",""),
            "Shirt":       r.get("shirt",""),
            "Contact":     r.get("contact",""),
            "Fee":         f"₱{r.get('fee',0):,}",
            "Proof":       "✓" if r.get("has_proof") else "—",
            "Registered":  r.get("registered_at","")[:16],
        } for i,r in enumerate(reversed(regs))]
        st.dataframe(pd.DataFrame(df_data), use_container_width=True, hide_index=True)
    else:
        st.info("No registrations yet. Use the Register tab to add participants.")

# ════════════════════════════════════════════════════════════════
#  TAB 3 — MATERIALS & QR
# ════════════════════════════════════════════════════════════════
with tab_materials:
    st.markdown("#### Race Materials")
    st.caption("Upload images for shirts, medals, QR codes and event materials. QR codes will automatically appear in the payment step.")

    ASSET_GROUPS = [
        ("Race Items",    ["shirt_7km","shirt_15km","shirt_21km","medal"]),
        ("Payment QR",   ["qr_gcash","qr_maya"]),
        ("Event Assets", ["banner","route_map"]),
    ]

    for group_label, keys in ASSET_GROUPS:
        st.markdown(f"**{group_label}**")
        cols = st.columns(len(keys))
        for col, key in zip(cols, keys):
            asset = st.session_state.assets[key]
            with col:
                if asset["img"]:
                    st.image(asset["img"], caption=asset["label"], use_container_width=True)
                    up = st.file_uploader(f"Replace {asset['label']}", type=["jpg","jpeg","png","webp"],
                                         key=f"mat_{key}", label_visibility="collapsed")
                else:
                    st.markdown(f"<div style='border:1px dashed rgba(201,168,76,.35);border-radius:3px;padding:12px;text-align:center;font-size:11px;letter-spacing:1px;text-transform:uppercase;color:#8a9a8a;margin-bottom:6px;'>{asset['label']}<br><span style=\"font-size:18px\">↑</span></div>", unsafe_allow_html=True)
                    up = st.file_uploader(f"Upload {asset['label']}", type=["jpg","jpeg","png","webp"],
                                         key=f"mat_{key}", label_visibility="collapsed")
                if up:
                    img = Image.open(up)
                    st.session_state.assets[key]["img"]  = img
                    st.session_state.assets[key]["name"] = up.name
                    st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
