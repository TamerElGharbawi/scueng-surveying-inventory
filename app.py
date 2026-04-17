# ================================================================
#  نظام جرد معمل المساحة | Surveying Lab Inventory Agent
#  جامعة قناة السويس – كلية الهندسة
#  Version 2.0 — Mobile-First | AI-Powered | Professional Arabic Reports
# ================================================================

import streamlit as st
import google.generativeai as genai
from PIL import Image
import io, base64, json, os, uuid, re, requests, datetime
from pathlib import Path

# ── PDF ──────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image as RLImage, PageBreak, HRFlowable, KeepTogether,
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
import arabic_reshaper
from bidi.algorithm import get_display

# ── DOCX ─────────────────────────────────────────────────────────
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ================================================================
# ⚙️  CONSTANTS
# ================================================================
PAGE_W, PAGE_H = A4
FONT_DIR = Path("/tmp/ar_fonts")
FONT_DIR.mkdir(exist_ok=True)
FONT_REG  = str(FONT_DIR / "Cairo-Regular.ttf")
FONT_BOLD = str(FONT_DIR / "Cairo-Bold.ttf")
FONT_SEMI = str(FONT_DIR / "Cairo-SemiBold.ttf")

UNIV_NAME   = "جامعة قناة السويس"
FAC_NAME    = "كلية الهندسة"
LAB_NAME    = "معمل المساحة"
RPT_TITLE   = "كشف جرد أجهزة ومعدات المعمل"

STATUS_OPT  = ["ممتاز", "جيد جداً", "جيد", "يحتاج صيانة", "معطل"]

# PDF colour palette
C_PRI  = colors.HexColor("#1a5276")
C_SEC  = colors.HexColor("#2e86c1")
C_ACC  = colors.HexColor("#f39c12")
C_LITE = colors.HexColor("#eaf0fb")
C_GRAY = colors.HexColor("#aab7b8")
C_OK   = colors.HexColor("#1e8449")
C_WARN = colors.HexColor("#e67e22")
C_ERR  = colors.HexColor("#c0392b")
C_BLK  = colors.HexColor("#1c2833")


# ================================================================
# 🔤  FONT SETUP (downloaded once, cached)
# ================================================================
@st.cache_resource
def setup_fonts() -> bool:
    urls = {
        FONT_REG:  "https://raw.githubusercontent.com/google/fonts/main/ofl/cairo/static/Cairo-Regular.ttf",
        FONT_BOLD: "https://raw.githubusercontent.com/google/fonts/main/ofl/cairo/static/Cairo-Bold.ttf",
        FONT_SEMI: "https://raw.githubusercontent.com/google/fonts/main/ofl/cairo/static/Cairo-SemiBold.ttf",
    }
    for path, url in urls.items():
        if not os.path.exists(path):
            try:
                r = requests.get(url, timeout=20)
                r.raise_for_status()
                with open(path, "wb") as f:
                    f.write(r.content)
            except Exception:
                return False
    try:
        pdfmetrics.registerFont(TTFont("Cairo",      FONT_REG))
        pdfmetrics.registerFont(TTFont("Cairo-Bold", FONT_BOLD))
        pdfmetrics.registerFont(TTFont("Cairo-Semi", FONT_SEMI))
        return True
    except Exception:
        return False


def ar(txt: str) -> str:
    """Reshape + BiDi Arabic text for ReportLab."""
    return get_display(arabic_reshaper.reshape(str(txt)))


# ================================================================
# 🎨  PAGE CONFIG & CSS
# ================================================================
st.set_page_config(
    page_title="جرد معمل المساحة | SCU",
    page_icon="🔭",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');

html, body, .stApp                    { direction:rtl!important; font-family:'Cairo',sans-serif!important; }
.main .block-container                { padding:1rem .75rem!important; max-width:100%!important; }
h1,h2,h3,h4                          { font-family:'Cairo',sans-serif!important; text-align:right!important; color:#1a5276!important; }
.stButton>button                      { width:100%!important; min-height:3rem!important; font-size:1.05rem!important;
                                        font-family:'Cairo',sans-serif!important; font-weight:700!important;
                                        border-radius:12px!important; transition:all .2s!important; }
.stButton>button:hover                { transform:translateY(-2px)!important; box-shadow:0 4px 14px rgba(0,0,0,.18)!important; }
.stTextInput input, .stTextArea textarea { direction:rtl!important; text-align:right!important;
                                           font-family:'Cairo',sans-serif!important; font-size:1rem!important; border-radius:10px!important; }
label, .stTextInput label, .stSelectbox label, .stTextArea label
                                      { font-family:'Cairo',sans-serif!important; font-size:1rem!important;
                                        font-weight:600!important; color:#1a5276!important; text-align:right!important; }
[data-testid="metric-container"]      { direction:rtl!important; text-align:right!important; background:#eaf0fb!important;
                                        border-radius:12px!important; padding:.75rem!important; border-right:4px solid #1a5276!important; }
[data-testid="metric-container"] label { font-family:'Cairo',sans-serif!important; }
[data-testid="metric-container"] [data-testid="metric-value"] { font-family:'Cairo',sans-serif!important; color:#1a5276!important; }
.stExpander                           { border:1.5px solid #2e86c1!important; border-radius:12px!important; direction:rtl!important; }
.stExpander summary                   { font-family:'Cairo',sans-serif!important; font-weight:700!important; direction:rtl!important; }
.stAlert                              { direction:rtl!important; font-family:'Cairo',sans-serif!important; border-radius:10px!important; }
.stTabs [data-baseweb="tab"]          { font-family:'Cairo',sans-serif!important; font-size:1rem!important; font-weight:600!important; }
.stTabs [data-baseweb="tab-list"]     { direction:rtl!important; }
.stFileUploader                       { direction:rtl!important; }
.stFileUploader label                 { font-family:'Cairo',sans-serif!important; }
.stRadio>div, .stCheckbox>label       { direction:rtl!important; font-family:'Cairo',sans-serif!important; }
.stSelectbox [data-baseweb="select"]  { direction:rtl!important; }
#MainMenu, footer, header             { visibility:hidden; }
::-webkit-scrollbar                   { width:5px; }
::-webkit-scrollbar-thumb             { background:#2e86c1; border-radius:3px; }

/* custom component styles */
.app-header { background:linear-gradient(135deg,#1a5276,#2e86c1); color:white;
               padding:1rem 1.5rem; border-radius:14px; text-align:center; margin-bottom:1.2rem; }
.app-header h1 { color:white!important; font-size:1.4rem!important; margin:0!important; }
.app-header p  { color:#aed6f1!important; margin:.2rem 0 0!important; font-size:.85rem!important; }

.step-label { background:#eaf0fb; border-right:4px solid #1a5276; padding:.5rem 1rem;
               border-radius:0 8px 8px 0; font-family:'Cairo',sans-serif; font-weight:700;
               color:#1a5276; margin-bottom:.75rem; }

.photo-card { border:2px solid #2e86c1; border-radius:12px; padding:8px;
               text-align:center; background:white; direction:rtl; }
.badge-primary { background:#1a5276; color:white; padding:2px 10px; border-radius:20px;
                  font-size:.72rem; font-family:'Cairo',sans-serif; font-weight:700; }
.badge-ref { background:#f39c12; color:white; padding:2px 10px; border-radius:20px;
              font-size:.72rem; font-family:'Cairo',sans-serif; font-weight:700; }

.inv-card { background:white; border:1.5px solid #d5d8dc; border-radius:14px;
             padding:.9rem; margin-bottom:.9rem; direction:rtl;
             box-shadow:0 2px 8px rgba(0,0,0,.06); }
.inv-card:hover { border-color:#2e86c1; box-shadow:0 4px 16px rgba(46,134,193,.15); }

@keyframes dupPulse { 0%{box-shadow:0 0 0 0 rgba(192,57,43,.5)} 70%{box-shadow:0 0 0 10px rgba(192,57,43,0)} 100%{box-shadow:0 0 0 0 rgba(192,57,43,0)} }
.dup-warning { background:#fdedec; border:2px solid #c0392b; border-radius:12px;
                padding:1rem; direction:rtl; animation:dupPulse 1.4s infinite; }
.success-flash { background:linear-gradient(135deg,#1e8449,#27ae60); color:white;
                  padding:1rem; border-radius:12px; text-align:center;
                  font-family:'Cairo',sans-serif; font-weight:700; font-size:1.05rem; }
.nav-bar { display:flex; gap:8px; margin-bottom:1rem; }
</style>
""", unsafe_allow_html=True)


# ================================================================
# 🗃️  SESSION STATE
# ================================================================
def init_state():
    defaults = dict(
        inventory         = [],
        captured_photos   = [],   # staging: list of {id,data,is_primary,label}
        photo_hashes      = set(),
        page              = "capture",
        professor_name    = "",
        report_date       = datetime.date.today().strftime("%Y/%m/%d"),
        ai_result         = None,
        gemini_key        = "",
    )
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ================================================================
# 🤖  GEMINI AI EXTRACTION
# ================================================================
def extract_device_info(images: list[bytes], api_key: str) -> dict:
    if not api_key:
        return {"error": "أدخل Gemini API Key في الإعدادات أولاً"}
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")

        parts: list = ["""
أنت نظام متخصص في التعرف على الأجهزة والمعدات الهندسية من الصور.
حلّل الصور المُرفقة بدقة واستخرج البيانات التالية.
اقرأ لوحة البيانات (nameplate) أو جسم الجهاز بعناية.

أجب بـ JSON صالح فقط، بدون أي نص خارجي أو backticks:
{
  "device_type": "نوع الجهاز (مثل: جهاز مستوي آلي، محطة شاملة، GPS، ميزان، إلخ)",
  "brand": "الماركة أو المصنّع",
  "serial_number": "الرقم التسلسلي أو 'غير مقروء'",
  "condition": "ممتاز أو جيد جداً أو جيد أو يحتاج صيانة أو معطل",
  "notes": "ملاحظات مفيدة من الصور",
  "confidence": "high أو medium أو low"
}
"""]

        for img_bytes in images:
            pil = Image.open(io.BytesIO(img_bytes))
            if max(pil.size) > 1600:
                pil.thumbnail((1600, 1600), Image.LANCZOS)
            buf = io.BytesIO()
            pil.save(buf, format="JPEG", quality=85)
            parts.append({"mime_type": "image/jpeg",
                           "data": base64.b64encode(buf.getvalue()).decode()})

        resp = model.generate_content(parts)
        raw = re.sub(r"```json|```", "", resp.text).strip()
        return json.loads(raw)

    except json.JSONDecodeError:
        return {"device_type":"","brand":"","serial_number":"",
                "condition":"جيد","notes":"تعذّر تحليل الاستجابة","confidence":"low"}
    except Exception as e:
        return {"error": f"خطأ Gemini: {e}"}


# ================================================================
# 🔍  DUPLICATE DETECTION
# ================================================================
def is_duplicate(serial: str, exclude_idx: int = -1) -> bool:
    if not serial or serial.strip().upper() in ["", "غير مقروء", "غير محدد"]:
        return False
    for i, item in enumerate(st.session_state.inventory):
        if i == exclude_idx:
            continue
        if item.get("serial_number","").strip().upper() == serial.strip().upper():
            return True
    return False


def get_duplicates() -> list[tuple[str, list[int]]]:
    seen: dict[str, list[int]] = {}
    for i, item in enumerate(st.session_state.inventory):
        sn = item.get("serial_number","").strip().upper()
        if sn and sn not in ["غير مقروء", "غير محدد"]:
            seen.setdefault(sn, []).append(i + 1)
    return [(sn, idx) for sn, idx in seen.items() if len(idx) > 1]


# ================================================================
# 🖼️  IMAGE HELPERS
# ================================================================
def resize_img(data: bytes, mx: int = 900) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    if max(img.size) > mx:
        img.thumbnail((mx, mx), Image.LANCZOS)
    buf = io.BytesIO(); img.save(buf, "JPEG", quality=82)
    return buf.getvalue()

def thumb(data: bytes, sz=(150,150)) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.thumbnail(sz, Image.LANCZOS)
    canvas = Image.new("RGB", sz, (255,255,255))
    canvas.paste(img, ((sz[0]-img.size[0])//2, (sz[1]-img.size[1])//2))
    buf = io.BytesIO(); canvas.save(buf, "JPEG", quality=85)
    return buf.getvalue()

def b64img(data: bytes) -> str:
    return base64.b64encode(data).decode()


# ================================================================
# 📄  PDF REPORT
# ================================================================
class PDFReport:
    def __init__(self, inv, prof, date):
        self.inv   = inv
        self.prof  = prof
        self.date  = date
        self.dups  = get_duplicates()
        self.buf   = io.BytesIO()

    # ── styles ───────────────────────────────────────────────────
    def _ps(self, size=10, bold=False, align=TA_RIGHT, color=C_BLK):
        fn = "Cairo-Bold" if bold else "Cairo"
        return ParagraphStyle(f"s{uuid.uuid4().hex[:4]}",
            fontName=fn, fontSize=size, textColor=color,
            alignment=align, leading=int(size*1.65), wordWrap="CJK")

    def _p(self, txt, **kw):
        return Paragraph(ar(txt), self._ps(**kw))

    # ── header/footer callback ────────────────────────────────────
    def _hf(self, canvas, doc):
        canvas.saveState()
        W, H = A4

        # ── header bar ──
        canvas.setFillColor(C_PRI)
        canvas.rect(0, H-2.3*cm, W, 2.3*cm, fill=1, stroke=0)

        # logo placeholder box
        canvas.setFillColor(colors.white)
        canvas.setStrokeColor(C_ACC)
        canvas.setLineWidth(1.2)
        canvas.rect(W-2.9*cm, H-2.1*cm, 2.3*cm, 1.9*cm, fill=1, stroke=1)
        canvas.setFillColor(C_GRAY)
        canvas.setFont("Cairo", 6)
        canvas.drawCentredString(W-1.75*cm, H-1.25*cm, ar("شعار"))
        canvas.drawCentredString(W-1.75*cm, H-1.55*cm, ar("الجامعة"))

        # university text
        canvas.setFillColor(colors.white)
        canvas.setFont("Cairo-Bold", 11)
        canvas.drawRightString(W-3.3*cm, H-1.0*cm, ar(f"{UNIV_NAME}  |  {FAC_NAME}"))
        canvas.setFont("Cairo", 9)
        canvas.drawRightString(W-3.3*cm, H-1.6*cm, ar(f"{LAB_NAME}   |   {self.date}"))

        # professor (left side)
        if self.prof:
            canvas.setFillColor(colors.HexColor("#aed6f1"))
            canvas.setFont("Cairo", 8)
            canvas.drawString(0.6*cm, H-1.2*cm, ar(f"أ.د / {self.prof}"))

        # ── footer bar ──
        canvas.setFillColor(C_PRI)
        canvas.rect(0, 0, W, 1.2*cm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Cairo", 8)
        canvas.drawCentredString(W/2, 0.42*cm, ar(f"{UNIV_NAME} — {FAC_NAME} — {LAB_NAME}"))
        canvas.setFont("Cairo-Bold", 8)
        canvas.drawString(0.8*cm, 0.42*cm, ar(f"صفحة {doc.page}"))
        canvas.restoreState()

    # ── COVER PAGE ────────────────────────────────────────────────
    def _cover(self):
        el = [Spacer(1, 2.5*cm)]

        # logo box centred
        lbox = Table([[self._p("[ شعار جامعة قناة السويس ]",
                                size=10, align=TA_CENTER, color=C_GRAY)]],
                     colWidths=[6*cm], rowHeights=[5*cm])
        lbox.setStyle(TableStyle([
            ("ALIGN",      (0,0),(-1,-1),"CENTER"),
            ("VALIGN",     (0,0),(-1,-1),"MIDDLE"),
            ("BOX",        (0,0),(-1,-1), 2, C_ACC),
            ("BACKGROUND", (0,0),(-1,-1), C_LITE),
        ]))
        outer = Table([[lbox]], colWidths=[PAGE_W-4*cm])
        outer.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
        el.extend([outer, Spacer(1,.8*cm)])

        el.append(self._p(UNIV_NAME,  size=20, bold=True, align=TA_CENTER, color=C_PRI))
        el.append(Spacer(1,.25*cm))
        el.append(self._p(FAC_NAME,   size=15, bold=True, align=TA_CENTER, color=C_SEC))
        el.append(Spacer(1,.2*cm))
        el.append(self._p(LAB_NAME,   size=13, align=TA_CENTER, color=C_BLK))
        el.append(Spacer(1,1.2*cm))
        el.append(HRFlowable(width="75%", thickness=2, color=C_ACC, hAlign="CENTER"))
        el.append(Spacer(1,.8*cm))
        el.append(self._p(RPT_TITLE,  size=17, bold=True, align=TA_CENTER, color=C_BLK))
        el.append(Spacer(1,.4*cm))
        yr = datetime.datetime.now().year
        el.append(self._p(f"العام الدراسي {yr-1} / {yr}", size=12, align=TA_CENTER))
        el.append(Spacer(1,1.8*cm))

        rows = [
            [self._p("تاريخ الجرد :", size=11, bold=True), self._p(self.date, size=11)],
            [self._p("المُعِد :", size=11, bold=True), self._p(self.prof or "—", size=11)],
            [self._p("إجمالي الأجهزة :", size=11, bold=True),
             self._p(str(len(self.inv)), size=12, bold=True, color=C_PRI)],
        ]
        info = Table(rows, colWidths=[4.5*cm, 8*cm])
        info.setStyle(TableStyle([
            ("ALIGN",       (0,0),(-1,-1),"RIGHT"),
            ("VALIGN",      (0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,0),(-1,-1),[C_LITE, colors.white, C_LITE]),
            ("BOX",         (0,0),(-1,-1), 1, C_SEC),
            ("INNERGRID",   (0,0),(-1,-1), .5, C_GRAY),
            ("TOPPADDING",  (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1),8),
            ("LEFTPADDING", (0,0),(-1,-1), 8),
            ("RIGHTPADDING",(0,0),(-1,-1), 8),
        ]))
        outer2 = Table([[info]], colWidths=[PAGE_W-4*cm])
        outer2.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
        el.append(outer2)
        return el

    # ── INTRO ─────────────────────────────────────────────────────
    def _intro(self):
        el = [self._p("تمهيد", size=13, bold=True, color=C_PRI),
              HRFlowable(width="100%", thickness=1, color=C_SEC),
              Spacer(1,.3*cm)]
        txt = (
            f"بناءً على التوجيهات الصادرة من إدارة كلية الهندسة بجامعة قناة السويس، "
            f"وفي إطار برنامج الجرد الدوري للأجهزة والمعدات، أُعِدَّ هذا الكشف الشامل "
            f"لأجهزة ومعدات {LAB_NAME} بتاريخ {self.date} تحت إشراف "
            f"{'أ.د / ' + self.prof if self.prof else 'المسؤول المختص'}. "
            f"يشمل هذا الكشف جميع الأجهزة المتاحة بالمعمل مع توثيق حالتها الراهنة "
            f"وصورها وأرقامها التسلسلية، ويُعدُّ وثيقةً رسميةً معتمدةً من الكلية."
        )
        el.append(self._p(txt, size=10))
        return el

    # ── STATS ─────────────────────────────────────────────────────
    def _stats(self):
        el = [self._p("الملخص الإحصائي", size=13, bold=True, color=C_PRI),
              HRFlowable(width="100%", thickness=1, color=C_SEC),
              Spacer(1,.3*cm)]

        brands = {}; conds = {}
        for item in self.inv:
            b = item.get("brand","غير محدد"); brands[b] = brands.get(b,0)+1
            c = item.get("condition","غير محدد"); conds[c] = conds.get(c,0)+1

        need_maint = conds.get("يحتاج صيانة",0) + conds.get("معطل",0)
        excellent  = conds.get("ممتاز",0) + conds.get("جيد جداً",0)

        rows = [
            [self._p("البيان",size=10,bold=True,color=colors.white),
             self._p("القيمة",size=10,bold=True,color=colors.white)],
            [self._p("إجمالي عدد الأجهزة",size=10),
             self._p(str(len(self.inv)),size=10,bold=True)],
            [self._p("عدد الماركات المختلفة",size=10),
             self._p(str(len(brands)),size=10,bold=True)],
            [self._p("أجهزة في حالة ممتازة / جيد جداً",size=10),
             self._p(str(excellent),size=10,bold=True,color=C_OK)],
            [self._p("أجهزة تحتاج صيانة / معطلة",size=10),
             self._p(str(need_maint),size=10,bold=True,color=C_ERR)],
        ]
        if self.dups:
            rows.append([self._p("أرقام تسلسلية مكررة",size=10,color=C_ERR),
                          self._p(str(len(self.dups)),size=10,bold=True,color=C_ERR)])

        tbl = Table(rows, colWidths=[11*cm, 4*cm])
        style = [
            ("BACKGROUND",    (0,0),(-1,0),  C_PRI),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [C_LITE, colors.white]),
            ("BOX",           (0,0),(-1,-1),  1, C_PRI),
            ("INNERGRID",     (0,0),(-1,-1), .5, C_GRAY),
            ("ALIGN",         (0,0),(-1,-1), "RIGHT"),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 6),
            ("BOTTOMPADDING", (0,0),(-1,-1), 6),
            ("LEFTPADDING",   (0,0),(-1,-1), 6),
            ("RIGHTPADDING",  (0,0),(-1,-1), 6),
        ]
        tbl.setStyle(TableStyle(style))
        el.append(tbl)

        if brands:
            el.append(Spacer(1,.4*cm))
            line = " | ".join(f"{b}: {c}" for b,c in sorted(brands.items(), key=lambda x:-x[1]))
            el.append(self._p(f"توزيع الماركات:  {line}", size=9, color=C_SEC))

        return el

    # ── MAIN INVENTORY TABLE ──────────────────────────────────────
    def _inv_table(self):
        el = [PageBreak(),
              self._p("كشف الأجهزة والمعدات", size=13, bold=True, color=C_PRI),
              HRFlowable(width="100%", thickness=1, color=C_SEC),
              Spacer(1,.3*cm)]

        # Columns: م | نوع | ماركة | رقم | حالة | ملاحظات | صورة
        CW = [1*cm, 3.3*cm, 2.5*cm, 2.4*cm, 2*cm, 3*cm, 3*cm]

        hdr = [
            self._p("م",               size=9, bold=True, align=TA_CENTER, color=colors.white),
            self._p("نوع الجهاز",       size=9, bold=True, color=colors.white),
            self._p("الماركة / المصنّع",size=9, bold=True, color=colors.white),
            self._p("الرقم التسلسلي",   size=9, bold=True, color=colors.white),
            self._p("الحالة",           size=9, bold=True, align=TA_CENTER, color=colors.white),
            self._p("الملاحظات",        size=9, bold=True, color=colors.white),
            self._p("صورة الجهاز",      size=9, bold=True, align=TA_CENTER, color=colors.white),
        ]
        data = [hdr]
        ROW_H = 3.2*cm

        for i, item in enumerate(self.inv):
            cond = item.get("condition","")
            cc = C_OK if cond in ["ممتاز","جيد جداً"] else \
                 C_SEC if cond == "جيد" else \
                 C_WARN if cond == "يحتاج صيانة" else C_ERR

            # primary photo
            ph = None
            for p in item.get("photos",[]):
                if p.get("is_primary"):
                    ph = p["data"]; break
            if ph is None and item.get("photos"):
                ph = item["photos"][0]["data"]

            if ph:
                t = thumb(ph, (113, 113))
                img_cell = RLImage(io.BytesIO(t), width=3*cm, height=3*cm)
            else:
                img_cell = self._p("لا توجد صورة", size=8, align=TA_CENTER, color=C_GRAY)

            row = [
                self._p(str(i+1),               size=9,  align=TA_CENTER),
                self._p(item.get("device_type","—"), size=9),
                self._p(item.get("brand","—"),   size=9),
                self._p(item.get("serial_number","—"), size=8),
                self._p(cond,                    size=8,  align=TA_CENTER, color=cc),
                self._p(item.get("notes","—"),   size=8),
                img_cell,
            ]
            data.append(row)

        tbl = Table(data, colWidths=CW, rowHeights=[0.7*cm] + [ROW_H]*len(self.inv),
                    repeatRows=1)
        tstyle = [
            ("BACKGROUND",    (0,0),(-1,0),  C_PRI),
            ("BOX",           (0,0),(-1,-1),  1.5, C_PRI),
            ("INNERGRID",     (0,0),(-1,-1), .5,  C_GRAY),
            ("ALIGN",         (0,0),(0,-1),  "CENTER"),
            ("ALIGN",         (4,1),(4,-1),  "CENTER"),
            ("ALIGN",         (6,1),(6,-1),  "CENTER"),
            ("ALIGN",         (0,0),(-1,0),  "RIGHT"),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("LEFTPADDING",   (0,0),(-1,-1), 4),
            ("RIGHTPADDING",  (0,0),(-1,-1), 4),
        ]
        for r in range(1, len(data)):
            bg = C_LITE if r % 2 == 0 else colors.white
            tstyle.append(("BACKGROUND",(0,r),(-1,r),bg))
        tbl.setStyle(TableStyle(tstyle))
        el.append(tbl)
        return el

    # ── DUPLICATES TABLE ──────────────────────────────────────────
    def _dups_table(self):
        el = [PageBreak(),
              self._p("تحذير: أرقام تسلسلية مكررة", size=13, bold=True, color=C_ERR),
              HRFlowable(width="100%", thickness=2, color=C_ERR),
              Spacer(1,.3*cm),
              self._p("يرجى مراجعة الأجهزة التالية والتحقق من أرقامها التسلسلية:", size=10),
              Spacer(1,.3*cm)]

        rows = [[self._p("الرقم التسلسلي",    size=10, bold=True, color=colors.white),
                  self._p("أرقام الصفوف المكررة",size=10, bold=True, color=colors.white)]]
        for sn, idx in self.dups:
            rows.append([self._p(sn, size=10, color=C_ERR),
                          self._p("، ".join(str(x) for x in idx), size=10)])

        tbl = Table(rows, colWidths=[9*cm, 6*cm])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0),  C_ERR),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.HexColor("#fdedec"), colors.white]),
            ("BOX",           (0,0),(-1,-1),  1.5, C_ERR),
            ("INNERGRID",     (0,0),(-1,-1), .5,  C_GRAY),
            ("ALIGN",         (0,0),(-1,-1), "RIGHT"),
            ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",    (0,0),(-1,-1), 7),
            ("BOTTOMPADDING", (0,0),(-1,-1), 7),
            ("LEFTPADDING",   (0,0),(-1,-1), 7),
            ("RIGHTPADDING",  (0,0),(-1,-1), 7),
        ]))
        el.append(tbl)
        return el

    # ── SIGNATURE ─────────────────────────────────────────────────
    def _sig(self):
        el = [Spacer(1,2*cm),
              HRFlowable(width="100%", thickness=.5, color=C_GRAY),
              Spacer(1,.5*cm)]
        rows = [
            [self._p("اعتُمد بمعرفة:",           size=10),
             self._p("المراجع / الرئيس المباشر:", size=10)],
            [self._p("أ.د / " + (self.prof or "............................"), size=11, bold=True),
             self._p("د / ....................................", size=11)],
            [Spacer(1,1.5*cm), Spacer(1,1.5*cm)],
            [self._p("التوقيع:  ___________________", size=10, color=C_GRAY),
             self._p("التوقيع:  ___________________", size=10, color=C_GRAY)],
            [self._p(f"التاريخ:  {self.date}", size=10),
             self._p("التاريخ:  _____  /  _____  /  _____", size=10)],
        ]
        tbl = Table(rows, colWidths=[8.5*cm, 8.5*cm])
        tbl.setStyle(TableStyle([
            ("ALIGN",  (0,0),(0,-1),"RIGHT"),
            ("ALIGN",  (1,0),(1,-1),"LEFT"),
            ("VALIGN", (0,0),(-1,-1),"MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1),6),
            ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]))
        el.extend([tbl, Spacer(1,.4*cm)])
        el.append(self._p(
            f"هذا الكشف صادر من إدارة {LAB_NAME} — {FAC_NAME} — {UNIV_NAME}",
            size=8, align=TA_CENTER, color=C_GRAY))
        return el

    # ── BUILD ─────────────────────────────────────────────────────
    def build(self) -> bytes:
        doc = SimpleDocTemplate(
            self.buf, pagesize=A4,
            rightMargin=2*cm, leftMargin=2*cm,
            topMargin=2.8*cm, bottomMargin=2*cm,
        )
        story  = self._cover()
        story += [PageBreak()]
        story += self._intro()
        story += [Spacer(1,.5*cm)]
        story += self._stats()
        story += self._inv_table()
        if self.dups:
            story += self._dups_table()
        story += self._sig()
        doc.build(story, onFirstPage=self._hf, onLaterPages=self._hf)
        return self.buf.getvalue()


# ================================================================
# 📝  DOCX REPORT
# ================================================================
class DOCXReport:
    def __init__(self, inv, prof, date):
        self.inv  = inv
        self.prof = prof
        self.date = date
        self.dups = get_duplicates()
        self.doc  = Document()

    def _rtl(self, para):
        pPr = para._p.get_or_add_pPr()
        b = OxmlElement("w:bidi"); b.set(qn("w:val"),"1"); pPr.append(b)
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def _run(self, para, text, bold=False, size=11, color=None):
        run = para.add_run(text)
        run.bold = bold
        run.font.name = "Cairo"
        run.font.size = Pt(size)
        run._element.rPr.rFonts.set(qn("w:hint"), "cs")
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        return run

    def _heading(self, text, level=1, color="1a5276"):
        p = self.doc.add_paragraph(); self._rtl(p)
        sz = {1:26,2:18,3:14,4:12}.get(level,11)
        self._run(p, text, bold=True, size=sz, color=color)
        return p

    def _para(self, text, size=11, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.RIGHT):
        p = self.doc.add_paragraph(); self._rtl(p); p.alignment = align
        self._run(p, text, bold=bold, size=size, color=color)
        return p

    def _shd(self, cell, fill_hex):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill_hex)
        shd.set(qn("w:val"),  "clear")
        cell._tc.get_or_add_tcPr().append(shd)

    def _setup(self):
        sec = self.doc.sections[0]
        sec.right_margin = Cm(2); sec.left_margin  = Cm(2)
        sec.top_margin   = Cm(2.5); sec.bottom_margin = Cm(2.5)
        # doc-level RTL
        for el in ["w:bidi","w:defaultTabStop"]:
            e = OxmlElement(el)
            try: self.doc.settings.element.append(e)
            except: pass
        # default font
        style = self.doc.styles["Normal"]
        style.font.name = "Cairo"; style.font.size = Pt(11)

    def _hdr_ftr(self):
        sec = self.doc.sections[0]
        hdr = sec.header
        # simple header para
        hp = hdr.paragraphs[0] if hdr.paragraphs else hdr.add_paragraph()
        self._rtl(hp)
        run = hp.add_run(f"{UNIV_NAME}  |  {FAC_NAME}  |  {LAB_NAME}  |  {self.date}")
        run.font.name = "Cairo"; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x1a,0x52,0x76); run.bold = True
        run._element.rPr.rFonts.set(qn("w:hint"), "cs")

        ftr = sec.footer
        fp = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run(f"{UNIV_NAME} — {FAC_NAME} — {LAB_NAME}")
        frun.font.name = "Cairo"; frun.font.size = Pt(8)
        frun.font.color.rgb = RGBColor(0x1a,0x52,0x76)
        frun._element.rPr.rFonts.set(qn("w:hint"), "cs")

    def _build_cover(self):
        lp = self.doc.add_paragraph(); lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lr = lp.add_run("[ شعار جامعة قناة السويس ]")
        lr.font.size = Pt(14); lr.font.color.rgb = RGBColor(0xaa,0xbb,0xcc)
        lr._element.rPr.rFonts.set(qn("w:hint"), "cs")

        self.doc.add_paragraph()
        self._heading(UNIV_NAME, 1, "1a5276")
        self._heading(FAC_NAME,  2, "2e86c1")
        self._heading(LAB_NAME,  3, "1c2833")
        self.doc.add_paragraph()
        div = self.doc.add_paragraph("─" * 42); div.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.doc.add_paragraph()
        self._heading(RPT_TITLE, 2, "1c2833")
        self.doc.add_paragraph()

        tbl = self.doc.add_table(3, 2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.style = "Table Grid"
        for r,(lbl,val) in enumerate([
            ("تاريخ الجرد:", self.date),
            ("المُعِد:", self.prof or "—"),
            ("إجمالي الأجهزة:", str(len(self.inv)))
        ]):
            row = tbl.rows[r]
            for c,txt in enumerate([lbl,val]):
                p = row.cells[c].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                self._rtl(p)
                run = p.add_run(txt); run.bold=(c==0)
                run.font.name="Cairo"; run.font.size=Pt(11)
                run._element.rPr.rFonts.set(qn("w:hint"), "cs")

    def _build_intro(self):
        self._heading("تمهيد", 3, "1a5276")
        self._para(
            f"بناءً على التوجيهات الصادرة من إدارة كلية الهندسة بجامعة قناة السويس، "
            f"أُعِدَّ هذا الكشف الشامل لأجهزة ومعدات {LAB_NAME} بتاريخ {self.date} "
            f"تحت إشراف {'أ.د / ' + self.prof if self.prof else 'المسؤول المختص'}."
        )
        self.doc.add_paragraph()

    def _build_stats(self):
        self._heading("الملخص الإحصائي", 3, "1a5276")
        brands = {}; conds = {}
        for item in self.inv:
            brands[item.get("brand","غير محدد")] = brands.get(item.get("brand","غير محدد"),0)+1
            conds[item.get("condition","غير محدد")] = conds.get(item.get("condition","غير محدد"),0)+1

        rows_data = [
            ("إجمالي عدد الأجهزة", str(len(self.inv))),
            ("عدد الماركات المختلفة", str(len(brands))),
            ("أجهزة تحتاج صيانة / معطلة",
             str(conds.get("يحتاج صيانة",0)+conds.get("معطل",0))),
        ]
        if self.dups:
            rows_data.append(("أرقام تسلسلية مكررة", str(len(self.dups))))

        tbl = self.doc.add_table(len(rows_data)+1, 2)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT
        for c, txt in enumerate(["البيان","القيمة"]):
            cell = tbl.rows[0].cells[c]
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            self._rtl(p)
            run = p.add_run(txt); run.bold = True
            run.font.name = "Cairo"; run.font.size = Pt(11)
            run._element.rPr.rFonts.set(qn("w:hint"), "cs")
            self._shd(cell, "1a5276")
        for r,(lbl,val) in enumerate(rows_data):
            row = tbl.rows[r+1]
            for c,txt in enumerate([lbl,val]):
                p = row.cells[c].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                self._rtl(p)
                run = p.add_run(txt); run.font.name = "Cairo"; run.font.size = Pt(11)
                run._element.rPr.rFonts.set(qn("w:hint"), "cs")
        self.doc.add_paragraph()

    def _build_inv_table(self):
        self._heading("كشف الأجهزة والمعدات", 2, "1a5276")
        hdrs = ["م","نوع الجهاز","الماركة / المصنّع","الرقم التسلسلي","الحالة","الملاحظات","صورة"]
        CW = [Cm(1),Cm(3.5),Cm(2.5),Cm(2.5),Cm(2),Cm(3),Cm(2.5)]
        tbl = self.doc.add_table(len(self.inv)+1, 7)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT

        # header
        for c,(h,w) in enumerate(zip(hdrs,CW)):
            cell = tbl.rows[0].cells[c]
            cell.width = w
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self._rtl(p)
            run = p.add_run(h); run.bold = True
            run.font.name = "Cairo"; run.font.size = Pt(10)
            run._element.rPr.rFonts.set(qn("w:hint"), "cs")
            self._shd(cell, "1a5276")

        # rows
        for r, item in enumerate(self.inv):
            row = tbl.rows[r+1]
            # set row height ~3cm
            trPr = row._tr.get_or_add_trPr()
            trH = OxmlElement("w:trHeight")
            trH.set(qn("w:val"),"1700"); trH.set(qn("w:hRule"),"exact")
            trPr.append(trH)

            vals = [str(r+1), item.get("device_type","—"), item.get("brand","—"),
                    item.get("serial_number","—"), item.get("condition","—"),
                    item.get("notes","—"), None]

            for c, val in enumerate(vals):
                cell = row.cells[c]
                cell.width = CW[c]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c in (0,6) else WD_ALIGN_PARAGRAPH.RIGHT
                self._rtl(p)

                if c == 6:
                    ph = None
                    for ph_item in item.get("photos",[]):
                        if ph_item.get("is_primary"):
                            ph = ph_item["data"]; break
                    if ph is None and item.get("photos"):
                        ph = item["photos"][0]["data"]
                    if ph:
                        t = thumb(ph,(150,150))
                        try:
                            run = p.add_run()
                            run.add_picture(io.BytesIO(t), width=Cm(2.5), height=Cm(2.5))
                        except Exception:
                            run = p.add_run("—"); run.font.name="Cairo"
                            run._element.rPr.rFonts.set(qn("w:hint"), "cs")
                    else:
                        run = p.add_run("—"); run.font.name="Cairo"
                        run._element.rPr.rFonts.set(qn("w:hint"), "cs")
                else:
                    run = p.add_run(str(val) if val else "—")
                    run.font.name = "Cairo"; run.font.size = Pt(9)
                    run._element.rPr.rFonts.set(qn("w:hint"), "cs")

                if r % 2 == 0:
                    self._shd(cell, "eaf0fb")

        self.doc.add_paragraph()

    def _build_dups(self):
        self._heading("أجهزة بأرقام تسلسلية مكررة", 3, "c0392b")
        self._para("يرجى مراجعة الأجهزة التالية والتحقق من أرقامها:")
        tbl = self.doc.add_table(len(self.dups)+1, 2)
        tbl.style = "Table Grid"
        for c, txt in enumerate(["الرقم التسلسلي","أرقام الصفوف"]):
            cell = tbl.rows[0].cells[c]
            p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            self._rtl(p)
            run = p.add_run(txt); run.bold = True
            run.font.name="Cairo"; run.font.size=Pt(11)
            run._element.rPr.rFonts.set(qn("w:hint"), "cs")
            self._shd(cell,"c0392b")
        for r,(sn,idx) in enumerate(self.dups):
            row = tbl.rows[r+1]
            for c,txt in enumerate([sn,"، ".join(str(x) for x in idx)]):
                p = row.cells[c].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                self._rtl(p)
                run = p.add_run(txt); run.font.name="Cairo"; run.font.size=Pt(10)
                run._element.rPr.rFonts.set(qn("w:hint"), "cs")
        self.doc.add_paragraph()

    def _build_sig(self):
        self.doc.add_paragraph()
        tbl = self.doc.add_table(4,2)
        sig_rows = [
            ("اعتُمد بمعرفة:","المراجع / الرئيس المباشر:"),
            (f"أ.د / {self.prof or '......................'}","د / ......................"),
            ("",""),
            ("التوقيع: _______________","التوقيع: _______________"),
        ]
        for r,(l,rr) in enumerate(sig_rows):
            row = tbl.rows[r]
            for c,txt in enumerate([l,rr]):
                p = row.cells[c].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if c==0 else WD_ALIGN_PARAGRAPH.LEFT
                self._rtl(p)
                run = p.add_run(txt); run.font.name="Cairo"; run.font.size=Pt(11)
                run.bold = (r==1); run._element.rPr.rFonts.set(qn("w:hint"), "cs")
        self.doc.add_paragraph()
        fp = self.doc.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER; self._rtl(fp)
        run = fp.add_run(f"هذا الكشف صادر من إدارة {LAB_NAME} — {FAC_NAME} — {UNIV_NAME}")
        run.font.name="Cairo"; run.font.size=Pt(9)
        run.font.color.rgb = RGBColor(0x99,0x99,0x99)
        run._element.rPr.rFonts.set(qn("w:hint"), "cs")

    def build(self) -> bytes:
        self._setup()
        self._hdr_ftr()
        self._build_cover()
        self.doc.add_page_break()
        self._build_intro()
        self._build_stats()
        self.doc.add_page_break()
        self._build_inv_table()
        if self.dups:
            self.doc.add_page_break()
            self._build_dups()
        self._build_sig()
        buf = io.BytesIO(); self.doc.save(buf)
        return buf.getvalue()


# ================================================================
# 📱  UI HELPERS
# ================================================================
def nav():
    items = [
        ("📷", "إضافة جهاز", "capture"),
        ("📋", f"القائمة  ({len(st.session_state.inventory)})", "list"),
        ("📤", "تصدير التقرير", "export"),
    ]
    cols = st.columns(3)
    for col, (icon, label, pg) in zip(cols, items):
        active = "✅ " if st.session_state.page == pg else ""
        with col:
            if st.button(f"{active}{icon}  {label}", key=f"nav_{pg}"):
                st.session_state.page = pg
                st.rerun()


def add_photo_to_staging(data: bytes, is_primary: bool, label: str):
    h = hash(data)
    if h not in st.session_state.photo_hashes:
        st.session_state.photo_hashes.add(h)
        st.session_state.captured_photos.append({
            "id": uuid.uuid4().hex[:8],
            "data": data,
            "is_primary": is_primary,
            "label": label,
        })


# ================================================================
# 📷  PAGE: CAPTURE
# ================================================================
def page_capture():

    # ── Settings ─────────────────────────────────────────────────
    with st.expander("⚙️  إعدادات التقرير والـ API", expanded=not st.session_state.gemini_key):
        st.session_state.gemini_key = st.text_input(
            "🔑  Gemini API Key",
            value=st.session_state.gemini_key,
            type="password",
            help="احصل على مفتاحك من: aistudio.google.com/app/apikey"
        )
        c1, c2 = st.columns(2)
        st.session_state.professor_name = c1.text_input(
            "👤  اسم الأستاذ المسؤول",
            value=st.session_state.professor_name,
            placeholder="أ.د / محمد أحمد..."
        )
        st.session_state.report_date = c2.text_input(
            "📅  تاريخ الجرد",
            value=st.session_state.report_date
        )
    st.markdown("---")

    # ── STEP 1: Upload ────────────────────────────────────────────
    st.markdown('<div class="step-label">📸  الخطوة 1 — التقاط صور الجهاز</div>',
                unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        prim_file = st.file_uploader(
            "⭐  الصورة الرئيسية (تظهر في التقرير)",
            type=["jpg","jpeg","png","webp"],
            accept_multiple_files=False,
            key="up_primary"
        )
    with c2:
        ref_files = st.file_uploader(
            "📎  صور مرجعية (لوحة البيانات، أسفل الجهاز...)",
            type=["jpg","jpeg","png","webp"],
            accept_multiple_files=True,
            key="up_refs",
            help="تُستخدم للاستخراج الذكي فقط — لا تظهر في التقرير"
        )

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("📥  تثبيت الصور في القائمة"):
            if prim_file:
                add_photo_to_staging(resize_img(prim_file.read()), True, "الصورة الرئيسية")
            if ref_files:
                for i, rf in enumerate(ref_files):
                    add_photo_to_staging(resize_img(rf.read()), False, f"مرجعية {i+1}")
            st.rerun()
    with col_b:
        if st.button("🗑️  مسح جميع الصور الحالية"):
            st.session_state.captured_photos = []
            st.session_state.photo_hashes = set()
            st.rerun()

    # display staged photos
    shots = st.session_state.captured_photos
    if shots:
        st.markdown(f"**معاينة الصور المُثبَّتة ({len(shots)} صورة):**")
        cols = st.columns(min(len(shots), 3))
        to_del = []
        for i, ph in enumerate(shots):
            with cols[i % 3]:
                badge_cls = "badge-primary" if ph["is_primary"] else "badge-ref"
                badge_lbl = ph["label"]
                b64 = b64img(ph["data"])
                st.markdown(f"""
                <div class="photo-card">
                  <span class="{badge_cls}">{badge_lbl}</span><br><br>
                  <img src="data:image/jpeg;base64,{b64}"
                       style="width:100%;border-radius:8px;max-height:155px;object-fit:cover;" />
                </div>""", unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_{ph['id']}"):
                    to_del.append(ph["id"])
        if to_del:
            st.session_state.photo_hashes = {
                hash(p["data"]) for p in shots if p["id"] not in to_del}
            st.session_state.captured_photos = [
                p for p in shots if p["id"] not in to_del]
            st.rerun()

    st.markdown("---")

    # ── STEP 2: AI Extraction ─────────────────────────────────────
    st.markdown('<div class="step-label">🤖  الخطوة 2 — الاستخراج الذكي بـ Gemini AI</div>',
                unsafe_allow_html=True)

    ca, cb = st.columns(2)
    with ca:
        if st.button("🚀  استخراج البيانات بالذكاء الاصطناعي",
                     disabled=not st.session_state.captured_photos):
            if not st.session_state.gemini_key:
                st.error("⚠️  أدخل Gemini API Key في الإعدادات أولاً")
            else:
                with st.spinner("🔍  جاري التحليل..."):
                    imgs = [p["data"] for p in st.session_state.captured_photos]
                    res  = extract_device_info(imgs, st.session_state.gemini_key)
                    st.session_state.ai_result = res
                st.rerun()
    with cb:
        if st.button("✏️  إدخال يدوي بدون AI"):
            st.session_state.ai_result = {
                "device_type":"","brand":"","serial_number":"",
                "condition":"جيد","notes":"","confidence":"manual"}
            st.rerun()

    ai = st.session_state.ai_result or {}
    if "confidence" in ai:
        conf_map = {"high":("🟢","عالية"),"medium":("🟡","متوسطة"),
                    "low":("🔴","منخفضة"),"manual":("✏️","يدوي")}
        ic, lb = conf_map.get(ai["confidence"],("⚪","غير محدد"))
        st.info(f"{ic}  دقة استخراج الذكاء الاصطناعي: **{lb}**")

    if "error" in ai:
        st.error(f"❌  {ai['error']}")

    st.markdown("---")

    # ── STEP 3: Form ──────────────────────────────────────────────
    if ai and "device_type" in ai:
        st.markdown('<div class="step-label">📝  الخطوة 3 — مراجعة البيانات وتأكيدها</div>',
                    unsafe_allow_html=True)

        fc1, fc2 = st.columns(2)
        device_type = fc1.text_input("🔧  نوع الجهاز",   value=ai.get("device_type",""))
        brand       = fc2.text_input("🏭  الماركة / المصنّع", value=ai.get("brand",""))
        serial      = fc1.text_input("🔢  الرقم التسلسلي",   value=ai.get("serial_number",""))
        cond_idx    = STATUS_OPT.index(ai.get("condition","جيد")) \
                        if ai.get("condition","") in STATUS_OPT else 2
        condition   = fc2.selectbox("📊  الحالة", STATUS_OPT, index=cond_idx)
        notes       = st.text_area("📒  ملاحظات الأستاذ", value=ai.get("notes",""),
                                   height=90, placeholder="أدخل أي ملاحظات...")

        # duplicate check
        dup = is_duplicate(serial)
        if dup and serial.strip():
            st.markdown(f"""
            <div class="dup-warning">
              ⚠️ <strong>تحذير: الرقم التسلسلي مكرر!</strong><br>
              الرقم <code>{serial}</code> موجود مسبقاً في قائمة الجرد — يرجى التحقق.
            </div>
            """, unsafe_allow_html=True)
            st.markdown("""<script>
                if(navigator&&navigator.vibrate)navigator.vibrate([300,100,300,100,300]);
            </script>""", unsafe_allow_html=True)

        st.markdown("---")
        sb1, sb2 = st.columns([3,1])
        with sb1:
            if st.button("✅  إضافة إلى قائمة الجرد", type="primary"):
                st.session_state.inventory.append({
                    "id": uuid.uuid4().hex,
                    "device_type":   device_type,
                    "brand":         brand,
                    "serial_number": serial,
                    "condition":     condition,
                    "notes":         notes,
                    "photos":        [{"id":p["id"],"data":p["data"],
                                       "is_primary":p["is_primary"],"label":p["label"]}
                                      for p in st.session_state.captured_photos],
                    "is_duplicate":  dup,
                    "added_at":      datetime.datetime.now().strftime("%H:%M"),
                })
                st.session_state.captured_photos = []
                st.session_state.photo_hashes    = set()
                st.session_state.ai_result       = None
                st.markdown('<div class="success-flash">✅  تمت الإضافة بنجاح!</div>',
                            unsafe_allow_html=True)
                st.rerun()
        with sb2:
            if st.button("↩️  إعادة تعيين"):
                st.session_state.captured_photos = []
                st.session_state.photo_hashes    = set()
                st.session_state.ai_result       = None
                st.rerun()


# ================================================================
# 📋  PAGE: LIST
# ================================================================
def page_list():
    inv = st.session_state.inventory
    if not inv:
        st.info("📭  قائمة الجرد فارغة — أضف أجهزة من صفحة 'إضافة جهاز'.")
        return

    total = len(inv)
    dups  = len(get_duplicates())
    brands= len(set(i.get("brand","") for i in inv))
    maint = sum(1 for i in inv if i.get("condition") in ["يحتاج صيانة","معطل"])

    c1,c2,c3,c4 = st.columns(4)
    c1.metric("📦  إجمالي الأجهزة", total)
    c2.metric("🏭  الماركات",        brands)
    c3.metric("⚠️  تحتاج صيانة",    maint)
    c4.metric("🔄  مكررة",           dups)

    if dups:
        st.error(f"⚠️  يوجد **{dups}** رقم تسلسلي مكرر — راجع الأجهزة بعناية.")

    st.markdown("---")

    for i, item in enumerate(inv):
        dup = item.get("is_duplicate") or is_duplicate(item.get("serial_number",""), i)
        ph  = next((p["data"] for p in item.get("photos",[]) if p.get("is_primary")), None)
        if ph is None and item.get("photos"):
            ph = item["photos"][0]["data"]

        c_img, c_info, c_del = st.columns([1,4,1])
        with c_img:
            if ph:
                st.image(thumb(ph,(120,120)), use_container_width=True)
            else:
                st.markdown("🖼️")
        with c_info:
            dup_badge = "  🔴 **مكرر**" if dup else ""
            st.markdown(f"**{i+1}. {item.get('device_type','غير محدد')}**{dup_badge}")
            st.markdown(f"🏭 `{item.get('brand','—')}`  |  🔢 `{item.get('serial_number','—')}`  |  📊 {item.get('condition','—')}  |  🕐 {item.get('added_at','')}")
            if item.get("notes"):
                st.caption(f"📝 {item['notes'][:100]}")
        with c_del:
            if st.button("🗑️", key=f"del_inv_{i}"):
                st.session_state.inventory.pop(i)
                st.rerun()

        st.markdown("---")


# ================================================================
# 📤  PAGE: EXPORT
# ================================================================
def page_export():
    inv = st.session_state.inventory
    if not inv:
        st.warning("⚠️  أضف أجهزة إلى قائمة الجرد أولاً.")
        return

    st.markdown("### 📋  ملخص التقرير المزمع تصديره")
    dups = get_duplicates()
    st.markdown(f"""
    <div style="background:#eaf0fb;border-radius:12px;padding:1.1rem;direction:rtl;
                border-right:4px solid #1a5276;margin-bottom:1rem;">
      <strong>👤</strong> المُعِد: {st.session_state.professor_name or '<em>غير محدد</em>'}<br>
      <strong>📅</strong> التاريخ: {st.session_state.report_date}<br>
      <strong>📦</strong> عدد الأجهزة: {len(inv)}<br>
      <strong>⚠️</strong> أرقام مكررة: {len(dups)}
    </div>""", unsafe_allow_html=True)

    fonts_ok = setup_fonts()
    if not fonts_ok:
        st.warning("⚠️  فشل تحميل الخطوط — تأكد من الاتصال بالإنترنت وأعد المحاولة.")

    st.markdown("---")
    col_pdf, col_doc = st.columns(2)

    # ── PDF ──────────────────────────────────────────────────────
    with col_pdf:
        st.markdown("#### 📄  ملف PDF")
        st.caption("للأرشفة الرسمية والطباعة")
        if st.button("🖨️  توليد PDF", disabled=not fonts_ok, type="primary"):
            with st.spinner("⏳  جاري إنشاء PDF..."):
                try:
                    pdf = PDFReport(inv, st.session_state.professor_name,
                                    st.session_state.report_date).build()
                    fn = f"جرد_{LAB_NAME}_{st.session_state.report_date.replace('/','-')}.pdf"
                    st.download_button("⬇️  تحميل PDF", pdf, fn, "application/pdf")
                    st.success("✅  PDF جاهز!")
                except Exception as e:
                    st.error(f"❌  {e}")

    # ── DOCX ─────────────────────────────────────────────────────
    with col_doc:
        st.markdown("#### 📝  ملف Word (DOCX)")
        st.caption("للتعديل والمراجعة")
        if st.button("📄  توليد DOCX"):
            with st.spinner("⏳  جاري إنشاء DOCX..."):
                try:
                    docx = DOCXReport(inv, st.session_state.professor_name,
                                      st.session_state.report_date).build()
                    fn = f"جرد_{LAB_NAME}_{st.session_state.report_date.replace('/','-')}.docx"
                    MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    st.download_button("⬇️  تحميل DOCX", docx, fn, MIME)
                    st.success("✅  DOCX جاهز!")
                except Exception as e:
                    st.error(f"❌  {e}")

    # ── clear ─────────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("⚠️  خيارات متقدمة"):
        if st.button("🗑️  مسح قائمة الجرد بالكامل", type="secondary"):
            st.session_state.inventory = []
            st.session_state.captured_photos = []
            st.session_state.photo_hashes    = set()
            st.session_state.ai_result       = None
            st.success("✅  تم مسح القائمة.")
            st.rerun()


# ================================================================
# 🚀  MAIN
# ================================================================
def main():
    st.markdown(f"""
    <div class="app-header">
      <h1>🔭  نظام جرد معمل المساحة</h1>
      <p>{UNIV_NAME}  |  {FAC_NAME}  |  مدعوم بـ Gemini AI</p>
    </div>""", unsafe_allow_html=True)

    nav()
    st.markdown("")

    pg = st.session_state.page
    if   pg == "capture": page_capture()
    elif pg == "list":    page_list()
    elif pg == "export":  page_export()


if __name__ == "__main__":
    main()
