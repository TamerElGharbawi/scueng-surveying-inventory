# ================================================================
#  نظام جرد معمل المساحة | Surveying Lab Inventory
#  جامعة قناة السويس – كلية الهندسة
#  Version 5.0 — DOCX Only | RTL Fixed | Single Image | Logo Support
# ================================================================

import streamlit as st
from PIL import Image, ImageEnhance, ImageOps, ExifTags
import io, base64, os, uuid, datetime
from pathlib import Path

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
UNIV_NAME  = "جامعة قناة السويس"
FAC_NAME   = "كلية الهندسة"
LAB_NAME   = "معمل المساحة"
RPT_TITLE  = "كشف جرد أجهزة ومعدات المعمل"
STATUS_OPT = ["ممتاز", "جيد جداً", "جيد", "يحتاج صيانة", "معطل"]

LOGO_PATH  = Path("logo.png")   # place logo.png in the same folder as app.py


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
html,body,.stApp{direction:rtl!important;font-family:'Cairo',sans-serif!important;}
.main .block-container{padding:1rem .75rem!important;max-width:100%!important;}
h1,h2,h3,h4{font-family:'Cairo',sans-serif!important;text-align:right!important;color:#1a5276!important;}
.stButton>button{width:100%!important;min-height:3rem!important;font-size:1rem!important;
  font-family:'Cairo',sans-serif!important;font-weight:700!important;border-radius:12px!important;transition:all .2s!important;}
.stButton>button:hover{transform:translateY(-2px)!important;box-shadow:0 4px 14px rgba(0,0,0,.18)!important;}
.stTextInput input,.stTextArea textarea{direction:rtl!important;text-align:right!important;
  font-family:'Cairo',sans-serif!important;font-size:1rem!important;border-radius:10px!important;}
label,.stSelectbox label,.stTextArea label{font-family:'Cairo',sans-serif!important;
  font-size:1rem!important;font-weight:600!important;color:#1a5276!important;text-align:right!important;}
[data-testid="metric-container"]{direction:rtl!important;text-align:right!important;
  background:#eaf0fb!important;border-radius:12px!important;padding:.75rem!important;
  border-right:4px solid #1a5276!important;}
[data-testid="metric-container"] label,
[data-testid="metric-container"] [data-testid="metric-value"]{
  font-family:'Cairo',sans-serif!important;color:#1a5276!important;}
.stExpander{border:1.5px solid #2e86c1!important;border-radius:12px!important;direction:rtl!important;}
.stExpander summary{font-family:'Cairo',sans-serif!important;font-weight:700!important;}
.stAlert{direction:rtl!important;font-family:'Cairo',sans-serif!important;border-radius:10px!important;}
.stTabs [data-baseweb="tab"]{font-family:'Cairo',sans-serif!important;font-size:1rem!important;font-weight:600!important;}
.stTabs [data-baseweb="tab-list"]{direction:rtl!important;}
#MainMenu,footer,header{visibility:hidden;}
::-webkit-scrollbar{width:5px;}
::-webkit-scrollbar-thumb{background:#2e86c1;border-radius:3px;}

.app-header{background:linear-gradient(135deg,#1a5276,#2e86c1);color:white;
  padding:1rem 1.5rem;border-radius:14px;text-align:center;margin-bottom:1.2rem;}
.app-header h1{color:white!important;font-size:1.35rem!important;margin:0!important;}
.app-header p{color:#aed6f1!important;margin:.2rem 0 0!important;font-size:.82rem!important;}
.step-label{background:#eaf0fb;border-right:4px solid #1a5276;padding:.5rem 1rem;
  border-radius:0 8px 8px 0;font-family:'Cairo',sans-serif;font-weight:700;
  color:#1a5276;margin-bottom:.75rem;}
.photo-card{border:2px solid #2e86c1;border-radius:12px;padding:8px;
  text-align:center;background:white;direction:rtl;margin-bottom:6px;}
.badge-primary{background:#1a5276;color:white;padding:2px 10px;border-radius:20px;
  font-size:.72rem;font-family:'Cairo',sans-serif;font-weight:700;}
.dup-warning{background:#fdedec;border:2px solid #c0392b;border-radius:12px;
  padding:1rem;direction:rtl;animation:dupPulse 1.4s infinite;}
@keyframes dupPulse{
  0%{box-shadow:0 0 0 0 rgba(192,57,43,.5)}
  70%{box-shadow:0 0 0 10px rgba(192,57,43,0)}
  100%{box-shadow:0 0 0 0 rgba(192,57,43,0)}}
.success-flash{background:linear-gradient(135deg,#1e8449,#27ae60);color:white;
  padding:1rem;border-radius:12px;text-align:center;font-family:'Cairo',sans-serif;
  font-weight:700;font-size:1.05rem;}
.editor-box{background:#f0f8ff;border:1.5px solid #2e86c1;border-radius:14px;
  padding:1rem;margin:8px 0;}
.info-box{background:#eaf0fb;border-right:4px solid #1a5276;border-radius:0 12px 12px 0;
  padding:.8rem 1rem;direction:rtl;font-family:'Cairo',sans-serif;margin-bottom:.75rem;}
</style>
""", unsafe_allow_html=True)


# ================================================================
# 🗃️  SESSION STATE
# ================================================================
def init_state():
    defs = dict(
        inventory=[],
        captured_photos=[],
        photo_hashes=set(),
        page="capture",
        professor_name="",
        report_date=datetime.date.today().strftime("%Y/%m/%d"),
        edit_photo_idx=None,
        form_counter=0,          # incremented on every "add" to force widget reset
    )
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ================================================================
# 🖼️  IMAGE PROCESSING
# ================================================================
def fix_exif_rotation(img: Image.Image) -> Image.Image:
    try:
        exif = img._getexif()
        if exif is None:
            return img
        orient_tag = next(
            (k for k, v in ExifTags.TAGS.items() if v == "Orientation"), None)
        if orient_tag and orient_tag in exif:
            o = exif[orient_tag]
            if   o == 3: img = img.rotate(180, expand=True)
            elif o == 6: img = img.rotate(270, expand=True)
            elif o == 8: img = img.rotate(90,  expand=True)
            elif o == 2: img = img.transpose(Image.FLIP_LEFT_RIGHT)
            elif o == 4: img = img.transpose(Image.FLIP_TOP_BOTTOM)
            elif o == 5: img = img.transpose(Image.FLIP_LEFT_RIGHT).rotate(90,  expand=True)
            elif o == 7: img = img.transpose(Image.FLIP_LEFT_RIGHT).rotate(270, expand=True)
    except Exception:
        pass
    return img


def auto_enhance(img: Image.Image) -> Image.Image:
    img = ImageOps.autocontrast(img, cutoff=1)
    img = ImageEnhance.Sharpness(img).enhance(1.4)
    img = ImageEnhance.Contrast(img).enhance(1.12)
    return img


def process_image(data: bytes,
                  rotate_deg: int = 0,
                  crop_box: tuple = (0.0, 0.0, 1.0, 1.0),
                  enhance: bool = True) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img = fix_exif_rotation(img)
    if rotate_deg:
        img = img.rotate(-rotate_deg, expand=True)
    l, t, r, b = crop_box
    if (l, t, r, b) != (0.0, 0.0, 1.0, 1.0):
        W, H = img.size
        img = img.crop((int(l*W), int(t*H), int(r*W), int(b*H)))
    if enhance:
        img = auto_enhance(img)
    if max(img.size) > 1400:
        img.thumbnail((1400, 1400), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=88)
    return buf.getvalue()


def resize_img(data: bytes, mx: int = 1200) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img = fix_exif_rotation(img)
    if max(img.size) > mx:
        img.thumbnail((mx, mx), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    return buf.getvalue()


def thumb(data: bytes, sz=(160, 160)) -> bytes:
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.thumbnail(sz, Image.LANCZOS)
    canvas = Image.new("RGB", sz, (255, 255, 255))
    canvas.paste(img, ((sz[0]-img.size[0])//2, (sz[1]-img.size[1])//2))
    buf = io.BytesIO()
    canvas.save(buf, "JPEG", quality=85)
    return buf.getvalue()


def b64img(data: bytes) -> str:
    return base64.b64encode(data).decode()


# ================================================================
# 🔍  DUPLICATE DETECTION
# ================================================================
IGNORED_SERIALS = {"", "غير مقروء", "غير محدد", "-", "—"}

def is_duplicate(serial: str, exclude_idx: int = -1) -> bool:
    """Return True only if the same serial already exists in inventory
    (ignoring the entry at exclude_idx, used when re-checking saved items)."""
    s = serial.strip().upper() if serial else ""
    if not s or s in {v.upper() for v in IGNORED_SERIALS}:
        return False
    for i, item in enumerate(st.session_state.inventory):
        if i == exclude_idx:
            continue
        if item.get("serial_number", "").strip().upper() == s:
            return True
    return False


def get_duplicates() -> list:
    seen: dict = {}
    for i, item in enumerate(st.session_state.inventory):
        sn = item.get("serial_number", "").strip().upper()
        if sn and sn not in {v.upper() for v in IGNORED_SERIALS}:
            seen.setdefault(sn, []).append(i + 1)
    return [(sn, idx) for sn, idx in seen.items() if len(idx) > 1]


# ================================================================
# 📝  DOCX HELPERS
# ================================================================
def _safe_set_hint(run):
    """Safely set w:hint=cs on run's rFonts, creating elements if missing."""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:hint"), "cs")


def _safe_run(para, text, bold=False, size=11, color_hex=None):
    run = para.add_run(text)
    run.bold = bold
    run.font.name = "Amiri"
    run.font.size = Pt(size)
    if color_hex:
        run.font.color.rgb = RGBColor.from_string(color_hex)
    _safe_set_hint(run)
    return run


def _rtl_para(para):
    p = para._p
    pPr = p.get_or_add_pPr()

    # Remove existing bidi if any (avoid duplication)
    for el in pPr.findall(qn("w:bidi")):
        pPr.remove(el)

    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)

    # Ensure right alignment
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def _set_cell_rtl(cell):
    """Force RTL inside table cell (fixes mixed direction issue in Word)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Remove existing bidi if present
    for el in tcPr.findall(qn("w:bidi")):
        tcPr.remove(el)

    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    tcPr.append(bidi)
    
def _cell_shd(cell, fill_hex):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill_hex)
    shd.set(qn("w:val"), "clear")
    cell._tc.get_or_add_tcPr().append(shd)


def _set_table_rtl(tbl):
    """Force table to render right-to-left (columns order reversed visually)."""

    tbl_elem = tbl._tbl

    # --- FIX FOR OLD python-docx ---
    # Instead of: get_or_add_tblPr()
    tblPr = tbl_elem.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_elem.insert(0, tblPr)

    # Avoid duplicate bidiVisual
    existing = tblPr.findall(qn("w:bidiVisual"))
    if not existing:
        bidi = OxmlElement("w:bidiVisual")
        bidi.set(qn("w:val"), "1")
        tblPr.append(bidi)


def _set_col_widths(tbl, widths_cm):
    """Set individual column widths on a table."""
    tbl_elem = tbl._tbl

    # Ensure tblGrid exists safely
    tblGrid = tbl_elem.tblGrid
    if tblGrid is None:
        tblGrid = OxmlElement("w:tblGrid")
        tbl_elem.insert(1, tblGrid)

    # Clear existing grid
    for child in list(tblGrid):
        tblGrid.remove(child)

    for w in widths_cm:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(int(w * 567)))  # cm → twips
        tblGrid.append(gridCol)


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

    def _heading(self, text, size=14, color="1a5276"):
        p = self.doc.add_paragraph()
        _rtl_para(p)
        _safe_run(p, text, bold=True, size=size, color_hex=color)

    def _para(self, text, size=11, bold=False, color_hex=None):
        p = self.doc.add_paragraph()
        _rtl_para(p)
        _safe_run(p, text, bold=bold, size=size, color_hex=color_hex)

    def _hdr_run(self, para, text, size=9, bold=False, color_hex="1a5276"):
        run = para.add_run(text)
        run.bold = bold
        run.font.name = "Amiri"
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color_hex)
        _safe_set_hint(run)

    def build(self) -> bytes:
        doc = self.doc

        # ── document setup ─────────────────────────────────────
        sec = doc.sections[0]
        sec.right_margin  = Cm(2)
        sec.left_margin   = Cm(2)
        sec.top_margin    = Cm(3)
        sec.bottom_margin = Cm(2)
        doc.styles["Normal"].font.name = "Amiri"
        doc.styles["Normal"].font.size = Pt(11)

        # ── header ─────────────────────────────────────────────
        hp = (sec.header.paragraphs[0]
               if sec.header.paragraphs else sec.header.add_paragraph())
        _rtl_para(hp)

        # Try to insert logo in header
        if LOGO_PATH.exists():
            try:
                run_logo = hp.add_run()
                run_logo.add_picture(str(LOGO_PATH), height=Cm(1.0))
                hp.add_run("  ")   # small spacer
            except Exception:
                pass

        self._hdr_run(
            hp,
            f"{UNIV_NAME}  |  {FAC_NAME}  |  {LAB_NAME}  |  {self.date}",
            size=9, bold=True)

        # ── footer ─────────────────────────────────────────────
        fp = (sec.footer.paragraphs[0]
               if sec.footer.paragraphs else sec.footer.add_paragraph())
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._hdr_run(fp, f"{UNIV_NAME} — {FAC_NAME} — {LAB_NAME}", size=8)

        # ── COVER ──────────────────────────────────────────────
        # Logo (centred)
        lp = doc.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if LOGO_PATH.exists():
            try:
                lr = lp.add_run()
                lr.add_picture(str(LOGO_PATH), height=Cm(3.5))
            except Exception:
                lr = lp.add_run("[ شعار جامعة قناة السويس ]")
                lr.font.size = Pt(14)
                lr.font.color.rgb = RGBColor(0xaa, 0xbb, 0xcc)
                _safe_set_hint(lr)
        else:
            lr = lp.add_run("[ شعار جامعة قناة السويس ]")
            lr.font.size = Pt(14)
            lr.font.color.rgb = RGBColor(0xaa, 0xbb, 0xcc)
            _safe_set_hint(lr)

        doc.add_paragraph()
        self._heading(UNIV_NAME, 22, "1a5276")
        self._heading(FAC_NAME,  16, "2e86c1")
        self._heading(LAB_NAME,  13, "1c2833")
        doc.add_paragraph()
        div = doc.add_paragraph("─" * 42)
        div.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        self._heading(RPT_TITLE, 16, "1c2833")
        doc.add_paragraph()

        # cover info table (RTL)
        ct = doc.add_table(3, 2)
        ct.style = "Table Grid"
        ct.alignment = WD_TABLE_ALIGNMENT.RIGHT
        _set_table_rtl(ct)
        for r, (lbl, val) in enumerate([
            ("تاريخ الجرد:", self.date),
            ("المُعِد:", self.prof or "—"),
            ("إجمالي الأجهزة:", str(len(self.inv))),
        ]):
            for c, txt in enumerate([lbl, val]):
                cell = ct.rows[r].cells[c]
                _set_cell_rtl(cell)
                p = cell.paragraphs[0]
                
                _rtl_para(p)
                _safe_run(p, txt, bold=(c == 0), size=11)

        doc.add_page_break()

        # ── STATS ──────────────────────────────────────────────
        self._heading("الملخص الإحصائي", 13, "1a5276")
        brands = {}; conds = {}
        for item in self.inv:
            b = item.get("brand", "غير محدد")
            brands[b] = brands.get(b, 0) + 1
            c = item.get("condition", "غير محدد")
            conds[c] = conds.get(c, 0) + 1
        rows_data = [
            ("إجمالي عدد الأجهزة",  str(len(self.inv))),
            ("عدد الماركات",         str(len(brands))),
            ("تحتاج صيانة / معطلة", str(conds.get("يحتاج صيانة", 0) + conds.get("معطل", 0))),
        ]
        if self.dups:
            rows_data.append(("أرقام مكررة", str(len(self.dups))))

        st2 = doc.add_table(len(rows_data) + 1, 2)
        st2.style = "Table Grid"
        st2.alignment = WD_TABLE_ALIGNMENT.RIGHT
        _set_table_rtl(st2)
        for c, txt in enumerate(["البيان", "القيمة"]):
            cell = st2.rows[0].cells[c]
            _set_cell_rtl(cell)
            p = cell.paragraphs[0]
            _rtl_para(p)
            _safe_run(p, txt, bold=True, size=11, color_hex="ffffff")
            _cell_shd(cell, "1a5276")
        for r, (lbl, val) in enumerate(rows_data):
            for c, txt in enumerate([lbl, val]):

                cell = st2.rows[r + 1].cells[c]
                _set_cell_rtl(cell)
                p = cell.paragraphs[0]
                _rtl_para(p)
                _safe_run(p, txt, size=11)
                if r % 2 == 0:
                    _cell_shd(st2.rows[r + 1].cells[c], "eaf0fb")

        doc.add_paragraph()
        doc.add_page_break()

        # ── INVENTORY TABLE ─────────────────────────────────────
        self._heading("كشف الأجهزة والمعدات", 14, "1a5276")

        # Columns: م | نوع الجهاز | الماركة | الرقم التسلسلي | الحالة | الملاحظات | صورة
        hdrs    = ["م", "نوع الجهاز", "الماركة / المصنّع",
                   "الرقم التسلسلي", "الحالة", "الملاحظات", "صورة"]
        CW_CM   = [1.0, 3.5, 2.5, 2.5, 2.0, 3.5, 2.5]

        tbl = doc.add_table(len(self.inv) + 1, 7)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.RIGHT
        _set_table_rtl(tbl)

        # header row
        for c, h in enumerate(hdrs):
            cell = tbl.rows[0].cells[c]
            _set_cell_rtl(cell)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _rtl_para(p)
            _safe_run(p, h, bold=True, size=10, color_hex="ffffff")
            _cell_shd(cell, "1a5276")

        # data rows
        for r, item in enumerate(self.inv):
            row = tbl.rows[r + 1]

            # row height
            trPr = row._tr.get_or_add_trPr()
            trH  = OxmlElement("w:trHeight")
            trH.set(qn("w:val"), "1700")
            trH.set(qn("w:hRule"), "exact")
            trPr.append(trH)

            vals = [
                str(r + 1),
                item.get("device_type", "—"),
                item.get("brand", "—"),
                item.get("serial_number", "—"),
                item.get("condition", "—"),
                item.get("notes", "—"),
                None,   # photo placeholder
            ]

            for c, val in enumerate(vals):
                cell = row.cells[c]
                _set_cell_rtl(cell)
                p = cell.paragraphs[0]
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER
                                if c in (0, 6) else WD_ALIGN_PARAGRAPH.RIGHT)
                _rtl_para(p)

                if c == 6:   # photo cell
                    ph = None
                    for pi in item.get("photos", []):
                        if pi.get("is_primary"):
                            ph = pi["data"]
                            break
                    if ph is None and item.get("photos"):
                        ph = item["photos"][0]["data"]
                    if ph:
                        try:
                            run = p.add_run()
                            run.add_picture(
                                io.BytesIO(thumb(ph, (150, 150))),
                                width=Cm(2.4), height=Cm(2.4))
                        except Exception:
                            _safe_run(p, "—", size=9)
                    else:
                        _safe_run(p, "—", size=9)
                else:
                    _safe_run(p, str(val) if val else "—", size=9)

                if r % 2 == 0:
                    _cell_shd(cell, "eaf0fb")

        doc.add_paragraph()

        # ── DUPLICATES ─────────────────────────────────────────
        if self.dups:
            doc.add_page_break()
            self._heading("أجهزة بأرقام تسلسلية مكررة", 13, "c0392b")
            dt = doc.add_table(len(self.dups) + 1, 2)
            dt.style = "Table Grid"
            dt.alignment = WD_TABLE_ALIGNMENT.RIGHT
            _set_table_rtl(dt)
            for c, txt in enumerate(["الرقم التسلسلي", "أرقام الصفوف"]):
                cell = dt.rows[0].cells[c]
                _set_cell_rtl(cell)
                p = cell.paragraphs[0]
                _rtl_para(p)
                _safe_run(p, txt, bold=True, size=11, color_hex="ffffff")
                _cell_shd(cell, "c0392b")
            for r, (sn, idx) in enumerate(self.dups):
                for c, txt in enumerate([sn, "، ".join(str(x) for x in idx)]):
                    p = dt.rows[r + 1].cells[c].paragraphs[0]
                    _rtl_para(p)
                    _safe_run(p, txt, size=10)
            doc.add_paragraph()

        # ── SIGNATURE ──────────────────────────────────────────
        doc.add_page_break()
        doc.add_paragraph()

        sig = doc.add_table(4, 2)
        sig.style = "Table Grid"
        sig.alignment = WD_TABLE_ALIGNMENT.RIGHT
        _set_table_rtl(sig)

        # NOTE: "أ.د/" prefix removed from signature per requirements
        sig_rows = [
            ("اعتُمد بمعرفة:",                      "المراجع / الرئيس المباشر:"),
            (self.prof or "......................",    "......................"),
            ("", ""),
            ("التوقيع: _______________",             "التوقيع: _______________"),
        ]
        for r, (left_txt, right_txt) in enumerate(sig_rows):
            for c, txt in enumerate([left_txt, right_txt]):
                cell = sig.rows[r].cells[c]
                _set_cell_rtl(cell)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _rtl_para(p)
                _safe_run(p, txt, bold=(r == 1), size=11)

        doc.add_paragraph()
        fp2 = doc.add_paragraph()
        fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _rtl_para(fp2)
        _safe_run(fp2,
                  f"هذا الكشف صادر من إدارة {LAB_NAME} — {FAC_NAME} — {UNIV_NAME}",
                  size=9, color_hex="999999")

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()


# ================================================================
# 📱  UI HELPERS
# ================================================================
def nav():
    cols = st.columns(3)
    items = [
        ("📷", "إضافة جهاز",                          "capture"),
        ("📋", f"القائمة ({len(st.session_state.inventory)})", "list"),
        ("📤", "تصدير",                                "export"),
    ]
    for col, (icon, label, pg) in zip(cols, items):
        active = "✅ " if st.session_state.page == pg else ""
        with col:
            if st.button(f"{active}{icon} {label}", key=f"nav_{pg}"):
                st.session_state.page = pg
                st.rerun()


def add_photo(data: bytes, is_primary: bool, label: str):
    """Add a photo only if it hasn't been added before (hash-based dedup)."""
    h = hash(data)
    if h not in st.session_state.photo_hashes:
        st.session_state.photo_hashes.add(h)
        st.session_state.captured_photos.append({
            "id":         uuid.uuid4().hex[:8],
            "data":       data,
            "is_primary": is_primary,
            "label":      label,
        })


# ================================================================
# ✏️  IMAGE EDITOR
# ================================================================
def image_editor(idx: int):
    ph   = st.session_state.captured_photos[idx]
    data = ph["data"]
    st.markdown('<div class="editor-box">', unsafe_allow_html=True)
    st.markdown(f"**✏️ تعديل: {ph['label']}**")

    c1, c2 = st.columns([3, 1])
    with c1:
        rotate = st.select_slider(
            "↩️ تدوير",
            options=[0, 90, 180, 270],
            value=0,
            key=f"rot_{ph['id']}",
            help="اضبط حتى تظهر الصورة بشكل صحيح")
    with c2:
        enh = st.checkbox("✨ تحسين", value=True, key=f"enh_{ph['id']}")

    cc1, cc2 = st.columns(2)
    cl = cc1.slider("✂️ قص يسار %",  0, 45, 0, key=f"cl_{ph['id']}")
    cr = cc1.slider("✂️ قص يمين %", 0, 45, 0, key=f"cr_{ph['id']}")
    ct = cc2.slider("✂️ قص أعلى %",  0, 45, 0, key=f"ct_{ph['id']}")
    cb = cc2.slider("✂️ قص أسفل %", 0, 45, 0, key=f"cb_{ph['id']}")
    crop_box = (cl/100, ct/100, 1.0-cr/100, 1.0-cb/100)

    try:
        preview = process_image(data, rotate_deg=rotate,
                                crop_box=crop_box, enhance=enh)
        st.image(preview, use_container_width=True, caption="👁️ معاينة حية")
    except Exception as e:
        st.error(f"خطأ معاينة: {e}")
        preview = data

    ba, bb = st.columns(2)
    with ba:
        if st.button("✅ تطبيق التعديلات", key=f"apply_{ph['id']}", type="primary"):
            st.session_state.captured_photos[idx]["data"] = preview
            st.session_state.edit_photo_idx = None
            st.rerun()
    with bb:
        if st.button("↩️ إلغاء", key=f"cancel_{ph['id']}"):
            st.session_state.edit_photo_idx = None
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ================================================================
# 📷  PAGE: CAPTURE
# ================================================================
def page_capture():
    # Settings
    with st.expander("⚙️ إعدادات التقرير",
                     expanded=not st.session_state.professor_name):
        c1, c2 = st.columns(2)
        st.session_state.professor_name = c1.text_input(
            "👤 اسم الأستاذ المسؤول",
            value=st.session_state.professor_name,
            placeholder="تامر الغرباوي...")
        st.session_state.report_date = c2.text_input(
            "📅 تاريخ الجرد",
            value=st.session_state.report_date)

    st.markdown("---")

    # ── STEP 1: Single Photo ─────────────────────────────────────
    st.markdown('<div class="step-label">📸 الخطوة 1 — رفع صورة الجهاز</div>',
                unsafe_allow_html=True)

    # form_counter key ensures the uploader resets after each "add" action
    fk = st.session_state.form_counter
    prim = st.file_uploader(
        "⭐ صورة الجهاز (تظهر في التقرير)",
        type=["jpg", "jpeg", "png", "webp"],
        key=f"up_primary_{fk}")

    ca, cb_btn = st.columns(2)
    with ca:
        if st.button("📥 تثبيت الصورة"):
            if prim:
                # Replace any existing photo (single-image mode)
                st.session_state.captured_photos = []
                st.session_state.photo_hashes    = set()
                add_photo(resize_img(prim.read()), True, "الصورة الرئيسية")
                st.success("✅ تمت إضافة الصورة")
                st.rerun()
            else:
                st.warning("⚠️ اختر صورة أولاً")
    with cb_btn:
        if st.button("🗑️ مسح الصورة"):
            st.session_state.captured_photos = []
            st.session_state.photo_hashes    = set()
            st.session_state.edit_photo_idx  = None
            st.rerun()

    # Photo preview
    shots = st.session_state.captured_photos
    if shots:
        st.markdown("**صورة مُثبَّتة — اضغط ✏️ لتعديل الاتجاه والقص:**")

        if st.session_state.edit_photo_idx is not None:
            idx = st.session_state.edit_photo_idx
            if 0 <= idx < len(shots):
                image_editor(idx)
            else:
                st.session_state.edit_photo_idx = None
        else:
            ph = shots[0]
            col_img, col_btns = st.columns([2, 1])
            with col_img:
                st.markdown(f"""
                <div class="photo-card">
                  <span class="badge-primary">{ph['label']}</span><br><br>
                  <img src="data:image/jpeg;base64,{b64img(ph['data'])}"
                       style="width:100%;border-radius:8px;max-height:200px;object-fit:cover;"/>
                </div>""", unsafe_allow_html=True)
            with col_btns:
                if st.button("✏️ تعديل الصورة", key=f"ed_{ph['id']}"):
                    st.session_state.edit_photo_idx = 0
                    st.rerun()

        if st.button("✨ تحسين تلقائي (تباين + حدة)"):
            with st.spinner("⏳ معالجة..."):
                d = st.session_state.captured_photos[0]["data"]
                st.session_state.captured_photos[0]["data"] = process_image(
                    d, enhance=True)
            st.success("✅ تمت المعالجة")
            st.rerun()

    st.markdown("---")

    # ── STEP 2: Manual Data Entry ─────────────────────────────────
    st.markdown('<div class="step-label">📝 الخطوة 2 — إدخال بيانات الجهاز يدوياً</div>',
                unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box">
      📋 أدخل بيانات الجهاز يدوياً من خلال قراءة لوحة البيانات أو الجهاز نفسه
    </div>""", unsafe_allow_html=True)

    # Keyed inputs reset automatically when form_counter increments
    fc1, fc2 = st.columns(2)
    device_type = fc1.text_input(
        "🔧 نوع الجهاز",
        placeholder="مثال: جهاز مستوي آلي، ترازيت، GPS، محطة شاملة...",
        key=f"device_type_{fk}")
    brand = fc2.text_input(
        "🏭 الماركة / المصنّع",
        placeholder="مثال: Leica, Trimble, Topcon, Sokkia...",
        key=f"brand_{fk}")
    serial = fc1.text_input(
        "🔢 الرقم التسلسلي",
        placeholder="اقرأه من لوحة بيانات الجهاز...",
        key=f"serial_{fk}")
    condition = fc2.selectbox("📊 الحالة", STATUS_OPT, index=2, key=f"cond_{fk}")
    notes = st.text_area(
        "📒 ملاحظات الأستاذ",
        height=100,
        placeholder="أي ملاحظات تقنية أو حالة الجهاز أو ملحقاته...",
        key=f"notes_{fk}")

    # Duplicate warning — only fires when serial is non-empty AND already in inventory
    if serial and is_duplicate(serial):
        st.markdown(f"""
        <div class="dup-warning">
          ⚠️ <strong>تحذير: الرقم التسلسلي مكرر!</strong><br>
          الرقم <code>{serial}</code> مسجّل مسبقاً في القائمة — تحقق قبل الإضافة.
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    s1, s2 = st.columns([3, 1])
    with s1:
        if st.button("✅ إضافة إلى قائمة الجرد", type="primary",
                     disabled=not device_type):
            st.session_state.inventory.append({
                "id":            uuid.uuid4().hex,
                "device_type":   device_type,
                "brand":         brand or "—",
                "serial_number": serial or "—",
                "condition":     condition,
                "notes":         notes,
                "photos":        list(st.session_state.captured_photos),
                "is_duplicate":  is_duplicate(serial),
                "added_at":      datetime.datetime.now().strftime("%H:%M"),
            })
            # Reset form by incrementing the counter
            st.session_state.form_counter  += 1
            st.session_state.captured_photos = []
            st.session_state.photo_hashes    = set()
            st.session_state.edit_photo_idx  = None
            st.markdown('<div class="success-flash">✅ تمت الإضافة بنجاح!</div>',
                        unsafe_allow_html=True)
            st.rerun()
    with s2:
        if st.button("↩️ مسح النموذج"):
            st.session_state.form_counter  += 1
            st.session_state.captured_photos = []
            st.session_state.photo_hashes    = set()
            st.session_state.edit_photo_idx  = None
            st.rerun()


# ================================================================
# 📋  PAGE: LIST
# ================================================================
def page_list():
    inv = st.session_state.inventory
    if not inv:
        st.info("📭 قائمة الجرد فارغة — أضف أجهزة من صفحة 'إضافة جهاز'.")
        return

    dups   = len(get_duplicates())
    brands = len(set(i.get("brand", "") for i in inv))
    maint  = sum(1 for i in inv
                  if i.get("condition") in ["يحتاج صيانة", "معطل"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 إجمالي الأجهزة", len(inv))
    c2.metric("🏭 الماركات",        brands)
    c3.metric("⚠️ تحتاج صيانة",    maint)
    c4.metric("🔄 مكررة",           dups)

    if dups:
        st.error(f"⚠️ يوجد **{dups}** رقم تسلسلي مكرر")

    st.markdown("---")

    for i, item in enumerate(inv):
        dup = item.get("is_duplicate") or is_duplicate(
            item.get("serial_number", ""), i)
        ph  = next((p["data"] for p in item.get("photos", [])
                     if p.get("is_primary")), None)
        if ph is None and item.get("photos"):
            ph = item["photos"][0]["data"]

        ci, cd, cx = st.columns([1, 4, 1])
        with ci:
            if ph:
                st.image(thumb(ph, (120, 120)), use_container_width=True)
            else:
                st.markdown("🖼️")
        with cd:
            dup_b = "  🔴 **مكرر**" if dup else ""
            st.markdown(
                f"**{i+1}. {item.get('device_type','غير محدد')}**{dup_b}")
            st.markdown(
                f"🏭 `{item.get('brand','—')}` | "
                f"🔢 `{item.get('serial_number','—')}` | "
                f"📊 {item.get('condition','—')} | "
                f"🕐 {item.get('added_at','')}")
            if item.get("notes"):
                st.caption(f"📝 {item['notes'][:100]}")
        with cx:
            if st.button("🗑️", key=f"di_{i}"):
                st.session_state.inventory.pop(i)
                st.rerun()
        st.markdown("---")


# ================================================================
# 📤  PAGE: EXPORT
# ================================================================
def page_export():
    inv = st.session_state.inventory
    if not inv:
        st.warning("⚠️ أضف أجهزة إلى قائمة الجرد أولاً.")
        return

    dups = get_duplicates()
    st.markdown(f"""
    <div style="background:#eaf0fb;border-radius:12px;padding:1.1rem;direction:rtl;
                border-right:4px solid #1a5276;margin-bottom:1rem;">
      <strong>👤</strong> المُعِد: {st.session_state.professor_name or '<em>غير محدد</em>'}<br>
      <strong>📅</strong> التاريخ: {st.session_state.report_date}<br>
      <strong>📦</strong> عدد الأجهزة: {len(inv)} &nbsp;|&nbsp;
      <strong>⚠️</strong> أرقام مكررة: {len(dups)}
    </div>""", unsafe_allow_html=True)

    st.markdown("#### 📝 ملف Word (DOCX)")
    st.caption("للتعديل والمراجعة — جدول RTL محاذٍ لليمين")

    if st.button("📄 توليد DOCX", type="primary"):
        with st.spinner("⏳ جاري إنشاء DOCX..."):
            try:
                docx_b = DOCXReport(
                    inv,
                    st.session_state.professor_name,
                    st.session_state.report_date).build()
                fn = (f"جرد_{LAB_NAME}_"
                      f"{st.session_state.report_date.replace('/','-')}.docx")
                MIME = ("application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.document")
                st.download_button(
                    "⬇️ تحميل DOCX", docx_b, fn, MIME, key="dl_docx")
                st.success("✅ DOCX جاهز للتحميل!")
            except Exception as e:
                st.error(f"❌ خطأ DOCX: {e}")
                import traceback
                st.code(traceback.format_exc(), language="text")

    st.markdown("---")
    with st.expander("⚠️ خيارات متقدمة"):
        if st.button("🗑️ مسح قائمة الجرد بالكامل"):
            for k in ["inventory", "captured_photos"]:
                st.session_state[k] = []
            st.session_state.photo_hashes  = set()
            st.session_state.form_counter += 1
            st.success("✅ تم المسح")
            st.rerun()


# ================================================================
# 🚀  MAIN
# ================================================================
def main():
    st.markdown(f"""
    <div class="app-header">
      <h1>🔭 نظام جرد معمل المساحة</h1>
      <p>{UNIV_NAME} | {FAC_NAME} | إدخال يدوي مع معالجة الصور</p>
    </div>""", unsafe_allow_html=True)

    nav()
    st.markdown("")

    pg = st.session_state.page
    if   pg == "capture": page_capture()
    elif pg == "list":    page_list()
    elif pg == "export":  page_export()


if __name__ == "__main__":
    main()
