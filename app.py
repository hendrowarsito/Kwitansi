"""
SRR Billing Generator — Generator Kwitansi & Surat Tagihan KJPP SRR
v1.0 | Streamlit Cloud App
"""

import streamlit as st
import anthropic
import json
import io
import zipfile
import base64
import re
from datetime import datetime, date
from copy import deepcopy

# ── openpyxl (template xlsx)
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ── python-docx (template docx)
from docx import Document
from docx.shared import Pt

# ── PDF reading
import pdfplumber

# ─────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="SRR Billing Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# CUSTOM CSS  — clean professional look
# ─────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif; }

.srr-header {
    background: linear-gradient(135deg, #0f3460 0%, #16213e 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.srr-header h1 { margin:0; font-size:1.6rem; font-weight:700; letter-spacing:-0.5px; }
.srr-header p  { margin:0; font-size:0.85rem; opacity:0.75; }

.step-pill {
    display: inline-flex; align-items: center; gap: 6px;
    background: #eef2ff; color: #3730a3;
    border-radius: 20px; padding: 4px 12px;
    font-size: 0.78rem; font-weight: 600;
    margin-bottom: 0.5rem;
}

.project-card {
    border: 1.5px solid #e2e8f0;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.8rem;
    background: white;
    transition: box-shadow 0.2s;
}
.project-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.08); }

.tag-pill {
    display: inline-block;
    background: #f0fdf4; color: #166534;
    border-radius: 6px; padding: 2px 8px;
    font-size: 0.75rem; font-weight: 600;
    font-family: 'JetBrains Mono', monospace;
}

.amount-display {
    font-family: 'JetBrains Mono', monospace;
    font-weight: 600;
    color: #0f3460;
    font-size: 1.1rem;
}

.stButton > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# TERBILANG RUPIAH
# ─────────────────────────────────────────
def terbilang_rupiah(n: int) -> str:
    satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan',
              'Sepuluh', 'Sebelas', 'Dua Belas', 'Tiga Belas', 'Empat Belas', 'Lima Belas',
              'Enam Belas', 'Tujuh Belas', 'Delapan Belas', 'Sembilan Belas']
    puluhan = ['', '', 'Dua Puluh', 'Tiga Puluh', 'Empat Puluh', 'Lima Puluh',
               'Enam Puluh', 'Tujuh Puluh', 'Delapan Puluh', 'Sembilan Puluh']

    def _t(n):
        if n == 0:
            return ''
        elif n < 20:
            return satuan[n]
        elif n < 100:
            ones = satuan[n % 10]
            return (puluhan[n // 10] + (' ' + ones if ones else '')).strip()
        elif n < 200:
            rest = _t(n - 100)
            return 'Seratus' + (' ' + rest if rest else '')
        elif n < 1000:
            rest = _t(n % 100)
            return satuan[n // 100] + ' Ratus' + (' ' + rest if rest else '')
        elif n < 2000:
            rest = _t(n - 1000)
            return 'Seribu' + (' ' + rest if rest else '')
        elif n < 1_000_000:
            rest = _t(n % 1000)
            return _t(n // 1000) + ' Ribu' + (' ' + rest if rest else '')
        elif n < 1_000_000_000:
            rest = _t(n % 1_000_000)
            return _t(n // 1_000_000) + ' Juta' + (' ' + rest if rest else '')
        elif n < 1_000_000_000_000:
            rest = _t(n % 1_000_000_000)
            return _t(n // 1_000_000_000) + ' Miliar' + (' ' + rest if rest else '')
        else:
            rest = _t(n % 1_000_000_000_000)
            return _t(n // 1_000_000_000_000) + ' Triliun' + (' ' + rest if rest else '')

    if n == 0:
        return 'Nol Rupiah'
    return _t(int(n)) + ' Rupiah'


# ─────────────────────────────────────────
# FEE CALCULATOR
# ─────────────────────────────────────────
def hitung_tagihan(imbalan_jasa: float, tarif_ppn: float = 0.12) -> dict:
    """Hitung DPP + PPN + Total sesuai pola KJPP SRR (DPP = IJ × 11/12)"""
    ij = round(imbalan_jasa)
    dpb = round(ij * 11 / 12)
    ppn = round(dpb * tarif_ppn)
    total = ij + ppn
    return {
        "imbalan_jasa": ij,
        "dpb_ppn": dpb,
        "ppn": ppn,
        "total": total,
        "terbilang": terbilang_rupiah(total),
    }


# ─────────────────────────────────────────
# NOMOR SURAT GENERATOR
# ─────────────────────────────────────────
def generate_nomor(tanggal: date, seq: int, tipe: str, kode_klien: str) -> str:
    """
    Format: YYMMDD.SEQ/SRR-JK/TIPE/KODEKLIEN
    Contoh KWT : 260115.001/SRR-JK/KWT.PJK/IDXSTI
    Contoh SK-OR: 260115.001/SRR-JK/SK-OR/IDXSTI
    """
    prefix = tanggal.strftime("%y%m%d")
    seq_str = f"{seq:03d}"
    kode = kode_klien.upper().replace(" ", "").replace(".", "")[:8]
    return f"{prefix}.{seq_str}/SRR-JK/{tipe}/{kode}"


def format_tanggal_indo(d: date) -> str:
    bulan = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    return f"{d.day} {bulan[d.month]} {d.year}"


# ─────────────────────────────────────────
# CLAUDE API — EXTRACT FROM PDF/DOCX
# ─────────────────────────────────────────
def extract_proposal_data(file_bytes: bytes, filename: str) -> dict:
    """Use Claude API to extract structured data from proposal PDF/DOCX"""

    # Extract text from PDF
    if filename.lower().endswith(".pdf"):
        text_parts = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_parts.append(t)
        raw_text = "\n".join(text_parts)
    else:
        # DOCX: extract paragraphs
        doc = Document(io.BytesIO(file_bytes))
        raw_text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())

    # Limit text to ~6000 chars for API efficiency
    raw_text = raw_text[:6000]

    client = anthropic.Anthropic()
    prompt = f"""Kamu adalah asisten ekstraksi data dari proposal jasa penilaian KJPP SRR.

Ekstrak informasi berikut dari teks proposal di bawah dan kembalikan HANYA JSON (tanpa komentar, tanpa markdown).

Field yang dibutuhkan:
- nama_klien: nama lengkap perusahaan klien
- nama_klien_singkat: akronim/kode klien (contoh: PTRO, IDXSTI, KMS)
- alamat_baris1: baris pertama alamat
- alamat_baris2: baris kedua alamat (jalan, dll)
- kota: nama kota
- kode_pos: kode pos (string)
- up: jabatan yang dituju (contoh: Direksi)
- jenis_pekerjaan: deskripsi singkat jenis pekerjaan (contoh: Penilaian Saham dan Pendapat Kewajaran)
- nomor_proposal: nomor proposal (contoh: 260112.001/SRR-JK/SPN-ABF/PTRO/OR)
- tanggal_proposal: tanggal proposal dalam format "DD Bulan YYYY" (contoh: 12 Januari 2026)
- imbalan_jasa_total: total imbalan jasa sebelum PPN dalam angka integer (ambil dari baris "Jumlah Imbalan Jasa/Fee")
- tanggal_tagihan: tanggal penilaian/cut-off yang disebutkan, atau tanggal surat jika tidak ada, format "DD Bulan YYYY"
- receiver: nama penandatangan dari SRR (biasanya Ocky Rinaldy)

Jika suatu field tidak ditemukan, isi dengan string kosong "".
Untuk imbalan_jasa_total: ambil angka yang paling besar yang disebutkan sebagai total fee, jika ada beberapa komponen jumlahkan.

TEKS PROPOSAL:
{raw_text}

JSON:"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = response.content[0].text.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Fallback: return empty template
        data = {
            "nama_klien": "", "nama_klien_singkat": "",
            "alamat_baris1": "", "alamat_baris2": "",
            "kota": "", "kode_pos": "", "up": "Direksi",
            "jenis_pekerjaan": "", "nomor_proposal": "",
            "tanggal_proposal": "", "imbalan_jasa_total": 0,
            "tanggal_tagihan": "", "receiver": "Ocky Rinaldy"
        }

    # Ensure numeric
    try:
        data["imbalan_jasa_total"] = int(str(data.get("imbalan_jasa_total", 0)).replace(",", "").replace(".", ""))
    except (ValueError, TypeError):
        data["imbalan_jasa_total"] = 0

    return data


# ─────────────────────────────────────────
# KWITANSI GENERATOR (openpyxl)
# ─────────────────────────────────────────
def generate_kwitansi(template_bytes: bytes, project: dict, seq: int) -> bytes:
    """Inject project data into kwitansi XLSX template."""
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    tanggal = project["tanggal_tagihan_date"]
    fin = hitung_tagihan(project["imbalan_jasa_total"])
    nomor_kwt = generate_nomor(tanggal, seq, "KWT.PJK", project["nama_klien_singkat"])
    tanggal_str = format_tanggal_indo(tanggal)

    # ── Helper: find & write cell safely (preserve existing formatting)
    def write_cell(row, col, value, bold=False, align=None):
        cell = ws.cell(row=row, column=col)
        cell.value = value
        if bold:
            cell.font = Font(bold=True, name=cell.font.name if cell.font else "Arial",
                             size=cell.font.size if cell.font else 10)
        if align:
            cell.alignment = Alignment(horizontal=align, wrap_text=True)

    # ─── MAPPING ke template (berdasarkan analisis TemplateInformation)
    # Row/col disesuaikan agar cocok dengan template KJPP SRR yang diupload user
    # Template pakai col J (10) untuk Receipt No dan Amount, col E(5) untuk data klien

    # Receipt No (J12)
    write_cell(12, 10, nomor_kwt)

    # Klien Name (E11)
    write_cell(11, 5, project["nama_klien"])

    # Address (E13-14)
    write_cell(13, 5, project["alamat_baris1"])
    write_cell(14, 5, project["alamat_baris2"])

    # City (E15), ZIP (G15)
    write_cell(15, 5, project["kota"])
    write_cell(15, 7, project["kode_pos"])

    # Description rows (B20-22)
    write_cell(20, 2, f"Pembayaran jasa {project['jenis_pekerjaan']} {project['nama_klien']}")
    write_cell(21, 2, f"sesuai dengan proposal No. {project['nomor_proposal']}")
    write_cell(22, 2, f"tanggal {project['tanggal_proposal']}, sebesar:")

    # Amounts
    write_cell(24, 10, fin["imbalan_jasa"])   # Imbalan Jasa
    write_cell(26, 10, fin["dpb_ppn"])        # DPP PPN
    write_cell(28, 10, fin["ppn"])            # PPN 12%
    write_cell(38, 10, fin["total"])          # Total

    # Terbilang (C40) + Tanggal (I40)
    write_cell(40, 3, fin["terbilang"])
    write_cell(40, 9, f"Jakarta, {tanggal_str}")

    # Receiver (H47)
    write_cell(47, 8, project["receiver"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────
# SURAT TAGIHAN GENERATOR (python-docx)
# ─────────────────────────────────────────
def replace_in_run(run, placeholder, value):
    if placeholder in run.text:
        run.text = run.text.replace(placeholder, str(value))

def replace_in_paragraph(para, replacements):
    """Replace placeholders in a paragraph, handling split-run cases."""
    # First try simple per-run replacement
    for run in para.runs:
        for k, v in replacements.items():
            if k in run.text:
                run.text = run.text.replace(k, str(v))

    # If placeholder is split across runs, fix full paragraph text
    full_text = "".join(r.text for r in para.runs)
    for k, v in replacements.items():
        if k in full_text:
            full_text = full_text.replace(k, str(v))
            # Rewrite all text into first run, clear rest
            if para.runs:
                para.runs[0].text = full_text
                for r in para.runs[1:]:
                    r.text = ""
            break

def replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_in_paragraph(para, replacements)

def generate_surat(template_bytes: bytes, project: dict, seq: int) -> bytes:
    """Inject project data into surat tagihan DOCX template."""
    doc = Document(io.BytesIO(template_bytes))

    tanggal = project["tanggal_tagihan_date"]
    fin = hitung_tagihan(project["imbalan_jasa_total"])
    nomor_sk = generate_nomor(tanggal, seq, "SK-OR", project["nama_klien_singkat"])
    tanggal_str = format_tanggal_indo(tanggal)
    total_fmt = f"Rp {fin['total']:,.0f}".replace(",", ".")

    replacements = {
        "{{NOMOR_SURAT}}":    nomor_sk,
        "{{TANGGAL_SURAT}}":  tanggal_str,
        "{{NAMA_KLIEN}}":     project["nama_klien"],
        "{{ALAMAT1}}":        project["alamat_baris1"],
        "{{ALAMAT2}}":        project["alamat_baris2"],
        "{{KOTA_POS}}":       f"{project['kota']} {project['kode_pos']}".strip(),
        "{{UP}}":             project["up"],
        "{{JENIS_PEKERJAAN}}":project["jenis_pekerjaan"],
        "{{NOMOR_PROPOSAL}}": project["nomor_proposal"],
        "{{TGL_PROPOSAL}}":   project["tanggal_proposal"],
        "{{TERBILANG}}":      fin["terbilang"],
        "{{TOTAL_ANGKA}}":    total_fmt,
        "{{RECEIVER}}":       project["receiver"],
        "{{NAMA_KLIEN_SINGKAT}}": project["nama_klien_singkat"],
    }

    for para in doc.paragraphs:
        replace_in_paragraph(para, replacements)

    for table in doc.tables:
        replace_in_table(table, replacements)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────
# TEMPLATE DOCX BUILDER (jika user tidak upload template)
# ─────────────────────────────────────────
def build_default_surat_template() -> bytes:
    """Build a basic surat tagihan DOCX template with placeholders."""
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # Margins
    section = doc.sections[0]
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)
    section.top_margin    = Cm(2.5)
    section.bottom_margin = Cm(2.0)

    def add_para(text, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, size=11, space_after=6):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        run.font.name = "Times New Roman"
        return p

    # Header KOP SURAT (simplified)
    add_para("SUWENDHO RINALDY DAN REKAN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13)
    add_para("KANTOR JASA PENILAI PUBLIK", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=11)
    add_para("Nomor Izin Usaha KJPP: 2.09.0059 | Nomor Izin Cabang KJPP: 1138/KM.1/2017",
             align=WD_ALIGN_PARAGRAPH.CENTER, size=9)
    add_para("Komplek Kalibata Indah Blok K16-17, Jl. Rawajati Timur, Pancoran, Jakarta Selatan 12750",
             align=WD_ALIGN_PARAGRAPH.CENTER, size=9)
    add_para("─" * 80, align=WD_ALIGN_PARAGRAPH.CENTER, size=9, space_after=12)

    # Nomor & Tanggal
    add_para("No. : {{NOMOR_SURAT}}                                    {{TANGGAL_SURAT}}", size=11)
    add_para("")

    # Kepada
    add_para("Kepada Yth.", size=11)
    add_para("{{NAMA_KLIEN}}", bold=True, size=11)
    add_para("{{ALAMAT1}}", bold=True, size=11)
    add_para("{{ALAMAT2}}", bold=True, size=11)
    add_para("{{KOTA_POS}}", bold=True, size=11)
    add_para("")
    add_para("U.p.  :  {{UP}}", size=11)
    add_para("")

    # Hal
    add_para("Hal   :  Penagihan Pembayaran Jasa {{JENIS_PEKERJAAN}}", bold=True, size=11)
    add_para("          {{NAMA_KLIEN_SINGKAT}}", bold=True, size=11)
    add_para("")
    add_para("Dengan hormat,", size=11)
    add_para("")

    # Body
    body = (
        "Menunjuk surat penawaran kami No. {{NOMOR_PROPOSAL}} tanggal {{TGL_PROPOSAL}}, "
        "maka dengan ini kami mohon agar pembayaran untuk penugasan tersebut sebesar:"
    )
    add_para(body, size=11, space_after=12)
    add_para("")

    # Nominal box
    p_box = doc.add_paragraph()
    p_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_box.add_run("{{TOTAL_ANGKA}}")
    r.bold = True
    r.font.size = Pt(14)
    r.font.name = "Times New Roman"

    p_terb = doc.add_paragraph()
    p_terb.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_terb.paragraph_format.space_after = Pt(12)
    r2 = p_terb.add_run("(  {{TERBILANG}}  )")
    r2.font.size = Pt(11)
    r2.font.name = "Times New Roman"

    add_para("")
    add_para(
        "(kwitansi dan faktur PPN terlampir) dapat dibayarkan kepada kami dengan bilyet giro "
        "atau ditransfer ke rekening kami atas nama KJPP SUWENDHO RINALDY & REKAN di "
        "Bank Mandiri KCP JKT Kalibata Rawajati dengan nomor rekening 126-0005748719 "
        "pada kesempatan pertama.",
        size=11, space_after=12
    )
    add_para("")
    add_para("Demikianlah permohonan kami, atas perhatian dan kerja sama Bapak/Ibu kami ucapkan terima kasih.", size=11)
    add_para("")
    add_para("Hormat kami,", size=11)
    add_para("")
    add_para("")
    add_para("{{RECEIVER}}", size=11)
    add_para("Rekan", size=11)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def build_default_kwitansi_template() -> bytes:
    """Build a minimal kwitansi XLSX template with the correct cell positions."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    thin = Side(style="thin")
    border_box = Border(left=thin, right=thin, top=thin, bottom=thin)

    def c(row, col, value, bold=False, align="left", size=10, wrap=False, number=False):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = Font(name="Arial", bold=bold, size=size)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        if number and isinstance(value, (int, float)):
            cell.number_format = '#,##0'
        return cell

    # KOP
    c(3, 2, "SUWENDHO RINALDY DAN REKAN", bold=True, align="center", size=13)
    c(4, 2, "KANTOR JASA PENILAI PUBLIK", align="center", size=9)
    c(5, 2, "Nomor Izin Usaha KJPP: 2.09.0059 | Izin Cabang: 1138/KM.1/2017", align="center", size=8)

    # Title KWITANSI
    c(7, 5, "KWITANSI", bold=True, align="center", size=16)

    # Receipt No label + value
    c(11, 4, "Name", bold=False, size=9)
    c(11, 5, ":  [NAMA KLIEN]", size=9)      # placeholder untuk referensi user
    c(12, 10, "[NOMOR KWT]", size=9)         # Receipt No

    c(13, 4, "Address", size=9)
    c(13, 5, ":  [ALAMAT 1]", size=9)
    c(14, 5, "   [ALAMAT 2]", size=9)
    c(15, 4, "City", size=9)
    c(15, 5, ":  [KOTA]", size=9)
    c(15, 6, "ZIP", size=9)
    c(15, 7, "[KODEPOS]", size=9)
    c(10, 10, "Receipt No.", bold=True, size=9)

    # Description header
    c(18, 2, "Description", bold=True, size=10)
    c(18, 10, "Amount (Rp)", bold=True, align="right", size=10)

    c(20, 2, "[KETERANGAN PEKERJAAN]", size=10)
    c(21, 2, "[NOMOR PROPOSAL]", size=10)
    c(22, 2, "[TANGGAL PROPOSAL]", size=10)

    c(24, 2, "Imbalan Jasa", size=10)
    c(24, 10, 0, align="right", number=True, size=10)

    c(26, 2, "Dasar Pengenaan PPN (Imbalan Jasa x 11/12)", size=10)
    c(26, 10, 0, align="right", number=True, size=10)

    c(28, 2, "PPN 12%", size=10)
    c(28, 10, 0, align="right", number=True, size=10)

    c(38, 9, "Total", bold=True, size=11)
    c(38, 10, 0, bold=True, align="right", number=True, size=11)

    c(39, 2, "Say Rupiah :", bold=True, size=10)
    c(40, 3, "[TERBILANG]", size=10)
    c(40, 9, "Jakarta,", size=10)
    c(40, 10, "[TANGGAL]", size=10)

    c(47, 8, "[PENERIMA]", size=10)

    # Footer bank
    c(51, 2, "Pembayaran melalui Bank Mandiri KCP JKT Kalibata Rawajati, atas nama KJPP Suwendho Rinaldy & Rekan No. Rek. 126-0005748719", size=8)

    # Column widths
    col_widths = {1: 4, 2: 30, 3: 20, 4: 12, 5: 28, 6: 6, 7: 8, 8: 18, 9: 10, 10: 18}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────
if "projects"        not in st.session_state: st.session_state.projects = []
if "kwt_template"    not in st.session_state: st.session_state.kwt_template = None
if "sk_template"     not in st.session_state: st.session_state.sk_template = None
if "seq_counter"     not in st.session_state: st.session_state.seq_counter = 1
if "edit_idx"        not in st.session_state: st.session_state.edit_idx = None


# ─────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────
st.markdown("""
<div class="srr-header">
  <div>
    <h1>📄 SRR Billing Generator</h1>
    <p>Generator Kwitansi & Surat Penagihan — KJPP Suwendho Rinaldy dan Rekan</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# SIDEBAR — SETTINGS + TEMPLATES
# ─────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Pengaturan")

    receiver = st.text_input("Nama Penandatangan", value="Ocky Rinaldy")
    seq_start = st.number_input("Nomor Urut Awal", min_value=1, max_value=999,
                                 value=st.session_state.seq_counter)
    if seq_start != st.session_state.seq_counter:
        st.session_state.seq_counter = seq_start

    st.divider()
    st.markdown("### 📎 Upload Template")

    kwt_upload = st.file_uploader("Template Kwitansi (.xlsx)", type=["xlsx", "xls"],
                                   key="kwt_up",
                                   help="Upload template kwitansi .xlsx Anda")
    if kwt_upload:
        if kwt_upload.name.endswith(".xls"):
            st.warning("Format .xls terdeteksi. Konversi ke .xlsx disarankan untuk hasil terbaik.")
        else:
            st.session_state.kwt_template = kwt_upload.read()
            st.success("✅ Template kwitansi tersimpan")

    sk_upload = st.file_uploader("Template Surat Tagihan (.docx)", type=["docx"],
                                  key="sk_up",
                                  help="Upload template surat tagihan .docx dengan placeholder {{NAMA_KLIEN}} dll")
    if sk_upload:
        st.session_state.sk_template = sk_upload.read()
        st.success("✅ Template surat tersimpan")

    if not st.session_state.kwt_template:
        st.info("ℹ️ Belum ada template kwitansi — akan gunakan template default SRR")
    if not st.session_state.sk_template:
        st.info("ℹ️ Belum ada template surat — akan gunakan template default SRR")

    st.divider()
    st.markdown("#### 💡 Placeholder Surat")
    with st.expander("Lihat daftar placeholder"):
        st.code("""{{NOMOR_SURAT}}
{{TANGGAL_SURAT}}
{{NAMA_KLIEN}}
{{ALAMAT1}}
{{ALAMAT2}}
{{KOTA_POS}}
{{UP}}
{{JENIS_PEKERJAAN}}
{{NOMOR_PROPOSAL}}
{{TGL_PROPOSAL}}
{{TERBILANG}}
{{TOTAL_ANGKA}}
{{RECEIVER}}
{{NAMA_KLIEN_SINGKAT}}""", language="text")


# ─────────────────────────────────────────
# MAIN TABS
# ─────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📥 Input Proyek", "📋 Review & Edit", "📦 Download"])

# ════════════════════════════════════════
# TAB 1 — INPUT
# ════════════════════════════════════════
with tab1:
    st.markdown('<div class="step-pill">LANGKAH 1 — Tambah Proyek</div>', unsafe_allow_html=True)

    input_mode = st.radio(
        "Mode input:",
        ["🤖 Upload Proposal (Ekstraksi AI)", "✏️ Input Manual"],
        horizontal=True
    )

    # ── AI MODE
    if "🤖" in input_mode:
        uploaded_proposals = st.file_uploader(
            "Upload Proposal / Kontrak (PDF atau DOCX)",
            type=["pdf", "docx"],
            accept_multiple_files=True,
            help="Bisa upload beberapa file sekaligus untuk batch processing"
        )

        if uploaded_proposals:
            if st.button("🤖 Ekstrak Data dengan AI", type="primary",
                         disabled=len(uploaded_proposals) == 0):
                progress = st.progress(0)
                for i, f in enumerate(uploaded_proposals):
                    with st.spinner(f"Membaca {f.name}..."):
                        try:
                            data = extract_proposal_data(f.read(), f.name)
                            data["receiver"] = receiver
                            data["_source_file"] = f.name
                            # Parse tanggal tagihan
                            data["tanggal_tagihan_date"] = date.today()
                            st.session_state.projects.append(data)
                        except Exception as e:
                            st.error(f"Gagal proses {f.name}: {e}")
                    progress.progress((i + 1) / len(uploaded_proposals))

                st.success(f"✅ {len(uploaded_proposals)} proyek berhasil diekstrak. Lanjut ke tab **Review & Edit** untuk memeriksa & download.")
                st.balloons()

    # ── MANUAL MODE
    else:
        with st.form("manual_form", clear_on_submit=True):
            st.markdown("#### Data Klien")
            c1, c2 = st.columns(2)
            with c1:
                m_nama       = st.text_input("Nama Klien *")
                m_singkat    = st.text_input("Kode/Singkatan Klien *", help="Contoh: PTRO, IDXSTI")
                m_alamat1    = st.text_input("Alamat Baris 1")
                m_alamat2    = st.text_input("Alamat Baris 2")
            with c2:
                m_kota       = st.text_input("Kota")
                m_kodepos    = st.text_input("Kode Pos")
                m_up         = st.text_input("U.p. (Jabatan)", value="Direksi")
                m_receiver_f = st.text_input("Nama Penandatangan", value=receiver)

            st.markdown("#### Detail Penugasan")
            c3, c4 = st.columns(2)
            with c3:
                m_pekerjaan  = st.text_input("Jenis Pekerjaan *",
                                              help="Contoh: Penilaian Saham dan Pendapat Kewajaran")
                m_noprop     = st.text_input("Nomor Proposal")
                m_tglprop    = st.text_input("Tanggal Proposal", help="Contoh: 12 Januari 2026")
            with c4:
                m_fee        = st.number_input("Imbalan Jasa (Rp, sebelum PPN) *",
                                               min_value=0, step=1000000)
                m_tgltagih   = st.date_input("Tanggal Tagihan", value=date.today())

            # Preview perhitungan
            if m_fee > 0:
                fin_prev = hitung_tagihan(m_fee)
                st.info(
                    f"**Preview:** DPP = Rp {fin_prev['dpb_ppn']:,.0f} | "
                    f"PPN 12% = Rp {fin_prev['ppn']:,.0f} | "
                    f"**Total = Rp {fin_prev['total']:,.0f}**"
                )

            submitted = st.form_submit_button("➕ Tambah ke Daftar Proyek", type="primary")
            if submitted:
                if not m_nama or not m_singkat or not m_pekerjaan or m_fee <= 0:
                    st.error("⚠️ Isi semua field wajib (*)")
                else:
                    proj = {
                        "nama_klien": m_nama,
                        "nama_klien_singkat": m_singkat,
                        "alamat_baris1": m_alamat1,
                        "alamat_baris2": m_alamat2,
                        "kota": m_kota,
                        "kode_pos": m_kodepos,
                        "up": m_up,
                        "jenis_pekerjaan": m_pekerjaan,
                        "nomor_proposal": m_noprop,
                        "tanggal_proposal": m_tglprop,
                        "imbalan_jasa_total": int(m_fee),
                        "tanggal_tagihan_date": m_tgltagih,
                        "receiver": m_receiver_f,
                        "_source_file": "Manual",
                    }
                    st.session_state.projects.append(proj)
                    st.success(f"✅ {m_nama} ditambahkan!")

    # Show current count
    if st.session_state.projects:
        st.markdown(f"---\n**{len(st.session_state.projects)} proyek** dalam antrian. Lanjut ke tab **Review & Edit**.")


# ════════════════════════════════════════
# TAB 2 — REVIEW & EDIT
# ════════════════════════════════════════
with tab2:
    st.markdown('<div class="step-pill">LANGKAH 2 — Review & Edit Data</div>', unsafe_allow_html=True)

    if not st.session_state.projects:
        st.info("Belum ada proyek. Tambahkan di tab **Input Proyek** terlebih dahulu.")
    else:
        # Summary table
        summary_rows = []
        for i, p in enumerate(st.session_state.projects):
            fin = hitung_tagihan(p["imbalan_jasa_total"])
            summary_rows.append({
                "#": i + 1,
                "Klien": p["nama_klien"],
                "Kode": p["nama_klien_singkat"],
                "Pekerjaan": p["jenis_pekerjaan"][:40] + "..." if len(p["jenis_pekerjaan"]) > 40 else p["jenis_pekerjaan"],
                "Imbalan Jasa": f"Rp {p['imbalan_jasa_total']:,.0f}",
                "Total + PPN": f"Rp {fin['total']:,.0f}",
                "Sumber": p.get("_source_file", "-"),
            })

        import pandas as pd
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("### ✏️ Edit Detail Proyek")

        edit_idx = st.selectbox(
            "Pilih proyek untuk diedit:",
            options=list(range(len(st.session_state.projects))),
            format_func=lambda i: f"#{i+1} — {st.session_state.projects[i]['nama_klien']}"
        )

        p = st.session_state.projects[edit_idx]

        with st.form(f"edit_form_{edit_idx}"):
            st.markdown(f"#### Edit: {p['nama_klien']}")
            c1, c2 = st.columns(2)
            with c1:
                e_nama    = st.text_input("Nama Klien", value=p["nama_klien"])
                e_singkat = st.text_input("Kode Klien", value=p["nama_klien_singkat"])
                e_alamat1 = st.text_input("Alamat 1", value=p["alamat_baris1"])
                e_alamat2 = st.text_input("Alamat 2", value=p["alamat_baris2"])
                e_kota    = st.text_input("Kota", value=p["kota"])
                e_kodepos = st.text_input("Kode Pos", value=str(p["kode_pos"]))
            with c2:
                e_up        = st.text_input("U.p.", value=p["up"])
                e_pekerjaan = st.text_input("Jenis Pekerjaan", value=p["jenis_pekerjaan"])
                e_noprop    = st.text_input("Nomor Proposal", value=p["nomor_proposal"])
                e_tglprop   = st.text_input("Tanggal Proposal", value=p["tanggal_proposal"])
                e_fee       = st.number_input("Imbalan Jasa (Rp)",
                                               value=int(p["imbalan_jasa_total"]),
                                               min_value=0, step=1000000)
                e_tgltagih  = st.date_input("Tanggal Tagihan",
                                             value=p.get("tanggal_tagihan_date", date.today()))
                e_receiver  = st.text_input("Penandatangan", value=p["receiver"])

            # Preview
            if e_fee > 0:
                fin_e = hitung_tagihan(e_fee)
                st.info(
                    f"**Preview Tagihan:** Imbalan Jasa = Rp {fin_e['imbalan_jasa']:,.0f} | "
                    f"DPP = Rp {fin_e['dpb_ppn']:,.0f} | PPN = Rp {fin_e['ppn']:,.0f} | "
                    f"**Total = Rp {fin_e['total']:,.0f}**\n\n"
                    f"*{fin_e['terbilang']}*"
                )

            c_save, c_del = st.columns([3, 1])
            with c_save:
                if st.form_submit_button("💾 Simpan Perubahan", type="primary"):
                    st.session_state.projects[edit_idx].update({
                        "nama_klien": e_nama,
                        "nama_klien_singkat": e_singkat,
                        "alamat_baris1": e_alamat1,
                        "alamat_baris2": e_alamat2,
                        "kota": e_kota,
                        "kode_pos": e_kodepos,
                        "up": e_up,
                        "jenis_pekerjaan": e_pekerjaan,
                        "nomor_proposal": e_noprop,
                        "tanggal_proposal": e_tglprop,
                        "imbalan_jasa_total": int(e_fee),
                        "tanggal_tagihan_date": e_tgltagih,
                        "receiver": e_receiver,
                    })
                    st.success("✅ Data diperbarui!")
                    st.rerun()
            with c_del:
                if st.form_submit_button("🗑️ Hapus Proyek"):
                    st.session_state.projects.pop(edit_idx)
                    st.success("Proyek dihapus.")
                    st.rerun()


# ════════════════════════════════════════
# TAB 3 — DOWNLOAD
# ════════════════════════════════════════
with tab3:
    st.markdown('<div class="step-pill">LANGKAH 3 — Generate & Download</div>', unsafe_allow_html=True)

    if not st.session_state.projects:
        st.info("Belum ada proyek. Tambahkan terlebih dahulu di tab **Input Proyek**.")
    else:
        n = len(st.session_state.projects)
        st.success(f"**{n} proyek** siap di-generate.")

        kwt_tpl  = st.session_state.kwt_template or build_default_kwitansi_template()
        sk_tpl   = st.session_state.sk_template  or build_default_surat_template()

        if st.button("⚡ Generate Semua Dokumen", type="primary", use_container_width=True):
            zip_buf = io.BytesIO()
            errors  = []

            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, proj in enumerate(st.session_state.projects):
                    seq  = st.session_state.seq_counter + i
                    kode = proj["nama_klien_singkat"].upper()[:8]
                    tanggal = proj.get("tanggal_tagihan_date", date.today())
                    prefix  = tanggal.strftime("%y%m%d")

                    try:
                        kwt_bytes = generate_kwitansi(kwt_tpl, proj, seq)
                        kwt_fname = f"{prefix}_{seq:03d}-KW-B-{kode}_Tagihan.xlsx"
                        zf.writestr(kwt_fname, kwt_bytes)
                    except Exception as e:
                        errors.append(f"Kwitansi #{i+1} ({kode}): {e}")

                    try:
                        sk_bytes  = generate_surat(sk_tpl, proj, seq)
                        sk_fname  = f"{prefix}_{seq:03d}-SK-OR-{kode}_Tagihan.docx"
                        zf.writestr(sk_fname, sk_bytes)
                    except Exception as e:
                        errors.append(f"Surat #{i+1} ({kode}): {e}")

            zip_buf.seek(0)

            if errors:
                for err in errors:
                    st.error(f"⚠️ {err}")

            st.download_button(
                label=f"📦 Download ZIP ({n} kwitansi + {n} surat tagihan)",
                data=zip_buf.read(),
                file_name=f"SRR_Tagihan_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary",
            )

        st.divider()
        st.markdown("#### 📄 Download Individual")

        kwt_tpl_ind = st.session_state.kwt_template or build_default_kwitansi_template()
        sk_tpl_ind  = st.session_state.sk_template  or build_default_surat_template()

        for i, proj in enumerate(st.session_state.projects):
            seq   = st.session_state.seq_counter + i
            kode  = proj["nama_klien_singkat"].upper()[:8]
            tanggal = proj.get("tanggal_tagihan_date", date.today())
            fin   = hitung_tagihan(proj["imbalan_jasa_total"])

            with st.expander(f"#{i+1} — {proj['nama_klien']} | Total: Rp {fin['total']:,.0f}"):
                cA, cB = st.columns(2)
                with cA:
                    try:
                        kwt_b = generate_kwitansi(kwt_tpl_ind, proj, seq)
                        st.download_button(
                            f"📊 Kwitansi ({kode})",
                            data=kwt_b,
                            file_name=f"KWT_{kode}_{tanggal.strftime('%y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"kwt_{i}"
                        )
                    except Exception as e:
                        st.error(f"Error kwitansi: {e}")

                with cB:
                    try:
                        sk_b = generate_surat(sk_tpl_ind, proj, seq)
                        st.download_button(
                            f"📝 Surat Tagihan ({kode})",
                            data=sk_b,
                            file_name=f"SKOR_{kode}_{tanggal.strftime('%y%m%d')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"sk_{i}"
                        )
                    except Exception as e:
                        st.error(f"Error surat: {e}")

        st.divider()
        if st.button("🗑️ Hapus Semua Proyek", type="secondary"):
            st.session_state.projects = []
            st.session_state.seq_counter = 1
            st.rerun()
