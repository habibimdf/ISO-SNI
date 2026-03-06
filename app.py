import streamlit as st
import os
import re
import time
from io import BytesIO

# --- IMPORT ENGINE ---
from engine2 import DocxOptimizerEngine
from engine4 import CoverPageEngine
from engine5 import DaftarIsiEngine
from engine6 import PrakataPendahuluanEngine
from engine7 import InfoPendukungEngine
from engine9 import CustomDictionary

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="ISO Doc Master",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOM ---
st.markdown("""
    <style>
    /* ── Tombol global ── */
    .stButton>button {
        width: 100%;
        border-radius: 8px;
        height: 2.8em;
        font-weight: 600;
        transition: background 0.2s, box-shadow 0.2s;
    }
    .stButton>button:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.15); }

    /* ── Badge status kamus di sidebar ── */
    .kamus-badge {
        background: #e8f5e9;
        border-left: 4px solid #43a047;
        border-radius: 6px;
        padding: 0.5rem 0.8rem;
        font-size: 0.85rem;
        margin-bottom: 0.5rem;
    }

    /* ── Card menu pilihan ── */
    .menu-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8edf2 100%);
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.5rem;
        border-left: 4px solid #1976d2;
    }

    /* ── Header section ── */
    .section-header {
        background: linear-gradient(90deg, #1565c0 0%, #0d47a1 100%);
        color: white;
        padding: 0.6rem 1rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        font-weight: 700;
    }

    /* ── Info box ── */
    .info-step {
        background: #e3f2fd;
        border-radius: 8px;
        padding: 0.7rem 1rem;
        margin: 0.3rem 0;
        font-size: 0.9rem;
    }
    </style>
""", unsafe_allow_html=True)

# --- KONSTANTA STANDAR ISO (HARDCODED) ---
ISO_FONT_NAME = "Arial"
ISO_FONT_SIZE = 11


# ─────────────────────────────────────────────────────────────
# AUTO EKSTRAK JUDUL DARI DOKUMEN
# Strategi:
#   1. Cari paragraf dengan font size ≥ 14pt atau heading style
#   2. Judul Bahasa Indonesia  → baris pertama yang memenuhi kriteria
#   3. Judul Bahasa Inggris   → baris kedua yang memenuhi kriteria
#      (atau baris pertama yang terdeteksi italic / font size lebih kecil)
# ─────────────────────────────────────────────────────────────
def extract_titles_from_docx(docx_path: str):
    """
    Mengekstrak judul Bahasa Indonesia dan Bahasa Inggris dari dokumen .docx.

    Returns:
        (title_id, title_en) – keduanya string, bisa kosong jika tidak ditemukan.
    """
    try:
        from docx import Document
        from docx.shared import Pt

        doc = Document(docx_path)
        candidates = []

        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or len(text) < 5:
                continue

            # Cek apakah paragraf menggunakan heading style
            is_heading = para.style.name.lower().startswith('heading')

            # Cek font size dari run pertama yang punya teks
            max_size = 0
            is_italic = False
            for run in para.runs:
                if run.text.strip():
                    sz = run.font.size
                    if sz:
                        max_size = max(max_size, sz.pt if hasattr(sz, 'pt') else sz / 12700)
                    if run.font.italic:
                        is_italic = True

            # Fallback: cek dari style paragraf jika run tidak punya size
            if max_size == 0 and para.style.font.size:
                max_size = para.style.font.size.pt

            # Kandidat judul: heading style ATAU font ≥ 13pt
            if is_heading or max_size >= 12:
                candidates.append({
                    'text': text,
                    'size': max_size,
                    'italic': is_italic,
                    'heading': is_heading,
                })

            # Ambil maksimal 10 kandidat pertama (area awal dokumen)
            if len(candidates) >= 10:
                break

        if not candidates:
            return "", ""

        # Judul ID → kandidat pertama (biasanya lebih besar, tidak italic)
        title_id = candidates[0]['text']

        # Judul EN → sama dengan judul ID (diambil dari paragraf judul pertama)
        title_en = title_id

        return title_id, title_en

    except Exception as e:
        return "", ""

# --- INISIALISASI ENGINE ---
@st.cache_resource
def load_engines():
    e2 = DocxOptimizerEngine()
    e4 = CoverPageEngine()
    e5 = DaftarIsiEngine()
    e6 = PrakataPendahuluanEngine()
    e7 = InfoPendukungEngine()
    return e2, e4, e5, e6, e7

engine2, engine4, engine5, engine6, engine7 = load_engines()


# --- SIDEBAR ---
# ─────────────────────────────────────────────────────────────────────────────
# AUTO-LOAD KAMUS DARI SPREADSHEET (tanpa input manual)
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _load_kamus_from_sheet():
    d = CustomDictionary()
    count = d.load_defaults()
    return d, count

if 'custom_dict' not in st.session_state:
    _d, _n = _load_kamus_from_sheet()
    st.session_state['custom_dict'] = _d
    st.session_state['kamus_count'] = _n

_kamus = st.session_state.get('custom_dict')
_count = len(_kamus) if _kamus else 0

# --- HALAMAN UTAMA ---
st.markdown('<div class="section-header">🛠️ ISO Doc Master</div>', unsafe_allow_html=True)

if _count > 0:
    st.success(f"📚 Kamus aktif: **{_count} istilah**")
else:
    st.warning("⚠️ Kamus tidak aktif.")

import datetime
_tahun = str(datetime.date.today().year)

col_upload, col_judul = st.columns([2, 1])
with col_upload:
    uploaded_file = st.file_uploader("Upload File Word (.docx)", type=["docx"])
with col_judul:
    doc_title = st.text_input(
        "Judul Dokumen (untuk header)",
        value="SNI ISO 15118-1:2019",
        help="Muncul di header halaman. Nomor SNI cover mengikuti nilai ini."
    )

copyright_text = f"©BSN {_tahun}"
enable_headers = True

target_file = None
if uploaded_file:
    target_file = f"temp_{uploaded_file.name}"
    with open(target_file, "wb") as f:
        f.write(uploaded_file.getbuffer())

    enable_cover      = True
    enable_daftar_isi = True
    enable_prakata    = True
    enable_info_pendukung = True

    cover_settings = {
        "sni_number":   doc_title if doc_title else "SNI ISO XXXXX:20XX",
        "bsn_year":     _tahun,
        "ics_number":   "XX.XXX.XX",
        "ref_standard": "",
    }

    st.markdown(
        "✅ **Aktif otomatis:**  "
        "Cover (Engine 4) &nbsp;·&nbsp; "
        "Daftar Isi (Engine 5) &nbsp;·&nbsp; "
        "Prakata & Pendahuluan (Engine 6) &nbsp;·&nbsp; "
        "Info Pendukung Perumus (Engine 7)",
        unsafe_allow_html=True
    )


    if st.button("✨ Jalankan Optimasi", type="primary"):
        st.session_state['_action'] = 'optimasi'
        st.rerun()

# ── EKSEKUSI ─────────────────────────────────────────────────────────────────
LANG_OPTIONS = {
    "auto": "🔍 Deteksi Otomatis", "en": "🇬🇧 Inggris",
    "fr": "🇫🇷 Prancis", "de": "🇩🇪 Jerman", "es": "🇪🇸 Spanyol",
    "it": "🇮🇹 Italia", "nl": "🇳🇱 Belanda", "pt": "🇵🇹 Portugis",
    "ru": "🇷🇺 Rusia", "ja": "🇯🇵 Jepang", "zh-CN": "🇨🇳 Mandarin",
    "ko": "🇰🇷 Korea", "ar": "🇸🇦 Arab",
}

_action = st.session_state.get('_action')

def _run_optimasi(target_file, output_file, doc_title, copyright_text, cover_settings,
                  enable_cover, enable_daftar_isi, enable_prakata, enable_info_pendukung):
    """Jalankan pipeline optimasi, kembalikan (success, final_output_file, elapsed)."""
    import time as _time
    _t0 = _time.time()
    success, msg = engine2.process(
        target_file, output_file, ISO_FONT_NAME, ISO_FONT_SIZE,
        enable_headers=True, doc_title=doc_title, copyright_text=copyright_text
    )
    if not success:
        return False, None, _time.time() - _t0

    final_output_file = output_file

    if enable_cover and cover_settings:
        cover_output = f"cover_{os.path.basename(output_file)}"
        auto_title_id, auto_title_en = extract_titles_from_docx(output_file)
        ok_cover, _ = engine4.prepend_cover(
            input_docx=output_file, output_docx=cover_output,
            sni_number=cover_settings["sni_number"], bsn_year=cover_settings["bsn_year"],
            title_id=auto_title_id, title_en=auto_title_en,
            ref_standard=cover_settings["ref_standard"], ics_number=cover_settings["ics_number"],
        )
        if ok_cover:
            final_output_file = cover_output
            if enable_daftar_isi:
                di_output = f"di_{os.path.basename(cover_output)}"
                ok_di, _ = engine5.insert(
                    input_docx=final_output_file, output_docx=di_output,
                    doc_title=cover_settings["sni_number"],
                    copyright_text=f"©BSN {cover_settings['bsn_year']}",
                )
                if ok_di:
                    final_output_file = di_output
                    if enable_prakata:
                        pp_output = f"pp_{os.path.basename(di_output)}"
                        ref_std = cover_settings.get("ref_standard", "").strip() or \
                                  re.sub(r'^SNI\s+', '', cover_settings["sni_number"]).strip()
                        ok_pp, _ = engine6.insert(
                            input_docx=final_output_file, output_docx=pp_output,
                            sni_number=cover_settings["sni_number"],
                            title_id=auto_title_id or 'Judul Bahasa Indonesia',
                            title_en=auto_title_en or 'Title in English',
                            ref_standard=ref_std, bsn_year=cover_settings["bsn_year"],
                        )
                        if ok_pp:
                            final_output_file = pp_output

    if enable_info_pendukung:
        ip_output = f"ip_{os.path.basename(final_output_file)}"
        ok_ip, _ = engine7.append(input_docx=final_output_file, output_docx=ip_output)
        if ok_ip:
            final_output_file = ip_output

    return True, final_output_file, _time.time() - _t0

if _action == 'optimasi' and target_file:
    import time as _time
    output_file = f"opt_{os.path.basename(target_file)}"
    _timer_slot = st.empty()

    with st.spinner("Memproses dokumen..."):
        ok, final_output_file, elapsed = _run_optimasi(
            target_file, output_file, doc_title, copyright_text, cover_settings,
            enable_cover, enable_daftar_isi, enable_prakata, enable_info_pendukung
        )

    if ok:
        st.success("✅ Berhasil!")
        _timer_slot.caption(f"⏱️ Selesai dalam **{elapsed:.1f} detik**")
        st.session_state['_opt_file'] = final_output_file
        st.session_state['_action'] = 'pilih'
        st.rerun()
    else:
        st.error("❌ Gagal.")
        st.session_state['_action'] = None

if st.session_state.get('_action') == 'pilih' and st.session_state.get('_opt_file'):
    final_output_file = st.session_state['_opt_file']
    st.success("✅ Optimasi selesai!")
    st.divider()
    col_dl, col_tr_lang, col_tr_btn = st.columns([1, 2, 1])
    with col_dl:
        with open(final_output_file, "rb") as f:
            st.download_button("📥 Download", f, file_name="ISO_Fixed_Document.docx", use_container_width=True)
    with col_tr_lang:
        src_lang = st.selectbox(
            "Bahasa Sumber",
            options=list(LANG_OPTIONS.keys()),
            index=0,
            format_func=lambda x: LANG_OPTIONS[x],
            key="e9_src_lang",
        )
    with col_tr_btn:
        if st.button("🌍 Terjemahkan", type="primary", use_container_width=True):
            st.session_state['_action'] = 'terjemah'
            st.session_state['_src_lang'] = src_lang
            st.rerun()

if st.session_state.get('_action') == 'terjemah' and st.session_state.get('_opt_file'):
    from engine9 import DocxFinalTranslatorEngine
    import time as _time
    final_output_file = st.session_state['_opt_file']
    _src   = st.session_state.get('_src_lang', 'auto')
    tr_out = f"ID_{os.path.basename(final_output_file)}"
    _prog  = st.progress(0, text="Memulai terjemahan...")
    _timer_tr = st.empty()
    _t0_tr = _time.time()

    def _cb(pct, msg):
        _prog.progress(min(pct, 100), text=f"{pct}% — {msg[:60]}  ⏱️ {round(_time.time()-_t0_tr)}s")

    _engine9 = DocxFinalTranslatorEngine(
        source_lang=_src, target_lang='id',
        custom_dict=st.session_state.get('custom_dict'),
    )
    with st.spinner("Menerjemahkan..."):
        ok_tr, _ = _engine9.translate(
            input_docx=final_output_file, output_docx=tr_out,
            progress_callback=_cb, translate_headers=False,
        )
    _prog.empty()
    if ok_tr:
        st.success("✅ Terjemahan berhasil!")
        _timer_tr.caption(f"⏱️ Selesai dalam **{_time.time()-_t0_tr:.1f} detik**")
        with open(tr_out, "rb") as f:
            st.download_button("📥 Download Terjemahan", f,
                               file_name=f"ID_{os.path.basename(final_output_file)}",
                               key="e9_download")
    else:
        st.error("❌ Terjemahan gagal.")
    st.session_state['_action'] = 'pilih'





# --- FOOTER ---
st.markdown("---")
st.markdown(
    "<div style='text-align:center; color:gray; font-size:0.8rem;'>"
    "© 2026 ISO Doc Master. All rights reserved."
    "</div>",
    unsafe_allow_html=True
)