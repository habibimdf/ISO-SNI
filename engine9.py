"""
Engine9: DocxFinalTranslatorEngine + Custom Dictionary (Upgraded)
=================================================================
Upgrade:
  - Dual Title Sync: Mengambil Judul ID (Bold) & EN (Italic) dari Cover.
  - Sinkronisasi ke Kata Pengantar: Mengganti Judul ID & Judul EN
    dengan formatting Italic yang benar.
"""

import re
import copy
import time
import uuid
import traceback
import csv
import os

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import lxml.etree as etree


# ─────────────────────────────────────────────────────────────────────────────
# NAMESPACE & KONSTANTA
# ─────────────────────────────────────────────────────────────────────────────

_NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_W    = f'{{{_NS_W}}}'

_RE_PURE_NUMBER = re.compile(r'^[\d\s\.\,\:\;\-\(\)\[\]\/\\\+\=\*\%\&\^\$\#\@\!\"\'`~<>{}|_]+$')
_RE_COPYRIGHT = re.compile(r'©|BSN\s*\d{4}', re.IGNORECASE)

_SKIP_STYLES = {
    'caption', 'header', 'footer',
    'toc 1', 'toc 2', 'toc 3', 'toc 4', 'toc 5',
    'table of figures', 'footnote text', 'endnote text', 'macro text',
}
_BIBLIO_TITLE_STYLES = {'biblioti', 'bibliotitle', 'bibliography title'}
_BIBLIO_KEYWORDS_EXACT = {
    'bibliografi', 'bibliography',
    'daftar acuan', 'daftar pustaka', 'daftar referensi',
}
_ANNEX_STYLE_IDS = {'ANNEX', 'Annex', 'annex'}
_HEADING_STYLES_WITH_NUM = {
    'Heading1', 'Heading2', 'Heading3',
    'ANNEX', 'a2', 'a3',
    'Heading4', 'Heading5', 'Heading6',
}
_TRANSLATE_DELAY = 0.15
_EM_DASH = '—'

# ─────────────────────────────────────────────────────────────────────────────
# URL KAMUS ISTILAH — HARDCODED DARI GOOGLE SPREADSHEET
# Ganti URL di bawah dengan link Google Sheet kamus Anda.
# Sheet harus bisa diakses publik (Anyone with the link → Viewer).
# Format kolom: kolom pertama = istilah asing, kolom kedua = terjemahan Indonesia.
# ─────────────────────────────────────────────────────────────────────────────
KAMUS_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1BBPCMPwvbBk5LPdoDQwnjQzcPHv7_RDKENqeMsklF-8/edit?usp=sharing"


# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM DICTIONARY CLASS
# ─────────────────────────────────────────────────────────────────────────────

class CustomDictionary:
    def __init__(self):
        self._entries: dict[str, tuple[str, str]] = {}

    def add_term(self, source: str, target: str) -> None:
        s = source.strip()
        t = target.strip()
        if s and t:
            self._entries[s.lower()] = (s, t)

    def clear(self) -> None:
        self._entries.clear()

    def load_defaults(self) -> int:
        """Load kamus otomatis dari Google Spreadsheet yang sudah dikonfigurasi."""
        try:
            return self.load_from_google_sheet(KAMUS_SPREADSHEET_URL)
        except Exception:
            # Jika gagal (offline / URL belum diset), return 0 tanpa error
            return 0

    def load_from_csv(self, filepath: str, src_col: str = 'source', tgt_col: str = 'target', delimiter: str = ',', encoding: str = 'utf-8-sig') -> int:
        if not os.path.isfile(filepath): raise FileNotFoundError(f"File CSV tidak ditemukan: {filepath}")
        count = 0
        with open(filepath, newline='', encoding=encoding) as f:
            sample = f.read(1024); f.seek(0)
            has_header = csv.Sniffer().has_header(sample)
            reader = csv.DictReader(f, delimiter=delimiter) if has_header else csv.reader(f, delimiter=delimiter)
            for row in reader:
                if has_header:
                    src = row.get(src_col, row.get('source', '')).strip()
                    tgt = row.get(tgt_col, row.get('target', '')).strip()
                else:
                    row = list(row)
                    if len(row) < 2: continue
                    src, tgt = row[0].strip(), row[1].strip()
                if src and tgt:
                    self._entries[src.lower()] = (src, tgt)
                    count += 1
        return count

    def load_from_excel(self, filepath: str, sheet_name: str | int = 0, src_col: str = 'source', tgt_col: str = 'target') -> int:
        if not os.path.isfile(filepath): raise FileNotFoundError(f"File Excel tidak ditemukan: {filepath}")
        try: import pandas as pd
        except ImportError: raise ImportError("Jalankan: pip install pandas openpyxl")
        df = pd.read_excel(filepath, sheet_name=sheet_name, dtype=str).fillna('')
        col_src = _find_col(df.columns.tolist(), [src_col, 'source', 'Inggris'])
        col_tgt = _find_col(df.columns.tolist(), [tgt_col, 'target', 'Indonesia'])
        if col_src is None: col_src = df.columns[0]
        if col_tgt is None and len(df.columns) >= 2: col_tgt = df.columns[1]
        if col_tgt is None: raise ValueError("Kolom target tidak ditemukan.")
        count = 0
        for _, row in df.iterrows():
            src = str(row[col_src]).strip()
            tgt = str(row[col_tgt]).strip()
            if src and tgt and src.lower() not in ('nan', '') and tgt.lower() not in ('nan', ''):
                self._entries[src.lower()] = (src, tgt); count += 1
        return count

    def load_from_google_sheet(self, url: str, src_col: str = 'source', tgt_col: str = 'target', timeout: int = 15) -> int:
        try: import urllib.request, io
        except ImportError: raise ImportError("urllib tidak tersedia.")
        csv_url = _google_sheet_to_csv_url(url)
        try:
            req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=timeout) as resp: raw = resp.read().decode('utf-8-sig')
        except Exception as e: raise ConnectionError(f"Gagal mengambil data: {e}")
        f = io.StringIO(raw); reader = csv.DictReader(f); fieldnames = reader.fieldnames or []
        col_src = _find_col(fieldnames, [src_col, 'source', 'Inggris'])
        col_tgt = _find_col(fieldnames, [tgt_col, 'target', 'Indonesia'])
        if col_src is None and len(fieldnames) >= 1: col_src = fieldnames[0]
        if col_tgt is None and len(fieldnames) >= 2: col_tgt = fieldnames[1]
        if col_tgt is None: raise ValueError(f"Kolom target tidak ditemukan. Header: {fieldnames}")
        count = 0
        for row in reader:
            src = str(row.get(col_src, '')).strip()
            tgt = str(row.get(col_tgt, '')).strip()
            if src and tgt and src.lower() not in ('', 'nan') and tgt.lower() not in ('', 'nan'):
                self._entries[src.lower()] = (src, tgt); count += 1
        return count

    def __len__(self) -> int: return len(self._entries)
    def list_terms(self) -> list[tuple[str, str]]: return [(s, t) for _, (s, t) in sorted(self._entries.items())]

    def _apply_pre(self, text: str) -> tuple[str, dict]:
        if not self._entries: return text, {}
        token_map = {}; result = text
        sorted_entries = sorted(self._entries.items(), key=lambda x: len(x[0]), reverse=True)
        for src_lower, (src_orig, tgt) in sorted_entries:
            pattern = re.compile(r'(?<![A-Za-z0-9])' + re.escape(src_lower) + r'(?![A-Za-z0-9])', re.IGNORECASE)
            if pattern.search(result):
                token = f'@@TK_{uuid.uuid4().hex[:8].upper()}@@'
                token_map[token] = tgt
                result = pattern.sub(token, result)
        return result, token_map

    def _apply_post(self, translated: str, token_map: dict) -> str:
        result = translated
        for token, tgt in token_map.items(): result = re.sub(re.escape(token), tgt, result, flags=re.IGNORECASE)
        return result

def _find_col(columns: list, candidates: list) -> str | None:
    for c in candidates:
        if c in columns: return c
    return None

def _google_sheet_to_csv_url(url: str) -> str:
    url = url.strip()
    if 'output=csv' in url or 'format=csv' in url: return url
    if 'docs.google.com/spreadsheets' not in url: return url
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m: raise ValueError(f"Tidak dapat mengekstrak Sheet ID: {url}")
    sheet_id = m.group(1)
    gid_match = re.search(r'[#&?]gid=(\d+)', url)
    gid_param = f'&gid={gid_match.group(1)}' if gid_match else ''
    return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}'


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def _skip_text(text: str) -> bool:
    t = text.strip()
    if len(t) < 3: return True
    if _RE_PURE_NUMBER.fullmatch(t): return True
    if _RE_COPYRIGHT.search(t): return True
    return False

def _skip_paragraph(para, past_bibliography: bool = False) -> bool:
    if past_bibliography: return True
    if not para.text.strip(): return True
    for tag in [f'{_W}drawing', f'{_W}pict']:
        if para._element.find('.//' + tag) is not None: return True
    style_name = (para.style.name or '').lower()
    if any(style_name.startswith(s) for s in _SKIP_STYLES): return True
    return False

def _is_biblio_title_para(para) -> bool:
    style_id = ''
    try:
        pStyle = para._element.find(f'{_W}pPr/{_W}pStyle')
        if pStyle is not None: style_id = pStyle.get(f'{_W}val', '').lower()
    except Exception: pass
    if style_id in _BIBLIO_TITLE_STYLES: return True
    txt = para.text.strip().lower()
    return bool(txt) and not txt[0].isdigit() and txt in _BIBLIO_KEYWORDS_EXACT

def _get_para_style_id(para) -> str:
    pStyle = para._element.find(f'{_W}pPr/{_W}pStyle')
    return pStyle.get(f'{_W}val', '') if pStyle is not None else ''

def _get_style_id(el) -> str:
    pStyle = el.find(f'{_W}pPr/{_W}pStyle')
    return pStyle.get(f'{_W}val', '') if pStyle is not None else 'Normal'

def _has_inline_sectpr(para) -> bool:
    pPr = para._element.find(f'{_W}pPr')
    return pPr is not None and pPr.find(f'{_W}sectPr') is not None

def _all_runs_italic(para) -> bool:
    text_runs = [r for r in para.runs if r.text.strip()]
    if not text_runs: return False
    para_style_italic = False
    try:
        if para.style and para.style.font and para.style.font.italic: para_style_italic = True
    except Exception: pass
    for run in text_runs:
        italic = False
        if run.font.italic is True: italic = True
        if not italic:
            rPr = run._element.find(f'{_W}rPr')
            if rPr is not None:
                i_el = rPr.find(f'{_W}i')
                if i_el is not None:
                    val = i_el.get(f'{_W}val', 'true')
                    if val.lower() not in ('false', '0'): italic = True
        if not italic and para_style_italic:
            rPr = run._element.find(f'{_W}rPr')
            if rPr is not None:
                i_el = rPr.find(f'{_W}i')
                if i_el is not None:
                    val = i_el.get(f'{_W}val', 'true')
                    if val.lower() not in ('false', '0'): italic = True
                else: italic = True
            else: italic = True
        if not italic: return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 1: REKONSTRUKSI ANNEX (FIXED 12PT)
# ─────────────────────────────────────────────────────────────────────────────

def _fix_annex_style_para(para, annex_letter: str = None) -> None:
    sid = _get_para_style_id(para)
    if sid not in _ANNEX_STYLE_IDS: return
    # Ekstraksi teks aman — abaikan run yang hanya berisi <w:br/>
    full_text = ''.join(r.text for r in para.runs if r.text is not None).strip()
    if not full_text: return
    tag_norm = None; title_part = full_text
    tags_to_find = ['(informatif)', '(normatif)', '(informative)', '(normative)', '(informasi)']
    for t in tags_to_find:
        idx = full_text.lower().find(t)
        if idx != -1:
            if 'norm' in t.lower(): tag_norm = '(normatif)'
            else: tag_norm = '(informatif)'
            part_before = full_text[:idx]
            part_after  = full_text[idx + len(t):]
            title_part  = part_before + " " + part_after
            break
    # Ekstrak huruf dari "Annex X" atau "Lampiran X", fallback ke annex_letter
    annex_label = None
    m_annex = re.match(r'^(?:Annex|Lampiran)\s+([A-Z0-9\.]+)\s*', title_part, flags=re.IGNORECASE)
    if m_annex:
        annex_label = f'Lampiran {m_annex.group(1).upper()}'
        title_part  = title_part[m_annex.end():].strip()
    elif annex_letter:
        annex_label = f'Lampiran {annex_letter.upper()}'
        title_part  = re.sub(r'^(?:Annex|Lampiran)\s*\S*\s*', '', title_part, flags=re.IGNORECASE).strip()
    else:
        title_part  = re.sub(r'^(?:Annex|Lampiran)\s*\S*\s*', '', title_part, flags=re.IGNORECASE).strip()
    title_part = title_part.lstrip('\n').strip()
    pPr = para._element.find(f'{_W}pPr')
    if pPr is None:
        pPr = etree.SubElement(para._element, f'{_W}pPr')
    # Set numId=0 untuk matikan autonumbering "Annex %1" yang di-inherit dari style
    old_numPr = pPr.find(f'{_W}numPr')
    if old_numPr is not None: pPr.remove(old_numPr)
    pPr.insert(0, etree.fromstring(
        f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
    ))
    for child in list(para._element):
        if child is not pPr: para._element.remove(child)
    
    # FIXED: sz w:val="24" -> 12pt
    def make_arial_run(text, is_bold=False):
        esc = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        b_tag = '<w:b/><w:bCs/>' if is_bold else ''
        return etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>{b_tag}<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{esc}</w:t></w:r>')
    def make_br_run(): return etree.fromstring(f'<w:r xmlns:w="{_NS_W}"><w:br/></w:r>')
    
    new_runs = []
    if annex_label:
        new_runs.append(make_arial_run(annex_label, is_bold=True))
        new_runs.append(make_br_run())
    else:
        new_runs.append(make_br_run())
    if tag_norm: new_runs.append(make_arial_run(tag_norm, is_bold=False))
    if title_part:
        new_runs.append(make_br_run())
        new_runs.append(make_arial_run(title_part, is_bold=True))
    for run_el in new_runs: para._element.append(run_el)


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 2: EM DASH TO BULLETS
# ─────────────────────────────────────────────────────────────────────────────

def _get_or_create_emdash_numid(doc: Document) -> str:
    try:
        np = doc.part.numbering_part
        if np is None: return None
        nxml = np._element
        target_abstract_id = None
        for ab in nxml.findall(f'{_W}abstractNum'):
            for lvl in ab.findall(f'{_W}lvl'):
                txt_el = lvl.find(f'{_W}lvlText')
                if txt_el is not None and txt_el.get(f'{_W}val') == _EM_DASH:
                    target_abstract_id = ab.get(f'{_W}abstractNumId'); break
            if target_abstract_id: break
        if target_abstract_id:
            existing_nums = nxml.findall(f'{_W}num')
            max_num_id = max((int(n.get(f'{_W}numId', 0)) for n in existing_nums), default=0)
            new_num_id = str(max_num_id + 1)
            new_num = etree.fromstring(f'<w:num xmlns:w="{_NS_W}" w:numId="{new_num_id}"><w:abstractNumId w:val="{target_abstract_id}"/></w:num>')
            nxml.append(new_num)
            return new_num_id
        existing_abstracts = nxml.findall(f'{_W}abstractNum')
        max_abstract_id = max((int(a.get(f'{_W}abstractNumId', 0)) for a in existing_abstracts), default=0)
        new_abstract_id = str(max_abstract_id + 1)
        existing_nums = nxml.findall(f'{_W}num')
        max_num_id = max((int(n.get(f'{_W}numId', 0)) for n in existing_nums), default=0)
        new_num_id = str(max_num_id + 1)
        abstract_xml = f'<w:abstractNum xmlns:w="{_NS_W}" w:abstractNumId="{new_abstract_id}"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="{_EM_DASH}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
        nxml.insert(0, etree.fromstring(abstract_xml))
        num_xml = f'<w:num xmlns:w="{_NS_W}" w:numId="{new_num_id}"><w:abstractNumId w:val="{new_abstract_id}"/></w:num>'
        nxml.append(etree.fromstring(num_xml))
        return new_num_id
    except Exception as e: print(f"[Engine9] Error numbering: {e}"); return None

def _convert_emdash_to_bullets(doc: Document) -> None:
    num_id = _get_or_create_emdash_numid(doc)
    if not num_id: return
    for para in doc.paragraphs:
        sid = _get_para_style_id(para)
        if sid in _ANNEX_STYLE_IDS or sid.startswith('Heading'): continue
        text = para.text.strip()
        if text.startswith(_EM_DASH):
            for r in para.runs:
                if _EM_DASH in r.text: r.text = r.text.replace(_EM_DASH, "", 1).lstrip(); break
            pPr = para._element.find(f'{_W}pPr')
            if pPr is None: pPr = etree.SubElement(para._element, f'{_W}pPr')
            old_num = pPr.find(f'{_W}numPr')
            if old_num is not None: pPr.remove(old_num)
            pPr.insert(0, etree.fromstring(f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'))
            old_ind = pPr.find(f'{_W}ind')
            if old_ind is not None: pPr.remove(old_ind)
            pPr.insert(1, etree.fromstring(f'<w:ind xmlns:w="{_NS_W}" w:left="360" w:hanging="360"/>'))


# ─────────────────────────────────────────────────────────────────────────────
# FITUR 3: FIX NOTE / CATATAN FORMATTING
# ─────────────────────────────────────────────────────────────────────────────

def _fix_note_para(para) -> None:
    """
    Split paragraf Note/Catatan menjadi dua run:
      - Label (CATATAN, NOTE, CATATAN 1:, dll.) → Bold
      - Isi teks                                 → Not Bold
    """
    full_text = para.text.strip()
    if not full_text:
        return
    txt_lower = full_text.lower()
    if not (txt_lower.startswith('note') or txt_lower.startswith('catatan')):
        return

    m = re.match(
        r'^((?:NOTE|CATATAN)\s*\d*\s*(?:to\s+entry\s*)?:?)\s*',
        full_text, re.IGNORECASE
    )
    if m:
        bold_part   = m.group(1).rstrip()
        normal_part = full_text[m.end():]
    elif ':' in full_text:
        parts       = full_text.split(':', 1)
        bold_part   = parts[0] + ':'
        normal_part = parts[1] if len(parts) > 1 else ''
    else:
        words       = full_text.split(None, 1)
        bold_part   = words[0]
        normal_part = words[1] if len(words) > 1 else ''

    if not normal_part.strip():
        return

    font_name = 'Arial'
    font_size = None
    for run in para.runs:
        if run.text.strip():
            if run.font.name:  font_name = run.font.name
            if run.font.size:  font_size = run.font.size
            break

    for run in list(para.runs):
        run._element.getparent().remove(run._element)

    run_bold           = para.add_run(bold_part)
    run_bold.bold      = True
    run_bold.font.name = font_name
    if font_size: run_bold.font.size = font_size

    run_normal           = para.add_run(' ' + normal_part.lstrip())
    run_normal.bold      = False
    run_normal.font.name = font_name
    if font_size: run_normal.font.size = font_size


def _fix_all_notes(doc: Document) -> None:
    """Post-processing: terapkan _fix_note_para ke seluruh body dan tabel."""
    for para in doc.paragraphs:
        _fix_note_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _fix_note_para(para)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: BIBLIOGRAFI & AUTONUMBERING
# ─────────────────────────────────────────────────────────────────────────────

def _el_text(el) -> str: return ''.join(t.text or '' for t in el.findall(f'.//{_W}t')).strip()
def _is_bibliography_el(el) -> bool:
    if el.tag != f'{_W}p': return False
    sid = _get_style_id(el).lower()
    if sid in _BIBLIO_TITLE_STYLES: return True
    text = _el_text(el).lower().strip()
    if not text or text[0].isdigit(): return False
    return text in _BIBLIO_KEYWORDS_EXACT

def _find_bib_index(body_els: list) -> int:
    for i, el in enumerate(body_els):
        if _is_bibliography_el(el): return i
    return -1

def _find_content_start_before_bib(body_els: list, bib_idx: int) -> int:
    search_limit = bib_idx if bib_idx >= 0 else len(body_els)
    last_sectpr  = -1
    for i, el in enumerate(body_els):
        if i >= search_limit: break
        if el.tag != f'{_W}p': continue
        pPr = el.find(f'{_W}pPr')
        if pPr is not None and pPr.find(f'{_W}sectPr') is not None: last_sectpr = i
    return last_sectpr + 1

def _read_style_numpr(doc: Document) -> dict:
    result = {}
    try:
        styles_xml = doc.part.styles._element
        for style_el in styles_xml.findall(f'{_W}style'):
            sid = style_el.get(f'{_W}styleId', '')
            if sid not in _HEADING_STYLES_WITH_NUM: continue
            pPr = style_el.find(f'{_W}pPr')
            if pPr is None: continue
            numPr = pPr.find(f'{_W}numPr')
            if numPr is None: continue
            numId_el = numPr.find(f'{_W}numId')
            ilvl_el  = numPr.find(f'{_W}ilvl')
            nid  = numId_el.get(f'{_W}val', '0') if numId_el is not None else '0'
            ilvl = ilvl_el.get(f'{_W}val',  '0') if ilvl_el  is not None else '0'
            result[sid] = {'numId': nid, 'ilvl': ilvl}
    except Exception: pass
    return result

def _create_restart_numids(doc: Document, style_numpr: dict) -> dict:
    remapping = {}
    if not style_numpr: return remapping
    try:
        np   = doc.part.numbering_part
        nxml = np._element
        unique_numids = {info['numId'] for info in style_numpr.values()}
        existing      = nxml.findall(f'{_W}num')
        max_id        = max((int(n.get(f'{_W}numId', 0)) for n in existing), default=0)
        for old_nid in unique_numids:
            if old_nid == '0': continue
            old_num_el = None
            for n in existing:
                if n.get(f'{_W}numId') == old_nid: old_num_el = n; break
            if old_num_el is None: continue
            max_id += 1
            new_nid    = str(max_id)
            new_num_el = copy.deepcopy(old_num_el)
            new_num_el.set(f'{_W}numId', new_nid)
            for lo in new_num_el.findall(f'{_W}lvlOverride'): new_num_el.remove(lo)
            override_el = etree.fromstring(f'<w:lvlOverride xmlns:w="{_NS_W}" w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>')
            new_num_el.append(override_el)
            nxml.append(new_num_el)
            remapping[old_nid] = new_nid
    except Exception: pass
    return remapping

def _apply_numpr_restart_to_headings(elements: list, style_numpr: dict, remapping: dict) -> None:
    if not remapping: return
    for el in elements:
        if el.tag != f'{_W}p': continue
        sid = _get_style_id(el)
        if sid not in style_numpr: continue
        info    = style_numpr[sid]
        old_nid = info['numId']
        ilvl    = info['ilvl']
        new_nid = remapping.get(old_nid)
        if not new_nid: continue
        pPr = el.find(f'{_W}pPr')
        if pPr is None: pPr = etree.SubElement(el, f'{_W}pPr'); el.insert(0, pPr)
        old_numPr = pPr.find(f'{_W}numPr')
        if old_numPr is not None: pPr.remove(old_numPr)
        pPr.insert(0, etree.fromstring(f'<w:numPr xmlns:w="{_NS_W}"><w:ilvl w:val="{ilvl}"/><w:numId w:val="{new_nid}"/></w:numPr>'))


# ─────────────────────────────────────────────────────────────────────────────
# CORE: INSERT ORIGINAL
# ─────────────────────────────────────────────────────────────────────────────

def _page_break_para() -> etree._Element:
    return etree.fromstring(f'<w:p xmlns:w="{_NS_W}"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>')

def _intro_heading() -> etree._Element:
    return etree.fromstring(f'<w:p xmlns:w="{_NS_W}"><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>Introduction</w:t></w:r></w:p>')

def _empty_para() -> etree._Element:
    return etree.fromstring(f'<w:p xmlns:w="{_NS_W}"><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr></w:p>')

def _insert_original_before_bib(translated_doc: Document, orig_body_els: list, progress_callback=None) -> None:
    bib_in_orig   = _find_bib_index(orig_body_els)
    content_start = _find_content_start_before_bib(orig_body_els, bib_in_orig)
    content_end   = bib_in_orig if bib_in_orig >= 0 else len(orig_body_els)
    if bib_in_orig < 0 and content_end > content_start:
        if orig_body_els[content_end - 1].tag == f'{_W}sectPr': content_end -= 1
    els_to_insert = orig_body_els[content_start:content_end]
    if not els_to_insert:
        _notify(progress_callback, 50, "⚠️ Tidak ada konten asli yang bisa disisipkan.")
        return
    _notify(progress_callback, 10, f"Ditemukan {len(els_to_insert)} elemen.")
    style_numpr = _read_style_numpr(translated_doc)
    _notify(progress_callback, 20, f"Styles: {list(style_numpr.keys())}")
    remapping = _create_restart_numids(translated_doc, style_numpr)
    _notify(progress_callback, 30, f"Remapping: {remapping}")
    new_content_els = [copy.deepcopy(el) for el in els_to_insert]
    _apply_numpr_restart_to_headings(new_content_els, style_numpr, remapping)
    _notify(progress_callback, 40, "Reset heading numbers.")
    trans_body     = translated_doc.element.body
    trans_children = list(trans_body)
    bib_idx_trans  = _find_bib_index(trans_children)
    if bib_idx_trans < 0:
        bib_idx_trans = len(trans_children)
        for i in range(len(trans_children) - 1, -1, -1):
            if trans_children[i].tag == f'{_W}sectPr': bib_idx_trans = i; break
        _notify(progress_callback, 45, "Bib tidak ditemukan → akhir.")
    else:
        _notify(progress_callback, 45, f"Bib posisi [{bib_idx_trans}].")
    new_els = [_page_break_para(), _intro_heading(), _page_break_para()] + new_content_els + [_empty_para()]
    total = len(new_els)
    for offset, el in enumerate(new_els):
        trans_body.insert(bib_idx_trans + offset, el)
        if progress_callback and offset % 20 == 0:
            pct = 45 + int(offset / max(total, 1) * 55)
            progress_callback(pct, f"Menyisipkan... ({offset}/{total})")
    _notify(progress_callback, 100, f"✅ Selesai ({len(els_to_insert)} item).")

def _notify(cb, pct: int, msg: str) -> None:
    if cb:
        try: cb(pct, msg)
        except Exception: pass


# ─────────────────────────────────────────────────────────────────────────────
# TRANSLATOR WRAPPER
# ─────────────────────────────────────────────────────────────────────────────

class _Translator:
    def __init__(self, source: str = 'auto', target: str = 'id', custom_dict: CustomDictionary | None = None):
        try: from deep_translator import GoogleTranslator
        except ImportError: raise ImportError("Jalankan: pip install deep-translator")
        self._cls = GoogleTranslator
        self.source = source; self.target = target; self.custom_dict = custom_dict

    def translate_one(self, text: str) -> str:
        t = text.strip()
        if not t or _skip_text(t): return text
        token_map = {}
        if self.custom_dict and len(self.custom_dict) > 0:
            t, token_map = self.custom_dict._apply_pre(t)
        try:
            result = self._cls(source=self.source, target=self.target).translate(t)
            if not result: result = t
        except Exception:
            time.sleep(0.8)
            try:
                result = self._cls(source=self.source, target=self.target).translate(t)
                if not result: result = t
            except Exception: result = t
        if token_map: result = self.custom_dict._apply_post(result, token_map)
        return result

def _match_capitalization(original: str, translated: str) -> str:
    orig = original.strip(); tran = translated.strip()
    if not orig or not tran: return translated
    letters = [c for c in orig if c.isalpha()]
    if not letters: return translated
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio >= 0.8: return tran.upper()
    if orig[0].isupper(): return tran[0].upper() + tran[1:] if len(tran) > 1 else tran.upper()
    return tran

def _translate_para(para, tr, past_bibliography: bool = False) -> None:
    if _skip_paragraph(para, past_bibliography): return
    text_runs = [(i, r) for i, r in enumerate(para.runs) if r.text and r.text.strip()]
    if not text_runs: return
    combined = ''.join(r.text for _, r in text_runs)
    if _skip_text(combined): return
    translated = tr.translate_one(combined.strip())
    time.sleep(_TRANSLATE_DELAY)
    if not translated or translated == combined: return
    translated = _match_capitalization(combined.strip(), translated)
    _, first_run = text_runs[0]
    first_run.text = translated
    for _, run in text_runs[1:]: run.text = ''

def _translate_table(table, tr) -> None:
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs: _translate_para(para, tr)

def _translate_hf(hf_part, tr) -> None:
    if hf_part is None: return
    try:
        for para in hf_part.paragraphs:
            if _RE_COPYRIGHT.search(para.text or ''): continue
            _translate_para(para, tr)
        for table in hf_part.tables: _translate_table(table, tr)
    except Exception: pass


# ─────────────────────────────────────────────────────────────────────────────
# SINKRONISASI JUDUL COVER & KATA PENGANTAR (UPGRADED)
# ─────────────────────────────────────────────────────────────────────────────

def _extract_cover_titles(doc: Document) -> tuple[str, str]:
    """
    Mengekstrak Judul Indonesia dan Judul Inggris dari halaman Cover.
    Return: (id_title, en_title)
    """
    id_title = ""; en_title = ""
    
    for para in doc.paragraphs:
        if _has_inline_sectpr(para): break # Stop di section break pertama
        
        text = para.text.strip()
        if not text: continue
        
        is_bold = False; is_italic = False; max_size = 0
        
        for run in para.runs:
            if run.text.strip():
                if run.font.bold: is_bold = True
                if run.font.italic: is_italic = True
                sz = run.font.size
                if sz and sz.pt > max_size: max_size = sz.pt
        
        # 1. Cari Judul ID (Bold, Non-Italic) - Biasanya ukuran paling besar
        if not id_title:
            if max_size >= 16 and is_bold and not is_italic:
                id_title = text
                continue
        
        # 2. Cari Judul EN (Italic) - Biasanya setelah ID
        if id_title and not en_title:
            # Engine 4 membuat EN dengan size 16pt, Bold Italic.
            # Kriteria: Italic adalah pembeda utama.
            if max_size >= 14 and is_italic:
                en_title = text
    
    return id_title, en_title


def _sync_body_title(doc: Document, cover_id: str) -> bool:
    """
    Mencari judul dokumen di awal konten body (sebelum Heading 1 pertama),
    lalu menggantinya dengan judul Indonesia yang sudah diterjemahkan dari cover.

    Strategi deteksi (prioritas urutan):
    1. Style 'Main Title 2', 'boxedTitle', atau varian serupa
    2. Paragraf bold+centered setelah section break TERAKHIR sebelum Heading 1
    """
    if not cover_id:
        return False

    _BODY_TITLE_STYLES = {'main title 2', 'main title2', 'maintitle2',
                          'boxedtitle', 'boxed title', 'body title'}
    _RE_HEADING1 = re.compile(r'^\d+\s+\S')

    # ── Pass 1: cari berdasarkan style name ──────────────────────────────────
    for para in doc.paragraphs:
        style_name = (para.style.name or '').lower().strip() if para.style else ''
        if style_name not in _BODY_TITLE_STYLES:
            continue
        txt = para.text.strip()
        if not txt:
            continue
        _replace_para_text(para, cover_id)
        return True

    # ── Pass 2: fallback — bold+centered setelah section break terakhir ──────
    # Temukan dulu indeks section break terakhir sebelum Heading 1
    last_sec_idx  = -1
    heading1_idx  = -1
    paragraphs    = doc.paragraphs

    for i, para in enumerate(paragraphs):
        if _has_inline_sectpr(para):
            last_sec_idx = i
        txt = para.text.strip()
        if txt and _RE_HEADING1.match(txt) and (para.style and 'heading' in (para.style.name or '').lower()):
            heading1_idx = i
            break

    if last_sec_idx < 0:
        return False

    limit = heading1_idx if heading1_idx > 0 else len(paragraphs)
    for i in range(last_sec_idx + 1, limit):
        para = paragraphs[i]
        txt  = para.text.strip()
        if not txt:
            continue
        is_bold     = any(r.bold for r in para.runs if r.text.strip())
        is_centered = (para.alignment == WD_ALIGN_PARAGRAPH.CENTER)
        if is_bold and is_centered:
            _replace_para_text(para, cover_id)
            return True

    return False


def _replace_para_text(para, new_text: str) -> None:
    """Ganti seluruh teks paragraf dengan new_text, pertahankan font dari run pertama."""
    font_name = 'Arial'
    font_size = None
    for run in para.runs:
        if run.text.strip():
            if run.font.name:  font_name = run.font.name
            if run.font.size:  font_size = run.font.size
            break

    for run in list(para.runs):
        run._element.getparent().remove(run._element)

    new_run           = para.add_run(new_text)
    new_run.bold      = True
    new_run.font.name = font_name
    if font_size:
        new_run.font.size = font_size
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _sync_foreword_title(doc: Document, cover_id: str, cover_en: str) -> None:
    """
    Menyinkronkan judul ID dan EN di Kata Pengantar.
    Struktur Paragraf (Engine 6):
    "SNI [No], [Judul ID], merupakan standar ... identik dari [Ref], [Judul EN], dengan metode..."
    """
    if not cover_id: return

    # Regex untuk memecah paragraf menjadi 5 bagian:
    # 1. Awal: "SNI [No], "
    # 2. Judul ID Lama
    # 3. Tengah: ", merupakan standar ... identik dari [Ref], "
    # 4. Judul EN Lama
    # 5. Akhir: ", dengan metode ..."
    
    pattern = re.compile(
        r'^(SNI [^,]+,\s*)'                 # Group 1: Awal
        r'(.*?)'                            # Group 2: Judul ID Lama
        r'(,\s*merupakan standar.*?identik dari [^,]+,\s*)' # Group 3: Tengah (termasuk Ref Standard)
        r'(.*?)'                            # Group 4: Judul EN Lama
        r'(,\s*dengan metode.*)',           # Group 5: Akhir
        re.IGNORECASE | re.DOTALL
    )

    for para in doc.paragraphs:
        text = para.text.strip()
        match = pattern.match(text)
        
        if match:
            pPr = para._element.pPr
            if pPr is None: continue
            
            # Clear old runs
            for run in para.runs:
                run._element.getparent().remove(run._element)
            
            # Reconstruct Paragraph
            
            # 1. SNI Number
            run_1 = para.add_run(match.group(1))
            run_1.font.name = "Arial"
            run_1.font.size = Pt(11)
            run_1.italic = False
            
            # 2. Judul ID (dari Cover) - ITALIC
            run_2 = para.add_run(cover_id)
            run_2.font.name = "Arial"
            run_2.font.size = Pt(11)
            run_2.italic = True  # Sesuai standar SNI
            
            # 3. Tengah (Merupakan standar... identik dari...)
            run_3 = para.add_run(match.group(3))
            run_3.font.name = "Arial"
            run_3.font.size = Pt(11)
            run_3.italic = False
            
            # 4. Judul EN (dari Cover) - ITALIC (jika ada)
            if cover_en:
                run_4 = para.add_run(cover_en)
                run_4.font.name = "Arial"
                run_4.font.size = Pt(11)
                run_4.italic = True # Sesuai permintaan
            
            # 5. Akhir (Dengan metode...)
            run_5 = para.add_run(match.group(5))
            run_5.font.name = "Arial"
            run_5.font.size = Pt(11)
            run_5.italic = False
            
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            break # Selesai


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENGINE CLASS
# ─────────────────────────────────────────────────────────────────────────────

class DocxFinalTranslatorEngine:
    def __init__(self, source_lang: str = 'auto', target_lang: str = 'id', custom_dict: CustomDictionary | None = None):
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.custom_dict = custom_dict

    def set_dictionary(self, d: CustomDictionary) -> None: self.custom_dict = d
    def get_dictionary(self) -> CustomDictionary:
        if self.custom_dict is None: self.custom_dict = CustomDictionary()
        return self.custom_dict

    def translate(self, input_docx: str, output_docx: str, progress_callback=None, translate_headers: bool = False) -> tuple[bool, str]:
        try:
            dict_info = f"{len(self.custom_dict)} istilah kamus" if self.custom_dict else "tanpa kamus"
            _notify(progress_callback, 2, f"Snapshot dokumen... ({dict_info})")
            doc_orig      = Document(input_docx)
            orig_body_els = [copy.deepcopy(el) for el in doc_orig.element.body]
            del doc_orig

            _notify(progress_callback, 5, "Init translator...")
            tr  = _Translator(source=self.source_lang, target=self.target_lang, custom_dict=self.custom_dict)
            doc = Document(input_docx)

            body     = doc.element.body
            para_map = {p._element: p for p in doc.paragraphs}
            tbl_map  = {t._element: t for t in doc.tables}

            items    = []
            sec_brks = 0
            COVER_END = 1

            for child in body:
                if child in para_map:
                    in_cover = (sec_brks < COVER_END)
                    items.append(('para', para_map[child], in_cover))
                    if _has_inline_sectpr(para_map[child]): sec_brks += 1
                elif child in tbl_map:
                    in_cover = (sec_brks < COVER_END)
                    items.append(('table', tbl_map[child], in_cover))

            total             = len(items)
            done              = 0
            past_bibliography = False
            annex_counter     = 0  # counter untuk Lampiran A, B, C, ...

            for kind, obj, in_cover in items:
                done += 1
                pct = 5 + int(done / max(total, 1) * 60)

                if kind == 'para':
                    para           = obj
                    is_bib_heading = _is_biblio_title_para(para)
                    is_annex       = _get_para_style_id(para) in _ANNEX_STYLE_IDS

                    if in_cover and _all_runs_italic(para):
                        _notify(progress_callback, pct, "[Cover-italic] skip")
                    elif is_bib_heading:
                        _translate_para(para, tr, past_bibliography=False)
                        past_bibliography = True
                        _notify(progress_callback, pct, "[Bib-heading]")
                    elif is_annex and not past_bibliography:
                        _translate_para(para, tr, past_bibliography=False)
                        annex_letter = chr(ord('A') + annex_counter)
                        _fix_annex_style_para(para, annex_letter=annex_letter)
                        annex_counter += 1
                        _notify(progress_callback, pct, f"[ANNEX] Lampiran {annex_letter} fixed")
                    elif not _skip_paragraph(para, past_bibliography):
                        _translate_para(para, tr, past_bibliography=False)

                elif kind == 'table':
                    if not past_bibliography: _translate_table(obj, tr)

            _notify(progress_callback, 66, "Formatting em-dash bullets...")
            _convert_emdash_to_bullets(doc)

            _notify(progress_callback, 67, "Fixing Note/Catatan formatting...")
            _fix_all_notes(doc)

            if translate_headers:
                _notify(progress_callback, 70, "Translating headers/footers...")
                for section in doc.sections:
                    for hf in [section.header, section.footer, section.even_page_header, section.even_page_footer, section.first_page_header, section.first_page_footer]:
                        _translate_hf(hf, tr)

            _notify(progress_callback, 75, "Inserting original content...")
            def _insert_cb(pct_inner, msg): _notify(progress_callback, 75 + int(pct_inner * 0.20), msg)

            _insert_original_before_bib(translated_doc=doc, orig_body_els=orig_body_els, progress_callback=_insert_cb)

            # ── SINKRONISASI JUDUL COVER (UPDATED) ────────────────────────────
            _notify(progress_callback, 96, "Menyinkronkan judul Cover dengan Kata Pengantar...")
            
            # 1. Ambil KEDUA judul dari Cover
            final_id_title, final_en_title = _extract_cover_titles(doc)
            
            # 2. Sinkronkan ke Kata Pengantar
            if final_id_title:
                _sync_foreword_title(doc, final_id_title, final_en_title)
                msg = f"Judul ID: {final_id_title[:30]}..."
                if final_en_title: msg += f" | Judul EN: {final_en_title[:30]}..."
                _notify(progress_callback, 97, f"Judul disinkronkan: {msg}")

                # 3. Sinkronkan judul body (bold centered sebelum Heading 1)
                body_synced = _sync_body_title(doc, final_id_title)
                _notify(progress_callback, 98, f"Judul body {'diperbarui' if body_synced else 'tidak ditemukan'}.")
            else:
                _notify(progress_callback, 97, "Judul cover tidak ditemukan, skip sinkronisasi.")
            # ─────────────────────────────────────────────────────────────

            _notify(progress_callback, 97, "Saving...")
            doc.save(output_docx)
            _notify(progress_callback, 100, "✅ Done!")
            return True, output_docx

        except ImportError as e: return False, f"Dependensi tidak ditemukan: {e}"
        except Exception as e: return False, f"Engine9 Error: {str(e)}\n{traceback.format_exc()}"