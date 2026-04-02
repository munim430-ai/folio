#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tag_template.py
Creates the tagged DOCX template (templates/mokwon.docx) from the unpacked
original DOCX in folio/unpacked_mokwon/.

Strategy:
- For empty cells (no text run): inject a marker run before </w:p>
- For cells with existing text: replace the <w:t> content with marker
- For consent checkboxes YES: replace □ with ☑
- For the underlined signature run in 4A4B848C: replace spaces with FOLIOSIG3
- For the date3 line paragraph: replace the underline run text with FOLIODATE3
- Pack all files into templates/mokwon.docx as a ZIP
"""

import re
import os
import sys
import shutil
import zipfile
import io

# Force UTF-8 output on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ─── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
UNPACKED_DIR = os.path.join(PROJECT_DIR, 'folio', 'unpacked_mokwon')
DOC_XML_PATH = os.path.join(UNPACKED_DIR, 'word', 'document.xml')
OUTPUT_DIR = os.path.join(PROJECT_DIR, 'templates')
OUTPUT_DOCX = os.path.join(OUTPUT_DIR, 'mokwon.docx')

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── Read document.xml ─────────────────────────────────────────────────────────
print(f'Reading {DOC_XML_PATH}')
with open(DOC_XML_PATH, 'r', encoding='utf-8') as f:
    xml = f.read()

# ─── Helper: inject marker run into empty paragraph ───────────────────────────
MARKER_RUN_TPL = (
    '<w:r>'
    '<w:rPr>'
    '<w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman"'
    ' w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
    '<w:color w:val="000000"/>'
    '</w:rPr>'
    '<w:t>{marker}</w:t>'
    '</w:r>'
)

def inject_marker_into_empty_cell(xml_str, para_id, marker):
    """Find the paragraph by paraId and inject marker run before </w:p>."""
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str

    # Find the </w:p> that closes THIS paragraph
    p_end = xml_str.find('</w:p>', idx)
    if p_end < 0:
        print(f'  WARNING: closing </w:p> not found for {para_id}')
        return xml_str

    run = MARKER_RUN_TPL.format(marker=marker)
    new_xml = xml_str[:p_end] + run + xml_str[p_end:]
    print(f'  Injected {marker} into paraId {para_id}')
    return new_xml


def replace_wt_in_para(xml_str, para_id, new_text):
    """Replace ALL <w:t>...</w:t> text nodes within a paragraph with new_text.
    Keeps the first <w:t>, replaces its content, removes subsequent ones."""
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str

    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')

    para_xml = xml_str[p_start:p_end_close]

    # Replace all <w:t ...>...</w:t> with a single one containing new_text
    # First replace: replace first occurrence
    replaced = False

    def replacer(m):
        nonlocal replaced
        if not replaced:
            replaced = True
            return f'<w:t>{new_text}</w:t>'
        else:
            return ''  # Remove subsequent w:t nodes

    new_para = re.sub(r'<w:t[^>]*>.*?</w:t>', replacer, para_xml, flags=re.DOTALL)

    result = xml_str[:p_start] + new_para + xml_str[p_end_close:]
    print(f'  Replaced text in paraId {para_id} -> {new_text!r}')
    return result


def replace_checkbox_in_para(xml_str, para_id, old_char, new_char):
    """Replace a specific character in all <w:t> nodes within a paragraph."""
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str

    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')

    para_xml = xml_str[p_start:p_end_close]
    new_para = para_xml.replace(old_char, new_char)

    result = xml_str[:p_start] + new_para + xml_str[p_end_close:]
    if old_char in para_xml:
        print(f'  Replaced checkbox in paraId {para_id}: {old_char!r} -> {new_char!r}')
    else:
        print(f'  WARNING: checkbox char {old_char!r} not found in paraId {para_id}')
    return result


def replace_underlined_run_in_para(xml_str, para_id, marker):
    """Replace the text of the underlined run inside a paragraph with marker."""
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str
    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')
    para_xml = xml_str[p_start:p_end_close]

    def replace_underlined(m):
        run_content = m.group(0)
        if '<w:u w:val="single"' in run_content:
            new_run = re.sub(
                r'<w:t[^>]*>.*?</w:t>',
                f'<w:t>{marker}</w:t>',
                run_content,
                flags=re.DOTALL
            )
            print(f'  Replaced underlined run in {para_id} with {marker}')
            return new_run
        return run_content

    new_para = re.sub(r'<w:r>.*?</w:r>', replace_underlined, para_xml, flags=re.DOTALL)
    return xml_str[:p_start] + new_para + xml_str[p_end_close:]


# ─── TABLE 0 — Admission Form ──────────────────────────────────────────────────
print('\n--- Table 0: Admission Form ---')

# Cell[24] T0: full_name in Korean name field (same value — Bangladeshi students use English name)
xml = inject_marker_into_empty_cell(xml, '6F047129', 'FOLIOFULL')

# Cell[28] T0: full_name in English name field (empty paragraph)
xml = inject_marker_into_empty_cell(xml, '7F8FFEB5', 'FOLIOFULL')

# Cell[32] T0: gender male checkbox - replace □ 남/Male
xml = replace_wt_in_para(xml, '1863850F', 'FOLIOGM\u00a0\ub0a8/Male')

# Cell[33] T0: gender female checkbox - replace □ 여/Female
xml = replace_wt_in_para(xml, '3CF448B9', 'FOLIOGF\u00a0\uc5ec/Female')

# Cell[37] T0: birth_year (empty)
xml = inject_marker_into_empty_cell(xml, '227A15BA', 'FOLIOBIRY')

# Cell[39] T0: birth_month (empty)
xml = inject_marker_into_empty_cell(xml, '5EE22C0B', 'FOLIOBIRMTH')

# Cell[41] T0: birth_day (empty)
xml = inject_marker_into_empty_cell(xml, '6888E24D', 'FOLIOBIRDAY')

# Cell[45] T0: nationality (empty)
xml = inject_marker_into_empty_cell(xml, '789239C0', 'FOLIONAT')

# Cell[47] T0: passport_number (empty)
xml = inject_marker_into_empty_cell(xml, '4F85F6BC', 'FOLIOPASS')

# Cell[50] T0: place_of_birth (empty)
xml = inject_marker_into_empty_cell(xml, '7D3D656C', 'FOLIOPOB')

# Cell[53] T0: address (empty)
xml = inject_marker_into_empty_cell(xml, '125D3336', 'FOLIOADDR')

# Cell[56] T0: phone (empty)
xml = inject_marker_into_empty_cell(xml, '634A6B36', 'FOLIOPHONE')

# Cell[60] T0: school_name (has bullet "ᆞ" - replace it)
xml = replace_wt_in_para(xml, '2F4E0883', 'FOLIOSCHOOL')

# Cell[64] T0: agency_name (has bullet "ᆞ" - replace it)
xml = replace_wt_in_para(xml, '19844208', 'FOLIOAGNCY')

# Cell[85] T0: full_name for signature line (visible cell above vMerge continuation)
xml = replace_underlined_run_in_para(xml, '1C17C11B', 'FOLIOSIG1')

# Cell[87] T0: date page 1 - replace ": _____________" with marker
xml = replace_wt_in_para(xml, '1BF401C1', 'FOLIODATE1')

# ─── Semester checkbox — always tick Summer ────────────────────────────────────
print('\n--- Semester: pre-tick Summer ---')

# Cell[6]: Summer checkbox (paraId 2C076C8F) — replace □ with ☑
xml = replace_checkbox_in_para(xml, '2C076C8F', '□', '☑')

# ─── TABLE 1 — Consent Form ────────────────────────────────────────────────────
print('\n--- Table 1: Consent Form ---')

# Cell[20] T1: consent 1 YES - always tick (□ -> ☑)
xml = replace_checkbox_in_para(xml, '5E6CF235', '□', '☑')

# Cell[21] T1: consent 1 NO - leave as □

# Cell[48] T1: consent 2 YES - always tick
xml = replace_checkbox_in_para(xml, '7BF3349F', '□', '☑')

# Cell[49] T1: consent 2 NO - leave as □

# Cell[84] T1: consent 3 YES - always tick
xml = replace_checkbox_in_para(xml, '30CD4F0B', '□', '☑')

# Cell[85] T1: consent 3 NO - leave as □

# Cell[94] T1: full_name for signature (visible cell above vMerge continuation)
xml = replace_underlined_run_in_para(xml, '7E19F301', 'FOLIOSIG2')

# Cell[96] T1: date page 2 - replace ": ______________" with marker
xml = replace_wt_in_para(xml, '01C617D9', 'FOLIODATE2')

# ─── TABLE 3 — Affidavit ───────────────────────────────────────────────────────
print('\n--- Table 3: Affidavit ---')

# Cell[2] T3: sponsor_name (empty)
xml = inject_marker_into_empty_cell(xml, '54C0A82E', 'FOLISPONM')

# Cell[4] T3: sponsor_address (empty)
xml = inject_marker_into_empty_cell(xml, '527FAFF0', 'FOLISPONAD')

# Cell[6] T3: sponsor_occupation (empty)
xml = inject_marker_into_empty_cell(xml, '6B05DE58', 'FOLISPONOCC')

# Cell[8] T3: sponsor_relation (empty)
xml = inject_marker_into_empty_cell(xml, '458777C7', 'FOLISPONREL')

# Cell[18] T3: sponsor_contact_HP (empty)
xml = inject_marker_into_empty_cell(xml, '4776947B', 'FOLISPONHP')

# ─── TABLE 4 — Consent checkbox row ───────────────────────────────────────────
print('\n--- Table 4: Consent checkbox ---')

# Cell[1] T4: tick first □ in "□ 예(Yes)         □ 아니오(No)"
# Replace only the FIRST □
def replace_first_checkbox(xml_str, para_id):
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str
    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')
    para_xml = xml_str[p_start:p_end_close]
    # Replace only first occurrence of □
    new_para = para_xml.replace('□', '☑', 1)
    result = xml_str[:p_start] + new_para + xml_str[p_end_close:]
    print(f'  Ticked first checkbox in paraId {para_id}')
    return result

xml = replace_first_checkbox(xml, '6345B7F5')

# ─── Free paragraph 4A4B848C — Signature line (Page 3) ────────────────────────
print('\n--- Free paragraph: Signature (4A4B848C) ---')

# The underlined run has spaces between " : " and "(SIGNATURE)"
# Replace the <w:t> content of the underlined run with FOLIOSIG3
def replace_underlined_run_in_para(xml_str, para_id, marker):
    idx = xml_str.find(para_id)
    if idx < 0:
        print(f'  WARNING: paraId {para_id} not found')
        return xml_str
    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')
    para_xml = xml_str[p_start:p_end_close]

    # Find the run with <w:u w:val="single"/>
    # Pattern: <w:r>...<w:u w:val="single"/>...</w:r> containing <w:t>...</w:t>
    def replace_underlined(m):
        run_content = m.group(0)
        if '<w:u w:val="single"' in run_content:
            # Replace the <w:t ...>...</w:t> within this run
            new_run = re.sub(
                r'<w:t[^>]*>.*?</w:t>',
                f'<w:t>{marker}</w:t>',
                run_content,
                flags=re.DOTALL
            )
            print(f'  Replaced underlined run in {para_id} with {marker}')
            return new_run
        return run_content

    new_para = re.sub(r'<w:r>.*?</w:r>', replace_underlined, para_xml, flags=re.DOTALL)
    return xml_str[:p_start] + new_para + xml_str[p_end_close:]

xml = replace_underlined_run_in_para(xml, '4A4B848C', 'FOLIOSIG3')

# ─── Free paragraph date3 — Date line (Page 3) ─────────────────────────────────
print('\n--- Free paragraph: Date3 ---')

# The date3 paragraph has paraId 4E28D8F5 and contains ":_______________/_________________/..."
# Replace the underlined run (or the run with underscores) with FOLIODATE3
def replace_date3_line(xml_str):
    # Find the paragraph containing the date line with underscores
    # Pattern: find ":___" or ": ___"
    patterns = [':_______________', ': _______________', ':________________']
    idx = -1
    for pat in patterns:
        idx = xml_str.find(pat)
        if idx >= 0:
            break
    if idx < 0:
        # Try to find by paraId
        idx = xml_str.find('4E28D8F5')
        if idx < 0:
            print('  WARNING: date3 paragraph not found')
            return xml_str

    p_start = xml_str.rfind('<w:p ', 0, idx)
    p_end_close = xml_str.find('</w:p>', idx) + len('</w:p>')
    para_xml = xml_str[p_start:p_end_close]

    # Replace all <w:t ...>...</w:t> nodes containing underscores or the colon pattern
    # We want to keep the label run (연월일(DATE)) and replace the rest
    # Strategy: find the run with underscores and replace its <w:t> with FOLIODATE3
    def replace_underscore_run(m):
        run_content = m.group(0)
        # Check if this run's <w:t> contains underscores or colon-underscore pattern
        t_match = re.search(r'<w:t[^>]*>(.*?)</w:t>', run_content, re.DOTALL)
        if t_match:
            t_text = t_match.group(1)
            if '_' in t_text or (': ' in t_text and '___' in t_text):
                new_run = re.sub(
                    r'<w:t[^>]*>.*?</w:t>',
                    '<w:t>FOLIODATE3</w:t>',
                    run_content,
                    flags=re.DOTALL
                )
                print('  Replaced date3 underscore run with FOLIODATE3')
                return new_run
        return run_content

    new_para = re.sub(r'<w:r>.*?</w:r>', replace_underscore_run, para_xml, flags=re.DOTALL)
    return xml_str[:p_start] + new_para + xml_str[p_end_close:]

xml = replace_date3_line(xml)

# ─── Write modified document.xml to a temp location ───────────────────────────
MODIFIED_XML_PATH = os.path.join(PROJECT_DIR, 'folio', 'tagged_document.xml')
with open(MODIFIED_XML_PATH, 'w', encoding='utf-8') as f:
    f.write(xml)
print(f'\nTagged document.xml written to {MODIFIED_XML_PATH}')

# ─── Pack into DOCX (ZIP) ──────────────────────────────────────────────────────
print(f'\nPacking {OUTPUT_DOCX}...')

with zipfile.ZipFile(OUTPUT_DOCX, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(UNPACKED_DIR):
        for file in files:
            file_abs = os.path.join(root, file)
            # Compute relative path within the zip (use forward slashes)
            rel_path = os.path.relpath(file_abs, UNPACKED_DIR).replace(os.sep, '/')

            if rel_path == 'word/document.xml':
                # Use the tagged version
                zf.write(MODIFIED_XML_PATH, rel_path)
            else:
                zf.write(file_abs, rel_path)

print(f'\nDone! Tagged template saved to: {OUTPUT_DOCX}')

# ─── Cleanup temp file ────────────────────────────────────────────────────────
os.remove(MODIFIED_XML_PATH)
print('Temp file cleaned up.')
