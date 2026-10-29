#!/usr/bin/env python3
"""
CMS LCD / Article Processor
Extracts structured information from all Article PDFs in a directory:
- Outputs a single Excel file with 4 sheets:
    1. Covered HCPCS-CPT Codes
    2. Covered ICD Codes
    3. HCPCS-CPT to ICD Mapping
    4. Non-Covered ICD Codes
- Outputs a single JSON file with all metadata:
    Article ID, Original Effective Date, Revised Effective Date,
    Revision End Date, Retirement Date, Contractor Information, Article Status

Pure rule-based parsing with PyMuPDF (no LLM).
"""

import os
import re
import sys
import glob
import json
import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
#  Text & Layout Helpers
# ---------------------------------------------------------------------------

def clean_line(text):
    """Normalize whitespace and strip text."""
    if not text:
        return ""
    return " ".join(text.split()).strip()


def is_header_or_footer(line):
    """Detect page header / footer / pagination lines to ignore."""
    line = line.strip()
    if not line:
        return True
    # Date-time stamps like "9/4/26, 4:25 PM"
    if re.match(r'^\d{1,2}/\d{1,2}/\d{2,4},\s+\d{1,2}:\d{2}\s+(AM|PM)$', line):
        return True
    # Article title headers like "Article - Billing and Coding: ..."
    if line.startswith("Article -") and ("(" in line and ")" in line):
        return True
    if line.startswith("Article - Billing and Coding:"):
        return True
    # URLs like cms.gov...
    if "cms.gov/medicare-coverage-database" in line:
        return True
    # Page numbers like "1/29"
    if re.match(r'^\d+/\d+$', line):
        return True
    return False


def get_document_lines(doc):
    """Extract all text lines from document, filtering out page headers/footers."""
    lines = []
    for pno in range(len(doc)):
        text = doc[pno].get_text("text")
        for line in text.split("\n"):
            if not is_header_or_footer(line):
                lines.append(line.rstrip())
    return lines


# ---------------------------------------------------------------------------
#  Code Pattern Matching Helpers
# ---------------------------------------------------------------------------

DASH_CHARS = r"[-–—−]"
CODE_TOKEN = r"(?:[A-TV-Z][0-9][0-9A-Z]?(?:\.[0-9A-Z]{1,7})?|\d{5}|[A-Z]\d{4}|[A-Z0-9]{5}|XX\d+|[0-9]{3,5}[xX]?)[\*\†\‡\#\^]*"

RE_SINGLE_CODE = re.compile(rf"^{CODE_TOKEN}$")
RE_SAME_LINE_RANGE = re.compile(rf"^({CODE_TOKEN})\s*{DASH_CHARS}\s*({CODE_TOKEN})$")
RE_SPLIT_START = re.compile(rf"^({CODE_TOKEN})\s*{DASH_CHARS}$")
RE_SPLIT_DASH_END = re.compile(rf"^{DASH_CHARS}\s*({CODE_TOKEN})$")


def match_code(lines, i):
    """
    Check if line(s) starting at index i represent a code or code range.
    Handles:
      - Single codes (e.g. C33, 44388, J9035)
      - Single-line ranges (e.g. C46.0 - C46.4)
      - Multi-line ranges:
          Line 1: C34.00 -
          Line 2: C34.92
        or
          Line 1: C34.00
          Line 2: - C34.92
    Returns: (code_string, lines_consumed) or (None, 0).
    """
    if i >= len(lines):
        return None, 0
    line = lines[i].strip()

    # 1. Same-line range: e.g. "C46.0 - C46.4" or "C34.00 - C34.92"
    m_range = RE_SAME_LINE_RANGE.match(line)
    if m_range:
        return f"{m_range.group(1)} - {m_range.group(2)}", 1

    # 2. Split-line range where line 1 has trailing dash: e.g. "C34.00 -" followed by "C34.92"
    m_start = RE_SPLIT_START.match(line)
    if m_start:
        j = i + 1
        while j < len(lines) and (not lines[j].strip() or lines[j].strip() in ("Code", "Description")):
            j += 1
        if j < len(lines):
            next_l = lines[j].strip()
            m_end = RE_SINGLE_CODE.match(next_l)
            if m_end:
                return f"{m_start.group(1)} - {m_end.group(0)}", (j - i + 1)

    # 3. Single code or split-line range where line 2 has leading dash
    m_single = RE_SINGLE_CODE.match(line)
    if m_single:
        j = i + 1
        while j < len(lines) and (not lines[j].strip() or lines[j].strip() in ("Code", "Description")):
            j += 1
        if j < len(lines):
            next_l = lines[j].strip()
            m_dash_end = RE_SPLIT_DASH_END.match(next_l)
            if m_dash_end:
                return f"{m_single.group(0)} - {m_dash_end.group(1)}", (j - i + 1)
            # Solitary dash on next_l
            if next_l in ("-", "–", "—", "−"):
                k = j + 1
                while k < len(lines) and (not lines[k].strip() or lines[k].strip() in ("Code", "Description")):
                    k += 1
                if k < len(lines) and RE_SINGLE_CODE.match(lines[k].strip()):
                    return f"{m_single.group(0)} - {lines[k].strip()}", (k - i + 1)
        return m_single.group(0), 1

    return None, 0


# ---------------------------------------------------------------------------
#  Metadata & Contractor Extraction
# ---------------------------------------------------------------------------

def extract_contractor_info(doc):
    """
    Extract contractor table records from page 1 using block layout clustering.
    """
    page = doc[0]
    blocks = page.get_text("blocks")
    
    start_y = None
    end_y = None
    for b in blocks:
        text = b[4].strip()
        if text == "Contractor Information":
            start_y = b[1]
        elif text in ("Article Information", "General Information") and start_y is not None:
            end_y = b[1]
            break

    table_blocks = [
        b for b in blocks 
        if start_y is not None and (start_y + 40) < b[1] < (end_y if end_y else 999)
    ]
    table_blocks.sort(key=lambda b: (b[1], b[0]))

    # Group blocks into rows by vertical proximity
    rows = []
    current_row = []
    for b in table_blocks:
        if not current_row:
            current_row.append(b)
        else:
            row_y1 = max(item[3] for item in current_row)
            if b[1] < row_y1 + 5:
                current_row.append(b)
            else:
                rows.append(current_row)
                current_row = [b]
    if current_row:
        rows.append(current_row)

    contractors = []
    for r in rows:
        r.sort(key=lambda b: b[0])
        name_parts = []
        type_parts = []
        num_jur_state = []
        for b in r:
            x0, text = b[0], b[4].strip()
            if x0 < 180:
                name_parts.append(text)
            elif x0 < 280:
                type_parts.append(text)
            else:
                num_jur_state.append(text)

        name = clean_line(" ".join(name_parts))
        ctype = clean_line(" ".join(type_parts))
        rest = "\n".join(num_jur_state).split("\n")
        rest = [l.strip() for l in rest if l.strip()]

        # Skip table column headers
        if not name or "Contractor Name" in name:
            continue

        num = rest[0] if len(rest) > 0 else ""
        jur = rest[1] if len(rest) > 1 else ""
        state = clean_line(" ".join(rest[2:])) if len(rest) > 2 else ""

        contractors.append({
            "contractor_name": name,
            "contract_type": ctype,
            "contract_number": num,
            "jurisdiction": jur,
            "states": state
        })

    return contractors


def extract_metadata(doc):
    """
    Extract article metadata:
    Article ID, Original Effective Date, Revised Effective Date,
    Revision End Date, Retirement Date, Contractor Information, Article Status.
    """
    full_text = ""
    for p in range(len(doc)):
        full_text += doc[p].get_text("text") + "\n"

    meta = {}

    # Article ID
    m = re.search(r'Article ID\n(\S+)', full_text)
    article_id = m.group(1).strip() if m else ""
    meta["article_id"] = article_id

    # Article Title
    m = re.search(r'Article Title\n(.+?)(?:\nArticle Type)', full_text, re.DOTALL)
    meta["article_title"] = clean_line(m.group(1)) if m else ""

    # Original Effective Date
    m = re.search(r'Original Effective Date\n(.+?)(?:\nRevision Effective Date)', full_text, re.DOTALL)
    meta["original_effective_date"] = clean_line(m.group(1)) if m else "N/A"

    # Revised Effective Date
    m = re.search(r'Revision Effective Date\n(.+?)(?:\nRevision Ending Date)', full_text, re.DOTALL)
    meta["revised_effective_date"] = clean_line(m.group(1)) if m else "N/A"

    # Revision End Date
    m = re.search(r'Revision Ending Date\n(.+?)(?:\nRetirement Date)', full_text, re.DOTALL)
    meta["revision_end_date"] = clean_line(m.group(1)) if m else "N/A"

    # Retirement Date
    m = re.search(r'Retirement Date\n(.+?)(?:\n(?:CPT codes|CMS National|Article Guidance))', full_text, re.DOTALL)
    meta["retirement_date"] = clean_line(m.group(1)) if m else "N/A"

    # Article Status (Active or Inactive)
    # Check if Retirement Date is set (other than N/A) or if status is 'Currently in Effect'
    retirement_val = meta["retirement_date"].upper()
    if retirement_val not in ("N/A", "", "NONE") and re.search(r'\d{1,2}/\d{1,2}/\d{2,4}', retirement_val):
        status = "Inactive"
    elif "Currently in Effect" in full_text:
        status = "Active"
    elif "Retired" in full_text and "Retirement Date" in full_text:
        status = "Inactive"
    else:
        status = "Active"
    meta["article_status"] = status

    # Contractor Information
    meta["contractor_information"] = extract_contractor_info(doc)

    return meta


# ---------------------------------------------------------------------------
#  Code Groups Extraction
# ---------------------------------------------------------------------------

def extract_article_groups(doc, article_id):
    """
    Extract code groups and their complete descriptions:
    1. Covered HCPCS/CPT groups (group_number, group_description)
    2. Covered ICD groups (group_number, group_description)

    Includes both the paragraph before the code table and any description/notes/explanations
    after the code table.
    """
    lines = get_document_lines(doc)

    section_cpt = "CPT/HCPCS Codes"
    section_icd = "ICD-10-CM Codes that Support Medical Necessity"
    tracked_sections = [section_cpt, section_icd]

    section_terminators = [
        "CPT/HCPCS Modifiers",
        "ICD-10-CM Codes that Support Medical Necessity",
        "ICD-10-CM Codes that DO NOT Support Medical Necessity",
        "ICD-10-PCS Codes",
        "Additional ICD-10 Information",
        "Bill Type Codes",
        "Revenue Codes",
        "Other Coding Information",
        "Revision History Information",
        "Associated Documents",
    ]

    cpt_groups = []
    icd_groups = []

    current_section = ""
    current_group_name = ""
    para_lines = []
    post_lines = []
    mode = ""  # "PARAGRAPH", "CODES", "POST_TABLE"
    expecting_desc = False

    def finalize_current_group():
        nonlocal current_group_name, para_lines, post_lines, mode, expecting_desc
        if current_group_name and current_section:
            full_desc_parts = []
            if para_lines:
                full_desc_parts.append(clean_line(" ".join(para_lines)))
            if post_lines:
                full_desc_parts.append(clean_line(" ".join(post_lines)))
            full_desc = " ".join(full_desc_parts).strip()

            record = {
                "article_id": article_id,
                "group_number": current_group_name,
                "group_description": full_desc
            }
            if current_section == section_cpt:
                cpt_groups.append(record)
            elif current_section == section_icd:
                icd_groups.append(record)

        current_group_name = ""
        para_lines = []
        post_lines = []
        mode = ""
        expecting_desc = False

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # Check for section header
        if line in tracked_sections:
            finalize_current_group()
            current_section = line
            i += 1
            continue

        # Check for section exit/terminator
        if line in section_terminators and line not in tracked_sections:
            finalize_current_group()
            current_section = ""
            i += 1
            continue

        if not current_section:
            i += 1
            continue

        # Check for Group definition header: e.g. "Group 1 (8 Codes)" or "Group 1"
        m_grp = re.match(r'^(Group \d+)(?:\s*\(\d+\s*Codes?\))?$', line, re.IGNORECASE)
        if m_grp:
            finalize_current_group()
            current_group_name = m_grp.group(1).title()
            mode = "PARAGRAPH"
            i += 1
            continue

        # Check for "Group N Paragraph"
        m_para = re.match(r'^(Group \d+)\s+Paragraph$', line, re.IGNORECASE)
        if m_para:
            if not current_group_name or current_group_name != m_para.group(1).title():
                finalize_current_group()
                current_group_name = m_para.group(1).title()
            mode = "PARAGRAPH"
            i += 1
            continue

        # Check for "Group N Codes"
        m_codes = re.match(r'^(Group \d+)\s+Codes$', line, re.IGNORECASE)
        if m_codes:
            if not current_group_name:
                current_group_name = m_codes.group(1).title()
            mode = "CODES"
            expecting_desc = False
            i += 1
            continue

        # Table headers: "Code" and "Description"
        if line in ("Code", "Description"):
            i += 1
            continue

        if mode == "PARAGRAPH":
            para_lines.append(line)
            i += 1
            continue

        if mode == "CODES":
            code_str, consumed = match_code(lines, i)
            if code_str:
                expecting_desc = True
                i += consumed
                continue
            elif expecting_desc:
                # Description for the code; lookahead for wrapped description lines
                j = i + 1
                while j < len(lines):
                    next_l = lines[j].strip()
                    if not next_l:
                        j += 1
                        continue
                    next_code, _ = match_code(lines, j)
                    if (next_code or 
                        re.match(r'^Group \d+', next_l, re.IGNORECASE) or
                        re.match(r'^(Note|Asterisk|\*|Explanation)', next_l, re.IGNORECASE) or
                        next_l in tracked_sections or 
                        next_l in section_terminators or
                        next_l in ("Code", "Description")):
                        break
                    j += 1
                expecting_desc = False
                i = j
                continue
            else:
                # Line after the table has finished
                mode = "POST_TABLE"
                post_lines.append(line)
                i += 1
                continue

        if mode == "POST_TABLE":
            post_lines.append(line)
            i += 1
            continue

        i += 1

    finalize_current_group()
    return cpt_groups, icd_groups


# ---------------------------------------------------------------------------
#  Excel Generation
# ---------------------------------------------------------------------------

def apply_excel_styling(ws, header_color="1F4E79"):
    """Format table with corporate dark-blue header, borders, and column widths."""
    header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    cell_font = Font(name="Calibri", size=10)
    thin_side = Side(style='thin', color="D3D3D3")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Style header
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Style data rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = cell_font
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 90


def create_excel_workbook(all_cpt_groups, all_icd_groups, output_path):
    """
    Generate single Excel workbook containing exactly 2 sheets:
      Sheet 1: HCPCS Code Groups (Article ID, HCPCS Code Group, Group Description)
      Sheet 2: ICD Code Groups (Article ID, ICD Code Group, Group Description)
    """
    wb = Workbook()

    # Sheet 1: HCPCS Code Groups
    ws1 = wb.active
    ws1.title = "HCPCS Code Groups"
    ws1.append([
        "Article ID",
        "HCPCS Code Group",
        "Group Description"
    ])
    for r in all_cpt_groups:
        ws1.append([
            r["article_id"],
            r["group_number"],
            r["group_description"]
        ])
    apply_excel_styling(ws1)

    # Sheet 2: ICD Code Groups
    ws2 = wb.create_sheet("ICD Code Groups")
    ws2.append([
        "Article ID",
        "ICD Code Group",
        "Group Description"
    ])
    for r in all_icd_groups:
        ws2.append([
            r["article_id"],
            r["group_number"],
            r["group_description"]
        ])
    apply_excel_styling(ws2)

    wb.save(output_path)
    return output_path


# ---------------------------------------------------------------------------
#  Main Pipeline
# ---------------------------------------------------------------------------

def process_all_articles(input_dir=None, output_folder="output folder", output_excel="articles_data.xlsx", output_json="articles_metadata.json"):
    """
    Process all PDFs in input_dir and output:
      - 1 combined Excel file with 2 sheets:
          Sheet 1: HCPCS Code Groups (Article ID, HCPCS Code Group, Group Description)
          Sheet 2: ICD Code Groups (Article ID, ICD Code Group, Group Description)
      - 1 combined JSON file with metadata
    Saves outputs into the specified output_folder (creates it if missing).
    """
    if input_dir is None:
        input_dir = os.getcwd()

    if os.path.isfile(input_dir):
        pdf_files = [input_dir]
        base_dir = os.path.dirname(input_dir) or os.getcwd()
    else:
        base_dir = input_dir
        pdf_files = sorted(glob.glob(os.path.join(input_dir, "*.pdf")))

    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        return

    print(f"Found {len(pdf_files)} article PDF(s) to process.")

    # Ensure output folder exists (if folder exists, existing files inside are overwritten)
    out_dir = os.path.join(base_dir, output_folder)
    os.makedirs(out_dir, exist_ok=True)

    all_metadata = []
    all_cpt_groups = []
    all_icd_groups = []

    for pdf_path in pdf_files:
        pdf_name = os.path.basename(pdf_path)
        print(f"\nProcessing article: {pdf_name}")
        doc = pymupdf.open(pdf_path)

        # 1. Metadata
        meta = extract_metadata(doc)
        article_id = meta.get("article_id") or os.path.splitext(pdf_name)[0]
        all_metadata.append(meta)
        print(f"  Article ID: {article_id} | Status: {meta['article_status']}")
        print(f"  Contractors: {len(meta['contractor_information'])}")

        # 2. Code Groups
        cpt, icd = extract_article_groups(doc, article_id)
        print(f"  HCPCS Code Groups: {len(cpt)}")
        print(f"  ICD Code Groups:   {len(icd)}")

        all_cpt_groups.extend(cpt)
        all_icd_groups.extend(icd)

        doc.close()

    # Save JSON metadata
    json_path = os.path.join(out_dir, output_json)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(all_metadata, f, indent=2, ensure_ascii=False)
    print(f"\n✅ Saved metadata JSON: {json_path}")

    # Save Excel
    excel_path = os.path.join(out_dir, output_excel)
    create_excel_workbook(
        all_cpt_groups,
        all_icd_groups,
        excel_path
    )
    print(f"✅ Saved consolidated Excel: {excel_path}")

    print("\n--- Summary of All Extracted Data ---")
    print(f"  Articles Processed:            {len(all_metadata)}")
    print(f"  Sheet 1 (HCPCS Code Groups):   {len(all_cpt_groups)} rows")
    print(f"  Sheet 2 (ICD Code Groups):     {len(all_icd_groups)} rows")
    print(f"  Output Directory:              {out_dir}")


if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
    process_all_articles(target_dir)
