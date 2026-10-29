#!/usr/bin/env python3
"""
Extract all structured information from CMS LCD Billing & Coding Article PDFs
into a multi-sheet Excel file. No LLM is used — pure text/regex parsing.

Usage:
    python extract_article.py            # processes all PDFs in current directory
    python extract_article.py /path/to   # processes all PDFs in given directory
"""

import os
import re
import sys
import glob
import pymupdf
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------------

def get_full_text(doc):
    """Return full text of all pages concatenated, with page markers."""
    pages = []
    for p in range(len(doc)):
        pages.append(doc[p].get_text("text"))
    return "\n".join(pages)


def clean_line(line):
    """Strip and normalize whitespace."""
    return " ".join(line.split()).strip()


def is_footer_or_header(line):
    """Return True if the line is a page header/footer we should skip."""
    line = line.strip()
    if not line:
        return True
    # Pattern: "9/4/26, 4:25 PM" or similar date/time header
    if re.match(r'^\d{1,2}/\d{1,2}/\d{2,4},\s+\d{1,2}:\d{2}\s+(AM|PM)$', line):
        return True
    # Pattern: "Article - Billing and Coding: ..."
    if line.startswith("Article - Billing and Coding:"):
        return True
    if line.startswith("Article -") and "(" in line and ")" in line:
        return True
    # Pattern: URL footer
    if "cms.gov/medicare-coverage-database" in line:
        return True
    # Pattern: page number like "1/29"
    if re.match(r'^\d+/\d+$', line):
        return True
    return False


def get_clean_lines(doc):
    """Get all lines from the document, excluding headers/footers."""
    lines = []
    for pno in range(len(doc)):
        page_text = doc[pno].get_text("text")
        for line in page_text.split("\n"):
            if not is_footer_or_header(line):
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
#  Section-level parsing
# ---------------------------------------------------------------------------

def extract_contractor_info(doc):
    """Extract the Contractor Information table from page 1."""
    page = doc[0]
    text = page.get_text("text")
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    rows = []
    # Find 'Contractor Information' and parse blocks after it
    # The structure is: Contractor Name | Contract Type | Contract Number | Jurisdiction | States
    # Each contractor is "First Coast Service Options, Inc." followed by details

    i = 0
    while i < len(lines):
        if lines[i] == "Contractor Information":
            i += 1
            break
        i += 1

    # Skip header row labels
    while i < len(lines) and lines[i] in ("Contractor Name", "Contract", "Type", "Number", "Jurisdiction", "States"):
        i += 1

    # Now parse contractor entries
    while i < len(lines):
        if lines[i] == "Article Information" or lines[i] == "General Information":
            break
        # Look for "First Coast..." pattern
        if "First Coast" in lines[i] or "Novitas" in lines[i] or "National Government" in lines[i]:
            contractor_name = lines[i]
            i += 1
            # Collect remaining parts of the name (e.g., "Inc.")
            while i < len(lines) and lines[i] in ("Inc.", "Inc"):
                contractor_name += ", " + lines[i]
                i += 1
            # Next should be contract type
            contract_type = ""
            if i < len(lines):
                contract_type = lines[i]
                i += 1
            # might continue on next line
            while i < len(lines) and lines[i] in ("MAC", "DME MAC", "HHH"):
                contract_type += " " + lines[i] if contract_type else lines[i]
                i += 1
            # Contract number + jurisdiction + states
            contract_number = ""
            jurisdiction = ""
            states_parts = []
            while i < len(lines) and not ("First Coast" in lines[i] or "Novitas" in lines[i] or
                                          "National Government" in lines[i] or
                                          "Article Information" in lines[i] or
                                          "General Information" in lines[i]):
                line = lines[i]
                if re.match(r'^\d{5}\s*-', line):
                    # contract number line, possibly followed by jurisdiction and states on same or next lines
                    parts = line.split("\n")
                    contract_number = parts[0].strip()
                    for p in parts[1:]:
                        p = p.strip()
                        if p.startswith("J -") or p.startswith("J-"):
                            jurisdiction = p
                        elif p:
                            states_parts.append(p)
                elif line.startswith("J -") or line.startswith("J-"):
                    jurisdiction = line
                elif line in ("A and B", "MAC"):
                    # already captured
                    pass
                elif line in ("Article Information", "General Information"):
                    break
                else:
                    states_parts.append(line)
                i += 1

            states = " ".join(states_parts).strip()
            # Clean up contract_type - remove MAC if it got duplicated
            contract_type = contract_type.strip()

            rows.append({
                "Contractor Name": contractor_name,
                "Contract Type": contract_type,
                "Contract Number": contract_number,
                "Jurisdiction": jurisdiction,
                "States": states
            })
        else:
            i += 1

    return pd.DataFrame(rows) if rows else pd.DataFrame()


def extract_contractor_info_v2(doc):
    """Alternative extraction: parse the contractor table by examining text blocks."""
    page = doc[0]
    blocks = page.get_text("blocks")
    
    # Find all blocks and group by y-coordinate to reconstruct table rows
    # Filter out headers/footers
    data_blocks = []
    for b in blocks:
        text = b[4].strip()
        if is_footer_or_header(text):
            continue
        if text in ("Contractor Information",):
            continue
        data_blocks.append({
            "x0": b[0], "y0": b[1], "x1": b[2], "y1": b[3],
            "text": text.replace("\n", " ").strip()
        })

    # Group blocks by approximate y-center (rows)
    # For the contractor table, rows have similar y ranges
    rows = []
    # We'll just parse the full text approach
    text = page.get_text("text")
    
    # Use regex to find contractor entries
    # Pattern: "First Coast Service Options,\nInc." followed by fields
    contractor_pattern = re.compile(
        r'(First Coast Service Options,\s*Inc\.|Novitas Solutions,\s*Inc\.|National Government Services,\s*Inc\.)',
        re.IGNORECASE
    )
    
    parts = text.split("Article Information")[0]  # Only contractor section
    parts = parts.split("Contractor Information")[-1]  # After header
    
    # Remove column headers
    for header in ["Contractor Name", "Contract", "Type", "Number", "Jurisdiction", "States"]:
        parts = parts.replace(header, "", 1)
    
    # Split by contractor name occurrences
    segments = contractor_pattern.split(parts)
    
    entries = []
    i = 1  # segments[0] is before first match
    while i < len(segments):
        name = segments[i].strip()
        i += 1
        if i < len(segments):
            detail = segments[i].strip()
            i += 1
        else:
            detail = ""
        
        # Parse detail lines
        detail_lines = [l.strip() for l in detail.split("\n") if l.strip()]
        
        contract_type = ""
        contract_number = ""
        jurisdiction = ""
        states = ""
        
        j = 0
        while j < len(detail_lines):
            line = detail_lines[j]
            if line in ("A and B", "A", "B", "DME"):
                contract_type = line
                j += 1
                if j < len(detail_lines) and detail_lines[j] in ("MAC", "DME MAC", "HHH"):
                    contract_type += " " + detail_lines[j]
                    j += 1
            elif re.match(r'^\d{5}\s*-', line):
                contract_number = line
                j += 1
            elif line.startswith("J -") or line.startswith("J-") or line.startswith("J –"):
                jurisdiction = line
                j += 1
            elif line == "MAC":
                if not contract_type:
                    contract_type = "MAC"
                j += 1
            else:
                if states:
                    states += " " + line
                else:
                    states = line
                j += 1
        
        entries.append({
            "Contractor Name": name,
            "Contract Type": contract_type,
            "Contract Number": contract_number,
            "Jurisdiction": jurisdiction,
            "States": states
        })
    
    return pd.DataFrame(entries) if entries else pd.DataFrame()


def extract_article_info(doc):
    """Extract Article Information section: Article ID, Title, Type, Dates, etc."""
    text = get_full_text(doc)
    
    info = {}
    
    # Article ID
    m = re.search(r'Article ID\n(\S+)', text)
    if m:
        info["Article ID"] = m.group(1)
    
    # Article Title
    m = re.search(r'Article Title\n(.+?)(?:\nArticle Type)', text, re.DOTALL)
    if m:
        info["Article Title"] = clean_line(m.group(1))
    
    # Article Type
    m = re.search(r'Article Type\n(.+?)(?:\nOriginal Effective Date)', text, re.DOTALL)
    if m:
        info["Article Type"] = clean_line(m.group(1))
    
    # Original Effective Date
    m = re.search(r'Original Effective Date\n(.+?)(?:\nRevision Effective Date)', text, re.DOTALL)
    if m:
        info["Original Effective Date"] = clean_line(m.group(1))
    
    # Revision Effective Date
    m = re.search(r'Revision Effective Date\n(.+?)(?:\nRevision Ending Date)', text, re.DOTALL)
    if m:
        info["Revision Effective Date"] = clean_line(m.group(1))
    
    # Revision Ending Date
    m = re.search(r'Revision Ending Date\n(.+?)(?:\nRetirement Date)', text, re.DOTALL)
    if m:
        info["Revision Ending Date"] = clean_line(m.group(1))
    
    # Retirement Date
    m = re.search(r'Retirement Date\n(.+?)(?:\n(?:CPT codes|CMS National|Article Guidance))', text, re.DOTALL)
    if m:
        info["Retirement Date"] = clean_line(m.group(1))
    
    # LCD Reference
    m = re.search(r'LCD Reference Article\n(.*?)(?:\nBilling and Coding Article)', text, re.DOTALL)
    if m:
        val = clean_line(m.group(1))
        if val:
            info["LCD Reference Article"] = val
    
    # Billing and Coding Article title
    m = re.search(r'Billing and Coding Article\n(.+?)(?:\nLinks in PDF)', text, re.DOTALL)
    if m:
        info["Billing and Coding Article"] = clean_line(m.group(1))
    
    return info


def extract_article_text(doc):
    """Extract Article Text / Guidance, Documentation Requirements, Coding Guidance sections."""
    text = get_full_text(doc)
    
    def clean_section(raw):
        """Clean a section of text by removing footer/header lines."""
        cleaned_lines = []
        for line in raw.split("\n"):
            if not is_footer_or_header(line) and line.strip():
                cleaned_lines.append(line.strip())
        return " ".join(cleaned_lines).strip()
    
    sections = {}
    
    # CMS National Coverage Policy
    m = re.search(r'CMS National Coverage Policy\n(.*?)(?=Article Guidance)', text, re.DOTALL)
    if m:
        content = clean_section(m.group(1))
        if content:
            sections["CMS National Coverage Policy"] = content
    
    # Social Security Act references
    m = re.search(r'Social Security Act Standard References:\n(.*?)(?=Article Guidance)', text, re.DOTALL)
    if m:
        sections["Social Security Act Standard References"] = clean_section(m.group(1))
    
    # Article Text
    m = re.search(r'Article Text\n(.*?)(?=Coding Guidance)', text, re.DOTALL)
    if m:
        sections["Article Text"] = clean_section(m.group(1))
    
    # Coding Guidance
    m = re.search(r'Coding Guidance\n(.*?)(?=Documentation Requirements)', text, re.DOTALL)
    if m:
        sections["Coding Guidance"] = clean_section(m.group(1))
    
    # Documentation Requirements
    m = re.search(r'Documentation Requirements\n(.*?)(?=Coding Information)', text, re.DOTALL)
    if m:
        sections["Documentation Requirements"] = clean_section(m.group(1))
    
    return sections


def extract_group_codes(doc):
    """
    Extract ALL Group codes from the document.
    Returns a list of dicts, each with:
      - Section: e.g. "CPT/HCPCS Codes", "ICD-10-CM Codes that Support Medical Necessity", etc.
      - Group: e.g. "Group 1"
      - Group Count: e.g. "29 Codes"
      - Group Paragraph: the paragraph text
      - Code: the code value
      - Description: the code description
    """
    lines = get_clean_lines(doc)
    
    results = []
    current_section = ""
    current_group = ""
    current_group_count = ""
    current_group_paragraph = ""
    in_paragraph = False
    in_codes = False
    expecting_description = False
    pending_code = ""
    
    # Define known section headers
    section_headers = [
        "CPT/HCPCS Codes",
        "ICD-10-CM Codes that Support Medical Necessity",
        "ICD-10-CM Codes that DO NOT Support Medical Necessity",
        "ICD-10-PCS Codes",
        "Bill Type Codes",
        "Revenue Codes",
    ]
    
    # Sections that go directly to Code/Description without Group headers
    ungrouped_sections = [
        "Bill Type Codes",
        "Revenue Codes",
    ]
    
    # Define terminators: sections that end code extraction
    terminators = [
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
    
    # Track whether we're in an ungrouped section waiting for Code/Desc headers
    awaiting_code_header = False
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        if not line:
            i += 1
            continue
        
        # Check for section header
        matched_section = False
        for sh in section_headers:
            if line == sh:
                current_section = sh
                current_group = ""
                current_group_count = ""
                current_group_paragraph = ""
                in_paragraph = False
                in_codes = False
                matched_section = True
                # For ungrouped sections, prepare to enter codes mode on Code/Description header
                if sh in ungrouped_sections:
                    awaiting_code_header = True
                else:
                    awaiting_code_header = False
                break
        if matched_section:
            i += 1
            continue
        
        # Check for terminators that are NOT section headers we track
        if line in terminators and line not in section_headers:
            in_codes = False
            in_paragraph = False
            awaiting_code_header = False
            i += 1
            continue
        
        # Check for "Group N (X Codes)" header
        m = re.match(r'^(Group \d+)\s*\((\d+ Codes?)\)', line)
        if m:
            current_group = m.group(1)
            current_group_count = m.group(2)
            current_group_paragraph = ""
            in_paragraph = False
            in_codes = False
            awaiting_code_header = False
            i += 1
            continue
        
        # Check for "Group N Paragraph"
        m = re.match(r'^Group \d+ Paragraph$', line)
        if m:
            in_paragraph = True
            in_codes = False
            current_group_paragraph = ""
            i += 1
            continue
        
        # Check for "Group N Codes"
        m = re.match(r'^Group \d+ Codes$', line)
        if m:
            in_paragraph = False
            in_codes = True
            i += 1
            continue
        
        # Skip column header "Code" / "Description"
        # For ungrouped sections, seeing "Code" header means codes start next
        if line == "Code" or line == "Description":
            if awaiting_code_header and line == "Code":
                in_codes = True
                awaiting_code_header = False
            i += 1
            continue
        
        # For ungrouped sections, accumulate paragraph text before Code/Description header
        if awaiting_code_header and current_section:
            current_group_paragraph += (" " if current_group_paragraph else "") + line
            i += 1
            continue
        
        # If we're in paragraph mode, accumulate text
        if in_paragraph:
            # Stop paragraph if we hit "Group N Codes"
            if re.match(r'^Group \d+ Codes$', line):
                in_paragraph = False
                in_codes = True
                i += 1
                continue
            current_group_paragraph += (" " if current_group_paragraph else "") + line
            i += 1
            continue
        
        # If we're in codes mode, parse code-description pairs
        if in_codes and current_section:
            code_str, consumed = match_code(lines, i)
            if code_str:
                if expecting_description and pending_code:
                    # Previous code had no description
                    results.append({
                        "Section": current_section,
                        "Group": current_group,
                        "Group Count": current_group_count,
                        "Group Paragraph": current_group_paragraph,
                        "Code": pending_code,
                        "Description": ""
                    })
                pending_code = code_str
                expecting_description = True
                i += consumed
                continue
            elif expecting_description:
                # This line is the description for the pending code
                desc = line
                # Check if description continues on next line(s) - if next line is NOT a code
                j = i + 1
                while j < len(lines):
                    next_line = lines[j].strip()
                    if not next_line:
                        j += 1
                        continue
                    # Check if next line is a code, section header, group header, or terminator
                    next_code, _ = match_code(lines, j)
                    if (next_code or
                        re.match(r'^Group \d+', next_line) or
                        next_line in section_headers or
                        next_line in terminators or
                        next_line in ("Code", "Description")):
                        break
                    desc += " " + next_line
                    j += 1
                
                results.append({
                    "Section": current_section,
                    "Group": current_group,
                    "Group Count": current_group_count,
                    "Group Paragraph": current_group_paragraph,
                    "Code": pending_code,
                    "Description": clean_line(desc)
                })
                pending_code = ""
                expecting_description = False
                i = j
                continue
            else:
                # Post-table text for the current group
                if current_group:
                    current_group_paragraph += (" " if current_group_paragraph else "") + line
                    for r in results:
                        if r["Section"] == current_section and r["Group"] == current_group:
                            r["Group Paragraph"] = current_group_paragraph
                i += 1
                continue
        
        i += 1
    
    # Don't forget any pending code
    if expecting_description and pending_code:
        results.append({
            "Section": current_section,
            "Group": current_group,
            "Group Count": current_group_count,
            "Group Paragraph": current_group_paragraph,
            "Code": pending_code,
            "Description": ""
        })
    
    return results


def extract_modifiers(doc):
    """Extract CPT/HCPCS Modifiers section."""
    text = get_full_text(doc)
    m = re.search(r'CPT/HCPCS Modifiers\n(.*?)(?=ICD-10-CM Codes that Support Medical Necessity)', text, re.DOTALL)
    if m:
        val = clean_line(m.group(1))
        return val if val and val != "N/A" else "N/A"
    return "N/A"


def extract_additional_sections(doc):
    """Extract: ICD-10-PCS Codes, Additional ICD-10, Other Coding Info."""
    text = get_full_text(doc)
    sections = {}
    
    m = re.search(r'ICD-10-PCS Codes\n(.*?)(?=Additional ICD-10 Information)', text, re.DOTALL)
    if m:
        sections["ICD-10-PCS Codes"] = clean_line(m.group(1))
    
    m = re.search(r'Additional ICD-10 Information\n(.*?)(?=Bill Type Codes)', text, re.DOTALL)
    if m:
        sections["Additional ICD-10 Information"] = clean_line(m.group(1))
    
    m = re.search(r'Other Coding Information\n(.*?)(?=Revision History Information)', text, re.DOTALL)
    if m:
        sections["Other Coding Information"] = clean_line(m.group(1))
    
    return sections


def extract_revision_history(doc):
    """Extract Revision History entries."""
    text = get_full_text(doc)
    
    # Find revision history section
    m = re.search(r'Revision History Information\n(.*?)(?=Associated Documents)', text, re.DOTALL)
    if not m:
        return pd.DataFrame()
    
    rev_text = m.group(1)
    lines = [l.strip() for l in rev_text.split("\n") if l.strip()]
    
    # Filter out headers/footers
    lines = [l for l in lines if not is_footer_or_header(l)]
    
    # Skip the column headers: "Revision", "History", "Date", "Number", "Revision History Explanation"
    # Parse date-number-explanation triples
    entries = []
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Skip repeated column headers
        if line in ("Revision", "History", "Date", "Number", "Revision History Explanation"):
            i += 1
            continue
        
        # Look for a date pattern: MM/DD/YYYY
        date_match = re.match(r'^(\d{2}/\d{2}/\d{4})$', line)
        if date_match:
            date = date_match.group(1)
            i += 1
            
            # Next should be revision number (e.g., R13, R12, etc.)
            rev_num = ""
            if i < len(lines):
                rev_match = re.match(r'^(R\d+)$', lines[i])
                if rev_match:
                    rev_num = rev_match.group(1)
                    i += 1
            
            # Collect explanation until next date or end
            explanation_parts = []
            while i < len(lines):
                if re.match(r'^\d{2}/\d{2}/\d{4}$', lines[i]):
                    break
                if lines[i] in ("Revision", "History", "Date", "Number", "Revision History Explanation"):
                    i += 1
                    continue
                explanation_parts.append(lines[i])
                i += 1
            
            entries.append({
                "Revision History Date": date,
                "Revision History Number": rev_num,
                "Revision History Explanation": " ".join(explanation_parts)
            })
        else:
            i += 1
    
    return pd.DataFrame(entries) if entries else pd.DataFrame()


def extract_associated_documents(doc):
    """Extract Associated Documents section."""
    text = get_full_text(doc)
    
    info = {}
    
    # Related LCDs
    m = re.search(r'Related Local Coverage Documents\n(?:LCDs\n)?(.*?)(?=Related National Coverage Documents)', text, re.DOTALL)
    if m:
        info["Related LCDs"] = clean_line(m.group(1))
    
    # Related NCDs
    m = re.search(r'Related National Coverage Documents\n(?:NCDs\n)?(.*?)(?=Statutory Requirements URLs)', text, re.DOTALL)
    if m:
        info["Related NCDs"] = clean_line(m.group(1))
    
    # Statutory Requirements URLs
    m = re.search(r'Statutory Requirements URLs\n(.*?)(?=Rules and Regulations URLs)', text, re.DOTALL)
    if m:
        info["Statutory Requirements URLs"] = clean_line(m.group(1))
    
    # Rules and Regulations URLs
    m = re.search(r'Rules and Regulations URLs\n(.*?)(?=CMS Manual Explanations URLs)', text, re.DOTALL)
    if m:
        info["Rules and Regulations URLs"] = clean_line(m.group(1))
    
    # CMS Manual Explanations URLs
    m = re.search(r'CMS Manual Explanations URLs\n(.*?)(?=Other URLs)', text, re.DOTALL)
    if m:
        info["CMS Manual Explanations URLs"] = clean_line(m.group(1))
    
    # Other URLs
    m = re.search(r'Other URLs\n(.*?)(?=Public Versions)', text, re.DOTALL)
    if m:
        info["Other URLs"] = clean_line(m.group(1))
    
    return info


def extract_public_versions(doc):
    """Extract Public Versions table."""
    text = get_full_text(doc)
    
    m = re.search(r'Public Versions\n(.*?)(?=Keywords|$)', text, re.DOTALL)
    if not m:
        return pd.DataFrame()
    
    version_text = m.group(1)
    lines = [l.strip() for l in version_text.split("\n") if l.strip()]
    lines = [l for l in lines if not is_footer_or_header(l)]
    
    entries = []
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Skip column headers
        if line in ("Updated On", "Effective Dates", "Status"):
            i += 1
            continue
        
        # Skip informational text
        if "archived" in line.lower() or "MCD Archive" in line:
            i += 1
            continue
        if "You are here" in line:
            i += 1
            continue
        
        # Look for date pattern MM/DD/YYYY
        date_match = re.match(r'^(\d{2}/\d{2}/\d{4})$', line)
        if date_match:
            updated_on = date_match.group(1)
            i += 1
            
            # Next: effective dates (e.g., "02/04/2026 - N/A")
            eff_dates = ""
            if i < len(lines):
                eff_dates = lines[i]
                i += 1
            
            # Next: status
            status = ""
            if i < len(lines) and not re.match(r'^\d{2}/\d{2}/\d{4}$', lines[i]):
                status = lines[i]
                i += 1
                # Skip "You are here" or "View"
                while i < len(lines) and lines[i] in ("You are here", "View"):
                    i += 1
            
            entries.append({
                "Updated On": updated_on,
                "Effective Dates": eff_dates,
                "Status": status
            })
        else:
            i += 1
    
    return pd.DataFrame(entries) if entries else pd.DataFrame()


def extract_keywords(doc):
    """Extract Keywords section."""
    text = get_full_text(doc)
    m = re.search(r'Keywords\n(.*?)$', text, re.DOTALL)
    if m:
        raw = m.group(1)
        # Clean out footer/header lines
        cleaned_lines = []
        for line in raw.split("\n"):
            if not is_footer_or_header(line) and line.strip():
                cleaned_lines.append(line.strip())
        val = " ".join(cleaned_lines).strip()
        return val if val else "N/A"
    return "N/A"


# ---------------------------------------------------------------------------
#  Excel formatting
# ---------------------------------------------------------------------------

def style_worksheet(ws, header_fill_color="1F4E79"):
    """Apply consistent styling to a worksheet."""
    header_fill = PatternFill(start_color=header_fill_color, end_color=header_fill_color, fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Style header row
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    # Style data rows
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = thin_border
    
    # Auto-width columns (capped)
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    max_len = max(max_len, min(len(str(cell.value)), 60))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)


# ---------------------------------------------------------------------------
#  Main processing
# ---------------------------------------------------------------------------

def process_pdf(pdf_path, output_dir):
    """Process a single PDF and write an Excel file."""
    print(f"\n{'='*60}")
    print(f"Processing: {os.path.basename(pdf_path)}")
    print(f"{'='*60}")
    
    doc = pymupdf.open(pdf_path)
    
    # --- Extract all sections ---
    
    # 1. Contractor Information
    print("  Extracting: Contractor Information...")
    contractor_df = extract_contractor_info_v2(doc)
    print(f"    Found {len(contractor_df)} contractor entries")
    
    # 2. Article Information
    print("  Extracting: Article Information...")
    article_info = extract_article_info(doc)
    print(f"    Found {len(article_info)} fields")
    
    # 3. Article Text / Guidance
    print("  Extracting: Article Text & Guidance...")
    article_text = extract_article_text(doc)
    print(f"    Found {len(article_text)} text sections")
    
    # 4. CPT/HCPCS Modifiers
    print("  Extracting: CPT/HCPCS Modifiers...")
    modifiers = extract_modifiers(doc)
    print(f"    Modifiers: {modifiers}")
    
    # 5. All Group Codes (CPT, ICD-10, Bill Type, Revenue)
    print("  Extracting: Group Codes (CPT/HCPCS, ICD-10-CM, Bill Type, Revenue)...")
    all_codes = extract_group_codes(doc)
    print(f"    Found {len(all_codes)} total code entries")
    
    # Break into separate DataFrames by section
    codes_by_section = {}
    for code in all_codes:
        section = code["Section"]
        if section not in codes_by_section:
            codes_by_section[section] = []
        codes_by_section[section].append(code)
    
    for section, codes in codes_by_section.items():
        print(f"      {section}: {len(codes)} codes")
    
    # 6. Additional sections
    print("  Extracting: Additional sections...")
    additional = extract_additional_sections(doc)
    
    # 7. Revision History
    print("  Extracting: Revision History...")
    revision_df = extract_revision_history(doc)
    print(f"    Found {len(revision_df)} revision entries")
    
    # 8. Associated Documents
    print("  Extracting: Associated Documents...")
    assoc_docs = extract_associated_documents(doc)
    
    # 9. Public Versions
    print("  Extracting: Public Versions...")
    versions_df = extract_public_versions(doc)
    print(f"    Found {len(versions_df)} version entries")
    
    # 10. Keywords
    print("  Extracting: Keywords...")
    keywords = extract_keywords(doc)
    
    doc.close()
    
    # --- Write to Excel ---
    basename = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = os.path.join(output_dir, f"{basename}_extracted.xlsx")
    
    print(f"\n  Writing Excel: {output_path}")
    
    wb = Workbook()
    
    # Sheet 1: Contractor Information
    ws = wb.active
    ws.title = "Contractor Information"
    if not contractor_df.empty:
        ws.append(list(contractor_df.columns))
        for _, row in contractor_df.iterrows():
            ws.append(list(row))
    else:
        ws.append(["No contractor information found"])
    style_worksheet(ws)
    
    # Sheet 2: Article Information
    ws = wb.create_sheet("Article Information")
    ws.append(["Field", "Value"])
    for k, v in article_info.items():
        ws.append([k, v])
    # Add modifiers
    ws.append(["CPT/HCPCS Modifiers", modifiers])
    # Add additional sections
    for k, v in additional.items():
        ws.append([k, v])
    # Add keywords
    ws.append(["Keywords", keywords])
    style_worksheet(ws)
    
    # Sheet 3: Article Text & Guidance
    ws = wb.create_sheet("Article Text & Guidance")
    ws.append(["Section", "Content"])
    for k, v in article_text.items():
        ws.append([k, v])
    style_worksheet(ws)
    
    # Sheet 4+: Group Codes (one sheet per section)
    section_sheet_names = {
        "CPT/HCPCS Codes": "CPT-HCPCS Codes",
        "ICD-10-CM Codes that Support Medical Necessity": "ICD-10 Support MedNecessity",
        "ICD-10-CM Codes that DO NOT Support Medical Necessity": "ICD-10 NOT Support MedNec",
        "Bill Type Codes": "Bill Type Codes",
        "Revenue Codes": "Revenue Codes",
    }
    
    for section, codes in codes_by_section.items():
        sheet_name = section_sheet_names.get(section, section[:31])
        ws = wb.create_sheet(sheet_name)
        ws.append(["Section", "Group", "Group Count", "Group Paragraph", "Code", "Description"])
        for c in codes:
            ws.append([c["Section"], c["Group"], c["Group Count"], c["Group Paragraph"], c["Code"], c["Description"]])
        style_worksheet(ws)
    
    # Sheet: Revision History
    ws = wb.create_sheet("Revision History")
    if not revision_df.empty:
        ws.append(list(revision_df.columns))
        for _, row in revision_df.iterrows():
            ws.append(list(row))
    else:
        ws.append(["No revision history found"])
    style_worksheet(ws)
    
    # Sheet: Associated Documents
    ws = wb.create_sheet("Associated Documents")
    ws.append(["Field", "Value"])
    for k, v in assoc_docs.items():
        ws.append([k, v])
    style_worksheet(ws)
    
    # Sheet: Public Versions
    ws = wb.create_sheet("Public Versions")
    if not versions_df.empty:
        ws.append(list(versions_df.columns))
        for _, row in versions_df.iterrows():
            ws.append(list(row))
    else:
        ws.append(["No public versions found"])
    style_worksheet(ws)
    
    wb.save(output_path)
    print(f"  ✅ Excel saved: {output_path}")
    
    # Print summary
    print(f"\n  --- Extraction Summary ---")
    print(f"  Contractor entries:   {len(contractor_df)}")
    print(f"  Article info fields:  {len(article_info)}")
    print(f"  Text sections:        {len(article_text)}")
    total_codes = len(all_codes)
    print(f"  Total codes extracted: {total_codes}")
    for section, codes in codes_by_section.items():
        print(f"    {section}: {len(codes)}")
    print(f"  Revision history:     {len(revision_df)}")
    print(f"  Public versions:      {len(versions_df)}")
    
    return output_path


def main():
    """Main entry point."""
    # Determine input directory or file
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        target = os.getcwd()

    if os.path.isfile(target):
        if not target.lower().endswith(".pdf"):
            print(f"Error: '{target}' is not a PDF file.")
            sys.exit(1)
        pdf_files = [target]
        base_dir = os.path.dirname(target) or os.getcwd()
    elif os.path.isdir(target):
        base_dir = target
        pdf_files = sorted(glob.glob(os.path.join(target, "*.pdf")))
    else:
        print(f"Error: '{target}' is not a valid file or directory.")
        sys.exit(1)
    
    output_dir = os.path.join(base_dir, "output folder")
    os.makedirs(output_dir, exist_ok=True)
    
    if not pdf_files:
        print(f"No PDF files found in: {target}")
        sys.exit(1)
    
    print(f"Found {len(pdf_files)} PDF file(s) to process.")
    results = []
    
    for pdf_path in pdf_files:
        try:
            out = process_pdf(pdf_path, output_dir)
            results.append((os.path.basename(pdf_path), "✅ Success", out))
        except Exception as e:
            print(f"  ❌ Error processing {os.path.basename(pdf_path)}: {e}")
            import traceback
            traceback.print_exc()
            results.append((os.path.basename(pdf_path), f"❌ Error: {e}", ""))
    
    print(f"\n{'='*60}")
    print("FINAL RESULTS")
    print(f"{'='*60}")
    for pdf_name, status, out_path in results:
        print(f"  {pdf_name}: {status}")
        if out_path:
            print(f"    -> {out_path}")


if __name__ == "__main__":
    main()
