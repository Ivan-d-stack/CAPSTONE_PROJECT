"""
Master Black Book Builder & PDF Generator for RetainIQ.
Assembles all chapters, applies professional styling, converts DOCX to PDF via Word COM,
and validates the exact final page count against the 51-60 page requirement (target: 55-57 pages).
"""

import os
import sys
import subprocess
import pypdf
import docx
from docx import Document
from report_helpers import add_header_footer

from content_frontmatter import build_frontmatter
from content_chapter1 import build_chapter_1
from content_chapter2 import build_chapter_2
from content_chapter3 import build_chapter_3
from content_chapter4 import build_chapter_4
from content_chapter5 import build_chapter_5
from content_chapter6_7 import build_chapter_6_and_7

def build_complete_blackbook():
    print("=" * 60)
    print("STARTING RETAINIQ BLACK BOOK GENERATION")
    print("=" * 60)

    doc = Document()
    add_header_footer(doc)

    print("Building Preliminary Front Matter...")
    build_frontmatter(doc)

    print("Building Chapter 1: Introduction...")
    build_chapter_1(doc)

    print("Building Chapter 2: System Analysis...")
    build_chapter_2(doc)

    print("Building Chapter 3: System Design...")
    build_chapter_3(doc)

    print("Building Chapter 4: Implementation and Testing...")
    build_chapter_4(doc)

    print("Building Chapter 5: Results and Discussions...")
    build_chapter_5(doc)

    print("Building Chapter 6 & 7: Conclusion, Future Work & References...")
    build_chapter_6_and_7(doc)

    docx_path = "Ivan_Dsouza_RetainIQ_Blackbook.docx"
    pdf_path = "Ivan_Dsouza_RetainIQ_Blackbook.pdf"

    print(f"Saving compiled DOCX to {docx_path}...")
    doc.save(docx_path)
    print(f"DOCX saved successfully. Size: {os.path.getsize(docx_path):,} bytes.")

    print("\nConverting DOCX to PDF via Microsoft Word COM...")
    cmd = [
        "powershell", "-ExecutionPolicy", "Bypass",
        "-File", ".\\convert_docx_to_pdf.ps1",
        "-DocxPath", f".\\{docx_path}",
        "-PdfPath", f".\\{pdf_path}"
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    print("PowerShell Output:", res.stdout)
    if res.stderr:
        print("PowerShell Error:", res.stderr)

    if os.path.exists(pdf_path):
        reader = pypdf.PdfReader(pdf_path)
        total_pages = len(reader.pages)
        print("=" * 60)
        print(f"PDF GENERATED SUCCESSFULLY: {pdf_path}")
        print(f"TOTAL PAGE COUNT: {total_pages} PAGES")
        print("=" * 60)
        
        if 51 <= total_pages <= 60:
            print(f"SUCCESS: Page count {total_pages} strictly satisfies the 51-60 page requirement!")
            if 55 <= total_pages <= 57:
                print(f"PERFECT TARGET: Page count {total_pages} is in the optimal 55-57 page target!")
        else:
            print(f"WARNING: Page count {total_pages} is outside the 51-60 range. Calibration needed.")
        return total_pages
    else:
        print("ERROR: PDF was not generated.")
        return None

if __name__ == '__main__':
    build_complete_blackbook()
