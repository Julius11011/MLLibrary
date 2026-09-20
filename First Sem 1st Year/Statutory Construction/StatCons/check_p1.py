# build_phase1_clean.py
import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core_builder import (
    create_docx_builder, add_header_box, add_heading_1, add_heading_2,
    add_body_p, add_student_explanation_box, add_alac_box,
    add_latin_maxims_box, add_statcon_table
)

from statcon_database_phase1 import cases_p1
from generate_phase1_full import cases_p1_part2

# Add 16 to 21
from statcon_database_phase1_complete import cases_p1_part3 if os.path.exists("statcon_database_phase1_complete.py") else []
