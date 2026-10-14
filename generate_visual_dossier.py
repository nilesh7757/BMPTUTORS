#!/usr/bin/env python3
"""
Generate Comprehensive Visual Analytics DOCX Report for BMP Tutors.
Now delegates to build_custom_docx.py to generate 100% native, editable Word charts and clean tables.
"""

import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

import build_custom_docx

if __name__ == "__main__":
    print("Running modern editable docx generator...")
    build_custom_docx.build_report()
    print("Complete.")
