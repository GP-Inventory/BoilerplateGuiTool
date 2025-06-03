# Import necessary libraries (consistent across all sheets)

# Native to Python
# Import necessary libraries (consistent across all sheets)
import sys  # Provides access to system-specific parameters and functions
import re  # Supports regular expression matching and manipulation
import json  # Enables conversion between Python dictionaries and JSON strings (e.g., for saving, exporting, or pretty-printing data)
import shutil  # High-level file operations (copying, moving, deleting files and directories)
import csv  # Reading from and writing to CSV (Comma-Separated Values) files; useful for lightweight tabular data I/O
import tkinter as tk  # Base Tkinter GUI framework (windows, widgets)
from pathlib import Path  # Offers an object-oriented interface for filesystem paths
from datetime import datetime, timedelta  # Handling date/time parsing, formatting, and calculations
from io import BytesIO  # Enables handling of in-memory binary streams
from typing import Union, List, Optional, Dict  # Type hinting for function parameters and return values

# Install Required
import pandas as pd  # (pandas) Data analysis and manipulation using DataFrames
import win32com.client  # (pywin32) Accessing COM objects (e.g., automating Word for DOCX to PDF conversion)
import pdfplumber  # (pdfplumber) Extracts tables and structured content from PDFs using OCR-like analysis
from PyPDF2 import PdfReader, PdfWriter  # (PyPDF2) Facilitates reading from and writing to PDF files
from pdfminer.high_level import extract_text  # (pdfminer.six) Extracts text content from PDF files

# Word document generation and formatting (python-docx)
from docx import Document  # Creating and editing Word .docx files
from docx.shared import Pt  # Setting font sizes
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT  # Text alignment options (e.g., left, center)
from docx.oxml.ns import qn  # Accessing and modifying XML namespaces (for font compatibility)