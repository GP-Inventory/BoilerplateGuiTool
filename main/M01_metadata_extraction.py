# Import Libraries that are required to adjust sys path
import sys                      # Provides access to system-specific parameters and functions
from pathlib import Path        # Offers an object-oriented interface for filesystem paths

# Adjust sys.path so we can import modules from the parent folder
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.dont_write_bytecode = True  # Prevents _pycache_ creation

# Import Project Libraries
from processes.P00_set_packages import *

# ====================================================================================================

# Import shared functions and file paths from other folders
from processes.P01_set_file_paths import downloaded_invoices_folder, processed_invoices_folder, mfc_name_converter, cover_sheet_template, download_folder
from processes.P03_shared_functions import load_mfc_names, rotate_pdf_in_memory, extract_text, extract_field, extract_target_pages, clean_currency
from processes.P04_static_lists import METADATA_PATTERNS
from processes.P06_class_items import InvoiceMetadata, InvoiceFinancials, MFCNames

