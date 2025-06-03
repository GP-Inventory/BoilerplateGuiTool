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





# ====================================================================================================

# Set Users Download Folder
download_folder = Path.home() / "Downloads"

# Define Root Windows Directory (Shared Drive)
root_directory_windows_folder = Path(r"I:\Shared drives\Invoices")

# Set template directories
templates_folder = root_directory_windows_folder / '01 Templates'

# Set required documents
cover_sheet_template = templates_folder / 'Invoice Metadata Cover Sheet.docx'
mfc_name_converter = templates_folder / 'mfc_names.csv'

# Set child directories
vendor = r'03 Blakemore' # Change this to the required vendor for the correct path
downloaded_invoices_folder = root_directory_windows_folder / vendor / '01 Downloaded Invoices'
processed_invoices_folder = root_directory_windows_folder / vendor / '02 Processed Invoices'
downloaded_credit_memo_folder = root_directory_windows_folder / vendor / '03 Downloaded Credit Memos'
processed_credit_memo_folder = root_directory_windows_folder / vendor / '04 Processed Credit Memos'