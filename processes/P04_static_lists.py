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

VENDOR_METADATA = {
    "short_vendor_name": "Blakemore",
    "vendor_name": "A F Blakemore & Son Ltd",
    "vendor_address": "Long Acre Industrial Estate, Rosehill, Willenhall, WV13 2JP",
    "vendor_vat_reg_no": "431 3902 80",
    "vendor_awrs": "XEAW00000120927",
    "vendor_customer_no": "26244"
}

# Set Patterns for Invoice Metadata
METADATA_PATTERNS = { }

# Set Patterns for Invoice Delivery Address
ADDRESS_PATTERNS = { }

# Set Patterns for Invoice Financials
FINANCIALS_PATTERNS = { }

# Set Vendor Specific Other Patterns
REGEN_PATTERNS = { }