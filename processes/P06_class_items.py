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
from processes.P04_static_lists import VENDOR_METADATA




# ====================================================================================================

class InvoiceMetadata:
    """
    Stores core invoice identifiers and vendor information.
    """
    # Class-level vendor info (constants)
    short_vendor_name = VENDOR_METADATA.get('short_vendor_name', '')
    vendor_name = VENDOR_METADATA.get('vendor_name', '')
    vendor_address = VENDOR_METADATA.get('vendor_address', '')
    vendor_vat_reg_no = VENDOR_METADATA.get('vendor_vat_reg_no', '')
    vendor_awrs = VENDOR_METADATA.get('vendor_awrs', '')
    vendor_customer_no = VENDOR_METADATA.get('vendor_customer_no', '')

    # ================================================================================================

    def __init__(self, invoice_no, invoice_date, po_reference, invoice_type):
        self.invoice_no = invoice_no
        self.po_reference = po_reference
        self.invoice_type = invoice_type
        self.original_invoice_date = None

        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%b-%Y", "%d-%B-%Y"):
            try:
                parsed_date = datetime.strptime(invoice_date, fmt)
                # store for filenames
                self.original_invoice_date = parsed_date
                # canonical invoice_date (for due‐date logic & templates)
                self.invoice_date = parsed_date.strftime("%Y-%m-%d")
                self.due_date     = (parsed_date + timedelta(days=7)).strftime("%Y-%m-%d")
                self.po_month     = parsed_date.strftime("%Y-%m")
                break
            except ValueError:
                continue
        else:
            print(f"⚠️ Couldn't parse invoice_date '{invoice_date}'.")
            self.invoice_date = self.due_date = self.po_month = "UNKNOWN"

    # ================================================================================================

    def to_dict(self) -> dict:
        return {
            "Invoice No": self.invoice_no,
            "Invoice Date": self.invoice_date,
            "Invoice Type": self.invoice_type,
            "Due Date": self.due_date,
            "PO Reference": self.po_reference,
            "PO Month": self.po_month,
            "Short Vendor Name": self.short_vendor_name,
            "Vendor Name": self.vendor_name,
            "Vendor Address": self.vendor_address,
            "Vendor VAT Reg No": self.vendor_vat_reg_no,
            "Vendor AWRS": self.vendor_awrs,
            "Vendor Customer No": self.vendor_customer_no
        }

    # ================================================================================================

    def adjust_invoice_date_if_overdue(self):
        """
        Adjust invoice_date and po_reference if overdue.

        Always set `po_month` to the original invoice_date (MM/YY).
        If that date is more than three days before the first of the
        current month:

        1. Reset `invoice_date` → first of current month (YYYY-MM-DD)
        2. If PO starts with "GUK", prefix it with "<MM/YY> invoice\n",
            otherwise just set it to "<MM/YY> invoice".

        On parse error, prints a warning and sets `po_month` to "UNKNOWN".
        """
        today = datetime.today()
        first_of_month = datetime(today.year, today.month, 1)
        date_threshold = first_of_month - timedelta(days=3)

        try:
            original_invoice_date = datetime.strptime(self.invoice_date, "%Y-%m-%d")
            # Always capture the MM/YY of the original date
            self.po_month = original_invoice_date.strftime("%m/%y")

            if original_invoice_date < date_threshold:
                # Move invoice date to the first of this month
                self.invoice_date = first_of_month.strftime("%Y-%m-%d")

                # Update the PO reference
                if self.po_reference.startswith("GUK"):
                    self.po_reference = f"{self.po_month} invoice\n{self.po_reference}"
                else:
                    self.po_reference = f"{self.po_month} invoice"

        except Exception as e:
            print(f"⚠️ Date parsing error in InvoiceMetadata: {e}")
            self.po_month = "UNKNOWN"

    # ================================================================================================

    def generate_standard_filename(self, mfc, financials):
        # Prefer the original parsed date for the filename:
        if isinstance(self.original_invoice_date, datetime):
            date_part = self.original_invoice_date.strftime("%y.%m.%d")
        else:
            # fallback if parse failed
            try:
                dt = datetime.strptime(self.invoice_date, "%Y-%m-%d")
                date_part = dt.strftime("%y.%m.%d")
            except:
                date_part = "unknown-date"

        vendor_part   = self.short_vendor_name
        mfc_part      = mfc.friendly_mfc_name or "UnknownMFC"
        invoice_part  = self.invoice_no or "UnknownInvoice"
        amount_part   = f"£{financials.total_gross:.2f}"

        # include PO only if it's GUK…
        po_part = f" - {self.po_reference}" if self.po_reference.startswith("GUK") else ""

        return f"{date_part} - {vendor_part} - {mfc_part} - {invoice_part}{po_part} - {amount_part}"

# ====================================================================================================

class MFCNames:
    def __init__(self, row: dict):
        self.id = row.get("ID")
        self.mfc_name = row.get("LOCATION_NAME")
        self.friendly_mfc_name = row.get("FRIENDLY_LOCATION_NAME")
        self.xero_name = row.get("XERO_NAME")
        self.location_address = row.get("LOCATION_ADDRESS")
        self.postcode = row.get("POSTCODE")
        self.morrisons_id = row.get("MORRISONS_ID")
        self.morrisons_postcode = row.get("MORRISONS_POSTCODE")
        self.blakemore_postcode = row.get("BLAKEMORE_POSTCODE")
        self.wholegood_postcode = row.get("WHOLEOOD_POSTCODE")
        self.htdrinks_postcode = row.get("HTDRINKS_POSTCODE")
        self.on_the_rocks_postcode = row.get("ONTHEROCKS_POSTCODE")

    # ================================================================================================

    def is_valid(self) -> bool:
        """Checks if key fields are not empty."""
        return all([
            self.mfc_name,
            self.friendly_mfc_name,
            self.xero_name,
            self.postcode
        ])

    # ================================================================================================

    def to_dict(self) -> dict:
        return {
            "ID": self.id,
            "LOCATION_NAME": self.mfc_name,
            "FRIENDLY_LOCATION_NAME": self.friendly_mfc_name,
            "XERO_NAME": self.xero_name,
            "LOCATION_ADDRESS": self.location_address,
            "POSTCODE": self.postcode,
            "MORRISONS_ID": self.morrisons_id,
            "MORRISONS_POSTCODE": self.morrisons_postcode,
            "BLAKEMORE_POSTCODE": self.blakemore_postcode,
            "WHOLEOOD_POSTCODE": self.wholegood_postcode,
            "HTDRINKS_POSTCODE": self.htdrinks_postcode,
            "ONTHEROCKS_POSTCODE": self.on_the_rocks_postcode
        }

class InvoiceFinancials:
    """
    Stores VAT band totals, auto-computes gross totals, and validates against expected values.
    """

    def __init__(
        self,
        net_0=0.0,
        vat_0=0.0,
        net_5=0.0,
        vat_5=0.0,
        net_20=0.0,
        vat_20=0.0,
        expected_net_total=None,
        expected_vat_total=None,
        expected_gross_total=None
    ) -> None:
        self.net_0 = net_0
        self.vat_0 = vat_0
        self.net_5 = net_5
        self.vat_5 = vat_5
        self.net_20 = net_20
        self.vat_20 = vat_20

        self.expected_net_total = expected_net_total
        self.expected_vat_total = expected_vat_total
        self.expected_gross_total = expected_gross_total

    # ================================================================================================

    # Gross by VAT rate
    @property
    def gross_0(self):
        return self.net_0 + self.vat_0

    @property
    def gross_5(self):
        return self.net_5 + self.vat_5

    @property
    def gross_20(self):
        return self.net_20 + self.vat_20

    # Totals
    @property
    def total_net(self):
        return self.net_0 + self.net_5 + self.net_20

    @property
    def total_vat(self):
        return self.vat_0 + self.vat_5 + self.vat_20

    @property
    def total_gross(self):
        return self.total_net + self.total_vat

    # ================================================================================================

    # Validation
    def validate_totals(self, tolerance=0.01) -> dict:
        """
        Compares calculated totals to expected values (if provided),
        within a specified tolerance.
        """
        results = {}
        if self.expected_net_total is not None:
            results["net_match"] = abs(self.total_net - self.expected_net_total) <= tolerance
        if self.expected_vat_total is not None:
            results["vat_match"] = abs(self.total_vat - self.expected_vat_total) <= tolerance
        if self.expected_gross_total is not None:
            results["gross_match"] = abs(self.total_gross - self.expected_gross_total) <= tolerance
        return results

    @property
    def is_valid(self):
        """Returns True only if all provided expected totals match."""
        results = self.validate_totals()
        return all(results.values()) if results else True

    # ================================================================================================

    def to_dict(self) -> dict:
        return {
            "Net 0%": round(self.net_0, 2),
            "VAT 0%": round(self.vat_0, 2),
            "Gross 0%": round(self.gross_0, 2),
            "Net 5%": round(self.net_5, 2),
            "VAT 5%": round(self.vat_5, 2),
            "Gross 5%": round(self.gross_5, 2),
            "Net 20%": round(self.net_20, 2),
            "VAT 20%": round(self.vat_20, 2),
            "Gross 20%": round(self.gross_20, 2),
            "Total Net": round(self.total_net, 2),
            "Total VAT": round(self.total_vat, 2),
            "Total Gross": round(self.total_gross, 2),
        }

    # ================================================================================================

    def __str__(self):
        return f"Total Net: £{self.total_net:.2f}, VAT: £{self.total_vat:.2f}, Gross: £{self.total_gross:.2f}"