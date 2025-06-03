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
from processes.P01_set_file_paths import mfc_name_converter
from processes.P06_class_items import MFCNames, InvoiceMetadata, InvoiceFinancials



# ====================================================================================================

def load_mfc_names(csv_path: Path) -> list[MFCNames]:
    """
    Loads MFC (Micro Fulfilment Centre) name mappings from a CSV file and returns
    a list of MFCNames objects, each representing one row of metadata.

    Args:
        csv_path (Path): Path to the CSV file containing MFC name and location mappings.

    Returns:
        list[MFCNames]: A list of MFCNames objects constructed from the CSV rows.
    """
    df = pd.read_csv(csv_path)
    return [MFCNames(row.to_dict()) for _, row in df.iterrows()]

# ====================================================================================================

def rotate_pdf_in_memory(input_pdf_path: Path, rotation: int = 270) -> BytesIO:
    """
    Rotates all pages of a PDF by a specified angle (e.g., 90, 180, 270).

    Args:
        input_pdf_path (Path): Path to the input PDF file.
        rotation (int): Rotation angle in degrees (must be a multiple of 90).

    Returns:
        BytesIO: Rotated PDF in memory.
    """
    if rotation % 90 != 0:
        raise ValueError(f"Rotation angle must be a multiple of 90. Got: {rotation}")

    reader = PdfReader(str(input_pdf_path))
    writer = PdfWriter()

    for page in reader.pages:
        page.rotate(rotation)
        writer.add_page(page)

    output_stream = BytesIO()
    writer.write(output_stream)
    output_stream.seek(0)
    return output_stream

# ====================================================================================================

def extract_field(text: str, field_pattern: str, patterns: Optional[dict[str, str]] = None) -> Optional[str]:
    """
    Extracts a value from the given text using either:
    - a named regex pattern from a dictionary (if `patterns` is provided), or
    - a direct regex string.

    Args:
        text (str): The full text block to search within.
        field_pattern (str): The name of the pattern key (if using `patterns`) or the regex itself.
        patterns (dict[str, str], optional): Dictionary of field name → regex string.

    Returns:
        Optional[str]: The first captured group (group 1), or None if not found.
    """
    pattern: str = patterns[field_pattern] if patterns and field_pattern in patterns else field_pattern

    match = re.search(pattern, text, flags=re.IGNORECASE)
    return match.group(1).strip() if match else None

# ====================================================================================================

def extract_target_pages(reader: PdfReader, search_pattern: str) -> BytesIO | None:
    """
    Searches all pages of a PDF for a text pattern and extracts matching pages into a new in-memory PDF.

    Args:
        reader (PdfReader): The PyPDF2 PdfReader object for the source PDF.
        search_pattern (str): Regex pattern to match text on each page.

    Returns:
        BytesIO | None: A BytesIO PDF stream containing only matching pages, or None if no matches found.
    """
    writer = PdfWriter()

    for page in reader.pages:
        text = page.extract_text() or ""
        if re.search(search_pattern, text, re.IGNORECASE):
            writer.add_page(page)

    if writer.pages:
        stream = BytesIO()
        writer.write(stream)
        stream.seek(0)
        return stream

    print(f'⚠️ No pages found matching pattern: "{search_pattern}"')
    return None

# ====================================================================================================

def clean_currency(value) -> float:
    """
    Cleans a currency input (e.g., '£1,234.56') and converts it to a float.

    Accepts:
    - Float or int: returned as-is (cast to float)
    - String with symbols like '£' or commas: cleaned and converted
    - None or empty string: returns 0.0

    Args:
        value: Currency value as str, int, or float.

    Returns:
        float: The numeric value as a float.
    """
    if isinstance(value, (int, float)):
        return float(value)
    if not value or not isinstance(value, str):
        return 0.0
    try:
        return float(value.replace("£", "").replace(",", "").strip())
    except ValueError:
        print(f"⚠️ Could not parse currency value: {value!r}")
        return 0.0

# ====================================================================================================

def extract_until_pattern(text: str, stop_pattern: str) -> list[str]:
    """
    Extracts the header section of an invoice from raw OCR text.

    This function scans line by line through the provided invoice text and collects lines
    until it encounters a user-defined pattern, which marks the beginning of the main invoice body.

    Parameters:
        text (str): Full extracted text content from the invoice PDF.
        stop_pattern (str): Regex pattern that marks the start of the main invoice section.

    Returns:
        list[str]: A list of lines representing the header section before the main invoice body.
    """
    lines = text.split('\n')
    header_lines = []
    for line in lines:
        if re.search(stop_pattern, line, re.IGNORECASE):
            break
        header_lines.append(line)
    return header_lines

# ====================================================================================================

def extract_between_patterns(text: str, start_pattern: str, stop_pattern: str, include_start: bool = False, include_stop: bool = False) -> list[str]:
    """
    Extracts lines from text between two regex-defined patterns.

    Parameters:
        text (str): Full extracted text content.
        start_pattern (str): Regex pattern that marks the start of the section.
        end_pattern (str): Regex pattern that marks the end of the section.
        include_start (bool): Whether to include the start line in the result.
        include_end (bool): Whether to include the end line in the result.

    Returns:
        list[str]: Lines found between the start and end patterns.
    """
    lines = text.split('\n')
    extracting = False
    extracted_lines = []

    for line in lines:
        if not extracting and re.search(start_pattern, line, re.IGNORECASE):
            extracting = True
            if include_start:
                extracted_lines.append(line)
            continue
        if extracting:
            if re.search(stop_pattern, line, re.IGNORECASE):
                if include_stop:
                    extracted_lines.append(line)
                break
            extracted_lines.append(line)

    return extracted_lines

# ====================================================================================================

def remove_cover_sheet_pages(pdf_path: Path, target_phrase: str = "INVOICE METADATA COVER SHEET") -> BytesIO:
    """
    Removes pages from a PDF that contain a specific marker phrase (typically used to identify
    cover sheets) and returns the cleaned PDF in memory. Also overwrites the original PDF.

    Parameters
    ----------
    pdf_path : Path
        Path to the original PDF file from which cover sheet pages should be removed.
    target_phrase : str, optional
        Phrase that identifies pages to skip (default is "INVOICE METADATA COVER SHEET").

    Returns
    -------
    BytesIO
        A seekable in-memory bytes buffer containing the new PDF (with unwanted pages removed).
        The buffer’s cursor is positioned at the start.

    Raises
    ------
    IOError
        If the input file cannot be read or the output file cannot be written.
    """
    try:
        reader = PdfReader(str(pdf_path))
    except Exception as e:
        raise IOError(f"Could not read PDF file: {pdf_path}") from e

    writer = PdfWriter()
    removed_pages = 0

    for page in reader.pages:
        text = page.extract_text() or ""
        if target_phrase.lower() in text.lower():
            removed_pages += 1
            continue
        writer.add_page(page)

    output = BytesIO()
    writer.write(output)
    output.seek(0)

    try:
        with open(pdf_path, "wb") as f:
            f.write(output.getvalue())
    except Exception as e:
        raise IOError(f"Failed to write cleaned PDF to {pdf_path}") from e

    output.seek(0)

    print(f"✅ Removed {removed_pages} cover sheet page(s) from {pdf_path.name}")
    return output

# ====================================================================================================

def check_existing_filename(
    metadata: InvoiceMetadata,
    financials: InvoiceFinancials,
    matched_mfc: MFCNames,
    destination_folder: Path,
    current_folder: Path
) -> tuple[str, bool]:
    """
    Generates a unique filename for the invoice PDF based on metadata and financials.

    The filename is constructed using a standard naming convention from the metadata. If a file with
    the same name already exists in the destination folder, it is considered a duplicate and will be
    prefixed with "Duplicate -". If further duplicates exist in the current folder (e.g., in-memory
    or working directory), a number is appended to avoid overwriting.

    Parameters
    ----------
    metadata : InvoiceMetadata
        Invoice metadata object containing filename details.
    financials : InvoiceFinancials
        Financial information used in filename generation.
    matched_mfc : MFCNames
        Location naming conventions to apply to the filename.
    destination_folder : Path
        Final storage location — used to detect processed duplicates.
    current_folder : Path
        Temporary or intermediate folder — used to avoid overwriting local copies.

    Returns
    -------
    tuple[str, bool]
        A safe filename and a boolean indicating whether it was renamed due to duplication.
    """
    base_filename = metadata.generate_standard_filename(matched_mfc, financials) + ".pdf"
    is_duplicate = False

    # Check if the filename already exists in the final destination folder
    if (destination_folder / base_filename).exists():
        is_duplicate = True
        prefix = "Duplicate"
        filename = f"{prefix} - {base_filename}"

        # If multiple duplicates exist in the current folder, append numbering
        counter = 2
        while (current_folder / filename).exists():
            filename = f"{prefix} ({counter}) - {base_filename}"
            counter += 1

        return filename, True

    # If no duplicate exists, return the base name
    return base_filename, False

# ====================================================================================================

def create_cover_sheet(template_path: str, output_path: str, invoice_data: dict[str, Any]):
    """
    Populates a Word document cover sheet template with values from invoice_data.
    Replaces placeholders and applies Calibri size 11 formatting.

    Parameters
    ----------
    template_path : str
        Path to the DOCX template containing placeholder tags (e.g., {{InvoiceNumber}}).
    output_path : str
        Path to save the completed Word document.
    invoice_data : dict[str, Any]
        Dictionary containing invoice metadata and financial values.
    """
    doc = Document(template_path)

    def fmt(value: Any) -> str:
        return f"{value:.2f}" if isinstance(value, float) else str(value)

    substitution_map = {
        "{{VendorName}}": fmt(invoice_data.get("Vendor Name", "")),
        "{{VendorAddress}}": fmt(invoice_data.get("Vendor Address", "")),
        "{{VendorVatNo}}": fmt(invoice_data.get("Vendor VAT Reg No", "")),
        "{{InvoiceNumber}}": fmt(invoice_data.get("Invoice No", "")),
        "{{InvoiceDate}}": fmt(invoice_data.get("Invoice Date", "")),
        "{{DueDate}}": fmt(invoice_data.get("Due Date", "")),
        "{{DescriptionFinal}}": fmt(invoice_data.get("Description", "")),
        "{{MFCName}}": fmt(invoice_data.get("Friendly MFC Name", "")),
        "{{MFCAddress}}": fmt(invoice_data.get("Delivery Address", "")),
        "{{XeroLocation}}": fmt(invoice_data.get("Xero Location", "")),
        "{{Net0}}": fmt(invoice_data.get("Net 0%", 0.0)),
        "{{VAT0}}": fmt(invoice_data.get("VAT 0%", 0.0)),
        "{{Gross0}}": fmt(invoice_data.get("Gross 0%", 0.0)),
        "{{Net5}}": fmt(invoice_data.get("Net 5%", 0.0)),
        "{{VAT5}}": fmt(invoice_data.get("VAT 5%", 0.0)),
        "{{Gross5}}": fmt(invoice_data.get("Gross 5%", 0.0)),
        "{{Net20}}": fmt(invoice_data.get("Net 20%", 0.0)),
        "{{VAT20}}": fmt(invoice_data.get("VAT 20%", 0.0)),
        "{{Gross20}}": fmt(invoice_data.get("Gross 20%", 0.0)),
        "{{NetTotal}}": fmt(invoice_data.get("Total Net", 0.0)),
        "{{NetVAT}}": fmt(invoice_data.get("Total VAT", 0.0)),
        "{{NetGross}}": fmt(invoice_data.get("Total Gross", 0.0)),
    }

    def replace_text_in_paragraph(paragraph):
        for placeholder, value in substitution_map.items():
            if placeholder in paragraph.text:
                paragraph.clear()
                run = paragraph.add_run(value)
                run.font.name = 'Calibri'
                run.font.size = Pt(11)
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')

    # Replace in body paragraphs
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph)

    try:
        doc.save(output_path)
        print(f"✅ Cover sheet saved to: {output_path}")
    except Exception as e:
        raise IOError(f"Failed to save cover sheet to {output_path}") from e

# ====================================================================================================

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path):
    """
    Converts a DOCX file to PDF using Microsoft Word (Windows only).

    Parameters
    ----------
    docx_path : Path
        Path to the .docx file to be converted.
    pdf_path : Path
        Destination path for the resulting PDF file.

    Raises
    ------
    IOError
        If Microsoft Word cannot open or save the file.
    """
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(str(docx_path))
        doc.SaveAs(str(pdf_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
    except Exception as e:
        raise IOError(f"Failed to convert {docx_path} to PDF: {e}") from e
    finally:
        word.Quit()

# ====================================================================================================

from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter


def merge_pdfs(cover_pdf_path: Path, original_pdf_path: Path, output_path: Path):
    """
    Merges a cover sheet PDF with an original invoice PDF and saves the result.

    Parameters
    ----------
    cover_pdf_path : Path
        Path to the cover sheet PDF (typically generated from a DOCX).
    original_pdf_path : Path
        Path to the original invoice PDF.
    output_path : Path
        Destination where the merged PDF will be saved.

    Raises
    ------
    IOError
        If any of the input files cannot be read or the output cannot be written.
    """
    writer = PdfWriter()

    try:
        for path in [cover_pdf_path, original_pdf_path]:
            reader = PdfReader(str(path))
            for page in reader.pages:
                writer.add_page(page)

        with open(output_path, "wb") as f:
            writer.write(f)

        print(f"✅ Final PDF saved to: {output_path}")

    except Exception as e:
        raise IOError(f"Failed to merge PDFs: {cover_pdf_path} + {original_pdf_path}") from e

# ====================================================================================================

def move_and_rename_final_pdf(final_pdf_path: Path, metadata: InvoiceMetadata, financials: InvoiceFinancials, matched_mfc: MFCNames, destination_folder: Path) -> Optional[Path]:
    """
    Renames and moves the final merged PDF to the destination folder using a standardized filename.

    Parameters
    ----------
    final_pdf_path : Path
        Path to the final merged PDF file.
    metadata : InvoiceMetadata
        Invoice metadata object (e.g., invoice number, date).
    financials : InvoiceFinancials
        Financial summary object used for filename formatting.
    matched_mfc : MFCNames
        MFC naming object providing location identifiers.
    destination_folder : Path
        Folder to move the renamed PDF into.

    Returns
    -------
    Optional[Path]
        The path of the moved file if successful, or None if it failed.
    """
    if not matched_mfc or not financials:
        print("⚠️ Missing required components for renaming/moving file.")
        return None

    new_filename = metadata.generate_standard_filename(matched_mfc, financials) + ".pdf"
    new_path = destination_folder / new_filename

    try:
        shutil.move(str(final_pdf_path), str(new_path))  # Cross-drive compatible
        print(f"📁 Moved to: {new_path}")
        return new_path
    except Exception as e:
        print(f"❌ Failed to move file: {e}")
        return None