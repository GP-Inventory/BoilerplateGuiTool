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
from processes.P06_class_items import MFCNames



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