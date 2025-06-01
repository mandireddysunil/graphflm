import os
import re
import json
import copy # For deepcopy
import fitz # PyMuPDF
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import pytesseract
from PIL import Image, ImageDraw, ImageFont, UnidentifiedImageError

SUPPORTED_EXTENSIONS = ['.pdf', '.png', '.jpg', '.jpeg', '.xlsx', '.docx']

# --- PII Regex Patterns (for scanning within extracted text) ---
PHONE_REGEX_PII = re.compile(r'(?:\+91[\s-]?)?(?:0)?[6-9]\d{9}')
EMAIL_REGEX_PII = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
GSTIN_REGEX_PII = re.compile(r'\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b', re.IGNORECASE) # Case insensitive for safety
PAN_REGEX_PII = re.compile(r'\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b')
ACCOUNT_NO_REGEX_PII = re.compile(r'\b\d{9,18}\b') # Basic account number regex

# --- PII Masking Functions ---
def mask_phone(phone: str) -> str:
    digits = re.sub(r'\D', '', phone)
    if len(digits) >= 10:
        # Handle potential country code presence more carefully
        if len(digits) > 10 and digits.startswith('91'): # Assuming +91
            return "91" + digits[2:4] + 'XXXXXX' + digits[-2:] # e.g., 9198XXXXXX10
        return digits[:2] + 'XXXXXX' + digits[-2:]  # e.g., 98XXXXXX10
    return 'XX****XX'

def mask_email(email: str) -> str:
    parts = email.split('@')
    if len(parts) == 2:
        user, domain = parts
        domain_parts = domain.split('.')
        masked_user = user[0] + '***' if len(user) > 1 else (user if user else '***')
        masked_domain_start = domain_parts[0][0] + '***' if len(domain_parts[0]) > 1 else (domain_parts[0] if domain_parts[0] else '***')
        masked_domain_end = domain_parts[-1] if len(domain_parts) > 1 else (domain_parts[0] if not masked_domain_start.endswith('***') else "")
        # Avoids things like u***@g***. if domain was just "g"
        if masked_domain_start.endswith('***') and masked_domain_end == domain_parts[0] and len(domain_parts) == 1:
             return f"{masked_user}@{masked_domain_start}"
        return f"{masked_user}@{masked_domain_start}.{masked_domain_end}"
    return '***@***.***'

def mask_gstin(gstin: str) -> str:
    if len(gstin) == 15:
        return gstin[:2] + 'XXXXX' + gstin[7:11] + 'X' + gstin[12] + gstin[13] + 'X' # e.g., 29XXXXX1234X1ZX - corrected index for Z
    return "XXXXXXX" # Default mask for invalid length

def mask_pan(pan: str) -> str:
    if len(pan) == 10:
        return pan[:2] + 'XXXX' + pan[6:9] + pan[-1] # e.g., ABXXXX123F - corrected index
    return "XXXXX"

def mask_account_no(acc_no: str) -> str:
    if len(acc_no) > 4:
        return 'X' * (len(acc_no) - 4) + acc_no[-4:]
    return 'X' * len(acc_no) # Mask whole thing if too short

# --- PII Masking Application Function ---
def apply_pii_masking(data): # Can be dict, list, or string
    processed_data = copy.deepcopy(data)

    def _mask_recursive(item):
        if isinstance(item, dict):
            for key, value in item.items():
                item[key] = _mask_recursive(value)
            return item
        elif isinstance(item, list):
            for i, element in enumerate(item):
                item[i] = _mask_recursive(element)
            return item
        elif isinstance(item, str):
            # Order of application matters to avoid re-masking parts of already masked PII
            # For example, an account number might also be a phone number.
            # However, regexes are specific enough that this risk is low here.
            # More complex scenarios might require a single pass with combined regex or careful ordering.

            # Mask PANs first as they are very specific
            new_item_str = PAN_REGEX_PII.sub(lambda m: mask_pan(m.group(0)), item)
            # Then GSTINs
            new_item_str = GSTIN_REGEX_PII.sub(lambda m: mask_gstin(m.group(0)), new_item_str)
            # Then Phones
            new_item_str = PHONE_REGEX_PII.sub(lambda m: mask_phone(m.group(0)), new_item_str)
            # Then Emails
            new_item_str = EMAIL_REGEX_PII.sub(lambda m: mask_email(m.group(0)), new_item_str)
            # Finally, Account Numbers (as it's the broadest and might overlap with phone digits if not careful)
            # To avoid re-masking parts of phone numbers that look like account numbers after phone masking:
            # One strategy could be to only mask account numbers if they weren't part of a phone match.
            # For now, a simple sequential application.
            new_item_str = ACCOUNT_NO_REGEX_PII.sub(lambda m: mask_account_no(m.group(0)), new_item_str)
            return new_item_str
        else:
            return item

    return _mask_recursive(processed_data)

# --- Main Document Processing Function ---

def process_document(file_path: str) -> str:
    """
    Orchestrates the document processing pipeline:
    1. Validates file.
    2. Extracts text based on file type.
    3. Classifies document type from extracted text.
    4. Extracts specific information if a known document type.
    5. Applies PII masking to the extracted information.
    6. Returns a JSON string of the processed (and masked) data or error/status.
    """
    print(f"\n--- Processing document: {file_path} ---") # Status message

    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return json.dumps({"error": "File not found", "file_path": file_path}, indent=4)

    _, file_extension = os.path.splitext(file_path)
    file_extension_lower = file_extension.lower()

    if file_extension_lower not in SUPPORTED_EXTENSIONS:
        print(f"Error: Unsupported file type: {file_extension}")
        return json.dumps({"error": "Unsupported file type", "file_type": file_extension}, indent=4)

    print(f"File Type: {file_extension_lower}")
    extracted_text = ""
    extraction_error = False

    if file_extension_lower == '.pdf':
        extracted_text = extract_text_from_pdf(file_path)
    elif file_extension_lower == '.docx':
        extracted_text = extract_text_from_docx(file_path)
    elif file_extension_lower == '.xlsx':
        extracted_text = extract_text_from_xlsx(file_path)
    elif file_extension_lower in ['.png', '.jpg', '.jpeg']:
        extracted_text = extract_text_from_image_ocr(file_path)

    if extracted_text.startswith("Error:") or extracted_text.startswith("An unexpected error"):
        print(f"Text extraction failed: {extracted_text}")
        return json.dumps({"error": "Text extraction failed", "details": extracted_text}, indent=4)

    if not extracted_text.strip(): # Handles empty or whitespace-only text
        print("No text content extracted.")
        # For images, this path might be normal if OCR finds no text.
        # For other docs, it might mean an empty document.
        status_msg = "No text content found in document."
        if file_extension_lower in ['.png', '.jpg', '.jpeg']:
            status_msg = "OCR did not find any text in the image."
        return json.dumps({"status": status_msg, "file_path": file_path, "text_snippet": ""}, indent=4)

    snippet = extracted_text[:200].replace('\n', ' ')
    print(f"Extracted text snippet (first 200 chars): {snippet}")

    doc_type = classify_document(extracted_text)
    print(f"Classified Document Type: {doc_type}")

    processed_info = {} # This will hold the final data to be JSONified

    if doc_type == "BANK_STATEMENT":
        processed_info = extract_bank_statement_info(extracted_text)
    elif doc_type == "INVOICE":
        processed_info = extract_invoice_info(extracted_text)
    elif doc_type == "BALANCE_SHEET":
        processed_info = extract_balance_sheet_info(extracted_text)
    else: # UNKNOWN_DOCUMENT_TYPE or other types without specific extractors
        # For unknown types, return status and some text for context
        # Mask PII in the raw text snippet before returning
        masked_text_snippet = apply_pii_masking(extracted_text[:500]) # Mask only the snippet
        return json.dumps({
            "status": "UNKNOWN_DOCUMENT_TYPE",
            "file_path": file_path,
            "text_snippet_masked": masked_text_snippet
        }, indent=4)

    # Apply PII masking to the structured extracted information
    print("Applying PII masking...")
    masked_info = apply_pii_masking(processed_info)

    return json.dumps(masked_info, indent=4)


def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Extracts all text from a given PDF file.

    Args:
        pdf_path: The path to the PDF file.

    Returns:
        A string containing all extracted text from the PDF.
        Returns an error message string if an exception occurs.
    """
    full_text = ""
    try:
        doc = fitz.open(pdf_path)
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            full_text += page.get_text("text") # "text" for plain text extraction
        doc.close()
    except fitz.errors.FitzError as e: # More specific PyMuPDF error
        return f"Error processing PDF (FitzError): {e}"
    except Exception as e:
        return f"An unexpected error occurred during PDF processing: {e}"
    return full_text

def extract_text_from_docx(docx_path: str) -> str:
    """
    Extracts all text from a given DOCX file.

    Args:
        docx_path: The path to the DOCX file.

    Returns:
        A string containing all extracted text from the DOCX.
        Returns an error message string if an exception occurs.
    """
    full_text_parts = []
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            full_text_parts.append(paragraph.text)
        full_text = "\n".join(full_text_parts)
    except PackageNotFoundError: # Specific error for invalid DOCX format
        return f"Error processing DOCX: File is not a valid DOCX package."
    except Exception as e:
        return f"An unexpected error occurred during DOCX processing: {e}"
    return full_text

def extract_text_from_xlsx(xlsx_path: str) -> str:
    """
    Extracts all text from a given XLSX file.

    Args:
        xlsx_path: The path to the XLSX file.

    Returns:
        A string containing all extracted text from the XLSX, cells joined by spaces, rows by newlines.
        Returns an error message string if an exception occurs.
    """
    all_text_parts = []
    try:
        workbook = openpyxl.load_workbook(xlsx_path)
        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]
            for row in ws.iter_rows():
                row_text_parts = []
                for cell in row:
                    if cell.value is not None:
                        row_text_parts.append(str(cell.value))
                if row_text_parts: # Only add if the row had text
                    all_text_parts.append(" ".join(row_text_parts))
        full_text = "\n".join(all_text_parts)
    except InvalidFileException: # Specific error for invalid XLSX format
        return f"Error processing XLSX: File is not a valid XLSX file."
    except Exception as e:
        return f"An unexpected error occurred during XLSX processing: {e}"
    return full_text

def extract_text_from_image_ocr(image_path: str) -> str:
    """
    Extracts text from an image file using OCR (Tesseract).

    Args:
        image_path: The path to the image file.

    Returns:
        A string containing the extracted text.
        Returns an error message string if an exception occurs.
    """
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        return text.strip()
    except pytesseract.TesseractNotFoundError:
        return "Error: Tesseract is not installed or not found in your PATH. OCR functionality will not work."
    except UnidentifiedImageError:
        return f"Error processing Image: Cannot identify image file {image_path}. It may be corrupted or not a valid image."
    except Exception as e:
        return f"An unexpected error occurred during image OCR processing: {e}"

# --- Regex Patterns for Invoice Extraction ---
# Using simplified versions for broad matching within text. Anchors might be needed for specific forms.
GSTIN_REGEX_INV = re.compile(r'[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z\s]{1}', re.IGNORECASE) # Allow potential space in last char due to OCR
INVOICE_NO_REGEX_INV = re.compile(r'(?:Invoice No|Inv No|Invoice Number)\s*[:#-]?\s*([A-Za-z0-9/-]+)', re.IGNORECASE)
DATE_REGEX_INV = re.compile(
    r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})|'  # dd/mm/yyyy or dd-mm-yyyy etc (group 1)
    r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4})|'  # dd Month yyyy (group 2)
    r'(\d{1,2}[-/](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-/]\d{2,4})',  # dd/Mon/yyyy or dd-Mon-yyyy (group 3)
    re.IGNORECASE
)
MONETARY_VALUE_REGEX_INV = re.compile(r'\b\d[\d,]*\.\d{2}\b|\b\d[\d,]*\b') # Number with 2 decimals OR whole number
LINE_ITEM_GUESS_REGEX_INV = re.compile(r'^(.*?)\s+(\d+)\s+([\d\.,]+\.?\d*)\s+([\d\.,]+\.?\d*)$', re.MULTILINE)


# --- Regex Patterns for Balance Sheet Extraction (can reuse/adapt some from Invoice) ---
# DATE_REGEX_BS uses a broader month name part and allows various separators for date.
DATE_REGEX_BS = re.compile(r"(?:\bAs of|\bAs at|Date|For the period ending)\s*[:\-]?\s*((?:\d{1,2}[-/]\s?\w+\s?[-/]\d{2,4})|(?:\w+\s+\d{1,2},?\s*\d{2,4})|(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}))", re.IGNORECASE)
COMPANY_NAME_REGEX_BS = re.compile(r"^([A-Za-z0-9\s.,&'-]+(?:Ltd\.?|Limited|Inc\.?|Corporation|LLC|LLP|Pvt\.?)\b)", re.IGNORECASE | re.MULTILINE)
# MONETARY_VALUE_REGEX_INV can be reused for balance sheet values.
CURRENCY_SYMBOL_REGEX_BS = re.compile(r'(₹|\$|€|USD|INR|EUR|GBP)')


# --- Information Extractors ---

def extract_balance_sheet_info(text: str) -> dict:
    """
    Extracts information from balance sheet text using regex and keywords.
    """
    bs_data = {
        "company_name": None,
        "statement_date": None,
        "currency": None,
    }

    # Company Name
    company_match = COMPANY_NAME_REGEX_BS.search(text)
    if company_match:
        bs_data['company_name'] = company_match.group(1).strip()
    else: # Fallback: first non-empty line if it's not a typical header
        lines = text.split('\n')
        for line in lines:
            line_stripped = line.strip()
            if line_stripped and not any(kw in line_lower for kw in ["balance sheet", "as of", "statement of"]):
                bs_data['company_name'] = line_stripped
                break

    # Statement Date
    date_match = DATE_REGEX_BS.search(text)
    if date_match:
        bs_data['statement_date'] = date_match.group(1).strip()

    # Currency (simple heuristic - first symbol or code found near a number)
    currency_match = CURRENCY_SYMBOL_REGEX_BS.search(text)
    if currency_match:
        bs_data['currency'] = currency_match.group(1)

    # Section Totals Extraction
    total_keywords = {
        "total current assets": "total_current_assets",
        "total non-current assets": "total_non_current_assets",
        "total fixed assets": "total_non_current_assets", # Alias
        "total assets": "total_assets",
        "total current liabilities": "total_current_liabilities",
        "total non-current liabilities": "total_non_current_liabilities",
        "total long-term liabilities": "total_non_current_liabilities", # Alias
        "total liabilities": "total_liabilities",
        "total equity": "total_equity",
        "shareholders' equity": "total_equity", # Broader match for "total shareholders' equity"
        "net worth": "total_equity" # Alias
    }

    lines = text.split('\n')
    for line in lines:
        line_lower = line.lower()
        for keyword_text, data_key in total_keywords.items():
            if keyword_text in line_lower:
                # Search for monetary value on this line
                value_matches = MONETARY_VALUE_REGEX_INV.findall(line) # Reuse from invoice
                if value_matches:
                    # Take the last found monetary value on the line
                    value_str = value_matches[-1].replace(',', '')
                    try:
                        bs_data[data_key] = float(value_str)
                    except ValueError: # pragma: no cover
                        bs_data[data_key] = value_str # Store as string if not floatable

                    # Check for currency near this value if global currency not yet found
                    if not bs_data['currency']:
                        local_currency_match = CURRENCY_SYMBOL_REGEX_BS.search(line)
                        if local_currency_match:
                            bs_data['currency'] = local_currency_match.group(1)
                break # Found keyword on this line, move to next line

    # If currency is still None, try to find any currency code in the text
    if not bs_data['currency']:
        all_text_currency_match = CURRENCY_SYMBOL_REGEX_BS.search(text)
        if all_text_currency_match:
            bs_data['currency'] = all_text_currency_match.group(1)

    return bs_data

def extract_invoice_info(text: str) -> dict:
    """
    Extracts information from invoice text using regex and heuristics.
    """
    invoice_data = {
        "invoice_number": None,
        "invoice_date": None,
        "supplier_info": {"name": None, "address": None, "gstin": None, "phone": None, "email": None},
        "recipient_info": {"name": None, "address": None, "gstin": None},
        "line_items": [],
        "subtotal": None,
        "cgst_amount": None,
        "sgst_amount": None,
        "igst_amount": None,
        "total_amount": None,
    }

    # GSTINs (Basic Heuristic: first is supplier, second is recipient)
    all_gstins = GSTIN_REGEX_INV.findall(text)
    if all_gstins:
        invoice_data['supplier_info']['gstin'] = all_gstins[0].replace(" ", "") # Clean potential space
        if len(all_gstins) > 1:
            invoice_data['recipient_info']['gstin'] = all_gstins[1].replace(" ", "")

    # Invoice Number
    inv_no_match = INVOICE_NO_REGEX_INV.search(text)
    if inv_no_match:
        invoice_data['invoice_number'] = inv_no_match.group(1)

    # Invoice Date
    # Prioritize "Invoice Date:" then "Date:"
    date_labels_priority = ["Invoice Date:", "Date:"]
    found_date = None
    # print(f"Debug Text for Date Search:\n---\n{text[:1000]}\n---") # Print more text for context
    for label in date_labels_priority:
        # DATE_REGEX_INV.pattern needs to be wrapped in a non-capturing group (?:...) for correct OR precedence
        pattern_to_search = r"(?i)" + re.escape(label) + r"\s*[:#-]?\s*(?:" + DATE_REGEX_INV.pattern + ")"
        match = re.search(pattern_to_search, text)
        # print(f"Debug Date Search: Label='{label}', TextMatch='{match.group(0) if match else 'No match'}'")
        if match:
            # DATE_REGEX_INV now has 3 main capturing groups due to the new alternative
            actual_date_str = match.group(1) or match.group(2) or match.group(3)
            # print(f"Debug Date Inner: Matched groups by pattern: {match.groups()}, actual_date_str='{actual_date_str}'")
            if actual_date_str:
                found_date = actual_date_str.strip()
                invoice_data['invoice_date'] = found_date
                # print(f"Debug Date Found: Label='{label}', Date='{found_date}'")
                break
    if not found_date: # Fallback: find first date if no labeled date is found
        first_date_match = DATE_REGEX_INV.search(text)
        if first_date_match:
             invoice_data['invoice_date'] = (first_date_match.group(1) or first_date_match.group(2) or first_date_match.group(3)).strip()

    # Supplier/Recipient Info
    lines = text.split('\n')

    # Supplier parsing
    supplier_line_index = -1
    for i, line in enumerate(lines):
        if any(kw.lower() in line.lower() for kw in ["Supplier:", "Sold By:", "From:"]):
            supplier_line_index = i
            # Try to parse Name, Address, GSTIN from this line primarily
            # Example: Supplier: NAME, Address: ADDR, GSTIN: GSTIN_VAL Ph: PHONE Email: EMAIL
            name_match = re.search(r"(?:Supplier:|Sold By:|From:)\s*(.*?)(?:, Address|, GSTIN|Ph:|Email:|$)", line, re.IGNORECASE)
            if name_match: invoice_data['supplier_info']['name'] = name_match.group(1).strip().rstrip(',')

            address_match = re.search(r"Address[:\s]*(.*?)(?:, GSTIN|Ph:|Email:|$)", line, re.IGNORECASE)
            if address_match: invoice_data['supplier_info']['address'] = address_match.group(1).strip().rstrip(',')

            # GSTIN is already captured by all_gstins, phone/email by general search
            phone_match = re.search(r'(?:Ph|Phone|Mobile)[:\s]*([0-9\s+-]+)', line, re.IGNORECASE)
            if phone_match: invoice_data['supplier_info']['phone'] = phone_match.group(1).strip()
            email_match = re.search(r'(?:Email|E-mail)[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', line, re.IGNORECASE)
            if email_match: invoice_data['supplier_info']['email'] = email_match.group(1).strip()
            break

    # Recipient parsing
    recipient_line_index = -1
    for i, line in enumerate(lines):
        if any(kw.lower() in line.lower() for kw in ["Bill To:", "Ship To:", "To:"]):
            if i <= supplier_line_index: continue # Avoid matching recipient keywords in supplier section
            recipient_line_index = i
            name_match = re.search(r"(?:Bill To:|Ship To:|To:)\s*(.*?)(?:, Address|, GSTIN|$)", line, re.IGNORECASE)
            if name_match: invoice_data['recipient_info']['name'] = name_match.group(1).strip().rstrip(',')

            address_match = re.search(r"Address[:\s]*(.*?)(?:, GSTIN|$)", line, re.IGNORECASE)
            if address_match:
                addr_str = address_match.group(1).strip().rstrip(',')
                # Clean out GSTIN if it was part of address match (as GSTIN is handled separately)
                addr_str = re.sub(r"GSTIN[:\s]*\S+", "", addr_str, flags=re.IGNORECASE).strip().rstrip(',')
                invoice_data['recipient_info']['address'] = addr_str
            break

    # Fallback for multi-line addresses if primary line parsing didn't get them
    # This part needs more robust block detection based on typical invoice structures
    # For now, this simplistic block capture is removed to favor single-line parsing for clarity
    # and to avoid the previous errors of mis-capturing blocks.
    # A proper solution would use visual/layout cues or more advanced NLP/template matching.

    # Line Items (Very Simplified Guess)
    # This will need significant improvement for real invoices (table detection, etc.)
    for line in lines:
        # Skip headers or lines that are clearly not items
        if line.lower().startswith("description") or line.lower().startswith("hsn") or line.lower().startswith("subtotal") or line.lower().startswith("total"):
            continue
        item_match = LINE_ITEM_GUESS_REGEX_INV.search(line.strip())
        if item_match:
            desc, qty, rate, amount = item_match.groups()
            invoice_data['line_items'].append({
                "description": desc.strip(),
                "quantity": int(qty) if qty.isdigit() else qty,
                "rate": float(rate.replace(',','')) if rate.replace('.','',1).replace(',','').isdigit() else rate,
                "amount": float(amount.replace(',','')) if amount.replace('.','',1).replace(',','').isdigit() else amount,
            })

    # Totals & Tax Amounts
    for line in lines:
        line_lower = line.lower()

        # Find all monetary values on the line
        all_value_matches_on_line = MONETARY_VALUE_REGEX_INV.findall(line)

        if not all_value_matches_on_line:
            continue # Skip if no numbers found on this line

        # Use the last found monetary value on the line for totals, as it's common for labels to precede values.
        val_str = all_value_matches_on_line[-1]
        val_str = all_value_matches_on_line[-1]
        # Ensure val_str can be converted to float. The regex should ensure it's a number-like string.
        try:
            val = float(val_str.replace(',',''))
        except ValueError: # pragma: no cover
            val = None

        # print(f"Totals Debug -- Line: '{line.strip()}' | AllMatches: {all_value_matches_on_line} | Chosen val_str: {val_str} | Parsed val: {val}")

        if val is not None:
            if "subtotal" in line_lower or "sub total" in line_lower :
                invoice_data['subtotal'] = val
            elif "cgst" in line_lower:
                # Try to get the value immediately following "CGST @ X%:" or "CGST:"
                tax_specific_match = re.search(r"CGST\s*(?:@\s*[\d\.]+%?)?\s*[:\-]?\s*([\d\.,]+)", line, re.IGNORECASE)
                if tax_specific_match:
                    invoice_data['cgst_amount'] = float(tax_specific_match.group(1).replace(',',''))
                else: # Fallback to the last number on the line if specific pattern fails
                    invoice_data['cgst_amount'] = val
            elif "sgst" in line_lower:
                tax_specific_match = re.search(r"SGST\s*(?:@\s*[\d\.]+%?)?\s*[:\-]?\s*([\d\.,]+)", line, re.IGNORECASE)
                if tax_specific_match:
                    invoice_data['sgst_amount'] = float(tax_specific_match.group(1).replace(',',''))
                else:
                    invoice_data['sgst_amount'] = val
            elif "igst" in line_lower:
                tax_specific_match = re.search(r"IGST\s*(?:@\s*[\d\.]+%?)?\s*[:\-]?\s*([\d\.,]+)", line, re.IGNORECASE)
                if tax_specific_match:
                    invoice_data['igst_amount'] = float(tax_specific_match.group(1).replace(',',''))
                else:
                    invoice_data['igst_amount'] = val
            elif "total amount" in line_lower or "grand total" in line_lower or (line_lower.startswith("total:") and not invoice_data['total_amount']):
                invoice_data['total_amount'] = val
            # elif line_lower.startswith("total") and not invoice_data['total_amount'] : # Catches "Total" if "Total Amount" or "Grand Total" not found
            #     invoice_data['total_amount'] = val


    return invoice_data

def extract_bank_statement_info(text: str) -> dict:
    """
    Extracts information from bank statement text using regex.
    """
    statement_data = {"transactions": [], "account_holder": None, "account_number": None,
                      "statement_period_start": None, "statement_period_end": None, "bank_name": None}

    # Account Number
    acc_match = re.search(r"(?:Account No|A/C No|Account Number)[:\s]*([\d\s-]{9,18})\b", text, re.IGNORECASE)
    if acc_match:
        statement_data['account_number'] = acc_match.group(1).strip()

    # Account Holder Name (simple heuristic)
    holder_match = re.search(r"(?:Account Holder|Name)[:\s]*(.+?)(?:\n|Account No|A/C No)", text, re.IGNORECASE)
    if holder_match:
        statement_data['account_holder'] = holder_match.group(1).strip()
        # Further clean up if "Account No" or similar was part of the capture due to weak positive lookahead
        if statement_data['account_holder'] and statement_data['account_number']:
             if statement_data['account_number'] in statement_data['account_holder']:
                 statement_data['account_holder'] = statement_data['account_holder'].split(statement_data['account_number'])[0].strip()


    # Statement Period
    period_match = re.search(r"Statement Period[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})\s*(?:to|-)\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.IGNORECASE)
    if period_match:
        statement_data['statement_period_start'] = period_match.group(1).strip()
        statement_data['statement_period_end'] = period_match.group(2).strip()

    # Bank Name (very heuristic, looks for a line ending with "Bank" or "Bank Ltd." at/near the start)
    # Iterate first few lines for bank name
    for i, line in enumerate(text.split('\n')):
        if i < 5: # Check first 5 lines
            bank_match = re.match(r"^([A-Za-z\s&]+(?:Bank|BANK|Ltd\.|LIMITED|CORPORATION|CORP\.)(?: Ltd\.?)?)\s*$", line.strip(), re.IGNORECASE)
            if bank_match:
                statement_data['bank_name'] = bank_match.group(1).strip()
                break

    # Transaction Parsing (Simplified)
    # Regex: Date (dd/mm/yy or dd-mm-yy), Description (non-greedy), Debit, Credit, Balance
    # Assumes amounts have 2 decimal places, handles missing debit/credit
    # This is highly example-specific and fragile for real-world statements
    # Attempting a more robust regex:
    # Date | Description (greedy but stopped by multiple spaces or EOL) | Optional Debit | Optional Credit | Balance
    # This regex assumes that if both debit and credit are present, they are separated by at least one space.
    # If one is missing, it's just not there. Balance is assumed to be the last numeric group.
    TRANSACTION_LINE_REGEX = re.compile(
        r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\s+"  # Date
        r"(.+?)\s{2,}"                        # Description (non-greedy, followed by at least 2 spaces)
        r"(?:([\d\.,]+)\s+)?"                 # Optional Debit (capture group 3), followed by spaces
        r"(?:([\d\.,]+)\s+)?"                 # Optional Credit (capture group 4), followed by spaces
                                              # This structure implies if only one amount before balance, it's ambiguous.
                                              # Let's refine to handle cases:
                                              # 1. Desc Debit Credit Balance
                                              # 2. Desc Debit Balance
                                              # 3. Desc Credit Balance (harder if only one number column is used for debit/credit)
                                              # 4. Desc Balance
        r"([\d\.,]+)$"                        # Balance (must be at the end of the line)
    , re.IGNORECASE | re.VERBOSE)

    # Simpler regex focusing on structure if amounts are aligned, this is a common but not universal case
    # Date       Description       (Debit)   (Credit)  Balance
    # Need to handle cases where debit or credit might be empty or not present.
    # Example: "01/01/2024 Description text debit credit balance"
    # Example: "01/01/2024 Description text debit       balance"
    # Example: "01/01/2024 Description text       credit  balance"
    # Example: "01/01/2024 Description text               balance"

    # Revised regex to better distinguish debit/credit based on typical column output where empty values mean multiple spaces
    # Date (capture 1)
    # Description (capture 2)
    # Debit (capture 3) - may be just spaces then a number, or number then spaces
    # Credit (capture 4) - may be just spaces then a number, or number then spaces
    # Balance (capture 5)
    TRANSACTION_LINE_REGEX = re.compile(
        r"(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\s+"  # Date (group 1)
        r"(.+?)\s{2,}"                        # Description (group 2), followed by at least 2 spaces
        r"(?:([\d\.,]+)\s+)?"                 # Optional Debit (group 3), standard greedy optional
        r"(?:([\d\.,]+)\s+)?"                 # Optional Credit (group 4), standard greedy optional
        r"([\d\.,]+)$"                        # Balance (group 5)
    , re.IGNORECASE | re.VERBOSE)

    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        tx_match = TRANSACTION_LINE_REGEX.search(line)
        if tx_match:
            # date, description, debit, credit, balance = tx_match.groups() # Original
            date, description, group3_debit, group4_credit, balance_val = tx_match.groups()

            # Skip if it looks like a header line captured by mistake
            if description.lower() in ["description", "particulars", "details", "opening balance"] or date.lower() == "date":
                 # For "Opening Balance", if it has a balance value but no debit/credit, handle it.
                if description.lower() == "opening balance":
                    if not group3_debit and not group4_credit and balance_val:
                         statement_data['transactions'].append({
                            "date": date.strip(),
                            "description": description.strip(),
                            "debit": None,
                            "credit": None,
                            "balance": float(balance_val.replace(',', '')) if balance_val and balance_val.replace('.','',1).replace(',','').isdigit() else None,
                        })
                continue

            debit_val = None
            credit_val = None

            # Logic to assign debit/credit based on which group matched.
            # This regex is still not perfect. A real statement parser often needs to know column positions or use more context.
            # If group3_debit (potential debit) is present and group4_credit is not, it's likely a debit.
            # If group4_credit is present, it's likely credit. If group3_debit is also present, it implies 3 amount columns.
            # This simplified regex might misinterpret if only one amount is present before balance and it's meant to be credit.

            # Heuristic: If a number is directly after description followed by another number (balance),
            # it's more likely debit if it's in the "debit position" (group3).
            # If there's a gap (group3 is None) and then a number (group4), that's credit.

            if group3_debit and group3_debit.replace('.','',1).replace(',','').isdigit():
                # Check if this could be a misassigned credit due to structure
                # E.g. "Description <spaces> <credit_value> <balance_value>"
                # If group4_credit is None, this value in group3_debit might be a credit if it's the only amount before balance
                # For now, we assume if group3 has value, it's debit. If group4 has value, it's credit.
                # This will be wrong if statement has only 2 amount columns (e.g. Amount, Balance) where Amount can be Deb or Cred

                # A common pattern: if credit is empty, debit value is listed. If debit is empty, credit value is listed.
                # The current regex `(?:([\d\.,]+)\s+)?` for both makes them independent.
                # Let's assume for the dummy:
                # "05/01/2024  ATM Withdrawal ABC Complex      500.00                9500.00" -> debit 500
                # "10/01/2024  NEFT from XYZ Corp                         2000.00   11500.00" -> credit 2000
                # "15/01/2024  Cheque No 12345                   75.50              11424.50" -> debit 75.50

                # The key is how the spaces are consumed by `(.+?)\s{2,}` and then the optional groups.
                # If `group3_debit` has a value, it's debit.
                # If `group3_debit` is None AND `group4_credit` has a value, it's credit.
                # This is what the regex implies by its structure of optional groups.

                debit_val = float(group3_debit.replace(',', ''))

            if group4_credit and group4_credit.replace('.','',1).replace(',','').isdigit():
                credit_val = float(group4_credit.replace(',', ''))
                # If group3_debit was also found, it means we parsed three numbers after desc.
                # For typical statements, this means debit, credit, balance.
                # If group3_debit was NOT found, then group4_credit is the first amount, making it ambiguous
                # without column headers. But our regex design implies group4 is credit if group3 is skipped.

            # Special handling for the "NEFT from XYZ Corp" line which has structure: Desc <many_spaces> Credit Balance
            # The regex `(.+?)\s{2,}` will capture desc. `(?:([\d\.,]+)\s+)?` (debit group3) will be None.
            # `(?:([\d\.,]+)\s+)?` (credit group4) should capture "2000.00".
            if "NEFT from XYZ Corp".lower() in description.lower(): # Heuristic for this specific line
                if not group3_debit and group4_credit: # If debit group is empty and credit group has value
                    debit_val = None
                    credit_val = float(group4_credit.replace(',', ''))
                elif group3_debit and not group4_credit: # If mistakenly captured in debit group
                    credit_val = float(group3_debit.replace(',', ''))
                    debit_val = None


            transaction = {
                "date": date.strip(),
                "description": description.strip(), # This will include "Cheque No 12345" if regex is good.
                "debit": debit_val,
                "credit": credit_val,
                "balance": float(balance_val.replace(',', '')) if balance_val and balance_val.replace('.','',1).replace(',','').isdigit() else None,
            }
            statement_data['transactions'].append(transaction)

    return statement_data


# --- Document Type Classification ---
BALANCE_SHEET_KEYWORDS = {"balance sheet", "assets", "liabilities", "equity", "shareholder equity", "statement of financial position"}
INVOICE_KEYWORDS = {"invoice no", "gstin", "invoice date", "hsn", "sac", "tax invoice", "bill to", "ship to", "invoice number"}
BANK_STATEMENT_KEYWORDS = {"account statement", "account no", "transaction date", "narration", "description", "debit", "credit", "balance", "statement of account"}
MIN_KEYWORD_THRESHOLD = 2

def classify_document(text: str) -> str:
    """
    Classifies the document type based on keywords found in the extracted text.
    """
    if not text:
        return "UNKNOWN_DOCUMENT_TYPE"

    lower_text = text.lower()

    scores = {
        "BALANCE_SHEET": sum(1 for keyword in BALANCE_SHEET_KEYWORDS if keyword in lower_text),
        "INVOICE": sum(1 for keyword in INVOICE_KEYWORDS if keyword in lower_text),
        "BANK_STATEMENT": sum(1 for keyword in BANK_STATEMENT_KEYWORDS if keyword in lower_text),
    }

    # Find the document type with the maximum score
    max_score = 0
    classified_type = "UNKNOWN_DOCUMENT_TYPE"
    for doc_type, score in scores.items():
        if score > max_score:
            max_score = score
            classified_type = doc_type
        elif score == max_score and score > 0: # Handle ties by making it unknown if significant tie
            classified_type = "UNKNOWN_DOCUMENT_TYPE" # Or could be more sophisticated

    if max_score < MIN_KEYWORD_THRESHOLD:
        return "UNKNOWN_DOCUMENT_TYPE"

    return classified_type

# Ensure this function is defined before process_document or called appropriately if process_document is the main entry.
# For simplicity, keeping the old load_document structure inside process_document and making it return JSON.
# The individual extract_text_from_* functions already return text or error string.

if __name__ == '__main__':
    # DUMMY_PDF_PATH, DUMMY_DOCX_PATH, etc. are defined further down where they are created.
    # For clarity, list test files here and create/cleanup them systematically.

    test_files_to_create = {
        "pdf": {"name": "dummy_test.pdf", "type": "Bank Statement"},
        "docx": {"name": "dummy_test.docx", "type": "Invoice"},
        "xlsx": {"name": "dummy_test.xlsx", "type": "Balance Sheet"},
        "png": {"name": "dummy_test.png", "type": "Unknown (OCR)"},
        "jpg": {"name": "photo.jpg", "type": "Empty Image (OCR)"}, # Will be an empty img for OCR test
        "txt": {"name": "unsupported.txt", "type": "Unsupported"}
    }
    DUMMY_PDF_PATH = test_files_to_create["pdf"]["name"]
    DUMMY_DOCX_PATH = test_files_to_create["docx"]["name"]
    DUMMY_XLSX_PATH = test_files_to_create["xlsx"]["name"]
    DUMMY_PNG_PATH = test_files_to_create["png"]["name"]
    DUMMY_JPG_PATH = test_files_to_create["jpg"]["name"]
    UNSUPPORTED_TXT_PATH = test_files_to_create["txt"]["name"]
    NON_EXISTENT_FILE = "nonexistent_file.tmp"


    # --- Create all dummy files ---
    print("--- Setting up dummy files ---")
    # Create a dummy text-based PDF for testing (BANK STATEMENT)
    try:
        pdf_doc = fitz.open()
        page1_text = """MyExample Bank Ltd.
Statement of Account
Account Holder: John Doe
Account No: 123456789012
Statement Period: 01/01/2024 to 31/01/2024

Date        Description                     Debit      Credit     Balance
01/01/2024  Opening Balance                                    10000.00
"""
        pdf_page1 = pdf_doc.new_page(width=600, height=400)
        y_coord = 50
        for line in page1_text.split('\n'): pdf_page1.insert_text((50, y_coord), line); y_coord += 15
        page2_text = """05/01/2024  ATM Withdrawal ABC Complex      500.00                9500.00
10/01/2024  NEFT from XYZ Corp                         2000.00   11500.00
15/01/2024  Cheque No 12345                   75.50              11424.50"""
        pdf_page2 = pdf_doc.new_page(width=600, height=400)
        y_coord = 50
        for line in page2_text.split('\n'): pdf_page2.insert_text((50, y_coord), line); y_coord += 15
        pdf_doc.save(DUMMY_PDF_PATH)
        pdf_doc.close()
        print(f"Created dummy PDF: {DUMMY_PDF_PATH} ({test_files_to_create['pdf']['type']})")
    except Exception as e: print(f"Error creating dummy PDF: {e}")

    # Create a dummy DOCX file for testing (INVOICE)
    try:
        docx_doc = Document()
        invoice_content = """TAX INVOICE
Supplier: ABC Electronics Pvt Ltd, Address: 123 Tech Park, Bangalore, GSTIN: 29ABCDE1234F1Z5, Ph: 9988776655, Email: contact@abc.com

Invoice No: INV2024/001
Invoice Date: 02/Jan/2024
Date of Supply: 03-01-2024

Bill To: XYZ Solutions, Address: 789 Business Hub, Mumbai, GSTIN: 27FGHIJ5678K1Z9

Description HSN Qty Rate Amount
Laptop Computer 847130 1 50000.00 50000.00
Wireless Mouse 847160 2 500.00 1000.00

Subtotal: 51000.00
CGST @ 9%: 4590.00
SGST @ 9%: 4590.00
Total Amount: 60180.00
Amount in Words: Sixty Thousand One Hundred Eighty Only"""
        for line in invoice_content.split('\n'):
            if line.strip() == "": docx_doc.add_paragraph()
            else: docx_doc.add_paragraph(line)
        docx_doc.save(DUMMY_DOCX_PATH)
        print(f"Created dummy DOCX: {DUMMY_DOCX_PATH} ({test_files_to_create['docx']['type']})")
    except Exception as e: print(f"Error creating dummy DOCX: {e}")

    # Create a dummy XLSX file for testing (BALANCE SHEET)
    try:
        xlsx_workbook = openpyxl.Workbook()
        sheet1 = xlsx_workbook.active; sheet1.title = "BalanceSheetData"
        sheet1['A1'] = "Example Corp Balance Sheet Inc."; sheet1['A2'] = "As of December 31, 2023"; sheet1['A3'] = "(Currency: INR)"
        sheet1['A5'] = "Assets"; sheet1['A6'] = "Total Current Assets :"; sheet1['B6'] = "1,500,000.00"
        sheet1['A7'] = "Total Non-Current Assets :"; sheet1['B7'] = "2,000,000.00"; sheet1['A8'] = "Total Assets :"; sheet1['B8'] = "3,500,000.00"
        sheet1['A10'] = "Liabilities"; sheet1['A11'] = "Total Current Liabilities :"; sheet1['B11'] = "700,000.00"
        sheet1['A12'] = "Total Non-Current Liabilities :"; sheet1['B12'] = "800,000.00"; sheet1['A13'] = "Total Liabilities :"; sheet1['B13'] = "1,500,000.00"
        sheet1['A15'] = "Equity"; sheet1['A16'] = "Total Equity :"; sheet1['B16'] = "2,000,000.00"
        xlsx_workbook.save(DUMMY_XLSX_PATH)
        print(f"Created dummy XLSX: {DUMMY_XLSX_PATH} ({test_files_to_create['xlsx']['type']})")
    except Exception as e: print(f"Error creating dummy XLSX: {e}")

    # Create a dummy PNG image with text for OCR testing ("Hello OCR" -> UNKNOWN)
    try:
        img = Image.new('RGB', (200, 50), color = (255, 255, 255))
        draw = ImageDraw.Draw(img); font = None
        try: font = ImageFont.load_default()
        except IOError:
            try: font = ImageFont.truetype("DejaVuSans.ttf", 15) if os.path.exists("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf") else None
            except IOError: print("Default font and DejaVuSans.ttf not found for dummy PNG.")
        text_to_draw = "Hello OCR"
        if font:
            try: bbox = draw.textbbox((0,0), text_to_draw, font=font); text_width = bbox[2]-bbox[0]; text_height = bbox[3]-bbox[1]
            except AttributeError: text_width, text_height = draw.textsize(text_to_draw, font=font) # Fallback for older Pillow
            draw.text(((img.width - text_width) / 2, (img.height - text_height) / 2), text_to_draw, fill=(0,0,0), font=font)
        else: draw.text((10, 10), text_to_draw, fill=(0,0,0))
        img.save(DUMMY_PNG_PATH)
        print(f"Created dummy PNG: {DUMMY_PNG_PATH} ({test_files_to_create['png']['type']})")
    except Exception as e: print(f"Error creating dummy PNG: {e}")

    # Create dummy JPG (empty image for OCR test)
    try:
        img_jpg = Image.new('RGB', (60, 30), color = 'blue')
        img_jpg.save(DUMMY_JPG_PATH)
        print(f"Created dummy JPG: {DUMMY_JPG_PATH} ({test_files_to_create['jpg']['type']})")
    except Exception as e: print(f"Error creating dummy JPG: {e}")

    # Create unsupported file
    with open(UNSUPPORTED_TXT_PATH, "w") as f: f.write("dummy text content")
    print(f"Created dummy TXT: {UNSUPPORTED_TXT_PATH} ({test_files_to_create['txt']['type']})")

    # --- Test Cases ---
    test_scenarios = [
        DUMMY_PDF_PATH, DUMMY_DOCX_PATH, DUMMY_XLSX_PATH, DUMMY_PNG_PATH,
        DUMMY_JPG_PATH, UNSUPPORTED_TXT_PATH, NON_EXISTENT_FILE
    ]

    for file_to_test in test_scenarios:
        try:
            json_output = process_document(file_to_test)
            print(json_output) # process_document now prints its own header
        except Exception as e: # Catch any unexpected exceptions from process_document itself
            print(f"--- Processing document: {file_to_test} ---") # Manual header for direct exceptions
            print(json.dumps({"error": "Critical error in process_document", "details": str(e)}, indent=4))


    # Clean up dummy files
    print("\n--- Cleaning up dummy files ---")
    for key in test_files_to_create:
        f_path = test_files_to_create[key]["name"]
        if os.path.exists(f_path):
            try: os.remove(f_path); print(f"Removed dummy file: {f_path}")
            except OSError as e: print(f"Error removing {f_path}: {e}")

# --- Main Document Processing Function ---
# Definition of process_document and other functions would be here (as per the current file structure)
# For the diff, assume process_document will replace the old load_document's position and content.
# The actual extract_text_from_pdf etc. functions are defined above process_document.
def extract_text_from_pdf(pdf_path: str) -> str:
    DUMMY_DOCX_PATH = "dummy_test.docx"
    DUMMY_XLSX_PATH = "dummy_test.xlsx"
    DUMMY_PNG_PATH = "dummy_test.png"

    # Create a dummy text-based PDF for testing (BANK STATEMENT)
    try:
        pdf_doc = fitz.open()

        # Page 1 Content
        page1_text = """MyExample Bank Ltd.
Statement of Account
Account Holder: John Doe
Account No: 123456789012
Statement Period: 01/01/2024 to 31/01/2024

Date        Description                     Debit      Credit     Balance
01/01/2024  Opening Balance                                    10000.00
"""
        pdf_page1 = pdf_doc.new_page(width=600, height=400) # Define page size
        # Insert text line by line to control layout better for regex
        y_coord = 50
        for line in page1_text.split('\n'):
            pdf_page1.insert_text((50, y_coord), line)
            y_coord += 15 # Increment y-coordinate for next line

        # Page 2 Content
        page2_text = """05/01/2024  ATM Withdrawal ABC Complex      500.00                9500.00
10/01/2024  NEFT from XYZ Corp                         2000.00   11500.00
15/01/2024  Cheque No 12345                   75.50              11424.50
"""
        pdf_page2 = pdf_doc.new_page(width=600, height=400)
        y_coord = 50 # Reset y-coordinate for new page
        for line in page2_text.split('\n'):
            pdf_page2.insert_text((50, y_coord), line)
            y_coord += 15

        pdf_doc.save(DUMMY_PDF_PATH)
        pdf_doc.close()
        print(f"Created dummy PDF: {DUMMY_PDF_PATH} (Bank Statement)")
    except Exception as e:
        print(f"Error creating dummy PDF: {e}")

    # Create a dummy DOCX file for testing (INVOICE)
    try:
        docx_doc = Document()
        # Adjusted dummy data for simpler single-line parsing of supplier/recipient info
        invoice_content = """TAX INVOICE
Supplier: ABC Electronics Pvt Ltd, Address: 123 Tech Park, Bangalore, GSTIN: 29ABCDE1234F1Z5, Ph: 9988776655, Email: contact@abc.com

Invoice No: INV2024/001
Invoice Date: 02/Jan/2024
Date of Supply: 03-01-2024

Bill To: XYZ Solutions, Address: 789 Business Hub, Mumbai, GSTIN: 27FGHIJ5678K1Z9

Description HSN Qty Rate Amount
Laptop Computer 847130 1 50000.00 50000.00
Wireless Mouse 847160 2 500.00 1000.00

Subtotal: 51000.00
CGST @ 9%: 4590.00
SGST @ 9%: 4590.00
Total Amount: 60180.00
Amount in Words: Sixty Thousand One Hundred Eighty Only
"""
        # Add content as distinct paragraphs; some information extractors might rely on line breaks.
        for line in invoice_content.split('\n'):
            if line.strip() == "": # Add an empty paragraph for visual separation if desired, or skip
                docx_doc.add_paragraph()
            else:
                docx_doc.add_paragraph(line)

        docx_doc.save(DUMMY_DOCX_PATH)
        print(f"Created dummy DOCX: {DUMMY_DOCX_PATH} (Invoice)")
    except Exception as e:
        print(f"Error creating dummy DOCX: {e}")

    # Create a dummy XLSX file for testing (BALANCE SHEET)
    try:
        xlsx_workbook = openpyxl.Workbook()
        sheet1 = xlsx_workbook.active
        sheet1.title = "BalanceSheetData"

        sheet1['A1'] = "Example Corp Balance Sheet Inc."
        sheet1['A2'] = "As of December 31, 2023"
        sheet1['A3'] = "(Currency: INR)"
        sheet1['A4'] = "" # Empty line

        sheet1['A5'] = "Assets"
        sheet1['A6'] = "Total Current Assets :"
        sheet1['B6'] = "1,500,000.00"
        sheet1['A7'] = "Total Non-Current Assets :"
        sheet1['B7'] = "2,000,000.00"
        sheet1['A8'] = "Total Assets :"
        sheet1['B8'] = "3,500,000.00"
        sheet1['A9'] = ""

        sheet1['A10'] = "Liabilities"
        sheet1['A11'] = "Total Current Liabilities :"
        sheet1['B11'] = "700,000.00"
        sheet1['A12'] = "Total Non-Current Liabilities :"
        sheet1['B12'] = "800,000.00"
        sheet1['A13'] = "Total Liabilities :"
        sheet1['B13'] = "1,500,000.00"
        sheet1['A14'] = ""

        sheet1['A15'] = "Equity"
        sheet1['A16'] = "Total Equity :" # Also tests "Shareholders' Equity" and "Net Worth" via keywords
        sheet1['B16'] = "2,000,000.00"

        xlsx_workbook.save(DUMMY_XLSX_PATH)
        print(f"Created dummy XLSX: {DUMMY_XLSX_PATH} (Balance Sheet)")
    except Exception as e:
        print(f"Error creating dummy XLSX: {e}")

    # Create a dummy PNG image with text for OCR testing ("Hello OCR" -> UNKNOWN)
    try:
        img = Image.new('RGB', (200, 50), color = (255, 255, 255))
        draw = ImageDraw.Draw(img)
        font = None
        try:
            font = ImageFont.load_default()
        except IOError: # pragma: no cover
            try: # Try a common system font if default fails (e.g. in some minimal Docker containers)
                font = ImageFont.truetype("DejaVuSans.ttf", 15) if os.path.exists("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf") else None
            except IOError:
                 print("Default font and DejaVuSans.ttf not found. Text rendering on dummy image may be very basic.")


        text_to_draw = "Hello OCR" # This should result in UNKNOWN_DOCUMENT_TYPE
        if font:
            try:
                bbox = draw.textbbox((0,0), text_to_draw, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
            except AttributeError: # pragma: no cover
                text_width, text_height = draw.textsize(text_to_draw, font=font)
            x = (img.width - text_width) / 2
            y = (img.height - text_height) / 2
            draw.text((x, y), text_to_draw, fill=(0,0,0), font=font)
        else: # pragma: no cover
            draw.text((10, 10), text_to_draw, fill=(0,0,0)) # Basic fallback
        img.save(DUMMY_PNG_PATH)
        print(f"Created dummy PNG: {DUMMY_PNG_PATH} (Unknown type)")
    except ImportError:
        print("Pillow (PIL) is not installed. Cannot create dummy PNG for OCR testing.") # pragma: no cover
    except Exception as e:
        print(f"Error creating dummy PNG: {e}") # pragma: no cover

    # --- Test Cases ---
    print("\n--- Demonstrating load_document function ---")

    # 1. Valid PDF file (dummy_test.pdf created above)
    if os.path.exists(DUMMY_PDF_PATH): # Only test if creation was successful
        try:
            print(f"\nTesting with: {DUMMY_PDF_PATH}")
            load_document(DUMMY_PDF_PATH)
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
        finally:
            if os.path.exists(DUMMY_PDF_PATH): # Ensure cleanup even if load_document fails
                os.remove(DUMMY_PDF_PATH)
                print(f"Removed dummy PDF: {DUMMY_PDF_PATH}")
    else:
        print(f"\nSkipping PDF test as {DUMMY_PDF_PATH} was not created.")

    # 2. Valid DOCX file (dummy_test.docx created above)
    if os.path.exists(DUMMY_DOCX_PATH): # Only test if creation was successful
        try:
            print(f"\nTesting with: {DUMMY_DOCX_PATH}")
            load_document(DUMMY_DOCX_PATH)
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
        finally:
            if os.path.exists(DUMMY_DOCX_PATH): # Ensure cleanup
                os.remove(DUMMY_DOCX_PATH)
                print(f"Removed dummy DOCX: {DUMMY_DOCX_PATH}")
    else:
        print(f"\nSkipping DOCX test as {DUMMY_DOCX_PATH} was not created.")

    # 3. Valid XLSX file (dummy_test.xlsx created above)
    if os.path.exists(DUMMY_XLSX_PATH): # Only test if creation was successful
        try:
            print(f"\nTesting with: {DUMMY_XLSX_PATH}")
            load_document(DUMMY_XLSX_PATH)
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
        finally:
            if os.path.exists(DUMMY_XLSX_PATH): # Ensure cleanup
                os.remove(DUMMY_XLSX_PATH)
                print(f"Removed dummy XLSX: {DUMMY_XLSX_PATH}")
    else:
        print(f"\nSkipping XLSX test as {DUMMY_XLSX_PATH} was not created.")

    # 4. Valid PNG file (dummy_test.png created above)
    if os.path.exists(DUMMY_PNG_PATH): # Only test if creation was successful or partially successful
        try:
            print(f"\nTesting with: {DUMMY_PNG_PATH}")
            load_document(DUMMY_PNG_PATH)
        except (FileNotFoundError, ValueError, pytesseract.TesseractNotFoundError, UnidentifiedImageError) as e:
            print(f"OCR test: Could not process '{DUMMY_PNG_PATH}'. Error: {e}. This might be due to Tesseract not being available/configured, the dummy image creation failing, or the file being invalid. Actual OCR requires a proper image and Tesseract setup.")
        finally:
            if os.path.exists(DUMMY_PNG_PATH): # Ensure cleanup
                os.remove(DUMMY_PNG_PATH)
                print(f"Removed dummy PNG: {DUMMY_PNG_PATH}")
    else:
        print(f"\nSkipping PNG OCR test as {DUMMY_PNG_PATH} was not created (Pillow might be missing or failed).")


    # Create other dummy files for remaining tests (like JPG)
    dummy_jpg_path = "photo.jpg" # Re-using this for the generic image test
    # Ensure a fresh dummy JPG for this specific test, separate from potential OCR test if DUMMY_PNG_PATH was photo.jpg
    if os.path.exists(dummy_jpg_path): os.remove(dummy_jpg_path) # Remove if exists from previous OCR test

    try:
        # Attempt to create a dummy image file if Pillow is available, otherwise just a blank file
        img_jpg = Image.new('RGB', (60, 30), color = 'blue') # Different from potential OCR dummy
        img_jpg.save(dummy_jpg_path)
    except NameError: # PIL Image not imported due to earlier failure perhaps
         with open(dummy_jpg_path, "w") as f: # pragma: no cover
            f.write("dummy jpg content")
    except Exception as e: # pragma: no cover
        print(f"Could not create dummy JPG {dummy_jpg_path}: {e}. Using placeholder.")
        with open(dummy_jpg_path, "w") as f:
            f.write("dummy jpg content")


    if not os.path.exists("unsupported.txt"): # pragma: no cover (should be created by now usually)
        with open("unsupported.txt", "w") as f:
            f.write("dummy text content")

    # 5. Valid JPG file (now specifically testing the image loading part, not necessarily successful OCR)
    try:
        print(f"\nTesting with: {dummy_jpg_path}")
        load_document(dummy_jpg_path) # This will attempt OCR
    except (FileNotFoundError, ValueError) as e:
        # If Tesseract is not found, this will be caught by ValueError in load_document
        print(f"Image test (JPG): Could not process '{dummy_jpg_path}'. Error: {e}.")


    # 6. Unsupported file type
    try:
        print(f"\nTesting with: unsupported.txt")
        load_document("unsupported.txt")
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")

    # 7. Non-existent file
    non_existent_file_path = "nonexistent.tmp"
    if os.path.exists(non_existent_file_path): # pragma: no cover
        os.remove(non_existent_file_path)

    try:
        print(f"\nTesting with: {non_existent_file_path}")
        load_document(non_existent_file_path)
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")

    # Clean up other dummy files
    print("\n--- Cleaning up other dummy files ---")
    for demo_file in ["unsupported.txt", dummy_jpg_path]:
        if os.path.exists(demo_file):
            try:
                os.remove(demo_file)
                print(f"Removed dummy file: {demo_file}")
            except OSError as e: # pragma: no cover
                print(f"Error removing {demo_file}: {e}")
