import re

# Regex patterns
# Basic Indian phone number regex (10 digits, optional +91, optional space/hyphen)
# Covers +91 XXXXXXXXXX, 0XXXXXXXXXX, XXXXXXXXXX, +91-XXXXXXXXXX etc.
PHONE_REGEX = re.compile(r'(?:\+91[\s-]?)?(?:0)?[6-9]\d{9}')

# Standard email regex
EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

# GSTIN regex (India)
# Removed anchors ^ and $ to allow findall to match anywhere in the string.
GSTIN_REGEX = re.compile(r'[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}')

# PAN number regex (India)
# Removed anchors ^ and $ to allow findall to match anywhere in the string.
PAN_REGEX = re.compile(r'[A-Z]{5}[0-9]{4}[A-Z]{1}')

# Basic Bank Account Number Regex (heuristic, typically 9-18 digits)
# This is a general pattern and might need refinement for specific banks or to reduce false positives.
ACCOUNT_NO_REGEX = re.compile(r'\b\d{9,18}\b')


def find_phone_numbers(text):
    """Finds phone numbers in the given text."""
    return PHONE_REGEX.findall(text)

def find_emails(text):
    """Finds email addresses in the given text."""
    return EMAIL_REGEX.findall(text)

def find_gstins(text):
    """Finds GSTINs in the given text.
    Note: This regex checks format. Further validation (e.g., checksum) is not included.
    It expects GSTINs to be standalone words/tokens. May need adjustment if embedded.
    """
    return GSTIN_REGEX.findall(text)

def find_pan_numbers(text):
    """Finds PAN numbers in the given text.
    Note: This regex checks format. Further validation is not included.
    It expects PANs to be standalone words/tokens. May need adjustment if embedded.
    """
    return PAN_REGEX.findall(text)

def find_account_numbers(text):
    """Finds potential bank account numbers in the given text.
    This is heuristic and might identify other long numbers as account numbers.
    """
    return ACCOUNT_NO_REGEX.findall(text)

if __name__ == '__main__':
    sample_text_invoice = """
    Invoice To: John Doe, Ph: +91-9876543210, email: john.doe@example.com.
    His PAN is ABCDE1234F and GSTIN is 29ABCDE1234F1Z5. Another number 08023456789.
    Payment to account 123456789012. Supplier GST: 27AAAAA0000A1Z4.
    For support, contact support@example.co.in or call 8887776665.
    Order ref: ORD12345, Account: 001234567890000.
    """

    sample_text_statement = """
    Account Statement for AC No. 500001234567890.
    Beneficiary: Jane K Smith (PAN: FGHIJ5678K), email jane.s@example.net
    Mobile: 7001112222. Transaction with 22BBBBB1111B1Z0.
    Paid to account 987654321.
    """

    print("--- Testing with Invoice Sample ---")
    # --- Post-processing to refine results ---
    phones_invoice = find_phone_numbers(sample_text_invoice)
    emails_invoice = find_emails(sample_text_invoice)
    gstins_invoice = find_gstins(sample_text_invoice)
    pans_invoice_raw = find_pan_numbers(sample_text_invoice)
    accounts_invoice_raw = find_account_numbers(sample_text_invoice)

    # Refine PANs
    gstin_pan_fragments_invoice = [g[2:12] for g in gstins_invoice]
    temp_pans_invoice = list(pans_invoice_raw)
    pans_invoice_refined = []
    for g_pan_frag in gstin_pan_fragments_invoice:
        if g_pan_frag in temp_pans_invoice:
            temp_pans_invoice.remove(g_pan_frag)
    pans_invoice_refined = temp_pans_invoice

    # Refine Account Nums: Compare account numbers with the last 10 digits of phone numbers
    phone_numeric_parts_invoice = {re.sub(r'\D', '', p)[-10:] for p in phones_invoice if re.sub(r'\D', '', p)} # Ensure non-empty after strip
    accounts_invoice_refined = [acc for acc in accounts_invoice_raw if acc not in phone_numeric_parts_invoice]

    print(f"Phone Numbers: {phones_invoice}")
    print(f"Email Addresses: {emails_invoice}")
    print(f"GSTINs: {gstins_invoice}")
    print(f"PAN Numbers (raw): {pans_invoice_raw}")
    print(f"PAN Numbers (refined): {pans_invoice_refined}")
    print(f"Account Numbers (raw): {accounts_invoice_raw}")
    print(f"Account Numbers (refined): {accounts_invoice_refined}")
    print("\n")

    print("--- Testing with Statement Sample ---")
    phones_statement = find_phone_numbers(sample_text_statement)
    emails_statement = find_emails(sample_text_statement)
    gstins_statement = find_gstins(sample_text_statement)
    pans_statement_raw = find_pan_numbers(sample_text_statement)
    accounts_statement_raw = find_account_numbers(sample_text_statement)

    gstin_pan_fragments_statement = [g[2:12] for g in gstins_statement]
    temp_pans_statement = list(pans_statement_raw)
    for g_pan_frag in gstin_pan_fragments_statement:
        if g_pan_frag in temp_pans_statement:
            temp_pans_statement.remove(g_pan_frag)
    pans_statement_refined = temp_pans_statement

    phone_numeric_parts_statement = {re.sub(r'\D', '', p)[-10:] for p in phones_statement if re.sub(r'\D', '', p)}
    accounts_statement_refined = [acc for acc in accounts_statement_raw if acc not in phone_numeric_parts_statement]

    print(f"Phone Numbers: {phones_statement}")
    print(f"Email Addresses: {emails_statement}")
    print(f"GSTINs: {gstins_statement}")
    print(f"PAN Numbers (raw): {pans_statement_raw}")
    print(f"PAN Numbers (refined): {pans_statement_refined}")
    print(f"Account Numbers (raw): {accounts_statement_raw}")
    print(f"Account Numbers (refined): {accounts_statement_refined}")
    print("\n")

    # Example of how to use with a block of text
    print("--- Testing with General Text Block ---")
    text_block = "Contact: 9998887776, mail_me@test.com, ID: BCDEF2345G, GST: 01ABCDE2345F1Z6, Acc: 1122334455667788"
    all_phones = find_phone_numbers(text_block)
    all_emails = find_emails(text_block)
    all_gstins = find_gstins(text_block)
    all_pans_raw = find_pan_numbers(text_block)
    all_accounts_raw = find_account_numbers(text_block)

    gstin_pan_fragments_general = [g[2:12] for g in all_gstins]
    temp_pans_general = list(all_pans_raw)
    for g_pan_frag in gstin_pan_fragments_general:
        if g_pan_frag in temp_pans_general:
            temp_pans_general.remove(g_pan_frag)
    all_pans_refined = temp_pans_general

    phone_numeric_parts_general = {re.sub(r'\D', '', p)[-10:] for p in all_phones if re.sub(r'\D', '', p)}
    all_accounts_refined = [acc for acc in all_accounts_raw if acc not in phone_numeric_parts_general]

    print(f"Phones: {all_phones}")
    print(f"Emails: {all_emails}")
    print(f"GSTINs: {all_gstins}")
    print(f"PANs (raw): {all_pans_raw}")
    print(f"PANs (refined): {all_pans_refined}")
    print(f"Account Nums (raw): {all_accounts_raw}")
    print(f"Account Nums (refined): {all_accounts_refined}")
    print("\n")

    # Test cases for GSTIN and PAN where they might be part of other strings
    # With findall and removed anchors, these should now be found.
    print("--- Testing edge cases for GSTIN/PAN (now using findall) ---")
    edge_case_text = "MyGSTINis29ABCDE1234F1Z5andPANisABCDE1234Fthanks"
    gstins_edge = find_gstins(edge_case_text)
    pans_edge_raw = find_pan_numbers(edge_case_text)

    gstin_pan_fragments_edge = [g[2:12] for g in gstins_edge]
    temp_pans_edge = list(pans_edge_raw)
    for g_pan_frag in gstin_pan_fragments_edge:
        if g_pan_frag in temp_pans_edge:
            temp_pans_edge.remove(g_pan_frag)
    pans_edge_refined = temp_pans_edge

    print(f"GSTINs in edge case: {gstins_edge}")
    print(f"PANs in edge case (raw): {pans_edge_raw}")
    print(f"PANs in edge case (refined): {pans_edge_refined}")
    print("\n")

    edge_case_text_separated = "My GSTIN is 29ABCDE1234F1Z5 and PAN is ABCDE1234F thanks"
    gstins_edge_sep = find_gstins(edge_case_text_separated)
    pans_edge_sep_raw = find_pan_numbers(edge_case_text_separated)

    gstin_pan_fragments_edge_sep = [g[2:12] for g in gstins_edge_sep]
    temp_pans_edge_sep = list(pans_edge_sep_raw)
    for g_pan_frag in gstin_pan_fragments_edge_sep:
        if g_pan_frag in temp_pans_edge_sep:
            temp_pans_edge_sep.remove(g_pan_frag)
    pans_edge_sep_refined = temp_pans_edge_sep

    print(f"GSTINs in separated edge case: {gstins_edge_sep}")
    print(f"PANs in separated edge case (raw): {pans_edge_sep_raw}")
    print(f"PANs in separated edge case (refined): {pans_edge_sep_refined}")
