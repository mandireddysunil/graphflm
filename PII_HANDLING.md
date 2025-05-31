# PII Detection and Handling Strategy

## 1. Introduction

Personally Identifiable Information (PII) refers to any data that can be used to identify a specific individual. The proper handling of PII is crucial for maintaining privacy and complying with data protection regulations. This document outlines the strategy for detecting and handling PII within the financial documents processed by this system.

## 2. Expected PII Types in Financial Documents

The following types of PII are commonly expected or may potentially be found in balance sheets, GST/tax invoices, and bank statements:

*   **General Identifiers:**
    *   **Names:** Names of individuals (e.g., account holders, proprietors, customer names, contact persons, signatories).
    *   **Addresses:** Residential, business, billing, and shipping addresses.
    *   **Phone Numbers:** Personal and business phone numbers.
    *   **Email Addresses:** Personal and business email addresses.

*   **Financial Identifiers:**
    *   **Bank Account Numbers:** Full or partial account numbers, often found in transaction details, supplier payment information, or within bank statements themselves.
    *   **Credit/Debit Card Numbers:** While full card numbers are less common in the primary documents, partial numbers or transaction IDs related to card payments might appear. (The system will be cautious if full PANs are detected).

*   **Government-Issued Identifiers (India specific):**
    *   **GSTIN (Goods and Services Tax Identification Number):** Unique ID for taxpayers under GST. Can identify businesses, including proprietorships which are linked to individuals.
    *   **PAN (Permanent Account Number):** Income tax identification number for individuals and entities.
    *   **Aadhaar Number:** (Highly Sensitive) While not expected to be prevalent in these business documents, any occurrence would be treated as high-risk PII.

*   **Transactional & Contextual PII:**
    *   **Transaction Descriptions/Narratives:** Details in bank statements or payment notes on invoices can contain names of individuals, vendor names (especially if sole proprietors), or other identifying payment purposes.
    *   **Signatures:** If present as images, they are a form of PII. OCR might attempt to convert them or they might be processed as image data.
    *   **Company/Business Names:** While often public, in the context of proprietorships or small partnerships, these can be directly linked to individuals.

## 3. PII Detection Methods

The system will employ a combination of methods for PII detection:

*   **Regular Expressions (Regex):**
    *   This will be the primary method for detecting PII with well-defined patterns, such as:
        *   Phone Numbers (various Indian formats)
        *   Email Addresses
        *   GSTINs (format: 2 numbers, 5 letters, 4 numbers, 1 letter, 1 number, 1 letter, 1 number/letter)
        *   PANs (format: 5 letters, 4 numbers, 1 letter)
        *   Bank Account Numbers (common length and numeric patterns, though these vary widely)
        *   IFSC Codes (format: 4 letters, 0, 6 numbers/letters)
    *   Regex patterns will be designed to be as accurate as possible, but may require refinement based on real-world data.

*   **Named Entity Recognition (NER):**
    *   NLP libraries such as `spaCy` or `NLTK` will be used to identify entities like:
        *   `PERSON`: For individual names.
        *   `ORG`: For organization names (which can be PII for small businesses/proprietorships).
        *   `GPE` (Geopolitical Entity) / `LOC` (Location): To help identify parts of addresses.
    *   NER helps capture PII that doesn't follow strict alphanumeric patterns.

*   **Keyword-Based Detection:**
    *   Identifying common labels or keywords that typically precede PII can help confirm or locate PII. Examples: "Name:", "Address:", "Tel:", "Email:", "A/C No:", "GSTIN:", "PAN:", "Proprietor:".
    *   This method can improve the precision of regex and NER outputs.

*   **Luhn Algorithm (for card numbers):**
    *   If patterns resembling credit/debit card numbers are detected, the Luhn algorithm can be used as a checksum to validate their potential authenticity. However, extraction of full card numbers will be approached with extreme caution and might be out of scope for default processing.

## 4. PII Handling Strategies (User Configurable)

The system will provide the user with options to define how detected PII should be handled before the data is outputted. The user must specify their preferred method.

*   **1. No Action (Default if not specified):**
    *   PII is identified and extracted but no changes are made to the data.
    *   This is suitable if the user has a separate, secure downstream process for PII management.

*   **2. Full Redaction:**
    *   The identified PII string is completely replaced with a generic placeholder indicating the type of PII.
    *   Examples: `[REDACTED_NAME]`, `[REDACTED_PHONE]`, `[REDACTED_ADDRESS]`, `[REDACTED_GSTIN]`, `[REDACTED_ACCOUNT_NUMBER]`.

*   **3. Partial Masking:**
    *   Part of the PII string is replaced with a masking character (e.g., 'X' or '*'), while some parts remain visible for contextual utility.
    *   Examples:
        *   Phone Number: `+91-XXXXXX****` or `******1234`
        *   Email Address: `u***@e***.com` or `user*****@example.com`
        *   Bank Account Number: `XXXXXXXXXX1234` or `****5678`
        *   GSTIN: `XXAAAAA0000A1ZX` (masking first few characters)
        *   PAN: `AAAAX0000X` (masking some characters)
    *   The specific rules for masking (e.g., how many characters to show/hide) will use sensible defaults but could be made more configurable in future iterations.

*   **4. Tagging/Flagging (Future Consideration):**
    *   Instead of altering the data, add metadata to the output indicating which fields contain detected PII and the type of PII.

*   **5. Encryption (Future Consideration):**
    *   Encrypt specific PII fields. This is a more complex approach requiring robust key management and is not part of the initial core offering.

## 5. Process Flow for PII Handling

1.  Data is extracted by the Information Extraction Module.
2.  The extracted text fields are then passed to the PII Detection Module.
3.  The PII Detection Module applies regex, NER, and keyword spotting to identify potential PII.
4.  Detected PII (along with its type, e.g., "PHONE", "EMAIL", "GSTIN") is then processed by the PII Handling Module.
5.  The PII Handling Module applies the user-configured strategy (e.g., redact, mask) to the PII.
6.  The processed data (with PII handled) is then sent to the Output Module.

## 6. Security and Compliance Considerations

*   **User Responsibility:** While this system provides tools for PII detection and handling, the end-user is ultimately responsible for overall compliance with applicable data privacy laws and regulations (e.g., India's Digital Personal Data Protection Act, GDPR if applicable).
*   **Secure Environment:** It is assumed that the system will be deployed and operated in a secure environment, especially when dealing with sensitive financial documents.
*   **Data at Rest and in Transit:** If files are fetched from GCS or if extracted data is stored, appropriate security measures for data at rest (e.g., GCS bucket security, disk encryption) and in transit (e.g., HTTPS) should be implemented by the user.
*   **Access Controls:** Strict access controls should be applied to the documents, the system itself, and any extracted data containing PII.
*   **Minimization:** The principle of data minimization should be applied; only extract and retain PII that is strictly necessary for the defined purpose.
