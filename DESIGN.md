# Solution Design: Financial Document Processing System

## 1. Overview

This document outlines the design for a comprehensive solution to process and extract information from various financial documents, including balance sheets, GST invoices, tax invoices, and bank statements. The system supports multiple file formats (PDF, images, Excel, Word), includes PII detection, and allows for configurable PII handling. The design prioritizes modularity and aims for offline processing capabilities, with file inputs potentially sourced from local storage or Google Cloud Storage.

## 2. Core Components

The system is designed as a pipeline of core components:

### 2.1. Input Module
*   **Responsibilities:**
    *   Accept input files in various formats: PDF (text-based and image-based), images (JPEG, PNG, etc.), Microsoft Excel (XLSX), and Microsoft Word (DOCX).
    *   Handle file paths from local storage or references to files stored in Google Cloud Storage.
    *   Perform initial validation for file accessibility and basic format checks.
*   **Technology Suggestions:** Standard Python file handling libraries. For GCS, `google-cloud-storage` library.
*   **Considerations:** Robust error handling for unsupported file types or corrupted files.

### 2.2. Text Extraction Module
*   **Responsibilities:**
    *   For image-based documents (scanned PDFs, images): Utilize Optical Character Recognition (OCR) to convert images of text into machine-readable text.
    *   For text-based PDFs: Directly extract embedded text streams.
    *   For Excel files: Parse and extract text content from relevant cells.
    *   For Word files: Extract text content from the document structure.
*   **Technology Suggestions:**
    *   OCR: Tesseract OCR (via `pytesseract` Python wrapper). Google Cloud Vision API or AWS Textract could be considered for higher accuracy if cloud connectivity is permissible for this step.
    *   PDF (text-based): `PyMuPDF` (fitz) or `pdfminer.six`.
    *   Excel: `openpyxl` or `pandas`.
    *   Word: `python-docx`.
*   **Considerations:**
    *   Image preprocessing (deskewing, noise removal, binarization) to improve OCR accuracy.
    *   Handling various layouts, including multi-column text and tables.
    *   Extracting text in the correct reading order.

### 2.3. Document Type Classifier
*   **Responsibilities:**
    *   Analyze the extracted raw text (and potentially file metadata) to automatically identify the type of financial document (e.g., Balance Sheet, GST Invoice, Bank Statement, General Tax Invoice).
*   **Technology Suggestions:**
    *   Primarily keyword-based classification: Searching for unique and common phrases/terms specific to each document type (e.g., "Balance Sheet As At", "GSTIN", "Invoice No.", "Account Statement", "Transactions").
    *   Rule-based logic combining keyword presence and structural cues.
    *   (Future Enhancement) Machine learning models (e.g., Naive Bayes, SVM, or a simple neural network) trained on a labeled dataset of document examples for more complex scenarios or a wider variety of documents.
*   **Considerations:** Potential ambiguity if documents share common keywords. The classifier should ideally output a classification label and a confidence score.

### 2.4. Information Extraction Module
*   **Responsibilities:** This module will contain specialized sub-modules for each supported document type to extract relevant data fields.
    *   **Balance Sheet Extractor:**
        *   Fields: Company Name, Statement Date, Currency, Line items for Assets (e.g., Cash, Accounts Receivable, Inventory, Fixed Assets with their values), Liabilities (e.g., Accounts Payable, Loans, Current Liabilities, Long-term Liabilities with their values), and Shareholder Equity (e.g., Share Capital, Retained Earnings with their values).
        *   Logic: Utilize regex for dates and monetary values. Identify sections based on keywords and common accounting structures. Parse table-like structures.
    *   **GST/Tax Invoice Extractor:**
        *   Fields: Supplier Details (Name, Address, GSTIN), Recipient Details (Name, Address, GSTIN), Invoice Number, Invoice Date, Date of Supply, Place of Supply, Line Items (Description, HSN/SAC Code, Quantity, Unit, Rate, Taxable Value, Discount), Tax Details (CGST, SGST/UTGST, IGST rates and amounts), Total Invoice Value, E-way Bill Number (if present), IRN (if e-invoicing).
        *   Logic: Regex for GSTINs, invoice numbers, dates, HSN/SAC codes, and amounts. Table extraction for line items. Keyword spotting for field identification.
    *   **Bank Statement Extractor:**
        *   Fields: Account Holder Name, Account Number, Bank Name & Branch, Statement Period, Transaction Date, Value Date, Description/Narrative, Cheque Number (if any), Withdrawal Amount (Debit), Deposit Amount (Credit), Running Balance.
        *   Logic: Parse header for account details. Extract tabular transaction data. Regex for dates, amounts, and potentially patterns within transaction narratives (e.g., "ATM WDL", "NEFT", "UPI").
*   **Technology Suggestions:**
    *   Regular Expressions (Python `re` module).
    *   NLP libraries like spaCy or NLTK for Named Entity Recognition (names, organizations, dates, monetary values) and tokenization.
    *   Table extraction libraries for PDFs (e.g., `camelot-py`, `tabula-py`) or custom logic for text-based tables.
    *   Template-based approaches for common invoice/statement layouts from major banks/companies can significantly improve accuracy for those specific templates.
*   **Considerations:** High variability in document layouts is the primary challenge. Extractors need to be robust or adaptable, potentially using a combination of rules, templates, and ML for future enhancements.

### 2.5. PII Detection Module
*   **Responsibilities:**
    *   Scan all extracted text fields to identify potential Personally Identifiable Information (PII).
*   **Technology Suggestions:**
    *   Regular Expressions: For common PII patterns like phone numbers, email addresses, GSTINs, PAN numbers, Aadhaar numbers (if applicable), bank account numbers (common structures), credit/debit card number patterns.
    *   Named Entity Recognition (NER) using libraries like spaCy: To identify person names, organization names, and locations which might be PII depending on context.
    *   Lookup lists: For keywords that often precede PII (e.g., "Name:", "Address:", "A/C No:").
*   **Considerations:** Balancing precision and recall to minimize false positives (flagging non-PII) and false negatives (missing actual PII). Context is important (e.g., "Name" as a column header vs. "Name: John Doe").

### 2.6. PII Handling Module
*   **Responsibilities:**
    *   Apply a user-defined action to the PII identified by the PII Detection Module.
*   **Technology Suggestions:** String manipulation functions in Python for redaction or masking.
*   **Configurable Options (to be confirmed by user):**
    *   **Full Redaction:** Replace identified PII with a placeholder (e.g., `[REDACTED_NAME]`, `[REDACTED_PHONE]`).
    *   **Partial Masking:** Mask parts of the PII (e.g., Account Number: `******1234`, Email: `j***@e***.com`).
    *   **No Action:** Pass the PII through unchanged (if handling is done by a separate, downstream system).
    *   **(Future Enhancement) Encryption:** Encrypt PII, requiring a key for decryption.
*   **Considerations:** The chosen handling method must be applied consistently across all detected PII. The level of masking should be configurable.

### 2.7. Output Module
*   **Responsibilities:**
    *   Structure the extracted information (with PII handled as per configuration) into a well-defined output format.
    *   Provide a clear linkage between the extracted data and the source document.
*   **Technology Suggestions (to be confirmed by user):**
    *   **JSON:** Highly flexible, can represent hierarchical data from all document types. Each document could be a JSON object.
    *   **CSV:** Suitable for flatter data structures. Might require separate CSVs for main document info and line items (e.g., for invoices).
    *   **Database:** Store extracted information in a structured database (e.g., PostgreSQL, MySQL, NoSQL like MongoDB) for querying and analysis (more of a storage/application layer consideration beyond basic extraction output).
*   **Considerations:** A clear, consistent schema for the output for each document type is essential for downstream consumption.

## 3. Data Flow

The general data flow through the system will be as follows:

```
Input File (PDF, Image, XLSX, DOCX)
    |
    v
[Input Module]  ->  (Raw File Data / Path)
    |
    v
[Text Extraction Module]  ->  (Raw Extracted Text)
    |
    v
[Document Type Classifier]  ->  (Document Type + Raw Extracted Text)
    |
    v
[Information Extraction Module (Type-specific)]  ->  (Structured Extracted Data with field labels)
    |
    v
[PII Detection Module]  ->  (Structured Extracted Data + Identified PII locations/types)
    |
    v
[PII Handling Module]  ->  (Processed Structured Data with PII handled as per rules)
    |
    v
[Output Module]  ->  (Final Output, e.g., JSON or CSV)
```

## 4. Offline Processing and Google Cloud Storage

*   **Offline Capability:** All core processing logic (text extraction from local files, classification, information extraction, PII handling) will be designed to run offline without requiring continuous internet access, once necessary libraries are installed.
*   **Google Cloud Storage (GCS) Integration:** The Input Module can be extended to accept file paths from a GCS bucket. This would involve using the `google-cloud-storage` library to download files temporarily for local processing or stream them if feasible. The core processing pipeline remains the same after the file is fetched.
*   **Cloud Services (Optional):** For enhanced accuracy in OCR or NLP tasks (like NER or classification), cloud-based services (e.g., Google Cloud Vision API, Google Cloud Natural Language API) could be integrated. This would be an optional configuration and would require internet connectivity for the specific components utilizing them. The primary design will focus on offline-capable libraries first.

## 5. Modularity and Scalability

*   Each component is designed to be modular, allowing for independent development, testing, and upgrades.
*   For scalability, especially for high-volume processing, individual components (e.g., OCR, Information Extraction) could be scaled independently if deployed in a microservices-like architecture or using parallel processing techniques (e.g., a task queue like Celery with multiple workers). This is a consideration for future deployment rather than the initial library/script development.
