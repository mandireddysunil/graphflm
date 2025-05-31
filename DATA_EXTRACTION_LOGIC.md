# Data Extraction Logic for Financial Documents

## 1. Introduction

This document details the specific data fields to be extracted from each supported financial document type (Balance Sheets, GST/Tax Invoices, Bank Statements) and outlines the general logic and techniques to be employed for this extraction. The strategies aim to handle common patterns and structures, while acknowledging that real-world documents exhibit considerable variation.

## 2. Balance Sheet Extraction

*   **Objective:** To extract key financial figures, dates, and company identification from balance sheets.
*   **Key Fields to Extract:**
    *   **Company Name:** The name of the entity for which the balance sheet is prepared.
    *   **Statement Date / "As of" Date:** The specific date to which the financial position pertains.
    *   **Currency:** The currency in which monetary values are reported (e.g., INR, USD).
    *   **Assets Section:**
        *   Total Current Assets
        *   Total Non-Current (or Long-Term) Assets
        *   Total Assets
        *   *Optional Sub-Categories (if clearly identifiable):* Cash & Cash Equivalents, Accounts Receivable, Inventory, Investments, Property, Plant & Equipment (PP&E), Intangible Assets.
    *   **Liabilities Section:**
        *   Total Current Liabilities
        *   Total Non-Current (or Long-Term) Liabilities
        *   Total Liabilities
        *   *Optional Sub-Categories (if clearly identifiable):* Accounts Payable, Short-Term Borrowings, Long-Term Debt, Provisions.
    *   **Equity Section (Shareholders' Equity / Net Worth):**
        *   Total Equity
        *   *Optional Sub-Categories (if clearly identifiable):* Share Capital, Retained Earnings, Other Equity.
*   **Extraction Logic & Techniques:**
    *   **Company Name:** Typically found at the top of the document. Look for text with larger font sizes or text adjacent to keywords like "Limited", "Ltd.", "Inc.", "Corporation". Named Entity Recognition (NER) for `ORG` (Organization) can be applied.
    *   **Statement Date:** Use regular expressions (regex) to find date patterns (e.g., `dd-mm-yyyy`, `dd/mm/yy`, `Month dd, yyyy`, `dd MMMM yyyy`). Often preceded by phrases like "As of", "As at", "For the period ending".
    *   **Currency:** Search for currency symbols (e.g., ₹, $, €) or ISO codes (e.g., INR, USD) typically found near monetary values. If not explicit, a default (e.g., INR) might be assumed based on context or user setting.
    *   **Section Identification:** Use keywords and phrases to identify main sections: "ASSETS", "LIABILITIES", "EQUITY", "SHAREHOLDERS' EQUITY", "Current Assets", "Non-Current Assets", "Fixed Assets", "Current Liabilities", etc.
    *   **Line Item and Value Extraction:**
        *   Identify line items by matching text against a predefined list of common balance sheet account names.
        *   Extract associated monetary values, which are typically numeric and may appear in adjacent columns or on the same line. Use regex for capturing monetary formats (e.g., `\d{1,3}(,\d{3})*(\.\d{1,2})?`).
        *   Handle negative values, often represented in parentheses (e.g., `(1,234.56)`).
        *   Utilize table structure analysis if items are in clear tabular format. For less structured layouts, proximity logic (label-value pairing) will be used.
        *   Verify extracted totals (e.g., "Total Assets" should reconcile with the sum of its components if individual items are extracted; "Total Assets" should equal "Total Liabilities + Total Equity").

## 3. GST Invoice / General Tax Invoice Extraction

*   **Objective:** To extract supplier and recipient details, invoice identification, line-item details of goods/services, tax information, and total amounts.
*   **Key Fields to Extract:**
    *   **Supplier Details:** Name, Full Address, GSTIN, Phone Number, Email (if available).
    *   **Recipient (Buyer) Details:** Name, Billing Address, Shipping Address (if different), GSTIN (if registered), Phone Number (if available).
    *   **Invoice Identifiers:** Invoice Number (unique serial), Invoice Date, Date of Supply (if different), Place of Supply (State/UT code or name).
    *   **Line Item Details (repeated for each item/service):**
        *   Description of Goods or Services
        *   HSN (Harmonized System of Nomenclature) Code / SAC (Services Accounting Code)
        *   Quantity (Qty)
        *   Unit of Measurement (UoM / UQC - e.g., pcs, kg, ltr, hrs, bags)
        *   Rate (Price per Unit)
        *   Taxable Value (Amount before tax)
        *   Discount (if provided per item)
    *   **Tax Details (can be per item and/or summarized):**
        *   CGST (Central GST): Rate (%) and Amount
        *   SGST (State GST) / UTGST (Union Territory GST): Rate (%) and Amount
        *   IGST (Integrated GST): Rate (%) and Amount (for inter-state supplies)
        *   Cess (if applicable): Rate (%) and Amount
    *   **Total Values:**
        *   Total Taxable Value (sum of taxable values of all line items)
        *   Total CGST, SGST/UTGST, IGST, Cess amounts
        *   Grand Invoice Total (Total Taxable Value + Total Taxes)
        *   Invoice Total in Words
    *   **Other Potential Fields:** E-way Bill Number, IRN (Invoice Reference Number for e-invoicing), Purchase Order Number, Terms of Payment, Supplier's Bank Details (Account Number, IFSC).
*   **Extraction Logic & Techniques:**
    *   **GSTIN/PAN:** Use specific regex patterns for validation (e.g., GSTIN: `^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$`).
    *   **Invoice Number & Dates:** Regex for common alphanumeric and date formats, often located near labels like "Invoice No.", "Inv No", "Date:", "Dated:".
    *   **Addresses:** Search for keywords like "Address:", "Billed To:", "Shipped To:", "Place of Supply:", followed by capturing multi-line text. NER can assist in identifying location entities.
    *   **Line Items:** These are typically presented in a table.
        *   Identify table boundaries and headers (e.g., "Description", "HSN", "Qty", "Rate", "Amount", "CGST", "SGST", "IGST").
        *   Parse each row to extract individual item details.
    *   **HSN/SAC Codes:** Regex for numeric (HSN often 4, 6, or 8 digits) or alphanumeric codes, usually under "HSN" or "SAC" columns/labels.
    *   **Tax Amounts & Rates:** Look for keywords "CGST", "SGST", "IGST", "Cess" often associated with percentage symbols (%) and monetary values. Validate calculations where possible (e.g., Taxable Value * Rate = Tax Amount).
    *   **Totals:** Find labels like "Total Amount", "Grand Total", "Amount Payable". Extract corresponding monetary values using regex. For "Amount in Words", capture text following labels like "Total in Words:".
    *   **Bank Details/IRN/E-Way Bill:** Regex for common account number formats, IFSC codes, and specified formats for IRN/E-Way Bill.

## 4. Bank Statement Extraction

*   **Objective:** To extract account holder and bank information, statement period, and a chronological list of transactions.
*   **Key Fields to Extract:**
    *   **Account Holder Details:** Full Name, Registered Address, Customer ID (CIF No.).
    *   **Bank Details:** Bank Name, Branch Name (and Address if available), IFSC Code, MICR Code.
    *   **Account Details:** Account Number, Account Type (e.g., Savings Account, Current Account, Overdraft).
    *   **Statement Summary:** Statement Period (Start Date and End Date), Opening Balance, Closing Balance, Currency.
    *   **Transaction Details (repeated for each transaction):**
        *   Transaction Date
        *   Value Date (date of posting, if different from transaction date)
        *   Description / Narrative / Particulars (details of the transaction)
        *   Cheque Number (if applicable)
        *   Withdrawal Amount (Debit / DR)
        *   Deposit Amount (Credit / CR)
        *   Running Balance (balance after the transaction)
*   **Extraction Logic & Techniques:**
    *   **Header/Account Information:**
        *   This information is usually located at the top of the statement.
        *   Use regex for Account Numbers (often 10-16 digits), dates (for statement period), IFSC/MICR codes.
        *   Use keyword spotting for labels like "Account Holder:", "Account No.:", "Customer ID:", "Statement Period:", "Opening Balance:", "Closing Balance:".
        *   NER can assist in extracting names and addresses.
    *   **Transaction Table:**
        *   Transactions are almost always presented in a tabular format. Identify table headers (e.g., "Date", "Transaction Details", "Narration", "Debit", "Credit", "Balance", "Chq No.").
        *   Parse each row to extract individual transaction components.
        *   Apply regex for dates in the transaction date column and for monetary amounts in debit, credit, and balance columns.
        *   The "Description/Narrative" column is often free-form; extract its content as is. This field is critical for understanding transaction nature and may contain PII (e.g., names of payees/payers, merchant IDs, UPI transaction details).
    *   **Balances:** Opening and Closing balances are usually clearly labeled or are the first and last entries in the balance column of the transaction list.

## 5. General Extraction Considerations

*   **Layout Analysis:** Before attempting field-specific extraction, an initial analysis to identify blocks of text, tables, headers, and footers can be beneficial.
*   **Template-Free vs. Template-Based:** While the initial aim is for general logic, for very common and consistent formats (e.g., statements from major banks or standard invoice layouts), creating specific templates or rule-sets can significantly improve accuracy.
*   **Confidence Scoring:** Where feasible, the extraction process should associate a confidence score with extracted fields, especially when dealing with ambiguous layouts or relying on probabilistic methods (like NER).
*   **Data Cleaning & Standardization:** Extracted text often requires cleaning (e.g., removing extraneous spaces, special characters) and standardization (e.g., converting all date formats to `YYYY-MM-DD`, standardizing currency symbols).
*   **Iterative Refinement:** The extraction logic will likely require iterative refinement based on testing with a diverse set of real-world documents.
