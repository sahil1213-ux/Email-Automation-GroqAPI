# Gmail Freight Parser

A Google Apps Script project that automates the extraction of freight forwarding data from emails, processes both text and image-based rate tables, and logs the data into Google Sheets.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup Instructions](#setup-instructions)
- [Usage](#usage)
- [File Structure](#file-structure)
- [Dependencies](#dependencies)
- [How It Works](#how-it-works)
- [Error Handling](#error-handling)
- [Contributing](#contributing)
- [License](#license)

## Overview

The **Gmail Freight Parser for Discord** is designed to streamline the processing of freight forwarding emails by extracting structured data from email bodies and image attachments (e.g., rate tables). It uses the Groq API for natural language processing and OCR.space for image-to-text conversion, then organizes the extracted data into Google Sheets for further use, such as Discord notifications or reporting.

This project is tailored for freight forwarding companies, specifically handling formats like those from P.D.T Logistics, which include multi-route rate tables with a single origin port and multiple destinations.

## Features

- **Email Processing**: Automatically scans Gmail for unread emails labeled "FreightParser".
- **Data Extraction**: Extracts freight forwarding details (e.g., origin, destination, rates, validity) using the Groq API.
- **Image Processing**: Supports OCR for image attachments (e.g., rate tables) using the OCR.space API, with fallback to Google Drive OCR.
- **Multi-Route Support**: Handles emails with multiple routes (e.g., one origin to multiple destinations) and single-route formats.
- **Google Sheets Integration**: Logs extracted data into two sheets: "Rate Updates" (single routes) and "Multi Route Rates" (multi-route tables).
- **Custom Menu**: Adds a Google Sheets menu for manual triggering of email processing.
- **Error Handling**: Robust retry logic for OCR failures and fallback mechanisms to ensure data continuity.
- **P.D.T Logistics Format**: Optimized for parsing weekly rate tables from P.D.T Logistics.

## Prerequisites

- **Google Account**: For access to Gmail, Google Sheets, and Google Apps Script.
- **Google Sheets**: A spreadsheet with the ID specified in the script (e.g., `1ZGzg1M8oQT5AFT7Vkpq5fL9vbay-IyQHNui8eG5cRX4`).
- **API Keys**:
  - **Groq API Key**: For natural language processing (replace `GROK_API` in the script).
  - **OCR.space API Key**: For image-to-text conversion (replace `OCRSPACE_API_KEY` in the script).
- **Gmail Label**: A Gmail label named "FreightParser" to filter emails for processing.

## Setup Instructions

1. **Create a Google Apps Script Project**:

   - Open Google Sheets and navigate to `Extensions > Apps Script`.
   - Copy and paste the contents of `GMAIL FOR DISCORD.txt` into the script editor.
   - Save the project with a descriptive name (e.g., `GmailFreightParser`).

2. **Configure API Keys**:

   - Replace `GROK_API` with your Groq API key.
   - Replace `OCRSPACE_API_KEY` with your OCR.space API key (free tier available at [ocr.space](https://ocr.space)).
   - Replace `SHEET_ID` with your Google Sheets spreadsheet ID.

3. **Set Up Google Sheets**:

   - Create or use an existing Google Sheet with the ID specified in the script.
   - Ensure two sheets are present: `Rate Updates` and `Multi Route Rates`. If not, the script will create the `Multi Route Rates` sheet automatically.
   - Run the `setQuotationHeaders` function to initialize the `Rate Updates` sheet headers.

4. **Set Up Gmail Label**:

   - In Gmail, create a label named `FreightParser`.
   - Apply this label to emails containing freight forwarding data (text or image-based rate tables).

5. **Authorize the Script**:

   - Run any function (e.g., `onOpen`) in the Apps Script editor to trigger authorization.
   - Grant necessary permissions for Gmail, Google Sheets, and Google Drive access.

6. **Deploy Triggers** (Optional):
   - Set up time-driven triggers in Apps Script to run `Rate_Updates` or `Multi_Route_Prices` periodically (e.g., every hour).
   - Alternatively, use the custom menu in Google Sheets to manually trigger processing.

## Usage

1. **Manual Execution**:

   - Open the Google Sheet.
   - Click the `📩 Email Actions` menu.
   - Select `🔄 Refresh Rate Updates` to process single-route emails or `🚚 Refresh Multi Route Prices` for multi-route emails (e.g., P.D.T Logistics format).

2. **Automated Execution**:

   - Set up a time-driven trigger in Apps Script to run `Rate_Updates` or `Multi_Route_Prices` at regular intervals.

3. **Output**:

   - Single-route data is appended to the `Rate Updates` sheet.
   - Multi-route data (e.g., weekly rate tables) is appended to the `Multi Route Rates` sheet.
   - Emails are marked as read after processing to avoid duplicate processing.

4. **Discord Integration** (Not Included):
   - To send data to Discord, add a webhook or bot integration to push data from the Google Sheets to a Discord channel (e.g., using Apps Script's `UrlFetchApp`).

## File Structure

```
GmailFreightParser/
├── Emails                          # All used Emails present here
├── Multi Route Prices.js           # Appscript Code for sheet named Multi Route Rates
├── Price Updates File.xlsx         # SpreadSheet Where Data is Stored.
├── README.markdown                 # This file
├── Rate Update Structure.js        # Appscript Code for UI Elements Of GoogleSheet
├── Rate Prices.js                  # Appscript Code for sheet named Rate Updates
```

## Dependencies

- **Google Apps Script Services**:
  - GmailApp: For accessing and processing emails.
  - SpreadsheetApp: For writing data to Google Sheets.
  - DriveApp: For fallback OCR using Google Drive.
  - UrlFetchApp: For making API calls to Groq and OCR.space.
- **External APIs**:
  - Groq API: For natural language processing and data extraction.
  - OCR.space API: For extracting text from image attachments.

## How It Works

1. **Email Processing**:

   - The script scans Gmail for unread emails labeled `FreightParser`.
   - For each email, it extracts the sender, subject, date, and body.

2. **Image Processing**:

   - If an email contains image attachments (e.g., PNG, JPG), the script uses the OCR.space API to extract text.
   - A fallback mechanism uses Google Drive's OCR if OCR.space fails.

3. **Data Extraction**:

   - The email body and extracted image text are sent to the Groq API with a structured prompt.
   - The prompt is tailored for freight forwarding data, especially P.D.T Logistics' multi-route format.
   - The API returns a JSON object with metadata and rate data.

4. **Data Logging**:

   - Single-route emails are logged to the `Rate Updates` sheet.
   - Multi-route emails (e.g., weekly rate tables) are logged to the `Multi Route Rates` sheet with one row per route.
   - Headers are automatically set, and columns are auto-resized for readability.

5. **Error Handling**:
   - The script includes retry logic for OCR.space API failures and a fallback to Google Drive OCR.
   - If parsing fails, a fallback row with basic email info is logged to avoid data loss.

## Error Handling

- **API Failures**: The script retries OCR.space API calls up to 3 times with a 2-second delay for transient errors (e.g., 503, 429).
- **Invalid JSON**: If the Groq API returns invalid JSON, the script logs an error and adds a fallback row with basic email details.
- **OCR Failures**: If both OCR.space and Google Drive OCR fail, a structured placeholder is returned to allow manual review.
- **Logging**: All errors are logged using `Logger.log` for debugging.

## GoogleSheet View
<img width="1851" height="927" alt="Image" src="https://github.com/user-attachments/assets/9312c596-dceb-4ad4-8f0d-78399ac3027e" />
<img width="1735" height="922" alt="Image" src="https://github.com/user-attachments/assets/ef5db363-6a30-4c39-9fe8-03453f75a01e" />
<img width="598" height="884" alt="Image" src="https://github.com/user-attachments/assets/4cc9cc2b-19ba-4957-aa47-c01e12965ec7" />

## Contributing

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/YourFeature`).
3. Commit your changes (`git commit -m 'Add YourFeature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Open a pull request.

Please ensure your code follows the existing style and includes comments for clarity.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
