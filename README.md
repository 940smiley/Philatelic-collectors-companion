# Philatelic Collector's Companion

A Google Apps Script add-on for managing stamp collections in Google Sheets. It helps philatelists log stamps, pull in metadata, and export listings.

## Features
- Reverse image lookup to suggest country, year, and theme using Google Cloud Vision.
- Import catalog data from Colnect or StampWorld by pasting a URL.
- Export selected stamps to an eBay listing template.
- Sidebar to configure required API keys.

## Setup
1. Open your spreadsheet and launch **Extensions > Apps Script**.
2. Add the files from this repository to the Apps Script project.
3. Ensure the manifest includes the listed OAuth scopes and enable the Drive advanced service.
4. In the sheet, run **Configure API Keys** and supply a Google Cloud Vision API key.

## Usage
Use the "Stamp Logging Assistant" sheet to enter stamp details. Tools are available from the custom Philatelic menu to import metadata, fetch image suggestions, and export rows to eBay.

