# Doc Merger: Batch Generator 🚀

An elegant, client-side web application designed to parse rows from a Google Sheet and dynamically generate individual, perfectly styled Google Documents based on a master template. 

By isolating each row into its own file, this tool completely bypasses Google Docs REST API formatting limitations—guaranteeing **100% preservation** of your template's layout, custom fonts, tables, cell borders, alignment, and background colors.

---

## ✨ Features

*   **Pixel-Perfect Formatting:** Zero style degradation. Tables, borders, fills, and margins remain exactly as designed in your template.
*   **Automatic Organization:** Every execution dynamically creates a uniquely timestamped destination folder in your Google Drive to hold the generated files.
*   **Real-time Progress Tracker:** Watch the application securely initialize the OAuth handshake, read the sheet arrays, and process each row sequentially.
*   **Built with Modern Tech:** Styled with Tailwind CSS, built with vanilla JavaScript, and powered securely by the official Google Identity Services (GSI) Client Library.

---

## 📋 How It Works

1.  **Template Injection:** The script uses your original Google Doc template as a structural baseline.
2.  **Isolated Cloning:** For every row in your spreadsheet, the tool makes an exact clone of the template directly inside a new batch output folder.
3.  **In-Place Text Swapping:** It uses the cloud-native `replaceAllText` method to instantly swap markers (e.g., `<<Student_Name>>`) with row data without touching the table cells' structural styling properties.

---
