# AutoCrat Custom Hub: Single-Doc Generator

A lightweight, open-source document merging web application hosted entirely on GitHub Pages. It connects seamlessly to a Google Apps Script backend API to read row data from a Google Sheet and compile it into a single, beautifully formatted master Google Doc with dynamic page breaks.

Unlike standard tools that generate hundreds of individual files, this tool consolidates everything into **one master document**, keeping your Google Drive organized.

---

## 🚀 Features

* **Consolidated Output:** Merges multiple data rows into a single Google Doc file instead of cluttering your Drive.
* **Dynamic Template Layouts:** Preserves your original text styles, tables, lists, and formatting.
* **Serverless Frontend:** Hosted 100% free on GitHub Pages.
* **Simple Mapping Engine:** Uses the classic `<<HeaderName>>` tag syntax to automatically detect where data should go.

---

## 📂 Project Architecture

```text
  [ GitHub Pages Frontend ] 
            │  (Secure JSON POST Request)
            ▼
  [ Google Apps Script Web App ]
            │  (Accesses via Link Sharing)
            ├──► Reads [ Google Sheet Data ]
            └──► Clones & Appends [ Google Doc Template ] ──► Saves Master Doc File
