<div align="center">

# 📌 OBS Pro

**An Intelligent, High-Performance OBS Task & Database Management Platform**

[![GitHub Pages](https://img.shields.io/badge/Live-Demo-brightgreen?style=for-the-badge&logo=github)](https://mesuraj.github.io/ObsPro/)
[![GitHub Repository](https://img.shields.io/badge/GitHub-Repository-blue?style=for-the-badge&logo=github)](https://github.com/MeSuraj/ObsPro)
[![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web/HTML)
[![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-38B2AC?style=for-the-badge&logo=tailwind-css&logoColor=white)](https://tailwindcss.com/)
[![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)

---

### 🚀 [Launch Live Application](https://mesuraj.github.io/ObsPro/) · [Report Bug](https://github.com/MeSuraj/ObsPro/issues) · [Request Feature](https://github.com/MeSuraj/ObsPro/issues)

</div>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Key Features](#-key-features)
- [Tech Stack](#-tech-stack)
- [Project Architecture](#-project-architecture)
- [Directory Structure](#-directory-structure)
- [Getting Started](#-getting-started)
  - [Prerequisites](#prerequisites)
  - [Local Setup](#local-setup)
- [Detailed User Guide](#-detailed-user-guide)
  - [1. Document Parsing (Word to JSON)](#1-document-parsing-word-to-json)
  - [2. JSON Merging Engine](#2-json-merging-engine)
  - [3. Excel Matrix & Coordinate Lookup](#3-excel-matrix--coordinate-lookup)
  - [4. Non-Conformance (NC) Management](#4-non-conformance-nc-management)
  - [5. Satellite Vector & KML Integration](#5-satellite-vector--kml-integration)
- [Data Security & Privacy](#-data-security--privacy)
- [Contributing](#-contributing)
- [License](#-license)
- [Contact & Support](#-contact--support)

---

## 📖 Overview

**OBS Pro** is an end-to-end web application designed to streamline and automate daily operational workflows for OBS task validation, spatial data matching, report compilation, and database indexing.

In field operations and quality management, handling unstructured task reports in `.docx` format, cross-referencing geographical coordinates (Latitude/Longitude), and manually tracking Non-Conformance (NC) records can be time-consuming and error-prone. **OBS Pro** solves these challenges by providing a fast, client-side tool that converts documents to structured JSON, links directly with Excel databases, dynamically tracks NC issues, and visualizes KML satellite vectors.

---

## 🔥 Key Features

### 📄 1. Word to JSON Engine
- Converts Microsoft Word (`.docx`) reports directly into formatted, queryable JSON schemas.
- Powered by `Mammoth.js` for fast client-side extraction without server reliance.
- Preserves report headers, task IDs, image references, and custom metadata fields.

### 🗂️ 2. JSON Consolidation & Merge Tool
- Allows users to combine multiple batch-parsed JSON report files into a unified master database.
- Deduplicates repetitive records and normalizes task attributes automatically.

### 📊 3. Excel Matrix Integration & Spatial Lookup
- Loads spreadsheet data (`.xlsx`, `.xls`, `.csv`) via `SheetJS`.
- Performs instant cross-referencing between Job/Task IDs and geographical coordinates (Latitude & Longitude).
- Updates dataset status fields in real-time.

### 🔍 4. Deep Global Indexing & Search
- Provides real-time search across active workstreams and historical backup repositories.
- Allows filtering by Task ID, Job Reference, Date Range, and Non-Conformance status.

### 📌 5. Automated KML & Vector Linker
- Dynamically resolves job references to Satellite Vector KML files.
- Provides immediate one-click downloading and map visualization options.

### ⚙️ 6. Customizable Non-Conformance (NC) Manager
- Built-in NC manager for adding, modifying, or resetting custom quality audit remarks.
- Enables standardized tagging across task entries for consistent team reporting.

### 🌗 7. Adaptive UI & Export Options
- Fully responsive layout built with Tailwind CSS.
- Supports both Light Mode and Dark Mode.
- One-click export for localized single-page HTML/PDF views or updated `.xlsx` database backups.

---

## 🛠️ Tech Stack

| Domain | Technology | Description |
| :--- | :--- | :--- |
| **Frontend Framework** | HTML5 / JavaScript (ES6+) | Core logic executed client-side |
| **Styling** | Tailwind CSS | Modern utility-first CSS styling |
| **Document Processing** | Mammoth.js | Converts `.docx` documents to HTML/JSON |
| **Spreadsheet Engine** | SheetJS (`xlsx`) | Parses and modifies Excel files in-browser |
| **Icons** | Lucide Icons / Heroicons | Vector interface icon sets |
| **Hosting & CI/CD** | GitHub Pages | Zero-downtime static hosting |

---

## 📐 Project Architecture

```text
               +-------------------------------------------------------+
               |                  User Web Browser                     |
               +-------------------------------------------------------+
                                           |
      +---------------------+--------------+--------------+---------------------+
      |                     |                             |                     |
[ DOCX Files ]        [ JSON Files ]              [ Excel Matrix ]       [ KML Query ]
      |                     |                             |                     |
      v                     v                             v                     v
+------------+     +-----------------+          +------------------+   +----------------+
| Mammoth.js |     | JSON Merge Engine|          | SheetJS (XLSX)   |   | Auto Vector    |
| Parser     |     | & Normalizer    |          | Spatial Matcher  |   | KML Resolver   |
+------------+     +-----------------+          +------------------+   +----------------+
      |                     |                             |                     |
      +---------------------+--------------+--------------+---------------------+
                                           |
                                           v
                        +-------------------------------------+
                        |     Centralized State & NC Indexer  |
                        +-------------------------------------+
                                           |
                       +-------------------+-------------------+
                       |                                       |
                       v                                       v
             [ Interactive UI View ]                 [ Export Options ]
             (Search, Dark Mode, Render)             (JSON, Excel, PDF)
```

---

## 📁 Directory Structure

```text
ObsPro/
├── index.html            # Primary application UI layout & template
├── app.js                # Main application logic, parsing engine, state management
├── styles/
│   └── styles.css        # Theme variables, custom styling & animation overrides
├── assets/               # Static icons, favicons, and branding assets
├── LICENSE               # Project license (MIT)
└── README.md             # Project documentation
```

---

## ⚡ Getting Started

### Prerequisites

All processing takes place in the user's web browser. No server environments (Node.js, Python, PHP) or database backends (MySQL, MongoDB) are strictly required to run the UI.

- A modern web browser (Google Chrome 90+, Mozilla Firefox 88+, Microsoft Edge 90+, Safari 14+).
- Git installed on your local machine (for development/cloning).

### Local Setup

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/MeSuraj/ObsPro.git
   ```

2. **Navigate to the Directory:**
   ```bash
   cd ObsPro
   ```

3. **Launch the Application:**
   - Double-click `index.html` to open it in your default web browser.
   - Alternatively, serve it using VS Code's **Live Server** extension or Python's built-in HTTP server:
     ```bash
     python -m http.server 8000
     ```
     Then open `http://localhost:8000` in your browser.

---

## 📖 Detailed User Guide

### 1. Document Parsing (Word to JSON)
1. Open **OBS Pro** and navigate to the **Converter** panel.
2. Drag and drop your `.docx` OBS report file or select it via the file browser.
3. Click **Parse Document**. The system extracts task IDs, descriptions, and media links, returning a clean JSON structure ready for download or merging.

### 2. JSON Merging Engine
1. Navigate to the **JSON Engine** tab.
2. Select multiple JSON files generated from previous steps.
3. Click **Merge Files** to compile them into a unified dataset.

### 3. Excel Matrix & Coordinate Lookup
1. Load your master Excel database (`.xlsx`) containing job references.
2. The system maps Task IDs against rows to retrieve Latitude, Longitude, and current status.
3. Any edits made in the interface are synced back into the memory matrix and can be downloaded as an updated Excel sheet.

### 4. Non-Conformance (NC) Management
1. Highlight tasks flagged for review.
2. Open the **NC Manager** to assign standardized remarks, severity tags, and corrective action notes.
3. Custom NC categories can be added or restored to default at any time via settings.

### 5. Satellite Vector & KML Integration
1. Enter a valid Job ID into the spatial query field.
2. The system scans the repository index and builds a direct download link for associated KML vector files for spatial validation.

---

## 🔒 Data Security & Privacy

- **Client-Side Processing:** All conversions (Word parsing, JSON operations, Excel updates) are performed strictly inside your browser environment using Web APIs.
- **Zero Server Uploads:** Your document data, coordinates, and internal reports are **never** uploaded to external servers or third-party tracking services.

---

## 🤝 Contributing

Contributions are welcome! To contribute:

1. **Fork** the Repository.
2. **Create a Feature Branch:**
   ```bash
   git checkout -b feature/AmazingFeature
   ```
3. **Commit your Changes:**
   ```bash
   git commit -m "Add some AmazingFeature"
   ```
4. **Push to the Branch:**
   ```bash
   git push origin feature/AmazingFeature
   ```
5. **Open a Pull Request** for review.

---

## 📜 License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

## 👨‍💻 Contact & Support

**Project Author:** [MeSuraj](https://github.com/MeSuraj)  
**Live Site:** [https://mesuraj.github.io/ObsPro/](https://mesuraj.github.io/ObsPro/)  
**GitHub Repository:** [https://github.com/MeSuraj/ObsPro](https://github.com/MeSuraj/ObsPro)

If you find this tool helpful, please give it a **⭐️ Star** on GitHub!