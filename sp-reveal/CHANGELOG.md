# SP Reveal — Changelog

All notable changes to this project will be documented in this file.

This project adheres to semantic versioning.

---

## **1.1.0 — February 2026**

### ✨ New Features
- **Clickable internal field names**  
  Logical (internal) SharePoint column names added to forms are now interactive.  
  Clicking a tag copies the internal name to the clipboard, with visual toast confirmation.

- **Enhanced “Show All Fields” dialog**  
  - Fluent-style modal layout with light/dark mode support  
  - Search and highlight function  
  - Copy full JSON response  
  - Reopens the originating item panel after closing  
  - Faster rendering and better large‑item handling

### 🛠 Improvements
- **Fixed duplicate site URL issue**  
  Item URLs now normalize correctly using origin-only concatenation, preventing repeated `/sites/.../sites/...` paths.

- **More resilient SharePoint context detection**  
  Improved handling of:
  - Modern dialogs and edit panels  
  - Classic forms  
  - Grid-selected items  
  - Nested folder paths  
  - Lists lacking `_spPageContextInfo`  
  - Fallbacks via `GetListUsingPath`, `GetList`, folder resolution, file resolution, list enumeration, and GUID probing

- **Improved UX for tag rendering**  
  - No duplicate internal-name tags  
  - Keyboard-accessible (Enter/Space)  
  - Clear outline feedback  
  - Pointer cursor applied even without stylesheet loading

- **Safer dialog reopen flow**  
  - More reliable close detection for SharePoint modern item panels  
  - Better fallback handling with overlays, ESC‑key dispatch, and history cleanup

### 🔧 Manifest & Store Preparation
- Updated description for Chrome/Edge 132‑char limit compliance  
- Added `"tabs"` permission for Chrome stability  
- Cleaned SharePoint host permissions  
- Polished extension metadata for store certification  
- Added privacy policy, permission justifications, and certification notes

---

## **1.0.0 — Initial Release**

### 🚀 Core Features
- **Show internal names** — Displays logical field names next to SharePoint form labels  
- **Copy item data**  
  - Item ID  
  - Item URL  
  - Item API URL  
- **Duplicate item** — Copies simple field values using SharePoint REST API  
- **Show all fields** — Opens a detailed modal view with every field/value for the current item

### 📄 List-Level Features
- Copy list GUID  
- Copy list URL  
- Copy list API URL  

### 🧩 Context Detection
- Detects list ID, title, form mode, and item ID from:
  - URL  
  - Dialog panels  
  - Classic pages  
  - Grid selections  

### 🎨 Extension UI
- Popup interface with item and list actions  
- Toast notifications  
- Initial styling for injected UI elements

---

