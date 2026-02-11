# SP Reveal — Store Listing Content (Chrome + Edge)

This file contains the recommended long description, short description, promotional text, screenshots, and test instructions for publishing SP Reveal to the Chrome Web Store and Microsoft Edge Add-ons Store.

---

## 📝 Short Description (Chrome & Edge)

A SharePoint productivity toolkit to reveal internal names, copy URLs, duplicate items, and inspect all item fields.

(Under 132 characters — compliant with both stores)

---

## 📝 Long Description (Chrome & Edge)

SP Reveal is a modern productivity extension for SharePoint Online that adds powerful tools directly into list and item forms. It is designed for SharePoint administrators, developers, and power users who need fast access to internal field names, item metadata, and API-ready URLs.

With SP Reveal, you can instantly inspect list items, duplicate existing items, extract URLs or API paths, and analyze all item fields in a searchable Fluent-style dialog — all without leaving the page.

### 🔍 Key Features

#### **Item Tools**
- Show internal (logical) column names — clickable to copy instantly  
- Clear internal names  
- Copy Item ID  
- Copy Item URL (correctly normalized)  
- Copy Item API URL (`/_api/...`)  
- Duplicate Item via SharePoint REST API  
- Show All Fields — searchable, with JSON export  

#### **List Tools**
- Copy list GUID  
- Copy list URL  
- Copy list API endpoint  

### ✨ Additional Highlights
- Works on modern and classic SharePoint forms  
- Detects selected items in list grid views  
- Fast and lightweight — no performance impact  
- Supports light and dark mode  
- Fully client-side — does not send or store data  

### 🔒 Privacy & Security

SP Reveal:
- Does **not** collect or transmit data  
- Stores no SharePoint content  
- Processes only the information visible on the page  
- Runs entirely in the browser  
- Never communicates with external servers  

See the full privacy policy at: `docs/privacy-policy.md`

### 🧑‍💻 Ideal For

- SharePoint administrators  
- SharePoint developers  
- Power users  
- Anyone troubleshooting list or item metadata  

---

## 📸 Screenshot Captions (with working URLs)

1. **"Quick tools for SharePoint items — copy IDs, URLs, and API endpoints instantly."**  
   https://raw.githubusercontent.com/rikdekoning/sharepoint/main/docs/screenshots/screenshot1.png

2. **"Reveal internal (logical) column names directly on the form."**  
   https://raw.githubusercontent.com/rikdekoning/sharepoint/main/docs/screenshots/screenshot2.png

3. **"Inspect all item fields with search and JSON export."**  
   https://raw.githubusercontent.com/rikdekoning/sharepoint/main/docs/screenshots/screenshot3.png

These URLs always display correctly in:
- Chrome Web Store  
- Edge Add-ons  
- GitHub  
- Any Markdown viewer  

---

## 🏷️ Promotional Text (Optional)

SP Reveal adds developer-friendly tools to SharePoint list forms: reveal internal names, copy URLs and IDs, duplicate items, and inspect full field data — all inside the browser.

---

## 🔧 Test Instructions (for store review teams)

SP Reveal requires no separate login.

To test:

1. Sign in to any Microsoft 365 tenant with SharePoint access.  
2. Open a SharePoint Online list.  
3. Select or open an item.  
4. Open the extension popup.  
5. Validate actions such as:  
   - Show Internal Names  
   - Copy Item URL  
   - Copy Item ID  
   - Duplicate Item  
   - Show All Fields  

All features use the user’s existing SharePoint permissions and operate entirely within the browser.

---

## ✔ Edge Add-ons Store Link

https://microsoftedge.microsoft.com/addons/detail/sp-reveal/mpknkmeflipbbmdepeeijhojamiblfif

---

This file is ready to use for both Chrome Web Store and Microsoft Edge Add-ons submission.
