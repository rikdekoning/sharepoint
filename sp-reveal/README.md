# SP Reveal

**SP Reveal** is a modern productivity extension for **SharePoint Online** that adds developer‑focused tools directly into list and item forms.  
It helps admins, developers, and power users work faster by revealing internal names, generating URLs, duplicating items, and inspecting full item payloads — all directly inside SharePoint.

---

## 🚀 Features

### Item actions
- Show internal column names (click to copy)
- Clear internal names
- Copy item ID
- Copy item URL (correctly normalized)
- Copy item API URL
- Duplicate item (copies simple fields using SharePoint REST API)
- Show all fields (searchable Fluent-style dialog with JSON export)

### List actions
- Copy list GUID  
- Copy list URL  
- Copy list API URL  

---

## 🛠 Installation

### Microsoft Edge Add-ons  
https://microsoftedge.microsoft.com/addons/detail/sp-reveal/mpknkmeflipbbmdepeeijhojamiblfif

### Chrome Web Store  
(Will be added once approved)

### Manual Installation (Developer Mode)

1. Download or clone this repository:

   ```bash
   git clone https://github.com/rikdekoning/sharepoint.git

2. Open your browser and navigate to the extensions page:
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`

3. Enable **Developer mode**.

4. Click **Load unpacked**.

5. Select the `src/` folder from this repository.

---

## 🎥 Screenshots

docs/screenshots/screenshot1.png  
docs/screenshots/screenshot2.png  
docs/screenshots/screenshot3.png

---

## 🔍 How It Works

SP Reveal injects a lightweight content script **only on SharePoint domains** (`*.sharepoint.com`, `*.sharepoint.us`, etc.).  
All processing occurs **entirely in the browser** and uses the user’s existing SharePoint permissions.

The extension:
- Reads SharePoint form structure and metadata
- Extracts list paths, list IDs, and item IDs
- Detects modern dialogs, classic forms, and list grid selections
- Retrieves field definitions and item data via SharePoint REST (`/_api/web/lists/...`)
- Never sends data to external servers

---

## 🔒 Privacy & Security

SP Reveal:
- Does **not** collect or transmit data  
- Uses permissions only to operate on the active SharePoint tab  
- Stores only tiny local settings  
- Never communicates with external services  

See: `docs/privacy-policy.md`

---

## 🔑 Permissions

| Permission       | Reason |
|------------------|--------|
| `activeTab`      | Communicate with the active SharePoint page |
| `tabs`           | Required by Chrome MV3 to get active tab ID |
| `scripting`      | Inject the content script |
| `storage`        | Store small local preferences |
| `clipboardWrite` | Copy IDs, URLs, API paths, internal names |
| Host permissions | Access SharePoint Online pages |

---

## 🧑‍💻 Development

1. Modify files inside the `src/` directory.  
2. Reload the extension via the browser’s extension page.  
3. For publishing, zip the **contents of `src/`** and upload to Chrome/Edge stores.

---

## ❗ Known Limitations

- Classic SharePoint pages may behave differently  
- SharePoint permission restrictions still apply  
- Some cached pages may require a refresh  

---

## 🤝 Contributing

Issues and pull requests are welcome.  
This project is designed for transparency, simplicity, and enterprise use.

---

## 📄 License

SP Reveal is licensed under the **MIT License**.  
See `LICENSE` for details.

---

## 🙌 Author

Created by **Rik de Koning**  
More at: https://about365.nl

