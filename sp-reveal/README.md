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

2. Navigate to the SP Reveal extension folder:

   ```bash
   cd sharepoint/sp-reveal

3. Open your browser and navigate to the extensions page:
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`

4. Enable **Developer mode**.

5. Click **Load unpacked**.

6. Select the `src/` folder inside the `sp-reveal` directory.

---

## 🎥 Screenshots

[![Screenshot 1](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot1.png)](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot1.png)

[![Screenshot 2](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot2.png)](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot2.png)

[![Screenshot 3](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot3.png)](https://raw.githubusercontent.com/rikdekoning/sharepoint/main/sp-reveal/docs/screenshots/screenshot3.png)

---

## 🔍 How It Works

SP Reveal injects a lightweight content script **only on SharePoint domains**  
(`*.sharepoint.com`, `*.sharepoint.us`, `*.sharepoint.de`, etc.).

All processing occurs **entirely in the browser** using the user’s existing SharePoint permissions.

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
- Stores no SharePoint or personal content  
- Processes only what is visible on the current SharePoint page  
- Never communicates with external servers  
- Uses **no privileged Chrome permissions**

See full privacy policy: `docs/privacy-policy.md`

---

## 🧑‍💻 Development

1. Modify files inside the `src/` directory.  
2. Reload the extension via Chrome/Edge extension settings.  
3. For publishing, zip the **contents of `src/`** (not the folder itself) and upload to the Chrome or Edge store dashboards.

---

## ❗ Known Limitations

- Classic SharePoint pages may behave differently  
- SharePoint permission restrictions still apply  
- Some cached SharePoint pages may require a refresh  

---

## 🤝 Contributing

Issues and pull requests are welcome.  
This project is designed for transparency, simplicity, and enterprise usability.

---

## 📄 License

SP Reveal is licensed under the **MIT License**.  
See `LICENSE` for details.

---

## 🙌 Author

Created by **Rik de Koning**  
More at: https://about365.nl
