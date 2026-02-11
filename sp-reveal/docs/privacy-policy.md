# Privacy Policy — SP Reveal
_Last updated: February 2026_

SP Reveal is a browser extension designed to enhance productivity when working with SharePoint Online lists and items. This privacy policy explains what information the extension accesses, how it is used, and how your data is protected.

---

## 1. Data Collection

SP Reveal **does not collect, store, transmit, or share any personal data**.

The extension does **not**:
- collect browsing history  
- collect user identifiers  
- collect login information  
- send data to external servers  
- track usage  
- use analytics or telemetry  

All processing happens **locally** within the user's browser.

---

## 2. Permissions

SP Reveal **does not request any optional or sensitive Chrome permissions**.

The extension operates entirely through:
- a statically declared content script  
- user‑initiated actions (e.g., clicking copy buttons)  
- built‑in browser clipboard APIs  

No additional Chrome APIs or permissions (tabs, storage, scripting, clipboardWrite, etc.) are used or required.

SP Reveal runs only on SharePoint pages due to its `content_scripts.matches` patterns, not because of host permissions.

---

## 3. Data Usage

SP Reveal interacts only with the **current SharePoint page** the user is viewing and only when the user triggers an action in the extension's UI.

Any information read from the page—such as field labels, list metadata, or item details—is:
- accessed **only** in response to a user action  
- used solely to perform the requested feature (e.g., copy a URL, reveal an internal name)  
- processed entirely within the browser  
- never logged, saved, uploaded, or shared  

The extension does **not** modify SharePoint data except when the user explicitly chooses **Duplicate Item**, which uses SharePoint’s own REST APIs and the user’s existing permissions.

---

## 4. Data Storage

SP Reveal does **not** use Chrome storage APIs.

No data—personal, SharePoint-related, or otherwise—is stored locally or remotely.

---

## 5. Third‑Party Sharing

SP Reveal does **not** share any data with third parties.

The extension contains:
- no trackers  
- no analytics  
- no external API calls  
- no advertising modules  

All functionality is self‑contained.

---

## 6. Security

Because the extension runs entirely on the user’s device:
- no external transmission reduces exposure risk  
- no data leaves the browser  
- SP Reveal cannot access anything the user is not already permitted to access within SharePoint  
- the extension cannot interact with or read data from non‑SharePoint websites  

All SharePoint operations respect the user’s existing Microsoft 365 authentication.

---

## 7. Changes to This Policy

Any updates to this privacy policy will be published in the extension store listings and the public GitHub repository.

---

## 8. Contact

For questions or support, please visit:  
**https://about365.nl**
