# LIFEINNO Daraz Helper — Chrome Extension

A specialized, high-productivity sidebar and popup assistant for **Daraz.pk** listing sellers. Load your product Excel workbook once, search products in a compact card list, and click any field to copy or dynamically injection-fill it directly into Daraz seller forms.

---

## 📂 FILE STRUCTURE IN THIS DIRECTORY

*   `manifest.json` — Chrome extension descriptor (Manifest V3)
*   `popup.html` — Layout for the toolbar popover
*   `popup.js` — Core functional logic for file ingestion, caching, and copying
*   `sidebar.html` — Layout for the tall Chrome Side Panel 
*   `sidebar.js` — Functions for side panel integration
*   `content.js` — Webpage triggers, styling overlays, and automated form fields injection
*   `background.js` — Service worker enabling programmatic side panel openings
*   `styles.css` — Consolidated custom LIFEINNO orange design variables
*   `icon.png` — Visual launcher icon (128x128 pixel block)

---

## 🛠️ HOW TO INSTALL IN CHROME

Follow these 4 simple steps to install this extension directly from this source:

1.  **Download the Extension files**:
    *   Click the **📥 Download Chrome Extension .ZIP** button inside the web app workspace, or copy these files into a local folder on your computer named `lifeinno-daraz-helper`.
2.  **Open Chrome Extensions manager**:
    *   In your Google Chrome url bar, navigate to exactly: `chrome://extensions/`
3.  **Enable Developer Mode**:
    *   In the top-right corner, toggle the **Developer mode** switch to **ON**.
4.  **Load the unpacked folder**:
    *   Click the **Load unpacked** button (top-left of screen).
    *   Select the `lifeinno-daraz-helper` folder containing all these files.
5.  **Success!**:
    *   The **LIFEINNO Daraz Helper** is now loaded! Pin it to your Chrome toolbar for fast access.

---

## 📊 EXCEL LAYOUT REQUIREMENTS

Your workbook sheet needs these standard columns starting on Row 1 (Header row):

| Column | Header | Description |
|---|---|---|
| Col A (0) | `Product Serial Number` | Unique product identifier (only on the first row of each product) |
| Col B (1) | `Professional Title` | Main listing name (only on the first row of each product) |
| Col C (2) | `Size / Dimensions` | Dimension descriptors e.g. "4x4 Inches" |
| Col D (3) | `Color Options` | Available colors e.g. "Gold / Silver" |
| Col E (4) | `Variation Slot Name` | Merged pack sizing e.g. "Pack of 4" (every row) |
| Col F (5) | `Base Unit Price (PKR)` | Per-unit base manufacturer price |
| Col G (6) | `Actual Price (Daraz Cut)` | Non-sale cost including commissions |
| Col H (7) | `Discounted Price (Special Price)` | Active listing sale value (every row) |

*The extension handles row spanning gracefully. First row of each product contains general fields; consecutive rows contain individual variation rows (Columns E, G, H).*

---

## ⚡ KEY CAPABILITIES

*   **Offline First**: Sheets are parsed fully locally in your web browser. No product data ever leaves your computer!
*   **One-Click Copy**: Paste name, sizes, color options, or prices instantly into forms.
*   **Active Field Injection**: Pressing the `⚡ Fill` button next to any field instantly writes it directly into whichever input is active on your Daraz Seller Center.
*   **Auto-Calculated Logistics**: Automatically estimates weight and package measurements matching LIFEINNO's warehouse rule-sets.
