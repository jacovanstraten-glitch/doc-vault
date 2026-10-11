# Document Vault

A personal document archive that runs in the browser on phone and PC and stores everything in your own OneDrive.

**Open the app:** https://jacovanstraten-glitch.github.io/document-vault/

## What it does

- **Scan** letters, invoices, receipts and warranties with the phone camera (crop, rotate, enhance, multi-page). Scans are saved as PDF.
- **Add files** on the PC by dropping PDFs or photos on the window, or by saving them into `OneDrive\Document Vault\Inbox`. The app picks those up automatically.
- **Search everything**, including the text inside documents. PDFs with a text layer are read directly; photos and scans are read on your own device with OCR (Tesseract, Dutch + English).
- **File** documents under Private or Business and a category. The file is renamed and moved in OneDrive to match, for example `Document Vault\Private\Purchases & Warranties\2025-03-14 Trek FX 3 bike, Piet Fiets.jpg`.
- **Learns from you:** a new document from a sender you filed before is filed the same way automatically.
- **AI only as a fallback (optional):** if OCR is unsure (faded receipt, handwriting), the page can be sent to Claude with your own API key. Without a key everything else works.

## Where your data lives

| What | Where |
| --- | --- |
| Documents | `OneDrive\Document Vault\Private\…`, `…\Business\…`, `…\Inbox\` |
| Index with all details and searchable text | `OneDrive\Document Vault\index.json` |
| Overview for Excel | `OneDrive\Document Vault\index.csv` |
| AI key (optional) | Only in the browser of the device where you enter it |

The code in this repository contains no personal data and no secrets. The client ID in `config.js` is a public identifier.

## Setup (already done)

1. Microsoft Entra app registration **Document Vault** — personal Microsoft accounts only, platform *Single-page application*, redirect URI `https://jacovanstraten-glitch.github.io/document-vault/`.
2. Its Application (client) ID is in `config.js`.
3. GitHub Pages serves the `main` branch root.

On first sign-in Microsoft asks once to allow access to your OneDrive files. The app only reads and writes inside the `Document Vault` folder.

## Install on the phone

Open the link in Chrome on Android → menu → **Add to Home screen** (or **Install app**).

## Changing the AI provider

All AI code is in one function, `aiRead()` in `app.js`. To use another provider, replace that function; the rest of the app does not depend on it.

## Libraries (bundled in `vendor/`)

MSAL.js 3.28.1 (Microsoft sign-in), PDF.js 3.11.174, jsPDF 2.5.1, Tesseract.js 5.1.1 (the OCR engine and language data load from jsDelivr on first use and are then cached).
