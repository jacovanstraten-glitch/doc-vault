# Document Vault

A personal document archive that runs in the browser on phone and PC and stores everything in your own OneDrive.

**Open the app:** https://jacovanstraten-glitch.github.io/document-vault/

## What it does

- **Scan** letters, invoices, receipts and warranties with the phone camera (crop, rotate, enhance, multi-page). Scans are saved as PDF.
- **Add files** on the PC by dropping PDFs or photos on the window, or by saving them into `OneDrive\Document Vault\Inbox`. The app picks those up automatically.
- **Search everything**, including the text inside documents. PDFs with a text layer are read directly; photos and scans are read on your own device with OCR (Tesseract, Dutch + English).
- **File** documents under Private or Business and a category. The file is renamed and moved in OneDrive to match, for example `Document Vault\Private\Purchases & Warranties\2025-03-14 Piet Fiets - Trek FX 3 bike.pdf` (date, sender, title).
- **Learns from you:** a new document from a sender you filed before is filed the same way automatically.
- **AI only as a fallback (optional):** if OCR is unsure (faded receipt, handwriting), the page can be sent to Claude with your own API key. Without a key everything else works.

## Security

- The app is a static web page: there is no server of ours. Your files go directly between your device and your OneDrive.
- All code, including libraries, OCR engine, language data and fonts, is served from this repository. A Content-Security-Policy blocks scripts from anywhere else.
- Sign-in is Microsoft's own (MSAL). The app never sees your password. Tokens are stored in the browser on your device.
- Access can be revoked at any time at https://account.live.com/consent/Manage.
- Protect the two accounts this depends on with two-step verification: your Microsoft account (your documents) and your GitHub account (the app code).

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

## Open on the PC

Double-click `pc/Document Vault.bat`. It opens the app in its own Chrome window (Edge if Chrome is not installed). Put a copy on your desktop or wherever you keep your other app launchers.

## Install on the phone

Open the link in Chrome on Android → menu → **Add to Home screen** (or **Install app**).

## Changing the AI provider

All AI code is in one function, `aiRead()` in `app.js`. To use another provider, replace that function; the rest of the app does not depend on it.

## Libraries (bundled in `vendor/`)

MSAL.js 3.28.1 (Microsoft sign-in), PDF.js 3.11.174, jsPDF 2.5.1, Tesseract.js 5.1.1 with its OCR engine and Dutch + English language data, and the fonts. Everything is served from this repository; the app loads no code from other websites.
