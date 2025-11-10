A lightweight rule-based NLP add-in for Microsoft Word that detects contracting parties in legal documents and highlights all textual references, including aliases and pronoun forms—fully offline and privacy-safe.


## ✨ Features

- **Automatic Party Detection**
  - Detects parties from the opening recital (e.g., *“This Agreement is made between X and Y…”*)
  - Also scans early lines for organization names (`Inc`, `LLC`, `Ltd`, `Holdings`, etc.)

- **Interactive Party Selection**
  - Displays detected parties as **radio buttons**
  - Option to **manually add** a party if needed

- **Smart Highlighting**
  - Highlights:
    - Exact name matches (case-insensitive)
    - Possessive forms (`Acme` → `Acme’s`)
    - Role aliases (e.g., *Company*, *Contractor*, *Client*, etc.)
    - Pronouns (e.g., *it / its / itself* or *they / their / themselves* based on plurality)

- **One-Click Reset**
  - Removes only the highlights inserted by the add-in

- **Document-Safe**
  - Works **body-only** (no headers/footers)
  - UTF-8 & emoji safe
  - Idempotent (can run repeatedly without stacking highlights)

---

## 🧠 AI / NLP Usage

This project does **not** use any AI or external NLP models.

All logic was implemented using:

- **Heuristic text pattern recognition**
- **Regex + capitalized entity scanning**
- **Alias and pronoun mapping rules**
- **Range-based Word highlighting using Office.js**

This ensures:
- ✅ 100% offline functionality  
- ✅ No document data leaves the user’s system  
- ✅ Works securely in confidential legal environments  

---

## 🖥 Tech Stack

- **Office JavaScript API (Office.js)**
- **TypeScript**
- **Webpack 5 & webpack-dev-server**
- **VS Code**

---

## 🚀 Getting Started (Development Mode)

```bash
npm install
npm start
