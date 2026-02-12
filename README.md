# 🖨️ PrintBot (Beta)

An automated bot for printing email attachments.
The script monitors an IMAP inbox, downloads attachments (PDF, Images), and sends them to specified printers.

## ✨ Features
- **Persistent Connection:** Keeps a constant connection to the server (no connection spam).
- **Auto-Recovery:** Automatically reconnects in case of network interruptions.
- **Smart Routing:** PDFs are printed via SumatraPDF, images via IrfanView or MS Paint.
- **Filtering:** Sender whitelist support.

## 🚀 Installation and Usage (from source)

1. **Install Python 3.10+**.
2. **Install dependencies**:
   ```bash
   pip install imap-tools pywin32
   ```
3. **Install IrfanView** (for images) and **SumatraPDF** (for PDFs).
4. **Run the script**. On the first run, it will create `settings.ini`.
5. **Configure `settings.ini`** (specify server, login, password, and printer paths).

## 📦 Build to EXE

To build into a single executable file using PyInstaller:

```bash
python -m PyInstaller --onefile --noconsole gui_print_bot.py
```

## ⚠️ Important
Requires printer usage rights and internet access (IMAP port 993).

---
**Status:** Beta v5.0
