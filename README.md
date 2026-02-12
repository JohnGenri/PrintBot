# PrintBot 🖨️

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows)
![GUI](https://img.shields.io/badge/GUI-Tkinter-green?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)

**PrintBot** is a robust automation utility designed to streamline your document workflow. It monitors your email inbox in real-time and automatically sends attachments to your specified printers. 🚀

No more manual downloading and opening files! Just forward them to your bot email and let **PrintBot** handle the rest. ✨

---

## 🌟 Features

*   📧 **IMAP Integration**: Connects securely to any standard IMAP email server (SSL/TLS supported).
*   📄 **PDF Automation**: Automatically detects and prints `.pdf` files using **SumatraPDF**.
*   🖼️ **Image Processing**: Supports printing of `.jpg`, `.png`, `.bmp`, and more via **IrfanView** or **MS Paint**.
*   🛡️ **Smart Filtering**:
    *   **Whitelist Mode**: Only print emails from trusted senders to save paper and ink. 🔒
    *   **Open Mode**: Print attachments from any incoming email. 🌍
*   ⚙️ **Flexible Configuration**:
    *   Choose specific printers for images vs. documents.
    *   Set custom paths for external handlers (IrfanView/SumatraPDF).
*   🔄 **Resilience**: Built-in "Keep-Alive" worker that automatically reconnects if the network drops. 🔌
*   📂 **Auto-Archiving**: Automatically moves processed emails to a `Printed` folder to keep your inbox clean. 🧹

---

## 🛠️ Requirements

To use the full potential of PrintBot, ensure you have the following installed on your Windows machine:

1.  **Windows OS** (7, 8, 10, 11) 🪟
2.  **[IrfanView](https://www.irfanview.com/)** (Recommended for image printing) 🎨
3.  **[SumatraPDF](https://www.sumatrapdfreader.org/)** (Required for PDF printing) 📑

---

## 🚀 How to Use

1.  **Launch the App**: Run `gui_print_bot.exe`.
2.  **Configure Email**: Enter your IMAP server details, email address, and password.
3.  **Set Paths**: 
    *   Point to your `SumatraPDF.exe` for PDF handling.
    *   Point to your `i_view64.exe` (IrfanView) for images.
4.  **Select Printers**: Choose which physical printer to use for each file type.
5.  **Start**: Click the **Start** button! The bot will begin monitoring your inbox. 🟢

---

## 📸 Screenshots

| Settings Panel ⚙️ | Log Output 📝 |
|:---:|:---:|
| *Configure your servers and printers easily.* | *Real-time status updates and error tracking.* |

---

## 🤝 Contributing

Feel free to open issues or submit pull requests if you have ideas for improvements! 💡

**Enjoy your automated printing experience!** 🎉
