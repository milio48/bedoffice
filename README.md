# ONLYOFFICE × bedoffice

<p align="center">
  <strong>Client-Side ONLYOFFICE WASM Document Suite & bedoffice Embed Bridge</strong><br>
  <em>Zero Server, 100% Privacy, Pure WebAssembly ONLYOFFICE Editor for Word, Excel, PowerPoint & PDF</em>
</p>

<p align="center">
  <a href="LICENSE.txt"><img src="https://img.shields.io/badge/license-AGPL--3.0-blue.svg" alt="License: AGPL-3.0"></a>
  <a href="https://github.com/milio48/bedoffice"><img src="https://img.shields.io/badge/engine-ONLYOFFICE_WASM-purple.svg" alt="ONLYOFFICE WASM Core"></a>
  <a href="docs/integration-guide.md"><img src="https://img.shields.io/badge/docs-integration_guide-success.svg" alt="Docs"></a>
</p>

---

## ⚡ Highlights

- **🔒 100% Client-Side Privacy**: Document editing, conversion, and exporting happen entirely in the browser using `x2t.wasm`. No files or keystrokes are ever sent to an external server.
- **📄 Full Format Support**: Native compatibility with `.docx`, `.xlsx`, `.pptx`, `.pdf`, ODF (`.odt`, `.ods`, `.odp`), `.csv`, and form filling.
- **🔌 Seamless Embed Protocol**: `onlyoffice.html` exposes a lightweight `postMessage` IPC protocol that delivers raw `ArrayBuffer` byte streams directly into your host application.
- **⚡ Static & Serverless**: Ready to deploy on GitHub Pages, Cloudflare Pages, Vercel, S3, or Nginx with zero backend configuration.
- **🚀 Web App & Developer Hub**: Includes an end-user web app (`index.html`) and an interactive developer documentation portal (`docs.html`).

---

## 🚀 Quick Start / Embed in 5 Lines of Code

Embed the editor into any web application using standard `<iframe>` and `postMessage` (using the hosted CDN bridge or your self-hosted URL):

```html
<!-- Use hosted CDN bridge or your own self-hosted onlyoffice.html -->
<iframe id="editor" 
        src="https://milio48.github.io/bedoffice/onlyoffice.html" 
        style="width: 100%; height: 100vh; border: none;"></iframe>

<script>
const iframe = document.getElementById('editor');

window.addEventListener('message', (event) => {
  const data = event.data;
  if (!data) return;

  // 1. Send configuration when iframe is ready
  if (data.type === 'onlyoffice-ready') {
    iframe.contentWindow.postMessage({
      type: 'onlyoffice-config',
      docConfig: {
        document: {
          url: 'https://example.com/document.docx',
          title: 'Document.docx',
          fileType: 'docx'
        },
        editorConfig: { lang: 'en-US', mode: 'edit' }
      }
    }, '*');
  }

  // 2. Receive edited file stream (ArrayBuffer)
  if (data.type === 'onlyoffice-saved' && data.ok) {
    const blob = new Blob([data.buffer]);
    console.log('Saved document bytes:', blob.size);
  }
});
</script>
```

---

## 📚 Documentation

- 📖 **[Integration Guide](docs/integration-guide.md)**: Full API specification, `docConfig` schema, `postMessage` protocol, and React / Vue / Svelte components.
- 🔬 **[File Stream Architecture](docs/file-stream-architecture.md)**: Deep dive into WASM `x2t.downloadFile` interception and zero-backend buffer extraction.

---

## 📁 Repository Structure

```
bedoffice/
├── index.html         # Standalone Web App for End-Users
├── docs.html          # Interactive Developer Documentation & Embed Guide
├── onlyoffice.html    # Core Iframe Integration Bridge (postMessage IPC)
├── assets/            # Config helper (office-config.js), favicon, PDF templates
├── blank/             # Blank templates for new documents (DOCX, XLSX)
├── docs/              # In-depth technical guides (Markdown)
├── LICENSE.txt        # AGPL-3.0 License
└── README.md          # Official documentation
```

---

## 📜 License & Credits

- Licensed under **[AGPL-3.0](LICENSE.txt)**.
- ONLYOFFICE core components and trademarks belong to **[Ascensio System SIA / ONLYOFFICE](https://www.onlyoffice.com/)**.
- Forked & enhanced from [fernfei/OnlyofficePersonal](https://github.com/fernfei/OnlyofficePersonal).
