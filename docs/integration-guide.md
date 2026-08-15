# bedoffice Integration Guide

> Comprehensive developer guide for embedding **bedoffice** (offline client-side ONLYOFFICE WASM suite) into any web application without a Document Server.

---

## Table of Contents

1. [Overview & Architecture](#1-overview--architecture)
2. [Integration Methods](#2-integration-methods)
3. [Configuration Reference (`docConfig`)](#3-configuration-reference-docconfig)
4. [postMessage IPC Protocol Reference](#4-postmessage-ipc-protocol-reference)
5. [Core Features & Code Recipes](#5-core-features--code-recipes)
   - [5.1 Saving File Streams (`ArrayBuffer`)](#51-saving-file-streams-arraybuffer)
   - [5.2 Dynamic Renaming](#52-dynamic-renaming)
   - [5.3 Export / Save As Different Formats](#53-export--save-as-different-formats)
   - [5.4 Read-Only / Preview Mode](#54-read-only--preview-mode)
   - [5.5 Creating Blank Documents](#55-creating-blank-documents)
   - [5.6 Automation Connector API](#56-automation-connector-api)
6. [Supported File Formats](#6-supported-file-formats)
7. [Frequently Asked Questions (FAQ)](#7-frequently-asked-questions-faq)

---

## 1. Overview & Architecture

**bedoffice** operates 100% in the client browser using WebAssembly (`x2t.wasm`). It does **not** connect to any backend Document Server.

### Window Hierarchy

When integrating via iframe, communication flows across three window layers:

```
┌─────────────────────────────────────────────────────────────┐
│  Host Application (Your Web App)                            │
│  - Triggers save / configuration                            │
│  - Receives ArrayBuffer byte streams via postMessage        │
└──────────────────────────────┬──────────────────────────────┘
                               │ postMessage IPC
┌──────────────────────────────▼──────────────────────────────┐
│  onlyoffice.html (Integration Bridge)                       │
│  - Initializes DocsAPI.DocEditor                            │
│  - Hooks byte extraction & suppresses browser auto-download │
└──────────────────────────────┬──────────────────────────────┘
                               │ Internal DocsAPI
┌──────────────────────────────▼──────────────────────────────┐
│  ONLYOFFICE Editor Core (WASM / x2t engine)                 │
└─────────────────────────────────────────────────────────────┘
```

---

## 2. Integration Methods

### Method 1: `<iframe>` + `postMessage` (Recommended)

Embed the bridge page into your app using the hosted CDN URL (or your self-hosted instance):

```html
<!-- Use the hosted bridge or your own self-hosted URL -->
<iframe id="bedoffice-frame" 
        src="https://milio48.github.io/bedoffice/onlyoffice.html" 
        style="width: 100%; height: 100vh; border: none;"></iframe>

<script>
const frame = document.getElementById('bedoffice-frame');

window.addEventListener('message', (event) => {
    const data = event.data;
    if (!data) return;

    // 1. When iframe is ready, send document configuration
    if (data.type === 'onlyoffice-ready') {
        frame.contentWindow.postMessage({
            type: 'onlyoffice-config',
            docConfig: {
                document: {
                    url: 'https://example.com/demo.docx',
                    title: 'Demo Document.docx',
                    fileType: 'docx',
                    key: 'doc_' + Date.now(),
                    permissions: { edit: true, download: true, print: true }
                },
                editorConfig: {
                    mode: 'edit',
                    lang: 'en-US',
                    user: { id: 'user-01', name: 'Alex Developer' }
                }
            }
        }, '*');
    }

    // 2. When document is fully loaded
    if (data.type === 'onlyoffice-document-ready') {
        console.log('Document loaded and ready for editing.');
    }

    // 3. When saved file stream arrives
    if (data.type === 'onlyoffice-saved' && data.ok) {
        const fileBlob = new Blob([data.buffer]);
        console.log(`Received saved file (${fileBlob.size} bytes):`, data.fileName);
        // Upload to your backend or save locally
    }
});
</script>
```

---

## 3. Configuration Reference (`docConfig`)

The `docConfig` follows the standard ONLYOFFICE configuration schema with specialized offline enhancements:

| Field | Type | Description |
| :--- | :--- | :--- |
| `documentType` | `string` | `'word'` \| `'cell'` \| `'slide'` \| `'pdf'` (Auto-detected if omitted) |
| `document.url` | `string` | Remote HTTP URL or `blob:` URL. |
| `document.title` | `string` | File name displayed in the top title bar. |
| `document.fileType` | `string` | File extension (e.g. `'docx'`, `'xlsx'`, `'pptx'`, `'pdf'`). |
| `document.key` | `string` | Unique document identifier (e.g. `docId_timestamp`). |
| `document.permissions`| `object` | `{ edit: true, download: true, print: true }` |
| `editorConfig.mode` | `string` | `'edit'` (editable) or `'view'` (read-only). |
| `editorConfig.lang` | `string` | UI language tag (e.g. `'en-US'`, `'zh-CN'`, `'es'`, `'fr'`). |
| `editorConfig.user` | `object` | `{ id: 'uuid', name: 'User Name' }` displayed in comments and revisions. |
| `editorConfig.customization`| `object` | Customize toolbar, feedback buttons, etc. |

---

## 4. postMessage IPC Protocol Reference

### Messages Received from `onlyoffice.html`

| Event Type | Payload Attributes | Description |
| :--- | :--- | :--- |
| `onlyoffice-ready` | none | Sent when the iframe has loaded and is ready for `onlyoffice-config`. |
| `onlyoffice-document-ready` | none | Sent when ONLYOFFICE has parsed the file and rendered the editor canvas. |
| `onlyoffice-saved` | `ok: boolean`, `buffer: ArrayBuffer`, `fileName: string`, `fileType: string`, `requestId?: string`, `error?: string` | Contains the resulting raw binary stream after saving or export. |
| `onlyoffice-open-error` | `error: string` | Triggered if loading the document fails. |
| `onlyoffice-dirty` | none | User has made unsaved modifications in the document. |
| `onlyoffice-clean` | none | Document is currently clean (all changes saved or undone). |

### Messages Sent to `onlyoffice.html`

| Command Type | Payload Parameters | Description |
| :--- | :--- | :--- |
| `onlyoffice-config` | `docConfig: object`, `streamFallback?: 'autosave' \| 'download'` | Injects configuration and mounts the editor. |
| `onlyoffice-save` | `requestId?: string`, `format?: string` | Triggers WASM conversion and requests the file stream. |
| `onlyoffice-set-name` | `name: string` | Updates the active document title in the top bar. |
| `onlyoffice-save-as` | `name: string`, `fileType: string` | Saves a copy of the document under a new name and format. |
| `onlyoffice-connector` | `action: string`, `data?: any` | Interacts with ONLYOFFICE Automation API. |

---

## 5. Core Features & Code Recipes

### 5.1 Saving File Streams (`ArrayBuffer`)

Triggering save with a unique `requestId` allows asynchronous promise matching in host applications:

```javascript
function saveDocument(frameWindow, requestId = 'req_' + Date.now()) {
    return new Promise((resolve, reject) => {
        const handler = (event) => {
            const d = event.data;
            if (d && d.type === 'onlyoffice-saved' && d.requestId === requestId) {
                window.removeEventListener('message', handler);
                if (d.ok && d.buffer) {
                    resolve(new Blob([d.buffer]));
                } else {
                    reject(new Error(d.error || 'Save failed'));
                }
            }
        };
        window.addEventListener('message', handler);
        frameWindow.postMessage({ type: 'onlyoffice-save', requestId }, '*');
    });
}
```

### 5.2 Dynamic Renaming

Update the title displayed in the editor header:

```javascript
iframe.contentWindow.postMessage({
    type: 'onlyoffice-set-name',
    name: 'Quarterly_Report_2026.docx'
}, '*');
```

### 5.3 Export / Save As Different Formats

Convert and extract the document into other formats (e.g. DOCX -> PDF):

```javascript
iframe.contentWindow.postMessage({
    type: 'onlyoffice-save',
    requestId: 'export_pdf_1',
    format: 'pdf'
}, '*');
```

---

## 6. Framework Integration Examples

### Vue 3 Component

```vue
<template>
  <div class="bedoffice-container">
    <!-- Use hosted bridge or your self-hosted /onlyoffice.html -->
    <iframe ref="iframeRef" src="https://milio48.github.io/bedoffice/onlyoffice.html" class="editor-frame" />
  </div>
</template>

<script setup>
import { ref, onMounted, onUnmounted } from 'vue'

const props = defineProps({
  documentUrl: { type: String, required: true },
  fileName: { type: String, required: true },
  fileType: { type: String, default: 'docx' }
})

const emit = defineEmits(['ready', 'saved', 'error'])
const iframeRef = ref(null)

function onMessage(e) {
  const data = e.data
  if (!data) return

  if (data.type === 'onlyoffice-ready') {
    iframeRef.value.contentWindow.postMessage({
      type: 'onlyoffice-config',
      docConfig: {
        document: {
          url: props.documentUrl,
          title: props.fileName,
          fileType: props.fileType
        },
        editorConfig: { lang: 'en-US', mode: 'edit' }
      }
    }, '*')
  } else if (data.type === 'onlyoffice-document-ready') {
    emit('ready')
  } else if (data.type === 'onlyoffice-saved' && data.ok) {
    emit('saved', new Blob([data.buffer]), data.fileName)
  }
}

onMounted(() => window.addEventListener('message', onMessage))
onUnmounted(() => window.removeEventListener('message', onMessage))
</script>

<style scoped>
.bedoffice-container, .editor-frame {
  width: 100%;
  height: 100vh;
  border: none;
}
</style>
```

### React / Next.js Component

```tsx
import React, { useEffect, useRef, useCallback } from 'react';

interface BedofficeProps {
  documentUrl: string;
  fileName: string;
  fileType?: string;
  onSave?: (blob: Blob, fileName: string) => void;
}

export const BedofficeEditor: React.FC<BedofficeProps> = ({
  documentUrl,
  fileName,
  fileType = 'docx',
  onSave
}) => {
  const iframeRef = useRef<HTMLIFrameElement>(null);

  const handleMessage = useCallback((event: MessageEvent) => {
    const data = event.data;
    if (!data) return;

    if (data.type === 'onlyoffice-ready' && iframeRef.current?.contentWindow) {
      iframeRef.current.contentWindow.postMessage({
        type: 'onlyoffice-config',
        docConfig: {
          document: { url: documentUrl, title: fileName, fileType },
          editorConfig: { lang: 'en-US', mode: 'edit' }
        }
      }, '*');
    }

    if (data.type === 'onlyoffice-saved' && data.ok && onSave) {
      onSave(new Blob([data.buffer]), data.fileName);
    }
  }, [documentUrl, fileName, fileType, onSave]);

  useEffect(() => {
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, [handleMessage]);

  return (
    <iframe
      ref={iframeRef}
      src="https://milio48.github.io/bedoffice/onlyoffice.html"
      style={{ width: '100%', height: '100vh', border: 'none' }}
      title="ONLYOFFICE x bedoffice Document Editor"
    />
  );
};
```

---

## 7. Supported File Formats

| Category | Primary Formats | Additional Supported Formats |
| :--- | :--- | :--- |
| **Word Processor** | `.docx` | `.doc`, `.odt`, `.rtf`, `.txt`, `.html` |
| **Spreadsheet** | `.xlsx` | `.xls`, `.ods`, `.csv` |
| **Presentation** | `.pptx` | `.ppt`, `.odp`, `.ppsx` |
| **PDF & Forms** | `.pdf` | `.oxps`, `.xps`, `.djvu` (Annotation & Form Filling) |

---

## 8. Frequently Asked Questions (FAQ)

#### Q: Is any data sent to external servers?
**A:** No. All document parsing, rendering, editing, and conversion occurs entirely on the client side using WebAssembly.

#### Q: How does caching work?
**A:** Static CDNs (like GitHub Pages or Cloudflare) deliver assets with edge compression and HTTP caching. Once loaded, cached WebAssembly binaries load instantaneously on repeat visits.
