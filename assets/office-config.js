(function (root, factory) {
    if (typeof module === 'object' && module.exports) {
        module.exports = factory()
        return
    }
    root.OfficeConfig = factory()
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict'

    const DEFAULT_LANG = 'en-US'
    const DEFAULT_MODE = 'edit'
    const LOCAL_USER_ID = 'local-user'
    const LOCAL_USER_NAME = 'Local User'
    const EMPTY_UPDATED_AT = 0
    const LOCAL_BLOB_URL_PREFIX = 'blob:'

    const CELL_EXTENSIONS = new Set(['xls', 'xlsx', 'xlsm', 'xlt', 'xltx', 'xltm', 'ods', 'fods', 'csv'])
    const SLIDE_EXTENSIONS = new Set(['ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'odp', 'fodp', 'pps', 'ppsx'])
    const PDF_EXTENSIONS = new Set(['pdf', 'djvu', 'oxps', 'xps'])

    const BEDOFFICE_LOGO_LIGHT = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA4NiAyMCIgd2lkdGg9Ijg2IiBoZWlnaHQ9IjIwIj4KICAgIDwhLS0gYmVkT2ZmaWNlIEJyYW5kIExvZ28gKExpZ2h0IFRoZW1lKSAtLT4KICAgIDxnIHRyYW5zZm9ybT0idHJhbnNsYXRlKDEsIDEpIHNjYWxlKDAuMDM4KSI+CiAgICAgICAgPGcgZmlsbC1ydWxlPSJub256ZXJvIiBzdHJva2U9IiMxQzUwNzYiIHN0cm9rZS13aWR0aD0iMTQiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+CiAgICAgICAgICAgIDwhLS0gSEVBREJPQVJEIChCbHVlKSAtLT4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIzODAsMzAwIDM5NSwyOTAgMzk1LDE1MCAzODAsMTYwIiBmaWxsPSIjNURCMEU0IiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjI4MCwxMDAgMzgwLDE2MCAzOTUsMTUwIDI5NSw5MCIgZmlsbD0iIzVEQjBFNCIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIzODAsMzAwIDI4MCwyNDAgMjgwLDEwMCAzODAsMTYwIiBmaWxsPSIjNURCMEU0IiAvPgoKICAgICAgICAgICAgPCEtLSBCRUQgQkFTRSAoQmx1ZSkgLS0+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjIwLDM4MCAxMjAsMzIwIDEyMCwzMDAgMjIwLDM2MCIgZmlsbD0iIzVEQjBFNCIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzgwIDM4MCwzMDAgMzgwLDI4MCAyMjAsMzYwIiBmaWxsPSIjNURCMEU0IiAvPgoKICAgICAgICAgICAgPCEtLSBNQVRUUkVTUyAoR3JlZW4pIC0tPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIyMCwzNjAgMTIwLDMwMCAxMjAsMjYwIDIyMCwzMjAiIGZpbGw9IiM5MEJBM0UiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjIwLDM2MCAzODAsMjgwIDM4MCwyNDAgMjIwLDMyMCIgZmlsbD0iIzkwQkEzRSIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzIwIDEyMCwyNjAgMjgwLDE4MCAzODAsMjQwIiBmaWxsPSIjOTBCQTNFIiAvPgoKICAgICAgICAgICAgPCEtLSBQSUxMT1cgKEdyZWVuKSAtLT4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMzAsMjEwIDMxMCwyNTAgMzEwLDI3MCAyMzAsMjMwIiBmaWxsPSIjOTBCQTNFIiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjMxMCwyNTAgMzQwLDIzNSAzNDAsMjU1IDMxMCwyNzAiIGZpbGw9IiM5MEJBM0UiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjMwLDIxMCAzMTAsMjUwIDM0MCwyMzUgMjYwLDE5NSIgZmlsbD0iIzkwQkEzRSIgLz4KICAgICAgICAgICAgPGxpbmUgeDE9IjI1MCIgeTE9IjIxNSIgeDI9IjMxNSIgeTI9IjI0OCIgLz4KCiAgICAgICAgICAgIDwhLS0gQkxBTktFVCAoT3JhbmdlKSAtLT4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzYwIDEyMCwzMDAgMTIwLDI2MCAyMjAsMzIwIiBmaWxsPSIjREU2MzNGIiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIyMCwzNjAgMzE2LDMxMiAzMTYsMjcyIDIyMCwzMjAiIGZpbGw9IiNERTYzM0YiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjIwLDMyMCAxMjAsMjYwIDIxNiwyMTIgMzE2LDI3MiIgZmlsbD0iI0RFNjMzRiIgLz4KICAgICAgICAgICAgPGxpbmUgeDE9IjI2MCIgeTE9IjI5NiIgeDI9IjE4MCIgeTI9IjI0OCIgLz4KICAgICAgICA8L2c+CiAgICA8L2c+CgogICAgPCEtLSBXb3JkbWFyayAtLT4KICAgIDx0ZXh0IHg9IjIxIiB5PSIxNC41IiBmb250LWZhbWlseT0iLWFwcGxlLXN5c3RlbSwgQmxpbmtNYWNTeXN0ZW1Gb250LCAnU2Vnb2UgVUknLCBSb2JvdG8sIEhlbHZldGljYSwgQXJpYWwsIHNhbnMtc2VyaWYiIGZvbnQtc2l6ZT0iMTEuNSIgZm9udC13ZWlnaHQ9IjgwMCIgZmlsbD0iIzBmMTcyYSIgbGV0dGVyLXNwYWNpbmc9Ii0wLjMiPmJlZDx0c3BhbiBmaWxsPSIjMjU2M2ViIiBmb250LXdlaWdodD0iNzAwIj5PZmZpY2U8L3RzcGFuPjwvdGV4dD4KPC9zdmc+Cg=='

    const BEDOFFICE_LOGO_DARK = 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA4NiAyMCIgd2lkdGg9Ijg2IiBoZWlnaHQ9IjIwIj4KICAgIDwhLS0gYmVkT2ZmaWNlIEJyYW5kIExvZ28gKERhcmsgVGhlbWUpIC0tPgogICAgPGcgdHJhbnNmb3JtPSJ0cmFuc2xhdGUoMSwgMSkgc2NhbGUoMC4wMzgpIj4KICAgICAgICA8ZyBmaWxsLXJ1bGU9Im5vbnplcm8iIHN0cm9rZT0iIzBmMTcyYSIgc3Ryb2tlLXdpZHRoPSIxNCIgc3Ryb2tlLWxpbmVjYXA9InJvdW5kIiBzdHJva2UtbGluZWpvaW49InJvdW5kIj4KICAgICAgICAgICAgPCEtLSBIRUFEQk9BUkQgKEJsdWUpIC0tPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjM4MCwzMDAgMzk1LDI5MCAzOTUsMTUwIDM4MCwxNjAiIGZpbGw9IiM1REIwRTQiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjgwLDEwMCAzODAsMTYwIDM5NSwxNTAgMjk1LDkwIiBmaWxsPSIjNURCMEU0IiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjM4MCwzMDAgMjgwLDI0MCAyODAsMTAwIDM4MCwxNjAiIGZpbGw9IiM1REIwRTQiIC8+CgogICAgICAgICAgICA8IS0tIEJFRCBCQVNFIChCbHVlKSAtLT4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzgwIDEyMCwzMjAgMTIwLDMwMCAyMjAsMzYwIiBmaWxsPSIjNURCMEU0IiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIyMCwzODAgMzgwLDMwMCAzODAsMjgwIDIyMCwzNjAiIGZpbGw9IiM1REIwRTQiIC8+CgogICAgICAgICAgICA8IS0tIE1BVFRSRVNTIChHcmVlbikgLS0+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjIwLDM2MCAxMjAsMzAwIDEyMCwyNjAgMjIwLDMyMCIgZmlsbD0iIzkwQkEzRSIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzYwIDM4MCwyODAgMzgwLDI0MCAyMjAsMzIwIiBmaWxsPSIjOTBCQTNFIiAvPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIyMCwzMjAgMTIwLDI2MCAyODAsMTgwIDM4MCwyNDAiIGZpbGw9IiM5MEJBM0UiIC8+CgogICAgICAgICAgICA8IS0tIFBJTExPVyAoR3JlZW4pIC0tPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIzMCwyMTAgMzEwLDI1MCAzMTAsMjcwIDIzMCwyMzAiIGZpbGw9IiM5MEJBM0UiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMzEwLDI1MCAzNDAsMjM1IDM0MCwyNTUgMzEwLDI3MCIgZmlsbD0iIzkwQkEzRSIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMzAsMjEwIDMxMCwyNTAgMzQwLDIzNSAyNjAsMTk1IiBmaWxsPSIjOTBCQTNFIiAvPgogICAgICAgICAgICA8bGluZSB4MT0iMjUwIiB5MT0iMjE1IiB4Mj0iMzE1IiB5Mj0iMjQ4IiAvPgoKICAgICAgICAgICAgPCEtLSBCTEFOS0VUIChPcmFuZ2UpIC0tPgogICAgICAgICAgICA8cG9seWdvbiBwb2ludHM9IjIyMCwzNjAgMTIwLDMwMCAxMjAsMjYwIDIyMCwzMjAiIGZpbGw9IiNERTYzM0YiIC8+CiAgICAgICAgICAgIDxwb2x5Z29uIHBvaW50cz0iMjIwLDM2MCAzMTYsMzEyIDMxNiwyNzIgMjIwLDMyMCIgZmlsbD0iI0RFNjMzRiIgLz4KICAgICAgICAgICAgPHBvbHlnb24gcG9pbnRzPSIyMjAsMzIwIDEyMCwyNjAgMjE2LDIxMiAzMTYsMjcyIiBmaWxsPSIjREU2MzNGIiAvPgogICAgICAgICAgICA8bGluZSB4MT0iMjYwIiB5MT0iMjk2IiB4Mj0iMTgwIiB5Mj0iMjQ4IiAvPgogICAgICAgIDwvZz4KICAgIDwvZz4KCiAgICA8IS0tIFdvcmRtYXJrIC0tPgogICAgPHRleHQgeD0iMjEiIHk9IjE0LjUiIGZvbnQtZmFtaWx5PSItYXBwbGUtc3lzdGVtLCBCbGlua01hY1N5c3RlbUZvbnQsICdTZWdvZSBVSScsIFJvYm90bywgSGVsdmV0aWNhLCBBcmlhbCwgc2Fucy1zZXJpZiIgZm9udC1zaXplPSIxMS41IiBmb250LXdlaWdodD0iODAwIiBmaWxsPSIjZmZmZmZmIiBsZXR0ZXItc3BhY2luZz0iLTAuMyI+YmVkPHRzcGFuIGZpbGw9IiM2MGE1ZmEiIGZvbnQtd2VpZ2h0PSI3MDAiPk9mZmljZTwvdHNwYW4+PC90ZXh0Pgo8L3N2Zz4K'

    /**
     * Normalizes a file extension string to lowercase.
     * @param {string} ext
     * @returns {string}
     */
    function normalizeExtension(ext) {
        return String(ext || '').toLowerCase()
    }

    /**
     * Resolves the ONLYOFFICE documentType from a file extension.
     * @param {string} ext
     * @returns {'word' | 'cell' | 'slide' | 'pdf'}
     */
    function documentTypeOf(ext) {
        const normalized = normalizeExtension(ext)
        if (CELL_EXTENSIONS.has(normalized)) return 'cell'
        if (SLIDE_EXTENSIONS.has(normalized)) return 'slide'
        if (PDF_EXTENSIONS.has(normalized)) return 'pdf'
        return 'word'
    }

    /**
     * Checks if a URL is a browser local blob URL.
     * @param {string} url
     * @returns {boolean}
     */
    function isLocalBlobUrl(url) {
        return String(url || '').startsWith(LOCAL_BLOB_URL_PREFIX)
    }

    /**
     * Determines whether document should use binary buffer loading (PDF only).
     * @param {string} documentType
     * @param {string} blobUrl
     * @returns {boolean}
     */
    function shouldOpenFromBinary(documentType, blobUrl) {
        return documentType === 'pdf' && isLocalBlobUrl(blobUrl)
    }

    function requireRecord(options) {
        if (!options || !options.record) {
            throw new Error('record is required')
        }
        if (!options.record.id) {
            throw new Error('record.id is required')
        }
        if (!options.record.name) {
            throw new Error('record.name is required')
        }
        if (!options.record.fileType) {
            throw new Error('record.fileType is required')
        }
        return options.record
    }

    function documentKeyOf(record) {
        return `${record.id}-${record.updatedAt || EMPTY_UPDATED_AT}`
    }

    function documentPermissions() {
        return { edit: true, download: true, print: true, fillForms: true, review: true, comment: true }
    }

    function localUser(userName) {
        return { id: LOCAL_USER_ID, name: userName || LOCAL_USER_NAME }
    }

    /**
     * Builds document configuration.
     * @param {Object} record
     * @param {string} [blobUrl]
     * @returns {Object}
     */
    function buildDocumentConfig(record, blobUrl) {
        const documentType = documentTypeOf(record.fileType)
        const openFromBinary = shouldOpenFromBinary(documentType, blobUrl)
        const config = {
            url: openFromBinary ? undefined : blobUrl || undefined,
            title: record.name,
            fileType: normalizeExtension(record.fileType),
            key: documentKeyOf(record),
            permissions: documentPermissions()
        }

        if (documentType === 'pdf') {
            const pdfConfig = Object.assign({}, config, { isForm: false })
            if (openFromBinary) pdfConfig.localOpenFromBinary = true
            return pdfConfig
        }
        return config
    }

    /**
     * Builds full ONLYOFFICE DocsAPI configuration.
     * @param {Object} options
     * @param {Object} options.record - Document metadata ({ id, name, fileType, updatedAt })
     * @param {string} [options.blobUrl] - Local blob or remote URL
     * @param {string} [options.lang] - UI language (defaults to 'en-US')
     * @param {string} [options.mode] - Editor mode ('edit' or 'view')
     * @returns {Object} Full docConfig
     */
    function buildOnlyofficeConfig(options) {
        const record = requireRecord(options)
        const documentType = documentTypeOf(record.fileType)
        const lang = options.lang || DEFAULT_LANG
        const mode = options.mode || DEFAULT_MODE
        const config = {
            document: buildDocumentConfig(record, options.blobUrl),
            documentType: documentType,
            editorConfig: { 
                mode: mode, 
                lang: lang, 
                user: localUser(options.userName),
                customization: {
                    close: {
                        visible: true,
                        text: 'Back to Documents'
                    },
                    goback: {
                        text: 'Back to Documents',
                        blank: false,
                        requestClose: true
                    },
                    logo: {
                        image: (options.logo && options.logo.image) || BEDOFFICE_LOGO_LIGHT,
                        imageDark: (options.logo && options.logo.imageDark) || BEDOFFICE_LOGO_DARK,
                        url: (options.logo && options.logo.url !== undefined) ? options.logo.url : 'https://milio48.github.io/bedoffice/'
                    },
                    autosave: true,
                    forcesave: true,
                    unit: 'cm',
                    compactHeader: false
                }
            }
        }
        if (shouldOpenFromBinary(documentType, options.blobUrl)) {
            config.localOpenFromBinary = true
        }
        return config
    }

    return {
        buildOnlyofficeConfig: buildOnlyofficeConfig,
        documentTypeOf: documentTypeOf,
        DEFAULT_LANG: DEFAULT_LANG
    }
}))
