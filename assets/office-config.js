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

    const CELL_EXTENSIONS = new Set(['xls', 'xlsx', 'xlsm', 'ods', 'csv'])
    const SLIDE_EXTENSIONS = new Set(['ppt', 'pptx', 'odp', 'pps', 'ppsx'])
    const PDF_EXTENSIONS = new Set(['pdf', 'oxps', 'xps', 'djvu'])

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
        return { edit: true, download: true, print: true }
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
        return {
            fileType: normalizeExtension(record.fileType),
            key: documentKeyOf(record),
            title: record.name,
            url: blobUrl || '',
            permissions: documentPermissions()
        }
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
                    goback: {
                        text: 'Back to Documents',
                        blank: false,
                        requestClose: true
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
