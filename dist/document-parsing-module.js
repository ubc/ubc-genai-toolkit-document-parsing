"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.DocumentParsingModule = void 0;
const ubc_genai_toolkit_core_1 = require("ubc-genai-toolkit-core");
const error_1 = require("./error");
const path = __importStar(require("path"));
const fs = __importStar(require("fs/promises"));
const crypto_1 = require("crypto");
// Import parsing libraries
const pdf2md_1 = __importDefault(require("@opendocsg/pdf2md"));
const mammoth_1 = __importDefault(require("mammoth"));
const turndown_1 = __importDefault(require("turndown"));
const markdown_to_text_1 = __importDefault(require("markdown-to-text"));
const jszip_1 = __importDefault(require("jszip"));
const pngjs_1 = require("pngjs");
const PPTX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
/**
 * Placeholder URI scheme used to mark where each embedded DOCX image sat in
 * the converted markdown, so its textual description can be spliced in at the
 * exact position after (async) description completes.
 */
const DOCX_IMAGE_PLACEHOLDER_SCHEME = 'docparse-image://';
/**
 * Embedded PDF images narrower or shorter than this (in pixels) are treated
 * as typographic decorations (bullets, rules, glyph fragments) and skipped.
 */
const MIN_PDF_IMAGE_DIMENSION = 40;
/**
 * Embedded PDF images are downscaled so their long edge is at most this many
 * pixels before PNG encoding — plenty for vision models while keeping the
 * payload well under provider size limits.
 */
const MAX_PDF_IMAGE_DIMENSION = 1500;
/**
 * Marker @opendocsg/pdf2md places between pages in its markdown output. Used
 * to inline each PDF image description on the page it belongs to; always
 * stripped from the final content.
 */
const PDF2MD_PAGE_BREAK_MARKER = '<!-- PAGE_BREAK -->';
/** Maps common image file extensions found in PPTX archives to MIME types. */
const IMAGE_EXTENSION_MIME_TYPES = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.bmp': 'image/bmp',
    '.tif': 'image/tiff',
    '.tiff': 'image/tiff',
    '.webp': 'image/webp',
    '.svg': 'image/svg+xml',
};
// Define default configuration if any specific defaults are needed for this module
const DEFAULT_DOC_PARSING_CONFIG = {
    logger: new ubc_genai_toolkit_core_1.NoopLogger(), // Default to NoopLogger if none provided
    debug: false,
};
/**
 * The main facade class for the Document Parsing module.
 * Provides a simple interface to parse various document types.
 */
class DocumentParsingModule {
    /**
     * Creates an instance of DocumentParsingModule.
     *
     * @param {Partial<DocumentParsingConfig>} config - Optional configuration settings.
     */
    constructor(config) {
        // Merge provided config with defaults
        this.config = (0, ubc_genai_toolkit_core_1.mergeWithDefaults)(config, DEFAULT_DOC_PARSING_CONFIG);
        // Ensure a logger is available
        this.logger = this.config.logger || new ubc_genai_toolkit_core_1.NoopLogger();
        // Initialize Turndown service for HTML -> Markdown conversion
        this.turndownService = new turndown_1.default();
        if (this.config.debug) {
            this.logger.debug('DocumentParsingModule initialized', { config: this.config });
        }
    }
    /**
     * Parses the content of a document specified by the input parameters.
     *
     * @param {ParseInput} input - The input specification (e.g., { filePath: string }).
     * @param {SupportedOutputFormat} outputFormat - The desired output format ('text' or 'markdown').
     * @returns {Promise<ParsingResult>} A promise resolving to the parsing result (content and metadata).
     * @throws {UnsupportedFileTypeError} If the input file type is not supported.
     * @throws {ParsingError} If an error occurs during file access or parsing.
     */
    async parse(input, outputFormat) {
        this.logger.debug('Starting document parsing', { input, outputFormat });
        const { filePath } = input;
        // Basic check for file existence
        try {
            await fs.access(filePath);
        }
        catch (error) {
            this.logger.error('File not accessible', { filePath, error });
            throw new error_1.ParsingError('File not found or inaccessible', filePath, error);
        }
        // --- Step 4: Implement File Type Detection ---
        let detectedType;
        let finalType = 'unknown';
        try {
            detectedType = await this._detectFileType(filePath);
            finalType = detectedType || 'unknown'; // Assign to finalType for metadata
            this.logger.info(`Detected file type: ${finalType}`, { filePath });
            if (!detectedType) {
                throw new error_1.UnsupportedFileTypeError(undefined, filePath);
            }
        }
        catch (error) {
            if (error instanceof ubc_genai_toolkit_core_1.ToolkitError) {
                throw error; // Re-throw known toolkit errors
            }
            this.logger.error('Error during file type detection', { filePath, error });
            throw new error_1.ParsingError('Failed to determine file type', filePath, error);
        }
        // --- Step 5: Implement Parsing Logic ---
        let content = '';
        try {
            switch (detectedType) {
                case 'application/pdf':
                    content = await this._parsePdf(filePath, outputFormat);
                    break;
                case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                    content = await this._parseDocx(filePath, outputFormat);
                    break;
                case PPTX_MIME_TYPE:
                    content = await this._parsePptx(filePath, outputFormat);
                    break;
                case 'text/html':
                    content = await this._parseHtml(filePath, outputFormat);
                    break;
                case 'text/markdown':
                    content = await this._parseMarkdown(filePath, outputFormat);
                    break;
                default:
                    // This case should technically be unreachable due to the check above
                    this.logger.error('Reached default case in parsing logic - should not happen', { detectedType });
                    throw new error_1.UnsupportedFileTypeError(detectedType, filePath, { internalError: 'Parsing logic mismatch' });
            }
        }
        catch (error) {
            if (error instanceof ubc_genai_toolkit_core_1.ToolkitError) {
                // If it's already a ParsingError or UnsupportedFileTypeError from helpers, re-throw
                throw error;
            }
            // Wrap unexpected errors from libraries in ParsingError
            this.logger.error('Error during document parsing process', { filePath, detectedType, error });
            const message = error instanceof Error ? error.message : 'Unknown parsing error';
            throw new error_1.ParsingError(message, filePath, error);
        }
        // --- Construct Result ---
        const result = {
            content: content,
            metadata: {
                detectedInputType: finalType,
            },
        };
        this.logger.debug('Parsing completed successfully', { filePath, outputFormat });
        return result;
    }
    // --- Internal Helper Methods ---
    /**
     * Detects the file type using MIME type first, then falling back to extension.
     * @param filePath The path to the file.
     * @returns The detected supported MIME type or extension, or undefined if unsupported/error.
     * @throws {ParsingError} If MIME detection library throws an unexpected error.
     * @throws {UnsupportedFileTypeError} If the file extension is not supported after fallback.
     */
    async _detectFileType(filePath) {
        // 1. Try MIME type detection using dynamic import (via new Function to bypass tsc transform)
        let mimeType;
        try {
            // Use Function constructor for dynamic import to prevent tsc commonjs transform
            const dynamicImport = new Function('specifier', 'return import(specifier)');
            const fileTypeModule = await dynamicImport('file-type');
            const fileTypeFromFile = fileTypeModule.fileTypeFromFile;
            // Check if the function was loaded correctly
            if (typeof fileTypeFromFile !== 'function') {
                throw new Error('Failed to dynamically load fileTypeFromFile function.');
            }
            const fileTypeResult = await fileTypeFromFile(filePath);
            mimeType = fileTypeResult?.mime;
            if (mimeType) {
                this.logger.debug(`Detected MIME type: ${mimeType}`, { filePath });
                // Check if the detected MIME type is in our supported list
                switch (mimeType) {
                    case 'application/pdf':
                    case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                    case PPTX_MIME_TYPE:
                    case 'text/html':
                    case 'text/markdown':
                        return mimeType;
                }
            }
        }
        catch (error) {
            // Log warning but continue to fallback. If fileTypeFromFile throws unexpectedly, wrap it.
            this.logger.warn('MIME type detection failed or library errored.', { filePath, error });
            if (!(error instanceof Error && error.message.includes('ENOENT'))) { // Ignore simple file not found here, handled earlier
                // Re-throw unexpected errors from file-type library as ParsingError
                throw new error_1.ParsingError('MIME type detection library failed', filePath, error);
            }
        }
        // 2. Fallback to file extension
        const extension = path.extname(filePath).toLowerCase();
        this.logger.debug(`Falling back to file extension: ${extension}`, { filePath });
        switch (extension) {
            case '.pdf':
                return 'application/pdf';
            case '.docx':
                return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            case '.pptx':
                return PPTX_MIME_TYPE;
            case '.html':
            case '.htm':
                return 'text/html';
            case '.md':
                return 'text/markdown';
            default:
                this.logger.warn(`Unsupported file extension: ${extension}`, { filePath });
                // Throw specific error here instead of returning undefined
                throw new error_1.UnsupportedFileTypeError(extension, filePath);
        }
    }
    /**
     * Parses PDF content. Text is extracted with pdf2md as before; when an
     * {@link ImageDescriber} is configured, embedded raster images are also
     * extracted (per page, via pdfjs) and their textual descriptions inlined
     * as `> [Image, page N] ...` blocks at the end of the page they appear on
     * — mirroring how PPTX keeps descriptions with their slide. Image
     * extraction failures never fail the overall parse — the text-only result
     * is returned instead.
     */
    async _parsePdf(filePath, outputFormat) {
        this.logger.debug('Parsing PDF', { filePath, outputFormat });
        try {
            // Read the file content into a buffer
            const pdfBuffer = await fs.readFile(filePath);
            // Pass the buffer to pdf2md. Its output joins pages with an HTML
            // comment marker, which we use to keep image descriptions with their
            // page (and strip from the final output either way).
            const rawMarkdown = await (0, pdf2md_1.default)(pdfBuffer);
            const pageMarkdowns = rawMarkdown.split(PDF2MD_PAGE_BREAK_MARKER);
            const describer = this.config.imageDescriber;
            if (describer) {
                try {
                    const extracted = await this._extractPdfImages(pdfBuffer, filePath);
                    if (extracted.length > 0) {
                        const descriptions = await this._describeImages(describer, extracted.map((entry) => entry.image), extracted.map((entry) => entry.image.pageNumber ?? 0), { filePath, docType: 'PDF' });
                        // Group each described image (with its vertical position)
                        // by the page it belongs to.
                        const entriesByPage = new Map();
                        extracted.forEach((entry, index) => {
                            const description = descriptions[index];
                            if (!description) {
                                return;
                            }
                            const pageNumber = entry.image.pageNumber ?? 0;
                            const block = `> [Image, page ${pageNumber}] ${description}`;
                            const list = entriesByPage.get(pageNumber);
                            if (list) {
                                list.push({ ratio: entry.splitRatio, block });
                            }
                            else {
                                entriesByPage.set(pageNumber, [
                                    { ratio: entry.splitRatio, block },
                                ]);
                            }
                        });
                        for (const [pageNumber, entries] of entriesByPage) {
                            // Inline each description at its position within the
                            // page's own text (snapped to a paragraph boundary).
                            // A page-marker mismatch falls back to the last page.
                            const targetIndex = pageNumber >= 1 && pageNumber <= pageMarkdowns.length
                                ? pageNumber - 1
                                : pageMarkdowns.length - 1;
                            pageMarkdowns[targetIndex] = this._insertPdfImageBlocks(pageMarkdowns[targetIndex], entries);
                        }
                    }
                }
                catch (error) {
                    this.logger.warn('PDF image extraction failed; returning text-only content', { filePath, error });
                }
            }
            const markdownContent = pageMarkdowns
                .map((page) => page.trim())
                .filter((page) => page.length > 0)
                .join('\n\n');
            if (outputFormat === 'markdown') {
                return markdownContent;
            }
            else {
                // Convert markdown to text
                return (0, markdown_to_text_1.default)(markdownContent);
            }
        }
        catch (error) {
            this.logger.error('PDF parsing failed', { filePath, error });
            throw new error_1.ParsingError('Failed to parse PDF file', filePath, error);
        }
    }
    /**
     * Inserts image-description blocks into a page's markdown at the vertical
     * position each image occupies, expressed as a `ratio` in [0, 1] of the
     * characters above it. Insertion is snapped to a paragraph boundary so a
     * description never lands mid-sentence or inside a table. When the page has
     * no text, the blocks (in top-to-bottom order) become the page's content.
     */
    _insertPdfImageBlocks(pageText, entries) {
        const ordered = [...entries].sort((a, b) => a.ratio - b.ratio);
        const paragraphs = pageText
            .split(/\n{2,}/)
            .map((paragraph) => paragraph.trim())
            .filter((paragraph) => paragraph.length > 0);
        if (paragraphs.length === 0) {
            return ordered.map((entry) => entry.block).join('\n\n');
        }
        const totalChars = paragraphs.reduce((sum, paragraph) => sum + paragraph.length, 0);
        // For each image, choose the paragraph index to insert *after* (-1 means
        // before the first paragraph) by walking paragraphs until the cumulative
        // character count passes the image's target position.
        const blocksAfterIndex = new Map();
        for (const entry of ordered) {
            const targetChar = Math.max(0, Math.min(totalChars, entry.ratio * totalChars));
            let cumulative = 0;
            let insertAfter = -1;
            for (let i = 0; i < paragraphs.length; i++) {
                cumulative += paragraphs[i].length;
                if (cumulative <= targetChar) {
                    insertAfter = i;
                }
                else {
                    break;
                }
            }
            const list = blocksAfterIndex.get(insertAfter);
            if (list) {
                list.push(entry.block);
            }
            else {
                blocksAfterIndex.set(insertAfter, [entry.block]);
            }
        }
        const out = [];
        const beforeFirst = blocksAfterIndex.get(-1);
        if (beforeFirst) {
            out.push(...beforeFirst);
        }
        for (let i = 0; i < paragraphs.length; i++) {
            out.push(paragraphs[i]);
            const after = blocksAfterIndex.get(i);
            if (after) {
                out.push(...after);
            }
        }
        return out.join('\n\n');
    }
    /**
     * Extracts embedded raster images from a PDF, page by page, using pdfjs.
     * Decoded pixel data is re-encoded as PNG so any vision model can consume
     * it regardless of the image's original encoding inside the PDF. Tiny
     * images (bullets, rules, glyph fragments) are skipped, and any single
     * image that fails to decode is skipped without aborting extraction.
     *
     * Each returned image carries a `splitRatio` in [0, 1] describing where it
     * sits vertically within its page's text (0 = above all text, 1 = below all
     * text), computed from the fraction of the page's characters that lie above
     * the image. This lets the caller inline the description at the right point
     * in reading order rather than dumping it at the end of the page. pdf2md
     * discards positions, so the ratio is necessarily approximate and snapped to
     * a paragraph boundary; it defaults to 1 (append) when text geometry is
     * unavailable.
     */
    async _extractPdfImages(pdfBuffer, filePath) {
        // pdfjs-dist v4 is ESM-only; dynamic-import it the same way as file-type
        // (via Function constructor to survive the tsc CommonJS transform).
        const dynamicImport = new Function('specifier', 'return import(specifier)');
        const pdfjs = await dynamicImport('pdfjs-dist/legacy/build/pdf.mjs');
        // pdf2md (via its `unpdf` dependency) bundles a DIFFERENT pdfjs version
        // and leaves its worker on `globalThis.pdfjsWorker`, which pdfjs checks
        // BEFORE loading its own matching worker — producing an API/Worker
        // version-mismatch error. Hide any foreign global worker while our
        // document initialises (pdfjs caches its own worker after first use).
        const globalScope = globalThis;
        const foreignWorker = globalScope.pdfjsWorker;
        if (foreignWorker) {
            delete globalScope.pdfjsWorker;
        }
        let doc;
        try {
            doc = await pdfjs.getDocument({
                data: new Uint8Array(pdfBuffer),
                useSystemFonts: true,
                isEvalSupported: false,
                verbosity: 0,
            }).promise;
        }
        finally {
            if (foreignWorker) {
                globalScope.pdfjsWorker = foreignWorker;
            }
        }
        const results = [];
        try {
            for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber++) {
                let page;
                try {
                    page = await doc.getPage(pageNumber);
                    const opList = await page.getOperatorList();
                    // Collect the ids of image XObjects painted on this page, in
                    // paint order (ignoring repeat paints of the same object), and
                    // record the vertical centre of each image's first paint by
                    // tracking the current transformation matrix.
                    const seenIds = new Set();
                    const orderedIds = [];
                    const centerYById = new Map();
                    let ctm = [1, 0, 0, 1, 0, 0];
                    const ctmStack = [];
                    for (let i = 0; i < opList.fnArray.length; i++) {
                        const fn = opList.fnArray[i];
                        if (fn === pdfjs.OPS.save) {
                            ctmStack.push(ctm.slice());
                        }
                        else if (fn === pdfjs.OPS.restore) {
                            ctm = ctmStack.pop() || ctm;
                        }
                        else if (fn === pdfjs.OPS.transform) {
                            ctm = this._multiplyMatrix(ctm, opList.argsArray[i]);
                        }
                        else if (fn === pdfjs.OPS.paintImageXObject) {
                            const objId = opList.argsArray[i]?.[0];
                            if (typeof objId === 'string' && !seenIds.has(objId)) {
                                seenIds.add(objId);
                                orderedIds.push(objId);
                                // Image unit square maps y=0 -> ctm[5], y=1 -> ctm[5]+ctm[3].
                                const yA = ctm[5];
                                const yB = ctm[5] + ctm[3];
                                centerYById.set(objId, (yA + yB) / 2);
                            }
                        }
                    }
                    // Page text geometry: total characters and a helper that sums
                    // the characters lying above a given y (higher y == higher on
                    // the page in PDF user space).
                    let textItems = [];
                    let totalChars = 0;
                    try {
                        const textContent = await page.getTextContent();
                        for (const item of textContent.items) {
                            const str = item.str || '';
                            if (!str) {
                                continue;
                            }
                            textItems.push({ y: item.transform[5], len: str.length });
                            totalChars += str.length;
                        }
                    }
                    catch (error) {
                        this.logger.warn('Failed to read PDF page text geometry', {
                            filePath,
                            pageNumber,
                            error,
                        });
                        textItems = [];
                        totalChars = 0;
                    }
                    const splitRatioForCenter = (centerY) => {
                        if (totalChars === 0) {
                            return 1; // No text to interleave with — append.
                        }
                        let charsAbove = 0;
                        for (const item of textItems) {
                            if (item.y >= centerY) {
                                charsAbove += item.len;
                            }
                        }
                        return charsAbove / totalChars;
                    };
                    let imageIndex = 0;
                    for (const objId of orderedIds) {
                        try {
                            const imgObj = await this._resolvePdfObject(page, objId);
                            const png = this._pdfImageToPng(pdfjs, imgObj);
                            if (!png) {
                                continue; // Tiny/undecodable image.
                            }
                            const centerY = centerYById.get(objId);
                            results.push({
                                image: {
                                    data: png,
                                    mimeType: 'image/png',
                                    source: 'pdf',
                                    pageNumber,
                                    imageIndex: imageIndex++,
                                    fileName: objId,
                                },
                                splitRatio: centerY === undefined
                                    ? 1
                                    : splitRatioForCenter(centerY),
                            });
                        }
                        catch (error) {
                            this.logger.warn('Failed to decode embedded PDF image; skipping', { filePath, pageNumber, objId, error });
                        }
                    }
                }
                finally {
                    page?.cleanup();
                }
            }
        }
        finally {
            await doc.destroy();
        }
        return results;
    }
    /**
     * Multiplies the current transformation matrix by a pdfjs `transform`
     * operator's matrix (both in [a, b, c, d, e, f] form), returning the new
     * CTM. Used to locate where each image is painted on the page.
     */
    _multiplyMatrix(m, t) {
        const [a, b, c, d, e, f] = t;
        return [
            m[0] * a + m[2] * b,
            m[1] * a + m[3] * b,
            m[0] * c + m[2] * d,
            m[1] * c + m[3] * d,
            m[0] * e + m[2] * f + m[4],
            m[1] * e + m[3] * f + m[5],
        ];
    }
    /**
     * Resolves a pdfjs object id to its decoded image object, checking the
     * page-local store first and falling back to the document-common store.
     */
    _resolvePdfObject(page, objId) {
        return new Promise((resolve, reject) => {
            try {
                page.objs.get(objId, (obj) => resolve(obj));
            }
            catch {
                try {
                    page.commonObjs.get(objId, (obj) => resolve(obj));
                }
                catch (error) {
                    reject(error);
                }
            }
        });
    }
    /**
     * Converts a decoded pdfjs image object ({width, height, kind, data}) into
     * a PNG buffer, or returns undefined for images too small to carry content
     * (bullets, rules) or in a pixel layout we don't handle.
     */
    _pdfImageToPng(pdfjs, imgObj) {
        if (!imgObj || !imgObj.data || !imgObj.width || !imgObj.height) {
            return undefined;
        }
        const { width, height, kind } = imgObj;
        // Skip tiny decorations — they carry no instructional content and can
        // number in the hundreds in a typeset PDF.
        if (width < MIN_PDF_IMAGE_DIMENSION || height < MIN_PDF_IMAGE_DIMENSION) {
            return undefined;
        }
        const src = imgObj.data;
        const rgba = Buffer.alloc(width * height * 4);
        if (kind === pdfjs.ImageKind.RGBA_32BPP) {
            rgba.set(src.subarray(0, width * height * 4));
        }
        else if (kind === pdfjs.ImageKind.RGB_24BPP) {
            for (let p = 0, s = 0, d = 0; p < width * height; p++, s += 3, d += 4) {
                rgba[d] = src[s];
                rgba[d + 1] = src[s + 1];
                rgba[d + 2] = src[s + 2];
                rgba[d + 3] = 255;
            }
        }
        else if (kind === pdfjs.ImageKind.GRAYSCALE_1BPP) {
            // 1 bit per pixel, each row padded to a whole byte.
            const rowBytes = Math.ceil(width / 8);
            for (let y = 0; y < height; y++) {
                for (let x = 0; x < width; x++) {
                    const bit = (src[y * rowBytes + (x >> 3)] >> (7 - (x & 7))) & 1;
                    const value = bit ? 255 : 0;
                    const d = (y * width + x) * 4;
                    rgba[d] = value;
                    rgba[d + 1] = value;
                    rgba[d + 2] = value;
                    rgba[d + 3] = 255;
                }
            }
        }
        else {
            return undefined; // Unknown pixel layout.
        }
        // Downscale very large images before PNG-encoding: vision models don't
        // need more than ~1500px on the long edge, and raw re-encoded PDF images
        // can otherwise exceed provider payload limits.
        let outWidth = width;
        let outHeight = height;
        let outData = rgba;
        const longEdge = Math.max(width, height);
        if (longEdge > MAX_PDF_IMAGE_DIMENSION) {
            const scale = MAX_PDF_IMAGE_DIMENSION / longEdge;
            outWidth = Math.max(1, Math.round(width * scale));
            outHeight = Math.max(1, Math.round(height * scale));
            outData = this._downscaleRgba(rgba, width, height, outWidth, outHeight);
        }
        const png = new pngjs_1.PNG({ width: outWidth, height: outHeight });
        outData.copy(png.data);
        return pngjs_1.PNG.sync.write(png);
    }
    /**
     * Downscales RGBA pixel data with a box filter (averaging every source
     * pixel that maps onto each destination pixel), which keeps thin lines and
     * text in figures legible better than nearest-neighbour sampling.
     */
    _downscaleRgba(src, srcWidth, srcHeight, dstWidth, dstHeight) {
        const dst = Buffer.alloc(dstWidth * dstHeight * 4);
        for (let dy = 0; dy < dstHeight; dy++) {
            const y0 = Math.floor((dy * srcHeight) / dstHeight);
            const y1 = Math.max(y0 + 1, Math.floor(((dy + 1) * srcHeight) / dstHeight));
            for (let dx = 0; dx < dstWidth; dx++) {
                const x0 = Math.floor((dx * srcWidth) / dstWidth);
                const x1 = Math.max(x0 + 1, Math.floor(((dx + 1) * srcWidth) / dstWidth));
                let r = 0;
                let g = 0;
                let b = 0;
                let a = 0;
                const count = (y1 - y0) * (x1 - x0);
                for (let sy = y0; sy < y1; sy++) {
                    for (let sx = x0; sx < x1; sx++) {
                        const s = (sy * srcWidth + sx) * 4;
                        r += src[s];
                        g += src[s + 1];
                        b += src[s + 2];
                        a += src[s + 3];
                    }
                }
                const d = (dy * dstWidth + dx) * 4;
                dst[d] = Math.round(r / count);
                dst[d + 1] = Math.round(g / count);
                dst[d + 2] = Math.round(b / count);
                dst[d + 3] = Math.round(a / count);
            }
        }
        return dst;
    }
    /**
     * Parses DOCX content. Embedded images are never emitted as inline base64
     * blobs (mammoth's default, which produced enormous unreadable data URIs in
     * the markdown output). Instead, when an {@link ImageDescriber} is
     * configured each image is extracted and replaced in-place with a
     * `> [Image] ...` textual description; without a describer, images are
     * simply omitted.
     */
    async _parseDocx(filePath, outputFormat) {
        this.logger.debug('Parsing DOCX', { filePath, outputFormat });
        try {
            const describer = this.config.imageDescriber;
            // 1. Convert DOCX to HTML, intercepting embedded images. Each
            // describable image is collected and replaced by a positional
            // placeholder URI so its description can be inlined at the exact spot
            // the image appeared; everything else becomes an empty <img> that is
            // stripped after markdown conversion.
            const collectedImages = [];
            const convertImage = mammoth_1.default.images.imgElement(async (image) => {
                if (!describer) {
                    return { src: '' };
                }
                try {
                    const mimeType = image.contentType || 'application/octet-stream';
                    if (!this._isDescribableImage(mimeType)) {
                        return { src: '' }; // Vector/unknown formats vision models can't read.
                    }
                    // Newer mammoth exposes readAsBase64String; older read('base64').
                    const base64 = typeof image.readAsBase64String === 'function'
                        ? await image.readAsBase64String()
                        : await image.read('base64');
                    const index = collectedImages.length;
                    collectedImages.push({
                        data: Buffer.from(base64, 'base64'),
                        mimeType,
                        source: 'docx',
                        imageIndex: index,
                    });
                    return { src: `${DOCX_IMAGE_PLACEHOLDER_SCHEME}${index}` };
                }
                catch (error) {
                    this.logger.warn('Failed to read embedded DOCX image; skipping', { filePath, error });
                    return { src: '' };
                }
            });
            const { value: htmlContent } = await mammoth_1.default.convertToHtml({ path: filePath }, { convertImage });
            // 2. Convert HTML to Markdown
            let markdownContent = this.turndownService.turndown(htmlContent);
            // 3. Describe collected images (de-duplicated, bounded concurrency)
            // and splice each description into the placeholder position.
            if (describer && collectedImages.length > 0) {
                const descriptions = await this._describeImages(describer, collectedImages, 
                // Each occurrence is its own "unit": an image repeated many
                // times throughout the document is treated as decorative.
                collectedImages.map((img) => img.imageIndex + 1), { filePath, docType: 'DOCX' });
                const placeholderRegex = new RegExp(`!\\[[^\\]]*\\]\\(${DOCX_IMAGE_PLACEHOLDER_SCHEME}(\\d+)\\)`, 'g');
                markdownContent = markdownContent.replace(placeholderRegex, (_match, indexStr) => {
                    const description = descriptions[parseInt(indexStr, 10)];
                    return description ? `> [Image] ${description}` : '';
                });
            }
            // 4. Strip any leftover empty image tags (non-describable images, or
            // all images when no describer is configured).
            markdownContent = markdownContent.replace(/!\[[^\]]*\]\(\s*\)/g, '');
            if (outputFormat === 'markdown') {
                return markdownContent;
            }
            else {
                // 5. Convert Markdown to Text
                return (0, markdown_to_text_1.default)(markdownContent);
            }
        }
        catch (error) {
            this.logger.error('DOCX parsing failed', { filePath, error });
            throw new error_1.ParsingError('Failed to parse DOCX file', filePath, error);
        }
    }
    /**
     * Parses PPTX (PowerPoint) content using a pure-JS approach (no system
     * dependencies). Reads slide text and speaker notes directly from the OOXML
     * archive and, when an {@link ImageDescriber} is configured, inlines textual
     * descriptions of embedded images (charts, screenshots, pictures).
     */
    async _parsePptx(filePath, outputFormat) {
        this.logger.debug('Parsing PPTX', { filePath, outputFormat });
        try {
            const buffer = await fs.readFile(filePath);
            const zip = await jszip_1.default.loadAsync(buffer);
            const slidePaths = await this._getOrderedSlidePaths(zip);
            if (slidePaths.length === 0) {
                this.logger.warn('No slides found in PPTX', { filePath });
            }
            const describer = this.config.imageDescriber;
            const onSlide = this.config.onSlide;
            const slides = [];
            for (let i = 0; i < slidePaths.length; i++) {
                const slideNumber = i + 1;
                const slidePath = slidePaths[i];
                const slideXml = await zip.file(slidePath)?.async('string');
                if (!slideXml) {
                    continue;
                }
                const images = describer
                    ? await this._extractSlideImages(zip, slidePath, slideXml, slideNumber)
                    : [];
                slides.push({
                    slideNumber,
                    slideText: this._extractTextFromSlideXml(slideXml),
                    notes: await this._extractSlideNotes(zip, slidePath),
                    images,
                    descriptions: new Array(images.length),
                });
            }
            // --- Phase 2: describe images concurrently (bounded), if enabled. ---
            // De-duplication and decorative-template skipping happen inside
            // _describeImages; slow LLM calls run in parallel across the deck while
            // output order is always preserved.
            if (describer) {
                const flatImages = [];
                const flatUnits = [];
                for (const slide of slides) {
                    for (const image of slide.images) {
                        flatImages.push(image);
                        flatUnits.push(slide.slideNumber);
                    }
                }
                const flatDescriptions = await this._describeImages(describer, flatImages, flatUnits, { filePath, docType: 'PPTX' });
                // Fan the per-image descriptions back out onto each slide.
                let cursor = 0;
                for (const slide of slides) {
                    slide.descriptions = slide.images.map(() => flatDescriptions[cursor++]);
                }
            }
            // --- Phase 3: assemble each slide in order; deliver via onSlide. ---
            let describedImageCount = 0;
            const slideBlocks = [];
            for (const slide of slides) {
                const parts = [`## Slide ${slide.slideNumber}`];
                if (slide.slideText.trim()) {
                    parts.push(slide.slideText.trim());
                }
                let slideDescribedCount = 0;
                for (const description of slide.descriptions) {
                    if (description) {
                        slideDescribedCount++;
                        parts.push(`> [Image] ${description}`);
                    }
                }
                if (slide.notes.trim()) {
                    parts.push(`### Notes\n\n${slide.notes.trim()}`);
                }
                describedImageCount += slideDescribedCount;
                const slideMarkdown = parts.join('\n\n');
                slideBlocks.push(slideMarkdown);
                // Deliver this slide for per-slide ("chunked") processing if requested.
                if (onSlide) {
                    await onSlide({
                        slideNumber: slide.slideNumber,
                        markdown: slideMarkdown,
                        text: (0, markdown_to_text_1.default)(slideMarkdown),
                        describedImageCount: slideDescribedCount,
                    });
                }
            }
            if (describer) {
                this.logger.debug('PPTX image description complete', {
                    filePath,
                    describedImageCount,
                });
            }
            const markdownContent = slideBlocks.join('\n\n');
            if (outputFormat === 'markdown') {
                return markdownContent;
            }
            else {
                return (0, markdown_to_text_1.default)(markdownContent);
            }
        }
        catch (error) {
            if (error instanceof ubc_genai_toolkit_core_1.ToolkitError) {
                throw error;
            }
            this.logger.error('PPTX parsing failed', { filePath, error });
            throw new error_1.ParsingError('Failed to parse PPTX file', filePath, error);
        }
    }
    /**
     * Determines slide files in presentation (display) order. Uses the
     * presentation relationships when available, falling back to a numeric sort
     * of the slide files if the ordering metadata is missing or malformed.
     */
    async _getOrderedSlidePaths(zip) {
        const numericSortFallback = () => Object.keys(zip.files)
            .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
            .sort((a, b) => this._slideFileNumber(a) - this._slideFileNumber(b));
        try {
            const presentationXml = await zip
                .file('ppt/presentation.xml')
                ?.async('string');
            const relsXml = await zip
                .file('ppt/_rels/presentation.xml.rels')
                ?.async('string');
            if (!presentationXml || !relsXml) {
                return numericSortFallback();
            }
            // Map relationship id -> slide path (targets are relative to ppt/).
            const rels = this._parseRelationships(relsXml);
            const relIdToPath = new Map();
            for (const rel of rels) {
                if (rel.type.endsWith('/slide')) {
                    relIdToPath.set(rel.id, this._resolveZipPath('ppt', rel.target));
                }
            }
            // Read the ordered list of slide relationship ids from the deck.
            const orderedPaths = [];
            const sldIdRegex = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
            let match;
            while ((match = sldIdRegex.exec(presentationXml)) !== null) {
                const slidePath = relIdToPath.get(match[1]);
                if (slidePath && zip.file(slidePath)) {
                    orderedPaths.push(slidePath);
                }
            }
            return orderedPaths.length > 0 ? orderedPaths : numericSortFallback();
        }
        catch (error) {
            this.logger.warn('Failed to read PPTX slide order; falling back to numeric sort', { error });
            return numericSortFallback();
        }
    }
    /** Extracts the trailing slide number from a slide file path. */
    _slideFileNumber(slidePath) {
        const match = slidePath.match(/slide(\d+)\.xml$/);
        return match ? parseInt(match[1], 10) : 0;
    }
    /**
     * Describes a batch of embedded images via the consumer-supplied describer,
     * with the optimisations shared by every document format:
     *   1. De-duplication. Identical image bytes — an icon reused on every
     *      slide/page, a vector graphic stored beside its raster fallback, or
     *      the same picture placed several times — are described ONCE and the
     *      result reused wherever the image recurs.
     *   2. Decorative-template skipping. An image that appears in many distinct
     *      units (slides for PPTX, pages for PDF, occurrences for DOCX) is
     *      almost always a non-content element (bullet icon, logo, divider).
     *      Describing it wastes calls and, with smaller vision models, invites
     *      confident hallucination, so such images are skipped entirely.
     * Describer calls run in parallel with bounded concurrency; the returned
     * array is parallel to `images` (undefined = no description / skipped).
     *
     * @param images - The images to describe, in document order.
     * @param unitNumbers - Parallel array giving the unit (slide/page/occurrence
     *   number) each image belongs to, used by the decorative heuristic.
     */
    async _describeImages(describer, images, unitNumbers, logContext) {
        if (images.length === 0) {
            return [];
        }
        // Map each unique image (by content hash) to the distinct units it
        // appears in, keeping one representative occurrence to describe.
        const hashes = images.map((img) => (0, crypto_1.createHash)('sha1').update(img.data).digest('hex'));
        const unitsByHash = new Map();
        const representativeByHash = new Map();
        hashes.forEach((hash, index) => {
            let unitSet = unitsByHash.get(hash);
            if (!unitSet) {
                unitSet = new Set();
                unitsByHash.set(hash, unitSet);
                representativeByHash.set(hash, images[index]);
            }
            unitSet.add(unitNumbers[index]);
        });
        // An image reused in at least this many distinct units is treated as a
        // decorative/template element and skipped. A value <= 0 disables the
        // heuristic (every unique image is still described only once).
        const decorativeThreshold = this.config.decorativeImageSlideThreshold ?? 5;
        const decorativeHashes = new Set();
        if (decorativeThreshold > 0) {
            for (const [hash, unitSet] of unitsByHash) {
                if (unitSet.size >= decorativeThreshold) {
                    decorativeHashes.add(hash);
                }
            }
        }
        // Describe each unique, non-decorative image exactly once.
        const hashesToDescribe = [...representativeByHash.keys()].filter((hash) => !decorativeHashes.has(hash));
        const descriptionByHash = new Map();
        const concurrency = Math.max(1, this.config.imageConcurrency ?? 5);
        await this._mapWithConcurrency(hashesToDescribe, concurrency, async (hash) => {
            const description = await this._describeImageSafely(describer, representativeByHash.get(hash));
            descriptionByHash.set(hash, description && description.trim() ? description.trim() : undefined);
        });
        this.logger.debug(`${logContext.docType} image de-duplication summary`, {
            filePath: logContext.filePath,
            totalImagePlacements: images.length,
            uniqueImages: unitsByHash.size,
            describedImages: hashesToDescribe.length,
            skippedDecorativeImages: decorativeHashes.size,
        });
        // Fan the per-image descriptions back out onto every occurrence.
        // Decorative (skipped) images stay undefined and are omitted by callers.
        return hashes.map((hash) => descriptionByHash.get(hash));
    }
    /**
     * Runs `fn` over `items` with at most `limit` in flight at once. Used to
     * parallelize slow per-image describer calls while bounding concurrency.
     */
    async _mapWithConcurrency(items, limit, fn) {
        let next = 0;
        const worker = async () => {
            while (next < items.length) {
                const current = next++;
                await fn(items[current]);
            }
        };
        const workers = Array.from({ length: Math.min(Math.max(1, limit), items.length) }, () => worker());
        await Promise.all(workers);
    }
    /**
     * Extracts visible text from a slide's XML, preserving paragraph breaks.
     * Each `<a:p>` becomes a line; text runs (`<a:t>`) within it are concatenated.
     */
    _extractTextFromSlideXml(slideXml) {
        const paragraphs = [];
        const paragraphRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
        let pMatch;
        while ((pMatch = paragraphRegex.exec(slideXml)) !== null) {
            const runs = [];
            const runRegex = /<a:t>([\s\S]*?)<\/a:t>/g;
            let rMatch;
            while ((rMatch = runRegex.exec(pMatch[1])) !== null) {
                runs.push(this._decodeXmlEntities(rMatch[1]));
            }
            const line = runs.join('').trim();
            if (line) {
                paragraphs.push(line);
            }
        }
        return paragraphs.join('\n');
    }
    /** Extracts speaker-notes text for a given slide, if a notes slide exists. */
    async _extractSlideNotes(zip, slidePath) {
        const relsXml = await this._readSlideRels(zip, slidePath);
        if (!relsXml) {
            return '';
        }
        const baseDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
        const notesRel = this._parseRelationships(relsXml).find((rel) => rel.type.endsWith('/notesSlide'));
        if (!notesRel) {
            return '';
        }
        const notesPath = this._resolveZipPath(baseDir, notesRel.target);
        const notesXml = await zip.file(notesPath)?.async('string');
        return notesXml ? this._extractTextFromSlideXml(notesXml) : '';
    }
    /**
     * Collects embedded raster images placed on a slide, in their on-slide
     * order (via `<a:blip r:embed>` references), de-duplicated by relationship.
     */
    async _extractSlideImages(zip, slidePath, slideXml, slideNumber) {
        const relsXml = await this._readSlideRels(zip, slidePath);
        if (!relsXml) {
            return [];
        }
        const baseDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
        const relIdToTarget = new Map();
        for (const rel of this._parseRelationships(relsXml)) {
            if (rel.type.endsWith('/image')) {
                relIdToTarget.set(rel.id, rel.target);
            }
        }
        if (relIdToTarget.size === 0) {
            return [];
        }
        // Preserve on-slide order via blip embeds; ignore duplicate references.
        const embedRegex = /r:embed="([^"]+)"/g;
        const seen = new Set();
        const orderedRelIds = [];
        let match;
        while ((match = embedRegex.exec(slideXml)) !== null) {
            const relId = match[1];
            if (relIdToTarget.has(relId) && !seen.has(relId)) {
                seen.add(relId);
                orderedRelIds.push(relId);
            }
        }
        const images = [];
        for (const relId of orderedRelIds) {
            const target = relIdToTarget.get(relId);
            const mediaPath = this._resolveZipPath(baseDir, target);
            const mimeType = this._imageMimeFromPath(mediaPath);
            if (!this._isDescribableImage(mimeType)) {
                continue; // Skip vector/unknown formats (e.g. EMF/WMF) vision models can't read.
            }
            const data = await zip.file(mediaPath)?.async('nodebuffer');
            if (!data) {
                continue;
            }
            images.push({
                data,
                mimeType,
                source: 'pptx',
                slideNumber,
                imageIndex: images.length,
                fileName: mediaPath.substring(mediaPath.lastIndexOf('/') + 1),
            });
        }
        return images;
    }
    /**
     * Invokes the consumer-supplied describer, isolating failures so a single
     * problematic image never aborts the overall parse.
     */
    async _describeImageSafely(describer, image) {
        try {
            return await describer(image);
        }
        catch (error) {
            this.logger.warn('imageDescriber failed for embedded image; skipping', {
                slideNumber: image.slideNumber,
                fileName: image.fileName,
                error,
            });
            return undefined;
        }
    }
    /** Reads the `.rels` file associated with a slide, if present. */
    async _readSlideRels(zip, slidePath) {
        const dir = slidePath.substring(0, slidePath.lastIndexOf('/'));
        const file = slidePath.substring(slidePath.lastIndexOf('/') + 1);
        return zip.file(`${dir}/_rels/${file}.rels`)?.async('string');
    }
    /** Parses OOXML relationship entries into {id, type, target} records. */
    _parseRelationships(relsXml) {
        const rels = [];
        const relRegex = /<Relationship\b[^>]*\/?>/g;
        let match;
        while ((match = relRegex.exec(relsXml)) !== null) {
            const tag = match[0];
            const id = this._attr(tag, 'Id');
            const type = this._attr(tag, 'Type');
            const target = this._attr(tag, 'Target');
            if (id && type && target) {
                rels.push({ id, type, target });
            }
        }
        return rels;
    }
    /** Reads a single double-quoted attribute value from an XML tag string. */
    _attr(tag, name) {
        const match = tag.match(new RegExp(`\\b${name}="([^"]*)"`));
        return match ? match[1] : undefined;
    }
    /**
     * Resolves an OOXML relationship target (which may be relative, e.g.
     * `../media/image1.png`) against a base directory into a normalized zip path.
     */
    _resolveZipPath(baseDir, target) {
        if (target.startsWith('/')) {
            return target.replace(/^\/+/, '');
        }
        const segments = baseDir ? baseDir.split('/') : [];
        for (const part of target.split('/')) {
            if (part === '' || part === '.') {
                continue;
            }
            if (part === '..') {
                segments.pop();
            }
            else {
                segments.push(part);
            }
        }
        return segments.join('/');
    }
    /** Infers an image MIME type from a media file's extension. */
    _imageMimeFromPath(mediaPath) {
        const ext = path.extname(mediaPath).toLowerCase();
        return IMAGE_EXTENSION_MIME_TYPES[ext] || 'application/octet-stream';
    }
    /** Whether a MIME type is a raster image a vision model can typically read. */
    _isDescribableImage(mimeType) {
        switch (mimeType) {
            case 'image/png':
            case 'image/jpeg':
            case 'image/gif':
            case 'image/bmp':
            case 'image/webp':
            case 'image/tiff':
                return true;
            default:
                return false;
        }
    }
    /** Decodes the small set of XML entities that appear in OOXML text runs. */
    _decodeXmlEntities(text) {
        return text
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
            .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
            .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
            .replace(/&amp;/g, '&');
    }
    /** Parses HTML content. */
    async _parseHtml(filePath, outputFormat) {
        this.logger.debug('Parsing HTML', { filePath, outputFormat });
        try {
            // 1. Read HTML file content
            const htmlContent = await fs.readFile(filePath, 'utf-8');
            // 2. Convert HTML to Markdown
            const markdownContent = this.turndownService.turndown(htmlContent);
            if (outputFormat === 'markdown') {
                return markdownContent;
            }
            else {
                // 3. Convert Markdown to Text
                return (0, markdown_to_text_1.default)(markdownContent);
            }
        }
        catch (error) {
            this.logger.error('HTML parsing failed', { filePath, error });
            throw new error_1.ParsingError('Failed to parse HTML file', filePath, error);
        }
    }
    /** Parses Markdown content. */
    async _parseMarkdown(filePath, outputFormat) {
        this.logger.debug('Parsing Markdown', { filePath, outputFormat });
        try {
            // 1. Read Markdown file content
            const markdownContent = await fs.readFile(filePath, 'utf-8');
            if (outputFormat === 'markdown') {
                return markdownContent;
            }
            else {
                // 2. Convert Markdown to Text
                return (0, markdown_to_text_1.default)(markdownContent);
            }
        }
        catch (error) {
            this.logger.error('Markdown parsing/reading failed', { filePath, error });
            throw new error_1.ParsingError('Failed to parse Markdown file', filePath, error);
        }
    }
}
exports.DocumentParsingModule = DocumentParsingModule;
//# sourceMappingURL=document-parsing-module.js.map