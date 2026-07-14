import { DocumentParsingConfig, ParseInput, ParsingResult, SupportedOutputFormat } from './types';
/**
 * The main facade class for the Document Parsing module.
 * Provides a simple interface to parse various document types.
 */
export declare class DocumentParsingModule {
    private config;
    private logger;
    private turndownService;
    /**
     * Creates an instance of DocumentParsingModule.
     *
     * @param {Partial<DocumentParsingConfig>} config - Optional configuration settings.
     */
    constructor(config?: Partial<DocumentParsingConfig>);
    /**
     * Parses the content of a document specified by the input parameters.
     *
     * @param {ParseInput} input - The input specification (e.g., { filePath: string }).
     * @param {SupportedOutputFormat} outputFormat - The desired output format ('text' or 'markdown').
     * @returns {Promise<ParsingResult>} A promise resolving to the parsing result (content and metadata).
     * @throws {UnsupportedFileTypeError} If the input file type is not supported.
     * @throws {ParsingError} If an error occurs during file access or parsing.
     */
    parse(input: ParseInput, outputFormat: SupportedOutputFormat): Promise<ParsingResult>;
    /**
     * Detects the file type using MIME type first, then falling back to extension.
     * @param filePath The path to the file.
     * @returns The detected supported MIME type or extension, or undefined if unsupported/error.
     * @throws {ParsingError} If MIME detection library throws an unexpected error.
     * @throws {UnsupportedFileTypeError} If the file extension is not supported after fallback.
     */
    private _detectFileType;
    /**
     * Parses PDF content. Text is extracted with pdf2md as before; when an
     * {@link ImageDescriber} is configured, embedded raster images are also
     * extracted (per page, via pdfjs) and their textual descriptions inlined
     * as `> [Image, page N] ...` blocks at the end of the page they appear on
     * — mirroring how PPTX keeps descriptions with their slide. Image
     * extraction failures never fail the overall parse — the text-only result
     * is returned instead.
     */
    private _parsePdf;
    /**
     * Inserts image-description blocks into a page's markdown at the vertical
     * position each image occupies, expressed as a `ratio` in [0, 1] of the
     * characters above it. Insertion is snapped to a paragraph boundary so a
     * description never lands mid-sentence or inside a table. When the page has
     * no text, the blocks (in top-to-bottom order) become the page's content.
     */
    private _insertPdfImageBlocks;
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
    private _extractPdfImages;
    /**
     * Multiplies the current transformation matrix by a pdfjs `transform`
     * operator's matrix (both in [a, b, c, d, e, f] form), returning the new
     * CTM. Used to locate where each image is painted on the page.
     */
    private _multiplyMatrix;
    /**
     * Resolves a pdfjs object id to its decoded image object, checking the
     * page-local store first and falling back to the document-common store.
     */
    private _resolvePdfObject;
    /**
     * Converts a decoded pdfjs image object ({width, height, kind, data}) into
     * a PNG buffer, or returns undefined for images too small to carry content
     * (bullets, rules) or in a pixel layout we don't handle.
     */
    private _pdfImageToPng;
    /**
     * Downscales RGBA pixel data with a box filter (averaging every source
     * pixel that maps onto each destination pixel), which keeps thin lines and
     * text in figures legible better than nearest-neighbour sampling.
     */
    private _downscaleRgba;
    /**
     * Parses DOCX content. Embedded images are never emitted as inline base64
     * blobs (mammoth's default, which produced enormous unreadable data URIs in
     * the markdown output). Instead, when an {@link ImageDescriber} is
     * configured each image is extracted and replaced in-place with a
     * `> [Image] ...` textual description; without a describer, images are
     * simply omitted.
     */
    private _parseDocx;
    /**
     * Parses PPTX (PowerPoint) content using a pure-JS approach (no system
     * dependencies). Reads slide text and speaker notes directly from the OOXML
     * archive and, when an {@link ImageDescriber} is configured, inlines textual
     * descriptions of embedded images (charts, screenshots, pictures).
     */
    private _parsePptx;
    /**
     * Determines slide files in presentation (display) order. Uses the
     * presentation relationships when available, falling back to a numeric sort
     * of the slide files if the ordering metadata is missing or malformed.
     */
    private _getOrderedSlidePaths;
    /** Extracts the trailing slide number from a slide file path. */
    private _slideFileNumber;
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
    private _describeImages;
    /**
     * Runs `fn` over `items` with at most `limit` in flight at once. Used to
     * parallelize slow per-image describer calls while bounding concurrency.
     */
    private _mapWithConcurrency;
    /**
     * Extracts visible text from a slide's XML, preserving paragraph breaks.
     * Each `<a:p>` becomes a line; text runs (`<a:t>`) within it are concatenated.
     */
    private _extractTextFromSlideXml;
    /** Extracts speaker-notes text for a given slide, if a notes slide exists. */
    private _extractSlideNotes;
    /**
     * Collects embedded raster images placed on a slide, in their on-slide
     * order (via `<a:blip r:embed>` references), de-duplicated by relationship.
     */
    private _extractSlideImages;
    /**
     * Invokes the consumer-supplied describer, isolating failures so a single
     * problematic image never aborts the overall parse.
     */
    private _describeImageSafely;
    /** Reads the `.rels` file associated with a slide, if present. */
    private _readSlideRels;
    /** Parses OOXML relationship entries into {id, type, target} records. */
    private _parseRelationships;
    /** Reads a single double-quoted attribute value from an XML tag string. */
    private _attr;
    /**
     * Resolves an OOXML relationship target (which may be relative, e.g.
     * `../media/image1.png`) against a base directory into a normalized zip path.
     */
    private _resolveZipPath;
    /** Infers an image MIME type from a media file's extension. */
    private _imageMimeFromPath;
    /** Whether a MIME type is a raster image a vision model can typically read. */
    private _isDescribableImage;
    /** Decodes the small set of XML entities that appear in OOXML text runs. */
    private _decodeXmlEntities;
    /** Parses HTML content. */
    private _parseHtml;
    /** Parses Markdown content. */
    private _parseMarkdown;
}
//# sourceMappingURL=document-parsing-module.d.ts.map