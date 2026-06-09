import {
	LoggerInterface,
	ModuleConfig,
	NoopLogger,
	mergeWithDefaults,
	ToolkitError,
} from 'ubc-genai-toolkit-core';
import {
	DocumentParsingConfig,
	ParseInput,
	ParsingResult,
	SupportedOutputFormat,
	SupportedInputMimeType,
	SupportedInputExtension,
	EmbeddedImage,
} from './types';
import { UnsupportedFileTypeError, ParsingError } from './error';
import * as path from 'path';
import * as fs from 'fs/promises';

// Import parsing libraries
import pdf2md from '@opendocsg/pdf2md';
import mammoth from 'mammoth';
import TurndownService from 'turndown';
import markdownToText from 'markdown-to-text';
import JSZip from 'jszip';

const PPTX_MIME_TYPE =
	'application/vnd.openxmlformats-officedocument.presentationml.presentation';

/** Maps common image file extensions found in PPTX archives to MIME types. */
const IMAGE_EXTENSION_MIME_TYPES: Record<string, string> = {
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
const DEFAULT_DOC_PARSING_CONFIG: Partial<DocumentParsingConfig> = {
	logger: new NoopLogger(), // Default to NoopLogger if none provided
	debug: false,
};

/**
 * The main facade class for the Document Parsing module.
 * Provides a simple interface to parse various document types.
 */
export class DocumentParsingModule {
	private config: DocumentParsingConfig;
	private logger: LoggerInterface;
	private turndownService: TurndownService;

	/**
	 * Creates an instance of DocumentParsingModule.
	 *
	 * @param {Partial<DocumentParsingConfig>} config - Optional configuration settings.
	 */
	constructor(config?: Partial<DocumentParsingConfig>) {
		// Merge provided config with defaults
		this.config = mergeWithDefaults<DocumentParsingConfig>(
			config,
			DEFAULT_DOC_PARSING_CONFIG
		);
		// Ensure a logger is available
		this.logger = this.config.logger || new NoopLogger();
		// Initialize Turndown service for HTML -> Markdown conversion
		this.turndownService = new TurndownService();

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
	async parse(
		input: ParseInput,
		outputFormat: SupportedOutputFormat
	): Promise<ParsingResult> {
		this.logger.debug('Starting document parsing', { input, outputFormat });

		const { filePath } = input;

		// Basic check for file existence
		try {
			await fs.access(filePath);
		} catch (error) {
			this.logger.error('File not accessible', { filePath, error });
			throw new ParsingError('File not found or inaccessible', filePath, error);
		}

		// --- Step 4: Implement File Type Detection ---
		let detectedType: SupportedInputMimeType | SupportedInputExtension | undefined;
		let finalType: SupportedInputMimeType | SupportedInputExtension | 'unknown' = 'unknown';

		try {
			detectedType = await this._detectFileType(filePath);
			finalType = detectedType || 'unknown'; // Assign to finalType for metadata
			this.logger.info(`Detected file type: ${finalType}`, { filePath });

			if (!detectedType) {
				throw new UnsupportedFileTypeError(undefined, filePath);
			}
		} catch (error) {
			if (error instanceof ToolkitError) {
				throw error; // Re-throw known toolkit errors
			}
			this.logger.error('Error during file type detection', { filePath, error });
			throw new ParsingError('Failed to determine file type', filePath, error);
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
					throw new UnsupportedFileTypeError(detectedType, filePath, { internalError: 'Parsing logic mismatch'});
			}
		} catch (error) {
			if (error instanceof ToolkitError) {
				// If it's already a ParsingError or UnsupportedFileTypeError from helpers, re-throw
				throw error;
			}
			// Wrap unexpected errors from libraries in ParsingError
			this.logger.error('Error during document parsing process', { filePath, detectedType, error });
			const message = error instanceof Error ? error.message : 'Unknown parsing error';
			throw new ParsingError(message, filePath, error);
		}

		// --- Construct Result ---
		const result: ParsingResult = {
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
	private async _detectFileType(
		filePath: string
	): Promise<SupportedInputMimeType | SupportedInputExtension | undefined> {
		// 1. Try MIME type detection using dynamic import (via new Function to bypass tsc transform)
		let mimeType: string | undefined;
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
						return mimeType as SupportedInputMimeType;
				}
			}
		} catch (error) {
			// Log warning but continue to fallback. If fileTypeFromFile throws unexpectedly, wrap it.
			this.logger.warn('MIME type detection failed or library errored.', { filePath, error });
			if (!(error instanceof Error && error.message.includes('ENOENT'))) { // Ignore simple file not found here, handled earlier
                // Re-throw unexpected errors from file-type library as ParsingError
                throw new ParsingError('MIME type detection library failed', filePath, error);
            }
		}

		// 2. Fallback to file extension
		const extension = path.extname(filePath).toLowerCase() as SupportedInputExtension;
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
				throw new UnsupportedFileTypeError(extension, filePath);
		}
	}

	/** Parses PDF content. */
	private async _parsePdf(
		filePath: string,
		outputFormat: SupportedOutputFormat
	): Promise<string> {
		this.logger.debug('Parsing PDF', { filePath, outputFormat });
		try {
			// Read the file content into a buffer
			const pdfBuffer = await fs.readFile(filePath);

			// Pass the buffer to pdf2md
			const markdownContent = await pdf2md(pdfBuffer);
			if (outputFormat === 'markdown') {
				return markdownContent;
			} else {
				// Convert markdown to text
				return markdownToText(markdownContent);
			}
		} catch (error) {
			this.logger.error('PDF parsing failed', { filePath, error });
			throw new ParsingError('Failed to parse PDF file', filePath, error);
		}
	}

	/** Parses DOCX content. */
	private async _parseDocx(
		filePath: string,
		outputFormat: SupportedOutputFormat
	): Promise<string> {
		this.logger.debug('Parsing DOCX', { filePath, outputFormat });
		try {
			// 1. Convert DOCX to HTML
			const { value: htmlContent } = await mammoth.convertToHtml({ path: filePath });

			// 2. Convert HTML to Markdown
			const markdownContent = this.turndownService.turndown(htmlContent);

			if (outputFormat === 'markdown') {
				return markdownContent;
			} else {
				// 3. Convert Markdown to Text
				return markdownToText(markdownContent);
			}
		} catch (error) {
			this.logger.error('DOCX parsing failed', { filePath, error });
			throw new ParsingError('Failed to parse DOCX file', filePath, error);
		}
	}

	/**
	 * Parses PPTX (PowerPoint) content using a pure-JS approach (no system
	 * dependencies). Reads slide text and speaker notes directly from the OOXML
	 * archive and, when an {@link ImageDescriber} is configured, inlines textual
	 * descriptions of embedded images (charts, screenshots, pictures).
	 */
	private async _parsePptx(
		filePath: string,
		outputFormat: SupportedOutputFormat
	): Promise<string> {
		this.logger.debug('Parsing PPTX', { filePath, outputFormat });
		try {
			const buffer = await fs.readFile(filePath);
			const zip = await JSZip.loadAsync(buffer);

			const slidePaths = await this._getOrderedSlidePaths(zip);
			if (slidePaths.length === 0) {
				this.logger.warn('No slides found in PPTX', { filePath });
			}

			const describer = this.config.imageDescriber;
			const onSlide = this.config.onSlide;

			// --- Phase 1: extract each slide's text, notes and embedded images. ---
			interface SlideWork {
				slideNumber: number;
				slideText: string;
				notes: string;
				images: EmbeddedImage[];
				descriptions: (string | undefined)[];
			}
			const slides: SlideWork[] = [];
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

			// --- Phase 2: describe all images concurrently (bounded), if enabled. ---
			// Slow LLM calls run in parallel across the whole deck so image-heavy
			// files don't take (sum of every call) time; output order is preserved.
			if (describer) {
				const tasks: Array<{ slide: SlideWork; index: number }> = [];
				for (const slide of slides) {
					slide.images.forEach((_image, index) => tasks.push({ slide, index }));
				}
				const concurrency = Math.max(1, this.config.imageConcurrency ?? 5);
				await this._mapWithConcurrency(tasks, concurrency, async (task) => {
					const description = await this._describeImageSafely(
						describer,
						task.slide.images[task.index]
					);
					task.slide.descriptions[task.index] =
						description && description.trim() ? description.trim() : undefined;
				});
			}

			// --- Phase 3: assemble each slide in order; deliver via onSlide. ---
			let describedImageCount = 0;
			const slideBlocks: string[] = [];
			for (const slide of slides) {
				const parts: string[] = [`## Slide ${slide.slideNumber}`];
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
						text: markdownToText(slideMarkdown),
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
			} else {
				return markdownToText(markdownContent);
			}
		} catch (error) {
			if (error instanceof ToolkitError) {
				throw error;
			}
			this.logger.error('PPTX parsing failed', { filePath, error });
			throw new ParsingError('Failed to parse PPTX file', filePath, error);
		}
	}

	/**
	 * Determines slide files in presentation (display) order. Uses the
	 * presentation relationships when available, falling back to a numeric sort
	 * of the slide files if the ordering metadata is missing or malformed.
	 */
	private async _getOrderedSlidePaths(zip: JSZip): Promise<string[]> {
		const numericSortFallback = (): string[] =>
			Object.keys(zip.files)
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
			const relIdToPath = new Map<string, string>();
			for (const rel of rels) {
				if (rel.type.endsWith('/slide')) {
					relIdToPath.set(rel.id, this._resolveZipPath('ppt', rel.target));
				}
			}

			// Read the ordered list of slide relationship ids from the deck.
			const orderedPaths: string[] = [];
			const sldIdRegex = /<p:sldId\b[^>]*\br:id="([^"]+)"/g;
			let match: RegExpExecArray | null;
			while ((match = sldIdRegex.exec(presentationXml)) !== null) {
				const slidePath = relIdToPath.get(match[1]);
				if (slidePath && zip.file(slidePath)) {
					orderedPaths.push(slidePath);
				}
			}

			return orderedPaths.length > 0 ? orderedPaths : numericSortFallback();
		} catch (error) {
			this.logger.warn(
				'Failed to read PPTX slide order; falling back to numeric sort',
				{ error }
			);
			return numericSortFallback();
		}
	}

	/** Extracts the trailing slide number from a slide file path. */
	private _slideFileNumber(slidePath: string): number {
		const match = slidePath.match(/slide(\d+)\.xml$/);
		return match ? parseInt(match[1], 10) : 0;
	}

	/**
	 * Runs `fn` over `items` with at most `limit` in flight at once. Used to
	 * parallelize slow per-image describer calls while bounding concurrency.
	 */
	private async _mapWithConcurrency<T>(
		items: T[],
		limit: number,
		fn: (item: T) => Promise<void>
	): Promise<void> {
		let next = 0;
		const worker = async (): Promise<void> => {
			while (next < items.length) {
				const current = next++;
				await fn(items[current]);
			}
		};
		const workers = Array.from(
			{ length: Math.min(Math.max(1, limit), items.length) },
			() => worker()
		);
		await Promise.all(workers);
	}

	/**
	 * Extracts visible text from a slide's XML, preserving paragraph breaks.
	 * Each `<a:p>` becomes a line; text runs (`<a:t>`) within it are concatenated.
	 */
	private _extractTextFromSlideXml(slideXml: string): string {
		const paragraphs: string[] = [];
		const paragraphRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
		let pMatch: RegExpExecArray | null;
		while ((pMatch = paragraphRegex.exec(slideXml)) !== null) {
			const runs: string[] = [];
			const runRegex = /<a:t>([\s\S]*?)<\/a:t>/g;
			let rMatch: RegExpExecArray | null;
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
	private async _extractSlideNotes(
		zip: JSZip,
		slidePath: string
	): Promise<string> {
		const relsXml = await this._readSlideRels(zip, slidePath);
		if (!relsXml) {
			return '';
		}
		const baseDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
		const notesRel = this._parseRelationships(relsXml).find((rel) =>
			rel.type.endsWith('/notesSlide')
		);
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
	private async _extractSlideImages(
		zip: JSZip,
		slidePath: string,
		slideXml: string,
		slideNumber: number
	): Promise<EmbeddedImage[]> {
		const relsXml = await this._readSlideRels(zip, slidePath);
		if (!relsXml) {
			return [];
		}
		const baseDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
		const relIdToTarget = new Map<string, string>();
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
		const seen = new Set<string>();
		const orderedRelIds: string[] = [];
		let match: RegExpExecArray | null;
		while ((match = embedRegex.exec(slideXml)) !== null) {
			const relId = match[1];
			if (relIdToTarget.has(relId) && !seen.has(relId)) {
				seen.add(relId);
				orderedRelIds.push(relId);
			}
		}

		const images: EmbeddedImage[] = [];
		for (const relId of orderedRelIds) {
			const target = relIdToTarget.get(relId)!;
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
	private async _describeImageSafely(
		describer: NonNullable<DocumentParsingConfig['imageDescriber']>,
		image: EmbeddedImage
	): Promise<string | null | undefined> {
		try {
			return await describer(image);
		} catch (error) {
			this.logger.warn('imageDescriber failed for embedded image; skipping', {
				slideNumber: image.slideNumber,
				fileName: image.fileName,
				error,
			});
			return undefined;
		}
	}

	/** Reads the `.rels` file associated with a slide, if present. */
	private async _readSlideRels(
		zip: JSZip,
		slidePath: string
	): Promise<string | undefined> {
		const dir = slidePath.substring(0, slidePath.lastIndexOf('/'));
		const file = slidePath.substring(slidePath.lastIndexOf('/') + 1);
		return zip.file(`${dir}/_rels/${file}.rels`)?.async('string');
	}

	/** Parses OOXML relationship entries into {id, type, target} records. */
	private _parseRelationships(
		relsXml: string
	): Array<{ id: string; type: string; target: string }> {
		const rels: Array<{ id: string; type: string; target: string }> = [];
		const relRegex = /<Relationship\b[^>]*\/?>/g;
		let match: RegExpExecArray | null;
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
	private _attr(tag: string, name: string): string | undefined {
		const match = tag.match(new RegExp(`\\b${name}="([^"]*)"`));
		return match ? match[1] : undefined;
	}

	/**
	 * Resolves an OOXML relationship target (which may be relative, e.g.
	 * `../media/image1.png`) against a base directory into a normalized zip path.
	 */
	private _resolveZipPath(baseDir: string, target: string): string {
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
			} else {
				segments.push(part);
			}
		}
		return segments.join('/');
	}

	/** Infers an image MIME type from a media file's extension. */
	private _imageMimeFromPath(mediaPath: string): string {
		const ext = path.extname(mediaPath).toLowerCase();
		return IMAGE_EXTENSION_MIME_TYPES[ext] || 'application/octet-stream';
	}

	/** Whether a MIME type is a raster image a vision model can typically read. */
	private _isDescribableImage(mimeType: string): boolean {
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
	private _decodeXmlEntities(text: string): string {
		return text
			.replace(/&lt;/g, '<')
			.replace(/&gt;/g, '>')
			.replace(/&quot;/g, '"')
			.replace(/&apos;/g, "'")
			.replace(/&#x([0-9a-fA-F]+);/g, (_, hex) =>
				String.fromCodePoint(parseInt(hex, 16))
			)
			.replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
			.replace(/&amp;/g, '&');
	}

	/** Parses HTML content. */
	private async _parseHtml(
		filePath: string,
		outputFormat: SupportedOutputFormat
	): Promise<string> {
		this.logger.debug('Parsing HTML', { filePath, outputFormat });
		try {
			// 1. Read HTML file content
			const htmlContent = await fs.readFile(filePath, 'utf-8');

			// 2. Convert HTML to Markdown
			const markdownContent = this.turndownService.turndown(htmlContent);

			if (outputFormat === 'markdown') {
				return markdownContent;
			} else {
				// 3. Convert Markdown to Text
				return markdownToText(markdownContent);
			}
		} catch (error) {
			this.logger.error('HTML parsing failed', { filePath, error });
			throw new ParsingError('Failed to parse HTML file', filePath, error);
		}
	}

	/** Parses Markdown content. */
	private async _parseMarkdown(
		filePath: string,
		outputFormat: SupportedOutputFormat
	): Promise<string> {
		this.logger.debug('Parsing Markdown', { filePath, outputFormat });
		try {
			// 1. Read Markdown file content
			const markdownContent = await fs.readFile(filePath, 'utf-8');

			if (outputFormat === 'markdown') {
				return markdownContent;
			} else {
				// 2. Convert Markdown to Text
				return markdownToText(markdownContent);
			}
		} catch (error) {
			this.logger.error('Markdown parsing/reading failed', { filePath, error });
			throw new ParsingError('Failed to parse Markdown file', filePath, error);
		}
	}
}