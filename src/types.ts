import { ModuleConfig } from 'ubc-genai-toolkit-core';

/**
 * Supported input document MIME types (or extensions as fallback).
 * Using common MIME types.
 */
export type SupportedInputMimeType =
  | 'application/pdf'
  | 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' // .docx
  | 'application/vnd.openxmlformats-officedocument.presentationml.presentation' // .pptx
  | 'text/html'
  | 'text/markdown';

/**
 * Supported file extensions (used as fallback or for explicit typing).
 */
export type SupportedInputExtension = '.pdf' | '.docx' | '.pptx' | '.html' | '.htm' | '.md';

/**
 * Supported output formats.
 */
export type SupportedOutputFormat = 'text' | 'markdown';

/**
 * An embedded image extracted from a document (currently PowerPoint slides),
 * passed to a consumer-supplied {@link ImageDescriber}.
 */
export interface EmbeddedImage {
  /** The raw image bytes. */
  data: Buffer;

  /**
   * The image MIME type (e.g. 'image/png', 'image/jpeg'), inferred from the
   * file extension inside the document.
   */
  mimeType: string;

  /** The 1-based slide number the image was found on. */
  slideNumber: number;

  /** The 0-based index of this image within its slide (in document order). */
  imageIndex: number;

  /** The original media filename inside the archive (e.g. 'image3.png'). */
  fileName?: string;
}

/**
 * Provider-agnostic extension point for turning embedded images into text.
 *
 * The document-parsing module never talks to an LLM or holds an API key itself.
 * A consumer (e.g. an app that already has a multi-modal model configured)
 * supplies this function; when provided, embedded images in supported documents
 * are passed to it and the returned text is inlined into the parsed output.
 * This keeps the module portable while letting any tool plug in OpenAI,
 * Anthropic, a local model, OCR, etc.
 *
 * Implementations should be resilient: returning an empty string / null /
 * undefined (or throwing) for a single image simply omits that image's
 * description and does not fail the overall parse.
 *
 * @param image - The embedded image to describe.
 * @returns A textual description of the image, or a falsy value to skip it.
 */
export type ImageDescriber = (
  image: EmbeddedImage
) => Promise<string | null | undefined>;

/**
 * Configuration for the DocumentParsingModule.
 * Extends the core ModuleConfig for common options like logger and debug.
 */
export interface DocumentParsingConfig extends ModuleConfig {
  /**
   * Optional hook for describing images embedded in documents (e.g. charts,
   * screenshots and pictures inside PowerPoint slides). When supplied, the
   * module extracts each embedded image and passes it to this function,
   * inlining the returned text alongside the slide's text content.
   *
   * When omitted, parsing is text-only: no images are processed and no external
   * calls are made. See {@link ImageDescriber}.
   */
  imageDescriber?: ImageDescriber;
}

/**
 * Input specification for the parse method.
 * Currently only supports file paths.
 */
export interface ParseInput {
  /**
   * The path to the input document file.
   */
  filePath: string;
}

/**
 * The result of a successful parsing operation.
 */
export interface ParsingResult {
  /**
   * The extracted content in the requested output format.
   */
  content: string;

  /**
   * Optional metadata about the parsing process or the document.
   * Examples: detected input type, warnings, character count, etc.
   */
  metadata?: {
    detectedInputType?: SupportedInputMimeType | SupportedInputExtension | 'unknown';
    [key: string]: any; // Allow for other arbitrary metadata
  };
}