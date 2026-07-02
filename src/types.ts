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
 * A single fully-parsed slide, delivered incrementally via {@link SlideCallback}
 * as PowerPoint parsing progresses. Lets a consumer process the deck one slide
 * at a time — e.g. embed/store each slide as its own unit — instead of waiting
 * for (and holding) the whole document at once.
 */
export interface ParsedSlide {
  /** The 1-based slide number, in presentation order. */
  slideNumber: number;

  /**
   * Markdown for this slide alone: its `## Slide N` heading, text, any inlined
   * image descriptions, and speaker notes.
   */
  markdown: string;

  /** Plain-text rendering of this slide alone (markdown flattened to text). */
  text: string;

  /** How many embedded images were described on this slide. */
  describedImageCount: number;
}

/**
 * Called once per slide as PowerPoint parsing completes it, in presentation
 * order. Enables per-slide ("chunked") processing and storage rather than
 * handling the whole file as one unit. If it returns a promise, parsing awaits
 * it before continuing to the next slide.
 *
 * @param slide - The fully-parsed slide.
 */
export type SlideCallback = (slide: ParsedSlide) => void | Promise<void>;

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

  /**
   * Optional callback invoked once per slide as a PowerPoint file is parsed
   * (in presentation order). Enables per-slide processing/storage — e.g.
   * embedding each slide independently — so large, image-heavy decks don't have
   * to be handled as a single all-or-nothing unit. `parse()` still returns the
   * full concatenated content as well. See {@link SlideCallback}.
   */
  onSlide?: SlideCallback;

  /**
   * Maximum number of embedded images to describe concurrently (across all
   * slides) when an {@link ImageDescriber} is configured. Higher values finish
   * image-heavy decks much faster; keep it modest to respect provider rate
   * limits. Defaults to 5. Slide output order is always preserved regardless of
   * this value.
   */
  imageConcurrency?: number;

  /**
   * Treat an embedded image as a decorative/template element (and skip
   * describing it) when the *same image* appears on at least this many distinct
   * slides. Recurring icons, logos, dividers and doodles carry no instructional
   * content, and describing them wastes describer calls and can trigger
   * hallucinated descriptions on smaller vision models.
   *
   * Regardless of this value, identical images are always de-duplicated so each
   * unique image is described at most once and the result reused.
   *
   * Defaults to 5. Set to 0 (or any value <= 0) to disable the heuristic.
   */
  decorativeImageSlideThreshold?: number;
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