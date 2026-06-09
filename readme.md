# UBC GenAI Toolkit - Document Parsing Module

## Overview

This module provides a standardized interface for parsing text content from various document formats. It follows the Facade pattern, simplifying interactions with underlying parsing libraries for formats like PDF, DOCX, and HTML, while shielding your application from their complexities.

Applications can use this module to extract text from files into either plain text or Markdown format through a consistent API.

## Installation

```bash
npm install ubc-genai-toolkit-document-parsing ubc-genai-toolkit-core
```

## Core Concepts

-   **`DocumentParsingModule`**: The main class and entry point for parsing documents.
-   **`parse(input, outputFormat)`**: The primary method that takes a file path and a desired output format (`'text'` or `'markdown'`) and returns the extracted content.
-   **Supported Formats**: The module currently supports:
    -   PDF (`.pdf`)
    -   Microsoft Word (`.docx`)
    -   Microsoft PowerPoint (`.pptx`)
    -   HTML (`.html`, `.htm`)
    -   Markdown (`.md`)
-   **`imageDescriber`** _(optional)_: A provider-agnostic hook for turning images
    embedded in documents (charts, screenshots, pictures in PowerPoint slides)
    into text. See [Parsing images](#parsing-images-powerpoint).

## Configuration

The `DocumentParsingModule` is configured during instantiation with a `DocumentParsingConfig` object, which extends the `ModuleConfig` from `ubc-genai-toolkit-core`.

```typescript
import { DocumentParsingModule } from 'ubc-genai-toolkit-document-parsing';
import { ConsoleLogger } from 'ubc-genai-toolkit-core';

const config = {
	logger: new ConsoleLogger(),
	debug: true,
};

const docParser = new DocumentParsingModule(config);
```

## Usage Example

The following example demonstrates how to use the module to parse a document from a file path.

```typescript
import { DocumentParsingModule } from 'ubc-genai-toolkit-document-parsing';
import path from 'path';

async function parseDocument(filePath: string) {
	const docParser = new DocumentParsingModule();

	console.log(`--- Parsing: ${path.basename(filePath)} ---`);

	try {
		// Parse to Markdown
		const markdownResult = await docParser.parse({ filePath }, 'markdown');
		console.log('Markdown Output (first 200 chars):');
		console.log(markdownResult.content.substring(0, 200) + '...');

		// Parse to Plain Text
		const textResult = await docParser.parse({ filePath }, 'text');
		console.log('\\nText Output (first 200 chars):');
		console.log(textResult.content.substring(0, 200) + '...');
	} catch (error) {
		console.error(`Failed to parse ${filePath}:`, error);
	}
}

// Example usage:
// const pathToDoc = path.resolve(__dirname, 'data/sample.docx');
// parseDocument(pathToDoc);
```

This example initializes the module and uses it to parse a file into both Markdown and plain text, printing the first 200 characters of each result.

## PowerPoint (`.pptx`)

PowerPoint files are parsed entirely in pure JavaScript (no LibreOffice or other
system dependency). For each slide, in presentation order, the module extracts:

-   slide text (under a `## Slide N` heading),
-   speaker notes (under a `### Notes` heading),
-   and — when an `imageDescriber` is configured — a text description of each
    embedded image.

### Parsing images (PowerPoint)

The module never calls an LLM or holds an API key. Instead it exposes an optional
`imageDescriber` hook. When provided, each embedded image is handed to your
function and the returned text is inlined under its slide. When omitted, parsing
is text-only and makes no external calls.

```typescript
import { DocumentParsingModule, EmbeddedImage } from 'ubc-genai-toolkit-document-parsing';

// `imageDescriber` receives the image bytes + which slide it came from.
// Plug in any vision model — here, the UBC GenAI Toolkit LLM module.
const docParser = new DocumentParsingModule({
	imageDescriber: async (image: EmbeddedImage) => {
		const response = await llm.sendConversation(
			[
				{
					role: 'user',
					content: 'Describe this image from a lecture slide in 1-2 sentences.',
					images: [{ data: image.data.toString('base64'), mimeType: image.mimeType }],
				},
			],
			{ model: 'gpt-5-nano' }
		);
		return response.content;
	},
});

const { content } = await docParser.parse({ filePath: 'lecture.pptx' }, 'markdown');
```

Each image is passed as an `EmbeddedImage` (`{ data: Buffer, mimeType, slideNumber, imageIndex, fileName }`).
The hook is resilient: if it throws or returns an empty value for one image, that
image is simply skipped and the rest of the parse continues.

### Per-slide processing (`onSlide`)

For large, image-heavy decks you may not want to handle the whole file as a
single unit. Provide an `onSlide` callback and the module invokes it once per
slide, in presentation order, as each slide finishes parsing — so you can embed
or store each slide independently instead of waiting for (and holding) the entire
document. `parse()` still returns the full concatenated content as well.

```typescript
const docParser = new DocumentParsingModule({
	imageDescriber,
	onSlide: async (slide) => {
		// slide = { slideNumber, markdown, text, describedImageCount }
		await indexSlide(slide.slideNumber, slide.text); // e.g. embed this slide alone
	},
});

await docParser.parse({ filePath: 'lecture.pptx' }, 'markdown');
```

Each `ParsedSlide` carries that slide's `markdown` and plain `text` (heading +
text + image descriptions + notes) plus `describedImageCount`. If the callback
returns a promise, parsing awaits it before moving to the next slide.

## Error Handling

The module uses the common error types from `ubc-genai-toolkit-core` and defines its own specific errors:

-   **`UnsupportedFileTypeError`**: Thrown if the file type of the input document is not supported.
-   **`ParsingError`**: A generic error for issues during the parsing process, such as file access problems or failures in the underlying parsing libraries.

Always wrap calls to the `parse` method in `try...catch` blocks to handle these potential errors.
