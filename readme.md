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
    -   HTML (`.html`, `.htm`)
    -   Markdown (`.md`)

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

## Error Handling

The module uses the common error types from `ubc-genai-toolkit-core` and defines its own specific errors:

-   **`UnsupportedFileTypeError`**: Thrown if the file type of the input document is not supported.
-   **`ParsingError`**: A generic error for issues during the parsing process, such as file access problems or failures in the underlying parsing libraries.

Always wrap calls to the `parse` method in `try...catch` blocks to handle these potential errors.
