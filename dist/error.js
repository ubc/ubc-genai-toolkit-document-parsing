"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ParsingError = exports.UnsupportedFileTypeError = void 0;
const ubc_genai_toolkit_core_1 = require("ubc-genai-toolkit-core");
/**
 * Error indicating that the input file type is not supported by the module.
 */
class UnsupportedFileTypeError extends ubc_genai_toolkit_core_1.ToolkitError {
    constructor(detectedType, filePath, details) {
        super(`Unsupported file type ('${detectedType || 'unknown'}') for file: ${filePath}`, 415, // HTTP 415 Unsupported Media Type
        details);
    }
}
exports.UnsupportedFileTypeError = UnsupportedFileTypeError;
/**
 * Generic error during the parsing process (e.g., file access issues, library errors).
 */
class ParsingError extends ubc_genai_toolkit_core_1.ToolkitError {
    constructor(message, filePath, details) {
        super(`Parsing error for file ${filePath}: ${message}`, 500, details);
    }
}
exports.ParsingError = ParsingError;
//# sourceMappingURL=error.js.map