import { ToolkitError } from 'ubc-genai-toolkit-core';
/**
 * Error indicating that the input file type is not supported by the module.
 */
export declare class UnsupportedFileTypeError extends ToolkitError {
    constructor(detectedType: string | undefined, filePath: string, details?: any);
}
/**
 * Generic error during the parsing process (e.g., file access issues, library errors).
 */
export declare class ParsingError extends ToolkitError {
    constructor(message: string, filePath: string, details?: any);
}
//# sourceMappingURL=error.d.ts.map