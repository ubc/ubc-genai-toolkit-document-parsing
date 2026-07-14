"use strict";
// Main entry point for the Document Parsing module
Object.defineProperty(exports, "__esModule", { value: true });
exports.UnsupportedFileTypeError = exports.ParsingError = exports.DocumentParsingModule = void 0;
// Export the main facade class
var document_parsing_module_1 = require("./document-parsing-module");
Object.defineProperty(exports, "DocumentParsingModule", { enumerable: true, get: function () { return document_parsing_module_1.DocumentParsingModule; } });
// Export custom error types
var error_1 = require("./error");
Object.defineProperty(exports, "ParsingError", { enumerable: true, get: function () { return error_1.ParsingError; } });
Object.defineProperty(exports, "UnsupportedFileTypeError", { enumerable: true, get: function () { return error_1.UnsupportedFileTypeError; } });
//# sourceMappingURL=index.js.map