const assert = require('node:assert/strict');
const test = require('node:test');
const { PNG } = require('pngjs');

const { DocumentParsingModule } = require('../dist');

function makePng(width, height, [red, green, blue, alpha = 255]) {
	const png = new PNG({ width, height });
	for (let offset = 0; offset < png.data.length; offset += 4) {
		png.data[offset] = red;
		png.data[offset + 1] = green;
		png.data[offset + 2] = blue;
		png.data[offset + 3] = alpha;
	}
	return PNG.sync.write(png);
}

function pdfImageEntry(data, bbox, imageIndex) {
	return {
		image: {
			data,
			mimeType: 'image/png',
			source: 'pdf',
			pageNumber: 1,
			imageIndex,
			fileName: `img-${imageIndex}`,
		},
		splitRatio: imageIndex / 10,
		placement: { bbox },
	};
}

test('reconstructs PDF text from explicit spaces, line endings and bullets', () => {
	const parser = new DocumentParsingModule();
	const markdown = parser._pdfTextItemsToMarkdown([
		{ str: 'Learning', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: 'Objectives', hasEOL: true },
		{ str: '•', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: 'Describe', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: 'Watson', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: '‐', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: 'Crick', hasEOL: true },
	]);

	assert.equal(markdown, 'Learning Objectives\n- Describe Watson-Crick');
});

test('keeps semantic dashes separate from word-joining hyphens', () => {
	const parser = new DocumentParsingModule();
	const markdown = parser._pdfTextItemsToMarkdown([
		{ str: 'nucleic acids', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: '–', hasEOL: false },
		{ str: ' ', hasEOL: false },
		{ str: 'solubility', hasEOL: true },
	]);

	assert.equal(markdown, 'nucleic acids - solubility');
});

test('stitches touching horizontal image strips into one PNG', () => {
	const parser = new DocumentParsingModule();
	const red = pdfImageEntry(makePng(8, 2, [255, 0, 0]), [0, 0, 8, 2], 0);
	const blue = pdfImageEntry(makePng(8, 2, [0, 0, 255]), [0, 2, 8, 4], 1);

	const [stitched] = parser._mergePdfImageTiles([red, blue]);
	const png = PNG.sync.read(stitched.image.data);

	assert.equal(png.width, 8);
	assert.equal(png.height, 4);
	assert.match(stitched.image.fileName, /^stitched:/);
	assert.deepEqual([...png.data.subarray(0, 4)], [255, 0, 0, 255]);
	const bottomPixel = ((png.height - 1) * png.width) * 4;
	assert.deepEqual(
		[...png.data.subarray(bottomPixel, bottomPixel + 4)],
		[0, 0, 255, 255]
	);
});

test('does not merge ordinary adjacent figures', () => {
	const parser = new DocumentParsingModule();
	const first = pdfImageEntry(makePng(4, 4, [255, 0, 0]), [0, 0, 4, 4], 0);
	const second = pdfImageEntry(makePng(4, 4, [0, 0, 255]), [4, 0, 8, 4], 1);

	assert.equal(parser._mergePdfImageTiles([first, second]).length, 2);
});
