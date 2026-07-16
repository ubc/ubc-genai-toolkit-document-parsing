const assert = require('node:assert/strict');
const test = require('node:test');

const { DocumentParsingModule } = require('../dist');

function parserResolver() {
	const parser = new DocumentParsingModule();
	return parser._resolvePdfObject.bind(parser);
}

test('resolves a page-local PDF image when the common store does not resolve', async () => {
	const resolvePdfObject = parserResolver();
	const expected = { store: 'page' };
	const page = {
		objs: { get: (_id, callback) => callback(expected) },
		commonObjs: { get: () => undefined },
	};

	await assert.doesNotReject(async () => {
		assert.equal(await resolvePdfObject(page, 'img_p0_1'), expected);
	});
});

test('resolves a globally cached PDF image when the page store does not resolve', async () => {
	const resolvePdfObject = parserResolver();
	const expected = { store: 'common' };
	const page = {
		objs: { get: () => undefined },
		commonObjs: { get: (_id, callback) => callback(expected) },
	};

	await assert.doesNotReject(async () => {
		assert.equal(await resolvePdfObject(page, 'g_d0_img_p20_3'), expected);
	});
});

test('rejects when neither PDF object store can accept the lookup', async () => {
	const resolvePdfObject = parserResolver();
	const page = {
		objs: { get: () => { throw new Error('page store failed'); } },
		commonObjs: { get: () => { throw new Error('common store failed'); } },
	};

	await assert.rejects(
		resolvePdfObject(page, 'missing-image'),
		/common store failed/
	);
});
