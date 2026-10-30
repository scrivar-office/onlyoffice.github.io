/*
 * Scrivar Office: tests for the document-rules chunk codec (docProps.js) and
 * the pure helpers in context.js.
 * Run: node --test sdkjs-plugins/content/ai/scripts/scrivar/docProps.test.js
 */
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

const docProps = require(path.join(__dirname, 'docProps.js'));
const context = require(path.join(__dirname, 'context.js'));

const { encodeChunks, decodeChunks, sanitizeLayer, storageKindFor, localKeyFor, CHUNK_MAX } = docProps;

// A deterministic layer whose JSON is ~3,000 chars, with non-ASCII text so
// surrogate pairs and escapes cross chunk boundaries.
function bigLayer() {
	const rules = [];
	for (let i = 0; i < 12; i++) {
		rules.push({
			id: 'r' + i,
			text: 'Rule ' + i + ': write plainly, "quote" carefully, café 😀 naïve — ' + 'x'.repeat(150),
			defaultKey: i === 0 ? 'no_em_dash' : null,
		});
	}
	return {
		voice: 'Warm, direct and concise. Ünïcödé 😀😀😀 ' + 'v'.repeat(200),
		avoid: ['leverage', 'synergy', 'delve', 'utilize'],
		rules,
		referenceIds: [3, 7, 12],
	};
}

test('CHUNK_MAX is 255', () => {
	assert.equal(CHUNK_MAX, 255);
});

test('a ~3,000-char layer round-trips exactly', () => {
	const layer = bigLayer();
	const json = JSON.stringify(Object.assign({ v: 1 }, layer));
	assert.ok(json.length >= 3000, 'fixture is at least 3,000 chars (got ' + json.length + ')');
	const chunks = encodeChunks(layer);
	assert.ok(chunks.length >= 12);
	assert.deepEqual(decodeChunks(chunks), layer);
});

test('every chunk is at most 255 chars, non-empty, and never splits a surrogate pair', () => {
	const chunks = encodeChunks(bigLayer());
	for (const c of chunks) {
		assert.equal(typeof c, 'string');
		assert.ok(c.length > 0 && c.length <= 255, 'chunk length ' + c.length);
		const last = c.charCodeAt(c.length - 1);
		assert.ok(!(last >= 0xd800 && last <= 0xdbff), 'chunk ends with a high surrogate');
		const first = c.charCodeAt(0);
		assert.ok(!(first >= 0xdc00 && first <= 0xdfff), 'chunk starts with a low surrogate');
	}
});

test('the encoded payload carries v:1', () => {
	const chunks = encodeChunks({ voice: 'a', avoid: [], rules: [] });
	assert.equal(chunks.length, 1);
	assert.equal(JSON.parse(chunks[0]).v, 1);
});

test('a missing middle chunk decodes to null', () => {
	const chunks = encodeChunks(bigLayer());
	const mid = Math.floor(chunks.length / 2);
	const withNull = chunks.slice();
	withNull[mid] = null;
	assert.equal(decodeChunks(withNull), null);
	const withEmpty = chunks.slice();
	withEmpty[mid] = '';
	assert.equal(decodeChunks(withEmpty), null);
	const dropped = chunks.slice(0, mid).concat(chunks.slice(mid + 1));
	assert.equal(decodeChunks(dropped), null);
});

test('corrupt JSON decodes to null (never throws)', () => {
	assert.equal(decodeChunks(['{"v":1,"voice":"abc"']), null);
	assert.equal(decodeChunks(['not json at all']), null);
	const chunks = encodeChunks(bigLayer());
	chunks[0] = chunks[0].slice(0, -3) + '###';
	assert.equal(decodeChunks(chunks), null);
});

test('a chunk truncated inside a string literal (still valid JSON) decodes to null via the checksum', () => {
	const chunks = encodeChunks(bigLayer());
	// find a chunk that is all 'x' padding so cutting it keeps the JSON parseable
	const idx = chunks.findIndex((c) => /^x{40,}$/.test(c.slice(0, 40)) && c.length > 60);
	assert.ok(idx > 0, 'fixture has a padding chunk');
	const cut = chunks.slice();
	cut[idx] = cut[idx].slice(0, 30);
	assert.doesNotThrow(() => JSON.parse(cut.join('')), 'the truncated text still parses');
	assert.equal(decodeChunks(cut), null);
});

test('a hand-written layer without a checksum is accepted', () => {
	const raw = JSON.stringify({ v: 1, voice: 'Plain', avoid: [], rules: [] });
	assert.deepEqual(decodeChunks([raw]), { voice: 'Plain', avoid: [], rules: [] });
});

test('wrong version, non-object payloads and empty input decode to null', () => {
	assert.equal(decodeChunks([JSON.stringify({ v: 2, voice: 'x' })]), null);
	assert.equal(decodeChunks([JSON.stringify([1, 2, 3])]), null);
	assert.equal(decodeChunks(['null']), null);
	assert.equal(decodeChunks([]), null);
	assert.equal(decodeChunks(null), null);
	assert.equal(decodeChunks([42]), null);
});

test('encodeChunks(null) is an empty list (clears the document layer)', () => {
	assert.deepEqual(encodeChunks(null), []);
	assert.deepEqual(encodeChunks(undefined), []);
});

test('sanitizeLayer clips to the server limits and drops junk', () => {
	const s = sanitizeLayer({
		voice: 'v'.repeat(2500),
		avoid: ['ok', 7, 'a'.repeat(101), ...Array.from({ length: 60 }, (_, i) => 'w' + i)],
		rules: [
			{ id: 'a', text: 'keep', defaultKey: 'no_em_dash' },
			{ id: 'b' },
			null,
			{ text: 't'.repeat(1200) },
			...Array.from({ length: 60 }, (_, i) => ({ id: 'x' + i, text: 'r' + i })),
		],
		referenceIds: [5, 5, -1, 2.5, '3', 8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
	});
	assert.equal(s.voice.length, 2000);
	assert.equal(s.avoid[0], 'ok');
	assert.ok(s.avoid.every((w) => typeof w === 'string' && w.length <= 100));
	assert.equal(s.avoid.length, 50);
	assert.deepEqual(s.rules[0], { id: 'a', text: 'keep', defaultKey: 'no_em_dash' });
	assert.equal(s.rules[1].text.length, 1000);
	assert.equal(typeof s.rules[1].id, 'string');
	assert.equal(s.rules[1].defaultKey, null);
	assert.equal(s.rules.length, 50);
	assert.deepEqual(s.referenceIds, [5, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
	assert.equal(sanitizeLayer(null), null);
	assert.equal(sanitizeLayer('x'), null);
});

test('storageKindFor picks document properties for OOXML and new files, local storage otherwise', () => {
	assert.equal(storageKindFor({ editorType: 'word', documentTitle: 'Report.docx' }), 'properties');
	assert.equal(storageKindFor({ editorType: 'cell', documentTitle: 'Budget.XLSX' }), 'properties');
	assert.equal(storageKindFor({ editorType: 'slide', documentTitle: 'Deck.pptx' }), 'properties');
	assert.equal(storageKindFor({ editorType: 'word', documentTitle: 'Document1' }), 'properties');
	assert.equal(storageKindFor({ editorType: 'word', documentTitle: 'notes.txt' }), 'local');
	assert.equal(storageKindFor({ editorType: 'cell', documentTitle: 'data.csv' }), 'local');
	assert.equal(storageKindFor({ editorType: 'word', editorSubType: 'pdf', documentTitle: 'a.pdf' }), 'local');
	assert.equal(storageKindFor({ editorType: 'unknown', documentTitle: 'a.docx' }), 'local');
	assert.equal(localKeyFor({ documentTitle: 'notes.txt' }), 'scrivar_doc_rules:notes.txt');
	assert.equal(localKeyFor({ documentId: 'abc' }), 'scrivar_doc_rules:abc');
});

// ---- context.js pure helpers -------------------------------------------------

test('every Scrivar action id is known', () => {
	assert.deepEqual(context.ACTION_IDS.slice().sort(), [
		'annotate_grammar', 'annotate_spelling', 'assistant', 'chat', 'explain', 'explain_link',
		'fix', 'keywords', 'longer', 'rewrite', 'shorter', 'summarize', 'translate',
	]);
});

test('buildContext uses the explicit action, else a mapping from the ActionType', () => {
	const layer = sanitizeLayer({ voice: 'Plain', avoid: ['delve'], rules: [{ id: 'r', text: 'Short sentences.' }], referenceIds: [4] });
	const c1 = context.buildContext({ explicit: 'shorter', actionType: 'TextAnalyze', layer, toolContract: 'X' });
	assert.deepEqual(c1, {
		action: 'shorter',
		documentRules: { voice: 'Plain', avoid: ['delve'], rules: [{ id: 'r', text: 'Short sentences.', defaultKey: null }] },
		referenceIds: [4],
		toolContract: null,
	});
	assert.equal(context.buildContext({ actionType: 'Summarization' }).action, 'summarize');
	assert.equal(context.buildContext({ actionType: 'Translation' }).action, 'translate');
	assert.equal(context.buildContext({ actionType: 'Chat' }).action, 'chat');
	assert.equal(context.buildContext({ actionType: 'TextAnalyze' }).action, 'chat');
	assert.equal(context.buildContext({ actionType: undefined }).action, 'chat');
	assert.equal(context.buildContext({ explicit: 'bogus', actionType: 'Summarization' }).action, 'summarize');
});

test('buildContext sends the tool contract only for an explicitly tagged chat', () => {
	assert.equal(context.buildContext({ explicit: 'chat', actionType: 'Chat', toolContract: 'CONTRACT' }).toolContract, 'CONTRACT');
	// helper sub-requests made by chat tools fall back to "chat" but must not carry the contract
	assert.equal(context.buildContext({ actionType: 'Chat', toolContract: 'CONTRACT' }).toolContract, null);
	assert.equal(context.buildContext({ explicit: 'chat', toolContract: 'y'.repeat(60001) }).toolContract, null);
});

test('buildContext with no or empty document layer sends null rules and no references', () => {
	const c = context.buildContext({ explicit: 'rewrite', layer: null });
	assert.equal(c.documentRules, null);
	assert.deepEqual(c.referenceIds, []);
	const e = context.buildContext({ explicit: 'rewrite', layer: sanitizeLayer({ voice: '', avoid: [], rules: [], referenceIds: [9] }) });
	assert.equal(e.documentRules, null);
	assert.deepEqual(e.referenceIds, [9]);
});

test('findStreamError finds an error frame anywhere in a parsed SSE batch', () => {
	const delta = { choices: [{ delta: { content: 'Hello' } }] };
	const notice = { error: { message: 'Scrivar AI hit a temporary problem.', type: 'scrivar_notice', code: null } };
	assert.equal(context.findStreamError([delta, delta]), null);
	assert.deepEqual(context.findStreamError([delta, notice]), notice.error);
	assert.deepEqual(context.findStreamError([delta, { data: notice }]), notice.error);
	assert.deepEqual(context.findStreamError(notice), notice.error);
	assert.equal(context.findStreamError(null), null);
	assert.equal(context.findStreamError([]), null);
});

test('isScrivarProvider recognises the seeded provider only', () => {
	assert.equal(context.isScrivarProvider({ name: 'Scrivar AI', url: 'http://127.0.0.1:41317/v1' }), true);
	assert.equal(context.isScrivarProvider({ name: 'Other', url: 'http://127.0.0.1:41317/v1' }), true);
	assert.equal(context.isScrivarProvider({ name: 'OpenAI', url: 'https://api.openai.com/v1' }), false);
	assert.equal(context.isScrivarProvider(null), false);
});

test('apiUrl only reaches the launcher settings routes on loopback', () => {
	assert.equal(context.apiUrl('/v1/scrivar/rules'), 'http://127.0.0.1:41317/v1/scrivar/rules');
	assert.equal(context.apiUrl('/v1/scrivar/references/12'), 'http://127.0.0.1:41317/v1/scrivar/references/12');
	assert.throws(() => context.apiUrl('/v1/chat/completions'));
	assert.throws(() => context.apiUrl('https://evil.example/v1/scrivar/rules'));
	assert.throws(() => context.apiUrl('/v1/scrivar/../status'));
});

// ---- context.js runtime (node stubs) ------------------------------------------

test('setAction tags exactly the next request; untagged requests fall back', () => {
	const S = context.Scrivar;
	S.setAction('rewrite');
	assert.equal(S.getContext('TextAnalyze').action, 'rewrite');
	assert.equal(S.getContext('TextAnalyze').action, 'chat', 'the tag is consumed');
	S.setAction('not_an_action');
	assert.equal(S.getContext('Summarization').action, 'summarize');
});

test('adoptToolContract: Scrivar AI carries the contract as context, others keep the system message', () => {
	const S = context.Scrivar;
	const scrivar = { name: 'Scrivar AI', url: 'http://127.0.0.1:41317/v1' };
	assert.equal(S.adoptToolContract(scrivar, 'CONTRACT'), true);
	S.setAction('chat');
	assert.equal(S.getContext('Chat').toolContract, 'CONTRACT');
	assert.equal(S.getContext('Chat').toolContract, null, 'untagged chat helper request');
	assert.equal(S.adoptToolContract(scrivar, 'z'.repeat(60001)), false, 'over the limit stays a system message');
	assert.equal(S.adoptToolContract({ name: 'OpenAI', url: 'https://api.openai.com/v1' }, 'CONTRACT'), false);
	S.setAction('chat');
	assert.equal(S.getContext('Chat').toolContract, null);
});

test('routeErrorToSubscribe: not subscribed -> opens Cloud & AI once (debounced), subscribed -> normal error', async () => {
	const opened = [];
	const realFetch = globalThis.fetch;
	globalThis.open = (url) => opened.push(url);
	let subscribed = false;
	globalThis.fetch = async (url) => {
		assert.equal(url, 'http://127.0.0.1:41317/status');
		return { json: async () => ({ subscribed }) };
	};
	try {
		// a fresh module instance so status cache / debounce start clean
		delete require.cache[require.resolve('./context.js')];
		const S = require('./context.js').Scrivar;
		assert.equal(await S.routeErrorToSubscribe({ message: 'x' }, { interactive: false }), true);
		assert.equal(await S.routeErrorToSubscribe({ message: 'x' }, { interactive: false }), true);
		assert.equal(await S.routeErrorToSubscribe({ message: 'x' }, { interactive: true }), true);
		assert.deepEqual(opened, ['scrivar-office://manage'], 'opened once for a burst');

		delete require.cache[require.resolve('./context.js')];
		subscribed = true;
		const S2 = require('./context.js').Scrivar;
		assert.equal(await S2.routeErrorToSubscribe({ message: 'offline' }, { interactive: true }), false);
		assert.equal(opened.length, 1);

		delete require.cache[require.resolve('./context.js')];
		globalThis.fetch = async () => { throw new Error('launcher down'); };
		const S3 = require('./context.js').Scrivar;
		assert.equal(await S3.routeErrorToSubscribe({ message: 'x' }, { interactive: true }), false);
	} finally {
		globalThis.fetch = realFetch;
		delete globalThis.open;
	}
});

test('api() posts JSON to the loopback route and reports OFFLINE on a network failure', async () => {
	const realFetch = globalThis.fetch;
	const calls = [];
	globalThis.fetch = async (url, init) => {
		calls.push({ url, init });
		return { status: 402, json: async () => ({ success: false, code: 'ADDON_INACTIVE' }) };
	};
	try {
		const S = context.Scrivar;
		const out = await S.api('PUT', '/v1/scrivar/rules', { voice: 'v' });
		assert.deepEqual(out, { success: false, code: 'ADDON_INACTIVE', httpStatus: 402 });
		assert.equal(calls[0].url, 'http://127.0.0.1:41317/v1/scrivar/rules');
		assert.equal(calls[0].init.method, 'PUT');
		assert.equal(calls[0].init.body, '{"voice":"v"}');
		assert.equal(calls[0].init.headers['Content-Type'], 'application/json');
		globalThis.fetch = async () => { throw new Error('ECONNREFUSED'); };
		assert.deepEqual(await S.api('GET', '/v1/scrivar/references'), { success: false, code: 'OFFLINE', httpStatus: 0 });
	} finally {
		globalThis.fetch = realFetch;
	}
});
