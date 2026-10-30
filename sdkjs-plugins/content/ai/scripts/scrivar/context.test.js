/*
 * Scrivar Office: tests for AI.Scrivar's per-request action tag and the
 * document-layer cache (context.js), and the serialised command queue
 * (docProps.runCommand).
 * Run: node --test sdkjs-plugins/content/ai/scripts/scrivar/*.test.js
 */
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');

function freshModules() {
	delete require.cache[require.resolve(path.join(__dirname, 'context.js'))];
	delete require.cache[require.resolve(path.join(__dirname, 'docProps.js'))];
	return {
		context: require(path.join(__dirname, 'context.js')),
		docProps: require(path.join(__dirname, 'docProps.js')),
	};
}

function stubEditor(callCommand, title) {
	globalThis.Asc = {
		scope: {},
		plugin: { info: { editorType: 'word', documentTitle: title || 'Report.docx' }, callCommand },
	};
}

test('interleaved setAction/takeAction keep each request its own tag', async () => {
	const S = freshModules().context.Scrivar;

	// models AI.Request._wrapRequest: take the tag synchronously, then await
	// (StartAction Block), then build the context
	async function request(actionType, blockMs) {
		const tag = S.takeAction();
		await new Promise((r) => setTimeout(r, blockMs));
		return S.getContext(actionType, tag).action;
	}

	S.setAction('shorter');
	const a = request('TextAnalyze', 30);        // right-click "Make shorter", blocks first
	S.setAction('annotate_spelling');
	const b = request('Chat', 0);                // background spelling check during the Block
	S.setAction('explain_link');
	const c = request('TextAnalyze', 10);

	assert.deepEqual(await Promise.all([a, b, c]), ['shorter', 'annotate_spelling', 'explain_link']);
	assert.equal(S._state().pending, null);
});

test('an untaken tag expires at the end of the turn (no leak to a later request)', async () => {
	const S = freshModules().context.Scrivar;
	S.setAction('rewrite');                       // e.g. AI.Request.create returned null
	await Promise.resolve();
	assert.equal(S.takeAction(), null);
	assert.equal(S.getContext('TextAnalyze', null).action, 'chat');
});

test('a later setAction replaces an earlier one in the same turn; takeAction clears it', () => {
	const S = freshModules().context.Scrivar;
	S.setAction('rewrite');
	S.setAction('fix');
	assert.equal(S.takeAction(), 'fix');
	assert.equal(S.takeAction(), null);
});

test('a read that gets no answer does not poison the cache (layer kept, timestamp not refreshed)', async () => {
	const { context } = freshModules();
	const S = context.Scrivar;
	const DocProps = S.docProps;

	let reply = DocProps.encodeChunks({ voice: 'Plain', avoid: [], rules: [{ id: 'r1', text: 'Short sentences.' }], referenceIds: [4] });
	stubEditor((func, isClose, isCalc, cb) => {
		const values = reply;
		setTimeout(() => cb(values === undefined ? undefined : JSON.stringify({ count: values.length, values })), 0);
	});
	try {
		await S.prefetchDocLayer();
		const first = S._state();
		assert.equal(first.docLayer.voice, 'Plain');
		assert.ok(first.docLayerLoadedAt > 0);

		// the editor skipped the command (Block long action): callback with undefined
		reply = undefined;
		await new Promise((r) => setTimeout(r, 5));
		await S.prefetchDocLayer();
		const second = S._state();
		assert.deepEqual(second.docLayer, first.docLayer, 'layer kept');
		assert.equal(second.docLayerLoadedAt, first.docLayerLoadedAt, 'timestamp not refreshed');

		const ctx = await S.attachDocLayer(S.getContext('TextAnalyze', 'rewrite'));
		assert.deepEqual(ctx.documentRules, { voice: 'Plain', avoid: [], rules: [{ id: 'r1', text: 'Short sentences.', defaultKey: null }] });
		assert.deepEqual(ctx.referenceIds, [4]);

		// a DEFINITE "no rules" answer does replace the cache
		reply = [];
		await S.prefetchDocLayer();
		assert.equal(S._state().docLayer, null);
	} finally {
		delete globalThis.Asc;
	}
});

test('a read that throws keeps the cache too', async () => {
	const S = freshModules().context.Scrivar;
	const real = S.docProps;
	S._setDocProps({
		read: async () => ({ ok: true, layer: { voice: 'V', avoid: [], rules: [] } }),
		sanitizeLayer: real.sanitizeLayer,
	});
	await S.prefetchDocLayer();
	const t = S._state().docLayerLoadedAt;
	S._setDocProps({ read: async () => { throw new Error('boom'); }, sanitizeLayer: real.sanitizeLayer });
	await new Promise((r) => setTimeout(r, 5));
	await S.prefetchDocLayer();
	assert.equal(S._state().docLayer.voice, 'V');
	assert.equal(S._state().docLayerLoadedAt, t);
});

test('attachDocLayer starts no read of its own', async () => {
	const S = freshModules().context.Scrivar;
	let reads = 0;
	S._setDocProps({ read: async () => { reads++; return { ok: true, layer: null }; }, sanitizeLayer: () => null });
	await S.attachDocLayer(S.getContext('Chat', 'annotate_spelling'));
	await S.attachDocLayer(S.getContext('Chat', 'annotate_grammar'));
	assert.equal(reads, 0);
});

test('saveDocLayer updates the cache directly', async () => {
	const S = freshModules().context.Scrivar;
	const real = S.docProps;
	let written = 'unset';
	S._setDocProps({ write: async (v) => { written = v; return 'properties'; }, sanitizeLayer: real.sanitizeLayer });
	await S.saveDocLayer({ voice: 'Warm', avoid: ['delve'], rules: [], referenceIds: [2] });
	assert.equal(written.voice, 'Warm');
	const ctx = await S.attachDocLayer(S.getContext('TextAnalyze', 'shorter'));
	assert.deepEqual(ctx.documentRules, { voice: 'Warm', avoid: ['delve'], rules: [] });
	assert.deepEqual(ctx.referenceIds, [2]);
});

// ---- the shared command queue ---------------------------------------------------

// A faithful single-slot model of Plugin.executeCommand / onCommandCallback:
// every call overwrites the one callback slot; the first reply consumes it.
function singleSlotEditor(replyDelay) {
	const sent = [];
	let slot = null;
	let inFlight = 0;
	let maxInFlight = 0;
	const callCommand = (func, isClose, isCalc, cb) => {
		const scope = JSON.parse(JSON.stringify(globalThis.Asc.scope));
		sent.push(scope);
		slot = cb;
		inFlight++;
		maxInFlight = Math.max(maxInFlight, inFlight);
		setTimeout(() => {
			inFlight--;
			const s = slot;
			slot = null;
			if (s) s('result:' + scope.text);
		}, replyDelay);
	};
	return { callCommand, sent, stats: () => ({ maxInFlight }) };
}

test('runCommand sends one command at a time, each with the Asc.scope it was queued with', async () => {
	const { docProps } = freshModules();
	const ed = singleSlotEditor(15);
	stubEditor(ed.callCommand);
	try {
		globalThis.Asc.scope.text = 'A';
		const a = docProps.runCommand(function() {});
		globalThis.Asc.scope.text = 'B';           // set before A was actually sent
		const b = docProps.runCommand(function() {});
		globalThis.Asc.scope.text = 'C';
		const c = docProps.runCommand(function() {});
		assert.deepEqual(await Promise.all([a, b, c]), ['result:A', 'result:B', 'result:C']);
		assert.deepEqual(ed.sent.map((s) => s.text), ['A', 'B', 'C']);
		assert.equal(ed.stats().maxInFlight, 1);
		assert.equal(globalThis.Asc.scope.text, 'C', 'the live scope is left as the caller set it');
	} finally {
		delete globalThis.Asc;
	}
});

test('runCommand with a timeout answers undefined, but the next command still waits for the reply', async () => {
	const { docProps } = freshModules();
	const ed = singleSlotEditor(40);
	stubEditor(ed.callCommand);
	try {
		globalThis.Asc.scope.text = 'slow';
		const ours = docProps.runCommand(function() {}, 10);
		globalThis.Asc.scope.text = 'upstream';
		const theirs = docProps.runCommand(function() {});
		assert.equal(await ours, undefined);
		assert.equal(await theirs, 'result:upstream', 'never gets the late reply meant for the timed-out command');
		assert.equal(ed.stats().maxInFlight, 1);
	} finally {
		delete globalThis.Asc;
	}
});
