/*
 * Scrivar Office: tests for the AI panel's pure helpers (panelModel.js) and
 * the background page's panel message handler (context.js handlePanelMessage).
 * Run: node --test sdkjs-plugins/content/ai/scripts/scrivar/*.test.js
 */
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const M = require(path.join(__dirname, 'panelModel.js'));

const LIB = [
	{ id: 1, kind: 'facts', title: 'Company profile', charCount: 12000, pageCount: 4, truncated: false },
	{ id: 2, kind: 'facts', title: 'Price list 2026', charCount: 150000, pageCount: 60, truncated: true },
	{ id: 3, kind: 'sample', title: 'Proposal: Patel kitchen', charCount: 30000, pageCount: null, truncated: false },
	{ id: 4, kind: 'sample', title: 'Huge sample', charCount: 400000, pageCount: 200, truncated: true },
];

// ---- meter maths ---------------------------------------------------------------

test('selectedChars sums the selected library docs, capping each at 150,000 and skipping unknown ids', () => {
	assert.equal(M.selectedChars(LIB, []), 0);
	assert.equal(M.selectedChars(LIB, [1, 3]), 42000);
	assert.equal(M.selectedChars(LIB, [4]), 150000, 'a doc counts at most maxChars');
	assert.equal(M.selectedChars(LIB, [1, 99]), 12000, 'an id not in the library counts 0');
	assert.equal(M.selectedChars(LIB, [1, 1]), 12000, 'duplicates count once');
});

test('meter reports used/budget, a clamped ratio and ok/near/over levels', () => {
	assert.deepEqual(M.meter(LIB, [1]), { used: 12000, budget: 200000, ratio: 0.06, level: 'ok' });
	const near = M.meter(LIB, [2, 3]);
	assert.equal(near.used, 180000);
	assert.equal(near.level, 'near');
	const over = M.meter(LIB, [2, 4]);
	assert.equal(over.used, 300000);
	assert.equal(over.ratio, 1);
	assert.equal(over.level, 'over');
	assert.equal(M.meter(LIB, [1], { requestBudget: 10000 }).level, 'over', 'limits from the server win');
});

test('toggleState: a doc that would exceed 200,000 chars is disabled (budget); selected docs can always turn off', () => {
	assert.deepEqual(M.toggleState(LIB, [2], LIB[0]), { on: false, disabled: false, reason: null }); // 162,000
	assert.deepEqual(M.toggleState(LIB, [2, 3], LIB[0]), { on: false, disabled: false, reason: null }); // 192,000
	assert.deepEqual(M.toggleState(LIB, [2, 3], LIB[3]), { on: false, disabled: true, reason: 'budget' });
	assert.deepEqual(M.toggleState(LIB, [2], LIB[1]), { on: true, disabled: false, reason: null });
	assert.deepEqual(M.toggleState(LIB, [2, 4], LIB[1]), { on: true, disabled: false, reason: null }, 'over budget: still can turn off');
	// exactly at the budget is allowed
	assert.equal(M.toggleState([{ id: 7, kind: 'facts', charCount: 50000 }, { id: 8, kind: 'facts', charCount: 150000 }], [8],
		{ id: 7, kind: 'facts', charCount: 50000 }).disabled, false);
});

test('toggleState: at most 10 selected docs in the file', () => {
	const many = Array.from({ length: 11 }, (_, i) => ({ id: i + 1, kind: 'facts', charCount: 10 }));
	const ten = many.slice(0, 10).map((r) => r.id);
	assert.deepEqual(M.toggleState(many, ten, many[10]), { on: false, disabled: true, reason: 'count' });
	assert.equal(M.toggleState(many, ten.slice(0, 9), many[10]).disabled, false);
});

test('toggleState: ids not in this library ("Not in your library") count against the 10-id cap', () => {
	// The file already carries 9 ids from another user's library plus 1 of ours:
	// docProps.sanitizeLayer keeps only 10 ids, so an 11th switch would be
	// trimmed by the background save and snap back off. It must be disabled.
	const mine = [{ id: 1, kind: 'facts', charCount: 10 }, { id: 2, kind: 'facts', charCount: 10 }];
	const stored = [901, 902, 903, 904, 905, 906, 907, 908, 909, 1];
	assert.deepEqual(M.toggleState(mine, stored, mine[1]), { on: false, disabled: true, reason: 'count' });
	// 9 stored (8 foreign + 1 mine) still leaves room for one more.
	assert.deepEqual(M.toggleState(mine, stored.slice(1), mine[1]), { on: false, disabled: false, reason: null });
	// Duplicates and invalid ids are dropped by sanitizeLayer, so they don't count.
	const dupes = [901, 901, 902, 903, 904, 905, 906, 907, 908, 'x', 0, -3, 1.5];
	assert.equal(M.toggleState(mine, dupes, mine[1]).disabled, false);
	// Turning a selected one off is always allowed, even over the cap.
	assert.deepEqual(M.toggleState(mine, stored, mine[0]), { on: true, disabled: false, reason: null });
});

test('withReference / missingIds / groupLibrary', () => {
	assert.deepEqual(M.withReference([1, 3], 2, true), [1, 3, 2]);
	assert.deepEqual(M.withReference([1, 3, 2], 3, false), [1, 2]);
	assert.deepEqual(M.withReference([1], 1, true), [1], 'no duplicates');
	assert.deepEqual(M.missingIds(LIB, [1, 42, 42, 7]), [42, 7]);
	const g = M.groupLibrary(LIB);
	assert.deepEqual(g.facts.map((r) => r.id), [1, 2]);
	assert.deepEqual(g.sample.map((r) => r.id), [3, 4]);
});

// ---- dirty state ----------------------------------------------------------------

test('isDirty compares normalized layers, ignoring junk and key order', () => {
	const saved = { voice: 'Plain', avoid: ['delve'], rules: [{ id: 'a', text: 'Short.', defaultKey: null }], referenceIds: [2] };
	assert.equal(M.isDirty(saved, JSON.parse(JSON.stringify(saved))), false);
	assert.equal(M.isDirty(saved, { referenceIds: [2], rules: [{ text: 'Short.', id: 'a' }], avoid: ['delve'], voice: 'Plain' }), false);
	assert.equal(M.isDirty(saved, M.addRule(saved, 'Formal tone.').layer), true);
	assert.equal(M.isDirty(saved, Object.assign({}, saved, { referenceIds: [] })), true);
	assert.equal(M.isDirty(null, M.emptyLayer()), false, 'no layer == empty layer');
	assert.equal(M.hasNoRules({ referenceIds: [3] }), true);
	assert.equal(M.hasNoRules({ avoid: ['x'] }), false);
});

test('rule and avoid edits respect the server limits and never mutate the input', () => {
	const base = { voice: '', avoid: [], rules: [] };
	const r = M.addRule(base, '  Formal tone, no contractions.  ');
	assert.equal(r.error, null);
	assert.equal(r.layer.rules[0].text, 'Formal tone, no contractions.');
	assert.match(r.layer.rules[0].id, /^doc-/);
	assert.equal(base.rules.length, 0, 'input untouched');
	assert.equal(M.addRule(base, '   ').error, 'Type a rule first.');
	assert.match(M.addRule(base, 'x'.repeat(1001)).error, /1,000 characters/);
	const full = { rules: Array.from({ length: 50 }, (_, i) => ({ id: 'r' + i, text: 'r' + i })) };
	assert.match(M.addRule(full, 'one more').error, /at most 50 rules/);

	const a = M.addAvoid(base, 'delve, Leverage ,  ,delve');
	assert.deepEqual(a.layer.avoid, ['delve', 'Leverage']);
	assert.deepEqual(M.addAvoid(a.layer, 'LEVERAGE').layer.avoid, ['delve', 'Leverage'], 'case-insensitive duplicate');
	assert.match(M.addAvoid(base, 'y'.repeat(101)).error, /100 characters/);
	assert.deepEqual(M.removeAvoid(a.layer, 0).layer.avoid, ['Leverage']);

	const u = M.updateRule(r.layer, 0, 'Formal tone.');
	assert.equal(u.layer.rules[0].text, 'Formal tone.');
	assert.equal(u.layer.rules[0].id, r.layer.rules[0].id, 'id kept on edit');
	assert.equal(M.updateRule(r.layer, 0, '  ').layer.rules.length, 0, 'emptying a rule removes it');
	assert.match(M.setVoice(base, 'v'.repeat(2001)).error, /2,000 characters/);
});

// ---- "also save to my account" merge ------------------------------------------------

test('mergeRuleIntoAccount appends the doc rule without an id and keeps every account rule', () => {
	const account = {
		voice: 'Write clearly.',
		avoid: ['delve'],
		rules: [{ id: 'u-1', text: 'Never use em dashes.', defaultKey: 'no_em_dash' }, { id: 'u-2', text: 'Short paragraphs.', defaultKey: null }],
	};
	const m = M.mergeRuleIntoAccount(account, 'Address the client as Mr. and Mrs. Patel.');
	assert.equal(m.added, true);
	assert.equal(m.error, null);
	assert.deepEqual(m.layer, {
		voice: 'Write clearly.',
		avoid: ['delve'],
		rules: [
			{ id: 'u-1', text: 'Never use em dashes.', defaultKey: 'no_em_dash' },
			{ id: 'u-2', text: 'Short paragraphs.', defaultKey: null },
			{ text: 'Address the client as Mr. and Mrs. Patel.', defaultKey: null },
		],
	});
	assert.equal('referenceIds' in m.layer, false, 'the account body carries no referenceIds');
});

test('mergeRuleIntoAccount skips an equal rule and reports a full account', () => {
	const account = { voice: '', avoid: [], rules: [{ id: 'u-2', text: 'Short  paragraphs.', defaultKey: null }] };
	const same = M.mergeRuleIntoAccount(account, 'short paragraphs.');
	assert.equal(same.added, false);
	assert.equal(same.error, null);
	const full = { rules: Array.from({ length: 50 }, (_, i) => ({ id: 'u' + i, text: 'rule ' + i })) };
	const f = M.mergeRuleIntoAccount(full, 'new rule');
	assert.equal(f.added, false);
	assert.match(f.error, /at most 50 rules/);
});

// ---- formatting --------------------------------------------------------------------

test('formatting helpers', () => {
	assert.equal(M.formatNumber(200000), '200,000');
	assert.equal(M.formatNumber(999), '999');
	assert.equal(M.refDetail(LIB[0]), '4 pages · 12,000 characters');
	assert.equal(M.refDetail(LIB[2]), '30,000 characters');
	assert.equal(M.refDetail({ pageCount: 1, charCount: 10 }), '1 page · 10 characters');
	assert.equal(M.titleFromDocument('Kitchen proposal.docx'), 'Kitchen proposal');
	assert.equal(M.titleFromDocument(''), 'This document');
	assert.equal(M.keepRulesExtension('word'), '.docx');
	assert.equal(M.keepRulesExtension('cell'), '.xlsx');
	assert.equal(M.keepRulesExtension('slide'), '.pptx');
});

// ---- background page: panel messages (context.js) --------------------------------

function freshContext() {
	delete require.cache[require.resolve(path.join(__dirname, 'context.js'))];
	delete require.cache[require.resolve(path.join(__dirname, 'docProps.js'))];
	return require(path.join(__dirname, 'context.js')).Scrivar;
}

function fakeDocProps(over) {
	const real = require(path.join(__dirname, 'docProps.js'));
	return Object.assign({}, real, over);
}

test('panel docLayer.get answers the definite layer and where it is stored', async () => {
	const S = freshContext();
	S._setDocProps(fakeDocProps({
		read: async () => ({ ok: true, layer: { voice: 'Warm', rules: [{ id: 'd1', text: 'Formal.' }], referenceIds: [3] }, storage: 'properties' }),
	}));
	const reply = await S.handlePanelMessage({ reqId: 'r1', op: 'docLayer.get' });
	assert.equal(reply.reqId, 'r1');
	assert.equal(reply.ok, true);
	assert.equal(reply.storage, 'properties');
	assert.deepEqual(reply.layer, { voice: 'Warm', avoid: [], rules: [{ id: 'd1', text: 'Formal.', defaultKey: null }], referenceIds: [3] });
});

test('panel docLayer.get with no answer and no earlier read is NOT ok (the window must not overwrite the file)', async () => {
	const S = freshContext();
	S._setDocProps(fakeDocProps({ read: async () => ({ ok: false }) }));
	const reply = await S.handlePanelMessage({ reqId: 'r2', op: 'docLayer.get' });
	assert.equal(reply.ok, false);
	assert.equal(typeof reply.message, 'string');
});

test('panel docLayer.get with no answer falls back to an earlier definite read', async () => {
	const S = freshContext();
	let answer = { ok: true, layer: { voice: 'Plain' }, storage: 'local' };
	S._setDocProps(fakeDocProps({ read: async () => answer }));
	await S.loadDocLayer();
	answer = { ok: false };
	const reply = await S.handlePanelMessage({ reqId: 'r3', op: 'docLayer.get' });
	assert.equal(reply.ok, true);
	assert.equal(reply.layer.voice, 'Plain');
	assert.equal(reply.storage, 'local');
});

test('panel docLayer.save writes through the background page and refreshes ITS cache (annotators see it at once)', async () => {
	const S = freshContext();
	const writes = [];
	S._setDocProps(fakeDocProps({ write: async (obj) => { writes.push(obj); return 'properties'; } }));
	const layer = { voice: '', avoid: [], rules: [{ id: 'd1', text: 'No contractions.', defaultKey: null }], referenceIds: [1] };
	const reply = await S.handlePanelMessage({ reqId: 'r4', op: 'docLayer.save', layer });
	assert.equal(reply.ok, true);
	assert.equal(writes.length, 1);
	assert.deepEqual(S._state().docLayer.rules[0].text, 'No contractions.');
	const ctx = await S.attachDocLayer(S.getContext('Chat', 'annotate_grammar'));
	assert.deepEqual(ctx.referenceIds, [1]);
});

test('panel docLayer.save failure answers ok:false with the message', async () => {
	const S = freshContext();
	S._setDocProps(fakeDocProps({ write: async () => { throw new Error('Could not save the document rules into this file.'); } }));
	const reply = await S.handlePanelMessage({ reqId: 'r5', op: 'docLayer.save', layer: { voice: 'x' } });
	assert.equal(reply.ok, false);
	assert.equal(reply.message, 'Could not save the document rules into this file.');
});

test('panel docText returns the document text and a title; unknown ops are refused', async () => {
	const S = freshContext();
	globalThis.Asc = { scope: {}, plugin: { info: { editorType: 'word', documentTitle: 'Kitchen proposal.docx' } } };
	S._setDocProps(fakeDocProps({ readDocumentText: async () => ({ text: 'Hello world', truncated: false }) }));
	const reply = await S.handlePanelMessage({ reqId: 'r6', op: 'docText' });
	assert.deepEqual(reply, { reqId: 'r6', ok: true, text: 'Hello world', truncated: false, title: 'Kitchen proposal.docx' });

	S._setDocProps(fakeDocProps({ readDocumentText: async () => ({ unsupported: true }) }));
	const r7 = await S.handlePanelMessage({ reqId: 'r7', op: 'docText' });
	assert.equal(r7.ok, false);
	assert.equal(r7.code, 'UNSUPPORTED');

	const bad = await S.handlePanelMessage({ reqId: 'r8', op: 'callCommand' });
	assert.equal(bad.ok, false);
	delete globalThis.Asc;
});

// ---- fix round 1: truncation note + empty-library suggestions ---------------------

test('truncationNote: visible sentence for truncated references (pages known vs unknown)', () => {
	assert.equal(M.truncationNote(LIB[0]), null, 'not truncated');
	assert.equal(M.truncationNote(LIB[1]), 'Only the first 50 pages are used.');
	assert.equal(M.truncationNote({ truncated: true, pageCount: null, charCount: 150000 }), 'Only the first part of this document is used.');
	assert.equal(M.truncationNote({ truncated: true, charCount: 150000 }), 'Only the first part of this document is used.');
});

test('EMPTY_SUGGESTIONS: Company profile, Price list, Writing sample with the exact hints', () => {
	assert.deepEqual(M.EMPTY_SUGGESTIONS, [
		{ title: 'Company profile', hint: 'Who you are, services, area, contact details', kind: 'facts' },
		{ title: 'Price list', hint: 'Products or services and what they cost', kind: 'facts' },
		{ title: 'Writing sample', hint: 'A past letter or proposal you liked', kind: 'sample' },
	]);
});
