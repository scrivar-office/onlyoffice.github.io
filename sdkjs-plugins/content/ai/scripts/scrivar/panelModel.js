/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * Pure state helpers for the AI panel's Rules and References tabs
 * (rulesTab.js, referencesTab.js): layer edits, dirty state, the
 * "also save to my account" merge, and the reference meter maths.
 * No DOM, no network; unit-tested under `node --test` (panelModel.test.js).
 */
(function(root) {
	'use strict';

	// Mirrors the server (rulesStore LIMITS, referencesStore constants).
	var LIMITS = { voice: 2000, avoidItems: 50, avoidLen: 100, ruleItems: 50, ruleLen: 1000 };
	var REFS = { maxDocs: 10, maxChars: 150000, requestBudget: 200000, maxSelected: 10 };
	var NEAR_FULL = 0.85;

	// ---- layers ---------------------------------------------------------------

	function str(v) { return typeof v === 'string' ? v : ''; }

	/** A comparable copy: { voice, avoid, rules:[{id,text,defaultKey}], referenceIds }. */
	function normalizeLayer(layer) {
		layer = layer && typeof layer === 'object' ? layer : {};
		return {
			voice: str(layer.voice),
			avoid: Array.isArray(layer.avoid) ? layer.avoid.filter(function(w) { return typeof w === 'string'; }) : [],
			rules: Array.isArray(layer.rules) ? layer.rules.filter(function(r) {
				return r && typeof r.text === 'string';
			}).map(function(r) {
				return { id: str(r.id), text: r.text, defaultKey: (typeof r.defaultKey === 'string' && r.defaultKey) ? r.defaultKey : null };
			}) : [],
			referenceIds: Array.isArray(layer.referenceIds) ? layer.referenceIds.filter(function(n) {
				return typeof n === 'number' && Number.isInteger(n) && n > 0;
			}) : []
		};
	}

	function emptyLayer() {
		return { voice: '', avoid: [], rules: [], referenceIds: [] };
	}

	function layerEquals(a, b) {
		return JSON.stringify(normalizeLayer(a)) === JSON.stringify(normalizeLayer(b));
	}

	/** True when the working copy differs from what was last saved. */
	function isDirty(saved, working) {
		return !layerEquals(saved, working);
	}

	/** True when a document layer holds no rules (voice, avoid or rules). */
	function hasNoRules(layer) {
		var l = normalizeLayer(layer);
		return !l.voice && l.avoid.length === 0 && l.rules.length === 0;
	}

	function cleanText(text) {
		return str(text).replace(/\s+/g, ' ').trim();
	}

	function sameText(a, b) {
		return cleanText(a).toLowerCase() === cleanText(b).toLowerCase();
	}

	var idCounter = 0;
	function newDocRuleId() {
		idCounter++;
		return 'doc-' + Date.now().toString(36) + '-' + idCounter.toString(36);
	}

	/**
	 * Append a rule. -> { layer, error }. `error` is a user-facing sentence and
	 * `layer` is unchanged when it is set. opts.id: false sends no id (account
	 * rules: the server assigns one), otherwise a local doc-* id.
	 */
	function addRule(layer, text, opts) {
		var l = normalizeLayer(layer);
		var t = str(text).trim();
		if (!t)
			return { layer: l, error: 'Type a rule first.' };
		if (t.length > LIMITS.ruleLen)
			return { layer: l, error: 'A rule can be at most ' + formatNumber(LIMITS.ruleLen) + ' characters.' };
		if (l.rules.length >= LIMITS.ruleItems)
			return { layer: l, error: 'You can have at most ' + LIMITS.ruleItems + ' rules.' };
		var rule = { id: (opts && opts.id === false) ? '' : newDocRuleId(), text: t, defaultKey: null };
		l.rules.push(rule);
		return { layer: l, error: null };
	}

	function updateRule(layer, index, text) {
		var l = normalizeLayer(layer);
		var t = str(text).trim();
		if (index < 0 || index >= l.rules.length)
			return { layer: l, error: null };
		if (!t)
			return removeRule(l, index);
		if (t.length > LIMITS.ruleLen)
			return { layer: l, error: 'A rule can be at most ' + formatNumber(LIMITS.ruleLen) + ' characters.' };
		l.rules[index].text = t;
		return { layer: l, error: null };
	}

	function removeRule(layer, index) {
		var l = normalizeLayer(layer);
		if (index >= 0 && index < l.rules.length)
			l.rules.splice(index, 1);
		return { layer: l, error: null };
	}

	function setVoice(layer, voice) {
		var l = normalizeLayer(layer);
		var v = str(voice).trim();
		if (v.length > LIMITS.voice)
			return { layer: l, error: 'The voice can be at most ' + formatNumber(LIMITS.voice) + ' characters.' };
		l.voice = v;
		return { layer: l, error: null };
	}

	/** Add words to the avoid list (comma-separated input allowed; duplicates ignored). */
	function addAvoid(layer, input) {
		var l = normalizeLayer(layer);
		var words = str(input).split(',').map(cleanText).filter(Boolean);
		if (!words.length)
			return { layer: l, error: null };
		for (var i = 0; i < words.length; i++) {
			var w = words[i];
			if (w.length > LIMITS.avoidLen)
				return { layer: normalizeLayer(layer), error: 'Each word or phrase can be at most ' + LIMITS.avoidLen + ' characters.' };
			if (l.avoid.some(function(x) { return sameText(x, w); }))
				continue;
			if (l.avoid.length >= LIMITS.avoidItems)
				return { layer: normalizeLayer(layer), error: 'You can avoid at most ' + LIMITS.avoidItems + ' words and phrases.' };
			l.avoid.push(w);
		}
		return { layer: l, error: null };
	}

	function removeAvoid(layer, index) {
		var l = normalizeLayer(layer);
		if (index >= 0 && index < l.avoid.length)
			l.avoid.splice(index, 1);
		return { layer: l, error: null };
	}

	/**
	 * "Also save to my account": the account layer with `text` appended as a
	 * new account rule (no id: the server assigns one). An equal rule
	 * (case/whitespace-insensitive) is not added twice.
	 * -> { layer: {voice, avoid, rules}, added, error }
	 */
	function mergeRuleIntoAccount(account, text) {
		var a = normalizeLayer(account);
		var out = { voice: a.voice, avoid: a.avoid, rules: a.rules };
		if (a.rules.some(function(r) { return sameText(r.text, text); }))
			return { layer: out, added: false, error: null };
		var res = addRule(a, text, { id: false });
		if (res.error)
			return { layer: out, added: false, error: res.error };
		return { layer: accountBody(res.layer), added: true, error: null };
	}

	/** The PUT /v1/scrivar/rules body: rules without an id go up without one. */
	function accountBody(layer) {
		var l = normalizeLayer(layer);
		return {
			voice: l.voice,
			avoid: l.avoid,
			rules: l.rules.map(function(r) {
				var o = { text: r.text, defaultKey: r.defaultKey };
				if (r.id) o.id = r.id;
				return o;
			})
		};
	}

	// ---- references ------------------------------------------------------------

	function refChars(ref, limits) {
		var cap = (limits && limits.maxChars) || REFS.maxChars;
		var n = ref && typeof ref.charCount === 'number' && ref.charCount > 0 ? ref.charCount : 0;
		return Math.min(n, cap);
	}

	function byId(library) {
		var map = {};
		(library || []).forEach(function(r) { if (r && typeof r.id === 'number') map[r.id] = r; });
		return map;
	}

	/** Characters the selected references add to one request (ids not in the library count 0). */
	function selectedChars(library, ids, limits) {
		var map = byId(library);
		var seen = {};
		var sum = 0;
		(ids || []).forEach(function(id) {
			if (map[id] && !seen[id]) {
				seen[id] = true;
				sum += refChars(map[id], limits);
			}
		});
		return sum;
	}

	/** { used, budget, ratio (0..1, clamped), level: 'ok'|'near'|'over' } */
	function meter(library, ids, limits) {
		var budget = (limits && limits.requestBudget) || REFS.requestBudget;
		var used = selectedChars(library, ids, limits);
		var ratio = budget > 0 ? Math.min(1, used / budget) : 1;
		var level = used > budget ? 'over' : (used >= budget * NEAR_FULL ? 'near' : 'ok');
		return { used: used, budget: budget, ratio: ratio, level: level };
	}

	/**
	 * The toggle for one library document in this document:
	 * { on, disabled, reason: null|'budget'|'count' }. A selected document can
	 * always be turned off; turning one on must keep the request in budget.
	 */
	function toggleState(library, ids, ref, limits) {
		ids = ids || [];
		var on = ids.indexOf(ref.id) !== -1;
		if (on)
			return { on: true, disabled: false, reason: null };
		var budget = (limits && limits.requestBudget) || REFS.requestBudget;
		// Count EVERY id the file stores (incl. "Not in your library" ones from
		// another account): docProps.sanitizeLayer keeps only the first 10
		// valid unique ids, so an 11th would be trimmed on save and snap off.
		var seen = {};
		var count = 0;
		ids.forEach(function(id) {
			if (typeof id === 'number' && Number.isInteger(id) && id > 0 && !seen[id]) {
				seen[id] = true;
				count++;
			}
		});
		if (count >= REFS.maxSelected)
			return { on: false, disabled: true, reason: 'count' };
		if (selectedChars(library, ids, limits) + refChars(ref, limits) > budget)
			return { on: false, disabled: true, reason: 'budget' };
		return { on: false, disabled: false, reason: null };
	}

	/** referenceIds with `id` turned on or off (order kept, no duplicates). */
	function withReference(ids, id, on) {
		var out = (ids || []).filter(function(x) { return x !== id; });
		if (on)
			out.push(id);
		return out;
	}

	/** Ids this document selects that are not in the library (deleted, other account). */
	function missingIds(library, ids) {
		var map = byId(library);
		var seen = {};
		return (ids || []).filter(function(id) {
			if (map[id] || seen[id]) return false;
			seen[id] = true;
			return true;
		});
	}

	function groupLibrary(library) {
		var g = { facts: [], sample: [] };
		(library || []).forEach(function(r) {
			if (r && (r.kind === 'facts' || r.kind === 'sample'))
				g[r.kind].push(r);
		});
		return g;
	}

	// ---- formatting --------------------------------------------------------------

	function formatNumber(n) {
		n = Math.max(0, Math.round(Number(n) || 0));
		return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
	}

	/** "12 pages · 34,210 characters" (pages only when known). */
	function refDetail(ref) {
		var parts = [];
		if (ref && typeof ref.pageCount === 'number' && ref.pageCount > 0)
			parts.push(ref.pageCount + (ref.pageCount === 1 ? ' page' : ' pages'));
		parts.push(formatNumber(ref && ref.charCount) + ' characters');
		return parts.join(' · ');
	}

	/**
	 * The visible note on a truncated reference, or null. The server keeps the
	 * first 50 pages of a long document (and at most 150,000 characters); when
	 * the page count is unknown, say so without a number.
	 */
	function truncationNote(ref) {
		if (!ref || !ref.truncated)
			return null;
		if (typeof ref.pageCount === 'number' && ref.pageCount > 0)
			return 'Only the first 50 pages are used.';
		return 'Only the first part of this document is used.';
	}

	/** Suggestions shown when the reference library is empty. */
	var EMPTY_SUGGESTIONS = [
		{ title: 'Company profile', hint: 'Who you are, services, area, contact details', kind: 'facts' },
		{ title: 'Price list', hint: 'Products or services and what they cost', kind: 'facts' },
		{ title: 'Writing sample', hint: 'A past letter or proposal you liked', kind: 'sample' }
	];

	/** Title for "Use this document as reference": the file name without its extension. */
	function titleFromDocument(name) {
		var t = str(name).split(/[\\/]/).pop().replace(/\.[A-Za-z0-9]{1,5}$/, '').trim();
		return (t || 'This document').substring(0, 200);
	}

	/** The format to "Save as" so document rules stay with the file. */
	function keepRulesExtension(editorType) {
		if (editorType === 'cell') return '.xlsx';
		if (editorType === 'slide') return '.pptx';
		return '.docx';
	}

	var api = {
		LIMITS: LIMITS,
		REFS: REFS,
		normalizeLayer: normalizeLayer,
		emptyLayer: emptyLayer,
		layerEquals: layerEquals,
		isDirty: isDirty,
		hasNoRules: hasNoRules,
		addRule: addRule,
		updateRule: updateRule,
		removeRule: removeRule,
		setVoice: setVoice,
		addAvoid: addAvoid,
		removeAvoid: removeAvoid,
		mergeRuleIntoAccount: mergeRuleIntoAccount,
		accountBody: accountBody,
		refChars: refChars,
		selectedChars: selectedChars,
		meter: meter,
		toggleState: toggleState,
		withReference: withReference,
		missingIds: missingIds,
		groupLibrary: groupLibrary,
		formatNumber: formatNumber,
		refDetail: refDetail,
		truncationNote: truncationNote,
		EMPTY_SUGGESTIONS: EMPTY_SUGGESTIONS,
		titleFromDocument: titleFromDocument,
		keepRulesExtension: keepRulesExtension
	};

	if (typeof module !== 'undefined' && module.exports)
		module.exports = api;
	else
		root.ScrivarPanelModel = api;
})(typeof window !== 'undefined' ? window : globalThis);
