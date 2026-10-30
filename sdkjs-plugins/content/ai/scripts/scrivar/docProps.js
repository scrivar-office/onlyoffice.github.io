/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * Document rules storage. A document's rule layer is stored INSIDE the file
 * as custom document properties so it travels with it:
 *
 *   ScrivarAI.count = N
 *   ScrivarAI.1 .. ScrivarAI.N = <=255-char pieces of
 *                                JSON.stringify({ v: 1, h, voice, avoid, rules, referenceIds })
 *
 * h is an FNV-1a checksum of the same JSON without h: a missing, dropped or
 * truncated chunk makes the whole layer read as "no document rules".
 *
 * (255 chars is the custom-property text limit Microsoft Word honours.)
 * There is no property delete in the editor API, so when an edit shrinks the
 * chunk count the stale chunks are overwritten with "" and only the first
 * ScrivarAI.count chunks are ever read.
 *
 * Files that cannot keep custom properties (PDF, .txt, .csv, ...) fall back
 * to the plugin's localStorage, keyed by the document title.
 *
 * Pure codec functions (encode/decode/sanitize/storage choice) are unit-
 * tested under `node --test` (docProps.test.js); the callCommand layer at the
 * bottom only runs inside the editor.
 */
(function(root) {
	'use strict';

	var CHUNK_MAX = 255;
	var PROP_PREFIX = 'ScrivarAI.';
	var PROP_COUNT = 'ScrivarAI.count';
	var LOCAL_PREFIX = 'scrivar_doc_rules:';

	// Mirrors the server's per-layer limits (rulesStore.validateRuleLayer) and
	// the chat context limit (officeChat.validateContext) so a document layer
	// can never make a request fail with BAD_INPUT.
	var LIMITS = { voice: 2000, avoidItems: 50, avoidLen: 100, ruleItems: 50, ruleLen: 1000, referenceIds: 10 };

	var OOXML_EXT = ['docx', 'docm', 'dotx', 'dotm', 'xlsx', 'xlsm', 'xltx', 'xltm',
		'pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm'];

	function isHigh(code) { return code >= 0xd800 && code <= 0xdbff; }

	/** FNV-1a 32-bit over UTF-16 code units, as 8 hex chars. */
	function checksum(str) {
		var h = 0x811c9dc5;
		for (var i = 0; i < str.length; i++) {
			h ^= str.charCodeAt(i);
			h = Math.imul(h, 0x01000193) >>> 0;
		}
		return ('0000000' + h.toString(16)).slice(-8);
	}

	/** Split JSON.stringify({v:1, h, ...obj}) into <=255-char pieces. null -> []. */
	function encodeChunks(obj) {
		if (obj === null || obj === undefined)
			return [];
		var base = { v: 1 };
		var k;
		for (k in obj) {
			if (Object.prototype.hasOwnProperty.call(obj, k) && k !== 'v' && k !== 'h')
				base[k] = obj[k];
		}
		// h = checksum of the JSON without h, so a dropped or truncated chunk
		// that still happens to parse is rejected instead of silently used.
		var payload = { v: 1, h: checksum(JSON.stringify(base)) };
		for (k in base) {
			if (k !== 'v')
				payload[k] = base[k];
		}
		var json = JSON.stringify(payload);
		var out = [];
		var pos = 0;
		while (pos < json.length) {
			var end = Math.min(pos + CHUNK_MAX, json.length);
			// Never split a surrogate pair: a lone surrogate does not survive
			// the file's XML/UTF-8 serialisation.
			if (end < json.length && isHigh(json.charCodeAt(end - 1)))
				end--;
			out.push(json.substring(pos, end));
			pos = end;
		}
		return out;
	}

	/** Reassemble chunk values. null on any gap (null/""/non-string), parse error, bad version or checksum. */
	function decodeChunks(values) {
		if (!Array.isArray(values) || values.length === 0)
			return null;
		var json = '';
		for (var i = 0; i < values.length; i++) {
			var v = values[i];
			if (typeof v !== 'string' || v.length === 0)
				return null;
			json += v;
		}
		var parsed;
		try {
			parsed = JSON.parse(json);
		} catch (e) {
			return null;
		}
		if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed) || parsed.v !== 1)
			return null;
		var base = { v: 1 };
		var out = {};
		for (var k in parsed) {
			if (Object.prototype.hasOwnProperty.call(parsed, k) && k !== 'v' && k !== 'h') {
				base[k] = parsed[k];
				out[k] = parsed[k];
			}
		}
		// A layer written without h (hand-made) is accepted; with h it must match.
		if (parsed.h !== undefined && parsed.h !== checksum(JSON.stringify(base)))
			return null;
		return out;
	}

	var idCounter = 0;
	function newRuleId() {
		idCounter++;
		return 'doc-' + Date.now().toString(36) + '-' + idCounter.toString(36);
	}

	/**
	 * Clip a stored layer to the server limits: { voice, avoid, rules, referenceIds }.
	 * Junk items are dropped, over-long text is truncated. Non-object -> null.
	 */
	function sanitizeLayer(layer) {
		if (!layer || typeof layer !== 'object' || Array.isArray(layer))
			return null;
		var voice = typeof layer.voice === 'string' ? layer.voice.substring(0, LIMITS.voice) : '';

		var avoid = [];
		if (Array.isArray(layer.avoid)) {
			for (var i = 0; i < layer.avoid.length && avoid.length < LIMITS.avoidItems; i++) {
				var w = layer.avoid[i];
				if (typeof w === 'string' && w.length > 0 && w.length <= LIMITS.avoidLen)
					avoid.push(w);
			}
		}

		var rules = [];
		if (Array.isArray(layer.rules)) {
			for (var j = 0; j < layer.rules.length && rules.length < LIMITS.ruleItems; j++) {
				var r = layer.rules[j];
				if (!r || typeof r !== 'object' || typeof r.text !== 'string' || r.text.length === 0)
					continue;
				rules.push({
					id: (typeof r.id === 'string' && r.id) ? r.id : newRuleId(),
					text: r.text.substring(0, LIMITS.ruleLen),
					defaultKey: (typeof r.defaultKey === 'string' && r.defaultKey) ? r.defaultKey : null
				});
			}
		}

		var referenceIds = [];
		if (Array.isArray(layer.referenceIds)) {
			for (var n = 0; n < layer.referenceIds.length && referenceIds.length < LIMITS.referenceIds; n++) {
				var id = layer.referenceIds[n];
				if (typeof id === 'number' && Number.isInteger(id) && id > 0 && referenceIds.indexOf(id) === -1)
					referenceIds.push(id);
			}
		}

		return { voice: voice, avoid: avoid, rules: rules, referenceIds: referenceIds };
	}

	function extensionOf(title) {
		if (typeof title !== 'string')
			return '';
		var base = title.split(/[\\/]/).pop();
		var dot = base.lastIndexOf('.');
		if (dot <= 0 || dot === base.length - 1)
			return '';
		return base.substring(dot + 1).toLowerCase();
	}

	/**
	 * 'properties' when the open file can keep custom document properties
	 * (OOXML, or a new not-yet-saved document, which saves as OOXML), else
	 * 'local'. `info` is Asc.plugin.info.
	 */
	function storageKindFor(info) {
		info = info || {};
		if (info.editorSubType === 'pdf')
			return 'local';
		if (['word', 'slide', 'cell'].indexOf(info.editorType) === -1)
			return 'local';
		var ext = extensionOf(info.documentTitle);
		if (ext === '' || OOXML_EXT.indexOf(ext) !== -1)
			return 'properties';
		return 'local';
	}

	function localKeyFor(info) {
		info = info || {};
		return LOCAL_PREFIX + (info.documentTitle || info.documentId || 'untitled');
	}

	// ---- editor layer (runs only inside the plugin page) ---------------------

	function pluginInfo() {
		return (root.Asc && root.Asc.plugin && root.Asc.plugin.info) || {};
	}

	function callCommandJson(func, timeoutMs) {
		return new Promise(function(resolve) {
			var done = false;
			var timer = setTimeout(function() {
				if (!done) { done = true; resolve(null); }
			}, timeoutMs);
			try {
				root.Asc.plugin.callCommand(func, false, true, function(ret) {
					if (done) return;
					done = true;
					clearTimeout(timer);
					var parsed = null;
					try { parsed = typeof ret === 'string' ? JSON.parse(ret) : ret; } catch (e) { parsed = null; }
					resolve(parsed);
				});
			} catch (e) {
				if (!done) { done = true; clearTimeout(timer); resolve(null); }
			}
		});
	}

	/**
	 * Read ScrivarAI.* through the editor API.
	 * -> { unsupported: true } | { count, values } | null (no answer).
	 * The command body runs in the EDITOR context: no closure variables, only Asc.scope.
	 */
	function readViaApi() {
		root.Asc.scope.scrivarEditor = pluginInfo().editorType;
		return callCommandJson(function() {
			var t = Asc.scope.scrivarEditor;
			var cp = null;
			try {
				if (t === 'word') cp = Api.GetDocument().GetCustomProperties();
				else if (t === 'slide') cp = Api.GetPresentation().GetCustomProperties();
				else if (t === 'cell') cp = Api.GetActiveWorkbook().GetCustomProperties();
			} catch (e) {
				cp = null;
			}
			if (!cp || typeof cp.Get !== 'function')
				return JSON.stringify({ unsupported: true });
			var n = parseInt(cp.Get('ScrivarAI.count'), 10);
			if (!(n > 0))
				return JSON.stringify({ count: 0, values: [] });
			var values = [];
			for (var i = 1; i <= n; i++) {
				var v = cp.Get('ScrivarAI.' + i);
				values.push(v === null || v === undefined ? null : String(v));
			}
			return JSON.stringify({ count: n, values: values });
		}, 3000);
	}

	/** Write chunks + count; stale chunks beyond the new count become "". -> { ok } | { unsupported } | null. */
	function writeViaApi(obj) {
		root.Asc.scope.scrivarEditor = pluginInfo().editorType;
		root.Asc.scope.scrivarChunks = encodeChunks(obj);
		return callCommandJson(function() {
			var t = Asc.scope.scrivarEditor;
			var cp = null;
			try {
				if (t === 'word') cp = Api.GetDocument().GetCustomProperties();
				else if (t === 'slide') cp = Api.GetPresentation().GetCustomProperties();
				else if (t === 'cell') cp = Api.GetActiveWorkbook().GetCustomProperties();
			} catch (e) {
				cp = null;
			}
			if (!cp || typeof cp.Add !== 'function' || typeof cp.Get !== 'function')
				return JSON.stringify({ unsupported: true });
			var chunks = Asc.scope.scrivarChunks || [];
			var old = parseInt(cp.Get('ScrivarAI.count'), 10);
			if (!(old > 0)) old = 0;
			var i;
			for (i = 0; i < chunks.length; i++)
				cp.Add('ScrivarAI.' + (i + 1), chunks[i]);
			for (i = chunks.length + 1; i <= old; i++)
				cp.Add('ScrivarAI.' + i, '');
			cp.Add('ScrivarAI.count', chunks.length);
			return JSON.stringify({ ok: true });
		}, 5000);
	}

	function readLocal(info) {
		try {
			var raw = root.localStorage.getItem(localKeyFor(info));
			if (!raw) return null;
			return decodeChunks([raw]);
		} catch (e) {
			return null;
		}
	}

	function writeLocal(info, obj) {
		var key = localKeyFor(info);
		if (obj === null || obj === undefined) {
			root.localStorage.removeItem(key);
			return;
		}
		root.localStorage.setItem(key, encodeChunks(obj).join(''));
	}

	/** The open document's raw stored layer object, or null. Never throws. */
	async function read() {
		var info = pluginInfo();
		if (storageKindFor(info) === 'local')
			return readLocal(info);
		var res = await readViaApi();
		if (res && res.unsupported)
			return readLocal(info);
		if (!res || !Array.isArray(res.values) || !(res.count > 0))
			return null;
		return decodeChunks(res.values);
	}

	/** Store `obj` (null clears). Throws when the editor refused the write. */
	async function write(obj) {
		var info = pluginInfo();
		if (storageKindFor(info) === 'local') {
			writeLocal(info, obj);
			return 'local';
		}
		var res = await writeViaApi(obj);
		if (res && res.unsupported) {
			writeLocal(info, obj);
			return 'local';
		}
		if (!res || !res.ok)
			throw new Error('Could not save the document rules into this file.');
		return 'properties';
	}

	var api = {
		CHUNK_MAX: CHUNK_MAX,
		PROP_PREFIX: PROP_PREFIX,
		PROP_COUNT: PROP_COUNT,
		LIMITS: LIMITS,
		encodeChunks: encodeChunks,
		decodeChunks: decodeChunks,
		sanitizeLayer: sanitizeLayer,
		storageKindFor: storageKindFor,
		localKeyFor: localKeyFor,
		readViaApi: readViaApi,
		writeViaApi: writeViaApi,
		read: read,
		write: write
	};

	root.ScrivarDocProps = api;
	if (typeof module !== 'undefined' && module.exports)
		module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
