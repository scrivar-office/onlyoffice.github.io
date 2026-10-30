/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * AI.Scrivar: the request context the Scrivar AI provider sends with every
 * chat completion, as body field `scrivar`:
 *
 *   { action, documentRules, referenceIds, toolContract }
 *
 * The resident launcher's loopback proxy (http://127.0.0.1:41317) forwards it
 * and the server builds the whole system prompt (persona + the user's rules +
 * reference documents + the chat tool contract). This file holds NO Scrivar
 * API URL and NO credentials: everything goes to the loopback proxy only.
 *
 * Flow:
 *   register.js / annotators call AI.Scrivar.setAction('<id>') right before a
 *   request; engine.js asks AI.Scrivar.getContext(this.actionType) as the
 *   request is built (consuming the tag; untagged requests map from their
 *   ActionType), then awaits AI.Scrivar.attachDocLayer(ctx) to fill in the
 *   document's own rules (stored in the file, see docProps.js).
 *
 * Load order: after docProps.js. Pure helpers are exported for node tests.
 */
(function(root) {
	'use strict';

	var PROXY_ORIGIN = 'http://127.0.0.1:41317';
	var MANAGE_URL = 'scrivar-office://manage';
	var MAX_TOOL_CONTRACT = 60000;
	var DOC_LAYER_MAX_AGE_MS = 3000;
	var DOC_LAYER_WAIT_MS = 1500;
	var SUBSCRIBE_DEBOUNCE_MS = 10000;
	var STATUS_CACHE_MS = 10000;

	var ACTION_IDS = ['chat', 'rewrite', 'longer', 'shorter', 'summarize', 'fix', 'translate',
		'explain', 'explain_link', 'keywords', 'assistant', 'annotate_spelling', 'annotate_grammar'];

	// Untagged requests (other windows, chat tool helpers, external callers).
	var FALLBACK_BY_ACTION_TYPE = {
		Chat: 'chat',
		Summarization: 'summarize',
		Translation: 'translate',
		TextAnalyze: 'chat'
	};

	var DocProps = (typeof module !== 'undefined' && module.exports)
		? require('./docProps.js')
		: root.ScrivarDocProps;

	// ---- pure helpers ---------------------------------------------------------

	function isKnownAction(id) {
		return typeof id === 'string' && ACTION_IDS.indexOf(id) !== -1;
	}

	/** Rules part of a sanitized layer, or null when it holds nothing. */
	function rulesOf(layer) {
		if (!layer)
			return null;
		var voice = layer.voice || '';
		var avoid = layer.avoid || [];
		var rules = layer.rules || [];
		if (!voice && avoid.length === 0 && rules.length === 0)
			return null;
		return { voice: voice, avoid: avoid.slice(), rules: rules.map(function(r) {
			return { id: r.id, text: r.text, defaultKey: r.defaultKey === undefined ? null : r.defaultKey };
		}) };
	}

	/**
	 * { explicit, actionType, layer (sanitized), toolContract } -> the `scrivar` body field.
	 * The tool contract rides only on an explicitly tagged chat turn: chat tool
	 * helpers make their own untagged Chat requests and must not get it.
	 */
	function buildContext(input) {
		input = input || {};
		var explicit = isKnownAction(input.explicit) ? input.explicit : null;
		var action = explicit || FALLBACK_BY_ACTION_TYPE[input.actionType] || 'chat';
		var layer = input.layer || null;
		var contract = null;
		if (explicit === 'chat' && typeof input.toolContract === 'string' && input.toolContract &&
			input.toolContract.length <= MAX_TOOL_CONTRACT)
			contract = input.toolContract;
		return {
			action: action,
			documentRules: rulesOf(layer),
			referenceIds: layer && Array.isArray(layer.referenceIds) ? layer.referenceIds.slice() : [],
			toolContract: contract
		};
	}

	/** The first OpenAI-style error object anywhere in a parsed SSE batch, or null. */
	function findStreamError(chunks) {
		if (!chunks || typeof chunks !== 'object')
			return null;
		var list = Array.isArray(chunks) ? chunks : [chunks];
		for (var i = 0; i < list.length; i++) {
			var c = list[i];
			if (!c || typeof c !== 'object')
				continue;
			if (c.error)
				return c.error;
			if (c.data && typeof c.data === 'object' && c.data.error)
				return c.data.error;
		}
		return null;
	}

	function isScrivarProvider(provider) {
		if (!provider)
			return false;
		if (provider.name === 'Scrivar AI')
			return true;
		return typeof provider.url === 'string' && provider.url.indexOf('127.0.0.1:41317') !== -1;
	}

	/** Loopback URL for a launcher settings route; anything else throws. */
	function apiUrl(path) {
		if (typeof path !== 'string' || path.indexOf('/v1/scrivar/') !== 0 || path.indexOf('..') !== -1)
			throw new Error('AI.Scrivar.api: only /v1/scrivar/* routes are allowed');
		return PROXY_ORIGIN + path;
	}

	// ---- runtime state (plugin page) -----------------------------------------

	var pendingAction = null;
	var toolContract = null;
	var docLayer = null;          // sanitized layer of the open document
	var docLayerLoadedAt = 0;
	var docLayerInflight = null;
	var lastSubscribeAt = 0;

	/** Tag the NEXT request with a Scrivar action id. Also refreshes the document rules in the background. */
	function setAction(id) {
		if (!isKnownAction(id)) {
			if (root.console) console.warn('AI.Scrivar.setAction: unknown action ' + id);
			pendingAction = null;
			return;
		}
		pendingAction = id;
		ensureDocLayer().catch(function() {});
	}

	/** Build the `scrivar` field for a request; consumes the pending tag. Synchronous. */
	function getContext(actionType) {
		var explicit = pendingAction;
		pendingAction = null;
		return buildContext({ explicit: explicit, actionType: actionType, layer: docLayer, toolContract: toolContract });
	}

	/**
	 * Chat: decide where the function-calling contract goes. For Scrivar AI it
	 * rides as context.toolContract (returns true: do NOT also send it as a
	 * system message); otherwise, or when it is over the 60,000-char limit,
	 * the caller keeps it as the system message (returns false).
	 */
	function adoptToolContract(provider, contract) {
		if (isScrivarProvider(provider) && typeof contract === 'string' && contract &&
			contract.length <= MAX_TOOL_CONTRACT) {
			toolContract = contract;
			return true;
		}
		toolContract = null;
		return false;
	}

	function ensureDocLayer() {
		if (docLayerInflight)
			return docLayerInflight;
		if (docLayerLoadedAt && (Date.now() - docLayerLoadedAt) < DOC_LAYER_MAX_AGE_MS)
			return Promise.resolve(docLayer);
		docLayerInflight = loadDocLayer().then(function(layer) {
			docLayerInflight = null;
			return layer;
		}, function() {
			docLayerInflight = null;
			return docLayer;
		});
		return docLayerInflight;
	}

	/** Await a fresh-enough document layer (bounded wait) and put it into ctx. */
	async function attachDocLayer(ctx) {
		if (!ctx)
			return ctx;
		var timer = null;
		try {
			await Promise.race([
				ensureDocLayer(),
				new Promise(function(resolve) { timer = setTimeout(resolve, DOC_LAYER_WAIT_MS); })
			]);
		} catch (e) {
		} finally {
			if (timer) clearTimeout(timer);
		}
		ctx.documentRules = rulesOf(docLayer);
		ctx.referenceIds = docLayer && Array.isArray(docLayer.referenceIds) ? docLayer.referenceIds.slice() : [];
		return ctx;
	}

	/** Read the open document's rule layer (null when it has none). */
	async function loadDocLayer() {
		var raw = null;
		try {
			raw = await DocProps.read();
		} catch (e) {
			raw = null;
		}
		docLayer = DocProps.sanitizeLayer(raw);
		docLayerLoadedAt = Date.now();
		return docLayer;
	}

	/** Store the document's rule layer in the file (null clears it). */
	async function saveDocLayer(layer) {
		var clean = layer === null || layer === undefined ? null : DocProps.sanitizeLayer(layer);
		await DocProps.write(clean);
		docLayer = clean;
		docLayerLoadedAt = Date.now();
	}

	/** 'properties' (stored in the file) or 'local' (this computer only). */
	function docStorageKind() {
		return DocProps.storageKindFor((root.Asc && root.Asc.plugin && root.Asc.plugin.info) || {});
	}

	/**
	 * Call a launcher settings route. Resolves the server JSON ({ success, ... })
	 * plus `httpStatus`; a network failure resolves
	 * { success: false, code: 'OFFLINE', httpStatus: 0 }. `body` may be a
	 * FormData (multipart upload) or any JSON-serialisable value.
	 */
	async function api(method, path, body) {
		var url = apiUrl(path);
		var init = { method: method || 'GET', headers: {} };
		if (body !== undefined && body !== null) {
			if (typeof FormData !== 'undefined' && body instanceof FormData) {
				init.body = body;
			} else {
				init.headers['Content-Type'] = 'application/json';
				init.body = JSON.stringify(body);
			}
		}
		var res;
		try {
			res = await fetch(url, init);
		} catch (e) {
			return { success: false, code: 'OFFLINE', httpStatus: 0 };
		}
		var json = null;
		try {
			json = await res.json();
		} catch (e) {
			json = null;
		}
		if (!json || typeof json !== 'object' || Array.isArray(json))
			json = { success: false, code: 'BAD_RESPONSE' };
		json.httpStatus = res.status;
		return json;
	}

	function openManage() {
		try { root.open(MANAGE_URL); } catch (e) {}
	}

	var statusCache = null;      // { at, promise } — shared by bursts of failures
	function launcherStatus() {
		if (statusCache && (Date.now() - statusCache.at) < STATUS_CACHE_MS)
			return statusCache.promise;
		var promise = (async function() {
			try {
				var ctrl = (typeof AbortController !== 'undefined') ? new AbortController() : null;
				var timer = setTimeout(function() { if (ctrl) ctrl.abort(); }, 2000);
				var r = await fetch(PROXY_ORIGIN + '/status', ctrl ? { signal: ctrl.signal } : {});
				clearTimeout(timer);
				return await r.json();
			} catch (e) {
				return null;
			}
		})();
		statusCache = { at: Date.now(), promise: promise };
		return promise;
	}

	/**
	 * One subscribe flow: when a request fails and Cloud & AI is not active,
	 * open the launcher's Cloud & AI window (the chat banner's Subscribe
	 * action) instead of an editor warning. Resolves true when the error was
	 * handled here; false means "show the normal error".
	 *   interactive (a right-click action): opens the window, at most once
	 *     per 10 s so a double click doesn't open it twice.
	 *   background (spelling/grammar/assistant checks running while typing):
	 *     opens it once per session and otherwise stays quiet, so a lapsed
	 *     subscription never pops a window or a warning on every keystroke.
	 */
	async function routeErrorToSubscribe(err, opts) {
		var interactive = !!(opts && opts.interactive);
		var st = await launcherStatus();
		if (!st || st.subscribed !== false)
			return false;
		var now = Date.now();
		if (interactive) {
			if (!lastSubscribeAt || (now - lastSubscribeAt) >= SUBSCRIBE_DEBOUNCE_MS) {
				lastSubscribeAt = now;
				openManage();
			}
		} else if (!lastSubscribeAt) {
			lastSubscribeAt = now;
			openManage();
		}
		return true;
	}

	var Scrivar = {
		ACTION_IDS: ACTION_IDS,
		setAction: setAction,
		getContext: getContext,
		attachDocLayer: attachDocLayer,
		adoptToolContract: adoptToolContract,
		loadDocLayer: loadDocLayer,
		saveDocLayer: saveDocLayer,
		docStorageKind: docStorageKind,
		api: api,
		openManage: openManage,
		routeErrorToSubscribe: routeErrorToSubscribe,
		isScrivarProvider: isScrivarProvider,
		findStreamError: findStreamError,
		docProps: DocProps
	};

	if (typeof module !== 'undefined' && module.exports) {
		module.exports = {
			ACTION_IDS: ACTION_IDS,
			buildContext: buildContext,
			findStreamError: findStreamError,
			isScrivarProvider: isScrivarProvider,
			apiUrl: apiUrl,
			Scrivar: Scrivar
		};
	} else {
		root.AI = root.AI || {};
		root.AI.Scrivar = Scrivar;
	}
})(typeof window !== 'undefined' ? window : globalThis);
