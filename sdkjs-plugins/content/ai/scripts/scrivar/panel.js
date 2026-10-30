/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * The AI panel shell in chat.html: the Chat / Rules / References tab strip
 * and what the two Scrivar tabs share (window.ScrivarPanel):
 *
 *  - bridge(op, payload): document reads/writes, done by the BACKGROUND
 *    plugin page (context.js handlePanelMessage). This window never runs an
 *    editor command itself: a window's command bypasses the shared queue and
 *    would take upstream's single callback slot.
 *      window -> background: Asc.plugin.sendToPlugin('onScrivarPanel', { reqId, op, ... })
 *      background -> window: chatWindow.command('onScrivarPanelReply', { reqId, ok, ... })
 *  - docStore / accountStore: the document's rule layer (with its
 *    referenceIds) and the account's rule layer; edits are applied at once
 *    and saved one at a time (the latest state always wins).
 *  - settings API calls go straight to the launcher's loopback proxy
 *    (AI.Scrivar.api, http://127.0.0.1:41317/v1/scrivar/*).
 *  - el(): DOM builder. Every user or server string goes in as textContent.
 */
(function(window) {
	'use strict';

	var M = window.ScrivarPanelModel;
	var TAB_IDS = ['chat', 'rules', 'references'];
	var BRIDGE_TIMEOUT_MS = 20000;
	var MANAGE_URL = 'scrivar-office://manage';

	function tr(s) {
		try {
			return (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === 'function') ? window.Asc.plugin.tr(s) : s;
		} catch (e) {
			return s;
		}
	}

	/** el('div', { class: 'x', title: '...', onclick: fn }, [child | 'text', ...]) */
	function el(tag, attrs, children) {
		var node = document.createElement(tag);
		attrs = attrs || {};
		Object.keys(attrs).forEach(function(k) {
			var v = attrs[k];
			if (v === null || v === undefined || v === false)
				return;
			if (k === 'class')
				node.className = v;
			else if (k === 'text')
				node.textContent = v;
			else if (k.indexOf('on') === 0 && typeof v === 'function')
				node.addEventListener(k.substring(2), v);
			else if (k === 'value')
				node.value = v;
			else
				node.setAttribute(k, v === true ? '' : String(v));
		});
		(children || []).forEach(function(c) {
			if (c === null || c === undefined || c === false)
				return;
			node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
		});
		return node;
	}

	function clear(node) {
		while (node && node.firstChild)
			node.removeChild(node.firstChild);
	}

	function openManage() {
		try { window.open(MANAGE_URL); } catch (e) {}
	}

	// ---- background bridge -------------------------------------------------------

	var pending = {};
	var seq = 0;

	function bridge(op, payload) {
		return new Promise(function(resolve) {
			var reqId = 'sp' + (++seq) + '.' + Date.now().toString(36);
			var timer = setTimeout(function() {
				delete pending[reqId];
				resolve({ reqId: reqId, ok: false, code: 'TIMEOUT', message: tr('The editor did not answer. Try again in a moment.') });
			}, BRIDGE_TIMEOUT_MS);
			pending[reqId] = function(reply) {
				clearTimeout(timer);
				resolve(reply || { ok: false, message: tr('Something went wrong. Try again.') });
			};
			var msg = { reqId: reqId, op: op };
			Object.keys(payload || {}).forEach(function(k) { msg[k] = payload[k]; });
			try {
				window.Asc.plugin.sendToPlugin('onScrivarPanel', msg);
			} catch (e) {
				delete pending[reqId];
				clearTimeout(timer);
				resolve({ reqId: reqId, ok: false, code: 'NO_PLUGIN', message: tr('Something went wrong. Try again.') });
			}
		});
	}

	function onBridgeReply(reply) {
		var cb = reply && pending[reply.reqId];
		if (cb) {
			delete pending[reply.reqId];
			cb(reply);
		}
	}

	// ---- settings API ----------------------------------------------------------------

	/** -> server JSON + httpStatus, with `state`: 'ok' | 'inactive' | 'offline' | 'error'. */
	async function api(method, path, body) {
		var res;
		try {
			res = await window.AI.Scrivar.api(method, path, body);
		} catch (e) {
			res = { success: false, code: 'BAD_REQUEST', httpStatus: 0 };
		}
		if (res.success)
			res.state = 'ok';
		else if (res.code === 'ADDON_INACTIVE' || res.httpStatus === 402)
			res.state = 'inactive';
		else if (res.code === 'OFFLINE' || res.httpStatus === 0 || res.httpStatus === 503)
			res.state = 'offline';
		else
			res.state = 'error';
		if (!res.success && !res.message) {
			res.message = res.state === 'offline'
				? tr('Scrivar can’t reach the AI service right now. Check your connection and try again.')
				: tr('Something went wrong. Try again.');
		}
		if (res.state === 'inactive')
			setSubscribed(false);
		else if (res.state === 'ok')
			setSubscribed(true);
		return res;
	}

	// ---- subscription -----------------------------------------------------------------

	var subscribed = null; // null = unknown
	var listeners = [];

	function setSubscribed(v) {
		if (subscribed === v)
			return;
		subscribed = v;
		emit('subscription');
	}

	function on(fn) { listeners.push(fn); }
	function emit(what) {
		listeners.forEach(function(fn) {
			try { fn(what); } catch (e) { if (window.console) console.error(e); }
		});
	}

	/** The chat's Subscribe banner (SCRIVAR-SUBSCRIBE): same copy, same action. */
	function subscribeBanner() {
		return el('div', { class: 'sv-subscribe' }, [
			el('div', { class: 'sv-subscribe-msg', text: tr('Cloud & AI is not active. Subscribe to use AI features.') }),
			el('button', { type: 'button', class: 'form-control btn-text-default sv-subscribe-btn', text: tr('Subscribe to Cloud & AI'), onclick: openManage })
		]);
	}

	// ---- stores: apply edits at once, save one at a time ------------------------------

	/**
	 * A layer with a saved copy and a working copy. update(fn) applies fn to
	 * the working copy, notifies, and saves; a save running while more edits
	 * arrive is followed by one more save with the latest state.
	 * saver(layer) -> Promise<{ ok, layer?, message? }>
	 */
	function makeStore(name, loader, saver) {
		var store = {
			name: name,
			status: 'idle',      // idle | loading | ready | error
			loadError: null,     // { message, state }
			saved: null,
			working: null,
			saving: false,
			saveError: null,
			extra: {},
			listeners: []
		};
		var again = false;

		function notify() {
			store.listeners.forEach(function(fn) {
				try { fn(store); } catch (e) { if (window.console) console.error(e); }
			});
		}

		store.on = function(fn) { store.listeners.push(fn); };
		store.notify = notify;

		/** Load, or quietly refresh when already loaded (never over unsaved edits). */
		store.load = async function() {
			if (store.status === 'loading' || store.refreshing)
				return;
			if (store.saving || (store.working && M.isDirty(store.saved, store.working)))
				return; // never replace unsaved edits
			var quiet = store.status === 'ready';
			if (quiet) {
				store.refreshing = true;
			} else {
				store.status = 'loading';
				notify();
			}
			var res = await loader();
			store.refreshing = false;
			if (quiet && (store.saving || M.isDirty(store.saved, store.working)))
				return; // edited while refreshing: keep the edits
			if (quiet && !res.ok && res.state !== 'inactive')
				return; // a failed quiet refresh keeps what is shown
			if (res.ok) {
				store.saved = M.normalizeLayer(res.layer);
				store.working = M.normalizeLayer(res.layer);
				store.extra = res.extra || {};
				store.status = 'ready';
				store.loadError = null;
			} else {
				store.status = 'error';
				store.loadError = { message: res.message, state: res.state || 'error' };
			}
			notify();
		};

		store.update = function(fn) {
			if (store.status !== 'ready')
				return { error: tr('Still loading. Try again in a moment.') };
			var res = fn(store.working);
			if (res.error)
				return res;
			store.working = M.normalizeLayer(res.layer);
			store.saveError = null;
			notify();
			save();
			return res;
		};

		/** Drop edits that could not be saved. */
		store.discard = function() {
			store.working = M.normalizeLayer(store.saved);
			store.saveError = null;
			notify();
		};

		store.retry = function() {
			store.saveError = null;
			save();
		};

		store.dirty = function() {
			return store.status === 'ready' && M.isDirty(store.saved, store.working);
		};

		async function save() {
			if (store.saving) {
				again = true;
				return;
			}
			if (!M.isDirty(store.saved, store.working))
				return;
			store.saving = true;
			again = false;
			notify();
			var sent = store.working;
			var res = await saver(sent);
			store.saving = false;
			if (res.ok) {
				var stored = M.normalizeLayer(res.layer || sent);
				store.saved = stored;
				// No edits while saving: take what the server/file stored (new rule ids).
				if (!M.isDirty(sent, store.working))
					store.working = M.normalizeLayer(stored);
				if (res.extra)
					Object.keys(res.extra).forEach(function(k) { store.extra[k] = res.extra[k]; });
				store.lastSavedAt = Date.now();
			} else {
				store.saveError = res.message || tr('Couldn’t save. Try again.');
				again = false;
			}
			notify();
			if (res.ok && (again || M.isDirty(store.saved, store.working)))
				save();
		}

		return store;
	}

	var docStore = makeStore('document',
		async function() {
			var r = await bridge('docLayer.get');
			return r.ok
				? { ok: true, layer: r.layer, extra: { storage: r.storage, editorType: r.editorType } }
				: { ok: false, message: r.message };
		},
		async function(layer) {
			var r = await bridge('docLayer.save', { layer: layer });
			return r.ok ? { ok: true, layer: r.layer, extra: { storage: r.storage } } : { ok: false, message: r.message };
		});

	var accountStore = makeStore('account',
		async function() {
			var r = await api('GET', '/v1/scrivar/rules');
			return r.state === 'ok' ? { ok: true, layer: r.rules } : { ok: false, message: r.message, state: r.state };
		},
		async function(layer) {
			var r = await api('PUT', '/v1/scrivar/rules', M.accountBody(layer));
			return r.state === 'ok' ? { ok: true, layer: r.rules } : { ok: false, message: r.message };
		});

	/** "Also save to my account": merge `text` into the latest account layer and PUT it. */
	async function addRuleToAccount(text) {
		if (accountStore.status !== 'ready')
			await accountStore.load();
		if (accountStore.status !== 'ready')
			return { error: (accountStore.loadError && accountStore.loadError.message) || tr('Couldn’t save to your account. Try again.') };
		var res = accountStore.update(function(working) {
			var m = M.mergeRuleIntoAccount(working, text);
			return { layer: m.layer, error: m.error };
		});
		return res;
	}

	// ---- tabs -----------------------------------------------------------------------------

	var current = 'chat';
	var tabHandlers = {};

	function registerTab(id, handler) {
		tabHandlers[id] = handler;
	}

	function showTab(id) {
		if (TAB_IDS.indexOf(id) === -1)
			id = 'chat';
		current = id;
		TAB_IDS.forEach(function(t) {
			document.body.classList.toggle('scrivar-view-' + t, t === id);
			var btn = document.querySelector('#scrivar_tabs [data-tab="' + t + '"]');
			if (btn) {
				btn.classList.toggle('selected', t === id);
				btn.setAttribute('aria-selected', t === id ? 'true' : 'false');
				btn.setAttribute('tabindex', t === id ? '0' : '-1');
			}
			var panel = document.getElementById('scrivar_tab_' + t);
			if (panel)
				panel.classList.toggle('hidden', t !== id);
		});
		if (tabHandlers[id] && typeof tabHandlers[id].shown === 'function')
			tabHandlers[id].shown();
		if (id === 'chat') {
			var input = document.getElementById('input_message');
			if (input) setTimeout(function() { try { input.focus(); } catch (e) {} }, 0);
		}
	}

	function setupTabs() {
		var strip = document.getElementById('scrivar_tabs');
		if (!strip)
			return;
		// Both <html> and <body> get the class: the CSS pins the page to the
		// panel height (no document scroll can push the strip out of view).
		document.documentElement.classList.add('scrivar-has-tabs');
		document.body.classList.add('scrivar-has-tabs');
		strip.addEventListener('click', function(e) {
			var btn = e.target.closest ? e.target.closest('[data-tab]') : null;
			if (btn)
				showTab(btn.getAttribute('data-tab'));
		});
		strip.addEventListener('keydown', function(e) {
			if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight')
				return;
			var i = TAB_IDS.indexOf(current) + (e.key === 'ArrowRight' ? 1 : -1);
			i = (i + TAB_IDS.length) % TAB_IDS.length;
			showTab(TAB_IDS[i]);
			var btn = strip.querySelector('[data-tab="' + TAB_IDS[i] + '"]');
			if (btn) btn.focus();
			e.preventDefault();
		});
		showTab('chat');

		// Coming back to the window (e.g. after subscribing in the manage
		// window): refresh what the visible tab shows.
		window.addEventListener('focus', function() {
			if (tabHandlers[current] && typeof tabHandlers[current].refocus === 'function')
				tabHandlers[current].refocus();
		});
	}

	function hookPluginEvents() {
		var plugin = window.Asc && window.Asc.plugin;
		if (!plugin || typeof plugin.attachEvent !== 'function')
			return;
		plugin.attachEvent('onScrivarPanelReply', onBridgeReply);
		// Text sent to the chat (right-click "Ask AI", forced send) brings the Chat tab forward.
		var events = plugin._events;
		if (events && typeof events.onAttachedText === 'function') {
			var original = events.onAttachedText;
			events.onAttachedText = function() {
				if (current !== 'chat') showTab('chat');
				return original.apply(this, arguments);
			};
		}
	}

	function init() {
		hookPluginEvents();
		setupTabs();
		Object.keys(tabHandlers).forEach(function(id) {
			if (typeof tabHandlers[id].mount === 'function')
				tabHandlers[id].mount(document.getElementById('scrivar_tab_' + id));
		});
	}

	window.ScrivarPanel = {
		tr: tr,
		el: el,
		clear: clear,
		bridge: bridge,
		api: api,
		openManage: openManage,
		subscribeBanner: subscribeBanner,
		isSubscribed: function() { return subscribed; },
		on: on,
		docStore: docStore,
		accountStore: accountStore,
		addRuleToAccount: addRuleToAccount,
		registerTab: registerTab,
		showTab: showTab,
		currentTab: function() { return current; },
		_onBridgeReply: onBridgeReply
	};

	if (document.readyState === 'loading')
		document.addEventListener('DOMContentLoaded', init);
	else
		init();
})(window);
