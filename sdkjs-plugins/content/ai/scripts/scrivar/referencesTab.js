/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * The AI panel's References tab: the account's reference library grouped
 * Facts / Writing samples, each with an on/off switch for THIS document
 * (saved as referenceIds in the document layer, through the background
 * page), a meter of the selected characters against the per-request budget,
 * Upload (multipart POST from this window) and "Use this document as
 * reference" (the background page reads the document text; this window
 * POSTs it as JSON). Every user/server string is set as textContent.
 */
(function(window) {
	'use strict';

	var P = window.ScrivarPanel;
	var M = window.ScrivarPanelModel;
	if (!P || !M)
		return;
	var el = P.el;
	var tr = P.tr;
	var SVG_NS = 'http://www.w3.org/2000/svg';
	var ACCEPT = '.docx,.pdf,.txt,.md';

	var root = null;
	var lib = { status: 'idle', list: [], limits: { maxDocs: M.REFS.maxDocs, maxChars: M.REFS.maxChars, requestBudget: M.REFS.requestBudget }, error: null };
	var ui = {
		busy: null,          // { kind: 'facts'|'sample'|'doc', text }
		message: null,       // { text, kind }
		useDocOpen: false
	};

	function icon(d, size) {
		var svg = document.createElementNS(SVG_NS, 'svg');
		svg.setAttribute('viewBox', '0 0 16 16');
		svg.setAttribute('width', String(size || 12));
		svg.setAttribute('height', String(size || 12));
		svg.setAttribute('aria-hidden', 'true');
		svg.setAttribute('class', 'sv-svg');
		var path = document.createElementNS(SVG_NS, 'path');
		path.setAttribute('d', d);
		path.setAttribute('fill', 'none');
		path.setAttribute('stroke', 'currentColor');
		path.setAttribute('stroke-width', '1.4');
		path.setAttribute('stroke-linecap', 'round');
		path.setAttribute('stroke-linejoin', 'round');
		svg.appendChild(path);
		return svg;
	}
	var ICON_UPLOAD = 'M8 11V3M5 6l3-3 3 3M3 11v2h10v-2';
	var ICON_DOC = 'M4 2h5l3 3v9H4zM9 2v3h3M6 8h4M6 10.5h4';
	var ICON_CLOSE = 'M4 4l8 8M12 4l-8 8';

	function setMessage(text, kind) {
		var msg = text ? { text: text, kind: kind || 'error' } : null;
		ui.message = msg;
		if (msg && msg.kind === 'ok') {
			setTimeout(function() {
				if (ui.message === msg) { ui.message = null; render(); }
			}, 5000);
		}
	}

	// ---- data ----------------------------------------------------------------------

	async function loadLibrary() {
		if (lib.status === 'loading')
			return;
		var quiet = lib.status === 'ready';
		if (!quiet) {
			lib.status = 'loading';
			render();
		}
		var res = await P.api('GET', '/v1/scrivar/references');
		if (res.state === 'ok') {
			lib.list = Array.isArray(res.references) ? res.references : [];
			if (res.limits && typeof res.limits === 'object') {
				['maxDocs', 'maxChars', 'requestBudget'].forEach(function(k) {
					if (typeof res.limits[k] === 'number' && res.limits[k] > 0) lib.limits[k] = res.limits[k];
				});
			}
			lib.status = 'ready';
			lib.error = null;
		} else if (!quiet || res.state === 'inactive') {
			lib.status = 'error';
			lib.error = { message: res.message, state: res.state };
		}
		render();
	}

	function ensureLoaded() {
		loadLibrary();
		if (P.docStore.status !== 'loading')
			P.docStore.load();
	}

	function selectedIds() {
		return P.docStore.status === 'ready' ? P.docStore.working.referenceIds : [];
	}

	function toggle(ref, on) {
		var res = P.docStore.update(function(layer) {
			var next = M.normalizeLayer(layer);
			next.referenceIds = M.withReference(next.referenceIds, ref.id, on);
			return { layer: next, error: null };
		});
		if (res && res.error) setMessage(res.error);
		else setMessage(null);
		render();
	}

	function removeMissing(id) {
		toggle({ id: id }, false);
	}

	/** Turn a newly added library document on for this document when it fits. */
	function enableIfFits(ref) {
		if (P.docStore.status !== 'ready' || P.isSubscribed() === false)
			return false;
		var st = M.toggleState(lib.list, selectedIds(), ref, lib.limits);
		if (st.disabled || st.on)
			return false;
		toggle(ref, true);
		return true;
	}

	function libraryFull() {
		return lib.list.length >= lib.limits.maxDocs;
	}

	async function upload(kind, file) {
		if (!file)
			return;
		ui.busy = { kind: kind, text: tr('Uploading %1…').replace('%1', file.name) };
		setMessage(null);
		render();
		var form = new FormData();
		form.append('file', file, file.name);
		form.append('kind', kind);
		var res = await P.api('POST', '/v1/scrivar/references', form);
		ui.busy = null;
		if (res.state === 'ok' && res.reference) {
			lib.list.push(res.reference);
			var on = enableIfFits(res.reference);
			setMessage(on
				? tr('Added “%1” and turned it on for this document.').replace('%1', res.reference.title)
				: tr('Added “%1” to your library.').replace('%1', res.reference.title), 'ok');
			loadLibrary();
		} else {
			setMessage(res.message);
		}
		render();
	}

	async function useThisDocument(kind) {
		ui.useDocOpen = false;
		ui.busy = { kind: 'doc', text: tr('Reading this document…') };
		setMessage(null);
		render();
		var doc = await P.bridge('docText');
		if (!doc.ok) {
			ui.busy = null;
			setMessage(doc.message || tr('Couldn’t read this document. Try again.'));
			render();
			return;
		}
		ui.busy = { kind: 'doc', text: tr('Adding to your library…') };
		render();
		var res = await P.api('POST', '/v1/scrivar/references', {
			kind: kind,
			title: M.titleFromDocument(doc.title),
			text: doc.text
		});
		ui.busy = null;
		if (res.state === 'ok' && res.reference) {
			lib.list.push(res.reference);
			setMessage(tr('Added “%1” to your library. Turn it on in any document.').replace('%1', res.reference.title), 'ok');
			loadLibrary();
		} else {
			setMessage(res.message);
		}
		render();
	}

	function pickFile(kind) {
		var input = el('input', { type: 'file', accept: ACCEPT, class: 'hidden' });
		input.addEventListener('change', function() {
			var f = input.files && input.files[0];
			input.remove();
			upload(kind, f);
		});
		document.body.appendChild(input);
		input.click();
	}

	// ---- rendering --------------------------------------------------------------------

	function meterView() {
		var ids = selectedIds();
		var m = M.meter(lib.list, ids, lib.limits);
		var known = ids.length - M.missingIds(lib.list, ids).length;
		var bar = el('div', { class: 'sv-meter-bar', role: 'meter', 'aria-valuemin': '0',
			'aria-valuemax': String(m.budget), 'aria-valuenow': String(m.used),
			'aria-label': tr('Reference characters used in one request') }, [
			el('div', { class: 'sv-meter-fill sv-meter-' + m.level, style: 'width:' + (m.ratio * 100).toFixed(1) + '%' })
		]);
		return el('div', { class: 'sv-meter' }, [
			el('div', { class: 'sv-meter-head' }, [
				el('span', { class: 'sv-meter-title', text: known
					? (known === 1 ? tr('1 reference on for this document') : tr('%1 references on for this document').replace('%1', known))
					: tr('No references on for this document') }),
				el('span', { class: 'sv-meter-num' }, [
					el('b', { text: M.formatNumber(m.used) }),
					' / ' + M.formatNumber(m.budget)
				])
			]),
			bar,
			el('div', { class: 'sv-meter-foot', text: m.level === 'near'
				? tr('Almost at the limit for one request.')
				: tr('Characters sent with each AI request in this document.') })
		]);
	}

	function switchButton(ref, state, ro) {
		var disabled = ro || state.disabled;
		var title = null;
		if (!ro && state.reason === 'budget')
			title = tr('Turning this on would go over the %1-character limit for one request. Turn another reference off first.').replace('%1', M.formatNumber(lib.limits.requestBudget));
		else if (!ro && state.reason === 'count')
			title = tr('You can turn on up to %1 references in one document.').replace('%1', String(M.REFS.maxSelected));
		var btn = el('button', {
			type: 'button', class: 'sv-switch' + (state.on ? ' on' : '') + (disabled ? ' disabled' : ''),
			role: 'switch', 'aria-checked': state.on ? 'true' : 'false',
			'aria-disabled': disabled ? 'true' : 'false',
			'aria-label': (state.on ? tr('Turn off for this document: ') : tr('Turn on for this document: ')) + ref.title,
			title: title
		}, [el('span', { class: 'sv-switch-knob' })]);
		btn.addEventListener('click', function() {
			if (disabled) return;
			toggle(ref, !state.on);
		});
		return btn;
	}

	function refRow(ref, ro) {
		var state = M.toggleState(lib.list, selectedIds(), ref, lib.limits);
		var meta = [el('span', { text: M.refDetail(ref) })];
		var note = M.truncationNote(ref);
		if (note)
			meta.push(el('span', { class: 'sv-badge', text: tr('Truncated') }));
		var row = el('li', { class: 'sv-ref' + (state.on ? ' on' : '') + (state.disabled ? ' blocked' : '') }, [
			el('div', { class: 'sv-ref-main' }, [
				el('div', { class: 'sv-ref-title', text: ref.title, title: ref.filename || ref.title }),
				el('div', { class: 'sv-ref-meta' }, meta),
				note ? el('div', { class: 'sv-ref-note', text: tr(note) }) : null
			]),
			switchButton(ref, state, ro)
		]);
		if (state.disabled && !ro)
			row.setAttribute('title', state.reason === 'budget'
				? tr('Turning this on would go over the %1-character limit for one request. Turn another reference off first.').replace('%1', M.formatNumber(lib.limits.requestBudget))
				: tr('You can turn on up to %1 references in one document.').replace('%1', String(M.REFS.maxSelected)));
		return row;
	}

	function group(kind, title, emptyText, refs, ro) {
		var full = libraryFull();
		var head = el('div', { class: 'sv-label' }, [
			el('span', { text: tr(title) }),
			refs.length ? el('span', { class: 'sv-label-aside', text: String(refs.length) }) : null,
			ro ? null : el('button', {
				type: 'button', class: 'sv-link sv-upload' + (full ? ' disabled' : ''),
				'aria-disabled': full ? 'true' : 'false',
				title: full ? tr('Your library is full (%1 documents). Delete one in Manage library to add another.').replace('%1', String(lib.limits.maxDocs)) : tr('Upload a .docx, .pdf, .txt or .md file'),
				onclick: function() { if (!full && !ui.busy) pickFile(kind); }
			}, [icon(ICON_UPLOAD), el('span', { text: tr('Upload') })])
		]);
		var children = [head];
		if (ui.busy && ui.busy.kind === kind)
			children.push(el('div', { class: 'sv-busy' }, [el('span', { class: 'sv-spinner', 'aria-hidden': 'true' }), el('span', { text: ui.busy.text })]));
		if (refs.length)
			children.push(el('ul', { class: 'sv-refs' }, refs.map(function(r) { return refRow(r, ro); })));
		else if (!(ui.busy && ui.busy.kind === kind))
			children.push(el('div', { class: 'sv-empty', text: tr(emptyText) }));
		return el('div', { class: 'sv-section' }, children);
	}

	/** Empty library: suggest what to add, each with a hint and an Upload button. */
	function emptyLibrary(ro) {
		var children = [
			el('div', { class: 'sv-label' }, [el('span', { text: tr('Your library is empty') })]),
			el('div', { class: 'sv-empty sv-suggest-intro', text: tr('Add documents the AI can use. Good places to start:') })
		];
		if (ui.busy && (ui.busy.kind === 'facts' || ui.busy.kind === 'sample'))
			children.push(el('div', { class: 'sv-busy' }, [el('span', { class: 'sv-spinner', 'aria-hidden': 'true' }), el('span', { text: ui.busy.text })]));
		children.push(el('ul', { class: 'sv-refs sv-suggest' }, M.EMPTY_SUGGESTIONS.map(function(sg) {
			return el('li', { class: 'sv-ref' }, [
				el('div', { class: 'sv-ref-main' }, [
					el('div', { class: 'sv-ref-title', text: tr(sg.title) }),
					el('div', { class: 'sv-ref-meta' }, [el('span', { text: tr(sg.hint) })])
				]),
				ro ? null : el('button', {
					type: 'button', class: 'form-control btn-text-default sv-suggest-upload',
					'aria-label': tr('Upload') + ' ' + tr(sg.title),
					title: tr('Upload a .docx, .pdf, .txt or .md file'),
					onclick: function() { if (!ui.busy) pickFile(sg.kind); }
				}, [icon(ICON_UPLOAD), el('span', { text: tr('Upload') })])
			]);
		})));
		return el('div', { class: 'sv-section' }, children);
	}

	function missingView(ids, ro) {
		if (!ids.length)
			return null;
		return el('div', { class: 'sv-section' }, [
			el('div', { class: 'sv-label' }, [el('span', { text: tr('Not in your library') })]),
			el('ul', { class: 'sv-refs' }, ids.map(function(id) {
				return el('li', { class: 'sv-ref sv-ref-missing' }, [
					el('div', { class: 'sv-ref-main' }, [
						el('div', { class: 'sv-ref-title', text: tr('Reference %1').replace('%1', String(id)) }),
						el('div', { class: 'sv-ref-meta' }, [el('span', { text: tr('Not in your library') })])
					]),
					ro ? null : el('button', { type: 'button', class: 'sv-icon-btn', title: tr('Remove from this document'),
						'aria-label': tr('Remove from this document'), onclick: function() { removeMissing(id); } }, [icon(ICON_CLOSE)])
				]);
			}))
		]);
	}

	function useDocView(ro) {
		if (ro)
			return null;
		var full = libraryFull();
		if (ui.busy && ui.busy.kind === 'doc')
			return el('div', { class: 'sv-busy sv-usedoc-busy' }, [el('span', { class: 'sv-spinner', 'aria-hidden': 'true' }), el('span', { text: ui.busy.text })]);
		if (ui.useDocOpen) {
			return el('div', { class: 'sv-usedoc-choose' }, [
				el('div', { class: 'sv-usedoc-q', text: tr('Add this document to your library as:') }),
				el('div', { class: 'sv-add-actions' }, [
					el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Cancel'),
						onclick: function() { ui.useDocOpen = false; render(); } }),
					el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Writing sample'),
						onclick: function() { useThisDocument('sample'); } }),
					el('button', { type: 'button', class: 'form-control btn-text-default sv-primary', text: tr('Facts'),
						onclick: function() { useThisDocument('facts'); } })
				])
			]);
		}
		return el('button', {
			type: 'button', class: 'sv-add' + (full || ui.busy ? ' disabled' : ''), 'aria-disabled': full ? 'true' : 'false',
			title: full ? tr('Your library is full (%1 documents). Delete one in Manage library to add another.').replace('%1', String(lib.limits.maxDocs)) : null,
			onclick: function() { if (full || ui.busy) return; ui.useDocOpen = true; setMessage(null); render(); }
		}, [icon(ICON_DOC), el('span', { text: tr('Use this document as reference') })]);
	}

	function footer() {
		return el('div', { class: 'sv-ref-footer' }, [
			el('span', { text: tr('%1 of %2 documents in your library').replace('%1', String(lib.list.length)).replace('%2', String(lib.limits.maxDocs)) }),
			el('button', { type: 'button', class: 'sv-link', text: tr('Manage library'), onclick: P.openManage })
		]);
	}

	function statusLine() {
		var s = P.docStore;
		if (ui.message)
			return el('div', { class: 'sv-status ' + (ui.message.kind === 'ok' ? 'sv-status-ok' : 'sv-status-err'), role: 'status' }, [el('span', { text: ui.message.text })]);
		if (s.saveError) {
			return el('div', { class: 'sv-status sv-status-err', role: 'alert' }, [
				el('span', { text: tr('Couldn’t save: ') + s.saveError }),
				el('button', { type: 'button', class: 'sv-link', text: tr('Retry'), onclick: function() { s.retry(); } }),
				el('button', { type: 'button', class: 'sv-link', text: tr('Undo changes'), onclick: function() { s.discard(); } })
			]);
		}
		if (s.saving || s.dirty())
			return el('div', { class: 'sv-status', role: 'status' }, [el('span', { text: tr('Saving…') })]);
		return null;
	}

	function render() {
		if (!root || P.currentTab() !== 'references')
			return;
		var scrollTop = root.scrollTop;
		P.clear(root);
		var body = el('div', { class: 'sv-body' });
		var inactive = P.isSubscribed() === false;
		if (inactive) {
			body.appendChild(P.subscribeBanner());
			body.appendChild(el('div', { class: 'sv-empty sv-empty-block', text: tr('Your reference library appears here when Cloud & AI is active. Add a company profile, a price list or something you wrote, and turn it on per document so the AI gets the facts and your style right.') }));
		} else if (lib.status === 'idle' || lib.status === 'loading') {
			body.appendChild(el('div', { class: 'sv-loading' }, [el('span', { class: 'sv-spinner', 'aria-hidden': 'true' }), el('span', { text: tr('Loading…') })]));
		} else if (lib.status === 'error') {
			body.appendChild(el('div', { class: 'sv-error-box' }, [
				el('div', { text: (lib.error && lib.error.message) || tr('Something went wrong. Try again.') }),
				el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Try again'), onclick: function() { loadLibrary(); } })
			]));
		} else {
			var docReady = P.docStore.status === 'ready';
			var ro = !docReady;
			if (P.docStore.status === 'error') {
				body.appendChild(el('div', { class: 'sv-error-box' }, [
					el('div', { text: (P.docStore.loadError && P.docStore.loadError.message) || tr('Couldn’t read this document’s settings.') }),
					el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Try again'), onclick: function() { P.docStore.load(); } })
				]));
			}
			body.appendChild(meterView());
			if (lib.list.length === 0) {
				body.appendChild(emptyLibrary(ro));
			} else {
				var g = M.groupLibrary(lib.list);
				body.appendChild(group('facts', 'Facts', 'Add a company profile or a price list so the AI gets names, numbers and details right.', g.facts, ro));
				body.appendChild(group('sample', 'Writing samples', 'Add something you wrote so the AI can match your style.', g.sample, ro));
			}
			var mv = missingView(M.missingIds(lib.list, selectedIds()), ro);
			if (mv) body.appendChild(mv);
			var ud = useDocView(ro);
			if (ud) body.appendChild(ud);
			var st = statusLine();
			if (st) body.appendChild(st);
			body.appendChild(footer());
		}
		root.appendChild(body);
		root.scrollTop = scrollTop;
	}

	P.registerTab('references', {
		mount: function(node) {
			root = node;
			P.docStore.on(function() { render(); });
			P.on(function() { render(); });
		},
		shown: function() {
			ensureLoaded();
			render();
		},
		refocus: function() {
			ensureLoaded();
		}
	});
})(window);
