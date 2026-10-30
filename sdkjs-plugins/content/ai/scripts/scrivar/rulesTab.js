/*
 * Scrivar Office modifications to the ONLYOFFICE AI plugin.
 * Copyright (c) 2026 Scrivar Inc. Licensed under the GNU AGPL v3, like the
 * plugin it modifies (see ../../LICENSE.txt).
 *
 * The AI panel's Rules tab: "My account" / "This document" switch, Voice,
 * Avoid chips and numbered rules. Account rules come from the launcher
 * (/v1/scrivar/rules); document rules live inside the file and are read and
 * written by the background plugin page (ScrivarPanel.docStore). Edits save
 * at once. Every user/server string is set as textContent.
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

	function icon(kind) {
		var svg = document.createElementNS(SVG_NS, 'svg');
		svg.setAttribute('viewBox', '0 0 16 16');
		svg.setAttribute('width', '12');
		svg.setAttribute('height', '12');
		svg.setAttribute('aria-hidden', 'true');
		svg.setAttribute('class', 'sv-svg');
		var path = document.createElementNS(SVG_NS, 'path');
		var d = {
			close: 'M4 4l8 8M12 4l-8 8',
			plus: 'M8 3v10M3 8h10',
			info: 'M8 7v4.5M8 4.6v.1'
		}[kind];
		path.setAttribute('d', d);
		path.setAttribute('fill', 'none');
		path.setAttribute('stroke', 'currentColor');
		path.setAttribute('stroke-width', kind === 'info' ? '1.6' : '1.5');
		path.setAttribute('stroke-linecap', 'round');
		svg.appendChild(path);
		if (kind === 'info') {
			var c = document.createElementNS(SVG_NS, 'circle');
			c.setAttribute('cx', '8'); c.setAttribute('cy', '8'); c.setAttribute('r', '6.5');
			c.setAttribute('fill', 'none'); c.setAttribute('stroke', 'currentColor'); c.setAttribute('stroke-width', '1.2');
			svg.insertBefore(c, path);
		}
		return svg;
	}

	var root = null;
	var ui = {
		scope: 'account',     // 'account' | 'document'
		adding: false,
		addDraft: '',
		alsoAccount: false,
		editing: -1,
		editDraft: '',
		voiceDraft: null,     // string while the voice field has focus
		avoidDraft: '',
		message: null,        // { text, kind: 'error'|'ok' } — validation/save feedback
		focusKey: null
	};

	function store() {
		return ui.scope === 'account' ? P.accountStore : P.docStore;
	}

	function readOnly() {
		if (P.isSubscribed() === false)
			return true;
		var s = store();
		return s.status !== 'ready';
	}

	function setMessage(text, kind) {
		var msg = text ? { text: text, kind: kind || 'error' } : null;
		ui.message = msg;
		if (msg && msg.kind === 'ok') {
			setTimeout(function() {
				if (ui.message === msg) { ui.message = null; render(); }
			}, 5000);
		}
	}

	function apply(fn) {
		var res = store().update(fn);
		if (res && res.error) {
			setMessage(res.error, 'error');
			render();
			return false;
		}
		setMessage(null);
		return true;
	}

	// ---- actions ------------------------------------------------------------------

	function commitAdd() {
		var text = ui.addDraft.trim();
		if (!text) {
			setMessage(tr('Type a rule first.'), 'error');
			render();
			return;
		}
		var isDoc = ui.scope === 'document';
		var ok = apply(function(layer) { return M.addRule(layer, text, isDoc ? null : { id: false }); });
		if (!ok)
			return;
		ui.addDraft = '';
		ui.focusKey = 'add';
		if (isDoc && ui.alsoAccount) {
			P.addRuleToAccount(text).then(function(res) {
				if (res && res.error)
					setMessage(tr('Saved to this document, but not to your account: ') + res.error, 'error');
				else
					setMessage(tr('Also saved to your account.'), 'ok');
				render();
			});
		}
		render();
	}

	function commitEdit(index) {
		if (ui.editing !== index)
			return;
		var text = ui.editDraft;
		ui.editing = -1;
		ui.editDraft = '';
		var current = store().working.rules[index];
		if (current && current.text !== text.trim())
			apply(function(layer) { return M.updateRule(layer, index, text); });
		render();
	}

	function commitVoice() {
		if (ui.voiceDraft === null)
			return;
		var v = ui.voiceDraft;
		ui.voiceDraft = null;
		if (store().status === 'ready' && v.trim() !== store().working.voice)
			apply(function(layer) { return M.setVoice(layer, v); });
		render();
	}

	function commitAvoid() {
		var text = ui.avoidDraft;
		if (!text.trim())
			return;
		if (apply(function(layer) { return M.addAvoid(layer, text); }))
			ui.avoidDraft = '';
		ui.focusKey = 'avoid';
		render();
	}

	function switchScope(scope) {
		if (ui.scope === scope)
			return;
		commitVoice();
		ui.scope = scope;
		ui.adding = false;
		ui.addDraft = '';
		ui.editing = -1;
		ui.avoidDraft = '';
		setMessage(null);
		ensureLoaded();
		render();
	}

	function ensureLoaded() {
		var s = store();
		if (s.status === 'idle' || s.status === 'error' || s.status === 'ready')
			s.load();
	}

	// ---- rendering ------------------------------------------------------------------

	function segmented() {
		var docCount = P.docStore.status === 'ready' ? P.docStore.working.rules.length : 0;
		function seg(scope, label, count) {
			var on = ui.scope === scope;
			return el('button', {
				type: 'button', class: 'sv-seg-btn' + (on ? ' selected' : ''), role: 'tab',
				'aria-selected': on ? 'true' : 'false',
				onclick: function() { switchScope(scope); }
			}, [
				el('span', { text: tr(label) }),
				count ? el('span', { class: 'sv-seg-count', text: String(count) }) : null
			]);
		}
		return el('div', { class: 'sv-seg', role: 'tablist', 'aria-label': tr('Rules for') }, [
			seg('account', 'My account', 0),
			seg('document', 'This document', docCount)
		]);
	}

	function sectionLabel(text, aside) {
		return el('div', { class: 'sv-label' }, [el('span', { text: tr(text) }), aside ? el('span', { class: 'sv-label-aside', text: aside }) : null]);
	}

	function voiceSection(layer, ro) {
		var isDoc = ui.scope === 'document';
		var value = ui.voiceDraft !== null ? ui.voiceDraft : layer.voice;
		var ta = el('textarea', {
			class: 'form-control sv-voice' + (isDoc && layer.voice ? ' sv-doc-edge' : ''),
			rows: value ? '2' : '1', spellcheck: 'false', 'data-focus': 'voice',
			maxlength: String(M.LIMITS.voice),
			placeholder: isDoc ? tr('Uses your account voice. Type to set a voice for this document.') : tr('Describe how the AI should sound.'),
			'aria-label': tr('Voice'),
			disabled: ro
		});
		ta.value = value;
		ta.addEventListener('focus', function() { if (ui.voiceDraft === null) ui.voiceDraft = ta.value; });
		ta.addEventListener('input', function() { ui.voiceDraft = ta.value; autosize(ta); });
		ta.addEventListener('blur', function() { setTimeout(commitVoice, 0); });
		ta.addEventListener('keydown', function(e) {
			if (e.key === 'Escape') { ui.voiceDraft = null; ta.value = layer.voice; ta.blur(); }
		});
		return el('div', { class: 'sv-section' }, [sectionLabel('Voice'), ta]);
	}

	function avoidSection(layer, ro) {
		var chips = layer.avoid.map(function(word, i) {
			return el('span', { class: 'sv-chip' }, [
				el('span', { class: 'sv-chip-text', text: word }),
				ro ? null : el('button', {
					type: 'button', class: 'sv-chip-x', title: tr('Remove'), 'aria-label': tr('Remove') + ' ' + word,
					onclick: function() { apply(function(l) { return M.removeAvoid(l, i); }); render(); }
				}, [icon('close')])
			]);
		});
		var input = ro ? null : el('input', {
			type: 'text', class: 'sv-chip-input', 'data-focus': 'avoid', spellcheck: 'false',
			placeholder: layer.avoid.length ? tr('Add…') : tr('Add a word or phrase'),
			'aria-label': tr('Add a word or phrase to avoid'),
			maxlength: String(M.LIMITS.avoidLen)
		});
		if (input) {
			input.value = ui.avoidDraft;
			input.addEventListener('input', function() { ui.avoidDraft = input.value; });
			input.addEventListener('keydown', function(e) {
				if (e.key === 'Enter' || e.key === ',') { e.preventDefault(); commitAvoid(); }
				else if (e.key === 'Backspace' && !input.value && layer.avoid.length) {
					apply(function(l) { return M.removeAvoid(l, l.avoid.length - 1); });
					ui.focusKey = 'avoid';
					render();
				}
			});
			input.addEventListener('blur', function() { if (ui.avoidDraft.trim()) setTimeout(commitAvoid, 0); });
		}
		var box = el('div', { class: 'sv-chips' + (ro ? ' readonly' : '') }, chips.concat([input]));
		box.addEventListener('click', function(e) { if (input && e.target === box) input.focus(); });
		return el('div', { class: 'sv-section' }, [
			sectionLabel('Avoid', layer.avoid.length ? String(layer.avoid.length) : ''),
			box
		]);
	}

	function ruleItem(rule, i, ro, isDoc) {
		var item = el('li', { class: 'sv-rule' + (isDoc ? ' sv-rule-doc' : '') });
		item.appendChild(el('span', { class: 'sv-rule-num', text: String(i + 1) }));
		if (ui.editing === i && !ro) {
			var ta = el('textarea', { class: 'form-control sv-rule-edit', rows: '1', 'data-focus': 'edit', spellcheck: 'false',
				maxlength: String(M.LIMITS.ruleLen), 'aria-label': tr('Edit rule') });
			ta.value = ui.editDraft;
			ta.addEventListener('input', function() { ui.editDraft = ta.value; autosize(ta); });
			ta.addEventListener('keydown', function(e) {
				if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitEdit(i); }
				else if (e.key === 'Escape') { ui.editing = -1; render(); }
			});
			ta.addEventListener('blur', function() { setTimeout(function() { commitEdit(i); }, 0); });
			item.appendChild(ta);
		} else {
			var text = el('div', { class: 'sv-rule-text', text: rule.text });
			if (!ro) {
				text.setAttribute('tabindex', '0');
				text.setAttribute('role', 'button');
				text.setAttribute('title', tr('Click to edit'));
				var startEdit = function() {
					ui.editing = i;
					ui.editDraft = rule.text;
					ui.focusKey = 'edit';
					render();
				};
				text.addEventListener('click', startEdit);
				text.addEventListener('keydown', function(e) { if (e.key === 'Enter') { e.preventDefault(); startEdit(); } });
			}
			item.appendChild(text);
			if (!ro) {
				item.appendChild(el('button', {
					type: 'button', class: 'sv-icon-btn sv-rule-del', title: tr('Delete rule'), 'aria-label': tr('Delete rule') + ' ' + (i + 1),
					onclick: function() { apply(function(l) { return M.removeRule(l, i); }); render(); }
				}, [icon('close')]));
			}
		}
		return item;
	}

	function rulesSection(layer, ro) {
		var isDoc = ui.scope === 'document';
		var list = el('ol', { class: 'sv-rules' }, layer.rules.map(function(r, i) { return ruleItem(r, i, ro, isDoc); }));
		var children = [sectionLabel(isDoc ? 'Rules for this document only' : 'Rules', layer.rules.length ? String(layer.rules.length) : '')];
		if (layer.rules.length)
			children.push(list);
		else
			children.push(el('div', { class: 'sv-empty', text: isDoc
				? tr('No rules for this document yet. They add to your account rules and win when the two disagree.')
				: tr('No rules yet.') }));
		if (!ro)
			children.push(addArea(isDoc));
		return el('div', { class: 'sv-section' }, children);
	}

	function addArea(isDoc) {
		if (!ui.adding) {
			return el('button', {
				type: 'button', class: 'sv-add', 'data-focus': 'add-open',
				onclick: function() { ui.adding = true; ui.focusKey = 'add'; setMessage(null); render(); }
			}, [icon('plus'), el('span', { text: isDoc ? tr('Add a rule for this document') : tr('Add a rule') })]);
		}
		var ta = el('textarea', { class: 'form-control sv-add-input', rows: '2', 'data-focus': 'add', spellcheck: 'false',
			maxlength: String(M.LIMITS.ruleLen),
			placeholder: isDoc ? tr('e.g. Address the client as Mr. and Mrs. Patel.') : tr('e.g. Never start a sentence with “So”.'),
			'aria-label': tr('New rule') });
		ta.value = ui.addDraft;
		ta.addEventListener('input', function() { ui.addDraft = ta.value; autosize(ta); });
		ta.addEventListener('keydown', function(e) {
			if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); commitAdd(); }
			else if (e.key === 'Escape') { ui.adding = false; ui.addDraft = ''; render(); }
		});
		var children = [ta];
		if (isDoc) {
			var cb = el('input', { type: 'checkbox', class: 'form-control', id: 'sv_also_account' });
			cb.checked = ui.alsoAccount;
			cb.addEventListener('change', function() { ui.alsoAccount = cb.checked; });
			children.push(el('label', { class: 'sv-check', 'for': 'sv_also_account' }, [cb, el('span', { text: tr('Also save to my account') })]));
		}
		children.push(el('div', { class: 'sv-add-actions' }, [
			el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Cancel'),
				onclick: function() { ui.adding = false; ui.addDraft = ''; setMessage(null); render(); } }),
			el('button', { type: 'button', class: 'form-control btn-text-default sv-primary', text: tr('Add rule'), onclick: commitAdd })
		]));
		return el('div', { class: 'sv-add-box' + (isDoc ? ' sv-doc-edge' : '') }, children);
	}

	function docNotice() {
		var s = P.docStore;
		var storage = s.extra && s.extra.storage;
		var text = storage === 'local'
			? tr('Save as %1 to keep document rules with this file.').replace('%1', M.keepRulesExtension(s.extra.editorType))
			: tr('Document rules are saved inside this file and travel with it.');
		return el('div', { class: 'sv-notice' + (storage === 'local' ? ' sv-notice-warn' : '') }, [icon('info'), el('span', { text: text })]);
	}

	function statusLine() {
		var s = store();
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
		if (s.lastSavedAt && Date.now() - s.lastSavedAt < 2500)
			return el('div', { class: 'sv-status sv-status-ok', role: 'status' }, [el('span', { text: tr('Saved') })]);
		return null;
	}

	function loadingView() {
		return el('div', { class: 'sv-loading' }, [el('span', { class: 'sv-spinner', 'aria-hidden': 'true' }), el('span', { text: tr('Loading…') })]);
	}

	function errorView(err, retry) {
		return el('div', { class: 'sv-error-box' }, [
			el('div', { text: (err && err.message) || tr('Something went wrong. Try again.') }),
			el('button', { type: 'button', class: 'form-control btn-text-default', text: tr('Try again'), onclick: retry })
		]);
	}

	// A render replaces the DOM; never do it between mousedown and click, or
	// the click (e.g. on "Add a rule" right after typing in a field whose blur
	// saved) would land on a node that no longer exists.
	var pointerDown = false;
	var renderPending = false;
	function render() {
		if (pointerDown) {
			renderPending = true;
			return;
		}
		renderPending = false;
		doRender();
	}

	function doRender() {
		if (!root)
			return;
		var active = document.activeElement;
		if (active && root.contains(active) && active.getAttribute('data-focus'))
			ui.focusKey = ui.focusKey || active.getAttribute('data-focus');
		var scrollTop = root.scrollTop;

		P.clear(root);
		var inactive = P.isSubscribed() === false;
		var s = store();
		var body = el('div', { class: 'sv-body' });

		if (inactive)
			body.appendChild(P.subscribeBanner());
		body.appendChild(segmented());

		if (ui.scope === 'account' && inactive) {
			body.appendChild(el('div', { class: 'sv-empty sv-empty-block', text: tr('Your writing rules appear here when Cloud & AI is active. They shape every AI answer: voice, words to avoid, and rules like “Never use em dashes”.') }));
		} else if (s.status === 'idle' || s.status === 'loading') {
			body.appendChild(loadingView());
		} else if (s.status === 'error') {
			body.appendChild(errorView(s.loadError, function() { s.load(); }));
		} else {
			var ro = readOnly();
			var layer = s.working;
			body.appendChild(el('div', { class: 'sv-scope-hint', text: ui.scope === 'document'
				? tr('Only for this document. These add to your account rules and win when the two disagree.')
				: tr('Used in every document you write with Scrivar AI.') }));
			body.appendChild(voiceSection(layer, ro));
			body.appendChild(avoidSection(layer, ro));
			body.appendChild(rulesSection(layer, ro));
			if (ui.scope === 'document')
				body.appendChild(docNotice());
			var st = statusLine();
			if (st) body.appendChild(st);
		}
		root.appendChild(body);
		root.scrollTop = scrollTop;

		Array.prototype.forEach.call(root.querySelectorAll('textarea'), autosize);
		if (ui.focusKey) {
			var target = root.querySelector('[data-focus="' + ui.focusKey + '"]');
			ui.focusKey = null;
			if (target) {
				target.focus();
				if (typeof target.setSelectionRange === 'function' && target.value !== undefined) {
					var n = target.value.length;
					try { target.setSelectionRange(n, n); } catch (e) {}
				}
			}
		}
	}

	function autosize(ta) {
		if (!ta || !ta.style) return;
		ta.style.height = 'auto';
		ta.style.height = Math.min(ta.scrollHeight + 2, 220) + 'px';
	}

	var savedTimer = null;
	function onStoreChange() {
		if (P.currentTab() !== 'rules')
			return;
		render();
		// let the "Saved" tick fade out
		clearTimeout(savedTimer);
		savedTimer = setTimeout(function() { if (P.currentTab() === 'rules') render(); }, 2600);
	}

	P.registerTab('rules', {
		mount: function(node) {
			root = node;
			root.addEventListener('mousedown', function() { pointerDown = true; });
			document.addEventListener('mouseup', function() {
				if (!pointerDown) return;
				setTimeout(function() {
					pointerDown = false;
					if (renderPending) render();
				}, 0);
			});
			P.accountStore.on(function(s) { if (ui.scope === 'account') onStoreChange(s); });
			P.docStore.on(function(s) { onStoreChange(s); }); // the switch shows the doc rule count
			P.on(function() { if (P.currentTab() === 'rules') render(); });
		},
		shown: function() {
			ensureLoaded();
			if (ui.scope === 'account')
				P.docStore.status === 'idle' && P.docStore.load(); // for the switch's count
			render();
		},
		refocus: function() {
			ensureLoaded();
		}
	});
})(window);
