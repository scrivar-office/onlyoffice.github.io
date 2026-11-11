/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

(function(exports, undefined)
{
	let Editor = {};
	
	Editor.callMethod = async function(name, args)
	{
		return new Promise(resolve => (function(){
			Asc.plugin.executeMethod(name, args || [], function(returnValue){
				resolve(returnValue);
			});
		})());
	};

	Editor.callCommand = async function(func)
	{
		// SCRIVAR-RULES: one command at a time through the queue shared with the
		// Scrivar document-rules reads (the plugin API has ONE callCommand
		// callback slot; overlapping commands swap or lose their results).
		if (window.ScrivarDocProps && window.ScrivarDocProps.runCommand)
			return await window.ScrivarDocProps.runCommand(func);
		return new Promise(resolve => (function(){
			Asc.plugin.callCommand(func, false, true, function(returnValue){
				resolve(returnValue);
			});
		})());
	};

	// SCRIVAR-AGENT: run model-written macro code in the editor.
	// The editor's command sandbox (AscCommon.safePluginEval) stubs out any
	// nested eval and the Function constructor, so the code cannot be handed
	// over in Asc.scope and eval'ed there. It is placed inline in the command
	// text instead. Asc.plugin.callCommand builds that text as
	//   "...; (" + func.toString() + ")();"
	// so toString() closes that wrapper early, runs the code inside try/catch,
	// and reopens a harmless declaration for the trailing ")();". The command's
	// completion value is then the macro's last expression (a `var` statement
	// has no completion value), or the error object from the catch.
	// -> { error: string } | { result: any } | {} (ran, nothing returned)
	Editor.runMacro = async function(code)
	{
		code = String(code === undefined || code === null ? "" : code);
		try {
			// Syntax check here, where Function still works: the editor drops
			// invalid JS without telling the plugin.
			new Function("\"use strict\";\n" + code);
		} catch (e) {
			return { error: e.name + ": " + e.message };
		}
		let command = {
			toString : function() {
				return "function(){})();\n" +
					"try {\n" + code + "\n} catch (__scrivarErr) { ({ onlyoffice_id_error_message: __scrivarErr.name + \": \" + __scrivarErr.message }) }\n" +
					"var __scrivarTail = (function(){ return 0; }";
			}
		};
		let ret = await Editor.callCommand(command);
		if (ret && typeof ret === "object" && ret.onlyoffice_id_error_message)
			return { error: ret.onlyoffice_id_error_message };
		if (ret === undefined || ret === null)
			return {};
		return { result: ret };
	};

	// SCRIVAR-CONTEXT: a compact snapshot of the open file for the chat agent.
	// -> string ("" when there is nothing useful or the editor has none)
	Editor.getContextSnapshot = async function()
	{
		let type = Editor.getType();
		let ret = null;
		try {
			if (type === "cell") {
				ret = await Editor.callCommand(function() {
					try {
						var ws = Api.GetActiveSheet();
						var out = { sheet: ws.GetName(), rows: [] };
						var sel = Api.GetSelection();
						if (sel) out.selection = sel.GetAddress(false, false, "xlA1", false);
						var used = ws.GetUsedRange();
						if (!used) return JSON.stringify(out);
						out.used = used.GetAddress(false, false, "xlA1", false);
						// Rows/columns come from the ADDRESS: in this engine build
						// GetRows().Count and GetCols().Count return each other's value.
						var m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(out.used);
						if (!m) return JSON.stringify(out);
						var colNum = function(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; };
						var r0 = parseInt(m[2], 10), c0 = colNum(m[1]);
						var nRows = (m[4] ? parseInt(m[4], 10) : r0) - r0 + 1, nCols = (m[3] ? colNum(m[3]) : c0) - c0 + 1;
						out.firstRow = r0; out.lastRow = r0 + nRows - 1;
						var maxR = Math.min(nRows, 25), maxC = Math.min(nCols, 12);
						var colName = function(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
						out.lastCol = colName(c0 + nCols - 1);
						for (var r = 0; r < maxR; r++) {
							var cells = [];
							for (var c = 0; c < maxC; c++) {
								var cell = ws.GetRange(colName(c0 + c) + (r0 + r));
								var f = cell.GetFormula();
								var v = cell.GetValue();
								var text = (f && String(f).charAt(0) === "=") ? (f + " -> " + v) : String(v === null || v === undefined ? "" : v);
								if (text.length > 40) text = text.substring(0, 40) + "...";
								cells.push(text);
							}
							out.rows.push({ n: r0 + r, cells: cells });
						}
						out.firstCol = colName(c0);
						out.moreRows = nRows - maxR; out.moreCols = nCols - maxC;
						return JSON.stringify(out);
					} catch (e) { return ""; }
				});
				if (!ret) return "";
				let o = JSON.parse(ret);
				let text = "Sheet: " + o.sheet + "\n";
				if (o.selection) text += "Selected: " + o.selection + "\n";
				if (!o.used || !o.rows.length) return text + "The sheet is empty.";
				text += "Data range: " + o.used + " (rows " + o.firstRow + " to " + o.lastRow + ", columns " + o.firstCol + " to " + o.lastCol + ")\n";
				text += "Cells, row by row, starting at column " + o.firstCol + ", separated by | (formulas shown as formula -> value):\n";
				o.rows.forEach(function(row) { text += row.n + ": " + row.cells.join(" | ") + "\n"; });
				if (o.moreRows > 0) text += "... " + o.moreRows + " more rows not shown\n";
				if (o.moreCols > 0) text += "... " + o.moreCols + " more columns not shown\n";
				return text.length > 6000 ? text.substring(0, 6000) + "\n... (truncated)" : text;
			}
			if (type === "slide") {
				ret = await Editor.callCommand(function() {
					try {
						var p = Api.GetPresentation();
						var n = p.GetSlidesCount();
						var out = { count: n, current: p.GetCurSlideIndex() + 1, slides: [] };
						for (var i = 0; i < n && i < 40; i++) {
							var title = "", body = 0;
							var drawings = p.GetSlideByIndex(i).GetAllDrawings();
							for (var j = 0; j < drawings.length; j++) {
								var ph = drawings[j].GetPlaceholder ? drawings[j].GetPlaceholder() : null;
								var dc = drawings[j].GetDocContent ? drawings[j].GetDocContent() : null;
								if (!dc) continue;
								var t = "";
								for (var k = 0; k < dc.GetElementsCount(); k++) { var el = dc.GetElement(k); if (el && el.GetText) t += el.GetText() + " "; }
								t = t.replace(/\s+/g, " ").trim();
								var kind = ph ? ph.GetType() : "";
								if (!title && (kind === "title" || kind === "ctrTitle")) title = t;
								else if (t) body++;
							}
							out.slides.push({ title: title.substring(0, 80), body: body });
						}
						return JSON.stringify(out);
					} catch (e) { return ""; }
				});
				if (!ret) return "";
				let o = JSON.parse(ret);
				let text = "Presentation: " + o.count + " slide(s). Current slide: " + o.current + ".\n";
				o.slides.forEach(function(sl, i) { text += "Slide " + (i + 1) + ": " + (sl.title ? "\"" + sl.title + "\"" : "(no title)") + (sl.body ? ", " + sl.body + " text block(s)" : ", no body text") + "\n"; });
				return text;
			}
		} catch (e) {
			return "";
		}
		return "";
	};

	// SCRIVAR-PDF: the text of the open PDF, page by page, for the chat.
	// Uses the PDF scripting API (page selection -> selected text), then
	// collapses the selection again. -> string ("" when there is no text layer)
	Editor.getPdfText = async function(maxChars)
	{
		if (Editor.getType() !== "pdf")
			return "";
		Asc.scope.scrivarPdfMax = maxChars || 40000;
		let ret = null;
		try {
			ret = await Editor.callCommand(function() {
				try {
					var doc = Api.GetDocument();
					var n = doc.GetPagesCount();
					var max = Asc.scope.scrivarPdfMax;
					var out = { pages: n, read: 0, text: "", cut: false };
					var far = { x: 100000, y: 100000 }, origin = { x: 0, y: 0 };
					var last = null;
					for (var i = 0; i < n; i++) {
						var page = doc.GetPage(i);
						if (!page) continue;
						page.SetSelection(origin, far);
						last = page;
						var t = page.GetSelectedText() || "";
						t = String(t).replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
						if (t) out.text += "[Page " + (i + 1) + "]\n" + t + "\n\n";
						out.read = i + 1;
						if (out.text.length >= max) { out.text = out.text.substring(0, max); out.cut = true; break; }
					}
					if (last) last.SetSelection(origin, origin);
					return JSON.stringify(out);
				} catch (e) { return ""; }
			});
		} catch (e) {
			return "";
		}
		if (!ret) return "";
		let o;
		try { o = JSON.parse(ret); } catch (e) { return ""; }
		if (!o.text)
			return "This PDF has " + o.pages + " page(s) but no readable text layer (it may be a scan). Tell the user that, and that you can only work with text they select or paste.";
		let head = "PDF with " + o.pages + " page(s).";
		if (o.cut || o.read < o.pages)
			head += " Only the first part is included below (through page " + o.read + "); say so if the answer may depend on later pages.";
		return head + "\n\n" + o.text;
	};

	Editor.pause = async function(msec)
	{
		return new Promise(resolve => (function(){
			setTimeout(function(){
				resolve();				
			}, msec);
		})());
	};

	Editor.getType = function() {
		if (Asc.plugin.info.editorSubType === "pdf")
			return "pdf";
		return window.Asc.plugin.info.editorType;
	};

	exports.Asc = exports.Asc || {};
	exports.Asc.Editor = Editor;

	function Library() {
		this.version = 0;
	}

	exports.Asc.PluginsMD = {
		latex: function(md) {
			// Inline: $...$
			md.inline.ruler.after("escape", "latex_inline", function(state, silent) {
				let start = state.pos;
				if (state.src[start] !== '$')
					return false;
				if (state.src[start + 1] === '$')
					return false;

				let content = "";
				let end = start + 1;
				while ((end = state.src.indexOf('$', end)) !== -1) {
					if (state.src.charCodeAt(end - 1) === 92/*\\*/) {
						end++;
						continue;
					}
					content = state.src.slice(start + 1, end);
					content = content.trim();
					break;
				}

				if (!content)
					return false;

				if (!silent) {
					let token = state.push("latex_inline", "span", 0);
					token.content = content;
					token.attrs = [["class", "oo-latex-inline"]];
				}
			
				state.pos = end + 1;
				return true;
			});
			md.renderer.rules.latex_inline = function(tokens, idx) {
				return `<span class="oo-latex-inline">${tokens[idx].content}</span>`;
			};
			
			// Block: $$...$$  
			md.block.ruler.before("fence", "latex_block", function(state, startLine, endLine, silent) {
				let startPos = state.bMarks[startLine] + state.tShift[startLine];
				let maxPos = state.eMarks[startLine];
				let line = state.src.slice(startPos, maxPos).trim();
				
				if (!line.startsWith("$$"))
					return false;
				if (silent)
					return true;
				
				let content = "";
				let found = false;
				
				for (let i = startLine + 1; i < endLine; i++) {
					let pos = state.bMarks[i] + state.tShift[i];
					let max = state.eMarks[i];
					let nextLine = state.src.slice(pos, max).trim();

					if (nextLine === "$$") {
						found = true;
						state.line = i + 1;
						break;
					}

					content += nextLine + "\n";
				}
				
				if (!found) return false;
				
				const token = state.push("latex_block", "span", 0);
				token.block = true;
				token.content = content.trim();
				token.attrs = [["class", "oo-latex"]];
				token.map = [startLine, state.line];

				return true;
			});
			md.renderer.rules.latex_block = function(tokens, idx) {
				return `<span class="oo-latex">${tokens[idx].content}</span>\n`;
			};
		},

		forms: function(md) {
			let fieldCounter = 1837335014;
			let fieldKeyCounter = 1;

			function parseField(state, silent) {
				const start = state.pos;
				const max = state.posMax;

				if (state.src.charCodeAt(start) !== 0x7B/*{*/ ||
					state.src.slice(start, start + 7) !== '{FIELD:') {
					return false;
				}

				let pos = start + 7;
				let depth = 1;
				while (pos < max && depth > 0) {
					if (state.src.charCodeAt(pos) === 0x7B) depth++;
					if (state.src.charCodeAt(pos) === 0x7D) depth--;
					pos++;
				}

				if (depth !== 0) return false;

				const content = state.src.slice(start + 7, pos - 1);
				
				if (!silent) {
					const token = state.push('field', '', 0);
					token.content = content;
					token.markup = '{FIELD:}';
				}

				state.pos = pos;
				return true;
			}

			function parseParams(content) {
				const params = {};
				const regex = /(\w+)[:=](?:'([^']*)'|"([^"]*)"|([^,}]*))/g;
				let match;

				while ((match = regex.exec(content)) !== null) {
					const key = match[1];
					const value = match[2] || match[3] || match[4];
					params[key] = value;
				}

				return params;
			}

			function renderField(tokens, idx) {
				const params = parseParams(tokens[idx].content);
				const type = params.type;

				switch (type) {
					case 'checkbox':
						return renderCheckbox(params);
					case 'radiobutton':
						return renderRadiobutton(params);
					case 'combobox':
						return renderCombobox(params);
					case 'textbox':
						return renderTextbox(params);
					case 'date':
						return renderDate(params);
					default:
						return '';
				}
			}

			function renderCheckbox(params) {
				const checked = params.checked === 'true';
				const symbolChecked = params.symbolChecked || '☑';
				const symbolUnchecked = params.symbolUnchecked || '☐';
				const key = params.key || ('Checkbox' + fieldKeyCounter++);
				const text = params.text || '';
				const displaySymbol = checked ? symbolChecked : symbolUnchecked;

				return `<w:Sdt CheckBox="t" Form="t" CheckBoxValueChecked="${symbolChecked}" CheckBoxValueUnchecked="${symbolUnchecked}" Key="${key}" Text="${text}"/></w:Sdt>`;
			}

			function renderRadiobutton(params) {
				const checked = params.checked === 'true';
				const symbolChecked = params.symbolChecked || '◉';
				const symbolUnchecked = params.symbolUnchecked || '○';
				const groupKey = params.groupKey || params.GroupKey || 'Group 1';
				const key = params.key || ('Radio' + fieldKeyCounter++);
				const text = params.text || '';
				const displaySymbol = checked ? symbolChecked : symbolUnchecked;

				return `<w:Sdt CheckBox="t" Form="t" CheckBoxValueChecked="${symbolChecked}" CheckBoxValueUnchecked="${symbolUnchecked}" GroupKey="${groupKey}" Key="${key}" Text="${text}"></w:Sdt>`;
			}

			function renderCombobox(params) {
				const items = params.items ? params.items.split(',') : [];
				const selected = params.selected || '';
				const key = params.key || ('Combobox' + fieldKeyCounter++);
				const id = fieldCounter++;

				let itemsHtml = items.map(item => {
					const trimmed = item.trim();
					return `<w:ListItem ListValue="${trimmed}" DataValue="${trimmed}"></w:ListItem>`;
				}).join('');

				return `<w:Sdt ComboBox="t" Form="t" Key="${key}" ID="${id}">${itemsHtml}\n</w:Sdt>`;
			}

			function renderTextbox(params) {
				const key = params.key || ('Textbox' + fieldKeyCounter++);
				const text = params.text || params.value || '';
				const placeholder = params.placeholder || '';

				if (placeholder === "" && text === "")
					text = "empty";
				
				if (placeholder === "")
					return `<w:Sdt Form="t" Key="${key}">${text}</w:Sdt>`;
				else
					return `<w:Sdt Form="t" Key="${key}" PlcHdr="${placeholder}" ShowingPlcHdr="t">${text}</w:Sdt>`;
			}

			function renderDate(params) {
				const key = params.key || ('DatePicker' + fieldKeyCounter++);
				const value = params.value || '';
				
				let isoDate = '';
				let displayDate = value;
			
				if (value) {
					const parts = value.split('.');
					if (parts.length === 3) {
						isoDate = `${parts[2]}-${parts[1]}-${parts[0]}`;
						displayDate = value;
					}
				}
				return `<w:Sdt Calendar="t" Form="t" MapToDateTime="t" CalendarType="Gregorian" Date="${isoDate}" DateFormat="dd.MM.yyyy" Key="${key}">${displayDate}</w:Sdt>`;
			}

			md.inline.ruler.before('emphasis', 'field', parseField);
			md.renderer.rules.field = renderField;
		},

		hr: function(md) {
			md.renderer.rules.hr = function() {
				return '<hr><br>';
			};
		}
	};

	function decodeHtmlText(text) {
		return text
			.replace(/&quot;/g, '"')
			.replace(/&apos;/g, "'")
			.replace(/&amp;/g, '&')
			.replace(/&lt;/g, '<')
			.replace(/&gt;/g, '>')			
			.replace(/&nbsp;/g, ' ');
	}

	Library.prototype.GetEditorVersion = async function()
	{
		if (this.version !== 0)
			return this.version;

		let version = await Editor.callMethod("GetVersion");
		if ("develop" == version)
			version = "99.99.99";

		let arrVer = version.split(".");
		while (3 > arrVer.length)
			arrVer.push("0");

		this.version = 1000000 * parseInt(arrVer[0]) +  1000 * parseInt(arrVer[1]) + parseInt(arrVer[2]);
		return this.version;
	};

	Library.prototype.GetCurrentWord = async function()
	{
		return await Editor.callMethod("GetCurrentWord");
	};

	Library.prototype.GetSelectedText = async function()
	{
		let result = await Editor.callMethod("GetSelectedText");
		if (result !== "")
			return result;

		return this.GetSelectedContent("text");
	};

	Library.prototype.GetSelectedContent = async function(type) {
		return await Editor.callMethod("GetSelectedContent", [{ type : type }]);
	};

	Library.prototype.GetSelectedImage = async function(type) {
		let res = await Editor.callMethod("GetSelectedContent", [{ type : "html" }]);
		let index1 = res.indexOf("src=\"data:image/");
		if (-1 === index1)
			return "";
		index1 += 5;
		let index2 = res.indexOf("\"", index1);
		if (-1 === index2)
			return "";
		return res.substring(index1, index2);
	};

	Library.prototype.ReplaceTextSmart = async function(text)
	{
		return await Editor.callMethod("ReplaceTextSmart", [text]);
	};

	Library.prototype.InsertAsText = async function(text)
	{
		Asc.scope.data = (text || "").split("\n\n");
		return await Editor.callCommand(function() {
			let oDocument = Api.GetDocument();
			for (let ind = 0; ind < Asc.scope.data.length; ind++) {
				let text = Asc.scope.data[ind];
				if (text.length) {
					let oParagraph = Api.CreateParagraph();
					oParagraph.AddText(text);
					oDocument.Push(oParagraph);
				}
			}
		});
	};

	Library.prototype.getHTMLFromMD = function(data, plugins, isStreaming)
	{
		return Asc.Library.ConvertMdToHTML(data, plugins, isStreaming !== false);
	};

	Library.prototype.InsertAsMD = async function(data, plugins)
	{
		let htmlContent = Asc.Library.ConvertMdToHTML(data, plugins);

		if (false) {
			console.log("MD Content:\n" + data);
			console.log("HTML Content:\n" + htmlContent);
		}

		return await Asc.Library.InsertAsHTML(htmlContent);
	};

	Library.prototype.ConvertMdToHTML = function(data, plugins, isStreaming)
	{
		let c = window.markdownit();
		if (plugins) {
			for (let i = 0, len = plugins.length; i < len; i++)
				c.use(plugins[i]);
		}
		return c.render(this.getMarkdownResult(data, isStreaming));
	};

	Library.prototype.InsertAsHTML = async function(data)
	{
		switch (Asc.Editor.getType()) {
			case "word": {
				if (true) {
					await Editor.callCommand(function() {
						let document = Api.GetDocument();
						document.RemoveSelection();
					}, false);
				} else {
					await Editor.callCommand(function() {
						let doc = Api.GetDocument();
						let paras = doc.GetAllParagraphs();
						if (paras.length)
						{
							let lastPara = paras[paras.length - 1];
							let lastElement = lastPara.GetElement(lastPara.GetElementsCount() - 1);
							if (lastElement && lastElement.MoveCursorToPos)
							{
								lastElement.MoveCursorToPos(100000);
							}
						}
					});
				}
			}
			default:
				break;
		}
		return await Editor.callMethod("PasteHtml", [data]);
	};

	Library.prototype.InsertAsComment = async function(text)
	{
		return await Editor.callMethod("AddComment", [{
			UserName : "AI",
			Text : decodeHtmlText(text),
			Time: Date.now(),
			Solver: false
		}]);
	};

	Library.prototype.InsertAsHyperlink = async function(content, hint)
	{
		let text = content;
		start = text.indexOf('htt');
		end = text.indexOf(' ', start);
		if (end == -1)
			end = text.length;

		Asc.scope.link = text.slice(start, end);
		return await Editor.callCommand(function(){
			let oDocument = Api.GetDocument();
			let oRange = oDocument.GetRangeBySelect();
			oRange.AddHyperlink(Asc.scope.link, "Meaning of the word");
		});
	};

	Library.prototype.InsertAsReview = async function(content, isHtml) 
	{
		let isTrackRevisions = await Editor.callCommand(function(){
			let res = Api.asc_GetLocalTrackRevisions();
			Api.asc_SetLocalTrackRevisions(true);
			return res;
		});

		Asc.scope.localTrackRevisions = isTrackRevisions;

		await Editor.callMethod(isHtml ? "PasteHtml" : "PasteText", [content.trim()]);

		if (true !== isTrackRevisions) 
		{
			await Editor.callCommand(function(){
				Api.asc_SetLocalTrackRevisions(Asc.scope.localTrackRevisions);
			});
		}
	};

	Library.prototype.PasteText = async function(text)
	{
		return await Editor.callMethod("PasteText", [text]);
	};

	Library.prototype.SendError = async function(text, errorLevel)
	{
		Asc.scope.errorText = text;
		Asc.scope.errorLevel = errorLevel;
		return await Editor.callCommand(function(){
			Api.sendEvent("asc_onError", Asc.scope.errorText, Asc.scope.errorLevel);
		});
	};

	Library.prototype.GetLocalImagePath = async function(url) {
		return await Editor.callMethod("getLocalImagePath", [url]);
	};

	Library.prototype.AddGeneratedImage = async function(base64) {
		let editorVersion = await Asc.Library.GetEditorVersion();

		if (Asc.Editor.getType() === "pdf") {
			return await Editor.callMethod("PasteHtml", ["<img src=\"" + base64 + "\" />"]);
		}
		
		if (editorVersion >= 9000000) {
			let urlLocal = await this.GetLocalImagePath(base64);
			if (urlLocal.error === true)
				return;

			Asc.scope.url = urlLocal.url;
		} else {
			Asc.scope.url = url;
		}

		switch (window.Asc.plugin.info.editorType) {
			case "word": {
				return await Editor.callCommand(function() {
					let document = Api.GetDocument();
					let paragraph = Api.CreateParagraph();
					let drawing = Api.CreateImage(Asc.scope.url, 100 * 36000, 100 * 36000);
					paragraph.AddDrawing(drawing);
					document.RemoveSelection();
					document.InsertContent([paragraph], true);
				}, false);
			}
			case "cell": {
				return await Editor.callCommand(function() {
					let worksheet = Api.GetActiveSheet();
					worksheet.AddImage(Asc.scope.url, 100 * 36000, 100 * 36000, 0, 2 * 36000, 2, 3 * 36000);
				}, false);
			}
			case "slide": {
				return await Editor.callCommand(function() {
					let presentation = Api.GetPresentation();
					let slide = presentation.GetCurrentSlide();
					let image = Api.CreateImage(Asc.scope.url, 150 * 36000, 150 * 36000);
					slide.AddObject(image);
				}, false);
			}
			default:
				break;
		}
	};

	Library.prototype.AddOleObject = async function(imageUrl, data) {
		switch (window.Asc.plugin.info.editorType) {
			case "word": {
				await Editor.callCommand(function(){
					let document = Api.GetDocument();
					document.RemoveSelection();
				});
				break;
			}
			default:
				break;
		}

		let W = 100;
		let H = 100;

		let info = window.Asc.plugin.info;
		var obj = {
			guid : info.guid,
			widthPix : info.mmToPx * W,
			heightPix : info.mmToPx * H,
			width : W,
			height : H,
			imgSrc : imageUrl,
			data : data
		};

		return await Editor.callMethod("AddOleObject", [obj]);
	};

	Library.prototype.trimResult = function(data, posStart, isSpaces, extraCharacters) {
		let pos = posStart || 0;
		if (-1 != pos) {
			let trimC = ["\"", "'", "\n", "\r", "`"];
			if (true === isSpaces)
				trimC.push(" ");
			while (pos < data.length && trimC.includes(data[pos]))
				pos++;

			let posEnd = data.length - 1;
			while (posEnd > 0 && trimC.includes(data[posEnd]))
				posEnd--;

			if (posEnd > pos)
				return data.substring(pos, posEnd + 1);				
		}
		return data;
	};

	Library.prototype.getTranslateResult = function(data, dataSrc) {
		data = this.trimResult(data, 0, true);
		let trimC = ["\"", "'", "\n", "\r", " "];
		if (dataSrc.length > 0 && trimC.includes(dataSrc[0])) {
			data = dataSrc[0] + data;
		}
		if (dataSrc.length > 1 && trimC.includes(dataSrc[dataSrc.length - 1])) {
			data = data + dataSrc[dataSrc.length - 1];
		}
		return data;
	};

	Library.prototype.getMarkdownResult = function(data, isStreaming) {
		let markdownEscape = data.indexOf("```md");
		if (-1 !== markdownEscape && markdownEscape < 5)
			data = data.substring(markdownEscape + 5);
		if (data.endsWith("```"))
			data = data.slice(0, -3);
		let correctData = data.replace(/\n---#/g, '\n---\n#');
		return isStreaming ? correctData : this.trimResult(correctData);
	};

	Library.prototype.getJSONResult = function(data) {
		const markdownMarker = "```json";
		let markdownEscape = data.indexOf(markdownMarker);
		if (-1 !== markdownEscape && markdownEscape < 5)
			data = data.substring(markdownEscape + markdownMarker.length);
		return this.trimResult(data);
	};

	exports.Asc = exports.Asc || {};
	exports.Asc.Library = new Library();

	exports.Asc.Prompts = {
		getFixAndSpellPrompt(content) {
			let prompt = `I want you to act as an editor and proofreader. \
I will provide you with some text that needs to be checked for spelling and grammar errors. \
Your task is to carefully review the text and correct any mistakes, \
ensuring that the corrected text is free of errors and maintains the original meaning. \
Only return the corrected text. \
Here is the text that needs revision: \"${content}\"`;
			return prompt;
		},
		getSummarizationPrompt(content, language) {
			let prompt = "Summarize the following text. ";
			if (language) {
				prompt += "and translate the result to " + language;
				prompt += ". Return only the resulting translated text.";
			} else {
				prompt += ". Return only the resulting text.";
			}
			prompt += "Text: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getTranslatePrompt(content, language) {
			let prompt = "Translate the following text to " + language;
			prompt += ". Return only the resulting text.";
			prompt += "Text: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getExplainPrompt(content) {
			let prompt = "Explain what the following text means. Return only the resulting text.";
			prompt += "Text: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getTextLongerPrompt(content) {
			let prompt = "Make the following text longer. Return only the resulting text.";
			prompt += "Text: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getTextShorterPrompt(content) {
			// SCRIVAR-RULES: "shorter" (upstream asked for "simpler").
			let prompt = "Make the following text shorter while keeping its meaning. Return only the resulting text.";
			prompt += "\nText: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getTextRewritePrompt(content) {
			let prompt = "Rewrite the following text differently. Return only the resulting text.";
			prompt += "Text: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getTextKeywordsPrompt(content) {
			// SCRIVAR-RULES: one plain line (upstream's answer came back as markdown lists).
			let prompt = "List the key words and phrases from this text as a single comma-separated line of plain text, with no heading, numbering or formatting.";
			prompt += "\nText: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getExplainAsLinkPrompt(content) {
			// SCRIVAR-RULES: no invented URLs; register.js treats NO_LINK as "no link found".
			let prompt = "Give one link to a well-known reference page (such as Wikipedia) that explains the following text. Return only the URL. If you are not certain such a page exists, return exactly: NO_LINK";
			prompt += "\nText: \"\"\"\n";
			prompt += content;
			prompt += "\n\"\"\"";
			return prompt;
		},
		getImageDescription() {
			return "Describe in detail everything you see in this image. Mention the objects, their appearance, colors, arrangement, background, and any noticeable actions or interactions. Be as specific and accurate as possible. Avoid making assumptions about things that are not clearly visible."
		},
		getImagePromptOCR() {
			return "Extract all text from this image as accurately as possible. Preserve original reading order and formatting if possible. Recognize tables and images if possible. Do not add or remove any content. Output recognized objects in md format if possible. If not, return plain text.";
		}
	};

})(window);