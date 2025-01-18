/* OpenDocument */
function write_styles_ods(wb/*:any*/, opts/*:any*/)/*:string*/ {
	var master_styles = opts?.stayStyle ? makeOdsStyles(wb, opts) : [
		'<office:master-styles>',
			'<style:master-page style:name="mp1" style:page-layout-name="mp1">',
				'<style:header/>',
				'<style:header-left style:display="false"/>',
				'<style:footer/>',
				'<style:footer-left style:display="false"/>',
			'</style:master-page>',
		'</office:master-styles>'
	].join("");

	var payload = '<office:document-styles ' + wxt_helper({
		'xmlns:office': "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		'xmlns:ooo': "http://openoffice.org/2004/office",
		'xmlns:fo': "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		'xmlns:xlink': "http://www.w3.org/1999/xlink",
		'xmlns:dc': "http://purl.org/dc/elements/1.1/",
		'xmlns:meta': "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
		'xmlns:style': "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		'xmlns:text': "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		'xmlns:rpt': "http://openoffice.org/2005/report",
		'xmlns:draw': "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
		'xmlns:dr3d': "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
		'xmlns:svg': "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		'xmlns:chart': "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
		'xmlns:table': "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		'xmlns:number': "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		'xmlns:ooow': "http://openoffice.org/2004/writer",
		'xmlns:oooc': "http://openoffice.org/2004/calc",
		'xmlns:css3t': "http://www.w3.org/TR/css3-text/",
		'xmlns:of': "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		'xmlns:tableooo': "http://openoffice.org/2009/table",
		'xmlns:calcext': "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
		'xmlns:drawooo': "http://openoffice.org/2010/draw",
		'xmlns:xhtml': "http://www.w3.org/1999/xhtml",
		'xmlns:loext': "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
		'xmlns:grddl': "http://www.w3.org/2003/g/data-view#",
		'xmlns:field': "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
		'xmlns:math': "http://www.w3.org/1998/Math/MathML",
		'xmlns:form': "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
		'xmlns:script': "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
		'xmlns:dom': "http://www.w3.org/2001/xml-events",
		'xmlns:presentation': "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
		'office:version': "1.3"
	}) + '>' + master_styles + '</office:document-styles>';

	return XML_HEADER + payload;
};

// TODO: find out if anyone actually read the spec.  LO has some wild errors
function write_number_format_ods(nf/*:string*/, nfidx/*:string*/)/*:string*/ {
	var type = "number", payload = "", nopts = { "style:name": nfidx }, c = "", i = 0;
	nf = nf.replace(/"[$]"/g, "$");
	/* TODO: replace with an actual parser based on a real grammar */
	j: {
		// TODO: support style maps
		if(nf.indexOf(";") > -1) {
			console.error("Unsupported ODS Style Map exported.  Using first branch of " + nf);
			nf = nf.slice(0, nf.indexOf(";"));
		}

		if(nf == "@") { type = "text"; payload = "<number:text-content/>"; break j; }

		/* currency flag */
		if(nf.indexOf(/\$/) > -1) { type = "currency"; }

		/* opening string literal */
		if(nf[i] == '"') {
			c = "";
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			if(nf[i+1] == "*") {
				i++;
				payload += '<number:fill-character>' + escapexml(c.replace(/""/g, '"')) + '</number:fill-character>';
			} else {
				payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
			}
			nf = nf.slice(i+1); i = 0;
		}

		/* fractions */
		var t = nf.match(/# (\?+)\/(\?+)/);
		if(t) { payload += writextag("number:fraction", null, {"number:min-integer-digits":0, "number:min-numerator-digits": t[1].length, "number:max-denominator-value": Math.max(+(t[1].replace(/./g, "9")), +(t[2].replace(/./g, "9"))) }); break j; }
		if((t=nf.match(/# (\?+)\/(\d+)/))) { payload += writextag("number:fraction", null, {"number:min-integer-digits":0, "number:min-numerator-digits": t[1].length, "number:denominator-value": +t[2]}); break j; }

		/* percentages */
		if((t=nf.match(/\b(\d+)(|\.\d+)%/))) { type = "percentage"; payload += writextag("number:number", null, {"number:decimal-places": t[2] && t.length - 1 || 0, "number:min-decimal-places": t[2] && t.length - 1 || 0, "number:min-integer-digits": t[1].length }) + "<number:text>%</number:text>"; break j; }

		/* datetime */
		var has_time = false;
		if(["y","m","d"].indexOf(nf[0]) > -1) {
			type = "date";
			k: for(; i < nf.length; ++i) switch((c = nf[i].toLowerCase())) {
				case "h": case "s": has_time = true; --i; break k;
				case "m":
					l: for(var h = i+1; h < nf.length; ++h) switch(nf[h]) {
						case "y": case "d": break l;
						case "h": case "s": has_time = true; --i; break k;
					}
					/* falls through */
				case "y": case "d":
					while((nf[++i]||"").toLowerCase() == c[0]) c += c[0]; --i;
					switch(c) {
						case "y": case "yy": payload += "<number:year/>"; break;
						case "yyy": case "yyyy": payload += '<number:year number:style="long"/>'; break;
						case "mmmmm": console.error("ODS has no equivalent of format |mmmmm|");
							/* falls through */
						case "m": case "mm": case "mmm": case "mmmm":
							payload += '<number:month number:style="' + (c.length % 2 ? "short" : "long") + '" number:textual="' + (c.length >= 3 ? "true" : "false") + '"/>';
							break;
						case "d": case "dd": payload += '<number:day number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "ddd": case "dddd": payload += '<number:day-of-week number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
					}
					break;
				case '"':
					while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
					payload += '<number:text>' + escapexml(c.slice(1).replace(/""/g, '"')) + '</number:text>';
					break;
				case '\\': c = nf[++i];
					payload += '<number:text>' + escapexml(c) + '</number:text>'; break;
				case '/': case ':': payload += '<number:text>' + escapexml(c) + '</number:text>'; break;
				default: console.error("unrecognized character " + c + " in ODF format " + nf);
			}
			if(!has_time) break j;
			nf = nf.slice(i+1); i = 0;
		}
		if(nf.match(/^\[?[hms]/)) {
			if(type == "number") type = "time";
			if(nf.match(/\[/)) {
				nf = nf.replace(/[\[\]]/g, "");
				nopts['number:truncate-on-overflow'] = "false";
			}
			for(; i < nf.length; ++i) switch((c = nf[i].toLowerCase())) {
				case "h": case "m": case "s":
					while((nf[++i]||"").toLowerCase() == c[0]) c += c[0]; --i;
					switch(c) {
						case "h": case "hh": payload += '<number:hours number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "m": case "mm": payload += '<number:minutes number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "s": case "ss":
							if(nf[i+1] == ".") do { c += nf[i+1]; ++i; } while(nf[i+1] == "0");
							payload += '<number:seconds number:style="' + (c.match("ss") ? "long" : "short") + '"' + (c.match(/\./) ? ' number:decimal-places="' + (c.match(/0+/)||[""])[0].length + '"' : "")+ '/>'; break;
					}
					break;
				case '"':
					while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
					payload += '<number:text>' + escapexml(c.slice(1).replace(/""/g, '"')) + '</number:text>';
					break;
				case '/': case ':': payload += '<number:text>' + escapexml(c) + '</number:text>'; break;
				case "a":
					if(nf.slice(i, i+3).toLowerCase() == "a/p") { payload += '<number:am-pm/>'; i += 2; break; } // Note: ODF does not support A/P
					if(nf.slice(i, i+5).toLowerCase() == "am/pm")  { payload += '<number:am-pm/>'; i += 4; break; }
					/* falls through */
				default: console.error("unrecognized character " + c + " in ODF format " + nf);
			}
			break j;
		}

		/* currency flag */
		if(nf.indexOf(/\$/) > -1) { type = "currency"; }

		/* should be in a char loop */
		if(nf[0] == "$") { payload += '<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>'; nf = nf.slice(1); i = 0; }
		i = 0; if(nf[i] == '"') {
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			if(nf[i+1] == "*") {
				i++;
				payload += '<number:fill-character>' + escapexml(c.replace(/""/g, '"')) + '</number:fill-character>';
			} else {
				payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
			}
			nf = nf.slice(i+1); i = 0;
		}

		/* number TODO: interstitial text e.g. 000)000-0000 */
		var np = nf.match(/([#0][0#,]*)(\.[0#]*|)(E[+]?0*|)/i);
		if(!np || !np[0]) console.error("Could not find numeric part of " + nf);
		else {
			var base = np[1].replace(/,/g, "");
			payload += '<number:' + (np[3] ? "scientific-" : "")+ 'number' +
				' number:min-integer-digits="' + (base.indexOf("0") == -1 ? "0" : base.length - base.indexOf("0")) + '"' +
				(np[0].indexOf(",") > -1 ? ' number:grouping="true"' : "") +
				(np[2] && ' number:decimal-places="' + (np[2].length - 1) + '"' || ' number:decimal-places="0"') +
				(np[3] && np[3].indexOf("+") > -1 ? ' number:forced-exponent-sign="true"' : "" ) +
				(np[3] ? ' number:min-exponent-digits="' + np[3].match(/0+/)[0].length + '"' : "" ) +
				'>' +
				/* TODO: interstitial text placeholders */
				'</number:' + (np[3] ? "scientific-" : "") + 'number>';
			i = np.index + np[0].length;
		}

		/* residual text */
		if(nf[i] == '"') {
			c = "";
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
		}
	}

	if(!payload) { console.error("Could not generate ODS number format for |" + nf + "|"); return ""; }
	return writextag("number:" + type + "-style", payload, nopts);
}

function write_names_ods(Names, SheetNames, idx) {
	//var scoped = Names.filter(function(name) { return name.Sheet == (idx == -1 ? null : idx); });
	var scoped = []; for(var namei = 0; namei < Names.length; ++namei) {
		var name = Names[namei];
		if(!name) continue;
		if(name.Sheet == (idx == -1 ? null : idx)) scoped.push(name);
	}
	if(!scoped.length) return "";
	return "<table:named-expressions>" + scoped.map(function(name) {
		var odsref =  (idx == -1 ? "$" : "") + csf_to_ods_3D(name.Ref);
		return writextag("table:named-range", null, {
			"table:name": name.Name,
			"table:cell-range-address": odsref,
			"table:base-cell-address": odsref.replace(/[\.][^\.]*$/, ".$A$1")
		});
	}).join("") + "</table:named-expressions>";
}
var write_content_ods/*:{(wb:any, opts:any):string}*/ = /* @__PURE__ */(function() {
	/* 6.1.2 White Space Characters */
	var write_text_p = function(text/*:string*/, span)/*:string*/ {
		return escapexml(text)
			.replace(/  +/g, function($$){return '<text:s text:c="'+$$.length+'"/>';})
			.replace(/\t/g, "<text:tab/>")
			.replace(/\n/g, span ? "<text:line-break/>": "</text:p><text:p>")
			.replace(/^ /, "<text:s/>").replace(/ $/, "<text:s/>");
	};

	var null_cell_xml = '<table:table-cell/>';
	var coveredTableCell_xml = function(count) {
		let s = '<table:covered-table-cell';
		if (count > 1) {
			s += ` table:number-columns-repeated="${count}"`
		}
		s += '/>';
		return s;
	}
	var write_ws = function(ws, wb/*:Workbook*/, i/*:number*/, opts, nfs, date1904)/*:string*/ {
		/* Section 9 Tables */
		var o/*:Array<string>*/ = [];
		var tstyle = opts?.stayStyle && ws?.sn;
		if (!tstyle) tstyle = ((((wb||{}).Workbook||{}).Sheets||[])[i]||{}).Hidden ? 'ta2' : 'ta1';
		o.push('<table:table table:name="' + escapexml(wb.SheetNames[i]) + '" table:style-name="' + tstyle + '">');
		var R=0,C=0, range = decode_range(ws['!ref']||"A1");
		var marr/*:Array<Range>*/ = ws['!merges'] || [], mi = 0;
		var dense = ws["!data"] != null;
		if(ws["!cols"]) {
			let ar = [];
			for(C = 0; C <= range.e.c; ++C) {
				let col = ws["!cols"][C] || {};
				let pre = ar.length > 0 ? ar.slice(-1)[0] : {};
				if (col?.ods === pre?.ods && col?.dsn === pre?.dsn) {
					if (pre?.count === undefined) pre.count = 2;
					else pre.count++;
				} else {
					ar.push(col);
				}
			}
			ar.forEach(function(col) {
				let s = '<table:table-column';
				if (col.ods !== undefined) s += ` table:style-name="co${col.ods}"`;
				if (col.dsn) s += ` table:default-cell-style-name="${col.dsn}"`;
				if (col.count) s += ` table:number-columns-repeated="${col.count}"`;
				o.push(s + '/>');
			});
		}
		var H = "", ROWS = ws["!rows"]||[];
		for(R = 0; R < range.s.r; ++R) {
			H = ROWS[R] ? ' table:style-name="ro' + ROWS[R].ods + '"' : "";
			o.push('<table:table-row' + H + '></table:table-row>');
		}
		for(; R <= range.e.r; ++R) {
			H = ROWS[R] ? ' table:style-name="ro' + ROWS[R].ods + '"' : "";
			o.push('<table:table-row' + H + '>');
			for(C=0; C < range.s.c; ++C) o.push(null_cell_xml);
			for(; C <= range.e.c; ++C) {
				var ct = {}, textp = "";
				let m = marr.find(m => m.s.r <= R && R <= m.e.r && m.s.c <= C && C <= m.e.c);
				let cCount = 0;
				if (m) {
					cCount = m.e.c - m.s.c;
					if (m.s.c === C) {
						if (m.s.r === R) {
							ct['table:number-columns-spanned'] = cCount + 1;
							ct['table:number-rows-spanned'] = m.e.r - m.s.r + 1;
						} else {
							o.push(coveredTableCell_xml(cCount + 1));
							C += cCount;
							continue;
						}
					} else {
						continue;
					}
				}
				var ref = encode_cell({r:R, c:C}), cell = dense ? (ws["!data"][R]||[])[C]: ws[ref];
				if(cell && cell.f) {
					ct['table:formula'] = escapexml(csf_to_ods_formula(cell.f));
					if(cell.F) {
						if(cell.F.slice(0, ref.length) == ref) {
							var _Fref = decode_range(cell.F);
							ct['table:number-matrix-columns-spanned'] = (_Fref.e.c - _Fref.s.c + 1);
							ct['table:number-matrix-rows-spanned'] =    (_Fref.e.r - _Fref.s.r + 1);
						}
					}
				}
				if(!cell) { o.push(null_cell_xml); continue; }
				switch(cell.t) {
					case 'b':
						textp = (cell.v ? 'TRUE' : 'FALSE');
						ct['office:value-type'] = "boolean";
						ct['office:boolean-value'] = (cell.v ? 'true' : 'false');
						break;
					case 'n':
						if(!isFinite(cell.v)) {
							if(isNaN(cell.v)) {
								textp = "#NUM!";
								ct['table:formula'] = "of:=#NUM!";
							} else {
								textp = "#DIV/0!";
								ct['table:formula'] = "of:=" + (cell.v < 0 ? "-" : "") + "1/0";
							}
							ct['office:string-value'] = "";
							ct['office:value-type'] = "string";
							ct['calcext:value-type'] = "error";
						} else {
							textp = (cell.w||String(cell.v||0));
							ct['office:value-type'] = cell.vt || "float";
							ct['office:value'] = (cell.v||0);
						}
						break;
					case 's': case 'str':
						textp = cell.v == null ? "" : cell.v;
						ct['office:value-type'] = "string";
						break;
					case 'd':
						ct['office:value-type'] = cell.vt || "date";
						switch (cell.vt) {
						case 'date':
							textp = (cell.w||String(cell.v));
							ct['date-value'] = convertToOfficeDateValue(cell.v);
							break;
						case 'time':
							textp = (cell.w||String(cell.v));
							ct['time-value'] = convertToOfficeTimeValue(cell.v);
							break;
						default:
							textp = (cell.w||(parseDate(cell.v, date1904).toISOString()));
							ct['office:date-value'] = (parseDate(cell.v, date1904).toISOString());
							ct['table:style-name'] = "ce1";
							break;
						}
						break;
					//case 'e': // TODO: translate to ODS errors
					default: o.push(null_cell_xml); continue; // TODO: empty cell with comments
				}
				var text_p = write_text_p(textp);
				if(cell.l && cell.l.Target) {
					var _tgt = cell.l.Target;
					_tgt = _tgt.charAt(0) == "#" ? "#" + csf_to_ods_3D(_tgt.slice(1)) : _tgt;
					// TODO: choose correct parent path format based on link delimiters
					if(_tgt.charAt(0) != "#" && !_tgt.match(/^\w+:/)) _tgt = '../' + _tgt;
					text_p = writextag('text:a', text_p, {'xlink:href': _tgt.replace(/&/g, "&amp;")});
				}
				let tsn = opts?.stayStyle && cell?.sn;
				if (!tsn && nfs[cell.z]) tsn = "ce" + nfs[cell.z].slice(1);
				if (tsn) ct["table:style-name"] = tsn;
				var payload = writextag('text:p', text_p, {});
				if(cell.c) {
					var acreator = "", apayload = "", aprops = {};
					for(var ci = 0; ci < cell.c.length; ++ci) {
						if(!acreator && cell.c[ci].a) acreator = cell.c[ci].a;
						apayload += "<text:p>" + write_text_p(cell.c[ci].t) + "</text:p>";
					}
					if(!cell.c.hidden) aprops["office:display"] = true;
					payload = writextag('office:annotation', apayload, aprops) + payload;
				}
				o.push(writextag('table:table-cell', payload, ct));
				if (cCount > 0) {
					o.push(coveredTableCell_xml(cCount));
					C += cCount;
				}
			}
			o.push('</table:table-row>');
		}
		if((wb.Workbook||{}).Names) o.push(write_names_ods(wb.Workbook.Names, wb.SheetNames, i));
		o.push('</table:table>');
		return o.join("");
	};

	var write_automatic_styles_ods = function(o/*:Array<string>*/, wb) {
		o.push('<office:automatic-styles>');

		/* column styles */
		var cidx = 0;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			if(ws["!cols"]) {
				for(var C = 0; C < ws["!cols"].length; ++C) if(ws["!cols"][C]) {
					var colobj = ws["!cols"][C];
					if(colobj.width == null && colobj.wpx == null && colobj.wch == null) continue;
					process_col(colobj);
					colobj.ods = cidx;
					var w = ws["!cols"][C].wpx + "px";
					o.push('<style:style style:name="co' + cidx + '" style:family="table-column">');
					o.push('<style:table-column-properties fo:break-before="auto" style:column-width="' + w + '"/>');
					o.push('</style:style>');
					++cidx;
				}
			}
		});

		/* row styles */
		var ridx = 0;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			if(ws["!rows"]) {
				for(var R = 0; R < ws["!rows"].length; ++R) if(ws["!rows"][R]) {
					ws["!rows"][R].ods = ridx;
					var h = ws["!rows"][R].hpx + "px";
					o.push('<style:style style:name="ro' + ridx + '" style:family="table-row">');
					o.push('<style:table-row-properties fo:break-before="auto" style:row-height="' + h + '"/>');
					o.push('</style:style>');
					++ridx;
				}
			}
		});

		/* table */
		o.push('<style:style style:name="ta1" style:family="table" style:master-page-name="mp1">');
		o.push('<style:table-properties table:display="true" style:writing-mode="lr-tb"/>');
		o.push('</style:style>');
		o.push('<style:style style:name="ta2" style:family="table" style:master-page-name="mp1">');
		o.push('<style:table-properties table:display="false" style:writing-mode="lr-tb"/>');
		o.push('</style:style>');

		o.push('<number:date-style style:name="N37" number:automatic-order="true">');
		o.push('<number:month number:style="long"/>');
		o.push('<number:text>/</number:text>');
		o.push('<number:day number:style="long"/>');
		o.push('<number:text>/</number:text>');
		o.push('<number:year/>');
		o.push('</number:date-style>');

		/* number formats, table cells, text */
		var nfs = {};
		var nfi = 69;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			var dense = (ws["!data"] != null);
			if(!ws["!ref"]) return;
			var range = decode_range(ws["!ref"]);
			for(var R = 0; R <= range.e.r; ++R) for(var C = 0; C <= range.e.c; ++C) {
				var c = dense ? (ws["!data"][R]||[])[C] : ws[encode_cell({r:R,c:C})];
				if(!c || !c.z || c.z.toLowerCase() == "general") continue;
				if(!nfs[c.z]) {
					var out = write_number_format_ods(c.z, "N" + nfi);
					if(out) { nfs[c.z] = "N" + nfi; ++nfi; o.push(out); }
				}
			}
		});
		o.push('<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>');
		keys(nfs).forEach(function(nf) {
			o.push('<style:style style:name="ce' + nfs[nf].slice(1) + '" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="' + nfs[nf] + '"/>');
		});

		/* page-layout */

		o.push('</office:automatic-styles>');
		return nfs;
	};

	return function wcx(wb, opts) {
		var o = [XML_HEADER];
		/* 3.1.3.2 */
		var attr = wxt_helper({
			'xmlns:office':       "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
			'xmlns:table':        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
			'xmlns:style':        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
			'xmlns:text':         "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
			'xmlns:draw':         "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
			'xmlns:fo':           "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
			'xmlns:xlink':        "http://www.w3.org/1999/xlink",
			'xmlns:dc':           "http://purl.org/dc/elements/1.1/",
			'xmlns:meta':         "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
			'xmlns:number':       "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
			'xmlns:presentation': "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
			'xmlns:svg':          "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
			'xmlns:chart':        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
			'xmlns:dr3d':         "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
			'xmlns:math':         "http://www.w3.org/1998/Math/MathML",
			'xmlns:form':         "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
			'xmlns:script':       "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
			'xmlns:ooo':          "http://openoffice.org/2004/office",
			'xmlns:ooow':         "http://openoffice.org/2004/writer",
			'xmlns:oooc':         "http://openoffice.org/2004/calc",
			'xmlns:dom':          "http://www.w3.org/2001/xml-events",
			'xmlns:xforms':       "http://www.w3.org/2002/xforms",
			'xmlns:xsd':          "http://www.w3.org/2001/XMLSchema",
			'xmlns:xsi':          "http://www.w3.org/2001/XMLSchema-instance",
			'xmlns:sheet':        "urn:oasis:names:tc:opendocument:sh33tjs:1.0",
			'xmlns:rpt':          "http://openoffice.org/2005/report",
			'xmlns:of':           "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
			'xmlns:xhtml':        "http://www.w3.org/1999/xhtml",
			'xmlns:grddl':        "http://www.w3.org/2003/g/data-view#",
			'xmlns:tableooo':     "http://openoffice.org/2009/table",
			'xmlns:drawooo':      "http://openoffice.org/2010/draw",
			'xmlns:calcext':      "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
			'xmlns:loext':        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
			'xmlns:field':        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
			'xmlns:formx':        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
			'xmlns:css3t':        "http://www.w3.org/TR/css3-text/",
			'office:version':     "1.2"
		});

		var fods = wxt_helper({
			'xmlns:config':    "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
			'office:mimetype': "application/vnd.oasis.opendocument.spreadsheet"
		});

		if(opts.bookType == "fods") {
			o.push('<office:document' + attr + fods + '>');
			o.push(write_meta_ods().replace(/<office:document-meta[^<>]*?>/, "").replace(/<\/office:document-meta>/, ""));
			// TODO: settings (equiv of settings.xml for ODS)
		} else o.push('<office:document-content' + attr  + '>');
		// o.push('<office:scripts/>');
		var nfs = opts?.stayStyle && writeOdsStayStyles(o, wb, opts)
		if (!nfs) nfs = write_automatic_styles_ods(o, wb);
		o.push('<office:body>');
		o.push('<office:spreadsheet>');
		let ss = opts?.stayStyle && wb?.content?.body?.spreadsheet;
		let n = 'calculation-settings';
		if (ss && ss[n]) writeOdsCalculation(o, ss[n], n);
		else if (((wb.Workbook||{}).WBProps||{}).date1904) o.push('<table:calculation-settings table:case-sensitive="false" table:search-criteria-must-apply-to-whole-cell="true" table:use-wildcards="true" table:use-regular-expressions="false" table:automatic-find-labels="false"><table:null-date table:date-value="1904-01-01"/></table:calculation-settings>');
		for(var i = 0; i != wb.SheetNames.length; ++i) o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts, nfs, ((wb.Workbook||{}).WBProps||{}).date1904));
		if((wb.Workbook||{}).Names) o.push(write_names_ods(wb.Workbook.Names, wb.SheetNames, -1));
		o.push('</office:spreadsheet>');
		o.push('</office:body>');
		if(opts.bookType == "fods") o.push('</office:document>');
		else o.push('</office:document-content>');
		return o.join("");
	};
})();

function write_ods(wb/*:any*/, opts/*:any*/) {
	if(opts.bookType == "fods") return write_content_ods(wb, opts);

	var zip = zip_new();
	var f = "";

	var manifest/*:Array<Array<string> >*/ = [];
	var rdf/*:Array<[string, string]>*/ = [];

	/* Part 3 Section 3.3 MIME Media Type */
	f = "mimetype";
	zip_add_file(zip, f, "application/vnd.oasis.opendocument.spreadsheet");

	/* Part 1 Section 2.2 Documents */
	f = "content.xml";
	zip_add_file(zip, f, write_content_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "ContentFile"]);

	/* TODO: these are hard-coded styles to satiate excel */
	f = "styles.xml";
	zip_add_file(zip, f, write_styles_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "StylesFile"]);

	/* TODO: this is hard-coded to satiate excel */
	f = "meta.xml";
	zip_add_file(zip, f, XML_HEADER + write_meta_ods(/*::wb, opts*/));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "MetadataFile"]);

	/* Part 3 Section 6 Metadata Manifest File */
	f = "manifest.rdf";
	zip_add_file(zip, f, write_rdf(rdf/*, opts*/));
	manifest.push([f, "application/rdf+xml"]);

	/* Part 3 Section 4 Manifest File */
	f = "META-INF/manifest.xml";
	zip_add_file(zip, f, write_manifest(manifest/*, opts*/));

	return zip;
}
function setSheetHidden(sheet, ass, Sheet) {
	let sn = sheet['style-name'];
	if (sn) {
		let s = ass.find(s => s?.family === 'table' && s?.name === sn);
		if (!s) return;
		let tp = s && s['table-properties'];
		if (!tp) return;
		tp.display = Sheet.Hidden ? false : true;
	}
}
function setBookHidden(wb) {
	let Sheets = wb?.Workbook?.Sheets;
	let sheets = wb?.content?.body?.spreadsheet?.table;
	let ass = wb?.content['automatic-styles']?.style;
	if (!Sheets || !sheets || !ass) return;
	if (!Array.isArray(sheets)) sheets = [sheets];
	sheets.forEach(sheet => setSheetHidden(sheet, ass, Sheets.find(s => s.name === sheet.name)));
}
function writeOdsStayStyles(o, wb, opts) {
	let content = wb?.content;
	if (!content) return null;
	let target = {
		'font-face-decls': writeOdsFonts,
		'automatic-styles': writeOdsStyles,
	};
	let nfs = {};
	setBookHidden(wb);
	for (let n in target) {
		target[n](o, content[n], n);
	}
	return nfs;
}
function makeOdsStyles(wb, opts) {
	let styles = wb?.styles;
	if (!styles) return '';
	let target = {
		'font-face-decls': writeOdsFonts,
		'styles': writeOdsStyles,
		'automatic-styles': writeOdsStyles,
		'master-styles': writeOdsStyles,
	};
	let o = [];
	for (let n in target) {
		target[n](o, styles[n], n);
	}
	return o.join('');
}
const ODS_STYLE_ATTRS = [
	'name', 'family', 'display-name',
	'parent-style-name', 'next-style-name',
	'master-page-name', 'data-style-name',
];
const ODS_STYLE_SUBS = [
	'table-cell-properties', 'table-column-properties', 'table-row-properties',
	'table-properties',
	'paragraph-properties', 'text-properties', 'graphic-properties', 
];
const ODS_STYLE_SUBS_FUNC = {
	'table-properties': makeTableProperty,
};
const ODS_FONTS_PREFIX = [
	'background-color', 'color', 'country', 'language',
	'font-family', 'font-size', 'font-style', 'font-variant', 'font-weight',
	'wrap-option', 'text-shadow', 'text-align', 'break-before',
	'padding', 'padding-top', 'padding-bottom', 'padding-left', 'padding-right',
	'margin', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right',
	'border', 'border-top', 'border-bottom', 'border-left', 'border-right',
	'min-height', 'page-width', 'page-height',
];
const ODS_SVG_PREFIX = [
	'stroke-color', 'viewBox', 'd',
];
const ODS_DRAW_PREFIX = [
	'fill', 'fill-color', 'stroke',
	'shadow', 'shadow-offset-x', 'shadow-offset-y',
	'marker-start', 'marker-start-width', 'marker-start-center',
	'auto-grow-height', 'auto-grow-width',
];
const ODS_LOEXT_PREFIX = [
	'opacity',
	'blank-width-char', 'max-blank-integer-digits',
	'vertical-justify',
	'char-complex-color', 'theme-type', 'color-type',
];
const ODS_STYLE_PREFIX = [
	'name', 'volatile', 'map', 'condition',
	'apply-style-name', 'text-properties',
];
const ODS_TEXT_PREFIX = [
	'p', 's', 'title', 'sheet-name', 'page-number', 'page-count',
	'date', 'date-value', 'time', 'time-value',
];
const ODS_CCS3T_PREFIX = [
	'text-justify',
];
const ODS_PREFIXES = {
	'fo': ODS_FONTS_PREFIX,
	'svg': ODS_SVG_PREFIX,
	'draw': ODS_DRAW_PREFIX,
	'loext': ODS_LOEXT_PREFIX,
	'text': ODS_TEXT_PREFIX,
	'css3t': ODS_CCS3T_PREFIX,
};
const ODS_NUMBER_PREFIXES = {
	'style': ODS_STYLE_PREFIX,
	'loext': ODS_LOEXT_PREFIX,
	'fo': ODS_FONTS_PREFIX,
};
const ODS_MARKER_PREFIXES = {
	'svg': ODS_SVG_PREFIX,
};
function getPrefix(obj, def, n) {
	for (let pre in obj) {
		if (obj[pre].includes(n)) return pre;
	}
	return def;
}
function getOdsPrefix(n) {
	return getPrefix(ODS_PREFIXES, 'style', n);
}
function getOdsNumberPrefix(n) {
	return getPrefix(ODS_NUMBER_PREFIXES, 'number', n);
}
function getOdsMarkerPrefix(n) {
	return getPrefix(ODS_MARKER_PREFIXES, 'draw', n);
}
function getOdsThemePrefix(n) {
	return 'loext';
}
function getOdsFontPrefix(n) {
	return getPrefix({
		'svg': ['font-family']
	}, 'style', n);
}
function getOdsAutomaticStylePrefix(n) {
	return getPrefix({
		'fo': ['break-before']
	}, 'style', n);
}
function getOdsTablePrefix(n) {
	return 'table';
}
function getOdsGradientPrefix(n) {
	return getPrefix({
		'loext': ['gradient-stop', 'color-type', 'color-value'],
		'svg': ['offset']
	}, 'draw', n);
}
function makeOdsTag(v, pro, fn) {
	let o = [];
	fn(o, v, pro);
	return o.join('');
}
function makeOdsStyle(v, pro) {
	return makeOdsTag(v, pro, writeOdsStyle);
}
function makeOdsNumberStyle(v, pro) {
	return makeOdsTag(v, pro, writeOdsNumberStyle);
}
function makeOdsMarkerStyle(v, pro) {
	return makeOdsTag(v, pro, writeOdsMarkerStyle);
}
function makeOdsThemeStyle(v, pro) {
	return makeOdsTag(v, pro, writeOdsThemeStyle);
}
function makeOdsPageLayout(v, pro) {
	return makeOdsTag(v, pro, writeOdsPageLayout);
}
function makeOdsMasterPage(v, pro) {
	return makeOdsTag(v, pro, writeOdsMasterPage);
}
function makeOdsGradient(v, pro) {
	return makeOdsTag(v, pro, writeOdsGradient);
}
function makeTableProperty(v, pro) {
	return makeOdsTag(v, pro, writeOdsTableProperty);
}
function writeOdsStylesItem(item) {
	let vs = [];
	for (let n in item) {
		switch (n) {
		case 'default-style':
		case 'style':
			vs.push(makeOdsStyle(item[n], n));
			break;
		case 'number-style':
		case 'text-style':
		case 'date-style':
		case 'time-style':
		case 'boolean-style':
		case 'percentage-style':
		case 'currency-style':
			vs.push(makeOdsNumberStyle(item[n], n));
			break;
		case 'marker':
			vs.push(makeOdsMarkerStyle(item[n], n));
			break;
		case 'theme':
			vs.push(makeOdsThemeStyle(item[n], n));
			break;
		case 'page-layout':
			vs.push(makeOdsPageLayout(item[n], n));
			break;
		case 'master-page':
			vs.push(makeOdsMasterPage(item[n], n));
			break;
		case 'gradient':
			vs.push(makeOdsGradient(item[n], n));
			break;
		default:
			console.warn('unknown property ' + n);
			break;
		}
	}
	return vs.join('');
}
function writeOdsStyles(o, v, pro) {
	if (!v) return;
	o.push(makeXmlTag('office:' + pro, v, function(ss) {
		if (Array.isArray(ss)) {
			let ar = [];
			ss.forEach(function(item) {
				ar.push(writeOdsStylesItem(item));
			});
			return ar.join('');
		} else {
			return writeOdsStylesItem(ss);
		}
	}));
}
function writeOdsStyle(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('style:' + pro, d, function(d) {
			let s = '';
			for (let n in d) {
				if (ODS_STYLE_SUBS.includes(n)) {
					let fn = ODS_STYLE_SUBS_FUNC[n];
					if (typeof fn === 'function') {
						s += fn(d[n], n);
					} else {
						s += makeXmlTag('style:' + n, d[n], null, '?', getOdsPrefix);
					}
				}
			}
			return s;
		}, function(d) {
			let s = '';
			for (let n in d) {
				if (ODS_STYLE_ATTRS.includes(n)) {
					s += ` ${getOdsPrefix(n)}:${n}="${escapexml(d[n])}"`;
				}
			}
			return s;
		}));
	});
}
function writeOdsNumberStyle(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('number:' + pro, d, null, '?', getOdsNumberPrefix, function(n, val, pre) {
			let c = '', s = '';
			switch (n) {
			case 'text':
			case 'fill-character':
				if (!Array.isArray(val)) val = [val];
				val.forEach(function(v) {
					c += makeXmlTag(pre + n, v, function(v) {
						let s = typeof v === 'object' ? v.value : v;
						return s === 0 ? ' ' : s || '';
					}, function(v) {
						let s = '';
						if (typeof v === 'object') {
							for (let vn in v) {
								if (vn !== 'value') {
									let pre = getOdsNumberPrefix(vn);
									if (pre) pre += ':';
									s += ` ${pre}${vn}="${escapexml(v[vn])}"`;
								}
							}
						}
						return s;
					});
				});
				break;
			default:
				return null;
			}
			return {c:c, s:s};
		}));
	});
}
function writeOdsMarkerStyle(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('draw:' + pro, d, null, '?', getOdsMarkerPrefix));
	});
}
function writeOdsThemeStyle(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('loext:' + pro, d, null, '?', getOdsThemePrefix));
	});
}
function writeOdsFonts(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('office:' + pro, d, null, '?', getOdsFontPrefix));
	});
}
function writeOdsPageLayout(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('style:' + pro, d, null, '?', getOdsPrefix));
	});
}
function writeOdsMasterPage(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('style:' + pro, d, null, '?', getOdsPrefix));
	});
}
function writeOdsCalculation(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('table:' + pro, d, null, '?', getOdsTablePrefix));
	});
}
function writeOdsGradient(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('draw:' + pro, d, null, '?', getOdsGradientPrefix));
	});
}
function writeOdsTableProperty(o, v, pro) {
	if (!v) return;
	else if (!Array.isArray(v)) v = [v];
	v.forEach(function(d) {
		o.push(makeXmlTag('style:' + pro, d, null, '?', function(n) {
			return getPrefix({
				'table': ['display'],
			}, 'style', n);			
		}));
	});
}
