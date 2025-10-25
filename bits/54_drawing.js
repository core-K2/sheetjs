/* 20.5 DrawingML - SpreadsheetML Drawing */
/* 20.5.2.35 wsDr CT_Drawing */
function parse_drawing(data, rels/*:any*/) {
	if(!data) return "??";
	/*
	  Chartsheet Drawing:
	   - 20.5.2.35 wsDr CT_Drawing
	    - 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor
	     - 20.5.2.16 graphicFrame CT_GraphicalObjectFrame
	      - 20.1.2.2.16 graphic CT_GraphicalObject
	       - 20.1.2.2.17 graphicData CT_GraphicalObjectData
          - chart reference
	   the actual type is based on the URI of the graphicData
		TODO: handle embedded charts and other types of graphics
	*/
	var id = (data.match(/<c:chart [^<>]*r:id="([^<>"]*)"/)||["",""])[1];

	return rels['!id'][id].Target;
}

/**
 * parse drawings core-K2 extended
 * @param {object} zip 
 * @param {string} dfile
 * @param {object} ws
 * @param {object} wb
 * @param {object} styles
 * @param {object} opts 
 */
function parseDrawings(zip, dfile, ws, wb, styles, opts) {
	const xmlOpts = {
		asSeqArray: [/^a:path$/],
	};
	let draw = parse_xml(getzipdata(zip, dfile, true), xmlOpts);
	let ar = draw?.twoCellAnchor;
	let ar2 = draw?.absoluteAnchor;
	if (typeof ar !== 'object') ar = [];
	else if (!Array.isArray(ar)) ar = [ar];
	if (typeof ar2 !== 'object') ar2 = [];
	else if (!Array.isArray(ar2)) ar2 = [ar2];
	addAlterContent(ar, draw?.AlternateContent);
	if (ar.length < 1 && ar2.length < 1) return null;
	let draws = {};
	let dss = styles?.Draws;
	if (!dss) {
		styles.Draws = dss = [];
	}
	let iRow = 0, iCol = 0;
	ar.forEach(a => {
		let from = a.from;
		if (!from) return;
		iRow = Math.max(a.to.row, iRow);
		iCol = Math.max(a.to.col, iCol);
		let cn = encode_col(from.col) + (from.row + 1);
		let sp = a.sp;
		if (sp?.style) {
			sp.si = getOrAddObject(dss, sp.style);
			delete sp.style;
		}
		let d = draws[cn];
		if (d) {
			if (!Array.isArray(d)) d = [d];
			addOrExShape(d, a);
			a = d;
		}
		draws[cn] = a;
	});
	const cn = 'A1';
	ar2.forEach(a => {
		let sp = a.sp;
		if (sp?.style) {
			sp.si = getOrAddObject(dss, sp.style);
			delete sp.style;
		}
		const pos = a.pos;
		const ext = a.ext;
		if (pos && ext) {
			if (!a.from) {
				a.from = {
					col: 0,
					row: 0,
					colOff: pos.x,
					rowOff: pos.y
				};
			}
			if (!a.to) {
				a.to = {
					col: 0,
					row: 0,
					colOff: pos.x + ext.cx,
					rowOff: pos.y + ext.cy
				};
			}
		}
		let d = draws[cn];
		if (d) {
			if (!Array.isArray(d)) d = [d];
			addOrExShape(d, a);
			a = d;
		}
		draws[cn] = a;
	});
	let d = decode_range(ws['!ref']);
	if (d.e.r < iRow || d.e.c < iCol) {
		d.e.r = Math.max(d.e.r, iRow);
		d.e.c = Math.max(d.e.c, iCol);
		ws['!ref'] = encode_range(d);
	}
	const rels = getRels(zip, dfile);
	if (rels) {
		const rs = getRelsObject(zip, wb, rels);
		for (let n in rs) {
			const r = rs[n];
			if (r?.diagramData) {
				const rid = r?.diagramData?.extLst?.ext?.dataModelExt?.relId;
				if (rid) {
					const o = rs[rid];
					if (o) {
						for (let m in draws) {
							const d = draws[m];
							if (d?.graphicFrame?.graphic?.graphicData?.relIds?.dm === n) {
								d.grpSp = o?.diagramDrawing?.spTree;
								break;
							}
						}
					}
				}
			}
		}
		ws['!drawRels'] = rs;
	}
	return draws;
}
// get path referrences
function getRels(zip, path) {
	const rStr = getzipstr(zip, convertToRelPath(path), true);
	if (rStr) {
		const rels = parse_xml(rStr)?.Relationship;
		return Array.isArray(rels) ? rels : [rels];
	}
	return null;
}
// get referrnce objects
function getRelsObject(zip, wb, rels) {
	const rs = {};
	rels.forEach(r => {
		rs[r.Id] = getMedia(zip, wb, r);
	});
	return rs;
}
// find same name shape from array
function findExistShape(ar, shape) {
	const n = shape?.sp?.nvSpPr?.cNvPr?.name;
	return n && ar.find(a => a?.sp?.nvSpPr?.cNvPr?.name === n);
}
function addOrExShape(ar, shape, extOnly) {
	let found = findExistShape(ar, shape);
	if (!found) {
		// const f = shape.from, t = shape.to;
		// if (f && t && (f.col || f.colOff || t.col || t.rowOff)) ar.push(shape);
		if (!extOnly) ar.push(shape);
	} else extendObj(found, shape);
}
function addAlterContent(ar, alt) {
	if (!alt) return;
	if (!Array.isArray(alt)) alt = [alt];
	alt.forEach(a => {
		for (let n in a) {
			let o = a[n], c;
			if (typeof o === 'object' && (c = o.twoCellAnchor)) {
				if (typeof c === 'object') {
					c.$type = n;
					addOrExShape(ar, c, true);
				}
			}
		}
	});
}
function findRelsType(type) {
	for (let n in RELS) {
		const v = RELS[n];
		if (Array.isArray(v)) {
			const found = v.find(x => x === type);
			if (found) return found;
		} else if (v === type) {
			return v;
		}
	}
	return null;
}
function getMedia(zip, wb, rel) {
	let media = wb['$media'];
	if (!media) media = wb['$media'] = {};
	const t = rel.Target;
	let m = media[t];
	if (!m) {
		try {
			const type = rel.Type;
			const path = t.replace('..', 'xl');
			switch (type) {
			case RELS.IMG:
				m = getImageAsBase64(zip, path);
				break;
			default:
				if (findRelsType(type)) {
					m = {};
					const obj = m[type.split('/').at(-1)] = parse_xml(getzipdata(zip, path, true));
					switch (type) {
					case RELS.CHART:
						const rels = getRels(zip, path);
						if (rels) m.$rels = getRelsObject(zip, wb, rels);
						break;
					}
				} else {
					console.warn('Not implement rels type:', rel.Type);
				}
				break;
			}
		} catch (e) {
			m = e;
		}
		media[t] = m;
	}
	return m;
}

function analyzeVmlDrawing(ws) {
	const vml = ws['!vml']?.[ws['!legrel']];
	let shape;
	if (!vml || !(shape = vml.shape)) return;
	let draws = ws['!drawings'];
	if (!draws) draws = ws['!drawings'] = {};
	if (!Array.isArray(shape)) shape = [shape];
	const getVml = (spid, id) => shape.find(o => spid && (o.spid === spid || o.id === spid) || id && o.id === id);
	const pixelToEmu = v => v * 914400 / 96;
	const setCell = (obj, ar, idx) => {
		for (let i = 0; i < 4; i++) {
			const v = Number(ar[idx + i].trim());
			switch (i) {
			case 0:	obj.col = v;	break;
			case 1:	obj.colOff = pixelToEmu(v);	break;
			case 2:	obj.row = v;	break;
			case 3:	obj.rowOff = pixelToEmu(v);	break;
			}
		}
	};
	const setCellPosition = (d, clid) => {
		const anchor = clid?.Anchor;
		if (!anchor) return false;
		const ar = anchor.trim().split(',');
		setCell(d.from, ar, 0);
		setCell(d.to, ar, 4);
		let c = clid.Column, r = clid.Row;
		if (c != null && r != null) {
			d.link = {col: c, row:r};
		}
		return true;
	};
	const addDraw = (draws, d, chk) => {
		const pos = d.link || d.from;
		const cn = encode_col(pos.col) + (pos.row + 1);
		let draw = draws[cn];
		if (!draw) {
			draws[cn] = d;
		} else {
			if (chk && chk(draw, d)) return false;
			if (Array.isArray(draw)) {
				draw.push(d);
			} else {
				draws[cn] = new Array(draw, d);
			}
		}
		return true;
	};
	for (let n in draws) {
		let dr = draws[n];
		if (!Array.isArray(dr)) dr = [dr];
		const isTop = n === 'A1';
		const move = [];
		dr.forEach((d, i) => {
			const cNvPr = d?.sp?.nvSpPr?.cNvPr;
			if (!cNvPr) return;
			let spid;
			let ext = cNvPr.extLst?.ext;
			if (typeof ext === 'object') {
				if (!Array.isArray(ext)) ext = [ext];
				ext.find(ex => {
					spid = ex.compatExt?.spid;
					return !!spid;
				}, this);
			}
			const vml = getVml(spid, cNvPr.name);
			if (vml) {
				d._vml = vml;
				vml.$found = true;
				if (isTop) {
					const from = d.from, to = d.to;
					if (!from || !to || from.col || from.colOff || from.row || from.rowOff || to.col || to.colOff || to.row || to.rowOff) return;
					if (!setCellPosition(d, vml.ClientData)) return;
					move.push(i);
				}
			}
		});
		const len = move.length;
		if (len > 0) {
			for (let i = len - 1; i >= 0; i--) {
				const del = move[i];
				addDraw(draws, dr[del]);
				dr.splice(del, 1);
			}
			if (dr.length < 1) {
				delete draws[n];
			}
		}
	}
	const toTwip = (v) => Math.round(pixelToEmu(getPixelSize(v)));
	const analyzeStyle = (d, style) => {
		if (!style) return;
		const ar = style.split(';');
		const xfrm = d.sp.spPr.xfrm;
		const $st = d.$style = {};
		ar.forEach(s => {
			const nv = s.split(':');
			if (nv.length !== 2) return;
			const n = nv[0].trim();
			const v = nv[1].trim();
			$st[n] = v;
			switch (n) {
			case 'margin-left': xfrm.off.x = toTwip(v); break;
			case 'margin-top': xfrm.off.y = toTwip(v); break;
			case 'width': xfrm.ext.cx = toTwip(v); break;
			case 'height': xfrm.ext.cy = toTwip(v); break;
			}
		});
	};
	const nearV = (a, b) => Math.abs(a - b) < 9525;
	const isSame = (a, b) => {
		const xf1 = a.sp.spPr.xfrm, xf2 = b.sp.spPr.xfrm;
		const e1 = xf1.ext, e2 = xf2.ext;
		const o1 = xf1.off, o2 = xf2.off;
		return nearV(e1.cx, e2.cx) && nearV(e1.cy, e2.cy) && nearV(o1.x, o2.x) && nearV(o1.y, o2.y);
	};
	const chkDraw = (draw, d) => {
		if (Array.isArray(draw)) {
			if (draw.find(a => isSame(a, d))) return true;
		} else if (isSame(draw, d)) return true;
		return false;
	};
	const getStyle = (s, cd) => {
		let style = {};
		switch (cd?.TextVAlign) {
		case 'Center': style.anchor = 'ctr'; break;
		case 'Top': style.anchor = 't'; break;
		case 'Bottom': style.anchor = 'b'; break;
		}
		if (s) {
			let ar = s.split(';');
			ar.forEach(s => {
				let nv = s.split(':')
				if (nv.length === 2) {
					switch (nv[0].trim()) {
					case 'text-align':
						let v;
						switch (nv[1].trim()) {
						case 'start':
						case 'left':	v = 'l'; break;
						case 'end':
						case 'right':	v = 'r'; break;
						case 'justify':
						case 'center':	v = 'ctr'; break;
						}
						style.algn = v;
						break;
					}
				}
			});
		}
		return style;
	};
	const analyzeText = (d, txt) => {
		let div;
		if (!txt || !(div = txt.div)) return;
		const font = div.font;
		if (font) {
			const style = getStyle(div.style, d._vml?.ClientData);
			let t, rPr = {};
			if (typeof font === 'object') {
				for (let n in font) {
					let v = font[n];
					switch (n) {
					case 'value':
						t = v;
						break;
					case 'size':
						rPr.sz = v * 5;
						break;
					case 'color':
						rPr.solidFill = {
							srgbClr: {val: v.substring(1)}
						};
						break;
					case 'face':
						rPr.ea = {
							typeface: v
						};
						break;
					}
				}
			} else {
				t = font;
			}
			const txBody = d.sp.txBody = {
				p: {
					r: {t}
				}
			};
			if (style.anchor) {
				txBody.bodyPr = {
					anchor: style.anchor
				};
			}
			const p = txBody.p;
			if (style.algn) {
				p.pPr = {
					algn: style.algn
				};
			}
			if (Object.keys(rPr).length > 0) {
				p.r.rPr = rPr;
			}
		}
	};
	shape.forEach(vml => {
		if (vml.$found) return;
		let d = {
			_vml: vml,
			from: {},
			to: {},
			sp: {
				spPr: {
					xfrm: {
						ext: {},
						off: {},
					},
				},
			},
		};
		if (!setCellPosition(d, vml.ClientData)) return;
		analyzeStyle(d, vml.style);
		if (!addDraw(draws, d, chkDraw)) return;
		analyzeText(d, vml.textbox);
	});
}
