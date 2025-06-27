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
	if (typeof ar !== 'object') ar = [];
	else if (!Array.isArray(ar)) ar = [ar];
	addAlterContent(ar, draw?.AlternateContent);
	if (ar.length < 1) return null;
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
	let d = decode_range(ws['!ref']);
	if (d.e.r < iRow || d.e.c < iCol) {
		d.e.r = Math.max(d.e.r, iRow);
		d.e.c = Math.max(d.e.c, iCol);
		ws['!ref'] = encode_range(d);
	}
	let rPath = dfile.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
	let rStr = getzipstr(zip, rPath, true);
	let rels = rStr ? parse_xml(rStr)?.Relationship : null;
	if (rels) {
		let rs = {};
		if (!Array.isArray(rels)) rels = [rels];
		rels.forEach(r => {
			rs[r.Id] = getMedia(zip, wb, r)
		});
		ws['!drawRels'] = rs;
	}
	return draws;
}
// find same name shape from array
function findExistShape(ar, shape) {
	const n = shape?.sp?.nvSpPr?.cNvPr?.name;
	return n && ar.find(a => a?.sp?.nvSpPr?.cNvPr?.name === n);
}
function addOrExShape(ar, shape) {
	let found = findExistShape(ar, shape);
	if (!found) ar.push(shape);
	else extendObj(found, shape);
}
function addAlterContent(ar, alt) {
	if (!alt) return;
	if (!Array.isArray(alt)) alt = [alt];
	alt.forEach(a => {
		for (let n in a) {
			let o = a[n], c;
			if (typeof o === 'object' && (c = o.twoCellAnchor)) {
				if (typeof c === 'object')
					addOrExShape(ar, c);
			}
		}
	});
}
function getMedia(zip, wb, rel) {
	let media = wb['$media'];
	if (!media) media = wb['$media'] = {};
	let t = rel.Target;
	let m = media[t];
	if (!m) {
		switch (rel.Type) {
		case RELS.IMG:
			m = getImageAsBase64(zip, t.replace('..', 'xl'));
			break;
		default:
			console.warn('Not implement rels type:', rel.Type);
			break;
		}
		media[t] = m;
	}
	return m;
}
function binaryStringToBase64(bstr) {
	const bytes = new Uint8Array(bstr.length);
	for (let i = 0; i < bstr.length; i++) {
		bytes[i] = bstr.charCodeAt(i);
	}
	return btoa(String.fromCharCode.apply(null, bytes));
}
