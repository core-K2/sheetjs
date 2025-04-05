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
	let draw = parse_xml(getzipdata(zip, dfile, true));
	let ar = draw?.twoCellAnchor;
	if (typeof ar !== 'object') return null;
	if (!Array.isArray(ar)) ar = [ar];
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
			d.push(a);
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
function getMedia(zip, wb, rel) {
	let media = wb['$media'];
	if (!media) media = wb['$media'] = {};
	let t = rel.Target;
	let m = media[t];
	if (!m) {
		switch (rel.Type) {
		case RELS.IMG:
			m = analyzeImageData(getzipdata(zip, t.replace('..', 'xl'), true));
			m.ext = t.split('.').at(-1);
			m.data = `data:image/${m.type};base64,${m.b64}`;
			delete m.b64;
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
function analyzeImageData(bstr) {
	const bytes = new Uint8Array(bstr.length);
	for (let i = 0; i < bstr.length; i++) {
		bytes[i] = bstr.charCodeAt(i);
	}
	let t, w, h;

	// PNG
	if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
		t = 'png';
		w = (bytes[16] << 24) + (bytes[17] << 16) + (bytes[18] << 8) + bytes[19];
		h = (bytes[20] << 24) + (bytes[21] << 16) + (bytes[22] << 8) + bytes[23];
	}

	// JPEG
	if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
		for (let i = 0; i < bytes.length - 9; i++) {
			if (bytes[i] === 0xFF && (bytes[i + 1] === 0xC0 || bytes[i + 1] === 0xC1 || bytes[i + 1] === 0xC2)) {
				t = 'jpeg';
				h = (bytes[i + 5] << 8) + bytes[i + 6];
				w = (bytes[i + 7] << 8) + bytes[i + 8];
				break;
			}
		}
	}

	// GIF
	if (String.fromCharCode(bytes[0], bytes[1], bytes[2]) === 'GIF') {
		t = 'gif';
		w = bytes[6] + (bytes[7] << 8);
		h = bytes[8] + (bytes[9] << 8);
	}

	// BMP
	if (String.fromCharCode(bytes[0], bytes[1]) === 'BM') {
		t = 'bmp';
		w = bytes[18] + (bytes[19] << 8) + (bytes[20] << 16) + (bytes[21] << 24);
		h = bytes[22] + (bytes[23] << 8) + (bytes[24] << 16) + (bytes[25] << 24);
	}

	if (!t) {
		throw new Error('Unsupported image format', );
	}
	return {
		type: t,
		w: w,
		h: h,
		b64: btoa(String.fromCharCode.apply(null, bytes))
	};
}
