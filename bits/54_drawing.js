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
 * @param {string} data 
 * @param {object} wb 
 * @param {object} opts 
 */
function parseDrawings(data, ws, styles, opts) {
	let draw = parse_xml(data);
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
		draws[cn] = a;
	});
	let d = decode_range(ws['!ref']);
	if (d.e.r < iRow || d.e.c < iCol) {
		d.e.r = Math.max(d.e.r, iRow);
		d.e.c = Math.max(d.e.c, iCol);
		ws['!ref'] = encode_range(d);
	}
	return draws;
}