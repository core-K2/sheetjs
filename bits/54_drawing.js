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
function parseDrawings(data, styles, opts) {
	let draw = parse_xml(data);
	let ar = draw?.twoCellAnchor;
	if (!Array.isArray(ar)) return null;
	let draws = {};
	let dss = styles?.Draws;
	if (!dss) {
		styles.Draws = dss = [];
	}
	ar.forEach(a => {
		let from = a.from;
		if (!from) return;
		let cn = encode_col(from.col) + (from.row + 1);
		let sp = a.sp;
		if (sp?.style) {
			sp.si = getOrAddObject(dss, sp.style);
			delete sp.style;
		}
		draws[cn] = a;
	});
	return draws;
}