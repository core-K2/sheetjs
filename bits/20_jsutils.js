function keys(o/*:any*/)/*:Array<any>*/ {
	var ks = Object.keys(o), o2 = [];
	for(var i = 0; i < ks.length; ++i) if(Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
	return o2;
}

function evert_key(obj/*:any*/, key/*:string*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) if(o[obj[K[i]][key]] == null) o[obj[K[i]][key]] = K[i];
	return o;
}

function evert(obj/*:any*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
	return o;
}

function evert_num(obj/*:any*/)/*:EvertNumType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i],10);
	return o;
}

function evert_arr(obj/*:any*/)/*:EvertArrType*/ {
	var o/*:EvertArrType*/ = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) {
		if(o[obj[K[i]]] == null) o[obj[K[i]]] = [];
		o[obj[K[i]]].push(K[i]);
	}
	return o;
}

var dnthresh  = /*#__PURE__*/Date.UTC(1899, 11, 30, 0, 0, 0); // -2209161600000
var dnthresh1 = /*#__PURE__*/Date.UTC(1899, 11, 31, 0, 0, 0); // -2209075200000
var dnthresh2 = /*#__PURE__*/Date.UTC(1904, 0, 1, 0, 0, 0); // -2209075200000
function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = /*#__PURE__*/v.getTime();
	var res = (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	if(date1904) { res -= 1462; return res < -1402 ? res - 1 : res; }
	return res < 60 ? res - 1 : res;
}
function numdate(v/*:number*/)/*:Date|number*/ {
	if(v >= 60 && v < 61) return v;
	var out = new Date();
	out.setTime((v>60 ? v : (v+1)) * 24 * 60 * 60 * 1000 + dnthresh);
	return out;
}

/* ISO 8601 Duration */
function parse_isodur(s) {
	var sec = 0, mt = 0, time = false;
	var m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
	if(!m) throw new Error("|" + s + "| is not an ISO8601 Duration");
	for(var i = 1; i != m.length; ++i) {
		if(!m[i]) continue;
		mt = 1;
		if(i > 3) time = true;
		switch(m[i].slice(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].slice(m[i].length-1));
			case 'D': mt *= 24;
				/* falls through */
			case 'H': mt *= 60;
				/* falls through */
			case 'M':
				if(!time) throw new Error("Unsupported ISO Duration Field: M");
				else mt *= 60;
				/* falls through */
			case 'S': break;
		}
		sec += mt * parseInt(m[i], 10);
	}
	return sec;
}

/* Blame https://bugs.chromium.org/p/v8/issues/detail?id=7863 for the regexide */
var pdre1 = /^(\d+):(\d+)(:\d+)?(\.\d+)?$/; // HH:MM[:SS[.UUU]]
var pdre2 = /^(\d+)-(\d+)-(\d+)$/; // YYYY-mm-dd
var pdre3 = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)?(\.\d+)?$/; // YYYY-mm-dd(T or space)HH:MM[:SS[.UUU]], sans "Z"
/* parses a date string as a UTC date */
function parseDate(str/*:string*/, date1904/*:boolean*/)/*:Date*/ {
	if(str instanceof Date) return str;
	var m = str.match(pdre1);
	if(m) return new Date((date1904 ? dnthresh2 : dnthresh1) + ((parseInt(m[1], 10)*60 + parseInt(m[2], 10))*60 + (m[3] ? parseInt(m[3].slice(1), 10) : 0))*1000 + (m[4] ? parseInt((m[4]+"000").slice(1,4), 10) : 0));
	m = str.match(pdre2);
	if(m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3], 0, 0, 0, 0));
	/* TODO: 1900-02-29T00:00:00.000 should return a flag to treat as a date code (affects xlml) */
	m = str.match(pdre3);
	if(m) return new Date(Date.UTC(+m[1], +m[2]-1, +m[3], +m[4], +m[5], ((m[6] && parseInt(m[6].slice(1), 10))|| 0), ((m[7] && parseInt((m[7] + "0000").slice(1,4), 10))||0)));
	var d = new Date(str);
	return d;
}

function cc2str(arr/*:Array<number>*/, debomit)/*:string*/ {
	if(has_buf && Buffer.isBuffer(arr)) {
		if(debomit && buf_utf16le) {
			// TODO: temporary patch
			if(arr[0] == 0xFF && arr[1] == 0xFE) return utf8write(arr.slice(2).toString("utf16le"));
			if(arr[1] == 0xFE && arr[2] == 0xFF) return utf8write(utf16beread(arr.slice(2).toString("binary")));
		}
		return arr.toString("binary");
	}

	if(typeof TextDecoder !== "undefined") try {
		if(debomit) {
			if(arr[0] == 0xFF && arr[1] == 0xFE) return utf8write(new TextDecoder("utf-16le").decode(arr.slice(2)));
			if(arr[0] == 0xFE && arr[1] == 0xFF) return utf8write(new TextDecoder("utf-16be").decode(arr.slice(2)));
		}
		var rev = {
			"\u20ac": "\x80", "\u201a": "\x82", "\u0192": "\x83", "\u201e": "\x84",
			"\u2026": "\x85", "\u2020": "\x86", "\u2021": "\x87", "\u02c6": "\x88",
			"\u2030": "\x89", "\u0160": "\x8a", "\u2039": "\x8b", "\u0152": "\x8c",
			"\u017d": "\x8e", "\u2018": "\x91", "\u2019": "\x92", "\u201c": "\x93",
			"\u201d": "\x94", "\u2022": "\x95", "\u2013": "\x96", "\u2014": "\x97",
			"\u02dc": "\x98", "\u2122": "\x99", "\u0161": "\x9a", "\u203a": "\x9b",
			"\u0153": "\x9c", "\u017e": "\x9e", "\u0178": "\x9f"
		};
		if(Array.isArray(arr)) arr = new Uint8Array(arr);
		return new TextDecoder("latin1").decode(arr).replace(/[€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/g, function(c) { return rev[c] || c; });
	} catch(e) {}

	var o = [], i = 0;
	// this cascade is for the browsers and runtimes of antiquity (and for modern runtimes that lack TextEncoder)
	try {
		for(i = 0; i < arr.length - 65536; i+=65536) o.push(String.fromCharCode.apply(0, arr.slice(i, i + 65536)));
		o.push(String.fromCharCode.apply(0, arr.slice(i)));
	} catch(e) { try {
			for(; i < arr.length - 16384; i+=16384) o.push(String.fromCharCode.apply(0, arr.slice(i, i + 16384)));
			o.push(String.fromCharCode.apply(0, arr.slice(i)));
		} catch(e) {
			for(; i != arr.length; ++i) o.push(String.fromCharCode(arr[i]));
		}
	}
	return o.join("");
}

function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined' && !Array.isArray(o)) return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || o == null) return o;
	if(o instanceof Date) return new Date(o.getTime());
	var out = {};
	for(var k in o) if(Object.prototype.hasOwnProperty.call(o, k)) out[k] = dup(o[k]);
	return out;
}

function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }

/* TODO: stress test */
function fuzzynum(s/*:string*/)/*:number*/ {
	var v/*:number*/ = Number(s);
	if(!isNaN(v)) return isFinite(v) ? v : NaN;
	if(!/\d/.test(s)) return v;
	var wt = 1;
	var ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
	if(!isNaN(v = Number(ss))) return v / wt;
	ss = ss.replace(/[(]([^()]*)[)]/,function($$, $1) { wt = -wt; return $1;});
	if(!isNaN(v = Number(ss))) return v / wt;
	return v;
}

/* NOTE: Chrome rejects bare times like 1:23 PM */
var FDRE1 = /^(0?\d|1[0-2])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))\s+([ap])m?$/;
var FDRE2 = /^([01]?\d|2[0-3])(?:|:([0-5]?\d)(?:|(\.\d+)(?:|:([0-5]?\d))|:([0-5]?\d)(|\.\d+)))$/;
var FDISO = /^(\d+)-(\d+)-(\d+)[T ](\d+):(\d+)(:\d+)(\.\d+)?[Z]?$/; // YYYY-mm-dd(T or space)HH:MM:SS[.UUU][Z]

/* TODO: 1904 adjustment */
var utc_append_works = new Date("6/9/69 00:00 UTC").valueOf() == -17798400000;
function fuzzytime1(M) /*:Date*/ {
	if(!M[2]) return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), 0, 0, 0));
	if(M[3]) {
			if(M[4]) return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], +M[4], parseFloat(M[3])*1000));
			else return new Date(Date.UTC(1899,11,31,(M[7] == "p" ? 12 : 0), +M[1], +M[2], parseFloat(M[3])*1000));
	}
	else if(M[5]) return new Date(Date.UTC(1899,11,31, (+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], +M[5], M[6] ? parseFloat(M[6]) * 1000 : 0));
	else return new Date(Date.UTC(1899,11,31,(+M[1]%12) + (M[7] == "p" ? 12 : 0), +M[2], 0, 0));
}
function fuzzytime2(M) /*:Date*/ {
	if(!M[2]) return new Date(Date.UTC(1899,11,31,+M[1], 0, 0, 0));
	if(M[3]) {
			if(M[4]) return new Date(Date.UTC(1899,11,31,+M[1], +M[2], +M[4], parseFloat(M[3])*1000));
			else return new Date(Date.UTC(1899,11,31,0, +M[1], +M[2], parseFloat(M[3])*1000));
	}
	else if(M[5]) return new Date(Date.UTC(1899,11,31, +M[1], +M[2], +M[5], M[6] ? parseFloat(M[6]) * 1000 : 0));
	else return new Date(Date.UTC(1899,11,31,+M[1], +M[2], 0, 0));
}
var lower_months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
function fuzzydate(s/*:string*/)/*:Date*/ {
	// See issue 2863 -- this is technically not supported in Excel but is otherwise useful
	if(FDISO.test(s)) return s.indexOf("Z") == -1 ? local_to_utc(new Date(s)) : new Date(s);
	var lower = s.toLowerCase();
	var lnos = lower.replace(/\s+/g, " ").trim();
	var M = lnos.match(FDRE1);
	if(M) return fuzzytime1(M);
	M = lnos.match(FDRE2);
	if(M) return fuzzytime2(M);
	M = lnos.match(pdre3);
	if(M) return new Date(Date.UTC(+M[1], +M[2]-1, +M[3], +M[4], +M[5], ((M[6] && parseInt(M[6].slice(1), 10))|| 0), ((M[7] && parseInt((M[7] + "0000").slice(1,4), 10))||0)));
	var o = new Date(utc_append_works && s.indexOf("UTC") == -1 ? s + " UTC": s), n = new Date(NaN);
	var y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	if(lower.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
		lower = lower.replace(/[^a-z]/g,"").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/,"");
		if(lower.length > 3 && lower_months.indexOf(lower) == -1) return n;
	} else if(lower.replace(/[ap]m?/, "").match(/[a-z]/)) return n;
	if(y < 0 || y > 8099 || s.match(/[^-0-9:,\/\\\ ]/)) return n;
	return o;
}

var split_regex = /*#__PURE__*/(function() {
	var safe_split_regex = "abacaba".split(/(:?b)/i).length == 5;
	return function split_regex(str/*:string*/, re, def/*:string*/)/*:Array<string>*/ {
		if(safe_split_regex || typeof re == "string") return str.split(re);
		var p = str.split(re), o = [p[0]];
		for(var i = 1; i < p.length; ++i) { o.push(def); o.push(p[i]); }
		return o;
	};
})();

function utc_to_local(utc) {
	return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate(), utc.getUTCHours(), utc.getUTCMinutes(), utc.getUTCSeconds(), utc.getUTCMilliseconds());
}
function local_to_utc(local) {
	return new Date(Date.UTC(local.getFullYear(), local.getMonth(), local.getDate(), local.getHours(), local.getMinutes(), local.getSeconds(), local.getMilliseconds()));
}

function remove_doctype(str) {
	var preamble = str.slice(0, 1024);
	var si = preamble.indexOf("<!DOCTYPE");
	if(si == -1) return str;
	var m = str.match(/<[\w]/);
	if(!m) return str;
	return str.slice(0, si) + str.slice(m.index);
}

/* str.match(/<!--[\s\S]*?-->/g) --> str_match_ng(str, "<!--", "-->") */
function str_match_ng(str, s, e) {
  var out = [];

  var si = str.indexOf(s);
  while(si > -1) {
    var ei = str.indexOf(e, si + s.length);
		if(ei == -1) break;

		out.push(str.slice(si, ei + e.length));
		si = str.indexOf(s, ei + e.length);
	}

  return out.length > 0 ? out : null;
}

/* str.replace(/<!--[\s\S]*?-->/g, "") --> str_remove_ng(str, "<!--", "-->") */
function str_remove_ng(str, s, e) {
  var out = [], last = 0;

  var si = str.indexOf(s);
	if(si == -1) return str;
  while(si > -1) {
		out.push(str.slice(last, si));
    var ei = str.indexOf(e, si + s.length);
		if(ei == -1) break;

		if((si = str.indexOf(s, (last = ei + e.length))) == -1) out.push(str.slice(last));
	}

  return out.join("");
}

/* str.match(/<tag\b[^>]*?>([\s\S]*?)</tag>/) --> str_match_xml(str, "tag") */
var xml_boundary = { " ": 1, "\t": 1, "\r": 1, "\n": 1, ">": 1 };
function str_match_xml(str, tag) {
	var si = str.indexOf('<' + tag), w = tag.length + 1, L = str.length;
	while(si >= 0 && si <= L - w && !xml_boundary[str.charAt(si + w)]) si = str.indexOf('<' + tag, si+1);
	if(si === -1) return null;
	var sf = str.indexOf(">", si + tag.length);
	if(sf === -1) return null;
	var et = "</" + tag + ">";
	var ei = str.indexOf(et, sf);
	if(ei == -1) return null;
	return [str.slice(si, ei + et.length), str.slice(sf + 1, ei)];
}

/* str.match(/<(?:\w+:)?tag\b[^<>]*?>([\s\S]*?)<\/(?:\w+:)?tag>/) --> str_match_xml(str, "tag") */
var str_match_xml_ns = /*#__PURE__*/(function() {
	var str_match_xml_ns_cache = {};
	// got:get tag value as Object (core-K2 expansion)
	return function str_match_xml_ns(str, tag, got) {
		var res = str_match_xml_ns_cache[tag];
		if(!res) str_match_xml_ns_cache[tag] = res = [
			new RegExp('<(?:\\w+:)?'+tag+'\\b[^<>]*>', "g"),
			new RegExp('</(?:\\w+:)?'+tag+'>', "g")
		];
		res[0].lastIndex = res[1].lastIndex = 0;
		let m = res[0].exec(str);
		if(!m) return null;
		let si = m.index;
		let sf = res[0].lastIndex;
		let ei, ef;
		if (m[0].endsWith('/>')) {
			// none end tag case
			ei = ef = sf;
		} else {
			res[1].lastIndex = res[0].lastIndex;
			m = res[1].exec(str);
			if(!m) return null;
			ei = m.index;
			ef = res[1].lastIndex;
		}
		let s = str.slice(si, ef);
		if (got) {
			let obj = Xml.xmlStrToObject(s);
			if (obj?.body?.parsererror) {
				let x = Xml.xmlStrToObject(str);
				obj = x[tag];
			}
			if (typeof got === 'object') {
				got[tag] = obj;
			}
			return obj;
		}
		return [s, str.slice(sf, ei)];
	};
})();

/* str.match(/<(?:\w+:)?tag\b[^<>]*?>([\s\S]*?)<\/(?:\w+:)?tag>/g) --> str_match_xml_ns_g(str, "tag") */
var str_match_xml_ns_g = /*#__PURE__*/(function() {
	var str_match_xml_ns_cache = {};
	return function str_match_xml_ns(str, tag) {
		var out = [];
		var res = str_match_xml_ns_cache[tag];
		if(!res) str_match_xml_ns_cache[tag] = res = [
			new RegExp('<(?:\\w+:)?'+tag+'\\b[^<>]*>', "g"),
			new RegExp('</(?:\\w+:)?'+tag+'>', "g")
		];
		res[0].lastIndex = res[1].lastIndex = 0;
		var m;
		while((m = res[0].exec(str))) {
			var si = m.index;
			res[1].lastIndex = res[0].lastIndex;
			m = res[1].exec(str);
			if(!m) return null;
			var ef = res[1].lastIndex;
			out.push(str.slice(si, ef));
			res[0].lastIndex = res[1].lastIndex;
		}
		return out.length == 0 ? null : out;
	};
})();
var str_remove_xml_ns_g = /*#__PURE__*/(function() {
	var str_remove_xml_ns_cache = {};
	return function str_remove_xml_ns_g(str, tag) {
		var out = [];
		var res = str_remove_xml_ns_cache[tag];
		if(!res) str_remove_xml_ns_cache[tag] = res = [
			new RegExp('<(?:\\w+:)?'+tag+'\\b[^<>]*>', "g"),
			new RegExp('</(?:\\w+:)?'+tag+'>', "g")
		];
		res[0].lastIndex = res[1].lastIndex = 0;
		var m;
		var si = 0, ef = 0;
		while((m = res[0].exec(str))) {
			si = m.index;
			out.push(str.slice(ef, si));
			ef = si;
			res[1].lastIndex = res[0].lastIndex;
			m = res[1].exec(str);
			if(!m) return null;
			ef = res[1].lastIndex;
			res[0].lastIndex = res[1].lastIndex;
		}
		out.push(str.slice(ef));
		return out.length == 0 ? "" : out.join("");
	};
})();

/* str.match(/<(?:\w+:)?tag\b[^>]*?>([\s\S]*?)<\/(?:\w+:)?tag>/gi) --> str_match_xml_ns_ig(str, "tag") */
var str_match_xml_ig = /*#__PURE__*/(function() {
	var str_match_xml_ns_cache = {};
	return function str_match_xml_ns(str, tag) {
		var out = [];
		var res = str_match_xml_ns_cache[tag];
		if(!res) str_match_xml_ns_cache[tag] = res = [
			new RegExp('<'+tag+'\\b[^<>]*>', "ig"),
			new RegExp('</'+tag+'>', "ig")
		];
		res[0].lastIndex = res[1].lastIndex = 0;
		var m;
		while((m = res[0].exec(str))) {
			var si = m.index;
			res[1].lastIndex = res[0].lastIndex;
			m = res[1].exec(str);
			if(!m) return null;
			var ef = res[1].lastIndex;
			out.push(str.slice(si, ef));
			res[0].lastIndex = res[1].lastIndex;
		}
		return out.length == 0 ? null : out;
	};
})();

/**
 * XML proceccing (core-K2 expansion)
 */
var Xml = {
	parser: new DOMParser(),
	// options
	opts: {
		noXmlns: true,		// delete XML name space (xmlns)
		noNamePrefix: true,	// delete nodeName prefix (prefix:name=value -> name:value)
		propName: '',		// object property name
		textIsValue: true,	// not object when text node only case
		textAttrs: true,	// text value has attributes
		textPName: 'name',	// property name when text array convert to object
		textTName: 'type',	// property name of text data type 
		textVName: 'value',	// property name of text value 
		prefixAttr: '',		// prefix for attribute
		prefixText: '',		// prefix for text data
		asValue: 3,			// get value as bitmask (1:Number, 2:Boolean, 4:Date, 8:trim)
		convNames: null,	// convert name map
		asSeqArray: null,	// as sequence array object names
		asText: null,		// as text value object names
		convValues: null,	// convert value map
		tagTrap: null,		// tag trap function (xmlNode, parent, bSeqParent) => {}
	},
	_opts: [],
	pushOpts: function() {
		this._opts.push(Object.assign({}, this.opts));
	},
	popOpts: function() {
		if (this._opts.length > 0) {
			this.opts = this._opts.pop();
		}
	},
	// set options
	setOpts: function(v) {
		if (!v) return;
		this.pushOpts();
		if (typeof v === 'object') {
			for (let n in v) {
				this.setOptsValue(n, v[n]);
			}
		} else {
			for (let i = 0; i < arguments.length; i++) {
				let n = arguments[i];
				if (this.opts.hasOwnProperty(n)) {
					this.setOptsValue(n, arguments[++i]);
				}
			}
		}
	},
	// set option value
	setOptsValue: function(n, v) {
		if (['asSeqArray', 'asText'].includes(n)) {
			if (typeof v === 'string') v = v.split(',');
			let old = this.opts[n];
			if (old) {
				if (typeof old === 'string') old = old.split(',');
				v = old.concat(v);
			}
		}
		this.opts[n] = v;
	},
	// get property name
	getName: function(n) {
		let cns = this.opts.convNames;
		if (cns && cns[n]) return cns[n];
		if (this.opts.noXmlns && /^xmlns(\:.*|$)/.test(n)) {
			return '';
		} else if (this.opts.noNamePrefix) {
			n = n.split(':').at(-1);
		}
		return n;
	},
	isSeqArray: function(n) {
		let seq = this.opts.asSeqArray;
		if (!seq) return false;
		return seq.findIndex(function(s) {
			let re = s instanceof RegExp ? s : new RegExp(s);
			return re.test(n);
		}) >= 0;
	},
	isText: function(n) {
		let ar = this.opts.asText;
		return ar && ar.includes(n);
	},
	getText: function(node) {
		let v = this.opts.convValues?.[node.nodeName];
		if (v === undefined) {
			// nodeType 1:Element 3:text 8:comment
			switch (node.nodeType) {
			case 3:
				v = this.toValue(node.nodeValue, true);
				break;
			default:
				v = node.textContent || '';
				break;
			}
		}
		return v;
	},
	getAsText: function(node) {
		let t = '';
		node.childNodes.forEach(n => {
			t += this.getText(n);
		}, this);
		return t;
	},
	// convert value
	toValue: function(v, bTrim) {
		if (v) {
			const asV = this.opts.asValue;
			if (asV) {
				const vt = v.trim();
				if (bTrim && asV & 8) {
					v = vt;
				}
				if (asV & 2) {
					switch (vt.toLowerCase()) {
					case 'true': return true;
					case 'false': return false;
					}
				}
				if (asV & 1 && /^[+-]?\d+(\.\d+)?$/.test(v)) {
					return Number(v);
				}
				if (asV & 4) {
					let dt
					if (!isNaN((dt = new Date(v)).getTime())) {
						return dt;
					}
				}
			}
		}
		return v;
	},
	// get XMLDocument object
	getXmlDocument: function(v, mime) {
		return v instanceof XMLDocument ? v : this.parser.parseFromString(v, mime || 'application/xml');
	},
	// get JavaScript Object from XML(XMLDocument or string)
	getXmlAsObject: function(v, opts, mime) {
		this.setOpts(opts);
		return this.xmlToObject(this.getXmlDocument(v, mime));
	},
	// XML string to JavaScript Object
	xmlStrToObject: function(xmlStr, opts) {
		this.setOpts(opts);
		return this.xmlToObject(this.getXmlDocument(xmlStr).documentElement);
	},
	// XML node to JavaScript Object
	xmlToObject: function(xmlNode, parent, bSeqParent) {
		let fn = this.opts.tagTrap?.[xmlNode.nodeName];
		if (typeof fn === 'function') {
			let v = fn.apply(this, [xmlNode, parent, bSeqParent]);
			if (v !== undefined) return v;
		}
		if (this.isText(xmlNode.nodeName)) {
			return this.getAsText(xmlNode);
		}
		return this.parseNode(xmlNode, parent, bSeqParent);
	},
	// parse node object
	parseNode: function(xmlNode, parent, bSeqParent) {
		let obj = bSeqParent ? [] : {};
		let attrs = this.parseAttributes(xmlNode.attributes);
		// child node process
		let len = xmlNode.childNodes.length;
		for (let i = 0; i < len; i++) {
			let node = xmlNode.childNodes[i];
			let name = this.getName(node.nodeName);
			if (!name) {
				continue;
			}
			// nodeType 1:Element 3:text 8:comment
			if (node.nodeType === 3) {
				if (len === 1) {
					let v = this.toValue(node.nodeValue, true);
					if (this.opts.textIsValue) {
						if (this.opts.textAttrs && Object.keys(attrs).length > 0) {
							let o, n;
							if ((n = this.opts.textPName) && attrs.hasOwnProperty(n)) {
								if (parent) {
									parent[attrs[n]] = v;
									return;
								}
								o = {};
								o[attrs[n]] = v;
							} else {
								o = attrs;
								o[this.opts.textVName] = v;
							}
							return o;
						}
						return v;
					}
					obj[this.opts.prefixText + name] = v;
					break;
				}
				continue;
			}
			let val = this.xmlToObject(node, obj, this.isSeqArray(node.nodeName));
			if (val === undefined) {
				continue;
			}
			if (name !== node.localName) {
				val.ln = node.localName;
			}
			if (bSeqParent) {
				let o = {};
				o[name] = val;
				obj.push(o);
			} else if (!obj[name]) {
				obj[name] = val;
			} else {
				if (!Array.isArray(obj[name])) {
					obj[name] = [obj[name]];
				}
				obj[name].push(val);
			}
		}
		let pn = this.opts.propName;
		if (pn && attrs.hasOwnProperty(pn)) {
			let n = this.getName(attrs[pn]);
			for (let a in attrs) {
				if (a !== pn) {
					obj[a] = attrs[a];
				}
			}
			if (Object.keys(obj).length < 1) {
				obj = null;
			}
			if (parent) {
				parent[n] = obj;
				return;
			}
			let ret = {};
			ret[n] = obj;
			return ret;
		}
		if (obj) {
			Object.assign(obj, attrs);
		}
		return obj;
	},
	// parse attributes
	parseAttributes: function(attrs, obj) {
		obj = obj || {};
		if (attrs) {
			let prefix = this.opts.prefixAttr;
			for (let i = 0; i < attrs.length; i++) {
				let attr = attrs[i];
				let name = this.getName(attr.name);
				if (!name) continue;
				let n = prefix + name;
				if (obj.hasOwnProperty(n)) n = prefix + attr.name;
				obj[n] = this.toValue(attr.value);
			}
		}
		return obj;
	},
};

function parse_xml(str, xmlOpts) {
	str = xlml_normalize(utf8read(str));
	let xml = Xml.xmlStrToObject(str, xmlOpts);
	if (xmlOpts) Xml.popOpts();
	return xml;
}

function isEmpty(v) {
	if (v === null || v === undefined) {
		return true;
	} else if (v instanceof Array) {
		return v.length < 1;
	} else if (typeof v === 'object') {
		return Object.keys(v).length < 1;
	}
	return false;
}
function toCamelCase(str) {
	return str.replace(/[-_](\w)/g, function() {
		var v1 = arguments[1];
		return v1 ? v1.toUpperCase() : '';
	});
}
function toUpperCamelCase(str) {
	return toCamelCase(str).replace(/^[a-z]/, function(match) {
		return match.toUpperCase();
	});
}
function extendObject(d, s) {
	for (let n in s) {
		if (d[n] === undefined) d[n] = s[n];
	}
}
function extendObj(d, s) {
	if (typeof s === 'object') {
		for (let n in s) {
			let v = s[n];
			if (typeof v === 'object') {
				d[n] = extendObj(d[n] || (Array.isArray(v) ? [] : {}), v);
			} else if (!d.hasOwnProperty(n)) {
				d[n] = v;	
			}
		}
	}
	return d;
}
function mergeObject(...objs) {
	return objs.reduce((a, b) => ({...a, ...b}));
}
function cloneObject(obj) {
	// if (!obj || typeof obj !== 'object') return obj;
	// const o = Array.isArray(obj) ? [] : {};
	// for (const key in obj) {
	// 	if (obj.hasOwnProperty(key)) {
	// 		o[key] = cloneObject(obj[key]);
	// 	}
	// }
	// return o;
	return structuredClone(obj)
}
function applyObject(obj, v, deep) {
	if (v && typeof v === 'object') {
		if (!obj || !Object.keys(obj).length) {
			obj = cloneObject(v);
		} else {
			for (let n in v) {
				if (!obj.hasOwnProperty(n)) {
					obj[n] = v[n];
				} else if (deep && typeof v[n] === 'object') {
					applyObject(obj[n], v[n], deep);
				}
			}
		}
	}
	return obj;
}
function toBoolean(v) {
	if (typeof v === 'boolean') return v;
	if (isNaN(v)) {
		switch (v.toLowerCase()) {
		case 'f':
		case 'false':
			return false;
		default:
			return true;
		}
	}
	return Number(v) !== 0;
}
function toNumber(v, def) {
	if (typeof v === 'number') {
		return v;
	}
	let ret = parseFloat(('' + v).replace(/[^+\-0-9.]/g, ''));
	return isNaN(ret) ? def === undefined ? 0 : def : ret;
}
function toNumberInObject(o, props) {
	if (typeof props === 'string') props = props.split(',');
	props.forEach(p => {
		o[p] = toNumber(o[p]);
	});
	return o;	
}
function getPixelSize(v, u) {
	if (!v || typeof v === 'number') return v;
	let m = /(\d+(\.\d+)?)([^0-9.]*)?/.exec(v);
	if (!m) throw new Error(`invalid number format "${v}"`);
	let n = parseFloat(m[1]);
	let unit = m[3].toLowerCase();
	if (u && u === unit) {
		return n;
	}
	switch (unit) {
	case 'px':
		return n;
	case 'mm':
		n *= 10;
	case 'cm':
		n *= 96 / 2.54;
		break;
	case 'pt':
		n *= 96 / 72;
		break;
	case 'inch':
	case 'in':
		n *= 96;
		break;
	}
	return n.toFixed(4);
}

/**
 * convert to valid Date object
 * @param {any} v input value ('now': get current date)
 * @param {number} flg bitmask flag (1:always get value, 2:think time zone 4:localtime)
 * @returns Date object or null
 */
function toDate(v, flg) {
	if (!v && !!(flg & 1)) return null;
	let dt = v === 'now' ? new Date() :
		new Date(v instanceof Date ? v.getTime() : v);
	if (isNaN(dt.getTime())) {
		dt = new Date(null);
	}
	if (flg & 2) {
		let off = dt.getTimezoneOffset();
		if (off) {
			if (flg & 4) off = -off;
			dt.setMinutes(dt.getMinutes() - off);
		}
	}
	return dt;
}
function toOdsDateTime(v) {
	let dt = toDate(arguments.length < 1 ? 'now' : v, 3);
	if (!dt) return dt;
	return dt.toISOString().replace('Z', '000000');
}
const JAPANESE_DATE_KEYS = '年月日時分秒';
function parseDateJp(dt) {
	if (dt instanceof Date) return dt;
	let d = new Date(dt);
	if (!isNaN(d.getTime())) return d;
	d = new Date(null);
	let len = JAPANESE_DATE_KEYS.length;
	let be = 0;
	for (let i = 0; i < len; i++) {
		let m = dt.match(`(\\d+)${JAPANESE_DATE_KEYS.charAt(i)}`);
		if (m) {
			be |= 1 << i;
			let n = Number(m[1]);
			switch (i) {
			case 0: d.setFullYear(n); break;
			case 1: d.setMonth(n-1); break;
			case 2: d.setDate(n); break;
			case 3: d.setHours(n); break;
			case 4: d.setMinutes(n); break;
			case 5: d.setSeconds(n); break;
			}
		}
	}
	if (!be) {
		let m = dt.match(/(\d+)\/(\d+)(\/(\d+))?/);
		if (m) {
			let mon, day;
			be |= 6;
			if (m[4]) {
				be |= 1;
				let y = Number(m[1]);
				mon = Number(m[2]);
				day = Number(m[4]);
				if (y < 100) y += 2000;
				d.setFullYear(y);
			} else {
				mon = Number(m[1]);
				day = Number(m[2]);
			}
			d.setMonth(mon-1);
			d.setDate(day);
		}
		m = dt.match(/(\d+):(\d+)(:(\d+))?/) || dt.match(/PT(\d+)H(\d+)M((\d+)S)?/);
		if (m) {
			be |= 24;
			d.setHours(Number(m[1]));
			d.setMinutes(Number(m[2]));
			if (m[4]) {
				be |= 32;
				d.setSeconds(Number(m[4]));
			}
		}
		if (!be) {
			d = parseDate(dt);
			if (isNaN(d.getTime())) return null;
		}
	}
	if (!(be & 1) && be & 6) {
		d.setFullYear((new Date()).getFullYear());
	}
	return d;
}
function toTimeString(v) {
	let dt = parseDateJp(v);
	return dt instanceof Date ? dt.toTimeString().substring(0, 5) : '';
}
function toDateTimeStringLong(v) {
	let dt = parseDateJp(v);
	if (dt instanceof Date) {
		let y = dt.getFullYear();
		let M = String(dt.getMonth() + 1).padStart(2, '0');
		let d = String(dt.getDate()).padStart(2, '0');
		let h = String(dt.getHours()).padStart(2, '0');
		let m = String(dt.getMinutes()).padStart(2, '0');
		let s = String(dt.getSeconds()).padStart(2, '0');
		return `${y}-${M}-${d}T${h}:${m}:${s}`;
	}
	return '';
}
function toDateTimeString(v) {
	return toDateTimeStringLong(v).slice(0, -3);
}
function convertToOfficeDateValue(v) {
	let dt = parseDateJp(v);
	if (dt instanceof Date) {
		let f = 'yyyy-MM-dd';
		if (v.indexOf('T') > 0) f += 'THH:mm:ss';
		return formatDate(dt, f);
	}
	return '';
}
function convertToOfficeTimeValue(v) {
	let dt = parseDateJp(v);
	if (dt instanceof Date) {
		const h = String(dt.getHours()).padStart(2, '0');
		const m = String(dt.getMinutes()).padStart(2, '0');
		const s = String(dt.getSeconds()).padStart(2, '0');
		return `PT${h}H${m}M${s}S`;
	}
	return '';
}
function singleObject(v) {
	let keys = typeof v === 'object' ? Object.keys(v) : null;
	return keys && keys.length === 1 ? v[keys[0]] : v;
}
function getOrAddObject(ar, obj) {
	let s = JSON.stringify(obj);
	let i = ar.findIndex(function(o) {
		return JSON.stringify(o) === s;
	});
	if (i >= 0) {
		return i;
	}
	ar.push(obj);
	return ar.length - 1;
}

const INTL_LOCATION = {
	'ja': 'ja-JP-u-ca-japanese',
};
const DATE_FORMAT_LENGTH = ['narrow', 'short', 'long'];
function getDateFormat(opts, options) {
	let loc = opts?.loc || navigator.language;
	let il = INTL_LOCATION[(loc.substring(0, 2))];
	return new Intl.DateTimeFormat(il || loc, options);
}
function getGengoYear(dt, opts, form, flg = 0) {
	let dtf = getDateFormat(opts, {
		era: DATE_FORMAT_LENGTH[(form.length - 2) % 3],
		year: 'numeric'
	});
	let s = dtf.format(dt);
	let m = s.match(/(.+)((\d+|元))/);
	return m?.[flg + 1] || '';
}
function getMonthText(dt, opts, form) {
	return getDateFormat(opts, {
		month: DATE_FORMAT_LENGTH[(form.length - 2) % 3]
	}).format(dt);
}
function getWeekday(dt, opts, form) {
	return getDateFormat(opts, {
		weekday: DATE_FORMAT_LENGTH[(form.length - 1) % 3]
	}).format(dt);
}
function getAmPm(dt, opts) {
	let ar = getDateFormat(opts, {
		hour: 'numeric',
		hour12: true
	}).formatToParts(dt);
	let p = ar.find(n => n.type === 'dayPeriod');
	return p ? p.value : '';
}
function to2Digit(n) {
	return ('0' + n).slice(-2);
}
function datePart(key, dt, opts) {
	switch (key) {
	case 'ggge':
	case 'gge':
	case 'ge':
		return getGengoYear(dt, opts, key, -1);
	case 'GGGE':
	case 'GGE':
	case 'GE':
		return getGengoYear(dt, opts, key);
	case 'YY':
		return getGengoYear(dt, opts, key, 1);
	case 'yy':
		return to2Digit(dt.getFullYear());
	case 'YYYY':
	case 'yyyy':
	case 'Y':
	case 'y':
		return dt.getFullYear();
	case 'Mmmm':
	case 'Mmm':
	case 'mmmm':
	case 'mmm':
		opts = {loc:'en-us'};
	case 'MMMM':
	case 'MMM':
		return getMonthText(dt, opts, key);
	case 'Mm':
	case 'MM':
		return to2Digit(dt.getMonth() + 1);
	case 'M':
		return dt.getMonth() + 1;
	case 'DD':
	case 'dd':
		return to2Digit(dt.getDate());
	case 'D':
	case 'd':
		return dt.getDate();
	case 'dddd':
	case 'ddd':
		return getWeekday(dt, {loc:'en-us'}, key.substring(1));
	case 'AAAA':
	case 'AAA':
	case 'AA':
		key = key.substring(1);
	case 'WWW':
	case 'WW':
	case 'W':
		return getWeekday(dt, opts, key);
	case 'HH':
		return to2Digit(dt.getHours());
	case 'H':
		return dt.getHours();
	case 'hh':
		return to2Digit(dt.getHours() % 12);
	case 'h':
		return (dt.getHours() % 12);
	case 'mm':
		return to2Digit(dt.getMinutes());
	case 'm':
		return dt.getMinutes();
	case 'ss':
		return to2Digit(dt.getSeconds());
	case 's':
		return dt.getSeconds();
	case 'am/pm':
		opts = {loc:'en-us'};
	case 'AM/PM':
	case 'ap':
		return getAmPm(dt, opts);
	}
	return key;
}
function formatDate(v, df, opts) {
	let dt = parseDateJp(v);
	if (!dt) return '';
	let re = /(GGGE|GGE|GE|YY|yyyy|yy|y|MMMM|MMM|MM|M|dd|d|WWW|WW|W|HH|H|hh|h|mm|m|ss|s|ap)/g;
	return df.replace(re, function(key) {
		return datePart(key, dt, opts);
	});
}
function getNumberDataStyle(ds, n) {
	const map = ds?.map;
	if (map) {
		for (let f in map) {
			if (Function('return ' + f.replace('?', n))()) {
				return map[f];	
			}
		}
	}
	return ds;
}
function applyDataStyle(ds, v, t) {
	switch (t) {
	case '':
		if (!v || isNaN(v)) break;
	case 'n':
		if (v === undefined) return '';
		let n = Number(v);
		if (!isNaN(n)) {
			ds = getNumberDataStyle(ds, n);
			let s = '';
			let text = ds?.text;
			let sign = '';
			if (text?.pre) {
				s += text.pre.join('');
				if (n < 0) sign = '-';
			}
			s += ds?.symbol || '';
			let sn = n.toLocaleString(ds?.loc || navigator.language, {
				minimumIntegerDigits: ds?.dig || 1,
				minimumFractionDigits: ds?.mdot || 0,
				maximumFractionDigits: ds?.dot || 0,
				useGrouping: ds?.grp || false,
			});
			if (sign && sn.startsWith(sign)) sn = sn.substring(1);
			if (text?.suf) {
				sn += text.suf.join('');
			}
			return s + sn;
		}
		break;
	case 'd':
		let df = ds?.df;
		if (df) {
			return formatDate(v, df, ds?.opts);
		}
		break;
	}
	return v;
}
function isGeneral(v) {
	return /general/i.test(v);
}
function getGeneralFormat(opts) {
	if (opts) {
		if (opts.gf) return opts.gf;
		else if (opts.isPercent) return '0%';
		else if (opts.isNumber) return '#,##0';
	}
	return '';
}
function analyzeDateFormat(f, opts) {
	if (isGeneral(f)) return {text: getGeneralFormat(opts)};
	let blk = 0;
	let quote = 0;
	let time = 0;
	let data = {
		g: '',
		y: '',
		M: '',
		d: '',
		h: '',
		m: '',
		s: '',
		w: '',
		ap: '',
	};
	let block = '';
	let text = '';
	let ln = '';
	let str = '';
	let c;
	if (f)
	for (let i = 0; i < f.length; i++) {
		if ((c = f.charAt(i)) === '[') {
			blk++;
		} else if (c === ']') {
			blk--;
		} else if (c === '"') {
			if (quote) quote--;
			else quote++;
		} else if (c === '\\' && blk === 0) {
			if (ln) {
				text += '}';
				ln = '';
			}
			text += f.charAt(++i);
		} else if (blk === 0) {
			if (quote) {
				if (ln) {
					text += '}';
					ln = '';
				}
				text += c;
				continue;
			}
			let n = '';
			switch (c) {
			case ';':
				str = f.substring(i + 1);
				i += str.length;
				continue;
			case 'G':
			case 'g':
			case 'E':
			case 'e':
				n = 'g';
				break;
			case 'Y':
			case 'y':
				n = 'y';
				break;
			case 'M':
				n = 'M';
				break;
			case 'm':
				if (time) n = 'm';
				else {
					n = 'M';
					if (ln !== n) {
						c = n;
					}
				}
				break;
			case 'D':
			case 'd':
				n = 'd';
				break;
			case 'H':
			case 'h':
				n = 'h';
				c = 'H';
				time++;
				break;
			case 'S':
			case 's':
				n = 's';
				break;
			case 'A':
			case 'a':
				let ampm = f.substring(i, i + 5).toUpperCase();
				if (ampm === 'AM/PM') {
					n = 'ap';
					c = block ? ampm.toLowerCase() : ampm;
					i += ampm.length;
					break;
				}
			case 'W':
			case 'w':
				n = 'w';
				break;
			}
			if (ln && ln !== n) text += '}';
			if (n) {
				data[n] += c;
				if (ln !== n) {
					text += '${';
				}
			}
			text += c;
			ln = n;
		} else {
			block += c;
		}
	}
	if (ln) text += '}';
	let flag = 0;
	if (data.y || data.g) flag |= 1;
	if (data.M) flag |= 2;
	if (data.d) flag |= 4;
	if (data.w) flag |= 8;
	if (data.h) flag |= 0x10;
	if (data.m) flag |= 0x20;
	if (data.s) flag |= 0x40;
	if (data.ap) flag |= 0x80;
	return {
		flag: flag,
		data: data,
		text: text,
		block: block,
		str: str,
		type: !flag ? '' :
			!(flag & 0xf0) ? 'date' :
			!(flag & 0x0f) ? 'time' :
			'datetime-local',
	};
}
function analyzeFormat(f, val, opts) {
	let df = analyzeDateFormat(f, opts);
	if (!df.flag && !isNaN(Number(val))) {
		df.flag = 0x100;
	}
	return df;
}
function getAsDate(v, flg = 0) {
	if (v instanceof Date) {
		return v;
	} else if (typeof v === 'number') {
		return toDate(numdate(v), flg);
	}
	return parseDateJp(v);
}
function getInputFormat(f, val) {
	if (f) {
		let t, v, d;
		let df = analyzeFormat(f, val);
		let flg = df.flag;
		if (flg & 0xff) {
			let dt = getAsDate(val, 6);
			if (dt == null && df.str === '@') {
				t = 'text';
				v = val;
			} else {
				let s = toDateTimeStringLong(dt);
				t = df.type;
				if (!(flg & 0xf0)) {
					v = s.substring(0, 10);
				} else if (!(flg & 0x0f)) {
					v = s.substring(11, 19);
				} else {
					v = s.substring(0, 19);
				}
			}
		} else if (flg & 0x100) {
			let m = f.match(/\.([0Z]+)/);
			if (m) d = m[1].length;
		}
		return {
			type: t,
			value: v,
			dot: d,
		};
	}
	return null;
}
function validTypeNumber(type, v) {
	try {
		switch (type) {
		case 'number':
			return Number(v);
		case 'date':
		case 'time':
		case 'datetime-local':
			if (typeof v === 'number') return v;
			let dt = parseDateJp(v);
			let n = datenum(toDate(dt, 2));
			let s = ('' + n).split('.');
			if (type === 'date') {
				return Number(s[0]);
			} else if (type === 'time') {
				return s.length > 1 ? Number('0.' + s[1]) : 0;
			}
			return n;
		}
	} catch (e) {
	}
	return v;
}
function validNumber(el) {
	return validTypeNumber(el.type, el.value);
}
function formatDateVariable(f, dt, opts) {
	if (!dt) return '';
	return f.replace(/\$\{([^\}]+)\}/g, (m, key) => {
		return datePart(key, dt, opts);
	});
}
const normalizeFormat = f => f.replace(/\\(.{1})/g, '$1').replace(/_(.{1})/g, ' ');
function formatNumber(f, v, opts) {
	const fs = f.split(';');
	const sz = fs.length;
	if (isNaN(v)) {
		return sz > 3 ? (fs[3] || '@').replace('@', v) : v;
	}
	let n = Number(v);
	f = normalizeFormat(fs[sz > 2 && !n ? 2 : sz > 1 && n < 0 ? 1 : 0]);
	if (isGeneral(f)) {
		const gf = getGeneralFormat(opts);
		if (!gf) return v;
		f = gf;
	}
	f = f.replace(/\"([^\"]+)\"/g, (m, p1) => {
		return p1;
	});
	const sts = [];
	f = f.replace(/\[([^\]]+)\]/g, m => {
		sts.push(m.substring(1, m.length - 1));
		return '';
	});
	const pad = f.indexOf('*');
	if (pad >= 0) f = f.substring(0, pad) + f.substring(pad + 1);
	if (typeof opts === 'object') {
		if (pad >= 0) opts.padding = pad;
		if (sts.length > 0) opts.style = sts.join(';');
	}
	const minus = n < 0 && f.indexOf('-') >= 0;
	if (f.includes('%')) n *= 100;
	return f.replace(/(([0-9#,]+)(\.[0-9#]+)?)/, (m, p1, p2, p3) => {
		let i = p2.indexOf('0');
		let dig = i >= 0 ? p2.length - i : 1, mdot, dot, grp = p2.includes(',');
		if (p3) {
			p3 = p3.substring(1);
			dot = p3.length;
			i = p3.indexOf('#');
			mdot = i >= 0 ? i : dot;
		}
		const s = n.toLocaleString(opts?.loc || navigator.language, {
			minimumIntegerDigits: dig || 1,
			minimumFractionDigits: mdot || 0,
			maximumFractionDigits: dot || 0,
			useGrouping: grp || false,
		});
		return minus ? s.replace('-', '') : s;
	});
}
function formatValue(v, f, opts, c) {
	let df = analyzeDateFormat(f, opts);
	if (df.flag) {
		let n, dt;
		if (typeof v === 'number') {
			dt = v >= 1000000 ? new Date(v) : toDate(numdate(v), 6);
			n = v;
		} else {
			dt = parseDateJp(v);
			n = validTypeNumber(df.type, dt);
		}
		if (c) c.v = n;
		return formatDateVariable(df.text, dt, opts);
	} else if (df.text) {
		return formatNumber(f, v, opts);
	}
	return f ? SSF.format(f, v) : v;
}
function applyFormatValue(c, v, opts) {
	let f = c.z;
	if (c.t === 'n' && f) {
		return formatValue(v, f, opts, c);
	}
	return f ? SSF.format(f, v) : v;
}
function blobCheck(blob, cmp, pos = 0) {
	if (typeof cmp === 'string') {
		let ar = [];
		for (let i = 0; i < cmp.length; i++) {
			ar[i] = cmp.charCodeAt(i);
		}
		cmp = ar;
	} else if (!Array.isArray(cmp)) {
		cmp = [cmp];
	}
	for (let i = 0; i < cmp.length; i++) {
		if (cmp[i] !== blob[pos + i]) return false;
	}
	return true;
}
function blobLsb(blob, pos = 0) {
	blob.l = pos;
	blob.numb2 = (_) => blob[blob.l++] | (blob[blob.l++] << 8);
	blob.numb4 = (_) => blob[blob.l++] | (blob[blob.l++] << 8) | (blob[blob.l++] << 16) | (blob[blob.l++] << 24);
}
function blobMsb(blob, pos = 0) {
	blob.l = pos;
	blob.numb2 = (_) => (blob[blob.l++] << 8) | (blob[blob.l++]);
	blob.numb4 = (_) => (blob[blob.l++] << 24) | (blob[blob.l++] << 16) | (blob[blob.l++] << 8) | (blob[blob.l++]);
}
function analyzeImageData(bstr) {
	// const blob = Uint8Array.from({length: bstr.length}, (_, i) => bstr.charCodeAt(i));
	// ↑ は返って遅くなるので、ループで回す
	const blob = new Uint8Array(bstr.length);
	for (let i = 0; i < bstr.length; i++) {
		blob[i] = bstr.charCodeAt(i);
	}
	let t, w, h, dat;

	if (blobCheck(blob, [0x89,0x50,0x4E,0x47])) {
		// PNG
		t = 'png';
		blobMsb(blob, 16);
		w = blob.numb4();
		h = blob.numb4();
	} else if (blobCheck(blob, [0xFF,0xD8])) {
		// JPEG
		for (let i = 0; i < blob.length - 9; i++) {
			if (blob[i] === 0xFF && (blob[i + 1] === 0xC0 || blob[i + 1] === 0xC1 || blob[i + 1] === 0xC2)) {
				t = 'jpeg';
				blobMsb(blob, i + 5);
				h = blob.numb2();
				w = blob.numb2();
				break;
			}
		}
	} else if (blobCheck(blob, 'GIF')) {
		// GIF
		t = 'gif';
		blobLsb(blob, 6);
		w = blob.numb2();
		h = blob.numb2();
	} else if (blobCheck(blob, 'BM')) {
		// BMP
		t = 'bmp';
		blobLsb(blob, 18);
		w = blob.numb4();
		h = blob.numb4();
	} else if (blobCheck(blob, [0x01,0x00,0x00,0x00])) {
		// EMF (Enhanced Metafile)
		// Note: EMF is a vector format, so width/height are derived from the bounds rectangle in pixels.
		// Additional validation could check if the header size (bytes 4-7) is at least 40, but kept simple here.
		t = 'emf';
		blobLsb(blob, 8);
		const left = blob.numb4();
		const top = blob.numb4();
		const right = blob.numb4();
		const bottom = blob.numb4();
		dat = {
			left: left,
			top: top,
			right: right,
			bottom: bottom,
			w: right - left,
			h: bottom - top,
		};
	} else if (blobCheck(blob, 'VCLMTF')) {
		// StarView Metafile (SVM)
		t = 'svm';
		blobLsb(blob, 6);
		dat = {
			version: blob.numb2(),
			compress: blob.numb4(),
			reserved: blob.numb4(),
			w: blob.numb4(),
			h: blob.numb4(),
		};
	} else if (/<.*\Wxml.*svg\W/i.test(bstr)) {
		t = 'svg+xml';
		const xml = Xml.xmlStrToObject(bstr);
		const st = xml?.style;
		if (st) {
			st.split(';').forEach(s => {
				const ar = s.split(':');
				if (ar.length === 2) {
					switch (ar[0].trim()) {
					case 'width':
						w = getPixelSize(ar[1]);
						break;
					case 'height':
						h = getPixelSize(ar[1]);
						break;
					}
				}
			});
		}
	}

	if (!t) {
		const msg = 'Unsupported image format:' + bstr.substring(0, 16);
		console.warn(msg);
		throw new Error(msg);
	}
	if (dat) dat.blob = blob;
	return dat ? dat : {
		type: t,
		w: w,
		h: h,
		isImage: true,
		b64: bytesToBase64(blob)
	};
}
function bytesToBase64(bytes, chunkSize) {
	chunkSize = chunkSize > 0 ? Math.ceil(chunkSize / 3) * 3 : 65535;
	let result = '';
	for (let i = 0; i < bytes.length; i += chunkSize) {
		const chunk = bytes.slice(i, i + chunkSize);
		result += String.fromCharCode(...chunk);
	}
	return btoa(result);
}

// convert path to referrence path
function convertToRelPath(path) {
	return path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
}

// get directory name
function getDirName(path) {
	const i = path.lastIndexOf('/');
	return i > 0 ? path.substring(0, i) : null;
}

// get relative path from parent
function getRelativePath(path, parent, root = 'xl') {
	const ar = path.split('/');
	const dirs = parent ? parent.split('/').concat(ar) : ar;
	for (let i = dirs.length; i >= 0; i--) {
		const d = dirs[i];
		if (!d || d === '.') {
			dirs.splice(i, 1);
		} else if (d === '..') {
			if (i > 1) {
				dirs.splice(--i, 2);
			} else {
				if (i === 1)
					dirs.splice(0, i, root);
				else
					dirs[i] = root;
				break;
			}
		}
	}
	return dirs.join('/');
}

/**
 * XLSX text parser contain rPr
 */
var XlsxTextParser = {
	themes: {},
	getAsArray: function(v) {
		if (!v) return null;
		if (!Array.isArray(v)) v = [v];
		return v;
	},
	getText: function(v) {
		if (!v) return '';
		switch (typeof v) {
		case 'string': return v;
		case 'object':
			if (Array.isArray(v)) {
				let ar = [];
				v.forEach(function(o) {
					ar.push(this.getText(o));
				}, this);
				return ar.join("");
			}
			if (v.hasOwnProperty('value')) return v.value;
			if (v.t) return this.getText(v.t);
			return '';
		}
		return String(v);
	},
	getRgbColor: function(c) {
		const m = /^[a-f0-9]+$/i.exec(c);
		if (m) {
			const a = c.length > 6 ? c.substring(0, 2) : '';
			return '#' + c.slice(-6) + a;
		}
		return c;
	},
	transparentColor: function(c, tp) {
		if (/^#[a-f0-9]{6}$/i.test(c)) {
			let n = Math.floor(tp * 255) % 256;
			return c + n.toString(16).padStart(2, '0');
		}
		return c;
	},
	getThemeColor: function(v, fore) {
		const th = this.themes?.themeElements?.clrScheme;
		if (Array.isArray(th)) {
			let t;
			if (isNaN(v)) {
				t = th.find(t => t.name === v);
			} else {
				let i = Number(v);
				if (fore) {
					if (i === 1) i = 0;
				}
				t = th[i];
			}
			if (t) return this.getRgbColor(t.rgb);
		}
	},
	getIndexColor: function(v) {
		let c;
		const idx = Number(v);
		if (!isNaN(idx)) {
			if (idx < 64) {
				const colors = this.themes?.indexedColors;
				if (Array.isArray(colors) && idx < colors.length) {
					c = colors[idx];
				}
			} else {
				const th = this.themes?.themeElements?.clrScheme;
				if (Array.isArray(th)) {
					let i = -1;
					switch (idx) {
					case 64: case 65:	// dk1,lt1
						i = idx - 64;
						break;
					case 80: case 81:	// dk2,lt2
						i = idx - 80 + 2;
						break;
					case 72: case 73:	// hlink,folHlink
						i = idx - 72 + 10;
						break;
					default:	// accent1～6
						i = idx - 66 + 4;
						break;
					}
					if (i >= 0 && i < th.length) {
						c = th[i].rgb;
					}
				}
			}
		}
		return c && this.getRgbColor(c);
	},
	getColor: function(o, def, fore) {
		switch (typeof o) {
		case 'string':
			return o;
		case 'object':
			let c, tp = 0;
			if (o.hasOwnProperty('rgb')) {
				c = this.getRgbColor(o.rgb);
			}
			for (let n in o) {
				let v = o[n];
				switch (n) {
				case 'auto':
				case 'rgb':
					break;
				case 'index':
				case 'indexed':
					if (!c) c = this.getIndexColor(v);
					break;
				case 'theme':
					if (!c) c = this.getThemeColor(v, fore);
					break;
				case 'tint':
					tp = parseFloat(v);
					break;
				default:
					console.warn('Not implement color:' + n, v);
					continue;
				}
			}
			if (c) {
				return tp ? this.transparentColor(c, tp) : c;
			}
			break;
		default:
			console.warn('Not implement color type:' + typeof o, o);
			break;
		}
		return def !== undefined ? def : 'initial';
	},
	getStyle: function(o) {
		let ar = [];
		for (let p in o) {
			switch (p) {
			case 'value':
			case 't':
				continue;
			case 'space':
				// ar.push('white-space:pre');
				break;
			case 'rPr':
				const rPr = o[p];
				if (rPr && typeof rPr === 'object') {
					let dec = [];
					let bold;
					for (let n in rPr) {
						let v = rPr[n];
						let st;
						switch (n) {
						case 'b':
							if (v.val || Object.keys(v).length === 0) bold = 'bold';
							break;
						case 'i':
							st = 'font-style:italic';
							break;
						case 'u':
							dec.push('underline');
							st = `text-decoration-style:${v.val}`;
							break;
						case 'strike':
							dec.push('line-through');
							break;
						case 'charset':
						case 'family':
						case 'scheme':
							break;
						case 'rFont':
							st = `font-family:${v.val}`;
							break;
						case 'sz':
							st = `font-size:${v.val}pt`;
							break;
						case 'color':
							st = `color:${this.getColor(v, 'auto', true)}`;
							break;
						default:
							console.warn('Not implement rPr:' + n, v);
							continue;
						}
						if (st) ar.push(st)
					}
					ar.push(`font-weight:${bold || 'normal'}`);
					if (dec.length > 0) ar.push(`text-decoration:${dec.join(' ')}`);
				}
				break;
			default:
				console.warn('Not implement style:' + p, o[p]);
				continue;
			}
		}
		return ar.length > 0 ? ar.join(';') : '';
	},
	getHtml: function(v) {
		let s = '';
		if (!v) return s;
		if (!Array.isArray(v)) v = [v];
		v.forEach(function(o) {
			let t;
			if (typeof o === 'object') {
				t = this.getText(o.t || o);
				let st = this.getStyle(o);
				if (st) t = `<span style="${st}">${t}</span>`;
			} else {
				t = String(o);
			}
			s += t;
		}, this);
		return s;
	},
};

// export core-K2 expansion
var CK2 = {
	Xml: Xml,
	XlsxTextParser: XlsxTextParser,
	parseDateJp: parseDateJp,
	toTimeString: toTimeString,
	toDateTimeString: toDateTimeString,
	formatDate: formatDate,
	formatDateVariable: formatDateVariable,
	applyDataStyle: applyDataStyle,
	getInputFormat: getInputFormat,
	validNumber: validNumber,
	applyFormatValue: applyFormatValue,
	formatValue: formatValue,
	formatNumber: formatNumber,
};
