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
				this.opts[n] = v[n];
			}
		} else {
			for (let i = 0; i < arguments.length; i++) {
				let n = arguments[i];
				if (this.opts.hasOwnProperty(n)) {
					this.opts[n] = arguments[++i];
				}
			}
		}
		this.setOptsAsArray('asSeqArray', 'asText');
	},
	setOptsAsArray: function() {
		for (let n in arguments) {
			let as = this.opts[n];
			if (as && !Array.isArray(as)) {
				this.opts[n] = typeof as === 'string' ? as.split(',') : [as];
			}
		}
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
			let asV;
			if ((asV = this.opts.asValue)) {
				if (bTrim && asV & 8) {
					v = v.trim();
				}
				if (asV & 1 && !isNaN(v) && v.trim().length > 0) {
					return Number(v);
				}
				if (asV & 2) {
					switch (v.toLowerCase()) {
					case 'true': return true;
					case 'false': return false;
					}
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
function cloneObject(obj) {
	if (!obj || typeof obj !== 'object') return obj;
	const o = Array.isArray(obj) ? [] : {};
	for (const key in obj) {
		if (obj.hasOwnProperty(key)) {
			o[key] = cloneObject(obj[key]);
		}
	}
	return o;
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
function applyDataStyle(ds, v, t) {
	switch (t) {
	case '':
		if (!v || isNaN(v)) break;
	case 'n':
		if (v === undefined) return '';
		let n = Number(v);
		if (!isNaN(n)) {
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
function analyzeDateFormat(f) {
	if (f.includes('General')) return {};
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
function analyzeFormat(f, val) {
	let df = analyzeDateFormat(f);
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
function formatNumber(f, v, opts) {
	if (isNaN(v)) return '';
	let n = Number(v);
	let fs = f.split(';');
	f = fs[fs.length > 1 && n < 0 ? 1 : 0];
	f = f.replace(/\"([^\"]+)\"/g, (m, p1) => {
		return p1;
	});
	f = f.replace(/\[.*\]|\_.{1}$/, '');
	return f.replace(/(([0-9#,]+).?([0-9#]*))/, (m, p1, p2, p3) => {
		let i = p2.indexOf('0');
		let dig = i >= 0 ? p2.length - i : 1, mdot, dot, grp = p2.includes(',');
		if (p3) {
			dot = p3.length;
			i = p3.indexOf('#');
			mdot = i >= 0 ? i : dot;
		}
		return n.toLocaleString(opts?.loc || navigator.language, {
			minimumIntegerDigits: dig || 1,
			minimumFractionDigits: mdot || 0,
			maximumFractionDigits: dot || 0,
			useGrouping: grp || false,
		});
	});
}
function applyFormatValue(c, v, opts) {
	let f = c.z;
	if (c.t === 'n' && f) {
		let df = analyzeDateFormat(f);
		if (df.flag) {
			let n, dt;
			if (typeof v === 'number') {
				dt = v >= 1000000 ? new Date(v) : toDate(numdate(v), 6);
				n = v;
			} else {
				dt = parseDateJp(v);
				n = validTypeNumber(df.type, dt);
			}
			c.v = n;
			return formatDateVariable(df.text, dt, opts);
		} else if (df.text) {
			return formatNumber(f, v, opts);
		}
	}
	return f ? SSF.format(f, v) : v;
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

// export core-K2 expansion
var CK2 = {
	Xml: Xml,
	parseDateJp: parseDateJp,
	toTimeString: toTimeString,
	toDateTimeString: toDateTimeString,
	formatDate: formatDate,
	formatDateVariable: formatDateVariable,
	applyDataStyle: applyDataStyle,
	getInputFormat: getInputFormat,
	validNumber: validNumber,
	applyFormatValue: applyFormatValue,
};
