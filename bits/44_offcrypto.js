function _JS2ANSI(str/*:string*/)/*:Array<number>*/ {
	if(typeof $cptable !== 'undefined') return $cptable.utils.encode(current_ansi, str);
	var o/*:Array<number>*/ = [], oo = str.split("");
	for(var i = 0; i < oo.length; ++i) o[i] = oo[i].charCodeAt(0);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.4 Version */
function parse_CRYPTOVersion(blob, length/*:?number*/) {
	var o/*:any*/ = {};
	o.Major = blob.read_shift(2);
	o.Minor = blob.read_shift(2);
	/*:: if(length == null) return o; */
	if(length >= 4) blob.l += length - 4;
	return o;
}

/* [MS-OFFCRYPTO] 2.1.5 DataSpaceVersionInfo */
function parse_DataSpaceVersionInfo(blob) {
	var o = {};
	o.id = blob.read_shift(0, 'lpp4');
	o.R = parse_CRYPTOVersion(blob, 4);
	o.U = parse_CRYPTOVersion(blob, 4);
	o.W = parse_CRYPTOVersion(blob, 4);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.6.1 DataSpaceMapEntry Structure */
function parse_DataSpaceMapEntry(blob) {
	var len = blob.read_shift(4);
	var end = blob.l + len - 4;
	var o = {};
	var cnt = blob.read_shift(4);
	var comps/*:Array<{t:number, v:string}>*/ = [];
	/* [MS-OFFCRYPTO] 2.1.6.2 DataSpaceReferenceComponent Structure */
	while(cnt-- > 0) comps.push({ t: blob.read_shift(4), v: blob.read_shift(0, 'lpp4') });
	o.name = blob.read_shift(0, 'lpp4');
	o.comps = comps;
	if(blob.l != end) throw new Error("Bad DataSpaceMapEntry: " + blob.l + " != " + end);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.6 DataSpaceMap */
function parse_DataSpaceMap(blob) {
	var o = [];
	blob.l += 4; // must be 0x8
	var cnt = blob.read_shift(4);
	while(cnt-- > 0) o.push(parse_DataSpaceMapEntry(blob));
	return o;
}

/* [MS-OFFCRYPTO] 2.1.7 DataSpaceDefinition */
function parse_DataSpaceDefinition(blob)/*:Array<string>*/ {
	var o/*:Array<string>*/ = [];
	blob.l += 4; // must be 0x8
	var cnt = blob.read_shift(4);
	while(cnt-- > 0) o.push(blob.read_shift(0, 'lpp4'));
	return o;
}

/* [MS-OFFCRYPTO] 2.1.8 DataSpaceDefinition */
function parse_TransformInfoHeader(blob) {
	var o = {};
	/*var len = */blob.read_shift(4);
	blob.l += 4; // must be 0x1
	o.id = blob.read_shift(0, 'lpp4');
	o.name = blob.read_shift(0, 'lpp4');
	o.R = parse_CRYPTOVersion(blob, 4);
	o.U = parse_CRYPTOVersion(blob, 4);
	o.W = parse_CRYPTOVersion(blob, 4);
	return o;
}

function parse_Primary(blob) {
	/* [MS-OFFCRYPTO] 2.2.6 IRMDSTransformInfo */
	var hdr = parse_TransformInfoHeader(blob);
	/* [MS-OFFCRYPTO] 2.1.9 EncryptionTransformInfo */
	hdr.ename = blob.read_shift(0, '8lpp4');
	hdr.blksz = blob.read_shift(4);
	hdr.cmode = blob.read_shift(4);
	if(blob.read_shift(4) != 0x04) throw new Error("Bad !Primary record");
	return hdr;
}

/* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
function parse_EncryptionHeader(blob, length/*:number*/) {
	var tgt = blob.l + length;
	var o = {};
	o.Flags = (blob.read_shift(4) & 0x3F);
	blob.l += 4;
	o.AlgID = blob.read_shift(4);
	var valid = false;
	switch(o.AlgID) {
		case 0x660E: case 0x660F: case 0x6610: valid = (o.Flags == 0x24); break;
		case 0x6801: valid = !!(o.Flags & 0x04); break;
		case 0: valid = (o.Flags == 0x10 || o.Flags == 0x04 || o.Flags == 0x24); break;
		default: throw 'Unrecognized encryption algorithm: ' + o.AlgID;
	}
	if(!valid) throw new Error("Encryption Flags/AlgID mismatch");
	o.AlgIDHash = blob.read_shift(4);
	o.KeySize = blob.read_shift(4);
	o.ProviderType = blob.read_shift(4);
	blob.l += 8;
	o.CSPName = blob.read_shift((tgt-blob.l)>>1, 'utf16le');
	blob.l = tgt;
	return o;
}

/* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
function parse_EncryptionVerifier(blob, length/*:number*/) {
	var o = {}, tgt = blob.l + length;
	blob.l += 4; // SaltSize must be 0x10
	o.Salt = blob.slice(blob.l, blob.l+16); blob.l += 16;
	o.Verifier = blob.slice(blob.l, blob.l+16); blob.l += 16;
	/*var sz = */blob.read_shift(4);
	o.VerifierHash = blob.slice(blob.l, tgt); blob.l = tgt;
	return o;
}

/* [MS-OFFCRYPTO] 2.3.4.* EncryptionInfo Stream */
function parse_EncryptionInfo(blob) {
	var vers = parse_CRYPTOVersion(blob);
	switch(vers.Minor) {
		case 0x02: return [vers.Minor, parse_EncInfoStd(blob, vers)];
		case 0x03: return [vers.Minor, parse_EncInfoExt(blob, vers)];
		case 0x04: return [vers.Minor, parse_EncInfoAgl(blob, vers)];
	}
	throw new Error("ECMA-376 Encrypted file unrecognized Version: " + vers.Minor);
}

/* [MS-OFFCRYPTO] 2.3.4.5  EncryptionInfo Stream (Standard Encryption) */
function parse_EncInfoStd(blob/*::, vers*/) {
	var flags = blob.read_shift(4);
	if((flags & 0x3F) != 0x24) throw new Error("EncryptionInfo mismatch");
	var sz = blob.read_shift(4);
	//var tgt = blob.l + sz;
	var hdr = parse_EncryptionHeader(blob, sz);
	var verifier = parse_EncryptionVerifier(blob, blob.length - blob.l);
	return { t:"Std", h:hdr, v:verifier };
}
/* [MS-OFFCRYPTO] 2.3.4.6  EncryptionInfo Stream (Extensible Encryption) */
function parse_EncInfoExt(/*::blob, vers*/) { throw new Error("File is password-protected: ECMA-376 Extensible"); }
/* [MS-OFFCRYPTO] 2.3.4.10 EncryptionInfo Stream (Agile Encryption) */
function parse_EncInfoAgl(blob/*::, vers*/) {
	var KeyData = ["saltSize","blockSize","keyBits","hashSize","cipherAlgorithm","cipherChaining","hashAlgorithm","saltValue"];
	blob.l+=4;
	var xml = blob.read_shift(blob.length - blob.l, 'utf8');
	var o = {};
	xml.replace(tagregex, function xml_agile(x) {
		var y/*:any*/ = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;
			case '<encryption': case '</encryption>': break;
			case '<keyData': KeyData.forEach(function(k) { o[k] = y[k]; }); break;
			case '<dataIntegrity': o.encryptedHmacKey = y.encryptedHmacKey; o.encryptedHmacValue = y.encryptedHmacValue; break;
			case '<keyEncryptors>': case '<keyEncryptors': o.encs = []; break;
			case '</keyEncryptors>': break;

			case '<keyEncryptor': o.uri = y.uri; break;
			case '</keyEncryptor>': break;
			case '<encryptedKey': o.encs.push(y); break;
			default: throw y[0];
		}
	});
	o.$raw = Xml.xmlStrToObject(xml);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
function parse_RC4CryptoHeader(blob, length/*:number*/) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4); length -= 4;
	if(vers.Minor != 2) throw new Error('unrecognized minor version code: ' + vers.Minor);
	if(vers.Major > 4 || vers.Major < 2) throw new Error('unrecognized major version code: ' + vers.Major);
	o.Flags = blob.read_shift(4); length -= 4;
	var sz = blob.read_shift(4); length -= 4;
	o.EncryptionHeader = parse_EncryptionHeader(blob, sz); length -= sz;
	o.EncryptionVerifier = parse_EncryptionVerifier(blob, length);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
function parse_RC4Header(blob/*::, length*/) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4);
	if(vers.Major != 1 || vers.Minor != 1) throw 'unrecognized version code ' + vers.Major + ' : ' + vers.Minor;
	o.Salt = blob.read_shift(16);
	o.EncryptedVerifier = blob.read_shift(16);
	o.EncryptedVerifierHash = blob.read_shift(16);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
function crypto_CreatePasswordVerifier_Method1(Password/*:string*/) {
	var Verifier = 0x0000, PasswordArray;
	var PasswordDecoded = _JS2ANSI(Password);
	var len = PasswordDecoded.length + 1, i, PasswordByte;
	var Intermediate1, Intermediate2, Intermediate3;
	PasswordArray = new_raw_buf(len);
	PasswordArray[0] = PasswordDecoded.length;
	for(i = 1; i != len; ++i) PasswordArray[i] = PasswordDecoded[i-1];
	for(i = len-1; i >= 0; --i) {
		PasswordByte = PasswordArray[i];
		Intermediate1 = ((Verifier & 0x4000) === 0x0000) ? 0 : 1;
		Intermediate2 = (Verifier << 1) & 0x7FFF;
		Intermediate3 = Intermediate1 | Intermediate2;
		Verifier = Intermediate3 ^ PasswordByte;
	}
	return Verifier ^ 0xCE4B;
}

/* [MS-OFFCRYPTO] 2.3.7.2 Binary Document XOR Array Initialization */
var crypto_CreateXorArray_Method1 = /*#__PURE__*/(function() {
	var PadArray = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00];
	var InitialCode = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3];
	var XorMatrix = [0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4];
	var Ror = function(Byte) { return ((Byte/2) | (Byte*128)) & 0xFF; };
	var XorRor = function(byte1, byte2) { return Ror(byte1 ^ byte2); };
	var CreateXorKey_Method1 = function(Password) {
		var XorKey = InitialCode[Password.length - 1];
		var CurrentElement = 0x68;
		for(var i = Password.length-1; i >= 0; --i) {
			var Char = Password[i];
			for(var j = 0; j != 7; ++j) {
				if(Char & 0x40) XorKey ^= XorMatrix[CurrentElement];
				Char *= 2; --CurrentElement;
			}
		}
		return XorKey;
	};
	return function(password/*:string*/) {
		var Password = _JS2ANSI(password);
		var XorKey = CreateXorKey_Method1(Password);
		var Index = Password.length;
		var ObfuscationArray = new_raw_buf(16);
		for(var i = 0; i != 16; ++i) ObfuscationArray[i] = 0x00;
		var Temp, PasswordLastChar, PadIndex;
		if((Index & 1) === 1) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[0], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			PasswordLastChar = Password[Password.length - 1];
			ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp);
		}
		while(Index > 0) {
			--Index;
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
		}
		Index = 15;
		PadIndex = 15 - Password.length;
		while(PadIndex > 0) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
			--Index;
			--PadIndex;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			--PadIndex;
		}
		return ObfuscationArray;
	};
})();

/* [MS-OFFCRYPTO] 2.3.7.3 Binary Document XOR Data Transformation Method 1 */
var crypto_DecryptData_Method1 = function(password/*:string*/, Data, XorArrayIndex, XorArray, O) {
	/* If XorArray is set, use it; if O is not set, make changes in-place */
	if(!O) O = Data;
	if(!XorArray) XorArray = crypto_CreateXorArray_Method1(password);
	var Index, Value;
	for(Index = 0; Index != Data.length; ++Index) {
		Value = Data[Index];
		Value ^= XorArray[XorArrayIndex];
		Value = ((Value>>5) | (Value<<3)) & 0xFF;
		O[Index] = Value;
		++XorArrayIndex;
	}
	return [O, XorArrayIndex, XorArray];
};

var crypto_MakeXorDecryptor = function(password/*:string*/) {
	var XorArrayIndex = 0, XorArray = crypto_CreateXorArray_Method1(password);
	return function(Data) {
		var O = crypto_DecryptData_Method1("", Data, XorArrayIndex, XorArray);
		XorArrayIndex = O[1];
		return O[0];
	};
};

/* 2.5.343 */
function parse_XORObfuscation(blob, length, opts, out) {
	var o = ({ key: parseuint16(blob), verificationBytes: parseuint16(blob) }/*:any*/);
	if(opts.password) o.verifier = crypto_CreatePasswordVerifier_Method1(opts.password);
	out.valid = o.verificationBytes === o.verifier;
	if(out.valid) out.insitu = crypto_MakeXorDecryptor(opts.password);
	return o;
}

/* 2.4.117 */
function parse_FilePassHeader(blob, length/*:number*/, oo) {
	var o = oo || {}; o.Info = blob.read_shift(2); blob.l -= 2;
	if(o.Info === 1) o.Data = parse_RC4Header(blob, length);
	else o.Data = parse_RC4CryptoHeader(blob, length);
	return o;
}
function parse_FilePass(blob, length/*:number*/, opts, data) {
	var o = ({ Type: opts.biff >= 8 ? blob.read_shift(2) : 0 }/*:any*/); /* wEncryptionType */
	if(o.Type) {
		parse_FilePassHeader(blob, length-2, o);
		if (Rc4.verifyPassword(o, opts)) {
			const content = Xls97.decrypt(o, data, opts);
			if (content) {
				prep_blob(content, 0);
				o.content = content;
			}
		}
	} else {
		parse_XORObfuscation(blob, opts.biff >= 8 ? length : length - 2, opts, o);
	}
	return o;
}

// Check for required libraries
function checkLibs() {
	for (var i = 0; i < arguments.length; ++i) {
		const lib = arguments[i];
		if (typeof window[lib] === 'undefined') throw new Error(lib + " is required for decryption");
	}
}

// decrypt password (Need CryptoJS)
function decrypt(einfo, data, cfb, opts) {
	if (Array.isArray(einfo) && einfo.length === 2) {
		let type = einfo[0];
		switch (type) {
		case 2: // Standard Encryption
			return Ecma376Standard.decrypt(einfo[1], data, opts);
		case 3: // Extensible Encryption
			return Ecma376Extensible.decrypt(einfo[1], data, opts);
		case 4: // Agile Encryption
			return Ecma376Agile.decrypt(einfo[1], data, opts);
		default:
			throw new Error("ECMA-376 Encrypted file unrecognized Version: " + type);
		}
	}
	throw new Error("Unsupported encryption info format:" + JSON.stringify(einfo), cfb);
}

// Convert an integer to little-endian
function toNumberLE(i) {
	return ((i & 0xff) << 24) |
		((i & 0xff00) << 8) |
		((i & 0xff0000) >>> 8) |
		((i & 0xff000000) >>> 24);
}
// Convert an integer to a WordArray
function intToWordArray(i) {
	return CryptoJS.lib.WordArray.create([i]);
}
// Convert an integer to a WordArray in little-endian format
function intToWordArrayLE(i) {
	return intToWordArray(toNumberLE(i));
}
// Convert a string or ArrayBuffer to a CryptoJS WordArray
function createWordArray(v) {
	if (v) {
		switch (typeof v) {
		case 'string':
			if (/^[0-9a-f]+$/i.test(v))
				return CryptoJS.enc.Hex.parse(v);
			else
				return CryptoJS.enc.Base64.parse(v);
		case 'object':
			if (v instanceof Uint8Array) {
				break;
			} else if (v.words && v.hasOwnProperty('sigBytes')) {
				return v;
			} else if (Array.isArray(v)) {
				v = new Uint8Array(v);
			}
			break;
		}
	}
	return CryptoJS.lib.WordArray.create(v);
}
function stringToUint8Array(str, encoding) {
	if (encoding === undefined || encoding === 'utf-16le') {
		// utf-16le
		const bytes = new Uint8Array(str.length * 2);
		for (let i = 0; i < str.length; i++) {
			const code = str.charCodeAt(i);
			bytes[i * 2] = code & 0xff;
			bytes[i * 2 + 1] = (code >> 8) & 0xff;
		}
		return bytes;
	}
	const encoder = new TextEncoder(encoding);
	return encoder.encode(str);
}
function intToArrayLE(i) {
	const buf = new Uint8Array(4);
	buf[0] = i & 0xff;
	buf[1] = (i >>> 8) & 0xff;
	buf[2] = (i >>> 16) & 0xff;
	buf[3] = (i >>> 24) & 0xff;
	return buf;
}
function toUint8Array(v) {
	if (v) {
		switch (typeof v) {
		case 'string':
			if (/^[0-9a-f]+$/i.test(v)) {
				try {
					return Uint8Array.fromHex
						? Uint8Array.fromHex(v)
						: new Uint8Array(v.match(/.{2}/g).map(byte => parseInt(byte, 16)));
				} catch (e) {
					throw new Error('Invalid hex string');
				}
			} else {
				try {
					return Uint8Array.fromBase64
						? Uint8Array.fromBase64(v)
						: new Uint8Array(atob(v).split('').map(c => c.charCodeAt(0)));
				} catch (e) {
					throw new Error('Invalid base64 string');
				}
			}
		case 'number':
			return intToArrayLE(v);
		case 'object':
			if (v instanceof Uint8Array) {
				return v;
			} else if (!Array.isArray(v)) {
				v = [v];
			}
			break;
		}
	}
	return new Uint8Array(v);
}
function uint8ArrayToHex(bytes) {
	return Array.from(bytes).map(b => b.toString(16).padStart(2, '0')).join('');
}
function fillUint8Array(ar, sz, fill = 0x0) {
	ar = toUint8Array(ar);
	const len = ar.length;
	if (len < sz) {
		const dt = new Uint8Array(sz).fill(fill, len);
		dt.set(ar);
		return dt;
	} else if (len > sz) {
		return ar.slice(0, sz);
	}
	return ar;
}
// Convert a CryptoJS WordArray to a Uint8Array
function wordArrayToUint8Array(wa, sz, fill = 0x0) {
	const bsz = wa.sigBytes;
	const size = sz || bsz;
	if (size < 0) throw new Error("invalid sigBytes:" + size);
	const words = wa.words;
	const ret = new Uint8Array(size);
	if (bsz < size) ret.fill(fill, bsz);
	for (let i = 0; i < bsz; i++) {
		ret[i] = (words[i >>> 2] >>> (24 - (i % 4) * 8)) & 0xff;
	}
	return ret;
}
// XOR a WordArray with a Uint8Array
function wordArrayXorUint8Array(wa, buf) {
	const size = wa.sigBytes;
    if (size < 0) throw new Error("invalid sigBytes:" + size);
    const words = wa.words;
	if (!buf) {
		buf = new Uint8Array(size);
	} else if (buf.length < size) {
		buf = fillUint8Array(buf, size);
	}
    for (let i = 0; i < size; i++) {
        buf[i] ^= (words[i >>> 2] >>> (24 - (i % 4) * 8)) & 0xff;
    }
    return buf;
}
// Compare two Uint8Arrays for equality
function compareUint8Array(a, b) {
	if (a.length !== b.length) return false;
	for (let i = 0; i < a.length; i++) {
		if (a[i] !== b[i]) return false;
	}
	return true;
}
// Concatenate multiple Uint8Arrays into a single Uint8Array
function concatUint8Arrays(chunks) {
	let ret = null;
	for (const chunk of chunks) {
		if (!ret) {
			ret = chunk;
		} else {
			const data = new Uint8Array(ret.length + chunk.length);
			data.set(ret, 0);
			data.set(chunk, ret.length);
			ret = data;
		}
	}
	return ret;
}
function cryptHash(algo, ...bufs) {
	const hash = CryptoJS.algo[algo].create();
	for (let i = 0; i < bufs.length; i++) {
		hash.update(bufs[i]);
	}
	return hash.finalize();
}
// bufs allow Uint8Array only
async function cryptWebHash(algo = 'SHA-512', ...bufs) {
	const data = concatUint8Arrays(bufs);
	const hash = await crypto.subtle.digest(algo, data);
	return new Uint8Array(hash);
}
async function cryptWebSpinHash(algo, hash, spinCount, startValue = 0) {
	const hlen = 4;
	const ia = new Uint8Array(hash.length + hlen);
	if (startValue) ia.set(intToArrayLE(startValue), 0);
	else ia.fill(0, hlen);
	for (let i = 0; i < spinCount; i++) {
		ia.set(hash, hlen);
		hash = await cryptWebHash(algo, ia);
		if (++ia[0]>255) if (++ia[1]>255) if (++ia[2]>255) ++ia[3];
	}
	return hash;
}
function getCipherMode(cipherChaining) {
	const m = /^ChainingMode([A-Z]+)$/.exec(cipherChaining);
	return m ? m[1] : 'CBC';
}
function getWebAPIAlogorithm(algorithm) {
	const m = /^([a-z]+)(\d+)$/i.exec(algorithm);
	return m ? m[1].toUpperCase() + '-' + m[2] : algorithm;
}
function isZip(blob) {
	return blob[0] === 0x50 && blob[1] === 0x4B && blob[2] < 0x09 && blob[3] < 0x09;
}

// ECMA-376 Encryption Decryption
const Ecma376Standard = {
	HASH_ALGO: 'SHA1',	// hash algorithm
	KEY_REPEAT_COUNT: 50000,	// Number of iterations for key derivation
	CONTENT_OFFSET: 8, // Offset to the content in the encrypted data
	HEADER_SIZE: 2,	// contents header size
	AES_BLOCK_SIZE: 16, // AES block size in bytes
	CHUNK_SIZE: 32768, // Chunk size for processing
	iv: [], // Initialization vector for AES decryption

	// Decrypt a single chunk of data using AES in ECB mode
	_decrypt: function(cipherW, keyW) {
		return CryptoJS.AES.decrypt(
			{
				ciphertext: cipherW
			},
			keyW,
			{
				iv: this.iv,
				mode: CryptoJS.mode.ECB,
				padding: CryptoJS.pad.NoPadding
			}
		);
	},

	// Derive a key from the password and salt
	// The key size can be specified, default is 128 bits
	passwordToKey: function(password, salt, keySize = 128) {
		const passwordW = CryptoJS.enc.Utf16LE.parse(password);
		const saltW = createWordArray(salt);
		let hash = cryptHash(this.HASH_ALGO, saltW, passwordW);
		for (let i = 0; i < this.KEY_REPEAT_COUNT; i++) {
			const iW = intToWordArrayLE(i);
			hash = cryptHash(this.HASH_ALGO, iW, hash);
		}
	    const dataW = createWordArray(wordArrayToUint8Array(hash, 24));
		const keyHash = cryptHash(this.HASH_ALGO, dataW);
		const buf = wordArrayXorUint8Array(keyHash, new Uint8Array(64).fill(0x36));
		const keyHashW = createWordArray(buf);
		const key = cryptHash(this.HASH_ALGO, keyHashW);
		return wordArrayToUint8Array(key, keySize / 8);
	},

	// Verify the key against the verifier and verifier hash
	// Returns true if the key is valid, false otherwise
	verifyKey: function(key, verifier, verifierHash) {
		const keyW = createWordArray(key);
		const verifierW = createWordArray(verifier);
		const verifierHashW = createWordArray(verifierHash);
		const decryptedVerifierW = this._decrypt(verifierW, keyW);
		const expectedHashW = cryptHash(this.HASH_ALGO, decryptedVerifierW);
		const expectedHash = wordArrayToUint8Array(expectedHashW);
		const checkW = this._decrypt(verifierHashW, keyW);
		const check = wordArrayToUint8Array(checkW, 20);
		return expectedHash.toString() === check.toString();
	},

	// Decrypt the content using the derived key
	// The content is expected to be in a specific format with a size header
	// Returns the decrypted content as a Uint8Array
	decryptContent: function(key, content) {
		const keyW = createWordArray(key);
		const len = content.length;
		const size = content.read_shift(this.HEADER_SIZE);
		const chunks = [];
		const blockSize = this.AES_BLOCK_SIZE;
		const chunkLen = this.CHUNK_SIZE;
		let sIdx, eIdx = this.CONTENT_OFFSET;
		while (eIdx < len) {
			sIdx = eIdx;
			eIdx = sIdx + chunkLen; 
			if (eIdx > len) eIdx = len;
			let buf = content.slice(sIdx, eIdx);
			const remaind = buf.length % blockSize;
			if (remaind) {
				const padding = new Uint8Array(blockSize - remaind).fill(0);
				buf = new Uint8Array([...buf, ...padding]); // Pad with zeros
			}
			const bufW = createWordArray(new Uint8Array(buf));
			const decryptedW = this._decrypt(bufW, keyW);
			chunks.push(wordArrayToUint8Array(decryptedW));
		}
		const result = concatUint8Arrays(chunks);
		return result.slice(0, size); // Return only the decrypted content up to the specified size
	},

	// Main decryption function that takes the encryption info, data, and options
	// It verifies the password, derives the key, and decrypts the content
	// Returns the decrypted content as a Uint8Array
	decrypt: function(einfo, data, opts) {
		if (!opts?.password) throw new Error('need password');
		checkLibs('CryptoJS');
		const {Salt, Verifier, VerifierHash} = einfo.v;
		const {Flags, AlgID, AlgIDHash, KeySize, ProviderType} = einfo.h;
		if (AlgID !== 0x660E && AlgID !== 0x6801) {
			throw new Error("Unsupported AlgID: " + AlgID);
		}
		this.iv = createWordArray(new Uint8Array(16)); // Zero IV
		const key = this.passwordToKey(opts.password, Salt, KeySize);
		if (!this.verifyKey(key, Verifier, VerifierHash)) {
			throw new Error("Password verification failed");
		}
		return readSync(this.decryptContent(key, data.content), opts);
	}
};

/**
 * Implements ECMA-376 Agile Encryption/Decryption for Office documents.
 * Provides methods for password-based key derivation, decryption, and verification
 * according to the ECMA-376 standard (Agile Encryption).
 *
 * @namespace Ecma376Agile
 *
 * @property {Object} BLOCK_KEYS - Predefined block keys for data integrity, key, and verifier hash.
 * @property {number} CONTENT_OFFSET - Offset to the content in the encrypted data.
 * @property {number} HEADER_SIZE - Size of the contents header.
 * @property {number} CHUNK_SIZE - Chunk size for processing encrypted data.
 * @property {number} FILL_VALUE - Value used to fill padding bytes.
 *
 * @method passwordToKey
 *   Derives a key from a password using the specified hash algorithm, salt, spin count, and key bits.
 *   @param {CryptoJS.lib.WordArray} passwordW - Password as a CryptoJS WordArray.
 *   @param {string} hashAlgorithm - Hash algorithm name.
 *   @param {CryptoJS.lib.WordArray} saltValueW - Salt value as a WordArray.
 *   @param {number} spinCount - Number of hash iterations.
 *   @param {number} keyBits - Desired key length in bits.
 *   @param {Array|CryptoJS.lib.WordArray} key - Block key for derivation.
 *   @returns {Uint8Array} Derived key.
 *
 * @method _decrypt
 *   Decrypts a cipher buffer using the specified key, algorithm, mode, and IV.
 *   @param {Uint8Array|CryptoJS.lib.WordArray} key - Decryption key.
 *   @param {Uint8Array} cipher - Ciphertext to decrypt.
 *   @param {string} cipherAlgorithm - Cipher algorithm name.
 *   @param {string} cipherMode - Cipher mode (e.g., CBC).
 *   @param {Uint8Array|CryptoJS.lib.WordArray} iv - Initialization vector.
 *   @returns {CryptoJS.lib.WordArray} Decrypted data.
 *
 * @method createIV
 *   Creates an initialization vector for a given block using hash algorithm, salt, and block key.
 *   @param {string} hashAlgorithm - Hash algorithm name.
 *   @param {CryptoJS.lib.WordArray} saltValueW - Salt value as a WordArray.
 *   @param {number} blockSize - Block size in bytes.
 *   @param {number|Array|CryptoJS.lib.WordArray} blockKey - Block key or block index.
 *   @returns {Uint8Array} Initialization vector.
 *
 * @method getEncryptor
 *   Retrieves and normalizes the encryptor object from encryption info.
 *   @param {Object} einfo - Encryption info object.
 *   @returns {Object} Encryptor object.
 *
 * @method verifyPassword
 *   Verifies the password against the encrypted verifier hash.
 *   @param {CryptoJS.lib.WordArray} passwordW - Password as a WordArray.
 *   @param {Object} encryptor - Encryptor object.
 *   @returns {boolean} True if password is correct, false otherwise.
 *
 * @method decryptContent
 *   Decrypts the main content using the derived package key.
 *   @param {CryptoJS.lib.WordArray} passwordW - Password as a WordArray.
 *   @param {Object} content - Encrypted content buffer.
 *   @param {Object} encryptor - Encryptor object.
 *   @returns {Uint8Array} Decrypted content.
 *
 * @method decrypt
 *   Main entry point for decryption. Verifies password and decrypts content.
 *   @param {Object} einfo - Encryption info object.
 *   @param {Object} data - Data object containing encrypted content.
 *   @param {Object} opts - Options object, must include 'password'.
 *   @returns {Uint8Array} Decrypted content.
 *   @throws {Error} If password is missing or incorrect.
 */
const Ecma376Agile = {
	BLOCK_KEYS: {
		input:		[0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79],
		value:		[0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e],
		key:		[0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6],
		hmacKey:	[0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6],
		hmacValue:	[0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33],
	},
	CONTENT_OFFSET: 8,	// Offset to the content in the encrypted data
	HEADER_SIZE: 4,		// contents header size
	CHUNK_SIZE: 4096,	// Chunk size for processing
	FILL_VALUE: 0x36,	// fill value

	passwordToKey: function(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, key) {
		// const t = performance.now();
		let hash = cryptHash(hashAlgorithm, saltValueW, passwordW);
		for (let i = 0; i < spinCount; i++) {
			hash = cryptHash(hashAlgorithm, intToWordArrayLE(i), hash);
		}
		hash = cryptHash(hashAlgorithm, hash, createWordArray(key));
		const ret = wordArrayToUint8Array(hash, keyBits / 8, this.FILL_VALUE);
		// console.log('passwordToKey', performance.now() - t);
		return ret;
	},
	passwordToKeyA: async function(password, hashAlgorithm, saltValue, spinCount, keyBits, key) {
		// const t = performance.now();
		let hash = await cryptWebHash(hashAlgorithm, saltValue, password);
		hash = await cryptWebSpinHash(hashAlgorithm, hash, spinCount);
		hash = await cryptWebHash(hashAlgorithm, hash, toUint8Array(key));
		hash = fillUint8Array(hash, keyBits / 8, this.FILL_VALUE);
		// console.log('passwordToKey', performance.now() - t);
		return hash;
	},
	_decrypt: function(key, cipher, cipherAlgorithm, cipherMode, iv, padding = CryptoJS.pad.NoPadding) {
		return CryptoJS[cipherAlgorithm].decrypt(
			{
				ciphertext: createWordArray(cipher)
			},
			createWordArray(key),
			{
				iv: createWordArray(iv),
				mode: CryptoJS.mode[cipherMode],
				padding: padding
			}
		);
	},
	createIV: function(hashAlgorithm, saltValueW, blockSize, blockKey) {
		if (typeof blockKey === 'number') blockKey = intToWordArrayLE(blockKey);
		let iv = cryptHash(hashAlgorithm, saltValueW, blockKey);
		return wordArrayToUint8Array(iv, blockSize, this.FILL_VALUE);
	},
	getEncryptor: function(einfo) {
		const encryptor = einfo.$raw.keyEncryptors.keyEncryptor.encryptedKey;
		encryptor.cipherMode = getCipherMode(encryptor.cipherChaining);
		encryptor.hashAlgorithmWeb = getWebAPIAlogorithm(encryptor.hashAlgorithm);
		return encryptor;
	},
	getKeyData: function(einfo) {
		const keyData = einfo.$raw.keyData;
		keyData.cipherMode = getCipherMode(keyData.cipherChaining);
		keyData.hashAlgorithmWeb = getWebAPIAlogorithm(keyData.hashAlgorithm);
		return keyData;
	},
	verifyPassword: function(passwordW, encryptor) {
		const {cipherAlgorithm, hashAlgorithm, saltValue, spinCount, keyBits, cipherMode, encryptedVerifierHashInput, encryptedVerifierHashValue} = encryptor;
		const saltValueW = createWordArray(saltValue);
		const keyInput = this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.input);
		const keyValue = this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.value);
		const hashInput = this._decrypt(keyInput, encryptedVerifierHashInput, cipherAlgorithm, cipherMode, saltValueW);
		const hashValue = this._decrypt(keyValue, encryptedVerifierHashValue, cipherAlgorithm, cipherMode, saltValueW);
		const verifierHash = cryptHash(hashAlgorithm, hashInput);
		return verifierHash.toString(CryptoJS.enc.Hex) === hashValue.toString(CryptoJS.enc.Hex);
	},
	verifyPasswordA: async function(password, encryptor) {
		const {cipherAlgorithm, hashAlgorithm, hashAlgorithmWeb, saltValue, spinCount, keyBits, cipherMode, encryptedVerifierHashInput, encryptedVerifierHashValue} = encryptor;
		const saltValueW = toUint8Array(saltValue);
		const keyInput = await this.passwordToKeyA(password, hashAlgorithmWeb, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.input);
		const keyValue = await this.passwordToKeyA(password, hashAlgorithmWeb, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.value);
		const hashInput = this._decrypt(keyInput, encryptedVerifierHashInput, cipherAlgorithm, cipherMode, saltValueW);
		const hashValue = this._decrypt(keyValue, encryptedVerifierHashValue, cipherAlgorithm, cipherMode, saltValueW);
		const verifierHash = cryptHash(hashAlgorithm, hashInput);
		return verifierHash.toString(CryptoJS.enc.Hex) === hashValue.toString(CryptoJS.enc.Hex);
	},
	makePackageKey: function(passwordW, encryptor) {
		const {cipherAlgorithm, hashAlgorithm, saltValue, spinCount, keyBits, cipherMode, encryptedKeyValue} = encryptor;
		const saltValueW = createWordArray(saltValue);
		const key = this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.key);
		return this._decrypt(key, encryptedKeyValue, cipherAlgorithm, cipherMode, saltValueW);
	},
	makePackageKeyA: async function(passwordW, encryptor) {
		const {cipherAlgorithm, hashAlgorithmWeb, saltValue, spinCount, keyBits, cipherMode, encryptedKeyValue} = encryptor;
		const saltValueW = toUint8Array(saltValue);
		const key = await this.passwordToKeyA(passwordW, hashAlgorithmWeb, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.key);
		return this._decrypt(key, encryptedKeyValue, cipherAlgorithm, cipherMode, saltValueW);
	},
	decryptContent: function(packageKey, content, keyData) {
		const {cipherAlgorithm, hashAlgorithm, saltValue, cipherMode, blockSize} = keyData;
		const saltValueW = createWordArray(saltValue);
		const len = content.length;
		const size = content.read_shift(this.HEADER_SIZE);
		const chunks = [];
		const chunkLen = this.CHUNK_SIZE;
		let sIdx, eIdx = this.CONTENT_OFFSET;
		for (let i = 0; eIdx < len; i++) {
			sIdx = eIdx;
			eIdx = sIdx + chunkLen; 
			if (eIdx > len) eIdx = len;
			let buf = content.slice(sIdx, eIdx);
			const iv = this.createIV(hashAlgorithm, saltValueW, blockSize, i);
			const decryptedW = this._decrypt(packageKey, buf, cipherAlgorithm, cipherMode, iv);
			chunks.push(wordArrayToUint8Array(decryptedW));
		}
		const result = concatUint8Arrays(chunks);
		return result.slice(0, size);
	},
	decrypt: async function(einfo, data, opts) {
		if (!opts?.password) throw new Error('need password');
		checkLibs('CryptoJS');
		const encryptor = this.getEncryptor(einfo);
		const passwordW = opts.useAsync ?
			stringToUint8Array(opts.password) :
			CryptoJS.enc.Utf16LE.parse(opts.password);
		const valid = opts.useAsync ?
			await this.verifyPasswordA(passwordW, encryptor) :
			this.verifyPassword(passwordW, encryptor);
		if (!valid) throw new Error('password is incorrect');
		const packageKey = opts.useAsync ?
			await this.makePackageKeyA(passwordW, encryptor) :
			this.makePackageKey(passwordW, encryptor);
		const keyData = this.getKeyData(einfo);
		const blob = this.decryptContent(packageKey, data.content, keyData);
		if (!isZip(blob)) {
			console.error('decrypt failed', einfo.$raw, blob.slice(0, 16));
			throw new Error('decrypt failed');
		}
		return readSync(blob, opts);
	}
};
const Ecma376AgileWebAPI = {
	BLOCK_KEYS: {
		input: new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
		value: new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
		key: new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
		hmacKey: new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
		hmacValue: new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
	},
	CONTENT_OFFSET: 8,
	HEADER_SIZE: 4,
	CHUNK_SIZE: 4096,
	FILL_VALUE: 0x36,

	async passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, key) {
		const saltValue = toUint8Array(saltValueW);
		const password = stringToUint8Array(passwordW);
		let hash = await cryptWebHash(hashAlgorithm, saltValue, password);
		for (let i = 0; i < spinCount; i++) {
			hash = await cryptWebHash(hashAlgorithm, intToArrayLE(i), hash);
		}
		hash = await cryptWebHash(hashAlgorithm, hash, toUint8Array(key));
		const size = keyBits / 8;
		if (hash.length !== size) {
			const result = new Uint8Array(size);
			result.fill(this.FILL_VALUE, hash.length);
			result.set(hash.slice(0, Math.min(hash.length, size)));
			return result;
		}
		return hash;
	},

	async _decrypt(key, cipher, cipherAlgorithm, cipherMode, iv) {
		const algo = cipherAlgorithm.toUpperCase() === 'AES' ? `AES-${cipherMode.toUpperCase()}` : cipherAlgorithm;
		const keyArray = toUint8Array(key);
		let cipherArray = toUint8Array(cipher);
		const ivArray = toUint8Array(iv);

		const keyObj = await crypto.subtle.importKey(
			'raw',
			keyArray,
			{ name: algo },
			false,
			['decrypt']
		);
		try {
			const decrypted = await crypto.subtle.decrypt(
				{
					name: algo,
					iv: ivArray,
				},
				keyObj,
				cipherArray
			);
			let result = new Uint8Array(decrypted);
			if (algo === 'AES-CBC') {
				const padLength = result[result.length - 1];
				if (padLength > 0 && padLength <= 16 && result.slice(-padLength).every(b => b === padLength)) {
					console.log(`Removing PKCS#7 padding of length ${padLength}`);
					result = result.slice(0, -padLength);
				} else {
					console.log('No valid PKCS#7 padding detected, assuming NoPadding');
				}
			}
			return result;
		} catch (e) {
			console.error('Decryption error:', e.message, {
				algo,
				key: uint8ArrayToHex(keyArray),
				iv: uint8ArrayToHex(ivArray),
				cipher: uint8ArrayToHex(cipherArray.slice(0, 32)),
			});
			throw e;
		}
	},

	async createIV(hashAlgorithm, saltValueW, blockSize, blockKey) {
		const saltValue = toUint8Array(saltValueW);
		const blockKeyArray = typeof blockKey === 'number' ? intToArrayLE(blockKey) : toUint8Array(blockKey);
		const iv = await cryptWebHash(hashAlgorithm, saltValue, blockKeyArray);
		const size = blockSize;
		if (iv.length !== size) {
			const result = new Uint8Array(size);
			result.fill(this.FILL_VALUE, iv.length);
			result.set(iv.slice(0, Math.min(iv.length, size)));
			return result;
		}
		return iv;
	},
	getEncryptor(einfo) {
		const encryptor = einfo.$raw.keyEncryptors.keyEncryptor.encryptedKey;
		encryptor.cipherMode = getCipherMode(encryptor.cipherChaining);
		encryptor.hashAlgorithm = getWebAPIAlogorithm(encryptor.hashAlgorithm);
		return encryptor;
	},
	getKeyData(einfo) {
		const keyData = einfo.$raw.keyData;
		keyData.cipherMode = getCipherMode(keyData.cipherChaining);
		keyData.hashAlgorithm = getWebAPIAlogorithm(keyData.hashAlgorithm);
		return keyData;
	},

	async verifyPassword(passwordW, encryptor) {
		const { cipherAlgorithm, hashAlgorithm, saltValue, spinCount, keyBits, cipherMode, encryptedVerifierHashInput, encryptedVerifierHashValue } = encryptor;
		const saltValueW = toUint8Array(saltValue);
		const keyInput = await this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.input);
		const keyValue = await this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.value);
		const hashInput = await this._decrypt(keyInput, encryptedVerifierHashInput, cipherAlgorithm, cipherMode, saltValueW);
		const hashValue = await this._decrypt(keyValue, encryptedVerifierHashValue, cipherAlgorithm, cipherMode, saltValueW);
		const verifierHash = await cryptWebHash(hashAlgorithm, hashInput);
		return uint8ArrayToHex(verifierHash) === uint8ArrayToHex(hashValue);
	},

	async makePackageKey(passwordW, encryptor) {
		const { cipherAlgorithm, hashAlgorithm, saltValue, spinCount, keyBits, cipherMode, encryptedKeyValue } = encryptor;
		const saltValueW = toUint8Array(saltValue);
		const key = await this.passwordToKey(passwordW, hashAlgorithm, saltValueW, spinCount, keyBits, this.BLOCK_KEYS.key);
		return await this._decrypt(key, encryptedKeyValue, cipherAlgorithm, cipherMode, saltValueW);
	},

	async decryptContent(packageKey, content, keyData) {
		const { cipherAlgorithm, hashAlgorithm, saltValue, cipherMode, blockSize } = keyData;
		const saltValueW = toUint8Array(saltValue);
		const len = content.length;
		const size = new DataView(content.buffer, content.byteOffset, this.HEADER_SIZE).getUint32(0, true);
		const chunks = [];
		const chunkLen = this.CHUNK_SIZE;
		let sIdx = this.CONTENT_OFFSET, eIdx = this.CONTENT_OFFSET;
		for (let i = 0; eIdx < len; i++) {
			sIdx = eIdx;
			eIdx = sIdx + chunkLen;
			if (eIdx > len) eIdx = len;
			let buf = content.slice(sIdx, eIdx);
			if (buf.length % 16 !== 0) {
			const padded = new Uint8Array(Math.ceil(buf.length / 16) * 16);
			padded.set(buf);
			padded.fill(16, buf.length);
			buf = padded;
			}
			const iv = await this.createIV(hashAlgorithm, saltValueW, blockSize, i);
			const decrypted = await this._decrypt(packageKey, buf, cipherAlgorithm, cipherMode, iv);
			chunks.push(decrypted);
		}
		const result = concatUint8Arrays(chunks);
		return result.slice(0, size);
	},

	async decrypt(einfo, data, opts) {
		if (!opts?.password) throw new Error('need password');
		const encryptor = this.getEncryptor(einfo);
		const passwordW = opts.password;
		if (!(await this.verifyPassword(passwordW, encryptor))) {
			throw new Error('password is incorrect');
		}
		const packageKey = await this.makePackageKey(passwordW, encryptor);
		const keyData = this.getKeyData(einfo);
		const blob = await this.decryptContent(packageKey, data.content, keyData);
		if (!isZip(blob)) {
			console.error('decrypt failed', einfo.$raw, blob.slice(0, 16));
			throw new Error('decrypt failed');
		}
		return readSync(blob, opts);
	},
};

const Ecma376Extensible = {
	decrypt: function(einfo, data, opts) {
		if (!opts?.password) throw new Error('need password');
		throw new Error("not implement yet Ecma376Extensible:", einfo);
	}
};

// RC4 Encryption Decryption
// This implementation uses the CryptoJS library for RC4 decryption.
// It supports two types of FilePass structures: Type 0 (XOR Obfuscation) and Type 1 (RC4 Encryption).
// The verifyPassword function checks if the provided password is correct by decrypting the verifier and verifier hash.
// The decrypt function decrypts the actual data using the derived key from the password.
// Note: The RC4 algorithm is considered weak and is not recommended for secure applications.
// This implementation is provided for compatibility with legacy formats that use RC4 encryption.
const Rc4 = {
	// Block size for processing data in chunks
	BLOCK_SIZE: 0x200, // 512 bytes

	// Convert password to key using MD5 hashing
	convertPasswordToKey(password, Salt, block) {
		const passwordW = CryptoJS.enc.Utf16LE.parse(password);
		const saltW = createWordArray(Salt);
		const salt = wordArrayToUint8Array(saltW);
		const h0 = CryptoJS.MD5(passwordW);
		const truncatedHash = wordArrayToUint8Array(h0, 5);
		const intermediateBuffer = concatUint8Arrays([truncatedHash, salt]);
		const chunks = [];
		for (let i = 0; i < 16; i++) {
			chunks.push(intermediateBuffer);
		}
		const concatenatedBuffer = concatUint8Arrays(chunks);
		const concatenatedBufferW = createWordArray(concatenatedBuffer);
		const hashW = CryptoJS.MD5(concatenatedBufferW);
		const hash = wordArrayToUint8Array(hashW, 5);
		const intermediate = concatUint8Arrays([hash, wordArrayToUint8Array(intToWordArrayLE(block))]);
		const intermediateW = createWordArray(intermediate);
		const keyW = CryptoJS.MD5(intermediateW);
		return wordArrayToUint8Array(keyW, 128 / 8);
	},

	// Verify the password against the FilePass structure
	verifyPassword: function(fpass, opts) {
		if (!opts?.password) throw new Error('need password');
		checkLibs('CryptoJS');
		const {Type, Data} = fpass;
		if (Type === 1) {
			const {Salt, EncryptedVerifier, EncryptedVerifierHash} = Data;
			const block = 0;
			const key = this.convertPasswordToKey(opts.password, Salt, block);
  
			// RC4デクリプタを作成
			const keyW = createWordArray(key);
  			const cipher = CryptoJS.algo.RC4.createDecryptor(keyW);
  			// encryptedVerifierを復号化してverifierを取得
			const encryptedVerifierW = createWordArray(EncryptedVerifier);
  			const verifier = cipher.finalize(encryptedVerifierW);
  
			// verifierのMD5ハッシュを計算
			const hash = CryptoJS.MD5(verifier);
			// encryptedVerifierHashを復号化
			const EncryptedVerifierHashW = createWordArray(EncryptedVerifierHash);
			const verifierHash = cipher.process(EncryptedVerifierHashW).concat(cipher.finalize());
  
			// ハッシュ値が一致するか比較
			return (fpass.valid = verifierHash.toString(CryptoJS.enc.Hex) === hash.toString(CryptoJS.enc.Hex));
		} else {
			throw new Error("Unsupported FilePass Type: " + Type);
		}
	},	

	// Decrypt data using RC4 algorithm
	decrypt: function(fpass, data, opts, blocksize) {
		if (!opts?.password) throw new Error('need password');
		checkLibs('CryptoJS');
		const {Type, Data} = fpass;
		let decrypted;
		if (Type === 1) {
			if (!blocksize) blocksize = this.BLOCK_SIZE;
			const {Salt} = Data;
			const outputChunks = [];
			for (let start, end = 0, block = 0; end < data.length; block++) {
				start = end;
				end = start + blocksize;
				if (end > data.length) end = data.length;

				// 次のチャンクを取得
				const inputChunk = data.slice(start, end);

				// パスワードからキーを生成
				const key = this.convertPasswordToKey(opts.password, Salt, block);

				// RC4デクリプタを作成し、チャンクを復号化
				const cipher = CryptoJS.algo.RC4.createDecryptor(createWordArray(key));
				const outputChunk = cipher.finalize(createWordArray(inputChunk));
				outputChunks.push(wordArrayToUint8Array(outputChunk));
			}
			// すべての出力チャンクを結合
			decrypted = concatUint8Arrays(outputChunks);
		} else {
			throw new Error("Unsupported FilePass Type: " + Type);
		}
		return decrypted;
	},
};

// XLS97 Decryption
// This implementation handles the decryption of XLS files encrypted with the RC4 algorithm.
// It processes the file in chunks, deriving a new key for each chunk based on the block number.
// The decrypt function takes the FilePass structure, the encrypted data, and options including the password.
// It returns the decrypted content as a Uint8Array.
const Xls97 = {
	// Size of each block for key derivation
	BLOCK_SIZE: 1024,

	// Record type numbers for specific records
	recordNameNum: {
		BOF: 0x0809,
		FilePass: 0x002F,
		UsrExcl: 0x0194,
		FileLock: 0x0195,
		InterfaceHdr: 0x00E1,
		RRDInfo: 0x0106,
		RRDHead: 0x0138,
		BoundSheet8: 0x0085,
	},

	// Iterate over records in the binary blob
	iterRecord: function(blob) {
		const dataList = [];
		prep_blob(blob, 0);
		for (;;) {
			const h = blob.read_shift(4);
			if (!h) {
				break;
			}
			blob.l = blob.l - 4;
			const header = blob.slice(blob.l, blob.l + 4);
			const num = blob.read_shift(2);
			const size = blob.read_shift(2);
			const record = blob.slice(blob.l, blob.l + size);
			const temp = {header, num, size, record};
			dataList.push(temp);
			blob.l = blob.l + size;
		}
		return dataList;
	},

	// Allocate a Uint8Array of specified size, optionally filled with a specific value
	alloc: function(size, fill = 0) {
		const buf = new Uint8Array(size);
		buf.fill(fill);
		return buf;
	},

	// Concatenate multiple Uint8Arrays into a single Uint8Array
	concat: function(chunks) {
		return concatUint8Arrays(chunks);
	},

	// Decrypt the data using the provided FilePass structure and options
	decrypt: function(fpass, data, opts) {
		const plainBuf = [];
		let encryptedBuf = [];
		const dataList = this.iterRecord(data);

		// header [num, size] 2 bytes each
		for (const {header, num, size, record} of dataList) {
			// Remove encryption, pad by zero to preserve stream size
			if (num === this.recordNameNum.FilePass) {
				// header.slice(2); // size
				plainBuf.push(0, 0, ...header.slice(2), ...Array(size).fill(0));
				encryptedBuf.push(this.alloc(4 + size));
			} else if ([
				// The following records MUST NOT be obfuscated or encrypted: BOF (section 2.4.21),
				// FilePass (section 2.4.117), UsrExcl (section 2.4.339), FileLock (section 2.4.116),
				// InterfaceHdr (section 2.4.146), RRDInfo (section 2.4.227), and RRDHead (section 2.4.226).
				this.recordNameNum.BOF,
				this.recordNameNum.FilePass,
				this.recordNameNum.UsrExcl,
				this.recordNameNum.FileLock,
				this.recordNameNum.InterfaceHdr,
				this.recordNameNum.RRDInfo,
				this.recordNameNum.RRDHead,
			].includes(num)) {
				plainBuf.push(...header, ...record);
				encryptedBuf.push(this.alloc(4 + size));
			} else if (num === this.recordNameNum.BoundSheet8) {
				// The lbPlyPos field of the BoundSheet8 record (section 2.4.28) MUST NOT be encrypted.
				const lbPlyPos = record.slice(0, 4);
				const restSize = size - 4;
				plainBuf.push(...header, ...lbPlyPos, ...Array(restSize).fill(-2));
				encryptedBuf.push(this.concat([this.alloc(4), this.alloc(4), record.slice(4)]));
			} else {
				plainBuf.push(...header, ...Array(size).fill(-1));
				encryptedBuf.push(this.concat([this.alloc(4), record]));
			}
		}
		const encrypted = this.concat(encryptedBuf);
		const dec = Rc4.decrypt(fpass, encrypted, opts, this.BLOCK_SIZE);
	
		for (let i = 0; i < plainBuf.length; i++) {
			const c = plainBuf[i];
			if (c !== -1 && c !== -2) {
				dec[i] = c;
			}
		}
		return dec;
	},
};

// Decrypt ODS (OpenDocument Spreadsheet) files
// This function handles the decryption of ODS files that are encrypted using the methods specified in the manifest.
// It supports AES-256-GCM encryption with Argon2id key derivation.
// The function retrieves the necessary parameters from the manifest, derives the encryption key, and decrypts the content.
// It returns the decrypted content as a CFB container.
// 必要なライブラリ: CryptoJS, argon2-browser
async function decrypt_ods(zip, manifest, opts) {
	checkLibs('CryptoJS', 'argon2');
	try {
		// マニフェストから暗号化パラメータを取得
		const entry = manifest["file-entry"];
		const path = entry["full-path"];
		const encryptionData = entry["encryption-data"];
		const algo = encryptionData.algorithm;
		const keyGen = encryptionData["start-key-generation"];
		const keyDeri = encryptionData["key-derivation"];
		const algorithm = algo["algorithm-name"];
		const ivBase64 = algo["initialisation-vector"];
		const startKeyGenName = keyGen["start-key-generation-name"];
		const startKeySize = parseInt(keyGen["key-size"]);
		const saltBase64 = keyDeri.salt;
		const iterations = parseInt(keyDeri["argon2-iterations"]);
		const memory = parseInt(keyDeri["argon2-memory"]);
		const lanes = parseInt(keyDeri["argon2-lanes"]);
		const keySize = parseInt(keyDeri["key-size"]);

		// Base64デコード
		const ivW = CryptoJS.enc.Base64.parse(ivBase64);
		const saltW = CryptoJS.enc.Base64.parse(saltBase64);
		const ivBytes = wordArrayToUint8Array(ivW);
		const saltBytes = wordArrayToUint8Array(saltW);

		// start-key-generation: パスワードをSHA-256でハッシュ
		let passwordInput = opts.password;
		const gens = startKeyGenName.split("#");
		switch (gens[1]) {
		case "sha256":
			const hash = CryptoJS.SHA256(passwordInput);
			const hashBytes = hash.sigBytes;
			if (startKeySize !== hashBytes) {
				throw new Error(`start-key-size ${startKeySize} not match SHA-256 size(${hashBytes})`);
			}
			// passwordInput = hash.toString(CryptoJS.enc.Hex);
			passwordInput = wordArrayToUint8Array(hash); 
			break;
		default:
			throw new Error("not supported start-key-generation algorithm:" + startKeyGenName);
		}

		// alorithm: AES-256-GCMのみ対応
		const algos = algorithm.split("#");
		let algoName, tagLength;
		switch (algos[1]) {
		case "aes256gcm":
		case "aes256-gcm":
			if (keySize !== 32) {
				throw new Error(`key-size ${keySize} not match AES-256 size(32)`);
			}
			algoName = "AES-GCM";
			tagLength = 128;
			break;
		default:
			throw new Error("not supported algorithm:" + algorithm);
		}

		// Argon2idで鍵導出
		const argon2Result = await argon2.hash({
			pass: passwordInput,
			salt: saltBytes,
			time: iterations,
			mem: memory,
			parallelism: lanes,
			hashLen: keySize,
			type: argon2.ArgonType.Argon2id,
		});
		const keyBytes = new Uint8Array(argon2Result.hash);
		const cryptoKey = await crypto.subtle.importKey(
			"raw",
			keyBytes.buffer,
			{ name: algoName },
			false,
			["decrypt"]
		);

		// 暗号化されたパッケージを取得
		const fi = zip.FileIndex.find(fi => fi.name === path);
		if (!fi) throw new Error("encrypted file not found in the zip: " + path);
		const encryptedDataRaw = fi.content;

		// IV を検証し、暗号文を準備(IV付きならIVを除去)
		const iv = new Uint8Array(ivBytes);
		const encrypted = compareUint8Array(encryptedDataRaw.slice(0, 12), iv)
			? encryptedDataRaw.slice(12)
			: encryptedDataRaw;

		// Web Crypto APIで復号化
		const decrypted = await crypto.subtle.decrypt(
			{
				name: algoName,
				iv: iv,
				tagLength: tagLength
			},
			cryptoKey,
			encrypted.buffer
		);

		const blob = new Uint8Array(decrypted);
		let data;
		if (isZip(blob)) {
			// ZIP ヘッダ検出
			data = blob;
		} else {
			// inflate 解凍
			prep_blob(blob, 0);
			data = CFB.utils._inflateRaw(blob, entry.size);
		}
		return readSync(data, opts);
	} catch (e) {
		let msg = e?.message || '';
		if (msg) msg = `(${msg})`;
		throw new Error((e.name === 'OperationError' ?
			'password is incorrect':
			'Decryption failed')
			+ msg
		);
	}
}
