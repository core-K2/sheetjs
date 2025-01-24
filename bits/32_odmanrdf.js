var CT_ODS = "application/vnd.oasis.opendocument.spreadsheet";
function parse_manifest(d, opts) {
  var str = xlml_normalize(d);
  var Rn;
  var FEtag;
  while (Rn = xlmlregex.exec(str))
    switch (Rn[3]) {
      case "manifest":
        break;
      case "file-entry":
        FEtag = parsexmltag(Rn[0], false);
        if (FEtag.path == "/" && FEtag.type !== CT_ODS)
          throw new Error("This OpenDocument is not a spreadsheet");
        break;
      case "encryption-data":
      case "algorithm":
      case "start-key-generation":
      case "key-derivation":
        throw new Error("Unsupported ODS Encryption");
      default:
        if (opts && opts.WTF)
          throw Rn;
    }
}
function write_manifest(manifest) {
  var o = [XML_HEADER];
  o.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
  o.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
  for (var i = 0; i < manifest.length; ++i)
    o.push('  <manifest:file-entry manifest:full-path="' + manifest[i][0] + '" manifest:media-type="' + manifest[i][1] + '"/>\n');
  o.push("</manifest:manifest>");
  return o.join("");
}
function write_rdf_type(file, res, tag) {
  return [
    '  <rdf:Description rdf:about="' + file + '">\n',
    '    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/' + (tag || "odf") + "#" + res + '"/>\n',
    "  </rdf:Description>\n"
  ].join("");
}
function write_rdf_has(base, file) {
  return [
    '  <rdf:Description rdf:about="' + base + '">\n',
    '    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="' + file + '"/>\n',
    "  </rdf:Description>\n"
  ].join("");
}
function write_rdf(rdf) {
  var o = [XML_HEADER];
  o.push('<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n');
  for (var i = 0; i != rdf.length; ++i) {
    o.push(write_rdf_type(rdf[i][0], rdf[i][1]));
    o.push(write_rdf_has("", rdf[i][0]));
  }
  o.push(write_rdf_type("", "Document", "pkg"));
  o.push("</rdf:RDF>");
  return o.join("");
}
function write_meta_ods(wb, opts) {
  let props = wb?.Props;
  if (!props)
    return '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2"><office:meta><meta:generator>SheetJS ' + XLSX.version + "</meta:generator></office:meta></office:document-meta>";
  let o = [
    '<office:document-meta',
    ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
    ' xmlns:ooo="http://openoffice.org/2004/office"',
    ' xmlns:xlink="http://www.w3.org/1999/xlink"',
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"',
    ' xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
    ' xmlns:grddl="http://www.w3.org/2003/g/data-view#"',
    ' office:version="1.3">',
 //   '<office:meta>',
  ];
  props.generator = `SheetJS ${XLSX.version} (Extended Edition by core-K2)`;
  props.date = toOdsDateTime();
  props['editing-cycles'] = toNumber(props?.['editing-cycles']) + 1;
  let prefix = n => {
    if (['language', 'date'].includes(n)) return 'dc';
    return 'meta';
  };
  o.push(makeXmlTag('office:meta', props, null, '?', prefix, (n, val, pre) => {
    let c = '', s = '';
    if (!Array.isArray(val)) val = [val];
    val.forEach(v => {
      let bObj = typeof v === 'object';
      c += makeXmlTag(pre + n, v, (v) => {
        return bObj ? v?.value || '' : v === 0 ? ' ' : v || '';
      }, bObj ? '?' : null, prefix, (n => {
        return n === 'value' ? {s:'', c:''} : null;
      }));
    });
    return {
      s: s,
      c: c
    };
  }));
/*
  let v = props?.['creation-date'];
  if (v) o.push(`<meta:creation-date>${v}</meta:creation-date>`);  
  o.push(`<dc:date>${toOdsDateTime()}</dc:date>`);
  o.push(`<meta:generator>SheetJS ${XLSX.version} (Extended Edition by core-K2)</meta:generator>`);
  o.push(`<meta:print-date>${toOdsDateTime(props?.['print-date'])}</meta:print-date>`);
  o.push(`<meta:editing-duration>${props?.['editing-duration']}</meta:editing-duration>`);
  v = props?.['editing-cycles'] || 0;
  o.push(`<meta:editing-cycles>${++v}</meta:editing-cycles>`);
  o.push(makeXmlTag('meta:document-statistic', props?.['document-statistic'], null, '?', 'meta'));
  o = o.concat([
    '</office:meta>',
    '</office:document-meta>',
  ]);
*/
  o.push('</office:document-meta>');
  return o.join('');
}
