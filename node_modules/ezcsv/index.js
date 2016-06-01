var fs = require('fs');
var _ = require('lazy.js');

var tsv = {
  lineSepRE: /[\r\n]/, lineSep: '\n',
  fieldSepRE: /\t/, fieldSep: '\t',
  readSkipEmpty: true, writeSkipEmpty: true
};

var csv = {
  lineSepRE: /[\r\n]/, lineSep: '\n',
  fieldSepRE: /,/, fieldSep: ',',
  readSkipEmpty: true, writeSkipEmpty: true
};

module.exports.csv = get(csv);
module.exports.tsv = get(tsv);
module.exports.get = get;

function get(cfg) {
  return {
    read: function (name) { return readFile(name, cfg.lineSepRE, cfg.fieldSepRE, cfg.readSkipEmpty); },
    write: function (name, json) { return writeFile(name, json, cfg.lineSep, cfg.fieldSep, cfg.writeSkipEmpty); }
  };
}

function readFile(fileName, lineSep, fieldSep, skipEmpty) {
  var lines = fs.readFileSync(fileName).toString().split(lineSep);
  if ((!lines.length) || (lines.length == 1)) return [];
  var header = lines.shift().split(fieldSep);
  return {
    header: header,
    items: _(lines).map(function (line) {
      var fields = line.split(fieldSep);
      if (skipEmpty && _(fields).compact().isEmpty()) return;
      return _(header).zip(fields).toObject();
    }).compact().toArray()
  };
}

function writeFile(fileName, json, lineSep, fieldSep, skipEmpty) {
  if (!json.header) {
    var header = {};
    _(json.items || []).each(function (obj) { _(obj).each(function (value, key) { header[key] = 1; }); });
    json.header = _(obj).keys(header);
  }
  var defaults = _(json.header).zip(_.generate(function () { return ''; }, json.header.length).value()).toObject();
  fs.writeFileSync(fileName, json.header.join(fieldSep));
  _(json.items || []).each(function (item) {
    var picked = _(item).pick(json.header).defaults(defaults).values();
    if (skipEmpty && picked.compact().isEmpty()) return;
    fs.appendFileSync(fileName, lineSep + picked.join(fieldSep));
  });
}
