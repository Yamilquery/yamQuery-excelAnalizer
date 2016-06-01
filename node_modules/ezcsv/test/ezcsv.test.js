//
// Unit test: Run using mocha
//

var ezcsv = require('..');
var fs = require('fs');

describe('TSV', function () {
  describe('Read', function () {
    it('should read tab-seprated-value files', function () {
      verifyExample('example1.tsv');
    });
  });
  describe('Write', function () {
    it('should write tab-seprated-value files', function () {
      verifyExample('example3.json', 'tsv');
    });
  });
});

describe('CSV', function () {
  describe('Read', function () {
    it('should read tab-seprated-value files', function () {
      verifyExample('example2.csv');
    });
  });
  describe('Write', function () {
    it('should write tab-seprated-value files', function () {
      verifyExample('example4.json', 'csv');
    });
  });
});


function verifyExample(fName, fmt) {
  var name = __dirname + '/../examples/' + fName;
  var other = name.slice(0, -3) + 'json';
  var data;
  if (name.slice(-3).toLowerCase() == 'csv') {
    data = JSON.stringify(ezcsv.csv.read(name), null, 2);
  } else if (name.slice(-3).toLowerCase() == 'tsv') {
    data = JSON.stringify(ezcsv.tsv.read(name), null, 2);
  } else if (fmt) {
    var json = JSON.parse(fs.readFileSync(name).toString());
    other = name.slice(0, -4) + fmt.toLowerCase();
    name = '/tmp/' + (new Date()).getTime().toString() + '-' + fName.slice(0, -4) + fmt.toLowerCase();
    if (fmt.toLowerCase() == 'csv') ezcsv.csv.write(name, json);
    else if (fmt.toLowerCase() == 'tsv') ezcsv.tsv.write(name, json);
    data = fs.readFileSync(name).toString();
  }
    
  if (data == fs.readFileSync(other).toString()) return;
  throw new Error('Serialized ' + fName + ' does not match expected JSON.');
}
