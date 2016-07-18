var csv = require('fast-csv');
var Promise = require('bluebird');
var fs = require('fs');

var promiseCsv = Promise.method(function(path, options) {
  return new Promise(function(resolve, reject) {
    var records = [];
    var stream = fs.createReadStream(path)
    var csvStream = csv
      .fromPath(path, options)
      .on('data', function(record) {
        records.push(record);
      })
      .on('end', function() {
        resolve(records);
      });
    stream.pipe(csvStream)
  })
})

module.exports = promiseCsv;