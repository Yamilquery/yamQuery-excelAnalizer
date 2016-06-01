# ezcsv

Convert between CSV and JSON types with very simple API.

[![NPM info](https://nodei.co/npm/ezcsv.png?downloads=true)](https://nodei.co/npm/ezcsv.png?downloads=true)

[![Travis build status](https://api.travis-ci.org/Like-Falling-Leaves/ezcsv.png?branch=master)](
https://travis-ci.org/Like-Falling-Leaves/ezcsv)


## Install

    npm install ezcsv


## API

This library tries to convert tables stored as CSV or TSV into an array of json objects.  Consider a table of the following format:

Name|Date
Yorkan, Yorkie|Sun Aug 31 2014 16:07:57 GMT-0700 (PDT)
Yorkie, Yorkan|Sun Aug 31 2014 16:08:23 GMT-0700 (PDT)

This module would convert that to the following JSON structure:

```json
{
  "header": [
    "Name",
    "Date"
  ],
  "items": [
    {
      "Name": "Yorkan, Yorkie",
      "Date": "Sun Aug 31 2014 16:07:57 GMT-0700 (PDT)"
    },
    {
      "Name": "Yorkie, Yorkan",
      "Date": "Sun Aug 31 2014 16:08:23 GMT-0700 (PDT)"
    }
  ]
}
```

### Read TSV

If a file is saved as TSV, the module expects the first line to be the 'header'.  It then creates an object for every subsequent line.

```javascript
    var ezcsv = require('ezcsv');
    var json = ezcsv.tsv.read('example1.tsv');
    ....
    json.items.forEach(function () { .... })
```

### Writing TSV as output

You can write a set of objects to TSV.  If the header field is omitted, it will calculate the header based on the values of the objects provided (default to empty string for any missing fields).

```javascript
    var ezcsv = require('ezcsv');
    var json = ezcsv.tsv.write(
      'example1.tsv', 
      {
        header: ['Name', 'Date'], // optional, it will auto-calculate
        items: [ .... ]
      }
    );
```

### CSV Files

Reading and writing CSV files is similar, just use <strong>ezcsv.csv</strong> instead of <strong>ezcsv.tsv</strong>/

### Empty lines, separators etc.

By default, empty lines are skipped both when reading and writing.  Line separator is newline when writing and also carriage-return when reading.  Field seperator is comma or tab as the case may be.  All of these are configurable by using the <strong>get</strong> method with a config (it returns the <strong>read</strong> and <strong>write</strong> methods).  Look at index.js to see how to specify these configuration parameters.