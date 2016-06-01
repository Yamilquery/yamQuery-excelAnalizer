'use strict';
var assert = require('chai').assert
var mocha = require('mocha')
var coMocha = require('co-mocha')
var analizer = require('../index')
var fs = require('fs')

describe('ExcelAnalizer', function() {
	describe('#getResults()', function () {

		it('should return true if get statistic from THESIS JOURNALS 2016 success', function*(){
			var example = JSON.parse(fs.readFileSync('./config/example.json', 'utf8'))
			var data = yield analizer.getResults(example)
			console.log(data)
			assert.equal(true, data.length>0 )
		})
		// TODO: yamquery-excel compatibilidad con tsv 
		it.skip('should return true if read successfull an tsv file',function*(){
			var example = JSON.parse(fs.readFileSync('./config/example-tsv.json', 'utf8'))
			var data = yield analizer.getResults(example)
			assert.equal(true, data.length>0 )
		})
  	})
})