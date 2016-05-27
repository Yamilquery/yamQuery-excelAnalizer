'use strict';
var fs = require('fs')
var xlsx = require('node-xlsx')
var co = require('co')
var csvdata = require('csvdata')

var Excel = function(){
	this.DEFAULT_SHEET_NAME = 'Sheet1'
}

Excel.prototype.isExcel = function(file){
	return file.indexOf('.xls') != -1
}

Excel.prototype.isCsv = function(file){
	return file.indexOf('.csv') != -1
}

Excel.prototype.readFile = co.wrap(function*(file){
	if(this.isExcel(file)){
		var data = []
		var sheets = xlsx.parse(fs.readFileSync(file))
		sheets.some(function(sheet){ data.push(sheet) })
		return Promise.resolve(data)
	} else if(this.isCsv(file)){
		var data = yield csvdata.load(file)
		return Promise.resolve(data)
	}
	return Promise.resolve(null)
})

Excel.prototype.readDirectory = co.wrap(function*(directory){
	var self = this
	var data = []
	for(var i in fs.readdirSync(directory) ){
		var file = fs.readdirSync(directory)[i]
		var full_path = directory+"/"+file
		var sheet = yield self.readFile(full_path)
		data.push(sheet)
	}
	return Promise.resolve(data)
})

module.exports = Excel;
 