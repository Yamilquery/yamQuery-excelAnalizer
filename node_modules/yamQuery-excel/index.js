'use strict';
var fs = require('co-fs')
var xlsx = require('node-xlsx')
var co = require('co')
var csv = require('csv-parse/lib/sync')
var _ = require('co-lodash')

var Excel = function(){

	this.readDirectory = co.wrap(function*(directory){
		var self = this
		var sourcesData = []
		var files = yield fs.readdir(directory)
		yield* _.coEach(files,function*(file){
			var full_path = directory+"/"+file
			var sheet = yield self.readFile(full_path)
			sourcesData.push(sheet)
		})
		return Promise.resolve(sourcesData)
	})

	this.readFile = co.wrap(function*(file){
		if(this.isExcel(file)){
			var sourceData = []
			var sheets = xlsx.parse(yield fs.readFile(file))
			sheets.some(function(sheet){ sourceData.push(sheet) })
			return Promise.resolve(sourceData)
		} else if(this.isCsv(file)){
			var fileData = yield fs.readFile(file)
			try{
				sourceData = [{
					'name':'Hoja1',
					'data': yield csv(fileData,{relax_column_count:true})
				}]
			} catch(e){ console.log(e) }
			return Promise.resolve(sourceData)
		} else if(this.isTsv(file)){
			// TODO: Dar soporte a ficheros CSV
			return Promise.resolve('Debe convertir el TSV a CSV')
		}
		return Promise.resolve(null)
	})

	this.isExcel = function(filePath){
		return filePath.indexOf('.xls') != -1
	}

	this.isCsv = function(filePath){
		return filePath.indexOf('.csv') != -1
	}

	this.isTsv = function(filePath){
		return filePath.indexOf('.tsv') != -1
	}

}

module.exports = Excel;
