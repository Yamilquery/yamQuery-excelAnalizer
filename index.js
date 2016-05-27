'use strict';
var Excel = require('yamQuery-excel')
var co = require('co')

var ExcelAnalizer = function(){
	this.excel = new Excel()
	var data = {
		//'file':'test/ejemplo.xlsx', // xsl, xslx or csv
		'directory':null,
		'resource_name':'',
		'sheet_name':'Sheet1',
		'conditions':[
			{
				'column':0,
				//'cell':[1,1], // Optional (cell or column),
				'alias':'column1',
				'where':{notEquals:''},
			}
		],
		'results':[
			{
				'alias':'total', // Asigna un alias a la celda
				//'cell':[1,1], // Toma el valor de la fila 1 con columna 1 (en el valor de row si el valor es negativo o positivo(-3 o +5) quiere decir que relativamente),
				//'column':6, // Toma el valor de la columna 6,
				'column_range':[4,16], // Suma el valor de la columna 4 a la 16 de la celda actual,
				//'columns':[4,5,7,10], // Suma el valor de las columnas 4,5,7 y 10 de la celda actual,
				'transform':'exampleMethodTransform()' // De el arreglo de valores obtenidos se le aplica una transformación
			}
		]
	}

	this.getData = function(){
		return data
	}

	this.setData = function(d){
		data = d
	}
}

ExcelAnalizer.prototype.getResults = co.wrap(function*(resourceConfig){
	var self = this
	self.setData(resourceConfig)
	var result = []
	if(self.getData().directory){
		var resources = yield self.excel.readDirectory(self.getData().directory)
		resources.some(function(resourceFile){
			result.push( getResultResource.call(self,resourceFile) )
		})
		return Promise.resolve(result)
	} else if(self.getData().file){
		var resourceFile = yield self.excel.readFile(self.getData().file)
		result.push( getResultResource.call(self,resourceFile) ) // TODO: No funciona

		return Promise.resolve(result)
	} else {
		return Promise.reject('Para obtener la estadística debes especificar un directory o file')
	}
})

var getResultResource = function(resourceFile){
	var self = this
	var result = []
	resourceFile.some(function(sheet){
		if(sheet['name']==self.getData().sheet_name){
			var indexRow = 0
			sheet['data'].some(function(row){
				var isValid = isValidRow.call(self,row)
				if(isValid){
					var resultRow = getResultRow.call(self,sheet,indexRow)
					result.push(resultRow)
				}
				indexRow++
			})
		}
	})
	return result
}

var isValidRow = function(row){
	var self = this
	var isValid = true
	self.getData().conditions.forEach(function(condition){
		var cellValue = row[condition['column']]
		var where = condition['where']
		if( !isValidCell(where,cellValue) ) isValid = false
	})
	return isValid
}

var isValidCell = function(where,cellValue){
	var condition = (typeof where == 'object') ? Object.keys(where)[0] : where
	var value = where[condition]
	if(condition == 'notEquals') return cellValue != value
	if(condition == 'equals') return cellValue == value
	if(condition == 'isNumber') return (!isNaN(cellValue))
	if(condition == 'contains') return cellValue.toString().indexOf( value.toString() )!=-1
	return false
}

var getResultRow = function(sheet,indexRow){
	var self = this
	var resultRow = {}
	var row = sheet['data'][indexRow]
	self.getData().results.some(function(result){
		var alias = result['alias']
		var cellValue = null
		row = (result['relative'] || result['relative']>0) ? sheet['data'][indexRow+result['relative']] : row
		if(result['cell']) cellValue = sheet['data'][result['cell'][0]][result['cell'][1]]
		if(result['column'] || result['column']==0) cellValue = row[result['column']]
		if(result['column_range']){
			cellValue = 0
			for(var i=result['column_range'][0]; i<result['column_range'][1]; i++){
				cellValue += parseFloat(row[i])
			}
		}
		if(result['regex']) cellValue = applyRegex(cellValue,result['regex'])
		resultRow[alias] = cellValue
	})
	return resultRow
}

var applyRegex = function(value,regex){
	var re = new RegExp(regex, "gi");
	return (value.match(re)) ? value.match(re)[0] : value
}

module.exports = new ExcelAnalizer();