'use strict';
var Excel = require('yamQuery-excel')
var co = require('co')

var ExcelAnalizer = function(){
	this.excel = new Excel()
	var resourceConfig = {}

	this.getResourceConfig = function(){
		return resourceConfig
	}

	this.setResourceConfig = function(d){
		resourceConfig = d
	}
}

ExcelAnalizer.prototype.getResults = co.wrap(function*(resourceConfig){
	var self = this
	self.setResourceConfig(resourceConfig)
	var result = []
	var resourceConfig = self.getResourceConfig()
	if(resourceConfig['directory']){
		var resources = yield self.excel.readDirectory(resourceConfig['directory'])
		resources.some(function(resourceFile){
			result.push( getResultResource.call(self,resourceFile) )
		})
		return Promise.resolve(result)
	} else if(resourceConfig['file']){
		var resourceFile = yield self.excel.readFile(self.getResourceConfig().file)
		result = getResultResource.call(self,resourceFile)

		return Promise.resolve(result)
	} else {
		return Promise.reject('Para obtener la estadística debes especificar un directory o file')
	}
})

var getResultResource = function(resourceFile){
	var self = this
	var resourceConfig = self.getResourceConfig()
	var result = {
		resource_name:resourceConfig['resource_name'],
		data:[]
	}
	resourceFile.some(function(sheet){
		if(sheet['name']==resourceConfig['sheet_name']){
			var indexRow = 0
			sheet['data'].some(function(row){
				var isValid = isValidRow.call(self,sheet,indexRow)
				if(isValid){
					var resultRow = getResultRow.call(self,sheet,indexRow)
					resultRow = getTransformedRow.call(self,resultRow)
					result['data'].push(resultRow)
				}
				indexRow++
			})
		}
	})
	return result
}

var getResultRow = function(sheet,indexRow){
	var self = this
	var resultRow = {}
	var resourceConfig = self.getResourceConfig()
	resourceConfig['results'].some(function(result){
		var alias = result['alias']
		resultRow[alias] = getResultCell(sheet,indexRow,result)
	})
	return resultRow
}

var getTransformedRow = function(resultRow){
	var self = this
	var transformedRow = {}
	var resourceConfig = self.getResourceConfig()
	if(resourceConfig['dataQuality']){
		resourceConfig['dataQuality'].some(function(dataQuality){
			var cellValue = resultRow[dataQuality['alias']]
			if(dataQuality['sumResults']){
				cellValue = 0
				dataQuality['sumResults'].some(function(sumResult){
					cellValue += resultRow[sumResult]
				})
			}
			transformedRow[dataQuality['alias']] = cellValue
		})
		return transformedRow
	}
	return resultRow
}

var getResultCell = function(sheet,indexRow,result){
	var cellValue = null
	var relativeX = (result['relative_x'] || result['relative_x']>0) ? result['relative_x'] : 0
	var relativeY = (result['relative_y'] || result['relative_y']>0) ? result['relative_y'] : 0
	var row = (relativeY || relativeY>0) ? sheet['data'][indexRow+relativeY] : sheet['data'][indexRow]

	if((result['bottomSearch']) && (result['where'])){
		var MAX_ROW_SEARCH = 20
		for(var i=indexRow; i<indexRow+(MAX_ROW_SEARCH); i++ ){
			if( isValidCell(result['where'],sheet['data'][i][result['column']]) ){
				var indexColumn = (result['relative_x']) ? result['column']+result['relative_x'] : result['column']
				cellValue = sheet['data'][i][indexColumn]
				i=indexRow+(MAX_ROW_SEARCH)
			}
		}
	} else {
		if(result['cell']) cellValue = sheet['data'][result['cell'][0]][result['cell'][1]+relativeX]
		if(result['column'] || result['column']==0) cellValue = row[result['column']+relativeX] 
		if(result['column_range']){
			cellValue = 0
			for(var i=result['column_range'][0]; i<result['column_range'][1]; i++){
				cellValue += parseFloat(row[i+relativeX])
			}
		}
	}

	if(result['regex']) cellValue = applyRegex(cellValue,result['regex'])

	return cellValue
}

var isValidRow = function(sheet,indexRow){
	var self = this
	var isValid = true
	var row = sheet['data'][indexRow]
	var resourceConfig = self.getResourceConfig()
	resourceConfig['conditions'].forEach(function(condition){
		row = (condition['relative_y'] || condition['relative_y']>0) ? sheet['data'][indexRow+condition['relative_y']] : row
		var cellValue = row[condition['column']]
		var where = condition['where']
		if((!cellValue) || ( !isValidCell(where,cellValue) )) isValid = false
	})
	return isValid
}

var isValidCell = function(where,cellValue){
	if(!cellValue) return false
	for(var i in Object.keys(where)){
		var condition = (typeof where == 'object') ? Object.keys(where)[i] : where
		var value = where[condition]
		if(condition == 'notEquals') if(!(cellValue != value)) return false
		if(condition == 'equals') if(!(cellValue == value)) return false
		if(condition == 'isNumber') if(!(!isNaN(cellValue))) return false
		if(condition == 'contains') if((cellValue.toString().indexOf( value.toString() ))==-1) return false
	}
	return true
}

var applyRegex = function(value,regex){
	var re = new RegExp(regex, "i")
	return (value.match(re)) ? value.match(re)[1] : value
}

module.exports = new ExcelAnalizer();