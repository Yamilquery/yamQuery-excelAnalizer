'use strict';
var Excel = require('yamQuery-excel')
var co = require('co')
var _ = require('co-lodash')
var alasql = require('alasql')

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
		resource_type:resourceConfig['resource_type'],
		data:[]
	}
	var result_ajust = []
	var i = 0
	resourceFile.some(function(sheet){
		if((!resourceConfig['sheet_name']) || (sheet['name']==resourceConfig['sheet_name'])){
			var indexRow = 0
			result_ajust[i] = []
			sheet['data'].some(function(row){
				var isValid = isValidRow.call(self,sheet,indexRow)
				if(isValid){
					var resultRow = getResultRow.call(self,sheet,indexRow)
					resultRow = getTransformedRow.call(self,resultRow)
					result['data'].push(resultRow)
					result_ajust[i].push(resultRow)
				}
				indexRow++
			})
			result['data'] = adjustResult.call(self,result['data'])
		}
		i++
	})

	if(result_ajust.length>2){
		var sql = 'SELECT * FROM ? arr1 JOIN ? arr2 USING institucion'
		for(var i in result_ajust){
			if(i==0) var res = alasql(sql, [result_ajust[0], result_ajust[1]])
			if(i>1) res = alasql(sql, [res, result_ajust[i]])
		}
		result['data'] = res
	}
 
	return result
}

var adjustResult = function(resultData){
	var self = this
	var resourceConfig = self.getResourceConfig()
	if ((resourceConfig['attributes']) && ((resourceConfig['order']) || (resourceConfig['group']))){
		var sql = 'SELECT '
		if(resourceConfig['attributes']){
			resourceConfig['attributes'].some(function(attribute){
				sql += attribute + ', '
			})
			sql += 'FROM ? '
		}
		if(resourceConfig['where']) sql += 'WHERE '+resourceConfig['where']+' '
		if(resourceConfig['group']) sql += 'GROUP BY '+resourceConfig['group']+' '
		if(resourceConfig['order']) sql += 'ORDER BY '+resourceConfig['order']+' ASC '
		var sql = sql.replace(', FROM',' FROM')

		var res = alasql(sql,[resultData]);
		return res
	} else {
		return resultData
	}
}

var getResultRow = function(sheet,indexRow){
	var self = this
	var resultRow = {}
	var resourceConfig = self.getResourceConfig()
	resourceConfig['results'].some(function(result){
		if(!result['sheet_name'] || sheet['name']==result['sheet_name']){
			var alias = result['alias']
			resultRow[alias] = getResultCell(sheet,indexRow,result)
		}
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
	if (typeof cellValue == "string") cellValue = cellValue.trim().replace(/(\r\n|\n|\r|\t)/gm,"")
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
		if( condition['type'] ){
			if(condition['type']=='date') cellValue = getJsDateFromExcel(cellValue)
			if(condition['type']=='month') cellValue = getMonthFromExcel(cellValue)
		}
		var where = condition['where']
		if(cellValue==0) cellValue = '0'
		if((!cellValue) || ( isValidCell(where,cellValue)==false )) isValid = false
	})
	return isValid
}

var getMonthFromExcel = function(excelDate) {
	return getJsDateFromExcel(excelDate).getMonth()
}

var getJsDateFromExcel = function(excelDate) {
	return new Date((excelDate - (25567 + 1))*86400*1000)
}

var isValidCell = function(where,cellValue){
	if(!cellValue) return false
	for(var i in Object.keys(where)){
		var condition = (typeof where == 'object') ? Object.keys(where)[i] : where
		var value = where[condition]
		if(condition == 'notEquals') return notEquals(value, cellValue)
		if(condition == 'equals') return equals(value, cellValue)
		if(condition == 'contains') return contains(value, cellValue)
		if(condition == 'isNumber') return isNumber(cellValue)
	}
	return true
}

var notEquals = function(value, cellValue){
	if(typeof value == 'object'){
		var values = value
		var noIncumple = true
		for(var i in values){
			var value = values[i]
			if(cellValue == value) noIncumple = false
		}
		if(!noIncumple){
			return false
		}
	} else {
		if(!(cellValue != value)) return false
	}
}

var equals = function(value, cellValue){
	if(typeof value == 'object'){
		var values = value
		var noIncumple = false
		for(var i in values){
			var value = values[i]
			if(cellValue == value) noIncumple = true
		}
		if(!noIncumple){
			return false
		}
	} else {
		if(!(cellValue == value)) return false 
	}
}

var contains = function(value, cellValue){
	if((cellValue.toString().indexOf( value.toString() ))==-1) return false
}

var isNumber = function(cellValue){
	if(!(!isNaN(cellValue))) return false
}

var applyRegex = function(value,regex){
	var re = (typeof regex=="object") ? new RegExp(regex[0], regex[1]) : new RegExp(regex, "g")
	var index = (typeof regex=="object" && regex[2] && regex[2]>0) ? regex[2] : 0
	return (value.match(re) && value.match(re)[index]) ? value.match(re)[index] : null
}

module.exports = new ExcelAnalizer();