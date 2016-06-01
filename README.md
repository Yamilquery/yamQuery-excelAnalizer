# YamilQuery-ExcelAnalizer

Permite analizar y procesar grandes volumentes de información de archivos xls, xlsx y csv, mediante archivos de configuración JSON.

+ Este analizador permite lo siguiente.
  - Leer ficheros xls, xslx y csv
  - Obtener datos de columnas y celdas especificando criterios o condiciones
  - Aplicar filtros a los datos obtenidos y almacenar el resultado en un objeto

### Requerimientos

YamilQuery-ExcelAnalizer requiere:

* [YamQuery-Excel](https://www.npmjs.com/package/yamQuery-excel) 

### Instalación

YamilQuery-ExcelAnalizer requiere [Node.js](https://nodejs.org/) v4+ para ejecutarse correctamente.

```sh
$ npm install yamquery-excelanalizer --save
```

### Uso

El analizador de archivos utiliza archivos de configuración en formato json.
A continuación un ejemplo con la estructura requerida:

###### example.json
```sh
{
	"file": "./test/xlsx/example.xlsx",
	"resource_name": "EXAMPLE YamQuery-ExcelAnalizer",
	"sheet_name": "usageReport (14)",
	"conditions": [{
		"column": 0,
		"alias": "information",
		"where": {
			"contains": "Account "
		}
	}],
	"results": [{
		"alias": "id_information",
		"column": 0,
		"regex": "\\d+"
	}, {
		"alias": "information",
		"column": 0
	}, {
		"alias": "jan-2016",
		"column": 10,
		"relative_y": 2
	}, {
		"alias": "feb-2016",
		"column": 11,
		"relative_y": 2
	}, {
		"alias": "mar-2016",
		"column": 12,
		"relative_y": 2
	}, {
		"alias": "abr-2016",
		"column": 13,
		"relative_y": 2
	}]
}
```
Las opciones del archivo de configuración son:
+ __file__ (Opcional): Ruta del archivo a analizar
+ __directory__ (Opcional): Ruta del directorio donde se localizan los archivos a analizar
+ __sheet_name__ (Opcional): Nombre de la hoja de excel por analizar
+ __conditions__: Objeto con criterios o condiciones necesarias que deben cumplir las celdas para obtener los resultados deseados
  - __column__: Especifica el número de columna donde buscará dentro de la fila en la ejecución actual
  - __cell__: Especifica una fila y columna en particular
  - __alias__: Alias para identificar el criterio
  - __where__: Objeto llave-valor con criterio y valor esperado o puede especificar un solo Criterio como tipo String 
    - __contains__ (Object)
    - __isNumber__ (String)
    - __equals__ (Object)
    - __notEquals__ (Object)
+ __results__: Objeto con posiciones de las columnas, celdas o rangos a obtener cuando se han cumplido todas las condiciones
  - __column__: Especifica el número de columna donde buscará dentro de la fila en la ejecución actual
  - __column_range__: Objecto con dos valores de rango de columnas (Se sumarán para el resultado)
  - __cell__: Especifica una fila y columna en particular
  - __alias__: Otorga un alias al resultado, mismo que será devuelto al final
  - __relative_y__: Cuando el valor que requieres no se encuentra en la fila actual puedes incrementar o disminuir el indice de la fila de donde obtendrá el valor

Nota: Al menos una de las opciones file o directory deberán especificarse en el archivo de configuración
##### Obtener resultados
Invocamos al analizador para obtener los resultados.
###### example.js
```sh
analizer = require('yamquery-excelanalizer')
var configAnalizer = JSON.parse(fs.readFileSync('./config/example.json', 'utf8'))
analizer.getResults(configAnalizer).then(function(data){
	console.log(data) // Your result is ready!
})
```

##### Usar con co generator
Gracias a Promises (https://www.promisejs.org/) podemos utilizar yield de la siguiente forma:
###### example.js
```sh
co = require('co')
analizer = require('yamquery-excelanalizer')
co(function*(){
	var configAnalizer = JSON.parse(fs.readFileSync('./config/example.json', 'utf8'))
	var data = yield analizer.getResults(configAnalizer)
	console.log(data) // Your result is ready!
})
```


### Test
Run ```npm test ```