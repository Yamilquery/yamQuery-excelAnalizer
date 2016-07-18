# YamilQuery-Excel

Lector especialmente adaptado para leer documentos de tipo xls, xlsx y csv.
+ Este lector permite lo siguiente.
  - Leer ficheros xls, xslx y csv
  - Leer todos los archivos especificando un directorio
  - Leer un archivo en particular
  - Leer y almacenar el contenido en un objeto

### Requerimientos

YamilQuery-Excel requiere los siguientes módulos:

* [node-xlsx](https://www.npmjs.com/package/node-xlsx) 
* [csvdata](https://www.npmjs.com/package/csvdata) 

### Instalación

YamilQuery-Excel requiere [Node.js](https://nodejs.org/) v4+ para ejecutarse correctamente.

```sh
$ npm install yamQuery-excel --save
```

### Uso

```sh
Excel = require('yamQuery-excel')
var test_excel = new Excel()
test_excel.readDirectory('test/xlsx').then(function(data){
	console.log(data)
})
```

Usar con co generator
```sh
co = require('co')
Excel = require('yamQuery-excel')
co(function*(){
	var test_excel = new Excel()
	var data = yield test_excel.readFile('test/csv/matriz.csv')
	console.log(data)
})
```