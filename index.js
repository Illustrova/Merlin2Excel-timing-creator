var parseXML = require('./modules/parsexml.js');
var transformData = require('./modules/transformdata.js');
var exportXLS = require('./modules/exportxls.js');
var glob = require('glob-fs')({ gitignore: true });
var file = process.argv[2] || glob.readdirSync('source/*.xml')[0];

parseXML.xml(file, function (error, result) {
	if (error) throw error;
	var project = transformData(result);
	exportXLS(project);
	console.log("Excel file is ready");
});