var fs = require('fs');
var xml2js = require('xml2js');

exports.xml = function(file, cb) {
	var parser = new xml2js.Parser();

	fs.readFile(file, function(err, data) {
		if(err)
			cb(err);
		else
			parser.parseString(data, cb);
	});
};