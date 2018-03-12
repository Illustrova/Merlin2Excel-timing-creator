var xl = require('excel4node');

/**************
* CREATE BOOK
**************/
var wb = new xl.Workbook({
	defaultFont: {
				size: 10,
				name: 'Calibri',
				color: '#000000'
		},
	logLevel: 2
});

/*********
* STYLES
*********/
//header/footer text style, general style
var mainStyle = wb.createStyle({
	font: {
		color: '#000000',
		size: 10,
		family: "roman",
		name: "Helvetica CE"
	},
	numberFormat: '$#,##0.00; ($#,##0.00); -'
});

/*Date format
*************/
var dateFormat = wb.createStyle({
	numberFormat: 'dd-mmm-yyyy'
});

/*Borders
**********/
var firstRow = wb.createStyle({
	border: {
		left: {
				style: "thin",
				color: "#000000"
		},
		right: {
				style: "thin",
				color: "#000000"
		},
		top: {
				 style: "thin",
				color: "#000000"
		}
	}});

var lastRow = wb.createStyle({
	border: {
		left: {
				style: "thin",
				color: "#000000"
		},
		right: {
				style: "thin",
				color: "#000000"
		},
		bottom: {
				 style: "thin",
				color: "#000000"
		}
	}});

var midRow = wb.createStyle({
	border: {
		left: {
				style: "thin",
				color: "#000000"
		},
		right: {
				style: "thin",
				color: "#000000"
		}

	}});

/* COLORS
**********/
//Backround  fill for weekday/weekend
var weekEnd = wb.createStyle({
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#f79646'
	}});

var weekDay = wb.createStyle({
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#c0c0c0'
	}});

//Public holiday
var holiday = wb.createStyle({
	font: {
		color: '#d96709',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#f9d5b6'
	}
});

//Presentation
var presentation = wb.createStyle({
	font: {
		color: '#ff0000',
		size: 10,
		bold: true
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#f2dbda'
	}
});

//JOBS COLORS
//animatic, previz
var green = wb.createStyle({
	font: {
		color: '#007e39',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#ccff99'
	}
});

//render
var red = wb.createStyle({
	font: {
		color: '#c00000',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#fcd5b4'
	}
});

//comp, online, revisions
var blue = wb.createStyle({
	font: {
		color: '#002060',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#dce6f1'
	}
});

//Animation, simulation, particles
var yellow = wb.createStyle({
	font: {
		color: '#cc6600',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#ffff66'
	}
});

//Tracking, additional
var grey = wb.createStyle({
	font: {
		color: '#6d6d6d',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#dddddd'
	}
});

//3D modelling
var pink = wb.createStyle({
	font: {
		color: '#cc0066',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#ffccff'
	}
});

//additional
var lightblue = wb.createStyle({
	font: {
		color: '#0066cc',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#caf7f7'
	}
});

//additional
var orange = wb.createStyle({
	font: {
		color: '#cb5c01',
		size: 10
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#fdd991'
	}
});

//Shooting, empty cells
var white = wb.createStyle({
	font: {
		color: '#000000',
		size: 10,
		bold: true
	},
	fill: {
		type: "pattern",
		patternType: "solid",
		fgColor: '#ffffff'
	}
});

/*Alignment
************/
var center  = wb.createStyle({
	alignment: {
		horizontal: 'center',
		vertical: 'center'
	}
});

/***********
* EXPORT
***********/
module.exports = {
	wb: wb,
	mainStyle: mainStyle,
	weekEnd: weekEnd,
	weekDay: weekDay,
	dateFormat: dateFormat,
	firstRow: firstRow,
	lastRow: lastRow,
	midRow: midRow,
	center: center,
	colors: {
		"green": green,
		"red": red,
		"blue": blue,
		"yellow": yellow,
		"grey": grey,
		"lightblue": lightblue,
		"orange": orange,
		"pink": pink,
		"holiday": holiday,
		"presentation": presentation,
		"white": white
	}
};