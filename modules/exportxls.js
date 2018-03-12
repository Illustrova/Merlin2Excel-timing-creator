var xl = require('excel4node');
var s = require('../modules/styles.js');
var c = require("../config.js");
var moment = require('moment');

module.exports = function(project) {
/***************
*CREATE BOOK AND SHEET
***************/
//Import book
var wb = s.wb;

//Worksheet options
var wsOptions = {'sheetFormat': {
				'defaultColWidth': 28.3,
				'defaultRowHeight': 12.75
		}};

// Add Worksheets to the workbook
var ws = wb.addWorksheet('Timing', wsOptions);

//Change first column width
ws.column(1).setWidth(7);

//Adjust day height to look more consistent
if ((project.maxworks - project.maxpres) > 1) {
	var emptyRows = project.maxworks - project.maxpres - 2;
	if (emptyRows > 0) {
		project.height += emptyRows;
	}
}

/***************
* SHEET HEADER
***************/
ws.cell(2,2).string("Project: " + project.projectName).style(s.mainStyle);
ws.cell(3,2).string("Duration: " + /*duration + */ "\"").style(s.mainStyle);

//Logo
ws.addImage({
	path: './assets/logo.png',
	type: 'picture',
	position: {
		type: 'twoCellAnchor',
		from: {
			col: 7,
			colOff: 0,
			row: 2,
			rowOff: 0
		},
		to: {
			col: 9,
			colOff: 0,
			row: 6,
			rowOff: 0

		}
	}
});

//Title
ws.cell(5, 2).string("TIME SCHEDULE");

/***************
*BODY
***************/
var timing = project.timing;
var fCell = {row: 6, column: 2};

//wrap public holidays with moment.js
var publicHolidays = c.publicHolidays.map(function(holiday) {
	return moment(holiday).set('hour', 12);

});
//Check if date is holiday
function isHoliday(date) {
	for (var i = 0; i < publicHolidays.length; i++) {
		if (date.isSame(publicHolidays[i])) {
			return true;
		}
	}
	return false;
}

function createDay(obj, row, col) {
	//Get style of header
	var dayStyle;
	if (obj.dayOfWeek === "Saturday" || obj.dayOfWeek === "Sunday") {
		dayStyle = s.weekEnd;
	}
	else {
		dayStyle = s.weekDay;
	}
	/*DAY HEADER
	**********/
	ws.cell(row++, col).string(obj.dayOfWeek.toUpperCase())
		.style(s.firstRow)
		.style(dayStyle)
		.style(s.center);

	ws.cell(row++, col).date(obj.date)
		.style(s.lastRow)
		.style(dayStyle)
		.style(s.center)
		.style(s.dateFormat);

	/*DAY CONTENT
	**************/
	//Create borders
	for (j = 0; j < project.height; j++) {
		ws.cell(row + j, col).style(s.midRow).style(s.colors.white);

		if (j === (project.height - 1)) {
			ws.cell(row + j, col).style(s.lastRow).style(s.colors.white);
		}
	}

	//Mark public holidays
	if (isHoliday(obj.date)) {
		for (j = 0; j < project.height; j++) {
		ws.cell(row + j, col).style(s.colors.holiday);
			if (j === Math.ceil(project.height / 2) - 1) {
				ws.cell(row + j, col).string("PUBLIC HOLIDAY").style(s.colors.holiday).style(s.center);
			}
		}
	}

	//Capitalize first letter of string
	function capitalizeFirst(string) {
		return string.charAt(0).toUpperCase() + string.slice(1);
	}

	//Recalculate positions according current cell and print
	var startRow;
	var endRow;
	if (obj.works.length === 0) {
		//Presentations only
		startRow = row;
		endRow = row + project.height;

		while (row < endRow) {
			obj.presentations.forEach(function(pres) {
				pos = startRow + pres.pos;

				if (row === pos) {
					ws.cell(row, col).string(capitalizeFirst(pres.name))
					.style(s.colors.presentation)
					.style(s.center);
				}
			});
			row++;
		}
	}
	else {
		//Presentations
		//Adjust positions in case there are empty lines on top (see line 25)
		startRow = row + project.height - project.maxpres - project.maxworks;
		endRow = row + project.height - project.maxworks;

		while (row < endRow) {
			obj.presentations.forEach(function(pres) {
				pos = startRow + pres.pos;

				if (row === pos) {
					ws.cell(row, col).string(capitalizeFirst(pres.name))
					.style(s.colors.presentation)
					.style(s.center);
				}
			});
			row++;
		}
		//Jobs
		startRow = row;

		endRow = row + project.maxworks;

		while (row < endRow) {
			obj.works.forEach(function(work) {
				pos = startRow + work.pos;
				var color = work.color;

				if (row === pos) {
					ws.cell(row, col).string(capitalizeFirst(work.name))
						.style(s.colors[color])
						.style(s.center);
					}
			});
			row++;
		}
	}
}

/*********************
* CREATE DAYS
**********************/
timing.forEach(function(day) {
	if (fCell.column > 8) {
		fCell.row += (project.height + 2);
		fCell.column = 2;
	}
	createDay(day, fCell.row, fCell.column++);
});

/**************
* EXPORT BOOK
**************/
var currentDate = new Date();
var dateString = currentDate.getUTCDate().toString() + ("0" + (currentDate.getUTCMonth() + 1).toString()).slice(-2) + (currentDate.getUTCFullYear().toString()).slice(-2);
var fileName = c.companyName + "_" + project.projectName + "_timing_" + dateString;
wb.write(fileName + '.xlsx');

};