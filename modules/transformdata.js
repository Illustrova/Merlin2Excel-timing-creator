var moment = require('moment');
var c = require("../config.js");

/***********
* Helpers
***********/

// Recursive function to get all children "Activity" objects
var arr = [];
function getChildren(data) {

	for (var i = 0; i < data.length; i++) {
		if (data[i].hasOwnProperty("Activity")) {
				var newData = data[i].Activity;
				getChildren(newData);
		}
		else {
			arr.push(data[i]);
		}

	}
	return arr;
}

//Transform day of week to string
function dowText(num) {
	var vals = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
	return vals[num - 1];
}

//Get maximum possible items in array
function getMax(type, arr) {
	var arrLength = [];
	arr.forEach(function(day) {
		var length =  day[type].length;
		arrLength.push(length);
	});
	return Math.max.apply(null, arrLength);
}

//Calculate the optimal height of each "day" section in excel. Minimum is 5.
function calcRows(arr) {
	var maxWorks = getMax("works", arr);
	var maxPres = getMax("presentations", arr);
	var numRows;
	if ((maxWorks + maxPres) <= 5) {
		numRows = 5;
	}
	else {
		numRows = maxWorks + maxPres;
	}
return numRows;
}

/******************************************************************
* EXPORT: Takes parsed xml data and transforms to array of objects,
* where each object represents day, planned activities and output styles
*******************************************************************/
module.exports = function(rawdata) {
	var projectName = rawdata.Project.title[0];
	var data = rawdata.Project.Activity;
	//Split Activities by type - job(long process) and presentation(one-day event)
	var jobsArr = [];
	var presArr = [];
	var initData = getChildren(data);
	/*Collect all jobs into arrays by type
	*************************************/
	initData.forEach( function(job) {
		var start = moment(job.plannedStartDate[0].$.value).set('hour', 12);
		var end = moment(job.plannedEndDate[0].$.value).set('hour', 12);

		//check for type and collect
		if (job.hasOwnProperty("isMilestone")) {
			if (Boolean(job.isMilestone[0].$.value)) {
				presArr.push({
					work: job.title[0],
					start: start,
					end: end,
				});
			}
		}
		else {
				var type = (("string" in job) && ("_" in job.string[0])) ? job.string[0]._ : "";
				jobsArr.push({
					work: job.title[0],
					type: type,
					start: start,
					end: end
				});
			}
	});

	/* Calculate and assign positions
	*******************/
	//Sort jobs array by start date - to ensure nice ordering
	jobsArr.sort(function(obj1, obj2) {
		return obj1.start - obj2.start;
	});

	//Add relative positions to the jobs.
	var pos = 0;
	//Do until each job in array has position defined
	while (!jobsArr.every(function(obj) {return obj.hasOwnProperty("position");})) {

		currentDate = moment(0);
		jobsArr.forEach( function(job){
			if ((job.start.isAfter(currentDate)) && (!job.hasOwnProperty("position"))) {
				job.position = pos;
				currentDate = job.end;
			}
		});
		pos++;
	}

	/* Calculate colors.
	*  String matches defined in config, checks the job type first, if it's not defined - checks job title string.
	*  If no match found, assigns one of the color of reserve colors array
	***********************************************************/
	var getColor = function(obj, defIndex) {
		//If type of job is defined, use it. Else check name if name string matches the pattern.
		var string = (obj.type.length > 0) ? obj.type : obj.work;
		//var string = "comp";
		var color;
		//If no matches found, use one of reserve colors
		var reserveColors = c.colors.reserve;

		Object.keys(c.colors.match).forEach(function(clr) {
			c.colors.match[clr].forEach(function(str) {
				//Transform string to valid RegExp object
				var pattern = str.slice(str.indexOf("/")+1, str.lastIndexOf("/"));
				var flag = str.slice(str.lastIndexOf("/") + 1);
				var regExp = new RegExp(pattern, flag);

				if (regExp.test(string) ) {
					color = clr;
				}
			});
			return color;
		});

		if(!color) {
			color = reserveColors[defaultIndex];
			defaultIndex++; //referencing variable in outer scope, in order to access next item in array
			if (defaultIndex >= reserveColors.length) {
				defaultIndex = 0;
			}
		}
		return color;
	};

	/* Assign colors
	********************/
	var defaultIndex = 0;
	jobsArr.forEach(function(job) {
		var color = getColor(job, defaultIndex);
		job.color = color;
	});

	/* Create array of days, represented as objects - "timing".
	*  Set hour 12:00 for each day to avoid DST time confusions
	*********************************************************/
	var startDate = moment(jobsArr[0].start).isBefore(moment(presArr[0].start)) ?
		moment(jobsArr[0].start).set('hour', 12) : moment(presArr[0].start).set('hour', 12);
	var endDate = moment(jobsArr[jobsArr.length - 1].end).isAfter(moment(presArr[presArr.length - 1].end)) ?
		moment(jobsArr[jobsArr.length - 1].end).set('hour', 12) : moment(presArr[presArr.length - 1].end).set('hour', 12);
	var timing =  [];

	//Create objects with dates and empty works/presentations arrays
	for (var d = moment(startDate); d <= endDate; d.add(1, 'd')) {
		var currDate = moment(d);
		timing.push({
			date: currDate,
			dayOfWeek: dowText(currDate.isoWeekday()),
			works: [],
			presentations: []
		});
	}

	//Collect jobs
	jobsArr.forEach( function(job){
		currentDate = moment(job.start);
		endDate = moment(job.end);
		for (i = 0; i < timing.length; i++) {
			if(timing[i].date.isSame(currentDate)) {
				timing[i].works.push({
					name:  job.work,
					pos: job.position,
					color: job.color
				});
				currentDate.add(1, 'd');
				if (currentDate.isAfter(endDate)) {
					break;
				}
			}
		}
	});

	//Collect presentations
	presArr.forEach( function(job){
		currentDate = moment(job.start);
		endDate = moment(job.end);

		for (i = 0; i < timing.length; i++) {
			if(timing[i].date.isSameOrAfter(currentDate)) {
				timing[i].presentations.push(job.work);
				currentDate.add(1, 'd');
				if (currentDate.isAfter(endDate)) {
					break;
				}
			}
		}
	});

	/*
	* Set positions for works and presentations
	*******************************************/
	//Get maximum possible jobs/presentations per day
	var maxWorks = getMax("works", timing);
	var maxPres = getMax("presentations", timing);
	//Calculate number of excel rows
	var numRows = calcRows(timing);

	//Assign positions for presentations
	timing.forEach(function(day) {
		var startPosPres;
		if(day.works.length === 0) {
			startPosPres = Math.floor(numRows/2 -  day.presentations.length/2);
			day.presentations = day.presentations.map(function(pres) {
				presName = pres;
				return {
					name: presName,
					pos: startPosPres++
				};
			});
		}
		else {
			startPosPres = maxPres - 1;

			day.presentations = day.presentations.map(function(pres) {
			presName = pres;
				return {
					name: presName,
					pos: startPosPres--
				};
			});
		}
	});

	/* Add empty days to align weeks (always start from Monday and end on Sunday)
	****************************************************************************/
	var newdate;
	while (timing[0].dayOfWeek !== "Monday") {

		newdate = moment(timing[0].date).subtract(1, 'd');
		timing.unshift({
			date: newdate,
			dayOfWeek: dowText(newdate.isoWeekday()),
			works: [],
			presentations: []
		});
	}

	while (timing[timing.length - 1].dayOfWeek !== "Sunday") {
		newdate = moment(timing[timing.length - 1].date).add(1, 'd');

		timing.push({
			date: newdate,
			dayOfWeek: dowText(newdate.isoWeekday()),
			works: [],
			presentations: []
		});
	}

	/* Export project data
	***********************/
	return {
		projectName: projectName,
		height: numRows,
		maxworks: maxWorks,
		maxpres: maxPres,
		timing: timing
	};
};