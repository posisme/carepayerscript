var milesrate = .56;

function onOpen() {
	var menu = SpreadsheetApp.getUi().createMenu("Update");
	menu.addItem("Update", "update").addToUi();
	menu.addItem("Start Over", "startover").addToUi();
	//var html = "<h1>Make Paycheck</h1>\
	//<p>Pay Date: <input type='text' id='dt' /></p>\
	//<p><input type='button' value='Go' onclick='google.script.run.getPay()' /></p>";
	//var sb = HtmlService.createHtmlOutput(html);
	//SpreadsheetApp.getUi().showSidebar(sb);
}
//function getPay(){
//
//}

function startover() {
	var really = SpreadsheetApp.getUi().alert("This will erase the spreadsheet. Are you sure you want to do this?", SpreadsheetApp.getUi().ButtonSet);
	if (really == "OK") {
		getstuff("startover");
	}
}

function getstuff(status) {
	var app = SpreadsheetApp;
	var sheet = app.getActive().getSheetByName("Paystubs");
	var cal = CalendarApp.getCalendarsByName("Childcare")[0];
	if (status == "startover") {
		sheet.clear();
		sheet.appendRow([
				"Week",
				"Hours",
				"Gross Pay",
				"Withholdings",
				"Mileage Reimbursement",
				"Total Pay",
				"Payee",
				"Paid"
			]);
	}
	var p = app.getUi().prompt("Month (English)");
	if (p.getResponseText() == "all") {
		var year = app.getUi().prompt("Year");
		sd = new Date("January 01, " + year.getResponseText() + " 6:00:00 -5:00");
		ed = new Date("December 31, " + year.getResponseText() + " 24:00:00 -5:00");
	} else {
		//var sd = "November 26,2017 11:00:00 -5:00";
		//var ed = SpreadsheetApp.getUi().prompt("End Date");
		//var ed = "December 22, 2017 21:00:00 -5:00";
		var today = new Date();
		var sd = new Date(p.getResponseText() + " 01, " + today.getFullYear() + " 6:00:00 -5:00");
		if (today > sd) {
			sd = new Date(p.getResponseText() + " 01, " + (today.getFullYear() + 1) + " 6:00:00 -5:00");
		}
		var ed = new Date(sd);
		while (sd.getDay() != 0) {
			sd.setDate(sd.getDate() + 1);
		}

		ed.setMonth(ed.getMonth() + 1);
		ed.setDate(ed.getDate() - 1);
		ed.setHours(24);
		while (ed.getDay() != 6) {
			ed.setDate(ed.getDate() + 1);
		}
	}
	var rate;
	var strw = 2;
	while (sd < ed) {
		var ned = new Date(sd);
		ned.setDate(ned.getDate() + 6);
		var row = [Utilities.formatDate(sd, "America/Chicago", "MM-dd-yy") + " to " + Utilities.formatDate(ned, "America/Chicago", "MM-dd-yy"), 0, 0, 0, 0, 0, ""];

		var events = cal.getEvents(sd, ned);

		for (i = 0; i < events.length; i++) {
			//cal.getEventById(events[i].getId()).setDescription(events[i].getDescription()+"\nssadd\n");
			//if(status == "update" && !events[i].getDescription().match(/ssadd/)){


			var emp = "";
			var pay = 0;
			var hours = ((new Date(events[i].getEndTime()) - new Date(events[i].getStartTime())) / 1000 / 60 / 60);
			if (events[i].getDescription().match(/regular pay/)) {
				rate = 10;
			} else if (events[i].getDescription().match(/\$\d+\.\d+ per hour/)) {
				rate = parseInt(events[i].getDescription().match(/\$(\d+\.\d+) per hour/)[1]);
			}
			pay = hours * rate;
			var miles = 0;
			if (events[i].getDescription().match(/\d miles/)) {
				miles = parseInt(events[i].getDescription().match(/(\d) miles/)[1]);
			}

			if (events[i].getDescription().match(/Name:(.*)\n/)) {
				row[6] = events[i].getDescription().match(/Name:(.*)\n/)[1].trim();
			}
			row[4] += miles * milesrate;
			row[3] += pay * .2;
			row[1] += hours;
			row[2] += pay;
			row[5] += (pay - (pay * .2)) + (miles * milesrate);
			//}
		}
		sheet.appendRow(row);
		sd = new Date(ned);
		sd.setDate(sd.getDate() + 1);
	}

}

function update() {
	getstuff("update");
}
