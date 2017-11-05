var milesrate = .56;

function onOpen() {
	var menu = SpreadsheetApp.getUi().createMenu("Update");
	menu.addItem("Get Timecard", "getstuff").addToUi();
  menu.addItem("Pay", "pay").addToUi();
}


function getstuff() {
	var app = SpreadsheetApp;
	var sheet = app.getActive().getSheetByName("From Calendar");
  var paidsheet = app.getActive().getSheetByName("Paid");
  var paid = paidsheet.getDataRange().getValues();
  var paidsheet = {};
  for(i=0;i<paid.length;i++){
    paidsheet[paid[i][0]] = paid[i];
  }
	var cal = CalendarApp.getCalendarsByName("Childcare")[0];
	sheet.clear();
		sheet.appendRow([
				"Week",
				"Hours",
				"Gross Pay",
				"Withholdings",
				"Mileage Reimbursement",
				"Total Pay",
				"Payee",
          "Pay"
			]);
	
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
		var row = [Utilities.formatDate(sd, "America/Chicago", "MM-dd-yy") + " to " + Utilities.formatDate(ned, "America/Chicago", "MM-dd-yy"), 0, 0, 0, 0, 0, "",""];

		var events = cal.getEvents(sd, ned);

		for (i = 0; i < events.length; i++) {
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

		}
      if(paidsheet[row[0]]){
        row[7] = "Paid";
      }
		sheet.appendRow(row);
		sd = new Date(ned);
		sd.setDate(sd.getDate() + 1);
	}
  sheet.getRange("H2:H99").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Paid','Pay Now','Pay Later'], true));
  
  
  
}


function pay(){
  var app = SpreadsheetApp;
	var sheet = app.getActive().getSheetByName("From Calendar");
  var paysheet = app.getActive().getSheetByName("Paid");
	var cal = CalendarApp.getCalendarsByName("Childcare")[0];
  var data = sheet.getDataRange().getValues();
  for(i=0;i<data.length;i++){
    if(data[i][7] == "Pay Now"){
      data[i].pop();
      paysheet.appendRow(data[i]);
      sheet.getRange(i+1, 8).setValue("Paid");
    }
  }
  
  
}