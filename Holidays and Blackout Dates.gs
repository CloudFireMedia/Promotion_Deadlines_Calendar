//
//// This script requires Advance Calendar Service 
//
//var objSet = {
//	strSheetName: 'Holidays + Blackout Dates',
//	strWorkCalendarID: 'vqf94jnhisf3onqlhu2mcoupvo@group.calendar.google.com',//[ 2 ] Work
//	strCCNConcertsAndSpecialEventsCalendarID: '4rcjvpcmkv73nn5t6s3hmtkuqo@group.calendar.google.com',//CCN Concerts and Special Events
//	strarrCalendarsID: ['vqf94jnhisf3onqlhu2mcoupvo@group.calendar.google.com', '0t6n7ctb4g0ke81tg6lncgtif4@group.calendar.google.com', 'sf9ub89maekcldmqg1vn3lfomo@group.calendar.google.com', 'rpa5jpfm1psq19f69c1p6n280s@group.calendar.google.com', '1lau0mdfsk6sst26hqt0toh8jo@group.calendar.google.com'],//[ 2 ] Work, CCN Classes and Groups, CCN Children and Youth, CCN All-Church Community Events,CCN NEW! Events
//}
//
//function churchCalendarBlackoutDate() {
//
//	var counter = 0;
//	var sp = SpreadsheetApp.getActiveSpreadsheet();
//	var sh = sp.getSheetByName(objSet.strSheetName);
//	var data = sh.getDataRange().getValues();
//
//	var currentDate = new Date();
//	var currentYear = currentDate.getFullYear();
//
//	for (var i = 0; i < data[1].length; i++) {
//
//		if (data[1][i] == currentYear) {
//			var refCol1 = i;
//			break;
//		}
//
//	}
//
//	if (refCol1) {
//		for (var j = 0; j < objSet.strarrCalendarsID.length; j++) {
//			var cal = CalendarApp.getCalendarById(objSet.strarrCalendarsID[j]);
//
//			//Logger.log(refCol1)
//
//			for (var i = 3; i < data.length; i++) {
//
//				if (('' + data[i][3]).trim().toUpperCase() == "YES") {
//
//					var startDate = data[i][refCol1];
//					startDate.setHours(0)
//					startDate.setMinutes(0)
//					startDate.setSeconds(0)
//					startDate.setMilliseconds(0)
//
//					var endDate = data[i][refCol1 + 1];
//					endDate.setHours(0)
//					endDate.setMinutes(0)
//					endDate.setSeconds(0)
//					endDate.setMilliseconds(0)
//
//					var eDate = new Date(endDate.valueOf() + 1000 * 3600 * 24);
//
//					if (startDate.valueOf() >= currentDate.valueOf()) {
//
//						var events = cal.getEvents(startDate, eDate)
//							if (events.length > 0) {
//
//								//Logger.log(events.length)
//								//Logger.log(startDate)
//
//								for (var k = events.length - 1; k >= 0; k--) {
//									if (!events[k].isAllDayEvent()) {
//
//										events[k].deleteEvent();
//										counter++;
//									}
//
//								}
//							}
//					}
//				}
//			}
//		}
//		sp.toast("Deleted " + counter + " events.", "Info");
//	}
//}
//
//function churchCalendarAllDayEvent() {
//
//	var counter = 0;
//	var sp = SpreadsheetApp.getActiveSpreadsheet();
//	var sh = sp.getSheetByName(objSet.strSheetName);
//	var data = sh.getDataRange().getValues();
//
//	var currentDate = new Date();
//	var currentYear = currentDate.getFullYear();
//
//	for (var i = 0; i < data[1].length; i++) {
//
//		if (data[1][i] == currentYear) {
//			var refCol1 = i;
//			break;
//		}
//
//	}
//
//	for (var i = refCol1 + 1; i < data[1].length; i++) {
//
//		if (data[1][i] == currentYear) {
//			var refCol2 = i;
//			break;
//		}
//
//	}
//
//	if (refCol1 && refCol2) {
//		var cal = CalendarApp.getCalendarById(objSet.strCCNConcertsAndSpecialEventsCalendarID);
//
//		//Logger.log(refCol1)
//		//Logger.log(refCol2)
//		for (var i = 3; i < data.length; i++) {
//
//			if (('' + data[i][2]).trim().toUpperCase() == "YES" && ('' + data[i][refCol2 + 1]).trim() == "") {
//
//				var startDate = data[i][refCol1];
//				var endDate = data[i][refCol1 + 1];
//
//				if (startDate.valueOf() >= currentDate.valueOf()) {
//					var title = '' + data[i][0];
//
//					if (endDate > startDate) {
//						var eDate = new Date(endDate.valueOf() + 1000 * 3600 * 24);
//						var event = cal.createAllDayEvent(title, startDate, eDate);
//					} else {
//						var event = cal.createAllDayEvent(title, startDate);
//					}
//					var evID = event.getId();
//					//var evS = evID.replace("@google.com", "");
//					sh.getRange(i + 1, refCol2 + 2, 1, 1).setValue(evID);
//					counter++;
//					//var event = Calendar.Events.get(objSet.strWorkCalendarID, evS);
//					//event.transparency = "opaque";
//					//Calendar.Events.patch(event, objSet.strWorkCalendarID, evS);
//
//				}
//			}
//		}
//
//	}
//	sp.toast("Created " + counter + " events.", "Info");
//
//}
//
//function paidHoliday() {
//
//	var counter = 0;
//	var sp = SpreadsheetApp.getActiveSpreadsheet();
//	var sh = sp.getSheetByName(objSet.strSheetName);
//	var data = sh.getDataRange().getValues();
//
//	var currentDate = new Date();
//	var currentYear = currentDate.getFullYear();
//
//	for (var i = 0; i < data[1].length; i++) {
//
//		if (data[1][i] == currentYear) {
//			var refCol1 = i;
//			break;
//		}
//
//	}
//
//	for (var i = refCol1 + 1; i < data[1].length; i++) {
//
//		if (data[1][i] == currentYear) {
//			var refCol2 = i;
//			break;
//		}
//
//	}
//
//	if (refCol1 && refCol2) {
//		var cal = CalendarApp.getCalendarById(objSet.strWorkCalendarID);
//
//		//Logger.log(refCol1)
//		//Logger.log(refCol2)
//		for (var i = 3; i < data.length; i++) {
//
//			if (('' + data[i][1]).trim().toUpperCase() == "YES" && ('' + data[i][refCol2]).trim() == "") {
//
//				var startDate = data[i][refCol1];
//				var endDate = data[i][refCol1 + 1];
//
//				if (startDate.valueOf() >= currentDate.valueOf()) {
//					var title = '' + data[i][0] + ' (Paid Vacation)';
//
//					if (endDate > startDate) {
//						var eDate = new Date(endDate.valueOf() + 1000 * 3600 * 24);
//						var event = cal.createAllDayEvent(title, startDate, eDate);
//					} else {
//						var event = cal.createAllDayEvent(title, startDate);
//					}
//					var evID = event.getId();
//					var evS = evID.replace("@google.com", "");
//					sh.getRange(i + 1, refCol2 + 1, 1, 1).setValue(evID);
//
//					var event = Calendar.Events.get(objSet.strWorkCalendarID, evS);
//					event.transparency = "opaque";
//					Calendar.Events.patch(event, objSet.strWorkCalendarID, evS);
//					counter++;
//				}
//			}
//		}
//
//	}
//	sp.toast("Created " + counter + " events.", "Info");
//
//}
//
//function debug() {
//	var event = Calendar.Events.get('4hfqdq65aju0ei55diuuasgfuc@group.calendar.google.com', '26fr6fdekqmoigvqo201ssghtk')
//
//}