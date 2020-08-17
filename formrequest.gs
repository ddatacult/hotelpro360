var yesterDay = Utilities.formatDate(new Date(), "GMT+5.30", "dd")-1;
var mString = Utilities.formatDate(new Date(), "GMT+1", "MMMM");
var year = new Date().getYear();
var month = new Date().getMonth()+1;
var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");

function addReservation(e) {
//  var formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
//  var records = formSheet.getRange(2, 2, 1, 7).getValues() //details from Form Response Sheet
  var items = e.response.getItemResponses();
  var record = []
  for (var j = 0; j < items.length; j++) {
    var itemResponse = items[j];
    var item = ['%s',itemResponse.getResponse()]    
    record.push(item[1])
    
    }
  
  var records = [record]
//  Logger.log(records)
  var cName = records[0][2];
  var roomNo = records[0][5]
  var resNumber = records[0][4]
  var boxNum = records[0][3];
  var checkIn = formatDate(records[0][0]);
  var checkOut = formatDate(records[0][1]);
//  var checkIn = records[0][0];
//  var checkOut = records[0][1];
  Logger.log(checkIn)
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservation");
  var dateList = ss.getRange(2, 2,1,ss.getMaxColumns()).getValues()  //date row details collection. dont change this
  if(records[0][6].toLowerCase() == "confirmed") {
    var weightResult = checkWeight(checkIn,checkOut)
    if (weightResult == 1){
      var col_in = getColumnIndex(checkIn)
      var col_out = getColumnIndex(checkOut)
      var lifetime = parseInt((checkOut - checkIn) / 1000 / 60 / 60 / 24)+1;
//      Logger.log(lifetime)
      var colStart = col_in;
      if( (col_in == -1) && (col_out == -1) ){        // Date Not Found (Check-in and Check-out)       
        colStart = setNextDateNotFound(checkOut,checkIn)
      }
     
      if( (col_in > -1) && (col_out == -1) ){         // Partial date Found (Check-in only)
        colStart = setNextDatePartial(checkOut,checkIn, col_in)
      }
//      Logger.log(colStart)
      var boxNumber = ss.getRange("A:A").getValues()
      bookReserve(boxNumber,records,lifetime,colStart+2) // Date Found (Check-in and Check-out)  
      
    } else {
      
      var rowId = formSheet.getLastRow()
      formSheet.getRange(rowId, 2,1,2).setBackground("#f6f917")
      formSheet.getRange(rowId, 8).setValue("Invalid Dates")
    }
  } 
  days365Check()
}

