var formSheet = "Form Responses 1"

function onEdit_o(e){
  var sh = e.range.getSheet();
  var ov = valueConvert(e.oldValue)
  var nv = valueConvert(e.value)
  
  
  if(sh.getName()==formSheet){
    if(e.range.getColumn() !== 8){
      sh.getRange(e.range.rowStart,8).setValue("Modified")
    }
    var status = sh.getRange(e.range.getRow(),2,1,7).getValues();
    var checkIn = status[0][0]
    var checkOut = status[0][1]
    var boxNumber = status[0][3]
    //  var lifetime = parseInt((checkOut - checkIn) / 1000 / 60 / 60 / 24)+1;
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservation");
    var dateList = ss.getRange(2, 2,1,ss.getMaxColumns()).getValues()
    if(status[0][6].toLowerCase() == "modified") {
      if(e.range.getColumn() == 2) {
        var weightResult = checkWeight(checkIn,checkOut)
        if (weightResult == 1){
          var col_in = getColumnIndex(ov)
          var col_nv = getColumnIndex(checkIn)
          var col_out = getColumnIndex(checkOut)
          var lifetime_before = parseInt((checkOut - ov) / 1000 / 60 / 60 / 24)+1;
          var lifetime_after = parseInt((checkOut - nv) / 1000 / 60 / 60 / 24)+1;
          if (col_nv !== -1) {
            cancelReserv_(boxNumber,col_in+2,lifetime_before)
            var boxList = ss.getRange("A:A").getValues();
            bookReserve(boxList,status,lifetime_after,col_nv+2)
            Logger.log(lifetime)
          }
          Browser.msgBox('Invalid Date or Expired Date', 'The change has been restored',Browser.Buttons.OK)        
          sh.getRange(e.range.rowStart,e.range.getColumn()).setValue(ov)
        } else {
          Browser.msgBox('Check-In date must be before Check-Out date', 'The change has been restored',Browser.Buttons.OK)        
          sh.getRange(e.range.rowStart,e.range.getColumn()).setValue(ov)
        }
      }
      if(e.range.getColumn() == 3){
        var col_out = getColumnIndex(ov)
        var col_nv = getColumnIndex(checkOut)
        var col_in = getColumnIndex(checkIn)
        var lifetime_before = parseInt((ov- checkIn) / 1000 / 60 / 60 / 24)+1;
        var lifetime_after = parseInt((nv - checkIn) / 1000 / 60 / 60 / 24)+1;
        //        Logger.log(lifetime_before)
        if (col_nv == -1) {                
          var pid = setNextDatePartial(checkOut,checkIn, col_out)
          cancelReserv_(boxNumber,col_in+2,lifetime_before)
          var boxList = ss.getRange("A:A").getValues();
          bookReserve(boxList,status,lifetime_after,col_in+2)
        } else {
          cancelReserv_(boxNumber,col_in+2,lifetime_before)
          var boxList = ss.getRange("A:A").getValues();
          bookReserve(boxList,status,lifetime_after,col_in+2)
        }
        Browser.msgBox('Information', 'The request is updated',Browser.Buttons.OK)
      }
    } else {
      Browser.msgBox('Status Error', 'Invalid Date range',Browser.Buttons.OK)
    }
  }
}

function onEdit(e) {
 
  var sh = e.range.getSheet();
  var rId = sh.getRange(e.range.rowStart,6).getValue();
  
  var msg = ['The Reservation # '+rId+' already been cancelled. Please add new booking',
             'Invalid Date range',
             'The cancel request is completed',
             'Direct edit not allowed for any booking status']
  
  if(sh.getName()==formSheet){
    if(e.range.getColumn() == 8) {
      if(e.value == 'Cancelled') {
        var sh = e.range.getSheet();
        var status = sh.getRange(e.range.getRow(),2,1,7).getValues();
        var checkIn = status[0][0]
        var checkOut = status[0][1]
        var boxNumber = status[0][3]
        var boxNumer = sh.getRange(e.range.getRow(),5).getValue();
        var lifetime_before = parseInt((checkOut - checkIn) / 1000 / 60 / 60 / 24)+1;
        var col_nv = getColumnIndex(checkIn)
        
        if (col_nv !== -1) {
          cancelReserv_(boxNumber,col_nv+2,lifetime_before)
          Browser.msgBox('Information', msg[2],Browser.Buttons.OK)
          sh.getRange(e.range.rowStart,1,1,8).protect().setDescription(e.range.rowStart+' protected range');
          sh.getRange(e.range.rowStart,1,1,8).setBackground("#ff6269")
        }
        
      } else 
//        if(e.value == "Confirmed"){
//        sh.getRange(e.range.rowStart,e.range.getColumn()).setValue(e.oldValue)
//        Browser.msgBox('Error', msg[1],Browser.Buttons.OK)
//      } else 
      {
        sh.getRange(e.range.rowStart,e.range.getColumn()).setValue(e.oldValue)
        Browser.msgBox('Error', msg[0],Browser.Buttons.OK)
      }
    } else {
      
      Browser.msgBox('Error', msg[3],Browser.Buttons.OK)
      
    }
    if ((e.range.getColumn() == 2) || (e.range.getColumn() == 3)) {
      sh.getRange(e.range.rowStart,e.range.getColumn())
                                     .setValue(e.oldValue)
                                     .setNumberFormat("dd/MM/yyyy HH:mm:ss")
    }
  }
}
