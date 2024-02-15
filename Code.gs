function testing(){
  var sheetindex = 6;
  var start = 5; //should be the labels row
  var end = 313;
  dupeCheck(sheetindex,start,end)
  doeverything(true, sheetindex,start,end)
}
// function parameters
// dupeCheck(sheetIndex, start, end)
// sheetIndex  (int) - index of Sheet to check
// start  (int) - starting Row (start at the labels row, 1 before actual entry)
// end (int)  - last actual row for entry
function dupeCheckRun(){
  dupeCheck(6,5,313)
}
function dupeCheck(sheetIndex, start, end){
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var sheetData = sheets[sheetIndex].getDataRange().getValues(); // change this index depending on which sheet
  var startRow = start; // has to be 1 before actual data starts aka the labels row
  var lastRow = end // last actual dataRow
  const dupes = new Map();
  var copy = ""
  for(var a = startRow; a < lastRow; a++) {
    var repName = sheetData[a][2].toString().trim().split(" ")[0]
    if(dupes.has(sheetData[a][1].toString())) {
      dupes.set(sheetData[a][1].toString(), dupes.get(sheetData[a][1].toString()) + " & " + repName)
      copy += "Row " + (a+1) + ": " + sheetData[a][1].toString() +" ("+ dupes.get(sheetData[a][1].toString()) + ")\n"
    }
    else {
      dupes.set(sheetData[a][1].toString(), repName)
    }
  }
  Logger.log(copy)
}
function doeverything(update, sheetindex,startrow,endrow){
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var repshow = new Map(); // names of reps : number of shown appts
  var repsold = new Map(); // name of rep : units sold
  var dateshow= new Map(); // date : shown appts
  var dateappts = new Map(); // date : total appts
  var repdeposit = new Map(); // name of reps : deposit
  var reppayment = new Map(); // name of rep : pending payment
  var repverify = new Map(); // name of rep : pending verification
  var repDNH = new Map(); // name of rep : DNH
  var repNSF = new Map(); // name of rep : NSF
  var repOutstanding = new Map(); // name of rep : outstanding (total DNH,NSF,verification or payment)
  var sheetData = sheets[sheetindex].getDataRange().getValues();

  for(var a = startrow; a < endrow; a++) {
    var date = getFullDate(sheetData[a][0].toString())
    var repName = sheetData[a][2].toString().trim().split(" ")[0]
    var show = sheetData[a][3].toString().toLowerCase() === "x"

    var sold = sheetData[a][9].toString().toLowerCase() !== ""
    var elevate = sheetData[a][10].toString()
    var dep = sheetData[a][11].toString() !== ""
    
    // rep : show
    if(!repshow.has(repName)){
      repshow.set(repName,0)
      repsold.set(repName,0)
      repdeposit.set(repName,0)
      reppayment.set(repName,0)
      repverify.set(repName,0)
      repDNH.set(repName,0)
      repNSF.set(repName,0)
      repOutstanding.set(repName,0)
    } if (show){
      repshow.set(repName, repshow.get(repName) + 1)
    }

    // rep : sold
    if (sold && !repsold.has(repName)) {
      repsold.set(repName, 1);
    } else if (sold && repsold.has(repName)) {
      repsold.set(repName, repsold.get(repName) + 1);
    }

    // date : shown appts
    if(!dateshow.has(date)){
      dateshow.set(date,0)
    } if (show){
      dateshow.set(date, dateshow.get(date) + 1)
    }

    // date : total appts
    if(!dateappts.has(date)){
      dateappts.set(date,1)
    } else {
      dateappts.set(date, dateappts.get(date) + 1)
    }

    // rep : deposit
    if (dep && !repdeposit.has(repName)) {
      repdeposit.set(repName, 1);
    } else if (dep && repdeposit.has(repName)) {
      repdeposit.set(repName, repdeposit.get(repName) + 1);
    }

    // rep : elevate
    if(elevate === "Pending Payment"){
      reppayment.set(repName, reppayment.get(repName) + 1)
      repOutstanding.set(repName, repOutstanding.get(repName) + 1)
    }
    else if(elevate === "Pending Verification"){
      repverify.set(repName, repverify.get(repName) + 1)
      repOutstanding.set(repName, repOutstanding.get(repName) + 1)
    }
    else if(elevate === "DNH"){
      repDNH.set(repName, repDNH.get(repName) + 1)
      repOutstanding.set(repName, repOutstanding.get(repName) + 1)
    }
    else if(elevate === "NSF"){
      repNSF.set(repName, repNSF.get(repName) + 1)
      repOutstanding.set(repName, repOutstanding.get(repName) + 1)
    }
    else if(elevate !== "") {
      Logger.log(repName + " needs to fix - " + a + " " + sheetData[a][1].toString())
    }
    /*
    if(show && repName==="Randy"){
      Logger.log(sheetData[a][1].toString())
    }
    */
  }
  if(update){ 
    var sumrepshow = mapSumValues(repshow);
    repshow.set("Total", sumrepshow);
    var sumrepsold = mapSumValues(repsold);
    repsold.set("Total", sumrepsold);
    var sumdateshow = mapSumValues(dateshow);
    dateshow.set("Total", sumdateshow);
    var sumdateappts = mapSumValues(dateappts);
    dateappts.set("Total", sumdateappts);
    var sumrepdeposit = mapSumValues(repdeposit);
    repdeposit.set("Total", sumrepdeposit);
    var sumreppayment = mapSumValues(reppayment);
    reppayment.set("Total", sumreppayment);
    var sumrepverify = mapSumValues(repverify);
    repverify.set("Total", sumrepverify);
    var sumrepDNH = mapSumValues(repDNH);
    repDNH.set("Total", sumrepDNH);
    var sumrepNSF = mapSumValues(repNSF);
    repNSF.set("Total", sumrepNSF);
    var sumrepOustanding = mapSumValues(repOutstanding)
    repOutstanding.set("Total", sumrepOustanding);

    var closerate = mapDivideValues(repsold,repshow)
    var showrate = mapDivideValues(dateshow, dateappts)
    var augmentedSales = sumValuesAcrossMaps(repsold,reppayment,repverify,repDNH,repNSF)
    var augmentedcloserate = mapDivideValues(augmentedSales,repshow)

    putKeys(1,"A", dateshow)
    putValues(1,"B", dateshow)
    putValues(1,"C", dateappts)
    putValues(1,"D",showrate)
    putKeys(1,"E",repsold)
    putValues(1,"F", repOutstanding)
    putValues(1,"G",repsold)
    putValues(1,"H",repshow)
    putValues(1,"I", closerate)
    putValues(1,"J", augmentedcloserate)

    sheets[1].getRange(`A${dateshow.size+3}`).setValue(getCurrentDate())

    var rowEntry = repsold.size + 2;
    putKeys(rowEntry, "E", reppayment)
    putValues(rowEntry, "F", reppayment)
    putValues(rowEntry, "G", repverify)
    putValues(rowEntry, "H", repDNH)
    putValues(rowEntry, "I", repNSF)
    putValues(rowEntry, "J", augmentedSales)
    putValues(rowEntry,"K", repdeposit)
    
  }
}
function mapDivideValues(map1, map2) {
  const result = new Map();

  for (const [key, value2] of map2) {
    if (map1.has(key)) {
      const value1 = map1.get(key);
      result.set(key, value1 / value2);
    }
  }
  return result;
}
function sumValuesAcrossMaps(...maps) {
  const result = new Map();

  for (const map of maps) {
    for (const [key, value] of map) {
      if (!result.has(key)) {
        result.set(key, value);
      } else {
        result.set(key, result.get(key) + value);
      }
    }
  }

  return result;
}
function mapSumValues(map){
  var sum = 0
  for (const [key, value] of map) {
    sum += value
  }
  return sum
}
function putValues(x,y,z){
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var a = x
  for (const [key, value] of z) {
    var cell = sheets[1].getRange(`${y}${a+1}`)
    cell.setValue(value)
    a++
  }
}

function putKeys(x,y,z){
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var a = x
  for (const [key, value] of z) {
    var cell = sheets[1].getRange(`${y}${a+1}`)
    cell.setValue(key)
    a++
  }
}

function getReps(){
  var reps = new Map();
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var sheetData = sheets[5].getDataRange().getValues();

  for(var a = 5; a < 316; a++) {
    var repName = sheetData[a][2].toString().trim().split(" ")[0]
    if(!reps.has(repName)){
      reps.set(repName,0)
      Logger.log(a + " " + sheetData[a][0].toString() + " " + sheetData[a][1].toString() + " "+ repName)
    }
    else{
      reps.set(repName, reps.get(repName) + 1)
    }
  }
  Logger.log([...reps.entries()])
}

function getDates(){
  var dateList = new Map();
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var sheetData = sheets[5].getDataRange().getValues();

  for(var a = 5; a < 316; a++) {
    var date = getFullDate(sheetData[a][0].toString())
    if(!dateList.has(date)){
      dateList.set(date,0)
      Logger.log(a + " " + sheetData[a][0].toString() + " " + sheetData[a][1].toString())
    }
    else{
      dateList.set(date, dateList.get(date) + 1)
    }
  }
  Logger.log([...dateList.entries()])
}

function getFullDate(date) {
  var dateSplit = date.split(" ");
  if (dateSplit[0] === "Sun") {
    return "Sunday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Mon") {
    return "Monday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Tue") {
    return "Tuesday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Wed") {
    return "Wednesday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Thu") {
    return "Thursday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Fri") {
    return "Friday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else if (dateSplit[0] === "Sat") {
    return "Saturday, " + dateSplit[1] + " " + dateSplit[2] + ", " + dateSplit[3];
  } else {
    return "Invalid day";
  }
}
function runPutDates(){
  putdates(8)
}


function deposits(startrow, endrow, sheetindex, depositCOLUMN){
  var namesList = ["Sean", "Cade", "Randy", "Jason", "Christian", "Patrick", "Joseph", "Smokey", "Karim", "Saloh", "Tyler", "Chad", "Daniel", "ChargedUp", "ChargedUpTheU", "Labrekia", "Marshall", "Brady"];
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var deposit = new Map();

  var startRow = startrow; // has to be 1 before actual data starts aka the labels row
  var lastRow = endrow // last actual dataRow
  var sheetIndex = sheetindex;
  var sheetData = sheets[sheetIndex].getDataRange().getValues();
  var pplList = ""

  for(var a = 0; a < namesList.length; a++) {
    deposit.set(namesList[a],0)
  }

  for(var a = startRow; a < lastRow; a++) {
    var rep = sheetData[a][2].toString().split(" ")[0].toString()
    var dep = sheetData[a][depositCOLUMN].toString() !== ""
    //Logger.log(a+" " + sheetData[a][1].toString())
    if(dep){
      // Logger.log(a + " " + rep + " " + sheetData[a][0].toString() + " " + sheetData[a][1].toString())
      pplList += sheetData[a][1].toString() + "\n"
      deposit.set(rep, deposit.get(rep) + 1)
    }
    
  }
  Logger.log(pplList)
  Logger.log([...deposit.entries()])
  dummyTest(31,"G",deposit)
}

function elevate(startrow, endrow, sheetindex, elevateCOLUMN){
  var namesList = ["Sean", "Cade", "Randy", "Jason", "Christian", "Patrick", "Joseph", "Smokey", "Karim", "Saloh", "Tyler", "Chad", "Daniel", "ChargedUp", "ChargedUpTheU", "Labrekia", "Marshall", "Brady"];
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var payment = new Map();
  var verified = new Map();
  var nsf = new Map();
  var dnh = new Map();

  var startRow = startrow; // has to be 1 before actual data starts aka the labels row
  var lastRow = endrow // last actual dataRow
  var sheetIndex = sheetindex;
  var sheetData = sheets[sheetIndex].getDataRange().getValues();
  var pplList = ""

  for(var a = 0; a < namesList.length; a++) {
    payment.set(namesList[a],0)
    verified.set(namesList[a],0)
    nsf.set(namesList[a],0)
    dnh.set(namesList[a],0)
  }

  for(var a = startRow; a < lastRow; a++) {
    //Logger.log(a + " " + sheetData[a][1].toString())
    var rep = sheetData[a][2].toString().split(" ")[0].toString()
    var dep = sheetData[a][elevateCOLUMN].toString()

    if(dep === "NSF"){
      //Logger.log("NSF" + " " + a + " " + sheetData[a][1].toString())
      nsf.set(rep, nsf.get(rep) + 1)
    }
    else if(dep === "DNH"){
      //Logger.log("DNH" + " " + a + " " + sheetData[a][1].toString())
      dnh.set(rep, dnh.get(rep) + 1)
    }
    else if(dep === "Pending Verification"){
      //Logger.log("Pending Verification" + " " + a + " " + sheetData[a][1].toString())
      verified.set(rep, verified.get(rep) + 1)
    }
    else if(dep === "Pending Payment"){
      //Logger.log("Pending Payment" + " " + a + " " + sheetData[a][1].toString())
      payment.set(rep, payment.get(rep) + 1)
    }
    else if(dep !== "") {
      Logger.log(rep + " needs to fix - " + a + " " + sheetData[a][1].toString())
    }
    
  }
  dummyTest(31,"H",payment)
    dummyTest(31,"I",verified)
    dummyTest(31,"J",dnh)
    dummyTest(31,"K",nsf)
    Logger.log([...payment.entries()])
    Logger.log([...verified.entries()])
    Logger.log([...dnh.entries()])
    Logger.log([...nsf.entries()])
}
function getCurrentDate() {
  let currentDate = new Date();
  var hour = currentDate.getHours()
  var ampm = "AM"
  var seconds = currentDate.getSeconds()
  var minutes = currentDate.getMinutes();
  seconds = seconds < 10 ? "0"+seconds : seconds
  minutes = minutes < 10 ? "0"+minutes : minutes
  if(currentDate.getHours()>=12){
    hour -= 12
    ampm = "PM"
    hour = hour == 0 ? 12 : hour
  }
  
  Logger.log(hour + ":" + minutes + ":" + seconds + " " + ampm)
  var date = (currentDate.getMonth()+1) + "/" + currentDate.getDate() + "/" + currentDate.getFullYear()
  return "Last Updated: " + date + " " + hour + ":" + minutes+ ":" + seconds + " " + ampm + " CST"
}

function putdates(month) {
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();

  for(var a = 1; a <= 31; a++){
    var cell = sheets[0].getRange(`A${a+1}`)
    cell.setValue(`${month}/${a}/2023`)
  }
}
function changeAgentToNum(){
  var start = 6;
  var end = 236;
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  for(var a = start; a < end; a++){
    var cell = sheets[1].getRange(`C${a+1}`)
    //Logger.log(cell.getValue())
    cell.setValue(numToName(cell.getValue()))
  }
}
function checkNames() {
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var sheetData = sheets[3].getDataRange().getValues(); // change this index depending on which sheet
  var startRow = 5; // has to be 1 before actual data starts aka the labels row
  var lastRow = 330 // last actual dataRow
  const dupes = new Set();
  
  for(var a = startRow; a < lastRow-1; a++) {
    dupes.add(sheetData[a][2].toString())

  }
  Logger.log(new Array(...dupes).join(', '))
  Logger.log(dupes.size)
}
  // use a = 1, if that's where the labels for the data are
  // B ,a=1for dateShow, C,a=1 for totalDateShow, 
  // G, a=1 for dayShow, H, a=1 for totalDayShow
  // G, a=10 for closedRep, h,a=10 for totalShownRep
function dummyTest(x,y,z) {
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var a = x
  for (const [key, value] of z) {
    var cell = sheets[0].getRange(`${y}${a+1}`)
    cell.setValue(value)
    a++
  }
}


// function parameters
// dateCheck(sheetIndex, start, end, print)
// sheetIndex (int) - index of Sheet to check
// start  (int) - starting Row (start at the labels row, 1 before actual entry)
// end  (int) - last actual row for entry
// print (boolean) - each row until error to catch where a runtime error occurs
// if prints nothing when print==false, then there are no errors, and checks by year
function dateCheckRun() {
  dateCheck(4,6,228,false) // will encounter runtime error
}

function dateCheck(sheetIndex, start, end, print) {
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();
  var sheetData = sheets[sheetIndex].getDataRange().getValues(); // change this index depending on which sheet
  var startRow = start; // has to be 1 before actual data starts aka the labels row
  var lastRow = end // last actual dataRow
  

  for(var a = startRow; a < lastRow-1; a++) {
    if(print){
      Logger.log(`On Row ${a}`)
    }
    var day = sheetData[a][0].toString().split(" ")[0]
    var date = Number(sheetData[a][0].toString().split(" ")[2].toString())
    var year = Number(sheetData[a][0].toString().split(" ")[3].toString())
    var name = sheetData[a][1].toString()
    var rep = sheetData[a][2].toString()
    if(year != 2023){
      Logger.log(`ERROR for YEAR: Row ${a+1} for ${name}, year is ${year}, Appt Owner: ${rep}`)
    }
  }
}

// function parameters
// dupeCheck(sheetIndex, start, end)
// sheetIndex - index of Sheet to check
// start - starting Row (start at the labels row, 1 before actual entry)
// end - last actual row for entry


function dummy() {
  var a = 10
  for (const [key, value] of dateShow) {
    var cell = sheets[2].getRange(`B${a+1}`)
    cell.setValue(value)
    a++
  }
}

function combineSheetData()
{
  var doc = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = doc.getSheets();

  var startSheet = 1
  var endSheet = 2;
  
  for(var day = 1; day < 32; day++){
    var total = 0
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`B${day+1}`).getValue())
    }
    sheets[0].getRange(`B${day+1}`).setValue(total)

    total = 0;
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`C${day+1}`).getValue())
    }
    sheets[0].getRange(`C${day+1}`).setValue(total)
  }
  
  for(var day = 1; day < 8; day++){
    var total = 0
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`G${day+1}`).getValue())
    }
    sheets[0].getRange(`G${day+1}`).setValue(total)

    total = 0;
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`H${day+1}`).getValue())
    }
    sheets[0].getRange(`H${day+1}`).setValue(total)
  }
  // if number of agents increase, change 18 to 19
  for(var day = 10; day < 18; day++){
    var total = 0
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`G${day+1}`).getValue())
    }
    sheets[0].getRange(`G${day+1}`).setValue(total)

    total = 0;
    for(var a = startSheet; a <= endSheet; a++){
      total += Number(sheets[a].getRange(`H${day+1}`).getValue())
    }
    sheets[0].getRange(`H${day+1}`).setValue(total)
  }
}


function numToName(num){
  var i = Number(num)
  if(i === 7)
    return "Sean"
  else if(i === 21)
    return "Daniel"
  else if(i === 46)
    return "Cade"
  else if(i === 54)
    return "Randy"
  else if(i === 56)
    return "Jason"
  else if(i === 59)
    return "Christian"
  else if(i === 60)
    return "Patrick"
  else if(i === 61)
    return "Joseph"
  else {
    Logger.log("unknown agent")
    return i
  }
    
}



function looper(){
  var namesList = ["Sean", "Daniel", "Cade", "Randy", "Jason", "Patrick", "Christian", "Joseph"];
  
  var closedRep = new Map();
  for(var a = 0; a < namesList.length; a++) {
    closedRep.set(namesList[a],0)
  }
  Logger.log([...closedRep.entries()])

}
