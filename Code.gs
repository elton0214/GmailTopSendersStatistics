// connect and activate "D3.html"
function doGet() {
  const html = HtmlService.createHtmlOutputFromFile('D3')
              .setSandboxMode(HtmlService.SandboxMode.IFRAME)
              .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=2.0, user-scalable=yes')
              .setTitle('Test D3')

  return html;
}


// Parameters Setting
const numberOfMail       = 30;      // if 99999 then regard as all mails
const numberOfTopSenders = 10;
const foldername         = 'Temp';  // folder to place the sheet of static result
const lineNotifyToken    = '*************';
//-----------------------------------------
// Gmail query 
// ref - syntax: https://support.google.com/mail/answer/7190?hl=en
let nDaysAgo   = ' after:'    + '';        // or use "nDaysAgo += getNDaysAgo_Date(nday);" to get output like "after:2022/5/3"
const category = ' category:' + 'promotions'; // primary,social,promotions,updates,forums,reservations



function main(){
  let top10data=[];
  
  nDaysAgo += getNDaysAgo_Date(60);
  const gmailQuery = category + nDaysAgo; // e.g. gmailQuery = subject + isunread + category;
  Logger.log('gmailQuery: ' + gmailQuery);

  top10data = getGmailTopSendersStatistics(gmailQuery);
  return top10data;
}


function getGmailTopSendersStatistics(gmailQuery_arg) {
  let sendersArr = [];
  let countsArr  = [];

  // Get all senders (threads.length / particular input numbers) by Gmail API 
  if (numberOfMail == 99999) {
    Logger.log('Searching all mails...');
    sendersArr = getGmailSenders_all(gmailQuery_arg);
  } else {
    Logger.log('Searching latest ' + numberOfMail + ' mails...');
    sendersArr = getGmailSenders_custNum(numberOfMail, gmailQuery_arg);
  }

  
  // Count the number of times each sender has sent a letter
  for (const sender of sendersArr) {
    countsArr[sender] = countsArr[sender] ? countsArr[sender] + 1 : 1;
  }


  // Map a sender and count to a json, and add them to array
  let senderCountArr = [];
  for (let i = 0; i < sendersArr.length; i++) {
    let senderCount={};
    senderCount.sender = sendersArr[i];
    senderCount.count  = Object.values(countsArr)[i];  //using Object.values(x) to get enum value
    senderCountArr.push(senderCount);
  }

  // Descending sort
  senderCountArr.sort(function(a, b) {
    return b.count - a.count;
  });

  // Get top10 senders and their counts
  let top10data=[];
  for (let i = 0; i < numberOfTopSenders; i++) {
    top10data.push(senderCountArr[i]);
  }

  // Backup data to Spreadsheet
  let spreadsheetId = createSpreadSheet();
  saveData(top10data, spreadsheetId);
  moveFileToParticularFolder(spreadsheetId, foldername);

  return top10data;
}


// Get all senders in all mails via Gmail API
//  throuth threads(1 thread get 1 sender) in batches of 500 each time (Gmail regulation)
function getGmailSenders_all(gmailQuery_arg){
  const startIndex = 0;
  const maxThreads = 500;
  let sendersArr = [];

  do {
    // const threads = GmailApp.getInboxThreads(startIndex, maxThreads);
    const myThreads = GmailApp.search(gmailQuery_arg, startIndex, maxThreads);


    for(let i=0; i < threads.length; i++) {
      const sender = threads[i].getMessages()[0].getFrom(); // Get first message
      sendersArr.push(sender); 

      //For debug: calculating time to prevent from timeout
      // if (sendersArr.length%50 == 0) {Logger.log("sendersArr.length:"+sendersArr.length)}; 
    }

    //For debug: calculating times to prevent from over quotas (10,000 or 20,000 per day)
    // if (sendersArr.length == 10) { break; }
    startIndex += maxThreads;
  } while (threads.length == maxThreads);

  return sendersArr;
}

// Get all senders in mails of custom numbers via Gmail API
function getGmailSenders_custNum(numberOfMail_arg, gmailQuery_arg) {
  const threads = GmailApp.search(gmailQuery_arg);
  // const threads = GmailApp.getInboxThreads();
  let sendersArr = [];
  if (numberOfMail_arg > threads.length) {
    numberOfMail_arg = threads.length;
  }

  for (let i = 0; i < numberOfMail_arg; i++) {
    let sender = threads[i].getMessages()[0].getFrom(); // Get first message
    sender = sender.replace(/"/g, '');                  //remove depulicate "
    sendersArr.push(sender);
  }
  
  return sendersArr;
}


function saveData(data, spreadsheetId) {
  // Logger.log("top10data:"+JSON.stringify(data));
    
  const dataSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Sheet1");
  for (let i = 0; i < 10; i++) {
    const range  = dataSheet.getRange(i+1,1);
    range.setValue(JSON.stringify(data[i]));
  }
}


function createSpreadSheet(){
  const dateTime = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmm");
  const spreadsheet = SpreadsheetApp.create('Testdata_'+dateTime,100,100);
  return spreadsheet.getId();
}


function moveFileToParticularFolder(spreadsheetId, foldername) {
  const folder   = DriveApp.getFoldersByName(foldername).next();//gets first folder with the given foldername
  const copyFile = DriveApp.getFileById(spreadsheetId);
  folder.addFile(copyFile);
  DriveApp.getRootFolder().removeFile(copyFile);
}


function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}


function getNumberOfMail() {
  return numberOfMail;
}


function send_line(lineNotifyTitle) {
  var payload = { 'message': lineNotifyTitle };
  var options = {
    "method": "post",
    "payload": payload,
    "headers": { "Authorization": "Bearer " + lineNotifyToken }
  };
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}


function getNDaysAgo_Date(nday) {

  const now_time          = new Date().getTime();  //getTime() returns milliseconds
  const nDaysAgo_DateTime = new Date( now_time - ( 1000 * 60 * 60 * 24 * nday ) ); // get the date n days ago
 
  const yyyy = nDaysAgo_DateTime.getFullYear();
  const mm   = nDaysAgo_DateTime.getMonth()+1;
  const dd   = nDaysAgo_DateTime.getDate();
  const nDaysAgo_Date =yyyy+'/'+ mm +'/'+ dd;

  // Logger.log('nDaysAgo_Date: '+nDaysAgo_Date);
  return nDaysAgo_Date;

}
