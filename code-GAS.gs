const email = ""; // Input sender email like reports@company.com
const subject = ""; //Input subject like Yahoo DSP Scheduled Report: Connect Ads/Yahoo DSP Report.
const fileName = ""; // Input Filename that should be moved
const downloadFolderId = ""; // Input Folder Id where new file should be added (In the begining old file should be also in this folder)
const moveFolderId = ""; // Input Folder Id where old file should be moved
const label = ""; // Not used currently  'reports'

function myFunction() {
  let file = getFileFromEmail(email, subject);
  manageFiles(moveFolderId, downloadFolderId, file);
}


function getFileFromEmail(email,subject){

  // Log the subject lines of the threads labeled with 'reports'
  // var labelObject = GmailApp.getUserLabelByName("reports");
  // var threads = labelObject.getThreads();

  // for (var i = 0; i < threads.length; i++) {
  //   Logger.log(threads[i].getFirstMessageSubject());
  //   }

  let threads = GmailApp.search(`from:"<${email}>", subject:"${subject}"`);

  for (let thread of threads) {
    let messageCount = thread.getMessageCount(); // Count the number of messages in that thread
    let messages = thread.getMessages(); // All messages in thread
    let lastMessage = messages[messageCount - 1]; // Get the last message in that thread. The first message is numbered 0.

    let attach = lastMessage.getAttachments()[0]; // Get first attachment
    let unzipblob = Utilities.unzip(attach)[0]; // Unzip file

    return unzipblob

    // for (let msg of messages){
    //   var attach = msg.getAttachments()[0];
    //   var unzipblob = Utilities.unzip(attach);
    //   Logger.log(unzipblob[0].getContentType());
    // }
  }
}

function manageFiles(moveFolderId, downloadFolderId, blob){

  let mainFolder = DriveApp.getFolderById(downloadFolderId)
  let moveToFolder = DriveApp.getFolderById(moveFolderId)
  let files = mainFolder.getFiles();

  let key = fileName;

  while (files.hasNext()){
    file = files.next();
    if(file.getName().includes(key)){
      moveToFolder.addFile(file);
      mainFolder.removeFile(file)
    }
  }
  mainFolder.createFile(blob);  // Download new file to folder
}

// Шаги
// 1. Проверить последнее пиьсмо с label 'reports' и темой "Yahoo DSP Scheduled Report: Connect Ads/Yahoo DSP Report."
// 2. Извлечь вложение
// 3. Разархивировтаь вложение
// 4. Переместить существующий файл 'Connect Ads Yahoo DSP.csv' в 'Archive' в папке 'YAHOO DSP CON'
// 5. Сохранить новый файл как Connect Ads Yahoo DSP.csv в папке YAHOO DSP CON
// -----
// Для Gmail
// Письма приходят с label 'reports', Yahoo DSP Scheduled Report: Connect Ads/Yahoo DSP Report.
// Письмо приходит каждый раз в 3:08 AM с почты 'reports@company.com'
// Вложение: zip файл 'Connect Ads Yahoo DSP Report дд-мм-гггг.zip'.
// -----
// Для Drive
// Вставить произвольный ID папки (будет заменен)
// Потом будут подставлены ID папок из shared drive и настроены goggle permissions.
