var source_folder_id = '';
var archive_folder_id = '';
var subject = '';
var sender_email = '';
function update() {
  var threads = GmailApp.getInboxThreads(0, 100);
  var msgs = GmailApp.getMessagesForThreads(threads);
  for (var i = 0 ; i < msgs.length; i++) {
    for (var j = 0; j < msgs[i].length; j++) {
      if(msgs[i][j].getSubject() == subject && msgs[i][j].isUnread() &&
       msgs[i][j].getFrom().includes(sender_email.toLowerCase())){
        Logger.log('New email with subject found.');
        Logger.log('Sender: '+msgs[i][j].getFrom());
        msgs[i][j].markRead();
        var email_body = msgs[i][j].getPlainBody().toString();
        Logger.log("Body: "+email_body)
        link = getFirstLink(email_body);
        if(link != null){
          Logger.log('Link found in the email.');
          var fileBlob = getFile(link);
          var archive_folder = DriveApp.getFolderById(archive_folder_id);
          replaceFile(source_folder_id, fileBlob);
          archive_folder.createFile(fileBlob);
        }         
      
      }
      
    }
  }
}

function getFile(fileURL) {
  var response = UrlFetchApp.fetch(fileURL);
  var fileBlob = response.getBlob();
  return fileBlob;
}

function getFirstLink(input){
  var expression = /(https?:\/\/(?:www\.|(?!www))[^\s\.]+\.[^\s]{2,}|www\.[^\s]+\.[^\s]{2,})/gi;
  var matches = input.match(expression);
  return matches[0];
}

function replaceFile(folderId,file_blob) {
  const files = DriveApp.getFolderById(folderId).getFiles();
  while (files.hasNext()) {
    const file = files.next();
    var file_name = file.getName();
    Logger.log(file_name.toString());
    if(file_name.toString().endsWith('.csv')){
      Logger.log('updating file');
      file.setContent(file_blob.getDataAsString());
    }
    // Delete File
    //Drive.Files.remove(file.getId())
  }
}