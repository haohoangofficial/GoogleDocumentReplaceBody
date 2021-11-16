/**
 * Author: Hao Hoang
 * Sheet url template: https://docs.google.com/spreadsheets/d/1KAiaqei2rVLnE98I9NVyqqNTmNXG9FRnJ3XJzLmrlas/edit?usp=sharing
 * Library required bellow
 * AlaSQL library ID: 1XWR3NzQW6fINaIaROhzsxXqRREfKXAdbKoATNbpygoune43oCmez1N8U
 * How to use this library: https://github.com/contributorpw/alasqlgs
 */
const alaSQL = AlaSQLGS.load()
const sheet_id = '1KAiaqei2rVLnE98I9NVyqqNTmNXG9FRnJ3XJzLmrlas' // replace your sheet ID
const sheet_name = 'Doc' // Replace your Sheet Name
const doc_id = '1Htp-S5yg_-mHaclEHm3QAOQ4DOkiSE9HUhFOnkx42OM' // Replace your google document id
const sheet = SpreadsheetApp.openById(sheet_id).getSheetByName(sheet_name)
function getSheetData() {
  let parameters = sheet.getDataRange().getValues()
  let i = 1 // Number of files you want to replicate in 1 request
  let sql = `SELECT MATRIX FROM ? WHERE [2] = "" LIMIT ${i}`
  let response = alaSQL(sql,[parameters])
  return response
}

function addCronTask() {
  let response = getSheetData()
  if(response[0]) {
    for(let resp of response) {
      var doc = DriveApp.getFileById(doc_id)
      var newDoc = Drive.newFile();
      var blob =doc.getBlob();
      var file=Drive.Files.insert(newDoc,blob,{convert:true});
      let newfile = DocumentApp.openById(file.id)
      let body = newfile.getBody()
      body.replaceText('{{name}}', resp[3]);
      newfile.setName(`${doc.getName()} ${resp[3]}`);
      sheet.getRange(`B${resp[0] + 1}`).setValue(file.id)
      sheet.getRange(`C${resp[0] + 1}`).setValue('UPDATED')
    }
    Logger.log(`https://docs.google.com/document/d/${file.id}`)
    return `https://docs.google.com/document/d/${file.id}`
  }
}