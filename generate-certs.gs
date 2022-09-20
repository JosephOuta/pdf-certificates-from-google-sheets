//Goal: Generate unique certs for each participant based on unique study done
//Status: works Sep 2022
//runtime: 7s pp
//How to use: maybe only do ~35 rows at a time; empty drive recycle bin after using

function generate_attach_certs() {
foldername = "test_script";
var parentFolder = DriveApp.getFolderById(DriveApp.getRootFolder().getId());
parentFolder.createFolder(foldername);
const destinationFolder = DriveApp.getFoldersByName(foldername).next();
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[5];
const lastRow = sheet.getLastRow();
var input_list_rev = sheet.getRange(2, 3, lastRow - 1).getValues().reverse(); // to counteract effect of while-loop below
const study_list_rev = sheet.getRange(2, 8, lastRow - 1).getValues().reverse(); //

if (input_list_rev.length != study_list_rev.length) {
console.log("Error: child & study names don't match up");
} else {
  for (let i = 0; i < input_list_rev.length; i++) {
    if (input_list_rev[i] != '') { 
      var QUERY = study_list_rev[i];
      if (QUERY == "studyname1") {
        var fileId = "studyname1_slide-id"; //alexa_embodiment deck
      } else if (`${QUERY}`.toLowerCase().includes("studyname2")) {   //(`${INPUT}`.toUpperCase()
        var fileId = "studyname2_slide-id"; //slide deck..etc
      } else if (`${QUERY}`.toLowerCase().includes("studyname3")) {
        var fileId = "studyname3_slide-id"; 
      } else if (`${QUERY}`.toLowerCase().includes("studyname4")) {
        var fileId = "studyname4_slide-id"; 
      } else if (`${QUERY}`.toLowerCase().includes("studyname5")) {
        var fileId = "studyname5_slide-id";
      } else if (`${QUERY}`.toLowerCase().includes("studyname6")) {
        var fileId = "studyname6_slide-id"; //change file id
      } else {
        var fileId = "studyname-neutral_slide-id";   ///fix this: create generic certificate
      };     
      var INPUT = `${input_list_rev[i]}`.replace(/\s+/g, "");
      var template = DriveApp.getFileById(fileId); //grab template deck
      var fileName = template.getName();
      var template_copy = template.makeCopy(); // duplicate entire deck using DriveApp
      template_copy.setName("temp copy of " + fileName);
      var getcopy = SlidesApp.openById(template_copy.getId()); //open duplicate deck
      var template_slide = getcopy.getSlides()[0]; 
      var shapes = (template_slide.getShapes()); //grab and replace text
      shapes.forEach(function(shape){
        shape.getText().replaceAllText('Childname',INPUT);
      });
      getcopy.saveAndClose();
      destinationFolder.createFile(template_copy.getBlob().setName(`${INPUT}`.toUpperCase() + "_Certificate.pdf"));
      template_copy.setTrashed(true);
      Logger.log("created " + INPUT + " certificate 🎉");
    } 
  }
}
Logger.log("Attaching certificates...");
var input_list = input_list_rev.reverse();
var files = destinationFolder.getFiles(); 
var certificates = [];
while (files.hasNext()){
  var file = files.next();
  certificates.push(file.getId()); //convert getFiles output into list to avoid while loop
};
var certificate_index = [];
for (i = 0; i < input_list.length; i++) {
  if (input_list[i] != "") {
    var output_index = [];
    output_index.push(i);
    certificate_index.push(i);
    var certificate_url = "https://drive.google.com/file/d/" + certificates[certificate_index.indexOf(i)] + "/view?usp=sharing"; 
    sheet.getRange(`E${output_index[0]+2}`).setValue(certificate_url);
    }
}
destinationFolder.setTrashed(true);
}
