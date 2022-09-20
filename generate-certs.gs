//Goal: Generate unique certs for each participant based on unique study done
//Status: works June 6 2022
//runtime: 12s pp (13 calls) > 9s pp > 8s > 9.5s > 8.5 > 7.5 > 9s > 8.5s > 9.5s > 7s
//How to use: only do ~35 rows at a time; empty drive recycle bin after using

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
      if (QUERY == "Alexa Embodiment") {
        var fileId = "17G3Jqjbbw9Q8FRlVHM9CQunNCBq6xqxShMHUTRXe8sg"; //alexa_embodiment deck
      } else if (`${QUERY}`.toLowerCase().includes("ikc")) {   //(`${INPUT}`.toUpperCase()
        var fileId = "1etZcYGY-gS0WwR0FeI11nNBKHAV-SXEG39F452IZObI"; //IKC slide deck..etc
      } else if (`${QUERY}`.toLowerCase().includes("convo")) {
        var fileId = "1ZFdfGokgvbUsF6REmLsh1_CsOyGpRtnLoE43IbGC8yU"; 
      } else if (`${QUERY}`.toLowerCase().includes("storybook")) {
        var fileId = "1YQM7puktsjRhIKyGQzIzJFDMzazdDKn0yIjk13kypOo"; 
      } else if (`${QUERY}`.toLowerCase().includes("strconseq")) {
        var fileId = "19c7pHYhmH9eE1JuIejoA489O8B5bXAmebx_1B645i1I";
      } else if (`${QUERY}`.toLowerCase().includes("racelang")) {
        var fileId = "1FH4v2HB-G7XKgsb_u0DI9RcyYt4-pH3EjXGrX1IuhoA"; //change file id
      } else {
        var fileId = "1FG3fw8dA_Sja0YtqHphDQEisNpzQnoAvfu3c7RCfzEk";   ///fix this: create generic certificate
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
