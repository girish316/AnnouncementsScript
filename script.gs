function updateAnnouncements(){
  
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1i41ajDwKHeUeyK-B6L9NT8H9wGEq6DNkG5wMwaPVXMI/edit#gid=577109827");
  var sheet = spreadsheet.getSheetByName("Form Responses"); // sheet initialized and object created
  var data = sheet.getDataRange().getValues(); // 2d array that stores values from spreadsheet

  const today = new Date(); // date for today
  var yesterday = new Date(today); // stores yesterday's date - used for removing expired announcements from sheet and array
  yesterday.setDate(today.getDate() - 1);


  for (var i = 0; i < data.length; i++) { // loops through the sheet data, removes unnecessary or expired rows of data from sheet and array
    if(data[i][4] == "" || (new Date(data[i][2]) < yesterday || data[i][2] == "") && new Date(data[i][0]) < yesterday){
      sheet.deleteRow(i+1) // row deleted from sheet
      data.splice(i, 1) // index deleted from array
    }
  }
  
  data.sort(function(x,y){ // sorts data from speradsheet based on end date for ordered announcement placement
      var xp = x[2];
      var yp = y[2];
      return xp == yp ? 0 : xp < yp ? -1 : 1;
    });

  var doc = DocumentApp.openByUrl("https://docs.google.com/document/d/1Arsu7fCRVi2VvKYoNNs1Lchxm-eNC-_9ko519Ecph5E/edit") // announcements doc accessed
  var body = doc.getBody().clear().setFontFamily("Open Sans") // doc cleared and font set

  for(let i = 0; i < data.length; i++){ // loop to go through sheets data and update announcements doc

      if((new Date(data[i][2]) >= yesterday || data[i][2] == "") && new Date(data[i][0]) <= today){ // if announcement to be displayed

        var category = body.appendParagraph(data[i][5]); // title text as department section with style
        category.setHeading(DocumentApp.ParagraphHeading.HEADING3)

        var para = body.appendParagraph(data[i][4]); // body text with style
        para.setHeading(DocumentApp.ParagraphHeading.NORMAL)
        para.setLineSpacing(1.75)

        var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        var endMessage = "- Ends "+months[today.getMonth()]+" "+today.getDate()+", "+today.getFullYear() // variable that contains the displayed ending date of the announcement

        if(data[i][2] != ""){ // if statement to decide between today and end date as end message
          var date = new Date(data[i][2])
          endMessage = "- Ends "+months[date.getMonth()]+" "+date.getDate()+", "+date.getFullYear()
        }

        var end = body.appendParagraph(endMessage) // text to specify end date of announcement
        end.setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
        end.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }        
    }
}
