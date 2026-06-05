function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("HR Tools")
    .addItem("Generate Status", "status")
    .addItem("Highlight Attrition", "highlightAttrition")
    .addSeparator()
    .addItem("Generate HR Report", "Reports")
    .addToUi();
}
function status(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("HR_Data");
  var lastrow = sheet.getLastRow();
  sheet.getRange(2, 25, sheet.getLastRow()-1, 25).clear();
  for (var i = 2; i<=lastrow; i++){
    var attrition = sheet.getRange(i,21).getValue();
    if (attrition == "Yes") {
  sheet.getRange(i, 25).setValue("left company");
  sheet.getRange(i, 25).setBackground("#FF0000"); // 👈 Red
} else {
  sheet.getRange(i, 25).setValue("active");
  sheet.getRange(i, 25).setBackground("#00FF00"); // 👈 Green
}
  }
SpreadsheetApp.getUi().alert("Status of employees added ✅")
}
function highlightAttrition() {
  var sheet = SpreadsheetApp.getActive()
                            .getSheetByName("HR_Data");
  var lastRow = sheet.getLastRow();
  sheet.getRange(2, 1, sheet.getLastRow()-1, 23).setBackground(null);
  for (var i = 2; i <= lastRow; i++) {
    var attrition = sheet.getRange(i, 21).getValue(); // 👈 read cell directly
    
    if (attrition == "Yes") {
      sheet.getRange(i, 1, 1, 23).setBackground("#FF0000");
    } else {
      sheet.getRange(i, 1, 1, 23).setBackground("#FFFFFF");
    }
  }
function Reports() {
  
  var ss = SpreadsheetApp.getActive();
  
  // Step 2 - Check if Reports exists
  var reportSheet = ss.getSheetByName("Reports");
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  
  // Step 3 - Create fresh Reports sheet
  reportSheet = ss.insertSheet("Reports");
  
  // Step 4 - Get HR_Data for reading
  var dataSheet = ss.getSheetByName("HR_Data");
  // Step 5 - Add merged title
reportSheet.getRange("A1:B1").merge();
reportSheet.getRange("A1").setValue("HR Analytics Report");
reportSheet.getRange("A1:B1").setBackground("#1f3864");
reportSheet.getRange("A1:B1").setFontColor("white");
reportSheet.getRange("A1:B1").setFontWeight("bold");
reportSheet.getRange("A1:B1").setFontSize(14);
reportSheet.getRange("A1:B1").setHorizontalAlignment("center");
  // Step 6 - Add column headers
reportSheet.getRange("A3").setValue("Metric");
reportSheet.getRange("B3").setValue("Value");
reportSheet.getRange("A3:B3").setBackground("#4a86e8");
reportSheet.getRange("A3:B3").setFontColor("white");
reportSheet.getRange("A3:B3").setFontWeight("bold");
var data = dataSheet.getRange(2, 1, dataSheet.getLastRow()-1, 23).getValues();
  
  // Step 8 - Calculate metrics
var totalEmp = data.length;
var maleCount = 0;
var femaleCount = 0;
var attritionYes = 0;
var totalSalary = 0;
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][3] == "Male") maleCount++;
    if (data[i][3] == "Female") femaleCount++;
    if (data[i][20] == "Yes") attritionYes++;
    totalSalary += data[i][12];
  }
  
  var avgSalary = Math.round(totalSalary / totalEmp);
  // Step 9 - Write metrics to report
  reportSheet.getRange("A4").setValue("Total Employees");
  reportSheet.getRange("B4").setValue(totalEmp);
  
  reportSheet.getRange("A5").setValue("Male Employees");
  reportSheet.getRange("B5").setValue(maleCount);
  
  reportSheet.getRange("A6").setValue("Female Employees");
  reportSheet.getRange("B6").setValue(femaleCount);
  
  reportSheet.getRange("A7").setValue("Attrition Count");
  reportSheet.getRange("B7").setValue(attritionYes);
  
  reportSheet.getRange("A8").setValue("Attrition Rate");
  reportSheet.getRange("B8").setValue(Math.round(attritionYes/totalEmp*100) + "%");
  
  reportSheet.getRange("A9").setValue("Average Salary");
  reportSheet.getRange("B9").setValue("$" + avgSalary);
  // Step 10 - Color code the rows
  reportSheet.getRange("A4:B4").setBackground("#e8f0fe");
  reportSheet.getRange("A5:B5").setBackground("#c9daf8");
  reportSheet.getRange("A6:B6").setBackground("#c9daf8");
  reportSheet.getRange("A7:B7").setBackground("#fce8b2");
  reportSheet.getRange("A8:B8").setBackground("#fce8b2");
  reportSheet.getRange("A9:B9").setBackground("#d9ead3");
  
  // Step 11 - Auto resize columns
  reportSheet.autoResizeColumn(1);
  reportSheet.autoResizeColumn(2);
  
  // Step 12 - Success message
  SpreadsheetApp.getUi().alert("HR Report Generated! ✅");
  }
  
  SpreadsheetApp.getUi().alert("Attrition highlighted! ✅");
}
