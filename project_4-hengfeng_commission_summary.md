# Hengfeng Commission Summary

In this project, I wrote scripts to automate the RESET of the entire workbook.

This scripts have been applied to my google sheets project: https://docs.google.com/spreadsheets/d/19EMZrtWN6tuxd4o8rALYRjWscQSnvSW4H3nN6LoDyi8/edit. Please contact me for more information.

```javascript
function reset() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('数据源');
  const teamLeaderSheet = ss.getSheetByName('组长提成');
  const seriesSheet = ss.getSheetByName('序列');
  const aliasSheet = ss.getSheetByName('花名-核对');

  const ui = SpreadsheetApp.getUi();
  const buttonPress = ui.alert('Reset','Are you sure to reset the whole sheets?', ui.ButtonSet.YES_NO_CANCEL);

  if (buttonPress == ui.Button.YES){
    // delete data in 序列
    seriesSheet.getRange('B2').clearContent();

    // delete data in 花名-核对
    aliasSheet.getRange('B2:B').clearContent();
    aliasSheet.getRange('I1').clearContent();

    // delete data in 数据源
    sourceSheet.getRange('B5:V').clearContent();
    sourceSheet.getRange('X5:X').clearContent();
    sourceSheet.getRange('Z5:AH').clearContent();
    sourceSheet.getRange('AJ5:AJ').clearContent();
    sourceSheet.getRange('AL5:AL').clearContent();

    // delete data in 组长提成
    teamLeaderSheet.getRange('J5:J').clearContent();
    // reset formula
    const rangeToFillDown3 = teamLeaderSheet.getRange('I5:I');

    teamLeaderSheet.getRange('I5').setFormula('=IF($B5="","",$F5+$H5)');
    teamLeaderSheet.getRange('I5').copyTo(rangeToFillDown3);
  }
}
```