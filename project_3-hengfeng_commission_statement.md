# Hengfeng Commission Statement

In this project, I wrote scripts to:
1. Automate the RESET of all sheets in the entire spreadsheet
2. Automate the styling of a working spreadsheet

These scripts have been applied to my google sheets project: https://docs.google.com/spreadsheets/d/1xrfwpa9PzHIiJPijKymj7be1uUmHVTzn2r-OgltaOmM/edit. Please contact me for more information.

## 1. Automate the RESET of a particular spreadsheet
```javascript
// get all sheet names
function getSheetNames() {
  const sheetNames = [];
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for(let i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  return sheetNames;
}

// reset function
function reset() {
  // get all sheet names using getSheetNames()
  const sheetNames = getSheetNames();

  // create a customized pop-up window before applying the reset function
  const ui = SpreadsheetApp.getUi();
  const resetButton = ui.alert('重设','您确定要重设此表格吗？表格内所有数据将被删除，但不影响其他表格',ui.ButtonSet.YES_NO);

  // if the user clicks YES, iterate through the sheet name list
  if(resetButton == ui.Button.YES){
    for(let i=0; i<sheetNames.length; i++) {
      const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNames[i]);

      // apply on all sheets except "序列"
      if(sheetNames[i] != '序列') {
        // variables
        const headNameRange = "C3";
        const perAreaCurrent1 = "C6:V20";
        const perAreaCurrent2 = "Z6:AH20";
        const perAreaCurrent3 = "AL6:AL20";
        const perAreaPrevious1 = "B23:V";
        const perAreaPrevious2 = "Z23:AH";
        const perAreaPrevious3 = "AL23:AL";
        const personalActualTotalRow6 = "X6";
        const personalActualTotalRange1 = "X6:X20";
        const personalActualTotalRow23 = "X23";
        const personalActualTotalRange2 = "X23:X";
        const xTotalRow6 = "AJ6";
        const xTotalRange1 = "AJ6:AJ20";
        const xTotalRow23 = "AJ23";
        const xTotalRange2 = "AJ23:AJ";


        ss.getRange(headNameRange).clearContent();
        ss.getRange(perAreaCurrent1).clearContent();
        ss.getRange(perAreaCurrent2).clearContent();
        ss.getRange(perAreaCurrent3).clearContent();
        
        ss.getRange(perAreaPrevious1).clearContent();
        ss.getRange(perAreaPrevious2).clearContent();
        ss.getRange(perAreaPrevious3).clearContent();

        // set a formula in 'X6' to copy down
        ss.getRange(personalActualTotalRow6).setFormula('=if($C6="","",IF(COUNTIFS($C6,"*闲*置*")>0,0,$W6+sumifs($X$23:$X,$C$23:$C,$C6)))');
        ss.getRange(personalActualTotalRow6).copyTo(ss.getRange(personalActualTotalRange1));

        // set a formula in 'X23' to copy down
        ss.getRange(personalActualTotalRow23).setFormula('=if($C23="","",IF(COUNTIFS($C23,"*闲*置*")>0,0,$W23))');
        ss.getRange(personalActualTotalRow23).copyTo(ss.getRange(personalActualTotalRange2));

        // set a formula in 'AJ6' to copy down
        ss.getRange(xTotalRow6).setFormula('if($C6="","",IF(COUNTIFS($C6,"*闲*置*")>0,0,$AI6+sumifs($AJ$23:$AJ,$C$23:$C,$C6)))');
        ss.getRange(xTotalRow6).copyTo(ss.getRange(xTotalRange1));

        // set a formula in 'AJ23' to copy down
        ss.getRange(xTotalRow23).setFormula('if($C23="","",IF(COUNTIFS($C23,"*闲*置*")>0,0,$AI23))');
        ss.getRange(xTotalRow23).copyTo(ss.getRange(xTotalRange2));
      }      
    }    
  }
}
```

```javascript
// helper function to iterate through all cells in a range
function forEachRangeCell(range, f) {
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  for (let c=1; c<=numCols; c++) {
    for (let r=1; r<=numRows; r++) {
      const cell = range.getCell(r, c);
      f(cell)
    }
  }
}

// function to run on edit
function onEdit() {
  // get active range and column
  const ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = ss.getActiveRange();
  const activeColNum = ss.getActiveRange().getColumn();
  
  // Columns Variables
  const nameColNum = 3;
  const xingGuangColNum = 4;
  const xingGuangPerColNum = 5;
  const kaiXuanColNum = 6;
  const kaiXuangPerColNum = 7;
  const huaYangColNum = 8;
  const huaYangPerColNum = 9;
  const jinZhuColNum = 10;
  const jinZhuPerColNum = 11;
  const huaYang2ColNum = 12;
  const huaYang2PerColNum = 13;
  const huaYang3ColNum = 14;
  const huaYang3PerColNum = 15;
  const jinLiColNum = 16;
  const jinLiPerColNum = 17;
  const jinLingColNum = 18;
  const jinLingPerColNum = 19;
  const zhaoHuiColNum = 20;
  const zhaoHui1PerColNum = 21;
  const zhaoHui3PerColNum = 22;
  const personalTotalColNum = 23;
  const personalActualTotalColNum = 24;
  const cardRemarksColNum = 26;
  const xTianbaoColNum = 27;
  const xTianbaoPerColNum = 28;
  const xHeshengColNum = 29;
  const xHeshengPerColNum = 30;
  const kuaiBoColNum = 31;
  const kuaiBoPerColNum = 32;
  const kuaiBo91MengMeiColNum = 33;
  const kuaiBo91MengMeiPerColNum = 34;
  const xTotalColNum = 35;
  const xTotalPerColNum = 36;
  const xRemarksColNum = 38;

  // Rows Variables
  const firstTopWorkingRowNum = 6;
  const firstBottomWorkingRowNum = 20;
  const secondTopWorkingRowNum = 23;

  if(ss.getName() !== '序列' &&
      ((activeColNum>=nameColNum && activeColNum <= zhaoHui3PerColNum) ||
      (activeColNum == personalActualTotalColNum) ||
      (activeColNum>=cardRemarksColNum && activeColNum<=kuaiBo91MengMeiPerColNum) ||
      (activeColNum == xTotalPerColNum) ||
      (activeColNum == xRemarksColNum))){
    const activeRowNum = ss.getActiveRange().getRow();

    if((activeRowNum>=firstTopWorkingRowNum && activeRowNum <=firstBottomWorkingRowNum) || activeRowNum>=secondTopWorkingRowNum){
      
      forEachRangeCell(activeRange, (cell) => {
        
        if(cell.getColumn() == nameColNum){
          // change from "*闲*置*" to "闲置", since there may be spaces within the characters
          if(cell.getValue().includes("闲") && cell.getValue().includes("置")){
            cell.setValue("闲置");
          }

          const sourceRange = ss.getRange(firstTopWorkingRowNum, cell.getColumn());
          sourceRange.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        }

        if(cell.getColumn() == xingGuangColNum || 
          cell.getColumn() == kaiXuanColNum || 
          cell.getColumn() == huaYangColNum || 
          cell.getColumn() == jinZhuColNum ||
          cell.getColumn() == huaYang2ColNum || 
          cell.getColumn() == huaYang3ColNum || 
          cell.getColumn() == jinLiColNum || 
          cell.getColumn() == jinLingColNum || 
          cell.getColumn() == zhaoHuiColNum || 
          cell.getColumn() == xTianbaoColNum || 
          cell.getColumn() == xHeshengColNum ||
          cell.getColumn() == kuaiBoColNum ||
          cell.getColumn() == kuaiBo91MengMeiColNum){

          cell.setNumberFormat('@');
          cell.setFontWeight('bold');
          cell.setFontSize('11');
          cell.setFontFamily('Arial');
          cell.setFontStyle('normal');
          cell.setFontColor('#0000ff');
          cell.setHorizontalAlignment('center');
          cell.setBackground(null);
          cell.setBorder(false, false, false, false, false, false);
        }
        if(cell.getColumn() == xingGuangPerColNum || 
          cell.getColumn() == kaiXuangPerColNum || 
          cell.getColumn() == huaYangPerColNum || 
          cell.getColumn() == jinZhuPerColNum || 
          cell.getColumn() == huaYang2PerColNum || 
          cell.getColumn() == huaYang3PerColNum || 
          cell.getColumn() == jinLiPerColNum || 
          cell.getColumn() == jinLingPerColNum || 
          cell.getColumn() == zhaoHui1PerColNum || 
          cell.getColumn() == zhaoHui3PerColNum || 
          cell.getColumn() == personalActualTotalColNum || 
          cell.getColumn() == xTianbaoPerColNum || 
          cell.getColumn() == xHeshengPerColNum || 
          cell.getColumn() == kuaiBoPerColNum ||
          cell.getColumn() == kuaiBo91MengMeiPerColNum ||
          cell.getColumn() == xTotalPerColNum) {

          cell.setNumberFormat('#,##0.00');
          cell.setFontWeight('normal');
          cell.setFontSize('11');
          cell.setFontFamily('Arial');
          cell.setFontStyle('normal');
          cell.setFontColor('#000000');
          cell.setHorizontalAlignment('center');
          cell.setBackground(null);
          cell.setBorder(false, false, false, false, false, false);
        }

        // set the format and the formulae for column X column AJ
        if(cell.getColumn() == personalActualTotalColNum) {
          const sourceFormatCell = ss.getRange(firstTopWorkingRowNum, cell.getColumn());
          sourceFormatCell.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

          if(cell.isBlank()){
            const nameCell = ss.getRange(cell.getRow(), nameColNum).getA1Notation();
            const sourceCell = ss.getRange(cell.getRow(), personalTotalColNum).getA1Notation();

            if(cell.getRow() >= firstTopWorkingRowNum && cell.getRow() <= firstBottomWorkingRowNum) {        
            cell.setFormula(`=if($${nameCell}="","",IF(COUNTIFS($${nameCell},"*闲*置*")>0,0,$${sourceCell}+sumifs($X$23:$X,$C$23:$C,$${nameCell})))`);
            }
            else if(cell.getRow() >= secondTopWorkingRowNum) {
              cell.setFormula(`=if($${nameCell}="","",IF(COUNTIFS($${nameCell},"*闲*置*")>0,0,$${sourceCell}))`);
            }
          }
        }
        if(cell.getColumn() == xTotalPerColNum) {
          const sourceFormatCell = ss.getRange(firstTopWorkingRowNum, cell.getColumn());
          sourceFormatCell.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

          if(cell.isBlank()){
            const nameCell = ss.getRange(cell.getRow(), nameColNum).getA1Notation();
            const sourceCell = ss.getRange(cell.getRow(), xTotalColNum).getA1Notation();
            if(cell.getRow() >= firstTopWorkingRowNum && cell.getRow() <= firstBottomWorkingRowNum) {        
              cell.setFormula(`=if($${nameCell}="","",IF(COUNTIFS($${nameCell},"*闲*置*")>0,0,$${sourceCell}+sumifs($AJ$23:$AJ,$C$23:$C,$${nameCell})))`);
            }
            else if(cell.getRow() >= secondTopWorkingRowNum) {
              cell.setFormula(`=if($${nameCell}="","",IF(COUNTIFS($${nameCell},"*闲*置*")>0,0,$${sourceCell}))`);
            }
          }          
        }

        // set format and conditional format rules
        if(cell.getColumn() == cardRemarksColNum || cell.getColumn() == xRemarksColNum){
         
          // copy down the format and conditional format rules
          const sourceRange = ss.getRange(firstTopWorkingRowNum, cell.getColumn());
          sourceRange.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
         
        }
      })
    }
  }
}
```