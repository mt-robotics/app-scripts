# TS 2023-2024 Payroll App Scripts

Google sheets applying these codes: https://docs.google.com/spreadsheets/d/1w8J6nzhV67U-R9h5xtBuWWIKQrmg38XWf6zZe_rQzRk/edit

(Make a copy with the google sheets above, then in the menu bar, click on `Extensions` > `App Scripts`, the scripts will show up)

```javascript
function copySheetWithProtections_(sheet, optNewSheetName, optSheetIndex, optTargetSpreadsheet) {
  // version 1.2, mainly written by --Hyde, 16 February 2022
  //  - add optTargetSpreadsheet
  //  - see https://support.google.com/docs/thread/147940588
  // version 1.1, written by --Hyde, 8 February 2022
  //  - see https://support.google.com/docs/thread/149743347?msgid=149890712
  //  - see https://support.google.com/docs/thread/144274437
  //  - see https://support.google.com/docs/thread/126744993
 
  const ss = sheet.getParent();
  const newSheetName = optNewSheetName || sheet.getName();
  const targetSs = optTargetSpreadsheet || ss;
  const sheetIndex = optSheetIndex ?? targetSs.getNumSheets() + 1;
  const me = Session.getEffectiveUser();
  if (!me.getEmail()) {
    throw new Error('Cannot retrieve identity of the effective user.');
  }
  const newSheet = sheet.copyTo(targetSs);
  newSheet.activate();
  targetSs.moveActiveSheet(sheetIndex);
  try {
    newSheet.setName(newSheetName);
  } catch (error) {
    ; // the sheet name is already taken, retain the default 'Copy of...' name
  }
  const sheetProt = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (sheetProt) {
    const newSheetProt = newSheet
      .protect()
      .setDescription(sheetProt.getDescription())
      .setWarningOnly(sheetProt.isWarningOnly());
    if (!sheetProt.isWarningOnly()) {
      newSheetProt
        .addEditor(me)
        .removeEditors(newSheetProt.getEditors().filter(user => user.getEmail() !== me))
        .addEditors(sheetProt.getEditors());
      try {
        newSheetProt.setDomainEdit(sheetProt.canDomainEdit());
      } catch (error) {
        ; // we are not in a Google Workspace domain
      }
    }
    const unprotectedRanges = sheetProt.getUnprotectedRanges()
      .map(range => newSheet.getRange(range.getA1Notation()));
    newSheetProt.setUnprotectedRanges(unprotectedRanges);
  } else {
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .forEach(rangeProt => {
        const rangeA1 = rangeProt.getRange().getA1Notation();
        const newRangeProt = newSheet.getRange(rangeA1)
          .protect()
          .setDescription(rangeProt.getDescription())
          .setWarningOnly(rangeProt.isWarningOnly());
        if (!rangeProt.isWarningOnly()) {
          newRangeProt
            .addEditor(me)
            .removeEditors(newRangeProt.getEditors().filter(user => user.getEmail() !== me))
            .addEditors(rangeProt.getEditors());
          try {
            newRangeProt.setDomainEdit(rangeProt.canDomainEdit());
          } catch (error) {
            ; // we are not in a Google Workspace domain
          }
        }
      });
  }
  return newSheet;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // create a menu in the spreadsheet
  ui.createMenu('复制模版带保护范围').addItem('复制【模版】工作簿带保护范围，避免误修改导致公式出错', 'copyTemplateWithProtection').addToUi();

}

function copyTemplateWithProtection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNameToCopy = "模版";
  const sheetToCopy = ss.getSheetByName(sheetNameToCopy);

  copySheetWithProtections_(sheetToCopy);
}
```