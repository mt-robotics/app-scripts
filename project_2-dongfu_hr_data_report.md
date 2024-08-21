# Dong Fu HR Data Report

In this project, I wrote scripts to automate the weekly data report of DongFu HR:
![alt text](images/dongfu_hr_weekly_report_sample.png)

These scripts have been applied to my google sheets project: https://docs.google.com/spreadsheets/d/1bxYt4pCqcvtozSJmmdMoEN03c9NL37n3oCdc8ublsOk/edit. Please contact me for more information.

## constants.gs
```javascript
const k_ss = SpreadsheetApp.getActiveSpreadsheet()
const k_importRangeSheet = k_ss.getSheetByName("IMPORTRANGE")
const k_infoSheet = k_ss.getSheetByName("员工信息表")
const k_onboardSheet = k_ss.getSheetByName("入职")
const k_resignSheet = k_ss.getSheetByName("离职")
const k_transferSheet = k_ss.getSheetByName("调岗调职")
const k_unPaidLeaveSheet = k_ss.getSheetByName("停薪留职")

// IMPORTRANGE
const k_imptrngShStartDate = k_importRangeSheet.getRange('B3');
const k_imptrngShEndDate = k_importRangeSheet.getRange('B4');

// 员工信息表
const k_infoSheetCurrentInfoRng = k_infoSheet.getRange('B4:S');
const k_infoSheetDeptRng = k_infoSheet.getRange('H4:H');
const k_infoSheetSubDeptRng = k_infoSheet.getRange('I4:I');
const k_infoSheetIdRng = k_infoSheet.getRange('B4:B');
const k_infoSheetDeptAndSubDeptRng = k_infoSheet.getRange('H4:I');
const k_infoSheetDeptColNum = 8;
const k_infoSheetSubDeptColNum = 9;
const k_InfoSheetIdColIndexFromB = 0;
// const k_InfoSheetDeptColIndexFromB = 6;
// const k_InfoSheetSubDeptColIndexFromB = 7;
const k_infoSheetStartingRow = 4;
const k_infoSheetStartingCol = 2;

// 入职
const k_onboardSheetCurrentNameRng = k_onboardSheet.getRange('C4:C');
const k_onboardSheetCurrentDataColNumFromB = 18;

// 离职
const k_resignSheetDataRng = k_resignSheet.getRange('B4:M');
const k_resignSheetNameRng = k_resignSheet.getRange('C4:C');

const k_resignSheetIdIndexFromB = 0;
const k_resignSheetNameIndexFromB = 1;
const k_resignSheetRankIndexFromB = 2;
const k_resignSheetSubRankIndexFromB = 3;
const k_resignSheetDeptIndexFromB = 4;
const k_resignSheetSubDeptIndexFromB = 5;
const k_resignSheetGroupIndexFromB = 6;
const k_resignSheetPositionIndexFromB = 7;
const k_resignSheetOnboardIndexFromB = 8;
const k_resignSheetPpOnboardIndexFromB = 9;
const k_resignSheetChannelIndexFromB = 11;

// others
const k_currentDeptDataRng = 'B4:I';  // except resignSheet
const k_nameRngIndexFromB = 1;
const k_currentDeptRngIndexFromB = 6;
const k_currentSubDeptRngIndexFromB = 7;
const k_currentDeptRng = 'H4:H';  // except resignSheet
const k_currentSubDeptRng = 'I4:I';  // except resignSheet
const k_resignSheetDeptDataRng = 'B4:G';
const k_resignSheetDeptRng = 'F4:F';
const k_resignSheetSubDeptRng = 'G4:G';
const k_resignSheetDeptRngIndexFromB = 4;
const k_resignSheetSubDeptRngIndexFromB = 5;

// initialize the class MyComponents
const myComponents = new MyComponents();
```

## components.gs
```javascript
class MyComponents {
  // 1- replace qb线事业部 with 部门's names in all sheets except 离职表
  replaceDeptNames() {
    const listOfWs = [k_infoSheet, k_onboardSheet, k_unPaidLeaveSheet];
    listOfWs.forEach(async ws => {
      let deptDataRng = ws.getRange(k_currentDeptDataRng).getValues();
    
      // 部门
      const newSubDeptDataRng = deptDataRng.map(function(row) {
        return new Array(row[k_nameRngIndexFromB].includes('阿炮') ? '顺丰公司' : row[k_currentSubDeptRngIndexFromB]);  // return a two-dim array in order to use setValues() for the subdept column
      });
      // 更新部门
      await ws.getRange(k_currentSubDeptRng).setValues(newSubDeptDataRng);

      deptDataRng = ws.getRange(k_currentDeptDataRng).getValues();  // run it again to get the latest updated 部门, or the old 部门 values will be return
      // 事业部
      const newDeptDataRng = deptDataRng.map(function(row) {
        return new Array(row[k_currentDeptRngIndexFromB].includes('qp线事业部') ? row[k_currentSubDeptRngIndexFromB] : row[k_currentDeptRngIndexFromB]);
      });
      // 更新事业部
      ws.getRange(k_currentDeptRng).setValues(newDeptDataRng);
    });
  }

  // 2- replace qb线事业部 with 部门's names in 离职表
  async replaceResignSheetDeptNames() {
    const resignWs = k_resignSheet;
    let deptDataRng = resignWs.getRange(k_resignSheetDeptDataRng).getValues();
    
    // 部门
    const newSubDeptDataRng = deptDataRng.map(function(row) {
      return new Array(row[k_nameRngIndexFromB].includes('阿炮') ? '顺丰公司' : row[k_resignSheetSubDeptRngIndexFromB]);  // return a two-dim array in order to use setValues() for the subdept column
    });
    // 更新部门
    await resignWs.getRange(k_resignSheetSubDeptRng).setValues(newSubDeptDataRng);

    deptDataRng = resignWs.getRange(k_resignSheetDeptDataRng).getValues();  // run it again to get the latest updated 部门, or the old 部门 values will be return
    // 事业部
    const newDeptDataRng = deptDataRng.map(function(row) {
      return new Array(row[k_resignSheetDeptRngIndexFromB].includes('qp线事业部') ? row[k_resignSheetSubDeptRngIndexFromB] : row[k_resignSheetDeptRngIndexFromB]);
    });
    // 更新事业部
    resignWs.getRange(k_resignSheetDeptRng).setValues(newDeptDataRng);
  }

  // 3- clear all notes
  clearAllNotes() {
    const listOfWs = [k_infoSheet, k_onboardSheet, k_resignSheet, k_unPaidLeaveSheet];
    listOfWs.forEach(ws => {
      ws.getRange(1, 1, ws.getLastRow(), ws.getLastColumn()).clearNote();
    });
  }

  // 4- fill Short Term Staff In the Onboard Sheet
  fillShortTermStaffData() {
    const onbordingStaffList = k_onboardSheetCurrentNameRng.getValues().map(function(row) {
      return row[0] !== '' ? row[0] : '';
    }).filter(r => r !== '');

    const resignStaffDataList = k_resignSheetDataRng.getValues().filter(r => r[0] !== '');
    const shortTermStaffDataList = [];

    resignStaffDataList.forEach(data => {
      const id = data[k_resignSheetIdIndexFromB];
      const name = data[k_resignSheetNameIndexFromB];
      const rank = data[k_resignSheetRankIndexFromB];
      const subRank = data[k_resignSheetSubRankIndexFromB];
      const dept = data[k_resignSheetDeptIndexFromB];
      const subDept = data[k_resignSheetSubDeptIndexFromB];
      const group = data[k_resignSheetGroupIndexFromB];
      const position = data[k_resignSheetPositionIndexFromB];
      const onboardDate = data[k_resignSheetOnboardIndexFromB];
      const ppOnboardDate = data[k_resignSheetPpOnboardIndexFromB];
      const channel = data[k_resignSheetChannelIndexFromB];

      const isInclude = ((onboardDate >= k_imptrngShStartDate.getValue() && onboardDate <= k_imptrngShEndDate.getValue()) || (ppOnboardDate >= k_imptrngShStartDate.getValue() && ppOnboardDate <= k_imptrngShEndDate.getValue())) && !onbordingStaffList.includes(name);

      return isInclude ? shortTermStaffDataList.push(new Array(id, name, '离职', '', rank, subRank, dept, subDept, group, position, '', onboardDate, ppOnboardDate, '', '试用期', channel, '', '')) : '';
    });

    if(shortTermStaffDataList.length == 0) {
      return
    } else {
      // range to start filling
      const startingRow = k_onboardSheetCurrentNameRng.getValues().filter(String).length + 4;
      const areaToFill = k_onboardSheet.getRange(startingRow, 2, shortTermStaffDataList.length, k_onboardSheetCurrentDataColNumFromB);

      areaToFill.setValues(shortTermStaffDataList);
    }
  }
}
```

## main.gs
```javascript
function createReport() {
  const ui = SpreadsheetApp.getUi();
  const buttonPress = ui.alert("创建报告","您确定要创建报告吗？期初单元格将会被替换",ui.ButtonSet.YES_NO_CANCEL);

  if(buttonPress == ui.Button.YES){

    // 复制员工信息
    const previousInfo = k_infoSheet.getRange("B4:S").getDisplayValues();
    const currentInfo = k_importRangeSheet.getRange("E4:V").getDisplayValues();
    // infoSheet.getRange("V4:AM").clearContent();
    // infoSheet.getRange("B4:S").clearContent();
    k_infoSheet.getRange("V4:AM").setValues(previousInfo);
    k_infoSheet.getRange("B4:S").setValues(currentInfo);

    // 复制入职人员信息
    const previousInfoRange = k_onboardSheet.getRange("C4:C");
    const previousNames = previousInfoRange.getDisplayValues();
    const currentOnboardInfoRange = k_importRangeSheet.getRange("AS4:BJ");
    const currentOnboardInfo = currentOnboardInfoRange.getDisplayValues();
    // onboarSheet.getRange("U4:U").clearContent();
    // onboarSheet.getRange("B4:S").clearContent();
    if(previousInfoRange.getValues().length > 0){
      k_onboardSheet.getRange("U4:U").setValues(previousNames);
    }
    else{
      k_onboardSheet.getRange("U4:U").clearContent();
    }

    if(currentOnboardInfoRange.getValues().length > 0){
      k_onboardSheet.getRange("B4:S").setValues(currentOnboardInfo);
    }
    else{
      k_onboardSheet.getRange("B4:S").clearContent();
    }

    // 复制离职人员信息
    const currentResignedInfoRange = k_importRangeSheet.getRange("BM4:CA");
    const currentResignedInfo = currentResignedInfoRange.getDisplayValues();
    if(currentResignedInfoRange.getValues().length > 0){
      k_resignSheet.getRange("B4:P").setValues(currentResignedInfo);
    }
    else{
      k_resignSheet.getRange("B4:P").clearContent();
    }
    k_resignSheet.getRange("Q4:Q").clearContent();

    // 复制调岗调职人员信息
    const currentTransferInfoRange = k_importRangeSheet.getRange("CD4:CM");
    const currentTransferInfo =  currentTransferInfoRange.getDisplayValues();
    if(currentTransferInfoRange.getValues().length > 0){
      k_transferSheet.getRange("B4:K").setValues(currentTransferInfo);
    }
    else{
      k_transferSheet.getRange("B4:K").clearContent();
    }

    // 复制期初停薪留职人员信息
    const previousUnpaidLeaveInfo = k_unPaidLeaveSheet.getRange("B4:S").getDisplayValues();
    const currentUnpaidLeaveInfo = k_importRangeSheet.getRange("Y4:AP").getDisplayValues();
    // unPaidLeaveSheet.getRange("Y4:AP").clearContent();
    // unPaidLeaveSheet.getRange("B4:S").clearContent();
    k_unPaidLeaveSheet.getRange("Y4:AP").setValues(previousUnpaidLeaveInfo);
    k_unPaidLeaveSheet.getRange("B4:S").setValues(currentUnpaidLeaveInfo);





    // replace all spaces in sheet 员工信息表 with ""
    const range1 = k_infoSheet.getRange("B:T");
    const range2 = k_infoSheet.getRange("V:AM");
    range1.setValues(range1.getValues()
          .map(function (row) {
            return row.map(function (cell) {
              return cell == ' ' ? '' : cell;
            });
          }));
    range2.setValues(range2.getValues()
          .map(function (row) {
            return row.map(function (cell) {
              return cell == ' ' ? '' : cell;
            });
          }));

    // replace the blank cells and / in the 职级 column in sheet 员工信息表 with "不适用"
    const levelRange = k_infoSheet.getRange(4, 6, k_infoSheet.getRange("B4:B").getValues().filter(String).length, 1);
    levelRange.createTextFinder("/").replaceAllWith("不适用");
    levelRange.setValues(levelRange.getValues()
          .map(function (row) {
              return row.map(function (cell) {
                  return cell === '' ? '不适用' : cell;
              });
          }));

    // replace the blank cells in the 招聘 colum in sheet 员工信息表 with "无可查证"
    const sourceRange = k_infoSheet.getRange(4, 17, k_infoSheet.getRange("B4:B").getValues().filter(String).length, 1);
    sourceRange.createTextFinder("/").replaceAllWith("无可查证");
    sourceRange.setValues(sourceRange.getValues()
            .map(function (row) {
              return row.map(function (cell) {
                return cell === '' ? '无可查证' : cell;
              });
            }));

    // replace the blank cells in 宿舍 column in sheet 员工信息表 with "其他"
    const accomRange = k_infoSheet.getRange(4, 18, k_infoSheet.getRange("B4:B").getValues().filter(String).length, 1);
    accomRange.createTextFinder("/").replaceAllWith("其他");
    accomRange.setValues(accomRange.getValues()
            .map(function (row) {
              return row.map(function (cell) {
                return cell === '' ? '其他' : cell;
              });
            }));

    // replace all spaces in sheet 离职 with ""
    const rangeResign = k_resignSheet.getRange("B:P");
    rangeResign.setValues(rangeResign.getValues()
            .map(function (row) {
              return row.map(function (cell) {
                return cell === ' ' ? '' : cell;
              });
            }));

    // replace the blank cells in 招聘管道 row in sheet 离职 with "无可查证"
    const sourceRangeResign = k_resignSheet.getRange(4, 13, k_resignSheet.getRange("B4:B").getValues().filter(String).length, 1);
    sourceRangeResign.createTextFinder("/").matchCase(true).matchEntireCell(true).replaceAllWith("无可查证");
    sourceRangeResign.setValues(sourceRangeResign.getValues()
            .map(function (row) {
              return row.map(function (cell){
                return cell === '' ? '无可查证' : cell;
              });
            }));

    // replace qb线事业部 with 部门's names in all sheets except the resign sheet
    myComponents.replaceDeptNames();
    // replace qb线事业部 with 部门's names in 离职 sheet
    myComponents.replaceResignSheetDeptNames();

    // clear all previous notes
    myComponents.clearAllNotes();

    // fill Short Term Staff In the Onboard Sheet
    myComponents.fillShortTermStaffData();

  }
}
```