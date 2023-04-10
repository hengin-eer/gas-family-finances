function selectPeriod() {
  const period = "2023/03" // 期間を指定 "yyyy/MM"
  myFunction(period)
  makeCopyFile(period)
}

let prop = PropertiesService.getScriptProperties().getProperties()
function myFunction(period) {
  const activeSpreadSheet = SpreadsheetApp.openById(prop.ACTIVE_SPREADSHEET)
  const getFormDataSheet = activeSpreadSheet.getSheetByName("フォームの回答")
  const familyFinancesSheet = SpreadsheetApp.openById(prop.FAMILY_FINANCES_SHEET).getSheetByName("家計簿")

  familyFinancesSheet.getRange('B1').setValue(`${period}月間予算表`) // タイトルの更新

  const lastRow = getFormDataSheet.getLastRow()
  const allFormData = getFormDataSheet.getRange(`B2:F${lastRow}`).getDisplayValues()
  const sortByPeriodData = allFormData.filter(data => data[0].includes(period))
  const allExpenditureData = sortByPeriodData.filter(data => data[2] === "支出")
  const allIncomeData = sortByPeriodData.filter(data => data[2] === "収入")
  console.log(allExpenditureData)
  console.log(allIncomeData)


  let eachExpenditureDataSum = [0, 0, 0, 0, 0, 0, 0, 0, 0]
  for (let i = 0; i < allExpenditureData.length; i++) {
    let catData = allExpenditureData[i][3]
    let data = allExpenditureData[i][1]
    switch (catData) {
      case '食費':
        eachExpenditureDataSum[0] += Number(data)
        break
      case "交通":
        eachExpenditureDataSum[1] += Number(data)
        break
      case "外食":
        eachExpenditureDataSum[2] += Number(data)
        break
      case "日用品":
        eachExpenditureDataSum[3] += Number(data)
        break
      case "医療":
        eachExpenditureDataSum[4] += Number(data)
        break
      case "お土産":
        eachExpenditureDataSum[5] += Number(data)
        break
      case "衣類":
        eachExpenditureDataSum[6] += Number(data)
        break
      case "施設利用":
        eachExpenditureDataSum[7] += Number(data)
        break
      default:
        eachExpenditureDataSum[8] += Number(data)
        break
    }
  } // eachExpenditureDataSumにそれぞれのジャンルごとに合計を代入

  let allExpenditureDataSum = 0
  for (let i = 0; i < 9; i++) {
    allExpenditureDataSum += eachExpenditureDataSum[i]
  }
  console.log(`支出の合計:\n${allExpenditureDataSum}`) // allExpenditureDataSumに合計を代入


  let eachIncomeDataSum = [0, 0, 0, 0]
  for (let i = 0; i < allIncomeData.length; i++) {
    let catData = allIncomeData[i][3]
    let data = allIncomeData[i][1]
    switch (catData) {
      case '定期お小遣い':
        eachIncomeDataSum[0] += Number(data)
        break
      case "特殊お小遣い":
        eachIncomeDataSum[1] += Number(data)
        break
      case "クーポン":
        eachIncomeDataSum[2] += Number(data)
        break
      default:
        eachIncomeDataSum[3] += Number(data)
        break
    }
  } // eachIncomeDataSumにそれぞれのジャンルごとに合計を代入

  let allIncomeDataSum = 0
  for (let i = 0; i < 4; i++) {
    allIncomeDataSum += eachIncomeDataSum[i]
  }
  console.log(`収入の合計:\n${allIncomeDataSum}`) // allIncomeDataSumに合計を代入


  for (let i = 0; i < 9; i++) {
    familyFinancesSheet.getRange(`B${5 + i}`).setValue(eachExpenditureDataSum[i])
  }
  for (let i = 0; i < 4; i++) {
    familyFinancesSheet.getRange(`E${5 + i}`).setValue(eachIncomeDataSum[i])
  }
  familyFinancesSheet.getRange('B4').setValue(allExpenditureDataSum)
  familyFinancesSheet.getRange('E4').setValue(allIncomeDataSum)
  // それぞれの金額の合計をセルに代入


  let remainingMoney = allIncomeDataSum - allExpenditureDataSum
  const remainingMoneyRange = familyFinancesSheet.getRange('E10')
  remainingMoneyRange.setValue(remainingMoney)
  // 残金を入力

  let remainingMoneyIsPlus = remainingMoney >= 0
  const moneyBalance = familyFinancesSheet.getRange('D12')
  remainingMoneyIsPlus ? moneyBalance.setValue("黒字です") && moneyBalance.setFontColor('#171717') 
  : moneyBalance.setValue("赤字です") && moneyBalance.setFontColor('#DA0037')
  // 残金から黒字・赤字を判別し代入


  console.log("Success!")
}


function makeCopyFile(period) {
  const templateFile = DriveApp.getFileById(prop.TEMPLATE_FILE)
  const OutputFolder = DriveApp.getFolderById(prop.OUTPUT_FOLDER)
  const newFileName = `${period}月間予算表`

  let newFile = templateFile.makeCopy(newFileName, OutputFolder)
  DriveApp.removeFile(newFile)
}