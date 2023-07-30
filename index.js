const csvtojson = require("csvtojson");
const fs = require("fs");
const Excel = require("exceljs");
const { sortBy } = require("lodash");

const { insightsObj, CATEGORY } = require("./enums");

const readExcel = async (excelFilePath) => {
    const resp = await csvtojson().fromStream(fs.createReadStream(excelFilePath));
    const docs = resp.reverse();
    const excelSheet = createExcel(docs);

    excelSheet.xlsx.writeFile(`${excelFilePath.split(".")[0]}_Insights.xlsx`);
}

const createExcel = (docs) => {
    const workbook = new Excel.Workbook();
    const insightsTab = workbook.addWorksheet("Insights");
    const transactionsTab = workbook.addWorksheet("Actual Transactions");

    insightsTab.columns = [{ header: "Category", key: "Category", width: 20 }, { header: "Amount", key: "Amount", width: 20 }, 
        { header: "Number of Transactions", key: "Number of Transactions", width: 50 }];

    for(const doc of docs) {
        doc.Deposit = sanitizeNumber(doc.Deposit);
        doc.Withdrawal = sanitizeNumber(doc.Withdrawal);
        if(doc.Deposit) {
            insightsObj[CATEGORY.INCOMING].Amount += parseInt(doc.Deposit);
            insightsObj[CATEGORY.INCOMING]["Number of Transactions"]++;
            doc.Color = insightsObj[CATEGORY.INCOMING].Color;
            doc.Category = CATEGORY.INCOMING;
        } else {
            if(includesText(insightsObj[CATEGORY.CASHED_OUT].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.CASHED_OUT, doc);
            } else if(includesText(insightsObj[CATEGORY.EATING_OUT].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.EATING_OUT, doc);
            }  else if(includesText(insightsObj[CATEGORY.GROCERIES].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.GROCERIES, doc);
            }else if(includesText(insightsObj[CATEGORY.HOSPITAL].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.HOSPITAL, doc);
            } else if(includesText(insightsObj[CATEGORY.INSURANCE].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.INSURANCE, doc);
            } else if(includesText(insightsObj[CATEGORY.INVESTMENT].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.INVESTMENT, doc);
            } else if(includesText(insightsObj[CATEGORY.LENDING].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.LENDING, doc);
            } else if(includesText(insightsObj[CATEGORY.MEDICINES].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.MEDICINES, doc);
            } else if(includesText(insightsObj[CATEGORY.MISCELLANEOUS].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.MISCELLANEOUS, doc);
            } else if(includesText(insightsObj[CATEGORY.SHOPPING].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.SHOPPING, doc);
            } else if(includesText(insightsObj[CATEGORY.TRAVEL].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.TRAVEL, doc);
            } else if(includesText(insightsObj[CATEGORY.UNCATEGORISED].Reciepients, doc.Transaction)) {
                processObj(insightsObj,CATEGORY.UNCATEGORISED, doc);
            } else {
                processObj(insightsObj,CATEGORY.NOT_FOUND, doc);
            }
        }
    }

    transactionsTab.columns = Object.keys(docs[0]).map(key => ({ header: key, key, width: 15 }));
    const insightsRows = Object.keys(insightsObj).map(key => ({ "Category": key, "Amount": insightsObj[key].Amount, "Number of Transactions": insightsObj[key]["Number of Transactions"] }));
    insightsTab.addRows(insightsRows);

    docs = sortBy(docs,["Color"]);
    transactionsTab.addRows(docs);

    transactionsTab.eachRow((row, rowNum) => {
        row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: row.values[7] == 'Color' ? 'ffb799ff' : row.values[7] }
        }
    })

    return workbook;
}

const processObj = (insightsObj, category, doc) => {
    insightsObj[category].Amount += parseInt(doc.Withdrawal);
    insightsObj[category]["Number of Transactions"]++;
    doc.Color = insightsObj[category].Color;
    doc.Category = category;
}

const sanitizeNumber = (str) => {
    if(!str)    return 0;

    return parseFloat(str.split(",").join(""))
}

const includesText = (dataArr, fullText) => {
    for(const item of dataArr) {
        if(fullText.includes(item)) return true;
    }

    return false;
}

readExcel("AccountTransactions07292023420590.csv");
