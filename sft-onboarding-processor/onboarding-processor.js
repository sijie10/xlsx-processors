const readXlsxFile = require('read-excel-file/node')
const fs = require('fs');
const inputFolder = "./input";
const outputFolder = "output/";
const arrayName = "sft-authorizer-sit";

function appendItem(item) {

    let appId = item.PutRequest.Item.pk.S.substring(0, item.PutRequest.Item.pk.S.search("#"));
    let appIdArr = appId.split("-");
    let fileName = "";
    for (let i = 0; i < 3; i++) {
        if (appIdArr[i]) {
            if (fileName.length > 0)
                fileName += "-" + appIdArr[i];
            else
                fileName += appIdArr[i];
        }
    }
    if (appIdArr.includes("dev"))
        fileName += "-dev";
    else if (appIdArr.includes("qa"))
        fileName += "-qa";
    else if (appIdArr.includes("sit"))
        fileName += "-sit";
    else if (appIdArr.includes("prod"))
        fileName += "-prod";
    fileName += "-authorizer.json";

    if (fs.existsSync(outputFolder + fileName)) {
        let content = fs.readFileSync(outputFolder + fileName);
        let outputData = JSON.parse(content);
        outputData[arrayName].push(item);
        let fileData = JSON.stringify(outputData, null, 2);
        let outputFileName = outputFolder + fileName;
        fs.writeFileSync(outputFileName, fileData);
        console.log(outputFileName + " Updated...");
    } else {
        let outputData = {};
        outputData[arrayName] = [];
        outputData[arrayName].push(item);
        let fileData = JSON.stringify(outputData, null, 2);
        let outputFileName = outputFolder + fileName;
        fs.writeFileSync(outputFileName, fileData);
        console.log(outputFileName + " generated...");
    }
}

async function processFile() {
    let filesProcessed = 0;
    let filesArr = [];
    console.log("--- Starting Processing ---");
    fs.readdirSync(inputFolder).forEach(file => {
        filesArr.push(file);
    });

    for (let file of filesArr) {
        if (file.includes(".xlsx")) {
            console.log("\nProcessing: ", file);
            let fileName = "input/" + file;

            let sheets = await readXlsxFile.readSheetNames(fileName).then((sheetNames) => {
                return sheetNames;
            });

            if (sheets.includes("Whitelist Form")) {
                // For processing single file
                console.log("Processing Whitelist Form...");
                let requestProcessed = await processLongSheet(fileName, "Whitelist Form");
                console.log("Items Processed: ", requestProcessed);
            }
            filesProcessed++;
        }
    }


    if (filesProcessed > 0) {
        console.log("\n--- Processing Completed ---");
        console.log("Total files processed: ", filesProcessed);
    }
    else
        console.log("No files in input folder");
}

function processLongSheet(fileName, sheetName) {
    return new Promise(function (resolve, reject) {
        try {
            readXlsxFile(fileName, { sheet: sheetName }).then((rows) => {
                let requestProcessed = 0;
                let itemsProcesed = 0;
                for (let i = 0; i < rows.length; i++) {
                    if (rows[i][1] == "Request Type:") {
                        i = i + 2;
                        let tempItem = [];
                        //Add each item into a temp array for processing
                        while (rows[i] && rows[i][1] != "Request Type:" && i < rows.length) {
                            tempItem.push(rows[i]);
                            i++;
                        }
                        itemsProcesed += processItem(tempItem);
                        requestProcessed++;
                        i--;
                    }
                }
                resolve(requestProcessed);
            })
        } catch (e) {
            reject(e);
        }
    })
}

function processItem(item) {
    //item is a 2D array with row item[row][col]
    let itemsProcessed = 0;
    let processedItem = {
        PutRequest: {
            Item: {}
        }
    };
    processedItem.PutRequest.Item.pk = {};
    processedItem.PutRequest.Item.pkstatus = {};
    processedItem.PutRequest.Item.pkstatus.S = "true";
    let applicationId = "";

    for (let i = 0; i < item.length; i++) {
        let row = item[i];
        for (let j = 0; j < row.length; j++) {
            let column = row[j];
            if (column == "AppId") {
                applicationId = row[j + 1];
                if (applicationId) {
                    applicationId = applicationId.replaceAll(" ", "");
                }
            }
            if (column == "Type") {
                let type = row[j + 1];
                console.log(type);
                processedItem.PutRequest.Item.pk.S = applicationId + "#" + type;
                appendItem(processedItem);
                itemsProcessed++;
            }
        }
    }
    return itemsProcessed;
}

processFile();