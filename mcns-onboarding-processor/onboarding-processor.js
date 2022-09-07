const readXlsxFile = require('read-excel-file/node')
const fs = require('fs');
const inputFolder = "./input";
const outputFolder = "output/";
const arrayName = "mcns-authorizer-sit";

function appendItem(item) {
    let type = "";
    if (item.PutRequest.Item.pk.S.includes("email"))
        type = "email";
    else if (item.PutRequest.Item.pk.S.includes("sms"))
        type = "sms";

    let appId = item.PutRequest.Item.pk.S.substring(0, item.PutRequest.Item.pk.S.search("#"));
    let fileName = appId + "-" + type + "-authorizer.json";

    if (fs.existsSync(outputFolder + fileName)) {
        let content = fs.readFileSync(outputFolder + fileName);
        let outputData = JSON.parse(content);
        item.PutRequest.Item.pk.S += outputData[arrayName].length + 1;
        outputData[arrayName].push(item);
        let fileData = JSON.stringify(outputData, null, 2);
        let outputFileName = outputFolder + fileName;
        fs.writeFileSync(outputFileName, fileData);
        console.log(outputFileName + " Updated...");
    } else {
        let outputData = {};
        outputData[arrayName] = [];
        item.PutRequest.Item.pk.S += "1";
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
            let sheetsProcessed = 0;

            let sheets = await readXlsxFile.readSheetNames(fileName).then((sheetNames) => {
                return sheetNames;
            });

            let totalSheets = sheets.length;

            if (sheets.includes("Email") || sheets.includes("SMS")) {
                // For processing single file
                console.log("Processing single sheet file with multiple items...");
                if (sheets.includes("Email")) {
                    let requestProcessed = await processLongSheet(fileName, "Email");
                    console.log("Items Processed: ", requestProcessed);
                }
                if (sheets.includes("SMS")) {
                    let requestProcessed = await processLongSheet(fileName, "SMS");
                    console.log("Items Processed: ", requestProcessed);
                }
            }
            else {
                // For processing Multiple sheets file
                console.log("Processing file with multiple sheets...");
                for (let i = 3; i < totalSheets; i++) {
                    let sheetItem = await processSheet(fileName, i);
                    appendItem(sheetItem);
                    sheetsProcessed++;
                }
                console.log("Items Processed: ", sheetsProcessed);
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

function processSheet(fileName, sheetNumber) {
    return new Promise(function (resolve, reject) {
        try {
            readXlsxFile(fileName, { sheet: sheetNumber }).then((rows) => {
                resolve(processItem(rows));
            })
        } catch (e) {
            reject(e);
        }
    })
}

function processLongSheet(fileName, sheetName) {
    return new Promise(function (resolve, reject) {
        try {
            readXlsxFile(fileName, { sheet: sheetName }).then((rows) => {
                let requestProcessed = 0;
                for (let i = 0; i < rows.length; i++) {
                    if (rows[i][1] == "Request Type:") {
                        i = i + 2;
                        let tempItem = [];
                        //Add each item into a temp array for processing
                        while (rows[i] && rows[i][1] != "Request Type:" && i < rows.length) {
                            tempItem.push(rows[i]);
                            i++;
                        }
                        appendItem(processItem(tempItem));
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
    let processedItem = {
        PutRequest: {
            Item: {}
        }
    };
    let templateValueSchemaObj = {
        $schema: "https://json-schema.org/draft/2020-12/schema",
        $id: "https://dsta.gov.sg/app1.schema.json",
        title: "app1",
        type: "object",
        properties: {},
        required: []
    }
    let applicationId = "";
    let subject;
    let templateValueSchema;
    let template;
    let regexEspression;

    let sms = false;
    for (let i = 0; i < item.length; i++) {
        let row = item[i];
        for (let j = 0; j < row.length; j++) {
            let column = row[j];
            if (column == "Application ID") {
                applicationId = row[j + 2];
                if (applicationId) {
                    applicationId = applicationId.replaceAll(" ", "");
                }
            }
            if (column == "Sender") {
                let sender = row[j + 2];
                if (sender) {
                    let insertIndex = applicationId.search("#")
                    applicationId = applicationId.slice(0, insertIndex) + "#" + sender.replaceAll(" ", "") + applicationId.slice(insertIndex);
                }
            }
            if (column == "Channel Type") {
                let channelType = row[j + 2];
                if (channelType) {
                    if (channelType.toLowerCase() == "email")
                        applicationId += "#email";
                    else if (channelType.toLowerCase() == "sms") {
                        applicationId += "#sms";
                        subject = "MCNS Notification";
                        sms = true;
                    }
                }
            }
            if (column == "Subject") {
                //if sms subject = MCNS Notification
                subject = row[j + 2];;
                if (sms) {
                    subject = "MCNS Notification";
                }
            }
            if (column == "Template Values") {
                for (let k = i; k < j + 20; k++) {
                    if (item[k][j] == "Template Values' Regular Expression") {
                        break;
                    }
                    else {
                        if (item[k][j + 1] == "Text" || item[k][j + 1] == "Number") {
                            templateValueSchemaObj.properties[item[k][j + 2]] = {};
                            templateValueSchemaObj.properties[item[k][j + 2]].description = {};
                            templateValueSchemaObj.properties[item[k][j + 2]].description = item[k][j + 2] + "field";
                            templateValueSchemaObj.properties[item[k][j + 2]].type = "string";
                            templateValueSchemaObj.required.push(item[k][j + 2]);
                        }
                    }
                }
                if (templateValueSchemaObj.required.length > 0) {
                    templateValueSchema = JSON.stringify(templateValueSchemaObj);
                }
            }
            if (column == "Template") {
                template = row[j + 2];
                if (template) {
                    template = template.replace(/\n/g, ' ');
                }
            }
            if (column == "Template Values' Regular Expression") {
                regexEspression = row[j + 2];
                if (regexEspression) {
                    regexEspression = regexEspression.replace(/\n/g, "");
                    regexEspression = regexEspression.replaceAll("    ", "");
                    regexEspression = regexEspression.replaceAll(": ", ":");
                    regexEspression = regexEspression.replace(/\\/g, "\\\\");
                }
            }
        }
    }
    processedItem.PutRequest.Item.pk = {};
    processedItem.PutRequest.Item.pk.S = applicationId;
    if (subject) {
        processedItem.PutRequest.Item.subject = {};
        processedItem.PutRequest.Item.subject.S = subject;
    }
    if (templateValueSchema) {
        processedItem.PutRequest.Item.templateValueSchema = {};
        processedItem.PutRequest.Item.templateValueSchema.S = templateValueSchema;
    }
    if (template) {
        processedItem.PutRequest.Item.template = {};
        processedItem.PutRequest.Item.template.S = template;
    }
    if (regexEspression) {
        processedItem.PutRequest.Item.regEx = {};
        processedItem.PutRequest.Item.regEx.S = regexEspression;
    }
    processedItem.PutRequest.Item.pkstatus = {};
    processedItem.PutRequest.Item.pkstatus.S = "true";
    processedItem.PutRequest.Item.attachmentAllowed = {};
    processedItem.PutRequest.Item.attachmentAllowed.S = "true";
    return processedItem;
}

processFile();