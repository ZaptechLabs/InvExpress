function process_invoice(appSheetFilePath) {
    var file = get_uploaded_file(appSheetFilePath);
    if (!file) {
        Logger.log('File not found or could not be retrieved for path: ' + appSheetFilePath);
        return;
    }

    Logger.log('Processing uploaded file: ' + file.getName());

    try {
        var ocrResult = extract_text_from_image(file);
        process_invoice_data(ocrResult);
        Logger.log('Successfully processed file: ' + file.getName());
    } catch (e) {
        Logger.log('Error processing file ' + file.getName() + ': ' + e.message);
    }
}

function get_uploaded_file(appSheetFilePath) {
    try {
        // Split the path and extract relevant parts
        var pathParts = appSheetFilePath.split('/');
        var folderName = pathParts[0]; // "Table 1_Files"
        var fileName = pathParts[pathParts.length - 1]; // The actual file name

        // Find the "Table 1_Files" folder
        var folders = DriveApp.getFoldersByName(folderName);
        if (!folders.hasNext()) {
            Logger.log('Folder not found: ' + folderName);
            return null;
        }
        var folder = folders.next();

        // Search for the file within this folder
        var files = folder.getFilesByName(fileName);
        if (files.hasNext()) {
            var file = files.next();
            Logger.log('Successfully retrieved file: ' + fileName);
            return file;
        } else {
            Logger.log('File not found: ' + fileName);
            return null;
        }
    } catch (e) {
        Logger.log('Error retrieving file: ' + e.message);
        return null;
    }
}

function extract_text_from_image(file) {
    var apiKey = getPrivateProperty('OCR_KEY');
    var invoiceOcrEndpoint = getPrivateProperty('End_URL');

    try {
        var imageBlob = file.getBlob();
        var boundary = Utilities.getUuid();

        var payload = Utilities.newBlob(
            '--' + boundary + '\r\n' +
            'Content-Disposition: form-data; name="document"; filename="' + file.getName() + '"\r\n' +
            'Content-Type: ' + imageBlob.getContentType() + '\r\n\r\n'
        ).getBytes();
        payload = payload.concat(imageBlob.getBytes());
        payload = payload.concat(Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes());

        var options = {
            'method': 'post',
            'payload': payload,
            'headers': {
                'Authorization': 'Token ' + apiKey,
                'Content-Type': 'multipart/form-data; boundary=' + boundary
            },
            'muteHttpExceptions': true
        };

        var response = UrlFetchApp.fetch(invoiceOcrEndpoint, options);
        var responseCode = response.getResponseCode();
        if (responseCode !== 201) {
            throw new Error('Error: ' + responseCode + ' - ' + response.getContentText());
        }
        var json = JSON.parse(response.getContentText());
        return json;
    } catch (e) {
        Logger.log('Error in OCR API request: ' + e.message);
        throw e;
    }
}

function process_invoice_data(jsonData) {
    Logger.log('Starting to process invoice data');
    var spreadsheetId = getPrivateProperty('SHEET_ID');
    Logger.log('Attempting to open spreadsheet with ID: ' + spreadsheetId);

    try {//throw error if cannot open
        var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        Logger.log('Successfully opened spreadsheet');
    } catch (e) {
        Logger.log('Error opening spreadsheet: ' + e.message);
        return;
    }

    var allDataSheetName = "Invoice Data";
    Logger.log('Attempting to access sheet: ' + allDataSheetName);

    var allDataSheet = spreadsheet.getSheetByName(allDataSheetName);

    if (!allDataSheet) {
        Logger.log('Sheet not found. Creating new sheet: ' + allDataSheetName);
        allDataSheet = spreadsheet.insertSheet(allDataSheetName);
    } else {
        Logger.log('Successfully accessed sheet: ' + allDataSheetName);
    }

    // New sheet for invoice information only
    var invoiceInfoSheetName = "Invoice Information";
    Logger.log('Attempting to access sheet: ' + invoiceInfoSheetName);

    var invoiceInfoSheet = spreadsheet.getSheetByName(invoiceInfoSheetName);

    if (!invoiceInfoSheet) {
        Logger.log('Sheet not found. Creating new sheet: ' + invoiceInfoSheetName);
        invoiceInfoSheet = spreadsheet.insertSheet(invoiceInfoSheetName);
    } else {
        Logger.log('Successfully accessed sheet: ' + invoiceInfoSheetName);
    }

    var lastRowAllData = allDataSheet.getLastRow();
    var lastRowInvoiceInfo = invoiceInfoSheet.getLastRow();
    Logger.log('Last row with content in all data sheet: ' + lastRowAllData);
    Logger.log('Last row with content in invoice info sheet: ' + lastRowInvoiceInfo);

    if (lastRowAllData === 0) {
        var allDataHeaders = ["Invoice Number", "Date", "Due Date", "Reference Number", "Customer Name", "Customer Address",
            "Supplier Name", "Supplier Address", "Supplier Phone", "Total Net", "Tax Amount", "Total Amount",
            "Item Description", "Product Code", "Quantity", "Unit Price", "Total Amount By Product", "Status"];
        allDataSheet.getRange(1, 1, 1, allDataHeaders.length).setValues([allDataHeaders]);
        lastRowAllData = 1;
        Logger.log('Headers set in the first row of all data sheet');
    }

    if (lastRowInvoiceInfo === 0) {
        var invoiceInfoHeaders = ["Invoice Number", "Date", "Due Date", "Reference Number",
            "Supplier Name", "Supplier Address", "Supplier Phone", "Total Net", "Tax Amount", "Total Amount", "Status"];
        invoiceInfoSheet.getRange(1, 1, 1, invoiceInfoHeaders.length).setValues([invoiceInfoHeaders]);
        lastRowInvoiceInfo = 1;
        Logger.log('Headers set in the first row of invoice info sheet');
    }

    Logger.log('Extracting invoice details from JSON data');//For comprehensive invoice information
    var prediction = jsonData.document.inference.prediction;
    var mainData = [
        prediction.invoice_number.value,
        prediction.date.value,
        prediction.due_date.value,
        prediction.reference_numbers && prediction.reference_numbers.length > 0 ? prediction.reference_numbers[0].value : '',
        prediction.customer_name.value,
        prediction.customer_address.value,
        prediction.supplier_name ? prediction.supplier_name.value : '',
        prediction.supplier_address ? prediction.supplier_address.value : '',
        prediction.supplier_phone_number ? prediction.supplier_phone_number.value : '',
        prediction.total_net ? prediction.total_net.value : '',
        prediction.total_tax ? prediction.total_tax.value : '',
        prediction.total_amount.value
    ];

    var invoiceData = [ //For Invoice Information
        prediction.invoice_number.value,
        prediction.date.value,
        prediction.due_date.value,
        prediction.reference_numbers && prediction.reference_numbers.length > 0 ? prediction.reference_numbers[0].value : '',
        prediction.supplier_name ? prediction.supplier_name.value : '',
        prediction.supplier_address ? prediction.supplier_address.value : '',
        prediction.supplier_phone_number ? prediction.supplier_phone_number.value : '',
        prediction.total_net ? prediction.total_net.value : '',
        prediction.total_tax ? prediction.total_tax.value : '',
        prediction.total_amount.value
    ];

    Logger.log('Processing line items');
    var lineItems = prediction.line_items; //For Line Items
    var lineItemData = lineItems.map(function (item) {
        return [
            item.description,
            item.product_code,
            item.quantity,
            item.unit_price,
            item.total_amount
        ];
    });

    var combinedData = lineItemData.map(function (item) {
        return mainData.concat(item);
    });

    Logger.log('Writing combined data to all data sheet'); //writing to invoice data sheet (record of information for each invoice + line items)
    try {
        allDataSheet.getRange(lastRowAllData + 1, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
        Logger.log('Successfully wrote combined data to all data sheet');
    } catch (e) {
        Logger.log('Error writing combined data to all data sheet: ' + e.message);
        return;
    }

    Logger.log('Writing invoice information to invoice info sheet'); //writing to invoice information sheet
    try {
        invoiceInfoSheet.getRange(lastRowInvoiceInfo + 1, 1, 1, invoiceData.length).setValues([invoiceData]);
        Logger.log('Successfully wrote invoice information to invoice info sheet');
    } catch (e) {
        Logger.log('Error writing invoice information to invoice info sheet: ' + e.message);
        return;
    }

    Logger.log('Auto-resizing columns for all data sheet');
    allDataSheet.autoResizeColumns(1, combinedData[0].length);

    Logger.log('Auto-resizing columns for invoice info sheet');
    invoiceInfoSheet.autoResizeColumns(1, mainData.length);

    Logger.log('Finished processing invoice data');
}


function doPost(e) {
    var appSheetFilePath = e.parameter.filePath;
    if (appSheetFilePath) {
        process_invoice(appSheetFilePath);
        return ContentService.createTextOutput("Invoice processing started for file: " + appSheetFilePath);
    } else {
        return ContentService.createTextOutput("No file path provided").setMimeType(ContentService.MimeType.TEXT);
    }
}