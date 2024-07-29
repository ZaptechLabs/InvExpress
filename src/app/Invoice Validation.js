function validate_invoices() {
    //setSheetId();
    const sheetId = getPrivateProperty("SHEET_ID");
    var spreadsheetId = sheetId;
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);


    var purchaseOrderSheet = spreadsheet.getSheetByName('Purchase Order');
    var invoiceSheet = spreadsheet.getSheetByName('Invoice Data');

    var validationSheetName = 'Invoice Validation';
    var validationSheet = spreadsheet.getSheetByName(validationSheetName);

    if (!validationSheet) {
        Logger.log('Sheet not found. Creating new sheet: ' + validationSheetName);
        validationSheet = spreadsheet.insertSheet(validationSheetName);
    } else {
        Logger.log('Successfully accessed sheet: ' + validationSheetName);
    }


    var purchaseOrderData = purchaseOrderSheet.getDataRange().getValues();
    var invoiceData = invoiceSheet.getDataRange().getValues();
    var purchaseOrderMap = {};

    var poHeaders = purchaseOrderData[0];
    var invoiceHeaders = invoiceData[0];




    for (var i = 1; i < purchaseOrderData.length; i++) {
        var poRow = purchaseOrderData[i];
        var poRecord = {};

        // Convert each row into a JSON object using headers as keys
        poHeaders.forEach((header, index) => {
            poRecord[header] = poRow[index];
        });

        var key = poRecord['order_id'] + "_" + poRecord['product_code'];
        purchaseOrderMap[key] = poRecord;
    }


    var validationResultsMap = {};

    // Iterate over each invoice and group them by invoice_number
    for (var j = 1; j < invoiceData.length; j++) {
        var invoiceRow = invoiceData[j];
        var invoiceRecord = {};

        // Convert each row into a JSON object using headers as keys
        invoiceHeaders.forEach((header, index) => {
            invoiceRecord[header] = invoiceRow[index];
        });

        var invoiceKey = invoiceRecord['Reference Number'] + "_" + invoiceRecord['Product Code'];
        var invoiceNumber = invoiceRecord['Invoice Number'];

        // Initialize the validation result for the invoice
        if (!validationResultsMap[invoiceNumber]) {
            validationResultsMap[invoiceNumber] = {
                isValid: true, // Assume valid until a mismatch is found
                orderId: invoiceRecord['Reference Number'],
                checkedProductCodes: new Set(),
                dueDate: invoiceRecord['Due Date']
            };
        }

        // Check each line item against the purchase order map
        if (purchaseOrderMap[invoiceKey]) {
            var poData = purchaseOrderMap[invoiceKey];


            if (
                poData['quantity'] !== invoiceRecord['Quantity'] ||
                poData['unit_price'] !== invoiceRecord['Unit Price'] ||
                //poData['tax_rate'] !== invoiceRecord['Tax Amou'] ||
                poData['total_net'] !== invoiceRecord['Total Net'] ||
                poData['tax_amount'] !== invoiceRecord['Tax Amount'] ||
                poData['total_amount'] !== invoiceRecord['Total Amount']
            ) {
                // If any line item is invalid, mark the entire invoice as invalid
                validationResultsMap[invoiceNumber].isValid = false;
            }
        } else {
            // No matching purchase order line found, mark as invalid
            validationResultsMap[invoiceNumber].isValid = false;
        }

        // Add the product code to the set for this invoice
        validationResultsMap[invoiceNumber].checkedProductCodes.add(invoiceRecord['Product Code']);
    }

    // Prepare validation results array
    var validationResults = [];
    var paymentStats = "Unpaid";
    // Populate validation results with date, invoice number, order id, and validation result
    for (var invoiceNumber in validationResultsMap) {
        var result = validationResultsMap[invoiceNumber];
        var validationStatus = result.isValid ? 'valid' : 'invalid';
        var validationApproval = validationStatus === 'valid' ? 'Pending' : 'Rejected';
        validationResults.push([invoiceNumber, result.dueDate, result.orderId, validationStatus, validationApproval, paymentStats]);
    }

    //update invoice validation sheet
    validationSheet.appendRow(["Invoice Number", "Due Date", "Order ID", "Validation Result", "Approval", "Payment Status"]);
    validationSheet.getRange(2, 1, validationResults.length, 6).setValues(validationResults);
}
