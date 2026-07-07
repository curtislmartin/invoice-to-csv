class Spreadsheet {
	constructor(
		instructionsSheet,
		inputInvoices,
		invoiceKey,
		invoiceKey_AccNum,
		invoice_date_cell,
		due_date_cell,
		templateSheetName,
		titleCell,
		counter_cell,
	) {
		this.ss = SpreadsheetApp.getActiveSpreadsheet();
		this.instructionsSheet =
			this.ss.getSheetByName(instructionsSheet);
		this.inputInvoices = this.ss.getSheetByName(inputInvoices);
		this.invoiceKey = this.ss.getSheetByName(invoiceKey);
		this.invoiceKey_AccNum =
			this.ss.getSheetByName(invoiceKey_AccNum);
		this.issue_date = this.instructionsSheet
			.getRange(invoice_date_cell)
			.getValue();
		this.due_date = this.instructionsSheet
			.getRange(due_date_cell)
			.getValue();
		this.templateSheet = this.ss.getSheetByName(templateSheetName);
		this.titleCell = titleCell;
		this.counter_cell = counter_cell;
		this.counter = this.instructionsSheet
			.getRange(counter_cell)
			.getValue();
		this.issue_date_col = 8;
		this.due_date_col = 9;
		this.invoiceCSVSheetName = null;
	}

	getInputSheetValues(...args) {
		return this.getInputSheetRange(...args).getValues();
	}

	getInputSheetRange(row, col, numRows, numCols) {
		if (!row && !col && !numRows && !numCols)
			return this.inputInvoices.getDataRange();
		else if (row && col && numRows && numCols) {
			return this.inputInvoices.getRange(
				row,
				col,
				numRows,
				numCols,
			);
		}
		throw new Error("Invalid number of arguments");
	}

	addDates(i, issueDate, dueDate) {
		this.inputInvoices
			.getRange(i + 1, this.issue_date_col)
			.setValue(issueDate);
		this.inputInvoices
			.getRange(i + 1, this.due_date_col)
			.setValue(dueDate);
	}

	duplicateSheet() {
		this.newSheetName = this.instructionsSheet
			.getRange(this.titleCell)
			.getValue();

		let newSheet = this.ss.getSheetByName(this.newSheetName);
		if (newSheet === null) {
			newSheet = this.templateSheet.copyTo(this.ss);
			newSheet.setName(this.newSheetName);
			newSheet.showSheet();
		}
		this.invoiceCSVSheetName = this.newSheetName;
	}

	clearInputSheet() {
		const lastRow = this.inputInvoices.getLastRow();
		const lastColumn = this.inputInvoices.getLastColumn();
		if (lastRow < 2) {
			return;
		}
		this.inputInvoices
			.getRange(2, 2, lastRow - 1, lastColumn - 1)
			.clearContent();
	}
}

function mainFunction() {
	const sheet = new Spreadsheet(
		"Instructions",
		"InputInvoices",
		"InvoiceKey_Items",
		"InvoiceKey_AccNum",
		"B13",
		"B14",
		"SalesInvoiceTemplate",
		"B15",
		"D6",
	);

	addIssueAndDueDate(sheet, sheet.issue_date, sheet.due_date);
	sheet.duplicateSheet();
	addInvoiceNumbers(sheet, sheet.counter); //adds invoice number to columns

	lookupAccNum(sheet.inputInvoices, sheet.invoiceKey_AccNum); // finds relevant account number by key
	lookupInvoiceKey(sheet.inputInvoices, sheet.invoiceKey, sheet); // finds relevant unit amount, description and unit code by key
	copyToCSVSheet(sheet.inputInvoices, sheet); // copies InputInvoices to newly created sheet
}

function addIssueAndDueDate(sheet, val1, val2) {
	const values = sheet.getInputSheetValues();

	for (let i = 1; i < values.length; i++) {
		let hasData = values[i].slice(2).some((cell) => cell != "");
		if (hasData) {
			sheet.addDates(i, val1, val2);
		}
	}
}

function addInvoiceNumbers(sheet, counter) {
	const dataValues = sheet.getInputSheetValues();
	if (dataValues.length < 2) {
		return;
	}

	const invoiceNumbers = dataValues.slice(1).reduce((acc, row) => {
		const name = row[1];
		if (name !== "" && !acc[name]) {
			acc[name] = "INV-" + counter++;
		}
		return acc;
	}, {});

	let outputRange = sheet.getInputSheetRange(
		2,
		10,
		dataValues.length - 1,
		1,
	);
	outputRange.setValues(
		dataValues.slice(1).map((row) => [invoiceNumbers[row[1]] ?? ""]),
	);

	sheet.instructionsSheet.getRange(sheet.counter_cell).setValue(counter);
}

function lookupAccNum(inputInvoices, invoiceKey_AccNum) {
	const lastRow = inputInvoices.getLastRow();
	if (lastRow < 2) {
		return;
	}
	const lastRowAccSheet = invoiceKey_AccNum.getLastRow();

	const accNumValues = invoiceKey_AccNum
		.getRange(1, 1, lastRowAccSheet, 2)
		.getValues();

	const inputValues = inputInvoices
		.getRange(2, 6, lastRow - 1, 1)
		.getValues();

	const outputValues = inputValues.map(([value]) => {
		if (value === "") {
			return [""];
		}

		const foundRow = accNumValues.find((row) => row[0] === value);
		if (foundRow) {
			return [foundRow[1]];
		} else {
			return [""];
		}
	});

	inputInvoices.getRange(2, 12, lastRow - 1, 1).setValues(outputValues);
}

function lookupInvoiceKey(inputInvoices, invoiceKey, sheet) {
	const lastRow = inputInvoices.getLastRow();
	if (lastRow < 2) {
		return;
	}

	const lastRowItemSheet = invoiceKey.getLastRow();
	const invoiceKeyValues = invoiceKey
		.getRange(1, 1, lastRowItemSheet, 5)
		.getValues();

	const inputValues = sheet
		.getInputSheetRange(2, 1, lastRow - 1, 15)
		.getValues();

	inputValues.forEach((row, i) => {
		const value = row[2];

		if (value === "") {
			return;
		}

		let matchFound = false;

		for (let j = 0; j < invoiceKeyValues.length; j++) {
			const [
				keyValue,
				column2Value,
				,
				column4Value,
				column5Value,
			] = invoiceKeyValues[j];

			if (value === keyValue) {
				if (row[6] === "Activity Cost") {
					inputInvoices
						.getRange(i + 2, 11)
						.setValue(row[0]);
				} else {
					inputInvoices
						.getRange(i + 2, 11)
						.setValue(column5Value);
				}
				inputInvoices
					.getRange(i + 2, 14)
					.setValue(column2Value);

				const column13Value =
					row[14] === true
						? column4Value +
							" Cancellation Fee - " +
							Utilities.formatDate(
								row[3],
								Session.getScriptTimeZone(),
								"dd/MM/yy",
							)
						: column4Value +
							" " +
							Utilities.formatDate(
								row[3],
								Session.getScriptTimeZone(),
								"dd/MM/yy",
							);

				inputInvoices
					.getRange(i + 2, 13)
					.setValue(column13Value);

				matchFound = true;
				break;
			}
		}

		if (!matchFound) {
			inputInvoices.getRange(i + 2, 11).setValue("");
			inputInvoices.getRange(i + 2, 13).setValue("");
			inputInvoices.getRange(i + 2, 14).setValue("");
		}
	});
}

function copyToCSVSheet(inputInvoices, sheet) {
	const csvSheet = sheet.ss.getSheetByName(sheet.invoiceCSVSheetName);

	if (csvSheet === null) {
		Logger.log("Sheet not found: " + sheet.invoiceCSVSheetName);
		return;
	}

	const inputValues = inputInvoices.getDataRange().getValues();
	const data = [];

	for (let i = 1; i < inputValues.length; i++) {
		const input = inputValues[i];
		if (input[1] === "") {
			break;
		}

		const row = new Array(27).fill("");
		row[0] = input[1]; // col 2
		row[10] = input[9]; // col 10
		row[11] = input[6]; // col 7
		row[12] = input[7]; // col 8
		row[13] = input[8]; // col 9
		row[14] = input[13]; // col 14
		row[15] = input[12]; // col 13
		row[16] = input[4]; // col 5
		row[17] = input[10]; // col 11
		row[19] = input[11]; // col 12
		row[20] =
			input[6] == "Activity Cost"
				? "GST on Income"
				: "GST Free Income";
		row[26] = "Custom Invoice template";
		data.push(row);
	}

	if (data.length > 0) {
		csvSheet.getRange(2, 1, data.length, data[0].length).setValues(
			data,
		);
	}
}
