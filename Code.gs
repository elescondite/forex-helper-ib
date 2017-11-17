/*
 * Support functions for Forex Helper
 *
 * Copyright 2017 Julian R. Smith <beachguy@codesmith.co>
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *     http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * v1.0.0 2017.10.15 Inital version
 * v1.0.2 2017.11.05 Remove most egregious hard-coding but still a messy hack
 * v1.0.3 2017.11.12 Preliminary TradingView charts
 * v1.0.4 2017.11.16 Use standard currency pair naming (Base/Quote)
 */
// "Define" some constants to avoid hard coding
// For column hiding/showing toggles. 
var namedRangeAccount = 'AccountColumns'
var namedRangeBase = 'BaseColumns'
var namedRangeQuote = 'QuoteColumns'

// Sheet names.
var tabBasket = 'Basket'
var tabEntry = 'Entry'
var tabConfig = 'Config'
var tabArchive = 'Archive'


function openChart() {
	// FIXME. Screen sizes? Pass currency pairs for watchlist.
	var t = HtmlService.createTemplateFromFile('Index')
	var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabEntry);
	var lastSourceRow = sourceSheet.getLastRow();
	var lastSourceCol = sourceSheet.getLastColumn();

	var sourceRange = sourceSheet.getRange(1, 1, lastSourceRow, lastSourceCol);
	var sourceData = sourceRange.getValues();
	var targetValues = [];
    var re = /(Pending|Submitted|Working).*/
    var regExp = new RegExp(re);

	//Loop through every retrieved row from the Source
	for (row in sourceData) {
		//IF Column A in this row is active then work on it.
		if (sourceData[row][0].match(regExp)) {
            tempvalue = sourceData[row][3] + sourceData[row][4];
            targetValues.push(tempvalue);
		}
	}
    t.currencies = targetValues;
    var html =  t.evaluate();
    html.setHeight(1000)
    html.setWidth(1600);

    Logger.log(t.currencies);
	SpreadsheetApp.getUi() 
		.showModalDialog(html, 'ForexHelper Live Chart');
}

function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu('Forex')
		.addItem('Generate Basket', 'createBasket')
		// .addItem('Toggle Inactive Rows', 'toggleInactiveRows')
		.addItem('Archive Inactive Rows', 'archiveInactiveRows')
		.addItem('Toggle Account Currency', 'toggleAccount')
		.addItem('Toggle Base Currency', 'toggleBase')
		.addItem('Toggle Quote Currency', 'toggleQuote')
		.addItem('Open Live Chart', 'openChart')
		.addToUi();
}

function toggleColumns2(namedRange) {
	var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabEntry);

	// Convert named range into starting column and number of columns
	var range = sheet.getRange(namedRange);
	var count = range.getLastColumn() - range.getColumn() + 1;
	Logger.log(range.getColumn());
	Logger.log(range.getLastColumn());
	Logger.log(count);
	// GAS does not have static variable per se so we store these as sheet properties which unfortunately triggers an OAUTH 
	var docProperties = PropertiesService.getDocumentProperties();
	var isHidden = (docProperties.getProperty(namedRange) == 'hidden' ? true :
		false);
	if (isHidden) {
		sheet.showColumns(range.getColumn(), count)
	} else {
		sheet.hideColumns(range.getColumn(), count)
	}
	// Save new state
	docProperties.setProperty(namedRange, (isHidden ? 'shown' : 'hidden'));
}

function toggleInactiveRows() {
	var ss = SpreadsheetApp.getActiveSpreadsheet()
	var sh = ss.getActiveSheet()
	var docProperties = PropertiesService.getdocumentProperties();
	var isHidden = (docProperties.getProperty('HIDDEN_INACTIVE') == 'hidden' ?
		true : false);

	if (isHidden) {
		sh.unhideRow(sh.getDataRange());
	} else {
		var sh;
		var re = /(Expired|Closed|Cancel).*/
		var regExp = new RegExp(re);

		sh = SpreadsheetApp.getActive()
			.getSheetByName(tabEntry)
		sh.getDataRange()
			.getValues()
			.forEach(function(r, i) {
				if (r[0].match(regExp)) {
					sh.hideRows(i + 1)
				}
			})
	}
	docProperties.setProperty('HIDDEN_INACTIVE', (isHidden ? 'shown' : 'hidden'));
}


// These are separate functions <sigh> as GApps cannot pass parameters from a button
function toggleAccount() {
	toggleColumns2(namedRangeAccount);
}

function toggleBase() {
	toggleColumns2(namedRangeBase);
}

function toggleQuote() {
	toggleColumns2(namedRangeQuote);
}

function showAll() {
	toggleColumns2(namedRangeAccount);
	toggleColumns2(namedRangeBase)
	toggleColumns2(namedRangeQuote)
}

/**
 * Make a checkbox menu item. Returns a string with the original text aligned to
 * leave room for a check mark. (Dependent on font, compatible with default
 * font used in Google Spreadsheets as of May 2013. YMMV.)
 *
 * @param {String} name Original item name
 * @param {boolean} check True if menu item should have a check mark
 *
 * From an idea by Kalyan Reddy, http://stackoverflow.com/a/13452486/1677912
 */
function buildItemName(name, check) {
	// Prepend with check+space, or EM Space
	return (check ? "✓ " : "\u2003") + name;
}


/*
 * Create a TWS format compatible basket of trades
 */

function createBasket() {
	var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		tabBasket);

	var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		tabEntry);
	var lastSourceRow = sourceSheet.getLastRow();
	var lastSourceCol = sourceSheet.getLastColumn();

	var sourceRange = sourceSheet.getRange(1, 1, lastSourceRow, lastSourceCol);
	var sourceData = sourceRange.getValues();
	var orderNum = 0;
	var basketTag = 'fhelper';

	var tempvalue = ["Action", "Quantity", "Symbol", "SecType", "Exchange",
		"Currency", "TimeInForce", "OrderType", "LmtPrice", "AuxPrice", "OrderId",
		"ParentOrderId", "BasketTag"
	];
	var targetValues = [];
	targetValues.push(tempvalue);

	// Double check they want to do this
	if (!confirmAction('This will overwrite any existing basket.'))
		return;

	//Loop through every retrieved row from the Source
	for (row in sourceData) {
		//IF Column A in this row is 'Pending then work on it.
		if (sourceData[row][0] === 'Pending') {
			orderNum += 1;
			basketTag = 'fhelper-' + sourceData[row][1] + "-" + formatDate(sourceData[
				row][2]); // Source
			// Create BUY or SELL order
			tempvalue = [sourceData[row][5], sourceData[row][10], sourceData[row][3],
				'CASH', 'IDEALPRO', sourceData[row][4], 'GTC', 'LMT', sourceData[row][7],
				'', orderNum, '', basketTag
			];
			targetValues.push(tempvalue);
			''
			// Create Take Profit order
			if (sourceData[row][5] == 'SELL') // reverse the first order
				target = 'BUY';
			else
				target = 'SELL';
			tempvalue = [target, sourceData[row][10], sourceData[row][3], 'CASH',
				'IDEALPRO', sourceData[row][4], 'GTC', 'LMT', sourceData[row][9], '', '',
				orderNum, basketTag
			];
			targetValues.push(tempvalue);
			// Create Stop Loss order. Note the price goes in AuxPrice column
			if (sourceData[row][5] == 'SELL') // reverse the first order
				target = 'BUY';
			else
				target = 'SELL';
			tempvalue = [target, sourceData[row][10], sourceData[row][3], 'CASH',
				'IDEALPRO', sourceData[row][4], 'GTC', 'STP', '', sourceData[row][8], '',
				orderNum, basketTag
			];
			targetValues.push(tempvalue);
		}
	}

	//Save the new range to the appropriate sheet 
	activeSheet.clearContents();
	activeSheet.getRange(2, 1, targetValues.length, 13).setValues(targetValues);

	// Generate a downloadable url
	var outFileName = 'IBbasket.csv';
	var ssID = SpreadsheetApp.getActiveSpreadsheet().getId();
	var downloadUrl = 'https://docs.google.com/spreadsheets/d/' + ssID +
		'/gviz/tq?tqx=out:csv;outFileName=' + outFileName + '&sheet=' + tabBasket;
	activeSheet.getRange(1, 1 , 1, 1).setValue(downloadUrl);
	Logger.log(downloadUrl);

	// Switch to the basket tab
	activeSheet.activate();

}

/*
 * Copy/delete rows that match a regex to another sheet
 */
function archiveInactiveRows() {

	var s, targetSheet;
	var re = /(Expired|Closed|Cancel).*/
	var regExp = new RegExp(re);

	// Double check they want to do this
	if (!confirmAction(
			'This will move all Expired/Closed/Cancelled rows to the Archive sheet'))
		return;

	s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabEntry);
	targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabArchive);
	var rowCount = s.getLastRow();
	Logger.log(rowCount);

	for (i = rowCount; i > 5; i--) {
		var r = s.getRange(i, 1, 1, s.getLastColumn());
		var values = r.getValues();
		Logger.log(values);
		Logger.log(values[0][0]);
		if (values[0][0].match(regExp)) {
			targetSheet.appendRow(r.getDisplayValues()[0])
			s.deleteRow(i);
			Logger.log("Deleting: " + i);
		}
	}
}


// Simple utility to return YYYY.MM.DD
function formatDate(date) {
	if (Object.prototype.toString.call(date) !== '[object Date]') return '';
	return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy.MM.dd');
}

// Simple Yay/Nay confirmation dialogue
function confirmAction(message, title) {
	// Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
	// close the dialog by clicking the close button in its title bar.
	var ui = SpreadsheetApp.getUi();
	var response = ui.alert(title || 'Confirm', message +
		'\n\nAre you sure you want to continue?', ui.ButtonSet.YES_NO);

	// Process the user's response.
	if (response == ui.Button.YES) {
		Logger.log('The user clicked "Yes."');
		return true;
	} else {
		Logger.log(
			'The user clicked "No" or the close button in the dialog\'s title bar.');
		return false;
	}
}


// Version string if required
function CODESMITH_VERSION() {
	return ("V1.0.4 (c) 2017 Codesmith Co")
}