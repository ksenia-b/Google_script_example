/**
 * Created by Ксюша on 17.11.2016.
 */
/**
 * Created by Ксюша on 17.11.2016.
 */
ssId = 'spreadsheet_id';
ss = SpreadsheetApp.openById(ssId);
CheckListSheet = ss.getSheetByName("CheckList");
LUSheet =  ss.getSheetByName("LU");
SetUp = ss.getSheetByName("SetUp");


function doGet(e) {

	var template = HtmlService.createTemplateFromFile('index');
	Logger.log("template = "+template);
	template.action = ScriptApp.getService().getUrl();
	Logger.log("template.action = "+template.action);
	return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setTitle('Example Model');


}

function processForm(next1, DataToSave, str, inputCol, inputRow, a) {

	var colCount = 0;
	var rowCount = 0;
	var string = CheckListSheet.getRange("E"+(next1-1)).getValue();
	var loading = "Loading";

	if (DataToSave) {
		saveDataIntoTable(DataToSave, next1 - 1, inputCol, inputRow); // save data from previous step
	}

	do
	{
		var randomWait = Math.floor(Math.random()*100+50);
		Utilities.sleep(randomWait);
		Logger.log("randomWait = "+randomWait);
	}
	while (string.search(loading) ==! null);

	var showIndicator = CheckListSheet.getRange("E"+next1).getValue();  //show questions? yes/no

	if (showIndicator === 'NO') {

		return {
			"next1": ++next1,
			"result": false
		};
	}
	if(!showIndicator) {
		var sheetData = {
			"next1": 0,
			"endOfData" : "true",
			"qValues" : ''
		}
		Logger.log("sheetData.endOfData = "+sheetData.endOfData);

	}
	else {

		var aTValues =  CheckListSheet.getRange("K"+next1).getValue();  //answer type
		var tSource =  CheckListSheet.getRange("J"+next1).getValue();
		var tableDataName = SetUp.getRange("B"+next1).getValue(); //find data for table
		var arrData = '';

		if (tSource) {  //if source exist
			if (aTValues === "Table") {
				var table = getDataIfSourceTable(tSource);
				arrData = table.body; //get data from table
				colCount = table.col;
				rowCount = table.row;
			}
			else {
				colWithFindData = 2; //look for in C-columl
				arrData = getDataForInputRadio(tSource, LUSheet ,colWithFindData); //get array data from spreadsheet
			}
		}
		else {
			tSource = 0;
		}
		var qValues = CheckListSheet.getRange("F"+next1).getValue();
		next1++;
		var sheetData = {
			"aTValues" : aTValues,
			"qValues" : qValues,
			"tSource" : tSource,
			"arrData" : arrData,
			"next1": next1,
			"col": colCount,
			"row": rowCount,
			"result": true
		}
	}
	Logger.log('data ' +sheetData.next1);
	return sheetData;
}

function saveDataIntoTable(data, index, inputCol, inputRow) {
	if (index === 2) return;
	if(typeof data === 'object' && !data.type ) {
		var prevTSource =  CheckListSheet.getRange("J"+(index)).getValue(); //get previous TSource
		setDataIfSourceTable(prevTSource, data, inputCol, inputRow);
		CheckListSheet.getRange("G"+(index)).setValue('wrote');
	} else {
		if(  typeof data === 'object' && data.type ){
			CheckListSheet.getRange("G"+(index)).setValue(data.Data.toString());
		}
		else{
			CheckListSheet.getRange("G"+(index)).setValue(data);
		}
	}
}

function include(filename) {

	return HtmlService.createHtmlOutputFromFile(filename)
		.getContent();

}

function getDataIfSourceTable(tSource){

	var colWithFindData = 1;  //first column with table name which we look for
	var outRow  =  finDataInRange(tSource, SetUp, colWithFindData);
	var range = SetUp.getDataRange();
	var values = range.getValues();
	var row = outRow;
	var firstEmptyRow;
	var rowArrNo = [];
	var jsonArr = [];

	//  first empty row
	var rowArrNo = [];
	var outRow2 = outRow;

	do {
		outRow2++;
		var outRowVal = outRow2;
		var flag = SetUp.getRange(outRow2 , 2).getValue();

	}
	while (flag);

	var firstEmptyRow = outRowVal;

	//  first empty col
	var cellArrNo = [];
	var outCell = 1;
	do {
		outCell++;
		var outCellVal = outCell;
		var flag = SetUp.getRange(outRow + 1, outCell).getValue();
	}
	while (flag);
	var lastCol = outCellVal;

	return {
		body: getJson((ss.getSheetByName("SetUp").getRange(outRow, colWithFindData, firstEmptyRow - outRow, lastCol - colWithFindData).getValues()),lastCol - colWithFindData),
		col: firstEmptyRow,
		row: lastCol
	}

}


function finDataInRange(SourceName, sheetName, colWithFindData) {
	var dataRange = sheetName.getDataRange();
	var values = dataRange.getValues() || [];

	for (var i = 0; i < values.length; i++)
	{
		if (values[i][colWithFindData] === SourceName )
		{
			return i+1;
		}
	}
	return undefined;
}

function getDataForInputRadio(tSource, sheetName, colWithFindData) {

	var arrData = [];
	var flag = false;
	var outRow = finDataInRange(tSource, sheetName, colWithFindData) + 1; //find data in source

	do {
		arrData.push(sheetName.getRange("B"+outRow).getValue());
		outRow++;
		flag = sheetName.getRange("B"+outRow).getValue();
	}
	while (flag);
	return arrData;
}

function getJson(range, colWithData){

	var json = Utilities.jsonStringify(range);
	var jsonArr = [];
	var colWithFindData = 2;  //first column with table name which we look for
	json = JSON.parse(json);

	for (var i = 0; i < json.length; i++) {
		if (json[i][1] === "NO")
			continue;

		var arrJ = [];
		var arrDropDown = [];

		for (var j = 0; j < json[i].length; j++) {

			if (json[1][j] === "NO") {
				continue;
			}
			var cell_val  = json[i][j];

			if (j > 2 && json[5][j] !== "EMPTY" ) {
				if(a) {
					var arrValue = a;
					cell_val = {
						"title": json[6][j],
						"value" : {
							"arrData": arrValue
						}
					};
				}
				else {
					var arrValue = getDataForInputRadio(json[5][j], LUSheet, colWithFindData);
					cell_val = {
						"title": json[6][j],
						"value" : {
							"arrData": arrValue
						}
					};
					var a = arrValue;
				}
			}
			arrJ.push(cell_val);
		}
		jsonArr.push(arrJ);
	}
	return jsonArr;
}



function setDataIfSourceTable(prevTSource, DataItem, colCount, rowCount){

	var start = new Date().getTime();

	var colWithFindData = 1;  //first column with table name which we look for
	var outRow  =  finDataInRange(prevTSource, SetUp, colWithFindData);
	Logger.log('DataItem ' + JSON.stringify(DataItem));

	var row = outRow;
	var firstEmptyRow;
	var countCall = 0;

	//  first empty row
	var rowArrNo = [];
	var outRow2 = outRow;

	var firstEmptyRow = outRow2 = colCount;

	var cellArrNo = [];
	var outCell = 1;

	var lastCol = outCell = rowCount;

	var cell = 4;

	var m = 0;
	var outRow5 = outRow;
	for (var i = (outRow5+7); i < firstEmptyRow; i++) {
		var k = 0;
		var yesNo1 = SetUp.getRange(i, 2).getValue();
		if(yesNo1 === "NO") {
			continue;
		}

		var currentRowData = DataItem[m];

		for (var j = 4; j < outCell; j++) {
			countCall++;
			var yesNo2 = SetUp.getRange(outRow+1, j).getValue();
			if(yesNo2 === "NO") {
				continue;
			}
			k++;
			var arrr = Object.keys(currentRowData).map(function (key) {return currentRowData[key]});
			Logger.log('arr ' + arrr);
			if (arrr[k]){
				SetUp.getRange(i, j).setValue(arrr[k]);
			}
		}
		++m;
	}

	var now = new Date().getTime();
	Logger.log('Works time, ms: ' + (now - start))

}