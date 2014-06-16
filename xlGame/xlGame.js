var xlsToken = null; //the SkyDrive filetoken string
var ewa = null; //the Excel Services Web Part
var workbook = null //the active workbook
var questionCell = null; //the cell to be labeled
var labels = []; //list with the labels the user has clicked, each represented as a tuple eg "C3, value"
var remainingCells; //the number of cells to be labeled before changing workbook
var userScore = 0;

if (window.attachEvent) {
    window.attachEvent("onload", loadEwaOnPageLoad);
} else {
    window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
}

function initCollectors() {
    selectQuestionCell();
    labels = [];
    //de-color cells
    workbook.getRangeA1Async("hidden!A2", setHiddenValue, "");
    showSkipText();
}

function loadEwaOnPageLoad() {
    postData({}, true);
}

function btnNextClick() {
    var data = {
        spreadsheet: workbook.getWorkbookPath(),
        xlsToken: xlsToken,
        cell: questionCell.getAddressA1(),
        skipExpl: $("#txtSkip").val(),
        userEmail: $("#userEmail").val(),
        labels: labels
    };

    $("#questionDiv").html("What is...");
    remainingCells--;
    if (remainingCells < 0) {
        //Change worksheet
        postData(data, true);
    } else {
        //Change question cell in the same worksheet
        postData(data, false);
        initCollectors();
    }
}

function postData(data, changeXls) {
    posting = $.post("get.aspx", data);

    posting.done(function (resp) {
        if (changeXls) {
            xlsToken = resp.xls;
            loadExcel(xlsToken);
            remainingCells = Math.floor(Math.random() * 3) + 2; //TODO: set number of tries before worksheet change, now 2 to 2+(3-1) = 4
        }
        $("#stats").html(resp.statsDay + " / " + resp.statsWeek + " / " + resp.statsMonth + " / " + resp.statsYear);
    });
}

function loadExcel(token) {

    var props = {
        uiOptions: {
            showGridlines: true,
            showRowColumnHeaders: true,
            showParametersTaskPane: true
        },
        interactivityOptions: {
            allowTypingAndFormulaEntry: false,
            allowParameterModification: false,
            allowSorting: false,
            allowFiltering: false,
            allowPivotTableInteractivity: false
        }
    };

    $("#excelDiv").remove();
    $("#excelDivParent").append("<div id=\"excelDiv\" style=\"width: 1000px; height: 600px\"></div>");

    Ewa.EwaControl.loadEwaAsync(token, "excelDiv", props, onExcelLoaded);
}

function onExcelLoaded(result) {
    ewa = Ewa.EwaControl.getInstances().getItem(Ewa.EwaControl.getInstances().getCount() - 1);
    workbook = ewa.getActiveWorkbook();
    ewa.add_activeCellChanged(onCellSelectionChange);
    initCollectors();
}

function selectQuestionCell() {
    var qCell = randomCell();
    questionCell = workbook.getRange(workbook.getActiveSheet().getName(), qCell.row, qCell.column, 1, 1);
    questionCell.getValuesAsync(1, checkNonEmpty, null);
}

function checkNonEmpty(asyncResult) {
    var value = asyncResult.getReturnValue()[0][0];
    if (value == "") {
        //select another cell
        selectQuestionCell();
    } else {
        //set value to hidden!A1 to make confitional formatting color the ewaCell
        workbook.getRangeA1Async("hidden!A1", setHiddenValue, questionCell.getAddressA1() + ",");
        $("#questionDiv").html("What is <div class=\"questionCell\">" + asyncResult.getReturnValue()[0][0] + "</div> in " + questionCell.getAddressA1().split("!")[1] + "?");
    }
}

// Handle the active cell changed event.
function onCellSelectionChange(rangeArgs) {
    var cell = workbook.getActiveCell();
    cell.getValuesAsync(1, updateLabels, cell);
}

function updateLabels(asyncResult) {
    var cellValue = asyncResult.getReturnValue()[0][0];
    if (cellValue == "") {
        return;
    }

    var cellAddr = asyncResult.getUserContext().getAddressA1().split("!")[1];
    var index = indexOfLabel(cellAddr);
    if (index >= 0) { //Cell already in list, de-selected -> remove from labels
        labels.splice(index, 1);
        userScore--;
    }
    else { //Cell selected -> add to labels
        labels[labels.length] = [cellAddr, "\"" + cellValue + "\""];
        userScore++;
    }
    updateLabelColors();
    $("#noOfLabels").html(userScore);

    labels.length > 0 ? hideSkipText() : showSkipText();
}

function showSkipText() {
    $("#results").html("");
    $("#txtSkip").val("");
    $("#txtSkip").show();
    $("#lblSkip").text("Select cells that apply, or skip challenge because ");
}

function hideSkipText() {
    $("#results").html("It is <div class=\"labelCell\">" + getLabelsStr(1, "</div> <div class=\"labelCell\">") + "</div>");
    $("#txtSkip").hide();
    $("#lblSkip").text("");
}

function indexOfLabel(cellAddr) {
    for (var i = 0; i < labels.length; i++) {
        if (labels[i][0] == cellAddr) {
            return (i);
        }
    }
    return (-1);
}

function updateLabelColors() {
    workbook.getRangeA1Async("hidden!A2", setHiddenValue, getLabelsStr(0, ","));
}

function setHiddenValue(asyncResult) {
    asyncResult.getReturnValue().setValuesAsync([[asyncResult.getUserContext()]], setRangeValues, null);
}

function setRangeValues(asyncResult) {
}

function getLabelsStr(index, delimeter) {
    var labelsList = "";
    for (var i = 0; i < labels.length; i++) {
        labelsList += labels[i][index] + delimeter;
    };
    return labelsList;
}

function randomCell() {
    var r = Math.floor(Math.random() * 10); //TODO: set constant to max number of visible rows
    var c = Math.floor(Math.random() * 15); //TODO: set constant to max number of visible columns
    return {
        row: r,
        column: c
    };
}