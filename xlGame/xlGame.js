var xlsToken = null; //the SkyDrive filetoken string
var ewa = null; //the Excel Services Web Part
var workbook = null //the active workbook
var questionCell = null; //the cell to be labeled
var labels = []; //list with the labels the user has clicked, each represented as a tuple eg "C3, value"
var remainingCells; //the number of cells to be labeled before changing workbook
var useremail = "";
var keepUpdated = false;
var userScore = 0;
var smileyScore = 0;

if (window.attachEvent) {
    window.attachEvent("onload", loadEwaOnPageLoad);
} else {
    window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
}

$('#welcomeModal').on('hidden.bs.modal', function (e) {
    if ($("#CheckKeep").is(':checked')) {
        useremail = $("#InputEmail").val();
        keepUpdated = $("#CheckUpdated").is(':checked');
        $("#scoreLabel").html((useremail.length > 0) ? useremail.split('@')[0] + "'s score" : "your score");
    }
})

$('.carousel').carousel({
    interval: 20000
})

$('#welcomeModal').modal({
    keyboard: false,
    backdrop: 'static'
})

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
    updateSmiley();

    var data = {
        spreadsheet: workbook.getWorkbookPath(),
        xlsToken: xlsToken,
        cell: questionCell.getAddressA1(),
        skipExpl: $("#txtSkip").val(),
        userEmail: useremail + "#" + keepUpdated,
        labels: labels
    };

    $("#questionHeader").html("What describes&hellip;");
    $("#questionExp").hide();
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

function updateSmiley() {
    (labels.length > 0 || $("#txtSkip").val().length > 0) ? smileyScore++ : smileyScore--;

    var evenScore = Math.ceil(smileyScore / 2.0) * 2;
    if (evenScore >= -6 && evenScore <= 12) {
        $("#smiley").attr("src", "images/smileys/smiley_" + evenScore + ".jpg");
    }
}

function postData(data, changeXls) {
    posting = $.post("get.aspx", data);

    posting.done(function (resp) {
        if (changeXls) {
            xlsToken = resp.xls;
            loadExcel(xlsToken);
            remainingCells = Math.floor(Math.random() * 3) + 2; //number of tries before worksheet change, 2 to 2+(3-1) = 4
        }
        $("#statsDay").html(resp.statsDay);
        $("#statsWeek").html(resp.statsWeek);
        $("#statsMonth").html(resp.statsMonth);
        $("#statsYear").html(resp.statsYear);
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
        var qCellContent = asyncResult.getReturnValue()[0][0];
        if (qCellContent.length > 35) {
            qCellContent = qCellContent.substring(0, 30) + "&hellip;";
        }
        workbook.getRangeA1Async("hidden!A1", setHiddenValue, questionCell.getAddressA1() + ",");
        $("#questionHeader").html("What describes <span class=\"label label-warning\">" + qCellContent + "</span> in " + questionCell.getAddressA1().split("!")[1] + "?");
        $("#questionExp").show();
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
    $("#lblSkip").text("Skip challenge because ");
}

function hideSkipText() {
    $("#results").html("It is <span class=\"label label-success\">" + getLabelsStr(1, "</span> <span class=\"label label-success\">") + "</span>");
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
    var r = Math.floor(Math.random() * 30);
    var c = Math.floor(Math.random() * 15);
    return {
        row: r,
        column: c
    };
}