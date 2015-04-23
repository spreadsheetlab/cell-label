var xlsToken = null; //the SkyDrive filetoken string
var ewa = null; //the Excel Services Web Part
var workbook = null //the active workbook
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
    $("#result").val('-1');
    resultChanged();
    $("#questionHeader").html("Of what type is <span class=\"label label-warning\">" + xlsToken.split("#")[0] + "</span>?");
    $("#questionExp").show();
}

function loadEwaOnPageLoad() {
    postData({});
}

function btnNextClick() {
    updateSmiley();

    var data = {
        spreadsheet: workbook.getWorkbookPath(),
        xlsToken: xlsToken,
        skipExpl: $("#txtOther").val(),
        userEmail: useremail + "#" + keepUpdated,
        labels: $('#result').val()
    };

    $("#questionHeader").html("Loading&hellip;");
    $("#questionExp").hide();
    //Change worksheet
    postData(data);
}

function updateSmiley() {
    if ($('#result').val() == null || $('#result').val() == '-1' || ($('#result').val() == "other" && $("#txtOther").val().length == 0)) {
        smileyScore--;
    }
    else {
        smileyScore++;
        userScore++;
    }

    var evenScore = Math.ceil(smileyScore / 2.0) * 2;
    if (evenScore >= -6 && evenScore <= 12) {
        $("#smiley").attr("src", "images/smileys/smiley_" + evenScore + ".jpg");
    }

    $("#noOfLabels").html(userScore);
}

function postData(data) {
    posting = $.post("getWhatIsIt.aspx", data);

    posting.done(function (resp) {
        xlsToken = resp.xls;
        loadExcel(xlsToken);
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
            allowHyperlinkNavigation: false
        }
    };

    $("#excelDiv").remove();
    $("#excelDivParent").append("<div id=\"excelDiv\" style=\"width: 1000px; height: 600px\"></div>");

    Ewa.EwaControl.loadEwaAsync(token.split("#")[1], "excelDiv", props, onExcelLoaded);
}

function onExcelLoaded(result) {
    ewa = Ewa.EwaControl.getInstances().getItem(Ewa.EwaControl.getInstances().getCount() - 1);
    workbook = ewa.getActiveWorkbook();
    initCollectors();
}

function resultChanged() {
    $("#txtOther").val("");
    ($('#result').val() == "other") ? $("#txtOther").show() : $("#txtOther").hide();
}