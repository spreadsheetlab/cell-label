using System;
using System.Collections.Generic;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    private const String oneDriveFileTokenPrefix = "SD";
    private const String oneDriveFileTokenSuffix = "/517479313637659748/t=0&s=0";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.ContentLength > 0)
        {
            saveData();
        }
        sendNextChallenge();
    }

    private void saveData(){
        string spreadsheet = Request.Form[0];
        string xlsToken = Request.Form[1];
        string cell = Request.Form[2];

        StreamWriter sw = File.AppendText(Request.PhysicalApplicationPath + "results.txt");

        for (int i = 3; i < Request.Form.Count; i++)
        {
            var commaIndex = Request.Form[i].IndexOf(',');
            sw.WriteLine(xlsToken + "\t" + spreadsheet + "\t" + cell + "\t" + Request.Form[i].Insert(commaIndex, "\t").Remove(commaIndex + 1, 1));
        }
        
        sw.Flush();
        sw.Close();
    }

    private void sendNextChallenge()
    {
        var xls = getNextXlsToken();
        Response.ContentType = "application/json";
        Response.Write("{ \"xls\": \"" + xls + "\" }");
    }

    private String getNextXlsToken() 
    {
        var xlsFiles = File.ReadAllLines(Request.PhysicalApplicationPath + "input.txt");
        var randomXls = new Random().Next(0, xlsFiles.Length);
        var line = xlsFiles[randomXls];

        //from line="myTest.xlsx#file.072e74b1abfc5464.72E74B1ABFC5464!197"
        //return spToken = "SD72E74B1ABFC5464!197/517479313637659748/t=0&s=0";
        return oneDriveFileTokenPrefix + line.Split('#')[1].Split('.')[2] + oneDriveFileTokenSuffix;
    }
}