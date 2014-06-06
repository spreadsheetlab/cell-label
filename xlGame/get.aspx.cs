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
        string userEmail = removeBreakChars(Request.Form[3]);
        string clientIP = getClientIPAddress();

        StreamWriter sw = File.AppendText(Request.PhysicalApplicationPath + "results.txt");

        for (int i = 4; i < Request.Form.Count; i++)
        {
            var commaIndex = Request.Form[i].IndexOf(',');
            sw.WriteLine(DateTime.Now
                + xlsToken + "\t" 
                + spreadsheet + "\t" 
                + cell + "\t"
                + removeBreakChars(Request.Form[i]).Insert(commaIndex, "\t").Remove(commaIndex + 1, 1) + "\t" 
                + clientIP + "\t" 
                + userEmail);
        }

        saveStatistics(Request.Form.Count - 4);
        
        sw.Flush();
        sw.Close();
    }

    private void saveStatistics(int newLabels)
    {
        StreamReader sr = new StreamReader(Request.PhysicalApplicationPath + "stats.txt");

        var firstLine = sr.ReadLine();
        var newLine = "";
        var now = DateTime.Now;

        if (now.DayOfYear.ToString() == firstLine.Split('\t')[3])
        {
            var tmp = firstLine.Split('#');
            firstLine = tmp[0] + "#" + (Convert.ToInt32(tmp[1]) + newLabels);
        }
        else
        {
            newLine = now.Year.ToString()
                + "\t" + now.Month.ToString()
                + "\t" + System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(now, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday).ToString()
                + "\t" + now.DayOfYear.ToString()
                + "\t#" + newLabels.ToString()
                + Environment.NewLine; 
        }

        String s = newLine + firstLine + Environment.NewLine + sr.ReadToEnd();

        sr.Close();

        StreamWriter sw = new StreamWriter(Request.PhysicalApplicationPath + "stats.txt");

        sw.Write(s);
        sw.Flush();
        sw.Close();
    }

    private string removeBreakChars(string s)
    {
        return s.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace("\r\n", "");
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

    private string getClientIPAddress()
    {
        System.Web.HttpContext context = System.Web.HttpContext.Current;
        string ipAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];

        if (!string.IsNullOrEmpty(ipAddress))
        {
            string[] addresses = ipAddress.Split(',');
            if (addresses.Length != 0)
            {
                return addresses[0];
            }
        }

        return context.Request.ServerVariables["REMOTE_ADDR"];
    }
}
