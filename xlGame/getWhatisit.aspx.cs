using System;
using System.Collections.Generic;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    private const String oneDriveFileTokenPrefix = "SD";
    private const String oneDriveFileTokenSuffix = "/517479313637659748/t=0&s=0";
    private const int prefixFieldsNo = 4;

    protected void Page_Load(object sender, EventArgs e)
    {
        Response.ContentType = "application/json";

        if (Request.ContentLength > 0)
        {
            saveData();
        }
        sendNextChallenge();
        sendStatistics();
    }

    private void saveData(){
        string spreadsheet = Request.Form[0];
        string xlsToken = Request.Form[1];
        string skipExpl = "\"" + removeBreakChars(Request.Form[2]) + "\"";
        string userEmail = removeBreakChars(Request.Form[3]);
        string clientIP = getClientIPAddress();
        string linePrefix = DateTime.Now + "\t"
                + clientIP + "\t"
                + "email:" + userEmail + "\t"
                + xlsToken + "\t"
                + spreadsheet + "\t"
                + "skip:" + skipExpl;

        StreamWriter sw = File.AppendText(Request.PhysicalApplicationPath + @"\whatisitData\results.txt");

        if (Request.Form.Count == prefixFieldsNo)
        {
            sw.WriteLine(linePrefix);
        }
        else
        {
            for (int i = prefixFieldsNo; i < Request.Form.Count; i++)
            {
                var commaIndex = Request.Form[i].IndexOf(',');
                sw.WriteLine(linePrefix + "\t"
                    + removeBreakChars(Request.Form[i]).Insert(commaIndex, "\t").Remove(commaIndex + 1, 1));
            }
            updateStatistics(Request.Form.Count - prefixFieldsNo);
        }
        
        sw.Flush();
        sw.Close();
    }

    private void updateStatistics(int newLabels)
    {
        if (newLabels == 0)
        {
            return;
        }

        StreamReader sr = new StreamReader(Request.PhysicalApplicationPath + @"\whatisitData\stats.txt");

        var firstLine = sr.ReadLine();
        var newLine = "";
        var now = DateTime.Now;

        if (now.DayOfYear.ToString() == firstLine.Split('\t')[3])
        {
            var tmp = firstLine.Split('#');
            firstLine = tmp[0] + "#" + (newLabels + Convert.ToInt32(tmp[1]));
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

        StreamWriter sw = new StreamWriter(Request.PhysicalApplicationPath + @"\whatisitData\stats.txt");
        sw.Write(s);

        sw.Flush();
        sw.Close();
    }

    private void sendStatistics()
    {
        var statsDay = 0;
        var statsWeek = 0;
        var statsMonth = 0;
        var statsYear = 0;
        var now = DateTime.Now;

        string[] lineData;
        var stat = 0;
        StreamReader sr = new StreamReader(Request.PhysicalApplicationPath + @"\whatisitData\stats.txt");
        while(!sr.EndOfStream){
            lineData = sr.ReadLine().Split('\t');
            stat = Convert.ToInt32(lineData[4].Remove(0, 1));
            if (now.Year == Convert.ToInt32(lineData[0]))
            {
                statsYear += stat;
                if (now.Month == Convert.ToInt32(lineData[1]))
                {
                    statsMonth += stat;
                    if (System.Threading.Thread.CurrentThread.CurrentCulture.Calendar.GetWeekOfYear(now, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Monday) == Convert.ToInt32(lineData[2]))
                    {
                        statsWeek += stat;
                        if (now.DayOfYear == Convert.ToInt32(lineData[3]))
                        {
                            statsDay += stat;
                        }
                    }
                }
            }
            else
            {
                break;
            }
        }

        sr.Close();

        Response.Write(" \"statsDay\": \"" + statsDay + "\" ,");
        Response.Write(" \"statsWeek\": \"" + statsWeek + "\" ,");
        Response.Write(" \"statsMonth\": \"" + statsMonth + "\" ,");
        Response.Write(" \"statsYear\": \"" + statsYear + "\" }");
    }

    private string removeBreakChars(string s)
    {
        return s.Replace("\t", "").Replace("\n", "").Replace("\r", "").Replace("\r\n", "");
    }

    private void sendNextChallenge()
    {
        var xls = getNextXlsToken();
        Response.Write("{ \"xls\": \"" + xls + "\" ,");
    }

    private String getNextXlsToken()  
    {
        var xlsFiles = File.ReadAllLines(Request.PhysicalApplicationPath + @"\whatisitData\input.txt");
        var randomXls = new Random().Next(0, xlsFiles.Length);
        var line = xlsFiles[randomXls];

        //from line="chris_germany__1813__strg_proxy.xlsx#file.072e74b1abfc5464.72E74B1ABFC5464!197"
        //return spToken = "strg_proxy.xlsx#SD72E74B1ABFC5464!197/517479313637659748/t=0&s=0";
        var enronFileName = line.Substring(line.LastIndexOf("__") + 2).Split('#')[0];
        return enronFileName + "#" + oneDriveFileTokenPrefix + line.Split('#')[1].Split('.')[2] + oneDriveFileTokenSuffix;
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
