using System;
using System.Collections.Generic;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    private const String oneDriveFileTokenPrefix = "SD";
    private const String oneDriveFileTokenSuffix = "/517479313637659748/t=0&s=0";
    private const int prefixFieldsNo = 5;

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
        string cell = Request.Form[2];
        string skipExpl = "\"" + removeBreakChars(Request.Form[3]) + "\"";
        string userEmail = removeBreakChars(Request.Form[4]);
        string clientIP = getClientIPAddress();
        string linePrefix = DateTime.Now + "\t"
                + clientIP + "\t"
                + "email:" + userEmail + "\t"
                + xlsToken + "\t"
                + spreadsheet + "\t"
                + cell + "\t"
                + "skip:" + skipExpl;

        StreamWriter sw = File.AppendText(Request.PhysicalApplicationPath + @"\data\results.txt");

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

        StreamReader sr = new StreamReader(Request.PhysicalApplicationPath + @"\data\stats.txt");

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

        StreamWriter sw = new StreamWriter(Request.PhysicalApplicationPath + @"\data\stats.txt");
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
        StreamReader sr = new StreamReader(Request.PhysicalApplicationPath + @"\data\stats.txt");
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
        var challenge = getNextChallenge();
        Response.Write("{ \"xls\": \"" + challenge.Item1 + "\" ,");
        Response.Write(" \"column\": \"" + challenge.Item2.Item1 + "\" ,");
        Response.Write(" \"row\": \"" + challenge.Item2.Item2 + "\" ,");
    }

    private Tuple<String, Tuple<int, int>> getNextChallenge() 
    {
        var xlsFiles = File.ReadAllLines(Request.PhysicalApplicationPath + @"\data\input.txt");
        var randomXls = new Random().Next(0, xlsFiles.Length);
        var line = xlsFiles[randomXls];
        var split = line.Split('\t');

        return new Tuple<string, Tuple<int, int>>(split[0], ConvertToInt(split[1]));
    }

    static public Tuple<int, int> ConvertToInt(String s)
    {
        int Column;
        int Row;

        int FinalLetter = GetFinalLetter(s);

        int teller = 0;
        //doorloop de letters

        int Total = 0;
        int r = (int)Math.Pow(26, (FinalLetter - 1)); //base 26

        while (teller < FinalLetter)
        {
            Total = Total + r * ((int)s[teller] - 64);
            r = r / 26;
            teller++;
        }

        int TotalChars = Total - 1;

        //doorloop de cijfers
        Total = 0;

        r = (int)Math.Pow(10, (s.Length - FinalLetter - 1)); //base 10

        while (teller < s.Length)
        {
            Total = Total + r * ((int)s[teller] - 48);
            r = r / 10;
            teller++;
        }

        int TotalDigits = Total - 1;

        if (TotalChars < 0)
        {
            throw new ArgumentOutOfRangeException("Column value below zero");
        }
        else
        {
            Column = TotalChars;
        }

        if (TotalDigits < 0)
        {
            throw new ArgumentOutOfRangeException("Row value below zero");
        }
        else
        {
            Row = TotalDigits;
        }

        return new Tuple<int,int>(Column, Row);
    }

    private static int GetFinalLetter(String s)
    {
        char[] cijfers = { '1', '2', '3', '4', '5', '6', '7', '8', '9', '0' };
        var lastChar = s.IndexOfAny(cijfers);
        return lastChar;
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
