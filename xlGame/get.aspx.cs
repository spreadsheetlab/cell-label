using System;
using System.Collections.Generic;
using System.IO;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string spreadsheet = Request.Form[0];
        string spToken = Request.Form[1];
        string cell = Request.Form[2];

        StreamWriter sw = File.AppendText(Request.PhysicalApplicationPath + "results.txt");

        for (int i = 3; i < Request.Form.Count; i++)
        {
            var commaIndex = Request.Form[i].IndexOf(',');
            sw.WriteLine(spToken + "\t" + spreadsheet + "\t" + cell + "\t" + Request.Form[i].Insert(commaIndex, "\t").Remove(commaIndex + 1, 1));
        }
        
        sw.Flush();
        sw.Close();
    }
}