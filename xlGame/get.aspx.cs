using System;
using System.Collections.Generic;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string spreadsheet = Request.Form[0];
        string spToken = Request.Form[1];
        string cell = Request.Form[2];
        ISet<String> labels = new HashSet<String>();

        for (int i = 3; i < Request.Form.Count; i++)
        {
            labels.Add(Request.Form[i]);
        }
    }
}