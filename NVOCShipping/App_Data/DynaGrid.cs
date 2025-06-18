using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Summary description for GridData
/// 
/// DataFields will be created as Field1, Field2, etc.,
/// Dummy Column Name ID, for an empty Record
/// 
/// In the Gridview Design, 
/// 
/// For the Footer Row - > TextBox Ctrls Named as InsField0, InsField1, and so on...
///                         DropDowns Named as ddlInsVal0 along with a corresponding invisible TextBox
/// 
/// For the Data Row - > TextBox Ctrls Named as txtVal0, txtVal1, and so on...
///                     Each TextBox ctrl has an invisible label control named as -> lblField0,
///                     Dropdowns will have an invisible textbox named as 
///                                 ddlVal0, and txtVal0 respectively
/// 
/// </summary>
public class DynaGrid
{
    public DynaGrid()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public DataTable FillGrid(int ColCount)
    {
        DataTable dt=new DataTable();
        dt.Columns.Add("ID", typeof(String));
        for (int i = 1; i <= ColCount; i++)
            dt.Columns.Add("Field" + i.ToString(), typeof(String));
        dt.Rows.Add(dt.NewRow());
        dt.Rows[0]["ID"] = "Empty";        
        return dt;
    }

    public DataTable FillGridEmpty(int ColCount)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("ID", typeof(String));
        for (int i = 1; i <= ColCount; i++)
            dt.Columns.Add("Field" + i.ToString(), typeof(String));
        //dt.Rows.Add(dt.NewRow());
        //dt.Rows[0]["ID"] = "Empty";
        return dt;
    }

    public DataTable AddRec(DataTable srcTbl, DataTable DestTbl)
    {
        bool IDexists = false;
        if (srcTbl.Columns.Contains("ID"))
        { IDexists = true; }
        DestTbl.Rows.Clear();
        DestTbl.Rows.Add(DestTbl.NewRow());
        DestTbl.Rows[0]["ID"] = srcTbl.Rows.Count.ToString();
        for (int i = 0; i <= srcTbl.Rows.Count - 1; i++)
        {
            DestTbl.Rows.Add(DestTbl.NewRow());
            if (IDexists)
            {
                DestTbl.Rows[i + 1]["ID"] = srcTbl.Rows[i]["ID"].ToString();
                for (int j = 1; j <= srcTbl.Columns.Count - 1; j++)
                    DestTbl.Rows[i + 1][j] = srcTbl.Rows[i][j].ToString();
            }
            else
            {
                for (int j = 0; j <= srcTbl.Columns.Count - 1; j++)
                    DestTbl.Rows[i + 1][j + 1] = srcTbl.Rows[i][j].ToString();
            }
        }
        return DestTbl;
    }
    public DataTable AddRec(GridView gv, DataTable dt)
    {
        dt = UpdTable(gv, dt);
        dt.Rows.Add(dt.NewRow());
        for (int i = 0; i <= dt.Columns.Count - 2; i++)
        {
            DropDownList ddl = (DropDownList)gv.FooterRow.FindControl("ddlInsVal" + i.ToString());
            TextBox tb = (TextBox)gv.FooterRow.FindControl("InsField" + i.ToString());
            CheckBox cb = (CheckBox)gv.FooterRow.FindControl("ChkInsField" + i.ToString());
            if (ddl != null)
                dt.Rows[dt.Rows.Count - 1][i + 1] = ddl.Text;
            if (tb != null)
                dt.Rows[dt.Rows.Count - 1][i + 1] = tb.Text;
            if (cb != null)
                dt.Rows[dt.Rows.Count - 1][i + 1] = cb.Checked.ToString();
        }
        TextBox txtID = (TextBox)gv.FooterRow.FindControl("InsID");
        if (txtID != null)
            dt.Rows[dt.Rows.Count - 1]["ID"] = txtID.Text;
        return dt;
    }
    public DataTable DelRec(string RecNo,GridView gv,DataTable dt)
    {
        UpdTable(gv, dt);
        dt.Rows.RemoveAt(int.Parse(RecNo));
        return dt;
    }

    public DataTable UpdTable(GridView gv, DataTable dt)
    {
        if (gv.Rows.Count > 1 && dt.Rows.Count > 1 && gv.Rows.Count == dt.Rows.Count)
        {
            for (int dr = 1; dr <= dt.Rows.Count - 1; dr++)
            {
                for (int dc = 0; dc <= dt.Columns.Count - 1; dc++)
                {
                    TextBox tb = gv.Rows[dr].FindControl("txtVal" + dc.ToString()) as TextBox;
                    DropDownList ddl = gv.Rows[dr].FindControl("ddlVal" + dc.ToString()) as DropDownList;
                    CheckBox cb = (CheckBox)gv.Rows[dr].FindControl("ChkField" + dc.ToString());
                    if (tb != null)
                        dt.Rows[dr]["Field" + (dc + 1).ToString()] = tb.Text;
                    if (ddl != null)
                        dt.Rows[dr]["Field" + (dc + 1).ToString()] = ddl.Text;
                    if (cb != null)
                        dt.Rows[dr]["Field" + (dc + 1).ToString()] = cb.Checked.ToString();
                }
            }
        }

        return dt;
    }

    //Added on June3 start
    /// <summary>
    /// Naming ddl1, ddl2, for dropdowns and txt1, txt2, for textbox, lblID for ID field as label box
    /// </summary>
    /// <param name="GridView1"></param>
    /// <param name="dt"></param>
    /// <returns></returns>

    public DataTable InitGrid(int ColCount)
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("ID", typeof(String));
        for (int i = 1; i <= ColCount; i++)
            dt.Columns.Add("Field" + i.ToString(), typeof(String));
        dt.Rows.Add(dt.NewRow());
        dt.Rows[0]["ID"] = "";
        return dt;
    }

    public DataTable UpdateTable(GridView GridView1, DataTable dt)
    {
        dt.Rows.Clear();
        int ColumnCount = GridView1.Columns.Count;
        foreach (GridViewRow gRow in GridView1.Rows)
        {
            dt.Rows.Add(dt.NewRow());
            Label lblID = (Label)gRow.FindControl("lblID");
            if (lblID != null)
                dt.Rows[dt.Rows.Count - 1]["ID"] = lblID.Text;
            for (int Col = 1; Col < ColumnCount; Col++)
            {
                TextBox txt = (TextBox)gRow.FindControl("txt" + Col.ToString());
                DropDownList ddl = (DropDownList)gRow.FindControl("ddl" + Col.ToString());
                if (txt != null)
                    dt.Rows[dt.Rows.Count - 1][Col] = txt.Text;
                if (ddl != null)
                    dt.Rows[dt.Rows.Count - 1][Col] = ddl.SelectedValue;
            }
        }
        return dt;

    }
}
