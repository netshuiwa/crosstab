using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

// Written by Anurag Gandhi.
// Url: http://www.gandhisoft.com
// Contact me at: soft.gandhi@gmail.com
public partial class _Default : System.Web.UI.Page 
{
    string _Separator = ".";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindGridView();
            GridView gv = (GridView)TabPanel3.FindControl("grdBothPivot");
            MergeRows(gv, 2);
        }
    }

    //Binds all the GridView used in the page.//
    private void BindGridView()
    {
        // Retrieve the data table from Excel Data Source.
        DataTable dt = ExcelLayer.GetDataTable("_Data\\DataForPivot.xls", "Sheet1$");
        /*
         Note:: If you wish to read the data from excel, uncomment the above code and comment the below code.//
         */
        //DataTable dt = SqlLayer.GetDataTable("GetEmployee");
        Pivot pvt = new Pivot(dt);

        grdRawData.DataSource = dt;
        grdRawData.DataBind();

        //Example of Pivot on Both the Axis.//
        grdBothPivot.DataSource = pvt.PivotData(new string[] { "CTC","IsActive" }, AggregateFunction.Sum, new string[] { "Designation", "Year" }, new string[] { "Company", "Department","Name" });
        grdBothPivot.DataBind();
        DataTable dtIndex = new DataTable();
        DataColumn dc;//创建列 
        DataRow dr;       //创建行 
        //构造列 
        for (int i = 0; i < grdBothPivot.Columns.Count; i++)
        {
            dc = new DataColumn();
            dc.ColumnName = grdBothPivot.Columns[i].HeaderText;
            dtIndex.Columns.Add(dc);
        }
        //构造行 
        for (int i = 0; i < grdBothPivot.Rows.Count; i++)
        {
            dr = dtIndex.NewRow();
            for (int j = 0; j < grdBothPivot.Columns.Count; j++)
            {
                dr[j] = grdBothPivot.Rows[i].Cells[j].Text;
            }
            dtIndex.Rows.Add(dr);
        }
    }


    protected void grdPivot2_RowCreated(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
            MergeHeader((GridView)sender, e.Row, 2);
    }

    protected void grdPivot3_RowCreated(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
            MergeHeader((GridView)sender, e.Row, 4);
    }

    /// <summary>
    /// Function used to Create and Merge the Header Cells based on the Pivot conditions.
    /// </summary>
    /// <param name="gv">GridView</param>
    /// <param name="row">Header Row of the GridView</param>
    /// <param name="PivotLevel">The no. of ColumnFields used to Pivot the data</param>
    private void MergeHeader(GridView gv, GridViewRow row, int PivotLevel)
    {
        for (int iCount = 1; iCount <= PivotLevel; iCount++)
        {
            GridViewRow oGridViewRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
            var Header = (row.Cells.Cast<TableCell>()
                .Select(x => GetHeaderText(x.Text, iCount, PivotLevel)))
                .GroupBy(x => x);

            foreach (var v in Header)
            {
                TableHeaderCell cell = new TableHeaderCell();
                cell.Text = v.Key.Substring(v.Key.LastIndexOf(_Separator) + 1);
                cell.ColumnSpan = v.Count();
                oGridViewRow.Cells.Add(cell);
            }
            gv.Controls[0].Controls.AddAt(row.RowIndex, oGridViewRow);
        }
        row.Visible = false;
    }
    private string GetHeaderText(string s, int i, int PivotLevel)
    {
        if (!s.Contains(_Separator) && i != PivotLevel)
            return string.Empty;
        else
        {
            int Index = NthIndexOf(s, _Separator, i);
            if (Index == -1)
                return s;
            return s.Substring(0, Index);
        }
    }

    private void MergeRows(GridView gv, int rowPivotLevel)
    {
        for (int rowIndex = gv.Rows.Count - 2; rowIndex >= 0; rowIndex--)
        {
            GridViewRow row = gv.Rows[rowIndex];
            GridViewRow prevRow = gv.Rows[rowIndex + 1];
            for (int colIndex = 0; colIndex < rowPivotLevel; colIndex++)
            {
                if (row.Cells[colIndex].Text == prevRow.Cells[colIndex].Text)
                {
                    row.Cells[colIndex].RowSpan = (prevRow.Cells[colIndex].RowSpan < 2) ? 2 : prevRow.Cells[colIndex].RowSpan + 1;
                    prevRow.Cells[colIndex].Visible = false;
                }
            }
        }
    }

    /// <summary>
    /// Returns the nth occurance of the SubString from string str
    /// </summary>
    /// <param name="str">source string</param>
    /// <param name="SubString">SubString whose nth occurance to be found</param>
    /// <param name="n">n</param>
    /// <returns>Index of nth occurance of SubString if found else -1</returns>
    private int NthIndexOf(string str, string SubString, int n)
    {
        int x = -1;
        for (int i = 0; i < n; i++)
        {
            x = str.IndexOf(SubString, x + 1);
            if (x == -1)
                return x;
        }
        return x;
    }
}
