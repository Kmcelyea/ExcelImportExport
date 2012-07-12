//|-------------------------------------------------------|
//|  Author		: Poonam Verma                            |
//|-------------------------------------------------------|


using System;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Configuration;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
       lblError.Visible = false;
    }
    // Upload excel files.
    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (FileUpload1.HasFile)
        {
            string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
            string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
            string FilePath = Server.MapPath(FolderPath + FileName);
            FileUpload1.SaveAs(FilePath);
            Import_To_Grid(FilePath, Extension);
        }
    }
    

    // Bind with Grid
 private void Import_To_Grid(string FilePath, string Extension)
    {
        string conStr = "";
        switch (Extension)
        {
            case ".xls": //Excel 97-03
                conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                break;
            case ".xlsx": //Excel 07
                conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                break;
        }
        conStr = String.Format(conStr, FilePath, 1);
        OleDbConnection connExcel = new OleDbConnection(conStr);
        OleDbCommand cmdExcel = new OleDbCommand();
        OleDbDataAdapter oda = new OleDbDataAdapter();
        DataTable dt = new DataTable();
        cmdExcel.Connection = connExcel;
 
        connExcel.Open();
        DataTable dtExcelSchema;
        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
        connExcel.Close();

        //Read Data from First Sheet
        connExcel.Open();
        cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
        oda.SelectCommand = cmdExcel;
        oda.Fill(dt);
        connExcel.Close();
        GridView1.DataSource = dt;
        GridView1.DataBind();

       // Creat Temproary table 
          DataTable dttable_data = new DataTable();
          dttable_data = dt; //Return Table consisting data

        //Create Tempory Table
        DataTable dtTemp = new DataTable();

        // Creating Header Row
        dtTemp.Columns.Add(" S. No.");
        dtTemp.Columns.Add(" Student Name ");
        dtTemp.Columns.Add(" Roll No ");
        dtTemp.Columns.Add(" Total Marks ");
        dtTemp.Columns.Add(" Percentage");
    
       
        
        int sum;
        DataRow drAddItem;
        for (int i = 0; i < dttable_data.Rows.Count; i++)
         {
          drAddItem = dtTemp.NewRow();
          drAddItem[0] = dttable_data.Rows[i]["S# No#"].ToString();
          drAddItem[1] = dttable_data.Rows[i]["Student Name"].ToString();//Student Name
          drAddItem[2] = dttable_data.Rows[i]["Roll No#"].ToString();//Roll No
          
            //Sum
          
            sum = (int.Parse(dttable_data.Rows[i]["Hindi"].ToString()) + int.Parse(dttable_data.Rows[i]["English"].ToString()) + int.Parse(dttable_data.Rows[i]["Maths"].ToString()) + int.Parse(dttable_data.Rows[i]["Physics"].ToString()));
            drAddItem[3] = sum.ToString();
         
            //%age
            int prcnt = (sum*100/800);
            drAddItem[4] = prcnt.ToString();
            dtTemp.Rows.Add(drAddItem);
          }

       //Bind Data with Grid View
        GridView2.DataSource = dtTemp;
        GridView2.DataBind();
      
      }

    // Page Index Changing
    protected void PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
        string FileName = GridView1.Caption;
        string Extension = Path.GetExtension(FileName);
        string FilePath = Server.MapPath(FolderPath + FileName);

        Import_To_Grid(FilePath, Extension);
        GridView1.PageIndex = e.NewPageIndex;
        GridView1.DataBind();
    }

 
    // Write in Excel Sheet
    private void creatExcel()
    {
        if (Int32.Parse(GridView2.Rows.Count.ToString()) < 65536)
        {
            GridView2.AllowPaging = true;
            //grvProdReport.DataBind()
            StringWriter tw = new StringWriter();
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(tw);
            HtmlForm frm = new HtmlForm();

            string strTmpTime = (System.DateTime.Today).ToString();
            if (strTmpTime.IndexOf("/") != -1)
            {
                strTmpTime = strTmpTime.Replace("/", "-").ToString().Trim();
            }
            if (strTmpTime.IndexOf(":") != -1)
            {
                strTmpTime = strTmpTime.Replace(":", "-").ToString().Trim();
            }

            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;filename=sheet.xls");
            Response.Charset = "UTF-8";
            EnableViewState = false;
            Controls.Add(frm);
            frm.Controls.Add(GridView2);
            frm.RenderControl(hw);
            hw.WriteLine("<b> <u> <font-size:'5'> Student Report </font> </u> </b>");
            Response.Write(tw.ToString());
            Response.End();
        }
        //grvProdReport.AllowPaging = "True"
        //grvProdReport.DataBind()
        else
        {
            lblError.Visible = true;
            lblError.Text = "Export to Excel not possible";
        }
    }
    // insert value in database
    private void inst_data()
    {
        int i;
        string s = ConfigurationManager.ConnectionStrings["test_"].ConnectionString;
        SqlConnection con = new SqlConnection(s);
       
     for (i = 0; i <= GridView1.Rows.Count - 1; i++)
        {
            string query = "insert into tbl_studentsummary values ('" + GridView2.Rows[i].Cells[0].Text + "','" + GridView2.Rows[i].Cells[1].Text.ToString() + "','" +  GridView2.Rows[i].Cells[2].Text  + "','" +  GridView2.Rows[i].Cells[3].Text  + "','" +  GridView2.Rows[i].Cells[4].Text + "')";
            SqlCommand cmd = new SqlCommand(query,con);
            con.Open();    
            cmd.ExecuteNonQuery();
            con.Close();
            lblError.Visible = true;
            lblError.Text = "Sucessfull";
        }
    }

    

    protected void Save_ExporttoExcel_Click(object sender, EventArgs e)
    {
         inst_data(); // insert data in database.
         creatExcel(); // Creat Excel  File
    }
}
     

 
