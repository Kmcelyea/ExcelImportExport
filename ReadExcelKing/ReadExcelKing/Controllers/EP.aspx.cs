
using System;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data;
using System.IO;
using System.Configuration;
using OfficeOpenXml;
using NPOI.HSSF.UserModel;
using NPOI.Util;
using NPOI.SS.UserModel;
using NPOI;



 

public partial class _Default2 : System.Web.UI.Page 
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {

        if (FileUpload1.HasFile)
        {
            string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
            string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
            string FilePath = Server.MapPath(FolderPath + FileName);
            FileUpload1.SaveAs(FilePath);

            Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Value);

            FileName = null;
            FilePath = null;
            Extension = null;
            FolderPath = null;
            

        } 
    }
    private void Import_To_Grid(string FilePath, string Extension, string isHDR)
    {
        if (Extension == ".xls")
        {
            //string notxlsx = ("This is not an xlsx file");
            //ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('" + notxlsx + "');", true);

            HSSFWorkbook hssfworkbook;

            using (FileStream file = new FileStream(FilePath, FileMode.Open, FileAccess.Read))

                hssfworkbook = new HSSFWorkbook(file);



            ISheet sheet = hssfworkbook.GetSheetAt(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            DataTable dt = new DataTable();

            //Counts the number of cells in a row and determines the columns from that.
            int counter = sheet.GetRow(0).Cells.Count;
            // J < number of columns needs to be exact at this moment
            for (int j = 0; j < counter; j++)
            {

                // set each column to a - ** letters
                // dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());

                //Get first row and set the headers for each cell
                //dt.Columns.Add(Convert.ToString((string)sheet.GetRow(0).GetCell(+j).StringCellValue).ToString());
                //Get each cell value in row 0 and return its string for a column name.
                dt.Columns.Add(sheet.GetRow(0).GetCell(+j).StringCellValue);
            }

            while (rows.MoveNext())
            {
                HSSFRow row = (HSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);


                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);

            }
            //Hackish way to remove the bad first row made by getting column names
            dt.Rows.RemoveAt(0);
            GridView1.Caption = Path.GetFileName(FilePath);
            GridView1.DataSource = dt;
            //Bind the data
            GridView1.DataBind();
            sheet.Dispose();
            hssfworkbook.Dispose();
            
        }
        else
        {
            //Create a new epplus package using openxml
            var pck = new OfficeOpenXml.ExcelPackage();

            //load the package with the filepath I got from my fileuploader above on the button
            //pck.Load(new System.IO.FileInfo(FilePath).OpenRead());

            //stream the package
            FileStream stream = new FileStream(FilePath, FileMode.Open);
            pck.Load(stream);

            //So.. I am basicly telling it that there is 1 worksheet or to just look at the first one. Not really sure what kind of mayham placing 2 in there would cause.
            //Don't put 0 in the box it will likely cause it to break since it won't have a worksheet page at all.
            var ws = pck.Workbook.Worksheets[1];


            //This will add a sheet1 if your doing pck.workbook.worksheets["Sheet1"];
            if (ws == null)
            {
                ws = pck.Workbook.Worksheets.Add("Sheet1");
                // Obiviously I didn't add anything to the sheet so probably can count on it being blank.
            }

            //I created this datatable for below.
            DataTable tbl = new DataTable();

            //My sad attempt at changing a radio button value into a bool value to check if there is a header on the xlsx
            var hdr = bool.Parse(isHDR);
            Console.WriteLine(hdr);

            //Set the bool value for from above.
            var hasHeader = hdr;

            //Setup the table based on the value from my bool
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            //Bind Data to GridView
            //I have all my info in the tbl dataTable so the datasource for the Gridview1 is set to tbl
            GridView1.Caption = Path.GetFileName(FilePath);
            GridView1.DataSource = tbl;
            //Bind the data
            GridView1.DataBind();

            pck.Save();
            pck.Dispose();
            stream.Close();
            // string pathD = FilePath;
            FilePath = null;
            stream = null;
            // var fileToDelete = new FileInfo(pathD);
            // fileToDelete.Delete();
        }

        }

    protected void PageIndexChanging(object sender, GridViewPageEventArgs e)
    {

        //VERY UNTESTED, worked well with the OleDB dataGridview but now it has a different purpose.


        //This is for handling paging like if you have more than the set amount of rows like 15 at this time I believe then the numbers will appear and this helps map it out.
        string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
        string FileName = GridView1.Caption;
        string Extension = Path.GetExtension(FileName);
        string FilePath = Server.MapPath(FolderPath + FileName);

        //Changed the rbHDR.SelectedItem.Text to Value since I don't want to ask the user true or false just asking them yes or no.
        Import_To_Grid(FilePath, Extension, rbHDR.SelectedItem.Value);
        GridView1.PageIndex = e.NewPageIndex;
        GridView1.DataBind();

    }
}
