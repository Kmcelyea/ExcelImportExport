/* EPPlus Importer for Excel into aspx pages
 * Coder: Kyle McElyea
 * Comments By: Kyle McElyea
 * 7-12-2012
 * Using EPPlus to import an excel document
 * 
 * The using OfficeOpenXml is the assembly for EPPlus which is added under references from NuGet packages
 * Code hosted at
 * http://epplus.codeplex.com/
 * 
 * The program uses a fileuploader to gather the path of the local xlsx. I am not supporting xls in this you can use OleDB to do both xlsx and xls.
 * Use the OleDB ACE for 2007+ xlsx
 * Use the OleDB Jet for 2003 xls
 * OleDB has a patch to all it's use on 64bit servers...
 * 
 * Once I have the local file path I feed that information into an OfficeOpenXml ExcelPackage eventually saving a dataTable with the gathered information
 * This dataTable is then fed into a GridView
 * 
 * Currently I have not found a solution for trying to reopen an excel that you already opened after launching the program
 * Ex. I upload a file named Buildings.xlsx and then I decide I want to view it again so I upload Buildings.xlsx again and it will likely through the exception
 * that Buildings.xlsx is already in use. Perhaps if you make changes to the file and reupload it may not do this but in my tests if I double upload a file with
 * no changes it will break
 * 
 * So I need a solution like ending file usage or clearing files after they are not on the grid.
 * 
 * 
 * I'm tired of this doesn't exist in current context stuff its a wordy term that I don't get. Nothing like that in Obj-c
 * 
 * */


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
        } 
    }
    private void Import_To_Grid(string FilePath, string Extension, string isHDR)
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
        //if (ws == null)
        //{
        //    ws = pck.Workbook.Worksheets.Add("Sheet1");
        // Obiviously I didn't add anything to the sheet so probably can count on it being blank.
        //}

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
                pck.Stream.Close();
                pck.Dispose();
               // string pathD = FilePath;
               // FilePath = null;
               // var fileToDelete = new FileInfo(pathD);
              // fileToDelete.Delete();


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
