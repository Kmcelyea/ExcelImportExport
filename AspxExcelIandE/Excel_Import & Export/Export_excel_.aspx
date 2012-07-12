<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Export_excel_.aspx.cs"
    Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Import Excel Data into GridView & Again Export to Excel Sheet</title>
    <style type="text/css">
        .hdr
        {
            background: #ccc;
            font-family: Arial;
            font-size: 12px;
            color: #fff;
        }
        .ftr
        {
            background: #000;
            color: #fff;
            font-family: Arial;
            font-size: 12px;
            color: #fff;
        }
        .hdr1
        {
            background: #FF0000;
            font-family: Arial;
            font-size: 12px;
            font-weight: bold;
            color: #fff;
        }
        .ftr1
        {
            background: #000;
            color: #fff;
            font-family: Arial;
            font-size: 12px;
            color: #fff;
        }
        .Row
        {
            background: #ccc;
            text-align: center;
            font-size: 12px;
            color: #000;
        }
        .Alt
        {
            background: ##808000;
            text-align: center;
            font-size: 12px;
            color: #000;
        }
    </style>

    <script language="javascript" type="text/javascript">
        function exportToExcel() {
            var oExcel = new ActiveXObject("Excel.Application");
            var oBook = oExcel.Workbooks.Add;
            var oSheet = oBook.Worksheets(1);
            var dt = document.getElementById('tbl')
            for (var y = 0; y < dt.rows.length; y++)
            // detailsTable is the table where the content to be exported is
            {
                for (var x = 0; x < dt.rows(y).cells.length; x++) 
                {
                    oSheet.Cells(y + 1, x + 1) = dt.rows(y).cells(x).innerText;
                }
            }
            oExcel.Visible = true;
            oExcel.UserControl = true;
        }
    </script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table cellpadding="10" cellspacing="10" style="font-family: Arial; font-size: 12px;
            border: solid 1px #ccc;" border="1" align="center">
            <tr>
                <td align="center">
                    <asp:FileUpload ID="FileUpload1" runat="server" />
                </td>
                <td>
                    <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="false" OnPageIndexChanging="PageIndexChanging">
                        <HeaderStyle CssClass="hdr" />
                        <FooterStyle CssClass="ftr" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <asp:GridView ID="GridView2" runat="server" AllowPaging="false">
                        <HeaderStyle CssClass="hdr1" />
                        <RowStyle CssClass="Row" />
                        <AlternatingRowStyle CssClass="Alt" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <asp:Button ID="Save_ExporttoExcel" runat="server" Text="Save to DB & Export to Excel"
                        OnClientClick="exportToExcel()" OnClick="Save_ExporttoExcel_Click" />
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="lblError" runat="server" Text="Label"></asp:Label>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
