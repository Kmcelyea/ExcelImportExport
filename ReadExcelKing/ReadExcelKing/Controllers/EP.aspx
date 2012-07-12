<%@ Page Language="C#" AutoEventWireup="true"  CodeFile="EP.aspx.cs" Inherits="_Default2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Read my Excel</title>
    <link type="text/css" rel="stylesheet" href="~/Content/Site.css" />
</head>
<body>
    <form id="form1" runat="server">
    <center><h1>Upload an Excel Document</h1></center>
        <p></p>
        <br /> 
    <div><center>
        <br />
        <asp:Label ID="Label1" runat="server" Text="Is there a Header?" fontsize="15">
        </asp:Label>
        <asp:RadioButtonList ID="rbHDR" runat="server" RepeatDirection="Horizontal" Width="55px">
            <asp:ListItem Text = "Yes" Value = "true" Selected = "True" ></asp:ListItem>
            <asp:ListItem Text = "No" Value = "false"></asp:ListItem>
        </asp:RadioButtonList>
        <br />  

        <asp:FileUpload ID="FileUpload1" runat="server" CssClass="btn active" BackColor="White" BorderColor="Black" BorderStyle="Outset"/>
        <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" Height="28px" Width="76px" />
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server" GridLines="None" OnPageIndexChanging = "PageIndexChanging" AllowPaging="true" PageSize="15" CssClass="mGrid"  PagerStyle-CssClass="pgr"  AlternatingRowStyle-CssClass="alt" EmptyDataText="No Data Available" Font-Size="Medium">
        </asp:GridView>
        </center>
    </div>
    <p>
        &nbsp;</p>
        <center>
    <p>
        &nbsp;</p>
        </center>
    </form>
</body>
</html>
