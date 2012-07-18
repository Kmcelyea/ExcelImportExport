<%@ Page Language="C#" AutoEventWireup="true" Inherits="_Default2" Codebehind="EP.aspx.cs" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Read my Excel</title>
    <link type="text/css" rel="stylesheet" href="~/Content/Site.css" />
    <link type="text/css" rel="stylesheet" href="~/Content/bootstrap.css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
        <div class="hero-unit">
    <h1>View an Excel Document</h1>
        <p></p>
        <br /> 
    <div>
        <br />
        <asp:Label ID="Label1" runat="server" Text="Is there a Header? (only for xlsx)" font-size="15">
        </asp:Label>
        <asp:RadioButtonList ID="rbHDR" runat="server" RepeatDirection="Horizontal" Width="55px">
            <asp:ListItem Text = "Yes" Value = "true" Selected = "True" ></asp:ListItem>
            <asp:ListItem Text = "No" Value = "false"></asp:ListItem>
        </asp:RadioButtonList>
        <br />  

        <asp:FileUpload ID="FileUpload1" runat="server" CssClass="epicbtn" BackColor="White" BorderColor="Black" BorderStyle="Outset"/>
        <asp:Button ID="btnUpload" runat="server" Text="View" OnClick="btnUpload_Click" Height="28px" Width="76px" CssClass="epicbtn" />
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server" GridLines="None" CssClass="mGrid"  PagerStyle-CssClass="pgr"  AlternatingRowStyle-CssClass="alt" EmptyDataText="No Data Available" Font-Size="Medium">
        </asp:GridView>
        
    </div>
            </div>
            </div>
    <p>
        &nbsp;</p>
        
    <p>
        &nbsp;</p>
        
    </form>
</body>
</html>
