<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:Label ID="Label1" runat="server" Text="網頁網址"></asp:Label>
    <asp:TextBox ID="TextBox1" runat="server" Height="16px" Width="416px"></asp:TextBox>
    <br />
    <asp:Button ID="Button1" runat="server" Text="輸出成txt" />
    <asp:Button ID="Button2" runat="server" Text="輸出成excel" />
    <br />
    </form>
</body>
</html>
