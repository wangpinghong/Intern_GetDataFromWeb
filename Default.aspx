<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <asp:Label ID="Label1" runat="server" Text="網頁位置"></asp:Label>
    <br />
    <asp:ListBox ID="ListBox1" runat="server" Height="203px" Width="223px">
        <asp:ListItem></asp:ListItem>
    </asp:ListBox>
    <br />
    <asp:Label ID="Label2" runat="server" Text="loading"></asp:Label>
    <asp:Button ID="Button3" runat="server" Text="Button" />
    <p>
    <asp:Button ID="Button1" runat="server" Text="全部輸出成txt" />
        <asp:Button ID="Button2" runat="server" Text="選取輸出成txt" />
    </p>
    </form>
</body>
</html>
