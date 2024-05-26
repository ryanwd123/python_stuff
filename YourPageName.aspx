<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="YourPageName.aspx.cs" Inherits="YourNamespace.YourPageName" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Create Text File Example</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button ID="btnCreateFile" runat="server" Text="Create File" OnClick="btnCreateFile_Click" />
        </div>
    </form>
</body>
</html>
