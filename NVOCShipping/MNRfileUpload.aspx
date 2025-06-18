<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MNRfileUpload.aspx.cs" Inherits="NVOCShipping.MNRfileUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
              <asp:FileUpload runat="server" ID="ExcelFileUploading" />
            <asp:Button runat="server" ID="btnfileUploading" Text="File Uploading" OnClick="btnfileUploading_Click" />
            <asp:Label runat="server" ID="lblError" Text="Test"></asp:Label>
        </div>
    </form>
</body>
</html>
