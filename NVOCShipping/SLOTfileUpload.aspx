<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="SLOTfileUpload.aspx.cs" Inherits="NVOCShipping.SLOTfileUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <link rel="stylesheet" href="assets/css/bootstrap.min.css" />
     <style>
        .lblmove {
            margin-top: 200px;
            margin-left: 250px;
        }

        .lblMovSucess {
            margin-top: 300px;
             text-align:center;
        }

        .lblerrSuc {
            margin-top: 50px;
            text-align:center;
        }
    </style>
</head>
<body style="background-color:#fff;">
    <form id="form1" runat="server">
         <div class="container">
            <div class="row">
                <div class="col-md-6 col-md-offset-3" style="padding-left:0;padding-right:0;">
                    <div class="panel">
                        <div class="panel-heading" style="background-color:#21277a;text-align:center;color:#fff;">
                            <h4 style="margin-top:0;margin-bottom:0;">Slot Upload</h4>
                        </div>
                        <div class="panel-body">
                            <div class="row">
                                <div class="col-md-6">
                                    <asp:FileUpload runat="server" ID="ExcelFileUploading" />
                                    </div>
                                <div class="col-md-6" style="text-align:right;">
                                    <asp:Button runat="server" ID="btnfileUploading" Text="MOVE TO DATA" OnClick="btnfileUploading_Click" />
                                 </div>
                                <div class="col-md-12" style="text-align:center; margin-top:20px;">
                                     <asp:Label runat="server" ID="lblError" Text="Test"></asp:Label>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
             </div>
      </div>
    </form>
</body>
</html>
