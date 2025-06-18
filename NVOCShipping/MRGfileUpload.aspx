<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MRGfileUpload.aspx.cs" Inherits="NVOCShipping.MRGfileUpload" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css" type="text/css">
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
</head>
<body style="position: absolute;margin-top: 52px;background-color: azure;">
    <form id="form1" runat="server">
        <div class="container">
            <div class="row">
                <div class="col-md-12 col-xs-12">
                    <div class="panel">
                        <div class="panel-head">
                            <h5>MRG UPLOAD</h5>
                        </div>
                        <div class="panel-body">
                            <div class="row">
                                  <div class="form-group">
                        <div class="row">
                            <div class="col-md-6 col-xs-6">
                                   <asp:FileUpload runat="server" ID="ExcelFileUploading" />
                            </div>
                              <div class="col-md-6 col-xs-6">
                                    <asp:Button runat="server" ID="btnfileUploading" Text="File Uploading" CssClass="btn btn-primary" OnClick="btnfileUploading_Click" />
                            </div>
                        </div>
                      
          
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
