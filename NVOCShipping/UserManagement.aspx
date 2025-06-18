<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserManagement.aspx.cs" Inherits="NVOCShipping.UserManagement" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link rel="stylesheet" href="~/assets/menu/bootstrap.min.css" />
     <link href="~/assets/css/responsive.css" rel="stylesheet" />
   <%-- <link href="~/assets/css/custom.css" rel="stylesheet" />--%>
    <link rel="stylesheet" href="~/assets/menu/_all-skins.min.css.css" />
    <link rel="stylesheet" href="~/assets/menu/AdminLTE.min.css" />
    <link rel="stylesheet" href="~/assets/menu/common.css" />
    <link href="https://fonts.googleapis.com/css?family=Rubik" rel="stylesheet"/>
    <link href="https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css" rel="stylesheet"/>
    <link href="https://fonts.googleapis.com/css?family=Open+Sans" rel="stylesheet"/>
     <link href="https://cdn.jsdelivr.net/npm/pretty-checkbox@3.0/dist/pretty-checkbox.min.css" rel="stylesheet"/>
     <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
   <script type="text/javascript" src="http://ajax.cdnjs.com/ajax/libs/json2/20110223/json2.js"></script>
     <script type='text/javascript' src='http://ajax.googleapis.com/ajax/libs/jquery/1.4/jquery.min.js'></script>
     <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
       <style type="text/css">
                .modal {
                    position: fixed;
                    top: 0;
                    left: 0;
                    background-color: black;
                    z-index: 99;
                    opacity: 0.8;
                    filter: alpha(opacity=80);
                    -moz-opacity: 0.8;
                    min-height: 100%;
                    width: 100%;
                }

                .loading {
                    font-family: Arial;
                    font-size: 10pt;
                    border: 5px solid #67CFF5;
                    width: 200px;
                    height: 100px;
                    display: none;
                    position: fixed;
                    background-color: White;
                    z-index: 999;
                }


                .lable-fond {
                    font-size: 11px;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    text-align: right;
                    width: 130px;
                }

                .disabledStyle {
                    background-color: #EBEBE4;
                }

                .EnableStyle {
                    background-color: white;
                }


                #spinner {
                    float: left;
                    width: 100%;
                }

                    #spinner img {
                        top: 0;
                        bottom: 0;
                        left: 0;
                        right: 0;
                        margin: 0 auto;
                        position: absolute;
                    }

                .ui-state-default, .ui-widget-content .ui-state-default, .ui-widget-header
                .ui-state-default {
                    background: #69B8D6 !important;
                    font-family: 'Oswald', sans-serif;
                }

                .ui-widget-header {
                    background: #69B8D6 !important;
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                }


                .checkbox td, th {
                    padding: 8px 0px;
                }

                input[type=checkbox] {
                    margin: 4px 10px;
                }

                .checkbox {
                    padding-left: 20px;
                }

                    .checkbox label {
                        display: inline-block;
                        position: relative;
                        padding-left: 46px;
                    }

                        .checkbox label::before {
                            content: "";
                            display: inline-block;
                            position: absolute;
                            width: 17px;
                            height: 17px;
                            left: 0;
                            margin-left: 20px;
                            border: 1px solid #cccccc;
                            border-radius: 3px;
                            background-color: #fff;
                            -webkit-transition: border 0.15s ease-in-out, color 0.15s ease-in-out;
                            -o-transition: border 0.15s ease-in-out, color 0.15s ease-in-out;
                            transition: border 0.15s ease-in-out, color 0.15s ease-in-out;
                        }

                        .checkbox label::after {
                            display: inline-block;
                            position: absolute;
                            width: 16px;
                            height: 16px;
                            left: 0;
                            top: 0;
                            margin-left: 20px;
                            padding-left: 3px;
                            padding-top: 1px;
                            font-size: 11px;
                            color: #555555;
                        }

                    .checkbox input[type="checkbox"] {
                        opacity: 0;
                    }

                        .checkbox input[type="checkbox"]:focus + label::before {
                            outline: thin dotted;
                            outline: 5px auto -webkit-focus-ring-color;
                            outline-offset: -2px;
                        }

                        .checkbox input[type="checkbox"]:checked + label::after {
                            font-family: 'FontAwesome';
                            content: "\f00c";
                        }

                        .checkbox input[type="checkbox"]:disabled + label {
                            opacity: 0.65;
                        }

                            .checkbox input[type="checkbox"]:disabled + label::before {
                                background-color: #eeeeee;
                                cursor: not-allowed;
                            }

                    .checkbox.checkbox-circle label::before {
                        border-radius: 50%;
                    }

                    .checkbox.checkbox-inline {
                        margin-top: 0;
                    }

                .checkbox-primary input[type="checkbox"]:checked + label::before {
                    background-color: #428bca;
                    border-color: #428bca;
                }

                .checkbox-primary input[type="checkbox"]:checked + label::after {
                    color: #fff;
                }

                .checkbox-danger input[type="checkbox"]:checked + label::before {
                    background-color: #d9534f;
                    border-color: #d9534f;
                }

                .checkbox-danger input[type="checkbox"]:checked + label::after {
                    color: #fff;
                }

                .checkbox-info input[type="checkbox"]:checked + label::before {
                    background-color: #5bc0de;
                    border-color: #5bc0de;
                }

                .checkbox-info input[type="checkbox"]:checked + label::after {
                    color: #fff;
                }

                .checkbox-warning input[type="checkbox"]:checked + label::before {
                    background-color: #f0ad4e;
                    border-color: #f0ad4e;
                }

                .checkbox-warning input[type="checkbox"]:checked + label::after {
                    color: #fff;
                }

                .checkbox-success input[type="checkbox"]:checked + label::before {
                    background-color: #5cb85c;
                    border-color: #5cb85c;
                }

                .checkbox-success input[type="checkbox"]:checked + label::after {
                    color: #fff;
                }
            </style>
 <script type="text/javascript">

     //function validation() {
     //    ShowProgress();
     //}
     function ShowSave(message) {

         $("#dialog").html(message);
         $("#dialog").show();
         $("#dialog").dialog({
             title: "LetIT",
             buttons: {
                 "Close": function () {
                     $(this).dialog('close');
                 }
             },
             modal: true
         });
     }

     function ShowPopup(message) {

         $("#dialog").html(message);
         $("#dialog").show();
         $("#dialog").dialog({
             title: "LetIT",
             buttons: {
                 "Close": function () {
                     $(this).dialog('close');
                 }
             },
             modal: true
         });
     }

     function ShowPopupDelete(message) {

         $("#dialog").html(message);
         $("#dialog").show();
         $("#dialog").dialog({
             title: "LetIT",
             buttons: {
                 "Yes": function () {
                     $("[id*=Button1]").click();
                     $(this).dialog('close');
                 },
                 "No": function () {
                     $('#ConfirmBit').val("NO");
                     $(this).dialog('close');
                 }
             },
             // modal: true
         });
     }

     function SelectAll() {
                   <%-- var intIndex = 0;
                    var CheckTrue = document.getElementById("<%=chkAll.ClientID %>").checked;
                    if (CheckTrue == true) {
                        var rowCount = document.getElementById("<%=multipleCheckBoxLoc.ClientID %>").getElementsByTagName("input").length;
                        for (i = 0; i < rowCount; i++) {
                            alert(rowCount);
                            if (document.getElementById("multipleCheckBoxLoc" + "_" + i)) {
                                if (document.getElementById("multipleCheckBoxLoc" + "_" + i).disabled != true)
                                    document.getElementById("multipleCheckBoxLoc" + "_" + i).checked = true;

                            }

                        }
                    }--%>

                    var intIndex = 0;

                    var rowCount = document.getElementById("<%=multipleCheckBoxLoc.ClientID %>").getElementsByTagName("input").length;

                    for (i = 0; i < rowCount; i++) {
                        if (document.getElementById("<%=chkAll.ClientID %>").checked == true) {
                            if (document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i)) {
                                if (document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i).disabled != true)
                                    document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i).checked = true;

                            }

                        }

                        else {
                            if (document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i)) {

                                if (document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i).disabled != true)

                                    document.getElementById("<%=multipleCheckBoxLoc.ClientID %>" + "_" + i).checked = false;

                 }

             }

         }
     }

 </script>
</head>
<body>
    <div id="dialog" style="display: none"></div>
    <form id="form1" runat="server">
        <div class="wrapper">
           <div class="content-wrapper">
                <!-- Content Header (Page header) -->
                <section class="content-header">
                    <div class="container-fluid topcls ">
                        <div class="row">
                            <div class="col-md-6">
                                <h1>
                                    <!-- InstanceBeginEditable name="header" -->
                                    USER MANAGEMENT<!-- InstanceEndEditable -->
                                </h1>
                            </div>
                            <div class="col-md-6">
                                <nav class="navbar navbar-default" role="navigation">
                                    <div class="container-fluid">
                                        <!-- Brand and toggle get grouped for better mobile display -->


                                        <!-- Collect the nav links, forms, and other content for toggling -->
                                        <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                                            <ul class="nav navbar-nav navbar-right">
                                                <li><a href="UserManagementView.aspx" class="round blue round-lg"><i class="fa fa-reply"></i></a></li>
                                                <li><a href="#" class="round orange round-lg"><i class="fa fa-edit"></i></a></li>
                                                <li><a href="#" class="round red round-lg"><i class="fa fa-trash"></i></a></li>

                                            </ul>

                                        </div>
                                        <!-- /.navbar-collapse -->
                                    </div>
                                    <!-- /.container-fluid -->
                                </nav>
                            </div>

                        </div>
                    </div>


                </section>

                <!-- Main content -->
                <section class="content">
                    <!-- InstanceBeginEditable name="EditRegion1" -->
                    <div class="row">
                        <div class="col-md-12">
                            <div class="box">
                                <div class="box-body">
                                    <div class="col-md-6">

                                        <div class="box">
                                            <div class="box-body">
                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                            <label>Agent Name</label>                                                             
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="lblAgentName"></asp:Label>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                           <label>User Name</label>                                                             
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="lblUserName"></asp:Label>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                           <label>Branch</label>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="lblBranch"></asp:Label>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                            <label>Email ID</label>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="lblEmailID" Text=""></asp:Label>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                          <label>Old Password</label>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="lblOldpwd" Text=""></asp:Label>
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                          <label>Reset Password</label>
                                                        </div>
                                                        <div class="col-md-6">
                                                            <asp:Label runat="server" ID="Label1" Text=""></asp:Label>
                                                            <asp:TextBox runat="server" ID="txtpassword" Width="200px" Height="30px" class="form-control"></asp:TextBox><br />
                                                            <asp:Button runat="server" ID="btnReset" class="btn btn-danger" OnClick="btnReset_Click"  Text="RESET" />
                                                        </div>
                                                    </div>
                                                </div>

                                                <div class="col-md-12">
                                                    <div class="form-group">
                                                        <div class="col-md-6">
                                                           <label>Not Active</label>
                                                                
                                                        </div>
                                                        <div class="col-md-6">
                                                            <div class="pretty p-default p-thick p-pulse">
                                                                <input type="checkbox" id="chkActive" runat="server" value="chkActive" />                               
                                                                <div class="state p-info-o">
                                                                    <label>.</label>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>


                                            </div>
                                        </div>



                                    </div>

                                    <div class="col-md-6">

                                        <div class="box">

                                            <div class="box-body">

                                                <asp:GridView runat="server" ID="mygrid" class="table table-bordred table-striped" HeaderStyle-BackColor="#094d70" AutoGenerateColumns="false"  OnRowDataBound="mygrid_RowDataBound">
                                                    <Columns>

                                                        <asp:TemplateField HeaderText="S.No" HeaderStyle-ForeColor="White">
                                                            <ItemTemplate>
                                                                <asp:Label ID="lblSNo" runat="server" Width="20px" Text='<%#Container.DataItemIndex+1%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>
                                                        <asp:TemplateField HeaderText="ROLE DETAILS" HeaderStyle-ForeColor="White" HeaderStyle-HorizontalAlign="Center">
                                                            <ItemTemplate>
                                                                <asp:DropDownList runat="server" ID="ddVal1" class="form-control"></asp:DropDownList>
                                                                <asp:Label runat="server" ID="lblField0" Visible="false" Text='<%#Eval("Field1") %>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>


                                                        <asp:TemplateField>
                                                            <ItemTemplate>
                                                                <asp:ImageButton ID="btnDelete" runat="server" OnClick="btnDelete_Click"  ImageUrl="~/assets/img/deleteicon.png" />
                                                                <asp:ImageButton ID="btInsert" runat="server" Visible="false" OnClick="btnNewonClick" ImageUrl="~/assets/img/plus.jpg" />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>


                                                    </Columns>

                                                </asp:GridView>
                                            </div>
                                        </div>

                                        <div class="box">
                                            <fieldset>
                                                <legend>
                                                <asp:CheckBox runat="server" CssClass="checkbox checkbox-inline" onchange="SelectAll();" Text="Select All" ID="chkAll" />
                                                </legend>

                                            </fieldset>
                                            <div class="box-body">

                                                <div class="checkbox checkbox-primary">
                                                    <asp:CheckBoxList runat="server" RepeatColumns="4" ID="multipleCheckBoxLoc"></asp:CheckBoxList>
                                                </div>
                                            </div>
                                        </div>
                                    </div>


                                    <div class="col-md-6">

                                        <div class="box">
                                            <div class="box-body">
                                            </div>
                                        </div>
                                    </div>

                                    <div class="col-md-12 text-right btnpad">
                                        <asp:Button runat="server" ID="btnSubmit" class="btn btn-primary" OnClick="btnSubmit_Click" Text="Save" />
                                        <asp:HiddenField ID="hdId" runat="server" />
                                         <asp:HiddenField runat="server" ID="ConfirmBit" />
                                        <asp:Button ID="Button1" runat="server" Text="Button" Style="display: none" OnClick="Button1_Click" />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
            </div>
            </div>
    </form>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <script src="assets/menu/bootstrap.min.js"></script>
</body>
</html>
