<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="RoleManagement.aspx.cs" Inherits="NVOCShipping.RoleManagement" %>

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
     <script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>
   <script type="text/javascript" src="https://ajax.cdnjs.com/ajax/libs/json2/20110223/json2.js"></script>
     <script type='text/javascript' src='https://ajax.googleapis.com/ajax/libs/jquery/1.4/jquery.min.js'></script>
     <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>

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
    <script>
function ShowSave(message) {

                    $("#dialog").html(message);
                    $("#dialog").show();
                    $("#dialog").dialog({
                        title: "Nerida",
                        buttons: {
                            "Close": function () {
                                $(this).dialog('close');
                            }
                        },
                        modal: true
                    });
                }
    </script>
</head>
<body>
    <form id="form1" runat="server">

         <div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <section class="content-header">
      <div class="container topcls">
      <div class="row">
 <div class="col-md-6">
      <h1><!-- InstanceBeginEditable name="header" -->ROLE MANAGMENT<!-- InstanceEndEditable -->
      </h1>
 </div>
          <div class="col-md-6"> <nav class="navbar navbar-default" role="navigation">
  <div class="container-fluid">
    <!-- Brand and toggle get grouped for better mobile display -->


    <!-- Collect the nav links, forms, and other content for toggling -->
    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
      <ul class="nav navbar-nav navbar-right">
         <li><a href="RoleManagementView.aspx" data-toggle="tooltip" data-placement="top" title="Go Back" class="round blue round-lg"><i class="fa fa-search"></i></a></li>
   
      </ul>

    </div><!-- /.navbar-collapse -->
  </div><!-- /.container-fluid -->
</nav></div>
         
      </div>
  </div>
     
    </section>

    <!-- Main content -->
    <section class="content">
	<!-- InstanceBeginEditable name="EditRegion1" -->
<div class="row">
    <div class="col-md-12">
    
        <div class="col-md-5">
            <div class="box">
            <div class="box-body">
                <div class="col-md-12">
                <div class="form-group">
                 <div class="col-md-3">
                    <label>Role Name<span> :</span><span style="color:red">*</span></label>
	            </div>
                    <div class="col-md-9">
                        <asp:TextBox runat="server" ID="txtRole"  class="form-control"></asp:TextBox>
                    </div>
                </div>
                    </div>

                 <div class="col-md-12">
                <div class="form-group">
                 <div class="col-md-3">
                    <label>Remarks<span> :</span><span style="color:red">*</span></label>
	            </div>
                    <div class="col-md-9">
                       <asp:TextBox runat="server" ID="txtRemarks" TextMode="MultiLine" style="width:260px; height:50px;" class="form-control"></asp:TextBox>
                    </div>
                </div>
                    </div>


                  <div class="col-md-12 text-right btnpad" style="margin-top:20px;">
                <asp:Button runat="server" ID="btnSubmit" class="btn btn-primary" OnClick="btnSubmit_Click"  OnClientClick="return validation();" Text="Save" />
                    <asp:HiddenField runat="server" ID="HDID" />
                  </div>
                    
                 </div>
        
         


            <!-- /.box-body -->
          </div>
            </div>

         <div class="col-md-7">

               <div class="box">
          
                  

                  
            <!-- /.box-header -->
            <div class="box-body">
                <div class="col-md-12 padl0">
                    <div class="form-group">
                        <div class="col-md-3 padl0">
                            <label>Modules<span> :</span><span style="color:red">*</span></label>
                        </div>
                        <div class="col-md-9">
                            <asp:DropDownList runat="server" ID="ddlRol" CssClass="form-control" Width="250px" OnClick="btnSubmit_Click"  AutoPostBack="true"></asp:DropDownList>
                        </div>

                    </div>
                </div>

         
               
              
                <asp:GridView ID="mygridRole" runat="server" AllowPaging="True" 
                    AutoGenerateColumns="False" class="table table-bordred table-striped" ShowHeaderWhenEmpty="true"  OnRowDataBound="mygridRole_RowDataBound" HeaderStyle-BackColor="#094d70">
                <Columns>
        

                 <asp:TemplateField HeaderStyle-Width="10px">
                <ItemTemplate>
                <img alt = "" style="cursor: pointer" src="images/plus.png"  />
                <asp:Panel ID="pnlOrders" runat="server" Style="display: none">

                    <asp:GridView ID="GridInvoice" runat="server" AutoGenerateColumns="false" OnRowDataBound="GridInvoice_RowDataBound"  
                        class="table table-bordred table-striped" HeaderStyle-BackColor="#99ccff" EnableTheming="true">
                        <Columns>

                                        
                            <asp:TemplateField>
                            <ItemTemplate>
                            <asp:CheckBox  runat="server" ID="chkTriedchk" />
                            </ItemTemplate>
                           </asp:TemplateField>

                             <asp:TemplateField>
                                <ItemTemplate>
                                <asp:Label runat="server" ID="lblThiredMenu"  Width="10px" Text='<%#Eval("ID")%>' ></asp:Label>
                                </ItemTemplate>
                              </asp:TemplateField>

                             <asp:TemplateField>
                            <ItemTemplate>
                            <asp:Label runat="server" ID="lblThiredMenuID"  Width="5px" Text='<%#Eval("MenuID")%>' ></asp:Label>
                            </ItemTemplate>
                            </asp:TemplateField>

                             <asp:TemplateField HeaderText="Trird Level Menus" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                            <asp:Label runat="server"  ID="lblThirdlevel"  Width="200px" Text='<%#Eval("FileName")%>' ></asp:Label>
                            </ItemTemplate>
                           </asp:TemplateField>

                            <asp:TemplateField  HeaderText="Search" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                            <asp:CheckBox runat="server"  ID="chkThSearch" />
                            </ItemTemplate>
                           </asp:TemplateField>

                            <asp:TemplateField  HeaderText="Edit" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                            <asp:CheckBox runat="server"  ID="chkThEdit" />
                            </ItemTemplate>
                           </asp:TemplateField>

                             <asp:TemplateField  HeaderText="Delete" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                            <asp:CheckBox runat="server"  ID="chkThdelete" />
                            </ItemTemplate>
                           </asp:TemplateField>

                               <asp:TemplateField  HeaderText="Print" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                            <asp:CheckBox runat="server"  ID="chkThPrint" />
                            </ItemTemplate>
                           </asp:TemplateField>

                        </Columns>
                    </asp:GridView>
                </asp:Panel>
                   
                  </ItemTemplate>
                </asp:TemplateField>
                
 
                 <asp:TemplateField>
                <ItemTemplate>
                <asp:Label runat="server" ID="lblScoundMenu"  Width="10px" Text='<%#Eval("ID")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>
                
                 <asp:TemplateField>
                  <ItemTemplate>
                     <asp:CheckBox runat="server"  ID="chkScoundchk" />
                </ItemTemplate>
                </asp:TemplateField>

              
               <asp:TemplateField>
                <ItemTemplate>
                <asp:Label runat="server" ID="lblMenuID"  Width="5px" Text='<%#Eval("MenuID")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

               <asp:TemplateField HeaderText="MODULE" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:Label runat="server" ID="lblModule"  Width="250px" Text='<%#Eval("FileName")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

                <asp:TemplateField>
                <ItemTemplate>
                <asp:Label runat="server" ID="lblUrl"  Width="10px" Text='<%#Eval("Url")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

                 <asp:TemplateField  HeaderText="Search" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                                      
                                <asp:CheckBox runat="server"  ID="chkMSearch" />
                              
                          
                            </ItemTemplate>
                           </asp:TemplateField>

                 <asp:TemplateField  HeaderText="Edit" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                                <asp:CheckBox runat="server"  ID="chkMEdit" />
                            </ItemTemplate>
                           </asp:TemplateField>

                 <asp:TemplateField  HeaderText="Delete" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                        
                                <asp:CheckBox runat="server"  ID="chkMdelete" />
                              
                            </ItemTemplate>
                           </asp:TemplateField>

                 <asp:TemplateField  HeaderText="Print" HeaderStyle-ForeColor="White">
                            <ItemTemplate>
                                
                                <asp:CheckBox runat="server"  ID="chkMPrint" />
                              
                            </ItemTemplate>
                           </asp:TemplateField>

                </Columns>

                    <EmptyDataTemplate>
                        <div align="center">No records found.</div>
                    </EmptyDataTemplate>
                </asp:GridView>
                 </div>
        
           


            <!-- /.box-body -->
          </div>
            </div>

          </div>
         
          
         
          </div>
          
       

<!-- InstanceEndEditable -->
     
    </section>
    <!-- /.content -->
  </div>
    </form>
</body>
</html>
