<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UserManagementView.aspx.cs" Inherits="NVOCShipping.UserManagementView" %>

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
</head>

<body>
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
                                                <li><a href="UserManagementSearch.aspx" class="round blue round-lg"><i class="fa fa-reply"></i></a></li>
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
    
        <div class="col-md-9">
            <div class="box">
          
            <!-- /.box-header -->
            <div class="box-body">
	
                  <asp:GridView ID="ddlProductView" runat="server" AllowPaging="True"  onrowdatabound="ddlProductView_RowDataBound" 
              AutoGenerateColumns="False" class="table table-bordred table-striped" HeaderStyle-BackColor="#094d70">
                <Columns>
        
                  <asp:TemplateField HeaderText="S.No" HeaderStyle-ForeColor="White">
              <ItemTemplate>
               <asp:Label ID="lblSNo" runat="server"  Text='<%#Container.DataItemIndex+1%>' ></asp:Label>
               </ItemTemplate>
               </asp:TemplateField>
                    
            <asp:TemplateField HeaderText="Employee Name" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:LinkButton ID="linkbtn" runat="server"  PostBackUrl='<%#string.Format("~/UserManagement.aspx?RegId={0}",(Eval("ID").ToString())) %>' Text='<%#Eval("UserName") %>'></asp:LinkButton>
                </ItemTemplate>
            </asp:TemplateField>

                <asp:TemplateField HeaderText="Agency" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:Label runat="server" ID="lblemail"  Text='<%#Eval("AgencyName")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

                 <asp:TemplateField HeaderText="Branch" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:Label runat="server" ID="lblemail"  Text='<%#Eval("CityName")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>
                     <asp:TemplateField HeaderText="Country" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:Label runat="server" ID="lblemail"  Text='<%#Eval("CountryName")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

                     <asp:TemplateField HeaderText="Status" HeaderStyle-ForeColor="White">
                <ItemTemplate>
                <asp:Label runat="server" ID="lblemail"  Text='<%#Eval("ActiveV")%>' ></asp:Label>
                </ItemTemplate>
                </asp:TemplateField>

                
             
                </Columns>
                </asp:GridView>

                      <div class="col-md-12 padlr0">
<div class="col-md-5 padlr0 page" id="paging">
<asp:Label CssClass="pageIntext" ID="lblTotalRecord" runat="server"></asp:Label>
<asp:Label CssClass="pageIntext" ID="lblPageInfo" runat="server"></asp:Label>
</div>

<div class="col-md-7 text-right padlr0">
<nav aria-label="...">
  <ul class="pagination">
    <li class="page-item">
      <span class="page-link"><asp:LinkButton ID="lbtnFirst" runat="server" CausesValidation="false" OnClick="lbtnFirst_Click">First</asp:LinkButton></span>
    </li>
    <li class="page-item">
      <span class="page-link"><asp:LinkButton ID="lbtnPrevious" runat="server" CausesValidation="false" OnClick="lbtnPrevious_Click">Previous</asp:LinkButton></span>
    </li>
    <li class="page-item"><a class="page-link" href="#"> <asp:DataList ID="ddlPager" runat="server" RepeatDirection="Horizontal" OnItemCommand="ddlPager_ItemCommand" OnItemDataBound="ddlPager_ItemDataBound">
                      <ItemTemplate>
                        <asp:LinkButton ID="lnkbtnPaging" runat="server" CommandArgument='<%# Eval("PageIndex") %>' CommandName="Paging" Text='<%# Eval("PageText") %>'></asp:LinkButton>
                      </ItemTemplate>
             </asp:DataList></a></li>
    
    <li class="page-item">
      <asp:LinkButton ID="lbtnNext" runat="server" CausesValidation="false" OnClick="lbtnNext_Click">Next</asp:LinkButton>
    </li>
    <li class="page-item">
      <asp:LinkButton ID="lbtnLast" runat="server" CausesValidation="false" OnClick="lbtnLast_Click">Last</asp:LinkButton>
    </li>
  </ul>
</nav>

</div>     
</div>
                 </div>
        
           


            <!-- /.box-body -->
          </div>
            </div>

         <div class="col-md-3">
              <div class="box">
          
               <div class="box-body">
	
             		<div class="col-md-12">
			<div class="form-group">
          <asp:TextBox ID="txtAgencyName" CssClass="form-control" runat="server" Placeholder="Agency Name"></asp:TextBox>
          </div>
		  </div>
          
		            <div class="col-md-12">
          <div class="form-group">
          <asp:TextBox ID="txtUserName" CssClass="form-control" runat="server" Placeholder="User Name"></asp:TextBox>
          </div>
		  </div>
      
 <div class="col-md-12">
          <div class="form-group">
          <asp:TextBox ID="txtBranch" CssClass="form-control" runat="server" Placeholder="Branch"></asp:TextBox>
          </div>
		  </div>
       

                   
		  
		            <div class="col-md-12" style="margin-top:0px;">
                        <asp:Button runat="server" ID="btnSearch" Text="Search" class="btn btn-info" OnClick="btnsearch_Click" />
			</div>

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
