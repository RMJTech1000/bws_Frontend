<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="EmailConfiguration.aspx.cs" Inherits="EmailConfiguration" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

        <asp:UpdatePanel runat="server" ID="uploading">
        <ContentTemplate>
            <script type="text/javascript" src="../js/AutoText.js"></script>
            <script type="text/javascript" src="../js/Processing.js"></script>
            <style>
.topcls .col-md-6 h1
{
  margin-top:-5px;
}
                 .disabledStyle {
              
                background-color:#EBEBE4;
            }
              .EnableStyle {
              
                background-color:white;
            }
              .form-control.address{
                  margin-left: -24px;
margin-bottom: 26px;
              }
              .nav-tabs {
    padding-left: 0px !important;
    border-bottom: 0;
}
 .cusjob
    {
        margin-top:-48px;
        margin-left: 133px;
        color:#fff;
    }
              label
    {
        font-weight:normal;
    }
    .checkbox td,th
    {
        padding:9px 20px;
    }
    input[type=checkbox]
    {
        margin:4px 10px;
    }
    .checkbox {
  padding-left: 20px; }
  .checkbox label {
    display: inline-block;
    position: relative;
    padding-left:46px; }
    .checkbox label::before {
      content: "";
      display: inline-block;
      position: absolute;
      width: 17px;
      height: 17px;
      left: 0;
      margin-left: 20px;
      border: 1px solid #304974;
      /*border-radius: 3px;*/
      background-color: #fff;
      -webkit-transition: border 0.15s ease-in-out, color 0.15s ease-in-out;
      -o-transition: border 0.15s ease-in-out, color 0.15s ease-in-out;
      transition: border 0.15s ease-in-out, color 0.15s ease-in-out; }
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
      color: #555555; }
  .checkbox input[type="checkbox"] {
    opacity: 0; }
    .checkbox input[type="checkbox"]:focus + label::before {
      outline: thin dotted;
      outline: 5px auto -webkit-focus-ring-color;
      outline-offset: -2px; }
    .checkbox input[type="checkbox"]:checked + label::after {
      font-family: 'FontAwesome';
      content: "\f00c"; }
    .checkbox input[type="checkbox"]:disabled + label {
      opacity: 0.65; }
      .checkbox input[type="checkbox"]:disabled + label::before {
        background-color: #eeeeee;
        cursor: not-allowed; }
  .checkbox.checkbox-circle label::before {
    border-radius: 50%; }
  .checkbox.checkbox-inline {
    margin-top: 0; }

.checkbox-primary input[type="checkbox"]:checked + label::before {
  background-color: #428bca;
  border-color: #428bca; }
.checkbox-primary input[type="checkbox"]:checked + label::after {
  color: #fff; }

.checkbox-danger input[type="checkbox"]:checked + label::before {
  background-color: #d9534f;
  border-color: #d9534f; }
.checkbox-danger input[type="checkbox"]:checked + label::after {
  color: #fff; }

.checkbox-info input[type="checkbox"]:checked + label::before {
  background-color: #5bc0de;
  border-color: #5bc0de; }
.checkbox-info input[type="checkbox"]:checked + label::after {
  color: #fff; }

.checkbox-warning input[type="checkbox"]:checked + label::before {
  background-color: #f0ad4e;
  border-color: #f0ad4e; }
.checkbox-warning input[type="checkbox"]:checked + label::after {
  color: #fff; }

.checkbox-success input[type="checkbox"]:checked + label::before {
  background-color: #5cb85c;
  border-color: #5cb85c; }
.checkbox-success input[type="checkbox"]:checked + label::after {
  color: #fff; }


@media(max-width:1600px)
{
   .topcls.container {
    width: 100%;
    padding-right: 14px;
    padding-left: 14px;
}
}
table
{
    margin-top:-26px;
    margin-left:-19px;
}
            </style>
             <script type="text/javascript">
                

                 function validation() {

                     var ddlservice = $("#<%= ddlservice.ClientID %>").val();
                     if (ddlservice == 0) {
                         ShowPopup("Please Select Service Type")
                         return false;
                     }

                     var ddlBusinessType = $("#<%= ddlBusinessType.ClientID %>").val();
                     if (ddlBusinessType == 0) {
                         ShowPopup("Please Select Business Type")
                         return false;
                     }

                  
                     var txtCheckCategory = $("#<%= txtCheckCategory.ClientID %>").val();
                     if (txtCheckCategory.length == 0) {
                         ShowPopup("Please Enter Check Category")
                         return false;
                     }
                 }


                
                 function ShowPopup(message) {

                     $("#dialog").html(message);
                     $("#dialog").show();
                     $("#dialog").dialog({
                         title: sessionStorage.getItem("CompnayName"),
                         buttons: {
                             "Close": function () {
                                 $(this).dialog('close');
                             }
                         },
                         modal: true
                     });
                 }
                 
                 

              

             

             </script>


         <div id="dialog" style="display: none"></div>
      
                   <!-- Content Wrapper. Contains page content -->
  <div class="content-wrapper">
    <!-- Content Header (Page header) -->
    <section class="content-header">
        <div class="container topcls">
      <div class="row">
 <div class="col-md-6">
     <h1><!-- InstanceBeginEditable name="header" -->EMAIL CONFIGURATION<!-- InstanceEndEditable -->
      </h1>
     
 </div>
          <div class="col-md-6"> <nav class="navbar navbar-default" role="navigation">
  <div class="container-fluid">
    <!-- Brand and toggle get grouped for better mobile display -->


    <!-- Collect the nav links, forms, and other content for toggling -->
    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
      <ul class="nav navbar-nav navbar-right">
           <li class="dropdown" runat="server" id="btbLinkAdmin" visible="false">
          <a href="#" class="dropdown-toggle round round-lg yellow" data-toggle="dropdown"><i class="fas fa-clock-o fa" style="font-size:16px;"></i></a>
           <ul class="dropdown-menu" runat="server" id="divLogDetails">
      
          </ul>
        </li>
          <li><a href="PartySearch.aspx" data-toggle="tooltip" data-placement="top" title="Search" class="round blue round-lg"><i class="fa fa-reply"></i></a></li>
          <li><a href="#" runat="server" data-toggle="tooltip" data-placement="top" title="Edit" class="round orange round-lg" id="btnCustomerEdit" visible="false"><i class="fa fa-edit"></i></a></li>
          <li><a href="#" data-toggle="tooltip" data-placement="top" title="Delete" class="round red round-lg" runat="server" id="btnLnkkdelete" visible="false" onclick="return ShowPopupConfirm('Are you sure  want to Delete!!(Y/N)')"><i class="fa fa-trash"></i></a></li>
          <li><a href="CustomerParty.aspx" data-toggle="tooltip" data-placement="top" title="New" class="round green round-lg" visible="false" runat="server" id="btnLnkAdd"><i class="fa fa-plus"></i></a></li>

        
          <%--<li><a href="CustomerParty.aspx" class="round green round-lg"><i class="fa fa-plus"></i></a></li>--%>
      </ul>

    </div><!-- /.navbar-collapse -->
  </div><!-- /.container-fluid -->
</nav></div>
         
      </div>
  </div>

    </section>
      
 

                 <div class="tab-content">
                <div role="tabpanel" class="tab-pane active" id="Profile">
      	<section class="content">
				<div class="row">

                    
					<div class="col-md-12">

						<div class="col-md-11">
							<div class="box">
								<div class="box-body">

                                    <fieldset>
										<legend>Basic Details</legend>
                                        <div class="col-md-12 cusjob">
                                            <div class="col-md-6 pull-right"><div class="col-md-3">Party Name</div>
              <div class="col-md-1">:</div>
              <div class="col-md-8"><asp:Label ID="lblPartyName" runat="server"></asp:Label></div>
                                                
                                            </div>
                                     
                                        </div>
									</fieldset>

                                    <div class="col-md-12">

                                        <div class="col-md-3">
                                            <div class="form-group">
                                                <div class="col-md-12">
                                                 <label class="lable-fond">Service Type<span style="color: red">*</span></label>
                                                    <asp:DropDownList ID="ddlservice" runat="server" Width="200px" AutoPostBack="true" OnSelectedIndexChanged="ddlservice_SelectedIndexChanged"  autocomplete="off" CssClass="form-control">
                                                          <asp:ListItem Selected="True" Value="0" Text="--Select--"></asp:ListItem>
                                                          <asp:ListItem  Value="1" Text="Sea"></asp:ListItem>
                                                          <asp:ListItem  Value="2" Text="Air"></asp:ListItem>
                                                          <asp:ListItem  Value="3" Text="Others"></asp:ListItem>
                                                         
                                                    </asp:DropDownList>
                                                   
                                                </div>
                                              
                                            </div>
                                        </div>
                                        
                                        <div class="col-md-3">
                                            <div class="form-group">
                                                <div class="col-md-12">
                                                 <label class="lable-fond">Business Type<span style="color: red">*</span></label>
                                                    <asp:DropDownList ID="ddlBusinessType" runat="server" Width="200px" autocomplete="off" CssClass="form-control"></asp:DropDownList>
                                                </div>
                                              
                                            </div>
                                        </div>

                                         <div class="col-md-3">
                                            <div class="form-group">
                                                <div class="col-md-12">
                                                 <label class="lable-fond">Checklist Category<span style="color: red">*</span></label>
                                                    <asp:TextBox ID="txtCheckCategory" TextMode="MultiLine" Height="40px"  runat="server" Width="200px" autocomplete="off" CssClass="form-control"></asp:TextBox>
                                                </div>
                                              
                                            </div>
                                        </div>

                                         <div class="col-md-3">
                                            <div class="form-group">
                                                <div class="col-md-12">
                                                    <asp:Button ID="ddlAdd" Text="Add" runat="server" OnClientClick="return validation();"  OnClick="ddlAdd_Click" />
                                                </div>
                                              
                                            </div>
                                        </div>


                                         <asp:GridView runat="server" ID="mygrid"  ShowHeaderWhenEmpty="true"  AutoGenerateColumns="False" 
                                                    class="table table-bordred table-striped" HeaderStyle-BackColor="#094d70">
                                            <Columns>

                                                <asp:TemplateField HeaderText="Select" HeaderStyle-ForeColor="White">
                                                    <ItemTemplate>
                                                      <asp:LinkButton runat="server" ID="btnchargeSelect" OnClick="btnchargeSelect_Click" Text="Select" ></asp:LinkButton>
                                                     </ItemTemplate>
                                                   </asp:TemplateField>

                                               

                                                <asp:TemplateField HeaderText="Service Type" HeaderStyle-ForeColor="White">
                                                    <ItemTemplate>
                                                         
                                                        <asp:HiddenField runat="server" ID="HdIdValue" value='<%#Eval("ID") %>'></asp:HiddenField>
                                                        <asp:HiddenField runat="server" ID="HDServiceId" value='<%#Eval("ServiceType") %>'  />
                                                      <asp:Label runat="server" ID="lblTaxString" Text='<%#Eval("serviceTypev").ToString().ToUpper() %>'></asp:Label>
                                                       
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                 <asp:TemplateField HeaderText="Business Type" HeaderStyle-ForeColor="White">
                                                    <ItemTemplate>
                                                        <asp:HiddenField runat="server" ID="HDBusinessType" value='<%#Eval("BusinessType") %>'  />
                                                      <asp:Label runat="server" ID="lblBuss" Text='<%#Eval("BusinessTypeV").ToString().ToUpper() %>'></asp:Label>
                                                      
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                 <asp:TemplateField HeaderText="Checklist Category" HeaderStyle-ForeColor="White">
                                                    <ItemTemplate>
                                                      <asp:Label runat="server" ID="lblcheck" Text='<%#Eval("Category").ToString().ToUpper() %>'></asp:Label>
                                                      
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                
                                                  <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="btnChrgeCodeDelete" OnClick="btnChrgeCodeDelete_Click" runat="server"  ImageUrl="~/images/delete.jpg" />
                                                    </ItemTemplate>
                                                        </asp:TemplateField>
                                                </Columns>
                                                    <EmptyDataTemplate>
                            <p style="color:blue;">
                         No Record found
                      </p>
                  </EmptyDataTemplate>
                                        </asp:GridView>

                                        </div>

                                         


                                </div>
                                </div>
                            </div>
                        </div>
                   
                    <asp:HiddenField runat="server" ID="EcrID" />
                    <asp:HiddenField runat="server" ID="hdId" />
                    <asp:HiddenField runat="server" ID="HHDPartyID" />
			    </div>
              </section>
                    </div>
                     </div>
      </div>




            </ContentTemplate>
            </asp:UpdatePanel>



</asp:Content>

