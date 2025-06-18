<%@ Page Title="Business Approval" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="QuotationFormNew.aspx.cs" Inherits="FFD_QuotationFormNew" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <script src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/jquery-ui.min.js" type = "text/javascript"></script> 
<link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" rel = "Stylesheet" type="text/css" />
    <style>
        .ui-dialog .ui-dialog-titlebar-close
        {
            width:30px!important;
        }
        .ui-dialog .ui-dialog-titlebar-close span
        {
            display:initial!important;
        }
        .ui-dialog .ui-dialog-titlebar-close span
        {
            margin:-8px!important;
        }
        .form-group
        {
            margin-bottom:35px;
        }
    .disabledStyle {              
background-color:#EBEBE4;
}
.EnableStyle {
    background-color: white;
}
label {
    font-weight: normal;  
}
input[type=file] {
    display: inline;
}
.form-inline .title {
    width: 300px !important;
}
.lable-fond
{
    /*width:120px;*/
    text-align:left;
    font-size:13px; 
}
.invoicetbl.table>thead>tr>td, .invoicetbl.table>tbody>tr>td, .invoicetbl.table>tfoot>tr>td
{
    background-color:transparent;
    border:none;
    width:100px;
}
       
.ui-autocomplete-loading {
    background: white url(http://ajax.googleapis.com/ajax/libs/jqueryui/1.8.2/themes/smoothness/images/ui-anim_basic_16x16.gif) right center no-repeat;
}
.CssDelete
{
    display:none;
}
.table>thead>tr>th, .table>tbody>tr>th, .table>tfoot>tr>th
{
    background-color:#304974;
    color:#fff;
}
.table>thead>tr>td, .table>tbody>tr>td, .table>tfoot>tr>td
{
    text-align:left;
    background-color: #f9f9f9;
    color: #4f4e4e;
    border:1px solid #ccc;
    padding: 5px 4px;
    font-family: 'Oswald', sans-serif;
    font-weight:400;
}
.no-margin-bottom {
    margin-bottom: 0;
}
.item-input {
    float: left;
    width: 65%;
    margin-left: 10px;
    }
.item-select {
float: left;
margin: 4px;
}
                .delete-row {
    float: left;
    margin-top: 4px;
}
                .panel.grid
                {
                    margin-bottom:42px;
                    margin-top: -58px;
                }
                .panel.tax
                {
                    padding-bottom:52px;
                }
                .panel-default>.panel-heading
                {
                    color: #fff;
    background-color: #69b8d6;
    border-color: #ddd;
                }
                .panel.panel-default .panel-body
                {
                    margin-bottom:-64px!important;
                    padding: 0px 14px;
                }

                .panel.panel-default .panel-body .col-md-12
                {
                    margin-bottom:-19px;
                }
                .input-group-addon
                {
                    padding: 4px 12px;
                }
                .text:after {
                    font-family: 'FontAwesome';
                    content: '\f274';
                    position: absolute;
                    right: 6px;
                }
                .no-padding-right {
                 padding-right: 0;
                }
                .form-control
                {
                    height:24px;
                    padding: 0px 12px;
                    /*width: 200px;*/
                }

                /*.form-control
                {
                    border-radius:0;
                }*/
                .amountloc
                {
                    text-align:right;
                }
                .col-md-5.lbl
                {
                    width:77px;
                }
                .col-md-5.lbl2
                {
                    width:79px;
                }
                .col-md-5.lbl3
                {
                    /*width:114px;*/
                }
               .col-md-5.lbl4
                {
                    width:80px;
                }
                /*.col-md-5 .txtbox
                {
                    height: 17px;
                    font-size: 12px;
                    width: 150px;
                }*/
                .select2-container
                {
                   width:200px!important;
                }
                .form-group-sm .form-control
                {
                    height:24px!important;
                    padding:0px 10px;
                    
                }
                .form-group-sm .form-control.txtroe
                {
                    width:60px!important;
                }
                .form-group-sm .form-control.txtrate
                {
                    width:60px!important;
                }
                .form-group-sm .form-control.txtamt
                {
                    width:60px!important;
                }
                .form-group-sm .form-control.txtlocamt
                {
                    width:60px!important;
                }
                .form-group-sm select.form-control
                {
                    width: 100px!important;
                }
                .table > thead > tr > td, .table > tbody > tr > td, .table > tfoot > tr > td
                {
                    padding: 4px 4px!important;
                    vertical-align:inherit;
                    font-size:12px;
                }
                .select2-container--default .select2-selection--single
                {
                    /*width:150px!important;
                    border: 1px solid #000!important;
                    height:18px!important;
                    border-radius:0!important;*/
                     border-radius: 5px!important;
                    border: 1px solid #a9a9a9!important;
                    display: block!important;
                    width: 100%!important;
                    height: 24px!important;
                    padding: 6px 12px!important;
                    font-size: 12px!important;
                    text-transform: uppercase!important;
                    line-height: 1.42857143!important;
                    color: #555!important;
                    background-color: #fff!important;
                    background-image: none!important;
                    border: 1px solid #ccc!important;
                    border-radius: 4px!important;
                    -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075)!important;
                    box-shadow: inset 0 1px 1px rgba(0,0,0,.075)!important;
                    -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s!important;
                    -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s!important;
                    transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s!important;

                }
                .select2-results__option[aria-selected]
                {
                    font-size:12px!important;
                }
                .select2-container--default .select2-selection--single .select2-selection__rendered
                {
                    line-height:10px!important;
                    font-size:11px!important;
                    text-align:left!important;
                    padding-left: 0px!important;
                    padding-right: 0px!important;
                    margin-top:0!important;
                }
                .select2-container--default .select2-selection--single .select2-selection__arrow b
                {
                    display:none!important;
                }
                .table > thead > tr > th, .table > tbody > tr > th, .table > tfoot > tr > th
                {
                    padding: 2px 4px;
                    font-size:12px;
                }
               .table-striped>tbody>tr:nth-child(odd)>td
 {
  background-color: #87acc81a;
  
}
.table-striped>tbody>tr:nth-child(even)>td
 {
  background-color:fef9f4;
  
}
.panel.entry .form-group
{
    margin-bottom:32px;
}
.panel.entry .form-control
{
    margin-bottom:0;
}
/*.panel .tax  .form-control.txtdes
{
    width:100px!important;
}
.panel .tax  .form-control.txtcharge
{
    width:180px!important;
}
.panel .tax  .form-control
{
    width:60px!important;
}*/

.panel .tax .txtcrncy
{
    width:97px;
}
.panel .tax .txtroe
{
    width:60px;
}
.panel .tax .txtrate
{
    width:60px;
}
.panel .tax .txtamt
{
    width:60px;
}
.panel .tax .txtlocamt
{
    width:60px;
}
.panel .tax .txtunit
{
    width:97px;
}
.panel .tax .table
{
    margin-bottom:0;
}
.panel.entry>.panel-heading
                {
                    color: #fff;
    background-color: #69b8d6;
    border-color: #ddd;
                }
.panel.entry .panel-body
{
    padding:0px 15px;
}
/*Responsive*/
@media(max-width:1080px)
{
    .panel.entry .form-control
    {
        width:119px!important;
    }
    .panel.entry .form-control.address
{
    width:386px!important;
    height:70px!important;
}
    .select2-container
    {
        width:119px!important;
    }
    .select2-results__option[aria-selected]
    {
        width:170px!important;
    }
    .select2-container--open .select2-dropdown--below
    {
         width:170px!important;
    }
    .panel .tax .txtcrncy
{
    width:94px;
}
.panel .tax .txtroe
{
    width:50px!important;
}
.panel .tax .txtrate
{
    width:50px!important;
}
.panel .tax .txtamt
{
    width:50px!important;
}
.panel .tax .txtlocamt
{
    width:50px!important;
}
.panel .tax .txtunit
{
    width:94px!important;
}
.panel .tax .txtdes
{
    width:110px!important;
}

}
@media(max-width:1280px) 
{
    .panel.entry .form-control.address
    {
        /*width:630px!important;*/
    }
    .panel.entry .bank.form-control
    {
        width:130px!important;
    }
     .panel.entry .form-control
    {
        width:175px!important;
    }
      .panel.entry .form-control.address
{
    width:507px!important;
    height:70px!important;
}
       .panel .tax .txtcrncy
{
    width:94px!important;
}
.panel .tax .txtroe
{
    width:70px!important;
}
.panel .tax .txtrate
{
    width:70px!important;
}
.panel .tax .txtamt
{
    width:70px!important;
}
.panel .tax .txtlocamt
{
    width:70px!important;
}
.panel .tax .txtunit
{
    width:94px!important;
}
.panel .tax .txtdes
{
    width:157px!important;
}
.select2-container
    {
        width:175px!important;
    }
}
@media(max-width:1300px)
{
    .panel.entry .form-control.address
{
    width:508px!important;
    height:70px!important;
}
}

@media(max-width:1600px)
{
                    .panel .bank.form-control
                {
                    width:145px!important;
                }
}

.panel.entry .form-control.address
{
    width:555px;
    height:70px;
}
.cusjob
    {
        margin-top:-48px;
        margin-left:216px;
        color:#fff;
    }
    .cusjob .lblcus
    {
        margin-left:-11px;
    }
                .checkbox td, th {
                    padding: 8px 20px;
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
@media(max-width:1600px)
{
   .roundoff input[type=checkbox]
{
  margin:0;
}

}
@media(max-width:1280px)
{
   .roundoff input[type=checkbox]
{
  margin:0;
}
}
.table>tbody>tr>td.rightalign
{
    text-align:right;
    /*width: 74px;*/
}
.table>tbody>tr>td.taxwidth{
    text-align:right;
    /*width:10px;*/
}
.table>tbody>tr>td.unitwidth{
    
    /*width:10px;*/
}
.table>tbody>tr>td.chargewidth{
    /*width:165px;*/
}
.table>tbody>tr>td.unitswidth{   
    /*width:10px;*/
}
            </style>
       <script  type="text/javascript">
           function numericvalidation()
           {
               var key = event.keyCode;
               if ((key >= 48) && (key <= 58) || (key == 46))

                   return true;
               else
                   return false;
           }
    </script>

   <script type="text/javascript">
       //On UpdatePanel Refresh.
       var prm = Sys.WebForms.PageRequestManager.getInstance();
       if (prm != null) {
           prm.add_endRequest(function (sender, e) {
               if (sender._postBackSettings.panelsToUpdate != null) {
                   SetAutoComplete();
                   SetDatePicker();
                  
               }
           });
       };
        $(function ()
       {
           SetAutoComplete();
           SetDatePicker();
         
       });
       function SetDatePicker() {
           $(document).ready(function () {
               $("[id$=txtRSDate]").datepicker({
                   changeMonth: true,
                   changeYear: true,
                   uiLibrary: 'bootstrap',
                   dateFormat: 'dd-mm-yy'
               });

               $("[id$=txtValidtyUpto]").datepicker({
                   changeMonth: true,
                   changeYear: true,
                   uiLibrary: 'bootstrap',
                   dateFormat: 'dd-mm-yy'
               });


           });
       }
   
       function SetAutoComplete()
       {
           
           $(document).ready(function () {
               $("[id*=txtBkgParty]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CustomerBkgParty.asmx/GetCustomerMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddBkgParty]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtPLR]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i)
                   {
                      $("[id*=hddPLR]", $(e.target).closest("div")).val(i.item.val);
                   }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtPOL]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddPOL]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });



           $(document).ready(function () {
               $("[id*=txtPOD]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddPOD]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtPLD]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddPLD]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });







           $(document).ready(function () {
               $("[id*=txtCommodity]").autocomplete({

                   source: function (request, response) {
                       $.ajax({
                           url: 'WebService/CommodityService.asmx/GetCommodityMaster', async: false, jsonpCallback: 'jsonCallback', data: "{ 'prefix': '" + request.term + "'}",
                           dataType: "json", type: "POST", contentType: "application/json; charset=utf-8",
                           success: function (data) { response($.map(data.d, function (item) { return { label: item.split('-')[0], val: item.split('-')[1] } })) },
                           error: function (response) { alert(response.responseText); }, failure: function (response) { alert(response.responseText); }
                       });
                   },
                   select: function (e, i) { $("[id*=hddCommodity]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1

               });
           });



           $(document).ready(function () {
               $("[id*=txtPayAt]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddPayAt]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });



           $(document).ready(function () {
               $("[id*=txtOverseas]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/OverSeaseWebService.asmx/GetOverSeaseCustomerMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddOverseas]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtShippingLine]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/ShipperLinesService.asmx/GetShipperVendorMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddShpLine]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtPLD_Carrier]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/WebServicePort.asmx/GetPortsMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddPLD_Carrier]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           $(document).ready(function () {
               $("[id*=txtOverseas_Carrier]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/OverSeaseWebService.asmx/GetOverSeaseCustomerMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=hddOverseas_Carrier]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });



           $(document).ready(function () {
               $("[id*=txtVal0]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/ChargeCodeService.asmx/GetChargeCodeExpance") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdVal0]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           $(document).ready(function () {
               var ID = 1;
               $("[id*=txtVal1]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/RateRcvdFromService.asmx/GetRateFromValue") %>',
                           data: "{ 'prefix': '" + request.term + "','BussID': '" + ID + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdVal1]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtVal2]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CntrService.asmx/GetCntnValueMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdVal2]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           $(document).ready(function () {
               $("[id*=txtVal3]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/BasicService.asmx/GetBasisMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdVal3]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtVal4]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CurrencyService.asmx/GetCurrencyMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdVal4]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });




           $(document).ready(function () {
               $("[id*=txtSRVal0]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/ChargeCodeServiceSellingService.asmx/GetChargeCodeSelling") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdSRVal0]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               var ID = 2;
               $("[id*=txtSRVal1]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/RateRcvdFromService.asmx/GetRateFromValue") %>',
                           data: "{ 'prefix': '" + request.term + "','BussID': '" + ID + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdSRVal1]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtSRVal2]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CntrService.asmx/GetCntnValueMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdSRVal2]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           $(document).ready(function () {
               $("[id*=txtSRVal3]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/BasicService.asmx/GetBasisMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdSRVal3]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtSRVal4]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CurrencyService.asmx/GetCurrencyMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdSRVal4]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtRRVal0]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/RefoundTransService.asmx/GetRefundTransMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdRRVal0]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {

               var ID = 0;

               $("[id*=GView_RR] [id*=txtRRVal1]").click(function () {
                   var row = $(this).closest("tr");
                   ID = row.find("[id*=HdRRVal0]").val();

               });
               $("[id*=txtRRVal1]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/RateRcvdFromService.asmx/GetRateFromValue") %>',
                           data: "{ 'prefix': '" + request.term + "','BussID': '" + ID + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdRRVal1]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtRRVal2]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CntrService.asmx/GetCntnValueMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdRRVal2]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtRRVal3]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/BasicService.asmx/GetBasisMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdRRVal3]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });


           $(document).ready(function () {
               $("[id*=txtRRVal4]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CurrencyService.asmx/GetCurrencyMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdRRVal4]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           $(document).ready(function () {
               $("[id*=txtCnVal1]").autocomplete({
                   source: function (request, response) {
                       $.ajax({
                           url: '<%=ResolveUrl("WebService/CntrService.asmx/GetCntnValueMaster") %>',
                           data: "{ 'prefix': '" + request.term + "'}", dataType: "json", type: "POST", contentType: "application/json; charset=utf-8", success: function (data) {
                               response($.map(data.d, function (item) {
                                   return { label: item.split('-')[0], val: item.split('-')[1] }
                               }))
                           }, error: function (response) { alert(response.responseText); },
                           failure: function (response) { alert(response.responseText); }
                       });
                   }, select: function (e, i) { $("[id*=HdCnVal1]", $(e.target).closest("div")).val(i.item.val); }, minLength: 1
               });
           });

           

       }

       function validationall() {
           var EmptyFields = "";
           
           var txtRSDate = $("#<%= txtRSDate.ClientID %>").val();
           if (txtRSDate == "")
           {
              EmptyFields += "<span style='color:red;'>*</span> Please Enter the Quotation Date</br>";
           }

           var ddTrade = $("#<%= ddTrade.ClientID %>").val();
           if (ddTrade == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please Select the Trade</br>";
           }

           var txtBkgParty = $("#<%= txtBkgParty.ClientID %>").val();
           if (txtBkgParty == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please Enter the  Booking Party</br>";
           }

           var txtPLR = $("#<%= txtPLR.ClientID %>").val();
           if (txtPLR == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Enter the  Place of Receipt</br>";
           }

           var txtPOL = $("#<%= txtPOL.ClientID %>").val();
           if (txtPOL == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Enter the  Port of Loading</br>";
           }

           var txtPOD = $("#<%= txtPOD.ClientID %>").val();
           if (txtPOD == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Enter the  Port of Discharge</br>";
           }

           var txtPLD = $("#<%= txtPLD.ClientID %>").val();
           if (txtPLD == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Enter the  Place of Delivery</br>";
           }

           var ddBusinessTypes = $("#<%= ddBusinessTypes.ClientID %>").val();
           if (ddBusinessTypes == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Business Type</br>";
           }

           var txtCommodity = $("#<%= txtCommodity.ClientID %>").val();
           if (txtCommodity == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Commodity</br>";
           }

           var ddServiceTypes = $("#<%= ddServiceTypes.ClientID %>").val();
           if (ddServiceTypes == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Service Type</br>";
           }

           var ddServiceMode = $("#<%= ddServiceMode.ClientID %>").val();
           if (ddServiceMode == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Service Mode</br>";
           }

           var ddTermsOfShipment = $("#<%= ddTermsOfShipment.ClientID %>").val();
           if (ddTermsOfShipment == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Terms of Shipment</br>";
           }

           var txtShippingLine = $("#<%= txtShippingLine.ClientID %>").val();
           if (txtShippingLine == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the Account Holder</br>";
           }

           var txtShippingLine = $("#<%= txtShippingLine.ClientID %>").val();
           if (txtShippingLine == "") {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Shipping Lines</br>";
           }

           var ddTermsOfShipment_Carrier = $("#<%= ddTermsOfShipment_Carrier.ClientID %>").val();
           if (ddTermsOfShipment_Carrier == 0) {
               EmptyFields += "<span style='color:red;'>*</span> Please  Select the  Shipment carrier</br>";
           }


          
           if (EmptyFields != "") {
               ShowPopup(EmptyFields);
               return false;
           }
           //var myExtender = $find('LoadingModal');
           //myExtender.show();
           //return true;
       }

       function ContValidationAdd() {
           var EmptyContGrid = "";

      if (EmptyContGrid != "")
                     {
                         ShowPopup(EmptyContGrid);
                         return false;
                     }
                     //var myExtender = $find('LoadingModal');
                     //myExtender.show();
                     //return true;

       }

       function BuyRateValidationAdd() {
           var EmptyBuyGrid = "";

           var ddlBuyNarration = $("#<%= ddlBuyNarration.ClientID %>").val();
           if (ddlBuyNarration == 0) {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Select the  Narration</br>";
           }

           var ddlBuyPayable = $("#<%= ddlBuyPayable.ClientID %>").val();
           if (ddlBuyPayable == 0) {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Select the  Payable To</br>";
           }

           var ddlBuyContainerTypes = $("#<%= ddlBuyContainerTypes.ClientID %>").val();
           if (ddlBuyContainerTypes == 0) {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Select the  Container Type</br>";
           }

           var ddlBuyBasis = $("#<%= ddlBuyBasis.ClientID %>").val();
           if (ddlBuyBasis == 0) {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Select the Basis</br>";
           }

           var ddlBuyCurrency = $("#<%= ddlBuyCurrency.ClientID %>").val();
           if (ddlBuyCurrency == 0) {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Select the Currency</br>";
           }

           var txtBuyAmount = $("#<%= txtBuyAmount.ClientID %>").val();
           if (txtBuyAmount == "") {
               EmptyBuyGrid += "<span style='color:red;'>*</span> Please  Enter Amount</br>";
           }
           

      if (EmptyBuyGrid != "")
                     {
                         ShowPopup(EmptyBuyGrid);
                         return false;
                     }
                     //var myExtender = $find('LoadingModal');
                     //myExtender.show();
                     //return true;

       }

       function SellRateValidationAdd() {
           var EmptySellGrid = "";

           var ddlSellNarration = $("#<%= ddlSellNarration.ClientID %>").val();
           if (ddlSellNarration == 0) {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Select the  Narration</br>";
           }

           var ddlSellPayable = $("#<%= ddlSellPayable.ClientID %>").val();
           if (ddlSellPayable == 0) {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Select the  Payable To</br>";
           }

           var ddlSellContainer = $("#<%= ddlSellContainer.ClientID %>").val();
           if (ddlSellContainer == 0) {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Select the  Container Type</br>";
           }

           var ddlSellBasis = $("#<%= ddlSellBasis.ClientID %>").val();
           if (ddlSellBasis == 0) {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Select the Basis</br>";
           }

           var ddlSellCurrency = $("#<%= ddlSellCurrency.ClientID %>").val();
           if (ddlSellCurrency == 0) {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Select the Currency</br>";
           }

           var txtSellAmount = $("#<%= txtSellAmount.ClientID %>").val();
           if (txtSellAmount == "") {
               EmptySellGrid += "<span style='color:red;'>*</span> Please  Enter Amount</br>";
           }
           

      if (EmptySellGrid != "")
                     {
                         ShowPopup(EmptySellGrid);
                         return false;
                     }
                     //var myExtender = $find('LoadingModal');
                     //myExtender.show();
                     //return true;

       }

       function RebvalidationAdd() {
           var EmptyRebGrid = "";

           var ddlRebTransaction = $("#<%= ddlRebTransaction.ClientID %>").val();
           if (ddlRebTransaction == 0) {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Select the Transaction</br>";
           }

           var ddlRebFromTo = $("#<%= ddlRebFromTo.ClientID %>").val();
           if (ddlRebFromTo == 0) {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Select the FromTo</br>";
           }

           var ddlRebContrType = $("#<%= ddlRebContrType.ClientID %>").val();
           if (ddlRebContrType == 0) {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Select the Container Type</br>";
           }

           var ddlRebBasis = $("#<%= ddlRebBasis.ClientID %>").val();
           if (ddlRebBasis == 0) {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Select the Basis</br>";
           }

           var ddlRebCurrency = $("#<%= ddlRebCurrency.ClientID %>").val();
           if (ddlRebCurrency == 0) {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Select the Currency</br>";
           }

           var txtRebAmount = $("#<%= txtRebAmount.ClientID %>").val();
           if (txtRebAmount == "") {
               EmptyRebGrid += "<span style='color:red;'>*</span> Please  Enter Amount</br>";
           }

           if (EmptyRebGrid != "")
                     {
                         ShowPopup(EmptyRebGrid);
                         return false;
                     }
                     //var myExtender = $find('LoadingModal');
                     //myExtender.show();
                     //return true;
       }



           

      

       function numericvalidation() {
           var key = event.keyCode;
           if ((key >= 48) && (key <= 58))
               return true;
           else
               return false;
       }

    function ShowPopup(message) {
        var myExtender = $find('LoadingModal');
         myExtender.hide();
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


     </script>

     <asp:UpdatePanel runat="server" ID="update">
    <%--   
          <Triggers>
              <asp:PostBackTrigger ControlID="btnSave" />
          </Triggers>--%>
    
    <ContentTemplate>


        <div id="dialog" style="display: none"></div>
        <div class="content-wrapper">
    <!-- Content Header (Page header) -->
 <section class="content-header">
        <div class="container topcls">
            
      <div class="row">
 <div class="col-md-6">
     <h1>BUSINESS APPROVAL</h1>
 </div>
          <div class="col-md-6"> <nav class="navbar navbar-default" role="navigation">
  <div class="container-fluid">
    


    
    <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
      <ul class="nav navbar-nav navbar-right">
           <li class="dropdown">
          <a href="#" class="dropdown-toggle round round-lg yellow" data-toggle="dropdown"><i class="fas fa-clock-o fa" style="font-size:16px;"></i></a>
          <ul class="dropdown-menu" runat="server" id="divLogDetails">
        
          </ul>
        </li>
                  <li><a href="QuotationSearchNew.aspx" class="round blue round-lg"><i class="fa fa-reply"></i></a></li>
          <li><a href="#" runat="server" class="round orange round-lg" id="btnCustomerEdit"><i class="fa fa-edit"></i></a></li>
          <li><a href="#" class="round red round-lg" runat="server" id="A2" onclick="return ShowPopupConfirm('Are you sure  want to Delete!!(Y/N)')"><i class="fa fa-trash"></i></a></li>
          <li><a href="QuotationFormNew.aspx" class="round green round-lg"><i class="fa fa-plus"></i></a></li>
      </ul>

    </div>
  </div>
</nav></div>
         
      </div>
  </div>
    </section>
            <section class="content">
                <div class="row">
                    <div class="col-md-12">
                       	<div class="box">
								<div class="box-body">

                                    <fieldset>
										<legend>Basic Details</legend>
									</fieldset>
                                    
                                            <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                    <div class="col-xs-12 padlr0" style="padding-top:7px; padding-bottom:7px;">
                          <div class="col-md-6 padlr0">
                          <div class="form-group">
                              <div class="col-md-5 lbl3"><label class="lable-fond">Quotation No</label></div>
                              <div class="col-md-1">:</div>
                              <div class="col-md-5"><asp:Label ID="lblRateSheetNo" runat="server" CssClass="AutoGenLbl" Text="***"></asp:Label></div>
                          </div>
                    </div>
                      <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Date</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtRSDate" runat="server"  CssClass="form-control" ></asp:TextBox></div>
                        </div>
                </div>
                                                         <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Trade</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddTrade" runat="server"   CssClass="form-control" AutoPostBack="True">
                                <asp:ListItem Value="0" Text=""></asp:ListItem>
                                <asp:ListItem Value="1" Text="Import"></asp:ListItem>
                                <asp:ListItem Value="2" Text="Export"></asp:ListItem>
                                <asp:ListItem Value="3" Text="Both"></asp:ListItem>
                                <asp:ListItem Value="4" Text="Others"></asp:ListItem>
                                                  </asp:DropDownList></div>
                        </div>
                </div>
                                                        <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Booking Party</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtBkgParty" runat="server" CssClass="form-control"></asp:TextBox>
                                                  <asp:HiddenField runat="server" ID="hddBkgParty" /></div>
                        </div>
                </div>
                                                        <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Place Of Receipt</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtPLR" runat="server"  CssClass="form-control" ></asp:TextBox>
                    <asp:HiddenField runat="server" ID="hddPLR" /></div>
                        </div>
                </div>
                                                         <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Port Of Loading</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtPOL" runat="server"  CssClass="form-control"></asp:TextBox>
                    <asp:HiddenField runat="server" ID="hddPOL" /></div>
                        </div>
                </div>
                                                          <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Port of Discharge</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtPOD" runat="server" CssClass="form-control"> </asp:TextBox> 
                    <asp:HiddenField runat="server" ID="hddPOD" /></div>
                        </div>
                </div>
                                                        <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Place of Delivery</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtPLD" runat="server"  CssClass="form-control"></asp:TextBox>
                        <asp:HiddenField runat="server" ID="hddPLD" /></div>
                        </div>
                </div>
                                                        <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Business Type</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddBusinessTypes"  CssClass="form-control"  runat="server"  AutoPostBack="True" ></asp:DropDownList></div>
                        </div>
                </div>
                <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Commodity</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtCommodity" runat="server" CssClass="form-control"></asp:TextBox>    
                        <asp:HiddenField runat="server" ID="hddCommodity" /></div>
                        </div>
                </div>
                                                          <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Service Type</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddServiceTypes"  CssClass="form-control" runat="server"></asp:DropDownList></div>
                        </div>
                </div>
                                                        <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Service Mode</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddServiceMode"  CssClass="form-control" runat="server"></asp:DropDownList></div>
                        </div>
                </div>
                                                         <div class="col-md-6 padlr0">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Terms of Shipment</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddTermsOfShipment"  CssClass="form-control" runat="server" AutoPostBack="True"></asp:DropDownList></div>
                        </div>
                </div>
                                                         <div class="col-md-6 padlr0" id="ACHolder" runat="server">
                        <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Account Holder</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:DropDownList ID="ddACHolder"  CssClass="form-control" runat="server"></asp:DropDownList></div>
                        </div>
                </div>
                                                        <%-- <div class="col-md-6 padlr0" id="FPayableAt" runat="server" visible="false">
                                                             <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Freight PayableAt</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtPayAt" runat="server"  CssClass="form-control"> </asp:TextBox>
                                 <asp:HiddenField runat="server" ID="hddPayAt" /></div>
                        </div>                                
                                                             </div>--%>
                                                      <%--  <div class="col-md-12 padlr0" id="CollectingAgent" runat="server" visible="false">
                                                            <div class="col-md-6 padlr0">
                                                                 <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Collecting Agent Involved</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:CheckBox ID="ChkCollectAgent" runat="server" AutoPostBack="true"  /></div>
                         <div id="DivCollectAgent" runat="server" visible="false">
                                 </div>
                                                                 </div>  
                                                            </div>
                                                            <div class="col-md-6 padlr0">
                                                                <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Collecting Agent</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"> <asp:TextBox ID="txtOverseas" runat="server"  CssClass="form-control-small" ></asp:TextBox>
                                <asp:HiddenField runat="server" ID="hddOverseas" /></div>
                        </div>  
                                                            </div>
            
                                                             
                                                            </div>
                                                        
                                               <div class="col-md-6 padlr0" id="FreeDays" runat="server"  visible="false">
                                                                <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Freedays Agreed</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:TextBox ID="txtFreeDays" runat="server"  CssClass="form-control-small" ></asp:TextBox></div>
                        </div>  
                                                            </div> 
                                                         <div class="col-md-6 padlr0" id="ImpOpt" runat="server" visible="false">
                                                             <div class="col-md-6">
                                                                  <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Freehand</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:RadioButton ID="optFH" runat="server"  GroupName="optG1"   AutoPostBack="True"  /></div>
                        </div>  
                                                             </div>
                                                             <div class="col-md-6">
                                                                  <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Nomination</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:RadioButton ID="optAN" runat="server"  GroupName="optG1"  AutoPostBack="True" /></div>
                        </div>  
                                                             </div>
                                                               
                                                            </div> 

                                                     <div class="col-md-6 padlr0" id="ExpOpt" runat="server" visible="false">
                                                             <div class="col-md-6">
                                                                  <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Self Generated</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:RadioButton ID="optSG" runat="server"  GroupName="optG2" AutoPostBack="True"  /></div>
                        </div>  
                                                             </div>
                                                             <div class="col-md-6">
                                                                  <div class="form-group">
                            <div class="col-md-5 lbl3"><label class="lable-fond">Nomination</label></div>
                            <div class="col-md-1">:</div>
                            <div class="col-md-5"><asp:RadioButton ID="optBN" runat="server"  GroupName="optG2"  AutoPostBack="True" /></div>
                        </div>  
                                                             </div>
                                                               
                                                            </div> --%>
                                                   </div>
                                                        
                                                   </div>
                                                        </div>

                                    </div>


                               <div class="box-body">
                                   <fieldset>
										<legend>&nbsp;</legend>
									</fieldset>
                                   <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                    <div class="col-xs-12 padlr0" style="padding-top:7px; padding-bottom:7px;">
                                                         <div class="col-md-6 padlr0">
                          <div class="form-group">
                              <div class="col-md-5 lbl3"><label class="lable-fond">Shipping Line</label></div>
                              <div class="col-md-1">:</div>
                              <div class="col-md-5"><asp:TextBox ID="txtShippingLine"  CssClass="form-control" runat="server"></asp:TextBox> 
                            <asp:HiddenField runat="server" ID="hddShpLine" /></div>
                          </div>
                    </div>
                                                        <div class="col-md-6 padlr0">
                          <div class="form-group">
                              <div class="col-md-5 lbl3"><label class="lable-fond">Terms of Shipment(with carrier)</label></div>
                              <div class="col-md-1">:</div>
                              <div class="col-md-5"><asp:DropDownList ID="ddTermsOfShipment_Carrier" CssClass="form-control" runat="server"></asp:DropDownList></div>
                          </div>
                    </div>
                                                        </div>
                                    </div>
                               </div>
                    </div>
                                 <div class="box-body">
                                     <fieldset>
										<legend>CONTAINER TYPES</legend>
									</fieldset>
                                       <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                     <div class="col-xs-6 tax padlr0">
                                                         <asp:GridView ID="GView_CntrTypes" runat="server" AutoGenerateColumns="False" 
                                    CssClass="gridTopHeader gridtwofit table table-bordred table-striped" OnRowDataBound="GView_CntrTypes_RowDataBound"  OnRowCommand="GView_CntrTypes_RowCommand" >
                                    <Columns>
                                    
                                          
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton  ID="btnCTInsert" OnClick="btnCTInsert_Click" runat="server" OnClientClick="return ContValidationAdd();"  Visible="false"  ImageUrl="~/Images/plus.jpg"/>
                                            </ItemTemplate>
                                        </asp:TemplateField> 
                                           
                                        <asp:TemplateField >
                                            <ItemTemplate>
                                                <asp:Label ID="lblSNo"  runat="server" Text='<%#Container.DataItemIndex +1 %>'></asp:Label>
                                                  <asp:Label ID="lblID"  Visible="false" runat="server" Text='<%#Eval("ID")%>'></asp:Label> 
                                            </ItemTemplate>                                            
                                        </asp:TemplateField>
                            
                                        <asp:TemplateField HeaderText = "Container Types">
                                            <ItemTemplate>
                                                  <asp:DropDownList ID="ddlVal0" runat="server" CssClass="form-control"></asp:DropDownList>
                                               <asp:HiddenField runat="server" ID="HdCnVal1" value='<%#Eval("Field1")%>' />
                                            
                                            </ItemTemplate>
                                        </asp:TemplateField> 
                                        
                                        <asp:TemplateField HeaderText = "Approx Units Per Job">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtCNVal2"  runat="server" Text='<%#Eval("Field3")%>' CssClass="form-control"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateField> 
                                        
                                   
                                    </Columns>
                                    <HeaderStyle CssClass="gridTopHeader" />
                                    <AlternatingRowStyle CssClass="gridAltText" />
                                </asp:GridView>
                                                         </div>
                                                    </div>
                                           </div>
                                     </div>

                                <div class="box-body">
                                     <fieldset>
										<legend>BUYING RATE</legend>
									</fieldset>
                                       <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                     <div class="col-xs-12 tax padlr0">
                                                        <table class="table table-bordred table-striped" style="border: 1px solid #808080;">
                                               <tr>
                                                   <th>NARRATION</th>
                                                   <th>PAYABLE TO</th>
                                                   <th>CONTR TYPE</th>
                                                   <th>BASIS</th>
                                                   <th>CURRENCY</th>
                                                   <th>AMOUNT</th>
                                                   <th>ADD</th>
                                               </tr>
                                               <tr>
                                                   
                                                   <td><asp:DropDownList runat="server"   placeholder="Select Charge Code" CssClass="form-control txtcharge" ID="ddlBuyNarration"></asp:DropDownList>
                                                       <asp:HiddenField runat="server" ID="HDChargeCode" /></td>
                                                   <td>
                                                       <asp:DropDownList runat="server"  ID="ddlBuyPayable" AutoPostBack="true" CssClass="form-control"></asp:DropDownList>
                                                    </td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlBuyContainerTypes" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                    <td><asp:DropDownList runat="server" ID="ddlBuyBasis" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlBuyCurrency" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:TextBox runat="server" CssClass="form-control" autocomplete="off"  ID="txtBuyAmount"></asp:TextBox></td>
                                                  
                                              
                                                   <td><asp:ImageButton runat="server" Width="30px" ID="btnAdd" OnClick="btnAdd_Click" OnClientClick="return BuyRateValidationAdd();" ImageUrl="~/images/plus.jpg" /> 
                                                       <asp:HiddenField runat="server" ID="HDModifyID" />
                                                   </td>
                                               </tr>
                                           </table>
                                                    </div>
                                                    <div class="col-xs-12 tax padlr0">
                                                         <asp:GridView runat="server" ID="GridView_Buying" ShowHeaderWhenEmpty="true" class="table table-bordred table-striped table-condensed" 
                                                    AutoGenerateColumns="false">
                                                    <Columns>
                                                     
                                                     <asp:TemplateField HeaderText="Select" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:LinkButton runat="server" ID="btnBuyingSelect" OnClick="btnBuyingSelect_Click" Text="Select" ></asp:LinkButton>
                                                     </ItemTemplate>
                                                   </asp:TemplateField>

                                                          <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:Label runat="server" ID="lblID" Text='<%#Eval("ID") %>'></asp:Label>
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="S.No" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:Label ID="lblSNo" runat="server" Width="20px" Text='<%#Container.DataItemIndex+1%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField HeaderText="NARRATION" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="chargewidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval1"  Text='<%#Eval("Field1")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue1" value='<%#Eval("Field2")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                          <asp:TemplateField HeaderText="PAYABLE TO" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval2" Text='<%#Eval("Field3")%>'></asp:Label>
                                                                     <asp:HiddenField runat="server" ID="HDValue2" value='<%#Eval("Field4")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                         <asp:TemplateField HeaderText="CONTAINER TYPE" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval3" Text='<%#Eval("Field5")%>'></asp:Label>
                                                                 <asp:HiddenField runat="server" ID="HDValue3" value='<%#Eval("Field6")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="BASIS" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitswidth">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server"  ID="lblval4" Text='<%#Eval("Field7")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue4" value='<%#Eval("Field8")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField> 

                                                          <asp:TemplateField HeaderText="CURRENCY" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                   <asp:Label runat="server"  ID="lblval5" Text='<%#Eval("Field9")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue5" value='<%#Eval("Field10")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>


                                                           <asp:TemplateField HeaderText="AMOUNT" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server" ID="lblval6" Text='<%#Eval("Field11")%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:ImageButton runat="server" ID="GrdDelete"  Height="20px" ImageUrl="~/images/delete.jpg" />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>
                                                        

                                                    </Columns>
                                                     <EmptyDataTemplate>
                                                            <asp:Label ID="lblEmptySearch" runat="server">No Results Found</asp:Label>
                                                        </EmptyDataTemplate>
                                                </asp:GridView>
                                                        <asp:HiddenField runat="server" ID="HdGrID" />
                                                    </div>
                                                    </div>
                                           </div>
                                    </div>
                                <div class="box-body">
                                     <fieldset>
										<legend>SELLING RATE</legend>
									</fieldset>
                                       <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                     <div class="col-xs-12 tax padlr0">
                                                        <table class="table table-bordred table-striped" style="border: 1px solid #808080;">
                                               <tr>
                                                   <th>NARRATION</th>
                                                   <th>PAYABLE TO</th>
                                                   <th>CONTR TYPE</th>
                                                   <th>BASIS</th>
                                                   <th>CURRENCY</th>
                                                   <th>AMOUNT</th>
                                                   <th>ADD</th>
                                               </tr>
                                               <tr>
                                                   
                                                  <td><asp:DropDownList runat="server"   placeholder="Select Narration" CssClass="form-control txtcharge" ID="ddlSellNarration"></asp:DropDownList>
                                                       <asp:HiddenField runat="server" ID="HiddenField3" /></td>
                                                   <td>
                                                       <asp:DropDownList runat="server"  ID="ddlSellPayable" AutoPostBack="true" CssClass="form-control"></asp:DropDownList>
                                                    </td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlSellContainer" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                    <td><asp:DropDownList runat="server" ID="ddlSellBasis" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlSellCurrency" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:TextBox runat="server" CssClass="form-control"  autocomplete="off"  ID="txtSellAmount"></asp:TextBox></td>
                                                  
                                              
                                                   <td><asp:ImageButton runat="server" Width="30px" ID="btnSellAdd"  OnClick="btnSellAdd_Click" OnClientClick="return SellRateValidationAdd();"   ImageUrl="~/images/plus.jpg" /> 
                                                       <asp:HiddenField runat="server" ID="HDSellID" />
                                                   </td>
                                               </tr>
                                           </table>
                                                    </div>
                                                     <div class="col-xs-12 tax padlr0">
                                                         <asp:GridView runat="server" ID="GridView_Sell" ShowHeaderWhenEmpty="true" class="table table-bordred table-striped table-condensed" 
                                                    AutoGenerateColumns="false">
                                                    <Columns>
                                                     
                                                     <asp:TemplateField HeaderText="Select" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:LinkButton runat="server" ID="btnchargeSelect"  Text="Select" ></asp:LinkButton>
                                                     </ItemTemplate>
                                                   </asp:TemplateField>

                                                          <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:Label runat="server" ID="lblID" Text='<%#Eval("ID") %>'></asp:Label>
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="S.No" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:Label ID="lblSNo" runat="server" Width="20px" Text='<%#Container.DataItemIndex+1%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField HeaderText="NARRATION" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="chargewidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval1"  Text='<%#Eval("Field1")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue1" value='<%#Eval("Field2")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                          <asp:TemplateField HeaderText="PAYABLE TO" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval2" Text='<%#Eval("Field3")%>'></asp:Label>
                                                                     <asp:HiddenField runat="server" ID="HDValue2" value='<%#Eval("Field4")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                         <asp:TemplateField HeaderText="CONTAINER TYPE" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval3" Text='<%#Eval("Field5")%>'></asp:Label>
                                                                 <asp:HiddenField runat="server" ID="HDValue3" value='<%#Eval("Field6")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="BASIS" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitswidth">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server"  ID="lblval4" Text='<%#Eval("Field7")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue4" value='<%#Eval("Field8")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField> 

                                                          <asp:TemplateField HeaderText="CURRENCY" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                   <asp:Label runat="server"  ID="lblva5" Text='<%#Eval("Field9")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue5" value='<%#Eval("Field10")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>


                                                           <asp:TemplateField HeaderText="AMOUNT" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server" ID="lblval6" Text='<%#Eval("Field11")%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:ImageButton runat="server" ID="GrdDelete"  Height="20px" ImageUrl="~/images/delete.jpg" />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>
                                                        

                                                    </Columns>
                                                     <EmptyDataTemplate>
                                                            <asp:Label ID="lblEmptySearch" runat="server">No Results Found</asp:Label>
                                                        </EmptyDataTemplate>
                                                </asp:GridView>
                                                        <asp:HiddenField runat="server" ID="HDGrdHD1" />
                                                    </div>
                                                    </div>
                                           </div>
                                    </div>
                               <div class="box-body">
                                     <fieldset>
										<legend>REBATE</legend>
									</fieldset>
                                       <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                     <div class="col-xs-12 tax padlr0">
                                                        <table class="table table-bordred table-striped" style="border: 1px solid #808080;">
                                               <tr>
                                                   <th>TRANSACTION</th>
                                                   <th>FROM TO</th>
                                                   <th>CONTR TYPE</th>
                                                   <th>BASIS</th>
                                                   <th>CURRENCY</th>
                                                   <th>AMOUNT</th>
                                                   <th>ADD</th>
                                               </tr>
                                               <tr>
                                                   
                                                   <td><asp:DropDownList ID="ddlRebTransaction" runat="server">
                                                       <asp:ListItem Value="0" Text="-Select-"></asp:ListItem>
                                                       <asp:ListItem Value="1" Text="Payable"></asp:ListItem>
                                                       <asp:ListItem Value="2" Text="Rebate"></asp:ListItem>
                                                       </asp:DropDownList></td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlRebFromTo" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:DropDownList runat="server"  ID="ddlRebContrType" AutoPostBack="true" CssClass="form-control"></asp:DropDownList></td>
                                                    <td><asp:DropDownList runat="server" ID="ddlRebBasis" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:DropDownList runat="server" ID="ddlRebCurrency" CssClass="form-control"></asp:DropDownList></td>
                                                   <td><asp:TextBox runat="server" CssClass="form-control" autocomplete="off"  ID="txtRebAmount"></asp:TextBox></td>
                                                  
                                              
                                                   <td><asp:ImageButton runat="server" Width="30px" ID="btnRebate" OnClick="btnRebate_Click"  OnClientClick="return RebvalidationAdd();"   ImageUrl="~/images/plus.jpg" /> 
                                                       <asp:HiddenField runat="server" ID="HiddenField2" />
                                                   </td>
                                               </tr>
                                           </table>
                                                    </div>
                                                    <div class="col-xs-12 tax padlr0">
                                                    <asp:GridView runat="server" ID="GridView_Rebate" ShowHeaderWhenEmpty="true" class="table table-bordred table-striped table-condensed" 
                                                    AutoGenerateColumns="false">
                                                    <Columns>
                                                     
                                                     <asp:TemplateField HeaderText="Select" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:LinkButton runat="server" ID="btnchargeSelect"  Text="Select" ></asp:LinkButton>
                                                     </ItemTemplate>
                                                   </asp:TemplateField>

                                                    <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                    <ItemTemplate>
                                                      <asp:Label runat="server" ID="lblID" Text='<%#Eval("ID") %>'></asp:Label>
                                                      </ItemTemplate>
                                                   </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="S.No" HeaderStyle-ForeColor="White" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:Label ID="lblSNo" runat="server" Width="20px" Text='<%#Container.DataItemIndex+1%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField HeaderText="TRANSACTION" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="chargewidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval1"  Text='<%#Eval("Field1")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue1" value='<%#Eval("Field2")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                          <asp:TemplateField HeaderText="FROM TO" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval2" Text='<%#Eval("Field3")%>'></asp:Label>
                                                                     <asp:HiddenField runat="server" ID="HDValue2" value='<%#Eval("Field4")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                         <asp:TemplateField HeaderText="CONTAINER TYPE" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                  <asp:Label runat="server" ID="lblval3" Text='<%#Eval("Field5")%>'></asp:Label>
                                                                 <asp:HiddenField runat="server" ID="HDValue3" value='<%#Eval("Field6")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                       <asp:TemplateField HeaderText="BASIS" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="unitswidth">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server"  ID="lblval4" Text='<%#Eval("Field7")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue4" value='<%#Eval("Field8")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField> 

                                                          <asp:TemplateField HeaderText="CURRENCY" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                   <asp:Label runat="server"  ID="lblva5" Text='<%#Eval("Field9")%>'></asp:Label>
                                                                <asp:HiddenField runat="server" ID="HDValue5" value='<%#Eval("Field10")%>' />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>


                                                           <asp:TemplateField HeaderText="AMOUNT" HeaderStyle-HorizontalAlign="Center" ItemStyle-CssClass="rightalign">
                                                            <ItemTemplate>
                                                                 <asp:Label runat="server" ID="lblval6" Text='<%#Eval("Field11")%>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateField>

                                                        <asp:TemplateField ItemStyle-CssClass="unitwidth">
                                                            <ItemTemplate>
                                                                <asp:ImageButton runat="server" ID="GrdDelete" Height="20px" ImageUrl="~/images/delete.jpg" />
                                                            </ItemTemplate>
                                                        </asp:TemplateField>
                                                        

                                                    </Columns>
                                                     <EmptyDataTemplate>
                                                            <asp:Label ID="lblEmptySearch" runat="server">No Results Found</asp:Label>
                                                        </EmptyDataTemplate>
                                                </asp:GridView>
                                                        <asp:HiddenField runat="server" ID="HDGrd2" />
                                                    </div>

                                                    </div>
                                           </div>
                                    </div>
                                <div class="box-body">
                                    <div class="panel-body" style="margin-top:-22px;">
                                        <div class="row">
                                            <div class="col-xs-12 padlr0">
                                                <div class="col-xs-6 padlr0">
                                                    <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Expire After Single Use</label></div>
                                                        <div class="col-xs-1">&nbsp;</div>
                                                        <div class="col-xs-5"><asp:CheckBox ID="chkSingleUse" runat="server" AutoPostBack="true"/></div>
                                                    </div>
                                                        </div>
                                                    <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Location to Handle*</label></div>
                                                        <div class="col-xs-1">&nbsp;</div>
                                                        <div class="col-xs-5"><asp:DropDownList ID="ddHandleLocation" CssClass="form-control"  runat="server">
                                </asp:DropDownList>
                                                    </div>
                                                        </div>
                                                        </div>
                                                     <div class="col-xs-12 padlr0">
                                                          <div class="form-group">
                                                               <div class="col-xs-5"><label class="lable-fond">Special Instructions</label></div>
                                                        <div class="col-xs-1">&nbsp;</div>
                                                        <div class="col-xs-5">
                                                            <asp:TextBox ID="txtRemarks" runat="server"  CssClass="form-control"  TextMode="MultiLine" Width="360px" Height="100px"></asp:TextBox>
                                                    </div>
                                                              </div>
                                                     </div>

                                                   
                                                </div>
                                                 <div class="col-xs-6">
                                                      <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Validity Upto*</label></div>
                                                        <div class="col-xs-1">&nbsp;</div>
                                                        <div class="col-xs-5"><asp:TextBox ID="txtValidtyUpto" runat="server"  CssClass="form-control"></asp:TextBox></div>
                                                    </div>
                                                        </div>
                                                 </div>
                                            </div>
                                        </div>
                                        </div>
                                    </div>
                               <div class="box-body">
                                   <div class="panel-body" style="margin-top:-22px;">
                                       <div class="row">
                                           <div class="col-xs-6 padlr0">
                                                 <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Status</label></div>
                                                        <div class="col-xs-1">:</div>
                                                        <div class="col-xs-5">
                                                    </div>
                                                        </div>
                                                        </div>
                                               <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Approved/Rejected Date</label></div>
                                                        <div class="col-xs-1">:</div>
                                                        <div class="col-xs-5">
                                                    </div>
                                                        </div>
                                                        </div>
                                                <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Approved/Rejected By</label></div>
                                                        <div class="col-xs-1">:</div>
                                                        <div class="col-xs-5">
                                                    </div>
                                                        </div>
                                                        </div>
                                                     <div class="col-xs-12 padlr0">
                                                        <div class="form-group">
                                                        <div class="col-xs-5"><label class="lable-fond">Comments</label></div>
                                                        <div class="col-xs-1">:</div>
                                                        <div class="col-xs-5">
                                                    </div>
                                                        </div>
                                                        </div>
                                           </div>
                                      
                                           </div>
                                   </div>
                               </div>
                              <%--  <div class="box-body">
                                     <fieldset>
										<legend>FRIEGHT COMMISSION</legend>
									</fieldset>
                                       <div class="panel-body" style="margin-top:-22px;">
                                                <div class="row">
                                                     <div class="col-xs-12 tax padlr0">
                                                         <div class="col-xs-3">
                                                             <div class="form-group">
                                                                 <asp:DropDownList ID="ddlFrightCommision"   CssClass="form-control"  runat="server">
                                                                 </asp:DropDownList>
                                                             </div>
                                                         </div>
                                                         </div>
                                                    </div>
                                           </div>
                                    </div>--%>

                                <div class="box-body" ID="Div4" runat="server">
                                   <div class="text-right">
                                    <asp:Button ID="btnSave" runat="server" Text="SAVE" OnClick="btnSave_Click" CssClass="btn btn-primary btn-sm"  OnClientClick="return validationall();" />
               <asp:Button ID="btnPDF" runat="server" Text="PREVIEW" OnClick="btnPDF_Click"  CssClass="btn btn-danger btn-sm" />
            <asp:Button ID="btnApproval" runat="server" Text="SEND TO APPROVAL" OnClick="btnApproval_Click" Visible="false"  CssClass="btn btn-warning btn-sm"  />
                      <asp:HiddenField runat="server" ID="HdRateID" />
                      <asp:HiddenField runat="server" ID="HdSellingID" />
                      <asp:HiddenField runat="server" ID="HdRebID" />
                       
                                       </div>
                               </div>
                </div>
                </section>
 
     

            </div>

       






       

            </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>

