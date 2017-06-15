<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AjaxPrinting.aspx.vb" Inherits="AjaxPrinting" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style2 {
            width: 500px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">

        <asp:HiddenField ID="HF_Error" runat="server" />
        <asp:HiddenField ID="HF_ErrorString" runat="server" />
        <asp:HiddenField ID="HF_Exchange" runat="server" />
        <asp:HiddenField ID="HF_NoOK" runat="server" />
        <asp:HiddenField ID="HF_CanDelete" runat="server" />
        <asp:HiddenField ID="HF_Release" runat="server" />
        <asp:HiddenField ID="HF_PayID2" runat="server" />
        <asp:HiddenField ID="HF_POID" runat="server" /> 
        
        <asp:HiddenField ID="HF_Temp" runat="server" /> 

<asp:Panel ID="PNL_ACC" visible="false" runat="server" ScrollBars="Auto">
    <asp:HiddenField ID="HF_CustID" runat="server" />
    <asp:HiddenField ID="HF_TopData" runat="server" />
    <asp:HiddenField ID="HF_Currency" runat="server" />
    <asp:HiddenField ID="HF_PayID" runat="server" />

    <table id="linestop" cellpadding="0" cellspacing="0" width="100%" style="padding-top:5px; position:relative; border-bottom: solid 1px lightgray; "></table>
    <div id="pnl_lines" style="overflow:auto">
        <table id="lines" cellpadding="0" cellspacing="0" width="100%" style="padding-top:0px; position:relative; border-bottom: solid 1px lightgray; "></table>
    </div>

    <asp:Repeater ID="RPT_SalesItems" runat="server">
        <ItemTemplate>
            <asp:HiddenField ID="HF_Lines" Value='<%# Eval("fk_Account_ID") & "~" & Eval("Item_No") & "~" & Eval("Description") & "~" & Eval("Qty", "{0:#,###.##}") & "~" & Eval("UOM") & "~" & Eval("Unit", "{0:#,###.00###}") & "~" & Eval("Amount", "{0:#,###.00}") & "~" & Eval("fk_Tax_Code_ID") & "~" & Eval("Tax_Rate") & "~" & Eval("fk_Job_ID") & "~" & Eval("Job_No") & "~" & Eval("Line_ID") & "~" & Eval("PO_SO_Line_ID") & "~" & Eval("Prev_Rec") & "~" & Eval("Prev_Inv")%>' runat="server" />
        </ItemTemplate>
    </asp:Repeater>

    <asp:Repeater ID="RPT_Rec" runat="server">
        <ItemTemplate>
            <asp:HiddenField ID="HF_Lines" Value='<%# Eval("Item_No") & "~" & Eval("Description") & "~" & Eval("Qty", "{0:#,###.##}") & "~" & Eval("UOM") & "~" & Eval("Lot_No") & "~" & Eval("Requires_Cert") & "~" & Eval("File_Name") & "~" & Eval("Line_ID")%>' runat="server" />
        </ItemTemplate>
    </asp:Repeater>

    <asp:Repeater ID="RPT_Payments" runat="server">
        <ItemTemplate>
            <asp:HiddenField ID="HF_Lines" Value='<%# Eval("Cust_Vend_ID") & "~" & Eval("Cust_Vend_Name") & "~" & Eval("Amount", "{0:#,###.00}") & "~" & Eval("Transaction_Code") & "~" & Eval("Cheque_No") & "~" & Eval("Printed") & "~" & Eval("Memo") & "~" & Eval("Applied") & "~" & Eval("Pay_Line_ID") & "~" & Eval("fk_Act_ID")%>' runat="server" />
        </ItemTemplate>
    </asp:Repeater>


</asp:Panel>


<asp:Panel ID="PNL_Summary" Visible="false" runat="server">
<asp:HiddenField ID="HF_RowCount" runat="server" />
<table id="summary" cellpadding="0" cellspacing="0" class="tablerow" style="width:100%; position:relative; border-bottom: solid 1px darkgray">
<asp:Repeater ID="RPT_List" runat="server">
    <ItemTemplate>
        <tr id='<%# "trsum_" & LCase(Eval("Doc_No").ToString) & LCase(Eval("Name").ToString) & LCase(Eval("Description").ToString)%>'>
            <td id="td_nobot" align="left" valign="top" class="tablecell3" style="width:125px; border-right: solid 1px lightgray;"><span id='<%# "sl_" & Eval("Doc_ID")%>' class="blacklink8"><%#Eval("Doc_No")%></span></td>
            <td align="left" valign="top" style="width:230px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label27" runat="server" Text='<%#Eval("Name")%>' CssClass="text8"/></td>
            <td align="center" valign="top" style="width:90px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label37" runat="server" style="text-align: center" Text='<%#Eval("Doc_Date", "{0:yyyy-MM-dd}")%>' CssClass="text8"/></td>
            <td id="td_tax" align="right" valign="top" style="width:80px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label3" runat="server" style="text-align: right" Text='<%#Eval("Tax", "{0:#,###.00}")%>' CssClass="text8"/></td>
            <td id="td_total" align="right" valign="top" style="width:100px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label2" runat="server" style="text-align: right" Text='<%#Eval("Total", "{0:#,###.00}")%>' CssClass="text8"/></td>
            <td id="tr_balance" align="right" valign="top" style="width:100px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label38" runat="server" style="text-align: right" Text='<%#Eval("Balance", "{0:#,###.00}")%>' CssClass="text8"/></td>
            <td id="td_currency" align="center" valign="top" style="width:80px; border-right: solid 1px lightgray;" class="tablecell3"><asp:Label ID="Label14" runat="server" style="text-align: right" Text='<%#Eval("Doc_Currency")%>' CssClass="text8"/></td>
            <td id="tr_type" align="left" valign="top" class="tablecell3" style="border-right: solid 1px lightgray; width: 90px; display:none"><asp:Label ID="Label4" runat="server" Text='<%#Eval("Type")%>' CssClass="text8"/></td>
            <td align="left" valign="top"  class="tablecell3" style="border-right: solid 1px lightgray;"><asp:Label ID="Label1" runat="server" Text='<%#Eval("Description")%>' CssClass="text8"/><asp:HiddenField ID="HF_InvID" Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
        </tr>
        </ItemTemplate>
    </asp:Repeater>
</table>
</asp:Panel>

<asp:Panel ID="PNL_RecSummary" Visible="false" runat="server">
<table id="summary" cellpadding="0" cellspacing="0" style="width:100%; position:relative; border-bottom: solid 1px darkgray">
<asp:Repeater ID="RPT_RecList" runat="server">
    <ItemTemplate>
        <tr id='<%# "trsum_" & LCase(Eval("Doc_No").ToString) & LCase(Eval("Name").ToString) & LCase(Eval("Description").ToString)%>' style="background-color:#e5f3ff">
            <td align="left" valign="top" style="width:125px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><span id='<%# "sl_" & Eval("Doc_ID")%>' class="blacklink8"><%#Eval("Doc_No")%></span></td>
            <td align="left" valign="top" style="width:230px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label27" runat="server" Text='<%#Eval("Name")%>' CssClass="text8"/></td>
            <td align="center" valign="top" style="width:90px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label37" runat="server" style="text-align: center" Text='<%#Eval("Doc_Date", "{0:yyyy-MM-dd}")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:120px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label10" runat="server" Text='<%#Eval("Location_Name")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:120px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label11" runat="server" Text='<%#Eval("FirstName") & " " & Eval("LastName")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:300px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label1" runat="server" Text='<%#Eval("Description")%>' CssClass="text8"/><asp:HiddenField ID="HF_InvID" Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
            <td align="center" valign="top" style="width:80px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label16" runat="server" Text='<%#Eval("password")%>' CssClass="text8"/></td>
            <td></td>
        </tr>
        </ItemTemplate>
        <AlternatingItemTemplate>
        <tr id='<%# "trsum_" & LCase(Eval("Doc_No").ToString) & LCase(Eval("Name").ToString) & LCase(Eval("Description").ToString)%>'>
            <td align="left" valign="top" style="width:125px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><span id='<%# "sl_" & Eval("Doc_ID")%>' class="blacklink8"><%#Eval("Doc_No")%></span></td>
            <td align="left" valign="top" style="width:230px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label27" runat="server" Text='<%#Eval("Name")%>' CssClass="text8"/></td>
            <td align="center" valign="top" style="width:90px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label37" runat="server" style="text-align: center" Text='<%#Eval("Doc_Date", "{0:yyyy-MM-dd}")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:120px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label10" runat="server" Text='<%#Eval("Location_Name")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:120px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label11" runat="server" Text='<%#Eval("FirstName") & " " & Eval("LastName")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:300px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label1" runat="server" Text='<%#Eval("Description")%>' CssClass="text8"/><asp:HiddenField ID="HF_InvID" Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
            <td align="center" valign="top" style="width:80px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label17" runat="server" Text='<%#Eval("password")%>' CssClass="text8"/></td>
            <td></td>
        </tr>
        </AlternatingItemTemplate>
    </asp:Repeater>
</table>
</asp:Panel>

<asp:Panel ID="PNL_PaySummary" Visible="false" runat="server">
<table id="summary" class="tablerow" cellpadding="0" cellspacing="0" style="width:100%; position:relative; border-bottom: solid 1px darkgray">
<asp:Repeater ID="RPT_Pay" runat="server">
    <ItemTemplate>
        <tr>
            <td align="left" valign="top" style="width:100px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><span id='<%# "sl_" & Eval("Pay_ID")%>' class="blacklink8"><%#Eval("DateString")%></span></td>
            <td align="left" valign="top" style="width:200px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label27" runat="server" Text='<%#Eval("Method_Name")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="width:170px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label38" runat="server" style="text-align: right" Text='<%#Eval("Name")%>' CssClass="text8"/></td>
            <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label2" runat="server" style="text-align: right" Text='<%#Eval("Total", "{0:#,###.00}")%>' CssClass="text8"/></td>
            <td align="center" valign="top" style="width:80px; border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label14" runat="server" style="text-align: right" Text='<%#Eval("Currency")%>' CssClass="text8"/></td>
            <td align="center" valign="top" style="border-right: solid 1px lightgray; padding: 7px 5px 7px 5px; width: 80px"><asp:Label ID="Label4" runat="server" Text='<%#Eval("Status")%>' CssClass="text8"/></td>
            <td align="left" valign="top" style="border-right: solid 1px lightgray; padding: 7px 5px 7px 5px"><asp:Label ID="Label1" runat="server" Text='<%#Eval("Description")%>' CssClass="text8"/><asp:HiddenField ID="HF_PayID" Value='<%# Eval("Pay_ID")%>' runat="server" /></td>
            <td></td>
        </tr>
        </ItemTemplate>
    </asp:Repeater>
</table>
</asp:Panel>


<asp:Panel ID="PNL_History" visible="false" runat="server">
        <table cellpadding="7" cellspacing="0" id="rechistory" width="100%" style="border-radius: 5px 5px; padding: 0px 0px 5px 0px; background-color:white;">
            <tr>
                    <td align="center" style="padding: 0px 0px 0px 0px; border: solid 1px lightgray; border-top-left-radius: 3px 3px; border-top-right-radius: 3px 3px">
                        <table cellpadding="3" cellspacing="0" width="100%" style="background-color: lightgrey; background-size:100%; border-top-left-radius: 3px 3px; border-top-right-radius: 3px 3px">
                            <tr>
                                <td align="center" style="padding: 5px 0px 5px 0px"><asp:Label ID="Label5" CssClass="text" Text="Receiving History" Width="210px" runat="server" /></td>
                            </tr>
                        </table>
                </td>
            </tr>
            <tr>
            <td align="center" style="padding: 0px 0px 0px 0px; background-color: white; border-left: solid 1px lightgray; border-bottom: solid 1px lightgray; border-right: solid 1px lightgray;">
            <table cellpadding="3" cellspacing="0" width="100%" style="padding:5px 3px 5px 3px">
                <tr>
                    <td style="width:50px" align="left"><asp:Label ID="Label7" CssClass="text9" Text="Number" Width="50px" runat="server" /></td>
                    <td align="left"><asp:Label ID="Label9" CssClass="text9" Text="PDF" Width="50px" runat="server" /></td>
                </tr>
                <asp:Repeater ID="RPT_History" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td align="left"><asp:Label ID="LBL_RecNum" CssClass="blacklink8" Text='<%# Eval("Name")%>' Width="50px" runat="server" /><asp:HiddenField ID="HF_RecID"  Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
                            <td style="width:10px"><img id='<%# "img_" & Eval("File_Name") %>' style="cursor:pointer" src='<%# Eval("Image")%>' alt="" /></td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </td>
        </tr>
    </table>
    <table cellpadding="7" id="invhistory" cellspacing="0" width="100%" style="border-radius: 5px 5px; padding: 10px 0px 5px 0px; background-color:white;">
    <tr>
        <td id="histheaderinv" align="center" class="sideheader">
            <asp:Label ID="Label20" CssClass="sideheadertext"  Text="INVOICE HISTORY" runat="server" />
        </td>
    </tr>
    <tr>
    <td align="center" style="padding: 0px 0px 0px 0px; background-color: white; ">
    <table cellpadding="0" cellspacing="0" width="100%" style="padding:0px 0px 5px 0px;">
        <tr style="background-color:#dddddd">
            <td align="left" style="height:20px; padding-left:5px"><asp:Label ID="Label21" CssClass="text9" style="color:black; font-size:8pt" Text="Date" Width="65px" runat="server" /></td>
            <td style="width:85px; padding-left:5px; border-left: solid 1px white; border-right: solid 1px white" align="left"><asp:Label ID="Label13" CssClass="text9" style="color:black; font-size:8pt"  Text="Invoice No" Width="85px" runat="server" /></td>
            <td align="center"><asp:Label ID="Label8" CssClass="text9" style="color:black; font-size:8pt; padding-right:5px" Width="30px" Text="Doc" runat="server" /></td>
        </tr>
        <asp:Repeater ID="RPT_InvHist" runat="server">
            <ItemTemplate>
                <tr>
                    <td align="left" valign="top" style="padding: 3px 0px 3px 5px"><asp:Label ID="Label6" CssClass="text9" style="font-size:8pt;" Text='<%# Eval("Doc_Date", "{0:yyyy-MM-dd}")%>' Width="70px" runat="server" /></td>
                    <td align="left" valign="top" style="padding: 3px 0px 3px 5px; border-left: solid 1px lightgray; border-right: solid 1px lightgray"><asp:Label ID="LBL_PurchNum" CssClass="blacklink8" style="font-size:8pt" Text='<%# Eval("Doc_No")%>' Width="110px" runat="server" /><asp:HiddenField ID="HF_PurchID" Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
                    <td align="center" valign="top" style="width:30px"><img id='<%# "img_" & Eval("File_Name") %>' style="cursor:pointer" src='<%# Eval("Image")%>' alt="" /></td>
                </tr>
            </ItemTemplate>
        </asp:Repeater>
    </table>
        </td>
        </tr>
    </table>
</asp:Panel>


<asp:Panel ID="PNL_ApplyCr" Visible="false" Style="padding: 4px 4px 4px 4px" runat="server">
            <table align="center" cellpadding="4" cellspacing="0" style="background-color: white;">
                <tr>
                    <td align="left">
                        <asp:Label ID="LBL_VendorName" CssClass="text11" Text="" runat="server" />
                    </td>
                    <td align="left" style="padding:15px">
                        <asp:Label ID="Label18" CssClass="text11" Text="Total Applied" runat="server" />
                    </td>
                    <td align="right" style="padding:15px">
                        <asp:Label ID="LBL_TotalApplied" CssClass="text11" Text="" runat="server" />
                    </td>
                    <td align="right">
                        <asp:Label ID="LBL_CloseApplyCr" CssClass="close" Text="X" runat="server" /></td>
                </tr>
                <tr>
                    <td colspan="4" style="padding-top: 12px">
                        <table align="center" cellpadding="3" cellspacing="0" width="675px" style="border-radius: 3px 3px; border-bottom: none">
                            <tr>
                                <td align="left" class="tabletop2" style="border-right: solid 1px lightgray; border-top-left-radius: 4px 4px; width: 115px; padding: 5px 5px 5px 5px"><asp:Label ID="Label3" Text="Invoice No" runat="server" /></td>
                                <td align="left" class="tabletop2" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label12" Text="Invoice Date" runat="server" /></td>
                                <td align="right" class="tabletop2" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label15" Text="Invoice Total" runat="server" /></td>
                                <td align="right" class="tabletop2" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label16" Text="Balance Due" runat="server" /></td>
                                <td align="right" class="tabletop2" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px; border-top-right-radius: 5px 5px;"><asp:Label ID="Label14" Text="Applied Amount" runat="server" /></td>
<%--                                <td align="right" style="display: none"><asp:Label ID="Label18" CssClass="text8" Text="Const Amount" runat="server" /></td>
                                <td align="right" style="display: none"><asp:Label ID="Label17" CssClass="text8" Text="Const Balance" runat="server" /></td>--%>
                                <td align="right" style="border-right: none; border-top: none; width: 15px"><asp:Label ID="Label19" CssClass="text8" Text="" runat="server" /></td>
                            </tr>
                        </table>
                        <div id="PNL_Scroll" style="overflow-y: auto; overflow-x: hidden; height: auto; max-height: 300px; border-bottom: none">
                            <table id="tbl" class="tablerow5" cellpadding="3" cellspacing="0" width="650px" style="border-bottom: solid 1px darkgray">
                                <asp:Repeater ID="RPT_ApplyCr" runat="server">
                                    <ItemTemplate>
                                        <tr id="row">
                                            <td id="cell" align="left" style="border-right: solid 1px lightgray; width: 115px; padding: 5px 5px 5px 5px"><asp:Label ID="LBL_InvNo" CssClass="text8" Text='<%# Eval("Doc_No")%>' runat="server" /></td>
                                            <td align="left" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px"><asp:Label ID="LBL_InvDate" CssClass="text8" Text='<%# Eval("Doc_Date", "{0:yyyy-MM-dd}")%>' runat="server" /></td>
                                            <td align="right" style="border-right: solid 1px lightgray; width: 100px; padding: 5px 5px 5px 5px"><asp:Label ID="LBL_InvTotal" CssClass="text8" Text='<%# Eval("Total", "{0:#,###.00}")%>' runat="server" /></td>
                                            <td align="right" style="font-size: 8.25pt; width: 100px; padding: 5px 5px 5px 5px; border-right: solid 1px lightgray;"><asp:Label ID="LBL_Balance" CssClass="text8" Text='<%# Eval("Balance", "{0:#,###.00}")%>' runat="server" /><asp:HiddenField ID="HF_InvID" Value='<%# Eval("Doc_ID")%>' runat="server" /></td>
                                            <td align="right" style="width: 100px; padding: 5px 5px 5px 5px"><asp:TextBox ID="LBL_Amount" Class="tbinput" autocomplete="off" Style="text-align: right" Text="" runat="server" /></td>
<%--                                            <td align="right" style="display: none"><asp:Label ID="LBL_ConstAmount" CssClass="text8" Text='<%# Eval("Amount")%>' runat="server" /></td>
                                            <td align="right" style="display: none"><asp:Label ID="LBL_ConstBalance" CssClass="text8" Text='<%# Eval("Balance")%>' runat="server" /></td>--%>
                                        </tr>
                                    </ItemTemplate>
                                </asp:Repeater>
                            </table>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" align="right">
                        <asp:Button ID="BTN_ApplyCr" CssClass="buttonnew" Text="APPLY INVOICE" Width="125px" Height="25px" runat="server" /></td>
                </tr>
            </table>
        </asp:Panel>

<asp:Panel ID="PNL_Apply" visible="false" style="padding:8px 8px 8px 8px" runat="server">
    <table cellpadding="4" cellspacing="0" style="background-color:white;">
        <tr>
            <td align="left">
                <asp:Label ID="LBL_VendorName2" CssClass="text11" Text="" runat="server" />
                &nbsp;&nbsp;&nbsp;<asp:Label ID="LBL_TotalDue3" CssClass="text9" Text="Total Due: " runat="server" />&nbsp;&nbsp;&nbsp;<asp:Label ID="LBL_TotalDue2" CssClass="text9" Text="" runat="server" />
                &nbsp;&nbsp;&nbsp;<asp:Label ID="LBL_Remaining3" CssClass="text9" Text="Remaining: " runat="server" />&nbsp;&nbsp;&nbsp;<asp:Label ID="LBL_Remaining2" CssClass="text9" Text="" runat="server" />
            </td> 
            <td align="right"><asp:Label ID="LBL_CloseApply" CssClass="close" Text="X" runat="server" /></td> 
        </tr> 
        <tr>
            <td colspan="2" style="padding-top:12px">
                <table cellpadding="3" cellspacing="0" style="border: solid 1px darkgray; width:700px; border-top-left-radius: 3px 3px; border-top-right-radius: 3px 3px">
                    <tr>
                        <td align="left" style="width:115px; padding: 5px 5px 5px 5px"><asp:Label ID="Label22" CssClass="text8" Text="Invoice No" runat="server" /></td> 
                        <td align="center" style="width:100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label23" CssClass="text8" Text="Invoice Date" runat="server" /></td> 
                        <td align="center" style="width:100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label11" CssClass="text8" Text="Invoice Age" runat="server" /></td> 
                        <td align="center" style="width:100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label24" CssClass="text8" Text="Invoice Total" runat="server" /></td> 
                        <td align="center" style="width:100px; padding: 5px 5px 5px 5px"><asp:Label ID="Label25" CssClass="text8" Text="Balance Due" runat="server" /></td> 
                        <td align="center" style="width:125px; padding: 5px 5px 5px 5px"><asp:Label ID="Label26" CssClass="text8" Text="Payment Amount" runat="server" /></td> 
                    </tr>
                </table>
                <div id="divapply" style="max-height:400px; overflow:auto">
                <table id="tableapply" class="tablerow" cellpadding="3" cellspacing="0" style="border-left: solid 1px darkgray; width:100%;">
 
                </table>
                </div>
                <table cellpadding="3" cellspacing="0" style="width:700px; ">
                    <tr>
                        <td align="right" style="border-top:solid 1px lightgray; padding-top:9px"><asp:Label ID="LBL_Pay" CssClass="text8" Text="Total to Pay:" runat="server" /></td>
                        <td align="right" style="border-top:solid 1px lightgray; width:125px; padding: 9px 5px 5px 4px"><asp:Label ID="LBL_TotalPay" CssClass="text8" Text="" runat="server" /></td>
                    </tr>
                </table>
                    <asp:Repeater ID="RPT_Apply" runat="server">
                        <ItemTemplate>
                            <asp:HiddenField ID="HF_ApplyLines" Value='<%# Eval("Doc_No") & "~" & Eval("Doc_Date", "{0:yyyy-MM-dd}") & "~" & Eval("Age") & "~" & Eval("Amount", "{0:#,###.00}") & "~" & Eval("Balance", "{0:#,###.00}") & "~" & Eval("PayAmount", "{0:#,###.00}") & "~" & Eval("Document_ID")%>' runat="server" />
                        </ItemTemplate>
                    </asp:Repeater>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="right"><asp:Button ID="BTN_Apply" CssClass="button" Text="Apply Payments" Width="125px" Height="25px" runat="server" /></td>
        </tr>
    </table>

</asp:Panel>

<asp:Panel ID="PNL_PrintReports" Visible="false" runat="server">
    <asp:HiddenField ID="HF_PrintTitle" runat="server" />
    <asp:HiddenField ID="HF_PrintHeader" runat="server" />
    <asp:HiddenField ID="HF_PrintHeaderOnce" runat="server" />
    <asp:HiddenField ID="HF_PrintPagePadding" Value ="15~10~15~10" runat="server" />
        <asp:Repeater ID="RPT_PrintReports" runat="server">
            <ItemTemplate>
                <asp:HiddenField ID="HF_PrintLines" Value='<%# Eval("Style1") & "~" & Eval("Field1") & "~" & Eval("Style2") & "~" & Eval("Field2") & "~" & Eval("Style3") & "~" & Eval("Field3") & "~" & Eval("Style4") & "~" & Eval("Field4") & "~" & Eval("Style5") & "~" & Eval("Field5") & "~" & Eval("Style6") & "~" & Eval("Field6") & "~" & Eval("Style7") & "~" & Eval("Field7") & "~" & Eval("Style8") & "~" & Eval("Field8") & "~" & Eval("Style9") & "~" & Eval("Field9") & "~" & Eval("Style10") & "~" & Eval("Field10") & "~" & Eval("Style11") & "~" & Eval("Field11") & "~" & Eval("Style12") & "~" & Eval("Field12") & "~" & Eval("Style13") & "~" & Eval("Field13") & "~" & Eval("Style14") & "~" & Eval("Field14") & "~" & Eval("Style15") & "~" & Eval("Field15")%>' runat="server" />
            </ItemTemplate>
        </asp:Repeater>
</asp:Panel>

<asp:Panel ID="PNL_PrintPO" Visible="false" runat="server">
    <asp:HiddenField ID="HF_POPrintTitle" runat="server" />
    <asp:HiddenField ID="HF_POPrintHeader" runat="server" />
    <asp:HiddenField ID="HF_POPrintPagePadding" Value ="25~45~25~45" runat="server" />
        <asp:Repeater ID="RPT_PrintPO" runat="server">
            <ItemTemplate>
                <asp:HiddenField ID="HF_PrintLines" Value='<%# Eval("Style1") & "~" & Eval("Field1") & "~" & Eval("Style2") & "~" & Eval("Field2") & "~" & Eval("Style3") & "~" & Eval("Field3") & "~" & Eval("Style4") & "~" & Eval("Field4") & "~" & Eval("Style5") & "~" & Eval("Field5") & "~" & Eval("Style6") & "~" & Eval("Field6") & "~" & Eval("Style7") & "~" & Eval("Field7") & "~" & Eval("Style8") & "~" & Eval("Field8") & "~" & Eval("Style9") & "~" & Eval("Field9") & "~" & Eval("Style10") & "~" & Eval("Field10") & "~" & Eval("Style11") & "~" & Eval("Field11") & "~" & Eval("Style12") & "~" & Eval("Field12") & "~" & Eval("Style13") & "~" & Eval("Field13") & "~" & Eval("Style14") & "~" & Eval("Field14") & "~" & Eval("Style15") & "~" & Eval("Field15")%>' runat="server" />
            </ItemTemplate>
        </asp:Repeater>
    <asp:HiddenField ID="DF_POPrintFooter" runat="server" />
</asp:Panel>

<asp:Panel ID="PNL_Categories" Visible="false" runat="server">
    <table cellpadding="3" cellspacing="0" style="width:auto">
        <tr>
            <td align="center"><asp:Label ID="LBL_ActNo" CssClass="text9" Text="Choose GL Account No: " runat="server" />&nbsp;&nbsp;&nbsp;<asp:DropDownList ID="DDL_Account" Width="250px" CssClass="text9box" runat="server" /></td>
        </tr>
        <tr>
            <td><div id="div_innercat"></div></td>
        </tr>
    </table>
</asp:Panel>

<asp:Panel ID="PNL_InsideCategories" Visible="false" runat="server">
    <table cellpadding="3" cellspacing="0" style="width:auto; margin-left:auto; margin-right:auto">
        <tr>
            <td align="center"><asp:Label ID="Label28" CssClass="text9" Text="Delete" runat="server" /></td>
            <td align="center"><asp:Label ID="Label29" CssClass="text9" Text="Category Name" runat="server" /></td>
            <td align="center"><asp:Label ID="Label30" CssClass="text9" Text="Active" runat="server" /></td>
        </tr>
        <tr>
            <td align="center"></td>
            <td align="center"><asp:TextBox ID="TB_NewCat" CssClass="text9box" Width="130px" Text="" runat="server" /></td>
            <td align="center"><asp:Button ID="BTN_SaveCat" CssClass="buttonnew" Width="100px" Height="25px" Text="ADD" runat="server" /></td>
        </tr>
        <asp:Repeater ID="RPT_Categories" runat="server">
            <ItemTemplate>
                <tr>
                    <td align="center"><asp:Image ID="IMG_DelCat" ImageUrl="images/delete.png" runat="server" /></td>
                    <td align="center"><asp:TextBox ID="TB_NewCat" CssClass="text9box" Width="130px" Text='<%# Eval("Name")%>' runat="server" /></td>
                    <td align="center"><asp:Checkbox ID="CB_CatActive" CssClass="text9" runat="server" /></td>
                </tr>
            </ItemTemplate>
        </asp:Repeater>
    </table>

</asp:Panel>

<asp:Panel ID="PNL_XMLReport" Visible="false" runat="server" >
    <asp:HiddenField ID="HF_XML" runat="server" />
</asp:Panel>

        <asp:Panel ID="PNL_Report" Visible="false" runat="server" >
            <asp:HiddenField ID="HF_Date_Today" runat="server" />
            <div id="customers" style="display:none">
                <asp:Repeater ID="RPT_Cust" runat="server">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_Print_CustCur" Value ='<%# Eval("CURRENCY") %>' runat="server" />
                        <asp:HiddenField ID="HF_Print_CustName" Value ='<%# Eval("Name") %>' runat="server" />
                        <asp:HiddenField ID="HF_Print_CustID" Value ='<%# Eval("Cust_ID") %>' runat="server" />
                    </ItemTemplate>
                </asp:Repeater>
            </div>
            <div id="vendors" style="display:none">
                <asp:Repeater ID="RPT_Vend" runat="server">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_Print_VendCur" Value ='<%# Eval("CURRENCY") %>' runat="server" />
                        <asp:HiddenField ID="HF_Print_VendName" Value ='<%# Eval("Name") %>' runat="server" />
                        <asp:HiddenField ID="HF_Print_VendID" Value ='<%# Eval("Cust_ID") %>' runat="server" />
                    </ItemTemplate>
                </asp:Repeater>
            </div>

           <%-- Includes Category, Type and Language--%>

            <td class="auto-style2">
                  <table id="table_print">                         
                <tr style ="width: 100px; float: center;">
                    <td valign="top" align="left" bgcolor="#6699FF" class="auto-style2">
                          
                    <asp:Label ID="LBL_ReportTitle" CssClass="title"  Text="Print Reports" runat="server" ForeColor="White" />
               </td>
                 </tr>
                <tr  style="width: 0px; float: left;">
                    <td valign="top" align="left" style="padding: 18px 0px;" class="auto-style2">
                        <asp:Label Text="Category: " CssClass="text9" runat="server" style="padding: 18px 0px 0px 13px;"  />

                        <td valign="top" align="left" style="padding: 18px 0px;" class="auto-style2" width="400">                            
                        <asp:DropDownList ID="DDL_Print_Category" CssClass="text9box" align="left" runat="server" style="padding: 3px 10px;" />
                            <td valign="top" align="left" style="padding: 18px 0px;" class="auto-style2">
                        <asp:Label ID="LBL_T_LangLabel" Text="Language: " CssClass="text9" align= "right" runat="server" style="padding:18px 0px 0px 90px;"/>
                           </td>
                        </td>
                            <td valign="top" align="right" style="padding: 18px 0px;" class="auto-style2" width="400">
                            <asp:DropDownList ID="DDL_Print_Language" CssClass="text9box" runat="server">
                            <asp:ListItem Text="English" Value="0"></asp:ListItem>
                            <asp:ListItem Text="Español" Value="1"></asp:ListItem>
                        </asp:DropDownList>
                            </td>
                            
                    </tr>

                      <tr>
                          <td id="typestandard" align="left" class="auto-style2" style="padding: 18px 15px;" valign="top">
                              <asp:Label runat="server" CssClass="text9" Text="Type: " />
                              <asp:DropDownList ID="DDL_Print_Report" runat="server" CssClass="text9box" Width="200px">
                              </asp:DropDownList>
                          </td>
                      </tr>

                <tr>
                   <td id= "typeMulti" valign="top" align="left" style="padding: 18px 15px; display: none" class="auto-style2">
                       <asp:Label Text="Type: " CssClass="text9" runat="server"/>
                       <asp:DropDownList ID="DDL_Print_MultiPeriod" runat="server" CssClass="text9box" Width="200px">
                               <asp:ListItem Text="Balance Sheet" Value="11" />
                               <asp:ListItem Text="Income Statement" Value="22" />                                        
                       </asp:DropDownList>   
                       
                </td>                           
                </tr>
                    </table>
            </td>
                
         
      
                <%--Includes Denomination, Detail Level, Round, Show Zero's, Show %, Show Act No--%>
 
                    <tr>
                    <td colspan="4" class="">
                        <hr />
                        <table id="table_general2" >
                            <tr id="DetailReport" style="display: none;">
                                <td valign="top" align="right"  style="padding: 18px 15px;">
                                    <asp:Label ID="LBL_T_23" CssClass="text9" Text="Account No." runat="server" /><span>
                                </td>
                                 <td valign="top" align="left"  style="padding: 15px 5px;">
                                    <asp:TextBox ID="TB_Print_AccNo" Style="text-align: center;" CssClass="text9box" Width="50px" runat="server" /></span>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top" align="left" style="padding: 18px 5px;">
                                    <asp:Label ID="Label31" Text="Denomination Level:" CssClass="text9" runat="server" />
                                </td>
                                 <td valign="top" align="left" style="padding: 15px 5px;">
                                    <asp:DropDownList ID="DDL_Print_Denomination" runat="server" CssClass="text9box" Width="100px"></asp:DropDownList>
                                </td>
                                <td valign="top" align="right" style="padding: 15px 0px;">
                                    <asp:Label ID="LBL_T_19" CssClass="text9" Text="Round: " runat="server" Width= "50px"/>
                                   
                                    <asp:CheckBox runat="server" ID="CB_Print_Round" AutoPostBack="false" Text="" />
                                 </td>
                                
                                 <td valign="top" align="right" style="padding: 15px 0px;">
                                    <asp:Label ID="LBL_T_20" CssClass="text9" Text="Show Zero's: " runat="server" Width= "100px" />
                                     
                                    <asp:CheckBox runat="server" ID="CB_Print_ShowZeros" AutoPostBack="false" Text="" />
                                </td>
                            </tr>

                            <tr  id="td_detail">
                                <td valign="top" align="left" style="padding: 18px 5px;">
                                    <asp:Label ID="LBL_T_DetailLabel" Text="Detail Level:" CssClass="text9" runat="server" />
                                </td>
                                 <td valign="top" align="left" style="padding: 15px 5px;">
                                    <asp:DropDownList ID="DDL_Print_Level" CssClass="text9box" runat="server" Height="24px" Width="100px">
                                        <asp:ListItem Text="1" Value="1"></asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                                <td valign="top" align="right" style="padding: 15px 0px;" id="Show_per">
                                    <asp:Label ID="Label32" CssClass="text9" Text="Show %: " runat="server" Width="100px" />
                                    
                                    <asp:CheckBox ID="CB_Print_ShowPer" Text="" AutoPostBack="false" runat="server" />
                                </td>
                                <td valign="top" align="right" style="padding: 15px 0px;" >
                                    <asp:Label ID="Label33" CssClass="text9" Text="Show Act No: " runat="server" Width="100px" />
                                    
                                    <asp:CheckBox ID="CB_Print_Accno" CssClass="text9" Text="" AutoPostBack="false" runat="server" />
                                </td>
                            </tr>
                        </table>
                        
                    </td>
                    </tr>
                
             
                


                
               <%-- Includes Type and Date--%>
                    <tr>
                    <td class="auto-style2" >
                        <hr />
                        <table id="table_general1">
                           
                            <tr>
                                <td class="auto-style2" valign="top" align="left" style="padding: 15px 15px;">

                                    <asp:Label Text="Date: " CssClass="text9" runat="server"  />
                                    <asp:TextBox ID="TB_Print_Date1" Style="text-align: center" CssClass="text9box" Width="100px" runat="server" /><span id="PrintDate2Span">
                                    <asp:TextBox ID="TB_Print_Date2" Style="text-align: center" CssClass="text9box" Width="100px" runat="server" /></span>
                                </td>
                               
                            </tr>
                            <tr>
                                <td class="auto-style2" valign="top" align="left" style="padding: 15px 15px;">
                                    </td>
                                
                            </tr>
                           
                        
                        <table id="table_MultiPeriod" style="display:none;">

                            <tr>
                                <td valign="top" align="left" style="padding: 15px 15px;">
                                    <asp:Label Text="Period: " CssClass="text9" runat="server" />
                                    <asp:DropDownList ID="DDL_Print_Period" runat="server" CssClass="text9box" Width="150px"></asp:DropDownList>
                                </td>
                                
                                <td id="dropdownlist" valign="top" align="right" style="padding: 15px 55px;">
                                    <asp:Label Text="No. of Periods: " CssClass="text9" runat="server" />
                                    <asp:DropDownList ID="DDL_Print_Previous" runat="server" CssClass="text9box">
                                        <asp:ListItem Text="2 Years" Value="2" /> 
                                        <asp:ListItem Text="3 Years" Value="3" />
                                        <asp:ListItem Text="4 Years" Value="4" />
                                        <asp:ListItem Text="5 Years" Value="5" />
                                      </asp:DropDownList>
                                </td> 

                                 <td id="MonthlySelector" valign="top" align="left" style="padding: 15px 15px;">
                                    <asp:Label Text="Date: " CssClass="text9" runat="server" />
                                    <asp:TextBox ID="TB_Print_Date11" Style="text-align: center" CssClass="text9box" Width="85px" runat="server" /><span id="PrintDate2Span1">&nbsp;-&nbsp;
                                    <asp:TextBox ID="TB_Print_Date22" Style="text-align: center" CssClass="text9box" Width="85px" runat="server" /></span>
                                </td>

                                 <td id= "QuarterlySelector1" valign="top" align="left" style="padding: 15px 15px;">
                                    <asp:Label Text="Year: " CssClass="text9" runat="server" />
                                    <asp:DropDownList ID="DDL_Print_Quarter" CssClass="text9box" runat="server" />
                                </td>     

<%--                                   <td id="QuarterlySelector2" valign="top" align="left" style="padding: 15px 10px;">
                                    <asp:CheckBox ID="CB_Q1" CssClass="text9" Text="Sept-Nov" runat="server"/>
                                    <asp:CheckBox ID="CB_Q2" CssClass="text9" Text="Dec-Feb" runat="server"/>                                     
                                    <asp:CheckBox ID="CB_Q3" CssClass="text9" Text="Mar-May" runat="server"/>
                                    <asp:CheckBox ID="CB_Q4" CssClass="text9" Text="Jun-Aug" runat="server"/>
                                </td>--%>
                                
                                <td id="YearlySelector" valign="top" align="left" style="padding: 15px 15px;">
                                    <asp:Label Text="Year: " CssClass="text9" runat="server" />
                                    <asp:DropDownList ID ="DDL_Print_YearFrom" CssClass="text9" runat="server" /><span>&nbsp;-&nbsp;
                                    <asp:DropDownList ID="DDL_Print_YearTo" CssClass="text9" runat="server" /></span>
                                </td>
                            </tr>

           

                            <tr id ="MonthToMonthSelector">
                              <td valign="top" align="left" style="padding: 15px 15px;">
                                      <asp:Label Text="Select Month: " CssClass="text9" runat="server" />
                                       <asp:DropDownList ID="DDL_Print_P" CssClass="text9box" runat="server" /><br/> <br/>
                                    </td>
                                    </tr>


<%--                            <tr id="QuarterlySelector1">
                                                      
                            </tr>    --%>
                            
                            <tr id= "QuarterlySelector2">
                                <td valign="top" align="left" style="padding: 15px 10px;">
                                    <asp:CheckBox ID="CB_Q1" CssClass="text9" Text="Sept-Nov" runat="server"/>
                                    <asp:CheckBox ID="CB_Q2" CssClass="text9" Text="Dec-Feb" runat="server"/>                                     
                                    <asp:CheckBox ID="CB_Q3" CssClass="text9" Text="Mar-May" runat="server"/>
                                    <asp:CheckBox ID="CB_Q4" CssClass="text9" Text="Jun-Aug" runat="server"/>
                                </td>
                        </tr>
                                

                            <tr id="QuarterToQuarterSelector">
                                <td valign="top" align="left" style="padding: 15px 15px;">
                                    <asp:Label Text="Select Quarter: " CssClass="text9" runat="server" />
                                    <asp:DropDownList ID="DDL_Print_Q" CssClass="text9" runat="server"/><br/> <br/>
                                </td>
                            </tr>
                            </table>
                        </table>                       
                    </td>
                        </tr>

                <tr>
                    <td class="auto-style2">
                        <table id="table_sales" style="display:none;">
                            <tr id="Report_AR">
                                <td align="left" style="padding: 15px 15px;">
                                    <asp:Label runat="server" Text="Type: " />
                                    <asp:DropDownList ID="DDL_Print_Details" runat="server" CssClass="text9box" Width="120px">
                                        <asp:ListItem Text="Summary" Value="Summary" />
                                        <asp:ListItem Text="Details" Value="Details" />
                                        <asp:ListItem Text="Report" Value="Report" />
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" style="padding: 15px 15px; " valign="top">
                                    <asp:Label runat="server" Text="Currency: " />
                                    <asp:DropDownList ID="DDL_Print_Currency" runat="server" CssClass="text9box" Width="90px" />
                                </td>
                            </tr>
                            <tr>
                                <td align="left" style="padding: 15px 15px;" valign="top">
                                    <asp:Label runat="server" Text="Date: " />
                                    <asp:TextBox ID="Date_Print_From" runat="server" class="text9box" style="width:100px; text-align:center" />
                                    <span id="Date_DTSpan">&nbsp;-&nbsp;
                                    <asp:TextBox ID="Date_Print_To" runat="server" class="text9box" style="width:100px; text-align:center" />
                                    </span></td>
                            </tr>
                            <tr>
                                <td id="td_customer" align="left" style="padding: 15px 15px; display:none" valign="top">
                                    <asp:Label runat="server" Text="Customer: " />
                                    <asp:DropDownList ID="DDL_Print_Customer" runat="server" class="text9box" style="width:200px;" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                
                        
            <table id="table button" style="float:right;">
                <tr>
                    <td style="width: 70px; padding-left: 10px; display:none;" align="left" class="noprint">
                        <asp:Button ID="BTN_Print_Export" CssClass="buttonnew" Width="85px" Height="25px" Text="EXPORT" runat="server" /></td>
                    <td style="width: 70px; padding-left: 10px" class="noprint">
                        <asp:Button ID="BTN_Print_Print" CssClass="buttonnew" Width="85px" Height="25px" Text="PRINT" runat="server" /></td>
                    <td style="width: 70px; padding-left: 10px" align="left" class="noprint">
                        <asp:Button ID="BTN_Print_Cancel" CssClass="buttonnew" Width="85px" Height="25px" Text="CANCEL" runat="server" /></td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
