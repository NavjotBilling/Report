<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AjaxPrintingCSharp.aspx.cs" Inherits="AjaxPrintingCSharp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Panel ID="PNL_PrintReports" Visible="false" runat="server">
            <asp:HiddenField ID="HF_PrintTitle" runat="server" />
            <asp:HiddenField ID="HF_PrintHeader" runat="server" />
            <asp:HiddenField ID="HF_PrintHeaderOnce" runat="server" />
            <asp:HiddenField ID="HF_PrintPagePadding" Value ="15~10~15~10" runat="server" />
                <asp:Repeater ID="RPT_PrintReports" runat="server">
                    <ItemTemplate>
                        <asp:HiddenField ID="HF_PrintLines" Value='<%# Eval("Style1") + "~" + Eval("Field1") + "~" + Eval("Style2") + "~" + Eval("Field2") + "~" + Eval("Style3") + "~" + Eval("Field3") + "~" + Eval("Style4") + "~" + Eval("Field4") + "~" + Eval("Style5") + "~" + Eval("Field5") + "~" + Eval("Style6") + "~" + Eval("Field6") + "~" + Eval("Style7") + "~" + Eval("Field7") + "~" + Eval("Style8") + "~" + Eval("Field8") + "~" + Eval("Style9") + "~" + Eval("Field9") + "~" + Eval("Style10") + "~" + Eval("Field10") + "~" + Eval("Style11") + "~" + Eval("Field11") + "~" + Eval("Style12") + "~" + Eval("Field12") + "~" + Eval("Style13") + "~" + Eval("Field13") + "~" + Eval("Style14") + "~" + Eval("Field14") + "~" + Eval("Style15") + "~" + Eval("Field15")%>' runat="server" />
                    </ItemTemplate>
                </asp:Repeater>
        </asp:Panel>

        <asp:Panel ID="PNL_Excel" runat="server" Visible="false">
            <table id="excel_table">             
                <asp:Repeater ID="RPT_Excel" runat="server">
                    <ItemTemplate>
                        <tr>
                            <td><asp:Label Text='<%#Eval("value0") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value1") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value2") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value3") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value4") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value5") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value6") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value7") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value8") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value9") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value10") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value11") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value12") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value13") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value14") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value15") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value16") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value17") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value18") %>' runat="server"></asp:Label></td>
                            <td><asp:Label Text='<%#Eval("value19") %>' runat="server"></asp:Label></td>
                        </tr>
                    </ItemTemplate>
                </asp:Repeater>
            </table>
        </asp:Panel>

        <asp:Panel ID="PNL_Ajax" runat="server" Visible="false">
            <asp:Panel ID="PNL_Details" runat="server" Visible="false">
                <asp:HiddenField ID="HF_TotalDet" runat="server" />
                <asp:Label ID="LBL_Total"  CssClass="text" style="font-size:11pt;" runat="server" />
                <table width="100%" cellpadding="0" cellspacing="0" style="margin-left:auto; margin-right:auto; position:relative"> 
                    <tr>
                        <td id="TOPVendNameDet" align="left" valign="top" class="tabletop2" style="width:260px;border-top-left-radius:5px 5px;"><asp:Label ID="Label3" runat="server"  Text="Customer"/></td>
                        <td id="td_age" align="center" valign="top" class="tabletop2" style="width:50px"><asp:Label ID="Label1" runat="server"  Text="Age"/></td>
                        <td id="td_date" align="center" valign="top" class="tabletop2" style="width:100px"><asp:Label ID="Label10" runat="server"  Text="Date"/></td>
                        <td id="td_invoice" align="left" valign="top" class="tabletop2" style="width:100px"><asp:Label ID="Label17" runat="server"  Text="Invoice No."/></td>
                        <td align="center" valign="top" class="tabletop2" style="width:110px"><asp:Label ID="Label8" runat="server"  Text="Total"/></td>
                        <td align="center" valign="top" class="tabletop2" style="width:110px"><asp:Label ID="Label11" runat="server"  Text="Current"/></td>
                        <td align="center" valign="top" class="tabletop2" style="width:110px"><asp:Label ID="Label12" runat="server"  Text="31-60"/></td>
                        <td align="center" valign="top" class="tabletop2" style="width:110px"><asp:Label ID="Label15" runat="server"  Text="61-90"/></td>
                        <td align="center" valign="top" class="tabletop2" style="width:110px"><asp:Label ID="Label16" runat="server"  Text="90+"/></td>
                    
                        <td class="tabletop2" style="border-top-right-radius:5px 5px;"><span></span></td> 
                    </tr>
                </table>

                <div id="PNL_Scroll_Det" style=" overflow:auto;  border-bottom: solid 1px lightgray">   
                <table id="summary_det" cellpadding="0" cellspacing="0" style="width:100%; position:relative; border-bottom:solid 1px lightgray" class="tablerowgreen">
                    <asp:Repeater ID="RPT_List_Details" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td id='<%# "NameField_" + Eval("Padding")%>' align="left" valign="top" style="width:260px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_CustName" runat="server" Height="20px" Text='<%#Eval("Name")%>' CssClass="text8"/></td>
                                <td id="td_age1" align="center" valign="top" style="width:50px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_Age" runat="server" Text='<%#Eval("Age")%>' CssClass="text8"/></td>
                                <td id="td_date1" align="center" valign="top" style="width:100px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_ARDate" runat="server" Text='<%#Eval("Document_Date", "{0:yyyy-MM-dd}")%>' CssClass="text8"/></td>
                                <td id="td_invoice1" align="left" valign="top" style="width:100px; border-right: solid 1px lightgray;" class="tablecell2">
                                    <asp:HyperLink ID="HL_invNO" Text ='<%# Eval("Doc_No")%>' NavigateUrl='<%# "~/ACC_SalesInv.aspx?docid=" + Eval("Doc_ID") %>' CssClass="blacklink8" runat="server" />
                                </td>
                                <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_APTotal" runat="server" style="text-align: center;" Text='<%#Eval("Total", "{0:#,###.00}")%>' CssClass="text8"/></td>
                                <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_APCurrent" runat="server" style="text-align: center" Text='<%#Eval("Current", "{0:#,###.00}")%>' CssClass="text8 tooltip"/></td>
                                <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_AP30" runat="server" style="text-align: center" Text='<%#Eval("AP30", "{0:#,###.00}")%>' CssClass="text8 tooltip"/></td>
                                <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_AP60" runat="server" style="text-align: center" Text='<%#Eval("AP60", "{0:#,###.00}")%>' CssClass="text8 tooltip"/></td>
                                <td align="right" valign="top" style="width:110px; border-right: solid 1px lightgray;" class="tablecell2"><asp:Label ID="LBL_AP90" runat="server" style="text-align: center" Text='<%#Eval("AP90", "{0:#,###.00}")%>' CssClass="text8 tooltip"/></td>
                                <td align="center" valign="top" style=" padding: 5px 5px 5px 5px">
                                    <asp:HiddenField ID="HF_Cust_ID" Value='<%#Eval("Cust_Vend_ID")%>' runat="server" />
                                </td>
                            </tr>
                        </ItemTemplate>                    
                    </asp:Repeater>    
                </table>
                </div>
            </asp:Panel>
        </asp:Panel>
    <div>
    
    </div>
    </form>
</body>
</html>
