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
    <div>
    
    </div>
    </form>
</body>
</html>
