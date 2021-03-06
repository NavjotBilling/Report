
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.IO
Imports System.Xml
Imports System.Threading
Imports System.Globalization

Partial Class AjaxPrinting
    Inherits System.Web.UI.Page

    Dim Conn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("ConnInfo"))
    Dim SQLCommand As New SqlCommand
    Dim DataAdapter As New SqlDataAdapter
    Dim DataAdapterNav As New SqlDataAdapter
    Dim DBTable As New DataTable

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Dim Action As String = Request.Form("action")

        'Printing Reports
        If Action = "BankRecPrint" Then PrintBankRec()
        If Action = "BankRecXML" Then XMLBankRec()
        If Action = "BalanceSheet" Then PrintBalance()
        If Action = "BalanceSheetXML" Then XMLBalance()
        If Action = "ProfitLoss" Then PickIncomeSheet()
        If Action = "ProfitLossXML" Then XMLProfitLoss()
        If Action = "ProfitLossXMLM2M" Then XMLProfitLossM2M()
        If Action = "DetailTrial" Then PrintDetailTrial()
        If Action = "DetailTrialChart" Then PrintDetailTrialChart()
        If Action = "DetailTrialXML" Then XMLDetailTrial()
        If Action = "SummaryTrail" Then PrintSummaryTrail()
        If Action = "SummaryTrailXML" Then XMLSummaryTrial()
        If Action = "SalesSummary" Then PrintSalesSummary()

        If Action = "IncStateMultiMonthly" Then PrintMonthlyIncStateMultiPer()
        If Action = "IncStateMultiMonth-to-Month" Then PrintMonthToMonthIncStateMultiPer()
        If Action = "IncStateMultiQuarterly" Then PrintQuarterlyIncStateMultiPer()
        If Action = "IncStateMultiQuarter-to-Quarter" Then PrintQuarterToQuarterIncStateMultiPer()
        If Action = "IncStateMultiYearly" Then PrintYearlyIncStateMultiPer()

        If Action = "BalSheetMultiMonthly" Then PrintMonthlyBalSheetMultiPer()
        If Action = "BalSheetMultiMonth-to-Month" Then PrintMonthToMonthBalSheetMultiPer()
        If Action = "BalSheetMultiQuarterly" Then PrintQuarterlyBalSheetMultiPer()
        If Action = "BalSheetMultiQuarter-to-Quarter" Then PrintQuarterToQuarterBalSheetMultiPer()
        If Action = "BalSheetMultiYearly" Then PrintYearlyBalSheetMultiPer()

        'If Action = "IncomeStatementSingle" Then PrintIncomeStatementSingle()
        If Action = "Report" Then Report()
        If Action = "ShowPanelReport" Then ShowPNLReport()

    End Sub
    Private Sub XMLBankRec()
        Dim Account_number1 As String = Request.Form("acct")
        Dim date1 As String = Request.Form("date1")
        Dim sort_param As String = Request.Form("sort_param")
        Dim updown As String = " DESC"
        If sort_param = "" Then sort_param = "Transaction_Date"

        If sort_param = "Debit Amount" Then
            sort_param = "Debit_Amount"
        End If

        If sort_param = "Credit Amount" Then
            sort_param = "Credit Amount"
        End If

        If sort_param = "Transaction Date" Then
            sort_param = "Transaction Date"
        End If

        If sort_param = "Transaction Date ▽" Then
            sort_param = "Transaction_Date"
            updown = " ASC"
        End If
        'Δ
        If sort_param = "Transaction Date ^" Then
            sort_param = "Transaction_Date"
            updown = " DESC"
        End If

        HF_PrintHeader.Value = "text-align:left; width:80px; font-size:8pt~~text-align:left; font-size:8pt~~text-align:right; width:200px; font-size:8pt~"

        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Bank Reconciliation<br/>As Of " + date1 + "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " </span><br/><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'>Transaction Date</span><span style='position: absolute; margin-left: 2in'>Memo</span><span style='position: absolute; margin-left: 4in'>Credit</span><span style='position: absolute; margin-left: 6in'>Debit</span><span style='position: absolute; margin-left: 7.8in'>Currency</span></div>"

        Dim COA, Bal, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "SELECT * FROM ACC_GL LEFT JOIN ACC_GL_Accounts on Acc_Gl.fk_Account_ID = ACC_GL_Accounts.Account_ID WHERE (fk_Account_ID = @Account_number) AND (locked = 0) AND Transaction_Date <= @date ORDER BY " + sort_param + updown
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@Account_number", Account_number1)
        SQLCommand.Parameters.AddWithValue("@date", date1)
        System.Diagnostics.Debug.WriteLine(SQLCommand.CommandText)
        DataAdapter.Fill(COA)

        COA.AcceptChanges()
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))

        Dim ds As New DataSet
        ds.Tables.Add(COA)

        Dim xmlData As String = ds.GetXml()

        System.Diagnostics.Debug.WriteLine(xmlData)
        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True

        Conn.Close()
    End Sub
    ' Sales/Purchases Report
    Private Sub PrintSalesSummary()

        Dim Language As Integer = Request.Form("language")
        Dim FromDate As String = Request.Form("date1")
        Dim ToDate As String = Request.Form("date2")
        Dim Currency As String = Request.Form("cur")
        Dim Type As String = Request.Form("type")
        Dim Percentage As String = Request.Form("Percentage")
        Dim TotalTitle As String = ""
        Dim PercentageHeader As String = ""
        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Language = 0 Then
            If Percentage = "on" Then PercentageHeader = "~text-align:right; font-size:9pt; width:90px; font-weight:bold~Percentage(%)"
            If Type = "P" Then
                If Request.Form("Ac") <> "on" Then
                    HF_PrintHeader.Value = "text-align:left; width:0px; font-size:9pt; font-weight:bold~~text-align:left; font-size:9pt; font-weight:bold~Vendor Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)" & PercentageHeader
                Else
                    HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Vendor Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)" & PercentageHeader
                End If

                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Purchase Summary Report<br/>From " & FromDate & " to " & ToDate & "<br/>Currency: " & Currency & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
                TotalTitle = "Total Purchases"
            End If

            If Type = "S" Then
                If Request.Form("Ac") <> "on" Then
                    HF_PrintHeader.Value = "text-align:left; width:0px; font-size:9pt; font-weight:bold~~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)" & PercentageHeader
                Else
                    HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)" & PercentageHeader
                End If

                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Sales Summary Report</br>From " & FromDate & " to " & ToDate & "<br/>Currency: " & Currency & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
                TotalTitle = "Total Sales"
            End If
        ElseIf Language = 1 Then
            If Percentage = "on" Then PercentageHeader = "~text-align:right; font-size:9pt; width:90px; font-weight:bold~Porcentaje(%)"
            If Type = "P" Then
                If Request.Form("Ac") <> "on" Then
                    HF_PrintHeader.Value = "text-align:left; width:0px; font-size:9pt; font-weight:bold~~text-align:left; font-size:9pt; font-weight:bold~Nombre Del Vendedor~text-align:right; font-size:9pt; width:180px; font-weight:bold~Las Ventas Netas ($)" & PercentageHeader
                Else
                    HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~Número~text-align:left; font-size:9pt; font-weight:bold~Nombre Del Vendedor~text-align:right; font-size:9pt; width:180px; font-weight:bold~Las Ventas Netas ($)" & PercentageHeader
                End If

                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Informe De Resumen De Compra<br/>Desde " & FromDate & " a " & ToDate & "<br/>Moneda: " & Currency & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
                TotalTitle = "Total De Compras"
            End If

                If Type = "S" Then
                If Request.Form("Ac") <> "on" Then
                    HF_PrintHeader.Value = "text-align:left; width:0px; font-size:9pt; font-weight:bold~~text-align:left; font-size:9pt; font-weight:bold~Nombre Del Cliente~text-align:right; font-size:9pt; width:180px; font-weight:bold~Las Ventas Netas ($)" & PercentageHeader
                Else
                    HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~Número~text-align:left; font-size:9pt; font-weight:bold~Nombre Del Cliente~text-align:right; font-size:9pt; width:180px; font-weight:bold~Las Ventas Netas ($)" & PercentageHeader
                End If

                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Informe De Resumen De Ventas</br>Desde " & FromDate & " a " & ToDate & "<br/>Moneda: " & Currency & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
                TotalTitle = "Ventas Totales"
            End If
        End If

        Dim Sales, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        Dim where As String = " Currency = '" & Currency & "' and"
        If Currency = "CAD" Then where = " (Currency = '' or Currency = '" & Currency & "') and"
        If Currency = "ALL" Then where = ""

        If Type = "P" Then SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency, (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PINV') as Total, (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PINV') as Tax, (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PC') as TotalC, (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PC') as TaxC from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where " & where & "  piv.Doc_Date between @date1 and @date2 order by Name"
        If Type = "S" Then SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency, (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SINV') as Total, (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SINV') as Tax, (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SC') as TotalC, (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SC') as TaxC from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where " & where & "  piv.Doc_Date between @date1 and @date2 order by Name"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date1", FromDate)
        SQLCommand.Parameters.AddWithValue("@date2", ToDate)
        DataAdapter.Fill(Sales)

        Sales.Columns.Add("SubTotal", GetType(Decimal))
        Sales.Columns.Add("Percentage", GetType(String))

        Dim SubTotal As Decimal = 0
        Dim Tax As Decimal = 0
        Dim Total As Decimal = 0
        For i = 0 To Sales.Rows.Count - 1
            If Sales.Rows(i)("Currency").ToString = "" Then Sales.Rows(i)("Currency") = "CAD"
            Sales.Rows(i)("SubTotal") = (Val(Sales.Rows(i)("Total").ToString) - Val(Sales.Rows(i)("Tax").ToString)) - (Val(Sales.Rows(i)("TotalC").ToString) - Val(Sales.Rows(i)("TaxC").ToString))
            SubTotal = SubTotal + (Val(Sales.Rows(i)("Total").ToString) - Val(Sales.Rows(i)("Tax").ToString)) - (Val(Sales.Rows(i)("TotalC").ToString) - Val(Sales.Rows(i)("TaxC").ToString))
            Tax = Tax + Val(Sales.Rows(i)("Tax").ToString) - Val(Sales.Rows(i)("TaxC").ToString)
            Total = Total + Val(Sales.Rows(i)("Total").ToString) - Val(Sales.Rows(i)("TotalC").ToString)

        Next

        ' Percentage
        If Percentage = "on" Then
            For i = 0 To Sales.Rows.Count - 1
                Sales.Rows(i)("Percentage") = (Val(Sales.Rows(i)("SubTotal").ToString) / SubTotal) * 100
                Sales.Rows(i)("Percentage") = Format(Val(Sales.Rows(i)("Percentage").ToString), "##.00") + "%"
            Next
        End If

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        If Percentage = "on" Then
            If Request.Form("Ac") <> "on" Then
                For i = 0 To Sales.Rows.Count - 1
                    Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Percentage").ToString, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                Next
            Else
                For i = 0 To Sales.Rows.Count - 1
                    Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Cust_Vend_ID").ToString, "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Percentage").ToString, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                Next
            End If

        Else
            If Request.Form("Ac") <> "on" Then
                For i = 0 To Sales.Rows.Count - 1
                    Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                Next
            Else
                For i = 0 To Sales.Rows.Count - 1
                    Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Cust_Vend_ID").ToString, "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
                Next
            End If

        End If

        Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", TotalTitle, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold; border-top:solid 1px black; border-bottom: double 3px black", Currency & " " & Format(SubTotal, "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        'Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", TotalTitle, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(SubTotal, "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(Tax, "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(Total, "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then
            Sales.Columns.Remove("Total")
            Sales.Columns.Remove("Tax")
            Sales.Columns.Remove("TotalC")
            Sales.Columns.Remove("TaxC")

            ' Create new Datatable
            Dim exportTable As New DataTable

            exportTable.Columns.Add("value0", GetType(String))
            exportTable.Columns.Add("value1", GetType(String))
            exportTable.Columns.Add("value2", GetType(String))
            exportTable.Columns.Add("value3", GetType(String))

            For i = 0 To Sales.Rows.Count - 1
                exportTable.Rows.Add(Sales.Rows(i)("Cust_Vend_ID").ToString, Sales.Rows(i)("Name").ToString, Format(Sales.Rows(i)("SubTotal"), "$#,###.00"), Sales.Rows(i)("Percentage").ToString)
            Next

            ' Adding the Total
            exportTable.Rows.Add("", "Total", Currency & " " & Format(SubTotal, "$#,###.00"), "")

            ' Rename the existing column name to value
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = exportTable.NewRow()
            If Language = 0 Then
                excelheader("value0") = "ID"

                If Type = "P" Then
                    excelheader("value1") = "Vendor Name"
                    excelheader("value2") = "Purchases"
                Else
                    excelheader("value1") = "Customer Name"
                    excelheader("value2") = "Net Sales"
                End If
                If Percentage = "on" Then
                    excelheader("value3") = "Percentage"
                End If
            ElseIf Language = 1 Then
                excelheader("value0") = "ID"

                If Type = "P" Then
                    excelheader("value1") = "Nombre del Vendedor"
                    excelheader("value2") = "Compras"
                Else
                    excelheader("value1") = "Nombre del Cliente"
                    excelheader("value2") = "Las Ventas Netas"
                End If
                If Percentage = "on" Then
                    excelheader("value3") = "Porcentaje"
                End If
            End If

            ' Account No
            If Request.Form("Ac") <> "on" Then
                exportTable.Columns.Remove("value0")
                ' Rename the header so the cell shifted to value0
                For i = 0 To exportTable.Columns.Count - 1
                    exportTable.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            exportTable.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If

    End Sub

    Private Sub PrintPo()
        'Print PO with matching id
        Dim formID As String
        formID = "9269"



        Dim COA, Bal1, Bal2, Report, rpt As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
        HF_PrintTitle.Value = "<table style='Width: 100%;'><tr><td><img style='background-repeat: no-repeat;' src='images/Axiom_A.bmp'></td><td><table><tr><td style='font-size: 20px;'>Axiom Group Inc</td></tr><tr><td>Phone: 905-727-2878 Inc</td></tr><tr><td>Email: info@axiomgroup.ca</td></tr><tr><td>115 Mary Street, Aurora,</td></tr><tr><td>Ontario, Canada, L4G 1G3</td></tr></table></td><td align='right' style='text-align:right; background-color: red;'><table><tr><td style='font-size: 20px;' align='right'>PURCHASE ORDER</td></tr><tr><td>PURCHASE ORDER NO</td></tr><tr><td>" & COA.Rows(0)("fk_Doc_ID").ToString & "</td></tr></table></td></tr></table>"

        SQLCommand.CommandText = " SELECT Template, RepeaterFields, RepeaterTitles, RepeaterWidths FROM ACC_Cust_FormTemplates WHERE (fk_FormName_ID=@formid AND Form_Type = @invtype)"
        SQLCommand.Parameters.Clear()
        'SQLCommand.Parameters.AddWithValue("@formid", formTypeID)
        'SQLCommand.Parameters.AddWithValue("@invtype", formTC)
        DataAdapter.Fill(DBTable)

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True
    End Sub

    ' Income Statement
    Private Sub PickIncomeSheet()
        'Select what versiom of the income statement we should print
        If (Request.Form("FirstDate") = Request.Form("SecondDate")) Then
            PrintIncomeStatementSingle()
        Else
            PrintProfitLoss()
        End If
    End Sub
    ' Print the single income statement
    Private Sub PrintIncomeStatementSingle()

        Dim Language As Integer = Request.Form("language")
        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As String
        Dim seconDate As String
        Dim Acc_No As String = Request.Form("Ac")
        Dim Percentage As String = Request.Form("Perce")
        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"
        Dim ProfitAndLoss As String
        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Default date give today's date and a year before
        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Dim HF_Acc As String = ""
        Dim HF_Pre As String = ""
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        ' Translate the Header
        If Language = 0 Then
            If Acc_No = "on" Then
                HF_Acc = "width:60px; font-size:8pt~Act No"
            Else
                HF_Acc = "width:0px; font-size:0pt~"
            End If

            If Percentage = "on" Then
                HF_Pre = "width:160px; font-size:8pt~Sales/Expenses(%)"
            Else
                HF_Pre = "width:0px; font-size:0pt~"
            End If
            HF_PrintHeader.Value = "text-align:left; " & HF_Acc & "~text-align:left; width:350px; font-size:8pt~Account Description~text-align:right; width:120px; font-size:8pt~Dollar Amount~text-align:right; " & HF_Pre & "~text-align:centre; width:0px; font-size:0pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Income Statement " + DenomString + "<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            If Acc_No = "on" Then
                HF_Acc = "width:60px; font-size:8pt~Act No"
            Else
                HF_Acc = "width:0px; font-size:0pt~"
            End If

            If Percentage = "on" Then
                HF_Pre = "width:160px; font-size:8pt~Ventas/Gastos(%)"
            Else
                HF_Pre = "width:0px; font-size:0pt~(%)"
            End If
            HF_PrintHeader.Value = "text-align:left; " & HF_Acc & "~text-align:left; width:350px; font-size:8pt~Descripción De Cuenta~text-align:right; width:140px; font-size:8pt~Monto En Dólares~text-align:right; " & HF_Pre & "~text-align:centre; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado De Resultados " + DenomString + "<br/>Desde " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No;"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No;"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No;"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No;"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("Dollar_Difference", GetType(Decimal))
        COA.Columns.Add("DifferenceString", GetType(String))
        COA.Columns.Add("Percent_Difference", GetType(String))
        COA.Columns.Add("Percent_DifferenceString", GetType(String))

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    Dim denominatedValueNext As Double = Convert.ToDouble(Val(COA.Rows(i)("NextDateBalance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValueCurrent
                    COA.Rows(i)("NextDateBalance") = denominatedValueNext
                End If

            Next
        End If

        ' Give Padding
        For i = 0 To COA.Rows.Count - 1
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
        Next

        Dim Total As Decimal = 0
        Dim Account As String = ""
        ' Calculating Sub-Total and Total
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                    For iii = 0 To COA.Rows.Count - 1
                        If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                        If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                    Next
                Next
            End If
            For ii = 0 To COA.Rows.Count - 1
                If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("Balance") = Total
            Next


        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        If rowIncome.Length > 0 Then
            TotalIncome = rowIncome(0).Item("Balance")
        End If
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        If rowCost.Length > 0 Then
            TotalCost = rowCost(0).Item("Balance")
        End If
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")
        If rowExpense.Length > 0 Then
            TotalExpenses = rowExpense(0).Item("Balance")
        End If

        'Set the percentages
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then
                        Endd = Trim(Dash(0))
                    Else
                        Endd = Trim(Dash(1))
                    End If
                    For iii = 0 To COA.Rows.Count - 1

                        If COA.Rows(iii)("Account_Type") < 90 Then
                            Try
                                If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "4" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalIncome)) * 100
                                End If
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "5" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalCost)) * 100
                                End If
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "6" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalExpenses)) * 100
                                End If
                            Catch Ex As Exception
                            End Try
                        End If
                    Next
                Next
            End If

        Next

        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            COA.Rows(i)("Percent_Difference") = Format(Val(COA.Rows(i)("Percent_Difference").ToString), "##.00") + "%"
            If COA.Rows(i)("Percent_Difference").ToString = "0.##%" Or COA.Rows(i)("Percent_Difference").ToString = "%" Then COA.Rows(i)("Percent_Difference") = ""
            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"

            If Request.Form("Round") = "on" Then
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
            Else
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            End If

            If Left(COA.Rows(i)("DifferenceString").ToString, 1) = "-" Then COA.Rows(i)("DifferenceString") = "(" & COA.Rows(i)("DifferenceString").replace("-", "") & ")"

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("Percent_Difference").ToString = ".00%" Or COA.Rows(i)("Percent_Difference").ToString = "00%" Then COA.Rows(i)("Percent_Difference") = ""
            If COA.Rows(i)("DifferenceString").ToString = "$.00" Or COA.Rows(i)("DifferenceString").ToString = "$" Then COA.Rows(i)("DifferenceString") = ""
            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

        Next
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            ' Delete the rows that arnt above the detail level 
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If COA.Rows(i)("BalanceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                ElseIf COA.Rows(i)("DifferenceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                End If

            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next


        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 1.5in; max-width: 1.5in;"
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 3.5in; max-width: 3.5in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 1.1in; max-width: 1.1in;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:15px;"
                Style2 = Style2 & "; padding-bottom:15px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"

            End If
            Dim Ac_Style = " font-size:0pt;"
            Dim Per_Style = "font-size:0pt;"


            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If


            If Percentage = "on" Then
                If (Left(Per_Style.ToString, 1) = "-") Then
                    Per_Style = "(" & Per_Style.Replace("-", "") & ")"
                    Per_Style = Per_Style & "color: red !important;"
                End If
            End If

            If Percentage = "on" Then
                Per_Style = "text-align:right;font-size:8pt;width: 10px;"
            End If
            Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("DifferenceString") + "</span>", Per_Style, COA.Rows(i)("Percent_Difference"), "font-size:8pt; width:100px", COA.Rows(i)("fk_Currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        ProfitAndLoss = Convert.ToDecimal(TotalIncome) - Convert.ToDecimal(TotalCost) - Convert.ToDecimal(TotalExpenses)
        ProfitAndLoss = Format(Val(ProfitAndLoss.ToString), "$#,###.00")

        ' Check ProfitAndLoss Value negative or positive
        If Left(ProfitAndLoss.ToString, 1) = "-" Then
            ProfitAndLoss = "(" & ProfitAndLoss.Replace("-", "") & ")"
            StyleFinish = StyleFinish & "color: red !important;"
        End If

        Style = Style & "padding-bottom:0px;"
        Style2 = "text-align:right; font-size:8pt; min-width: 1.5in; max-width: 1.5in; padding: 0px 0px 0px 0px; font-weight:bold;border-top: px solid black;"

        Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + ProfitAndLoss + "</span>", "font-size:8pt; width:50px ;text-align:right ", "", "font-size:8pt; width:100px", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString")
            COA.Columns.Remove("Dollar_Difference")
            COA.Columns.Remove("Percent_DifferenceString")
            COA.Columns.Remove("Balance")

            ' Adding Total Profit and Loss
            COA.Rows.Add("", "PROFIT AND LOSS", "", ProfitAndLoss, "")

            ' Rename the existing column name to value
            For i = 0 To COA.Columns.Count - 1
                COA.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = COA.Columns.Count To 25
                COA.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = COA.NewRow()
            excelheader("value0") = "Account No"
            excelheader("value1") = "Name"
            excelheader("value2") = "Currency"
            excelheader("value3") = "Balance"
            excelheader("value4") = "Percentage"
            COA.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To COA.Columns.Count - 1
                COA.Rows(0)(i) = "<b>" & COA.Rows(0)(i).ToString & "</b>"
            Next

            ' Percentage
            If Percentage <> "on" Then
                COA.Columns.Remove("value4")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            ' Account No
            If Request.Form("Ac") <> "on" Then
                COA.Columns.Remove("value0")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            RPT_Excel.DataSource = COA
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub
    'Print the summary trail sheet
    Private Sub PrintSummaryTrail()

        Dim Language As Integer = Request.Form("language")
        Dim firstDate As String
        Dim seconDate As String
        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try
        'Set the header
        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; width:550px; font-size:8pt~Account Name~text-align:right; width:120px; font-size:8pt~Beginning Balance~text-align:right; width:120px; font-size:8pt~Debit~text-align:right; width:120px; font-size:8pt~Credit~text-align:right; width:120px; font-size:8pt~Net actvity~text-align:right; width:120px; font-size:8pt~Closing Balance"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Summary Trial Balance " + DenomString + "<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'></span><span style='position: absolute; margin-left: 1.7in;'></span><span style='position: absolute; margin-left: 3.3in'></span><span style='position: absolute; margin-left: 4.5in'></span><span style='position: absolute; margin-left: 5.5in'></span><span style='position: absolute; margin-left: 6.8in;'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; width:550px; font-size:8pt~Nombre De La Cuenta~text-align:right; width:120px; font-size:8pt~Balance Inicial~text-align:right; width:120px; font-size:8pt~Débito~text-align:right; width:120px; font-size:8pt~Crédito~text-align:right; width:120px; font-size:8pt~Actividad Neto~text-align:right; width:120px; font-size:8pt~Balance De Cierre"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Resumen Del Balance De Prueba " + DenomString + "<br/>Desde " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso en " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'></span><span style='position: absolute; margin-left: 1.7in;'></span><span style='position: absolute; margin-left: 3.3in'></span><span style='position: absolute; margin-left: 4.5in'></span><span style='position: absolute; margin-left: 5.5in'></span><span style='position: absolute; margin-left: 6.8in;'></span></div>"
        End If
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @startDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No"
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No"
        End If
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@endDate", DatStart)
        SQLCommand.Parameters.AddWithValue("@date", DatSecond)
        SQLCommand.Parameters.AddWithValue("@startDate", DatStart.AddDays(-1))
        DataAdapter.Fill(COA)

        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("NetString", GetType(String))
        COA.Columns.Add("NextDateBalanceString", GetType(String))

        Dim finalCredit As Double
        Dim finalDebit As Double
        Dim finalNet As Double
        'Get the total for the end of the page
        Dim COACount As Int32 = COA.Rows.Count - 1
        For i = 0 To COA.Rows.Count - 1
            finalCredit = finalCredit + COA.Rows(i)("Credit")
            finalDebit = finalDebit + COA.Rows(i)("Debit")
            finalNet = finalNet + COA.Rows(i)("NetActivity")
        Next
        'create the final row
        Try
            Dim newRow As DataRow = COA.NewRow()
            newRow.BeginEdit()
            ' newRow("Balance") = COA.Rows(COACount)("Balance")
            newRow("Credit") = finalCredit
            newRow("Debit") = finalDebit
            newRow("NetActivity") = finalNet
            newRow("Name") = "Total"

            newRow("Account_Type") = "33"
            newRow.EndEdit()
            COA.Rows.Add(newRow)
        Catch ex As Exception

        End Try

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    Dim denominatedValueNext As Double = Convert.ToDouble(Val(COA.Rows(i)("NextDateBalance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValueCurrent
                    COA.Rows(i)("NextDateBalance") = denominatedValueNext
                End If
            Next
        End If

        'Formatting the output
        For i = 0 To COA.Rows.Count - 1

            ' Format all the output for the paper
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit").ToString), "$#,###")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit").ToString), "$#,###")
                COA.Rows(i)("NetString") = Format(Val(COA.Rows(i)("NetActivity").ToString), "$#,###")
            Else
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###.00")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit").ToString), "$#,###.00")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit").ToString), "$#,###.00")
                COA.Rows(i)("NetString") = Format(Val(COA.Rows(i)("NetActivity").ToString), "$#,###.00")
            End If

            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("NextDateBalanceString").ToString, 1) = "-" Then COA.Rows(i)("NextDateBalanceString") = "(" & COA.Rows(i)("NextDateBalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
            If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"

            If Left(COA.Rows(i)("NetString").ToString, 1) = "-" Then COA.Rows(i)("NetString") = "(" & COA.Rows(i)("NetString").replace("-", "") & ")"
            'If Val(COA.Rows(i)("Level").ToString) > 1 Then COA.Rows(i).Delete()
            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("NextDateBalanceString").ToString = "$.00" Or COA.Rows(i)("NextDateBalanceString").ToString = "$" Then COA.Rows(i)("NextDateBalanceString") = ""
            If COA.Rows(i)("CreditString").ToString = "$.00" Or COA.Rows(i)("CreditString").ToString = "$" Then COA.Rows(i)("CreditString") = ""
            If COA.Rows(i)("DebitString").ToString = "$.00" Or COA.Rows(i)("DebitString").ToString = "$" Then COA.Rows(i)("DebitString") = ""
            If COA.Rows(i)("NetString").ToString = "$.00" Or COA.Rows(i)("NetString").ToString = "$" Then COA.Rows(i)("NetString") = ""

        Next

        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            ' Delete the rows that arnt above the detail level 
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If COA.Rows(i)("CreditString").ToString = "" And COA.Rows(i)("DebitString").ToString = "" Then
                    COA.Rows(i).Delete()
                End If
            End If
        Next i
        COA.AcceptChanges()
        'Pringint to the page
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 0px 0px 0px; min-width: 2.5in; max-width: 2.5in;"
        Dim StyleID As String = "font-size:0pt; width: 0px;"
        If Request.Form("Ac") = "on" Then
            StyleID = "font-size:8pt; width: 25px;"
        End If
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; width: 5in;"
            If COA.Rows(i)("Account_Type") > 90 Then Style = Style & "; font-weight:bold"
            Report.Rows.Add("padding: 0px 0px 0px 0px;border-top:solid 0px; text-align:left; " & StyleID, COA.Rows(i)("Account_No").ToString, Style + "width: 1px;border-top:solid 0px; min-width: 1in; max-width: 1in;", COA.Rows(i)("Name").ToString, "padding: 3px 5px 3px 5px;border-top:solid 0px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;", COA.Rows(i)("NextDateBalanceString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;padding-left: 0.2in;", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; border-top:solid 0px;min-width: 1in;max-width: 1in;", COA.Rows(i)("NetString"), "padding: 3px 5px 3px 5px; text-align:right; border-top:solid 0px;font-size:8pt; padding-left: 0.2in;", COA.Rows(i)("BalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        Report.Rows.Add("padding: 0px 0px 0px 0px;font-size:1pt;", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "border-top:double 0px", "", "border-top:double 0px", "", "border-top:double 4px", "", "border-top:double 4px", "", "border-top:double 0px", "")
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Balance")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Transaction_No")
            COA.Columns.Remove("Memo")
            COA.Columns.Remove("memo2")
            COA.Columns.Remove("Credit")
            COA.Columns.Remove("Debit")
            COA.Columns.Remove("NetActivity")
            COA.Columns.Remove("NextDateBalance")

            ' Rename the existing column name to value
            For i = 0 To COA.Columns.Count - 1
                COA.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = COA.Columns.Count To 25
                COA.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = COA.NewRow()
            excelheader("value0") = "Account No"
            excelheader("value1") = "Account Name"
            excelheader("value2") = "Beginning Balance"
            excelheader("value3") = "Debit"
            excelheader("value4") = "Credit"
            excelheader("value5") = "Net Activity"
            excelheader("value6") = "Closing Balance"
            COA.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To COA.Columns.Count - 1
                COA.Rows(0)(i) = "<b>" & COA.Rows(0)(i).ToString & "</b>"
            Next

            ' Account No
            If Request.Form("Ac") <> "on" Then
                COA.Columns.Remove("value0")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            RPT_Excel.DataSource = COA
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub
    Private Sub PrintDetailTrial()
        'Print the Detail Trial Sheet
        Dim StartDate As String
        Dim EndDate As String
        Dim accNo As String
        Dim Denom As Int32 = Request.Form("Denom")
        Dim id As String = Request.Form("id")
        Dim DenomString As String = ""

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If (Denom > 1) Then
            If Denom = 10 Then
                DenomString = "(In Tenth)"
            ElseIf Denom = 100 Then
                DenomString = "(In Hundreds)"
            ElseIf Denom = 1000 Then
                DenomString = "(In Thousands)"
            End If
        End If

        StartDate = Request.Form("StartDate")
        EndDate = Request.Form("EndDate")
        accNo = Request.Form("accNo")

        If StartDate = "" Then StartDate = Now().ToString("yyyy-MM-dd")
        If EndDate = "" Then EndDate = Now().AddDays(-30).ToString("yyyy-MM-dd")

        HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"

        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()
        'Get the matching name with the id
        SQLCommand.CommandText = "SELECT Transaction_Date, fk_currency_id,Transaction_No, Document_Type, Debit_Amount, Credit_Amount, Balance, Memo, Document_ID, fk_Account_ID,rowID FROM ACC_GL WHERE ((Transaction_Date >= @startDate AND Transaction_Date <= @endDate) AND fk_Account_ID = @id) ORDER BY Transaction_Date asc, rowID asc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@startDate", StartDate)
        SQLCommand.Parameters.AddWithValue("@endDate", EndDate)
        SQLCommand.Parameters.AddWithValue("@id", id)
        DataAdapter.Fill(COA)
        'If we have a matching name output it to the header
        Try
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detail Trial Balance " + DenomString + "<br/>For the Period " & StartDate & " to " & EndDate & " - " + COA.Rows(0)("fk_Currency_ID").ToString + "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accName") + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'>Posting Date</span><span style='position: absolute; margin-left: 1in'>Doc No</span><span style='position: absolute; margin-left: 2.5in'>Description</span><span style='position: absolute; margin-left: 4.7in;'>Debit</span><span style='position: absolute; margin-left: 5.8in'>Credit</span><span style='position: absolute; margin-left: 6.7in'>Balance</span></div></div>"
        Catch ex As Exception
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detail Trial Balance " + DenomString + "<br/>For the Period " & StartDate & " to " & EndDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accNo") + " " + id + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'>Posting Date</span><span style='position: absolute; margin-left: 1in'>Doc No</span><span style='position: absolute; margin-left: 2.5in'>Description</span><span style='position: absolute; margin-left: 4.7in;'>Debit</span><span style='position: absolute; margin-left: 5.8in'>Credit</span><span style='position: absolute; margin-left: 6.7in'>Balance</span></div></div>"
        End Try

        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))

        'Prepare for the final row that shows all the chages
        Dim finalCredit As Double
        Dim finalDebit As Double
        Dim COACount As Int32 = COA.Rows.Count - 1
        For i = 0 To COA.Rows.Count - 1
            finalCredit = finalCredit + COA.Rows(i)("Credit_Amount")
            finalDebit = finalDebit + COA.Rows(i)("Debit_Amount")
        Next
        Try
            Dim newRow As DataRow = COA.NewRow()
            Dim transactionDate As Date
            transactionDate = "0001-01-01"

            newRow.BeginEdit()
            newRow("Balance") = COA.Rows(COACount)("Balance")
            newRow("Credit_Amount") = finalCredit
            newRow("Debit_Amount") = finalDebit
            newRow("memo") = Request.Form("accName")
            newRow("Transaction_Date") = transactionDate
            newRow.EndEdit()
            COA.Rows.Add(newRow)
        Catch ex As Exception

        End Try

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString))
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValue
                    Dim denominatedValue2 As Double = Convert.ToDouble(Val(COA.Rows(i)("Credit_Amount").ToString)) / Denom
                    COA.Rows(i)("Credit_Amount") = denominatedValue2
                    Dim denominatedValue3 As Double = Convert.ToDouble(Val(COA.Rows(i)("Debit_Amount").ToString)) / Denom
                    COA.Rows(i)("Debit_Amount") = denominatedValue3
                End If
            Next
        End If

        'formatting the user output
        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###")
            Else
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###.00")
            End If

            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("CreditString").ToString = "$.00" Or COA.Rows(i)("CreditString").ToString = "$" Then COA.Rows(i)("CreditString") = ""
            If COA.Rows(i)("DebitString").ToString = "$.00" Or COA.Rows(i)("DebitString").ToString = "$" Then COA.Rows(i)("DebitString") = ""

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
            If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"
        Next

        COA.AcceptChanges()

        'Preparing it for the page
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        For i = 0 To COA.Rows.Count - 1
            Dim Transaction_Date As Date = COA.Rows(i)("Transaction_Date").ToString()

            Report.Rows.Add("padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 0.7in;", Transaction_Date.ToString("yyyy-MM-dd"), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 0.5in;", COA.Rows(i)("Transaction_No").ToString, "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 1.5in;", COA.Rows(i)("memo"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;", COA.Rows(i)("BalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True
    End Sub
    Private Sub PrintDetailTrialChart()
        'Print the detail trial from the chart page
        Dim Language As Integer = Request.Form("language")
        Dim StartDate As String
        Dim EndDate As String
        Dim accNo As String
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        Dim BalDate As Date
        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        StartDate = Request.Form("StartDate")
        EndDate = Request.Form("EndDate")
        accNo = Request.Form("accNo")

        If StartDate = "" Then StartDate = Now().ToString("yyyy-MM-dd")
        If EndDate = "" Then EndDate = Now().AddDays(-30).ToString("yyyy-MM-dd")

        'Get account name
        Dim value As Object
        If Language = 0 Then
            Conn.Open()
            Dim querystr As String = "SELECT Name FROM ACC_GL_Accounts WHERE Account_No = " + accNo + ";"
            Dim mycmd As New SqlCommand(querystr, Conn)
            value = mycmd.ExecuteScalar()
            Conn.Close()
        ElseIf Language = 1 Then
            Conn.Open()
            Dim querystr As String = "SELECT TranslatedName as Name FROM ACC_GL_Accounts WHERE Account_No = " + accNo + ";"
            Dim mycmd As New SqlCommand(querystr, Conn)
            value = mycmd.ExecuteScalar()
            Conn.Close()
        End If

        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~ Posting Date~text-align:left; width:350px; font-size:8pt~Doc No~text-align:left; width:120px; font-size:8pt~Description~text-align:right; width:120px; font-size:8pt~Debit~text-align:right; width:120px; font-size:8pt~Credit~text-align:right; width:120px; font-size:8pt~Balance"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~ Fecha~text-align:left; width:350px; font-size:8pt~Núm Del Doc~text-align:left; width:120px; font-size:8pt~Descripción~text-align:right; width:120px; font-size:8pt~Débito~text-align:right; width:120px; font-size:8pt~Crédito~text-align:right; width:120px; font-size:8pt~El Balance"
        End If

        Dim COA, Bal, Report As New DataTable
        PNL_Summary.Visible = True

        Dim Fiscal As New DataTable

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        SQLCommand.CommandText = "SELECT rowID, CAST (Transaction_Date AS varchar(50)) as Transaction_Date, Transaction_No, Document_Type, Debit_Amount, Credit_Amount, Balance, Memo, Document_ID, fk_Account_ID,Account_No,ACC_GL.fk_Currency_ID FROM ACC_GL LEFT JOIN ACC_GL_Accounts on ACC_GL_Accounts.Account_ID = ACC_GL.fk_Account_ID WHERE ((Transaction_Date >= @startDate AND Transaction_Date <= @endDate) AND Account_No = @accNo) ORDER BY Transaction_Date asc, rowID asc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@startDate", StartDate)
        SQLCommand.Parameters.AddWithValue("@endDate", EndDate)
        SQLCommand.Parameters.AddWithValue("@accNo", accNo)
        DataAdapter.Fill(COA)

        If Language = 0 Then
            Try
                'Set the page header, this is below the SQL so we can get the currency
                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detail Trial Balance " + DenomString + "<br/>For the Period " & StartDate & " to " & EndDate & " - " + COA.Rows(0)("fk_Currency_ID").ToString + "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
            Catch ex As Exception
                'Set the page header, this is below the SQL so we can get the currency
                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detail Trial Balance " + DenomString + "<br/>For the Period " & StartDate & " to " & EndDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
            End Try
        ElseIf Language = 1 Then
            Try
                'Set the page header, this is below the SQL so we can get the currency
                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detalle Balance De Prueba " + DenomString + "<br/>Para el Periodo " & StartDate & " a " & EndDate & " - " + COA.Rows(0)("fk_Currency_ID").ToString + "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
            Catch ex As Exception
                'Set the page header, this is below the SQL so we can get the currency
                HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Detalle Balance De Prueba " + DenomString + "<br/>Para el Periodo " & StartDate & " a " & EndDate & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
            End Try
        End If

        COA.Columns.Add("DebitString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("BalanceString", GetType(String))

        Dim finalCredit As Double
        Dim finalDebit As Double
        Dim COACount As Int32 = COA.Rows.Count - 1
        For i = 0 To COA.Rows.Count - 1
            finalCredit = finalCredit + COA.Rows(i)("Credit_Amount")
            finalDebit = finalDebit + COA.Rows(i)("Debit_Amount")
        Next
        Try
            Dim newRow As DataRow = COA.NewRow()
            Dim transactionDate As Date
            transactionDate = "0001-01-01"

            newRow.BeginEdit()
            newRow("Balance") = COA.Rows(COACount)("Balance")
            newRow("Credit_Amount") = finalCredit
            newRow("Debit_Amount") = finalDebit
            newRow("memo") = Request.Form("accNo") + " " + value.ToString()
            newRow("Transaction_Date") = transactionDate
            newRow.EndEdit()
            COA.Rows.Add(newRow)
        Catch ex As Exception

        End Try

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString))
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValue
                    Dim denominatedValue2 As Double = Convert.ToDouble(Val(COA.Rows(i)("Credit_Amount").ToString)) / Denom
                    COA.Rows(i)("Credit_Amount") = denominatedValue2
                    Dim denominatedValue3 As Double = Convert.ToDouble(Val(COA.Rows(i)("Debit_Amount").ToString)) / Denom
                    COA.Rows(i)("Debit_Amount") = denominatedValue3
                End If
            Next
        End If

        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###")
            Else
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###.00")
            End If

            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("CreditString").ToString = "$.00" Or COA.Rows(i)("CreditString").ToString = "$" Then COA.Rows(i)("CreditString") = ""
            If COA.Rows(i)("DebitString").ToString = "$.00" Or COA.Rows(i)("DebitString").ToString = "$" Then COA.Rows(i)("DebitString") = ""

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
            If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"
        Next

        COA.AcceptChanges()

        ' Adding Beginning balance
        BalDate = StartDate
        SQLCommand.CommandText = "SELECT TOP 1 (rowID), CAST (Transaction_Date AS varchar(50)) as Transaction_Date, Transaction_No, Document_Type, Debit_Amount, Credit_Amount, Balance, Memo, Document_ID, fk_Account_ID,Account_No,ACC_GL.fk_Currency_ID FROM ACC_GL LEFT JOIN ACC_GL_Accounts on ACC_GL_Accounts.Account_ID = ACC_GL.fk_Account_ID WHERE (Transaction_Date >= @startDate AND Account_No = @accNo) ORDER BY Transaction_Date asc, rowID desc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@startDate", BalDate.AddDays(-1).ToString("yyyy-MM-dd"))
        SQLCommand.Parameters.AddWithValue("@accNo", accNo)
        DataAdapter.Fill(Bal)

        Bal.Columns.Add("BalanceString", GetType(String))

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        ' Formatting for Bal Datatable
        If Request.Form("Round") = "on" Then
            Bal.Rows(0)("BalanceString") = Format(Val(Bal.Rows(0)("Balance").ToString), "$#,###")
        Else
            Bal.Rows(0)("BalanceString") = Format(Val(Bal.Rows(0)("Balance").ToString), "$#,###.00")
        End If
        If Left(Bal.Rows(0)("BalanceString").ToString, 1) = "-" Then Bal.Rows(0)("BalanceString") = "(" & Bal.Rows(0)("BalanceString").replace("-", "") & ")"

        ' Create datarow to insert value to COA datatable
        Dim coaHeader As DataRow = COA.NewRow()
        coaHeader.BeginEdit()
        ' newRow("Balance") = COA.Rows(COACount)("Balance")
        coaHeader("Transaction_Date") = BalDate.ToString("yyyy-MM-dd")
        coaHeader("Memo") = "Beginning Balance"
        coaHeader("Balance") = Bal.Rows(0)("Balance").ToString
        coaHeader("BalanceString") = Bal.Rows(0)("BalanceString").ToString

        coaHeader.EndEdit()
        COA.Rows.InsertAt(coaHeader, 0)

        For i = 0 To COA.Rows.Count - 1
            Dim Transaction_Date As Date = COA.Rows(i)("Transaction_Date").ToString()

            Report.Rows.Add("padding: 3px 5px 3px 5px; text-align:left;border-top: solid black 0px; font-size:8pt; min-width: 0.7in;", Transaction_Date.ToString("yyyy-MM-dd"), "text-align:left;border-top: solid black 0px; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 0.7in;", COA.Rows(i)("Transaction_No").ToString, "padding: 3px 5px 3px 5px; border-top: solid black 0px;text-align:left; font-size:8pt; min-width: 2.7in;", COA.Rows(i)("memo"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 0.7in;", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right;font-size:8pt;min-width: 0.7in;", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 0.7in;", COA.Rows(i)("BalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        Next
        Report.Rows.Add("padding: 0px 0px 0px 0px;font-size:1pt;", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "border-top:double 0px", "", "border-top:double 0px", "", "border-top:double 4px", "", "border-top:double 4px", "", "border-top:double 4px", "")
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            COA.Columns.Remove("rowID")
            COA.Columns.Remove("Document_Type")
            COA.Columns.Remove("Debit_Amount")
            COA.Columns.Remove("Credit_Amount")
            COA.Columns.Remove("Balance")
            COA.Columns.Remove("Document_ID")
            COA.Columns.Remove("fk_Account_ID")
            COA.Columns.Remove("Account_No")
            COA.Columns.Remove("fk_Currency_ID")

            ' Rename the existing column name to value
            For i = 0 To COA.Columns.Count - 1
                COA.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = COA.Columns.Count To 25
                COA.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = COA.NewRow()
            excelheader("value0") = "Posting Date"
            excelheader("value1") = "Doc No"
            excelheader("value2") = "Description"
            excelheader("value3") = "Debit"
            excelheader("value4") = "Credit"
            excelheader("value5") = "Balance"
            COA.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To COA.Columns.Count - 1
                COA.Rows(0)(i) = "<b>" & COA.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = COA
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    Private Sub PrintBankRec()

        Dim TempDate As Date
        Dim ID As Integer = Request.Form("ID")
        Dim List, RecBalances, Rec, Report As New DataTable

        Dim Account As String = ""
        Dim RecDate As String = ""
        Dim StatementBalance As String = "$0.00"
        Dim UserName As String = ""
        Dim LastDate As String = ""

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Conn.Open()
        SQLCommand.CommandText = "Update ACC_GL set Reconciled='0' where Reconciled is null"
        SQLCommand.Parameters.Clear()
        SQLCommand.ExecuteNonQuery()
        Conn.Close()

        SQLCommand.CommandText = "Select *, (Select FirstName + ' ' + LastName as UserName From Web_Users where UserID = Reconciled_By) UserName, (Select Top 1 Rec_Date From ACC_Bank_Rec where fk_Account_ID = br.fk_Account_ID and Rec_Date <br.Rec_Date order by Rec_Date desc) LastRecDate  From ACC_Bank_Rec br join ACC_GL_Accounts on br.fk_Account_ID = Account_ID where br.Rec_ID = @id"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@id", ID)
        DataAdapter.Fill(Rec)

        Dim AccountName As String = Rec.Rows(0)("Account_No").ToString & " " & Rec.Rows(0)("Name").ToString

        If Rec.Rows.Count > 0 Then
            StatementBalance = Format(Val(Rec.Rows(0)("Statement_Balance").ToString), "$#,##0.00")
            Account = Rec.Rows(0)("fk_Account_ID").ToString
            TempDate = Rec.Rows(0)("Rec_Date").ToString
            UserName = Rec.Rows(0)("UserName").ToString
            RecDate = TempDate.ToString("yyyy-MM-dd")
            LastDate = Convert.ToDateTime(Rec.Rows(0)("LastRecDate").ToString).ToString("yyyy-MM-dd")
        End If

        SQLCommand.CommandText = "Select top 1 Balance From ACC_GL where Transaction_Date <=@date and fk_Account_ID = @account order by Transaction_Date desc, rowID desc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", RecDate)
        SQLCommand.Parameters.AddWithValue("@account", Account)
        SQLCommand.Parameters.AddWithValue("@id", ID)
        DataAdapter.Fill(RecBalances)

        SQLCommand.CommandText = "Select Sum(Debit_Amount) debit, Sum(Credit_Amount) credit From ACC_GL where (isnull(Reconciled, '0') = '0' or isnull(Reconciled, '0') ='0' or Reconciled in (Select Rec_ID from ACC_Bank_Rec where fk_Account_ID = @account and Rec_Date>@date)) and Transaction_Date <=@date and fk_Account_ID = @account"
        DataAdapter.Fill(RecBalances)

        Dim GLBal As String = Format(Val(RecBalances.Rows(0)("Balance").ToString), "$#,##0.00")
        Dim OutDebits As String = Format(Val(RecBalances.Rows(1)("Debit").ToString), "$#,##0.00")
        Dim OutCredits As String = Format(Val(RecBalances.Rows(1)("Credit").ToString), "$#,##0.00")
        Dim AdJBal As String = Format(Val(RecBalances.Rows(0)("Balance").ToString) - Val(RecBalances.Rows(1)("Debit").ToString) + Val(RecBalances.Rows(1)("Credit").ToString), "$#,###.00")
        Dim OOB As String = Format(Val(Rec.Rows(0)("Statement_Balance").ToString) - Val(AdJBal.Replace(",", "").Replace("$", "")), "$#,###.00")

        'This is where we get the lines for the report. 

        Dim TDate As String = ""
        Dim Debit As String = ""
        Dim Credit As String = ""
        Dim TDebit As Decimal = 0
        Dim TCredit As Decimal = 0

        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@id", ID)
        SQLCommand.Parameters.AddWithValue("@account", Account)
        SQLCommand.Parameters.AddWithValue("@date", RecDate)

        SQLCommand.CommandText = "SELECT Transaction_Date, Debit_Amount, Credit_Amount, Memo FROM ACC_GL WHERE fk_Account_ID = @account and Reconciled= @id and Debit_Amount<>0 AND Transaction_Date <= @date  ORDER BY Transaction_Date, Transaction_No, rowID"
        DataAdapter.Fill(List)

        If List.Rows.Count > 0 Then
            TDebit = 0 : TCredit = 0
            Report.Rows.Add("HEADERfont-size:8pt; font-weight:bold", "Cleared Deposits", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            For i = 0 To List.Rows.Count - 1
                TDate = Convert.ToDateTime(List.Rows(i)("Transaction_Date").ToString).ToString("yyyy-MM-dd")
                If List.Rows(i)("Debit_Amount").ToString = "0.00" Then Debit = "" Else Debit = Format(Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString), "$#,###.00")
                If List.Rows(i)("Credit_Amount").ToString = "0.00" Then Credit = "" Else Credit = Format(Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                TDebit = TDebit + Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString)
                TCredit = TCredit + Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString)
                Report.Rows.Add("text-align:left", TDate, "text-align:left", List.Rows(i)("Memo").ToString, "text-align:right", Debit, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            Next
            Report.Rows.Add("text-align:left", "", "text-align:left; font-weight:bold", "Total Cleared Deposits", "text-align:right; font-weight:bold", "<span style = ""border-top:solid 1px black; border-bottom:double 3px black"">" + Format(TDebit, "$#,###.00") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ' Report.Rows.Add("LINE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End If

        SQLCommand.CommandText = "SELECT Transaction_Date, Debit_Amount, Credit_Amount, Memo FROM ACC_GL WHERE fk_Account_ID = @account and Reconciled= @id and Credit_Amount<>0 AND Transaction_Date <= @date  ORDER BY Transaction_Date, Transaction_No, rowID"
        List.Reset()
        DataAdapter.Fill(List)

        If List.Rows.Count > 0 Then
            TDebit = 0 : TCredit = 0
            Report.Rows.Add("HEADERfont-size:8pt; font-weight:bold", "Cleared Cheques", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            For i = 0 To List.Rows.Count - 1
                TDate = Convert.ToDateTime(List.Rows(i)("Transaction_Date").ToString).ToString("yyyy-MM-dd")
                If List.Rows(i)("Debit_Amount").ToString = "0.00" Then Debit = "" Else Debit = Format(Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString), "$#,###.00")
                If List.Rows(i)("Credit_Amount").ToString = "0.00" Then Credit = "" Else Credit = Format(Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                TDebit = TDebit + Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString)
                TCredit = TCredit + Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString)
                Report.Rows.Add("text-align:left", TDate, "text-align:left", List.Rows(i)("Memo").ToString, "text-align:right", Credit, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            Next
            Report.Rows.Add("text-align:left", "", "text-align:left; font-weight:bold", "Total Cleared Cheques", "text-align:right; font-weight:bold; ", "<span style = ""border-top:solid 1px black; border-bottom:double 3px black"">" + Format(TCredit, "$#,###.00") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            'Report.Rows.Add("LINE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End If

        SQLCommand.CommandText = "SELECT Transaction_Date, Debit_Amount, Credit_Amount, Memo FROM ACC_GL WHERE fk_Account_ID = @account and Debit_Amount<>0 AND (isnull(Reconciled,0)=0 or Reconciled in (Select Rec_ID from ACC_Bank_Rec where fk_Account_ID = @account and Rec_Date>@date))  AND Transaction_Date <= @date  ORDER BY Transaction_Date, Transaction_No, rowID"
        List.Reset()
        DataAdapter.Fill(List)

        If List.Rows.Count > 0 Then
            TDebit = 0 : TCredit = 0
            Report.Rows.Add("HEADERfont-size:8pt; font-weight:bold", "Outstanding Deposits", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            For i = 0 To List.Rows.Count - 1
                TDate = Convert.ToDateTime(List.Rows(i)("Transaction_Date").ToString).ToString("yyyy-MM-dd")
                If List.Rows(i)("Debit_Amount").ToString = "0.00" Then Debit = "" Else Debit = Format(Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString), "$#,###.00")
                If List.Rows(i)("Credit_Amount").ToString = "0.00" Then Credit = "" Else Credit = Format(Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                TDebit = TDebit + Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString)
                TCredit = TCredit + Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString)
                Report.Rows.Add("text-align:left", TDate, "text-align:left", List.Rows(i)("Memo").ToString, "text-align:right", Debit, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            Next
            Report.Rows.Add("text-align:left", "", "text-align:left; font-weight:bold", "Total Outstanding Deposits", "text-align:right; font-weight:bold; ", "<span style = ""border-top:solid 1px black;border-bottom:double 3px black"">" + Format(TDebit, "$#,###.00") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            'Report.Rows.Add("LINE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End If

        SQLCommand.CommandText = "SELECT Transaction_Date, Debit_Amount, Credit_Amount, Memo FROM ACC_GL WHERE fk_Account_ID = @account and Credit_Amount<>0 AND (isnull(Reconciled,0)=0 or Reconciled in (Select Rec_ID from ACC_Bank_Rec where fk_Account_ID = @account and Rec_Date>@date))  AND Transaction_Date <= @date  ORDER BY Transaction_Date, Transaction_No, rowID"
        List.Reset()
        DataAdapter.Fill(List)

        If List.Rows.Count > 0 Then
            TDebit = 0 : TCredit = 0
            Report.Rows.Add("HEADERfont-size:8pt; font-weight:bold", "Outstanding Cheques", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            For i = 0 To List.Rows.Count - 1
                TDate = Convert.ToDateTime(List.Rows(i)("Transaction_Date").ToString).ToString("yyyy-MM-dd")
                If List.Rows(i)("Debit_Amount").ToString = "0.00" Then Debit = "" Else Debit = Format(Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString), "$#,###.00")
                If List.Rows(i)("Credit_Amount").ToString = "0.00" Then Credit = "" Else Credit = Format(Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                TDebit = TDebit + Convert.ToDecimal(List.Rows(i)("Debit_Amount").ToString)
                TCredit = TCredit + Convert.ToDecimal(List.Rows(i)("Credit_Amount").ToString)
                Report.Rows.Add("text-align:left", TDate, "text-align:left", List.Rows(i)("Memo").ToString, "text-align:right", Credit, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            Next
            Report.Rows.Add("text-align:left", "", "text-align:left; font-weight:bold", "Total Outstanding Cheques", "text-align:right; font-weight:bold;; ", "<span style = ""border-top:solid 1px black;border-bottom:double 3px black"">" + Format(TCredit, "$#,###.00") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            'Report.Rows.Add("LINE", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End If



        HF_PrintHeaderOnce.Value = "<table style=""padding:15px 0px 15px 0px"">"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2""><span>Account</span></td><td class=""tablecellprint2"" style=""text-align:left""><span>" & AccountName & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2"" style=""padding-bottom:10px""><span>Last Reconciled On</span></td><td class=""tablecellprint2"" style=""text-align:right; padding-bottom:10px""><span>" & LastDate & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2"" style=""font-weight:bold""><span>Bank Statement Balance</span></td><td class=""tablecellprint2"" style=""font-weight:bold; text-align:right""><span>" & StatementBalance & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2""><span>General Ledger Balance</span></td><td class=""tablecellprint2"" style=""text-align:right""><span>" & GLBal & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2""><span>Outstanding Deposits</span></td><td class=""tablecellprint2"" style=""text-align:right""><span>" & OutDebits & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2""><span>Outstanding Cheques</span></td><td class=""tablecellprint2"" style=""text-align:right""><span>" & OutCredits & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2"" style=""font-weight:bold""><span>Calculated General Ledger Balance</span></td><td class=""tablecellprint2"" style=""font-weight:bold;  text-align:right""><span>" & AdJBal & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "<tr><td class=""tablecellprint2""><span>Out of Balance</span></td><td class=""tablecellprint2"" style=""text-align:right""><span>" & OOB & "</span></td></tr>"
        HF_PrintHeaderOnce.Value = HF_PrintHeaderOnce.Value + "</table>"


        HF_PrintHeader.Value = "text-align:left; width:15%;~Date~text-align:left;width:60%;~Description~text-align:right;width:15%~Amount"

        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Bank Reconciliation<br/>Reconciled On " + RecDate + " by " + UserName + "<br/></span><span style=""font-size:7pt""><br>Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " </span>"

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

    End Sub
    ' Balance Sheet
    Private Sub PrintBalance()

        Dim Language As Integer = Request.Form("language")
        Dim AsAt As String = Request.Form("date1")
        Dim StyleFinish As String
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim NoZeros As String = Request.Form("showZeros")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DenomString As String = ""
        Dim HF_Acc As String = ""

        Dim RE As Decimal = 0

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        If AsAt = "" Then AsAt = Now().ToString("yyyy-MM-dd")

        If DetailLevel = 0 Then DetailLevel = 7
        If Acc_No = "on" Then
            HF_Acc = "width:50px; font-size:8pt~Act No"
        Else
            HF_Acc = "width:0px; font-size:0pt~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; " + HF_Acc & "~text-align:left; font-size:8pt~Account Name~text-align:right; width:100px; font-size:8pt~Balance~text-align:left; width:0px; font-size:0pt~"

            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Balance Sheet " + DenomString + "<br/>As Of " & AsAt & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"

        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; " + HF_Acc & "~text-align:left; font-size:8pt~Nombre De La Cuenta~text-align:right; width:100px; font-size:8pt~El Balance~text-align:left; width:0px; font-size:0pt~"

            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Hoja De Balance " + DenomString + "<br/>A Partir De " & AsAt & "<br/></span><span style=""font-size:7pt"">Impreso el " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
        End If

        Dim COA, Bal, DataT, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling Active, Cash, Exchange_Account_ID, Associated_Totalling From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(DataT)

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Totalling_Minus From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 10000 and Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, (Select Top 1 Balance from ACC_GL where Transaction_Date <= @date and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", AsAt)
            DataAdapter.Fill(Bal)
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Totalling_Minus From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 10000 and Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, (Select Top 1 Balance from ACC_GL where Transaction_Date <= @date and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", AsAt)
            DataAdapter.Fill(Bal)
        End If

        COA.Columns.Add("Balance", GetType(String))
        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))

        Dim Padding As Integer = 0
        Dim Level As Integer = 1

        ' Calculation for 39000 Current Retained Earning
        RE = 0

        For j = 0 To DataT.Rows.Count - 1

            For jj = 0 To Bal.Rows.Count - 1

                If DataT.Rows(j)("Account_ID").ToString = Bal.Rows(jj)("Account_ID").ToString Then

                    If DataT.Rows(j)("Account_Type").ToString = "4" Then

                        If Bal.Rows(jj)("Balance").ToString = "" Then

                        Else
                            RE = RE + Bal.Rows(jj)("Balance")
                        End If
                    End If
                    If DataT.Rows(j)("Account_Type").ToString = "5" Or DataT.Rows(j)("Account_Type").ToString = "6" Then

                        If Bal.Rows(jj)("Balance").ToString = "" Then

                        Else
                            RE = RE - Bal.Rows(jj)("Balance")
                        End If
                        Exit For
                    End If
                End If
            Next
        Next

        For j = 0 To Bal.Rows.Count - 1
            If Bal.Rows(j)("Account_No") = "39000" Then Bal.Rows(j)("Balance") = RE
            COA.AcceptChanges()

        Next

        For i = 0 To COA.Rows.Count - 1
            For ii = 0 To Bal.Rows.Count - 1
                ' Copying the Balance value from table Bal to table COA
                If COA.Rows(i)("Account_ID").ToString = Bal.Rows(ii)("Account_ID").ToString Then
                    COA.Rows(i)("Balance") = Bal.Rows(ii)("Balance")
                    Exit For
                End If
            Next
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 10 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 10 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "<div style='min-width: 0.5in; max-width:0.5in;'></div>" ' hard coded
        Next

        Dim Total As Decimal = 0
        Dim Account As String = ""
        ' Totalling Total Equity (ACC_NO 39999)
        For j = 1 To 2
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If

                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("Balance") = Total
                Next
            Next
        Next

        Total = 0
        Account = ""
        For j = 1 To 2
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                        Next
                    Next
                End If
                If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total - Val(COA.Rows(iii)("BeforeBalance").ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If

            Next
        Next


        For i = 0 To COA.Rows.Count - 1
            If Left(COA.Rows(i)("Account_No").ToString, 1) > "3" Then COA.Rows(i).Delete()
        Next

        COA.AcceptChanges()

        ' Formating
        ' Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValue
                End If
            Next
        End If

        For i = 0 To COA.Rows.Count - 1
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("Balance") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
            Else
                COA.Rows(i)("Balance") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            End If


            If COA.Rows(i)("Balance").ToString = "$.00" Or COA.Rows(i)("Balance").ToString = "$" Then COA.Rows(i)("Balance") = ""

            If Left(COA.Rows(i)("Balance").ToString, 1) = "-" Then COA.Rows(i)("Balance") = "(" & COA.Rows(i)("Balance").replace("-", "") & ")"
            If Val(COA.Rows(i)("Level").ToString) > DetailLevel Then COA.Rows(i).Delete()
        Next

        COA.AcceptChanges()

        If NoZeros = "off" Then
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Balance") = "" And COA.Rows(i)("Account_Type").ToString < 90 Then COA.Rows(i).Delete()
            Next
        End If

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style1, Style2, Style3, Style4 As String
        For i = 0 To COA.Rows.Count - 1
            Style1 = "text-align:left; font-size:0pt; padding: 0px 0px 0px 0px"
            Style2 = "text-align:left; font-size:8pt; padding: 1px 1px 1px " & Val(COA.Rows(i)("Padding").ToString) + 15 & "px"
            Style3 = "text-align:right; font-size:8pt; padding: 0px 0px 0px 0px; max-width: 1in; min-width: 1in;"
            Style4 = "font-size:0pt"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style1 = Style1 & "; font-weight:bold;font-size:0pt;padding-top:30px"
                Style2 = Style2 & "; font-weight:bold"
                Style3 = Style3 & "; font-weight:bold"
                Style4 = Style4 & ";text-align:left"
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Or COA.Rows(i)("Totalling_Minus").ToString <> "" Then
                Style1 = Style1 & ";border-bottom:solid 0px;border-color:black;"
                Style2 = Style2 & "; border-top: 0x solid black;border-bottom:solid 0px;border-color:black;"
                Style3 = Style3 & ";border-color:black;"
                Style4 = Style4 & ";border-bottom:solid 0px;border-color:black;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style2, COA.Rows(i)("Name").ToString, Style3, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Balance") + "</span>", Style4, COA.Rows(i)("fk_currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("Totalling_Minus")

            ' Rename the existing column name to value
            For i = 0 To COA.Columns.Count - 1
                COA.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = COA.Columns.Count To 25
                COA.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = COA.NewRow()
            excelheader("value0") = "Account No"
            excelheader("value1") = "Name"
            excelheader("value2") = "Currency"
            excelheader("value3") = "Balance"
            COA.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To COA.Columns.Count - 1
                COA.Rows(0)(i) = "<b>" & COA.Rows(0)(i).ToString & "</b>"
            Next

            ' Account No
            If Request.Form("Ac") <> "on" Then
                COA.Columns.Remove("value0")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            RPT_Excel.DataSource = COA
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub
    Private Sub XMLBalance()
        Dim AsAt As String = Request.Form("date1")
        Dim ToDate As String = Request.Form("date2")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim NoZeros As String = Request.Form("showZeros")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        If AsAt = "" Then AsAt = Now().ToString("yyyy-MM-dd")
        Dim Dat As Date

        Try
            Dat = AsAt
        Catch ex As Exception
            Dat = Now()
        End Try
        Dim year As New DateTime(Dat.Year, 1, 1)
        If DetailLevel = 0 Then DetailLevel = 7

        Dim COA, Bal, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Totalling_Minus From ACC_GL_Accounts order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID,(Select top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <=@dateBefore order by Transaction_Date desc, rowID desc) as BeforeBalance, (Select top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <=@date order by Transaction_Date desc, rowID desc) as Balance from ACC_GL gl1 where Transaction_Date <=@date"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", Dat)
        SQLCommand.Parameters.AddWithValue("@dateBefore", year)
        DataAdapter.Fill(Bal)
        'System.diagnostics.Debug.WriteLine(SQLCommand.CommandText)
        COA.Columns.Add("Balance", GetType(String))
        COA.Columns.Add("BeforeBalance", GetType(String))
        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))

        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        For i = 0 To COA.Rows.Count - 1
            For ii = 0 To Bal.Rows.Count - 1
                If COA.Rows(i)("Account_ID").ToString = Bal.Rows(ii)("Account_ID").ToString Then
                    COA.Rows(i)("Balance") = Bal.Rows(ii)("Balance")
                    COA.Rows(i)("BeforeBalance") = Bal.Rows(ii)("BeforeBalance")
                    Exit For
                End If
            Next
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 10 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 10 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "<div style='min-width: 0.5in; max-width:0.5in;'></div>" ' hard coded
        Next

        Dim Total As Decimal = 0
        Dim Account As String = ""
        For j = 1 To 2
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                ' If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
                '     Account = COA.Rows(i)("Account_No").ToString
                '     Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
                '     For ii = 0 To Plus.Length - 1
                '         Dim Dash() As String = Plus(ii).Split("-")
                '         Dim Start As String = Trim(Dash(0))
                '         Dim Endd As String
                '         If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                '         For iii = 0 To COA.Rows.Count - 1
                '             If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                '             If Trim(COA.Rows(iii)("Account_No").ToString) >= Start Then Total = Total - Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                '         Next
                '     Next
                ' End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("Balance") = Total
                Next
            Next
        Next

        Total = 0
        Account = ""
        For j = 1 To 2
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start Then Total = Total + Val(COA.Rows(iii)("BeforeBalance").ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total - Val(COA.Rows(iii)("BeforeBalance").ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("BeforeBalance") = Total
                Next
            Next
        Next


        For i = 0 To COA.Rows.Count - 1
            If Left(COA.Rows(i)("Account_No").ToString, 1) > "3" Then COA.Rows(i).Delete()
        Next

        COA.AcceptChanges()

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("BeforeBalance") = Math.Round(Val(COA.Rows(i)("BeforeBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValue
                    Dim denominatedValue2 As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("BeforeBalance") = denominatedValue2
                End If
            Next
        End If

        For i = 0 To COA.Rows.Count - 1
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("Balance") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("BeforeBalance") = Format(Val(COA.Rows(i)("BeforeBalance").ToString), "$#,###")
            Else
                COA.Rows(i)("Balance") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("BeforeBalance") = Format(Val(COA.Rows(i)("BeforeBalance").ToString), "$#,###.00")
            End If


            If COA.Rows(i)("Balance").ToString = "$.00" Or COA.Rows(i)("Balance").ToString = "$" Then COA.Rows(i)("Balance") = ""
            If COA.Rows(i)("BeforeBalance").ToString = "$.00" Or COA.Rows(i)("BeforeBalance").ToString = "$" Then COA.Rows(i)("BeforeBalance") = ""

            If Left(COA.Rows(i)("Balance").ToString, 1) = "-" Then COA.Rows(i)("Balance") = "(" & COA.Rows(i)("Balance").replace("-", "") & ")"
            If Left(COA.Rows(i)("BeforeBalance").ToString, 1) = "-" Then COA.Rows(i)("BeforeBalance") = "(" & COA.Rows(i)("BeforeBalance").replace("-", "") & ")"
            If Val(COA.Rows(i)("Level").ToString) > DetailLevel Then COA.Rows(i).Delete()
        Next

        COA.AcceptChanges()


        If NoZeros = "off" Then
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Balance") = "" And COA.Rows(i)("Account_Type").ToString < 90 Then COA.Rows(i).Delete()
            Next
        End If



        COA.AcceptChanges()

        'Remoce the columns we dont want to show
        COA.Columns.Remove("Account_ID")
        COA.Columns.Remove("fk_Currency_ID")
        COA.Columns.Remove("Account_Type")
        COA.Columns.Remove("Totalling")
        COA.Columns.Remove("Padding")
        COA.Columns.Remove("Level")
        COA.Columns("Balance").ColumnName = "Current_Balance"
        COA.Columns("BeforeBalance").ColumnName = "Start_Of_Year"

        Dim ds As New DataSet
        ds.Tables.Add(COA)

        Dim xmlData As String = ds.GetXml()

        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True

        Conn.Close()


    End Sub
    Private Sub XMLSummaryTrial()
        Dim Language As Integer = Request.Form("language")
        Dim firstDate As String
        Dim seconDate As String
        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Summary Trial Balance<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'>Account Name</span><span style='position: absolute; margin-left: 1.7in;'>Beginning Balance</span><span style='position: absolute; margin-left: 3.3in'>Debit</span><span style='position: absolute; margin-left: 4.5in'>Credit</span><span style='position: absolute; margin-left: 5.5in'>Net actvity</span><span style='position: absolute; margin-left: 6.8in;'>Closing Balance</span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Resumen Del Balance De Prueba<br/>Desde " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso en " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'>Nombre De La Cuenta</span><span style='position: absolute; margin-left: 1.7in;'>Balance Inicial</span><span style='position: absolute; margin-left: 3.3in'>Débito</span><span style='position: absolute; margin-left: 4.5in'>Crédito</span><span style='position: absolute; margin-left: 5.5in'>Actividad Neto</span><span style='position: absolute; margin-left: 6.8in;'>Balance De Cierre</span></div>"
        End If

        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No;"
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No;"
        End If
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@enddate", DatStart)
        SQLCommand.Parameters.AddWithValue("@date", DatSecond)
        DataAdapter.Fill(COA)


        'System.diagnostics.Debug.WriteLine(SQLCommand.CommandText + DatSecond.ToString)

        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("NetString", GetType(String))
        COA.Columns.Add("NextDateBalanceString", GetType(String))



        Dim finalCredit As Double
        Dim finalDebit As Double
        Dim finalNet As Double

        Dim COACount As Int32 = COA.Rows.Count - 1
        For i = 0 To COA.Rows.Count - 1
            finalCredit = finalCredit + COA.Rows(i)("Credit")
            finalDebit = finalDebit + COA.Rows(i)("Debit")
            finalNet = finalNet + COA.Rows(i)("NetActivity")
        Next
        Try
            Dim newRow As DataRow = COA.NewRow()
            newRow.BeginEdit()
            ' newRow("Balance") = COA.Rows(COACount)("Balance")
            newRow("Credit") = finalCredit
            newRow("Debit") = finalDebit
            newRow("NetActivity") = finalNet
            newRow("Name") = "0001-01-01"

            newRow("Account_Type") = "33"
            newRow.EndEdit()
            COA.Rows.Add(newRow)
        Catch ex As Exception

        End Try


        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    Dim denominatedValueNext As Double = Convert.ToDouble(Val(COA.Rows(i)("NextDateBalance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValueCurrent
                    COA.Rows(i)("NextDateBalance") = denominatedValueNext
                End If
            Next
        End If

        For i = 0 To COA.Rows.Count - 1

            ' Format all the output for the paper
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit").ToString), "$#,###")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit").ToString), "$#,###")
                COA.Rows(i)("NetString") = Format(Val(COA.Rows(i)("NetActivity").ToString), "$#,###")
            Else
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###.00")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit").ToString), "$#,###.00")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit").ToString), "$#,###.00")
                COA.Rows(i)("NetString") = Format(Val(COA.Rows(i)("NetActivity").ToString), "$#,###.00")
            End If

            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("NextDateBalanceString").ToString, 1) = "-" Then COA.Rows(i)("NextDateBalanceString") = "(" & COA.Rows(i)("NextDateBalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
            If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"

            If Left(COA.Rows(i)("NetString").ToString, 1) = "-" Then COA.Rows(i)("NetString") = "(" & COA.Rows(i)("NetString").replace("-", "") & ")"
            'If Val(COA.Rows(i)("Level").ToString) > 1 Then COA.Rows(i).Delete()
            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("NextDateBalanceString").ToString = "$.00" Or COA.Rows(i)("NextDateBalanceString").ToString = "$" Then COA.Rows(i)("NextDateBalanceString") = ""
            If COA.Rows(i)("CreditString").ToString = "$.00" Or COA.Rows(i)("CreditString").ToString = "$" Then COA.Rows(i)("CreditString") = ""
            If COA.Rows(i)("DebitString").ToString = "$.00" Or COA.Rows(i)("DebitString").ToString = "$" Then COA.Rows(i)("DebitString") = ""
            If COA.Rows(i)("NetString").ToString = "$.00" Or COA.Rows(i)("NetString").ToString = "$" Then COA.Rows(i)("NetString") = ""

        Next

        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            ' Delete the rows that arnt above the detail level 
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If COA.Rows(i)("CreditString").ToString = "" And COA.Rows(i)("DebitString").ToString = "" Then
                    COA.Rows(i).Delete()
                End If
            End If
        Next i
        COA.AcceptChanges()

        COA.Columns.Remove("Account_ID")
        COA.Columns.Remove("fk_Currency_ID")
        COA.Columns.Remove("Account_Type")
        COA.Columns.Remove("Totalling")
        COA.Columns.Remove("Cash")
        COA.Columns.Remove("Active")
        COA.Columns.Remove("memo2")
        COA.Columns.Remove("Balance")
        COA.Columns.Remove("NextDateBalance")
        COA.Columns.Remove("Credit")
        COA.Columns.Remove("Debit")
        COA.Columns.Remove("NetActivity")

        COA.Columns("BalanceString").ColumnName = "Beginning_Balance"
        COA.Columns("NextDateBalanceString").ColumnName = "Closing_Balance"
        COA.Columns("CreditString").ColumnName = "Credit"
        COA.Columns("DebitString").ColumnName = "Debit"
        COA.Columns("NetString").ColumnName = "Net_Activity"

        Dim ds As New DataSet
        ds.Tables.Add(COA)


        Dim xmlData As String = ds.GetXml()


        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True


        Conn.Close()


    End Sub
    Private Sub XMLDetailTrial()
        Dim StartDate As String
        Dim EndDate As String
        Dim accNo As String
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        StartDate = Request.Form("StartDate")
        EndDate = Request.Form("EndDate")
        accNo = Request.Form("accNo")

        If StartDate = "" Then StartDate = Now().ToString("yyyy-MM-dd")
        If EndDate = "" Then EndDate = Now().AddDays(-30).ToString("yyyy-MM-dd")

        'Get account name
        Conn.Open()
        Dim querystr As String = "SELECT Name FROM ACC_GL_Accounts WHERE Account_No = " + accNo + ";"
        Dim mycmd As New SqlCommand(querystr, Conn)
        Dim value As Object = mycmd.ExecuteScalar()
        Conn.Close()

        HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
        'Set the page header, this is below the SQL so we can get the currency
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Detail Trial Balance<br/>From " & StartDate & " to " & EndDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><br/>" + Request.Form("accNo") + " " + value.ToString() + "</span><br><div style='Width: 8.5in; position: absolute; font-weight: bold !important;'><span style='position: absolute; margin-left: -0.2in'>Posting Date</span><span style='position: absolute; margin-left: 1in'>Doc No</span><span style='position: absolute; margin-left: 2in'>Description</span><span style='position: absolute; margin-left: 3.7in;'>Debit</span><span style='position: absolute; margin-left: 5in'>Credit</span><span style='position: absolute; margin-left: 5.7in'>Balance</span><span style='position: absolute; margin-left: 6.5in'>Entry No.</span><span style='position: absolute; margin-left: 7.3in'>Currency</span></div>"

        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "SELECT rowID,Transaction_Date, Transaction_No, Document_Type, Debit_Amount, Credit_Amount, Balance, Memo, Document_ID, fk_Account_ID,Account_No,ACC_GL.fk_Currency_ID FROM ACC_GL LEFT JOIN ACC_GL_Accounts on ACC_GL_Accounts.Account_ID = ACC_GL.fk_Account_ID WHERE ((Transaction_Date >= @startDate AND Transaction_Date <= @endDate) AND Account_No = @accNo) ORDER BY Transaction_Date asc, rowID desc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@startDate", StartDate)
        SQLCommand.Parameters.AddWithValue("@endDate", EndDate)
        SQLCommand.Parameters.AddWithValue("@accNo", accNo)
        DataAdapter.Fill(COA)




        COA.Columns.Add("DebitString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("BalanceString", GetType(String))

        Dim finalCredit As Double
        Dim finalDebit As Double
        Dim COACount As Int32 = COA.Rows.Count - 1
        For i = 0 To COA.Rows.Count - 1
            finalCredit = finalCredit + COA.Rows(i)("Credit_Amount")
            finalDebit = finalDebit + COA.Rows(i)("Debit_Amount")
        Next
        Try
            Dim newRow As DataRow = COA.NewRow()
            Dim transactionDate As Date
            transactionDate = "0001-01-01"

            newRow.BeginEdit()
            newRow("Balance") = COA.Rows(COACount)("Balance")
            newRow("Credit_Amount") = finalCredit
            newRow("Debit_Amount") = finalDebit
            newRow("memo") = Request.Form("accNo") + " " + value.ToString()
            newRow("Transaction_Date") = transactionDate
            newRow.EndEdit()
            COA.Rows.Add(newRow)
        Catch ex As Exception

        End Try




        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString))
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValue
                    Dim denominatedValue2 As Double = Convert.ToDouble(Val(COA.Rows(i)("Credit_Amount").ToString)) / Denom
                    COA.Rows(i)("Credit_Amount") = denominatedValue2
                    Dim denominatedValue3 As Double = Convert.ToDouble(Val(COA.Rows(i)("Debit_Amount").ToString)) / Denom
                    COA.Rows(i)("Debit_Amount") = denominatedValue3
                End If
            Next
        End If

        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###")
            Else
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###.00")
                COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###.00")
            End If

            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("CreditString").ToString = "$.00" Or COA.Rows(i)("CreditString").ToString = "$" Then COA.Rows(i)("CreditString") = ""
            If COA.Rows(i)("DebitString").ToString = "$.00" Or COA.Rows(i)("DebitString").ToString = "$" Then COA.Rows(i)("DebitString") = ""

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
            If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
            If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"
        Next

        COA.AcceptChanges()
        COA.Columns.Remove("rowID")
        COA.Columns.Remove("Transaction_Date")
        COA.Columns.Remove("Document_Type")
        COA.Columns.Remove("Debit_Amount")
        COA.Columns.Remove("Credit_Amount")
        COA.Columns.Remove("Balance")
        COA.Columns.Remove("Document_ID")
        COA.Columns.Remove("fk_Account_ID")
        COA.Columns.Remove("Account_No")
        COA.Columns.Remove("fk_Currency_ID")


        COA.Columns("BalanceString").ColumnName = "Balance"
        COA.Columns("CreditString").ColumnName = "Credit"
        COA.Columns("DebitString").ColumnName = "Debit"


        Dim ds As New DataSet
        ds.Tables.Add(COA)


        Dim xmlData As String = ds.GetXml()

        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True

        Conn.Close()


    End Sub
    Private Sub XMLProfitLossM2M()

        Dim firstDate As String
        Dim seconDate As String
        Dim dateArray(100) As String
        Dim monthArray() As String = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")



        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        Dim monthDifference As Int32
        'Need to figure out how many months are between the two selected dates
        monthDifference = (DatSecond.Month - DatStart.Month) + 12 * (DatSecond.Year - DatStart.Year)

        'Loop through and add all the months to the array
        Dim sqlInsert As String = ""
        Dim sqlInsertHeaders As String = ""
        Dim tempDate As Date
        For i = 0 To monthDifference
            tempDate = DatStart.AddMonths(i)
            dateArray(i) = tempDate.ToString("yyyy-MM-dd")
            sqlInsert = sqlInsert + "outer apply(select top 1 (Balance) as Month" & i & "Balance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <=@date" & i & " order by Transaction_Date desc, rowID desc)  as Month" & i & " "
            sqlInsertHeaders = sqlInsertHeaders + ", CONVERT(varchar(100), Month" & i & "Balance) as " + monthArray(Month(tempDate) - 1) + ""

        Next

        Dim COA, Bal1, Bal2, Report As New DataTable

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()
        SQLCommand.CommandTimeout = 500
        SQLCommand.CommandText = "Select Totalling,Totalling_Minus,Account_Type,Account_No, Name, Totalling_Minus " + sqlInsertHeaders + " From ACC_GL_Accounts outer apply( select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @dateStart AND @dateEnd order by Transaction_Date desc, rowID desc) as tid " + sqlInsert + "WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 order by Account_No"
        SQLCommand.Parameters.Clear()

        For i = 0 To monthDifference
            SQLCommand.Parameters.AddWithValue("@date" & i, dateArray(i))
        Next
        SQLCommand.Parameters.AddWithValue("@dateStart", DatStart.ToString("yyyy-MM-dd"))
        SQLCommand.Parameters.AddWithValue("@dateEnd", DatSecond.ToString("yyyy-MM-dd"))
        'System.diagnostics.Debug.WriteLine(SQLCommand.CommandText)
        DataAdapter.Fill(COA)


        'Get the totals 
        For a = 0 To monthDifference
            tempDate = DatStart.AddMonths(a)
            Dim Total As Decimal = 0
            Dim Account As String = ""
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(monthArray(Month(tempDate) - 1)).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                ' If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
                '     Account = COA.Rows(i)("Account_No").ToString
                '     Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
                '     For ii = 0 To Plus.Length - 1
                '         Dim Dash() As String = Plus(ii).Split("-")
                '         Dim Start As String = Trim(Dash(0))
                '         Dim Endd As String
                '         If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                '         For iii = 0 To COA.Rows.Count - 1
                '             If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                '             If Trim(COA.Rows(iii)("Account_No").ToString) >= Start Then Total = Total - Val(COA.Rows(iii)(monthArray(Month(tempDate) - 1)).ToString.Replace(",", "").Replace("$", ""))
                '         Next
                '     Next
                ' End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(monthArray(Month(tempDate) - 1)) = Total
                Next
            Next
        Next


        'Format everything before we put it to XML
        For a = 0 To COA.Rows.Count - 1
            For i = 0 To monthDifference
                tempDate = DatStart.AddMonths(i)
                If Request.Form("Round") = "on" Then
                    COA.Rows(a)(monthArray(Month(tempDate) - 1)) = Math.Round(Val(COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString)) / Denom
                    COA.Rows(a)(monthArray(Month(tempDate) - 1)) = denominatedValue
                End If
                If Request.Form("Round") = "on" Then
                    COA.Rows(a)(monthArray(Month(tempDate) - 1)) = Format(Val(COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString), "$#,###")
                Else
                    COA.Rows(a)(monthArray(Month(tempDate) - 1)) = Format(Val(COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString), "$#,###.00")
                End If
                If COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString = "$.00" Or COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString = "$" Then COA.Rows(a)(monthArray(Month(tempDate) - 1)) = ""
            Next
        Next




        For a = 0 To COA.Rows.Count - 1
            Dim Deleted As Boolean = False
            For i = 0 To monthDifference
                If (Deleted = False) Then
                    If Request.Item("showZeros") = "off" And COA.Rows(a)("Account_Type") < 90 Then
                        If COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString = "" Then
                            COA.Rows(a).Delete()
                            Deleted = True
                        End If
                    End If
                End If
            Next
        Next




        COA.AcceptChanges()
        COA.Columns.Remove("Account_Type")
        COA.Columns.Remove("Totalling_Minus")
        COA.Columns.Remove("Totalling")
        Dim ds As New DataSet
        ds.Tables.Add(COA)


        Dim xmlData As String = ds.GetXml()

        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True


        Conn.Close()

    End Sub
    Private Sub XMLProfitLoss()

        Dim firstDate As String
        Dim seconDate As String
        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If
        'System.diagnostics.Debug.WriteLine("HERE")

        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        Dim COA, Bal1, Bal2, Report As New DataTable

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ThisDateBalance AS Balance, Totalling_Minus, Exchange_Account_ID, Transaction_No,NextDateBalance From ACC_GL_Accounts outer apply(select top 1 * from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <=@date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select top 1 (Balance) as NextDateBalance from ACC_GL where fk_Account_ID=Account_ID and Transaction_Date <=@endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 order by Account_No;"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", DatStart)
        SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
        DataAdapter.Fill(COA)


        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("Dollar_Difference", GetType(Decimal))
        COA.Columns.Add("Percent_Difference", GetType(String))
        COA.Columns.Add("NextDateBalanceString", GetType(String))
        COA.Columns.Add("DifferenceString", GetType(String))

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    Dim denominatedValueNext As Double = Convert.ToDouble(Val(COA.Rows(i)("NextDateBalance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValueCurrent
                    COA.Rows(i)("NextDateBalance") = denominatedValueNext
                End If

            Next
        End If

        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        For i = 0 To COA.Rows.Count - 1
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
            'If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded
        Next

        Dim Total As Decimal = 0
        Dim Account As String = ""
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                    For iii = 0 To COA.Rows.Count - 1
                        If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                        If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                    Next
                Next
            End If
            ' If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
            '     Account = COA.Rows(i)("Account_No").ToString
            '     Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
            '     For ii = 0 To Plus.Length - 1
            '         Dim Dash() As String = Plus(ii).Split("-")
            '         Dim Start As String = Trim(Dash(0))
            '         Dim Endd As String
            '         If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
            '         For iii = 0 To COA.Rows.Count - 1
            '             If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
            '             If Trim(COA.Rows(iii)("Account_No").ToString) >= Start Then Total = Total - Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
            '         Next
            '     Next
            ' End If
            For ii = 0 To COA.Rows.Count - 1
                If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("Balance") = Total
            Next
        Next

        Total = 0
        Account = ""
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                    For iii = 0 To COA.Rows.Count - 1
                        If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                        If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("NextDateBalance").ToString.Replace(",", "").Replace("$", ""))
                    Next
                Next
            End If
            ' If COA.Rows(i)("Totalling_Minus").ToString <> "" Then
            '     Account = COA.Rows(i)("Account_No").ToString
            '     Dim Plus() As String = COA.Rows(i)("Totalling_Minus").ToString.Split("+")
            '     For ii = 0 To Plus.Length - 1
            '         Dim Dash() As String = Plus(ii).Split("-")
            '         Dim Start As String = Trim(Dash(0))
            '         Dim Endd As String
            '         If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
            '         For iii = 0 To COA.Rows.Count - 1
            '             If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
            '             If Trim(COA.Rows(iii)("Account_No").ToString) >= Start Then Total = Total - Val(COA.Rows(iii)("NextDateBalance").ToString.Replace(",", "").Replace("$", ""))
            '         Next
            '     Next
            ' End If
            For ii = 0 To COA.Rows.Count - 1
                If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("NextDateBalance") = Total
            Next
            Try
                COA.Rows(i)("Dollar_Difference") = COA.Rows(i)("NextDateBalance") - COA.Rows(i)("Balance")
                COA.Rows(i)("Percent_Difference") = FormatPercent((COA.Rows(i)("NextDateBalance") - COA.Rows(i)("Balance")) / COA.Rows(i)("Balance"), , TriState.True, TriState.True)
            Catch
            End Try

        Next

        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"



            If Left(COA.Rows(i)("NextDateBalanceString").ToString, 1) = "-" Then COA.Rows(i)("NextDateBalanceString") = "(" & COA.Rows(i)("NextDateBalanceString").replace("-", "") & ")"

            If Request.Form("Round") = "on" Then
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###")
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Dollar_Difference").ToString), "$#,###")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
            Else
                COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###.00")
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Dollar_Difference").ToString), "$#,###.00")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            End If


            If Left(COA.Rows(i)("DifferenceString").ToString, 1) = "-" Then COA.Rows(i)("DifferenceString") = "(" & COA.Rows(i)("DifferenceString").replace("-", "") & ")"
            'If Val(COA.Rows(i)("Level").ToString) > 1 Then COA.Rows(i).Delete()

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("NextDateBalanceString").ToString = "$.00" Or COA.Rows(i)("NextDateBalanceString").ToString = "$" Then COA.Rows(i)("NextDateBalanceString") = ""
            If COA.Rows(i)("Percent_Difference").ToString = "0.00%" Then COA.Rows(i)("Percent_Difference") = ""
            If COA.Rows(i)("DifferenceString").ToString = "$.00" Or COA.Rows(i)("DifferenceString").ToString = "$" Then COA.Rows(i)("DifferenceString") = ""
            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded
        Next
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            ' Delete the rows that arnt above the detail level 
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If COA.Rows(i)("BalanceString").ToString = "" And COA.Rows(i)("NextDateBalanceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                ElseIf COA.Rows(i)("DifferenceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                End If

            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        'Remoce the columns we dont want to show
        COA.Columns.Remove("Account_ID")
        COA.Columns.Remove("fk_Currency_ID")
        COA.Columns.Remove("Account_Type")
        COA.Columns.Remove("Direct_Posting")
        COA.Columns.Remove("fk_Linked_ID")
        COA.Columns.Remove("Totalling")
        COA.Columns.Remove("Active")
        COA.Columns.Remove("Cash")
        COA.Columns.Remove("Exchange_Account_ID")
        COA.Columns.Remove("Balance")
        COA.Columns.Remove("NextDateBalance")
        COA.Columns.Remove("Padding")
        COA.Columns.Remove("Level")

        COA.Columns("BalanceString").ColumnName = "Beginning_Balance"
        COA.Columns("NextDateBalanceString").ColumnName = "Closing_Balance"

        Dim ds As New DataSet
        ds.Tables.Add(COA)


        Dim xmlData As String = ds.GetXml()

        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True


        Conn.Close()

    End Sub
    ' Income Statement
    Private Sub PrintProfitLoss()

        Dim Language As Integer = Request.Form("language")
        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As String
        Dim seconDate As String
        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"
        Dim ProfitAndLoss As String
        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim Acc_No As String = Request.Form("Ac")
        Dim Percentage As String = Request.Form("Perce")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Default date give today's date and a year before
        If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
        Dim DatStart, DatSecond As Date
        Dim HF_Acc As String = ""
        Dim HF_Pre As String = ""
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        If Acc_No = "on" Then
            HF_Acc = "width:60px; font-size:8pt~Act No"
        Else
            HF_Acc = "width:0px; font-size:0pt~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            If Percentage = "on" Then
                HF_Pre = "text-align:right; width:160px; font-size:8pt~Sales/Expenses(%)"
            Else
                HF_Pre = "text-align:right; width:0px; font-size:0pt~"
            End If
            HF_PrintHeader.Value = "text-align:left; " & HF_Acc & "~text-align:left; width:350px; font-size:8pt~Account Description~text-align:right; width:120px; font-size:8pt~Dollar Amount~" & HF_Pre & "~text-align:centre; width:0px; font-size:0pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Income Statement " + DenomString + "<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            If Percentage = "on" Then
                HF_Pre = "text-align:right; width:160px; font-size:8pt~Ventas/Gastos(%)"
            Else
                HF_Pre = "text-align:right; width:0px; font-size:0pt~"
            End If
            HF_PrintHeader.Value = "text-align:left; " & HF_Acc & "~text-align:left; width:350px; font-size:8pt~Descripción De Cuenta~text-align:right; width:120px; font-size:8pt~Monto En Dólares~" & HF_Pre & "~text-align:centre; width:0px; font-size:0pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado De Resultados " + DenomString + "<br/>Desde " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            SQLCommand.Parameters.AddWithValue("@date", DatStart)
            SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("Dollar_Difference", GetType(Decimal))
        COA.Columns.Add("DifferenceString", GetType(String))
        COA.Columns.Add("Percent_Difference", GetType(String))
        COA.Columns.Add("Percent_DifferenceString", GetType(String))

        'Denomination And rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString))
                    'COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)("Balance").ToString)) / Denom
                    'Dim denominatedValueNext As Double = Convert.ToDouble(Val(COA.Rows(i)("NextDateBalance").ToString)) / Denom
                    COA.Rows(i)("Balance") = denominatedValueCurrent
                    'COA.Rows(i)("NextDateBalance") = denominatedValueNext
                End If

            Next
        End If

        ' Give Padding
        For i = 0 To COA.Rows.Count - 1
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
            'If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded
        Next

        Dim Total As Decimal = 0
        Dim Account As String = ""
        ' Calculating Sub-Total and Total
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                    For iii = 0 To COA.Rows.Count - 1
                        If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                        If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)("Balance").ToString.Replace(",", "").Replace("$", ""))
                    Next
                Next
            End If
            For ii = 0 To COA.Rows.Count - 1
                If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)("Balance") = Total
            Next
        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        If rowIncome.Length > 0 Then
            TotalIncome = rowIncome(0).Item("Balance")
        End If
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        If rowCost.Length > 0 Then
            TotalCost = rowCost(0).Item("Balance")
        End If
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")
        If rowExpense.Length > 0 Then
            TotalExpenses = rowExpense(0).Item("Balance")
        End If

        'Set the percentages
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Totalling").ToString <> "" Then
                Total = 0
                Account = COA.Rows(i)("Account_No").ToString
                Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                For ii = 0 To Plus.Length - 1
                    Dim Dash() As String = Plus(ii).Split("-")
                    Dim Start As String = Trim(Dash(0))
                    Dim Endd As String
                    If Dash.Length = 1 Then
                        Endd = Trim(Dash(0))
                    Else
                        Endd = Trim(Dash(1))
                    End If
                    For iii = 0 To COA.Rows.Count - 1

                        If COA.Rows(iii)("Account_Type") < 90 Then
                            Try
                                If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "4" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalIncome)) * 100
                                End If
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "5" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalCost)) * 100
                                End If
                                If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And Trim(COA.Rows(iii)("Account_No").ToString).Substring(0, 1) = "6" Then
                                    COA.Rows(iii)("Percent_Difference") = (Double.Parse(COA.Rows(iii)("Balance").ToString) / Double.Parse(TotalExpenses)) * 100
                                End If
                            Catch Ex As Exception
                            End Try
                        End If
                    Next
                Next
            End If

        Next

        For i = 0 To COA.Rows.Count - 1
            ' Format all the output for the paper
            COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            COA.Rows(i)("Percent_Difference") = Format(Val(COA.Rows(i)("Percent_Difference").ToString), "##.00") + "%"

            If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"

            If Request.Form("Round") = "on" Then
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###")
            Else
                COA.Rows(i)("DifferenceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
                COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
            End If

            If Left(COA.Rows(i)("DifferenceString").ToString, 1) = "-" Then COA.Rows(i)("DifferenceString") = "(" & COA.Rows(i)("DifferenceString").replace("-", "") & ")"
            'If Val(COA.Rows(i)("Level").ToString) > 1 Then COA.Rows(i).Delete()

            If COA.Rows(i)("BalanceString").ToString = "$.00" Or COA.Rows(i)("BalanceString").ToString = "$" Then COA.Rows(i)("BalanceString") = ""
            If COA.Rows(i)("Percent_Difference").ToString = ".00%" Or COA.Rows(i)("Percent_Difference").ToString = "00%" Then COA.Rows(i)("Percent_Difference") = ""
            If COA.Rows(i)("DifferenceString").ToString = "$.00" Or COA.Rows(i)("DifferenceString").ToString = "$" Then COA.Rows(i)("DifferenceString") = ""
            If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "" ' hard coded
        Next
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            ' Delete the rows that arnt above the detail level 
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If COA.Rows(i)("BalanceString").ToString = "" Then
                    'If COA.Rows(i)("BalanceString").ToString = "" And COA.Rows(i)("NextDateBalanceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                ElseIf COA.Rows(i)("DifferenceString").ToString = "" Then
                    COA.Rows(i).Delete()
                    AlreadyDeleted = True
                End If

            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next


        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 1.5in; max-width: 1.5in;"
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 3.5in; max-width: 3.5in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 1.1in; max-width: 1.1in;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:15px;"
                Style2 = Style2 & "; padding-bottom:15px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
            End If
            Dim Ac_Style = " font-size:0pt;"
            Dim Per_Style = "font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            If Percentage = "on" Then
                Per_Style = "text-align:right;font-size:8pt;width: 10px;"
            End If

            Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("DifferenceString") + "</span>", Per_Style, COA.Rows(i)("Percent_Difference"), "font-size:0pt; width:0px", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        ProfitAndLoss = Convert.ToDecimal(TotalIncome) - Convert.ToDecimal(TotalCost) - Convert.ToDecimal(TotalExpenses)
        ProfitAndLoss = Format(Val(ProfitAndLoss.ToString), "$#,###.00")

        ' Check ProfitAndLoss Value negative or positive
        If Left(ProfitAndLoss.ToString, 1) = "-" Then
            ProfitAndLoss = "(" & ProfitAndLoss.Replace("-", "") & ")"
            StyleFinish = StyleFinish & "color: red !important;"
        End If

        Style = Style & "padding-bottom:0px;"
        Style2 = "text-align:right; font-size:8pt; min-width: 1.5in; max-width: 1.5in; padding: 0px 0px 0px 0px; font-weight:bold;border-top: px solid black;"

        Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + ProfitAndLoss + "</span>", "font-size:8pt; width:0px ;text-align:right ", "", "font-size:0pt; width:0px", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString")
            COA.Columns.Remove("Dollar_Difference")
            COA.Columns.Remove("Percent_DifferenceString")
            COA.Columns.Remove("Balance")

            ' Adding Total Profit and Loss
            COA.Rows.Add("", "PROFIT AND LOSS", "", ProfitAndLoss, "")

            ' Rename the existing column name to value
            For i = 0 To COA.Columns.Count - 1
                COA.Columns(i).ColumnName = "value" + i.ToString()
            Next

            ' Creating new column to value20
            For i = COA.Columns.Count To 25
                COA.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the header as "value"
            Dim excelheader = COA.NewRow()
            excelheader("value0") = "Account No"
            excelheader("value1") = "Name"
            excelheader("value2") = "Currency"
            excelheader("value3") = "Balance"
            excelheader("value4") = "Percentage"
            COA.Rows.InsertAt(excelheader, 0)

            ' Bold the header.
            For i = 0 To COA.Columns.Count - 1
                COA.Rows(0)(i) = "<b>" & COA.Rows(0)(i).ToString & "</b>"
            Next

            ' Percentage
            If Percentage <> "on" Then
                COA.Columns.Remove("value4")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            ' Account No
            If Request.Form("Ac") <> "on" Then
                COA.Columns.Remove("value0")
                ' Rename the header so the cell shifted to value0
                For i = 0 To COA.Columns.Count - 1
                    COA.Columns(i).ColumnName = "value" + i.ToString()
                Next
            End If

            RPT_Excel.DataSource = COA
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Income Statement Multiperiod Monthly
    Private Sub PrintMonthlyIncStateMultiPer()
        Dim Language As Integer = Request.Form("language")
        Dim firstDate As String = Request.Form("FirstDate")
        Dim seconDate As String = Request.Form("SecondDate")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        'Dim j As Integer = 0
        Dim Level As Integer = 1

        Dim startDate, endDate As Date
        Dim startDate1, endDate1, StyleMonth, Asterix As String
        Dim StyleFinish As String = ""
        Dim HF_Acc As String = ""

        Dim DenomString As String = ""
        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""

        Dim Bal0, Bal1, Bal2 As Decimal
        Dim MonthCount As Integer

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        ' Get the MonthCount Value
        Try
            startDate = firstDate
            endDate = seconDate
            While startDate <> endDate
                startDate = startDate.AddMonths(1)
                MonthCount += 1
                'MonthCount = Convert.ToInt32(Request.Form["SecondDate"].Substring(5, 2)) - Convert.ToInt32(Request.Form["FirstDate"].Substring(5, 2));
            End While
        Catch ex As Exception
            MonthCount = 0
        End Try

        Dim Profitloss(MonthCount) As String

        ' Denomination
        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        Asterix = ""

        ' If seconDate Year and Month is Today's Year and Month then change the date to today
        If Request.Form("SecondDate") = Now().ToString("yyyy-MM") Then
            seconDate = Now().ToString("yyyy-MM-dd")
            firstDate = Now().AddMonths(-MonthCount).ToString("yyyy-MM-01")

            endDate = Now().AddMonths(-MonthCount)
            endDate1 = Now().AddMonths(-MonthCount).ToString("yyyy-MM-dd")
            Asterix = "(*)"
        Else
            ' Default date give today's date
            If firstDate = "" Then
                firstDate = Now().ToString("yyyy-MM-01")
                Asterix = "(*)"
            Else
                ' If exist, take the the first day of month
                startDate = firstDate
            End If
            If seconDate = "" Then
                seconDate = Now().ToString("yyyy-MM-dd")
                endDate = Now()
                Asterix = "(*)"
            Else
                ' If exist, take the the last day of month
                endDate = seconDate
                endDate1 = startDate.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
            End If
        End If

        startDate = firstDate
        startDate1 = startDate.ToString("yyyy-MM-dd")

        Dim Fr_Date As DateTime
        For i = 0 To MonthCount
            If Language = 0 Then
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + startDate.AddMonths(i).ToString("MMMM") + Asterix
            ElseIf Language = 1 Then
                Fr_Date = DateTime.Parse(startDate.AddMonths(i))
                Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + StrConv(Fr_Date.ToString("MMMM"), VbStrConv.ProperCase) + Asterix
            End If

            'startDate = startDate.AddMonths(i + 1)
            'startDate1 = startDate.ToString("yyyy-MM-dd")
        Next

        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt" & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Account Description" & StyleMonth & "~Text-align: Right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Multiperiod Income Statement(Monthly) " + DenomString + "<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt" & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Descripción De Cuenta" & StyleMonth & "~Text-align: Right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Mensual) " + DenomString + "<br/>Desde " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        startDate1 = startDate.ToString("yyyy-MM-dd")

        For i = 0 To MonthCount
            ' Check to it compare partial with partial
            If Request.Form("SecondDate") = Now().ToString("yyyy-MM") Then
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01")
                endDate1 = endDate.AddMonths(i).ToString("yyyy-MM-dd")
            Else
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01")
                endDate1 = startDate.AddMonths(i + 1).AddDays(-1).ToString("yyyy-MM-dd")
            End If

            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between '" & startDate1 & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between '" & startDate1 & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & i.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between '" & startDate1 & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between '" & startDate1 & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & i.ToString
        Next

        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))

        Dim Balance As String
        Dim BalanceString As String

        Balance = ""
        BalanceString = ""

        ' For loop for Calculation and  Formatting
        For j = 0 To MonthCount
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            ' Format all the output for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If

            COA.AcceptChanges()
        Next
        'End While
        ' End of While loop

        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If MonthCount = 0 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf MonthCount = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf MonthCount = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        ' Calculation for Total
        For i = 0 To COA.Rows.Count - 1

            If COA.Rows(i)("Balance0").ToString = "" Then
                Bal0 = 0
            Else
                Bal0 = COA.Rows(i)("Balance0")
            End If
            If MonthCount = 1 Or MonthCount = 2 Then
                If COA.Rows(i)("Balance1").ToString = "" Then
                    Bal1 = 0
                Else
                    Bal1 = COA.Rows(i)("Balance1")
                End If
            End If
            If MonthCount = 2 Then
                If COA.Rows(i)("Balance2").ToString = "" Then
                    Bal2 = 0
                Else
                    Bal2 = COA.Rows(i)("Balance2")
                End If
            End If

            COA.Rows(i)("Total") = (Bal0 + Bal1 + Bal2).ToString

            COA.AcceptChanges()
            ' Format all the output for the paper

            If Request.Form("Round") = "on" Then
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###")
            Else
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            End If

            If Left(COA.Rows(i)("Total").ToString, 1) = "-" Then COA.Rows(i)("Total") = "(" & COA.Rows(i)("Total").replace("-", "") & ")"

            If COA.Rows(i)("Total").ToString = "$.00" Or COA.Rows(i)("Total").ToString = "$" Then COA.Rows(i)("Total") = ""

            COA.AcceptChanges()

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            If MonthCount = 0 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf MonthCount = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf MonthCount = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")

        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish1 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish2 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinishTotal As String = "border-bottom: Double 3px black; border-top: 1px solid black;"

        ' Check if rowIncome, rowCost, and rowExpense have value
        If rowIncome.Length > 0 And rowCost.Length > 0 And rowExpense.Length > 0 Then
            ' Calculating Profit/Loss
            If MonthCount = 0 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                'If Request.Form("Round") = "on" Then
                '    TotalProfitloss = Math.Round(Val(TotalProfitloss.ToString))
                'End If

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf MonthCount = 1 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1)
                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf MonthCount = 2 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                Profitloss2 = Convert.ToDecimal(rowIncome(0).Item("Balance2")) - Convert.ToDecimal(rowCost(0).Item("Balance2")) - Convert.ToDecimal(rowExpense(0).Item("Balance2"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1) + Convert.ToDecimal(Profitloss2)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(Profitloss2.ToString, 1) = "-" Then
                    Profitloss2 = "(" & Profitloss2.Replace("-", "") & ")"
                    StyleFinish2 = StyleFinish2 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinish2 + """>" + Profitloss2 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End If

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export Function
        If Request.Form("expStat") = "on" Then

            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")

            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the total
            Dim excelTotal = exportTable.NewRow()
            excelTotal("value1") = "Profit/Loss"
            For i = 0 To MonthCount
                Profitloss(i) = Convert.ToDecimal(rowIncome(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowCost(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowExpense(0).Item("Balance" + i.ToString))
                excelTotal("value" + (i + 2).ToString) = Profitloss(i)
            Next

            exportTable.Rows.Add(excelTotal)

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To MonthCount
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next
            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i As Integer = 0 To MonthCount
                excelHeader("value" + (i + 2).ToString()) = startDate.AddMonths(i).ToString("MMMM") + Asterix
                'If i = MonthCount Then
                '    excelHeader("value" + (i + 3).ToString()) = "Total"
                'End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If

    End Sub

    ' Income Statement Multiperiod Month-to-Month
    Private Sub PrintMonthToMonthIncStateMultiPer()

        Dim Language As Integer = Request.Form("language")
        Dim seconDate As String = Request.Form("SecMonth")
        Dim Goback As String = Request.Form("goback")
        Dim Show_Per As String = Request.Form("Percentage")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim j As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As Date
        Dim secondDate As Date
        Dim startDate1 As String
        Dim startDate2 As String
        Dim endDate1 As String
        Dim endDate2 As String

        Dim Date1 As String
        Dim Date2 As String
        Dim endDate As Date = seconDate.ToString

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        startDate1 = endDate.AddYears(-(Goback - 1)).ToString("yyyy-MM-dd")

        If endDate.ToString("yyyy-MM") = Now().ToString("yyyy-MM") Then
            endDate1 = Now().AddYears(-(Goback - 1)).ToString("yyyy-MM-dd")
        Else
            endDate1 = endDate.AddYears(-(Goback - 1)).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
        End If

        Date1 = endDate.AddYears(-(Goback - 1)).ToString("MMMM yyyy")

        Dim StyleFinish As String = ""

        Dim DenomString As String = ""
        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""
        Dim Profitloss(Goback - 1) As String

        Dim Bal0 As Decimal
        Dim Bal1 As Decimal
        Dim Bal2 As Decimal
        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Default date give today's date and a year before
        'If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yy-MM")

        Dim StyleMonth As String
        Dim HF_Acc As String = ""
        Dim Fr_Date As DateTime
        'Dim period As Integer = Integer.Parse(Goback)
        firstDate = startDate1
        secondDate = endDate1
        'Header
        If Language = 0 Then
            For j = 0 To Goback - 1

                If endDate = Now().ToString("yyyy-MM") Then
                    Date1 = firstDate.ToString("MMMM yyyy") + "(*)"
                Else
                    Date1 = firstDate.ToString("MMMM yyyy")
                End If

                If j = (Goback - 1) Then
                    Date2 = Date2 + " and " + firstDate.ToString("MMMM yyyy")
                ElseIf j = (Goback - 2) Then
                    Date2 = Date2 + "" + firstDate.ToString("MMMM yyyy")
                Else
                    Date2 = Date2 + "" + firstDate.ToString("MMMM yyyy") + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + Date1

                startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd")
                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                firstDate = startDate2
                secondDate = endDate2

            Next
        ElseIf Language = 1 Then
            For j = 0 To Goback - 1
                Fr_Date = DateTime.Parse(firstDate)
                Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")

                If endDate = Now().ToString("yyyy-MM") Then
                    Date1 = Fr_Date.ToString("MMMM yyyy") + "(*)"
                Else
                    Date1 = Fr_Date.ToString("MMMM yyyy")
                End If

                If j = (Goback - 1) Then
                    Date2 = Date2 + " y " + Fr_Date.ToString("MMMM yyyy")
                ElseIf j = (Goback - 2) Then
                    Date2 = Date2 + "" + Fr_Date.ToString("MMMM yyyy")
                Else
                    Date2 = Date2 + "" + Fr_Date.ToString("MMMM yyyy") + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + StrConv(Fr_Date.ToString("MMMM yyyy"), VbStrConv.ProperCase)
                startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd")
                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                firstDate = startDate2
                secondDate = endDate2

            Next
        End If

        Dim Per_opt As String = ""
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If
        If Show_Per = "on" Then
            If Language = 0 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Percentage(%)"
            ElseIf Language = 1 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Porcentaje(%)"
            End If
        Else
            Per_opt = "~width:0px;font-size:0pt~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; font-size:8pt~Account Description" & StyleMonth & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Income Statement(Month To Month) " + DenomString + "<br/>For " & Date2 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Descripción De Cuenta" & StyleMonth & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Mensual) " + DenomString + "<br/>Para  " & Date2 & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        firstDate = startDate1
        secondDate = endDate1
        startDate2 = startDate1
        endDate2 = endDate1


        For j = 0 To Goback - 1

            Query1 = Query1 + ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + j.ToString()
            Query2 = Query2 + ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + j.ToString()

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd")
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
            firstDate = startDate2
            secondDate = endDate2
        Next
        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))
        COA.Columns.Add("Per", GetType(String))

        Dim Balance As String = ""
        Dim BalanceString As String = ""

        ' For loop for Calculation and  Formatting
        For j = 0 To Goback - 1

            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            ' Format all the output for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If

            COA.AcceptChanges()
        Next
        ' End of for loop

        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        ' Calculation for Per
        For i = 0 To COA.Rows.Count - 1

            If COA.Rows(i)("Balance0").ToString = "" Then
                Bal0 = 0
            ElseIf COA.Rows(i)("Balance0") <> 0 Then
                Bal0 = COA.Rows(i)("Balance0")
                If j >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal1 = 0
                    Else
                        Bal1 = COA.Rows(i)("Balance1")
                        COA.Rows(i)("Per") = (((Bal1 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                        If j = 3 Then
                            If COA.Rows(i)("Balance2").ToString = "" Then
                                Bal2 = 0
                            Else
                                Bal2 = COA.Rows(i)("Balance2")
                                COA.Rows(i)("Per") = (((Bal2 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                            End If

                        End If
                    End If
                End If
            End If

            COA.AcceptChanges()
            ' Format all the output for the paper  
            If COA.Rows(i)("Per").ToString <> "" Then
                COA.Rows(i)("Per") = (Math.Round(Val(COA.Rows(i)("Per").ToString), 2)).ToString() & "%"
            End If

            'String.Format("{0:#.##%}", COA.Rows(i)("Per"))

            'COA.Rows(i)("Per") = FormatPercent(Val(COA.Rows(i)("Per").ToString), , TriState.True)
            'Format(Val(COA.Rows(i)("Per").ToString), "00.##%")
            'COA.Rows(i)("Per") = Format(Val(COA.Rows(i)("Per").ToString), "##.00") + "%"
            Dim style_per As String

            If Left(COA.Rows(i)("Per").ToString, 1) = "-" Then COA.Rows(i)("Per") = "(" & COA.Rows(i)("Per").replace("-", "") & ")"

            'If Request.Form("Round") = "on" Then
            '    COA.Rows(i)("Per") = Format(Val(COA.Rows(i)("Per").ToString), "##")
            'End If

            If COA.Rows(i)("Per").ToString = "0.##%" Or COA.Rows(i)("Per").ToString = "%" Then COA.Rows(i)("Per") = ""

            COA.AcceptChanges()

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 2px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If


            If Show_Per = "on" Then
                style_per = Style2
                If (Left(COA.Rows(i)("Per").ToString, 1) = "(") Then
                    style_per = style_per & "color: red !important;"
                End If
            End If

            If j = 1 Then
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", style_per, "<span style=""" + StyleFinish + """ > " + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If


        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")

        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish1 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish2 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinishTotal As String = "border-bottom: Double 3px black; border-top: 1px solid black;"

        ' Check if rowIncome, rowCost, and rowExpense have value
        If rowIncome.Length > 0 And rowCost.Length > 0 And rowExpense.Length > 0 Then
            ' Calculating Profit/Loss

            If j = 2 Then

                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))

                TotalProfitloss = (((Profitloss1 - Profitloss0) / Profitloss0) * 100).ToString

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                If Show_Per = "on" Then

                    TotalProfitloss = (Math.Round(Val(TotalProfitloss.ToString), 2)).ToString() & "%"
                Else
                    TotalProfitloss = ""
                End If

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + " </span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                Profitloss2 = Convert.ToDecimal(rowIncome(0).Item("Balance2")) - Convert.ToDecimal(rowCost(0).Item("Balance2")) - Convert.ToDecimal(rowExpense(0).Item("Balance2"))
                'TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1) + Convert.ToDecimal(Profitloss2)
                TotalProfitloss = (((Profitloss2 - Profitloss0) / Profitloss0) * 100).ToString

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.##")
                If Show_Per = "on" Then

                    TotalProfitloss = (Math.Round(Val(TotalProfitloss.ToString), 2)).ToString() & "%"
                Else
                    TotalProfitloss = ""
                End If

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(Profitloss2.ToString, 1) = "-" Then
                    Profitloss2 = "(" & Profitloss2.Replace("-", "") & ")"
                    StyleFinish2 = StyleFinish2 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinish2 + """>" + Profitloss2 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """> " + TotalProfitloss + " </span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End If

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        'Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")
            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the total
            Dim excelTotal = exportTable.NewRow()
            excelTotal("value1") = "Profit/Loss"
            For i = 0 To Convert.ToInt32(Goback) - 1
                Profitloss(i) = Convert.ToDecimal(rowIncome(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowCost(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowExpense(0).Item("Balance" + i.ToString))
                excelTotal("value" + (i + 2).ToString) = Profitloss(i)
            Next

            exportTable.Rows.Add(excelTotal)

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Convert.ToInt32(Goback) - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To Convert.ToInt32(Goback) - 1
                excelHeader("value" + (i + 2).ToString) = endDate.AddYears(i - (Goback - 1)).ToString("MMMM yyyy")
                If i = Convert.ToInt32(Goback) - 1 Then
                    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Income Statement Multiperiod Quarterly
    Private Sub PrintQuarterlyIncStateMultiPer()

        Dim Language As Integer = Request.Form("language")
        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim Acc_No As String = Request.Form("Ac")
        Dim j As Integer = 0
        Dim Level As Integer = 1

        Dim Year As String = ""
        Dim Qua_1 As String = ""
        Dim Qua_2 As String = ""
        Dim Qua_3 As String = ""
        Dim Qua_4 As String = ""

        Dim Qua_1_StartDate As String = ""
        Dim Qua_1_EndDate As String = ""
        Dim Qua_2_StartDate As String = ""
        Dim Qua_2_EndDate As String = ""
        Dim Qua_3_StartDate As String = ""
        Dim Qua_3_EndDate As String = ""
        Dim Qua_4_StartDate As String = ""
        Dim Qua_4_EndDate As String = ""

        Dim seconDate, startDate, Asterix As String

        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"

        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""

        Year = Request.Form("YearForQuater")

        Dim FiscalDate, date1 As String

        Dim d1 As Date

        Qua_1 = Request.Item("Q1")
        Qua_2 = Request.Item("Q2")
        Qua_3 = Request.Item("Q3")
        Qua_4 = Request.Item("Q4")
        Dim Q As Integer = 0

        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        Dim StyleMonth As String
        Dim Quarter(4) As String

        Dim COA, Bal1, Bal2, Report, Fiscal As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = (Year - 1) & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        Else
            FiscalDate = (Year - 1) & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        End If

        date1 = FiscalDate

        Qua_1_StartDate = d1
        Qua_1_EndDate = d1.AddMonths(3).AddDays(-1)

        Qua_2_StartDate = d1.AddMonths(3)
        Qua_2_EndDate = d1.AddMonths(6).AddDays(-1)

        Qua_3_StartDate = d1.AddMonths(6)
        Qua_3_EndDate = d1.AddMonths(9).AddDays(-1)

        Qua_4_StartDate = d1.AddMonths(9)
        Qua_4_EndDate = d1.AddMonths(12).AddDays(-1)

        Asterix = ""

        ' Check if the quarter picked is today's quarter
        ' Check the year first then check the month compare to fiscal
        If ((Year = DateTime.Now.Year And DateTime.Now.Month < Fiscal.Rows(0)("Fiscal_Year_Start_Month")) Or (Year = (DateTime.Now.Year - 1) And DateTime.Now.Month >= Fiscal.Rows(0)("Fiscal_Year_Start_Month"))) Then
            ' Check the month to see if it's today's quarter got selected
            ' Need to Mark which quarter is today's
            If (Today() >= d1 And Now() <= d1.AddMonths(3).AddDays(-1) And Qua_1 = "on") Then
                ' It's in Q1
                Qua_1_EndDate = Now().ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(3) And Now() <= d1.AddMonths(6).AddDays(-1) And Qua_2 = "on") Then
                ' It's in Q2
                Qua_2_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(6) And Now() <= d1.AddMonths(9).AddDays(-1) And Qua_3 = "on") Then
                ' It's in Q3
                Qua_3_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_2_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-6).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(9) And Now() <= d1.AddMonths(12).AddDays(-1) And Qua_4 = "on") Then
                ' It's in Q4
                Qua_4_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_3_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Qua_2_EndDate = Now().AddMonths(-6).ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-9).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            End If
        End If

        If (Qua_1 = "on") Then
            If Language = 0 Then
                Quarter(0) = "Q-1"
            ElseIf Language = 1 Then
                Quarter(0) = "T-1"
            End If

            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & Q.ToString
            seconDate = Qua_1_EndDate
            startDate = Qua_1_StartDate
            Q += 1
        End If
        If (Qua_2 = "on") Then
            If Language = 0 Then
                Quarter(1) = "Q-2"
            ElseIf Language = 1 Then
                Quarter(1) = "T-2"
            End If

            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_2_EndDate
            If Q = 0 Then
                startDate = Qua_2_StartDate
            End If
            Q += 1
        End If
        If (Qua_3 = "on") Then
            If Language = 0 Then
                Quarter(2) = "Q-3"
            ElseIf Language = 1 Then
                Quarter(2) = "T-3"
            End If

            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_3_EndDate
            If Q = 0 Then
                startDate = Qua_3_StartDate
            End If
            Q += 1
        End If
        If (Qua_4 = "on") Then

            If Language = 0 Then
                Quarter(3) = "Q-4"
            ElseIf Language = 1 Then
                Quarter(3) = "T-4"
            End If

            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID  AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & Q.ToString
            seconDate = Qua_4_EndDate
            If Q = 0 Then
                startDate = Qua_4_StartDate
            End If
            Q += 1
        End If

        Dim Profitloss(Q) As String
        Dim H_Quarter, H_Qua_1 As String
        Dim HF_Acc As String = ""
        Dim Temp As Integer
        For l = 0 To 3
            If Quarter(l) <> "" Then
                H_Quarter = Quarter(l) + Asterix
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + H_Quarter
                If Temp < (Q - 1) Then
                    If Temp < (Q - 2) Then
                        H_Qua_1 = H_Qua_1 + Quarter(l) + ", "
                    Else
                        H_Qua_1 = H_Qua_1 + Quarter(l)
                    End If
                ElseIf Temp = (Q - 1) Then
                    If Language = 0 Then
                        H_Qua_1 = H_Qua_1 + " and " + Quarter(l)
                    ElseIf Language = 1 Then
                        H_Qua_1 = H_Qua_1 + " y " + Quarter(l)
                    End If

                End If
                Temp += 1
            End If

            'startDate1 = startDate.ToString("yyyy-MM")
        Next
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"

        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Account Description" & StyleMonth & "~Text-align:right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Multiperiod Income Statement(Quarterly) " + DenomString + "<br/>For " & H_Qua_1 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; width:50 px; font-size:8pt~Descripción De Cuenta" & StyleMonth & "~Text-align:right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Trimestral) " + DenomString + "<br/>Para " & H_Qua_1 & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        ' Translate the query
        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))

        Dim Balance As String
        Dim BalanceString As String = ""
        Dim ColMonth As String = ""

        j = 0
        Dim k As Int32 = 0

        Balance = ""
        BalanceString = ""
        For col = 0 To Q - 1
            'While (startDate1 <= seconDate)
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next


            Next

            If Q < 4 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(BalanceString).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If


            COA.AcceptChanges()
            j += 1
            'startDate = startDate.AddMonths(1)
            'startDate1 = startDate.ToString("yyyy-MM")
        Next
        'End While
        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next

        COA.AcceptChanges()
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        'startDate11 = firstDate
        'startDate2 = firstDate
        Dim Bal0 As Decimal
        Dim Bal11 As Decimal
        Dim Bal22 As Decimal

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
        For i = 0 To COA.Rows.Count - 1

            ' Calculation for Total
            If Q >= 1 Then
                If COA.Rows(i)("Balance0").ToString = "" Then
                    Bal0 = 0
                Else
                    Bal0 = COA.Rows(i)("Balance0")
                End If
                If Q >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal11 = 0
                    Else
                        Bal11 = COA.Rows(i)("Balance1")
                    End If
                    If Q = 3 Then
                        If COA.Rows(i)("Balance2").ToString = "" Then
                            Bal22 = 0
                        Else
                            Bal22 = COA.Rows(i)("Balance2")
                        End If
                    End If

                End If

            End If

            COA.Rows(i)("Total") = (Bal0 + Bal11 + Bal22).ToString
            Bal0 = 0
            Bal11 = 0
            Bal22 = 0
            COA.AcceptChanges()
            ' Format all the output for the paper

            If Request.Form("Round") = "on" Then
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###")
            Else
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            End If

            If Left(COA.Rows(i)("Total").ToString, 1) = "-" Then COA.Rows(i)("Total") = "(" & COA.Rows(i)("Total").replace("-", "") & ")"



            If COA.Rows(i)("Total").ToString = "$.00" Or COA.Rows(i)("Total").ToString = "$" Then COA.Rows(i)("Total") = ""

            COA.AcceptChanges()
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width:7px; max-width: 7px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"
            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            If Q = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")

        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish1 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish2 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinishTotal As String = "border-bottom: Double 3px black; border-top: 1px solid black;"

        ' Check if rowIncome, rowCost, and rowExpense have value
        If rowIncome.Length > 0 And rowCost.Length > 0 And rowExpense.Length > 0 Then
            ' Calculating Profit/Loss
            If j = 1 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                Profitloss2 = Convert.ToDecimal(rowIncome(0).Item("Balance2")) - Convert.ToDecimal(rowCost(0).Item("Balance2")) - Convert.ToDecimal(rowExpense(0).Item("Balance2"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1) + Convert.ToDecimal(Profitloss2)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(Profitloss2.ToString, 1) = "-" Then
                    Profitloss2 = "(" & Profitloss2.Replace("-", "") & ")"
                    StyleFinish2 = StyleFinish2 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinish2 + """>" + Profitloss2 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End If

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")

            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the total
            Dim excelTotal = exportTable.NewRow()
            excelTotal("value1") = "Profit/Loss"
            For i = 0 To Q - 1
                Profitloss(i) = Convert.ToDecimal(rowIncome(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowCost(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowExpense(0).Item("Balance" + i.ToString))
                excelTotal("value" + (i + 2).ToString) = Profitloss(i)
            Next

            exportTable.Rows.Add(excelTotal)

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Q - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            Dim count As Integer = 0
            For i = 0 To Q
                While (Quarter(i + count) = "" And (i + count) < Quarter.Length - 1)
                    count += 1
                End While
                If Quarter(i + count) <> "" Then
                    excelHeader("value" + (i + 2).ToString()) = Quarter(i + count) + Asterix
                End If
                'If i = Q - 1 Then
                '    excelHeader("value" + (i + 3).ToString()) = "Total"
                'End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Income Statement Multiperiod Quarter-to-Quarter
    Private Sub PrintQuarterToQuarterIncStateMultiPer()

        Dim Language As Integer = Request.Form("language")
        Dim seconDate As String = Request.Form("Quarter")
        Dim words As String() = seconDate.Split(New Char() {" "c})
        Dim Qua_No As Integer = Integer.Parse(words(0))
        Dim Qua_Year As String = words(1).ToString
        Dim Goback As String = Request.Form("goback")
        Dim Show_Per As String = Request.Form("Percentage")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")

        Dim Quarter(4) As String

        Dim Fiscal As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        Dim FiscalDate, Qdate1, Qdate2 As String

        Dim d1 As Date

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = Qua_Year & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        Else
            FiscalDate = Qua_Year & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        End If

        'date1 = FiscalDate

        If Qua_No = 1 Then
            Qdate1 = d1
            Qdate2 = d1.AddMonths(3).AddDays(-1)
        ElseIf Qua_No = 2 Then
            Qdate1 = d1.AddMonths(3)
            Qdate2 = d1.AddMonths(6).AddDays(-1)
        ElseIf Qua_No = 3 Then
            Qdate1 = d1.AddMonths(6)
            Qdate2 = d1.AddMonths(9).AddDays(-1)
        ElseIf Qua_No = 4 Then
            Qdate1 = d1.AddMonths(9)
            Qdate2 = d1.AddMonths(12).AddDays(-1)
        End If
        Dim S_Date As Date = Qdate1
        Dim E_Date As Date = Qdate2
        ' Check if the quarter picked is today's quarter
        ' Check the year first then check the month compare to fiscal+
        If S_Date < Now() And E_Date > Today() Then
            'If (((Qua_Year + 1) = DateTime.Now.Year And DateTime.Now.Month < Fiscal.Rows(0)("Fiscal_Year_Start_Month")) Or ((Qua_Year + 1) = (DateTime.Now.Year - 1) And DateTime.Now.Month >= Fiscal.Rows(0)("Fiscal_Year_Start_Month"))) Then
            ' Check the month to see if it's today's quarter got selected
            ' Need to Mark which quarter is today's
            If (Today() >= d1 And Now() <= d1.AddMonths(3).AddDays(-1)) Then
                ' It's in Q1
                Qdate1 = d1
                Qdate2 = Now().ToString("yyyy-MM-dd")
            ElseIf (Today() >= d1.AddMonths(3) And Now() <= d1.AddMonths(6).AddDays(-1)) Then
                ' It's in Q2
                Qdate1 = d1.AddMonths(3)
                Qdate2 = Now().ToString("yyyy-MM-dd")

            ElseIf (Today() >= d1.AddMonths(6) And Now() <= d1.AddMonths(9).AddDays(-1)) Then
                ' It's in Q3
                Qdate1 = d1.AddMonths(6)
                Qdate2 = Now().ToString("yyyy-MM-dd")

            ElseIf (Today() >= d1.AddMonths(9) And Now() <= d1.AddMonths(12).AddDays(-1)) Then
                ' It's in Q4
                Qdate1 = d1.AddMonths(9)
                Qdate2 = Now().ToString("yyyy-MM-dd")

            End If
        End If
        S_Date = Qdate1
        E_Date = Qdate2

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim j As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As Date = S_Date.AddYears(-(Goback - 1))
        Dim secondDate As Date = E_Date.AddYears(-(Goback - 1))

        Dim startDate2, endDate2, Date1, Date2 As String

        Dim StyleFinish As String = ""

        Dim DenomString As String = ""
        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""
        Dim Profitloss(Goback - 1) As String

        Dim Bal0, Bal1, Bal2 As Decimal

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Default date give today's date and a year before
        'If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yy-MM")

        Dim StyleMonth As String
        Dim HF_Acc As String = ""

        'Dim period As Integer = Integer.Parse(Goback)

        'Translation for Header

        For j = 0 To Goback - 1
            If Language = 0 Then
                If S_Date < Now() And E_Date >= Today() Then
                    Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2)) & "(*)"
                Else
                    Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                End If

                If j = (Goback - 1) Then
                    Date2 = Date2 & " and Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                ElseIf j = (Goback - 2) Then
                    Date2 = Date2 & "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                Else
                    Date2 = Date2 & "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2)).ToString() + ", "
                End If
            End If
            If Language = 1 Then
                If S_Date < Now() And E_Date >= Today() Then
                    Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2)) & "(*)"
                Else
                    Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                End If

                If j = (Goback - 1) Then
                    Date2 = Date2 & " y T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                ElseIf j = (Goback - 2) Then
                    Date2 = Date2 & "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2))
                Else
                    Date2 = Date2 & "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(Goback) - 2)).ToString() + ", "
                End If
            End If

            StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + Date1

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd")
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
            firstDate = startDate2
            secondDate = endDate2

        Next

        Dim Per_opt As String = ""
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If
        If Show_Per = "on" Then
            If Language = 0 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Percentage(%)"
            ElseIf Language = 1 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Porcentaje(%)"
            End If

        Else
            Per_opt = "~width:0px;font-size:0pt~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; font-size:8pt~Account Description" & StyleMonth & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Income Statement(Quarter To Quarter) " + DenomString + "<br/>For " & Date2 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Descripción De Cuenta" & StyleMonth & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Mensual) " + DenomString + "<br/>Para " & Date2 & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand


        firstDate = S_Date.AddYears(-(Goback - 1))
        secondDate = E_Date.AddYears(-(Goback - 1))
        startDate2 = firstDate
        endDate2 = secondDate

        For j = 0 To Goback - 1

            Query1 = Query1 + ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + j.ToString()
            Query2 = Query2 + ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate2 & "' and '" & endDate2 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + j.ToString()

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd")
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
            firstDate = startDate2
            secondDate = endDate2
        Next
        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))
        COA.Columns.Add("Per", GetType(String))

        Dim Balance As String = ""
        Dim BalanceString As String = ""

        ' For loop for Calculation and  Formatting
        For j = 0 To Goback - 1

            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            ' Format all the output for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If

            COA.AcceptChanges()
        Next
        ' End of for loop

        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()
            End If
        Next

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        ' Calculation for Per
        For i = 0 To COA.Rows.Count - 1

            If COA.Rows(i)("Balance0").ToString = "" Then
                Bal0 = 0
            ElseIf COA.Rows(i)("Balance0") <> 0 Then
                Bal0 = COA.Rows(i)("Balance0")
                If j >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal1 = 0
                    Else
                        Bal1 = COA.Rows(i)("Balance1")
                        COA.Rows(i)("Per") = (((Bal1 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                        If j = 3 Then
                            If COA.Rows(i)("Balance2").ToString = "" Then
                                Bal2 = 0
                            Else
                                Bal2 = COA.Rows(i)("Balance2")
                                COA.Rows(i)("Per") = (((Bal2 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                            End If

                        End If
                    End If
                End If
            End If


            COA.AcceptChanges()
            ' Format all the output for the paper  
            If COA.Rows(i)("Per").ToString <> "" Then
                COA.Rows(i)("Per") = (Math.Round(Val(COA.Rows(i)("Per").ToString), 2)).ToString() & "%"
            End If

            If Left(COA.Rows(i)("Per").ToString, 1) = "-" Then COA.Rows(i)("Per") = "(" & COA.Rows(i)("Per").replace("-", "") & ")"

            'If Request.Form("Round") = "on" Then
            '    COA.Rows(i)("Per") = Format(Val(COA.Rows(i)("Per").ToString), "##")
            'End If

            If COA.Rows(i)("Per").ToString = "0.##%" Or COA.Rows(i)("Per").ToString = "%" Then COA.Rows(i)("Per") = ""

            COA.AcceptChanges()

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 2px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If
            Dim style_per As String



            If Show_Per = "on" Then
                style_per = Style2

                If (Left(COA.Rows(i)("Per").ToString, 1) = "(") Then
                    style_per = style_per & "color: red !important;"

                End If
            End If

            If j = 1 Then
                'Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", style_per, "<span style=""" + StyleFinish + """ > " + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")

        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish1 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish2 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinishTotal As String = "border-bottom: Double 3px black; border-top: 1px solid black;"

        ' Check if rowIncome, rowCost, and rowExpense have value
        If rowIncome.Length > 0 And rowCost.Length > 0 And rowExpense.Length > 0 Then
            ' Calculating Profit/Loss

            If j = 2 Then

                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))

                'Profitloss0 = (Convert.ToDecimal(rowIncome(0).Item("Balance0")))
                'Profitloss1 = (Convert.ToDecimal(rowIncome(0).Item("Balance1")))
                TotalProfitloss = (((Profitloss1 - Profitloss0) / Profitloss0) * 100).ToString
                'TotalProfitloss = ((((Convert.ToDecimal(rowIncome(0).Item("Balance1"))) - (Convert.ToDecimal(rowIncome(0).Item("Balance0")))) / (Convert.ToDecimal(rowIncome(0).Item("Balance0")))) * 100).ToString

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                If Show_Per = "on" Then

                    TotalProfitloss = (Math.Round(Val(TotalProfitloss.ToString), 2)).ToString() & "%"
                Else
                    TotalProfitloss = ""
                End If


                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + " </span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                Profitloss2 = Convert.ToDecimal(rowIncome(0).Item("Balance2")) - Convert.ToDecimal(rowCost(0).Item("Balance2")) - Convert.ToDecimal(rowExpense(0).Item("Balance2"))
                'TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1) + Convert.ToDecimal(Profitloss2)
                TotalProfitloss = (((Profitloss2 - Profitloss0) / Profitloss0) * 100).ToString

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.##")
                If Show_Per = "on" Then

                    TotalProfitloss = (Math.Round(Val(TotalProfitloss.ToString), 2)).ToString() & "%"
                Else
                    TotalProfitloss = ""
                End If

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(Profitloss2.ToString, 1) = "-" Then
                    Profitloss2 = "(" & Profitloss2.Replace("-", "") & ")"
                    StyleFinish2 = StyleFinish2 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinish2 + """>" + Profitloss2 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """> " + TotalProfitloss + " </span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End If

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")
            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the total
            Dim excelTotal = exportTable.NewRow()
            excelTotal("value1") = "Profit/Loss"
            For i = 0 To Convert.ToInt32(Goback) - 1
                Profitloss(i) = Convert.ToDecimal(rowIncome(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowCost(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowExpense(0).Item("Balance" + i.ToString))
                excelTotal("value" + (i + 2).ToString) = Profitloss(i)
            Next

            exportTable.Rows.Add(excelTotal)

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Convert.ToInt32(Goback) - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To Convert.ToInt32(Goback) - 1
                If Language = 0 Then
                    If S_Date < Now() And E_Date >= Today() Then
                        Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + i) - (Convert.ToInt32(Goback) - 2)) & "(*)"
                    Else
                        Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + i) - (Convert.ToInt32(Goback) - 2))
                    End If
                End If
                If Language = 1 Then
                    If S_Date < Now() And E_Date >= Today() Then
                        Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + i) - (Convert.ToInt32(Goback) - 2)) & "(*)"
                    Else
                        Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + i) - (Convert.ToInt32(Goback) - 2))
                    End If
                End If

                excelHeader("value" + (i + 2).ToString) = Date1
                If i = Convert.ToInt32(Goback) - 1 Then
                    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If

    End Sub

    ' Income Statement Multiperiod Yearly
    Private Sub PrintYearlyIncStateMultiPer()

        Dim Language As Integer = Request.Form("language")
        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim j As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As String
        Dim seconDate As String
        Dim startDate As Date
        Dim startDate1 As String
        Dim endDate1 As Date
        Dim endDate As String
        Dim endDate2 As String
        Dim Acc_No As String = Request.Form("Ac")
        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"

        firstDate = Request.Form("FirstDate")
        seconDate = Request.Form("SecondDate")
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""


        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Default date give today's date and a year before
        If firstDate = "" Then firstDate = Now().ToString("yy")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yy")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try

        Dim COA, Bal, Bal1, Bal2, Report, Fiscal As New DataTable

        Dim FiscalDate, FiscalDateEnd As String
        Dim date1, date2 As String
        Dim d1, d2, d3, dtemp As Date
        Dim YearCount As Integer = seconDate - (firstDate - 1)
        Dim HF_Acc As String = ""
        Dim Asterix As String = ""
        ' Get the fiscal month
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        ' Because it is '9' not '09'
        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = (firstDate - 1) & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
            d1 = FiscalDate
            FiscalDateEnd = d1.AddDays(-1).AddYears(1).ToString("yyyy-MM-dd")
            d2 = FiscalDateEnd
            date2 = seconDate & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
            date2 = seconDate & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
            d3 = date2
            d3 = d3.AddDays(-1).ToString("yyyy-MM-dd")
            date2 = d3
        Else
            FiscalDate = (firstDate - 1) & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
            d1 = FiscalDate
            FiscalDateEnd = d1.AddDays(-1).AddYears(1).ToString("yyyy-MM-dd")
            d2 = FiscalDateEnd
            date2 = seconDate & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
            d3 = date2
            d3 = d3.AddDays(-1).ToString("yyyy-MM-dd")
            date2 = d3
        End If

        date1 = FiscalDate

        Dim seconDate1 = seconDate

        If seconDate = Now().ToString("yyyy") Then
            FiscalDateEnd = Now().AddYears(-(YearCount - 1)).ToString("yyyy-MM-dd")
            date2 = Now().ToString("yyyy-MM-dd")
            'Else
            '    endDate1 = endDate.AddYears(-(Goback - 1)).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
        End If
        dtemp = date2
        startDate1 = FiscalDate
        startDate = startDate1
        endDate = FiscalDateEnd
        endDate1 = endDate
        seconDate = date2

        endDate2 = endDate1.Year
        Dim StyleMonth, H_year As String

        While (endDate <= seconDate)
            If dtemp.Year >= DateTime.Now.Year Then
                endDate2 = endDate1.ToString("yyyy") + "(*)"
                Asterix = "(*)"
            Else
                endDate2 = endDate1.ToString("yyyy")
            End If
            If endDate1 = dtemp Then
                If Language = 0 Then
                    H_year = H_year + " and " & startDate.ToString("yyyy") + "-" + endDate1.ToString("yyyy")
                ElseIf Language = 1 Then
                    H_year = H_year + " y " & startDate.ToString("yyyy") + "-" + endDate1.ToString("yyyy")
                End If

            ElseIf endDate1 = (dtemp.AddYears(-1)) Then
                H_year = H_year & startDate.ToString("yyyy") + "-" + endDate1.ToString("yyyy")
            Else
                H_year = H_year & startDate.ToString("yyyy") + "-" + endDate1.ToString("yyyy") & ", "
            End If
            StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + startDate.ToString("yyyy") + "-" + endDate2
            startDate = startDate.AddYears(1)
            startDate1 = startDate
            endDate1 = endDate1.AddYears(1)
            endDate = endDate1.ToString("yyyy/MM/dd")

        End While
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Account Description" & StyleMonth & "~Text-align:right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Multiperiod Income Statement(Yearly) " + DenomString + "<br/>For " & H_year & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Descripción De Cuenta" & StyleMonth & "~Text-align:right; width:120px; font-size:8pt~Total"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Anual) " + DenomString + "<br/>Para " & H_year & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        startDate1 = FiscalDate
        startDate = startDate1
        endDate = FiscalDateEnd
        endDate1 = endDate
        Dim Q As Integer = 0

        While (endDate <= seconDate)
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & j.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" & j.ToString
            j += 1
            Q += 1
            startDate = startDate.AddYears(1).ToString("yyyy/MM/dd")
            startDate1 = startDate
            endDate1 = endDate1.AddYears(1).ToString("yyyy/MM/dd")
            endDate = endDate1
        End While

        ' Translation Query
        If Language = 0 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            ' Getting Total Sales and Other Income (49999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

            ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  '0' and Account_ID > 1 and Account_No >= '50000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))

        startDate1 = FiscalDate
        startDate = startDate1
        endDate = FiscalDateEnd
        endDate1 = endDate
        j = 0
        Dim Balance As String
        Dim BalanceString As String = ""
        Dim ColMonth As String = ""


        Dim k As Int32 = 0
        Balance = ""
        BalanceString = ""
        While (endDate1 <= seconDate)
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If
                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            ' Formatting for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If

            COA.AcceptChanges()

            j += 1
            startDate = startDate.AddYears(1).ToString("yyyy/MM/dd")
            startDate1 = startDate
            endDate1 = endDate1.AddYears(1).ToString("yyyy/MM/dd")
            endDate = endDate1
        End While

        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next

        COA.AcceptChanges()
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        'startDate11 = firstDate
        'startDate2 = firstDate + "-01-01"
        Dim Bal0 As Decimal
        Dim Bal11 As Decimal
        Dim Bal22 As Decimal

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
        For i = 0 To COA.Rows.Count - 1

            ' Calculation for Total
            If Q >= 1 Then
                If COA.Rows(i)("Balance0").ToString = "" Then
                    Bal0 = 0
                Else
                    Bal0 = COA.Rows(i)("Balance0")
                End If
                If Q >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal11 = 0
                    Else
                        Bal11 = COA.Rows(i)("Balance1")
                    End If
                    If Q = 3 Then
                        If COA.Rows(i)("Balance2").ToString = "" Then
                            Bal22 = 0
                        Else
                            Bal22 = COA.Rows(i)("Balance2")
                        End If
                    End If

                End If

            End If

            COA.Rows(i)("Total") = (Bal0 + Bal11 + Bal22).ToString
            Bal0 = 0
            Bal11 = 0
            Bal22 = 0
            COA.AcceptChanges()
            ' Format all the output for the paper

            'COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            If Request.Form("Round") = "on" Then
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###")
            Else
                COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            End If

            If Left(COA.Rows(i)("Total").ToString, 1) = "-" Then COA.Rows(i)("Total") = "(" & COA.Rows(i)("Total").replace("-", "") & ")"



            If COA.Rows(i)("Total").ToString = "$.00" Or COA.Rows(i)("Total").ToString = "$" Then COA.Rows(i)("Total") = ""

            COA.AcceptChanges()
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width:7px; max-width: 7px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If
            If j = 1 Then

                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        ' Get the value for Total Income, Total Cost, and Total Expenses
        Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
        Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
        Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")

        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish1 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinish2 As String = "border-bottom: Double 3px black; border-top: 1px solid black;"
        Dim StyleFinishTotal As String = "border-bottom: Double 3px black; border-top: 1px solid black;"

        ' Check if rowIncome, rowCost, and rowExpense have value
        If rowIncome.Length > 0 And rowCost.Length > 0 And rowExpense.Length > 0 Then
            ' Calculating Profit/Loss
            If j = 1 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Profitloss0 = Convert.ToDecimal(rowIncome(0).Item("Balance0")) - Convert.ToDecimal(rowCost(0).Item("Balance0")) - Convert.ToDecimal(rowExpense(0).Item("Balance0"))
                Profitloss1 = Convert.ToDecimal(rowIncome(0).Item("Balance1")) - Convert.ToDecimal(rowCost(0).Item("Balance1")) - Convert.ToDecimal(rowExpense(0).Item("Balance1"))
                Profitloss2 = Convert.ToDecimal(rowIncome(0).Item("Balance2")) - Convert.ToDecimal(rowCost(0).Item("Balance2")) - Convert.ToDecimal(rowExpense(0).Item("Balance2"))
                TotalProfitloss = Convert.ToDecimal(Profitloss0) + Convert.ToDecimal(Profitloss1) + Convert.ToDecimal(Profitloss2)

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.##")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.##")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.##")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.##")

                ' Check ProfitAndLoss Value negative or positive
                If Left(Profitloss0.ToString, 1) = "-" Then
                    Profitloss0 = "(" & Profitloss0.Replace("-", "") & ")"
                    StyleFinish = StyleFinish & "color: red !important;"
                End If
                If Left(Profitloss1.ToString, 1) = "-" Then
                    Profitloss1 = "(" & Profitloss1.Replace("-", "") & ")"
                    StyleFinish1 = StyleFinish1 & "color: red !important;"
                End If
                If Left(Profitloss2.ToString, 1) = "-" Then
                    Profitloss2 = "(" & Profitloss2.Replace("-", "") & ")"
                    StyleFinish2 = StyleFinish2 & "color: red !important;"
                End If
                If Left(TotalProfitloss.ToString, 1) = "-" Then
                    TotalProfitloss = "(" & TotalProfitloss.Replace("-", "") & ")"
                    StyleFinishTotal = StyleFinishTotal & "color: red !important;"
                End If

                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", "", Style, "PROFIT/LOSS", Style2, "<span style=""" + StyleFinish + """>" + Profitloss0 + "</span>", Style2, "<span style=""" + StyleFinish1 + """>" + Profitloss1 + "</span>", Style2, "<span style=""" + StyleFinish2 + """>" + Profitloss2 + "</span>", Style2, "<span style=""" + StyleFinishTotal + """>" + TotalProfitloss + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        End If
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            Dim Profitloss(Q) As String

            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")

            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Add the total
            Dim excelTotal = exportTable.NewRow()
            excelTotal("value1") = "Profit/Loss"
            For i = 0 To Q - 1
                Profitloss(i) = Convert.ToDecimal(rowIncome(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowCost(0).Item("Balance" + i.ToString)) - Convert.ToDecimal(rowExpense(0).Item("Balance" + i.ToString))
                excelTotal("value" + (i + 2).ToString) = Profitloss(i)
            Next

            exportTable.Rows.Add(excelTotal)

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To YearCount - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To YearCount - 1
                excelHeader("value" + (i + 2).ToString) = (firstDate - 1 + i) & "-" & (firstDate + i) & Asterix
                'If i = YearCount - 1 Then
                '    excelHeader("value" + (i + 3).ToString) = "Total"
                'End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If

    End Sub

    ' Balance Sheet Multiperiod Monthly
    Private Sub PrintMonthlyBalSheetMultiPer()

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim StyleFinish As String
        Dim Language As Integer = Request.Form("language")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim NoZeros As String = Request.Form("showZeros")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DenomString As String = ""

        Dim j As Integer = 0
        Dim k As Integer = 0

        Dim startDate, endDate, firstDate1, seconDate1 As Date
        Dim startDate1, Asterix As String
        Dim firstDate As String = Request.Form("FirstDate")
        Dim seconDate As String = Request.Form("SecondDate")

        Dim Padding As Integer = 0
        Dim Level As Integer = 1

        Dim MonthCount As Integer

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        ' Get the MonthCount Value
        Try
            startDate = firstDate
            endDate = seconDate
            While startDate <> endDate
                startDate = startDate.AddMonths(1)
                MonthCount += 1
            End While
        Catch ex As Exception
            MonthCount = 0
        End Try

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        Asterix = ""

        ' If seconDate Year and Month is Today's Year and Month then change the date to today
        If Request.Form("SecondDate") = Now().ToString("yyyy-MM") Then
            seconDate1 = Now()
            seconDate = Now().ToString("yyyy-MM-dd")
            firstDate = Now().AddMonths(-MonthCount).ToString("yyyy-MM-dd")
            firstDate1 = firstDate
            Asterix = "(*)"
        Else
            ' Default date give today's date
            If firstDate = "" Then
                firstDate = Now().ToString("yyyy-MM-dd")
                Asterix = "(*)"
            Else
                ' If exist, take the last day
                firstDate1 = firstDate
                firstDate = firstDate1.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
            End If
            If seconDate = "" Then
                seconDate = Now().ToString("yyyy-MM-dd")
                Asterix = "(*)"
            Else
                seconDate1 = seconDate
                seconDate = seconDate1.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
            End If
        End If
        Dim Fr_Date As DateTime
        startDate1 = firstDate
        startDate = firstDate
        Dim StyleMonth As String
        Dim HF_Acc As String = ""

        ' Translate the month Header
        If Language = 0 Then
            While (startDate <= seconDate)
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + startDate.ToString("MMMM") + Asterix
                startDate = startDate.AddMonths(1)
                startDate1 = startDate.ToString("yyyy-MM-dd")
            End While

        ElseIf Language = 1 Then
            While (startDate <= seconDate)
                Fr_Date = DateTime.Parse(startDate)
                Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + StrConv(Fr_Date.ToString("MMMM"), VbStrConv.ProperCase) + Asterix
                startDate = startDate.AddMonths(1)
                startDate1 = startDate.ToString("yyyy-MM-dd")
            End While
        End If

        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Account Name" + StyleMonth + "~Text-align:right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>MultiPeriod Balance Sheet(Monthly) " + DenomString + "<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"

        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + "~Text-align:right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Hoja de Balance Multi Período (Mensual) " + DenomString + "<br/>De " & firstDate & " a " & seconDate & "<br/></span><span style=""font-size:7pt"">Impreso el " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If


        Dim COA, DataT, Bal, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()
        'startDate1 = firstDate
        startDate = firstDate
        startDate1 = startDate.ToString("yyyy-MM-dd")

        While (startDate1 <= seconDate)
            Query1 = Query1 & ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & startDate1 & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" & j.ToString
            Query2 = Query2 & ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & startDate1 & "' order by Transaction_Date desc, rowID desc) as Balance" & j.ToString
            j += 1

            ' Check if the SecondDate is Today's date or not
            If Request.Form("SecondDate") = Now().ToString("yyyy-MM") Then
                startDate = firstDate1.AddMonths(j).ToString("yyyy-MM-dd")
            Else
                startDate = firstDate1.AddMonths(j + 1).AddDays(-1).ToString("yyyy-MM-dd")
            End If

            startDate1 = startDate.ToString("yyyy-MM-dd")
        End While

        ' Translation
        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        'SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
        'SQLCommand.Parameters.Clear()
        'DataAdapter.Fill(COA)

        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID" & Query2 & " From ACC_GL As gl1"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Bal)

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, Exchange_Account_ID, Associated_Totalling From ACC_GL_Accounts order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(DataT)

        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))

        startDate1 = firstDate
        startDate = firstDate
        k = 0
        Dim Balance As String
        Dim BalanceString As String
        Dim ColMonth As String = ""

        Dim RE As Decimal = 0
        'Dim k As Int32 = 0
        Balance = ""
        BalanceString = ""
        While (startDate <= seconDate)
            Balance = "Balance" + k.ToString
            BalanceString = "BalanceString" + k.ToString
            RE = 0
            For j = 0 To DataT.Rows.Count - 1
                For jj = 0 To Bal.Rows.Count - 1
                    If DataT.Rows(j)("Account_ID").ToString = Bal.Rows(jj)("Account_ID").ToString Then
                        If DataT.Rows(j)("Account_Type").ToString = "4" And Not IsDBNull(Bal.Rows(jj)(Balance)) Then RE = RE + Bal.Rows(jj)(Balance)
                        If (DataT.Rows(j)("Account_Type").ToString = "5" Or DataT.Rows(j)("Account_Type").ToString = "6") And Not IsDBNull(Bal.Rows(jj)(Balance)) Then RE = RE - Bal.Rows(jj)(Balance)
                        Exit For
                    End If
                Next
            Next

            For i = 0 To COA.Rows.Count - 1
                For ii = 0 To Bal.Rows.Count - 1
                    ' Copying the Balance value from table Bal to table COA
                    If COA.Rows(i)("Account_ID").ToString = Bal.Rows(ii)("Account_ID").ToString Then
                        COA.Rows(i)(Balance) = Bal.Rows(ii)(Balance)
                        Exit For
                    End If
                Next

                ' Give Padding
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 10 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 10 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
                If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "<div style='min-width: 0.5in; max-width:0.5in;'></div>" ' hard coded
            Next

            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Account_No") = "39000" Then COA.Rows(i)(Balance) = RE
                COA.AcceptChanges()
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Totalling Total Equity (ACC_NO 39999)

            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If

                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            Total = 0
            Account = ""
            For j = 1 To 2
                For i = 0 To COA.Rows.Count - 1
                    If COA.Rows(i)("Totalling").ToString <> "" Then
                        Total = 0
                        Account = COA.Rows(i)("Account_No").ToString
                        Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                        For ii = 0 To Plus.Length - 1
                            Dim Dash() As String = Plus(ii).Split("-")
                            Dim Start As String = Trim(Dash(0))
                            Dim Endd As String
                            If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                            For iii = 0 To COA.Rows.Count - 1
                                If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            Next
                        Next
                    End If

                Next
            Next

            COA.AcceptChanges()

            ' Formating
            For i = 0 To COA.Rows.Count - 1
                ' Denomination and rounding
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                End If
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                    COA.Rows(i)(Balance) = denominatedValue
                End If
            Next

            If k < 3 Then
                For i = 0 To COA.Rows.Count - 1
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                Next
            End If

            COA.AcceptChanges()
            k += 1
            startDate = startDate.AddMonths(1)
            startDate1 = startDate.ToString("yyyy-MM-dd")
        End While

        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If k = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf k = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf k = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2.5in; max-width: 2.5in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 9px; max-width: 9px;"
            StyleFinish = ""
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 1in; max-width: 1in;"
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 4px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"

            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            If j = 1 Then

                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)

                If COA.Rows(i)(0).ToString = "39999" Then
                    Exit For
                End If
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To MonthCount
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next
            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To MonthCount
                excelHeader("value" + (i + 2).ToString) = firstDate1.AddMonths(i).ToString("MMMM") + Asterix
                'If i = MonthCount Then
                '    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                'End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Balance Sheet Multiperiod Month-to-Month
    Private Sub PrintMonthToMonthBalSheetMultiPer()
        Dim Language As Integer = Request.Form("language")
        Dim seconDate As String = Request.Form("SecMonth")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim Show_Per As String = Request.Form("Percentage")
        Dim GoPrevious As String = Request.Form("prev")
        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        Dim i As Integer = 0
        Dim HF_Acc As String = ""
        Dim StyleFinish As String = ""
        Dim DenomString As String = ""
        Dim RE As Decimal = 0
        Dim j As Integer = 0
        Dim endDate1 As String
        Dim endDate2 As String
        Dim Date1 As String
        Dim Date2 As String
        Dim Balance As String
        Dim BalanceString As String = ""
        Dim endDate As Date = seconDate.ToString
        Dim secondDate As Date
        Dim StyleMonth As String
        Dim Bal0 As Decimal
        Dim Bal1 As Decimal
        Dim Bal2 As Decimal

        Dim Fiscal As New DataTable

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If endDate.ToString("yyyy-MM") = Now().ToString("yyyy-MM") Then
            endDate1 = Now().AddYears(-(GoPrevious - 1)).ToString("yyyy-MM-dd")
        Else
            endDate1 = endDate.AddYears(-(GoPrevious - 1)).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
        End If

        Date1 = endDate.AddYears(-(GoPrevious - 1)).ToString("dd MMMM yyyy")

        secondDate = endDate1

        Dim Fr_Date As DateTime

        If Language = 0 Then
            For j = 0 To GoPrevious - 1

                If endDate = Now().ToString("yyyy-MM") Then
                    Date1 = secondDate.ToString("MMMM yyyy") + "(*)"
                Else
                    Date1 = secondDate.ToString("MMMM yyyy")
                End If

                If j = (GoPrevious - 1) Then
                    Date2 = Date2 + " and " + secondDate.ToString("dd MMMM yyyy")
                ElseIf j = (GoPrevious - 2) Then
                    Date2 = Date2 + "" + secondDate.ToString("dd MMMM yyyy")
                Else
                    Date2 = Date2 + "" + secondDate.ToString("dd MMMM yyyy") + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + Date1
                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                secondDate = endDate2
            Next
        ElseIf Language = 1 Then
            For j = 0 To GoPrevious - 1
                Fr_Date = DateTime.Parse(secondDate)
                Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")

                If endDate = Now().ToString("yyyy-MM") Then
                    Date1 = Fr_Date.ToString("MMMM yyyy") + "(*)"
                Else
                    Date1 = Fr_Date.ToString("MMMM yyyy")
                End If

                If j = (GoPrevious - 1) Then
                    Date2 = Date2 + " y " + Fr_Date.ToString("dd MMMM yyyy")
                ElseIf j = (GoPrevious - 2) Then
                    Date2 = Date2 + "" + Fr_Date.ToString("dd MMMM yyyy")
                Else
                    Date2 = Date2 + "" + Fr_Date.ToString("dd MMMM yyyy") + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + StrConv(Fr_Date.ToString("MMMM yyyy"), VbStrConv.ProperCase)
                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                secondDate = endDate2
            Next
        End If

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        Dim Per_opt As String = ""
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If
        If Show_Per = "on" Then
            If Language = 0 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Percentage(%)"
            ElseIf Language = 1 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Porcentaje(%)"
            End If
        Else
            Per_opt = "~width:0px;font-size:0pt~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; font-size:8pt~Account Name" & StyleMonth & "" & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Balance Sheet(Month To Month) " + DenomString + "<br/>For " & Date2 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Descripción De Cuenta" & StyleMonth & "" & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Mensual) " + DenomString + "<br/>Para " & Date2 & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        Dim COA, Bal, Report, DataT As New DataTable

        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()
        ' Getting the Query

        secondDate = endDate1
        endDate2 = endDate1

        For j = 0 To GoPrevious - 1
            Query1 = Query1 + ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & endDate2 & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" + j.ToString()
            Query2 = Query2 + ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & endDate2 & "' order by Transaction_Date desc, rowID desc) as Balance" + j.ToString()

            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
            secondDate = endDate2
        Next

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts  WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID" & Query2 & " from ACC_GL gl1"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Bal)

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling Active, Cash, Exchange_Account_ID, Associated_Totalling From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(DataT)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))
        COA.Columns.Add("Per", GetType(String))

        Balance = ""
        BalanceString = ""

        For j = 0 To GoPrevious - 1
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    ' If Denom > 1 Then
                    '     Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                    '     COA.Rows(i)(Balance) = denominatedValueCurrent
                    ' End If

                Next
            End If

            ' Give Padding
            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            RE = 0

            ' Calculation for Current Retained Earning (39000)
            For i = 0 To DataT.Rows.Count - 1

                For jj = 0 To Bal.Rows.Count - 1
                    If DataT.Rows(i)("Account_ID").ToString = Bal.Rows(jj)("Account_ID").ToString Then

                        If DataT.Rows(i)("Account_Type").ToString = "4" Then
                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE + Bal.Rows(jj)(Balance)
                            End If
                        End If
                        If DataT.Rows(i)("Account_Type").ToString = "5" Or DataT.Rows(i)("Account_Type").ToString = "6" Then

                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE - Bal.Rows(jj)(Balance)
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next

            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next

                ' Assign the Value to 39000
                If COA.Rows(i)("Account_No").ToString = "39000" Then
                    COA.Rows(i)(Balance) = RE
                End If

            Next

            For i = 0 To COA.Rows.Count - 1
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                    COA.Rows(i)(Balance) = denominatedValue
                End If
            Next

            ' Format all the output for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1

                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                    If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = ""
                Next
            End If

            COA.AcceptChanges()
        Next

        ' End of For loop
        ' Delete the rows that arnt above the detail level 
        For i = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False
            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()
            End If
        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        'Calculation for Per
        For i = 0 To COA.Rows.Count - 1
            If COA.Rows(i)("Balance0").ToString = "" Then
                Bal0 = 0
            ElseIf COA.Rows(i)("Balance0") <> 0 Then
                Bal0 = COA.Rows(i)("Balance0")
                If j >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal1 = 0
                    Else
                        Bal1 = COA.Rows(i)("Balance1")
                        COA.Rows(i)("Per") = (((Bal1 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                        If j = 3 Then
                            If COA.Rows(i)("Balance2").ToString = "" Then
                                Bal2 = 0
                            Else
                                Bal2 = COA.Rows(i)("Balance2")
                                COA.Rows(i)("Per") = (((Bal2 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                            End If

                        End If
                    End If
                End If
            End If

            COA.AcceptChanges()
            ' Format all the output for the paper  
            If COA.Rows(i)("Per").ToString <> "" Then
                COA.Rows(i)("Per") = (Math.Round(Val(COA.Rows(i)("Per").ToString), 2)).ToString() & "%"
            End If


            If Left(COA.Rows(i)("Per").ToString, 1) = "-" Then COA.Rows(i)("Per") = "(" & COA.Rows(i)("Per").replace("-", "") & ")"

            ' If Request.Form("Round") = "on" Then
            '     COA.Rows(i)("Per") = Format(Val(COA.Rows(i)("Per").ToString), "##")
            ' End If

            If COA.Rows(i)("Per").ToString = "0.##%" Or COA.Rows(i)("Per").ToString = "%" Then COA.Rows(i)("Per") = ""

            COA.AcceptChanges()

            ' Post on Report DataTable

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            Dim style_per As String


            If Show_Per = "on" Then
                style_per = Style2
                If (Left(COA.Rows(i)("Per").ToString, 1) = "(") Then
                    style_per = style_per & "color: red !important;"
                End If
            End If

            If j = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", style_per, "<span style=""" + StyleFinish + """ > " + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
            If COA.Rows(i)("Account_No").ToString = "39999" Then Exit For
        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")

            ' Calculate the Percentage
            'For i = 0 To COA.Rows.Count - 1
            '    If COA.Rows(i)("Balance0").ToString = "" Then
            '        Bal0 = 0
            '    ElseIf COA.Rows(i)("Balance0") <> 0 Then
            '        Bal0 = COA.Rows(i)("Balance0")
            '        If COA.Rows(i)("Balance" & Convert.ToInt32(GoPrevious) - 1).ToString = "" Then
            '            Bal1 = 0
            '        Else
            '            Bal1 = COA.Rows(i)("Balance" & Convert.ToInt32(GoPrevious) - 1)
            '            COA.Rows(i)("Per") = (((Bal1 - Bal0) / Math.Abs(Bal0)) * 100).ToString
            '        End If
            '    End If
            'Next

            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)

                If COA.Rows(i)(0).ToString = "39999" Then
                    Exit For
                End If
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Convert.ToInt32(GoPrevious) - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To Convert.ToInt32(GoPrevious) - 1
                excelHeader("value" + (i + 2).ToString) = endDate.AddYears(i - (GoPrevious - 1)).ToString("MMMM yyyy")
                If i = Convert.ToInt32(GoPrevious) - 1 Then
                    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Balance Sheet Multiperiod Quarterly
    Private Sub PrintQuarterlyBalSheetMultiPer()
        Dim Language As Integer = Request.Form("language")
        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim MonthCount As Integer = 3
        Dim Padding As Integer = 0
        Dim j As Integer = 0
        Dim Level As Integer = 1

        Dim Year As String = ""
        Dim Qua_1 As String = ""
        Dim Qua_2 As String = ""
        Dim Qua_3 As String = ""
        Dim Qua_4 As String = ""

        Dim Qua_1_StartDate As String = ""
        Dim Qua_1_EndDate As String = ""
        Dim Qua_2_StartDate As String = ""
        Dim Qua_2_EndDate As String = ""
        Dim Qua_3_StartDate As String = ""
        Dim Qua_3_EndDate As String = ""
        Dim Qua_4_StartDate As String = ""
        Dim Qua_4_EndDate As String = ""

        Dim seconDate, startDate, Asterix As String

        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"
        Dim Date2 As String

        Year = Request.Form("YearForQuater")

        Qua_1 = Request.Item("Q1")
        Qua_2 = Request.Item("Q2")
        Qua_3 = Request.Item("Q3")
        Qua_4 = Request.Item("Q4")
        Dim Q As Integer = 0

        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        Dim StyleMonth As String
        Dim Quarter(4) As String
        Dim HF_Acc As String = ""

        Dim COA, DataT, Bal, Bal1, Bal2, Fiscal, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        Dim FiscalDate, date1 As String

        Dim d1 As Date

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = (Year - 1) & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate
        Else
            FiscalDate = (Year - 1) & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        End If

        date1 = FiscalDate

        Qua_1_StartDate = d1
        Qua_1_EndDate = d1.AddMonths(3).AddDays(-1)

        Qua_2_StartDate = d1.AddMonths(3)
        Qua_2_EndDate = d1.AddMonths(6).AddDays(-1)

        Qua_3_StartDate = d1.AddMonths(6)
        Qua_3_EndDate = d1.AddMonths(9).AddDays(-1)

        Qua_4_StartDate = d1.AddMonths(9)
        Qua_4_EndDate = d1.AddMonths(12).AddDays(-1)

        Asterix = ""

        ' Check if the quarter picked is today's quarter
        ' Check the year first then check the month compare to fiscal
        If ((Year = DateTime.Now.Year And DateTime.Now.Month < Fiscal.Rows(0)("Fiscal_Year_Start_Month")) Or (Year = (DateTime.Now.Year - 1) And DateTime.Now.Month >= Fiscal.Rows(0)("Fiscal_Year_Start_Month"))) Then
            ' Check the month to see if it's today's quarter got selected
            ' Need to Mark which quarter is today's
            If (Today() >= d1 And Now() <= d1.AddMonths(3).AddDays(-1) And Qua_1 = "on") Then
                ' It's in Q1
                Qua_1_EndDate = Now().ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(3) And Now() <= d1.AddMonths(6).AddDays(-1) And Qua_2 = "on") Then
                ' It's in Q2
                Qua_2_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(6) And Now() <= d1.AddMonths(9).AddDays(-1) And Qua_3 = "on") Then
                ' It's in Q3
                Qua_3_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_2_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-6).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            ElseIf (Today() >= d1.AddMonths(9) And Now() <= d1.AddMonths(12).AddDays(-1) And Qua_4 = "on") Then
                ' It's in Q4
                Qua_4_EndDate = Now().ToString("yyyy-MM-dd")
                Qua_3_EndDate = Now().AddMonths(-3).ToString("yyyy-MM-dd")
                Qua_2_EndDate = Now().AddMonths(-6).ToString("yyyy-MM-dd")
                Qua_1_EndDate = Now().AddMonths(-9).ToString("yyyy-MM-dd")
                Asterix = "(*)"
            End If
        End If

        If (Qua_1 = "on") Then
            If Language = 0 Then
                Quarter(0) = "Q-1"
            ElseIf Language = 1 Then
                Quarter(0) = "T-1"
            End If

            Query1 = Query1 & ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            Query2 = Query2 & ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & Qua_1_EndDate & "' order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            seconDate = Qua_1_EndDate
            startDate = Qua_1_StartDate
            Q += 1
        End If
        If (Qua_2 = "on") Then
            If Language = 0 Then
                Quarter(1) = "Q-2"
            ElseIf Language = 1 Then
                Quarter(1) = "T-2"
            End If

            Query1 = Query1 & ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            Query2 = Query2 & ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & Qua_2_EndDate & "' order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            seconDate = Qua_2_EndDate
            If Q = 0 Then
                startDate = Qua_2_StartDate
            End If

            Q += 1
        End If
        If (Qua_3 = "on") Then
            If Language = 0 Then
                Quarter(2) = "Q-3"
            ElseIf Language = 1 Then
                Quarter(2) = "T-3"
            End If

            Query1 = Query1 & ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            Query2 = Query2 & ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & Qua_3_EndDate & "' order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            seconDate = Qua_3_EndDate
            If Q = 0 Then
                startDate = Qua_3_StartDate
            End If
            Q += 1
        End If
        If (Qua_4 = "on") Then
            If Language = 0 Then
                Quarter(3) = "Q-4"
            ElseIf Language = 1 Then
                Quarter(3) = "T-4"
            End If

            Query1 = Query1 & ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            Query2 = Query2 & ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & Qua_4_EndDate & "' order by Transaction_Date desc, rowID desc) as Balance" & Q.ToString
            seconDate = Qua_4_EndDate
            If Q = 0 Then
                startDate = Qua_4_StartDate
            End If
            Q += 1
        End If

        Dim H_Quarter, H_Qua_1 As String
        Dim Temp As Integer

        For l = 0 To 3
            If Quarter(l) <> "" Then
                H_Quarter = Quarter(l) + Asterix
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + H_Quarter
                If Temp < (Q - 1) Then
                    If Temp < (Q - 2) Then
                        H_Qua_1 = H_Qua_1 + Quarter(l) + ", "
                    Else
                        H_Qua_1 = H_Qua_1 + Quarter(l)
                    End If
                ElseIf Temp = (Q - 1) Then
                    If Language = 0 Then
                        H_Qua_1 = H_Qua_1 + " and " + Quarter(l)
                    ElseIf Language = 1 Then
                        H_Qua_1 = H_Qua_1 + " y " + Quarter(l)
                    End If

                End If
                Temp += 1
            End If
        Next

        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Account Description" + StyleMonth + "~Text-align: Right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Multiperiod Balance Sheet(Quarterly) " + DenomString + "<br/>For " & H_Qua_1 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; width:5px; font-size:8pt~Nombre De La Cuenta" + StyleMonth + "~Text-align:right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Hoja de Balance Multi Período (Trimestral) " + DenomString + "<br/>Para " & H_Qua_1 & "<br/></span><span style=""font-size:7pt"">Impreso el " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        ' Translation
        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        'SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash" & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 And Account_No >= 10000 And Account_No<'40000' order by Account_No"
        'SQLCommand.Parameters.Clear()
        'DataAdapter.Fill(COA)

        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID" & Query2 & " From ACC_GL As gl1"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Bal)

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling Active, Cash, Exchange_Account_ID, Associated_Totalling From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(DataT)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))

        Dim Balance As String
        Dim BalanceString As String = ""
        Dim ColMonth As String = ""

        j = 0
        Dim k As Int32 = 0
        Dim RE As Decimal = 0
        Balance = ""
        BalanceString = ""

        For col = 0 To Q - 1
            Balance = "Balance" + k.ToString
            BalanceString = "BalanceString" + k.ToString
            RE = 0

            ' Calculation for Current Retained Earning (39000)
            For j = 0 To DataT.Rows.Count - 1

                For jj = 0 To Bal.Rows.Count - 1
                    If DataT.Rows(j)("Account_ID").ToString = Bal.Rows(jj)("Account_ID").ToString Then

                        If DataT.Rows(j)("Account_Type").ToString = "4" Then

                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE + Bal.Rows(jj)(Balance)
                            End If
                        End If
                        If DataT.Rows(j)("Account_Type").ToString = "5" Or DataT.Rows(j)("Account_Type").ToString = "6" Then

                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE - Bal.Rows(jj)(Balance)
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next



            For i = 0 To COA.Rows.Count - 1
                For ii = 0 To Bal.Rows.Count - 1
                    ' Copying the Balance value from table Bal to table COA
                    If COA.Rows(i)("Account_ID").ToString = Bal.Rows(ii)("Account_ID").ToString Then
                        COA.Rows(i)(Balance) = Bal.Rows(ii)(Balance)
                        Exit For
                    End If
                Next

                ' Give Padding
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
                If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = "<div style='min-width: 0.5in; max-width:0.5in;'></div>" ' hard coded
            Next

            For j = 0 To COA.Rows.Count - 1
                If COA.Rows(j)("Account_No") = "39000" Then COA.Rows(j)(Balance) = RE
                COA.AcceptChanges()
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""
            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If

                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next
            Next

            Total = 0
            Account = ""
            For j = 1 To 2
                For i = 0 To COA.Rows.Count - 1
                    If COA.Rows(i)("Totalling").ToString <> "" Then
                        Total = 0
                        Account = COA.Rows(i)("Account_No").ToString
                        Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                        For ii = 0 To Plus.Length - 1
                            Dim Dash() As String = Plus(ii).Split("-")
                            Dim Start As String = Trim(Dash(0))
                            Dim Endd As String
                            If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                            For iii = 0 To COA.Rows.Count - 1
                                If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            Next
                        Next
                    End If

                Next
            Next
            COA.AcceptChanges()

            ' Denomination and Rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Format all the output for the paper
            For i = 0 To COA.Rows.Count - 1
                COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                If Request.Form("Round") = "on" Then
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                Else
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                End If

                If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
            Next
            COA.AcceptChanges()
            k += 1
        Next

        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If k = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf k = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf k = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()
            End If
        Next

        COA.AcceptChanges()

        'End While
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
        For i = 0 To COA.Rows.Count - 1


            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If

            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If
            If Q = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2 + "min-width: 5in; max-width: 5in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2 + "min-width: 3.0in; max-width: 3.0in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2 + "min-width: 2.0in; max-width: 2.0in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2 + "min-width: 1.5in; max-width: 1.5in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2 + "min-width: 1.5in; max-width: 1.5in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2 + "min-width: 1.5in; max-width: 1.5in", "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If
        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)

                If COA.Rows(i)(0).ToString = "39999" Then
                    Exit For
                End If
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Q
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next
            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            Dim count As Integer = 0
            For i = 0 To Q
                While (Quarter(i + count) = "" And (i + count) < Quarter.Length - 1)
                    count += 1
                End While
                If Quarter(i + count) <> "" Then
                    excelHeader("value" + (i + 2).ToString()) = Quarter(i + count) + Asterix
                End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Balance Sheet Multiperiod Quarter-to-Quarter
    Private Sub PrintQuarterToQuarterBalSheetMultiPer()
        Dim Language As Integer = Request.Form("language")
        Dim seconDate As String = Request.Form("SecQuarter")
        Dim Acc_No As String = Request.Form("Ac")
        Dim words As String() = seconDate.Split(New Char() {" "c})
        Dim Qua_No As Integer = Integer.Parse(words(0))
        Dim Qua_Year As String = words(1).ToString
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim Show_Per As String = Request.Form("Percentage")
        Dim GoPrevious As String = Request.Form("prev")
        Dim Quarter(4) As String
        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        Dim i As Integer = 0
        Dim HF_Acc As String = ""
        Dim StyleFinish As String = ""
        Dim DenomString As String = ""
        Dim RE As Decimal = 0
        Dim j As Integer = 0
        Dim endDate2 As String
        Dim Date1 As String
        Dim Date2 As String
        Dim Balance As String
        Dim BalanceString As String = ""
        Dim StyleMonth As String
        Dim Bal0 As Decimal
        Dim Bal1 As Decimal
        Dim Bal2 As Decimal

        Dim COA, Bal, Report, DataT, Fiscal As New DataTable

        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        Dim FiscalDate, Qua_date1, Qua_date2 As String

        Dim d1 As Date

        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = Qua_Year & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate
        Else
            FiscalDate = Qua_Year & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        End If

        If Qua_No = 1 Then
            Qua_date1 = d1
            Qua_date2 = d1.AddMonths(3).AddDays(-1)
        ElseIf Qua_No = 2 Then
            Qua_date1 = d1.AddMonths(3)
            Qua_date2 = d1.AddMonths(6).AddDays(-1)
        ElseIf Qua_No = 3 Then
            Qua_date1 = d1.AddMonths(6)
            Qua_date2 = d1.AddMonths(9).AddDays(-1)
        ElseIf Qua_No = 4 Then
            Qua_date1 = d1.AddMonths(9)
            Qua_date2 = d1.AddMonths(12).AddDays(-1)
        End If

        Dim endDate As Date = Qua_date2

        If endDate >= Today() Then
            'If (((Qua_Year + 1) = DateTime.Now.Year And DateTime.Now.Month < Fiscal.Rows(0)("Fiscal_Year_Start_Month")) Or ((Qua_Year + 1) = (DateTime.Now.Year - 1) And DateTime.Now.Month >= Fiscal.Rows(0)("Fiscal_Year_Start_Month"))) Then
            ' Check the month to see if it's today's quarter got selected
            ' Need to Mark which quarter is today's
            If (Today() >= d1 And Now() <= d1.AddMonths(3).AddDays(-1)) Then
                ' It's in Q1
                Qua_date1 = d1
                Qua_date2 = Now().ToString("yyyy-MM-dd")
            ElseIf (Today() >= d1.AddMonths(3) And Now() <= d1.AddMonths(6).AddDays(-1)) Then
                ' It's in Q2
                Qua_date1 = d1.AddMonths(3)
                Qua_date2 = Now().ToString("yyyy-MM-dd")

            ElseIf (Today() >= d1.AddMonths(6) And Now() <= d1.AddMonths(9).AddDays(-1)) Then
                ' It's in Q3
                Qua_date1 = d1.AddMonths(6)
                Qua_date2 = Now().ToString("yyyy-MM-dd")

            ElseIf (Today() >= d1.AddMonths(9) And Now() <= d1.AddMonths(12).AddDays(-1)) Then
                ' It's in Q4
                Qua_date1 = d1.AddMonths(9)
                Qua_date2 = Now().ToString("yyyy-MM-dd")

            End If
        End If

        endDate = Qua_date2
        Dim E_date As Date = endDate
        endDate = endDate.AddYears(-(GoPrevious - 1))
        Dim secondDate As Date = endDate

        'Translation for Header
        If Language = 0 Then
            For j = 0 To GoPrevious - 1

                If E_date >= Today() Then
                    Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2)) & "(*)"
                Else
                    Date1 = "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                End If

                If j = (GoPrevious - 1) Then
                    Date2 = Date2 & " and Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                ElseIf j = (GoPrevious - 2) Then
                    Date2 = Date2 & "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                Else
                    Date2 = Date2 & "Q" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2)).ToString() + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + Date1

                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                secondDate = endDate2
            Next
        ElseIf Language = 1 Then
            For j = 0 To GoPrevious - 1

                If E_date >= Today() Then
                    Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2)) & "(*)"
                Else
                    Date1 = "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                End If

                If j = (GoPrevious - 1) Then
                    Date2 = Date2 & " y T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                ElseIf j = (GoPrevious - 2) Then
                    Date2 = Date2 & "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2))
                Else
                    Date2 = Date2 & "T" & Qua_No & " " & ((Convert.ToInt32(Qua_Year) + j) - (Convert.ToInt32(GoPrevious) - 2)).ToString() + ", "
                End If

                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + Date1

                endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
                secondDate = endDate2
            Next
        End If

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Show Account No
        Dim Per_opt As String = ""
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If
        ' Show Percentage
        If Show_Per = "on" Then
            If Language = 0 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Percentage(%)"
            ElseIf Language = 1 Then
                Per_opt = "~Text-align: Right; width:80px; font-size:8pt~Porcentaje(%)"
            End If

        Else
            Per_opt = "~width:0px;font-size:0pt~"
        End If


        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; font-size:8pt~Account Description" & StyleMonth & "" & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Balance Sheet(Quarter To Quarter) " + DenomString + " <br/>For " & Date2 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; " & HF_Acc & "~text-align:left; width:50px; font-size:8pt~Descripción De Cuenta" & StyleMonth & "" & Per_opt
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Estado de Resultados de Varios Períodos (Mensual) " + DenomString + "<br/>Para " & Date2 & "<br/></span><span style=""font-size:7pt"">Impreso En " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        ' Getting the Query

        secondDate = endDate
        endDate2 = endDate

        For j = 0 To GoPrevious - 1
            Query1 = Query1 + ", (Select Top 1 Balance from ACC_GL where Transaction_Date <= '" & endDate2 & "' and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance" + j.ToString()
            Query2 = Query2 + ", (Select Top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <= '" & endDate2 & "' order by Transaction_Date desc, rowID desc) as Balance" + j.ToString()

            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd")
            secondDate = endDate2

        Next

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)


        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)

        End If


        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID" & Query2 & " from ACC_GL gl1"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Bal)

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling Active, Cash, Exchange_Account_ID, Associated_Totalling From ACC_GL_Accounts WHERE Account_Type >=0 and Account_ID > 1 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(DataT)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))
        COA.Columns.Add("Per", GetType(String))

        Balance = ""
        BalanceString = ""

        For j = 0 To GoPrevious - 1
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    ' If Denom > 1 Then
                    '     Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                    '     COA.Rows(i)(Balance) = denominatedValueCurrent
                    ' End If

                Next
            End If

            ' Give Padding

            For i = 0 To COA.Rows.Count - 1
                If i > 0 Then
                    If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                    If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                    If Padding < 0 Then Padding = 0
                    If Level < 1 Then Level = 1
                End If
                COA.Rows(i)("Padding") = Padding
                COA.Rows(i)("Level") = Level
            Next

            Dim Total As Decimal = 0
            Dim Account As String = ""

            RE = 0

            ' Calculation for Current Retained Earning (39000)
            For i = 0 To DataT.Rows.Count - 1

                For jj = 0 To Bal.Rows.Count - 1

                    If DataT.Rows(i)("Account_ID").ToString = Bal.Rows(jj)("Account_ID").ToString Then

                        If DataT.Rows(i)("Account_Type").ToString = "4" Then

                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE + Bal.Rows(jj)(Balance)
                            End If
                        End If
                        If DataT.Rows(i)("Account_Type").ToString = "5" Or DataT.Rows(i)("Account_Type").ToString = "6" Then

                            If Bal.Rows(jj)(Balance).ToString = "" Then

                            Else
                                RE = RE - Bal.Rows(jj)(Balance)
                            End If
                            Exit For
                        End If
                    End If
                Next
            Next

            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next

                ' Assign the Value to 39000
                If COA.Rows(i)("Account_No").ToString = "39000" Then
                    COA.Rows(i)(Balance) = RE
                End If

            Next

            For i = 0 To COA.Rows.Count - 1
                If Denom > 1 Then
                    Dim denominatedValue As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                    COA.Rows(i)(Balance) = denominatedValue
                End If
            Next

            ' Format all the output for the paper
            If j < 3 Then
                For i = 0 To COA.Rows.Count - 1

                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                    If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = ""
                Next
            End If

            COA.AcceptChanges()
        Next

        ' End of For loop

        ' Delete the rows that arnt above the detail level 
        For i = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If j = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf j = 3 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        'Calculation for Per

        For i = 0 To COA.Rows.Count - 1

            If COA.Rows(i)("Balance0").ToString = "" Then
                Bal0 = 0
            ElseIf COA.Rows(i)("Balance0") <> 0 Then
                Bal0 = COA.Rows(i)("Balance0")
                If j >= 2 Then
                    If COA.Rows(i)("Balance1").ToString = "" Then
                        Bal1 = 0
                    Else
                        Bal1 = COA.Rows(i)("Balance1")
                        COA.Rows(i)("Per") = (((Bal1 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                        If j = 3 Then
                            If COA.Rows(i)("Balance2").ToString = "" Then
                                Bal2 = 0
                            Else
                                Bal2 = COA.Rows(i)("Balance2")
                                COA.Rows(i)("Per") = (((Bal2 - Bal0) / Math.Abs(Bal0)) * 100).ToString
                            End If

                        End If
                    End If
                End If
            End If

            COA.AcceptChanges()
            ' Format all the output for the paper  
            If COA.Rows(i)("Per").ToString <> "" Then
                COA.Rows(i)("Per") = (Math.Round(Val(COA.Rows(i)("Per").ToString), 2)).ToString() & "%"
            End If

            If Left(COA.Rows(i)("Per").ToString, 1) = "-" Then COA.Rows(i)("Per") = "(" & COA.Rows(i)("Per").replace("-", "") & ")"

            ' If Request.Form("Round") = "on" Then
            '     COA.Rows(i)("Per") = Format(Val(COA.Rows(i)("Per").ToString), "##")
            ' End If

            If COA.Rows(i)("Per").ToString = "0.##%" Or COA.Rows(i)("Per").ToString = "%" Then COA.Rows(i)("Per") = ""

            COA.AcceptChanges()

            ' Post on Report DataTable

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            Dim style_per As String


            If Show_Per = "on" Then
                style_per = Style2
                If (Left(COA.Rows(i)("Per").ToString, 1) = "(") Then
                    style_per = style_per & "color: red !important;"
                End If
            End If

            If j = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", style_per, "<span style=""" + StyleFinish + """ > " + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", style_per, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Per") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

            If COA.Rows(i)("Account_No").ToString = "39999" Then Exit For

        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Direct_Posting")
            COA.Columns.Remove("fk_Linked_ID")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Active")
            COA.Columns.Remove("Cash")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")
            COA.Columns.Remove("Total")
            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)

                If COA.Rows(i)(0).ToString = "39999" Then
                    Exit For
                End If
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To Convert.ToInt32(GoPrevious) - 1
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To Convert.ToInt32(GoPrevious) - 1
                excelHeader("value" + (i + 2).ToString) = endDate.AddYears(i - (GoPrevious - 1)).ToString("MMMM yyyy")
                If i = Convert.ToInt32(GoPrevious) - 1 Then
                    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    ' Balance Sheet Multiperiod Yearly
    Private Sub PrintYearlyBalSheetMultiPer()
        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        Dim COA, Bal, Report, Fiscal As New DataTable

        Dim Language As Integer = Request.Form("language")
        Dim firstDate As String = Request.Form("FirstDate")
        Dim secondDate As String = Request.Form("SecondDate")
        Dim Acc_No As String = Request.Form("Ac")
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")

        Dim startDate, FiscalDate As String
        Dim date1 As String
        Dim d1, d2, d3, dtemp As Date

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim YearCount As Integer = secondDate - firstDate
        Dim Padding As Integer = 0
        Dim Level As Integer = 1
        Dim StyleMonth As String

        Dim HF_Acc As String = ""
        Dim StyleFinish As String = ""
        Dim DenomString As String = ""

        Dim RE As Decimal = 0
        Dim Asterix As String = ""

        If (Denom > 1) Then
            If Language = 0 Then
                If Denom = 10 Then
                    DenomString = "(In Tenth)"
                ElseIf Denom = 100 Then
                    DenomString = "(In Hundreds)"
                ElseIf Denom = 1000 Then
                    DenomString = "(In Thousands)"
                End If
            ElseIf Language = 1 Then
                If Denom = 10 Then
                    DenomString = "(En Décimo)"
                ElseIf Denom = 100 Then
                    DenomString = "(En Centenares)"
                ElseIf Denom = 1000 Then
                    DenomString = "(En Miles)"
                End If
            End If
        End If

        ' Get the fiscal month
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        ' Because it is '9' not '09'
        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") > 10 Then
            FiscalDate = secondDate & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
        Else
            FiscalDate = secondDate & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString & "-01"
        End If

        date1 = FiscalDate

        ' if date picked is current year, check if today's month is > fiscal month
        If secondDate >= DateTime.Now.Year Then
            ' Check if today's date already pass the fiscal month
            If DateTime.Now.Month < Fiscal.Rows(0)("Fiscal_Year_Start_Month") Then
                ' use today's date to compare with previous year
                date1 = Now().ToString("yyyy-MM-dd")
                d1 = date1
                dtemp = d1
                Asterix = "(*)"
            Else
                d1 = date1
                d1 = d1.AddDays(-1)
                dtemp = d1
            End If
        Else
            d1 = date1
            d1 = d1.AddDays(-1)
            dtemp = d1
        End If

        If YearCount = 1 Then
            d2 = d1.AddYears(-1)
            dtemp = d2
        ElseIf YearCount = 2 Then
            d2 = d1.AddYears(-1)
            d3 = d2.AddYears(-1)
            dtemp = d3
        Else
            dtemp = dtemp.AddYears(-YearCount)
        End If

        Dim H_year As String
        ' Header
        For i = 0 To YearCount
            If i = YearCount Then
                If Language = 0 Then
                    H_year = H_year + " and " & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy")
                ElseIf Language = 1 Then
                    H_year = H_year + " y " & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy")
                End If

            ElseIf i = (YearCount - 1) Then
                H_year = H_year & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy")
            Else
                H_year = H_year & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy") & ", "
            End If
            'If Asterix = "1" Then
            '    StyleMonth = StyleMonth + "~Text-align:right; width:120px; font-size:8pt~" & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy") & "(*)"
            'Else
            StyleMonth = StyleMonth + "~Text-align:right; width:120px; font-size:8pt~" & dtemp.AddYears(i - 1).ToString("yyyy") + "-" & dtemp.AddYears(i).ToString("yyyy") & Asterix
            'End If
        Next
        If Acc_No = "on" Then
            HF_Acc = "width:60px; ~Act No"
        Else
            HF_Acc = "~"
        End If

        ' Translate the Header and Title
        If Language = 0 Then
            HF_PrintHeader.Value = "Text-align: Left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Account Description" + StyleMonth + "~Text-align:right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Multiperiod Balance Sheet(Yearly) " + DenomString + "<br/>For " & H_year & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        ElseIf Language = 1 Then
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt" & HF_Acc & "~text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + "~Text-align:right; width:0px; font-size:8pt~"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">" & Fiscal.Rows(0)("Company_Name").ToString & "<br/>Hoja de Balance Multi Período (Anual) " + DenomString + "<br/>Para " & H_year & "<br/></span><span style=""font-size:7pt"">Impreso el " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        End If

        PNL_Summary.Visible = True

        ' Getting the Query
        For i = 0 To YearCount
            startDate = dtemp.AddYears(i).ToString("yyyy/MM/dd")
            Query1 = Query1 & ", (Select top 1 Balance from ACC_GL where gl1.fk_Account_ID = fk_Account_ID and Transaction_Date <='" & startDate & "' order by Transaction_Date desc, rowID desc) as Balance" + i.ToString
        Next

        If Language = 0 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling From ACC_GL_Accounts order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        ElseIf Language = 1 Then
            SQLCommand.CommandText = "Select Account_ID, Account_No, TranslatedName as Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling From ACC_GL_Accounts order by Account_No"
            SQLCommand.Parameters.Clear()
            DataAdapter.Fill(COA)
        End If

        SQLCommand.CommandText = "Select Distinct(gl1.fk_Account_ID) as Account_ID" & Query1 & " from ACC_GL gl1"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Bal)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        For i = 0 To YearCount
            COA.Columns.Add("Balance" + i.ToString, GetType(String))
        Next
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))
        COA.Columns.Add("Total", GetType(String))

        Dim Balance As String
        Dim BalanceString As String = ""

        Balance = ""
        BalanceString = ""

        ' Give Padding
        For i = 0 To COA.Rows.Count - 1
            If i > 0 Then
                If COA.Rows(i - 1)("Account_Type").ToString = "98" Then Padding = Padding + 20 : Level = Level + 1
                If COA.Rows(i)("Account_Type").ToString = "99" Then Padding = Padding - 20 : Level = Level - 1
                If Padding < 0 Then Padding = 0
                If Level < 1 Then Level = 1
            End If
            COA.Rows(i)("Padding") = Padding
            COA.Rows(i)("Level") = Level
        Next

        ' For loop for Calculation and  Formatting
        For j = 0 To YearCount
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString


            Dim Total As Decimal = 0
            Dim Account As String = ""

            RE = 0

            ' calculation for 39000
            For i = 0 To COA.Rows.Count - 1
                ' Calculation for 39000
                For ii = 0 To Bal.Rows.Count - 1
                    If COA.Rows(i)("Account_ID").ToString = Bal.Rows(ii)("Account_ID").ToString Then
                        COA.Rows(i)(Balance) = Bal.Rows(ii)(Balance)
                        If COA.Rows(i)("Account_Type").ToString = "4" And Bal.Rows(ii)(Balance).ToString <> "" Then RE = RE + Val(Bal.Rows(ii)(Balance))
                        If (COA.Rows(i)("Account_Type").ToString = "5" Or COA.Rows(i)("Account_Type").ToString = "6") And Bal.Rows(ii)(Balance).ToString <> "" Then RE = RE - Val(Bal.Rows(ii)(Balance))
                        Exit For
                    End If
                Next
            Next

            ' Calculating Sub-Total and Total
            For i = 0 To COA.Rows.Count - 1
                If COA.Rows(i)("Totalling").ToString <> "" Then
                    Total = 0
                    Account = COA.Rows(i)("Account_No").ToString
                    Dim Plus() As String = COA.Rows(i)("Totalling").ToString.Split("+")
                    For ii = 0 To Plus.Length - 1
                        Dim Dash() As String = Plus(ii).Split("-")
                        Dim Start As String = Trim(Dash(0))
                        Dim Endd As String
                        If Dash.Length = 1 Then Endd = Trim(Dash(0)) Else Endd = Trim(Dash(1))
                        For iii = 0 To COA.Rows.Count - 1
                            If Trim(COA.Rows(iii)("Account_No").ToString) > Endd Then Exit For
                            If Trim(COA.Rows(iii)("Account_No").ToString) >= Start And COA.Rows(iii)("Account_Type") < 90 Then Total = Total + Val(COA.Rows(iii)(Balance).ToString.Replace(",", "").Replace("$", ""))
                        Next
                    Next
                End If
                For ii = 0 To COA.Rows.Count - 1
                    If COA.Rows(ii)("Account_No") = Account Then COA.Rows(ii)(Balance) = Total
                Next

                ' Assign the Value to 39000
                If COA.Rows(i)("Account_No").ToString = "39000" Then
                    COA.Rows(i)(Balance) = RE
                End If

            Next

            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString))
                    End If
                    If Denom > 1 Then
                        Dim denominatedValueCurrent As Double = Convert.ToDouble(Val(COA.Rows(i)(Balance).ToString)) / Denom
                        COA.Rows(i)(Balance) = denominatedValueCurrent
                    End If

                Next
            End If

            ' Format all the output for the paper
            If YearCount < 3 Then
                For i = 0 To COA.Rows.Count - 1

                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                    Else
                        COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                    End If

                    If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Or COA.Rows(i)(BalanceString).ToString = "$,00" Then COA.Rows(i)(BalanceString) = ""
                    If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
                    If COA.Rows(i)("fk_Currency_ID").ToString = "CAD" Then COA.Rows(i)("fk_Currency_ID") = ""
                Next
            End If

            COA.AcceptChanges()
        Next
        ' End of For loop

        ' Delete the rows that arnt above the detail level 
        For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            Dim AlreadyDeleted As Boolean = False

            If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                If YearCount = 0 Then
                    If COA.Rows(i)("BalanceString0") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf YearCount = 1 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                ElseIf YearCount = 2 Then
                    If COA.Rows(i)("BalanceString0") = "" And COA.Rows(i)("BalanceString1") = "" And COA.Rows(i)("BalanceString2") = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                    End If
                End If
            End If
            If (AlreadyDeleted = False) Then
                If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            End If

        Next i

        COA.AcceptChanges()

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"

        ' Post on Report DataTable
        For i = 0 To COA.Rows.Count - 1

            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2in; max-width: 2in;"
            Style2 = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
            Dim Style3 As String = "padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; min-width: 5px; max-width: 5px;"
            StyleFinish = ""
            If COA.Rows(i)("Account_Type") > 90 Then
                Style = Style & "; font-weight:bold;border-top: px solid black "
                Style2 = Style2 & "; font-weight:bold;border-top: px solid black; font-size:8pt;text-align:right "
            End If
            If COA.Rows(i)("Totalling").ToString <> "" Then
                'Style1 = Style1 & "; font-weight:bold"
                Style = Style & "; border-bottom: 0px solid black;padding-bottom:10px;"
                Style2 = Style2 & "; padding-bottom:10px;"
                StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;"
                Style3 = Style3 & ";padding-bottom:10px;"
            End If
            Dim Ac_Style = " font-size:0pt;"

            If Acc_No = "on" Then
                Ac_Style = "text-align:left;font-size:8pt;width: 10px;"
            End If

            If YearCount = 0 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf YearCount = 1 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf YearCount >= 2 Then
                Report.Rows.Add(Ac_Style, COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

            If COA.Rows(i)("Account_No").ToString = "39999" Then Exit For

        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

        ' Export function
        If Request.Form("expStat") = "on" Then
            ' Remove the columns that do not need to be display in excel
            COA.Columns.Remove("Account_ID")
            COA.Columns.Remove("fk_Currency_ID")
            COA.Columns.Remove("Account_Type")
            COA.Columns.Remove("Totalling")
            COA.Columns.Remove("Padding")
            COA.Columns.Remove("Level")
            COA.Columns.Remove("BalanceString0")
            COA.Columns.Remove("BalanceString1")
            COA.Columns.Remove("BalanceString2")

            ' Create new Datatable
            Dim exportTable As New DataTable

            For i = 0 To COA.Columns.Count
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Copy the data value
            For i = 0 To COA.Rows.Count - 1
                Dim excelRow As DataRow = exportTable.NewRow()
                For ii = 0 To COA.Columns.Count - 1
                    excelRow("value" + ii.ToString) = COA.Rows(i)(ii).ToString
                Next

                exportTable.Rows.Add(excelRow)

                If COA.Rows(i)(0).ToString = "39999" Then
                    Exit For
                End If
            Next

            ' Creating new column to value20
            For i = exportTable.Columns.Count To 25
                exportTable.Columns.Add("value" + i.ToString, GetType(String))
            Next

            ' Formatting the value
            For i = 0 To exportTable.Rows.Count - 1
                For ii = 0 To YearCount
                    exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###.00")

                    If Request.Form("Round") = "on" Then
                        exportTable.Rows(i)("value" + (ii + 2).ToString) = Format(Val(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString), "$#,###")
                    End If

                    If exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$.00" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$" Or exportTable.Rows(i)("value" + (ii + 2).ToString).ToString = "$,00" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = ""
                    If Left(exportTable.Rows(i)("value" + (ii + 2).ToString).ToString, 1) = "-" Then exportTable.Rows(i)("value" + (ii + 2).ToString) = "(" & exportTable.Rows(i)("value" + (ii + 2).ToString).replace("-", "") & ")"
                Next

            Next

            ' Add the header as "value"
            Dim excelHeader = exportTable.NewRow()
            excelHeader("value0") = "Account No"
            excelHeader("value1") = "Account Description"

            ' Add the header with dynamic number of columns
            For i = 0 To YearCount
                excelHeader("value" + (i + 2).ToString) = dtemp.AddYears(i - 1).ToString("yyyy") & "-" & dtemp.AddYears(i).ToString("yyyy") & Asterix
                'If i = YearCount Then
                '    excelHeader("value" + (i + 3).ToString) = "Growth (%)"
                'End If
            Next

            exportTable.Rows.InsertAt(excelHeader, 0)

            ' Bold the header.
            For i = 0 To exportTable.Columns.Count - 1
                exportTable.Rows(0)(i) = "<b>" & exportTable.Rows(0)(i).ToString & "</b>"
            Next

            RPT_Excel.DataSource = exportTable
            RPT_Excel.DataBind()

            PNL_Excel.Visible = True
        End If
    End Sub

    Private Sub Report()
        Dim Cur, Cust, Vend As New DataTable
        Dim intYear As Integer = DateTime.Now.Year
        Dim intMonth As Date = Now().ToString("yyyy-MM")
        Dim intY As Date = Now().ToString("yyyy-MM")
        Dim i As Integer = 0
        'Dim s As String = DDL_Print_Q.SelectedValue
        Dim Q As Int32 = 0

        HF_Date_Today.Value = Now().ToString("yyyy-MM-dd")

        DDL_Print_Category.Items.Clear()
        DDL_Print_Category.Items.Add(New ListItem("General", "1"))
        DDL_Print_Category.Items.Add(New ListItem("MultiPeriod", "10"))
        DDL_Print_Category.Items.Add(New ListItem("Sales", "2"))
        DDL_Print_Category.Items.Add(New ListItem("Purchases", "3"))
        DDL_Print_Category.SelectedValue = "1"


        DDL_Print_Report.Items.Clear()
        DDL_Print_Report.Items.Add(New ListItem("Balance Sheet", "1"))
        DDL_Print_Report.Items.Add(New ListItem("Income Statement", "2"))
        DDL_Print_Report.Items.Add(New ListItem("Summary Trial Balance", "3"))
        DDL_Print_Report.Items.Add(New ListItem("Detailed Trial Balance", "4"))
        DDL_Print_Report.SelectedValue = "2"

        DDL_Print_Denomination.Items.Add(New ListItem("1", "1"))
        DDL_Print_Denomination.Items.Add(New ListItem("10", "10"))
        DDL_Print_Denomination.Items.Add(New ListItem("100", "100"))
        DDL_Print_Denomination.Items.Add(New ListItem("1000", "1000"))

        For i = 0 To 4 ' 5 years
            DDL_Print_Quarter.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next
        For i = 0 To 4 ' 5 years
            DDL_Print_YearFrom.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next
        For i = 0 To 4 ' 5 years
            DDL_Print_YearTo.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next
        For i = 0 To 11
            DDL_Print_P.Items.Add(New ListItem(intMonth.AddMonths(-i).ToString("MMM yyyy"), intMonth.AddMonths(-i).ToString("MM-yyyy")))
        Next


        'Adding the drop down list for Quarter to Quarter

        Dim Fiscal As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        Dim FiscalDate, date1 As String
        Dim d1 As Date

        ' Getting the fiscal year
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Fiscal)

        If Fiscal.Rows(0)("Fiscal_Year_Start_Month") >= 10 Then
            FiscalDate = (Now().ToString("yyyy") - 1) & "-" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate

        Else
            FiscalDate = (Now().ToString("yyyy") - 1) & "-0" & Fiscal.Rows(0)("Fiscal_Year_Start_Month").ToString
            d1 = FiscalDate
        End If

        date1 = FiscalDate
        Dim L_Qua, chk As Integer
        If (Today() >= d1 And Now() <= d1.AddMonths(3).AddDays(-1)) Then
            ' It's in Q1
            L_Qua = 1

        ElseIf (Today() >= d1.AddMonths(3) And Now() <= d1.AddMonths(6).AddDays(-1)) Then
            ' It's in Q2
            L_Qua = 2

        ElseIf (Today() >= d1.AddMonths(6) And Now() <= d1.AddMonths(9).AddDays(-1)) Then
            ' It's in Q3
            L_Qua = 3

        ElseIf (Today() >= d1.AddMonths(9) And Now() <= d1.AddMonths(12).AddDays(-1)) Then
            ' It's in Q4
            L_Qua = 4

        End If
        chk = L_Qua

        Dim Qua_p, Qua_year As String
        For i = 0 To 4 ' 5 Quarter
            If L_Qua = 0 Then
                L_Qua = 4
            End If
            If chk <= i Then
                Qua_p = (intMonth.AddYears(-2)).ToString("yyyy") + " - " + (intMonth.AddYears(-1)).ToString("yy")
                Qua_year = (intMonth.AddYears(-2)).ToString("yyyy")
            Else
                Qua_p = (intMonth.AddYears(-1)).ToString("yyyy") + " - " + intMonth.ToString("yy")
                Qua_year = (intMonth.AddYears(-1)).ToString("yyyy")
            End If
            DDL_Print_Q.Items.Add(New ListItem("Q" & L_Qua & " " & Qua_p, L_Qua & " " & Qua_year))
            L_Qua -= 1
        Next

        ' Period Type
        DDL_Print_Period.Items.Clear()
        DDL_Print_Period.Items.Add(New ListItem("Monthly", "Monthly"))
        DDL_Print_Period.Items.Add(New ListItem("Month-to-Month", "Month-to-Month"))
        DDL_Print_Period.Items.Add(New ListItem("Quarterly", "Quarterly"))
        DDL_Print_Period.Items.Add(New ListItem("Quarter-to-Quarter", "Quarter-to-Quarter"))
        DDL_Print_Period.Items.Add(New ListItem("Yearly", "Yearly"))

        ' Adding Currency
        SQLCommand.CommandText = "Select * from ACC_Currency order by Local desc, Currency_ID"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Cur)

        DDL_Print_Currency.Items.Clear()
        DDL_Print_Currency.DataTextField = "Currency_ID"
        DDL_Print_Currency.DataValueField = "Currency_ID"
        DDL_Print_Currency.DataSource = Cur
        DDL_Print_Currency.DataBind()

        ' Adding Customer
        SQLCommand.CommandText = "Select * from Customer where Customer_Type = 'Customer' order by Name"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Cust)

        For i = 0 To Cust.Rows.Count - 1
            If Cust.Rows(i)("Currency").ToString = "" Then Cust.Rows(i)("Currency") = "CAD"
        Next

        'Date_DT.Text = DateTime.Now.ToString("yyyy-MM-dd")
        'Date_DT.Text = Now().ToString("yyyy-MM-dd")
        RPT_Cust.DataSource = Cust
        RPT_Cust.DataBind()

        ' Adding Vendor
        SQLCommand.CommandText = "Select * from Customer where Customer_Type = 'Vendor' and Name <>'' order by Name"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(Vend)

        For i = 0 To Vend.Rows.Count - 1
            If Vend.Rows(i)("Currency").ToString = "" Then Vend.Rows(i)("Currency") = "CAD"
        Next

        RPT_Vend.DataSource = Vend
        RPT_Vend.DataBind()

        PNL_Report.Visible = True
        ' Return HF_PrintTitle and HF_PrintHeader
        ' PNL_PrintReports.Visible = True
    End Sub

    ' Required to print Report for AR and AP
    Private Sub ShowPNLReport()
        PNL_PrintReports.Visible = True
    End Sub

End Class