
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.IO
Imports System.Xml
Partial Class AjaxPrinting
    Inherits System.Web.UI.Page

    Dim Conn As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("ConnInfo"))
    Dim ConnNav As New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("NAV_Data"))
    Dim SQLCommand As New SqlCommand
    Dim SQLNav As New SqlCommand
    Dim DataAdapter As New SqlDataAdapter
    Dim DataAdapterNav As New SqlDataAdapter
    Dim DBTable As New DataTable

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        SQLNav.Connection = ConnNav
        DataAdapterNav.SelectCommand = SQLNav

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
        If Action = "IncStateMulti" Then PrintIncStateMultiRep()
        If Action = "QuarIncStateMulti" Then PrintQuarIncStateMultiRep()
        If Action = "YearIncStateMulti" Then PrintYearIncStateMultiRep()
        '''''''''
        'If Action = " ThenIncomeStatementSingle" Then PrintIncomeStatementSingle()
        If Action = "Report" Then Report()

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
    Private Sub PrintSalesSummary()

        Dim FromDate As String = Request.Form("date1")
        Dim ToDate As String = Request.Form("date2")
        Dim Currency As String = Request.Form("cur")
        Dim Type As String = Request.Form("type")
        Dim TotalTitle As String = ""

        If Type = "P" Then
            HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc.<br/>Purchase Summary Report<br/>From " & FromDate & " to " & ToDate & "<br/>Currency: " & Currency & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
            TotalTitle = "Total Purchases"
        End If

        If Type = "S" Then
            HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)"
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc.<br/>Sales Summary Report</br>From " & FromDate & " to " & ToDate & "<br/>Currency: " & Currency & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span>"
            TotalTitle = "Total Sales"
        End If

        Dim Sales, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        Dim where As String = " Currency = '" & Currency & "' and"
        If Currency = "CAD" Then where = " (Currency = '' or Currency = '" & Currency & "') and"
        If Currency = "ALL" Then where = ""

        If Type = "P" Then SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency, (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PINV') as Total, (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='PINV') as Tax from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where " & where & " piv.Doc_Date between @date1 and @date2 order by Name"
        If Type = "S" Then SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency, (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SINV') as Total, (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SINV') as Tax, (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SC') as TotalC, (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between @date1 and @date2 and Doc_Type='SC') as TaxC from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where " & where & "  piv.Doc_Date between @date1 and @date2 order by Name"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date1", FromDate)
        SQLCommand.Parameters.AddWithValue("@date2", ToDate)
        DataAdapter.Fill(Sales)

        Sales.Columns.Add("SubTotal", GetType(Decimal))

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

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        For i = 0 To Sales.Rows.Count - 1
            'Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Cust_Vend_ID").ToString, "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("Tax"), "#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("Total"), "#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Cust_Vend_ID").ToString, "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", Sales.Rows(i)("Name").ToString, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px;", Format(Sales.Rows(i)("SubTotal"), "#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", TotalTitle, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold; border-top:solid 1px black; border-bottom: double 3px black", Currency & " " & Format(SubTotal, "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        'Report.Rows.Add("text-align:left; font-size:9pt; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", TotalTitle, "", "", "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(SubTotal, "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(Tax, "$#,###.00"), "text-align:right; font-size:9pt; padding: 3px 5px 3px 5px; font-weight:bold", Format(Total, "$#,###.00"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True


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
    Private Sub PrintIncomeStatementSingle()
        'Print the single income statement
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
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        ' Default date give today's date and a year before
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

        HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; width:350px; font-size:8pt~Account Description~text-align:right; width:120px; font-size:8pt~Dollar Amount~text-align:right; width:160px; font-size:8pt~Sales/Expenses(%)~text-align:centre; width:70px;  font-size:8pt~"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Income Statement<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        ' Getting Total Sales and Other Income (49999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No;"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", DatStart)
        DataAdapter.Fill(COA)

        ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date = @date and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 50000 order by Account_No;"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", DatStart)
        DataAdapter.Fill(COA)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("Dollar_Difference", GetType(Decimal))
        COA.Columns.Add("Percent_Difference", GetType(String))
        COA.Columns.Add("Percent_DifferenceString", GetType(String))
        COA.Columns.Add("DifferenceString", GetType(String))

        'Denomination and rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString) / 5) * 5
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
            Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("DifferenceString") + "</span>", "font-size:8pt; width:50px ;text-align:right ", COA.Rows(i)("Percent_Difference"), "font-size:8pt; width:100px", COA.Rows(i)("fk_Currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
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
    End Sub
    Private Sub PrintSummaryTrail()
        'Print the summary trail sheet
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
        'Set the header
        HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; width:550px; font-size:8pt~Account Name~text-align:right; width:120px; font-size:8pt~Beginning Balance~text-align:right; width:120px; font-size:8pt~Debit~text-align:right; width:120px; font-size:8pt~Credit~text-align:right; width:120px; font-size:8pt~Net actvity~text-align:right; width:120px; font-size:8pt~Closing Balance"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Summary Trial Balance<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'></span><span style='position: absolute; margin-left: 1.7in;'></span><span style='position: absolute; margin-left: 3.3in'></span><span style='position: absolute; margin-left: 4.5in'></span><span style='position: absolute; margin-left: 5.5in'></span><span style='position: absolute; margin-left: 6.8in;'></span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No;"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@enddate", DatStart)
        SQLCommand.Parameters.AddWithValue("@date", DatSecond)
        DataAdapter.Fill(COA)


        'System.diagnostics.Debug.WriteLine(SQLCommand.CommandText + DatSecond.ToString)

        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("NextDateBalanceString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))
        COA.Columns.Add("NetString", GetType(String))

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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString) / 5) * 5
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
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; width: 5in;"
            If COA.Rows(i)("Account_Type") > 90 Then Style = Style & "; font-weight:bold"
            Report.Rows.Add("padding: 0px 0px 0px 0px;border-top:solid 0px; text-align:left; font-size:0pt; width: 25px;", COA.Rows(i)("Account_No").ToString, Style + "width: 1px;border-top:solid 0px; min-width: 1in; max-width: 1in;", COA.Rows(i)("Name").ToString, "padding: 3px 5px 3px 5px;border-top:solid 0px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;", COA.Rows(i)("BalanceString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 1in;max-width: 1in;padding-left: 0.2in;", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; border-top:solid 0px;min-width: 1in;max-width: 1in;", COA.Rows(i)("NetString"), "padding: 3px 5px 3px 5px; text-align:right; border-top:solid 0px;font-size:8pt; padding-left: 0.2in;", COA.Rows(i)("NextDateBalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        Report.Rows.Add("padding: 0px 0px 0px 0px;font-size:1pt;", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "border-top:double 0px", "", "border-top:double 0px", "", "border-top:double 4px", "", "border-top:double 4px", "", "border-top:double 0px", "")
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True
    End Sub
    Private Sub PrintDetailTrial()
        'Print the Detail Trial Sheet
        Dim StartDate As String
        Dim EndDate As String
        Dim accNo As String
        Dim Denom As Int32 = Request.Form("Denom")
        Dim id As String = Request.Form("id")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
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
        SQLCommand.CommandText = "SELECT Transaction_Date, fk_currency_id,Transaction_No, Document_Type, Debit_Amount, Credit_Amount, Balance, Memo, Document_ID, fk_Account_ID,rowID FROM ACC_GL WHERE ((Transaction_Date >= @startDate AND Transaction_Date <= @endDate) AND fk_Account_ID = @id) ORDER BY Transaction_Date asc, rowID desc"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@startDate", StartDate)
        SQLCommand.Parameters.AddWithValue("@endDate", EndDate)
        SQLCommand.Parameters.AddWithValue("@id", id)
        DataAdapter.Fill(COA)
        'If we have a matching name output it to the header
        Try
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Detail Trial Balance<br/>For the Period " & StartDate & " to " & EndDate & " - " + COA.Rows(0)("fk_Currency_ID").ToString + "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><br/><br/><span>" + Request.Form("accName") + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'>Posting Date</span><span style='position: absolute; margin-left: 1in'>Doc No</span><span style='position: absolute; margin-left: 2.5in'>Description</span><span style='position: absolute; margin-left: 4.7in;'>Debit</span><span style='position: absolute; margin-left: 5.8in'>Credit</span><span style='position: absolute; margin-left: 6.7in'>Balance</span></div></div>"
        Catch ex As Exception
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Detail Trial Balance<br/>For the Period " & StartDate & " to " & EndDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><br/><br/><span>" + Request.Form("accNo") + " " + id + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'>Posting Date</span><span style='position: absolute; margin-left: 1in'>Doc No</span><span style='position: absolute; margin-left: 2.5in'>Description</span><span style='position: absolute; margin-left: 4.7in;'>Debit</span><span style='position: absolute; margin-left: 5.8in'>Credit</span><span style='position: absolute; margin-left: 6.7in'>Balance</span></div></div>"
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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString) / 5) * 5
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString) / 5) * 5
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

        HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~ Posting Date~text-align:left; width:350px; font-size:8pt~Doc No~text-align:left; width:120px; font-size:8pt~Description~text-align:right; width:120px; font-size:8pt~Debit~text-align:right; width:120px; font-size:8pt~Cridit~text-align:right; width:120px; font-size:8pt~Balance"


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

        Try
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Detail Trial Balance<br/>For the Period " & StartDate & " to " & EndDate & " - " + COA.Rows(0)("fk_Currency_ID").ToString + "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
        Catch ex As Exception
            'Set the page header, this is below the SQL so we can get the currency
            HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Detail Trial Balance<br/>For the Period " & StartDate & " to " & EndDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><br/><br/><span>" + Request.Form("accNo") + " " + value.ToString() + "</span></span><br><div style='Width: 8.5in; position: absolute; margin-top: -1px;'><span style='position: absolute; margin-left: -0.2in'></span><span style='position: absolute; margin-left: 1in'></span><span style='position: absolute; margin-left: 2.5in'></span><span style='position: absolute; margin-left: 4.7in;'></span><span style='position: absolute; margin-left: 5.8in'></span><span style='position: absolute; margin-left: 6.7in'></span></div></div>"
        End Try


        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("CreditString", GetType(String))
        COA.Columns.Add("DebitString", GetType(String))

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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString) / 5) * 5
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString) / 5) * 5
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

        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        For i = 0 To COA.Rows.Count - 1
            Dim Transaction_Date As Date = COA.Rows(i)("Transaction_Date").ToString()

            Report.Rows.Add("padding: 3px 5px 3px 5px; text-align:left;border-top: solid black 0px; font-size:8pt; min-width: 0.7in;", Transaction_Date.ToString("yyyy-MM-dd"), "text-align:left;border-top: solid black 0px; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 0.7in;", COA.Rows(i)("Transaction_No").ToString, "padding: 3px 5px 3px 5px; border-top: solid black 0px;text-align:left; font-size:8pt; min-width: 2.7in;", COA.Rows(i)("memo"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 0.7in;", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right;font-size:8pt;min-width: 0.7in;", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt;min-width: 0.7in;", COA.Rows(i)("BalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

        Next
        Report.Rows.Add("padding: 0px 0px 0px 0px;font-size:1pt;", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "border-top:double 0px", "", "border-top:double 0px", "", "border-top:double 4px", "", "border-top:double 4px", "", "border-top:double 4px", "")
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True

    End Sub
    'Private Sub PrintDetailTrail()
    '    Dim firstDate As String
    '    Dim seconDate As String
    '    firstDate = Request.Form("FirstDate")
    '    seconDate = Request.Form("SecondDate")
    '    Dim DetailLevel As Integer
    '    DetailLevel = Request.Form("detailLevel")

    '    If firstDate = "" Then firstDate = Now().ToString("yyyy-MM-dd")
    '    If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yyyy-MM-dd")
    '    Dim DatStart, DatSecond As Date
    '    Try
    '        DatStart = firstDate
    '        DatSecond = seconDate
    '    Catch ex As Exception
    '        DatStart = Now()
    '        DatSecond = Now().AddDays(-365)
    '    End Try

    '    HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
    '    HF_PrintTitle.Value = "Axiom Plastics Detail Trail Report From " & DatStart & " to " & DatSecond & "<br/><span style=""font-size:6pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'>ID #</span><span style='position: absolute; margin-left: 0.5in'>Account Name</span><span style='position: absolute; margin-left: 1.5in'>Memo</span><span style='position: absolute; margin-left: 2.5in;'>" & DatStart & "</span><span style='position: absolute; margin-left: 3.5in'>Credit</span><span style='position: absolute; margin-left: 4.5in'>Debit</span><span style='position: absolute; margin-left: 5.5in'>Net Activity</span><span style='position: absolute; margin-left: 6.5in;'>" & DatSecond & "</span></div>"
    '    Dim COA, Bal1, Bal2, Report As New DataTable
    '    PNL_Summary.Visible = True

    '    SQLCommand.Connection = Conn
    '    DataAdapter.SelectCommand = SQLCommand

    '    Conn.Open()

    '    SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No;"
    '    SQLCommand.Parameters.Clear()
    '    SQLCommand.Parameters.AddWithValue("@enddate", DatStart)
    '    SQLCommand.Parameters.AddWithValue("@date", DatSecond)
    '    DataAdapter.Fill(COA)


    '    'System.diagnostics.Debug.WriteLine(SQLCommand.CommandText + DatSecond.ToString)

    '    COA.Columns.Add("BalanceString", GetType(String))
    '    COA.Columns.Add("NextDateBalanceString", GetType(String))
    '    COA.Columns.Add("CreditString", GetType(String))
    '    COA.Columns.Add("DebitString", GetType(String))
    '    COA.Columns.Add("NetString", GetType(String))

    '    For i = 0 To COA.Rows.Count - 1
    '        ' Format all the output for the paper
    '        COA.Rows(i)("BalanceString") = Format(Val(COA.Rows(i)("Balance").ToString), "$#,###.00")
    '        If Left(COA.Rows(i)("BalanceString").ToString, 1) = "-" Then COA.Rows(i)("BalanceString") = "(" & COA.Rows(i)("BalanceString").replace("-", "") & ")"
    '        COA.Rows(i)("NextDateBalanceString") = Format(Val(COA.Rows(i)("NextDateBalance").ToString), "$#,###.00")
    '        If Left(COA.Rows(i)("NextDateBalanceString").ToString, 1) = "-" Then COA.Rows(i)("NextDateBalanceString") = "(" & COA.Rows(i)("NextDateBalanceString").replace("-", "") & ")"
    '        COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit").ToString), "$#,###.00")
    '        If Left(COA.Rows(i)("CreditString").ToString, 1) = "-" Then COA.Rows(i)("CreditString") = "(" & COA.Rows(i)("CreditString").replace("-", "") & ")"
    '        COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit").ToString), "$#,###.00")
    '        If Left(COA.Rows(i)("DebitString").ToString, 1) = "-" Then COA.Rows(i)("DebitString") = "(" & COA.Rows(i)("DebitString").replace("-", "") & ")"
    '        COA.Rows(i)("NetString") = Format(Val(COA.Rows(i)("NetActivity").ToString), "$#,###.00")

    '        If Left(COA.Rows(i)("NetString").ToString, 1) = "-" Then COA.Rows(i)("NetString") = "(" & COA.Rows(i)("NetString").replace("-", "") & ")"
    '        'If Val(COA.Rows(i)("Level").ToString) > 1 Then COA.Rows(i).Delete()

    '        If COA.Rows(i)("BalanceString").ToString = "$.00" Then COA.Rows(i)("BalanceString") = ""
    '        If COA.Rows(i)("NextDateBalanceString").ToString = "$.00" Then COA.Rows(i)("NextDateBalanceString") = ""
    '        If COA.Rows(i)("CreditString").ToString = "$.00" Then COA.Rows(i)("CreditString") = ""
    '        If COA.Rows(i)("DebitString").ToString = "$.00" Then COA.Rows(i)("DebitString") = ""
    '        If COA.Rows(i)("NetString").ToString = "$.00" Then COA.Rows(i)("NetString") = ""
    '    Next
    '    For i As Integer = COA.Rows.Count - 1 To 0 Step -1
    '        ' Delete the rows that arnt above the detail level 
    '        If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
    '            If COA.Rows(i)("BalanceString").ToString = "" And COA.Rows(i)("NextDateBalanceString").ToString = "" And COA.Rows(i)("CreditString").ToString = "" And COA.Rows(i)("DebitString").ToString = "" And COA.Rows(i)("NetString").ToString = "" Then
    '                COA.Rows(i).Delete()
    '            End If
    '        End If
    '    Next i

    '    COA.AcceptChanges()

    '    For i = 1 To 15
    '        Report.Columns.Add("Style" + i.ToString, GetType(String))
    '        Report.Columns.Add("Field" + i.ToString, GetType(String))
    '    Next

    '    Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
    '    For i = 0 To COA.Rows.Count - 1
    '        Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; width: 5in;"
    '        If COA.Rows(i)("Account_Type") > 90 Then Style = Style & "; font-weight:bold"
    '        Report.Rows.Add("padding: 3px 5px 3px 5px; text-align:left; font-size:8pt; width: 25px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("memo2"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("BalanceString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("DebitString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("CreditString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("NetString"), "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt", COA.Rows(i)("NextDateBalanceString"), "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    '    Next

    '    RPT_PrintReports.DataSource = Report
    '    RPT_PrintReports.DataBind()

    '    Conn.Close()

    '    PNL_PrintReports.Visible = True
    'End Sub

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


        'Dim COA, Bal, Report As New DataTable
        'PNL_Summary.Visible = True

        'SQLCommand.Connection = Conn
        'DataAdapter.SelectCommand = SQLCommand

        'Conn.Open()

        'SQLCommand.CommandText = "SELECT * FROM ACC_GL LEFT JOIN ACC_GL_Accounts on Acc_Gl.fk_Account_ID = ACC_GL_Accounts.Account_ID WHERE (fk_Account_ID = @Account_number) AND (locked = 0) AND Transaction_Date <= @date ORDER BY " + sort_param + updown
        'SQLCommand.Parameters.Clear()
        'SQLCommand.Parameters.AddWithValue("@Account_number", Account_number1)
        'SQLCommand.Parameters.AddWithValue("@date", date1)
        'System.Diagnostics.Debug.WriteLine(SQLCommand.CommandText)
        'DataAdapter.Fill(COA)

        'COA.AcceptChanges()
        'COA.Columns.Add("CreditString", GetType(String))
        'COA.Columns.Add("DebitString", GetType(String))

        'For i = 1 To 15
        '    Report.Columns.Add("Style" + i.ToString, GetType(String))
        '    Report.Columns.Add("Field" + i.ToString, GetType(String))
        'Next
        'Dim smallStyle, medStyle, largeStyle As String
        'For i = 0 To COA.Rows.Count - 1

        '    If (COA.Rows(i)("Credit_Amount").ToString = "0.00") Then
        '        COA.Rows(i)("CreditString") = ""
        '    Else
        '        COA.Rows(i)("CreditString") = Format(Val(COA.Rows(i)("Credit_Amount").ToString), "$#,###.00")

        '    End If

        '    If (COA.Rows(i)("Debit_Amount").ToString = "0.00") Then
        '        COA.Rows(i)("DebitString") = ""
        '    Else
        '        COA.Rows(i)("DebitString") = Format(Val(COA.Rows(i)("Debit_Amount").ToString), "$#,###.00")
        '    End If

        '    smallStyle = "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; min-width: 1in; max-width: 1in;"
        '    medStyle = "text-align:left; font-size:9pt; padding: 3px 5px 3px 5px; min-width: 1.5in; max-width: 1.5in;"
        '    largeStyle = "text-align:left; font-size:9pt; padding: 3px 5px 3px; min-width: 2in; max-width: 2in; "
        '    Dim Transaction_Date As Date = COA.Rows(i)("Transaction_Date").ToString()

        '    Report.Rows.Add(medStyle, Transaction_Date.ToString("yyyy-MM-dd"), largeStyle, COA.Rows(i)("Memo").ToString, largeStyle, COA.Rows(i)("CreditString"), largeStyle, COA.Rows(i)("DebitString"), smallStyle, COA.Rows(i)("fk_currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        'Next

    End Sub

    ' Balance Sheet
    Private Sub PrintBalance()

        Dim AsAt As String = Request.Form("date1")
        Dim StyleFinish As String
        Dim DetailLevel As Integer = Request.Form("detailLevel")
        Dim NoZeros As String = Request.Form("showZeros")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        If AsAt = "" Then AsAt = Now().ToString("yyyy-MM-dd")

        If DetailLevel = 0 Then DetailLevel = 7

        HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; font-size:10pt~Account Name~text-align:right; width:100px; font-size:10pt~Balance~text-align:left; width:60px; font-size:8pt~"

        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Balance Sheet<br/>As Of " & AsAt & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span>"

        Dim COA, Bal, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Totalling_Minus From ACC_GL_Accounts order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, (Select Top 1 Balance from ACC_GL where Transaction_Date <= @date and fk_Account_Id = Account_ID order by Transaction_Date desc, rowID desc) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 10000 and Account_No<'40000' order by Account_No"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", AsAt)
        DataAdapter.Fill(Bal)

        COA.Columns.Add("Balance", GetType(String))
        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))

        Dim Padding As Integer = 0
        Dim Level As Integer = 1
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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("BeforeBalance") = Math.Round(Val(COA.Rows(i)("BeforeBalance").ToString) / 5) * 5
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
            Report.Rows.Add(Style1, COA.Rows(i)("Account_No").ToString, Style2, COA.Rows(i)("Name").ToString, Style3, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Balance") + "</span>", Style4, COA.Rows(i)("fk_currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next

        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True


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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("BeforeBalance") = Math.Round(Val(COA.Rows(i)("BeforeBalance").ToString) / 5) * 5
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

        HF_PrintHeader.Value = "text-align:left; width:100px; font-size:8pt~~text-align:left; width:350px; font-size:8pt~~text-align:right; width:120px; font-size:8pt~"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Summary Trial Balance<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 0'></span><span style='position: absolute; margin-left: 0.5in'>Account Name</span><span style='position: absolute; margin-left: 1.7in;'>Beginning Balance</span><span style='position: absolute; margin-left: 3.3in'>Debit</span><span style='position: absolute; margin-left: 4.5in'>Credit</span><span style='position: absolute; margin-left: 5.5in'>Net actvity</span><span style='position: absolute; margin-left: 6.8in;'>Closing Balance</span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Totalling, Active, Cash, COALESCE(ThisDateBalance,0.00) AS Balance, Transaction_No,COALESCE(NextDateBalance,0.00) AS NextDateBalance, Memo,memo2,ISNULL(creditSum,0) as Credit,ISNULL(debitSum,0) as Debit, ISNULL((creditSum - debitSum),0) as NetActivity From ACC_GL_Accounts outer apply(select top 1 * from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance,Memo as memo2 from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select sum(Credit_Amount) as creditSum, sum(Debit_Amount) as debitSum from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @endDate and @date)  as Summary outer apply(select top 1 (Balance) as NextDateBalance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <= @endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type != 99 and Account_Type != 98 order by Account_No;"
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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString) / 5) * 5
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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("Credit_Amount") = Math.Round(Val(COA.Rows(i)("Credit_Amount").ToString) / 5) * 5
                    COA.Rows(i)("Debit_Amount") = Math.Round(Val(COA.Rows(i)("Debit_Amount").ToString) / 5) * 5
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
            sqlInsert = sqlInsert + "outer apply(select top 1 (Balance) as Month" & i & "Balance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <=@date" & i & " order by Transaction_Date desc, rowID desc)  as Month" & i & " "
            sqlInsertHeaders = sqlInsertHeaders + ", CONVERT(varchar(100), Month" & i & "Balance) as " + monthArray(Month(tempDate) - 1) + ""

        Next

        Dim COA, Bal1, Bal2, Report As New DataTable

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()
        SQLCommand.CommandTimeout = 500
        SQLCommand.CommandText = "Select Totalling,Totalling_Minus,Account_Type,Account_No, Name, Totalling_Minus " + sqlInsertHeaders + " From [AXIOMGROUP].[dbo].[ACC_GL_Accounts] outer apply( select top 1 * from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @dateStart AND @dateEnd order by Transaction_Date desc, rowID desc) as tid " + sqlInsert + "WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 order by Account_No"
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
                    COA.Rows(a)(monthArray(Month(tempDate) - 1)) = Math.Round(Val(COA.Rows(a)(monthArray(Month(tempDate) - 1)).ToString) / 5) * 5
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

        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ThisDateBalance AS Balance, Totalling_Minus, Exchange_Account_ID, Transaction_No,NextDateBalance From [AXIOMGROUP].[dbo].[ACC_GL_Accounts] outer apply(select top 1 * from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date BETWEEN @date AND @endDate order by Transaction_Date desc, rowID desc) as tid outer apply(select top 1 (Balance) as ThisDateBalance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <=@date order by Transaction_Date desc, rowID desc )  as ThisDateTotal outer apply(select top 1 (Balance) as NextDateBalance from [AXIOMGROUP].[dbo].[ACC_GL] where fk_Account_ID=Account_ID and Transaction_Date <=@endDate order by Transaction_Date desc, rowID desc)  as NextDateTotal WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 order by Account_No;"
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
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString) / 5) * 5
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
        'DataTable.Columns["Marks"].ColumnName = "SubjectMarks";
        'DataTable.Columns["Marks"].ColumnName = "SubjectMarks";
        '    DataTable.Columns["Marks"].ColumnName = "SubjectMarks";

        Dim ds As New DataSet
        ds.Tables.Add(COA)


        Dim xmlData As String = ds.GetXml()

        HF_XML.Value = xmlData
        PNL_XMLReport.Visible = True


        Conn.Close()

    End Sub
    Private Sub PrintProfitLoss()

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
        Dim DetailLevel As Integer
        DetailLevel = Request.Form("detailLevel")
        Dim Denom As Int32 = Request.Form("Denom")
        Dim DenomString As String = ""
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        ' Default date give today's date and a year before
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

        HF_PrintHeader.Value = "text-align:left; width:0px; font-size:0pt~~text-align:left; width:350px; font-size:8pt~Account Description~text-align:right; width:120px; font-size:8pt~Dollar Amount~text-align:right; width:160px; font-size:8pt~Sales/Expenses(%)~text-align:centre; width:70px;  font-size:8pt~"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Income Statement<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand

        Conn.Open()

        ' Getting Total Sales and Other Income (49999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", DatStart)
        SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
        DataAdapter.Fill(COA)

        ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash, ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date between @date and @enddate and fk_Account_Id = Account_ID)) as Balance From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 50000 order by Account_No"
        SQLCommand.Parameters.Clear()
        SQLCommand.Parameters.AddWithValue("@date", DatStart)
        SQLCommand.Parameters.AddWithValue("@enddate", DatSecond)
        DataAdapter.Fill(COA)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString", GetType(String))
        COA.Columns.Add("Dollar_Difference", GetType(Decimal))
        COA.Columns.Add("Percent_Difference", GetType(String))
        COA.Columns.Add("Percent_DifferenceString", GetType(String))
        COA.Columns.Add("DifferenceString", GetType(String))

        'Denomination And rounding
        If Denom > 1 Or Request.Form("Round") = "on" Then
            For i = 0 To COA.Rows.Count - 1
                If Request.Form("Round") = "on" Then
                    COA.Rows(i)("Balance") = Math.Round(Val(COA.Rows(i)("Balance").ToString) / 5) * 5
                    'COA.Rows(i)("NextDateBalance") = Math.Round(Val(COA.Rows(i)("NextDateBalance").ToString) / 5) * 5
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
            Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("DifferenceString") + "</span>", "font-size:8pt; width:50px ;text-align:right ", COA.Rows(i)("Percent_Difference"), "font-size:8pt; width:100px", COA.Rows(i)("fk_Currency_id"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
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

    End Sub
    Private Sub PrintIncStateMultiRep()

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim MonthCount As Integer = 3
        Dim Padding As Integer = 0
        Dim j As Integer = 0
        Dim Level As Integer = 1
        Dim firstDate As String
        Dim seconDate As String
        Dim startDate As Date
        Dim startDate1 As String
        Dim startDate2 As Date
        Dim startDate11 As String
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
        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
        End If

        ' Default date give today's date and a year before
        If firstDate = "" Then firstDate = Now().ToString("yy-MM")
        If seconDate = "" Then seconDate = Now().AddDays(-365).ToString("yy-MM")
        Dim DatStart, DatSecond As Date
        Try
            DatStart = firstDate
            DatSecond = seconDate
        Catch ex As Exception
            DatStart = Now()
            DatSecond = Now().AddDays(-365)
        End Try


        startDate1 = firstDate
        startDate = firstDate
        Dim StyleMonth As String

        While (startDate <= seconDate)
            StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + startDate.ToString("MMMM")
            startDate = startDate.AddMonths(1)
            startDate1 = startDate.ToString("yyyy-MM")
        End While
        HF_PrintHeader.Value = "text-align:left; width:10px; font-size:8pt~Account No~text-align:left; width:5px; font-size:8pt~Account Description" + StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~Total"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Multiperiod Income Statement(Monthly)<br/>From " & firstDate & " to " & seconDate & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        startDate1 = firstDate
        startDate = firstDate

        While (startDate <= seconDate)

            startDate1 = "'" + startDate1 + "%'"
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date LIKE " & startDate1 & " and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date LIKE " & startDate1 & " and fk_Account_Id = Account_ID)) as Balance" & j.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date LIKE " & startDate1 & " and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date LIKE " & startDate1 & " and fk_Account_Id = Account_ID)) as Balance" & j.ToString
            j += 1
            startDate = startDate.AddMonths(1)
            startDate1 = startDate.ToString("yyyy-MM")
        End While

        ' Getting Total Sales and Other Income (49999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 50000 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        COA.Columns.Add("Padding", GetType(Integer))
        COA.Columns.Add("Level", GetType(Integer))
        COA.Columns.Add("BalanceString0", GetType(String))
        COA.Columns.Add("BalanceString1", GetType(String))
        COA.Columns.Add("BalanceString2", GetType(String))

        startDate1 = firstDate
        startDate = firstDate
        j = 0
        Dim Balance As String
        Dim BalanceString As String = ""
        Dim ColMonth As String = ""


        Dim k As Int32 = 0
        Balance = ""
        BalanceString = ""
        While (startDate <= seconDate)
            Balance = "Balance" + j.ToString
            BalanceString = "BalanceString" + j.ToString
            'Denomination And rounding
            If Denom > 1 Or Request.Form("Round") = "on" Then
                For i = 0 To COA.Rows.Count - 1
                    If Request.Form("Round") = "on" Then
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString) / 5) * 5
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

            ' Get the value for Total Income, Total Cost, and Total Expenses
            Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
            If rowIncome.Length > 0 Then
                TotalIncome = rowIncome(0).Item(Balance)
            End If
            Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
            If rowCost.Length > 0 Then
                TotalCost = rowCost(0).Item(Balance)
            End If
            Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")
            If rowExpense.Length > 0 Then
                TotalExpenses = rowExpense(0).Item(Balance)
            End If

            ' Format all the output for the paper
            For i = 0 To COA.Rows.Count - 1
                COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                If Left(COA.Rows(i)(BalanceString).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"

                If Request.Form("Round") = "on" Then
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                Else
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                End If

                If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Then COA.Rows(i)(BalanceString) = ""
            Next
            For i As Integer = COA.Rows.Count - 1 To 0 Step -1
                Dim AlreadyDeleted As Boolean = False

                ' Delete the rows that arnt above the detail level 
                If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
                    If COA.Rows(i)(BalanceString).ToString = "" Then
                        COA.Rows(i).Delete()
                        AlreadyDeleted = True
                        'ElseIf COA.Rows(i)("DifferenceString").ToString = "" Then
                        '    COA.Rows(i).Delete()
                        '    AlreadyDeleted = True
                    End If

                End If
                If (AlreadyDeleted = False) Then
                    If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

                End If

            Next i

            COA.AcceptChanges()
            j += 1
            startDate = startDate.AddMonths(1)
            startDate1 = startDate.ToString("yyyy-MM")
        End While
        For i = 1 To 15
            Report.Columns.Add("Style" + i.ToString, GetType(String))
            Report.Columns.Add("Field" + i.ToString, GetType(String))
        Next

        startDate11 = firstDate
        startDate2 = firstDate
        Dim Style As String = "text-align:left; font-size:8pt; padding: 3px 5px 3px; "
        Dim Style2 As String = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;"
        For i = 0 To COA.Rows.Count - 1
            Style = "text-align:left; font-size:8pt; padding: 3px 5px 3px " & Val(COA.Rows(i)("Padding").ToString) + 5 & "px; min-width: 2.5in; max-width: 2.5in;"
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

            k = 0
            'While (startDate11 <= seconDate)
            '    Balance1 = "'" & "Balance" & k.ToString & "'"
            '    ColMonth = ColMonth + ", Style2, COA.Rows(" & i.ToString & ")(" & Balance1 & ")"
            '    k += 1
            '    startDate2 = startDate2.AddMonths(1)
            '    startDate11 = startDate2.ToString("yyyy-MM")
            'End While

            If j = 1 Then
                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, COA.Rows(i)("BalanceString1"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, COA.Rows(i)("BalanceString1"), Style2, COA.Rows(i)("BalanceString2"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            End If

        Next

        BalanceString = "BalanceString0"
        ColMonth = "Style2, COA.Rows(i)(" & BalanceString & ")"
        For i = 0 To COA.Rows.Count - 1
            BalanceString = "BalanceString0"
            ColMonth = COA.Rows(i)(BalanceString) + ", Style2," + COA.Rows(i)("BalanceString1")
            'Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, COA.Rows(i)("BalanceString1"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")

            Report.Rows.Add(" text-align:left; font-size:0pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, ColMonth, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        Next
        RPT_PrintReports.DataSource = Report
        RPT_PrintReports.DataBind()

        Conn.Close()

        PNL_PrintReports.Visible = True


    End Sub

    Private Sub PrintQuarIncStateMultiRep()

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

        Dim seconDate As String
        Dim startDate As String


        Dim StyleFinish As String = ""
        Dim TotalIncome As String = "0"
        Dim TotalCost As String = "0"
        Dim TotalExpenses As String = "0"
        Dim ProfitAndLoss As String
        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""

        Year = Request.Form("YearForQuater")

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
            DenomString = "Denomination x" + Denom.ToString()
        End If

        Dim StyleMonth As String
        Dim Quarter(4) As String

        If (Qua_1 = "on") Then
            Quarter(0) = "Q-1"
            Qua_1_StartDate = Year - 1 & "-09-01"
            Qua_1_EndDate = Year - 1 & "-11-30"
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_1_StartDate & "' and '" & Qua_1_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_1_EndDate
            startDate = Qua_1_StartDate
            Q += 1
        End If
        If (Qua_2 = "on") Then
            Quarter(1) = "Q-2"
            Qua_2_StartDate = Year - 1 & "-12-01"
            Qua_2_EndDate = Year & "-02-28"
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_2_StartDate & "' and '" & Qua_2_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_2_EndDate
            If Q = 0 Then
                startDate = Qua_2_StartDate
            End If

            Q += 1
        End If
        If (Qua_3 = "on") Then
            Quarter(2) = "Q-3"
            Qua_3_StartDate = Year & "-03-01"
            Qua_3_EndDate = Year & "-05-31"
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_3_StartDate & "' and '" & Qua_3_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_3_EndDate
            If Q = 0 Then
                startDate = Qua_3_StartDate
            End If
            Q += 1
        End If
        If (Qua_4 = "on") Then
            Quarter(3) = "Q-4"
            Qua_4_StartDate = Year & "-06-01"
            Qua_4_EndDate = Year & "-08-31"
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & Qua_4_StartDate & "' and '" & Qua_4_EndDate & "' and fk_Account_Id = Account_ID)) as Balance" & Q.ToString
            seconDate = Qua_4_EndDate
            If Q = 0 Then
                startDate = Qua_4_StartDate
            End If
            Q += 1
        End If

        Dim H_Quarter As String
        For l = 0 To 3
            If Quarter(l) <> "" Then
                H_Quarter = Quarter(l)
                StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + H_Quarter
            End If

            'startDate1 = startDate.ToString("yyyy-MM")
        Next
        HF_PrintHeader.Value = "text-align:left; width:50px; font-size:8pt~A/C No~text-align:left; width:5px; font-size:8pt~Account Description" + StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~Total"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Multiperiod Income Statement(Quarterly)<br/>From " + startDate + "  to " + seconDate + " <br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"
        Dim COA, Bal1, Bal2, Report As New DataTable
        PNL_Summary.Visible = True

        SQLCommand.Connection = Conn
        DataAdapter.SelectCommand = SQLCommand
        Conn.Open()

        ' Getting Total Sales and Other Income (49999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 50000 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

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
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString) / 5) * 5
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

            ' Get the value for Total Income, Total Cost, and Total Expenses
            'Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
            'If rowIncome.Length > 0 Then
            '    TotalIncome = rowIncome(0).Item(Balance)
            'End If
            'Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
            'If rowCost.Length > 0 Then
            '    TotalCost = rowCost(0).Item(Balance)
            'End If
            'Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")
            'If rowExpense.Length > 0 Then
            '    TotalExpenses = rowExpense(0).Item(Balance)
            'End If

            ' Format all the output for the paper
            For i = 0 To COA.Rows.Count - 1
                COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                If Request.Form("Round") = "on" Then
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                Else
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                End If

                If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Then COA.Rows(i)(BalanceString) = ""
                If Left(COA.Rows(i)(BalanceString).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
            Next

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

            COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")

            If Left(COA.Rows(i)("Total").ToString, 1) = "-" Then COA.Rows(i)("Total") = "(" & COA.Rows(i)("Total").replace("-", "") & ")"

            'If Request.Form("Round") = "on" Then
            '    COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###")
            'Else
            '    COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            'End If

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

            If Q = 1 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 2 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf Q = 3 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.00")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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


    End Sub

    Private Sub PrintYearIncStateMultiRep()

        Dim Query1 As String = ""
        Dim Query2 As String = ""
        Dim MonthCount As Integer = 3
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


        Dim Profitloss0 As String = ""
        Dim Profitloss1 As String = ""
        Dim Profitloss2 As String = ""
        Dim TotalProfitloss As String = ""

        If (Denom > 1) Then
            DenomString = "Denomination x" + Denom.ToString()
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
        Dim date1, date2, date3 As String
        Dim d1, d2, d3, dtemp As Date
        Dim YearCount As Integer = seconDate - (firstDate - 1)

        ' Get the fiscal month
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'"
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


        'd2 = firstDate
        ' if date picked is current year, check if today's month is > fiscal month
        'If seconDate >= DateTime.Now.Year Then
        '    ' Check if today's date already pass the fiscal month
        '    If DateTime.Now.Month <= Fiscal.Rows(0)("Fiscal_Year_Start_Month") Then
        '        ' use today's date to compare with previous year
        '        date1 = Now().ToString("yyyy-MM-dd")
        '        d1 = date1
        '        date2 = d1.AddDays(-1).AddYears(-1).ToString("yyyy-MM-dd")
        '        dtemp = d1
        '    Else
        '        d1 = date1
        '        date2 = d1.AddDays(-1).AddYears(-1).ToString("yyyy-MM-dd")
        '        d1 = d1.AddDays(-1)
        '        dtemp = d1
        '    End If
        'Else
        '    d1 = date1
        '    date2 = d1.AddDays(-1).AddYears(-1).ToString("yyyy-MM-dd")
        '    d1 = d1.AddDays(-1)
        '    dtemp = d1
        'End If

        'endDate = date1
        'If YearCount = 1 Then
        '    d2 = d1.AddYears(-1)
        '    dtemp = d2

        '    date2 = d2.ToString("yyyy-MM-dd")
        'ElseIf YearCount = 2 Then
        '    d2 = d1.AddYears(-1)
        '    d3 = d2.AddYears(-1)
        '    dtemp = d3

        '    date2 = d2.ToString("yyyy-MM-dd")
        '    date3 = d3.ToString("yyyy-MM-dd")
        'End If

        'startDate = dtemp.ToString("yyyy-MM-dd")





        Dim seconDate1 = seconDate
        startDate1 = FiscalDate
        startDate = startDate1
        endDate = FiscalDateEnd
        endDate1 = endDate
        seconDate = date2
        endDate2 = endDate1.Year
        Dim StyleMonth As String

        While (endDate <= seconDate)
            If endDate1.Year >= DateTime.Now.Year Then
                endDate2 = endDate1.ToString("yyyy") + "(*)"
            Else
                endDate2 = endDate1.ToString("yyyy")
            End If
            StyleMonth = StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~" + startDate.ToString("yyyy") + "-" + endDate2
            startDate = startDate.AddYears(1)
            startDate1 = startDate
            endDate1 = endDate1.AddYears(1)
            endDate = endDate1.ToString("yyyy/MM/dd")

        End While
        HF_PrintHeader.Value = "text-align:left; width:10px; font-size:8pt~Account No~text-align:left; width:5px; font-size:8pt~Account Description" + StyleMonth + "~Text-align: Right; width:120px; font-size:8pt~Total"
        HF_PrintTitle.Value = "<span style=""font-size:11pt"">Axiom Plastics Inc<br/>Multiperiod Income Statement(Yearly)<br/>From " & (firstDate - 1).ToString + "-" + firstDate & " to " & (seconDate1 - 1).ToString + "-" + seconDate1 & "<br/></span><span style=""font-size:7pt"">Printed on " & Now().ToString("yyyy-MM-dd hh:mm tt") & " " + DenomString + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>"

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
            Query1 = Query1 & ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID) - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID)) as Balance" & j.ToString
            Query2 = Query2 & ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID) - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" & startDate & "' and '" & endDate1 & "' and fk_Account_Id = Account_ID)) as Balance" & j.ToString
            j += 1
            Q += 1
            startDate = startDate.AddYears(1).ToString("yyyy/MM/dd")
            startDate1 = startDate
            endDate1 = endDate1.AddYears(1).ToString("yyyy/MM/dd")
            endDate = endDate1
        End While

        ' Getting Total Sales and Other Income (49999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query1 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 40000 and Account_No<'50000' order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

        ' Getting Total Direct Cost of Goods Sold (59999) and Total General & Administration Expenses (69999)
        SQLCommand.CommandText = "Select Account_ID, Account_No, Name, ACC_GL_Accounts.fk_Currency_ID, Account_Type, Direct_Posting, fk_Linked_ID, Totalling, Active, Cash " & Query2 & " From ACC_GL_Accounts WHERE Account_Type >=  0 and Account_ID > 1 and Account_No >= 50000 order by Account_No"
        SQLCommand.Parameters.Clear()
        DataAdapter.Fill(COA)

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
                        COA.Rows(i)(Balance) = Math.Round(Val(COA.Rows(i)(Balance).ToString) / 5) * 5
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

            ' Get the value for Total Income, Total Cost, and Total Expenses
            'Dim rowIncome() As DataRow = COA.Select("Account_No = '49999'")
            'If rowIncome.Length > 0 Then
            '    TotalIncome = rowIncome(0).Item(Balance)
            'End If
            'Dim rowCost() As DataRow = COA.Select("Account_No = '59999'")
            'If rowCost.Length > 0 Then
            '    TotalCost = rowCost(0).Item(Balance)
            'End If
            'Dim rowExpense() As DataRow = COA.Select("Account_No = '69999'")
            'If rowExpense.Length > 0 Then
            '    TotalExpenses = rowExpense(0).Item(Balance)
            'End If

            ' Format all the output for the paper
            For i = 0 To COA.Rows.Count - 1
                COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")

                If Request.Form("Round") = "on" Then
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###")
                Else
                    COA.Rows(i)(BalanceString) = Format(Val(COA.Rows(i)(Balance).ToString), "$#,###.00")
                End If

                If COA.Rows(i)(BalanceString).ToString = "$.00" Or COA.Rows(i)(BalanceString).ToString = "$" Then COA.Rows(i)(BalanceString) = ""
                If Left(COA.Rows(i)(Balance).ToString, 1) = "-" Then COA.Rows(i)(BalanceString) = "(" & COA.Rows(i)(BalanceString).replace("-", "") & ")"
            Next
            COA.AcceptChanges()
            'For i As Integer = COA.Rows.Count - 1 To 0 Step -1
            '    Dim AlreadyDeleted As Boolean = False

            '    ' Delete the rows that arnt above the detail level 
            '    If Request.Item("showZeros") = "off" And COA.Rows(i)("Account_Type") < 90 Then
            '        If COA.Rows(i)(BalanceString).ToString = "" Then
            '            COA.Rows(i).Delete()
            '            AlreadyDeleted = True
            '            'ElseIf COA.Rows(i)("DifferenceString").ToString = "" Then
            '            '    COA.Rows(i).Delete()
            '            '    AlreadyDeleted = True
            '        End If

            '    End If
            '    If (AlreadyDeleted = False) Then
            '        If COA.Rows(i)("Level") > DetailLevel Then COA.Rows(i).Delete()

            '    End If

            'Next

            'COA.AcceptChanges()
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

            COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")

            If Left(COA.Rows(i)("Total").ToString, 1) = "-" Then COA.Rows(i)("Total") = "(" & COA.Rows(i)("Total").replace("-", "") & ")"

            'If Request.Form("Round") = "on" Then
            '    COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###")
            'Else
            '    COA.Rows(i)("Total") = Format(Val(COA.Rows(i)("Total").ToString), "$#,###.00")
            'End If

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

            If j = 1 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 2 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
            ElseIf j = 3 Then
                Report.Rows.Add(" text-align:left; font-size:8pt; width: 10px;", COA.Rows(i)("Account_No").ToString, Style, COA.Rows(i)("Name").ToString, Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString0") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString1") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("BalanceString2") + "</span>", Style2, "<span style=""" + StyleFinish + """>" + COA.Rows(i)("Total") + "</span>", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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

                Profitloss0 = Format(Val(Profitloss0.ToString), "$#,###.00")
                Profitloss1 = Format(Val(Profitloss1.ToString), "$#,###.00")
                Profitloss2 = Format(Val(Profitloss2.ToString), "$#,###.00")
                TotalProfitloss = Format(Val(TotalProfitloss.ToString), "$#,###.00")

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


    End Sub

    Private Sub Report()
        Dim Cur, Cust, Vend As New DataTable
        Dim intYear As Integer = DateTime.Now.Year
        Dim i As Integer = 0

        HF_Date_Today.Value = Now().ToString("yyyy-MM-dd")

        DDL_Print_Category.Items.Clear()
        DDL_Print_Category.Items.Add(New ListItem("General", "1"))
        DDL_Print_Category.Items.Add(New ListItem("General-MultiPeriod", "10"))
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

        'Add Year for multiperiod Quaterly
        For i = 0 To 4 ' 5 years
            DDL_Print_Quarter.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next

        'Add Year for multiperiod Yearly
        For i = 0 To 4 ' 5 years
            DDL_Print_YearFrom.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next
        For i = 0 To 4 ' 5 years
            DDL_Print_YearTo.Items.Add(New ListItem((intYear - i - 1).ToString() + " - " + (intYear - i).ToString(), (intYear - i).ToString()))
        Next


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
End Class