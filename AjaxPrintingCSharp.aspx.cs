using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Globalization;

public partial class AjaxPrintingCSharp :  System.Web.UI.Page
{
    SqlConnection Conn = new SqlConnection(System.Configuration.ConfigurationManager.AppSettings["ConnInfo"]);
    SqlCommand SQLCommand = new SqlCommand();
    SqlDataAdapter DataAdapter = new SqlDataAdapter();
    protected void Page_Load(object sender, EventArgs e)
    {
        string Action = Request.Form["action"];
        DataAdapter.SelectCommand = SQLCommand;
        SQLCommand.Connection = Conn;

        if (Action == "SalesMultiMonthly")
        {
            PrintMonthlySalesMultiPer();
        }
        else if (Action == "SalesMultiMonth-to-Month")
        {
            PrintMonthToMonthSalesMultiPer();
        }
        else if (Action == "SalesMultiQuarterly")
        {
            PrintQuarterlySalesMultiPer();
        }
        else if (Action == "SalesMultiQuarter-to-Quarter")
        {
            PrintQuarterToQuarterSalesMultiPer();
        }
        else if (Action == "SalesMultiYearly")
        {
            PrintYearlySalesMultiPer();
        }
        else if (Action == "PurchasesMultiMonthly")
        {
            PrintMonthlyPurchasesMultiPer();
        }
        else if (Action == "PurchasesMultiMonth-to-Month")
        {
            PrintMonthToMonthPurchasesMultiPer();
        }
        else if (Action == "PurchasesMultiQuarterly")
        {
            PrintQuarterlyPurchasesMultiPer();
        }
        else if (Action == "PurchasesMultiQuarter-to-Quarter")
        {
            PrintQuarterToQuarterPurchasesMultiPer();
        }
        else if (Action == "PurchasesMultiYearly")
        {
            PrintYearlyPurchasesMultiPer();
        }
        else if (Action == "exportAR")
        {
            ExportAR();
        }
        else if (Action == "exportAP")
        {
            ExportAP();
        }
    }

    // Sales Multiperiod Monthly
    private void PrintMonthlySalesMultiPer()
    {
        int Language = Convert.ToInt32(Request.Form["language"]);
        string firstDate = Request.Form["FirstDate"];
        string seconDate = Request.Form["SecondDate"];
        Int32 Denom = Convert.ToInt32(Request.Form["Denom"]);
        string Query = "";
        System.DateTime startDate = default(System.DateTime);
        System.DateTime endDate = default(System.DateTime);
        string startDate1 = "";
        string endDate1 = "";
        string StyleMonth = "";
        string Asterix = "";
        string DenomString = "";
        int MonthCount = 0;
        int i = 0;
        int ii = 0;
        double percentage = 0;
        string percent = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        String QueryDate = "";

        DataTable fiscal = new DataTable();

        // Get the MonthCount Value

        try
        {
            startDate = DateTime.Parse(firstDate);
            endDate = DateTime.Parse(seconDate);

            while (startDate != endDate)
            {
                startDate = startDate.AddMonths(1);
                MonthCount += 1;
            }
            //MonthCount = Convert.ToInt32(Request.Form["SecondDate"].Substring(5, 2)) - Convert.ToInt32(Request.Form["FirstDate"].Substring(5, 2));
        }
        catch (Exception ex)
        {
            MonthCount = 0;
        }

        double[] Total = new double[MonthCount + 1];

        if ((Denom > 1))
        {
            if (Language == 0)
            {
                if (Denom == 10)
                {
                    DenomString = "(In Tenth)";
                }
                else if (Denom == 100)
                {
                    DenomString = "(In Hundreds)";
                }
                else if (Denom == 1000)
                {
                    DenomString = "(In Thousands)";
                }
            }
            else if (Language == 1)
            {
                if (Denom == 10)
                {
                    DenomString = "(En D�cimo)";
                }
                else if (Denom == 100)
                {
                    DenomString = "(En Centenares)";
                }
                else if (Denom == 1000)
                {
                    DenomString = "(En Miles)";
                }
            }
        }

        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        Asterix = "";
        
        if (Request.Form["SecondDate"] == DateTime.Now.ToString("yyyy-MM"))
        {
            seconDate = DateTime.Now.ToString("yyyy-MM-dd");
            firstDate = DateTime.Now.AddMonths(-MonthCount).ToString("yyyy-MM-01");

            endDate = DateTime.Now.AddMonths(-MonthCount);
            endDate1 = DateTime.Now.AddMonths(-MonthCount).ToString("yyyy-MM-dd");
            Asterix = "(*)";

        }
        else
        {
            // Default date give today's date
            if (string.IsNullOrEmpty(firstDate))
            {
                firstDate = DateTime.Now.ToString("yyyy-MM-01");
                Asterix = "(*)";
            }
            else
            {
                // If exist, take the the first day of month
                startDate = DateTime.Parse(firstDate);
            }
            if (string.IsNullOrEmpty(seconDate))
            {
                seconDate = DateTime.Now.ToString("yyyy-MM-dd");
                endDate = DateTime.Now;
                Asterix = "(*)";
            }
            else
            {
                // If exist, take the the last day of month
                endDate = DateTime.Parse(seconDate);
                endDate1 = startDate.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            }
        }

        startDate = DateTime.Parse(firstDate);
        startDate1 = startDate.ToString("yyyy-MM-dd");

        for (i = 0; i <= MonthCount; i++)
        {
            if (Language == 0)
            {
                StyleMonth = StyleMonth + "~Text-align: right; font-size:8pt~" + startDate.AddMonths(i).ToString("MMMM") + Asterix;
            }

        }

        // Percentage.
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
            percent = "~Text-align:right; width:80px; font-size:8pt~Percentage(%)";
        else
            percent = "~Text-align:right; width:0px; font-size:0pt~";

        if (Language == 0)
        {
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt; width:60px;~ID~text-align:left; width:50px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Monthly) " + DenomString + "<br/>From " + firstDate + " to " + seconDate + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt; width:60px;~ID ~text-align:left; width:50px; font-size:8pt~Descripci�n De Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Estado de Resultados de Varios Per�odos (Mensual) " + DenomString + "<br/>Desde " + firstDate + " a " + seconDate + "<br/></span><span style=\"font-size:7pt\">Impreso En " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        DataTable COA = new DataTable();
        DataTable Report = new DataTable();

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();

        startDate1 = startDate.ToString("yyyy-MM-dd");
        
        for (i = 0; i <= MonthCount; i++)
        {
            // Check to it compare partial with partial
            if (Request.Form["SecondDate"] == DateTime.Now.ToString("yyyy-MM"))
            {
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01");
                endDate1 = endDate.AddMonths(i).ToString("yyyy-MM-dd");
            }
            else
            {
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01");
                endDate1 = startDate.AddMonths(i + 1).AddDays(-1).ToString("yyyy-MM-dd");
            }

            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='SINV') as Total" + i.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='SINV') as Tax" + i.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='SC') as TotalC" + i.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='SC') as TaxC" + i.ToString();


            QueryDate = QueryDate + "piv.Doc_Date between '" + startDate1 + "' and '" + endDate1 + "'";
            if (i < MonthCount)
                QueryDate = QueryDate + " OR ";

        }

        //SQL Command for queries
        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(COA);
        
        for (i = 0; i <= MonthCount; i++)
        {
            COA.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        COA.AcceptChanges();

        // Add percentage column.
        COA.Columns.Add("Percentage", typeof(String));

        COA.AcceptChanges();

        for (i = 0; i <= MonthCount; i++)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(COA.Rows[ii]["Total" + i.ToString()]))
                    COA.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["Tax" + i.ToString()]))
                    COA.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TotalC" + i.ToString()]))
                    COA.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TaxC" + i.ToString()]))
                    COA.Rows[ii]["TaxC" + i.ToString()] = 0;
                
                // Calculating SubTotal.
                COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]);

                // Denomination Calculation
                if (Denom > 1)
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble((COA.Rows[ii]["SubTotal" + i.ToString()])) / Denom;
                    COA.Rows[ii]["Total" + i.ToString()] = Convert.ToDouble((COA.Rows[ii]["Total" + i.ToString()])) / Denom;
                }

                // Rounding

                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    COA.Rows[ii]["Total" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]));
                }

                // Calculating Total.
                Total[i] += Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]);
            }
        }

        COA.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            // Calculate Percentage.
            for (i = 0; i < COA.Rows.Count; i++)
            {
                if ((double.Parse(COA.Rows[i]["SubTotal" + MonthCount.ToString()].ToString()) != 0) && (double.Parse(COA.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    COA.Rows[i]["Percentage"] = ((double.Parse(COA.Rows[i]["SubTotal" + MonthCount.ToString()].ToString()) - double.Parse(COA.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(COA.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((Total[MonthCount] - Total[0]) / Math.Abs(Total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            COA.AcceptChanges();
        }

        // Format all the output for the paper.
        for (i = 0; i <= MonthCount; i++)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                {

                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDecimal(COA.Rows[ii]["SubTotal" + i.ToString()]);
                }

                else
                { COA.Rows[ii]["SubTotal" + i.ToString()] = ""; }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }

                else
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Negative value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";


                // 0 value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = "";




            }
        }
        COA.AcceptChanges();

        //Percentage
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (COA.Rows[ii]["Percentage"].ToString() != "")
                    COA.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(COA.Rows[ii]["Percentage"]));

                if (COA.Rows[ii]["Percentage"].ToString() != "" && COA.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["Percentage"] = COA.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }
        
        COA.AcceptChanges();

        for (i = 0; i <= 15; i++)
        {
            Report.Columns.Add("Style" + i.ToString(), typeof(string));
            Report.Columns.Add("Field" + i.ToString(), typeof(string));
        }

        String Style2 = "padding: 3px 5px 3px 5px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
        String StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;";

        // PRINTING REPORT
        for (i = 0; i < COA.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && MonthCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (COA.Rows[i]["Percentage"].ToString() != "" && COA.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (MonthCount == 0)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (MonthCount == 1)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], Style2, COA.Rows[i]["SubTotal1"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], Style2, COA.Rows[i]["SubTotal1"], Style2, COA.Rows[i]["Subtotal2"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // FOR TOTAL    
        Style2 = Style2 + " font-weight:bold;";

        // Percent formatting
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (MonthCount == 0)

                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if ((MonthCount == 1) && Request.Form["Round"] == "on")
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[1])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        }
        else
        {
            if (MonthCount == 0)

                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (MonthCount == 1)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[1])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = Report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            COA.Columns.Remove("Currency");
            for (i = 0; i <= MonthCount; i++)
            {
                COA.Columns.Remove("Total" + i.ToString());
                COA.Columns.Remove("Tax" + i.ToString());
                COA.Columns.Remove("TotalC" + i.ToString());
                COA.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (i = 0; i < COA.Columns.Count; i++)
            {
                COA.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (i = COA.Columns.Count; i < 25; i++)
            {
                COA.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = COA.NewRow();
            excelHeader["value0"] = "Customer ID";
            excelHeader["value1"] = "Customer Name";

            // Add the header with dynamic number of columns
            for (i = 0; i <= MonthCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = startDate.AddMonths(i).ToString("MMMM") + Asterix;
                if (i == MonthCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            COA.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (i = 0; i < COA.Columns.Count - 1; i++)
            {
                COA.Rows[0][i] = "<b>" + COA.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = COA.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (i = 0; i <= MonthCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(Total[i]));
                if (i == MonthCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            COA.Rows.Add(excelTotal);

            RPT_Excel.DataSource = COA;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Sales Multiperiod Month to Month
    private void PrintMonthToMonthSalesMultiPer() {
        string language = Request.Form["language"];
        string firstDate = Request.Form["SecMonth"];
        string secondDate = Request.Form["SecMonth"].Substring(3, 4) + "-" + Request.Form["SecMonth"].Substring(0, 2) + "-" + "01";
        int denom = Int32.Parse(Request.Form["Denom"]);

        int yearCount = Int32.Parse(Request.Form["goback"]) - 1; ;
        DateTime[] startDate = new DateTime[yearCount + 1];
        DateTime[] endDate = new DateTime[yearCount + 1];
        DateTime selectionDate = DateTime.Parse(secondDate);
        string asterix = "";
        string denomString = "";
        string dateRange = "";
        string styleMonth = "";
        string query = "";
        string queryDate = "";
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        string styleFinish = "font-weight:bold; border-bottom: Double 3px black; border-top: 1px solid black;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable sales = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get fiscal date
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        // If date picked is current year, check if today's month is > fiscal month.
        if (selectionDate.Year == DateTime.Now.Year)
        {
            if (selectionDate.Month == DateTime.Now.Month)
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = new DateTime(DateTime.Today.AddYears(i - yearCount).Year, DateTime.Today.Month, 1);
                    endDate[i] = DateTime.Today.AddYears(i - yearCount);
                    asterix = "(*)";
                }
            }
            else
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = selectionDate.AddYears(i - yearCount);
                    endDate[i] = selectionDate.AddYears(i - yearCount).AddMonths(1).AddDays(-1);
                }
            }
        }
        else
        {
            for (int i = 0; i <= yearCount; i++)
            {
                startDate[i] = selectionDate.AddYears(i - yearCount);
                endDate[i] = selectionDate.AddYears(i - yearCount).AddMonths(1).AddDays(-1);
            }
        }


        // language 0 is english, language 1 is spanish
        if (language == "0")
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(In Tenth)";
                        break;
                    case 100:
                        denomString = "(In Hunreds)";
                        break;
                    case 1000:
                        denomString = "(In Thousands)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " and " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].ToString("MMMM yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; font-size:8pt~Account Description" + styleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Month To Month) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }
        else
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(En Décimo)";
                        break;
                    case 100:
                        denomString = "(En Centenares)";
                        break;
                    case 1000:
                        denomString = "(En Miles)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " y " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].ToString("MMMM yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + styleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Ventas de Varios Períodos (Mes A Mes)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TaxC" + i.ToString();

            queryDate = queryDate + "piv.Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "'";
            if (i < yearCount)
                queryDate = queryDate + " OR ";
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + queryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(sales);

        // Add the subtotal column.
        for (int i = 0; i <= yearCount; i++)
        {
            sales.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }
        // Add percentage column.
        sales.Columns.Add("Percentage", typeof(String));

        sales.AcceptChanges();

        // Rounding Calculation

        // Calculating Total and Sub-Total
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(sales.Rows[ii]["Total" + i.ToString()]))
                    sales.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["Tax" + i.ToString()]))
                    sales.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["TotalC" + i.ToString()]))
                    sales.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["TaxC" + i.ToString()]))
                    sales.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                sales.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(sales.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(sales.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["TaxC" + i.ToString()]));

                // Denomination Calculation.
                if (denom > 1)
                {
                    sales.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]) / denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    sales.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]);

            }
        }

        sales.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < sales.Rows.Count; i++)
            {
                if ((double.Parse(sales.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) != 0) && (double.Parse(sales.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    sales.Rows[i]["Percentage"] = ((double.Parse(sales.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(sales.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(sales.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((total[yearCount] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            sales.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "off")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));
                else
                    sales.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));

                // Negative value.
                if (sales.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = sales.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";

                // 0 value.
                if (sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }

        // Formatting Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                if (sales.Rows[ii]["Percentage"].ToString() != "")
                    sales.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(sales.Rows[ii]["Percentage"]));
                if (sales.Rows[ii]["Percentage"].ToString() != "" && sales.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    sales.Rows[ii]["Percentage"] = sales.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        sales.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            report.Columns.Add("Style" + i.ToString(), typeof(String));
            report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
        

        for (int i = 0; i < sales.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (sales.Rows[i]["Percentage"].ToString() != "" && sales.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], balanceStyle, sales.Rows[i]["Subtotal2"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }
        if (Request.Form["Round"] == "on")
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            sales.Columns.Remove("Currency");
            for (int i = 0; i <= yearCount; i++)
            {
                sales.Columns.Remove("Total" + i.ToString());
                sales.Columns.Remove("Tax" + i.ToString());
                sales.Columns.Remove("TotalC" + i.ToString());
                sales.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < sales.Columns.Count; i++)
            {
                sales.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = sales.Columns.Count; i < 25; i++)
            {
                sales.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = sales.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = endDate[i].ToString("MMMM yyyy") + asterix;
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            sales.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < sales.Columns.Count - 1; i++)
            {
                sales.Rows[0][i] = "<b>" + sales.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = sales.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            sales.Rows.Add(excelTotal);

            RPT_Excel.DataSource = sales;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Sales Multiperiod Quarterly
    private void PrintQuarterlySalesMultiPer()
    {
        DataTable COA = new DataTable();
        DataTable Report = new DataTable();
        DataTable fiscal = new DataTable();

        int Language = int.Parse(Request.Form["language"]);
        String Query = "";
        int Q = 0;

        String StyleMonth = "", Year, Qua_1, Qua_2, Qua_3, Qua_4, Qua_1_StartDate, Qua_1_EndDate, Qua_2_StartDate, Qua_2_EndDate, Qua_3_StartDate, Qua_3_EndDate, Qua_4_StartDate, Qua_4_EndDate, seconDate = "", startDate = "", Asterix, StyleFinish, fiscalDate, date1 = "";
        String[] Quarter = new String[4];
        string percent = "";
        Year = Request.Form["YearForQuater"];
        string balanceStyle = "";

        DateTime d1;

        Qua_1 = Request.Form["Q1"];
        Qua_2 = Request.Form["Q2"];
        Qua_3 = Request.Form["Q3"];
        Qua_4 = Request.Form["Q4"];

        String QueryDate = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        int Counter = int.Parse(Request.Form["count"]);


        // Denomination
        // Translation
        int Denom;
        Denom = int.Parse(Request.Form["Denom"]);
        String DenomString = "";
        if (Denom > 1)
        {
            if (Language == 0)
            {
                if (Denom == 10) { DenomString = "(In Tenth)"; }

                else if (Denom == 100) { DenomString = "(In Hundreds)"; }

                else if (Denom == 1000) { DenomString = "(In Thousands)"; }
            }
            else if (Language == 1)
            {
                if (Denom == 10) { DenomString = "(En Décimo)"; }

                else if (Denom == 100) { DenomString = "(En Centenares)"; }

                else if (Denom == 1000) { DenomString = "(En Miles)"; }
            }
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();
        // Get fiscal date
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);


        if (Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"]) >= 10)
        {
            fiscalDate = (int.Parse(Year) - 1) + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"];
            d1 = Convert.ToDateTime(fiscalDate);
        }
        else
        {
            fiscalDate = (int.Parse(Year) - 1) + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"];
            d1 = Convert.ToDateTime(fiscalDate);
        }

        date1 = fiscalDate;

        Qua_1_StartDate = d1.ToString("yyyy-MM-dd");
        Qua_1_EndDate = (d1.AddMonths(3).AddDays(-1)).ToString();

        Qua_2_StartDate = (d1.AddMonths(3)).ToString("yyyy-MM-dd");
        Qua_2_EndDate = (d1.AddMonths(6).AddDays(-1)).ToString();

        Qua_3_StartDate = (d1.AddMonths(6)).ToString("yyyy-MM-dd");
        Qua_3_EndDate = (d1.AddMonths(9).AddDays(-1)).ToString();

        Qua_4_StartDate = (d1.AddMonths(9)).ToString("yyyy-MM-dd");
        Qua_4_EndDate = (d1.AddMonths(12).AddDays(-1)).ToString();

        Asterix = "";
        DateTime now = DateTime.Now;
        //Check if the quarter picked is today's quarter
        //Check the year first { check the month compare to fiscal
        if ((int.Parse(Year) == DateTime.Now.Year && DateTime.Now.Month < Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"])) || (int.Parse(Year) == (DateTime.Now.Year - 1) && DateTime.Now.Month >= Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"])))
        {
            // Check the month to see if it's today's quarter got selected
            // Need to Mark which quarter is today's
            if (DateTime.Today >= d1 && now <= d1.AddMonths(3).AddDays(-1) && Qua_1 == "on")
            {
                //It's in Q1
                Qua_1_EndDate = now.ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(3) && now <= d1.AddMonths(6).AddDays(-1) && Qua_2 == "on")
            {
                //It's in Q2
                Qua_2_EndDate = now.ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(6) && now <= d1.AddMonths(9).AddDays(-1) && Qua_3 == "on")
            {
                //It's in Q3
                Qua_3_EndDate = now.ToString("yyyy-MM-dd");
                Qua_2_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-6).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(9) && now <= d1.AddMonths(12).AddDays(-1) && Qua_4 == "on")
            {
                //It's in Q4
                Qua_4_EndDate = now.ToString("yyyy-MM-dd");
                Qua_3_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Qua_2_EndDate = now.AddMonths(-6).ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-9).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
        }

        if (Qua_1 == "on")
        {
            if (Language == 0)
            {
                Quarter[0] = "Q-1";
            }
            else if (Language == 1)
            {
                Quarter[0] = "T-1";
            }

            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='SINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='SINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='SC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='SC') as TaxC" + Q.ToString();

            
            seconDate = Qua_1_EndDate;
            startDate = Qua_1_StartDate;
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }
        }
        if (Qua_2 == "on")
        {
            if (Language == 0)
            {
                Quarter[1] = "Q-2";
            }
            else if (Language == 1)
            {
                Quarter[1] = "T-2";
            }

            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='SINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='SINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='SC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='SC') as TaxC" + Q.ToString();


            seconDate = Qua_2_EndDate;
            if (Q == 0)
            {
                startDate = Qua_2_StartDate;
            }
            Q += 1;
            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }
        }
        if (Qua_3 == "on")
        {
            if (Language == 0)
            {
                Quarter[2] = "Q-3";
            }
            else if (Language == 1)
            {
                Quarter[2] = "T-3";
            }
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='SINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='SINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='SC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='SC') as TaxC" + Q.ToString();

            
            seconDate = Qua_3_EndDate;
            if (Q == 0)
            {
                startDate = Qua_3_StartDate;
            }
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }
        }
        if (Qua_4 == "on")
        {

            if (Language == 0)
            {
                Quarter[3] = "Q-4";
            }
            else if (Language == 1)
            {
                Quarter[3] = "T-4";
            }
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='SINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='SINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='SC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='SC') as TaxC" + Q.ToString();

            //Query1 = Query1 + ", ((Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + Q.ToString();
            //Query2 = Query2 + ", ((Select Sum(Debit_Amount) from ACC_GL where Transaction_Date Between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND') - (Select Sum(Credit_Amount) from ACC_GL where Transaction_Date Between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and fk_Account_Id = Account_ID AND Document_Type <> 'YEND')) as Balance" + Q.ToString();
            seconDate = Qua_4_EndDate;
            if (Q == 0)
            {
                startDate = Qua_4_StartDate;
            }
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "'";

        }

        double[] total = new double[Q + 1];
        double percentage = 0;
        // Header
        String H_Quarter, H_Qua_1 = "";
        int Temp = 0;
        for (int l = 0; l <= 3; l++)
        {
            if (Quarter[l] != null)
            {
                H_Quarter = Quarter[l] + Asterix;
                StyleMonth = StyleMonth + "~Text-align: Right;font-size:8pt~" + H_Quarter;
                if (Temp < (Q - 1))
                {
                    if (Temp < (Q - 2))
                    {
                        H_Qua_1 = H_Qua_1 + Quarter[l] + ", ";
                    }
                    else
                    {
                        H_Qua_1 = H_Qua_1 + Quarter[l];
                    }
                }
                else if (Temp == (Q - 1))
                {
                    if (Language == 0)
                    {
                        H_Qua_1 = H_Qua_1 + " and " + Quarter[l];
                    }
                    else if (Language == 1)
                    {
                        H_Qua_1 = H_Qua_1 + " y " + Quarter[l];
                    }
                }
                Temp += 1;
            }
        }
        
        //Translate the Header and Title
        if (Language == 0)
        {
            // Percentage.
            if (Request.Form["Percentage"] == "on")
                percent = "~Text-align:right; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right;font-size:0pt~";

            HF_PrintHeader.Value = "text-align:left; font-size:8pt;~ID~text-align:left; width:120px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Quarterly) " + DenomString + "<br/>For " + H_Qua_1 + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            if (Request.Form["Percentage"] == "on")
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Ventas de Varios Períodos (Anuales)" + DenomString + "<br/>Para " + H_Qua_1 + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        // Get the query
        if (Language == 0)
        {
            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(COA);
        }
        else if (Language == 1)
        {
            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(COA);
        }

        // Add the subtotal column.
        for (int i = 0; i < Q; i++)
        {
            COA.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        // Add percentage column.
        COA.Columns.Add("Percentage", typeof(String));

        COA.AcceptChanges();

        // Calculating Total and Sub-Total
        for (int i = 0; i <= (Q - 1); i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(COA.Rows[ii]["Total" + i.ToString()]))
                    COA.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["Tax" + i.ToString()]))
                    COA.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TotalC" + i.ToString()]))
                    COA.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TaxC" + i.ToString()]))
                    COA.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]));
                
                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }
                
                // Denomination Calculation.
                if (Denom > 1)
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]) / Denom;
                }
                // Calculating Total.
                total[i] += Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]);


            }
        }
        COA.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on")
        {
            // Calculate Percentage.
            for (int i = 0; i < COA.Rows.Count; i++)
            {
                if ((double.Parse(COA.Rows[i]["SubTotal" + (Q - 1).ToString()].ToString()) != 0) && (double.Parse(COA.Rows[i]["SubTotal0"].ToString()) != 0))
                    COA.Rows[i]["Percentage"] = ((double.Parse(COA.Rows[i]["SubTotal" + (Q - 1).ToString()].ToString()) - double.Parse(COA.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(COA.Rows[i]["SubTotal0"].ToString())));
            }
            percentage = ((total[(Q - 1)] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            COA.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i < Q; i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = "";
                }
                else
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = "";
                }
                // Negative value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() != "")
                {
                    if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                        COA.Rows[ii]["SubTotal" + i.ToString()] = COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";
                }
                // 0 value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }
        if (Request.Form["Percentage"] == "on")
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (COA.Rows[ii]["Percentage"].ToString() != "")
                    COA.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(COA.Rows[ii]["Percentage"]));
                if (COA.Rows[ii]["Percentage"].ToString() != "" && COA.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["Percentage"] = COA.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        COA.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            Report.Columns.Add("Style" + i.ToString(), typeof(String));
            Report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt;";
        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black";
       
        for (int i = 0; i < COA.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && Q > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (COA.Rows[i]["Percentage"].ToString() != "" && COA.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }
       
            if (Q == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], balanceStyle, COA.Rows[i]["Subtotal2"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        
        // Percent formatting
        if (Request.Form["Percentage"] == "on" && Q > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (Q == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (Q == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = Report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            COA.Columns.Remove("Currency");
            for (int i = 0; i < Q; i++)
            {
                COA.Columns.Remove("Total" + i.ToString());
                COA.Columns.Remove("Tax" + i.ToString());
                COA.Columns.Remove("TotalC" + i.ToString());
                COA.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < COA.Columns.Count; i++)
            {
                COA.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = COA.Columns.Count; i < 25; i++)
            {
                COA.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = COA.NewRow();
            excelHeader["value0"] = "Customer ID";
            excelHeader["value1"] = "Customer Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i < Q; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = Quarter[i] + Asterix;
                if (i == Q-1 && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            COA.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < COA.Columns.Count - 1; i++)
            {
                COA.Rows[0][i] = "<b>" + COA.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = COA.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i < Q; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == (Q - 1) && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = String.Format("{0:P}", percentage);
            }
            COA.Rows.Add(excelTotal);

            RPT_Excel.DataSource = COA;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Sales Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterSalesMultiPer()
    {

        int Language = int.Parse(Request.Form["language"]);
        string seconDate = Request.Form["Quarter"];
        string[] words = seconDate.Split(' ');
        int Qua_No = int.Parse(words[0]);
        string Qua_Year = words[1].ToString();
        string Goback = Request.Form["goback"];
        string Show_Per = Request.Form["Percentage"];
        int Denom = int.Parse(Request.Form["Denom"]);
        string[] Quarter = new string[4];
        string asterix = "";

        DataTable fiscal = new DataTable();
        DataTable COA = new DataTable();
        DataTable Report = new DataTable();
        string StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black";
        string percentStyle = "font-size:0pt; width: 0px;";
        String DenomString = "", StyleMonth = "";
        String QueryDate = "";
        if (Denom > 1)
        {
            if (Language == 0)
            {
                if (Denom == 10) { DenomString = "(In Tenth)"; }

                else if (Denom == 100) { DenomString = "(In Hundreds)"; }

                else if (Denom == 1000) { DenomString = "(In Thousands)"; }
            }
            else if (Language == 1)
            {
                if (Denom == 10) { DenomString = "(En D�cimo)"; }

                else if (Denom == 100) { DenomString = "(En Centenares)"; }

                else if (Denom == 1000) { DenomString = "(En Miles)"; }
            }
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();

        string balanceStyle = "", Query = "", percent = "", startDate2 = "", endDate2 = "", fiscalDate = "", Qdate1 = "", Qdate2 = "", Date1 = "";
        DateTime d1;
        double[] total = new double[(int.Parse(Goback)) + 1];
        double percentage = 0;

        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);


        if (Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"]) >= 10)
        {
            fiscalDate = (Qua_Year + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"]);
            d1 = Convert.ToDateTime(fiscalDate);
        }
        else
        {
            fiscalDate = (Qua_Year + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"]);
            d1 = Convert.ToDateTime(fiscalDate);
        }
        if (Qua_No == 1)
        {
            Qdate1 = d1.ToString();
            Qdate2 = (d1.AddMonths(3).AddDays(-1)).ToString();
        }
        else if (Qua_No == 2)
        {
            Qdate1 = (d1.AddMonths(3)).ToString();
            Qdate2 = (d1.AddMonths(6).AddDays(-1)).ToString();
        }
        else if (Qua_No == 3)
        {
            Qdate1 = (d1.AddMonths(6)).ToString();
            Qdate2 = (d1.AddMonths(9).AddDays(-1)).ToString();
        }
        else if (Qua_No == 4)
        {
            Qdate1 = (d1.AddMonths(9)).ToString();
            Qdate2 = (d1.AddMonths(12).AddDays(-1)).ToString();
        }
        DateTime S_Date = Convert.ToDateTime(Qdate1);
        DateTime E_Date = Convert.ToDateTime(Qdate2);

        // Check the year first then check the month compare to fiscal+

        DateTime now = DateTime.Now;
        //Check if the quarter picked is today's quarter
        //Check the year first { check the month compare to fiscal
        if (S_Date < now && E_Date > now)
        {
            // Check the month to see if it's today's quarter got selected
            // Need to Mark which quarter is today's
            if (DateTime.Today >= d1 && now <= d1.AddMonths(3).AddDays(-1))
            {
                //It's in Q1
                Qdate1 = d1.ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
            else if (DateTime.Today >= d1.AddMonths(3) && now <= d1.AddMonths(6).AddDays(-1))
            {
                //It's in Q2
                Qdate1 = (d1.AddMonths(3)).ToString();
                Qdate2 = (now.ToString("yyyy-MM-dd")).ToString();
            }
            else if (DateTime.Today >= d1.AddMonths(6) && now <= d1.AddMonths(9).AddDays(-1))
            {
                //It's in Q3
                Qdate1 = (d1.AddMonths(6)).ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
            else if (DateTime.Today >= d1.AddMonths(9) && now <= d1.AddMonths(12).AddDays(-1))
            {
                //It's in Q4
                Qdate1 = (d1.AddMonths(9)).ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
        }
        S_Date = Convert.ToDateTime(Qdate1);
        E_Date = Convert.ToDateTime(Qdate2);

        int j = 0;
        DateTime firstDate = S_Date.AddYears(-(int.Parse(Goback) - 1));
        DateTime secondDate = E_Date.AddYears(-(int.Parse(Goback) - 1));
        Qdate1 = ""; Qdate2 = "";
        for (j = 0; j < (int.Parse(Goback)); j++)
        {
            if (Language == 0)
            {
                if (S_Date < now && E_Date > now)
                {
                    Qdate1 = ("Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2)) + "(*)").ToString();
                    asterix = "(*)";
                }                    
                else
                    Qdate1 = "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2));

                if (j == ((int.Parse(Goback)) - 1))
                    Qdate2 = Qdate2 + " and Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2));
                else if (j == ((int.Parse(Goback)) - 2))
                    Qdate2 = Qdate2 + "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2));
                else
                    Qdate2 = Qdate2 + "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2)) + ", ";

            }
            if (Language == 1)
            {
                if (S_Date < now && E_Date >= now)
                {
                    Qdate1 = "T" + Qua_No + " " + firstDate.ToString("yyyy") + "(*)";
                    asterix = ("*");
                }                    
                else
                    Qdate1 = "T" + Qua_No + " " + firstDate.ToString("yyyy");

                if (j == ((int.Parse(Goback)) - 1))
                    Qdate2 = Qdate2 + " y T" + Qua_No + " " + firstDate.ToString("yyyy");
                else if (j == ((int.Parse(Goback)) - 2))
                    Qdate2 = Qdate2 + "T" + Qua_No + " " + firstDate.ToString("yyyy");
                else
                    Qdate2 = Qdate2 + "T" + Qua_No + " " + firstDate.ToString("yyyy") + ", ";
            }

            StyleMonth = StyleMonth + "~Text-align: right; font-size:8pt~" + Qdate1;

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd");
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd");
            firstDate = Convert.ToDateTime(startDate2);
            secondDate = Convert.ToDateTime(endDate2);
        }
        //Translate the Header and Title
        if (Language == 0)
        {
            // Percentage.
            if (Show_Per == "on")
                percent = "~Text-align:right; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; font-size:0pt~";

            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; width:120px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Quarter to Quarter) " + DenomString + "<br/>For " + Qdate2 + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            if (Show_Per == "on")
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Ventas de Varios Períodos (Trimestre a Trimestre)" + DenomString + "<br/>Para " + Qdate2 + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;

        firstDate = S_Date.AddYears(-(int.Parse(Goback) - 1));
        secondDate = E_Date.AddYears(-(int.Parse(Goback) - 1));
        startDate2 = firstDate.ToString("yyyy-MM-dd");
        endDate2 = secondDate.ToString("yyyy-MM-dd");

        for (j = 0; j < (int.Parse(Goback)); j++)
        {
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='SINV') as Total" + j.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='SINV') as Tax" + j.ToString();
            Query = Query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='SC') as TotalC" + j.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='SC') as TaxC" + j.ToString();


            QueryDate = QueryDate + "piv.Doc_Date between '" + startDate2 + "' and '" + endDate2 + "'";
            if (j < int.Parse(Goback) - 1)
                QueryDate = QueryDate + " OR ";

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd");
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd");
            firstDate = Convert.ToDateTime(startDate2);
            secondDate = Convert.ToDateTime(endDate2);
        }
        if (Language == 0)
        {

            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(COA);

        }
        else if (Language == 1)
        {

            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(COA);
        }

        for (int i = 0; i < int.Parse(Goback); i++)
        {
            COA.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        // Add percentage column.
        COA.Columns.Add("Percentage", typeof(String));

        COA.AcceptChanges();

        // Calculating Total and Sub-Total
        for (int i = 0; i < int.Parse(Goback); i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(COA.Rows[ii]["Total" + i.ToString()]))
                    COA.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["Tax" + i.ToString()]))
                    COA.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TotalC" + i.ToString()]))
                    COA.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TaxC" + i.ToString()]))
                    COA.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]));
                
                // Denomination Calculation.
                if (Denom > 1)
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]) / Denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }


                // Calculating Total.
                total[i] += Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]);
            }
        }
        COA.AcceptChanges();

        // Percentage.
        if (Show_Per == "on" && (int.Parse(Goback) - 1) >  0)
        {
            // Calculate Percentage.
            for (int i = 0; i < COA.Rows.Count; i++)
            {
                if ((double.Parse(COA.Rows[i]["SubTotal" + (int.Parse(Goback) - 1).ToString()].ToString()) != 0) && (double.Parse(COA.Rows[i]["SubTotal0"].ToString()) != 0))
                    COA.Rows[i]["Percentage"] = ((double.Parse(COA.Rows[i]["SubTotal" + (int.Parse(Goback) - 1).ToString()].ToString()) - double.Parse(COA.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(COA.Rows[i]["SubTotal0"].ToString())));
            }
            percentage = ((total[(int.Parse(Goback) - 1)] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; width: 80px;";
            COA.AcceptChanges();
        }
        
        // Format all the output for the paper.
        for (int i = 0; i <= (int.Parse(Goback) - 1); i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = "";
                }
                else
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = "";
                }

                // Negative value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() != "")
                {
                    if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                        COA.Rows[ii]["SubTotal" + i.ToString()] = COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";
                }
                // 0 value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }
        if (Show_Per == "on" && (int.Parse(Goback) - 1) > 0)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (COA.Rows[ii]["Percentage"].ToString() != "")
                    COA.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(COA.Rows[ii]["Percentage"]));
                if (COA.Rows[ii]["Percentage"].ToString() != "" && COA.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["Percentage"] = COA.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        COA.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            Report.Columns.Add("Style" + i.ToString(), typeof(String));
            Report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; max-width: 5px; min-width: 5px;";

        for (int i = 0; i < COA.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && (int.Parse(Goback) - 1) > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (COA.Rows[i]["Percentage"].ToString() != "" && COA.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 4in; max-width: 4.5in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3.5in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], balanceStyle, COA.Rows[i]["Subtotal2"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && (int.Parse(Goback) - 1) > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = Report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            COA.Columns.Remove("Currency");
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                COA.Columns.Remove("Total" + i.ToString());
                COA.Columns.Remove("Tax" + i.ToString());
                COA.Columns.Remove("TotalC" + i.ToString());
                COA.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < COA.Columns.Count; i++)
            {
                COA.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = COA.Columns.Count; i < 25; i++)
            {
                COA.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = COA.NewRow();
            excelHeader["value0"] = "Customer ID";
            excelHeader["value1"] = "Customer Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                excelHeader["value" + (i + 2).ToString()] = "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + i) - (int.Parse(Goback) - 2)) + asterix;
                if (i == int.Parse(Goback) - 1 && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            COA.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < COA.Columns.Count - 1; i++)
            {
                COA.Rows[0][i] = "<b>" + COA.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = COA.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == (int.Parse(Goback) - 1) && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = String.Format("{0:P}", percentage);
            }
            COA.Rows.Add(excelTotal);

            RPT_Excel.DataSource = COA;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Sales Multiperiod Yearly
    private void PrintYearlySalesMultiPer() {

        string language = Request.Form["language"];
        string firstDate = Request.Form["FirstDate"];
        string secondDate = Request.Form["SecondDate"];
        int denom = Int32.Parse(Request.Form["Denom"]);

        int yearCount = Int32.Parse(Request.Form["SecondDate"]) - Int32.Parse(Request.Form["FirstDate"]); ;
        DateTime[] startDate = new DateTime[yearCount + 1];
        DateTime[] endDate = new DateTime[yearCount + 1];
        DateTime fiscalDate;
        string asterix = "";
        string denomString = "";
        string dateRange = "";
        string styleMonth = "";
        string query = "";
        string queryDate = "";
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        string styleFinish = "font-weight:bold; border-bottom: Double 3px black; border-top: 1px solid black;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable sales = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get fiscal date.
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        // Because it is '9' not '09'.
        if (Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()) < 10)
            fiscalDate = DateTime.Parse(firstDate + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString() + "-01");
        else
            fiscalDate = DateTime.Parse(firstDate + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString() + "-01");

        // If date picked is current year, check if today's month is > fiscal month.
        if (Int32.Parse(secondDate) >= DateTime.Now.Year)
        {
            // Check if today's date already pass the fiscal month.
            //  If not, use today's date to compare with previous year.
            if (DateTime.Now.Month < Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()))
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = new DateTime(DateTime.Today.AddYears(i - yearCount - 1).Year, Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()), 1);
                    endDate[i] = DateTime.Today.AddYears(i - yearCount);
                    asterix = "(*)";
                }
            }
            else
            {
                for (int i = 0; i <=  yearCount; i++)
                {
                    startDate[i] = fiscalDate.AddYears(i - 1);
                    endDate[i] = fiscalDate.AddYears(i).AddDays(-1);
                } 
            }
        }
        else
        {
            for (int i = 0; i <= yearCount; i++)
            {
                startDate[i] = fiscalDate.AddYears(i - 1);
                endDate[i] = fiscalDate.AddYears(i).AddDays(-1);
            }
        }

        // language 0 is english, language 1 is spanish
        if (language == "0")
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(In Tenth)";
                        break;
                    case 100:
                        denomString = "(In Hunreds)";
                        break;
                    case 1000:
                        denomString = "(In Thousands)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i ++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " and " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; font-size:8pt~Account Description" + styleMonth + percent;
            //HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)";
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Yearly) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }
        else
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(En Décimo)";
                        break;
                    case 100:
                        denomString = "(En Centenares)";
                        break;
                    case 1000:
                        denomString = "(En Miles)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " y " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + styleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Ventas de Varios Períodos (Anuales)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TaxC" + i.ToString();

            queryDate = queryDate + "piv.Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "'";
            if (i < yearCount)
                queryDate = queryDate + " OR ";
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + queryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(sales);

        // Add the subtotal column.
        for (int i = 0; i <= yearCount; i++)
        {
            sales.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        // Add percentage column.
        sales.Columns.Add("Percentage", typeof(String));

        sales.AcceptChanges();

        // Calculating Total and Sub-Total
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(sales.Rows[ii]["Total" + i.ToString()]))
                    sales.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["Tax" + i.ToString()]))
                    sales.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["TotalC" + i.ToString()]))
                    sales.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(sales.Rows[ii]["TaxC" + i.ToString()]))
                    sales.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                sales.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(sales.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(sales.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["TaxC" + i.ToString()]));

                // Denomination Calculation.
                if (denom > 1)
                {
                    sales.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]) / denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    sales.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]);
            }
        }

        sales.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < sales.Rows.Count; i++)
            {
                if ((double.Parse(sales.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) != 0) && (double.Parse(sales.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    sales.Rows[i]["Percentage"] = ((double.Parse(sales.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(sales.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(sales.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((total[yearCount] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            sales.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "off")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));
                else
                    sales.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(sales.Rows[ii]["SubTotal" + i.ToString()]));
                    
                // Negative value.
                if (sales.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = sales.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";

                // 0 value.
                if (sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || sales.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    sales.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }

        // Formatting Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            for (int ii = 0; ii < sales.Rows.Count; ii++)
            {
                if (sales.Rows[ii]["Percentage"].ToString() != "")
                    sales.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(sales.Rows[ii]["Percentage"]));
                if (sales.Rows[ii]["Percentage"].ToString() != "" && sales.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    sales.Rows[ii]["Percentage"] = sales.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        sales.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            report.Columns.Add("Style" + i.ToString(), typeof(String));
            report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";

        for (int i = 0; i < sales.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (sales.Rows[i]["Percentage"].ToString() != "" && sales.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], balanceStyle, sales.Rows[i]["Subtotal2"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            sales.Columns.Remove("Currency");
            for (int i = 0; i <= yearCount; i++)
            {
                sales.Columns.Remove("Total" + i.ToString());
                sales.Columns.Remove("Tax" + i.ToString());
                sales.Columns.Remove("TotalC" + i.ToString());
                sales.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < sales.Columns.Count; i++)
            {
                sales.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = sales.Columns.Count; i < 25; i++)
            {
                sales.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = sales.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            sales.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < sales.Columns.Count - 1; i++)
            {
                sales.Rows[0][i] = "<b>" + sales.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = sales.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            sales.Rows.Add(excelTotal);

            RPT_Excel.DataSource = sales;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Purchases Multiperiod Monthly 
    private void PrintMonthlyPurchasesMultiPer()
    {
        int Language = Convert.ToInt32(Request.Form["language"]);
        string firstDate = Request.Form["FirstDate"];
        string seconDate = Request.Form["SecondDate"];
        Int32 Denom = Convert.ToInt32(Request.Form["Denom"]);
        string Query = "";
        System.DateTime startDate = default(System.DateTime);
        System.DateTime endDate = default(System.DateTime);
        string startDate1 = "";
        string endDate1 = "";
        string StyleMonth = "";
        string Asterix = "";
        string DenomString = "";
        int MonthCount = 0;
        int i = 0;
        int ii = 0;
        double percentage = 0;
        string percent = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        String Style2 = "";
        String StyleFinish = "";
        String QueryDate = "";

        DataTable fiscal = new DataTable();

        // Get fiscal date
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        // Get the MonthCount Value

        try
        {
            startDate = DateTime.Parse(firstDate);
            endDate = DateTime.Parse(seconDate);

            while (startDate != endDate)
            {
                startDate = startDate.AddMonths(1);
                MonthCount += 1;
            }
        }
        catch (Exception ex)
        {
            MonthCount = 0;
        }

        double[] Total = new double[MonthCount + 1];

        if ((Denom > 1))
        {
            if (Language == 0)
            {
                if (Denom == 10)
                {
                    DenomString = "(In Tenth)";
                }
                else if (Denom == 100)
                {
                    DenomString = "(In Hundreds)";
                }
                else if (Denom == 1000)
                {
                    DenomString = "(In Thousands)";
                }
            }
            else if (Language == 1)
            {
                if (Denom == 10)
                {
                    DenomString = "(En D�cimo)";
                }
                else if (Denom == 100)
                {
                    DenomString = "(En Centenares)";
                }
                else if (Denom == 1000)
                {
                    DenomString = "(En Miles)";
                }
            }
        }

        Asterix = "";

        if (Request.Form["SecondDate"] == DateTime.Now.ToString("yyyy-MM"))
        {
            seconDate = DateTime.Now.ToString("yyyy-MM-dd");
            firstDate = DateTime.Now.AddMonths(-MonthCount).ToString("yyyy-MM-01");

            endDate = DateTime.Now.AddMonths(-MonthCount);
            endDate1 = DateTime.Now.AddMonths(-MonthCount).ToString("yyyy-MM-dd");
            Asterix = "(*)";

        }
        else
        {
            // Default date give today's date
            if (string.IsNullOrEmpty(firstDate))
            {
                firstDate = DateTime.Now.ToString("yyyy-MM-01");
                Asterix = "(*)";
            }
            else
            {
                // If exist, take the the first day of month
                startDate = DateTime.Parse(firstDate);
            }
            if (string.IsNullOrEmpty(seconDate))
            {
                seconDate = DateTime.Now.ToString("yyyy-MM-dd");
                endDate = DateTime.Now;
                Asterix = "(*)";
            }
            else
            {
                // If exist, take the the last day of month
                endDate = DateTime.Parse(seconDate);
                endDate1 = startDate.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd");
            }
        }

        startDate = DateTime.Parse(firstDate);
        startDate1 = startDate.ToString("yyyy-MM-dd");

        for (i = 0; i <= MonthCount; i++)
        {
            if (Language == 0)
            {
                StyleMonth = StyleMonth + "~Text-align: right; font-size:8pt~" + startDate.AddMonths(i).ToString("MMMM") + Asterix;
            }
        }

        // Percentage.
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
            percent = "~Text-align: right; font-size:8pt~Percentage(%)";
        else
            percent = "~Text-align: right;  font-size:0pt~";

        if (Language == 0)
        {
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt; width:0px;~ID~text-align:left; width:50px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Purchase(Monthly) " + DenomString + "<br/>From " + firstDate + " to " + seconDate + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            HF_PrintHeader.Value = "Text-align:left;font-size:8pt; width:0px;~ID ~text-align:left; width:50px; font-size:8pt~Descripci�n De Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Estado de Resultados de Varios Per�odos (Mensual) " + DenomString + "<br/>Desde " + firstDate + " a " + seconDate + "<br/></span><span style=\"font-size:7pt\">Impreso En " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        DataTable COA = new DataTable();
        DataTable Report = new DataTable();

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();

        startDate1 = startDate.ToString("yyyy-MM-dd");

        for (i = 0; i <= MonthCount; i++)
        {
            // Check to it compare partial with partial
            if (Request.Form["SecondDate"] == DateTime.Now.ToString("yyyy-MM"))
            {
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01");
                endDate1 = endDate.AddMonths(i).ToString("yyyy-MM-dd");
            }
            else
            {
                startDate1 = startDate.AddMonths(i).ToString("yyyy-MM-01");
                endDate1 = startDate.AddMonths(i + 1).AddDays(-1).ToString("yyyy-MM-dd");
            }

            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='PINV') as Total" + i.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='PINV') as Tax" + i.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='PC') as TotalC" + i.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate1 + "' and '" + endDate1 + "' and Doc_Type='PC') as TaxC" + i.ToString();

            QueryDate = QueryDate + "piv.Doc_Date between '" + startDate1 + "' and '" + endDate1 + "'";
            if (i < MonthCount)
                QueryDate = QueryDate + " OR ";
        }

        //SQL Command for the queries
        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(COA);

        for (i = 0; i <= MonthCount; i++)
        {
            COA.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        COA.AcceptChanges();

        // Add percentage column.
        COA.Columns.Add("Percentage", typeof(String));

        COA.AcceptChanges();

        for (i = 0; i <= MonthCount; i++)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(COA.Rows[ii]["Total" + i.ToString()]))
                    COA.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["Tax" + i.ToString()]))
                    COA.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TotalC" + i.ToString()]))
                    COA.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TaxC" + i.ToString()]))
                    COA.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()])); ;

                // Denomination Calculation
                if (Denom > 1)
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble((COA.Rows[ii]["SubTotal" + i.ToString()].ToString())) / Denom;
                }

                // Rounding

                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }


                // Calculating Total.
                Total[i] += Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]);



            }
        }

        COA.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            // Calculate Percentage.
            for (i = 0; i < COA.Rows.Count; i++)
            {
                if ((double.Parse(COA.Rows[i]["SubTotal" + MonthCount.ToString()].ToString()) != 0) && (double.Parse(COA.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    COA.Rows[i]["Percentage"] = ((double.Parse(COA.Rows[i]["SubTotal" + MonthCount.ToString()].ToString()) - double.Parse(COA.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(COA.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((Total[MonthCount] - Total[0]) / Math.Abs(Total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 0px; max-width: 0px;";
            COA.AcceptChanges();
        }

        //Formatting all the output
        for (i = 0; i <= MonthCount; i++)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Rounding for subtotal
                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }

                else
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }
                // Negative value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";


                // 0 value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = "";


            }
        }

        COA.AcceptChanges();

        //Percentage Format
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            for (ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (COA.Rows[ii]["Percentage"].ToString() != "")
                    COA.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(COA.Rows[ii]["Percentage"]));

                if (COA.Rows[ii]["Percentage"].ToString() != "" && COA.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["Percentage"] = COA.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        COA.AcceptChanges();

        for (i = 0; i <= 15; i++)
        {
            Report.Columns.Add("Style" + i.ToString(), typeof(string));
            Report.Columns.Add("Field" + i.ToString(), typeof(string));
        }


        Style2 = "text-align: right; font-size: 8pt; min-width: 5px; max-width: 5px;";
        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;";

        // PRINTING REPORT
        for (i = 0; i < COA.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && MonthCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (COA.Rows[i]["Percentage"].ToString() != "" && COA.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (MonthCount == 0)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (MonthCount == 1)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], Style2, COA.Rows[i]["SubTotal1"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", COA.Rows[i]["Name"].ToString(), Style2, COA.Rows[i]["SubTotal0"], Style2, COA.Rows[i]["SubTotal1"], Style2, COA.Rows[i]["SubTotal2"] + "</span>", percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // FOR TOTAL    
        Style2 = Style2 + " font-weight:bold;";

        // Percent formatting
        if (Request.Form["Percentage"] == "on" && MonthCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }
        
        if (Request.Form["Round"] == "on")
        {
            if (MonthCount == 0)

                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if ((MonthCount == 1) && Request.Form["Round"] == "on")
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[1])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(Total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (MonthCount == 0)

                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (MonthCount == 1)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[0])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[1])) + "</span>", Style2, "<span style=\"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(Total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = Report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            COA.Columns.Remove("Currency");
            for (i = 0; i <= MonthCount; i++)
            {
                COA.Columns.Remove("Total" + i.ToString());
                COA.Columns.Remove("Tax" + i.ToString());
                COA.Columns.Remove("TotalC" + i.ToString());
                COA.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (i = 0; i < COA.Columns.Count; i++)
            {
                COA.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (i = COA.Columns.Count; i < 25; i++)
            {
                COA.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = COA.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (i = 0; i <= MonthCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = startDate.AddMonths(i).ToString("MMMM") + Asterix;
                if (i == MonthCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            COA.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for ( i = 0; i < COA.Columns.Count - 1; i++)
            {
                COA.Rows[0][i] = "<b>" + COA.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = COA.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (i = 0; i <= MonthCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(Total[i]));
                if (i == MonthCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            COA.Rows.Add(excelTotal);

            RPT_Excel.DataSource = COA;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Purchases Multiperiod Month to Month
    private void PrintMonthToMonthPurchasesMultiPer() {
        string language = Request.Form["language"];
        string firstDate = Request.Form["SecMonth"];
        string secondDate = Request.Form["SecMonth"].Substring(3,4) + "-" + Request.Form["SecMonth"].Substring(0,2) + "-" + "01";
        int denom = Int32.Parse(Request.Form["Denom"]);

        int yearCount = Int32.Parse(Request.Form["goback"]) - 1; ;
        DateTime[] startDate = new DateTime[yearCount + 1];
        DateTime[] endDate = new DateTime[yearCount + 1];
        DateTime selectionDate = DateTime.Parse(secondDate);
        string asterix = "";
        string denomString = "";
        string dateRange = "";
        string styleMonth = "";
        string query = "";
        string queryDate = "";
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        string styleFinish = "font-weight:bold; border-bottom: Double 3px black; border-top: 1px solid black;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable purchases = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get fiscal date
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        // If date picked is current year, check if today's month is > fiscal month.
        if (selectionDate.Year == DateTime.Now.Year)
        {
            if (selectionDate.Month == DateTime.Now.Month)
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = new DateTime(DateTime.Today.AddYears(i - yearCount).Year, DateTime.Today.Month, 1);
                    endDate[i] = DateTime.Today.AddYears(i - yearCount);
                    asterix = "(*)";
                }
            }
            else
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = selectionDate.AddYears(i - yearCount);
                    endDate[i] = selectionDate.AddYears(i - yearCount).AddMonths(1).AddDays(-1);
                }
            }
        }
        else
        {
            for (int i = 0; i <= yearCount; i++)
            {
                startDate[i] = selectionDate.AddYears(i - yearCount);
                endDate[i] = selectionDate.AddYears(i - yearCount).AddMonths(1).AddDays(-1);
            }
        }
        

        // language 0 is english, language 1 is spanish
        if (language == "0")
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(In Tenth)";
                        break;
                    case 100:
                        denomString = "(In Hunreds)";
                        break;
                    case 1000:
                        denomString = "(In Thousands)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " and " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].ToString("MMMM yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; font-size:8pt~Account Description" + styleMonth + percent;
            //HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)";
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Purchases(Month To Month) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }
        else
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(En Décimo)";
                        break;
                    case 100:
                        denomString = "(En Centenares)";
                        break;
                    case 1000:
                        denomString = "(En Miles)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " y " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].ToString("MMMM yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + styleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Compras Con Varios Períodos (Mes A Mes)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TaxC" + i.ToString();

            queryDate = queryDate + "piv.Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "'";
            if (i < yearCount)
                queryDate = queryDate + " OR ";
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + queryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(purchases);

        // Add the subtotal column.
        for (int i = 0; i <= yearCount; i++)
        {
            purchases.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }
        // Add percentage column.
        purchases.Columns.Add("Percentage", typeof(String));

        purchases.AcceptChanges();

        // Rounding Calculation

        // Calculating Total and Sub-Total
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(purchases.Rows[ii]["Total" + i.ToString()]))
                    purchases.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["Tax" + i.ToString()]))
                    purchases.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TotalC" + i.ToString()]))
                    purchases.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TaxC" + i.ToString()]))
                    purchases.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(purchases.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["TaxC" + i.ToString()]));

                // Denomination Calculation.
                if (denom > 1)
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]) / denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]);


            }
        }

        purchases.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < purchases.Rows.Count; i++)
            {
                if ((double.Parse(purchases.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) != 0) && (double.Parse(purchases.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    purchases.Rows[i]["Percentage"] = ((double.Parse(purchases.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(purchases.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(purchases.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((total[yearCount] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            purchases.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "off")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                else
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));

                // Negative value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";

                // 0 value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }

        // Formatting Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                if (purchases.Rows[ii]["Percentage"].ToString() != "")
                    purchases.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(purchases.Rows[ii]["Percentage"]));
                if (purchases.Rows[ii]["Percentage"].ToString() != "" && purchases.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    purchases.Rows[ii]["Percentage"] = purchases.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }
        
        purchases.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            report.Columns.Add("Style" + i.ToString(), typeof(String));
            report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";

        for (int i = 0; i < purchases.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (purchases.Rows[i]["Percentage"].ToString() != "" && purchases.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], balanceStyle, purchases.Rows[i]["Subtotal2"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            purchases.Columns.Remove("Currency");
            for (int i = 0; i <= yearCount; i++)
            {
                purchases.Columns.Remove("Total" + i.ToString());
                purchases.Columns.Remove("Tax" + i.ToString());
                purchases.Columns.Remove("TotalC" + i.ToString());
                purchases.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < purchases.Columns.Count; i++)
            {
                purchases.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = purchases.Columns.Count; i < 25; i++)
            {
                purchases.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = purchases.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = endDate[i].ToString("MMMM yyyy") + asterix;
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            purchases.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < purchases.Columns.Count - 1; i++)
            {
                purchases.Rows[0][i] = "<b>" + purchases.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = purchases.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            purchases.Rows.Add(excelTotal);

            RPT_Excel.DataSource = purchases;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Purchases Multiperiod Quarterly
    private void PrintQuarterlyPurchasesMultiPer()
    {
        DataTable report = new DataTable();
        DataTable purchases = new DataTable();
        DataTable fiscal = new DataTable();

        int Language = int.Parse(Request.Form["language"]);
        String Query = "";
        int Q = 0;

        String StyleMonth = "", Year, Qua_1, Qua_2, Qua_3, Qua_4, Qua_1_StartDate, Qua_1_EndDate, Qua_2_StartDate, Qua_2_EndDate, Qua_3_StartDate, Qua_3_EndDate, Qua_4_StartDate, Qua_4_EndDate, seconDate = "", startDate = "", Asterix, StyleFinish, fiscalDate, date1 = "";
        String[] Quarter = new String[4];
        string percent = "";
        Year = Request.Form["YearForQuater"];
        string balanceStyle = "";

        DateTime d1;

        Qua_1 = Request.Form["Q1"];
        Qua_2 = Request.Form["Q2"];
        Qua_3 = Request.Form["Q3"];
        Qua_4 = Request.Form["Q4"];

        string percentStyle = "text-align:right; font-size:0pt; width: 0px;";

        String QueryDate = "";
        int Counter = int.Parse(Request.Form["count"]);
        
        // Denomination
        // Translation
        int Denom;
        Denom = int.Parse(Request.Form["Denom"]);
        String DenomString = "";
        if (Denom > 1)
        {
            if (Language == 0)
            {
                if (Denom == 10) { DenomString = "(In Tenth)"; }

                else if (Denom == 100) { DenomString = "(In Hundreds)"; }

                else if (Denom == 1000) { DenomString = "(In Thousands)"; }
            }
            else if (Language == 1)
            {
                if (Denom == 10) { DenomString = "(En Décimo)"; }

                else if (Denom == 100) { DenomString = "(En Centenares)"; }

                else if (Denom == 1000) { DenomString = "(En Miles)"; }
            }
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();
        // Get fiscal date
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);
        
        if (Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"]) >= 10)
        {
            fiscalDate = (int.Parse(Year) - 1) + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"];
            d1 = Convert.ToDateTime(fiscalDate);
        }
        else
        {
            fiscalDate = (int.Parse(Year) - 1) + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"];
            d1 = Convert.ToDateTime(fiscalDate);
        }

        date1 = fiscalDate;

        Qua_1_StartDate = d1.ToString("yyyy-MM-dd");
        Qua_1_EndDate = (d1.AddMonths(3).AddDays(-1)).ToString();

        Qua_2_StartDate = (d1.AddMonths(3)).ToString("yyyy-MM-dd");
        Qua_2_EndDate = (d1.AddMonths(6).AddDays(-1)).ToString();

        Qua_3_StartDate = (d1.AddMonths(6)).ToString("yyyy-MM-dd");
        Qua_3_EndDate = (d1.AddMonths(9).AddDays(-1)).ToString();

        Qua_4_StartDate = (d1.AddMonths(9)).ToString("yyyy-MM-dd");
        Qua_4_EndDate = (d1.AddMonths(12).AddDays(-1)).ToString();

        Asterix = "";
        DateTime now = DateTime.Now;
        //Check if the quarter picked is today's quarter
        //Check the year first { check the month compare to fiscal
        if ((int.Parse(Year) == DateTime.Now.Year && DateTime.Now.Month < Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"])) || (int.Parse(Year) == (DateTime.Now.Year - 1) && DateTime.Now.Month >= Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"])))
        {
            // Check the month to see if it's today's quarter got selected
            // Need to Mark which quarter is today's
            if (DateTime.Today >= d1 && now <= d1.AddMonths(3).AddDays(-1) && Qua_1 == "on")
            {
                //It's in Q1
                Qua_1_EndDate = now.ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(3) && now <= d1.AddMonths(6).AddDays(-1) && Qua_2 == "on")
            {
                //It's in Q2
                Qua_2_EndDate = now.ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(6) && now <= d1.AddMonths(9).AddDays(-1) && Qua_3 == "on")
            {
                //It's in Q3
                Qua_3_EndDate = now.ToString("yyyy-MM-dd");
                Qua_2_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-6).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
            else if (DateTime.Today >= d1.AddMonths(9) && now <= d1.AddMonths(12).AddDays(-1) && Qua_4 == "on")
            {
                //It's in Q4
                Qua_4_EndDate = now.ToString("yyyy-MM-dd");
                Qua_3_EndDate = now.AddMonths(-3).ToString("yyyy-MM-dd");
                Qua_2_EndDate = now.AddMonths(-6).ToString("yyyy-MM-dd");
                Qua_1_EndDate = now.AddMonths(-9).ToString("yyyy-MM-dd");
                Asterix = "(*)";
            }
        }

        if (Qua_1 == "on")
        {
            if (Language == 0)
            {
                Quarter[0] = "Q-1";
            }
            else if (Language == 1)
            {
                Quarter[0] = "T-1";
            }

            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='PINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='PINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='PC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "' and Doc_Type='PC') as TaxC" + Q.ToString();

            seconDate = Qua_1_EndDate;
            startDate = Qua_1_StartDate;
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_1_StartDate + "' and '" + Qua_1_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }

        }
        if (Qua_2 == "on")
        {
            if (Language == 0)
            {
                Quarter[1] = "Q-2";
            }
            else if (Language == 1)
            {
                Quarter[1] = "T-2";
            }

            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='PINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='PINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='PC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "' and Doc_Type='PC') as TaxC" + Q.ToString();

            seconDate = Qua_2_EndDate;
            if (Q == 0)
            {
                startDate = Qua_2_StartDate;
            }
            Q += 1;
            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_2_StartDate + "' and '" + Qua_2_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }
        }
        if (Qua_3 == "on")
        {
            if (Language == 0)
            {
                Quarter[2] = "Q-3";
            }
            else if (Language == 1)
            {
                Quarter[2] = "T-3";
            }
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='PINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='PINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='PC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "' and Doc_Type='PC') as TaxC" + Q.ToString();
            seconDate = Qua_3_EndDate;
            if (Q == 0)
            {
                startDate = Qua_3_StartDate;
            }
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_3_StartDate + "' and '" + Qua_3_EndDate + "'";
            if (Q < Counter)
            {
                QueryDate = QueryDate + " OR ";
            }
        }
        if (Qua_4 == "on")
        {

            if (Language == 0)
            {
                Quarter[3] = "Q-4";
            }
            else if (Language == 1)
            {
                Quarter[3] = "T-4";
            }
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='PINV') as Total" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='PINV') as Tax" + Q.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='PC') as TotalC" + Q.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "' and Doc_Type='PC') as TaxC" + Q.ToString();
            seconDate = Qua_4_EndDate;
            if (Q == 0)
            {
                startDate = Qua_4_StartDate;
            }
            Q += 1;

            QueryDate = QueryDate + "piv.Doc_Date between '" + Qua_4_StartDate + "' and '" + Qua_4_EndDate + "'";
        }

        double[] total = new double[Q + 1];
        double percentage = 0;
        
        // Header
        String H_Quarter, H_Qua_1 = "";
        int Temp = 0;
        for (int l = 0; l <= 3; l++)
        {
            if (Quarter[l] != null)
            {
                H_Quarter = Quarter[l] + Asterix;
                StyleMonth = StyleMonth + "~Text-align: Right; font-size:8pt~" + H_Quarter;
                if (Temp < (Q - 1))
                {
                    if (Temp < (Q - 2))
                    {
                        H_Qua_1 = H_Qua_1 + Quarter[l] + ", ";
                    }
                    else
                    {
                        H_Qua_1 = H_Qua_1 + Quarter[l];
                    }
                }
                else if (Temp == (Q - 1))
                {
                    if (Language == 0)
                    {
                        H_Qua_1 = H_Qua_1 + " and " + Quarter[l];
                    }
                    else if (Language == 1)
                    {
                        H_Qua_1 = H_Qua_1 + " y " + Quarter[l];
                    }

                }
                Temp += 1;
            }
        }
        
        //Translate the Header and Title
        if (Language == 0)
        {
            // Percentage.
            if (Request.Form["Percentage"] == "on")
                percent = "~Text-align:right; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right;font-size:0pt~";

            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; width:120px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Purchases(Quarterly) " + DenomString + "<br/>For " + H_Qua_1 + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            if (Request.Form["Percentage"] == "on")
                percent = "~Text-align:right; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; font-size:0pt~";

            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Compras de Varios Períodos (Trimestral)" + DenomString + "<br/>Para " + H_Qua_1 + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        // Get the queryf
        if (Language == 0)
        {

            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(purchases);
        }
        else if (Language == 1)
        {
            SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
            SQLCommand.Parameters.Clear();
            SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
            DataAdapter.Fill(purchases);
        }

        // Add the subtotal column.
        for (int i = 0; i < Q; i++)
        {
            purchases.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }
        // Add percentage column.
        purchases.Columns.Add("Percentage", typeof(String));

        purchases.AcceptChanges();

        // Rounding Calculation

        // Calculating Total and Sub-Total
        for (int i = 0; i <= (Q - 1); i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(purchases.Rows[ii]["Total" + i.ToString()]))
                    purchases.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["Tax" + i.ToString()]))
                    purchases.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TotalC" + i.ToString()]))
                    purchases.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TaxC" + i.ToString()]))
                    purchases.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(purchases.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["TaxC" + i.ToString()]));

                // Denomination Calculation.
                if (Denom > 1)
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]) / Denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]);

            }
        }

        purchases.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && Q > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < purchases.Rows.Count; i++)
            {
                if ((double.Parse(purchases.Rows[i]["SubTotal" + (Q - 1).ToString()].ToString()) != 0) && (double.Parse(purchases.Rows[i]["SubTotal0"].ToString()) != 0))
                    purchases.Rows[i]["Percentage"] = ((double.Parse(purchases.Rows[i]["SubTotal" + (Q - 1).ToString()].ToString()) - double.Parse(purchases.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(purchases.Rows[i]["SubTotal0"].ToString())));
            }
            percentage = ((total[Q - 1] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; width: 80px;";
            purchases.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i < Q; i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    if (!((purchases.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        purchases.Rows[ii]["SubTotal" + i.ToString()] = "";
                }
                else
                {
                    if (!((purchases.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        purchases.Rows[ii]["SubTotal" + i.ToString()] = "";
                }

                // Negative value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() != "")
                {
                    if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                        purchases.Rows[ii]["SubTotal" + i.ToString()] = purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";
                }

                // 0 value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }

        purchases.AcceptChanges();

        // Percentage Format
        if (Request.Form["Percentage"] == "on")
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                if (purchases.Rows[ii]["Percentage"].ToString() != "")
                    purchases.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(purchases.Rows[ii]["Percentage"]));
                if (purchases.Rows[ii]["Percentage"].ToString() != "" && purchases.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    purchases.Rows[ii]["Percentage"] = purchases.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        purchases.AcceptChanges();

        // Post on report DataTable
        for (int i = 0; i <= 15; i++)
        {
            report.Columns.Add("Style" + i.ToString(), typeof(String));
            report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width:100px;";
        StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black;";

        for (int i = 0; i < purchases.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && Q > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (purchases.Rows[i]["Percentage"].ToString() != "" && purchases.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (Q == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 4in; max-width: 4.5in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, "<span style= + StyleFinish + >" + purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3.5in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], balanceStyle, purchases.Rows[i]["Subtotal2"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";

        // Percent formatting
        if (Request.Form["Percentage"] == "on" && Q > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (Q == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (Q == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        }
        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            purchases.Columns.Remove("Currency");
            for (int i = 0; i < Q; i++)
            {
                purchases.Columns.Remove("Total" + i.ToString());
                purchases.Columns.Remove("Tax" + i.ToString());
                purchases.Columns.Remove("TotalC" + i.ToString());
                purchases.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < purchases.Columns.Count; i++)
            {
                purchases.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = purchases.Columns.Count; i < 25; i++)
            {
                purchases.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = purchases.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i < Q; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = Quarter[i] + Asterix;
                if (i == Q - 1 && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            purchases.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < purchases.Columns.Count - 1; i++)
            {
                purchases.Rows[0][i] = "<b>" + purchases.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = purchases.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i < Q; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == Q - 1 && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = String.Format("{0:P}", percentage);
            }
            purchases.Rows.Add(excelTotal);

            RPT_Excel.DataSource = purchases;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Purchases Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterPurchasesMultiPer()
    {
        int Language = int.Parse(Request.Form["language"]);
        string seconDate = Request.Form["Quarter"];
        string[] words = seconDate.Split(' ');
        int Qua_No = int.Parse(words[0]);
        string Qua_Year = words[1].ToString();
        string Goback = Request.Form["goback"];
        string Show_Per = Request.Form["Percentage"];
        int Denom = int.Parse(Request.Form["Denom"]);
        string[] Quarter = new string[4];
        string asterix = "";

        DataTable fiscal = new DataTable();
        DataTable COA = new DataTable();
        DataTable Report = new DataTable();
        string StyleFinish = "border-bottom: Double 3px black; border-top: 1px solid black";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        String DenomString = "", StyleMonth = "";
        String QueryDate = "";
        if (Denom > 1)
        {
            if (Language == 0)
            {
                if (Denom == 10) { DenomString = "(In Tenth)"; }

                else if (Denom == 100) { DenomString = "(In Hundreds)"; }

                else if (Denom == 1000) { DenomString = "(In Thousands)"; }
            }
            else if (Language == 1)
            {
                if (Denom == 10) { DenomString = "(En D�cimo)"; }

                else if (Denom == 100) { DenomString = "(En Centenares)"; }

                else if (Denom == 1000) { DenomString = "(En Miles)"; }
            }
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;
        Conn.Open();

        string balanceStyle = "", Query = "", percent = "", startDate2 = "", endDate2 = "", fiscalDate = "", Qdate1 = "", Qdate2 = "", Date1 = "";
        DateTime d1;
        double[] total = new double[(int.Parse(Goback)) + 1];
        double percentage = 0;

        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);


        if (Convert.ToInt32(fiscal.Rows[0]["Fiscal_Year_Start_Month"]) >= 10)
        {
            fiscalDate = (Qua_Year + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"]);
            d1 = Convert.ToDateTime(fiscalDate);
        }
        else
        {
            fiscalDate = (Qua_Year + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"]);
            d1 = Convert.ToDateTime(fiscalDate);
        }
        if (Qua_No == 1)
        {
            Qdate1 = d1.ToString();
            Qdate2 = (d1.AddMonths(3).AddDays(-1)).ToString();
        }
        else if (Qua_No == 2)
        {
            Qdate1 = (d1.AddMonths(3)).ToString();
            Qdate2 = (d1.AddMonths(6).AddDays(-1)).ToString();
        }
        else if (Qua_No == 3)
        {
            Qdate1 = (d1.AddMonths(6)).ToString();
            Qdate2 = (d1.AddMonths(9).AddDays(-1)).ToString();
        }
        else if (Qua_No == 4)
        {
            Qdate1 = (d1.AddMonths(9)).ToString();
            Qdate2 = (d1.AddMonths(12).AddDays(-1)).ToString();
        }
        DateTime S_Date = Convert.ToDateTime(Qdate1);
        DateTime E_Date = Convert.ToDateTime(Qdate2);

        // Check the year first then check the month compare to fiscal+

        DateTime now = DateTime.Now;
        //Check if the quarter picked is today's quarter
        //Check the year first { check the month compare to fiscal
        if (S_Date < now && E_Date > now)
        {
            // Check the month to see if it's today's quarter got selected
            // Need to Mark which quarter is today's
            if (DateTime.Today >= d1 && now <= d1.AddMonths(3).AddDays(-1))
            {
                //It's in Q1
                Qdate1 = d1.ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
            else if (DateTime.Today >= d1.AddMonths(3) && now <= d1.AddMonths(6).AddDays(-1))
            {
                //It's in Q2
                Qdate1 = (d1.AddMonths(3)).ToString();
                Qdate2 = (now.ToString("yyyy-MM-dd")).ToString();
            }
            else if (DateTime.Today >= d1.AddMonths(6) && now <= d1.AddMonths(9).AddDays(-1))
            {
                //It's in Q3
                Qdate1 = (d1.AddMonths(6)).ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
            else if (DateTime.Today >= d1.AddMonths(9) && now <= d1.AddMonths(12).AddDays(-1))
            {
                //It's in Q4
                Qdate1 = (d1.AddMonths(9)).ToString();
                Qdate2 = now.ToString("yyyy-MM-dd");
            }
        }
        S_Date = Convert.ToDateTime(Qdate1);
        E_Date = Convert.ToDateTime(Qdate2);

        int j = 0;
        DateTime firstDate = S_Date.AddYears(-(int.Parse(Goback) - 1));
        DateTime secondDate = E_Date.AddYears(-(int.Parse(Goback) - 1));
        Qdate1 = ""; Qdate2 = "";
        for (j = 0; j < (int.Parse(Goback)); j++)
        {
            if (Language == 0)
            {
                if (S_Date < now && E_Date > now)
                {
                    Qdate1 = ("Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2)) + "(*)").ToString();
                    asterix = "(*)";
                }                    
                else
                    Qdate1 = "Q" + Qua_No + " " + ((int.Parse(Qua_Year)+j) - (int.Parse(Goback) - 2));

                if (j == ((int.Parse(Goback)) - 1))
                    Qdate2 = Qdate2 + " and Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2));
                else if (j == ((int.Parse(Goback)) - 2))
                    Qdate2 = Qdate2 + "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2));
                else
                    Qdate2 = Qdate2 + "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + j) - (int.Parse(Goback) - 2)) + ", ";

            }
            if (Language == 1)
            {
                if (S_Date < now && E_Date >= now)
                {
                    Qdate1 = "T" + Qua_No + " " + firstDate.ToString("yyyy") + "(*)";
                    asterix = "(*)";
                }                    
                else
                    Qdate1 = "T" + Qua_No + " " + firstDate.ToString("yyyy");

                if (j == ((int.Parse(Goback)) - 1))
                    Qdate2 = Qdate2 + " y T" + Qua_No + " " + firstDate.ToString("yyyy");
                else if (j == ((int.Parse(Goback)) - 2))
                    Qdate2 = Qdate2 + "T" + Qua_No + " " + firstDate.ToString("yyyy");
                else
                    Qdate2 = Qdate2 + "T" + Qua_No + " " + firstDate.ToString("yyyy") + ", ";
            }

            StyleMonth = StyleMonth + "~Text-align: Right;font-size:8pt~" + Qdate1;

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd");
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd");
            firstDate = Convert.ToDateTime(startDate2);
            secondDate = Convert.ToDateTime(endDate2);
        }

        //Translate the Header and Title
        if (Language == 0)
        {
            // Percentage.
            if (Show_Per == "on")
                percent = "~Text-align:right; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; font-size:0pt~";

            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; width:120px; font-size:8pt~Account Description" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Purchase(Quarter to Quarter) " + DenomString + "<br/>For " + Qdate2 + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }
        else if (Language == 1)
        {
            if (Show_Per == "on")
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + StyleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Compra Multiperiodos(Trimestre a Trimestre)" + DenomString + "<br/>Para " + Qdate2 + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";
        }

        SQLCommand.Connection = Conn;
        DataAdapter.SelectCommand = SQLCommand;

        firstDate = S_Date.AddYears(-(int.Parse(Goback) - 1));
        secondDate = E_Date.AddYears(-(int.Parse(Goback) - 1));
        startDate2 = firstDate.ToString("yyyy-MM-dd");
        endDate2 = secondDate.ToString("yyyy-MM-dd");

        for (j = 0; j < (int.Parse(Goback)); j++)
        {
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='PINV') as Total" + j.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='PINV') as Tax" + j.ToString();
            Query = Query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='PC') as TotalC" + j.ToString();
            Query = Query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate2 + "' and '" + endDate2 + "' and Doc_Type='PC') as TaxC" + j.ToString();


            QueryDate = QueryDate + "piv.Doc_Date between '" + startDate2 + "' and '" + endDate2 + "'";
            if (j < int.Parse(Goback)-1)
                 QueryDate = QueryDate + " OR ";

            startDate2 = firstDate.AddYears(1).ToString("yyyy-MM-dd");
            endDate2 = secondDate.AddYears(1).ToString("yyyy-MM-dd");
            firstDate = Convert.ToDateTime(startDate2);
            secondDate = Convert.ToDateTime(endDate2);

        }
        
        //SQL Command for the queries
        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + Query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + QueryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(COA);
        
        // Adding subtotal column
        for (int i = 0; i < int.Parse(Goback); i++)
        {
            COA.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }

        // Add percentage column.
        COA.Columns.Add("Percentage", typeof(String));

        COA.AcceptChanges();

        // Calculating Total and Sub-Total
        for (int i = 0; i < int.Parse(Goback); i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(COA.Rows[ii]["Total" + i.ToString()]))
                    COA.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["Tax" + i.ToString()]))
                    COA.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TotalC" + i.ToString()]))
                    COA.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(COA.Rows[ii]["TaxC" + i.ToString()]))
                    COA.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(COA.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(COA.Rows[ii]["TaxC" + i.ToString()]));
                
                // Denomination Calculation.
                if (Denom > 1)
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]) / Denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    COA.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]);


            }
        }
        COA.AcceptChanges();

        // Percentage.
        if (Show_Per == "on")
        {
            // Calculate Percentage.
            for (int i = 0; i < COA.Rows.Count; i++)
            {
                if ((double.Parse(COA.Rows[i]["SubTotal" + (int.Parse(Goback) - 1).ToString()].ToString()) != 0) && (double.Parse(COA.Rows[i]["SubTotal0"].ToString()) != 0))
                    COA.Rows[i]["Percentage"] = ((double.Parse(COA.Rows[i]["SubTotal" + (int.Parse(Goback) - 1).ToString()].ToString()) - double.Parse(COA.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(COA.Rows[i]["SubTotal0"].ToString())));
            }
            percentage = ((total[(int.Parse(Goback) - 1)] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            COA.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i <= (int.Parse(Goback) - 1); i++)
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                }
                else
                {
                    if (!((COA.Rows[ii]["SubTotal" + i.ToString()]) is DBNull))
                        COA.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(COA.Rows[ii]["SubTotal" + i.ToString()]));
                    else
                        COA.Rows[ii]["SubTotal" + i.ToString()] = "";
                }

                // Negative value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() != "")
                {
                    if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                        COA.Rows[ii]["SubTotal" + i.ToString()] = COA.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";
                }
                // 0 value.
                if (COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || COA.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    COA.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }
        if (Show_Per == "on")
        {
            for (int ii = 0; ii < COA.Rows.Count; ii++)
            {
                if (COA.Rows[ii]["Percentage"].ToString() != "")
                    COA.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(COA.Rows[ii]["Percentage"]));
                if (COA.Rows[ii]["Percentage"].ToString() != "" && COA.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    COA.Rows[ii]["Percentage"] = COA.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        COA.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            Report.Columns.Add("Style" + i.ToString(), typeof(String));
            Report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
        
        for (int i = 0; i < COA.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && (int.Parse(Goback) - 1) > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (COA.Rows[i]["Percentage"].ToString() != "" && COA.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", COA.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 3in; max-width: 3in;", COA.Rows[i]["Name"].ToString(), balanceStyle, COA.Rows[i]["SubTotal0"], balanceStyle, COA.Rows[i]["SubTotal1"], balanceStyle, COA.Rows[i]["Subtotal2"], percentStyle, COA.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && (int.Parse(Goback) - 1) > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (int.Parse(Goback) == 2)
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                Report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style= \"" + StyleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])) + "</span>", percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        }
        
        RPT_PrintReports.DataSource = Report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            COA.Columns.Remove("Currency");
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                COA.Columns.Remove("Total" + i.ToString());
                COA.Columns.Remove("Tax" + i.ToString());
                COA.Columns.Remove("TotalC" + i.ToString());
                COA.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < COA.Columns.Count; i++)
            {
                COA.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = COA.Columns.Count; i < 25; i++)
            {
                COA.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = COA.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                excelHeader["value" + (i + 2).ToString()] = "Q" + Qua_No + " " + ((int.Parse(Qua_Year) + i) - (int.Parse(Goback) - 2)) + asterix;
                if (i == int.Parse(Goback) - 1 && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            COA.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < COA.Columns.Count - 1; i++)
            {
                COA.Rows[0][i] = "<b>" + COA.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = COA.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i < int.Parse(Goback); i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == (int.Parse(Goback) - 1) && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = String.Format("{0:P}", percentage);
            }
            COA.Rows.Add(excelTotal);

            RPT_Excel.DataSource = COA;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // Purchases Multiperiod Yearly
    private void PrintYearlyPurchasesMultiPer()
    {
        string language = Request.Form["language"];
        string firstDate = Request.Form["FirstDate"];
        string secondDate = Request.Form["SecondDate"];
        int denom = Int32.Parse(Request.Form["Denom"]);

        int yearCount = Int32.Parse(Request.Form["SecondDate"]) - Int32.Parse(Request.Form["FirstDate"]); ;
        DateTime[] startDate = new DateTime[yearCount + 1];
        DateTime[] endDate = new DateTime[yearCount + 1];
        DateTime fiscalDate;
        string asterix = "";
        string denomString = "";
        string dateRange = "";
        string styleMonth = "";
        string query = "";
        string queryDate = "";
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";
        string styleFinish = "font-weight:bold; border-bottom: Double 3px black; border-top: 1px solid black;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable purchases = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get fiscal date.
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info";
        SQLCommand.Parameters.Clear();
        DataAdapter.Fill(fiscal);

        // Because it is '9' not '09'.
        if (Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()) < 10)
            fiscalDate = DateTime.Parse(firstDate + "-0" + fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString() + "-01");
        else
            fiscalDate = DateTime.Parse(firstDate + "-" + fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString() + "-01");

        // If date picked is current year, check if today's month is > fiscal month.
        if (Int32.Parse(secondDate) >= DateTime.Now.Year)
        {
            // Check if today's date already pass the fiscal month.
            //  If not, use today's date to compare with previous year.
            if (DateTime.Now.Month < Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()))
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = new DateTime(DateTime.Today.AddYears(i - yearCount - 1).Year, Int32.Parse(fiscal.Rows[0]["Fiscal_Year_Start_Month"].ToString()), 1);
                    endDate[i] = DateTime.Today.AddYears(i - yearCount);
                    asterix = "(*)";
                }
            }
            else
            {
                for (int i = 0; i <= yearCount; i++)
                {
                    startDate[i] = fiscalDate.AddYears(i - 1);
                    endDate[i] = fiscalDate.AddYears(i).AddDays(-1);
                }
            }
        }
        else
        {
            for (int i = 0; i <= yearCount; i++)
            {
                startDate[i] = fiscalDate.AddYears(i - 1);
                endDate[i] = fiscalDate.AddYears(i).AddDays(-1);
            }
        }

        // language 0 is english, language 1 is spanish
        if (language == "0")
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(In Tenth)";
                        break;
                    case 100:
                        denomString = "(In Hunreds)";
                        break;
                    case 1000:
                        denomString = "(In Thousands)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " and " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Percentage(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "text-align:left; font-size:8pt; width:60px;~ID~text-align:left; font-size:8pt~Account Description" + styleMonth + percent;
            //HF_PrintHeader.Value = "text-align:left; width:80px; font-size:9pt; font-weight:bold~No~text-align:left; font-size:9pt; font-weight:bold~Customer Name~text-align:right; font-size:9pt; width:180px; font-weight:bold~Net Sales ($)";
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Multiperiod Sales(Yearly) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }
        else
        {
            // Denomination.
            if (denom > 1)
            {
                switch (denom)
                {
                    case 10:
                        denomString = "(En Décimo)";
                        break;
                    case 100:
                        denomString = "(En Centenares)";
                        break;
                    case 1000:
                        denomString = "(En Miles)";
                        break;
                    default:
                        break;
                }
            }

            // Title and Header.
            for (int i = 0; i <= yearCount; i++)
            {
                // For Title.
                if (i == yearCount)
                    dateRange = dateRange + " y " + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else if (i == yearCount - 1)
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy");
                else
                    dateRange = dateRange + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + ", ";

                // For Header.
                styleMonth = styleMonth + "~Text-align:right; width:120px; font-size:8pt~" + endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
            }

            // Percentage.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
                percent = "~Text-align:right; width:80px; font-size:8pt~Porcentaje(%)";
            else
                percent = "~Text-align:right; width:0px; font-size:0pt~";

            // Header.
            HF_PrintHeader.Value = "Text-align:left; font-size:8pt; width:60px; ~ID~ text-align:left; font-size:8pt~Nombre De La Cuenta" + styleMonth + percent;
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">" + fiscal.Rows[0]["Company_Name"].ToString() + "<br/>Ventas de Varios Períodos (Anuales)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TaxC" + i.ToString();

            queryDate = queryDate + "piv.Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "'";
            if (i < yearCount)
                queryDate = queryDate + " OR ";
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND (" + queryDate + ")";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@cur", Request.Form["cur"]);
        DataAdapter.Fill(purchases);

        // Add the subtotal column.
        for (int i = 0; i <= yearCount; i++)
        {
            purchases.Columns.Add("SubTotal" + i.ToString(), typeof(String));
        }
        // Add percentage column.
        purchases.Columns.Add("Percentage", typeof(String));

        purchases.AcceptChanges();

        // Rounding Calculation

        // Calculating Total and Sub-Total
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Replacing null value with 0's.
                if (Convert.IsDBNull(purchases.Rows[ii]["Total" + i.ToString()]))
                    purchases.Rows[ii]["Total" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["Tax" + i.ToString()]))
                    purchases.Rows[ii]["Tax" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TotalC" + i.ToString()]))
                    purchases.Rows[ii]["TotalC" + i.ToString()] = 0;
                if (Convert.IsDBNull(purchases.Rows[ii]["TaxC" + i.ToString()]))
                    purchases.Rows[ii]["TaxC" + i.ToString()] = 0;

                // Calculating SubTotal.
                purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(purchases.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["TaxC" + i.ToString()]));

                // Denomination Calculation.
                if (denom > 1)
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]) / denom;
                }

                // Rounding.
                if (Request.Form["Round"] == "on")
                {
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = Math.Round(Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                }

                // Calculating Total.
                total[i] += Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]);
            }
        }

        purchases.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < purchases.Rows.Count; i++)
            {
                if ((double.Parse(purchases.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) != 0) && (double.Parse(purchases.Rows[i]["SubTotal0"].ToString()) != 0))
                {
                    purchases.Rows[i]["Percentage"] = ((double.Parse(purchases.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(purchases.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(purchases.Rows[i]["SubTotal0"].ToString())));
                }
            }
            percentage = ((total[yearCount] - total[0]) / Math.Abs(total[0]));
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
            purchases.AcceptChanges();
        }

        // Format all the output for the paper.
        for (int i = 0; i <= yearCount; i++)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                // Rounding.
                if (Request.Form["Round"] == "off")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C2}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));
                else
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = String.Format("{0:C0}", Convert.ToDouble(purchases.Rows[ii]["SubTotal" + i.ToString()]));

                // Negative value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Substring(0, 1) == "-")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = purchases.Rows[ii]["SubTotal" + i.ToString()].ToString().Replace("-", "(") + ")";

                // 0 value.
                if (purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0.00" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$" || purchases.Rows[ii]["SubTotal" + i.ToString()].ToString() == "$0")
                    purchases.Rows[ii]["SubTotal" + i.ToString()] = "";
            }
        }

        // Formatting Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            for (int ii = 0; ii < purchases.Rows.Count; ii++)
            {
                if (purchases.Rows[ii]["Percentage"].ToString() != "")
                    purchases.Rows[ii]["Percentage"] = String.Format("{0:P}", Convert.ToDouble(purchases.Rows[ii]["Percentage"]));
                if (purchases.Rows[ii]["Percentage"].ToString() != "" && purchases.Rows[ii]["Percentage"].ToString().Substring(0, 1) == "-")
                    purchases.Rows[ii]["Percentage"] = purchases.Rows[ii]["Percentage"].ToString().Replace("-", "(") + ")";
            }
        }

        purchases.AcceptChanges();

        // Post on Report DataTable
        for (int i = 0; i <= 15; i++)
        {
            report.Columns.Add("Style" + i.ToString(), typeof(String));
            report.Columns.Add("Field" + i.ToString(), typeof(String));
        }

        balanceStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";

        for (int i = 0; i < purchases.Rows.Count; i++)
        {
            // Percentage format if negatives.
            if (Request.Form["Percentage"] == "on" && yearCount > 0)
            {
                percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px;";
                if (purchases.Rows[i]["Percentage"].ToString() != "" && purchases.Rows[i]["Percentage"].ToString().Substring(0, 1) == "(")
                    percentStyle = percentStyle + "color: red;";
            }

            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], balanceStyle, purchases.Rows[i]["Subtotal2"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        // Percentage format if negatives.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:8pt; min-width: 5px; max-width: 5px; font-weight:bold;";
            if (percentage.ToString() != "" && percentage.ToString().Substring(0, 1) == "-")
                percentStyle = percentStyle + "color: red;";
        }

        if (Request.Form["Round"] == "on")
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C0}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        else
        {
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[0])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[1])) + "</span>", balanceStyle, "<span style=\"" + styleFinish + "\">" + String.Format("{0:C2}", Convert.ToDouble(total[2])) + "</span>", percentStyle, percentage.ToString("#,##0.00%;(#,##0.00%)"), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }
        

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;

        // Export Function
        if (Request.Form["expStat"] == "on")
        {
            // Remove columns that do not need to be display in excel
            purchases.Columns.Remove("Currency");
            for (int i = 0; i <= yearCount; i++)
            {
                purchases.Columns.Remove("Total" + i.ToString());
                purchases.Columns.Remove("Tax" + i.ToString());
                purchases.Columns.Remove("TotalC" + i.ToString());
                purchases.Columns.Remove("TaxC" + i.ToString());
            }

            // Rename the existing column name to value
            for (int i = 0; i < purchases.Columns.Count; i++)
            {
                purchases.Columns[i].ColumnName = "value" + i.ToString();
            }

            // Creating new column to value20
            for (int i = purchases.Columns.Count; i < 25; i++)
            {
                purchases.Columns.Add("value" + i.ToString(), typeof(String));
            }

            // Add the header as "value"
            DataRow excelHeader = purchases.NewRow();
            excelHeader["value0"] = "Vendor ID";
            excelHeader["value1"] = "Vendor Name";

            // Add the header with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelHeader["value" + (i + 2).ToString()] = endDate[i].AddYears(-1).ToString("yyyy") + "-" + endDate[i].ToString("yyyy") + asterix;
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelHeader["value" + (i + 3).ToString()] = "Growth(%)";
            }
            purchases.Rows.InsertAt(excelHeader, 0);

            // Bold the header.
            for (int i = 0; i < purchases.Columns.Count - 1; i++)
            {
                purchases.Rows[0][i] = "<b>" + purchases.Rows[0][i].ToString() + "</b>";
            }

            // Add the total
            DataRow excelTotal = purchases.NewRow();
            excelTotal["value1"] = "Total";
            // Add the total with dynamic number of columns
            for (int i = 0; i <= yearCount; i++)
            {
                excelTotal["value" + (i + 2).ToString()] = String.Format("{0:C2}", Convert.ToDouble(total[i]));
                if (i == yearCount && Request.Form["Percentage"] == "on")
                    excelTotal["value" + (i + 3).ToString()] = percentage.ToString("#,##0.00%;(#,##0.00%)");
            }
            purchases.Rows.Add(excelTotal);

            RPT_Excel.DataSource = purchases;
            RPT_Excel.DataBind();

            PNL_Excel.Visible = true;
        }
    }

    // AR/AP Summary Trial Balance and Detailed Trial Balance
    private void ExportAR()
    {
        DataTable List = new DataTable();
        string curr = Request.Form["currency"];
        string date = Request.Form["date"];
        string cust = Request.Form["cust"];
        string ARType = Request.Form["type"];
        string where = "Document_Date <=@date and Applies_To <>'0' and Currency_ID = @currency and (Select top 1 Balance_At_Date from ACC_AR where Applies_To = AR1.Applies_To and Document_Date <=@date order by Document_Date desc, AR_ID desc)<>0 ";
        if (ARType == "Details") { if (cust != "all") { where = where + " and AR1.Cust_Vend_ID = '" + cust + "' "; } }

        SQLCommand.CommandText = @"Select distinct(Applies_to) as Doc_ID, 
            (Select top 1 Balance_At_Date from ACC_AR where Applies_To = AR1.Applies_To and Document_Date <=@date order by Document_Date desc, AR_ID desc) as Balance_At_Date, 
            (Select top 1 Document_Date from ACC_AR where Document_ID = AR1.Applies_To and Document_Date <=@date order by Document_Date, AR_ID) as Document_Date,
            Doc_No, AR1.Cust_Vend_ID, Name from ACC_AR AR1 left join ACC_SalesInv on Applies_to = Doc_ID left join Customer on AR1.Cust_Vend_ID = Cust_ID where " + where + " order by Name, Document_Date, Doc_No";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@currency", curr);
        SQLCommand.Parameters.AddWithValue("@date", date);
        DataAdapter.Fill(List);

        SQLCommand.CommandText = @"Select Document_ID as Doc_ID, Balance_At_Date, Document_Date, Doc_No, ACC_AR.Cust_Vend_ID, Name from ACC_AR 
            left join ACC_SalesInv on Document_ID=Doc_ID 
            left join Customer on ACC_AR.Cust_Vend_ID = Cust_ID 
            where Applies_to = 0 and Balance_At_Date<>0 and Document_Date <=@date and Currency_ID=@currency ";
        DataAdapter.Fill(List);

        List.Columns.Add("Total", typeof(Double));
        List.Columns.Add("Current", typeof(Double));
        List.Columns.Add("AP30", typeof(Double));
        List.Columns.Add("AP60", typeof(Double));
        List.Columns.Add("AP90", typeof(Double));
        List.Columns.Add("Details", typeof(Double));
        List.Columns.Add("Padding", typeof(string));
        List.Columns.Add("Age", typeof(Int16));

        double Total = 0;
        double current = 0;
        double AP30 = 0;
        double AP60 = 0;
        double AP90 = 0;

        DateTime TempDate = new DateTime();
        DateTime BalDate = new DateTime();
        BalDate = Convert.ToDateTime(date);
        for (int i = 0; i < List.Rows.Count; i++)
        {
            TempDate = Convert.ToDateTime(List.Rows[i]["Document_Date"].ToString());
            double balance;
            if ((double.TryParse(List.Rows[i]["Balance_At_Date"].ToString(), out balance)))
            {
                List.Rows[i]["Total"] = List.Rows[i]["Balance_At_Date"];
            }
            else
            {
                List.Rows[i]["Total"] = "0";
            }

            TimeSpan age = BalDate - TempDate;
            List.Rows[i]["Age"] = age.TotalDays;
            List.Rows[i]["Current"] = "0";
            List.Rows[i]["AP30"] = "0";
            List.Rows[i]["AP60"] = "0";
            List.Rows[i]["AP90"] = "0";

            Total += Convert.ToDouble(List.Rows[i]["Total"].ToString());

            if (TempDate > BalDate.AddDays(-31))
            {
                List.Rows[i]["Current"] = Convert.ToDouble(balance);
                current += Convert.ToDouble(balance);
            }
            else if (TempDate > BalDate.AddDays(-61))
            {
                List.Rows[i]["AP30"] = Convert.ToDouble(balance);
                AP30 += Convert.ToDouble(balance);
            }
            else if (TempDate > BalDate.AddDays(-90))
            {
                List.Rows[i]["AP60"] = Convert.ToDouble(balance);
                AP60 += Convert.ToDouble(balance);
            }
            else
            {
                List.Rows[i]["AP90"] = Convert.ToDouble(balance);
                AP90 += Convert.ToDouble(balance);
            }
        }

        DataTable SummaryList = List.Copy(); //need to pass copy of table so as not to affect list table in the scope of popListDetail

        for (int i = 1; i < SummaryList.Rows.Count; i++)
        {
            if (SummaryList.Rows[i]["Cust_Vend_ID"].ToString() == SummaryList.Rows[i - 1]["Cust_Vend_ID"].ToString())
            {
                SummaryList.Rows[i]["Total"] = Convert.ToDouble(SummaryList.Rows[i]["Total"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["Total"].ToString());
                SummaryList.Rows[i]["AP30"] = Convert.ToDouble(SummaryList.Rows[i]["AP30"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP30"].ToString());
                SummaryList.Rows[i]["AP60"] = Convert.ToDouble(SummaryList.Rows[i]["AP60"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP60"].ToString());
                SummaryList.Rows[i]["AP90"] = Convert.ToDouble(SummaryList.Rows[i]["AP90"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP90"].ToString());
                SummaryList.Rows[i]["Current"] = Convert.ToDouble(SummaryList.Rows[i]["Current"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["Current"].ToString());
                SummaryList.Rows[i - 1].Delete();
            }
        }

        SummaryList.AcceptChanges();

        for (int i = 0; i < SummaryList.Rows.Count; i++)
        {
            for (int ii = i + 1; ii < SummaryList.Rows.Count; ii++)
            {
                if (SummaryList.Rows[i]["Cust_Vend_ID"].ToString() == SummaryList.Rows[ii]["Cust_Vend_ID"].ToString())
                {
                    SummaryList.Rows[i]["Total"] = Convert.ToDouble(SummaryList.Rows[i]["Total"].ToString()) + Convert.ToDouble(SummaryList.Rows[ii]["Total"].ToString());
                    SummaryList.Rows[i]["AP30"] = Convert.ToDouble(SummaryList.Rows[i]["AP30"].ToString()) + Convert.ToDouble(SummaryList.Rows[ii]["AP30"].ToString());
                    SummaryList.Rows[i]["AP60"] = Convert.ToDouble(SummaryList.Rows[i]["AP60"].ToString()) + Convert.ToDouble(SummaryList.Rows[ii]["AP60"].ToString());
                    SummaryList.Rows[i]["AP90"] = Convert.ToDouble(SummaryList.Rows[i]["AP90"].ToString()) + Convert.ToDouble(SummaryList.Rows[ii]["AP90"].ToString());
                    SummaryList.Rows[i]["Current"] = Convert.ToDouble(SummaryList.Rows[i]["Current"].ToString()) + Convert.ToDouble(SummaryList.Rows[ii]["Current"].ToString());
                    SummaryList.Rows[ii]["Cust_Vend_ID"] = "DELETE";
                }
            }
        }

        for (int i = 0; i < SummaryList.Rows.Count; i++) { if (SummaryList.Rows[i]["Cust_Vend_ID"] == "DELETE") { SummaryList.Rows[i].Delete(); } }

        SummaryList.AcceptChanges();

        DataTable Details = new DataTable();
        Details = List.Clone();
        for (int i = 0; i < SummaryList.Rows.Count; i++)
        {
            Details.Rows.Add("0", "0", date, "", "", SummaryList.Rows[i]["Name"], SummaryList.Rows[i]["Total"], SummaryList.Rows[i]["Current"], SummaryList.Rows[i]["AP30"], SummaryList.Rows[i]["AP60"], SummaryList.Rows[i]["AP90"], "0", "totalRow");
            for (int ii = 0; ii < List.Rows.Count; ii++)
            {
                if (List.Rows[ii]["Cust_Vend_ID"].ToString() == SummaryList.Rows[i]["Cust_Vend_ID"].ToString())
                {
                    Details.Rows.Add(List.Rows[ii]["Doc_ID"], "0", List.Rows[ii]["Document_Date"], List.Rows[ii]["Doc_No"], List.Rows[ii]["Cust_Vend_ID"], List.Rows[ii]["Name"], List.Rows[ii]["Total"], List.Rows[ii]["Current"], List.Rows[ii]["AP30"], List.Rows[ii]["AP60"], List.Rows[ii]["AP90"], "0", "20", List.Rows[ii]["Age"]);
                }
            }
        }
        if (ARType == "Details") { if (cust == "all") { Details.Rows.Add("0", "0", date, "", "", "Total", Total, current, AP30, AP60, AP90, "0", "totalRow"); } }

        if (ARType == "Details")
        {
            RPT_List_Details.DataSource = Details;
        }
        else
        {
            SummaryList.Rows.Add("0", "0", date, "", "", "Total", Total, current, AP30, AP60, AP90, "0", "totalRow");
            RPT_List_Details.DataSource = SummaryList;
        }
        RPT_List_Details.DataBind();

        LBL_Total.Text = Total.ToString("0,000.00");
        HF_TotalDet.Value = Total.ToString("0,000.00");

        // Export Function
        if (ARType == "Details")
        {
            if (Request.Form["expStat"] == "on")
            {
                // Remove columns that do not need to be display in excel
                Details.Columns.Remove("Doc_ID");
                Details.Columns.Remove("Balance_At_Date");
                Details.Columns.Remove("Details");
                Details.Columns.Remove("Padding");

                // Change the order of DataTable Columns
                Details.Columns["Cust_Vend_ID"].SetOrdinal(0);
                Details.Columns["Name"].SetOrdinal(1);
                Details.Columns["Age"].SetOrdinal(2);
                Details.Columns["Document_Date"].SetOrdinal(3);
                Details.Columns["Doc_No"].SetOrdinal(4);
                Details.Columns["Total"].SetOrdinal(5);
                Details.Columns["Current"].SetOrdinal(6);
                Details.Columns["AP30"].SetOrdinal(7);
                Details.Columns["AP60"].SetOrdinal(8);
                Details.Columns["AP90"].SetOrdinal(9);

                // Create new Datatable
                DataTable exportTable = new DataTable();

                for (int i = 0; i < Details.Columns.Count; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(string));
                }

                // Copy the data value
                for (int i = 0; i < Details.Rows.Count; i++)
                {
                    DataRow excelRow = exportTable.NewRow();
                    for (int ii = 0; ii < Details.Columns.Count; ii++)
                    {
                        excelRow["value" + ii.ToString()] = Details.Rows[i][ii].ToString();
                    }

                    exportTable.Rows.Add(excelRow);
                }

                // Creating new column to value20
                for (int i = exportTable.Columns.Count; i < 25; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(String));
                }

                // Formatting the numbers.
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    exportTable.Rows[i]["value3"] = exportTable.Rows[i]["value3"].ToString().Substring(0, 10);
                    exportTable.Rows[i]["value5"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value5"]));
                    exportTable.Rows[i]["value6"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value6"]));
                    exportTable.Rows[i]["value7"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value7"]));
                    exportTable.Rows[i]["value8"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value8"]));
                    exportTable.Rows[i]["value9"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value9"]));

                    // Make it empty if value is zero
                    if (exportTable.Rows[i]["value5"].ToString() == "$0.00")
                        exportTable.Rows[i]["value5"] = "";
                    if (exportTable.Rows[i]["value6"].ToString() == "$0.00")
                        exportTable.Rows[i]["value6"] = "";
                    if (exportTable.Rows[i]["value7"].ToString() == "$0.00")
                        exportTable.Rows[i]["value7"] = "";
                    if (exportTable.Rows[i]["value8"].ToString() == "$0.00")
                        exportTable.Rows[i]["value8"] = "";
                    if (exportTable.Rows[i]["value9"].ToString() == "$0.00")
                        exportTable.Rows[i]["value9"] = "";
                }
                exportTable.AcceptChanges();

                // Bold the header
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    if (exportTable.Rows[i]["value0"].ToString() == "")
                    {
                        for (int ii = 0; ii < exportTable.Columns.Count - 1; ii++)
                        {
                            exportTable.Rows[i][ii] = "<b>" + exportTable.Rows[i][ii].ToString() + "</b>";
                        }
                    }
                }

                // Add the header as "value"
                DataRow excelHeader = exportTable.NewRow();
                excelHeader["value0"] = "Cust/Vend ID";
                excelHeader["value1"] = "Cust/Vend Name";
                excelHeader["value2"] = "Age";
                excelHeader["value3"] = "Date";
                excelHeader["value4"] = "Invoice No.";
                excelHeader["value5"] = "Total";
                excelHeader["value6"] = "Current";
                excelHeader["value7"] = "31-60";
                excelHeader["value8"] = "61-90";
                excelHeader["value9"] = "90+";

                exportTable.Rows.InsertAt(excelHeader, 0);

                // Bold the header.
                for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                {
                    exportTable.Rows[0][i] = "<b>" + exportTable.Rows[0][i].ToString() + "</b>";
                }

                RPT_Excel.DataSource = exportTable;
                RPT_Excel.DataBind();

                PNL_Excel.Visible = true;
            }
        }
        else
        {
            if (Request.Form["expStat"] == "on")
            {
                // Remove columns that do not need to be display in excel
                SummaryList.Columns.Remove("Doc_ID");
                SummaryList.Columns.Remove("Balance_At_Date");
                SummaryList.Columns.Remove("Details");
                SummaryList.Columns.Remove("Padding");
                SummaryList.Columns.Remove("Document_Date");
                SummaryList.Columns.Remove("Doc_No");

                // Create new Datatable
                DataTable exportTable = new DataTable();

                for (int i = 0; i < SummaryList.Columns.Count; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(string));
                }

                // Copy the data value
                for (int i = 0; i < SummaryList.Rows.Count; i++)
                {
                    DataRow excelRow = exportTable.NewRow();
                    for (int ii = 0; ii < SummaryList.Columns.Count; ii++)
                    {
                        excelRow["value" + ii.ToString()] = SummaryList.Rows[i][ii].ToString();
                    }

                    exportTable.Rows.Add(excelRow);
                }

                // Creating new column to value20
                for (int i = exportTable.Columns.Count; i < 25; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(String));
                }

                // Formatting the numbers.
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    exportTable.Rows[i]["value2"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value2"]));
                    exportTable.Rows[i]["value3"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value3"]));
                    exportTable.Rows[i]["value4"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value4"]));
                    exportTable.Rows[i]["value5"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value5"]));
                    exportTable.Rows[i]["value6"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value6"]));

                    // Make it empty if value is zero
                    if (exportTable.Rows[i]["value2"].ToString() == "$0.00")
                        exportTable.Rows[i]["value2"] = "";
                    if (exportTable.Rows[i]["value3"].ToString() == "$0.00")
                        exportTable.Rows[i]["value3"] = "";
                    if (exportTable.Rows[i]["value4"].ToString() == "$0.00")
                        exportTable.Rows[i]["value4"] = "";
                    if (exportTable.Rows[i]["value5"].ToString() == "$0.00")
                        exportTable.Rows[i]["value5"] = "";
                    if (exportTable.Rows[i]["value6"].ToString() == "$0.00")
                        exportTable.Rows[i]["value6"] = "";
                }
                exportTable.AcceptChanges();

                // Add the header as "value"
                DataRow excelHeader = exportTable.NewRow();
                excelHeader["value0"] = "Cust/Vend ID";
                excelHeader["value1"] = "Cust/Vend Name";
                excelHeader["value2"] = "Total";
                excelHeader["value3"] = "Current";
                excelHeader["value4"] = "31-60";
                excelHeader["value5"] = "61-90";
                excelHeader["value6"] = "90+";
                excelHeader["value7"] = "Age";

                exportTable.Rows.InsertAt(excelHeader, 0);

                // Bold the header.
                for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                {
                    exportTable.Rows[0][i] = "<b>" + exportTable.Rows[0][i].ToString() + "</b>";
                }

                // Remove Account No if requested
                if (Request.Form["Ac"] != "on")
                {
                    exportTable.Columns.Remove("value0");
                    // Rename the column name so cell shifted to value0
                    for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                    {
                        exportTable.Columns[i].ColumnName = "value" + i.ToString();
                    }
                }

                RPT_Excel.DataSource = exportTable;
                RPT_Excel.DataBind();

                PNL_Excel.Visible = true;
            }
        }
    }

    private void ExportAP()
    {
        DataTable List = new DataTable();
        string curr = Request.Form["currency"];
        string date = Request.Form["date"];
        string cust = Request.Form["cust"];
        string ARType = Request.Form["type"];
        string where = "Document_Date <=@date  and Applies_To <>'0' and Currency_ID = @currency and (Select top 1 Balance_At_Date from ACC_AP where Applies_To = AP1.Applies_To and Document_Date <=@date order by Document_Date desc, AP_ID desc)<>0 ";
        if (ARType == "Details") { if (cust != "all") { where = where + " and AP1.Cust_Vend_ID = '" + cust + "' "; } }

        SQLCommand.CommandText = @"Select distinct(Applies_to) as Doc_ID, 
            (Select top 1 Balance_At_Date from ACC_AP where Applies_To = AP1.Applies_To and Document_Date <=@date order by Document_Date desc, AP_ID desc) as Balance_At_Date, 
            (Select top 1 Document_Date from ACC_AP where Applies_To = AP1.Applies_To and Document_Date <=@date order by Document_Date, AP_ID) as Document_Date,
            Doc_No, AP1.Cust_Vend_ID, Name from ACC_AP AP1 left join ACC_PurchInv on Applies_to = Doc_ID left join Customer on AP1.Cust_Vend_ID = Cust_ID where " + where + " order by Name, Document_Date";
        SQLCommand.Parameters.Clear();
        SQLCommand.Parameters.AddWithValue("@currency", curr);
        SQLCommand.Parameters.AddWithValue("@date", date);
        DataAdapter.Fill(List);

        SQLCommand.CommandText = @"Select Document_ID as Doc_ID, Balance_At_Date, Document_Date, Doc_No, ACC_AP.Cust_Vend_ID, Name from ACC_AP 
            left join ACC_PurchInv on Document_ID=Doc_ID 
            left join Customer on ACC_AP.Cust_Vend_ID = Cust_ID 
            where Applies_to = 0 and Balance_At_Date<>0 and Document_Date <=@date and Currency_ID=@currency ";
        DataAdapter.Fill(List);

        List.Columns.Add("Total", typeof(Double));
        List.Columns.Add("Current", typeof(Double));
        List.Columns.Add("AP30", typeof(Double));
        List.Columns.Add("AP60", typeof(Double));
        List.Columns.Add("AP90", typeof(Double));
        List.Columns.Add("Details", typeof(Double));
        List.Columns.Add("Padding", typeof(string));
        List.Columns.Add("Age", typeof(Int16));

        double Total = 0;
        double current = 0;
        double AP30 = 0;
        double AP60 = 0;
        double AP90 = 0;

        DateTime TempDate = new DateTime();
        DateTime BalDate = new DateTime();
        BalDate = Convert.ToDateTime(date);
        for (int i = 0; i < List.Rows.Count; i++)
        {
            TempDate = Convert.ToDateTime(List.Rows[i]["Document_Date"].ToString());
            double balance;
            if ((double.TryParse(List.Rows[i]["Balance_At_Date"].ToString(), out balance)))
            {
                List.Rows[i]["Total"] = List.Rows[i]["Balance_At_Date"];
            }
            else
            {
                List.Rows[i]["Total"] = "0";
            }

            TimeSpan age = BalDate - TempDate;
            List.Rows[i]["Age"] = age.TotalDays;
            List.Rows[i]["Current"] = "0";
            List.Rows[i]["AP30"] = "0";
            List.Rows[i]["AP60"] = "0";
            List.Rows[i]["AP90"] = "0";

            Total += Convert.ToDouble(List.Rows[i]["Total"].ToString());

            if (TempDate > BalDate.AddDays(-31))
            {
                List.Rows[i]["Current"] = Convert.ToDouble(balance);
                current += Convert.ToDouble(balance);
            }
            else if (TempDate > BalDate.AddDays(-61))
            {
                List.Rows[i]["AP30"] = Convert.ToDouble(balance);
                AP30 += Convert.ToDouble(balance);
            }
            else if (TempDate > BalDate.AddDays(-90))
            {
                List.Rows[i]["AP60"] = Convert.ToDouble(balance);
                AP60 += Convert.ToDouble(balance);
            }
            else
            {
                List.Rows[i]["AP90"] = Convert.ToDouble(balance);
                AP90 += Convert.ToDouble(balance);
            }
        }

        DataTable SummaryList = List.Copy(); //need to pass copy of table so as not to affect list table in the scope of popListDetail

        for (int i = 1; i < SummaryList.Rows.Count; i++)
        {
            if (SummaryList.Rows[i]["Cust_Vend_ID"].ToString() == SummaryList.Rows[i - 1]["Cust_Vend_ID"].ToString())
            {
                SummaryList.Rows[i]["Total"] = Convert.ToDouble(SummaryList.Rows[i]["Total"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["Total"].ToString());
                SummaryList.Rows[i]["AP30"] = Convert.ToDouble(SummaryList.Rows[i]["AP30"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP30"].ToString());
                SummaryList.Rows[i]["AP60"] = Convert.ToDouble(SummaryList.Rows[i]["AP60"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP60"].ToString());
                SummaryList.Rows[i]["AP90"] = Convert.ToDouble(SummaryList.Rows[i]["AP90"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["AP90"].ToString());
                SummaryList.Rows[i]["Current"] = Convert.ToDouble(SummaryList.Rows[i]["Current"].ToString()) + Convert.ToDouble(SummaryList.Rows[i - 1]["Current"].ToString());
                SummaryList.Rows[i - 1].Delete();
            }
        }

        SummaryList.AcceptChanges();

        DataTable Details = new DataTable();
        Details = List.Clone();
        for (int i = 0; i < SummaryList.Rows.Count; i++)
        {
            Details.Rows.Add("0", "0", date, "", "", SummaryList.Rows[i]["Name"], SummaryList.Rows[i]["Total"], SummaryList.Rows[i]["Current"], SummaryList.Rows[i]["AP30"], SummaryList.Rows[i]["AP60"], SummaryList.Rows[i]["AP90"], "0", "totalRow");
            for (int ii = 0; ii < List.Rows.Count; ii++)
            {
                if (List.Rows[ii]["Cust_Vend_ID"].ToString() == SummaryList.Rows[i]["Cust_Vend_ID"].ToString())
                {
                    Details.Rows.Add(List.Rows[ii]["Doc_ID"], "0", List.Rows[ii]["Document_Date"], List.Rows[ii]["Doc_No"], List.Rows[ii]["Cust_Vend_ID"], List.Rows[ii]["Name"], List.Rows[ii]["Total"], List.Rows[ii]["Current"], List.Rows[ii]["AP30"], List.Rows[ii]["AP60"], List.Rows[ii]["AP90"], "0", "20", List.Rows[ii]["Age"]);
                }
            }
        }
        if (ARType == "Details") { if (cust == "all") { Details.Rows.Add("0", "0", date, "", "", "Total", Total, current, AP30, AP60, AP90, "0", "totalRow"); } }

        PNL_Ajax.Visible = true;
        PNL_Details.Visible = true;
        if (ARType == "Details")
        {
            RPT_List_Details.DataSource = Details;
        }
        else
        {
            SummaryList.Rows.Add("0", "0", date, "", "", "Total", Total, current, AP30, AP60, AP90, "0", "totalRow");
            RPT_List_Details.DataSource = SummaryList;
        }
        RPT_List_Details.DataBind();

        LBL_Total.Text = Total.ToString("0,000.00");
        HF_TotalDet.Value = Total.ToString("0,000.00");

        // Export Function
        if (ARType == "Details")
        {
            if (Request.Form["expStat"] == "on")
            {
                // Remove columns that do not need to be display in excel
                Details.Columns.Remove("Doc_ID");
                Details.Columns.Remove("Balance_At_Date");
                Details.Columns.Remove("Details");
                Details.Columns.Remove("Padding");

                // Change the order of DataTable Columns
                Details.Columns["Cust_Vend_ID"].SetOrdinal(0);
                Details.Columns["Name"].SetOrdinal(1);
                Details.Columns["Age"].SetOrdinal(2);
                Details.Columns["Document_Date"].SetOrdinal(3);
                Details.Columns["Doc_No"].SetOrdinal(4);
                Details.Columns["Total"].SetOrdinal(5);
                Details.Columns["Current"].SetOrdinal(6);
                Details.Columns["AP30"].SetOrdinal(7);
                Details.Columns["AP60"].SetOrdinal(8);
                Details.Columns["AP90"].SetOrdinal(9);

                // Create new Datatable
                DataTable exportTable = new DataTable();

                for (int i = 0; i < Details.Columns.Count; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(string));
                }

                // Copy the data value
                for (int i = 0; i < Details.Rows.Count; i++)
                {
                    DataRow excelRow = exportTable.NewRow();
                    for (int ii = 0; ii < Details.Columns.Count; ii++)
                    {
                        excelRow["value" + ii.ToString()] = Details.Rows[i][ii].ToString();
                    }

                    exportTable.Rows.Add(excelRow);
                }

                // Creating new column to value20
                for (int i = exportTable.Columns.Count; i < 25; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(String));
                }

                // Formatting the numbers.
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    exportTable.Rows[i]["value3"] = exportTable.Rows[i]["value3"].ToString().Substring(0, 10);
                    exportTable.Rows[i]["value5"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value5"]));
                    exportTable.Rows[i]["value6"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value6"]));
                    exportTable.Rows[i]["value7"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value7"]));
                    exportTable.Rows[i]["value8"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value8"]));
                    exportTable.Rows[i]["value9"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value9"]));

                    // Make it empty if value is zero
                    if (exportTable.Rows[i]["value5"].ToString() == "$0.00")
                        exportTable.Rows[i]["value5"] = "";
                    if (exportTable.Rows[i]["value6"].ToString() == "$0.00")
                        exportTable.Rows[i]["value6"] = "";
                    if (exportTable.Rows[i]["value7"].ToString() == "$0.00")
                        exportTable.Rows[i]["value7"] = "";
                    if (exportTable.Rows[i]["value8"].ToString() == "$0.00")
                        exportTable.Rows[i]["value8"] = "";
                    if (exportTable.Rows[i]["value9"].ToString() == "$0.00")
                        exportTable.Rows[i]["value9"] = "";
                }
                exportTable.AcceptChanges();

                // Add the header as "value"
                DataRow excelHeader = exportTable.NewRow();
                excelHeader["value0"] = "Cust/Vend ID";
                excelHeader["value1"] = "Cust/Vend Name";
                excelHeader["value2"] = "Age";
                excelHeader["value3"] = "Date";
                excelHeader["value4"] = "Invoice No.";
                excelHeader["value5"] = "Total";
                excelHeader["value6"] = "Current";
                excelHeader["value7"] = "31-60";
                excelHeader["value8"] = "61-90";
                excelHeader["value9"] = "90+";

                exportTable.Rows.InsertAt(excelHeader, 0);

                // Bold the header
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    if (exportTable.Rows[i]["value0"].ToString() == "")
                    {
                        for (int ii = 0; ii < exportTable.Columns.Count - 1; ii++)
                        {
                            exportTable.Rows[i][ii] = "<b>" + exportTable.Rows[i][ii].ToString() + "</b>";
                        }                        
                    }
                }

                // Remove Account No if requested
                if (Request.Form["Ac"] != "on")
                {
                    exportTable.Columns.Remove("value0");
                    // Rename the column name so cell shifted to value0
                    for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                    {
                        exportTable.Columns[i].ColumnName = "value" + i.ToString();
                    }
                }

                RPT_Excel.DataSource = exportTable;
                RPT_Excel.DataBind();

                PNL_Excel.Visible = true;
            }
        }
        else
        {
            if (Request.Form["expStat"] == "on")
            {
                // Remove columns that do not need to be display in excel
                SummaryList.Columns.Remove("Doc_ID");
                SummaryList.Columns.Remove("Balance_At_Date");
                SummaryList.Columns.Remove("Details");
                SummaryList.Columns.Remove("Padding");
                SummaryList.Columns.Remove("Document_Date");
                SummaryList.Columns.Remove("Doc_No");

                // Create new Datatable
                DataTable exportTable = new DataTable();

                for (int i = 0; i < SummaryList.Columns.Count; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(string));
                }

                // Copy the data value
                for (int i = 0; i < SummaryList.Rows.Count; i++)
                {
                    DataRow excelRow = exportTable.NewRow();
                    for (int ii = 0; ii < SummaryList.Columns.Count; ii++)
                    {
                        excelRow["value" + ii.ToString()] = SummaryList.Rows[i][ii].ToString();
                    }

                    exportTable.Rows.Add(excelRow);
                }

                // Creating new column to value20
                for (int i = exportTable.Columns.Count; i < 25; i++)
                {
                    exportTable.Columns.Add("value" + i.ToString(), typeof(String));
                }

                // Formatting the numbers.
                for (int i = 0; i < exportTable.Rows.Count; i++)
                {
                    exportTable.Rows[i]["value2"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value2"]));
                    exportTable.Rows[i]["value3"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value3"]));
                    exportTable.Rows[i]["value4"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value4"]));
                    exportTable.Rows[i]["value5"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value5"]));
                    exportTable.Rows[i]["value6"] = String.Format("{0:C2}", Convert.ToDouble(exportTable.Rows[i]["value6"]));

                    // Make it empty if value is zero
                    if (exportTable.Rows[i]["value2"].ToString() == "$0.00")
                        exportTable.Rows[i]["value2"] = "";
                    if (exportTable.Rows[i]["value3"].ToString() == "$0.00")
                        exportTable.Rows[i]["value3"] = "";
                    if (exportTable.Rows[i]["value4"].ToString() == "$0.00")
                        exportTable.Rows[i]["value4"] = "";
                    if (exportTable.Rows[i]["value5"].ToString() == "$0.00")
                        exportTable.Rows[i]["value5"] = "";
                    if (exportTable.Rows[i]["value6"].ToString() == "$0.00")
                        exportTable.Rows[i]["value6"] = "";
                }
                exportTable.AcceptChanges();

                // Add the header as "value"
                DataRow excelHeader = exportTable.NewRow();
                excelHeader["value0"] = "Cust/Vend ID";
                excelHeader["value1"] = "Cust/Vend Name";
                excelHeader["value2"] = "Total";
                excelHeader["value3"] = "Current";
                excelHeader["value4"] = "31-60";
                excelHeader["value5"] = "61-90";
                excelHeader["value6"] = "90+";
                excelHeader["value7"] = "Age";

                exportTable.Rows.InsertAt(excelHeader, 0);

                // Bold the header.
                for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                {
                    exportTable.Rows[0][i] = "<b>" + exportTable.Rows[0][i].ToString() + "</b>";
                }

                // Remove Account No if requested
                if (Request.Form["Ac"] != "on")
                {
                    exportTable.Columns.Remove("value0");
                    // Rename the column name so cell shifted to value0
                    for (int i = 0; i < exportTable.Columns.Count - 1; i++)
                    {
                        exportTable.Columns[i].ColumnName = "value" + i.ToString();
                    }
                }

                RPT_Excel.DataSource = exportTable;
                RPT_Excel.DataBind();

                PNL_Excel.Visible = true;
            }
        }
    }
}