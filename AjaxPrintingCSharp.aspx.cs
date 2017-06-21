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
    }

    // Sales Multiperiod Monthly
    private void PrintMonthlySalesMultiPer() {

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
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable sales = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Multiperiod Sales(Month To Month) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Ventas de Varios Períodos (Mes A Mes)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TaxC" + i.ToString();
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND piv.Doc_Date between '" + startDate[0].ToString("yyyy-MM-dd") + "' and '" + endDate[yearCount].ToString("yyyy-MM-dd") + "'";
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
                total[i] += Convert.ToDouble(sales.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(sales.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["TaxC" + i.ToString()]));

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
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], balanceStyle, sales.Rows[i]["Subtotal2"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        percentStyle = percentStyle + " font-weight:bold;";
        if (yearCount == 0)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else if (yearCount == 1)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;
    }

    // Sales Multiperiod Quarterly
    private void PrintQuarterlySalesMultiPer() {

    }

    // Sales Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterSalesMultiPer() {

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
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable sales = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get Fiscal date.
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'";
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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Multiperiod Sales(Yearly) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Ventas de Varios Períodos (Anuales)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_SalesInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='SC') as TaxC" + i.ToString();
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_SalesInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND piv.Doc_Date between '" + startDate[0].ToString("yyyy-MM-dd") + "' and '" + endDate[yearCount].ToString("yyyy-MM-dd") + "'";
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
                total[i] += Convert.ToDouble(sales.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["Tax" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(sales.Rows[ii]["TaxC" + i.ToString()]);      
            }
        }

        sales.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < sales.Rows.Count; i++)
            {
                sales.Rows[i]["Percentage"] = ((double.Parse(sales.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(sales.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(sales.Rows[i]["SubTotal0"].ToString())));
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
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", sales.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", sales.Rows[i]["Name"].ToString(), balanceStyle, sales.Rows[i]["SubTotal0"], balanceStyle, sales.Rows[i]["SubTotal1"], balanceStyle, sales.Rows[i]["Subtotal2"], percentStyle, sales.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        percentStyle = percentStyle + " font-weight:bold;";
        if (yearCount == 0)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else if (yearCount == 1)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");


        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;
    }

    // Purchases Multiperiod Monthly 
    private void PrintMonthlyPurchasesMultiPer() {

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
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable purchases = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Multiperiod Purchases(Month To Month) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Compras Con Varios Períodos (Mes A Mes)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TaxC" + i.ToString();
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND piv.Doc_Date between '" + startDate[0].ToString("yyyy-MM-dd") + "' and '" + endDate[yearCount].ToString("yyyy-MM-dd") + "'";
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
                total[i] += Convert.ToDouble(purchases.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(purchases.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["TaxC" + i.ToString()]));


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
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], balanceStyle, purchases.Rows[i]["Subtotal2"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        percentStyle = percentStyle + " font-weight:bold;";
        if (yearCount == 0)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else if (yearCount == 1)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;
    }

    // Purchases Multiperiod QUarterly
    private void PrintQuarterlyPurchasesMultiPer() {

    }

    // Purchases Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterPurchasesMultiPer() {

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
        string percent = "";

        string balanceStyle = "";
        string percentStyle = "padding: 0px 0px 0px 0px; text-align:right; font-size:0pt; min-width: 0px; max-width: 0px;";

        double[] total = new double[yearCount + 1];
        double percentage = 0;

        DataTable purchases = new DataTable();
        DataTable fiscal = new DataTable();
        DataTable report = new DataTable();

        Conn.Open();

        // Get Fiscal date.
        SQLCommand.CommandText = "SELECT * FROM ACC_Comp_Info WHERE Company_ID = 'Plastics'";
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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Multiperiod Sales(Yearly) " + denomString + "<br/>For " + dateRange + "<br/></span><span style=\"font-size:7pt\">Printed on " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

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
            HF_PrintTitle.Value = "<span style=\"font-size:11pt\">Axiom Plastics Inc<br/>Ventas de Varios Períodos (Anuales)" + denomString + "<br/>Para " + dateRange + "<br/></span><span style=\"font-size:7pt\">Impreso el " + DateTime.Now.ToString("yyyy-MM-dd hh:mm tt") + "</span><div style='Width: 8.5in; position: absolute;'><span style='position: absolute; margin-left: 6in;'></span><span style='position: absolute; margin-left: 4.3in;'></span><span style='position: absolute; margin-left: 6in'></span><span style='position: absolute; margin-left: 4.3in'></span><span style='position: absolute; margin-left: 7.3in'></span></div>";

        }

        // Getting the Query.
        for (int i = 0; i <= yearCount; i++)
        {
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Total" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PINV') as Tax" + i.ToString();
            query = query + ", (Select sum(Total) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TotalC" + i.ToString();
            query = query + ", (Select sum(Tax) from ACC_PurchInv where Cust_Vend_ID = piv.Cust_Vend_ID and Doc_Date between '" + startDate[i].ToString("yyyy-MM-dd") + "' and '" + endDate[i].ToString("yyyy-MM-dd") + "' and Doc_Type='PC') as TaxC" + i.ToString();
        }

        SQLCommand.CommandText = "Select Distinct(Cust_Vend_ID), Name, Currency" + query + " from ACC_PurchInv piv left join Customer on Cust_Vend_ID=Cust_ID where Currency = @cur AND piv.Doc_Date between '" + startDate[0].ToString("yyyy-MM-dd") + "' and '" + endDate[yearCount].ToString("yyyy-MM-dd") + "'";
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
                total[i] += Convert.ToDouble(purchases.Rows[ii]["Total" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["Tax" + i.ToString()]) - (Convert.ToDouble(purchases.Rows[ii]["TotalC" + i.ToString()]) - Convert.ToDouble(purchases.Rows[ii]["TaxC" + i.ToString()]));
            }
        }

        purchases.AcceptChanges();

        // Percentage.
        if (Request.Form["Percentage"] == "on" && yearCount > 0)
        {
            // Calculate Percentage.
            for (int i = 0; i < purchases.Rows.Count; i++)
            {
                purchases.Rows[i]["Percentage"] = ((double.Parse(purchases.Rows[i]["SubTotal" + yearCount.ToString()].ToString()) - double.Parse(purchases.Rows[i]["SubTotal0"].ToString())) / Math.Abs(double.Parse(purchases.Rows[i]["SubTotal0"].ToString())));
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
            if (yearCount == 0)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else if (yearCount == 1)
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
            else
                report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", purchases.Rows[i]["Cust_Vend_ID"].ToString(), "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in;", purchases.Rows[i]["Name"].ToString(), balanceStyle, purchases.Rows[i]["SubTotal0"], balanceStyle, purchases.Rows[i]["SubTotal1"], balanceStyle, purchases.Rows[i]["Subtotal2"], percentStyle, purchases.Rows[i]["Percentage"].ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        }

        // Displaying Total.
        balanceStyle = balanceStyle + " font-weight:bold;";
        percentStyle = percentStyle + " font-weight:bold;";
        if (yearCount == 0)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else if (yearCount == 1)
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
        else
            report.Rows.Add("", "", "text-align:left;font-size:8pt;width: 60px; padding: 3px 5px 3px 5px;", "", "text-align:left; font-size:8pt; padding: 3px 5px 3px 5px; min-width: 2in; max-width: 2in; font-weight:bold;", "Total", balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[0])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[1])), balanceStyle, String.Format("{0:C2}", Convert.ToDouble(total[2])), percentStyle, String.Format("{0:P}", percentage), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");

        RPT_PrintReports.DataSource = report;
        RPT_PrintReports.DataBind();

        Conn.Close();

        PNL_PrintReports.Visible = true;
    }

}