using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class AjaxPrintingCSharp :  System.Web.UI.Page
{
    SqlConnection Conn = new SqlConnection(System.Configuration.ConfigurationManager.AppSettings["ConnInfo"]);
    SqlCommand SQLCommand = new SqlCommand();
    SqlDataAdapter DataAdapter = new SqlDataAdapter();
    protected void Page_Load(object sender, EventArgs e)
    {
        string Action = Request.Form["action"];

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

    }

    // Sales Multiperiod Quarterly
    private void PrintQuarterlySalesMultiPer() {

    }

    // Sales Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterSalesMultiPer() {

    }

    // Sales Multiperiod Yearly
    private void PrintYearlySalesMultiPer() {

        DataTable COA = new DataTable();
        DataTable Report = new DataTable();

        // Get Fiscal date


        // Denomination

        // Translation

        // Header

        // Get the query

        // Give Padding

        // Denomination Calculation

        // Rounding Calculation

        // Calculating Total and Sub-Total

        // Format all the output for the paper

        // Delete the rows that are not above the detail level

        // Post on Report DataTable

        PNL_PrintReports.Visible = true;
    }

    // Purchases Multiperiod Monthly 
    private void PrintMonthlyPurchasesMultiPer() {

    }

    // Purchases Multiperiod Month to Month
    private void PrintMonthToMonthPurchasesMultiPer() {

    }

    // Purchases Multiperiod QUarterly
    private void PrintQuarterlyPurchasesMultiPer() {

    }

    // Purchases Multiperiod Quarter to Quarter
    private void PrintQuarterToQuarterPurchasesMultiPer() {

    }

    // Purchases Multiperiod Yearly
    private void PrintYearlyPurchasesMultiPer() {

    }

}