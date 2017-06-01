/**
 * REQUIREMENTS:
 * BTN_Reports
 * <div id="PageCover" style="display:none;"></div>
 * <div id="popup_reports" style="display:none;"></div>
 * <div class="printinfo"></div>
 * <div id="ajaxData" style="display: none"></div>
 * <div id="tempData" style="display: none"></div>
 */


/**
 * To add spanish translations
 * On print.js, add these 2 lines to get the language:
 *  var lang = 0 //0 is english, 1 is spanish
    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1}
 * and add "language: lang" to the data in the ajax call.
 * On AjaxPrinting.aspx.vb, add this line to get the language:
 *  Dim Language As Integer = Request.Form("language")
 * then surround the lines where translation needs to happen with:
 *  If Language = 0 Then
        //English Version
    ElseIf Language = 1 Then
        //Spanish Version
    End If
 * For the SQL Commands, add "TranslatedName As Name" for 
 * the spanish version 
 */
``
// Report Button
$(document).on('click', '[id$=BTN_Reports]', function (e) {
    e.preventDefault();
    $('#po_param').hide();
    $('#table_print').show();
    //$('[id$=TB_Print_Date1]').val($('[id$=TB_Date]').val())
    //$('[id$=TB_Print_Date2]').val($('[id$=TB_Date]').val())
    $('[id$=LBL_T_Title]').text('Print Reports')

    // Added popup
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
        data: {
            action: "Report"
        },
        success: function (data, status, other) {
            $('#popup_reports').html(data)
            $('#PageCover').show();
            $('#popup_reports').show();

            $(function () { $('[id$=TB_Print_Date1]').datepicker({ dateFormat: "yy-mm-dd", onClose: function (dateText) { pickedDate(dateText); } }); });
            $(function () { $('[id$=TB_Print_Date2]').datepicker({ dateFormat: "yy-mm-dd", onClose: function (dateText) { pickedDate2(dateText); } }); });
            $(function () { $('[id$=Date_Print_From]').datepicker({ dateFormat: "yy-mm-dd", onClose: function (dateText) { choosenDate(dateText); } }); });
            $(function () { $('[id$=Date_Print_To]').datepicker({ dateFormat: "yy-mm-dd", onClose: function (dateText) { choosenDate2(dateText); } }); });
            $(function () {
                $('[id$=TB_Print_Date11]').datepicker({
                    changeMonth: true,
                    changeYear: true,
                    showButtonPanel: true,
                    dateFormat: 'yy-mm',
                    onClose: function (dateText, inst) {
                        $(this).datepicker('setDate', new Date(inst.selectedYear, inst.selectedMonth, 1));
                        //choosenMonth(inst)
                    }
                });
            });
            $(function () {
                $('[id$=TB_Print_Date22]').datepicker({
                    changeMonth: true,
                    changeYear: true,
                    showButtonPanel: true,
                    dateFormat: 'yy-mm',
                    onClose: function (dateText, inst) {
                        $(this).datepicker('setDate', new Date(inst.selectedYear, inst.selectedMonth, 1));
                        //choosenMonth2(inst)
                    }
                });
            });

            // Show only the Month and Year
            $('[id$=TB_Print_Date11]').focusin(function () {
                $('.ui-datepicker-calendar').css("display", "none");
            });
            $('[id$=TB_Print_Date22]').focusin(function () {
                $('.ui-datepicker-calendar').css("display", "none");
            });

            // Set Default date to Today
            $('[id$=TB_Print_Date1]').val($('[id$=HF_Date_Today]').val())
            $('[id$=TB_Print_Date2]').val($('[id$=HF_Date_Today]').val())
            $('[id$=Date_Print_From]').val($('[id$=HF_Date_Today]').val())
            $('[id$=Date_Print_To]').val($('[id$=HF_Date_Today]').val())
            // Required for DDL_Print_Level
            populate_DDL_Print_Level($('[id$=HF_Date_Today]').val())

            // Set Restriction on Date
            function pickedDate(date) {
                if (date) {
                    if ($('[id$=TB_Print_Date2]').val() == '' || $('[id$=TB_Print_Date2]').val() == 'To Date') { $('[id$=TB_Print_Date2]').val(date) }
                    if ($('[id$=TB_Print_Date2]').val() < date) { $('[id$=TB_Print_Date2]').val(date) }
                    if ($('[id$=TB_Print_Date1]').val() != 'From Date') { $('[id$=TB_Print_Date1]').css('color', 'black'); }
                    if ($('[id$=TB_Print_Date2]').val() != 'To Date') { $('[id$=TB_Print_Date2]').css('color', 'black'); }
                }
            }

            function pickedDate2(date) {
                if (date) {
                    if ($('[id$=TB_Print_Date1]').val() == '' || $('[id$=TB_Print_Date1]').val() == 'From Date') { $('[id$=TB_Print_Date1]').val($('[id$=TB_Print_Date2]').val()) }
                    if (date < $('[id$=TB_Print_Date1]').val()) { $('[id$=TB_Print_Date2]').val($('[id$=TB_Print_Date1]').val()) }
                    if ($('[id$=TB_Print_Date1]').val() != 'From Date') { $('[id$=TB_Print_Date1]').css('color', 'black'); }
                    if ($('[id$=TB_Print_Date2]').val() != 'To Date') { $('[id$=TB_Print_Date2]').css('color', 'black'); }
                }
            }

            function choosenDate(date) {
                if (date) {
                    if ($('[id$=Date_Print_To]').val() == '' || $('[id$=Date_Print_To]').val() == 'To Date') { $('[id$=Date_Print_To]').val(date) }
                    if ($('[id$=Date_Print_To]').val() < date) { $('[id$=Date_Print_To]').val(date) }
                    if ($('[id$=Date_Print_From]').val() != 'From Date') { $('[id$=Date_Print_From]').css('color', 'black'); }
                    if ($('[id$=Date_Print_To]').val() != 'To Date') { $('[id$=Date_Print_To]').css('color', 'black'); }
                }
            }

            function choosenDate2(date) {
                if (date) {
                    if ($('[id$=Date_Print_From]').val() == '' || $('[id$=Date_Print_From]').val() == 'From Date') { $('[id$=Date_Print_From]').val($('[id$=Date_Print_To]').val()) }
                    if (date < $('[id$=Date_Print_From]').val()) { $('[id$=Date_Print_To]').val($('[id$=Date_Print_From]').val()) }
                    if ($('[id$=Date_Print_From]').val() != 'From Date') { $('[id$=Date_Print_From]').css('color', 'black'); }
                    if ($('[id$=Date_Print_To]').val() != 'To Date') { $('[id$=Date_Print_To]').css('color', 'black'); }
                }
            }

            function choosenMonth(date) {
                if (date) {
                    if ($('[id$=TB_Print_Date22]').val() == '' || $('[id$=TB_Print_Date22]').val() == 'To Date') { $('[id$=TB_Print_Date22]').val(date) }
                    if ($('[id$=TB_Print_Date22]').val() < date) { $('[id$=TB_Print_Date22]').val(date) }
                    if ($('[id$=TB_Print_Date11]').val() != 'From Date') { $('[id$=TB_Print_Date11]').css('color', 'black'); }
                    if ($('[id$=TB_Print_Date22]').val() != 'To Date') { $('[id$=TB_Print_Date22]').css('color', 'black'); }
                }
            }

            function choosenMonth2(date) {
                if (date) {
                    if ($('[id$=TB_Print_Date11]').val() == '' || $('[id$=TB_Print_Date11]').val() == 'From Date') { $('[id$=TB_Print_Date11]').val($('[id$=TB_Print_Date22]').val()) }
                    if (date < $('[id$=TB_Print_Date11]').val()) { $('[id$=TB_Print_Date22]').val($('[id$=TB_Print_Date11]').val()) }
                    if ($('[id$=TB_Print_Date11]').val() != 'From Date') { $('[id$=TB_Print_Date11]').css('color', 'black'); }
                    if ($('[id$=TB_Print_Date22]').val() != 'To Date') { $('[id$=TB_Print_Date22]').css('color', 'black'); }
                }
            }
        },
        error: function (data, status, other) { alert(other); }
    });
});

// Has DDL_Print_Level on it
function populate_DDL_Print_Level(Date) {
    $('#spinner').show()
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxAccounting.aspx',
        data: { action: 'popchart', date: Date },
        success: function (data, status, other) {
            $('#ajaxData').html(data)
            $('[id$=DDL_Print_Level]').empty()
            for (i = 1; i <= parseInt($('[id$=HF_TopLevel]').val(), 10) ; i++) {
                $('[id$=DDL_Print_Level]').append('<option value="' + i + '">' + i + '</option>')
            }
            $('[id$=DDL_Print_Level]').val(parseInt($('[id$=HF_TopLevel]').val(), 10)).attr("selected",true);
        },
        error: function (data, status, other) { alert(other) }
    });
}

// Change on DDL_Print_Category
$(document).on('change', '[id$=DDL_Print_Category]', function () {
    // Hide things the user doesnt have to see when the change the category type
    // General Category is picked
    if ($(this).val() == "1") {
        $('#table_general1').show();
        $('#table_general2').show();
        $('#table_sales').hide();
        if ($('[id$=DDL_Print_Report]').val() == '2') {
            $('#Show_per').show();
        }
        else {
            $('#Show_per').hide();
        }        
        $('#table_MultiPeriod').hide();
        if ($('[id$=DDL_Print_Report]').val() == '4') {
            $('#DetailReport').show();
            $('#td_detail').hide();
        }
        $('[id$=BTN_Print_Export]').hide();
    }

    // General-Multiperiod is picked
    if ($(this).val() == "10") {
        $('#table_general1').hide();
        $('#table_general2').show();
        $('#Date_DTSpan').hide();
        $('#table_sales').hide();
        $('#DetailReport').hide();
        $('#Show_per').hide();
        $('#td_detail').show();
        $('#table_MultiPeriod').show();
        $('[id$=BTN_Print_Export]').hide();
        if ($('[id$=DDL_Print_Period]').val() == 'Monthly') {
            $('#MonthlySelector').show();
            $('#QuarterlySelector1').hide();
            $('#QuarterlySelector2').hide();
            $('#YearlySelector').hide();
        }
        //else if ($('[id$=DDL_Print_Period]').val() == 'Quarterly') {
        //    $('#MonthlySelector').hide();
        //    $('#QuarterlySelector1').show();
        //    $('#QuarterlySelector2').show();
        //    $('#YearlySelector').hide();
        //}
        //else if ($('[id$=RB_Yearly]').is(':checked')) {
        //    $('#MonthlySelector').hide();
        //    $('#QuarterlySelector1').hide();
        //    $('#QuarterlySelector2').hide();
        //    $('#YearlySelector').show();
        //}
    }

    // Sales is picked
    if ($(this).val() == "2") {
        $('#table_general1').hide();
        $('#table_MultiPeriod').hide();
        $('#table_general2').hide();
        $('#Date_DTSpan').hide();
        $('#Show_per').hide();
        $('#table_sales').show();
        $('[id$=BTN_Print_Export]').hide();
        if ($('[id$=DDL_Print_Details]').val() == 'Details') { printpopCustDD(); }
    }
    
    // Purchases Category is picked
    if ($(this).val() == "3") {
        $('#table_general1').hide();
        $('#table_general2').hide();
        $('#table_MultiPeriod').hide();
        $('#Date_DTSpan').hide();
        $('#Show_per').hide();
        $('#table_sales').show();
        $('[id$=BTN_Print_Export]').hide();
        if ($('[id$=DDL_Print_Details]').val() == 'Details') { printpopVendDD(); }
    }
});

// Change on DDL_Print_Report
$(document).on('change', '[id$=DDL_Print_Report]', function () {
    //Hide things the user doesnt have to see when the change the report type for example summary and detail don't need the detail
    $('#PrintDate2Span').show();
    $("#td_detail").show();
    $("#DetailReport").hide();
    $("#showZeros").show();
    $("#MonthToMonth").hide();
    if ($(this).val() == "1") { $('#PrintDate2Span').hide(); $('#Show_per').hide(); }//Balance Sheet Trial
    if ($(this).val() == "2") { $("#MonthToMonth").show(); $('#Show_per').show(); }//Profit and Loss sheet
    if ($(this).val() == "3") { $("#td_detail").hide(); $('#Show_per').hide();}//Detail Trial
    if ($(this).val() == "4") { $("#td_detail").hide(); $("#DetailReport").show(); $("#showZeros").hide(); $('#Show_per').hide();}//Detail Trial
    if ($(this).val() == "5") { ('#PrintDate2Span').hide(); $('#Show_per').hide(); }//Detail Trial
});

// Change on DDL_Print_Details
$(document).on('change', '[id$=DDL_Print_Details]', function () {
    
    if ($(this).val() == "Summary") {
        $('#td_customer').hide();
        $('#Date_DTSpan').hide();
    }
    else if ($(this).val() == 'Details') {
        $('#td_customer').show();
        $('#Date_DTSpan').hide();
        $('#spinner').show();
        if ($('[id$=DDL_Print_Category]').val() == "2") { printpopCustDD(); }//Sales Category is picked
        if ($('[id$=DDL_Print_Category]').val() == "3") { printpopVendDD(); }//Purchases Category is picked
    }
    else {
        $('#td_customer').hide();
        $('#Date_DTSpan').show();
    }
});

$(document).on('change', '[id$=DDL_Print_Period]', function () {
    if ($('[id$=DDL_Print_Period]').val() == 'Monthly') {
        $('#MonthlySelector').show();
        $('#QuarterlySelector1').hide();
        $('#QuarterlySelector2').hide();
        $('#YearlySelector').hide();
    }
    else if ($('[id$=DDL_Print_Period]').val() == 'Month-to-Month') {
        $('#MonthlySelector').hide();
        $('#QuarterlySelector1').hide();
        $('#QuarterlySelector2').hide();
        $('#YearlySelector').hide();
    }
    else if ($('[id$=DDL_Print_Period]').val() == 'Quarterly') { 
        $('#MonthlySelector').hide();
        $('#QuarterlySelector1').show();
        $('#QuarterlySelector2').show();
        $('#YearlySelector').hide();
    }
    else if ($('[id$=DDL_Print_Period]').val() == 'Quarter-to-Quarter') {
        $('#MonthlySelector').hide();
        $('#QuarterlySelector1').hide();
        $('#QuarterlySelector2').hide();
        $('#YearlySelector').hide();
    }
    else if ($('[id$=DDL_Print_Period]').val() == 'Yearly') {
        $('#MonthlySelector').hide();
        $('#QuarterlySelector1').hide();
        $('#QuarterlySelector2').hide();
        $('#YearlySelector').show();
    }    
});

// Cancel Button
$(document).on('click', '[id$=BTN_Print_Cancel]', function (e) {
    e.preventDefault()
    $('#po_param').show()
    $('#table_print').hide()
    $('[id$=LBL_T_Title]').text('Chart of Accounts')
    $('#popup_reports').hide();
    $('#PageCover').hide();
});

// Print Button
$(document).on('click', '[id$=BTN_Print_Print]', function (e) {
    e.preventDefault()
    $('#spinner').show();
    $('#printinfo').empty();
    $('#tempData').empty();

    if ($('[id$=DDL_Print_Category]').val() == "1") { printGeneral() }//General Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "2") {
        if ($('[id$=DDL_Print_Details]').val() == 'Details') { printSales(); }
        else if ($('[id$=DDL_Print_Details]').val() == 'Summary') { printSales(); }
        else { printSalesReport(); }
    }//Sales Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "3") {
        if ($('[id$=DDL_Print_Details]').val() == 'Details') { printPurchases(); }
        else if ($('[id$=DDL_Print_Details]').val() == 'Summary') { printPurchases(); }
        else { printPurchReport(); }
    }//Purchases Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "10") {
        if ($('[id$=DDL_Print_MultiPeriod]').val() == '22') { printIncStateMulti(); }
        else if ($('[id$=DDL_Print_MultiPeriod]').val() == '11') { printBalSheetMulti(); }
        else {  }
    }
});

// Multiperiod Income Statement
function printIncStateMulti() {
    var checked = "off"
    var roundChecked = "off"
    var accno = "off"
    var lang = 0 //0 is english, 1 is spanish

    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1 }
    if ($('[id$=CB_Print_Accno]').is(':checked')) { accno = "on" } else { accno = "off" }
    if ($('[id$=CB_Print_ShowZeros]').is(':checked')) { checked = "on" } else { checked = "off" }
    if ($('[id$=CB_Print_Round]').is(':checked')) { roundChecked = "on" } else { roundChecked = "off" }
    if ($('[id$=DDL_Print_Period]').val() == 'Monthly') {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "IncStateMultiMonthly", language: lang, FirstDate: $('[id$=TB_Print_Date11]').val(), SecondDate: $('[id$=TB_Print_Date22]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });       
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Month-to-Month') {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "IncStateMultiMonth-to-Month", language: lang, FirstDate: $('[id$=TB_Print_Date11]').val(), SecondDate: $('[id$=TB_Print_Date22]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Quarterly') {
        var Q1ch = "off"
        var Q2ch = "off"
        var Q3ch = "off"
        var Q4ch = "off"
        if ($('[id$=CB_Q1]').is(':checked')) { Q1ch = "on" } else { Q1ch = "off" }
        if ($('[id$=CB_Q2]').is(':checked')) { Q2ch = "on" } else { Q2ch = "off" }
        if ($('[id$=CB_Q3]').is(':checked')) { Q3ch = "on" } else { Q3ch = "off" }
        if ($('[id$=CB_Q4]').is(':checked')) { Q4ch = "on" } else { Q4ch = "off" }

        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "IncStateMultiQuarterly", language: lang, YearForQuater: $('[id$=DDL_Print_Quarter]').val(), Q1: Q1ch, Q2: Q2ch, Q3: Q3ch, Q4: Q4ch, detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Quarter-to-Quarter') {
        var Q1ch = "off"
        var Q2ch = "off"
        var Q3ch = "off"
        var Q4ch = "off"
        if ($('[id$=CB_Q1]').is(':checked')) { Q1ch = "on" } else { Q1ch = "off" }
        if ($('[id$=CB_Q2]').is(':checked')) { Q2ch = "on" } else { Q2ch = "off" }
        if ($('[id$=CB_Q3]').is(':checked')) { Q3ch = "on" } else { Q3ch = "off" }
        if ($('[id$=CB_Q4]').is(':checked')) { Q4ch = "on" } else { Q4ch = "off" }

        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "IncStateMultiQuarter-to-Quarter", language: lang, YearForQuater: $('[id$=DDL_Print_Quarter]').val(), Q1: Q1ch, Q2: Q2ch, Q3: Q3ch, Q4: Q4ch, detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Yearly') {
        // Year restriction
        if ($('[id=DDL_Print_YearTo]').val() - $('[id=DDL_Print_YearFrom]').val() > 2) {
            alert("Please select no more than 2 (Two) years in difference")
        }
        else {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "IncStateMultiYearly", language: lang, FirstDate: $('[id$=DDL_Print_YearFrom]').val(), SecondDate: $('[id$=DDL_Print_YearTo]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').removeClass('HideOnPage');
                    $('#printinfo').empty()
                    $('#printinfo').html(data)
                    printReport()
                },
                error: function (data, status, other) { alert(other); }
            });
        }        
    }
}

// Multiperiod Balance Sheet
function printBalSheetMulti() {
    var checked = "off"
    var roundChecked = "off"
    var accno = "off"
    var lang = 0 //0 is english, 1 is spanish

    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1 }
    if ($('[id$=CB_Print_Accno]').is(':checked')) { accno = "on" } else { accno = "off" }
    if ($('[id$=CB_Print_ShowZeros]').is(':checked')) { checked = "on" } else { checked = "off" }
    if ($('[id$=CB_Print_Round]').is(':checked')) { roundChecked = "on" } else { roundChecked = "off" }
    if ($('[id$=DDL_Print_Period]').val() == 'Monthly') {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalSheetMultiMonthly", language: lang, FirstDate: $('[id$=TB_Print_Date11]').val(), SecondDate: $('[id$=TB_Print_Date22]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Month-to-Month') {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalSheetMultiMonth-to-Month", language: lang, FirstDate: $('[id$=TB_Print_Date11]').val(), SecondDate: $('[id$=TB_Print_Date22]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Quarterly') {
        var Q1ch = "off"
        var Q2ch = "off"
        var Q3ch = "off"
        var Q4ch = "off"
        if ($('[id$=CB_Q1]').is(':checked')) { Q1ch = "on" } else { Q1ch = "off" }
        if ($('[id$=CB_Q2]').is(':checked')) { Q2ch = "on" } else { Q2ch = "off" }
        if ($('[id$=CB_Q3]').is(':checked')) { Q3ch = "on" } else { Q3ch = "off" }
        if ($('[id$=CB_Q4]').is(':checked')) { Q4ch = "on" } else { Q4ch = "off" }

        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalSheetMultiQuarterly", language: lang, YearForQuater: $('[id$=DDL_Print_Quarter]').val(), Q1: Q1ch, Q2: Q2ch, Q3: Q3ch, Q4: Q4ch, detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno,Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Quarter-to-Quarter') {
        var Q1ch = "off"
        var Q2ch = "off"
        var Q3ch = "off"
        var Q4ch = "off"
        if ($('[id$=CB_Q1]').is(':checked')) { Q1ch = "on" } else { Q1ch = "off" }
        if ($('[id$=CB_Q2]').is(':checked')) { Q2ch = "on" } else { Q2ch = "off" }
        if ($('[id$=CB_Q3]').is(':checked')) { Q3ch = "on" } else { Q3ch = "off" }
        if ($('[id$=CB_Q4]').is(':checked')) { Q4ch = "on" } else { Q4ch = "off" }

        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalSheetMultiQuarter-to-Quarter", language: lang, YearForQuater: $('[id$=DDL_Print_Quarter]').val(), Q1: Q1ch, Q2: Q2ch, Q3: Q3ch, Q4: Q4ch, detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    if ($('[id$=DDL_Print_Period]').val() == 'Yearly') {
        // Year restriction
        if ($('[id=DDL_Print_YearTo]').val() - $('[id=DDL_Print_YearFrom]').val() > 2) {
            alert("Please select no more than 2 (Two) years in difference")
        }
        else {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "BalSheetMultiYearly", language: lang, FirstDate: $('[id$=DDL_Print_YearFrom]').val(), SecondDate: $('[id$=DDL_Print_YearTo]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').removeClass('HideOnPage');
                    $('#printinfo').empty()
                    $('#printinfo').html(data)
                    printReport()
                },
                error: function (data, status, other) { alert(other); }
            });
        }        
    }
}

// Print for General
function printGeneral() {
    var checked = "off"
    var roundChecked = "off"
    var roundChecked = "off"
    var per = "off"
    var accno = "off"
    var lang = 0 //0 is english, 1 is spanish

    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1 }
    if ($('[id$=CB_Print_ShowPer]').is(':checked')) { per = "on" } else { per = "off" }
    if ($('[id$=CB_Print_Accno]').is(':checked')) { accno = "on" } else { accno = "off" }
    if ($('[id$=CB_Print_ShowZeros]').is(':checked')) { checked = "on" } else { checked = "off" }
    if ($('[id$=CB_Print_Round]').is(':checked')) { roundChecked = "on" } else { roundChecked = "off" }
    if ($('[id$=DDL_Print_Report]').val() == "1") {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalanceSheet", language: lang, date1: $('[id$=TB_Print_Date1]').val(), date2: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    else if ($('[id$=DDL_Print_Report]').val() == "2") {
        if ($('[id$=CB_Print_MonthToMonth]').is(':checked')) {
            alert("Can not Print Month To Month, uncheck to print or select Excel to download.")
            $('#spinner').hide();
        }
        else {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "ProfitLoss", language: lang, FirstDate: $('[id$=TB_Print_Date1]').val(), SecondDate: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Ac: accno, Perce: per, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').removeClass('HideOnPage');
                    $('#printinfo').empty()
                    $('#printinfo').html(data)
                    printReport()
                },
                error: function (data, status, other) { alert(other); }
            });
        }
    }
    else if ($('[id$=DDL_Print_Report]').val() == "4") {
        var StartDate = $('[id$=TB_Print_Date1]').val();
        var EndDate = $('[id$=TB_Print_Date2]').val();
        var accNo = $('[id$=TB_Print_AccNo]').val()

        if (accNo != "") {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "DetailTrialChart", language: lang, StartDate: StartDate, EndDate: EndDate, accNo: accNo, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').removeClass('HideOnPage');
                    $('#printinfo').html(data)
                    printReport()
                },
                error: function (data, status, other) { alert(other); }
            });
        }
        else {
            alert("No ID Inputed");
        }
    }
    else if ($('[id$=DDL_Print_Report]').val() == "3") {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "SummaryTrail", language: lang, FirstDate: $('[id$=TB_Print_Date1]').val(), SecondDate: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').removeClass('HideOnPage');
                $('#printinfo').empty()
                $('#printinfo').html(data)
                printReport()
            },
            error: function (data, status, other) { alert(other); }
        });
    }
}

// Print for Sales
function printSales() {
    printpopAR();
}

// Print for Purchases
function printPurchases() {
    printpopAP();
}


// Export Button
$(document).on('click', '[id$=BTN_Print_Export]', function (e) {
    e.preventDefault()
    $('#spinner').show();

    if ($('[id$=DDL_Print_Category]').val() == "1") { exportGeneral() }//General Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "2") { exportSales() }//Sales Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "3") { exportPurchases() }//Purchases Category is picked    
});

// Export for General
function exportGeneral() {
    var checked = "off"
    var roundChecked = "off"
    if ($('[id$=CB_Print_ShowZeros]').is(':checked')) { checked = "on" } else { checked = "off" }
    if ($('[id$=CB_Print_Round]').is(':checked')) { roundChecked = "on" } else { roundChecked = "off" }
    if ($('[id$=DDL_Print_Report]').val() == "1") {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "BalanceSheetXML", date1: $('[id$=TB_Print_Date1]').val(), date2: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').html(data)
                download("BalanceSheet-" + $('[id$=TB_Print_Date1]').val() + ".xml", $('[id$=HF_XML]').val());
            },
            error: function (data, status, other) { alert(other); }
        });
    }
    else if ($('[id$=DDL_Print_Report]').val() == "2") {
        //Get the start of every month for the set range
        if ($('[id$=CB_Print_MonthToMonth]').is(':checked')) {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "ProfitLossXMLM2M", FirstDate: $('[id$=TB_Print_Date1]').val(), SecondDate: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').html(data)
                    download("IncomeStatementM2M-" + $('[id$=TB_Print_Date1]').val() + "-" + $('[id$=TB_Print_Date2]').val() + ".xml", $('[id$=HF_XML]').val());
                },
                error: function (data, status, other) { alert(other); }
            });
        }
            //Get the default
        else {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "ProfitLossXML", FirstDate: $('[id$=TB_Print_Date1]').val(), SecondDate: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').html(data)
                    download("IncomeStatement-" + $('[id$=TB_Print_Date1]').val() + "-" + $('[id$=TB_Print_Date2]').val() + ".xml", $('[id$=HF_XML]').val());
                },
                error: function (data, status, other) { alert(other); }
            });
        }
    }
    else if ($('[id$=DDL_Print_Report]').val() == "4") {
        var StartDate = $('[id$=TB_Print_Date1]').val();
        var EndDate = $('[id$=TB_Print_Date2]').val();
        var accNo = $('[id$=TB_Print_AccNo]').val()

        if (accNo != "") {
            $.ajax({
                async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
                data: { action: "DetailTrialXML", StartDate: StartDate, EndDate: EndDate, accNo: accNo, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
                success: function (data, status, other) {
                    $('#printinfo').html(data)
                    download("DetailTrial-" + $('[id$=TB_Print_Date1]').val() + "-" + $('[id$=TB_Print_Date2]').val() + ".xml", $('[id$=HF_XML]').val());
                },
                error: function (data, status, other) { alert(other); }
            });
        }
        else {
            alert("No ID Inputed");
        }
    }
    else if ($('[id$=DDL_Print_Report]').val() == "3") {
        $.ajax({
            async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
            data: { action: "SummaryTrailXML", FirstDate: $('[id$=TB_Print_Date1]').val(), SecondDate: $('[id$=TB_Print_Date2]').val(), detailLevel: $('[id$=DDL_Print_Level]').val(), showZeros: checked, Denom: $('[id$=DDL_Print_Denomination]').val(), Round: roundChecked },
            success: function (data, status, other) {
                $('#printinfo').html(data)
                download("SummaryTrial-" + $('[id$=TB_Print_Date1]').val() + "-" + $('[id$=TB_Print_Date2]').val() + ".xml", $('[id$=HF_XML]').val());
            },
            error: function (data, status, other) { alert(other); }
        });
    }
}

// Export for Sales
function exportSales() {

}

// Export for Purchases
function exportPurchases() {

}

// Export/Download function
function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent('<?xml version="1.0" encoding="UTF-8"?>' + text));
    element.setAttribute('download', filename);

    element.style.display = 'none';
    document.body.appendChild(element);

    element.click();

    document.body.removeChild(element);
}


function resize() {
    if ($('[id*=PNL_Scroll_Det]').length > 0) {//checking if element exists
        var pos = $('[id*=PNL_Scroll_Det]').position()
        var pos2 = $('#content_table').position()
        var total = $(window).height()
        var hgt = total - pos.top - pos2.top - 110
        $('[id$=PNL_Scroll_Det]').css('height', hgt)
        $('#spinner').hide();
    }
}

// Remove Zeros before print
function dumpZeros() {
    $('[id*=LBL_AP]').each(function () { if ($(this).text() == '.00') { $(this).text('') } });
}

function addPercentPaid() {
    $("[id*=LBL_AP90], [id*=LBL_AP60], [id*=LBL_AP30], [id*=LBL_APCurrent]").each(function () {
        totalOwed = $(this).first().parent().siblings()[4].textContent
        //percentOwed = parseFloat($(this)[0].textContent) / parseFloat(totalOwed) * 100;
        percentOwed = parseFloat($(this)[0].textContent.replace(/,/g, '')) / parseFloat(totalOwed.replace(/,/g, '')) * 100;
        if (!isNaN(percentOwed)) {
            $(this).prepend("<span class='tooltiptext'>" + percentOwed.toFixed(2) + "%</span>");
        }
    });
}

// Show DropDown List for Customer
function printpopCustDD() {
    var cur = $('[id$=DDL_Print_Currency]').val()
    $('[id$=DDL_Print_Customer]').empty()
    $('[id$=DDL_Print_Customer]').append('<option value="all">All Customers</option>')
    $('[id$=CustCur]').each(function () {
        if ($(this).val() == cur) {
            header = $(this).attr('id').replace('HF_Print_CustCur', '')
            $('[id$=DDL_Print_Customer]').append('<option value="' + $('#' + header + 'HF_Print_CustID').val() + '">' + $('#' + header + 'HF_Print_CustName').val() + '</option>')
        }
    });
    $('[id$=DDL_Print_Customer]').val('all')
    $('#spinner').hide();
}

// Show DropDown List for Vendor
function printpopVendDD() {
    var cur = $('[id$=DDL_Print_Currency]').val()
    $('[id$=DDL_Print_Customer]').empty()
    $('[id$=DDL_Print_Customer]').append('<option value="all">All Vendors</option>')
    $('[id$=VendCur]').each(function () {
        if ($(this).val() == cur) {
            header = $(this).attr('id').replace('HF_Print_VendCur', '')
            $('[id$=DDL_Print_Customer]').append('<option value="' + $('#' + header + 'HF_Print_VendID').val() + '">' + $('#' + header + 'HF_Print_VendName').val() + '</option>')
        }
    });
    $('[id$=DDL_Print_Customer]').val('all')
    $('#spinner').hide();
}

// Populate the data before print/export
function printpopAR() {
    $('#spinner').show();
    // Temporary data to Print Sales/Purchase Summary/Detail
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
        data: {
            action: "ShowPanelReport"
        },
        success: function (data, status, other) {
            $('#tempData').html(data)
            $('#tempData').show();
            // Load the result to printinfo on masterpage
            $('#printinfo').load('ACC_Accounts_Receivables.aspx [id$=PNL_Details]',
                {
                    action: 'popList',
                    currency: $('[id$=DDL_Print_Currency]').val(),
                    type: $('[id$=DDL_Print_Details]').val(),
                    date: $('[id$=Date_Print_From]').val(),
                    cust: $('[id$=DDL_Print_Customer]').val()
                },
                function () {
                    $('[id$=LBL_Total]').text('$' + $("[id$=HF_TotalDet]").val());
                    dumpZeros()
                    if ($('[id$=DDL_Print_Details]').val() == 'Details') {
                        $('[id^=td_date]').show();
                        $('[id^=td_invoice]').show();
                        $('[id^=td_age]').show();
                    }
                    else if ($('[id$=DDL_Print_Details]').val() == 'Summary') {
                        $('[id^=td_date]').hide();
                        $('[id^=td_invoice]').hide();
                        $('[id^=td_age]').hide();
                    }
                    addPercentPaid();
                    pre_printAR();
                }
            );//End of Ajax
        },
        error: function (data, status, other) { alert(other); }
    });

}

// Populate the data before print/export
function printpopAP() {
    $('#spinner').show();
    // Temporary data to Print Sales/Purchase Summary/Detail
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
        data: {
            action: "ShowPanelReport"
        },
        success: function (data, status, other) {
            $('#tempData').html(data)
            $('#tempData').show();
            // Load the result to printinfo on masterpage
            $('#printinfo').load('ACC_Accounts_Payable.aspx [id$=PNL_Details]',
                {
                    action: 'popList',
                    currency: $('[id$=DDL_Print_Currency]').val(),
                    type: $('[id$=DDL_Print_Details]').val(),
                    date: $('[id$=Date_Print_From]').val(),
                    cust: $('[id$=DDL_Print_Customer]').val()
                },
                function () {
                    $('[id$=LBL_Total]').text('$' + $("[id$=HF_TotalDet]").val());
                    dumpZeros()
                    if ($('[id$=DDL_Print_Details]').val() == 'Details') {
                        $('[id^=td_date]').show();
                        $('[id^=td_invoice]').show();
                        $('[id^=td_age]').show();
                    }
                    else if ($('[id$=DDL_Print_Details]').val() == 'Summary') {
                        $('[id^=td_date]').hide();
                        $('[id^=td_invoice]').hide();
                        $('[id^=td_age]').hide();
                    }
                    addPercentPaid();
                    pre_printAP();
                }
            );//End of Ajax
        },
        error: function (data, status, other) { alert(other); }
    });// End of Ajax
}

// Printing AR
function pre_printAR() {
    $('.tooltiptext').remove()

    var dt = new Date($.now());
    dt = dt.toString().substr(0, 24)

    $('[id$=HF_PrintHeader]').val('text-align:left; width:80px; font-size:8pt~Customer~text-align:right; font-size:8pt~Total ($)~text-align:right; font-size:8pt~Current ($)~text-align:right; font-size:8pt~30-60 ($)~text-align:right; font-size:8pt~61-90 ($)~text-align:right; font-size:8pt~90+ ($)');
    $('[id$=HF_PrintTitle]').val('<span style="font-size:12pt">Axiom Plastics Inc.<br/>Aged Accounts Receivable ' + $('[id$=DDL_Print_Details]').val() + ' Report (' + $('[id$=DDL_Print_Currency]').val() + ')<br/>As Of ' + $('[id$=Date_Print_From]').val() + '<br/></span><span style="font-size:7pt">printed on: ' + dt + '</span>');

    var app = ''
    var i = 1;
    $('[id$=LBL_CustName]').each(function () {
        var header = $(this).attr('id').replace('LBL_CustName', '')
        if ($(this).text() == 'Total') {
            var total = 0;
            var apcurrent = 0;
            var ap30 = 0;
            var ap60 = 0;
            var ap90 = 0;

            app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px;border-top:solid 0px:; font-weight:bold~' + $(this).text() + '~font-size:8pt; border-top:solid 1px;border-bottom:double 4px;text-align:right; font-weight:bold~'
            if ($('#' + header + 'LBL_APTotal').text() != '') { app = app + '$' + $('#' + header + 'LBL_APTotal').text(); total = parseFloat($('#' + header + 'LBL_APTotal').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px; font-weight:bold~'
            if ($('#' + header + 'LBL_APCurrent').text() != '') { app = app + '$' + $('#' + header + 'LBL_APCurrent').text(); apcurrent = parseFloat($('#' + header + 'LBL_APCurrent').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px; font-weight:bold~'
            if ($('#' + header + 'LBL_AP30').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP30').text(); ap30 = parseFloat($('#' + header + 'LBL_AP30').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right;  border-top:solid 1px;border-bottom:double 4px;font-weight:bold~'
            if ($('#' + header + 'LBL_AP60').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP60').text(); ap60 = parseFloat($('#' + header + 'LBL_AP60').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px ; font-weight:bold~'
            if ($('#' + header + 'LBL_AP90').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP90').text(); ap90 = parseFloat($('#' + header + 'LBL_AP90').text().replace(/,/g, ''), 10) }
            app = app + '"/>'

            var percurrent = parseFloat(apcurrent, 10) / parseFloat(total, 10)
            var per30 = parseFloat(ap30, 10) / parseFloat(total, 10)
            var per60 = parseFloat(ap60, 10) / parseFloat(total, 10)
            var per90 = parseFloat(ap90, 10) / parseFloat(total, 10)

            if (percurrent == 0) { percurrent = '' } else { percurrent = '(' + (parseFloat(percurrent, 10) * 100).toFixed(1).toString() + '%)' }
            if (per30 == 0) { per30 = '' } else { per30 = '(' + (parseFloat(per30, 10) * 100).toFixed(1).toString() + '%)' }
            if (per60 == 0) { per60 = '' } else { per60 = '(' + (parseFloat(per60, 10) * 100).toFixed(1).toString() + '%)' }
            if (per90 == 0) { per90 = '' } else { per90 = '(' + (parseFloat(per90, 10) * 100).toFixed(1).toString() + '%)' }

            app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px~~font-size:8pt; text-align:right~~font-size:8pt; text-align:right; font-weight:bold~' + percurrent + '~font-size:8pt; text-align:right; font-weight:bold~' + per30 + '~font-size:8pt; text-align:right; font-weight:bold~' + per60 + '~font-size:8pt; text-align:right; font-weight:bold~' + per90 + '"/>'
        }
        else { app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px~' + $(this).text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_APTotal').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_APCurrent').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP30').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP60').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP90').text() + '"/>' }
    });

    $('#printinfo').html(app)
    printReport()
    $('#spinner').hide();
}

// Printing AP
function pre_printAP() {
    $('.tooltiptext').remove()

    var dt = new Date($.now());
    dt = dt.toString().substr(0, 24)

    $('[id$=HF_PrintHeader]').val('text-align:left; width:80px; font-size:8pt~Vendor~text-align:right; font-size:8pt~Total ($)~text-align:right; font-size:8pt~Current ($)~text-align:right; font-size:8pt~30-60 ($)~text-align:right; font-size:8pt~61-90 ($)~text-align:right; font-size:8pt~90+ ($)');
    $('[id$=HF_PrintTitle]').val('<span style="font-size:12pt">Axiom Plastics Inc.<br/>Aged Accounts Payable ' + $('[id$=DDL_Print_Details]').val() + ' Report (' + $('[id$=DDL_Print_Currency]').val() + ')<br/>As Of ' + $('[id$=Date_Print_From]').val() + '<br/></span><span style="font-size:7pt">printed on: ' + dt + '</span>');

    var app = ''
    var i = 1;
    $('[id$=LBL_CustName]').each(function () {
        var header = $(this).attr('id').replace('LBL_CustName', '')
        if ($(this).text() == 'Total') {
            var total = 0;
            var apcurrent = 0;
            var ap30 = 0;
            var ap60 = 0;
            var ap90 = 0;

            app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px;border-top:solid 0px:; font-weight:bold~' + $(this).text() + '~font-size:8pt; border-top:solid 1px;border-bottom:double 4px;text-align:right; font-weight:bold~'
            if ($('#' + header + 'LBL_APTotal').text() != '') { app = app + '$' + $('#' + header + 'LBL_APTotal').text(); total = parseFloat($('#' + header + 'LBL_APTotal').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px; font-weight:bold~'
            if ($('#' + header + 'LBL_APCurrent').text() != '') { app = app + '$' + $('#' + header + 'LBL_APCurrent').text(); apcurrent = parseFloat($('#' + header + 'LBL_APCurrent').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px; font-weight:bold~'
            if ($('#' + header + 'LBL_AP30').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP30').text(); ap30 = parseFloat($('#' + header + 'LBL_AP30').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right;  border-top:solid 1px;border-bottom:double 4px;font-weight:bold~'
            if ($('#' + header + 'LBL_AP60').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP60').text(); ap60 = parseFloat($('#' + header + 'LBL_AP60').text().replace(/,/g, ''), 10) }
            app = app + '~font-size:8pt; text-align:right; border-top:solid 1px;border-bottom:double 4px ; font-weight:bold~'
            if ($('#' + header + 'LBL_AP90').text() != '') { app = app + '$' + $('#' + header + 'LBL_AP90').text(); ap90 = parseFloat($('#' + header + 'LBL_AP90').text().replace(/,/g, ''), 10) }
            app = app + '"/>'

            var percurrent = parseFloat(apcurrent, 10) / parseFloat(total, 10)
            var per30 = parseFloat(ap30, 10) / parseFloat(total, 10)
            var per60 = parseFloat(ap60, 10) / parseFloat(total, 10)
            var per90 = parseFloat(ap90, 10) / parseFloat(total, 10)

            if (percurrent == 0) { percurrent = '' } else { percurrent = '(' + (parseFloat(percurrent, 10) * 100).toFixed(1).toString() + '%)' }
            if (per30 == 0) { per30 = '' } else { per30 = '(' + (parseFloat(per30, 10) * 100).toFixed(1).toString() + '%)' }
            if (per60 == 0) { per60 = '' } else { per60 = '(' + (parseFloat(per60, 10) * 100).toFixed(1).toString() + '%)' }
            if (per90 == 0) { per90 = '' } else { per90 = '(' + (parseFloat(per90, 10) * 100).toFixed(1).toString() + '%)' }

            app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px~~font-size:8pt; text-align:right~~font-size:8pt; text-align:right; font-weight:bold~' + percurrent + '~font-size:8pt; text-align:right; font-weight:bold~' + per30 + '~font-size:8pt; text-align:right; font-weight:bold~' + per60 + '~font-size:8pt; text-align:right; font-weight:bold~' + per90 + '"/>'
        }
        else { app = app + '<input id="PL_' + i + '_HF_PrintLines" type="hidden" value="font-size:8pt; width:250px~' + $(this).text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_APTotal').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_APCurrent').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP30').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP60').text() + '~font-size:8pt; text-align:right~' + $('#' + header + 'LBL_AP90').text() + '"/>' }
    });

    $('#printinfo').html(app)
    printReport()
    $('#spinner').hide();
}

// Change the Currency
$(document).on('change', '[id$=DDL_Print_Currency]', function () {
    if ($('[id$=DDL_Print_Category]').val() == "2") { printpopCustDD(); }//Sales Category is picked
    if ($('[id$=DDL_Print_Category]').val() == "3") { printpopVendDD(); }//Purchases Category is picked
});

// Print Purchase Report
function printPurchReport() {
    var lang = 0 //0 is english, 1 is spanish
    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1 }
    console.log($('[id$=Date_Print_From]').val(), $('[id$=Date_Print_To]').val())
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxAccounting2.aspx',
        data: { action: "SalesSummary", language: lang, date1: $('[id$=Date_Print_From]').val(), date2: $('[id$=Date_Print_To]').val(), cur: $('[id$=DDL_Print_Currency]').val(), type: "P" },
        success: function (data, status, other) {
            $('#printinfo').html(data)
            printReport()
        },
        error: function (data, status, other) { alert(other); }
    });
}

// Print Sales Report
function printSalesReport() {
    var lang = 0 //0 is english, 1 is spanish
    if ($('[id$=DDL_Print_Language]').val() == 1) { lang = 1 }
    $.ajax({
        async: true, type: 'POST', dataType: 'text', url: 'AjaxPrinting.aspx',
        data: { action: "SalesSummary", language: lang, date1: $('[id$=Date_Print_From]').val(), date2: $('[id$=Date_Print_To]').val(), cur: $('[id$=DDL_Print_Currency]').val(), type: "S" },
        success: function (data, status, other) {
            $('#printinfo').html(data)
            printReport()
        },
        error: function (data, status, other) { alert(other); }
    });
}

// Date Restriction
$(document).on('change', '[id$=CB_Q1], [id$=CB_Q2], [id$=CB_Q3], [id$=CB_Q4]', function () {
    if ($('[id$=CB_Q1]').is(':checked') && $('[id$=CB_Q2]').is(':checked') && $('[id$=CB_Q3]').is(':checked')) { $('[id$=CB_Q4]').attr("disabled", true); } else { $('[id$=CB_Q4]').removeAttr("disabled"); }
    if ($('[id$=CB_Q1]').is(':checked') && $('[id$=CB_Q2]').is(':checked') && $('[id$=CB_Q4]').is(':checked')) { $('[id$=CB_Q3]').attr("disabled", true); } else { $('[id$=CB_Q3]').removeAttr("disabled"); }
    if ($('[id$=CB_Q1]').is(':checked') && $('[id$=CB_Q3]').is(':checked') && $('[id$=CB_Q4]').is(':checked')) { $('[id$=CB_Q2]').attr("disabled", true); } else { $('[id$=CB_Q2]').removeAttr("disabled"); }
    if ($('[id$=CB_Q2]').is(':checked') && $('[id$=CB_Q3]').is(':checked') && $('[id$=CB_Q4]').is(':checked')) { $('[id$=CB_Q1]').attr("disabled", true); } else { $('[id$=CB_Q1]').removeAttr("disabled"); }
})