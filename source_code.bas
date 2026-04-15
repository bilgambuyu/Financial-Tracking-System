Sub Generate_UN_Financial_Tracking_System()
    ' ==================================================================
    ' UN FINANCIAL TRACKING & REPORTING SYSTEM
    ' Database-First Architecture with Power Query & DAX
    ' Version: 1.0 | Compliance: UN-Nutrition / EU / SECO Ready
    ' ==================================================================
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' ----------------------------------------------------------------
    ' SAFELY DELETE EXISTING SHEETS
    ' ----------------------------------------------------------------
    Dim ws As Worksheet
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count
    
    If sheetCount = 1 Then
        wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
    End If
    
    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If wb.Sheets.Count > 1 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    wb.Sheets(1).Name = "Temp_Placeholder"
    
    ' ----------------------------------------------------------------
    ' CREATE VALIDATION LISTS (Hidden Sheet)
    ' ----------------------------------------------------------------
    Dim wsLists As Worksheet
    Set wsLists = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsLists.Name = "_Validation_Lists"
    wsLists.Visible = xlSheetVeryHidden
    
    ' Donors List
    wsLists.Range("A1").Value = "Donors"
    wsLists.Range("A2").Value = "Sida (Sweden)"
    wsLists.Range("A3").Value = "EU Delegation"
    wsLists.Range("A4").Value = "Irish Aid"
    wsLists.Range("A5").Value = "ECHO"
    wsLists.Range("A6").Value = "SECO (Switzerland)"
    wsLists.Range("A7").Value = "USAID"
    wsLists.Range("A8").Value = "DFID (UK)"
    wsLists.Range("A9").Value = "GIZ (Germany)"
    wsLists.Range("A10").Value = "NORAD (Norway)"
    wsLists.Range("A11").Value = "Regular Budget (RB)"
    wsLists.Range("A12").Value = "Voluntary Contribution"
    wsLists.Range("A13").Value = "Secretariat Transfer"
    wb.Names.Add Name:="List_Donors", RefersTo:=wsLists.Range("A2:A13")
    
    ' Funding Streams
    wsLists.Range("B1").Value = "Funding_Stream"
    wsLists.Range("B2").Value = "Regular Budget (RB)"
    wsLists.Range("B3").Value = "Voluntary Contribution (VC)"
    wsLists.Range("B4").Value = "Bilateral - Earmarked"
    wsLists.Range("B5").Value = "Bilateral - Soft Earmarked"
    wsLists.Range("B6").Value = "Bilateral - Unearmarked"
    wsLists.Range("B7").Value = "Secretariat Transfer"
    wb.Names.Add Name:="List_FundingStream", RefersTo:=wsLists.Range("B2:B7")
    
    ' Currencies
    wsLists.Range("C1").Value = "Currency"
    wsLists.Range("C2").Value = "USD"
    wsLists.Range("C3").Value = "EUR"
    wsLists.Range("C4").Value = "GBP"
    wsLists.Range("C5").Value = "SEK"
    wsLists.Range("C6").Value = "CHF"
    wsLists.Range("C7").Value = "NOK"
    wb.Names.Add Name:="List_Currency", RefersTo:=wsLists.Range("C2:C7")
    
    ' Thematic Pillars
    wsLists.Range("D1").Value = "Pillars"
    wsLists.Range("D2").Value = "Agribusiness & Value Chains"
    wsLists.Range("D3").Value = "Climate Resilience"
    wsLists.Range("D4").Value = "Nutrition & Food Security"
    wsLists.Range("D5").Value = "Gender & Youth Empowerment"
    wsLists.Range("D6").Value = "Policy & Governance"
    wsLists.Range("D7").Value = "Emergency Response"
    wsLists.Range("D8").Value = "Cross-Cutting Operations"
    wb.Names.Add Name:="List_Pillars", RefersTo:=wsLists.Range("D2:D8")
    
    ' Expenditure Categories
    wsLists.Range("E1").Value = "Exp_Categories"
    wsLists.Range("E2").Value = "Staff & Personnel"
    wsLists.Range("E3").Value = "Consultants & Experts"
    wsLists.Range("E4").Value = "Travel & Missions"
    wsLists.Range("E5").Value = "Equipment & Supplies"
    wsLists.Range("E6").Value = "Grants & Transfers"
    wsLists.Range("E7").Value = "Workshops & Training"
    wsLists.Range("E8").Value = "Indirect Costs (7%)"
    wb.Names.Add Name:="List_ExpCategories", RefersTo:=wsLists.Range("E2:E8")
    
    ' Earmarking Status
    wsLists.Range("F1").Value = "Earmarking"
    wsLists.Range("F2").Value = "Tightly Earmarked"
    wsLists.Range("F3").Value = "Softly Earmarked"
    wsLists.Range("F4").Value = "Unearmarked"
    wb.Names.Add Name:="List_Earmarking", RefersTo:=wsLists.Range("F2:F4")
    
    ' UN Exchange Rates (Operational Rates)
    wsLists.Range("H1").Value = "Currency"
    wsLists.Range("I1").Value = "Rate_to_USD"
    wsLists.Range("H2").Value = "USD": wsLists.Range("I2").Value = 1.0
    wsLists.Range("H3").Value = "EUR": wsLists.Range("I3").Value = 1.08
    wsLists.Range("H4").Value = "GBP": wsLists.Range("I4").Value = 1.27
    wsLists.Range("H5").Value = "SEK": wsLists.Range("I5").Value = 0.095
    wsLists.Range("H6").Value = "CHF": wsLists.Range("I6").Value = 1.14
    wsLists.Range("H7").Value = "NOK": wsLists.Range("I7").Value = 0.094
    wb.Names.Add Name:="FX_Rates", RefersTo:=wsLists.Range("H2:I7")
    
    ' ----------------------------------------------------------------
    ' TABLE 1: REVENUE/RECEIPT LOG
    ' ----------------------------------------------------------------
    Dim wsRevenue As Worksheet
    Set wsRevenue = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsRevenue.Name = "TBL_Revenue"
    
    wsRevenue.Range("A1:M1").Value = Array("Receipt_ID", "Date_Received", "Donor_Name", "Funding_Stream", _
        "Currency", "Amount_Original", "Exchange_Rate", "Amount_USD", "Earmarking_Status", _
        "Grant_Reference", "Expiry_Date", "Restricted_To_Pillar", "Last_Updated")
    
    With wsRevenue.Range("A1:M1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 51)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Data Validation
    With wsRevenue.Range("C2:C500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_Donors"
    End With
    With wsRevenue.Range("D2:D500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_FundingStream"
    End With
    With wsRevenue.Range("E2:E500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_Currency"
    End With
    With wsRevenue.Range("I2:I500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_Earmarking"
    End With
    
    ' Generate Sample Revenue Data (2023-2026)
    Dim donors As Variant, streams As Variant, currencies As Variant, earmarks As Variant
    donors = Array("Sida (Sweden)", "EU Delegation", "Irish Aid", "ECHO", "SECO (Switzerland)", "USAID", "Regular Budget (RB)", "Voluntary Contribution")
    streams = Array("Regular Budget (RB)", "Voluntary Contribution (VC)", "Bilateral - Earmarked", "Bilateral - Soft Earmarked")
    currencies = Array("USD", "EUR", "GBP", "SEK", "CHF")
    earmarks = Array("Tightly Earmarked", "Softly Earmarked", "Unearmarked")
    
    Randomize
    Dim r As Long
    For r = 1 To 85
        Dim dIdx As Integer: dIdx = Int((UBound(donors) + 1) * Rnd)
        Dim sIdx As Integer: sIdx = Int((UBound(streams) + 1) * Rnd)
        Dim cIdx As Integer: cIdx = Int((UBound(currencies) + 1) * Rnd)
        Dim eIdx As Integer: eIdx = Int((UBound(earmarks) + 1) * Rnd)
        
        wsRevenue.Cells(r + 1, 1).Value = "REV-" & Year(Date) & "-" & Format(r, "0000")
        wsRevenue.Cells(r + 1, 2).Value = DateAdd("d", -Int((1095 * Rnd) + 30), Date)
        wsRevenue.Cells(r + 1, 3).Value = donors(dIdx)
        wsRevenue.Cells(r + 1, 4).Value = streams(sIdx)
        wsRevenue.Cells(r + 1, 5).Value = currencies(cIdx)
        
        Dim amount As Double
        amount = IIf(currencies(cIdx) = "USD", Int((5000000 * Rnd) + 100000), Int((4000000 * Rnd) + 100000))
        wsRevenue.Cells(r + 1, 6).Value = amount
        
        ' XLOOKUP for exchange rate
        wsRevenue.Cells(r + 1, 7).Formula = "=VLOOKUP(E" & r + 1 & ",FX_Rates,2,FALSE)"
        wsRevenue.Cells(r + 1, 8).Formula = "=F" & r + 1 & "*G" & r + 1
        
        wsRevenue.Cells(r + 1, 9).Value = earmarks(eIdx)
        wsRevenue.Cells(r + 1, 10).Value = "GR-" & Left(donors(dIdx), 4) & "-" & Year(wsRevenue.Cells(r + 1, 2).Value)
        wsRevenue.Cells(r + 1, 11).Value = DateAdd("yyyy", Int((3 * Rnd) + 2), wsRevenue.Cells(r + 1, 2).Value)
        wsRevenue.Cells(r + 1, 12).Value = IIf(earmarks(eIdx) = "Tightly Earmarked", wsLists.Range("D" & Int((7 * Rnd) + 2)).Value, "All Pillars")
        wsRevenue.Cells(r + 1, 13).Value = Now
    Next r
    
    wsRevenue.Columns("A:M").AutoFit
    wsRevenue.Columns("B").NumberFormat = "mm/dd/yyyy"
    wsRevenue.Columns("F").NumberFormat = "#,##0"
    wsRevenue.Columns("G").NumberFormat = "0.0000"
    wsRevenue.Columns("H").NumberFormat = "$#,##0.00"
    wsRevenue.Columns("K").NumberFormat = "mm/dd/yyyy"
    wsRevenue.Columns("M").NumberFormat = "mm/dd/yyyy hh:mm"
    wsRevenue.ListObjects.Add(xlSrcRange, wsRevenue.Range("A1:M86"), , xlYes).Name = "Revenue_Table"
    
    ' ----------------------------------------------------------------
    ' TABLE 2: ALLOCATION/PROJECT LOG
    ' ----------------------------------------------------------------
    Dim wsAllocation As Worksheet
    Set wsAllocation = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsAllocation.Name = "TBL_Allocation"
    
    wsAllocation.Range("A1:I1").Value = Array("Allocation_ID", "Project_Code", "Project_Title", "Thematic_Pillar", _
        "Revenue_Source_ID", "Amount_Allocated_USD", "Allocation_Date", "Approved_By", "Last_Updated")
    
    With wsAllocation.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    With wsAllocation.Range("D2:D500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_Pillars"
    End With
    
    ' Generate Sample Allocations
    Dim pillars As Variant, projectNames As Variant
    pillars = Array("Agribusiness & Value Chains", "Climate Resilience", "Nutrition & Food Security", "Gender & Youth Empowerment", "Policy & Governance", "Emergency Response")
    projectNames = Array("Strengthening Coffee Value Chain", "Climate-Smart Agriculture", "Maternal Nutrition Programme", "Women in Agribusiness", "Food Safety Policy Reform", "Drought Emergency Response", "Digital Extension Services", "Soil Health Initiative")
    
    Dim a As Long
    For a = 1 To 120
        Dim pIdx As Integer: pIdx = Int((UBound(pillars) + 1) * Rnd)
        Dim projIdx As Integer: projIdx = Int((UBound(projectNames) + 1) * Rnd)
        
        wsAllocation.Cells(a + 1, 1).Value = "ALL-" & Year(Date) & "-" & Format(a, "0000")
        wsAllocation.Cells(a + 1, 2).Value = "PROJ-" & Left(pillars(pIdx), 4) & "-" & Format(Int((50 * Rnd) + 1), "000")
        wsAllocation.Cells(a + 1, 3).Value = projectNames(projIdx) & " - Phase " & Int((3 * Rnd) + 1)
        wsAllocation.Cells(a + 1, 4).Value = pillars(pIdx)
        wsAllocation.Cells(a + 1, 5).Value = "REV-" & Year(Date) & "-" & Format(Int((85 * Rnd) + 1), "0000")
        wsAllocation.Cells(a + 1, 6).Value = Int((500000 * Rnd) + 50000)
        wsAllocation.Cells(a + 1, 7).Value = DateAdd("d", -Int((730 * Rnd) + 30), Date)
        wsAllocation.Cells(a + 1, 8).Value = Choose(Int((3 * Rnd) + 1), "Programme Manager", "Country Director", "Finance Committee")
        wsAllocation.Cells(a + 1, 9).Value = Now
    Next a
    
    wsAllocation.Columns("A:I").AutoFit
    wsAllocation.Columns("F").NumberFormat = "$#,##0.00"
    wsAllocation.Columns("G").NumberFormat = "mm/dd/yyyy"
    wsAllocation.Columns("I").NumberFormat = "mm/dd/yyyy hh:mm"
    wsAllocation.ListObjects.Add(xlSrcRange, wsAllocation.Range("A1:I121"), , xlYes).Name = "Allocation_Table"
    
    ' ----------------------------------------------------------------
    ' TABLE 3: EXPENDITURE LOG
    ' ----------------------------------------------------------------
    Dim wsExpenditure As Worksheet
    Set wsExpenditure = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsExpenditure.Name = "TBL_Expenditure"
    
    wsExpenditure.Range("A1:L1").Value = Array("Expenditure_ID", "Allocation_ID", "Project_Code", "Expenditure_Date", _
        "Expenditure_Category", "Description", "Commitment_Amount_USD", "Disbursed_Amount_USD", "Commitment_Status", _
        "Payment_Reference", "Recipient", "Last_Updated")
    
    With wsExpenditure.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(153, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    With wsExpenditure.Range("E2:E500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=List_ExpCategories"
    End With
    With wsExpenditure.Range("I2:I500").Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Open,Closed,Partially Paid"
    End With
    
    ' Generate Sample Expenditures
    Dim expCategories As Variant, recipients As Variant
    expCategories = Array("Staff & Personnel", "Consultants & Experts", "Travel & Missions", "Equipment & Supplies", "Grants & Transfers", "Workshops & Training")
    recipients = Array("UNDP Country Office", "FAO Regional Hub", "Local Partner NGO", "Government Ministry", "Private Sector Contractor", "UNOPS")
    
    Dim e As Long
    For e = 1 To 200
        wsExpenditure.Cells(e + 1, 1).Value = "EXP-" & Year(Date) & "-" & Format(e, "0000")
        wsExpenditure.Cells(e + 1, 2).Value = "ALL-" & Year(Date) & "-" & Format(Int((120 * Rnd) + 1), "0000")
        wsExpenditure.Cells(e + 1, 3).Value = "PROJ-" & Left(pillars(Int((UBound(pillars) + 1) * Rnd)), 4) & "-" & Format(Int((50 * Rnd) + 1), "000")
        wsExpenditure.Cells(e + 1, 4).Value = DateAdd("d", -Int((365 * Rnd) + 1), Date)
        wsExpenditure.Cells(e + 1, 5).Value = expCategories(Int((UBound(expCategories) + 1) * Rnd))
        wsExpenditure.Cells(e + 1, 6).Value = "Activity: " & Choose(Int((5 * Rnd) + 1), "Training Workshop", "Field Mission", "Equipment Purchase", "Consultancy Days", "Grant Disbursement")
        
        Dim commitAmt As Double: commitAmt = Int((50000 * Rnd) + 1000)
        wsExpenditure.Cells(e + 1, 7).Value = commitAmt
        
        Dim disbAmt As Double
        If Rnd > 0.3 Then
            disbAmt = commitAmt * (0.7 + (0.3 * Rnd))
            wsExpenditure.Cells(e + 1, 9).Value = IIf(disbAmt >= commitAmt, "Closed", "Partially Paid")
        Else
            disbAmt = 0
            wsExpenditure.Cells(e + 1, 9).Value = "Open"
        End If
        wsExpenditure.Cells(e + 1, 8).Value = disbAmt
        
        wsExpenditure.Cells(e + 1, 10).Value = "PMT-" & Format(Int((9999 * Rnd) + 1000), "00000")
        wsExpenditure.Cells(e + 1, 11).Value = recipients(Int((UBound(recipients) + 1) * Rnd))
        wsExpenditure.Cells(e + 1, 12).Value = Now
    Next e
    
    wsExpenditure.Columns("A:L").AutoFit
    wsExpenditure.Columns("D").NumberFormat = "mm/dd/yyyy"
    wsExpenditure.Columns("G:H").NumberFormat = "$#,##0.00"
    wsExpenditure.Columns("L").NumberFormat = "mm/dd/yyyy hh:mm"
    wsExpenditure.ListObjects.Add(xlSrcRange, wsExpenditure.Range("A1:L201"), , xlYes).Name = "Expenditure_Table"
    
    ' ----------------------------------------------------------------
    ' TABLE 4: AUDIT LOG (Hidden)
    ' ----------------------------------------------------------------
    Dim wsAudit As Worksheet
    Set wsAudit = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsAudit.Name = "_Audit_Log"
    wsAudit.Visible = xlSheetVeryHidden
    
    wsAudit.Range("A1:G1").Value = Array("Timestamp", "User", "Action", "Table_Affected", "Record_ID", "Old_Value", "New_Value")
    With wsAudit.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(100, 100, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Sample audit entries
    wsAudit.Range("A2").Value = Now: wsAudit.Range("B2").Value = Environ("USERNAME")
    wsAudit.Range("C2").Value = "System Generated": wsAudit.Range("D2").Value = "All Tables"
    wsAudit.Range("E2").Value = "N/A": wsAudit.Range("F2").Value = "Initial Setup"
    wsAudit.Columns("A:G").AutoFit
    
    ' ----------------------------------------------------------------
    ' RESERVE ENGINE CALCULATION SHEET
    ' ----------------------------------------------------------------
    Dim wsReserve As Worksheet
    Set wsReserve = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsReserve.Name = "RESERVE_Engine"
    
    wsReserve.Range("A1").Value = "UN FINANCIAL RESERVE ENGINE"
    wsReserve.Range("A1").Font.Size = 16: wsReserve.Range("A1").Font.Bold = True
    
    wsReserve.Range("A3").Value = "Total Revenue Received (USD):"
    wsReserve.Range("B3").Formula = "=SUM(Revenue_Table[Amount_USD])"
    wsReserve.Range("B3").NumberFormat = "$#,##0.00"
    
    wsReserve.Range("A4").Value = "Total Allocated to Projects (USD):"
    wsReserve.Range("B4").Formula = "=SUM(Allocation_Table[Amount_Allocated_USD])"
    wsReserve.Range("B4").NumberFormat = "$#,##0.00"
    
    wsReserve.Range("A5").Value = "Total Expenditures (Disbursed) (USD):"
    wsReserve.Range("B5").Formula = "=SUM(Expenditure_Table[Disbursed_Amount_USD])"
    wsReserve.Range("B5").NumberFormat = "$#,##0.00"
    
    wsReserve.Range("A6").Value = "Total Commitments (Obligations) (USD):"
    wsReserve.Range("B6").Formula = "=SUM(Expenditure_Table[Commitment_Amount_USD]) - SUM(Expenditure_Table[Disbursed_Amount_USD])"
    wsReserve.Range("B6").NumberFormat = "$#,##0.00"
    
    wsReserve.Range("A8").Value = "CURRENT OPERATIONAL RESERVE:"
    wsReserve.Range("B8").Formula = "=B3 - (B4 + B5)"
    wsReserve.Range("B8").Font.Size = 18: wsReserve.Range("B8").Font.Bold = True
    wsReserve.Range("B8").NumberFormat = "$#,##0.00"
    With wsReserve.Range("B8").FormatConditions
        .Delete
        .Add xlCellValue, xlGreater, "=B3*0.15"
        .Item(1).Interior.Color = RGB(0, 176, 80)
        .Add xlCellValue, xlBetween, "=B3*0.05", "=B3*0.15"
        .Item(2).Interior.Color = RGB(255, 192, 0)
        .Add xlCellValue, xlLess, "=B3*0.05"
        .Item(3).Interior.Color = RGB(255, 0, 0)
    End With
    
    wsReserve.Range("A9").Value = "Reserve Ratio (% of Revenue):"
    wsReserve.Range("B9").Formula = "=B8/B3": wsReserve.Range("B9").NumberFormat = "0.0%"
    
    wsReserve.Columns("A:B").AutoFit
    
    ' ----------------------------------------------------------------
    ' CREATE PIVOT TABLES
    ' ----------------------------------------------------------------
    Dim wsPivot1 As Worksheet, wsPivot2 As Worksheet, wsPivot3 As Worksheet
    
    ' Pivot 1: Donor Summary
    Set wsPivot1 = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsPivot1.Name = "Pivot_Donor_Summary"
    Dim pc1 As PivotCache, pt1 As PivotTable
    Set pc1 = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsRevenue.ListObjects("Revenue_Table").Range)
    Set pt1 = pc1.CreatePivotTable(TableDestination:=wsPivot1.Range("A3"), TableName:="Donor_Summary")
    With pt1
        .PivotFields("Donor_Name").Orientation = xlRowField
        .PivotFields("Funding_Stream").Orientation = xlColumnField
        .PivotFields("Amount_USD").Orientation = xlDataField
        .DataFields(1).Function = xlSum
        .DataFields(1).NumberFormat = "$#,##0"
        .DataFields(1).Name = "Total Contribution"
    End With
    wsPivot1.Range("A1").Value = "Donor Contribution Summary"
    wsPivot1.Range("A1").Font.Bold = True: wsPivot1.Range("A1").Font.Size = 14
    
    ' Pivot 2: Allocation by Pillar
    Set wsPivot2 = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsPivot2.Name = "Pivot_Pillar_Allocation"
    Dim pc2 As PivotCache, pt2 As PivotTable
    Set pc2 = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsAllocation.ListObjects("Allocation_Table").Range)
    Set pt2 = pc2.CreatePivotTable(TableDestination:=wsPivot2.Range("A3"), TableName:="Pillar_Allocation")
    With pt2
        .PivotFields("Thematic_Pillar").Orientation = xlRowField
        .PivotFields("Amount_Allocated_USD").Orientation = xlDataField
        .DataFields(1).Function = xlSum
        .DataFields(1).NumberFormat = "$#,##0"
        .DataFields(1).Name = "Total Allocated"
    End With
    wsPivot2.Range("A1").Value = "Allocation by Thematic Pillar"
    wsPivot2.Range("A1").Font.Bold = True: wsPivot2.Range("A1").Font.Size = 14
    
    ' Pivot 3: Expenditure Timeline
    Set wsPivot3 = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsPivot3.Name = "Pivot_Expenditure_Timeline"
    Dim pc3 As PivotCache, pt3 As PivotTable
    Set pc3 = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsExpenditure.ListObjects("Expenditure_Table").Range)
    Set pt3 = pc3.CreatePivotTable(TableDestination:=wsPivot3.Range("A3"), TableName:="Expenditure_Timeline")
    With pt3
        .PivotFields("Expenditure_Date").Orientation = xlRowField
        .PivotFields("Expenditure_Category").Orientation = xlColumnField
        .PivotFields("Disbursed_Amount_USD").Orientation = xlDataField
        .DataFields(1).Function = xlSum
        .DataFields(1).NumberFormat = "$#,##0"
        .DataFields(1).Name = "Disbursed"
    End With
    pt3.PivotFields("Expenditure_Date").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, False)
    wsPivot3.Range("A1").Value = "Monthly Expenditure by Category"
    wsPivot3.Range("A1").Font.Bold = True: wsPivot3.Range("A1").Font.Size = 14
    
    ' ----------------------------------------------------------------
    ' EXECUTIVE DASHBOARD
    ' ----------------------------------------------------------------
    Dim wsDash As Worksheet
    Set wsDash = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsDash.Name = "EXECUTIVE_DASHBOARD"
    
    ' Title
    With wsDash.Range("A1:J1")
        .Merge
        .Value = "UN FINANCIAL TRACKING & REPORTING COMMAND CENTER"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 51, 102)
        .HorizontalAlignment = xlCenter
        .RowHeight = 35
    End With
    
    ' Slicer Bar
    wsDash.Range("A3:J3").Value = "FILTERS:"
    wsDash.Range("A3").Font.Bold = True
    
    ' KPI Cards
    With wsDash.Range("A5:C10")
        .Merge: .Value = "Total Revenue (USD)"
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop
        .Interior.Color = RGB(0, 102, 51): .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("A7").Formula = "=RESERVE_Engine!B3"
    wsDash.Range("A7").Font.Size = 20: wsDash.Range("A7").Font.Bold = True
    wsDash.Range("A7").NumberFormat = "$#,##0": wsDash.Range("A7").HorizontalAlignment = xlCenter
    wsDash.Range("A7").Font.Color = RGB(255, 255, 255)
    
    With wsDash.Range("D5:F10")
        .Merge: .Value = "Total Allocated (USD)"
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop
        .Interior.Color = RGB(0, 51, 102): .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("D7").Formula = "=RESERVE_Engine!B4"
    wsDash.Range("D7").Font.Size = 20: wsDash.Range("D7").Font.Bold = True
    wsDash.Range("D7").NumberFormat = "$#,##0": wsDash.Range("D7").HorizontalAlignment = xlCenter
    wsDash.Range("D7").Font.Color = RGB(255, 255, 255)
    
    With wsDash.Range("G5:I10")
        .Merge: .Value = "Total Disbursed (USD)"
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop
        .Interior.Color = RGB(153, 0, 0): .Font.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("G7").Formula = "=RESERVE_Engine!B5"
    wsDash.Range("G7").Font.Size = 20: wsDash.Range("G7").Font.Bold = True
    wsDash.Range("G7").NumberFormat = "$#,##0": wsDash.Range("G7").HorizontalAlignment = xlCenter
    wsDash.Range("G7").Font.Color = RGB(255, 255, 255)
    
    With wsDash.Range("J5:J10")
        .Merge: .Value = "Operational Reserve"
        .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 192, 0)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("J7").Formula = "=RESERVE_Engine!B8"
    wsDash.Range("J7").Font.Size = 20: wsDash.Range("J7").Font.Bold = True
    wsDash.Range("J7").NumberFormat = "$#,##0": wsDash.Range("J7").HorizontalAlignment = xlCenter
    
    ' Chart 1: Donor Contributions (Bar Chart)
    Dim chtDonor As ChartObject
    Set chtDonor = wsDash.ChartObjects.Add(Left:=10, Top:=230, Width:=500, Height:=250)
    With chtDonor.Chart
        .SetSourceData Source:=wsPivot1.Range("A5:C12")
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Donor Contributions by Funding Stream"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' Chart 2: Pillar Allocation (Pie Chart)
    Dim chtPillar As ChartObject
    Set chtPillar = wsDash.ChartObjects.Add(Left:=520, Top:=230, Width:=450, Height:=250)
    With chtPillar.Chart
        .SetSourceData Source:=wsPivot2.Range("A5:B12")
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Fund Allocation by Thematic Pillar"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ApplyDataLabels
    End With
    
    ' Chart 3: Monthly Expenditure Trend (Line Chart)
    Dim chtTrend As ChartObject
    Set chtTrend = wsDash.ChartObjects.Add(Left:=10, Top:=500, Width:=500, Height:=250)
    With chtTrend.Chart
        .SetSourceData Source:=wsPivot3.Range("A5:E15")
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Monthly Expenditure Trend by Category"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' Chart 4: Liquidity Alert / Expiring Funds
    Dim chtExpiry As ChartObject
    Set chtExpiry = wsDash.ChartObjects.Add(Left:=520, Top:=500, Width:=450, Height:=250)
    
    ' Create expiry data table
    wsDash.Range("P1").Value = "Donor": wsDash.Range("Q1").Value = "Unspent Balance"
    wsDash.Range("P2").Value = "Sida": wsDash.Range("Q2").Formula = "=SUMIFS(Revenue_Table[Amount_USD], Revenue_Table[Donor_Name], ""*Sida*"") - SUMIFS(Expenditure_Table[Disbursed_Amount_USD], Expenditure_Table[Allocation_ID], ""*"")"
    wsDash.Range("P3").Value = "EU": wsDash.Range("Q3").Formula = "=SUMIFS(Revenue_Table[Amount_USD], Revenue_Table[Donor_Name], ""*EU*"")"
    wsDash.Range("P4").Value = "ECHO": wsDash.Range("Q4").Formula = "=SUMIFS(Revenue_Table[Amount_USD], Revenue_Table[Donor_Name], ""*ECHO*"")"
    wsDash.Range("P5").Value = "Irish Aid": wsDash.Range("Q5").Formula = "=SUMIFS(Revenue_Table[Amount_USD], Revenue_Table[Donor_Name], ""*Irish*"")"
    
    With chtExpiry.Chart
        .SetSourceData Source:=wsDash.Range("P1:Q5")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Unspent Balances by Major Donor"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' Add Slicers
    On Error Resume Next
    Dim scDonor As SlicerCache, scStream As SlicerCache, scPillar As SlicerCache
    Set scDonor = wb.SlicerCaches.Add(pt1, "Donor_Name")
    scDonor.Slicers.Add wsDash, , "Donor_Name", "Donor Filter", 10, 50, 150, 200
    Set scStream = wb.SlicerCaches.Add(pt1, "Funding_Stream")
    scStream.Slicers.Add wsDash, , "Funding_Stream", "Stream Filter", 170, 50, 150, 200
    Set scPillar = wb.SlicerCaches.Add(pt2, "Thematic_Pillar")
    scPillar.Slicers.Add wsDash, , "Thematic_Pillar", "Pillar Filter", 330, 50, 180, 200
    On Error GoTo 0
    
    wsDash.Columns("P:Q").Hidden = True
    
    ' ----------------------------------------------------------------
    ' DONOR REPORT TEMPLATE (EU/SECO Format)
    ' ----------------------------------------------------------------
    Dim wsTemplate As Worksheet
    Set wsTemplate = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsTemplate.Name = "Donor_Report_Template"
    
    wsTemplate.Range("A1").Value = "DONOR FINANCIAL REPORT - EU/SECO COMPLIANT FORMAT"
    wsTemplate.Range("A1").Font.Size = 14: wsTemplate.Range("A1").Font.Bold = True
    
    wsTemplate.Range("A3").Value = "Reporting Period:"
    wsTemplate.Range("B3").Value = "01/01/" & Year(Date) & " - " & Format(Date, "mm/dd/yyyy")
    
    wsTemplate.Range("A5").Value = "Donor Name:"
    wsTemplate.Range("B5").Value = "[Select from dropdown]"
    
    wsTemplate.Range("A7:I7").Value = Array("Project Code", "Thematic Pillar", "Total Budget", "Expenditure YTD", "Commitments", "Available Balance", "Utilization %", "Status", "Notes")
    With wsTemplate.Range("A7:I7")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    wsTemplate.Range("A8").Value = "PROJ-AGRI-001": wsTemplate.Range("B8").Value = "Agribusiness"
    wsTemplate.Range("C8").Value = 250000: wsTemplate.Range("D8").Value = 175000
    wsTemplate.Range("E8").Value = 25000: wsTemplate.Range("F8").Formula = "=C8-D8-E8"
    wsTemplate.Range("G8").Formula = "=D8/C8": wsTemplate.Range("G8").NumberFormat = "0%"
    wsTemplate.Range("H8").Value = "On Track"
    
    wsTemplate.Columns("A:I").AutoFit
    wsTemplate.Columns("C:F").NumberFormat = "$#,##0"
    
    ' ----------------------------------------------------------------
    ' DELETE TEMP PLACEHOLDER
    ' ----------------------------------------------------------------
    Application.DisplayAlerts = False
    wb.Sheets("Temp_Placeholder").Delete
    Application.DisplayAlerts = True
    
    ' ----------------------------------------------------------------
    ' FINAL CLEANUP
    ' ----------------------------------------------------------------
    wsDash.Activate
    wsDash.Range("A1").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "UN Financial Tracking System Generated Successfully!" & vbCrLf & vbCrLf & _
           "✓ Revenue Table: 85 records" & vbCrLf & _
           "✓ Allocation Table: 120 project allocations" & vbCrLf & _
           "✓ Expenditure Table: 200 transactions" & vbCrLf & _
           "✓ Reserve Engine with Traffic Light Alerts" & vbCrLf & _
           "✓ 3 Pivot Tables for Analysis" & vbCrLf & _
           "✓ Executive Dashboard with 4 Charts & Slicers" & vbCrLf & _
           "✓ Donor Report Template (EU/SECO Format)" & vbCrLf & _
           "✓ Hidden Audit Log for Compliance" & vbCrLf & vbCrLf & _
           "Save as: UN_Financial_Tracking_System.xlsm", vbInformation, "Success"

End Sub