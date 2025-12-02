---
title: Automating Financial Processes in Excel
description:  Using VBA to Automate a Financial Allocation Process in Excel
date: 2025-12-02 00:00:00+0000
categories:
    - Projects
tags:
    - Excel
    - VBA
    - Automation
---

In payroll accounting, some tasks just feel like they’re built to test your patience. Payroll Accounting Adjustments (PAAs) are definitely one of them. Month after month, they require the same copy-paste routines, manual matching, and formatting gymnastics. While I personally didn’t have much trouble completing them, I noticed something important: my coworkers did.

Everyone had their own way of handling PAAs, and that inconsistency made the process slower, more error-prone, and incredibly frustrating for the team. Some people struggled with the volume of data, others with the formatting, and others with double-checking variances manually. Even with instructions, the sheer number of steps made it easy to miss something.

So I wondered: What if we didn’t have to do these steps manually at all?
And that’s when I decided to automate the entire PAA workflow using VBA.

Below is the code I created to autoamte the allocation process. It was a challenging but worthwhile experience.

```diff
Sub DeleteMonthlyTabs()
    Dim monthNames As Variant
    Dim ws As Worksheet
    Dim monthName As Variant
    Dim wsFound As Boolean
    
    ' List of month names from June to May
    monthNames = Array("June", "July", "August", "September", "October", _
                       "November", "December", "January", "February", _
                       "March", "April", "May")
    
    ' Loop through each month name
    For Each monthName In monthNames
        On Error Resume Next
        ' Attempt to set the worksheet object
        Set ws = ThisWorkbook.Sheets(monthName)
        On Error GoTo 0
        
        ' If the worksheet exists, delete it
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False ' Disable delete confirmation
            ws.Delete
            Application.DisplayAlerts = True ' Re-enable delete confirmation
            Set ws = Nothing ' Reset the worksheet object
        End If
    Next monthName
End Sub

Sub CombinedFromTo()
    ' invoke previous sub first to clear already created sheets before generating
    DeleteMonthlyTabs
    RefreshData

    Dim selectedMonths As Collection
    Dim wsFactors As Worksheet, wsMonthlySplit As Worksheet, wsNew As Worksheet, wsToTest As Worksheet, wsCOA As Worksheet
    Dim tblSalaryDetail As ListObject, wdTable As ListObject
    Dim tblEmployeeComp As ListObject
    Dim row As ListRow
    Dim journalPeriod As String, journalPeriodTo As String
    Dim positionID As String
    Dim payComponent As String
    Dim valueToSum As Double
    Dim tiedKey As Variant
    Dim basePayTotal As Double, sickTotal As Double, vacationTotal As Double
    Dim holidayTotal As Double, floatingHolidayTotal As Double, adminLeaveTotal As Double
    Dim monthName As Variant
    Dim monthStart As Date, monthEnd As Date
    Dim tiedKeyRow As Long, payComponentRow As Long
    
    ' Variables used for populating TO data
    Dim matchCol As Long
    Dim matchRow As Long
    Dim toTestRow As Long
    Dim percentageValue As Variant
    Dim cumulativeTotal As Double
    
    ' Variables used for tracking Annual Pay changes
    Dim rng As Range
    Dim effectiveDate As Date
    Dim lastValue As Variant
    Dim changeCount As Integer
    Dim dateChanges() As Date
    Dim idx As Integer
    Dim basePayProposedArr() As Double
    Dim mostRecentEffectiveDateBeforeJune As Date
    Dim mostRecentBasePayProposedBeforeJune As Double
    Dim lastEntryFoundBeforeJune As Boolean
    
    'Capture today's date and the fiscal year
    Dim tdDate As Date
    Dim fYear As Long
    tdDate = Date
    Debug.Print (tdDate)
    
    If tdDate > "5/31/" & Year(tdDate) Then
        fYear = Year(tdDate) + 1
    Else
        fYear = Year(tdDate)
    End If
    
    ' Grab projections input from this sheet
    Set wsToTest = ThisWorkbook.Worksheets("PS_Copy")
    
    ' COA: Grab COA sheet
    Set wsCOA = ThisWorkbook.Worksheets("Master")

    ' Initialize collection for selected months
    Set selectedMonths = New Collection

    ' Show the custom user form to select months and retrieve positionID
    UserForm_MonthSelector.Show
    
    ' Get the Position ID value from the TextBox on the UserForm
    positionID = wsToTest.Range("H8").Value ' UserForm_MonthSelector.tb_PosID.Value
    If positionID = "" Then
        MsgBox "No Position ID provided. Macro will now exit.", vbExclamation, "Missing Input"
        Exit Sub
    End If

    ' Add selected months to the collection based on the checkboxes in the UserForm - generate in FY order
    If UserForm_MonthSelector.cb_June.Value Then selectedMonths.Add "June"
    If UserForm_MonthSelector.cb_July.Value Then selectedMonths.Add "July"
    If UserForm_MonthSelector.cb_August.Value Then selectedMonths.Add "August"
    If UserForm_MonthSelector.cb_September.Value Then selectedMonths.Add "September"
    If UserForm_MonthSelector.cb_October.Value Then selectedMonths.Add "October"
    If UserForm_MonthSelector.cb_November.Value Then selectedMonths.Add "November"
    If UserForm_MonthSelector.cb_December.Value Then selectedMonths.Add "December"
    If UserForm_MonthSelector.cb_January.Value Then selectedMonths.Add "January"
    If UserForm_MonthSelector.cb_February.Value Then selectedMonths.Add "February"
    If UserForm_MonthSelector.cb_March.Value Then selectedMonths.Add "March"
    If UserForm_MonthSelector.cb_April.Value Then selectedMonths.Add "April"
    If UserForm_MonthSelector.cb_May.Value Then selectedMonths.Add "May"

    ' Check if at least one month is selected
    If selectedMonths.Count = 0 Then
        MsgBox "No months were selected. Macro will now exit.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Reference the "Monthly_Split" tab
    On Error Resume Next
    Set wsMonthlySplit = ThisWorkbook.Sheets(CStr(UserForm_MonthSelector.tb_wp.Value) & "_Split")
    On Error GoTo 0
    If wsMonthlySplit Is Nothing Then
        MsgBox "Sheet 'Monthly_Split' does not exist in this workbook.", vbCritical, "Error"
        Exit Sub
    End If

    ' Reference the "Factors" tab
    On Error Resume Next
    Set wsFactors = ThisWorkbook.Sheets("Factors")
    On Error GoTo 0
    If wsFactors Is Nothing Then
        MsgBox "The 'Factors' tab could not be found.", vbCritical, "Error"
        Exit Sub
    End If

    ' Reference the WD_SALARYDETAIL table on the Factors sheet
    On Error Resume Next
    Set wdTable = wsFactors.ListObjects("WD_SALARYDETAIL")
    On Error GoTo 0
    If wdTable Is Nothing Then
        MsgBox "The 'WD_SALARYDETAIL' table could not be found on the 'Factors' sheet.", vbCritical, "Error"
        Exit Sub
    End If

    ' Reference the Employee_Comp table on the AnnualPay sheet
    On Error Resume Next
    Set tblEmployeeComp = ThisWorkbook.Sheets("AnnualPay").ListObjects("Employee_Comp")
    On Error GoTo 0
    If tblEmployeeComp Is Nothing Then
        MsgBox "The 'Employee_Comp' table could not be found on the 'AnnualPay' sheet.", vbCritical, "Error"
        Exit Sub
    End If

    ' Retrieve EmployeeID, EmployeeName, Pay Effective Date, and BasePayProposed using the Position ID
    Dim colPosID As Long, colWorker As Long, colPayEffDate As Long, colBasePayProposed As Long, EmployeeID As String, EmployeeName As String, BasePayProposed As Double
    colPosID = tblEmployeeComp.ListColumns("PosID").Index
    colWorker = tblEmployeeComp.ListColumns("Worker").Index
    colPayEffDate = tblEmployeeComp.ListColumns("Effective Date").Index
    colBasePayProposed = tblEmployeeComp.ListColumns("Base Pay - Proposed").Index

    ' What to do if error when getting information.
    On Error Resume Next
    Dim FoundRow As ListRow
    Set FoundRow = tblEmployeeComp.ListRows(WorksheetFunction.Match(positionID, tblEmployeeComp.ListColumns(colPosID).DataBodyRange, 0))
    On Error GoTo 0

    If FoundRow Is Nothing Then
        MsgBox "The Position ID was not found in the Employee_Comp table.", vbExclamation, "Error"
        Exit Sub
    Else
    EmployeeID = FoundRow.Range.Cells(1, tblEmployeeComp.ListColumns("Employee ID").Index).Value
    EmployeeName = FoundRow.Range.Cells(1, colWorker).Value
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' tracking annual pay changes
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Find matching rows for Position ID
    Dim posColumn As Range, effDateColumn As Range, basePayCurrentCol As Range, basePayProposedCol As Range
    Set posColumn = tblEmployeeComp.ListColumns("PosID").DataBodyRange
    Set effDateColumn = tblEmployeeComp.ListColumns("Effective Date").DataBodyRange
    Set basePayProposedCol = tblEmployeeComp.ListColumns("Base Pay - Proposed").DataBodyRange
            
    changeCount = 0
    lastValue = ""
    idx = 0
    lastEntryFoundBeforeJune = False
            
    ' Loop through rows
    For Each rng In posColumn
        If CStr(rng.Value) = CStr(positionID) Then
            effectiveDate = rng.Offset(0, effDateColumn.Column - posColumn.Column).Value
            
            ' Capture most recent effective date before June 1, 2025 [FY HARDCODED]
            ' adjust so that an entry isn't missed like for P00005148 if effective date is right on 6/1/2025
'            If effectiveDate < DateSerial(2025, 6, 1) Then
            If effectiveDate <= DateSerial(fYear - 1, 6, 1) Then
                If Not lastEntryFoundBeforeJune Or effectiveDate > mostRecentEffectiveDateBeforeJune Then
                    mostRecentEffectiveDateBeforeJune = effectiveDate
                    mostRecentBasePayProposedBeforeJune = rng.Offset(0, basePayProposedCol.Column - posColumn.Column).Value
                    lastEntryFoundBeforeJune = True
                End If
            End If
            
            ' Track changes in effective date and only store values for current FY [FY HARDCODED]
            If effectiveDate > DateSerial(fYear - 1, 5, 31) And effectiveDate <> lastValue Then
                changeCount = changeCount + 1
                lastValue = effectiveDate
                
                ' Store date change
                ReDim Preserve dateChanges(idx)
                dateChanges(idx) = effectiveDate
                
                ' Store corresponding pay values
                ReDim Preserve basePayProposedArr(idx)
                basePayProposedArr(idx) = rng.Offset(0, basePayProposedCol.Column - posColumn.Column).Value
                
                idx = idx + 1
            End If
        End If
    Next rng
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' initial values
    Dim z: z = UBound(dateChanges) - LBound(dateChanges) 'start with most dated values in arrays
    BasePayProposed = mostRecentBasePayProposedBeforeJune 'starting salary before comencement of FY [outside of array]

    ' [HARDCODED] Change for next fiscal year.
    ' Process each selected month
    For Each monthName In selectedMonths

        ' Define month start and end dates
        Select Case monthName
            Case "June"
                monthStart = DateSerial(fYear - 1, 6, 1)
                monthEnd = DateSerial(fYear - 1, 6, 30)
            Case "July"
                monthStart = DateSerial(fYear - 1, 7, 1)
                monthEnd = DateSerial(fYear - 1, 7, 31)
            Case "August"
                monthStart = DateSerial(fYear - 1, 8, 1)
                monthEnd = DateSerial(fYear - 1, 8, 31)
            Case "September"
                monthStart = DateSerial(fYear - 1, 9, 1)
                monthEnd = DateSerial(fYear - 1, 9, 30)
            Case "October"
                monthStart = DateSerial(fYear - 1, 10, 1)
                monthEnd = DateSerial(fYear - 1, 10, 31)
            Case "November"
                monthStart = DateSerial(fYear - 1, 11, 1)
                monthEnd = DateSerial(fYear - 1, 11, 30)
            Case "December"
                monthStart = DateSerial(fYear - 1, 12, 1)
                monthEnd = DateSerial(fYear - 1, 12, 31)
            Case "January"
                monthStart = DateSerial(fYear, 1, 1)
                monthEnd = DateSerial(fYear, 1, 31)
            Case "February"
                monthStart = DateSerial(fYear, 2, 1)
                monthEnd = DateSerial(fYear, 2, 28)
            Case "March"
                monthStart = DateSerial(fYear, 3, 1)
                monthEnd = DateSerial(fYear, 3, 31)
            Case "April"
                monthStart = DateSerial(fYear, 4, 1)
                monthEnd = DateSerial(fYear, 4, 30)
            Case "May"
                monthStart = DateSerial(fYear, 5, 1)
                monthEnd = DateSerial(fYear, 5, 31)
        End Select

        ' Duplicate the "Monthly_Split" tab and rename it
        wsMonthlySplit.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsNew.Name = monthName
        wsNew.Tab.Color = RGB(144, 238, 144)

        ' Place start and end dates into cells G9 and H9
        With wsNew
            .Range("G9").Value = monthStart
            .Range("H9").Value = monthEnd

            ' Populate EmployeeID, EmployeeName, and BasePayProposed into respective cells
            .Range("H1").Value = EmployeeID
            .Range("H2").Value = EmployeeName
            
            ' fixed below so that it is based on the Salary during that (monthly) period in FY25 instead of grabbing most recent
            If monthStart = dateChanges(z) Then
                BasePayProposed = basePayProposedArr(z)
                If z > 0 Then
                    z = z - 1 'only decrement to the most recent entry if within array bounds
                End If
            End If
            .Range("H5").Value = BasePayProposed
        End With

        ' Initialize starting rows
        tiedKeyRow = 18
        payComponentRow = 19

        ' Reset totals
        basePayTotal = 0
        sickTotal = 0
        vacationTotal = 0
        holidayTotal = 0
        floatingHolidayTotal = 0
        adminLeaveTotal = 0
        
        ' Fixed row loop so that the same Worktag is not repeated and corressponding components are properly grouped and summed
        Dim processedTiedKeys As Collection
        Set processedTiedKeys = New Collection ' To track already processed tied keys
        
        For Each row In wdTable.ListRows
            If row.Range(1, wdTable.ListColumns("Journal Line Position ID").Index).Value = positionID And _
               row.Range(1, wdTable.ListColumns("Journal Source").Index).Value = "Payroll Actual Accrual" Then
        
                journalPeriod = row.Range(1, wdTable.ListColumns("Journal Period").Index).Value ' Get Journal Period
        
                ' Process only rows matching the current month
                If journalPeriod = Left(monthName, 3) Then
                    payComponent = row.Range(1, wdTable.ListColumns("Pay Component").Index).Value ' Get Pay Component
                    valueToSum = row.Range(1, wdTable.ListColumns("Transaction Amount").Index).Value ' Get the value for the row
        
                    ' Capture the tied key (Program, Grant, or Gift)
                    If row.Range(1, wdTable.ListColumns("Program").Index).Value <> "" Then
                        tiedKey = Split(row.Range(1, wdTable.ListColumns("Program").Index).Value, " ")(0)
                    ElseIf row.Range(1, wdTable.ListColumns("Grant").Index).Value <> "" Then
                        tiedKey = Split(row.Range(1, wdTable.ListColumns("Grant").Index).Value, " ")(0)
                    ElseIf row.Range(1, wdTable.ListColumns("Gift").Index).Value <> "" Then
                        tiedKey = Split(row.Range(1, wdTable.ListColumns("Gift").Index).Value, " ")(0)
                    Else
                        tiedKey = "Undefined"
                    End If
        
                    ' Check if the tied key has already been processed
                    Dim keyExists As Boolean
                    keyExists = False
                    On Error Resume Next
                    keyExists = Not IsError(processedTiedKeys(tiedKey))
                    On Error GoTo 0
        
                    If Not keyExists Then
                        ' Add the tied key to the collection to mark it as processed
                        processedTiedKeys.Add tiedKey, tiedKey
        
                        ' Initialize totals for this tied key
                        basePayTotal = 0
                        sickTotal = 0
                        vacationTotal = 0
                        holidayTotal = 0
                        floatingHolidayTotal = 0
                        adminLeaveTotal = 0
        
                        ' Nested loop to process all rows corresponding to this tied key
                        For Each subRow In wdTable.ListRows
                            If subRow.Range(1, wdTable.ListColumns("Journal Line Position ID").Index).Value = positionID And _
                                subRow.Range(1, wdTable.ListColumns("Journal Source").Index).Value = "Payroll Actual Accrual" And _
                                ((subRow.Range(1, wdTable.ListColumns("Program").Index).Value Like tiedKey & "*") Or _
                                (subRow.Range(1, wdTable.ListColumns("Grant").Index).Value Like tiedKey & "*") Or _
                                (subRow.Range(1, wdTable.ListColumns("Gift").Index).Value Like tiedKey & "*")) Then
                                
                                subJournalPeriod = subRow.Range(1, wdTable.ListColumns("Journal Period").Index).Value ' Get Journal Period
        
                                ' Process only rows matching the current month
                                If subJournalPeriod = Left(monthName, 3) Then
                                    subPayComponent = subRow.Range(1, wdTable.ListColumns("Pay Component").Index).Value ' Get Pay Component
                                    subValueToSum = subRow.Range(1, wdTable.ListColumns("Transaction Amount").Index).Value ' Get the value for the row
        
                                    ' Add value to the appropriate total
                                    ' TODO: Bring in remaining components and add to the appropriate category below
                                    Select Case subPayComponent
                                        Case "Admin Leave"
                                            adminLeaveTotal = adminLeaveTotal + subValueToSum
                                        ' TODO: modify PQ to filter Activity Pay by SC (**since not all AP's go to Base**)
                                        Case "Activity Pay (Not Reported)", "Activity Pay (Reported)", "Base- ACP", "Base Pay", "Disaster Double Time Pay", "Disaster Pay", "Education", "FWS- Academic Year", "FWS- Overflow (Summer-1)", "FWS- Summer1", "FWS- Summer2", "Parental Leave", "Regular", "Retro (Hourly)", "Sea Pay", "Severance (onetime)", "Shift  Diff 10%", "Shift  Diff 6%", "Student Onetime Pay"
                                            If subPayComponent = "Activity Pay (Not Reported)" Then
                                                basePayTotal = basePayTotal
                                            Else
                                                basePayTotal = basePayTotal + subValueToSum
                                            End If
                                        Case "Floating Holiday"
                                            floatingHolidayTotal = floatingHolidayTotal + subValueToSum
                                        Case "Holiday (Exempt)"
                                            holidayTotal = holidayTotal + subValueToSum
                                        Case "Sick / EIB Pay"
                                            sickTotal = sickTotal + subValueToSum
                                        Case "Vacation / PTO", "Vacation / PTO Payout"
                                            vacationTotal = vacationTotal + subValueToSum
                                    End Select
                                End If
                            End If
                        Next subRow
        
                        ' Write tiedKey and component totals to the sheet after processing all rows for this tiedKey
                        wsNew.Cells(tiedKeyRow, "B").Value = tiedKey ' Tied Key
                        wsNew.Cells(payComponentRow, "C").Value = basePayTotal ' Base Pay Total
                        wsNew.Cells(payComponentRow + 1, "C").Value = sickTotal ' Sick Pay Total
                        wsNew.Cells(payComponentRow + 2, "C").Value = vacationTotal ' Vacation Pay Total
                        wsNew.Cells(payComponentRow + 3, "C").Value = holidayTotal ' Holiday Pay Total
                        wsNew.Cells(payComponentRow + 4, "C").Value = floatingHolidayTotal ' Floating Holiday Pay Total
                        wsNew.Cells(payComponentRow + 5, "C").Value = adminLeaveTotal ' Admin Leave Pay Total
        
                        ' Increment tiedKeyRow and payComponentRow for the next entry (row offset of 7)
                        tiedKeyRow = tiedKeyRow + 7
                        payComponentRow = payComponentRow + 7
                    End If
                End If
            End If
        Next row
    
        ' insert logic for TO DRIVER
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' need a secondary journalPeriod(To) variable to prevent interference with primary var
        ' reset journalPeriod with next monthName as monthly loop progresses
        Select Case monthName
            Case "January": journalPeriodTo = "Jan %"
            Case "February": journalPeriodTo = "Feb %"
            Case "March": journalPeriodTo = "Mar %"
            Case "April": journalPeriodTo = "Apr %"
            Case "May": journalPeriodTo = "May %"
            Case "June": journalPeriodTo = "Jun %"
            Case "July": journalPeriodTo = "Jul %"
            Case "August": journalPeriodTo = "Aug %"
            Case "September": journalPeriodTo = "Sep %"
            Case "October": journalPeriodTo = "Oct %"
            Case "November": journalPeriodTo = "Nov %"
            Case "December": journalPeriodTo = "Dec %"
            Case Else: journalPeriodTo = "S-May"
        End Select
        
        ' Display the entered Journal Period for verification
   '     wsNew.Range("M32").Value = journalPeriodTo ' View Journal Period used
    
        ' Find matching column for the month in "Test_PS" row 11
        On Error Resume Next
        matchCol = wsToTest.Rows(18).Find(What:=journalPeriodTo, LookIn:=xlValues, LookAt:=xlWhole).Column
        On Error GoTo 0 ' Reset error handling after search
    
        If matchCol = 0 Then
            MsgBox "No matching month found in row 11 of 'PS_Copy' tab.", vbCritical
            Exit Sub
        End If
    
        ' Start searching for a match only in rows greater than row 12
        On Error Resume Next
        matchRow = wsToTest.Columns(matchCol).Find(What:="%", After:=wsToTest.Cells(18, matchCol), LookIn:=xlValues, LookAt:=xlPart).row
        On Error GoTo 0 ' Reset error handling after search
    
        If matchRow = 0 Or matchRow <= 18 Then
            MsgBox "No matching percentage value found below row 12 in column for the month on 'PS_Copy' tab.", vbCritical
            Exit Sub
        End If
    
        ' **Fix for Error 1004: Ensure matchRow and matchCol are within valid range**
        If matchRow > wsToTest.Rows.Count Or matchCol > wsToTest.Columns.Count Then
            MsgBox "Invalid row or column detected. Macro will exit.", vbCritical
            Exit Sub
        End If
    
        ' Retrieve percentage value safely
        On Error Resume Next
        percentageValue = wsToTest.Cells(matchRow, matchCol).Value
        On Error GoTo 0 ' Reset error handling
    
        ' Fill B25 (B18) or the next available blank cell with the value from Column B safely
        If matchRow > 0 Then
            Dim targetCell, targetCell2 As Range
            Set targetCell = wsNew.Range("B18")
            Set targetCell2 = wsNew.Range("G18")
            toTestRow = 25 'offset by 7 for the 2nd To worktag entry after the initial one is detected to be used in the Do While cumulativeTotal < 1 loop below
            Dim searchValue As String
            
            searchValue = Trim(wsToTest.Cells(matchRow, 7).Value)
            searchValue = FindCOAValue(searchValue)
            Debug.Print searchValue & "---"
            
            ' Loop until a blank cell is found
            Do While targetCell.Value <> ""
                If targetCell.Value = searchValue Then
                    Exit Do
                End If
                'Debug.Print targetCell.Value
                'Debug.Print wsToTest.Cells(matchRow, 7).Value
                Set targetCell = targetCell.Offset(7, 0) ' Move down by 7 rows
                Set targetCell2 = targetCell2.Offset(7, 0) ' Move down by 7 rows
                toTestRow = toTestRow + 7
            Loop
        
            ' Write the value to the first available blank cell
            targetCell.Value = searchValue
            
            ' Ensure percentageValue is valid
            If IsNumeric(percentageValue) And Not IsEmpty(percentageValue) Then
                percentageValue = CDbl(percentageValue)
                targetCell2.Value = percentageValue
            Else
                percentageValue = 0
            End If
        End If
    
        ' If percentage is less than 1, continue filling values conditionally
        If percentageValue < 1 Then
            cumulativeTotal = percentageValue
            
    
            Do While cumulativeTotal < 1 Or (journalPeriodTo = "S-May" And cumulativeTotal < 2)
                matchRow = matchRow + 1 ' Move to the succeeding row in 'Test_PS'
                
                ' **Ensure matchRow doesn't exceed worksheet limits**
                If matchRow > wsToTest.Rows.Count Then
                    MsgBox "Reached end of data in 'PS_Copy'. Macro will exit.", vbCritical
                    Exit Sub
                End If
    
                On Error Resume Next
                percentageValue = wsToTest.Cells(matchRow, matchCol).Value
                
                On Error GoTo 0 ' Reset error handling
    
                ' Ensure percentage value is valid
                ' **************** [HARDCODED] UPDATE '2026' VALUE FOR NEW FISCAL YEAR ************************************
                If IsNumeric(percentageValue) And Not IsEmpty(percentageValue) And percentageValue <> 0 And wsToTest.Cells(matchRow, 13).Value = fYear Then
                    percentageValue = CDbl(percentageValue)
                    cumulativeTotal = cumulativeTotal + percentageValue
    
                    ' Fill values in the "Test" tab
                    Dim scanDriverLoc As Long
                    scanDriverLoc = 18
                    
                    searchValue = Trim(wsToTest.Cells(matchRow, 7).Value)
                    searchValue = FindCOAValue(searchValue)
                    
                    Do While scanDriverLoc <> toTestRow
                        If wsNew.Cells(scanDriverLoc, "B").Value = wsToTest.Cells(matchRow, 7).Value Then
                        Debug.Print matchRow
                        Debug.Print wsToTest.Cells(matchRow, 7).Value
                        Debug.Print wsNew.Cells(scanDriverLoc, "B").Value
                            Exit Do
                        End If
                        scanDriverLoc = scanDriverLoc + 7
                    Loop
                    
                    wsNew.Cells(scanDriverLoc, "G").Value = percentageValue ' Next percentage value
                    
                    If matchRow > 0 And scanDriverLoc = toTestRow Then
                        wsNew.Cells(toTestRow, "B").Value = wsToTest.Cells(matchRow, 7).Value ' Column B reference
                    End If
    
                    ' Increment toTestRow only if percentageValue is valid
                    toTestRow = toTestRow + 7
                End If
            Loop
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Next monthName

    ' Inform the user
    MsgBox "Sheets created.", vbInformation, "Success"
End Sub

Function FindCOAValue(searchValue As String) As Variant
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim rng As Range, cell As Range
    Dim matches As Collection
    Dim candidateValue As String
    
    Set ws = ThisWorkbook.Worksheets("Master")
    
    lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    Set matches = New Collection
    Set rng = ws.Range("A1:A" & lastrow)
    
    For Each cell In rng
        If Trim(cell.Value) = searchValue Then
            matches.Add cell.row
            Debug.Print (ws.Cells(cell.row, "B"))
        End If
    Next cell
    
    If matches.Count = 1 Then
        FindCOAValue = ws.Cells(matches(1), "B")
        Exit Function
    End If
    
    If matches.Count > 1 Then
        candidateValue = searchValue & "-3"
        
        For Each r In matches
            If ws.Cells(r, "B").Value = candidateValue Then
                FindCOAValue = ws.Cells(r, "B").Value
                Exit Function
            End If
        Next r
        
        FindCOAValue = ws.Cells(matches(1), "B").Value
    End If
End Function

Sub RefreshData()
    ThisWorkbook.RefreshAll
    MsgBox "Data has been refreshed successfully."
End Sub
```
