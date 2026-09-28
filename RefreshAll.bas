Attribute VB_Name = "Module1"
'===============================================================================
' RefreshAll.bas
' Optional VBA macro for Global_Supplement_Analytics.xlsx
'
' WHAT THIS DOES
'   - Refreshes all Power Query connections and PivotTables/PivotCaches in the
'     workbook (if/when you add Power Query connections and PivotTables — see
'     README.md "Adding real PivotTables, PivotCharts & Slicers").
'   - Forces a full recalculation of all formulas.
'   - Writes the current date/time into a cell named "LastRefreshed" (or, if
'     that named range doesn't exist yet, into Start_Here!B2).
'
' THIS FILE IS NOT EMBEDDED IN Global_Supplement_Analytics.xlsx.
' The .xlsx format cannot contain macros. To use this macro:
'
'   1. Open Global_Supplement_Analytics.xlsx in desktop Excel (Windows or Mac).
'   2. File > Save As > choose "Excel Macro-Enabled Workbook (*.xlsm)" and save
'      a copy — e.g. Global_Supplement_Analytics.xlsm.
'   3. Press Alt+F11 (Windows) / Fn+Option+F11 (Mac) to open the VBA editor.
'   4. Right-click your workbook's project in the left-hand tree > Insert > Module.
'   5. Open this RefreshAll.bas file in a text editor, copy all of its contents,
'      and paste them into the new module in the VBA editor.
'      (Alternatively: File > Import File... in the VBA editor and select
'      RefreshAll.bas directly — Excel will import it as a module.)
'   6. Close the VBA editor and save the .xlsm file.
'   7. Run the macro: Developer tab > Macros > RefreshAllData > Run
'      (or press Alt+F8, select RefreshAllData, click Run).
'
' MACRO SECURITY
'   - The first time you run a macro in a downloaded file, Excel will show a
'     security warning ("Macros have been disabled"). Click "Enable Content"
'     (Windows) or confirm you trust the file (Mac) to allow it to run.
'   - Only enable macros in files you trust and understand — that's why this
'     project ships the macro as a separate, readable .bas text file instead
'     of a pre-built .xlsm, so you can review exactly what it does first.
'   - This macro does not access the internet, other files, or the Windows/Mac
'     file system beyond the workbook it runs in. It only refreshes data
'     connections already defined in the workbook and writes one timestamp.
'
' The main Global_Supplement_Analytics.xlsx file is fully usable WITHOUT this
' macro or the .xlsm conversion — every KPI and chart works from plain Excel
' formulas alone. This macro is purely a convenience for one-click refresh
' once you've added your own Power Query connections/PivotTables.
'===============================================================================

Sub RefreshAllData()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim conn As WorkbookConnection
    Dim refreshedSomething As Boolean
    refreshedSomething = False

    On Error Resume Next

    ' 1) Refresh all workbook data connections (Power Query / external data)
    For Each conn In ThisWorkbook.Connections
        conn.Refresh
        refreshedSomething = True
    Next conn

    ' 2) Refresh every PivotTable on every sheet (if any have been added)
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
            refreshedSomething = True
        Next pt
    Next ws

    On Error GoTo 0

    ' 3) Force a full recalculation of every formula in the workbook
    Application.CalculateFullRebuild

    ' 4) Stamp the refresh timestamp
    Call StampLastRefreshed

    If refreshedSomething Then
        MsgBox "Data connections and PivotTables refreshed, and all formulas recalculated." & vbCrLf & _
               "Last Refreshed timestamp updated.", vbInformation, "Refresh Complete"
    Else
        MsgBox "No Power Query connections or PivotTables were found to refresh." & vbCrLf & _
               "All formulas were recalculated and the Last Refreshed timestamp was updated." & vbCrLf & vbCrLf & _
               "(This is expected if you haven't added Power Query connections or PivotTables yet " & _
               "— see README.md for how to add them.)", vbInformation, "Refresh Complete"
    End If

End Sub

Private Sub StampLastRefreshed()

    Dim targetCell As Range

    On Error Resume Next
    Set targetCell = ThisWorkbook.Names("LastRefreshed").RefersToRange
    On Error GoTo 0

    If targetCell Is Nothing Then
        ' Fallback: write into Start_Here!B2 if the named range doesn't exist.
        ' (Define a named range called "LastRefreshed" pointing at any cell to
        ' control exactly where this timestamp appears.)
        On Error Resume Next
        Set targetCell = ThisWorkbook.Worksheets("Start_Here").Range("B2")
        On Error GoTo 0
    End If

    If Not targetCell Is Nothing Then
        targetCell.Value = "Last Refreshed: " & Format(Now, "yyyy-mm-dd hh:mm")
        targetCell.NumberFormat = "@"
    End If

End Sub
