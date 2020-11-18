Attribute VB_Name = "Module1"
Option Explicit

Public bln_IMPORT_DONE As Boolean
Public bln_ICPMS_MERGE_DONE As Boolean, bln_HG_MERGE_DONE As Boolean

Dim WrkShtName As String

' For last column and "Analysis Code" column (worksheet Import).
Dim rngLastColumn As Range, rngAnalysisCode As Range
' For worksheet "Import".
Dim rngImportEnd As Range, rngImport As Range
' For worksheets "ICPMS" and "Hg".
Dim rngStartRow As Range, rngLastRow As Range
' Copied range from all worksheets.
' Used for quit procedure to clear coversheet contents.
Dim rngCurrent As Range

Dim wbkMetals As Workbook, wbkBacklog As Workbook
Dim wksImport As Worksheet, wksBacklog As Worksheet
Dim wksCurrent As Worksheet

Sub SetVariables(WrkShtName)
    
    Set wbkMetals = ThisWorkbook
    Set wksImport = wbkMetals.Worksheets("Import")
    Set wksCurrent = wbkMetals.Worksheets(WrkShtName)
    
    ' Sets first row on ICPMS or Hg coversheet where sample info -
    ' will be copied.
    With wksCurrent
        Set rngStartRow = .Cells(.Rows.Count, 1).End(xlUp)
    End With
    Set rngStartRow = rngStartRow.Cells(2, 1)
        
    ' Sets last row of sample information on worksheet Import.
    Set rngImportEnd = wksImport.Range("A2").End(xlDown)
    
    ' Sets column range variable for "Analysis Code" column.
    ' Would be better if you had a concise way to set rngAnalysisCode
    ' from row 1 to the row of rngImportEnd.
    With wksImport
        Set rngLastColumn = .Cells(1, .Columns.Count). _
            End(xlToLeft)
        Set rngCurrent = .Range(.Cells(1, 1), _
            .Cells(1, rngLastColumn.Column))
    End With
    Set rngAnalysisCode = rngCurrent.Find("Analysis Code", _
        LookIn:=xlValues)
    
End Sub

Sub Auto_Open()
'**********************************************************************
' Procedure "Auto_Open" (special name) automatically runs when workbook
' is opened.  Prompts user to verify standard and reagent expiration
' dates, lots numbers, etc.
' WAT 11/11/19
'**********************************************************************
    
    ' Sets procedure control variables for procedure Quit.
    bln_IMPORT_DONE = False
    bln_ICPMS_MERGE_DONE = False
    bln_HG_MERGE_DONE = False
    
    Worksheets("Master List").Select

    Dim strPrompt As String

    strPrompt = "Verify that standard and reagent information is " & _
        "correct."

    ' Prompts user to verify standard and reagent information.
    MsgBox (strPrompt)
    
End Sub

Sub Change_Standards()
'**********************************************************************
' Redirects user to QC / Lot #'s book for changing lot #s, expiration
' dates, etc.
' Created by WAT 11/11/19
'**********************************************************************

    Dim strFilePath As String, strUpdateQC As String
    Dim strMsg1 As String, strMsg2 As String
    
    strFilePath = "R:\DoRC\Environmental Laboratory\" & _
        "Update QC and Lot Numbers\"
    strUpdateQC = "Update QC and Lot Numbers Book.xlsm"
    
    strMsg1 = "Redirecting to workbook Update QC and Lot Numbers."
    strMsg2 = "  Be sure to save when finished updating."
    
    MsgBox (strMsg1 & strMsg2)
    
    ChDir (strFilePath) 'Not necessary, but seems like good practice.
    Workbooks.Open Filename:=(strFilePath & strUpdateQC)
    
End Sub

Sub ProceedImport()
' Created by WAT 11/11/19

    ThisWorkbook.Worksheets("Import").Select

End Sub

Sub ImportBacklog()
'**********************************************************************
' Imports sample information from LIMS backlog (.DAT file, similar to
' .CSV file).
' Taken from ICPMS_Sample_List_Build 11-26-19 and adapted.
'**********************************************************************
        
    Set wbkMetals = ThisWorkbook
    Set wksImport = wbkMetals.Worksheets("Import")
    
    ' Ensures all cells are clear and unhidden.
    wksImport.Range("A2:G1002").ClearContents
    wksImport.Range("A2:G1002").EntireRow.Hidden = False
    
    ' Opens backlog file in excel.
    ChDir "C:\lwuser6"
    Workbooks.OpenText Filename:="C:\lwuser6\BACKLOG.DAT", Origin:=437, _
        StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, _
        1), Array(7, 1), Array(68, 1), Array(78, 1), Array(86, _
        1), Array(126, 1), Array(150, 1)), _
        TrailingMinusNumbers:=True
            ' line break has to be in between "Array(7, _1)"
    
    Set wbkBacklog = ActiveWorkbook
    Set wksBacklog = wbkBacklog.Worksheets(1)
        
    ' Set range variables for Worksheet Backlog.
    Set rngLastRow = wksBacklog.Range("A1").End(xlDown)
    Set rngLastColumn = wksBacklog.Cells(1, wksBacklog.Columns.Count). _
        End(xlToLeft)
    With wksBacklog
        Set rngImport = .Range(.Cells(1, 1), .Cells( _
            rngLastRow.Row, rngLastColumn.Column))
    End With
        
    ' Copy data from Backlog.
    rngImport.Copy Destination:=wksImport.Range("A2")
    
    ' Close backlog.dat one data has been copied from it.
    Application.DisplayAlerts = False
    wbkBacklog.Close

    ' Set range variables for Worksheet Import.
    Set rngImportEnd = wksImport.Range("A2").End(xlDown)
    Set rngLastColumn = wksImport.Cells(1, wksImport.Columns.Count). _
        End(xlToLeft)
    With wksImport
        Set rngImport = .Range(.Cells(2, 1), .Cells( _
            rngImportEnd.Row, rngLastColumn.Column))
    End With
    
    ' Sort backlog list by LIMS number then analyte.
    With wksImport.Sort
        .SortFields.Clear
        .SortFields.Add key:=Range( _
            "A1:A" & rngImportEnd.Row), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add key:=Range( _
            "F1:F" & rngImportEnd.Row), SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange rngImport
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    bln_IMPORT_DONE = True
    
End Sub

Sub ICPMS_Merge()
'**********************************************************************
' Populates ICPMS cover sheet with sample information.
' Created by WAT 11/12/2019
'**********************************************************************
    
    Dim i As Range ' Iterator
    
    If bln_HG_MERGE_DONE = True Then
        MsgBox ("A merge procedure has been executed. Please quit " + _
            "open again to run another merge procedure.")
        End
    End If
    
    ' Clears ICPMS cover sheet if procedure already run.
    If bln_ICPMS_MERGE_DONE = True Then
        rngCurrent.ClearContents
    End If
    
    ' Sets worksheet and range variables.
    WrkShtName = ThisWorkbook.ActiveSheet.Name
    Call SetVariables(WrkShtName)
    
    ' Ensures all cells are unhidden and autofilter is reset on worksheet Import.
    wksImport.Range("A1:D502").EntireRow.Hidden = False
    wksImport.Range("A1:D502").AutoFilter Field:=4

    ' This code interacts with the excel (special) table in that it can
    ' de-select any analysis codes that aren't on the list (spreadsheet
    ' will hide rows with analysis codes not on the list)
    ' Will filter out unwanted analysis codes like "MET_DIG"
    ' Now impractical since the new ICPMS_XX_DW analysis codes have been implemented.
'    WrkSht1.Range("$A$2:$D$302").AutoFilter Field:=4, Criteria1:=Array( _
'        "AG_ICPMS", "AG_ICPMS_SL", "AL_ICPMS", "AL_ICPMS_SL", "AS_ICPMS", "AS_ICPMS_SL", _
'        "BA_ICPMS", "BE_ICPMS", "CA_ICPMS", "CA_ICPMS_SL", "CD_ICPMS", "CD_ICPMS_SL", _
'        "CR_ICPMS", "CR_ICPMS_SL", "CU_ICPMS", "CU_ICPMS_SL", "K_ICPMS", "K_ICPMS_SL", _
'        "FE_ICPMS", "FE_ICPMS_SL", "MN_ICPMS", "MN_ICPMS_SL", "MO_ICPMS", "MO_ICPMS_SL", _
'        "NA_ICPMS", "NA_ICPMS_SL", "NI_ICPMS", "NI_ICPMS_SL", "PB_ICPMS", "PB_ICPMS_SL", _
'        "SB_ICPMS", "SE_ICPMS", "SE_ICPMS_SL", "TL_ICPMS", "ZN_ICPMS", "ZN_ICPMS_SL"), _
'        Operator:=xlFilterValues

    ' Doesn't work; not sure why.
    'wksImport.Range("$A$2:$D$502").AutoFilter Field:=4, _
        Criteria1:="<>MET_DIG", Operator:=xlOr, Criteria2:="<>HG_DIG"

    ' Works but not enough criteria parameters (want to filter out *DRYWT*, others...).
    'wksImport.Range("$A$2").CurrentRegion.AutoFilter Field:=4, _
        Criteria1:="<>MET_DIG", Criteria2:="<>*HG_*"

    ' Does not work.
    ' https://stackoverflow.com/questions/32891223/autofilter-exceptions-with-more-than-two-criteria
    ' https://chandoo.org/forum/threads/vba-code-to-autofilter-and-hide-rows-based-on-criteria.23407/
    'wksImport.Range("$A$2").CurrentRegion.AutoFilter Field:=4, _
        Criteria1:=Array("<>AG_ICPMS", "<>AL_ICPMS", "<>AS_ICPMS"), _
        Operator:=xlFilterValues
    
    ' Set rngCurrent for Analysis Code column before applying filter criteria.
    With wksImport
        Set rngCurrent = .Range(rngAnalysisCode, .Cells( _
            rngImportEnd.Row, rngAnalysisCode.Column))
    End With
    
    ' Filter all non-ICPMS analysis codes using hidden cells.
    For Each i In rngCurrent.Cells
        If i.Value Like "*HG*" Or _
            i.Value Like "*DRYWT*" Or _
            i.Value Like "*DIG*" Or _
            i.Value Like "*WT*" Then
            i.EntireRow.Hidden = 1
        End If
    Next i

    '  To capture only non-hidden cells, need to use -
    ' .SpecialCells(xlCellTypeVisible).
    wksImport.Range("A2:C" & rngImportEnd.Row).SpecialCells(xlCellTypeVisible). _
        Copy Destination:=rngStartRow

    ' Define working range on current worksheet.
    Set rngLastRow = rngStartRow.End(xlDown)
    Set rngCurrent = wksCurrent.Range(rngStartRow, rngLastRow(1, 3))

    ' Delete duplicate instances of LIMS number.
    rngCurrent.RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Sort by LIMS number.
    With wksCurrent.Sort
        .SortFields.Clear
        .SortFields.Add key:=Range( _
            "A" & rngStartRow.Row & ":A" & rngLastRow.Row), _
            SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange rngStartRow.CurrentRegion
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Moves date column over to appropriate location.
    wksCurrent.Range("C" & rngStartRow.Row & ":C" & rngLastRow.Row). _
        Cut Destination:=wksCurrent.Range("E" & rngStartRow.Row)
        
    ' Define working range on current worksheet.
    Set rngLastRow = rngStartRow.End(xlDown)
    Set rngLastColumn = wksCurrent.Cells( _
        rngLastRow.Row, wksCurrent.Columns.Count).End(xlToLeft)
    Set rngCurrent = wksCurrent.Range(rngStartRow, rngLastColumn)

    ' Ensures all cells are unhidden and autofilter is reset on worksheet Import.
    wksImport.Range("A2:D" & rngImportEnd.Row).EntireRow.Hidden = False
    wksImport.Range("A1:D" & rngImportEnd.Row).AutoFilter Field:=4
    
    bln_ICPMS_MERGE_DONE = True

End Sub

Sub Hg_Merge()
'**********************************************************************
' Populates Hg cover sheet with sample information.
' Makes use of dynamic ranges.
' Created by WAT 11/12/19
' Issues:
' 1.) Not sure how to have dynamic range set to clear previous contents when -
'     clearing contents (ranges will stack if macro is run more than once)
'**********************************************************************

    Dim i As Range ' Iterator
      
    If bln_ICPMS_MERGE_DONE = True Then
        MsgBox ("A merge procedure has been executed. Please quit and " + _
            "open again to run another merge procedure.")
        End
    End If
    
    If bln_HG_MERGE_DONE = True Then
        rngCurrent.ClearContents
    End If
    
    ' Sets worksheet and range variables.
    WrkShtName = ThisWorkbook.ActiveSheet.Name
    Call SetVariables(WrkShtName)

    ' Ensures all cells are unhidden and autofilter is reset on worksheet Import.
    wksImport.Range("A1:D" & rngImportEnd.Row).EntireRow.Hidden = False
    wksImport.Range("A1:D" & rngImportEnd.Row).AutoFilter Field:=4
        
'    ' CAN'T USE VARIABLE FOR AUTOFILTER FIELD, SCRAPPING AUTOFILTER
'    ' Uses Autofilter to remove non-Hg analyses from wksImport.
'    rngCurrent.AutoFilter Field:=4, Criteria1:="=HG_CV", Operator:=xlOr, _
'        Criteria2:="=HG_CV_SL"
    
    ' Set rngCurrent for Analysis Code column before applying filter criteria.
    With wksImport
        Set rngCurrent = .Range(rngAnalysisCode, .Cells( _
            rngImportEnd.Row, rngAnalysisCode.Column))
    End With
    
    ' Filter all non-Hg analysis codes using hidden cells.
    For Each i In rngCurrent.Cells
        If i.Value Like "*ICPMS*" Or _
            i.Value Like "*DRYWT*" Or _
            i.Value Like "*DIG*" Or _
            i.Value Like "*WT*" Then
            i.EntireRow.Hidden = 1
        End If
    Next i
    
    ' Difficulty of moving ranges may stem from use of list-filter and hidden rows.
    ' To capture correct range, need to use .SpecialCells(xlCellTypeVisible)
    ' Note: xlCellTypeVisible may not be needed for Autofilter list.
    wksImport.Range("A2:C" & rngImportEnd.Row).SpecialCells(xlCellTypeVisible). _
        Copy Destination:=rngStartRow

    ' Define working range on current worksheet.
    Set rngLastRow = rngStartRow.End(xlDown)
    Set rngCurrent = wksCurrent.Range(rngStartRow, rngLastRow(1, 3))

   ' Moves date column over to appropriate location.
    wksCurrent.Range("C" & rngStartRow.Row & ":C" & rngLastRow.Row). _
        Cut Destination:=wksCurrent.Range("E" & rngStartRow.Row)
    
    ' Define working range on current worksheet.
    Set rngLastRow = rngStartRow.End(xlDown)
    Set rngLastColumn = wksCurrent.Cells( _
        rngLastRow.Row, wksCurrent.Columns.Count).End(xlToLeft)
    Set rngCurrent = wksCurrent.Range(rngStartRow, rngLastColumn)
    
    ' Ensures all cells are unhidden and autofilter is reset on worksheet Import.
    wksImport.Range("A1:D" & rngImportEnd.Row).EntireRow.Hidden = False
    wksImport.Range("A1:D" & rngImportEnd.Row).AutoFilter Field:=4
        
    bln_HG_MERGE_DONE = True
        
End Sub

Sub Quit()
'**********************************************************************
' Using dynamic range variables to clear workbook for next use.
' Needs to be executed next-in-line after either merge procedure run.
' Created by WAT 2019-12-4
'**********************************************************************

    ' Clears ranges on worksheet Import.
    If bln_IMPORT_DONE = True Then
        wksImport.Range("A2:G" & rngImportEnd.Row).ClearContents
    '    MsgBox (bln_IMPORT_DONE & _
            ": Clear Import Contents.") ' For testing.
    'Else
    '    Debug.Print ("Import " & bln_IMPORT_DONE)
    End If
    
    If bln_ICPMS_MERGE_DONE = True Or _
        bln_HG_MERGE_DONE = True _
        Then
        ' Clears range on whichever coversheet where a merge procedure
        ' was executed.
        rngCurrent.ClearContents
        'MsgBox (bln_ICPMS_MERGE_DONE & _
                    ": Clear Merge Contents.") ' For testing.
    'Else
    '    Debug.Print ("Merge " & bln_ICPMS_MERGE_DONE)
    End If
    
    'ThisWorkbook.Save
    Application.DisplayAlerts = False

    'ThisWorkbook.Close ' Only closes workbook, doesn't quit application.
    Application.Quit
           
End Sub


