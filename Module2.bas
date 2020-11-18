Attribute VB_Name = "Module2"
Sub ElementsBySample_python()
'*************************************************************************************
' 8/24/20 WAT
' This is the replacement procedure which utilizes xlwings to run a python script that
' returns the sample information with one row per unique sample with a corresponding
' list of analytes in an easy to read layout.
' The old procedure is commented out below.  It can be used again if difficulty is had
' with xlwings.
'*************************************************************************************
    
    mymodule = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
    'RunPython ("import " & mymodule & ";" & mymodule & ".main(" & rngCurrent & ")")
    RunPython ("import ICPMS_Sample_List; ICPMS_Sample_List.main(" & rngCurrent & ")")
    
End Sub

'Dim WrkShtName As String
'
'' For last column and "Analysis Code" column (worksheet Import).
'Dim rngLastColumn As Range ', rngAnalysisCode As Range
'' For worksheet "Import".
'Dim rngImportEnd As Range
'' For worksheets "ICPMS", "Hg", "ElementsBySample".
'Dim rngStartRow As Range, rngLastRow As Range
'' Copied range from all worksheets.
'Dim rngCurrent As Range
'
'Dim wbkMetals As Workbook, wbkBacklog As Workbook
'Dim wksImport As Worksheet, wksBacklog As Worksheet
'Dim wksCurrent As Worksheet
'
'Public Function Contains(col As Collection, key As Variant) As Boolean
'' https://stackoverflow.com/questions/137845/determining-whether-an-object-is-a-member-of-a-collection-in-vba
'
'Dim obj As Variant
'
'On Error GoTo err
'    Contains = True
'    IsObject (col.Item(key))
'    Exit Function
'err:
'    Contains = False
'
'End Function
'
'Sub SetVariables(WrkShtName)
'
'    Set wbkMetals = ThisWorkbook
'    Set wksImport = wbkMetals.Worksheets("Import")
'    Set wksCurrent = wbkMetals.Worksheets(WrkShtName)
'
'    ' Set Range Variables for worksheet "Import"
'    With wksImport
'        Set rngImportEnd = .Range("A1").End(xlDown)
'        Set rngLastColumn = .Cells( _
'            rngImportEnd.Row, .Columns.Count).End(xlToLeft)
'        Set rngCurrent = .Range(.Cells(2, 1), rngLastColumn)
'    End With
'
''    Debug.Print ("rngImportEnd: " & rngImportEnd.Row)
''    Debug.Print ("rngLastColumn: " & rngLastColumn.Column)
'
'End Sub
'
'Sub SetCurrentRange()
'' Set Range Variables for current worksheet.
'
'    With wksCurrent
'        Set rngLastRow = .Range("A1").End(xlDown)
'        Set rngLastColumn = .Cells( _
'            rngLastRow.Row, .Columns.Count).End(xlToLeft)
'        Set rngCurrent = .Range(.Cells(2, 1), rngLastColumn)
'    End With
'
'End Sub
'Sub ElementsBySample()
'' VBA FREEZES if Worksheet "Import" is not selected when running procedure
'    ' Did not freeze 7/9/20 when active sheet was Hg and this procedure was run; problem fixed?
'
'    ' Iterators.
'    Dim Sample As Range, i As Variant, Count As Long
'
'    ' Variables to set dynamic ranges for worksheet Import.
'    Dim rngCurrentColumn As Range
'    Dim rngSampleLoc As Range, rngAnalysisCode As Range
'    Dim rngLimsNum As Range
'
'    ' Variables for generating new list of samples, each with a list of elements.
'    Dim rngSearch As Range, rngSearchAfter As Range
'    Dim AddElement As Range, CurrentCell As Range
'    Dim rngNewTable As Range
'
'    'WrkShtName = ThisWorkbook.ActiveSheet.Name
'    WrkShtName = "ElementsBySample"
'
'    Call SetVariables(WrkShtName)
'
'    ' Ensures all target ranges are clear.
'    wksCurrent.Range("A2:F500", "J2:O100").ClearContents
'
'    ' Copy data from worksheet "Import" to worksheet "ElementsBySample".
'    rngCurrent.Copy Destination:=wksCurrent.Range("A2")
'    ' Clear unnecessary Location Code column.
'    ' Column "Location Code 2" sometimes has had pertinent info, so left that in place.
'    wksCurrent.Range("G:G").ClearContents
'
'    ' Set Range Variables for current worksheet.
'    Call SetCurrentRange
'
''    Dim rngSampleColumn As Range
''    'Dim SampleRow As Long ' Use rngImportEnd
''    Dim SampleDescription As Range, AnalysisCodes As Range, LimsNumbers As Range
'
'    Dim colFilter As New Collection
'
'    colFilter.Add "MET_DIG"
'    colFilter.Add "HG_CV"
'    colFilter.Add "HG_CV_SL"
'    colFilter.Add "HG_DIG"
'    colFilter.Add "DRYWT"
'    colFilter.Add "SLDG_WT"
'    colFilter.Add "SLG_WT_HG"
'    colFilter.Add "SLG_WT_H" '8/27/19 not sure why I'm seeing "SLG_WT_H" and not "SLG_WT_HG" in test data
'
'    '''
'    ' Filter out rows with non-pertinent analyses and truncates analysis codes to bare elements
'    ' Deletes duplicates
'    ' Sets dynamic range that is length of analysis codes column
'
'    ' Find columns automatically, but is this needed considering that the -
'    ' import header is made by you?
'    With wksCurrent
'        ' Sets dyanamic range for LIMS # column.
'        lngCurrentColumn = .Range("A1:Z1").Find("LIMS #").Column
'        Set rngLimsNum = .Range(.Cells(2, lngCurrentColumn), _
'            .Cells(rngLastRow.Row, lngCurrentColumn))
'        ' Sets dynamic range for Sample Description column.
'        lngCurrentColumn = .Range("A1:Z1").Find("Sample Location").Column
'        Set rngSampleLoc = .Range(.Cells(2, lngCurrentColumn), _
'            .Cells(rngLastRow.Row, lngCurrentColumn))
'        ' Sets dynamic range for Analysis Code column.
'        lngCurrentColumn = .Range("A1:Z1").Find("Analysis Code").Column
'        Set rngAnalysisCode = .Range(.Cells(2, lngCurrentColumn), _
'            .Cells(rngLastRow.Row, lngCurrentColumn))
'    End With
''    Debug.Print (wksCurrent.Name)
''    Debug.Print ("Sample Loc: " & rngSampleLoc.Column)
''    Debug.Print ("Analysis Code: " & rngAnalysisCode.Column)
''    Debug.Print ("Lims Number: " & rngLimsNum.Column)
'
'
'    For Each Sample In rngAnalysisCode
'        'See https://stackoverflow.com/questions/7851859/delete-a-row-in-excel-vba
'        For Each i In colFilter
'            ' Iterator for collection must be declared as variant or object.
'            ' Can only declare as object when ALL items in collection are objects.
'            If Sample = i Then
'                Sample.EntireRow = ""
'            ElseIf InStr(1, Sample, i) Then
'                Sample.EntireRow = ""
'            Else
'                Sample.Replace What:="_ICPMS", Replacement:=""
'                Sample.Replace What:="_SL", Replacement:=""
'            End If
'        Next i
'    Next Sample
'
'    ' Sort after clearing rows with non-pertinent analyses
'    With wksCurrent.Sort
'        .SortFields.Clear
'        .SortFields.Add key:=rngLimsNum, SortOn:=xlSortOnValues, _
'            Order:=xlAscending, DataOption:=xlSortNormal
'        .SortFields.Add key:=rngAnalysisCode, SortOn:=xlSortOnValues, _
'            Order:=xlAscending, DataOption:=xlSortNormal
'        .SetRange rngCurrent
'        .Header = xlTrue
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    Call SetCurrentRange
'
'    ' Delete duplicate entries of the LIMS# AND analyte.
'    rngCurrent.RemoveDuplicates Columns:= _
'        Array(rngLimsNum.Column, rngAnalysisCode.Column), Header:=xlYes
'
'    Call SetCurrentRange
'
'    '''
'    ' It would be ideal if you could have a dictionary with each key a new LIMS# and each item a collection -
'    ' - that can have new elements added (the collection doesn't have to be fully created before storing -
'    ' - in the dictionary; can be modified after insertion into the dictionary.
'        ' At the moment, only intent is to make analyte list more readable to user, but an iterable object would be
'        ' - more useful.
'        ' Tried these sources without luck:
'        ' https://stackoverflow.com/questions/38254337/how-to-create-dynamic-variable-names-vba
'        ' https://excelmacromastery.com/vba-dictionary/#Example_2_8211_Dealing_with_Multiple_Values
'        ' https://stackoverflow.com/questions/35444816/vba-can-i-put-a-collection-inside-a-scripting-dictionary
'    ' Main Strategy: Application.Match returns "relative position in array"; returns error if not found
'    ' Application.Match and Range.Find functions seem to freeze up VBA if search range is on a different worksheet
'    ' So had to build new table on same sheet.
'    ' Best practice to use dissimilar variable names; got into trouble using LimsNumbers and LimsNumber in For Each loop
'    ' Variables: SearchRange, Count, Last, LimsNumbers,
'
'    ' Where new table will be built (on the same worksheet).
'    Set rngSearch = wksCurrent.Range("J:J")
'    ' Separate iterator that will be the placeholder for the newly generated table.
'    ' Count will progress with each new sample location found and added to table.
'    Count = 2
'    'Placeholder that Find function will use for the search-after parameter.
'    Set rngSearchAfter = wksCurrent.Range("A2")
'
'    ' For this loop to work, sample information table needs to already be sorted by LIMS# then analyte.
'    For Each Sample In rngLimsNum
'    'For i = 1 To 10 'for testing without For Each Loop
'    'Set Sample = rngSampleLoc(i)
'        If Not IsError(Application.Match(Sample, rngSearch, 0)) Then
'        'Variable "Sample" (i.e. LIMS #)  is already in SearchRange; just add element
'            Set AddElement = rngLimsNum.Find(Sample, After:=rngSearchAfter)
'                ' Range.Find won't search before current LimsNumber
'            CurrentCell(1, 6) = CurrentCell(1, 6) & ", " & AddElement(1, 6)
'                ' ERROR "variable not set" - only if new range (at column J) not deleted (?)
'            'MsgBox (Sample & " already in search range - i = " & i & "; Count = " & Count & _
'                "; AddElement.Row = " & AddElement.Row) 'FOR TESTING
'        Else
'        'LimsNumber is not already in SearchRange and needs to be added w/ first element
'            wksCurrent.Range(Sample(1, 1), Sample(1, 6)).Copy _
'                Destination:=rngSearch(Count)
'            Set CurrentCell = rngSearch(Count)
'                'variable to use for adding elements to same LIMS#/row
'            Count = Count + 1
'            'MsgBox (Sample & " NOT in SearchRange - added; i = " & i & " Count = " & Count) ' FOR TESTING
'        End If
'        Set rngSearchAfter = Sample
'    Next Sample
'
''   'A different approach would be using a collection instead of a range.
''    Dim colSearch As New Collection
''    'Dim c As Long
''
''    For Each Sample In rngLimsNum
''        For Each i In colSearch
''            If Sample = i Then
''                rngAnalysisCode
''                Sample.EntireRow = ""
''            Else
''                colSearch.Add Sample
''            End If
'
''    ' Attempted using function "Contains", but couldn't get it to work.
''    ' Would be an elegant solution if it worked.
''    For Each i In rngSampleLoc
''        Debug.Print (i)
''        If Contains(colSearch, i) = True Then
''            Debug.Print ("True")
''        Else
''            Debug.Print ("False")
''            colSearch.Add i
''        End If
''    Next i
'
'    ' Cuts new table and places it in appropriate range.
'     rngCurrent.ClearContents
'    Set rngNewTable = wksCurrent.Range("J2").CurrentRegion
'    rngNewTable.Cut Destination:=wksCurrent.Range("A2")
'
'End Sub




