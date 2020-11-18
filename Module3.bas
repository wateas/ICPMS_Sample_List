Attribute VB_Name = "Module3"
Dim rngStartRow As Range, rngLastRow As Range
Dim Header As String
Dim rngSearch As Range

Dim wbkMetals As Workbook
Dim wksCurrent, wksTarget As Worksheet
    
Sub SetVariables(WrkShtName)
    
    Set wbkMetals = ThisWorkbook
    Set wksCurrent = wbkMetals.Worksheets(WrkShtName)
    ' How to do this better?
    Set wksTarget = wbkMetals.Worksheets("ManipulateList")
    
    Set rngStartRow = wksCurrent.Range("E1").End(xlDown)
    Set rngStartRow = rngStartRow.Cells(2, 1)
    
    ' ISSUE: Formulas in range prevent an accurate dynamic range.
    Set rngLastRow = wksCurrent.Range("F" & Rows.Count).End(xlUp)

End Sub

Sub UpdateLibrary()
'******************************************************************************
' Copies all samples from worksheet "PasteBacklog" to - "SampleLibrary" so
' that user can assign shortened sample names.
' Copied from DataPro_Protocol_Build_v2 5/27/2019
' Edited 12/9/19 by WAT
'******************************************************************************

    ' Finds first empty cell in library
    Dim LibraryPlacement As Long
    Dim LibraryStop As Boolean
    Dim rngCurrCell As Range
    
    Dim wksLibrary As Worksheet

    LibraryStop = False
    
    Set wbkMetals = ThisWorkbook
    Set wksLibrary = wbkMetals.Worksheets("SampleLibrary")
    
    Set rngCurCell = wksLibrary.Range("A11")

    Do While LibraryStop = False
        If rngCurCell.Value = "" Then
            LibraryStop = True
            LibraryPlacement = rngCurCell.Row
        Else
            Set rngCurCell = rngCurCell(2, 1)
        End If
    Loop
    
    ' Loops through all sample locations pasted from Cover Sheet and copies only
    ' ones missing from library to library.
    ' Utilizes a VBA Collection
    ' Used Range.Find method, which returns cell value; note that Worksheet.match
    ' method returns position.
    ' https://excelmacromastery.com/excel-vba-collections/
    ' (later note 6/18/19) Believe I originally tried to fill out collection then
    ' copy the whole thing.  The issue was that there was no way to prevent the
    ' same sample location/description being copied over twice.  Solution shown below;
    ' just add each new element of the collection simultaneously
    Dim NewSamples As New Collection
    Dim SearchRange As Range, SearchCell As Range
    Dim CheckCell As Range, CheckRange As Range
    Dim i As Long
    'Dim c As Long
    
    Set SearchRange = Worksheets("PasteCoverSheet").Range("B10:B310")
    Set CheckRange = Worksheets("SampleLibrary").Range("A10:A310")
    i = 1 'doesnt like Set
        ' not sure if collections can be zero indexed...

    For Each SearchCell In SearchRange
        Set CheckCell = CheckRange.Find(SearchCell, LookIn:=xlValues)
        ' Not sure how to have conditional logic to check if SearchCell was already
        ' added to collection
        ' Solution below; add new sample location to search range each iteration
        If CheckCell Is Nothing Then
            NewSamples.Add SearchCell.Value 'method for adding to collection
            Worksheets("SampleLibrary").Cells _
                ((LibraryPlacement + i - 1), 1).Value = NewSamples(i)
                ' Adds new cell to range so that it wont get added again to collection
            i = i + 1
        End If
    Next
    
    wksLibrary.Select

End Sub


Sub DeleteDuplicates()
'******************************************************************************
' Delete duplicate entries in Sample ID library (worksheet "SampleLibrary")
'******************************************************************************

    'Copied from "DataPro_Protocol_Build_v2" 5/22/19
    Worksheets("SampleLibrary").Select
        ' sometimes gives an "object out of range" error if text is
        ' selected/copied on the SampleLibrary sheet
    ActiveSheet.Range("$A$10:$B$310").RemoveDuplicates Columns:=1, Header:=xlNo

    ' Used macro recorder - modified range selections
    Range("A10:B310").Select
    ActiveWorkbook.Worksheets("SampleLibrary").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("SampleLibrary").Sort.SortFields.Add key:=Range( _
        "A10:A310"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("SampleLibrary").Sort
        .SetRange Range("A10:B310")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

Sub SaveLibrary()
'******************************************************************************
' Reminds user to save workbook once sample library has been updated
'******************************************************************************

    ThisWorkbook.Save
    
End Sub

Function TruncateName(CurrentCell As Range) As String
'******************************************************************************
' Uses regular expression to automate shortening of street names for truncated
' sample name.
'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
' To use regEx, must include "Microsoft VBScript Regular Expressions 5.5" in Tools -> References.
'******************************************************************************
    Dim regEx As New RegExp
    Dim strPattern As String
    Dim strInput As String
    Dim strReplace As String
    Dim strOutput As String

    strPattern = "[CDLPRST][dnlrt][.]"

    If strPattern <> "" Then
        strInput = CurrentCell.Value
        strReplace = ""

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            TruncateName = regEx.Replace(strInput, strReplace)
        Else
            TruncateName = strInput
        End If
    End If

End Function

Private Sub Regex_Test()
'******************************************************************************
' Testing regular expression for UDF TruncateName.
'******************************************************************************

    Dim strPattern As String: strPattern = "[CDLPRST][dnlrt][.]"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim MyStr As String

    MyStr = "11025 Graymarsh Pl."

    If strPattern <> "" Then
        strInput = MyStr

        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With

        If regEx.Test(strInput) Then
            Debug.Print ("Matched: " & regEx.Replace(strInput, strReplace))
        Else
            Debug.Print ("Not matched")
        End If
    End If
    
End Sub



