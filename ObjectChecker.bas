Attribute VB_Name = "ObjectChecker1"
Dim a As Variant 'this is the first dimension of the array 123
Dim counter1 As Integer 'this is the number of rows
Dim c As Integer 'this exits if 20 blanks are found
Dim myarray(1 To 1000, 1 To 1000) As Variant
Dim d As Integer 'this is the no of columns along Number can be found
Dim e As Integer 'this is the no of rows down Number can be found (i.e. header line)
Dim f As Integer 'this is the no of columns along Delivery can be found
Sub CopyToObjectCheckerForm()
Call Test_If_File_Is_Open
Call Sheet_Test
'below is the position in the first dimension of the array
a = 1
Call locaterow
Call locatecolumn
Call locatecolumnofDelivery
Call countnumbersincolumna
Call ArtNoToArray
Call pastearrayintoobjectchecker
Call fandrtosap
Call pastefontfrom13
Range("a1").Select
Call Highlightduplicateswithyellowfill
Call fandr_dash
Call datasort
End Sub
Sub Test_If_File_Is_Open() 'Exits if file Project checker form.xlsm does not exist
    Dim TestWorkbook As Workbook
    Set TestWorkbook = Nothing
    On Error Resume Next
    Set TestWorkbook = Workbooks("project checker form.xlsm")
    On Error GoTo 0
    If TestWorkbook Is Nothing Then
        MsgBox "'Project checker form.xlsm' is not open"
        End
    End If
End Sub
Sub Sheet_Test() 'Exits if sheet Object Checker - step 2 does not exist
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = Workbooks("project checker form.xlsm").Sheets("Object Checker - step 2") 'Workbooks("project checker form.xlsm").Activate
    If Err.Number <> 0 Then
        MsgBox "The sheet 'Object Checker - step 2' doesn't exist"
        End
    End If
End Sub
Sub locaterow()
Range("a1").Select
e = 0
    Do Until ActiveCell.Value = "Art no"
    ActiveCell.Offset(1, 0).Select
    e = e + 1
        If e >= 20 Then
            MsgBox "'Art no' must be in first twenty cells of column A in the active worksheet!"
            End 'Exit everything if row is not found
        End If
    Loop
End Sub
Sub locatecolumnofDelivery()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
f = 0  'this counts how far to right Number is
    Do Until (ActiveCell.Value = "Delivery")
            ActiveCell.Offset(0, 1).Select
            f = f + 1
            If f >= 20 Then
            MsgBox "The heading 'Delivery' must be aligned with the heading 'Art no.'"
                End
            End If
    Loop
End Sub
Sub countnumbersincolumna()
Range("a1").Select
counter1 = Range("A" & Rows.Count).End(xlUp).Row - e
End Sub
Sub locatecolumn()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
d = 0  'this counts how far to right Number is
    Do Until (ActiveCell.Value = "Number")
            ActiveCell.Offset(0, 1).Select
            d = d + 1
            If d >= 20 Then
            MsgBox "The heading 'Number' must be aligned with the heading 'Art no.'"
                End
            End If
    Loop
End Sub
Sub ArtNoToArray()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
Dim y As Integer 'this is the counter for how many times it seeks an article number
Dim zz As Variant 'this is the first six characters of a bracketed item
c = 1 'Set c to 1 so that when code is re-run it doesn't exit sub immediately
'Go down a cell until a six-figure number or mnumber is found then copy to array
For y = 1 To counter1
        Do Until (100000 <= ActiveCell.Value And ActiveCell.Value <= 999999) _
            Or (Len(ActiveCell) = 9 And Left(ActiveCell, 1) = "M") _
            Or (Len(ActiveCell) >= 17 And Len(ActiveCell) <= 25) And Right(ActiveCell, 1) = ")" _
            Or (Len(ActiveCell) = 8 And Mid(ActiveCell, 4, 1) = "-")
            ActiveCell.Offset(1, 0).Select
            'exit if 20 blanks are found
            If IsEmpty(ActiveCell) = True Then
                c = c + 1
            End If
            If c >= 20 Then
                Exit Sub
            End If
        Loop
line1:
            If Right(ActiveCell, 1) = ")" Then 'if bracketed number then copy out the six-figures
            zz = Left(ActiveCell.Value, 6)
            myarray(a, 1) = zz
            GoTo line2:
            End If
        myarray(a, 1) = ActiveCell.Value 'this copies the value found into the array
        'and do this
line2:  myarray(a, 2) = ActiveCell.Offset(0, d).Value 'this is where Number (quantity) is
        myarray(a, 3) = ActiveCell.Offset(0, f).Value 'this is where Delivery (PU etc.) is
            If ActiveCell.Offset(0, d + 3).Value = "+" Then 'if column m has a "+" then copy l to array
            myarray(a, 4) = ActiveCell.Offset(0, d + 2).Value
            Else: myarray(a, 4) = "" 'if you don’t add an else then the array holds the value of the last value
            End If
        a = a + 1
        ActiveCell.Offset(1, 0).Select
'repeat all this for number of rows there are in worksheet
Next y
End Sub

Sub pastearrayintoobjectchecker()
Dim i As Integer
i = 1
Workbooks("project checker form.xlsm").Activate
Range("b13").Select
For i = 1 To a
        'exit if a position in the array is empty
        If IsEmpty(myarray(i, 1)) = True Then
        Exit Sub
        End If
    ActiveCell.Value = myarray(i, 1)
    ActiveCell.Offset(0, 1).Value = myarray(i, 2)
    ActiveCell.Offset(0, 2).Value = myarray(i, 3)
    ActiveCell.Offset(0, 3).Value = myarray(i, 4)
    ActiveCell.Offset(1, 0).Select
    Next i
End Sub

'find schucal output codes and replace with sap ones
Sub fandrtosap()
Cells.Replace What:="PU", Replacement:="PAC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Number", Replacement:="PCE", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Sta", Replacement:="PCE", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Pair", Replacement:="PAA", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
End Sub
      
Sub pastefontfrom13()
    Range("B13").Select
    Selection.Copy
    Range("B14:B1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C13").Select
    Selection.Copy
    Range("C14:C1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D13").Select
    Selection.Copy
    Range("D14:D1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("E20:G20").Select
    Rows("100:1000").Select
    Selection.RowHeight = 19.5
    Range("B987").Select
        Range("D20").Select
End Sub

Sub Highlightduplicateswithyellowfill()
'Remove all conditional formatting to begin with
Cells.FormatConditions.Delete
Dim a As Variant 'this is the first dimension of the array 123
Dim counter1 As Integer 'this is the number of rows
Dim c As Integer 'this exits if 20 blanks are found
Dim myarray(1 To 1000, 1 To 1000) As Variant
Dim d As Integer 'this is the no of columns along Number can be found
Dim e As Integer 'this is the no of rows down Number can be found (i.e. header line)
Dim f As Integer 'this is the no of columns along Delivery can be found
Sub CopyToObjectCheckerForm()
Call Test_If_File_Is_Open
Call Sheet_Test
'below is the position in the first dimension of the array
a = 1
Call locaterow
Call locatecolumn
Call locatecolumnofDelivery
Call countnumbersincolumna
Call ArtNoToArray
Call pastearrayintoobjectchecker
Call fandrtosap
Call pastefontfrom13
Range("a1").Select
Call Highlightduplicateswithyellowfill
Call fandr_dash
Call datasort
End Sub
Sub Test_If_File_Is_Open() 'Exits if file Project checker form.xlsm does not exist
    Dim TestWorkbook As Workbook
    Set TestWorkbook = Nothing
    On Error Resume Next
    Set TestWorkbook = Workbooks("project checker form.xlsm")
    On Error GoTo 0
    If TestWorkbook Is Nothing Then
        MsgBox "'Project checker form.xlsm' is not open"
        End
    End If
End Sub
Sub Sheet_Test() 'Exits if sheet Object Checker - step 2 does not exist
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = Workbooks("project checker form.xlsm").Sheets("Object Checker - step 2") 'Workbooks("project checker form.xlsm").Activate
    If Err.Number <> 0 Then
        MsgBox "The sheet 'Object Checker - step 2' doesn't exist"
        End
    End If
End Sub
Sub locaterow()
Range("a1").Select
e = 0
    Do Until ActiveCell.Value = "Art no"
    ActiveCell.Offset(1, 0).Select
    e = e + 1
        If e >= 20 Then
            MsgBox "'Art no' must be in first twenty cells of column A in the active worksheet!"
            End 'Exit everything if row is not found
        End If
    Loop
End Sub
Sub locatecolumnofDelivery()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
f = 0  'this counts how far to right Number is
    Do Until (ActiveCell.Value = "Delivery")
            ActiveCell.Offset(0, 1).Select
            f = f + 1
            If f >= 20 Then
            MsgBox "The heading 'Delivery' must be aligned with the heading 'Art no.'"
                End
            End If
    Loop
End Sub
Sub countnumbersincolumna()
Range("a1").Select
counter1 = Range("A" & Rows.Count).End(xlUp).Row - e
End Sub
Sub locatecolumn()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
d = 0  'this counts how far to right Number is
    Do Until (ActiveCell.Value = "Number")
            ActiveCell.Offset(0, 1).Select
            d = d + 1
            If d >= 20 Then
            MsgBox "The heading 'Number' must be aligned with the heading 'Art no.'"
                End
            End If
    Loop
End Sub
Sub ArtNoToArray()
Range("a1").Select
ActiveCell.Offset(e, 0).Select
Dim y As Integer 'this is the counter for how many times it seeks an article number
Dim zz As Variant 'this is the first six characters of a bracketed item
c = 1 'Set c to 1 so that when code is re-run it doesn't exit sub immediately
'Go down a cell until a six-figure number or mnumber is found then copy to array
For y = 1 To counter1
        Do Until (100000 <= ActiveCell.Value And ActiveCell.Value <= 999999) _
            Or (Len(ActiveCell) = 9 And Left(ActiveCell, 1) = "M") _
            Or (Len(ActiveCell) >= 17 And Len(ActiveCell) <= 25) And Right(ActiveCell, 1) = ")" _
            Or (Len(ActiveCell) = 8 And Mid(ActiveCell, 4, 1) = "-")
            ActiveCell.Offset(1, 0).Select
            'exit if 20 blanks are found
            If IsEmpty(ActiveCell) = True Then
                c = c + 1
            End If
            If c >= 20 Then
                Exit Sub
            End If
        Loop
line1:
            If Right(ActiveCell, 1) = ")" Then 'if bracketed number then copy out the six-figures
            zz = Left(ActiveCell.Value, 6)
            myarray(a, 1) = zz
            GoTo line2:
            End If
        myarray(a, 1) = ActiveCell.Value 'this copies the value found into the array
        'and do this
line2:  myarray(a, 2) = ActiveCell.Offset(0, d).Value 'this is where Number (quantity) is
        myarray(a, 3) = ActiveCell.Offset(0, f).Value 'this is where Delivery (PU etc.) is
            If ActiveCell.Offset(0, d + 3).Value = "+" Then 'if column m has a "+" then copy l to array
            myarray(a, 4) = ActiveCell.Offset(0, d + 2).Value
            Else: myarray(a, 4) = "" 'if you don’t add an else then the array holds the value of the last value
            End If
        a = a + 1
        ActiveCell.Offset(1, 0).Select
'repeat all this for number of rows there are in worksheet
Next y
End Sub

Sub pastearrayintoobjectchecker()
Dim i As Integer
i = 1
Workbooks("project checker form.xlsm").Activate
Range("b13").Select
For i = 1 To a
        'exit if a position in the array is empty
        If IsEmpty(myarray(i, 1)) = True Then
        Exit Sub
        End If
    ActiveCell.Value = myarray(i, 1)
    ActiveCell.Offset(0, 1).Value = myarray(i, 2)
    ActiveCell.Offset(0, 2).Value = myarray(i, 3)
    ActiveCell.Offset(0, 3).Value = myarray(i, 4)
    ActiveCell.Offset(1, 0).Select
    Next i
End Sub

'find schucal output codes and replace with sap ones
Sub fandrtosap()
Cells.Replace What:="PU", Replacement:="PAC", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Number", Replacement:="PCE", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Sta", Replacement:="PCE", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
Cells.Replace What:="Pair", Replacement:="PAA", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=True
End Sub
      
Sub pastefontfrom13()
    Range("B13").Select
    Selection.Copy
    Range("B14:B1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C13").Select
    Selection.Copy
    Range("C14:C1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("D13").Select
    Selection.Copy
    Range("D14:D1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("E20:G20").Select
    Rows("100:1000").Select
    Selection.RowHeight = 19.5
    Range("B987").Select
        Range("D20").Select
End Sub

Sub Highlightduplicateswithyellowfill()
'Remove all conditional formatting to begin with
Cells.FormatConditions.Delete
Columns("B:B").Select
    Range("B3").Activate
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
Columns("C:C").FormatConditions.Delete
Columns("D:D").FormatConditions.Delete
End Sub
Sub datasort()
Range("E12:G1000").Select
    Selection.UnMerge
    Range("B13:E13").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("Object Checker - step 2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Object Checker - step 2").Sort.SortFields.Add Key _
        :=Range("B13:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Object Checker - step 2").Sort
        .SetRange Range("B13:E1000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   Range("E12:G12").Select
       With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
Range("a1").Select
End Sub

Sub fandr_dash()
Columns(e).Select
Selection.Replace What:="-", Replacement:=""
End Sub

'The assumptions are:
'a) All numbers to be copied are 6-figures, or 9 figures and begin with "M"
'or are between 17 and 25 characters ending in a bracket
'b) There is not a gap of more than 30 blanks between art. numbers
'c) a + is used to denote any special requirements in m and l contains the special data
'd) A contains all the articles numbers J is 'quantity and K is code in sheet
'e) This will work for up to a thousand article numbers

'Outstanding questions:
'1. What should go into non-standard requirements - only custom bar lengths
'2. Is PAA the right code for pair?
'3. Amalgamate duplicates
