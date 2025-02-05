Attribute VB_Name = "dataChecker"
Option Base 1

Function Col_Letter(lngCol As Variant) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
   
   Function RangeToArray(ws As Variant, startrow As Variant, endrow As Variant, col As Variant) As Variant
    Dim dataArray() As Variant
    Dim i As Long
    Dim rowIndex As Long
     
    ' Resize the array to hold the correct number of elements
    ReDim dataArray(1 To endrow - startrow + 1)
    
    ' Loop through each row in the specified range
    rowIndex = 1
    For i = startrow To endrow
        ' Assign each cell's value to the array
    On Error GoTo badVal
        dataArray(rowIndex) = ws.Cells(i, col).value
        rowIndex = rowIndex + 1
    Next i
    
    ' Return the filled array
    RangeToArray = dataArray
    Exit Function
badVal:
   ' colRef = Range.Cells(1, col)
    MsgBox "There is a bad value at cell row " & i & ", column " & Col_Letter(col) & vbCrLf & "Remove this value before continuing."
    Err.Raise Number:=vbObjectError + 513, Source:="RangeToArray"

End Function


Public Sub CheckData()
' Set the columns locations for the data dictionary sheet
Const vNameCol As Integer = 1
Const vNewNameCol As Integer = 2
Const Label_For_ReportCol As Integer = 3
Const TypeCol As Integer = 4
Const ValueCol As Integer = 5
Const Value_LabelCol As Integer = 6
Const MinimumCol As Integer = 7
Const MaximumCol As Integer = 8
Const MissingCol As Integer = 9
Const UnreadableCol As Integer = 10
Const DataCol As Integer = 11
Const ImportCol As Integer = 12
Dim n_unique_expected As Integer

'prompt the user to save a duplicate sheet if they don't want the formatting removed
  response = MsgBox("This macro will remove all formatting from the current sheet." & vbCrLf & "  Is it okay to remove all formatting from this worksheet?" & vbCrLf & _
                      "If not, please copy the sheet to a new one before running the CheckData macro.", _
                      vbYesNo + vbQuestion, "Remove Formatting")
                      
   If response = vbNo Then Exit Sub
                      
' prompt the user for the maximum number of codes expected
n_unique_expected = Application.InputBox("Please entered the maximum number of unique codes expected for a variable, or click OK to use the default:", Default:=10, Title:="PMH Datachecker", Type:=1)

If n_unique_expected = vbCancel Then Exit Sub

 ' Turn off screen updating
    Application.ScreenUpdating = False
    
 ' remove all formatting
 RemoveFormatting
 
    On Error GoTo errExit
    
    
 ' Store the current workbook /sheet
 Set thisBook = Application.ActiveWorkbook
    Set thissheet = Application.ActiveSheet
    
' Get the sheet range
PickedActualUsedRange
nrow = Selection.Rows.count
ncol = Selection.columns.count

' Find location of first number/date to check for multiple header rows
For r = 1 To nrow
    For c = 1 To ncol
            If IsEmpty(Cells(r, c).value) = False And (IsDate(Cells(r, c).value) = True Or IsNumeric(Cells(r, c).value) = True) Then GoTo FirstValue
    Next c
Next r

FirstValue:
firstdatarow = r
If firstdatarow > 2 Then
    With Cells(1, 1)
        .AddComment
        .Comment.Visible = True
        .Comment.text text:="Data Checks:" & Chr(10) & "Only one header row is permitted, with no special characters"
    End With
End If

' Check for special characters in headers
specialComment = False
For r = 1 To (firstdatarow - 1)
    For c = 1 To ncol
        If Cells(r, c).value Like "*[!A-Za-z0-9 _]*" Then
            Cells(r, c).Font.Color = -16776961
            If specialComment = False Then
                With Cells(r, c)
                .ClearComments
                    .AddComment
                    .Comment.Visible = True
                    .Comment.text text:="Data Checks:" & Chr(10) & "Please remove all spaces and special characters in column headers"
                End With
            specialComment = True
            End If
        End If
        If Cells(r, c).value Like "*[=]*" Then
            Cells(r, c).Font.Color = -16776961
                With Cells(r, c)
                    .ClearComments
                    .AddComment
                    .Comment.Visible = True
                    .Comment.text text:="Data Checks:" & Chr(10) & "All codes must be supplied in a separate data dictionary"
                End With

        End If
    Next c
Next r
        
' Check for empty rows, columns
Dim emptyFirstDataRow As Boolean
For r = 1 To nrow
    Set fullRow = Cells(r, 1).EntireRow
    If Application.WorksheetFunction.CountA(fullRow) = 0 Then
    If r = 2 Then emptyFirstDataRow = True
        With fullRow.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells(r, 1).AddComment
'        Cells(r, 1).Comment.Visible = True
'        Cells(r, 1).Comment.Text Text:="Data Checks:" & Chr(10) & "Please remove empy rows"
    End If
Next r


    ' Create a new sheet for the dictionary
    Dim dictCurrentRow As Single
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.count))
    
'    newSheet.Name = "DataCheckSheet"
    Dim varType As String
' Set up the column  headers
    newSheet.Cells(1, vNameCol).value = "Current_Variable_Name"
    newSheet.Cells(1, vNewNameCol).value = "Suggested_Name"
    newSheet.Cells(1, Label_For_ReportCol).value = "Label_For_Report"
    newSheet.Cells(1, TypeCol).value = "Type"
    newSheet.Cells(1, ValueCol).value = "Value"
    newSheet.Cells(1, Value_LabelCol).value = "Value_Label"
    newSheet.Cells(1, MinimumCol).value = "Minimum"
    newSheet.Cells(1, MaximumCol).value = "Maximum"
    newSheet.Cells(1, MissingCol).value = "Missing"
    newSheet.Cells(1, UnreadableCol).value = "Unreadable"
    newSheet.Cells(1, DataCol).value = "Column_Number"
    newSheet.Cells(1, ImportCol).value = "Import"
    dictCurrentRow = 2
    
    ' go back to the data sheet
    thissheet.Activate

' Highlight empty columns
For c = 1 To ncol
    Set fullCol = thissheet.Cells(1, c).EntireColumn
    If Application.WorksheetFunction.CountA(fullCol) = 0 Then
        With fullCol.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Cells(1, c).AddComment
'        Cells(1, c).Comment.Visible = True
 '       Cells(1, c).Comment.Text Text:="Data Checks:" & Chr(10) & "Please remove empy columns"
    End If
Next c

Dim vName As String
Dim vType As String
Dim nMissing As Single
Dim nUnreadable As Single
Dim uniqueColVals As Variant
Dim txtCols() As String
Dim txtCount As Integer
txtCount = 0
For c = 1 To ncol
' Identify unique columns values (Mac workaround)
    Set DictOne = New Dictionary
'    columnValues = Range(Cells(firstdatarow, c), Cells(nrow, c))
    
    columnValues = RangeToArray(thissheet, firstdatarow, nrow, c)
    
    On Error Resume Next 'if the key is already present
    For Each x In columnValues
        DictOne.Add CStr(x), CStr(x)
    Next x

    uniqueColVals = DictOne.Keys
    len_unique = UBound(uniqueColVals) + 1
    
    Set DictTwo = New Dictionary
    On Error Resume Next 'if the key is already present
    For Each x In columnValues
        DictTwo.Add LCase(CStr(x)), LCase(CStr(x))
    Next x
    On Error GoTo errExit
    uniqueLowerCase = DictTwo.Keys
    
    ' add warning if these don't match
    If len_unique > (UBound(uniqueLowerCase) + 1) Then
        If len_unique <= n_unique_expected Then
            msgText = "Data Checks:" & Chr(10) & "Case must be consistent, check values" & Chr(10) & Join(uniqueColVals, vbLf)
        Else
            msgText = "Data Checks:" & Chr(10) & "This column appears to have inconsistent formatting"
        End If
        Cells(firstdatarow, c).AddComment
        Cells(firstdatarow, c).Comment.Visible = True
        Cells(firstdatarow, c).Comment.text text:=msgText
    End If
    
    'Get the variable name
    vName = Cells(1, c).value
    
    ' Determine the data type
    vType = ""
    dateComment = False
    stringComment = False
    nMissing = 0
    nUnreadable = 0
   
    ' go to the first cell with data in the column
    r = firstdatarow
    Do Until IsEmpty(Cells(r, c).value) = False Or r = nrow
    r = r + 1
    Loop

    If IsDate(Cells(r, c).value) Then
        vType = "Date"
    ' Check date formatting
    ' Look in the second column to find date values
    ' If a date is found start looking down the column to check for consistency
        firstFormat = Cells(r, c).NumberFormat
        ' need to strip special characters excep '-' out of this and compare to new with special stripped
        firstFormat = Replace(firstFormat, "[$-409]", "")
        firstFormat = Replace(firstFormat, ";@", "")
        For rNext = (r + 1) To nrow
            If IsEmpty(Cells(rNext, c).value) = True Then nMissing = nMissing + 1
            If IsEmpty(Cells(rNext, c).value) = False Then
                nextCellFormat = Replace(Cells(rNext, c).NumberFormat, "[$-409]", "")
                nextCellFormat = Replace(nextCellFormat, ";@", "")
                If nextCellFormat <> firstFormat Then
                    nUnreadable = nUnreadable + 1
                    Cells(rNext, c).Font.Color = -16776961
                    Cells(rNext, c).Font.TintAndShade = 0
                    If dateComment = False Then
                        Cells(rNext, c).AddComment
                        Cells(rNext, c).Comment.Visible = True
                        Cells(rNext, c).Comment.text text:="Data Checks:" & Chr(10) & "Inconsistent date format"
                    End If
                    dateComment = True
                End If
            End If
         Next rNext
    End If

    ' Check Numeric Columns for Text -if mostly text then set to text
    If IsNumeric(Cells(r, c).value) Then
        Dim txtCells() As Integer
        Dim t As Integer
        t = 1
        For rNext = (r + 1) To nrow
            If IsEmpty(Cells(rNext, c).value) = True Then nMissing = nMissing + 1
            If IsEmpty(Cells(rNext, c).value) = False And IsNumeric(Cells(rNext, c).value) = False Then
                ReDim Preserve txtCells(t)
                txtCells(t) = rNext
                t = t + 1
            End If
        Next rNext
        If t = 1 Then
            ' All numeric
            If len_unique <= n_unique_expected Then vType = "Codes" Else vType = "Numeric"
            
        ElseIf len_unique <= n_unique_expected Then
            vType = "Categorical"
        ' If there are a lot of text cells (more than half) then assume this is a text column
        ' otherwise, if the first cell is numeric assume it is suppose to be numeric and that the data are messy
        ElseIf (t - 1) > (nrow / 2) Then
            vType = "Text"
        Else
            nUnreadable = t - 1
            vType = "Numeric"
            For t = LBound(txtCells) To UBound(txtCells)
                rw = txtCells(t)
                Cells(rw, c).Font.Color = -16776961
                Cells(rw, c).Font.TintAndShade = 0
                If stringComment = False Then
                    Cells(rw, c).AddComment
                    Cells(rw, c).Comment.Visible = True
                    Cells(rw, c).Comment.text text:="Data Checks:" & Chr(10) & "Text in numeric column"
                End If
                stringComment = True
            Next t
        End If
    End If
    
    ' Highlight the column if there are any date or string comments
    If (dateComment Or stringComment) Then
        Set fullCol = Cells(1, c).EntireColumn
        fullCol.Borders.ColorIndex = 3
    End If

    If vType = "" And len_unique <= n_unique_expected Then
        vType = "Categorical"
    ElseIf vType = "" Then vType = "Text"
    End If
    
    ' Add information to data summary sheet
'    If vType <> "Text" Then ' skip empty columns & text columns
     newSheet.Cells(dictCurrentRow, vNameCol).value = vName
     newSheet.Cells(dictCurrentRow, TypeCol).value = vType
    newSheet.Cells(dictCurrentRow, DataCol).value = c
    newSheet.Cells(dictCurrentRow, ImportCol).value = "TRUE"
     
 
        ' Calculate Min/Max for numeric columns
        If vType = "Numeric" Then
         ' Convert to numeric values
    For i = LBound(uniqueColVals) To UBound(uniqueColVals)
        If IsNumeric(uniqueColVals(i)) Then
            uniqueColVals(i) = CDbl(uniqueColVals(i))
        End If
    Next i
            minVal = Empty
            maxVal = Empty
                For x = LBound(uniqueColVals) To UBound(uniqueColVals)
                If uniqueColVals(x) <> "" And IsNumeric(uniqueColVals(x)) Then
                    If IsEmpty(minVal) Then
                        minVal = uniqueColVals(x)
                        maxVal = uniqueColVals(x)
                    End If
                    If uniqueColVals(x) < minVal Then minVal = uniqueColVals(x)
                    If uniqueColVals(x) > maxVal Then maxVal = uniqueColVals(x)
                End If
            Next x
            newSheet.Cells(dictCurrentRow, MinimumCol).value = minVal
            newSheet.Cells(dictCurrentRow, MaximumCol).value = maxVal
            newSheet.Cells(dictCurrentRow, MinimumCol).NumberFormat = "General"
            newSheet.Cells(dictCurrentRow, MaximumCol).NumberFormat = "General"
            newSheet.Cells(dictCurrentRow, MissingCol).value = nMissing
            newSheet.Cells(dictCurrentRow, UnreadableCol).value = nUnreadable
    
        End If
        ' ... or first / last for date columns
        If vType = "Date" Then
            minDate = Empty
            maxDate = Empty
            For x = LBound(uniqueColVals) To UBound(uniqueColVals)
                If uniqueColVals(x) <> "" And IsDate(uniqueColVals(x)) Then
                    If IsEmpty(minDate) Then
                        minDate = CDate(uniqueColVals(x))
                        maxDate = CDate(uniqueColVals(x))
                    End If
                    If CDate(uniqueColVals(x)) < minDate Then minDate = CDate(uniqueColVals(x))
                    If CDate(uniqueColVals(x)) > maxDate Then maxDate = CDate(uniqueColVals(x))
                End If
            Next x
            newSheet.Cells(dictCurrentRow, MinimumCol).value = Format(minDate, "d mmm yyyy")
            newSheet.Cells(dictCurrentRow, MaximumCol).value = Format(maxDate, "d mmm yyyy")
            newSheet.Cells(dictCurrentRow, MissingCol).value = nMissing
            newSheet.Cells(dictCurrentRow, UnreadableCol).value = nUnreadable
        End If
        
        ' Other wise for categorical data show each value
        If vType = "Categorical" Or vType = "Codes" Then
            startrow = dictCurrentRow
            For x = LBound(uniqueColVals) To UBound(uniqueColVals)
                If uniqueColVals(x) <> "" Then
                    newSheet.Cells(dictCurrentRow, ValueCol).value = uniqueColVals(x)
                    dictCurrentRow = dictCurrentRow + 1
                End If
            Next x
            dictCurrentRow = dictCurrentRow - 1
            endrow = dictCurrentRow
            ' Sort these - VBA DOESN'T HAVE INTERNAL SORT FUNCTION- FOR REAL????
            If (startrow <> endrow) Then
                Set Rng = newSheet.Range(newSheet.Cells(startrow, ValueCol), newSheet.Cells(endrow, ValueCol))
                With newSheet.Sort
                    .SortFields.Clear
                    .SortFields.Add Key:=Rng, Order:=xlAscending
                    .SetRange Rng
                    .Header = xlNo
                    .Apply
                End With
            End If
       End If
    

    dictCurrentRow = dictCurrentRow + 1
'   Else
'    txtCount = txtCount + 1
'    ReDim Preserve txtCols(txtCount)
'    txtCols(txtCount) = vName
'   End If
Next c
Dim lastVarRow As Integer
lastVarRow = dictCurrentRow - 1
'' List the txt columns that won't be imported
'If txtCount > 0 Then
'   dictCurrentRow = dictCurrentRow + 1
'    newSheet.Cells(dictCurrentRow, 1).value = "The following variables are mostly text and should not be imported:"
'    For t = LBound(txtCols) To UBound(txtCols)
'        dictCurrentRow = dictCurrentRow + 1
'        newSheet.Cells(dictCurrentRow, 1).value = txtCols(t)
'    Next t
'End If

' Format data dictionary
'    newSheet.Activate
Dim mac_msg_displayed As Boolean: mac_msg_displayed = False

'newSheet.Range("A2").Activate

Dim openBracketPos As Single
Dim colonPos As Single
Dim endPos As Single
Dim varNames As Collection
Set varNames = New Collection
    
For rw = 2 To lastVarRow
    varName = newSheet.Range("A1").Offset(rw - 1, 0).value
    If Not IsEmpty(varName) Then
        openBracketPos = InStr(varName, "(")
        colonPos = InStr(varName, ":")
    If openBracketPos > 0 And colonPos > 0 Then
        If openBracketPos > colonPos Then endPos = colonPos
        If colonPos > openBracketPos Then endPos = openBracketPos
    ElseIf openBracketPos > 0 Then
        endPos = openBracketPos
    Else
        endPos = colonPos
    End If
    On Error GoTo MacError
    newname = ReplaceSpecial(varName)
    On Error GoTo errExit

    If endPos > 0 Then newname = Left(newname, endPos - 1)
'        ' Check that the suggested variable name doesn't already exist
'     If varNames.Count > 0 Then
'     ' check if the key already exists
'
'        If Contains(varNames, CVar(newname)) Then
'
'        nm_ind = 2
'        Do ' increment new_name name until unique
'            ' creat new variable name
'            newname = newname & "_" & CStr(nm_ind)
'            If Contains(varNames, CVar(newname)) Then nm_ind = nm_ind + 1 Else Exit Do
'        Loop
'        End If
'      On Error Resume Next
'      varNames.Add newname, CStr(newname) ' this adds the name as the key
'     End If
'     If varNames.Count = 0 Then
'      On Error Resume Next
'      varNames.Add newname, CStr(newname)
'     End If

    
    If varName <> newname Then duplColFlag = False
    newSheet.Range("A1").Offset(rw - 1, 1).value = newname
    newname = Replace(newname, "_", " ")
    newname = Trim(newname)
    newSheet.Range("A1").Offset(rw - 1, 2).value = newname
    End If
Next rw

' Make all suggested names and labels unique
suggested_varnames = RangeToArray(newSheet, 2, lastVarRow, 2)
unique_varnames = AddSuffixToDuplicates(suggested_varnames, sep = "_")
suggested_varlbls = RangeToArray(newSheet, 2, lastVarRow, 3)
unique_varlbls = AddSuffixToDuplicates(suggested_varlbls, sep = Chr(32)) ' doesn't work, no space

' update the dictionary sheet
For rw = 2 To lastVarRow
    newSheet.Range("A1").Offset(rw - 1, 1).value = unique_varnames(rw - 1)
    newSheet.Range("A1").Offset(rw - 1, 2).value = unique_varlbls(rw - 1)
Next rw

' Autofit
    newSheet.Select
    newSheet.Range(Cells(1, 1), Cells(lastVarRow, UnreadableCol)).Select
    Selection.columns.AutoFit
    newSheet.Range(Cells(1, 1), Cells(lastVarRow, UnreadableCol)).columns.AutoFit
    
' Bold the first row and left-align codes and category values
    newSheet.columns("D:D").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.Font.Bold = True
    
' select first cell
'newSheet.Range("A1").Select
    
' removed suggested names if the same as varnames
If duplColFlag = True Then
    columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
End If

' Autosize Comments (on windows)
On Error Resume Next
For Each newComment In Application.ActiveSheet.Comments
    newComment.Shape.TextFrame.AutoSize = True
Next

' Move to just under the top row
commentTop = Application.ActiveSheet.Cells(2, 1).Top
    ' Loop through each shape on the worksheet
    For Each shp In Application.ActiveSheet.Shapes
        shp.Left = Application.ActiveSheet.Cells(shp.TopLeftCell.row, shp.TopLeftCell.Column).Right + 10
        shp.Top = commentTop
    Next shp

On Error GoTo 0

OnExit:
    ' Turn on screen updating & enable events
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Data Check Complete." & vbCrLf & "Data Dictionary stored on last sheet", , "PMH Datachecker"
    ' move back to the data sheet and select the first cell
    thissheet.Activate
    thissheet.Range("A1").Select
    Exit Sub

errExit:
    RemoveFormatting
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
MacError:
    If mac_msg_displayed = False Then
        mac_msg_displayed = True
        MsgBox "Not all functionality can be implemented on a mac. No suggested variable names provided", , "PMH Datachecker"
    End If
End Sub

Sub ProcessColumns()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim columnCount As Integer
    Dim uniqueValues As Collection
    Dim cellValue As Variant
    Dim uniqueCount As Integer
    Dim minValue As Variant
    Dim maxValue As Variant
    Dim isNumericOrDate As Boolean
    
    ' Set the active sheet
    Set ws = ActiveSheet
    
    ' Create a new sheet for the results
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.count))
    newSheet.Name = "UniqueValuesSheet"
    
    ' Loop through each column in the active sheet
    For columnCount = 1 To ws.columns.count
        ' Reset variables for each column
        Set uniqueValues = New Collection
        uniqueCount = 0
        minValue = Empty
        maxValue = Empty
        isNumericOrDate = False
        
        ' Check each cell in the column
        For Each cell In ws.columns(columnCount).Cells
            cellValue = cell.value
            
            ' Check if the cell is not empty
            If Not IsEmpty(cellValue) Then
                ' Check if the cell value is not already in the uniqueValues collection
                If Not Contains(uniqueValues, cellValue) Then
                    ' Add the unique value to the collection
                    uniqueValues.Add cellValue
                    uniqueCount = uniqueCount + 1
                End If
                
                ' Check if the value is numeric or a date
                If IsNumeric(cellValue) Or IsDate(cellValue) Then
                    isNumericOrDate = True
                    
                    ' Update min and max values
                    If IsEmpty(minValue) Or cellValue < minValue Then
                        minValue = cellValue
                    End If
                    
                    If IsEmpty(maxValue) Or cellValue > maxValue Then
                        maxValue = cellValue
                    End If
                End If
            End If
        Next cell
        
        ' Check the number of unique values
        If uniqueCount <= n_unique_expected Then
            ' Place unique values in a new row in the UniqueValuesSheet
            newSheet.Cells(1, columnCount).value = "VariableName" & columnCount
            For i = 1 To uniqueCount
                newSheet.Cells(i + 1, columnCount).value = uniqueValues(i)
            Next i
        Else
            ' More than n_unique_expected unique values
            ' Check if the data is numeric or date
            If isNumericOrDate Then
                ' Place min and max values in the second and third columns
                newSheet.Cells(1, columnCount * 2 - 1).value = "VariableName" & columnCount
                newSheet.Cells(2, columnCount * 2 - 1).value = "MinValue"
                newSheet.Cells(2, columnCount * 2).value = "MaxValue"
                newSheet.Cells(3, columnCount * 2 - 1).value = minValue
                newSheet.Cells(3, columnCount * 2).value = maxValue
            Else
                ' Text data
                newSheet.Cells(1, columnCount * 2 - 1).value = "VariableName" & columnCount
                newSheet.Cells(2, columnCount * 2 - 1).value = "Text"
            End If
        End If
    Next columnCount
End Sub

Function Contains(col As Collection, val As Variant) As Boolean
' This needs to check for error, but also test for change in collection count
' VBA seems to have gotten rid of the error when adding duplicate keys and
' simply doesn't change the collection (discovered Jan 2025)
    start_count = col.count
     Err.Clear
     On Error Resume Next
    col.Add val, CStr(val)
    Contains = (Err.Number = 0)
    Err.Clear
    If col.count <> start_count Then Contains = False
End Function

' Took this code from here:
' https://contexturesblog.com/archives/2012/03/01/select-actual-used-range-in-excel-sheet/
Sub PickedActualUsedRange()
  Range("A1").Resize(Cells.Find(What:="*", SearchOrder:=xlRows, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).row, _
      Cells.Find(What:="*", SearchOrder:=xlByColumns, _
      SearchDirection:=xlPrevious, LookIn:=xlValues).Column).Select
End Sub


Sub removeColumnFormatting()
   ActiveCell.EntireColumn.Select
    Selection.ClearComments
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
 Sub RemoveFormatting()
 
  ' Turn off screen updating
    Application.ScreenUpdating = False

    Cells.Select

' Add code to remove formatting from cells so that test can be re-run
    Selection.ClearComments
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone


OnExit:
    ' Turn on screen updating & enable events
    Application.ScreenUpdating = True
    Application.EnableEvents = True
 End Sub

' Not currently used
Sub SortArray(ByRef arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                ' Swap elements
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' Replace all non-alphanumeric characters
Function ReplaceSpecial(ByVal text As String) As String

Dim result As String
Dim allMatches As Object
Dim RE As Object
newtext = text

If IsFirstCharacterNumber(text) Then newtext = "v" & newtext

On Error GoTo macWorkAround

' Code for Windows
Set RE = CreateObject("vbscript.regexp")

RE.Pattern = "[^a-zA-Z0-9]"
RE.Global = True
RE.IgnoreCase = True
Set allMatches = RE.Execute(text)


If allMatches.count <> 0 Then
For i = 0 To allMatches.count
 For j = 0 To allMatches.Item(i).submatches.count - 1
        specialChar = allMatches.Item(i).submatches.Item(j)
        newtext = Replace(newtext, specialChar, "_")
    Next
Next i
End If
ReplaceSpecial = newtext
Exit Function

'Code for Mac:
macWorkAround:
    newtext = Replace(newtext, "!", "_")
    newtext = Replace(newtext, "@", "_")
    newtext = Replace(newtext, "#", "_")
    newtext = Replace(newtext, "$", "_")
    newtext = Replace(newtext, "%", "_")
    newtext = Replace(newtext, "^", "_")
    newtext = Replace(newtext, "&", "_")
    newtext = Replace(newtext, "*", "_")
    newtext = Replace(newtext, "+", "_")
    newtext = Replace(newtext, "=", "_")
    newtext = Replace(newtext, "-", "_")
    newtext = Replace(newtext, "?", "_")
    newtext = Replace(newtext, "/", "_")
    newtext = Replace(newtext, ">", "_")
    newtext = Replace(newtext, "<", "_")
    newtext = Replace(newtext, ";", "_")
    newtext = Replace(newtext, ";", "_")
    newtext = Replace(newtext, "`", "_")
    newtext = Replace(newtext, "'", "_")
    newtext = Replace(newtext, "[", "_")
    newtext = Replace(newtext, "]", "_")
    newtext = Replace(newtext, "|", "_")
    newtext = Replace(newtext, "~", "_")
    newtext = Replace(newtext, "(", "_")
    newtext = Replace(newtext, ")", "_")
    newtext = Replace(newtext, " ", "_")

    ReplaceSpecial = newtext

End Function

Function AddSuffixToDuplicates(arr As Variant, sep As String) As Variant
 Set dict = New Dictionary
    Dim i As Long
    Dim result() As String
    Dim value As String
    Dim count As Long
       
    ' Initialize result array with the same size as input array
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        value = arr(i)
        If value <> "" Then
        ' If the value is already in the dictionary, increment its count and add a suffix
        If dict.Exists(value) Then
            count = dict(value) + 1
            dict(value) = count
            result(i) = value & sep & count
        Else
            ' If it's the first occurrence, store it without a suffix
            dict.Add value, 1
            result(i) = value
        End If
        Else
            result(i) = ""
        End If
    Next i
    
    AddSuffixToDuplicates = result
End Function



Function IsFirstCharacterNumber(ByVal text As String) As Boolean
    Dim firstChar As String
    
    ' Check if the cell is not empty
    If Not IsEmpty(text) Then
        ' Get the first character of the cell value
        firstChar = Left(text, 1)
        
        ' Check if the first character is a number
        If IsNumeric(firstChar) Then
            IsFirstCharacterNumber = True
        Else
            IsFirstCharacterNumber = False
        End If
    Else
        IsFirstCharacterNumber = False
    End If
End Function

