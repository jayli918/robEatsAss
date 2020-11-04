Attribute VB_Name = "Module1"
Option Explicit


Sub MixMatchColumns(ByRef DataRange As Range, _
                    ByRef ResultRange As Range, _
                    Optional ByVal DataHasHeaders As Boolean = False, _
                    Optional ByVal HeadersInResult As Boolean = False)

Dim rngData As Range
Dim rngResults As Range
Dim lngCount As Long
Dim lngCol As Long
Dim lngNumberRows As Long
Dim ItemCount() As Long
Dim RepeatCount() As Long
Dim PatternCount() As Long
'Long Variables for the Variour For Loops
Dim lngForRow As Long
Dim lngForPattern As Long
Dim lngForItem As Long
Dim lngForRept As Long
'Temporary Arrays used to store the Data and Results
Dim DataArray() As Variant
Dim ResultArray() As Variant

'If the Data range has headers, adjust the
'Range to contain only data
Set rngData = DataRange
If DataHasHeaders Then
    Set rngData = rngData.Offset(1).Resize(rngData.Rows.Count - 1)
End If

'Initialize the Data Array
DataArray = rngData.Value
'Get the number of Columns
lngCol = rngData.Columns.Count

'Initialize the Arrays
ReDim ItemCount(1 To lngCol)
ReDim RepeatCount(1 To lngCol)
ReDim PatternCount(1 To lngCol)

'Get the number of items in each column
For lngCount = 1 To lngCol
    ItemCount(lngCount) = _
        Application.WorksheetFunction.CountA(rngData.Columns(lngCount))
    If ItemCount(lngCount) = 0 Then
        MsgBox "Column " & lngCount & " does not have any items in it."
        Exit Sub
    End If
Next

'Calculate the number of Permutations
lngNumberRows = Application.Product(ItemCount)
'Initialize the Results array
ReDim ResultArray(1 To lngNumberRows, 1 To lngCol)

'Get the number of times each of the items repeate
RepeatCount(lngCol) = 1
For lngCount = (lngCol - 1) To 1 Step -1
    RepeatCount(lngCount) = ItemCount(lngCount + 1) * _
                                RepeatCount(lngCount + 1)
Next lngCount

'Get howmany times the pattern repeates
For lngCount = 1 To lngCol
    PatternCount(lngCount) = lngNumberRows / _
            (ItemCount(lngCount) * RepeatCount(lngCount))
Next

'The Loop begins here, Goes through each column
For lngCount = 1 To lngCol
'Reset the row number for each column iteration
lngForRow = 1
    'Start the Pattern
    For lngForPattern = 1 To PatternCount(lngCount)
        'Loop through each item
        For lngForItem = 1 To ItemCount(lngCount)
            'Repeate the item
            For lngForRept = 1 To RepeatCount(lngCount)
                'Store the value in the array
                ResultArray(lngForRow, lngCount) = _
                        DataArray(lngForItem, lngCount)
                'Increment the Row number
                lngForRow = lngForRow + 1
            Next lngForRept
        Next lngForItem
    Next lngForPattern
Next lngCount

'Output the results
Set rngResults = ResultRange(1, 1).Resize(lngNumberRows, lngCol)
'If the user wants headers in the results
If DataHasHeaders And HeadersInResult Then
    rngResults.Rows(1).Value = DataRange.Rows(1).Value
    Set rngResults = rngResults.Offset(1)
End If
rngResults.Value = ResultArray()

End Sub
                        

Sub CoverMacro()

Dim rngData As Range
Dim rngResults As Range
Dim booDataHeader As Boolean
Dim booResultHeader As Boolean
Dim lngAns As Long
Dim strMessage As String
Dim strTitle As String

strTitle = "Mix 'n Match"

strMessage = "Select the Range that has the Lists:" _
    & vbNewLine & "Make sure there are no blank cells in between."

On Error Resume Next
Set rngData = Application.InputBox(strMessage, strTitle, , , , , , 8)
If Not Err.Number = 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
End If
    

strMessage = "Does the Data have headers in it?"
lngAns = MsgBox(strMessage, vbYesNo, strTitle)
If Not Err.Number = 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
End If

If lngAns = vbYes Then
    booDataHeader = True
Else
    booDataHeader = False
End If

strMessage = "Select the cell where you'd like the results to be pasted"
Set rngResults = Application.InputBox(strMessage, strTitle, , , , , , 8)
If Not Err.Number = 0 Then
    Err.Clear
    On Error GoTo 0
    Exit Sub
End If

If booDataHeader Then
    strMessage = "Do you want headers in your Result?"
    lngAns = MsgBox(strMessage, vbYesNo, strTitle)
    
    If Not Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    If lngAns = vbYes Then
        booResultHeader = True
    Else
        booResultHeader = False
    End If
Else
    booResultHeader = False
End If

Call MixMatchColumns(rngData, rngResults, booDataHeader, booResultHeader)
End Sub



