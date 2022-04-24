
##


### Changing the font

```vb
Cells.Select
selection.Font.Name = "Times New Roman"

```


### Manipulate sheets
Check `./CH02_AdvancedVBA/exampleFiles/MacroBasics_to_Advanced.xlsm`

```vb
Sub ChnageSheetName()
'
' Macro1 Macro
'

'

   'ActiveSheet.Name = "Third Quarter"


End Sub
Sub CreateSheet()
'
' CreateSheet Macro
'

'
Dim sheetNum As Integer

For sheetNum = 1 To 10
    Sheets.Add After:=ActiveSheet
    Sheets(sheetNum + 1).Select
    Cells(1, 1).Value = "Sheet Number:  " & sheetNum
    ActiveSheet.Name = "Quarter " & sheetNum

Next


For sheetNum = 1 To Worksheets.Count

Sheets(sheetNum).Select
    Cells(1, 5).Value = "Sheet Number:  " & sheetNum
    Next
Sheets(1).Activate
End Sub
Sub CleanMySheets()
'
' Macro12 Macro
'
'
Dim sheetNum As Integer

For sheetNum = 2 To Worksheets.Count
    'Sheets("Quarter 1").Select
        If sheetNum = Worksheets.Count Then
    Exit For 'Continue For
    End If
    'Sheets(sheetNum).Select
   Sheets(sheetNum).Delete
    'ActiveWindow.SelectedSheets.Delete

Next


```
### Object Browser window
This will allow us to see all the classes, methods and other attributes for the
Excel macro and VBA. It is the library of VBA for collections and properties
and methods that we have in Excel.

### USING IMMEDIATE WINDOW
This window will allow us to run some queries to check on some values of some
functions, Mainly it is used for testing and debugging.

```vb
?Worksheets.Count
?Range("B3").Value
'This will center your value in the selected cell'
Selection.HorizontalAlignment = xlCenterAcrossSelection
'Also you can center to the left'
Selection.HorizontalAlignment = xlLeft

```



### [Important] If you use filter
don't use `Use relative References` and you need to

### Using Excel function with Macro
I use usually
```vb
Application.WorksheetFunction.VLookup(prodNum, Range("A1:B51"), 2, FALSE)
```
- [vba with vlookup](https://spreadsheeto.com/vba-vlookup/)

### Advanced Example Macro and Function with fields (params)
#### Keywords: function, method, params, parameters
This example shows that we can also use `function`:`methods` with parameters (fields) that we can specify as shown in the `Macro1`

```vb
   Sub NumToText(ByRef sRng As String, Optional ByVal WS As Worksheet)
    '---Converting visible range form Numbers to Text
        Dim Temp As Double
        Dim vRng As Range
        Dim Cel As Object

        If WS Is Nothing Then Set WS = ActiveSheet
            Set vRng = WS.Range(sRng).SpecialCells(xlCellTypeVisible)
            For Each Cel In vRng
                If Not IsEmpty(Cel.Value) And IsNumeric(Cel.Value) Then
                    Temp = Cel.Value
                    Cel.ClearContents
                    Cel.NumberFormat = "@"
                    Cel.Value = CStr(Temp)
                End If
            Next Cel
    End Sub


    Sub Macro1()
        Call NumToText("A2:A100", ActiveSheet)
    End Sub
```
