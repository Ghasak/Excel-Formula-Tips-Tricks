# VBA useful Macro

#### Active a cell
[Selection for macro](https://excelchamps.com/vba/active-cell/)

```vb
Sub vba_activecell()
	'select and entire range
	Range("A1:A10").Select
	'select the cell A3 from the selected range
	Range("A3").Activate
	'clears everything from the active cell
	ActiveCell.Clear
End Sub
```

#### Go to specific cell
```vb
'Old way '
Sheets("Sheet2").Application.Goto Reference:="R1C1"
'New way '
Sheets("Shee2").Select
    Cells(1, 1).Select
```
#### Compare values with different colors
Problem want to solve:
- Highlight a column with `green` for positive value and `red` for negative value as shown below
	- [x] Compare values with different colors
	- First define the variable for the loop and for storing the value of the cell
		- datatype fr integer for the row, and `double` for the decimal value.
- Now you select the top of the row
```vb
Dim row As Integer
Dim val As Double
' Move the cursor to the cell 1,4 to start from there
Cells(1, 4).Select
For i = 1 To 30
	'val = Cells(i, 5).Value
	val = Cells(i, 4).Value
	Cells(i, 5) = val + 1
	' Selecting the cell thatwill be looped
	Cells(i, 4).Select
	If val >= 0 Then
		With Selection.Interior
			.Color = 65535
			Cells(i, 5) = val + 1
			Cells(i, 6) = "Positive"
		End With

	ElseIf val < 0 Then
		With Selection.Interior
			.Color = 255 '5296274
			Cells(i, 6) = "Negative"
		End With
	End If
Next i
' Move back your selection to the cell 1,4
Cells(1, 4).Select
```

For clearing the values and start the exercise again
`
```vb
ActiveCell.Offset(0, 1).Range("A1:B30").Select
Selection.ClearContents
ActiveCell.Offset(0, -1).Range("A1:A30").Select
With Selection.Interior
	.Pattern = xlNone
	.TintAndShade = 0
	.PatternTintAndShade = 0
End With
ActiveCell.Offset(6, 4).Range("A1").Select
```

#### Using
We can use either `Cells(i,j)` to refer to a cell and replace the value.
But, also we can use `ActiveCell.FormulaR1C1` also to refer to the value of the cell.
![[Screen Shot 2022-04-23 at 1.21.42.png]]


#### Casting int to str and replace value of a cell with For
First function I created:
- Using for loop
- Cast int to str using `CStr(CDbl())`
- with looping over every cell, change the value using `Cells`
```vb

Sub G01_Testing_formatting()
'
' G01_Testing_formatting Macro
'
' Keyboard Shortcut: Ctrl+r
'
    ActiveCell.Range("A1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone

    Dim i, mytext

    For i = 1 To 10  ' Set up 10 repetitions.
    mytext = "wow_" & CStr(CDbl(i))
    Cells(i, 2) = mytext


    Next i

End Sub

```
![[Screen Shot 2022-04-23 at 1.11.28.png]]o
