# Loop with VBA
- Selection
    ```vb
    ActiveSheet.Cells(5, 4).Select
    -or-
    ActiveSheet.Range("D5").Select
    ```
- Looping start always with the Cell(1,1), as the indices start counting from 1,2,3,4...N
- You can examine your code in the `immediate screen` for example
    ```vb
        Selection.HorizontalAlignment = xlCenter
        Cells(3,4).value = "yes"
    ```
- Looping in VB and assign a value to the cell
    ```vb
    Sub Ghasak_M01()
    '
    ' Ghasak_M01 Macro
    '
    ' Keyboard Shortcut: Ctrl+r
    ''

    Dim i As Integer
    Dim j As Integer
    Dim counter As Integer
    Dim shift_rows As Integer


    counter = 0
    For k = 0 To 3
    For i = 1 To 10
        For j = 1 To 10
        counter = counter + 1
        Worksheets("T_1").Cells(i + k * 12, j).Value = counter
        ActiveSheet.Cells(i + k * 12, j).Select
        Selection.HorizontalAlignment = xlLeft

    Next j
    Next i
    Next k


    End Sub
    ```
- Declare an array and store the values, We here using the `ReDim`.

    ```vb
    Sub Ghasak_M01()
    '
    ' Ghasak_M01 Macro
    '
    ' Keyboard Shortcut: Ctrl+r
    ''

    Dim i As Integer
    Dim j As Integer
    Dim counter As Integer
    Dim shift_rows As Integer
    ReDim M(0 To 10, 0 To 10)



    For k = 0 To 3
    counter = 0
    For i = 1 To 10
        For j = 1 To 10
        counter = counter + 1
        Worksheets("T_1").Cells(i + k * 12, j).Value = counter
        ActiveSheet.Cells(i + k * 12, j).Select
        Selection.HorizontalAlignment = xlLeft
        M(i, j) = counter * 10

    Next j
    Next i
    Next k

    For i = 1 To 10
    For j = 1 To 10
        Cells(i, j + 12).Value = M(i, j)
    Next j
    Next i

    End Sub
    ```
- Clear Content

    ```vb
    Sub clear()
    '
    ' clear Macro
    '
    ' You can select Columns,Rows or Range

        Columns("A:D").Select
        Selection.ClearContents
        Range("A1").Select
    End Sub
    ```


## REFERENCES
- [How to select cells/ranges by using Visual Basic procedures in Excel](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/select-cells-rangs-with-visual-basic)
