# How to handle the VB elements

- Our Example in the attached Excelsheet Path: "./loops_Fundamentals.xlsm"

    ```vb
        Sub Ghasak_M01()
        Dim i As Integer
        Dim j As Integer
        For k = 1 To 2
        counter = 0
        For i = 1 To 4
        For j = 1 To 4
            counter = counter + 1
                Worksheets("T_2").Cells(i + k * 12, j).Value = counter


        Next j
        Next i
        Next k
          ActiveSheet.Cells(1, 1).Select
        End Sub

        Sub Button2_Click()
                Columns("A:D").Select
                Selection.ClearContents
                Range("A1").Select

        End Sub
        Sub clear()

        If (Worksheets("T_2").CheckBoxes("Check Box 9").Value = xlOn) Then
                Columns("A:D").Select
                Selection.ClearContents
                Range("A1").Select

                'Worksheets("T_2").CheckBoxes("Check Box 9").Value = xlOff
                 Worksheets("T_2").CheckBoxes("Check Box 9").Value = False
                 End If

        End Sub

        Sub Macro5()
        '
        ' Macro5 Macro
        '
        ' Worksheets("T_2").CheckBoxes("Check Box 9").Value = xlOff
        If Worksheets("T_2").CheckBoxes("Check Box 9").Value = True Then
        Worksheets("T_2").CheckBoxes("Check Box 9").Value = False
        End If

        '
        End Sub
    ```

# Notes
- It seems that `VisualBasic` with `Excel` shows that these names are the one to refer to the name of the element
    - `CheckBoxes`
    - `Buttons`
    - `OptionBoxes`
- You can check the name of the element and how the VB refer to its name by
  reading the name property in the `top left corner in Excel`.
- The `False` Statment in Excel can be carried on using the `xlOff` -> `True`.
  ```vb
    Sub Unchecker()
        Dim chkBox As Excel.CheckBox
        Application.ScreenUpdating = False
        For Each chkBox In ActiveSheet.CheckBoxes
            chkBox.Value = xlOff
        Next chkBox
        Application.ScreenUpdating = True
    End Sub
  ```
## REFERENCES
- [Uncheck box in VBA](https://stackoverflow.com/questions/69252773/uncheck-box-in-vba)
