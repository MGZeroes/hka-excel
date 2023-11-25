Attribute VB_Name = "Makros"
Sub HelloWorld()
'
' Hello World Makro
' Prints Hallo on the Active Cell
'
    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "Hallo"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Columns("A:A").EntireColumn.EntireColumn.AutoFit
End Sub

Sub PictureMover()
'
' Picture Mover
' Shows a form with a movable picture
'
    PictureMoverForm.Show
End Sub

Sub InsertAndSelect()
'
' InsertAndSelect Makro
' Shows a form where you can select cells
' and insert values
'
    SelectionForm.Show
End Sub
