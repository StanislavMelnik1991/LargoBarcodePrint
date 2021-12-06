Sub ClearAll()
On Error Resume Next
CheckChanges = False
Call PageQuantity(4)
Erase Article
Erase Paint

Paint(0, 1, 0) = 1

    Cells(2, 3).FormulaR1C1 = Empty
    Sheets(1).Range("A2:D2,A4:D80,E:H").ClearContents
    Sheets(1).Range("A2:D2,A4:D80,E:H").ClearComments
    Sheets(2).Range("C2:D2,A4:D80,E:H").ClearContents
    Sheets(2).Range("C2:D2,A4:D80,E:H").ClearComments
    Paint(0, 1, 0) = 1
    Range("B2").FormulaR1C1 = "1"
CheckChanges = True

End Sub
Public Sub PasteColorants() 
Sheets(1).Select
Application.ScreenUpdating = False
On Error Resume Next
CheckChanges = False
Erase Article
Erase Paint
    Range("A2").ClearContents
    Range("A:D").ClearComments
    Paint(0, 1, 0) = 1
    Paint(0, 2, 0) = ""
    Cells(2, 2).FormulaR1C1 = "1"
    Cells(1, 5).Select
    ActiveSheet.Paste
Call PasteDefinition
Call PasteCalculation
Call PasteColorantsCalculation
    Sheets(1).Range("E:H").ClearContents
    Cells(2, 1).Select
CheckChanges = True
Application.ScreenUpdating = True
End Sub
Sub AddColorants()
Sheets(2).Select
On Error Resume Next
Application.ScreenUpdating = False
CheckChanges = False
Erase Article
    Sheets(2).Range("A:D").ClearComments
    Paint(0, 1, 0) = 1
    Cells(1, 5).Select
    ActiveSheet.Paste
Call PaintConvert2Article
Call ReDefinition
Call PasteAddDefinition
Call PasteCalculation
Call PasteColorantsCalculation
    Sheets(2).Range("E:H").ClearContents
    Cells(2, 1).Select
CheckChanges = True
Application.ScreenUpdating = True
End Sub
Sub AddPaint()
Sheets(2).Select
Application.ScreenUpdating = False
On Error Resume Next
CheckChanges = False
Erase Article
    Range("A:D").ClearComments
    Paint(0, 1, 0) = 1
Call PaintAddDefinition
Call PaintConvert2Article
Call ReDefinition
Call PasteCalculation
Call PasteColorantsCalculation
    Sheets(1).Range("E:H").ClearContents
    Cells(2, 1).Select
CheckChanges = True
Application.ScreenUpdating = True
End Sub

Sub Print1()
    Cells(2, 1).Select
Application.ScreenUpdating = False
Call BarCodeStart

    If PageNumber = 0 Then
    Cells(2, 1).Select
    Exit Sub
    End If

temp = 4 + PageNumber - 1
temp = "$A$4:$B$" & temp    
ActiveSheet.PageSetup.PrintArea = temp
Call PrintPage(1)
Call ClearBarcode
Cells(2, 1).Select
Application.ScreenUpdating = True
End Sub