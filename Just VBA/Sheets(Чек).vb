Private Sub Worksheet_Activate()
Call ClearAll
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'On Error Resume Next
   Dim KeyCells As Range
If CheckChanges = False Then
Exit Sub
End If
' The variable KeyCells contains the cells that will
    Set KeyCells = Range("A4:D40")
If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
'Target.Row 'номер строки
'Target.Column 'номер столбца
'
'
If Target.Column = 1 Then
'при изменении штрих-кода
CheckChanges = False
Application.ScreenUpdating = False
Target = Article(Target.Row - 4, 0, 0)
Application.ScreenUpdating = True
CheckChanges = True
End If
'
'
'
'
'
If Target.Column = 2 Then
Application.ScreenUpdating = False
'при изменении кол-ва
CheckChanges = False
    If IsNumeric(Target) Then
If Target = Empty Or Target = 0 Then 'delete
Sheets(1).Range("A2:D2,A4:D80,E:H").ClearComments
    If Cells(Target.Row, 3).Value = "Краска" Then
        temp = Cells(Target.Row, 1).Value
        For i = 1 To 8
        If temp = Paint(i, 0, 1) Then
            For i1 = 0 To 4
            Paint(i, i1, 1) = Empty
            Next i1
        Cells(2, 1) = Empty
        End If
        Next i
    End If
Erase Article
Call PaintConvert2Article
Call ReDefinition
Call PasteCalculation
Call PasteColorantsCalculation
CheckChanges = True
Application.ScreenUpdating = True
Exit Sub
End If
    Article(Target.Row - 4, 1, 0) = Target
 If Cells(Target.Row, Target.Column + 1).Value = "Краска" Then
 Article(Target.Row - 4, 1, 0) = Target * 10 ^ k / Paint(0, 1, 0)
 End If
    End If
    
If Cells(Target.Row, Target.Column + 1).Value <> Empty Then
Call PasteColorantsCalculation
Else
Target = Empty
End If
Application.ScreenUpdating = True
CheckChanges = True
End If
'
'
'
'
If Target.Column = 3 Then
'при изменении номера пигмента
CheckChanges = False
Application.ScreenUpdating = False
Find = Application.Match(Format(Target, "##00"), Sheets("красители").Range("A:A"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)
If Find = Empty Then
Target = Empty
Application.ScreenUpdating = True
CheckChanges = True
Exit Sub
End If

Article(Target.Row - 4, 2, 0) = Target
'Call AddCalculation(Target.Row - 4)
'Call lookArrayArticle(0)

If Cells(Target.Row, Target.Column - 1).Value > 0 Then
Erase Article
Call PaintConvert2Article
Call ReDefinition
Call PasteCalculation
Call PasteColorantsCalculation
Else


For i2 = 0 To (Target.Row - 5)
If Abs(Article(Target.Row - 4, 2, 0)) = Abs(Article(i2, 2, 0)) Then
        Article(i2, 1, 0) = InputBox("Введите кол-во пигмента № " & Article(i2, 2, 0), "Пигмент", Article(i2, 1, 0))
        Article(i2, 1, 0) = Article(i2, 1, 0) / 10 ^ k
        Article(Target.Row - 4, 2, 0) = Empty
        Article(i2, 1, 0) = Application.RoundUp(Article(i2, 1, 0) * 10 ^ k, 0)
Call AddCalculation(Target.Row - 4)
Call PasteColorantsCalculation
Application.ScreenUpdating = True
CheckChanges = True
Exit Sub
End If
Next i2


Article(Target.Row - 4, 1, 0) = InputBox("Введите количество для пигмента № " & Format(Article(Target.Row - 4, 2, 0), "##00"), "Краситель", 3)
Article(Target.Row - 4, 1, 0) = Article(Target.Row - 4, 1, 0)
Call AddCalculation(Target.Row - 4)
Call PasteColorantsCalculation
End If
Application.ScreenUpdating = True
CheckChanges = True
End If

If Target.Column = 4 Then
Application.ScreenUpdating = False
'при изменении цены
CheckChanges = False
Target = Application.RoundUp(Article(Target.Row - 4, 3, 0) * Article(Target.Row - 4, 1, 0) * Paint(0, 1, 0) / 1000, 2)
CheckChanges = True
Application.ScreenUpdating = True
End If

   Dim KeyCells2 As Range
If CheckChanges = False Then
Exit Sub
End If
End If




' подписываем столбцы
    Set KeyCells = Range("A1:D3")
If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
If Target.Row = 1 Or Target.Row = 3 Then
CheckChanges = False
Application.ScreenUpdating = False
Sheets(1).Cells(1, 1).FormulaR1C1 = "6 цифр штрихкода"
Sheets(1).Cells(1, 2).FormulaR1C1 = "Кол-во"
Cells(1, 3).FormulaR1C1 = "Стоимость колеровки:"
Cells(3, 1).FormulaR1C1 = "ШТРИХ-КОД"
Cells(3, 2).FormulaR1C1 = "Кол -во"
Cells(3, 3).FormulaR1C1 = "№"
Cells(3, 4).FormulaR1C1 = "Цена"
Application.ScreenUpdating = True
CheckChanges = True
End If

End If

'меняем краску или кол-во краски
    Set KeyCells = Range("A2:B2")
If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then
If Target.Column = 1 Then
Application.ScreenUpdating = False
CheckChanges = False
Erase Article
Erase Paint
Sheets(1).Range("A2:D2,A4:D80,E:H").ClearComments



Paint(0, 1, 0) = Sheets(1).Cells(2, 2).Value
Paint(1, 1, 1) = 1
Paint(1, 2, 1) = Sheets(1).Cells(2, 1).Value

If IsNumeric(Paint(1, 2, 1)) Then
Paint(1, 2, 1) = Int(Paint(1, 2, 1))
Else
Target = Empty
Application.ScreenUpdating = True
CheckChanges = True
Exit Sub
End If

Find = Application.Match(Paint(1, 2, 1), Sheets("Прайс").Range("L:L"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)

If Find = 0 Then
Erase Paint
Paint(0, 1, 0) = Sheets(1).Cells(2, 2).Value
Paint(1, 1, 1) = 1
'Call PaintConvert2Article
Call ReDefinition
Call PasteCalculation
Call PasteColorantsCalculation
Target = Empty
Application.ScreenUpdating = True
CheckChanges = True
    Exit Sub
End If
Paint(1, 0, 1) = Sheets("Прайс").Cells(Find, 8).Value
Paint(1, 3, 1) = Sheets("Прайс").Cells(Find, 9).Value
Paint(1, 4, 1) = Sheets("Прайс").Cells(Find, 2).Value
Call PaintConvert2Article
Call ReDefinition

Call PasteCalculation
Call PasteColorantsCalculation
Application.ScreenUpdating = True
CheckChanges = True
End If


If Target.Column = 2 Then
CheckChanges = False
If IsNumeric(Target) Then
Paint(0, 1, 0) = Target
Else
Target = Paint(0, 1, 0)
End If
Call PasteColorantsCalculation
CheckChanges = True
End If


End If
CheckChanges = True
End Sub