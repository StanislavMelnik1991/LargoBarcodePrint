Public Paint(8, 4, 1)
Public Article(32, 4, 1)        '| Штрих-код |кол-во |№ (или сокр. штрих-код для краски)|  цена   | *Наименование
                                '|     0     |   1   |                 2                |   3     |      *4
                                '0-32 - положение в списке
                                '0-4  - номер столбца *(№4 выводится в коментарий а не в ячейку)
                                '0-1  - новая и старая запись (1 - новая)
    'выбираем единицы измерения при помощи константы k:
     'может принимать значения "0"  и "3", определяет, в каких единицах измерения выводить расчеты
     '         для расчеётов в "мл" и "л", соответственно
Public Const k As Byte = 3  'кол-во значков после запятой в кол-ве
Public Const ErrorC As Integer = 2 'коэффициент погрешности оборудования (выбирается в зависимости от торговой точки)
Public PageNumber As Integer ' Считает кол-во страниц для печати
Public CheckChanges As Boolean 'закрывает взможность пересчета во время работы скриптов
Public ColorantCost 'запоминает сумарную стоимость колорантов


Public Sub PasteDefinition() 'надо
    'исправляем ошибку пользователя во время копирования из Largo
    Dim k As Integer
    k = 0
temp = Cells(1, 6).Value
If IsNumeric(temp) Then
k = 1
End If
    'построчно переносим информацию, вставленную из буфера, в массив ввода "Article(i, n, 1)"
For i = 0 To 32



Article(i, 1, 1) = Cells(i + 1, 7 + k).Value 'кол-во
            If IsNumeric(Article(i, 1, 1)) And Article(i, 1, 1) <> "" Then
Article(i, 2, 1) = Cells(i + 1, 5 + k).Value 'номер пигмента

Article(i, 1, 1) = Application.RoundUp(Article(i, 1, 1) + ErrorC, 0)
Else
Article(i, 1, 1) = Empty
End If
Next i
End Sub
Public Sub ReDefinition()  'надо
C = 0
For i = 0 To 32
Do While Cells(i + C + 4, 3).Value = Empty Or Cells(i + C + 4, 3).Value = "Краска"
C = C + 1
If i + C = 32 Then
Exit For
End If
Loop
Article(i, 2, 1) = Cells(i + C + 4, 3).Value 'номер пигмента
Article(i, 1, 1) = Cells(i + C + 4, 2).Value 'кол-во


Find = Application.Match(Format(Article(i, 2, 1), "##00"), Sheets("красители").Range("A:A"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)
If Find = 0 And Article(i, 2, 1) <> "Краска" Then
Article(i, 2, 1) = Empty
Cells(i + 4, 3).ClearContents
End If



If Article(i, 2, 1) = "Краска" Then
Article(i, 1, 1) = Cells(i + 4, 2).Value  'кол-во

End If
            If IsNumeric(Article(i, 1, 1)) And Article(i, 1, 1) <> Empty Then
If IsNumeric(Article(i, 2, 1)) Then
Article(i, 2, 1) = Int(Article(i, 2, 1))
End If
Article(i, 1, 1) = Application.RoundUp(Article(i, 1, 1) / Paint(0, 1, 0) * 10 ^ k, 0)
           End If
Next i
End Sub
Public Sub PasteCalculation() 'nado
            For i = 0 To 32
If Article(i, 2, 1) = "Краска" Then
Article(i, 0, 1) = Paint(0, 0, 0)
Article(i, 3, 1) = Paint(0, 3, 0)
i = i + 1
End If
            If Article(i, 1, 1) <> Empty Then
Find = Application.Match(Format(Article(i, 2, 1), "##00"), Sheets("красители").Range("A:A"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)

If Find <> Empty Then
Article(i, 3, 1) = Sheets("красители").Cells(Find, 3).Value
Article(i, 0, 1) = Sheets("красители").Cells(Find, 2).Value
'temp = Abs(Right(Article(0, 0, 0), 6))
Find = Application.Match(Abs(Right(Article(i, 0, 1), 6)), Sheets("Прайс").Range("L:L"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)
Article(i, 4, 1) = Sheets("Прайс").Cells(Find, 2).Value

End If
            End If
            Next i
End Sub


Sub PasteColorantsCalculation()
CheckChanges = False
Application.ScreenUpdating = False
'Call lookArrayArticle(0)
Range("A4:D30").ClearContents
ColorantCost = 0                'суммарная стоимость колеровки
Dim cost(32)                    'массив цены колоранта (с учетом кол-ва)

Call ConvertInToOut
'Call lookArrayArticle(0)
i1 = 0
For i = 0 To 32
If Article(i + i1, 1, 0) = Empty Then
Exit For
End If

temp = Application.RoundUp(Article(i, 1, 0) * Paint(0, 1, 0) / (10 ^ k), k)
                                    'считаем стоимость колоранта с учетом объема и округляем ее вверх до сотых
cost(i) = Application.RoundUp(Article(i, 3, 0) * Article(i, 1, 0) * Paint(0, 1, 0) / 1000, 2)


If Article(i, 2, 0) <> "Краска" Then
ColorantCost = ColorantCost + cost(i)                   'добавляем стоимость колоранта(i) к общей стоимости колеровки

    If k = 3 Then
    Cells(4 + i, 2).NumberFormat = "0.000"
    Else
    Cells(4 + i, 2).NumberFormat = "0"
    End If
Else
temp = Application.RoundUp(temp, 0)
cost(i) = Application.RoundUp(Article(i, 3, 0) * temp * 10 ^ k / 1000, 2)
Cells(4 + i, 2).NumberFormat = "0"  'убираем дробную часть из отоброжения (краска не может быть в кол-ве 0,5 банки)

approximateCost = approximateCost + cost(i)
End If
Cells(4 + i, 3).ClearComments
Cells(4 + i, 3).AddComment
Cells(4 + i, 3).Comment.Visible = False
Cells(4 + i, 3).Comment.Text Text:=Article(i, 4, 0)
Cells(4 + i, 4).FormulaR1C1 = cost(i)                   'выводим значение стоимости колоранта(i)
    'выводим значения элементов(i)
Cells(4 + i, 1).FormulaR1C1 = Article(i, 0, 0)



Cells(4 + i, 2).FormulaR1C1 = temp
    'костыль, убирающий "0" в последней строчке
If Cells(4 + i, 2).FormulaR1C1 = 0 Then
Cells(4 + i, 2).ClearContents
Cells(4 + i, 4).ClearContents
End If
Cells(4 + i, 3).FormulaR1C1 = Format(Article(i, 2, 0), "##00")

Next i
    'считаем стоимость колеровки (точная) вместе со стоимостью краски (не точная)

approximateCost = approximateCost + ColorantCost
approximateCost = "Примерная стоимость" & Chr(13) & Chr(10) & "вместе с краской: " & Chr(13) & Chr(10) & approximateCost & " руб."
ColorantCost = Format(ColorantCost, "#,##0.00 руб.")    'выводим стоимость колеровки
Cells(2, 3).FormulaR1C1 = ColorantCost

Cells(2, 3).ClearComments
Cells(2, 3).AddComment
Cells(2, 3).Comment.Visible = False
Cells(2, 3).Comment.Text Text:=approximateCost
Application.ScreenUpdating = True
'CheckChanges = True
End Sub

Sub ConvertInToOut()
n = 0
For i = 0 To 32
                'пропускаем пустые элементы вводимого массива (избегаем пустых строк в выводимом)
Do While Article(n, 1, 1) = Empty And n < 31
n = n + 1
If n = 32 Then
Exit Sub
End If
Loop
                'пропускаем заполненые элементы выводимого массива (избегаем пустых строк в выводимом)
Do While Article(i, 1, 0) <> Empty And i < 31
i = i + 1
If i = 32 Then
Exit For
End If
Loop

For i1 = 0 To 4
Article(i, i1, 0) = Article(n, i1, 1)       'переносим данные из вводимой части массива в выводимую
Article(n, i1, 1) = Empty
Next i1
n = n + 1
If n >= 31 Then
Exit Sub
End If
Next i
End Sub

Sub PasteAddDefinition()
      'исправляем ошибку пользователя во время копирования из Largo
    Dim k As Integer
    k = 0
temp = Cells(1, 6).Value
If IsNumeric(temp) Then
k = 1
End If
i1 = 0
Paint(0, 1, 1) = InputBox("Введите количество краски этого цвета", "Красители", 1)
If Paint(0, 1, 1) = "" Then
Exit Sub
End If
    'построчно переносим информацию, вставленную из буфера, в массив ввода "Article(i, n, 1)"
For i = 0 To 32
Do While Article(i1, 1, 1) > 0
i1 = i1 + 1
If i + i1 >= 32 Then
Exit For
End If
Loop
Article(i1, 1, 1) = Cells(i + 1, 7 + k).Value 'кол-во
            If IsNumeric(Article(i1, 1, 1)) And Article(i1, 1, 1) <> "" Then
Article(i1, 2, 1) = Cells(i + 1, 5 + k).Value 'номер пигмента
Article(i1, 1, 1) = Application.RoundUp((Article(i1, 1, 1) + ErrorC), 0) * Paint(0, 1, 1)
i2 = 0
    For i2 = 0 To (i1 - 1)
If Article(i2, 2, 1) = Article(i1, 2, 1) Then
Article(i2, 1, 1) = Article(i2, 1, 1) + Article(i1, 1, 1)
If Article(i2, 1, 1) <= 0 Then
Article(i2, 1, 1) = Empty
End If
Article(i1, 2, 1) = Empty
Article(i1, 1, 1) = Empty
End If
    Next i2
Else
Article(i1, 1, 1) = Empty
End If
Next i
End Sub


Sub PaintAddDefinition()
i = 1
Do While Paint(i, 2, 1) <> Empty
i = i + 1
If i = 8 Then
Exit Sub
End If
Loop
Paint(i, 2, 1) = InputBox("Введите последние 6 цифр штрих-кода краски", "Краска", 0)
If IsNumeric(Paint(i, 2, 1)) Then
Paint(i, 2, 1) = Int(Paint(i, 2, 1))
Else
Exit Sub
End If
If Paint(i, 2, 1) = "" Then
Exit Sub
End If
Find = Application.Match(Paint(i, 2, 1), Sheets("Прайс").Range("L:L"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)

If Find = Empty Then
For i1 = 0 To 4
Paint(i, i1, 1) = Empty
Next i1
    Exit Sub
End If
Paint(i, 1, 1) = Int(InputBox("Введите количество данной краски", "Краска", 1))
If Paint(i, 1, 1) = "" Then
Exit Sub
End If
For i1 = 0 To (i - 1)
If Paint(i, 2, 1) = Paint(i1, 2, 1) Then
Paint(i1, 1, 1) = Paint(i1, 1, 1) + Paint(i, 1, 1)
Paint(i, 1, 1) = Empty
Paint(i, 2, 1) = Empty
Exit Sub
End If
Next i1
Paint(i, 0, 1) = Sheets("Прайс").Cells(Find, 8).Value
Paint(i, 3, 1) = Sheets("Прайс").Cells(Find, 9).Value
Paint(i, 4, 1) = Sheets("Прайс").Cells(Find, 2).Value

End Sub
Sub PaintConvert2Article()
For A = 0 To 31
temp = Cells(4 + A, 3).Value
If temp = "Краска" Then
Cells(4 + A, 1).ClearContents
Cells(4 + A, 2).ClearContents
Cells(4 + A, 3).ClearContents
Cells(4 + A, 4).ClearContents
End If
Next A
i1 = 1
For i = 0 To 8
Do While Paint(i1, 2, 1) = Empty And i1 < 8
i1 = i1 + 1
Loop
If Paint(i1, 2, 1) = Empty Then
Exit For
End If
Article(i, 0, 0) = Paint(i1, 0, 1)
Article(i, 1, 0) = Paint(i1, 1, 1) * 10 ^ k
Article(i, 2, 0) = "Краска"
Article(i, 3, 0) = Paint(i1, 3, 1)
Article(i, 4, 0) = Paint(i1, 4, 1)
i1 = i1 + 1
Next i
End Sub
Public Sub AddCalculation(Pozition)

            If Article(Pozition, 1, 0) <> Empty Then
Find = Application.Match(Format(Article(Pozition, 2, 0), "##00"), Sheets("красители").Range("A:A"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)

If Find <> Empty Then
Article(Pozition, 3, 0) = Sheets("красители").Cells(Find, 3).Value
Article(Pozition, 0, 0) = Sheets("красители").Cells(Find, 2).Value
Find = Application.Match(Abs(Right(Article(Pozition, 0, 0), 6)), Sheets("Прайс").Range("L:L"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)
Article(Pozition, 4, 0) = Sheets("Прайс").Cells(Find, 2).Value
End If
            End If

End Sub

'скрипты для печати (module3)
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'



Public Sub BarCodeStart()    'просчитывает и отрисовывает полосы штрих-кода на выбранном листе
            str1 = 4         ' номер первой строки
            PageNumber = 0
            pozx = str1
Call PageQuantity(pozx)      'считаем кол-во заполненных листов
If PageNumber = 0 Then
Exit Sub
End If
                             'Рисуем полоски
            begy = 50        ' положение первой строки
            begx = 0
 Call BarCode(str1, begx, begy, PageNumber)
End Sub
Public Sub BarCode(str1, begx, begy, PageNumber)
        For L = 1 To PageNumber
    For str0 = 1 To 1
bk = Cells(str1, 1).Value
If bk = "" Then
Else
Dim k(4, 0 To 9) As String  'перечислим все значения...
k(1, 0) = "0001101": k(1, 1) = "0011001": k(1, 2) = "0010011": k(1, 3) = "0111101": k(1, 4) = "0100011": k(1, 5) = "0110001": k(1, 6) = "0101111": k(1, 7) = "0111011": k(1, 8) = "0110111": k(1, 9) = "0001011"
k(2, 0) = "0100111": k(2, 1) = "0110011": k(2, 2) = "0011011": k(2, 3) = "0100001": k(2, 4) = "0011101": k(2, 5) = "0111001": k(2, 6) = "0000101": k(2, 7) = "0010001": k(2, 8) = "0001001": k(2, 9) = "0010111"
k(3, 0) = "1110010": k(3, 1) = "1100110": k(3, 2) = "1101100": k(3, 3) = "1000010": k(3, 4) = "1011100": k(3, 5) = "1001110": k(3, 6) = "1010000": k(3, 7) = "1000100": k(3, 8) = "1001000": k(3, 9) = "1110100"
k(4, 0) = "000000": k(4, 1) = "001011": k(4, 2) = "001101": k(4, 3) = "001110": k(4, 4) = "010011": k(4, 5) = "011001": k(4, 6) = "011100": k(4, 7) = "010101": k(4, 8) = "010110": k(4, 9) = "011010"
kb = "101"      'бордюр
kr = "01010"    'разделитель

    'каждая цифра (кроме первой) кодируется 7 битами,
    'причём система сделана так, что всегда 1 цифра
    'выглядит как две полоски и два пробела, причём
    'толщина их зависит от количества идущих подряд
    'одинаковых битов;
    'всего 4 таблицы k(1,x)...k(4,x);
    'первая и вторая кодируют с 2 по 7 цифры;
    'третья кодирует с 8 по 13 цифру;
    'ноль равен 1 пробелу, два нуля подряд пробелу пошире, три нуля
    'ещё более широкому пробелу, а 4 сАмому широкому;
    'единица равна 1 линии, две подряд двойной, три тройной,
    'а 4 четверной;
    'четвёртая таблица кодирует 1 цифру;
    'первая цифра задаёт как кодировать цифры с 2 по 7;
    'ноль четвёртой таблицы показывает, что цифра берётся из первой таблицы;
    'единица показывает, что цифра берётся из второй таблицы;
    'например первая цифра = 2, тогда исходя из данных k(4,2)=001101:
    '2 цифра (0)01101 будет взята из 1 таблицы k(1,x)
    '3 цифра 0(0)1101 будет взята из 1 таблицы k(1,x)
    '4 цифра 00(1)101 будет взята из 2 таблицы k(2,x)
    '5 цифра 001(1)01 будет взята из 2 таблицы k(2,x)
    '6 цифра 0011(0)1 будет взята из 1 таблицы k(1,x)
    '7 цифра 00110(1) будет взята из 2 таблицы k(2,x)


'редактируем введёную строку доведя её до 12 знаков нулями слева

If Len(bk) < 12 Then bk = String(12 - Len(bk), "0") & bk
bk = Left(bk, 12)

'вычисляем контрольную цифру
    sy = 0
    For rt = 1 To 12        'перебираем все 12 значащих цифр
        sy = sy + Val(Mid(bk, rt, 1)) * (1 + 2 * ((rt + 1) Mod 2))  'суммируем
                            'все цифры кода, причём каждая вторая цифра
                            'домножается на 3
    Next rt
    sy = 10 - sy Mod 10     'теперь sy равно числу, дополняющему старое
                            'sy до ровного десятка (это и есть контрольная цифра)
    If sy = 10 Then sy = 0  'если получилось 10, то оставляем только 0
    bk = bk & sy            'дописываем контрольную цифру к коду


'далее составляем двоичной код для изображения
dk = kb     'сначала ставим бордюр
    For t = 1 To 6
        If Mid(k(4, Mid(bk, 1, 1)), t, 1) = "0" Then ki = 1 Else ki = 2 'выбираем таблицу кодирования по маске
        dk = dk & k(ki, Mid(bk, t + 1, 1))
    Next t
dk = dk & kr    'ставим разделитель
    For t = 7 To 12
        dk = dk & k(3, Mid(bk, t + 1, 1))
    Next t
dk = dk & kb    'опять бордюр

'рисуем полоски
For ee = 1 To 95
    If Mid(dk, ee, 1) = "1" Then
            'удлиняем служебные полосы
        If ee < 4 Or ee > 46 And ee < 50 Or ee > 92 Then yt = 2.6 Else yt = 0
            'рисуем полоску; 1.5 это коэффициент привязки экселевских координат
            'к ширине шага рисунка; 12 это отступ от левого края таблицы;
            '60 это высота линий (если это служебные, то высота получается 68)AddLine
            'AddLine(beginX,beginY,endX,endY)
        ActiveSheet.Shapes.AddLine(ee * 1.5 + 12 + begx, 3 + begy, ee * 1.5 + 12 + begx, 20 + yt + begy).Select
            'подгоняем ширину полосок под ширину пробелов
        Selection.ShapeRange.Line.Weight = 1.4
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)
            'даём имя типа Lхх, где хх - порядковый номер бита в коде
            'эта операция необходима для корректной очистки при расчёте нового кода
        Selection.Name = L & str0 & ee
    End If
Next ee
End If
str1 = str1 + 1
begy = begy + 36.75

    Next str0
'str1 = str1 + 1
'begy = begy + 13.5
        Next L
            Cells(2, 1).Select
End Sub

Public Sub PageQuantity(pozx) 'считает кол-во страниц для печати, содержащих информацию
For pageprint = 1 To 35                                         'пропускаем заполненные строки
findx2 = Cells(pozx, 1).Value
If findx2 <> "" Then
PageNumber = PageNumber + 1
Else
Exit For
End If
pozx = pozx + 1
Next pageprint
End Sub

Sub ClearBarcode()
        For L = 1 To PageNumber
    For str0 = 1 To 3
For ee = 1 To 95    'всего 95 битов данных
    On Error Resume Next
    ActiveSheet.Shapes(L & str0 & ee).Delete
Next ee
    Next str0
        Next L
End Sub

Public Sub PrintPage(copies)
    ActiveWindow.SelectedSheets.PrintOut copies:=copies, Collate:=True, _
        IgnorePrintAreas:=False
End Sub

'
'
'
'
'
'
'
'
'макросы, необходимые для отладки (module4)
'
'
'
'
'
'
'
'
'
'
'

Sub lookArrayArticle(version)
'Call AddColorants
message1 = Empty
'version = 0
For i = 0 To 32
message1 = message1 & Chr(13) & Chr(10) & i & "; " & Article(i, 0, version) & "; " & Article(i, 1, version) & "; " & Article(i, 2, version) & "; " & Article(i, 3, version)
Next i
MsgBox message1
End Sub
Sub lookArrayArticleManual()
'Call AddColorants
message1 = Empty
version = 0
For i = 0 To 32
message1 = message1 & Chr(13) & Chr(10) & i & "; " & Article(i, 0, version) & "; " & Article(i, 1, version) & "; " & Article(i, 2, version) & "; " & Article(i, 3, version) & "; " & Article(i, 4, version)
Next i
MsgBox message1
End Sub
Sub lookArrayPaint()
'Call AddColorants
message1 = Empty
version = 0
For i = 0 To 8
message1 = message1 & Chr(13) & Chr(10) & i & "; " & Paint(i, 0, version) & "; " & Paint(i, 1, version) & "; " & Paint(i, 2, version) & "; " & Paint(i, 3, version)
Next i
MsgBox message1
End Sub
Sub DoSomething()
'MsgBox Len(Article(0, 0, 0))
temp = Abs(Right(Article(0, 0, 0), 6))
Find = Application.Match(temp, Sheets("Ïðàéñ").Range("L:L"), 0)
Find = Application.WorksheetFunction.IfError(Find, Empty)
Article(0, 4, 0) = Sheets("Ïðàéñ").Cells(Find, 2).Value
MsgBox Article(0, 4, 0)
'CheckChanges = True
'Cells(2, 1).FormulaR1C1 = 7
'Cells(2, 2).Select

   ' ActiveSheet.Paste
End Sub
'test

