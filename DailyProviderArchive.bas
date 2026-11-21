REM  *****  BASIC  *****
Option Explicit

Sub DailyProviderArchive
    On Error GoTo ErrorHandler

    Dim doc As Object
    Dim main As Object
    Dim history As Object
    Dim todayDate As Double
    Dim previousDate As Double
    Dim nextRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim alreadyArchived As Boolean
    Dim cell As Object
    Dim cursor As Object

    doc = ThisComponent
    doc.calculateAll()

    main = doc.Sheets.getByName("Ежедневный учет")
    history = doc.Sheets.getByName("Статистика по дням")

    todayDate = main.getCellRangeByName("B3").Value
    previousDate = main.getCellRangeByName("H3").Value

    If todayDate <= 0 Then
        MsgBox "Не удалось получить дату из B3. Проверь формулу =СЕГОДНЯ().", 48, "Учет ДИ"
        Exit Sub
    End If

    If previousDate <= 0 Then
        main.getCellRangeByName("H3").Value = todayDate
        doc.store()
        MsgBox "Первый запуск выполнен. Дата записана в H3. Текущие данные сохранены.", 64, "Учет ДИ"
        Exit Sub
    End If

    If todayDate <= previousDate Then Exit Sub

    alreadyArchived = False
    cursor = history.createCursor()
    cursor.gotoEndOfUsedArea(True)
    lastRow = cursor.RangeAddress.EndRow

    For i = 1 To lastRow
        cell = history.getCellByPosition(0, i)
        If cell.Type <> com.sun.star.table.CellContentType.EMPTY Then
            If Abs(cell.Value - previousDate) < 0.0001 Then
                alreadyArchived = True
                Exit For
            End If
        End If
    Next i

    nextRow = 1
    Do While history.getCellByPosition(0, nextRow).Type <> com.sun.star.table.CellContentType.EMPTY
        nextRow = nextRow + 1
    Loop

    If Not alreadyArchived Then
        history.getCellByPosition(0, nextRow).Value = previousDate
        history.getCellRangeByName("A" & CStr(nextRow + 1)).NumberFormat = main.getCellRangeByName("B3").NumberFormat

        history.getCellByPosition(1, nextRow).Value = main.getCellRangeByName("C5").Value
        history.getCellByPosition(2, nextRow).Value = main.getCellRangeByName("E5").Value

        history.getCellByPosition(3, nextRow).Value = main.getCellRangeByName("C6").Value
        history.getCellByPosition(4, nextRow).Value = main.getCellRangeByName("E6").Value
    End If

    main.getCellRangeByName("C5:C9").clearContents(23)
    main.getCellRangeByName("E5:F9").clearContents(23)

    main.getCellRangeByName("H3").Value = todayDate
    doc.calculateAll()
    doc.store()
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка макроса: " & Error$ & Chr(10) & _
           "Проверь названия листов и адреса ячеек.", 16, "Учет ДИ"
End Sub
