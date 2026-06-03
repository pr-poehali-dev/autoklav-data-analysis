import { useState, useRef } from "react";
import Icon from "@/components/ui/icon";

const VBA_CODE = `'=============================================================
' АВТОКЛАВ — Расчёт стерилизационного эффекта F0
' Параметры: Tref = 121.1 C, z = 10 C
' Алгоритм: автоопределение циклов + расчёт F0 по формуле
'=============================================================

Option Explicit

' Глобальные константы стерилизации
Const T_REF As Double = 121.1   ' Эталонная температура (C)
Const Z_FACTOR As Double = 10#   ' Z-фактор (C)
Const T_START As Double = 30#    ' Порог начала цикла (C)
Const T_MIN_STERIL As Double = 100# ' Минимум для счёта F0

'-------------------------------------------------------------
' Главная процедура — точка входа
'-------------------------------------------------------------
Sub Autoclave_ProcessCSV()
    Dim wb As Workbook
    Dim wsData As Worksheet, wsReport As Worksheet
    Dim filePath As String
    Dim lastRow As Long

    filePath = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        Title:="Выберите CSV-файл с данными автоклава")
    If filePath = "False" Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ActiveWorkbook
    Call ImportAndParseCSV(wb, filePath, wsData)
    Call PrepareReportSheet(wb, wsReport)

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Call DetectCyclesAndCalculateF0(wsData, wsReport, lastRow)
    Call FormatReportSheet(wsReport)
    Call BuildTemperatureChart(wb, wsData, lastRow)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Расчёт завершён! Результаты на листе 'F0_Report'.", _
           vbInformation, "Автоклав F0 — Готово"
End Sub

'-------------------------------------------------------------
' Импорт CSV и организация по столбцам
'-------------------------------------------------------------
Sub ImportAndParseCSV(wb As Workbook, filePath As String, ByRef wsData As Worksheet)
    Dim ws As Worksheet
    Dim fileNum As Integer
    Dim lineText As String
    Dim fields() As String
    Dim rowNum As Long
    Dim i As Integer

    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name = "Data" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    Set wsData = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsData.Name = "Data"

    ' Заголовки — 17 столбцов реального CSV автоклава + F0
    Dim headers(1 To 18) As String
    headers(1)  = "ДАТА"
    headers(2)  = "ВРЕМЯ"
    headers(3)  = "МИЛЛИСЕКУНДЫ"
    headers(4)  = "ТЕМПЕРАТУРА СРЕДЫ"
    headers(5)  = "ТЕМПЕРАТУРА ПРОДУКТА"
    headers(6)  = "ДАВЛЕНИЕ"
    headers(7)  = "УРОВЕНЬ ВОДЫ"
    headers(8)  = "Q recirculación"
    headers(9)  = "ДАВЛЕНИЕ СЖАТОГО ВОЗДУХА"
    headers(10) = "SP presión"
    headers(11) = "ЗАДАНИЕ ТЕМПЕРАТУРЫ"
    headers(12) = "КЛАПАН ПОВЫШЕНИЯ ДАВЛЕНИЯ"
    headers(13) = "КЛАПАН СБРОСА ДАВЛЕНИЯ"
    headers(14) = "КЛАПАН НАГРЕВА"
    headers(15) = "КЛАПАН ОХЛАЖДЕНИЯ"
    headers(16) = "КЛАПАН ЗАПОЛНЕНИЯ"
    headers(17) = "НОМЕР ЗАМЕСА"
    headers(18) = "F0_накоп."

    For i = 1 To 18
        wsData.Cells(1, i).Value = headers(i)
    Next i

    fileNum = FreeFile
    Open filePath For Input As #fileNum
    rowNum = 2

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) = 0 Then GoTo NextLine
        If Left(lineText, 1) = "#" Then GoTo NextLine

        If InStr(lineText, ";") > 0 Then
            fields = Split(lineText, ";")
        Else
            fields = Split(lineText, ",")
        End If

        ' -------------------------------------------------------
        ' Очищаем кавычки из всех полей (CSV обёрнут в "...")
        ' -------------------------------------------------------
        Dim k As Integer
        For k = 0 To UBound(fields)
            fields(k) = Trim(fields(k))
            ' Убираем обрамляющие кавычки
            If Len(fields(k)) >= 2 Then
                If Left(fields(k), 1) = Chr(34) And Right(fields(k), 1) = Chr(34) Then
                    fields(k) = Mid(fields(k), 2, Len(fields(k)) - 2)
                End If
            End If
        Next k

        ' -------------------------------------------------------
        ' Пропускаем строку заголовков CSV (содержит нечисловое в поле 0)
        ' -------------------------------------------------------
        If rowNum = 2 Then
            Dim f0clean As String
            f0clean = LCase(Trim(fields(0)))
            If InStr(f0clean, "fecha") > 0 Or InStr(f0clean, "date") > 0 Or _
               InStr(f0clean, "дата") > 0 Or Left(f0clean, 4) = "data" Or _
               Not (f0clean Like "####/##/##" Or f0clean Like "##/##/####" Or IsDate(f0clean)) Then
                ' Дополнительная проверка: если поле 4 (темп. продукта) тоже не число — это заголовок
                Dim fCheck As String
                fCheck = Trim(fields(4))
                If Not IsNumeric(fCheck) Then GoTo NextLine
            End If
        End If

        ' -------------------------------------------------------
        ' Записываем все 17 столбцов с правильным преобразованием
        ' -------------------------------------------------------
        Dim totalCols As Integer
        totalCols = UBound(fields) + 1
        If totalCols > 17 Then totalCols = 17

        For i = 0 To totalCols - 1
            Dim cellVal As String
            cellVal = Trim(fields(i))

            Select Case i
                Case 0 ' Столбец A — Дата формата YYYY/MM/DD
                    ' Конвертируем YYYY/MM/DD → Excel дата
                    Dim dateParts() As String
                    If InStr(cellVal, "/") > 0 Then
                        dateParts = Split(cellVal, "/")
                        If UBound(dateParts) = 2 Then
                            Dim yr As Integer, mo As Integer, dy As Integer
                            yr = CInt(dateParts(0))
                            mo = CInt(dateParts(1))
                            dy = CInt(dateParts(2))
                            wsData.Cells(rowNum, 1).Value = DateSerial(yr, mo, dy)
                        Else
                            wsData.Cells(rowNum, 1).Value = cellVal
                        End If
                    ElseIf IsDate(cellVal) Then
                        wsData.Cells(rowNum, 1).Value = CDate(cellVal)
                    Else
                        wsData.Cells(rowNum, 1).Value = cellVal
                    End If

                Case 1 ' Столбец B — Время формата HH:MM:SS
                    If IsDate(cellVal) Then
                        wsData.Cells(rowNum, 2).Value = TimeValue(cellVal)
                    ElseIf InStr(cellVal, ":") > 0 Then
                        On Error Resume Next
                        wsData.Cells(rowNum, 2).Value = TimeValue(cellVal)
                        On Error GoTo 0
                    Else
                        wsData.Cells(rowNum, 2).Value = cellVal
                    End If

                Case Else ' Столбцы 3–17 — числа (возможна точка или запятая)
                    cellVal = Replace(cellVal, ",", ".")
                    If IsNumeric(cellVal) Then
                        wsData.Cells(rowNum, i + 1).Value = CDbl(cellVal)
                    Else
                        wsData.Cells(rowNum, i + 1).Value = cellVal
                    End If
            End Select
        Next i

        rowNum = rowNum + 1
NextLine:
    Loop
    Close #fileNum

    ' Форматирование
    wsData.Columns(1).NumberFormat = "dd.mm.yyyy"
    wsData.Columns(2).NumberFormat = "hh:mm:ss"
    wsData.Columns(3).NumberFormat = "0"
    wsData.Columns(4).Resize(, 14).NumberFormat = "0.00"
    wsData.Columns(18).NumberFormat = "0.0000"
    wsData.Rows(1).Font.Bold = True
    wsData.Rows(1).Interior.Color = RGB(15, 40, 65)
    wsData.Rows(1).Font.Color = RGB(0, 180, 220)
    wsData.Columns(1).Resize(, 18).AutoFit
End Sub

'-------------------------------------------------------------
' Подготовка листа F0_Report
'-------------------------------------------------------------
Sub PrepareReportSheet(wb As Workbook, ByRef wsReport As Worksheet)
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name = "F0_Report" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    Set wsReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsReport.Name = "F0_Report"

    With wsReport
        .Cells(1, 1).Value = "ПРОТОКОЛ СТЕРИЛИЗАЦИИ — РАСЧЁТ F0"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 180, 220)
        .Range("A1:J1").Merge

        .Cells(2, 1).Value = "Дата формирования: " & Format(Now, "dd.mm.yyyy hh:mm")
        .Cells(2, 1).Font.Color = RGB(120, 140, 160)
        .Range("A2:J2").Merge

        .Cells(3, 1).Value = "Параметры: Tref = 121.1C  |  z-фактор = 10C  |  Метод: трапеций"
        .Cells(3, 1).Font.Color = RGB(120, 140, 160)
        .Range("A3:J3").Merge

        .Rows(4).RowHeight = 6

        Dim colHeaders(1 To 10) As String
        colHeaders(1) = "Цикл"
        colHeaders(2) = "Начало"
        colHeaders(3) = "Конец"
        colHeaders(4) = "Длит. (мин)"
        colHeaders(5) = "T макс. (C)"
        colHeaders(6) = "T мин. (C)"
        colHeaders(7) = "F0 (мин)"
        colHeaders(8) = "Результат"
        colHeaders(9) = "Строк"
        colHeaders(10) = "Примечание"

        Dim c As Integer
        For c = 1 To 10
            .Cells(5, c).Value = colHeaders(c)
            .Cells(5, c).Font.Bold = True
            .Cells(5, c).Font.Color = RGB(255, 255, 255)
            .Cells(5, c).Interior.Color = RGB(15, 40, 65)
            .Cells(5, c).HorizontalAlignment = xlCenter
            .Cells(5, c).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Cells(5, c).Borders(xlEdgeBottom).Color = RGB(0, 180, 220)
        Next c
    End With
End Sub

'-------------------------------------------------------------
' Определение циклов стерилизации и расчёт F0
'-------------------------------------------------------------
Sub DetectCyclesAndCalculateF0(wsData As Worksheet, wsReport As Worksheet, lastRow As Long)
    Dim i As Long
    Dim cycleNum As Integer
    Dim inCycle As Boolean
    Dim cycleStart As Long
    Dim reportRow As Integer
    Dim tempProd As Double
    Dim f0Cycle As Double
    Dim tMax As Double, tMin As Double
    Dim cycleStartTime As Variant
    Dim cycleEndTime As Variant
    Dim durationMin As Double
    Dim peakReached As Boolean

    Const COL_DATE As Integer = 1      ' Столбец A — Дата
    Const COL_TIME As Integer = 2      ' Столбец B — Время
    Const COL_TEMP_PROD As Integer = 5 ' Столбец E — Температура продукта
    Const COL_F0 As Integer = 18       ' Столбец R — накопленный F0

    cycleNum = 0
    inCycle = False
    reportRow = 6
    peakReached = False
    tMax = -999
    tMin = 999
    f0Cycle = 0
    cycleStart = 2

    For i = 2 To lastRow
        Dim rawVal As Variant
        rawVal = wsData.Cells(i, COL_TEMP_PROD).Value
        If Not IsNumeric(rawVal) Or rawVal = "" Then GoTo NextRow
        tempProd = CDbl(rawVal)

        If Not inCycle Then
            ' Начало цикла: температура поднимается выше T_START
            ' Не требуем подтверждения роста — достаточно превышения порога
            If tempProd >= T_START Then
                inCycle = True
                cycleStart = i
                f0Cycle = 0
                tMax = tempProd
                tMin = tempProd
                peakReached = False

                ' Сохраняем дату+время начала цикла
                Dim dv As Variant, tv As Variant
                dv = wsData.Cells(i, COL_DATE).Value
                tv = wsData.Cells(i, COL_TIME).Value
                If IsNumeric(dv) And IsNumeric(tv) Then
                    cycleStartTime = CDbl(dv) + CDbl(tv)
                Else
                    cycleStartTime = 0
                End If
            End If
        Else
            If tempProd > tMax Then tMax = tempProd
            If tempProd < tMin Then tMin = tempProd
            If tMax >= T_MIN_STERIL Then peakReached = True

            If tempProd >= T_MIN_STERIL Then
                Dim lethality As Double
                lethality = 10 ^ ((tempProd - T_REF) / Z_FACTOR)

                ' Вычисляем Δt в минутах из столбцов A (дата) + B (время)
                ' Дата хранится как целое число Excel, время — как дробная часть (0..1)
                Dim deltaT As Double
                deltaT = 1  ' значение по умолчанию — 1 минута
                If i > cycleStart Then
                    Dim d1 As Variant, d2 As Variant
                    Dim t1v As Variant, t2v As Variant
                    d1 = wsData.Cells(i - 1, COL_DATE).Value
                    t1v = wsData.Cells(i - 1, COL_TIME).Value
                    d2 = wsData.Cells(i, COL_DATE).Value
                    t2v = wsData.Cells(i, COL_TIME).Value
                    ' Оба значения должны быть числами (Excel хранит дату/время как число)
                    If IsNumeric(d1) And IsNumeric(t1v) And IsNumeric(d2) And IsNumeric(t2v) Then
                        Dim dt As Double
                        ' Полная дата-время = дата (целая) + время (дробная часть)
                        dt = ((CDbl(d2) + CDbl(t2v)) - (CDbl(d1) + CDbl(t1v))) * 24# * 60#
                        ' Защита: интервал должен быть положительным и разумным (< 30 мин)
                        If dt > 0# And dt < 30# Then
                            deltaT = dt
                        ElseIf dt <= 0# Then
                            ' Возможно миллисекунды в столбце C дают точность
                            Dim ms1 As Variant, ms2 As Variant
                            ms1 = wsData.Cells(i - 1, 3).Value  ' столбец C — миллисекунды
                            ms2 = wsData.Cells(i, 3).Value
                            If IsNumeric(ms1) And IsNumeric(ms2) Then
                                Dim dtMs As Double
                                dtMs = (CDbl(ms2) - CDbl(ms1)) / 1000# / 60#  ' мс → минуты
                                If dtMs > 0# And dtMs < 30# Then deltaT = dtMs
                            End If
                        End If
                    End If
                End If

                f0Cycle = f0Cycle + lethality * deltaT
            End If

            ' Пишем накопленный F0 в столбец R (18)
            wsData.Cells(i, COL_F0).Value = Round(f0Cycle, 4)

            Dim cycleEnds As Boolean
            cycleEnds = False
            If peakReached And tempProd < T_START Then cycleEnds = True
            If i = lastRow Then cycleEnds = True

            If cycleEnds Then
                cycleNum = cycleNum + 1

                ' Время конца цикла — дата + время
                Dim deV As Variant, teV As Variant
                deV = wsData.Cells(i, COL_DATE).Value
                teV = wsData.Cells(i, COL_TIME).Value
                If IsNumeric(deV) And IsNumeric(teV) Then
                    cycleEndTime = CDbl(deV) + CDbl(teV)
                ElseIf IsDate(deV) And IsDate(teV) Then
                    cycleEndTime = CDbl(CDate(deV)) + CDbl(CDate(teV))
                Else
                    cycleEndTime = 0
                End If

                durationMin = 0
                If cycleStartTime > 0 And cycleEndTime > 0 Then
                    durationMin = (CDbl(cycleEndTime) - CDbl(cycleStartTime)) * 24 * 60
                End If

                Dim result As String
                Dim resultColor As Long
                Dim noteText As String

                If Not peakReached Then
                    result = "— Без стерилизации"
                    resultColor = RGB(100, 120, 140)
                    noteText = "Пик T < 100C — только простой"
                ElseIf f0Cycle >= 6 Then
                    result = "OK НОРМА (F0 >= 6)"
                    resultColor = RGB(0, 160, 80)
                    noteText = "Tref=121.1C, z=10C"
                ElseIf f0Cycle >= 3 Then
                    result = "! ПРЕДЕЛ (3 <= F0 < 6)"
                    resultColor = RGB(220, 140, 0)
                    noteText = "Tref=121.1C, z=10C"
                Else
                    result = "X НЕДОСТАТОЧНО (F0 < 3)"
                    resultColor = RGB(200, 40, 40)
                    noteText = "Tref=121.1C, z=10C"
                End If

                With wsReport
                    .Cells(reportRow, 1).Value = cycleNum
                    .Cells(reportRow, 2).Value = cycleStartTime
                    .Cells(reportRow, 3).Value = cycleEndTime
                    .Cells(reportRow, 4).Value = Round(durationMin, 1)
                    .Cells(reportRow, 5).Value = Round(tMax, 2)
                    .Cells(reportRow, 6).Value = Round(tMin, 2)
                    .Cells(reportRow, 7).Value = Round(f0Cycle, 4)
                    .Cells(reportRow, 8).Value = result
                    .Cells(reportRow, 9).Value = i - cycleStart + 1
                    .Cells(reportRow, 10).Value = noteText

                    .Cells(reportRow, 2).NumberFormat = "dd.mm.yyyy hh:mm:ss"
                    .Cells(reportRow, 3).NumberFormat = "dd.mm.yyyy hh:mm:ss"
                    ' Приводим числовые даты к отображению дата+время
                    If IsNumeric(.Cells(reportRow, 2).Value) Then
                        .Cells(reportRow, 2).Value = CDate(.Cells(reportRow, 2).Value)
                    End If
                    If IsNumeric(.Cells(reportRow, 3).Value) Then
                        .Cells(reportRow, 3).Value = CDate(.Cells(reportRow, 3).Value)
                    End If
                    .Cells(reportRow, 7).NumberFormat = "0.0000"
                    .Cells(reportRow, 8).Font.Color = resultColor
                    .Cells(reportRow, 8).Font.Bold = True

                    If cycleNum Mod 2 = 0 Then
                        .Rows(reportRow).Interior.Color = RGB(20, 35, 55)
                    Else
                        .Rows(reportRow).Interior.Color = RGB(12, 25, 42)
                    End If
                End With

                reportRow = reportRow + 1
                inCycle = False
                peakReached = False
                tMax = -999
                tMin = 999
                f0Cycle = 0
            End If
        End If
NextRow:
    Next i

    Call AddSummaryRow(wsReport, reportRow, cycleNum)
End Sub

'-------------------------------------------------------------
' Итоговая строка
'-------------------------------------------------------------
Sub AddSummaryRow(wsReport As Worksheet, reportRow As Integer, totalCycles As Integer)
    Dim lastDataRow As Integer
    lastDataRow = reportRow - 1
    reportRow = reportRow + 1

    With wsReport
        .Cells(reportRow, 1).Value = "ИТОГО:"
        .Cells(reportRow, 1).Font.Bold = True
        .Cells(reportRow, 1).Font.Color = RGB(0, 180, 220)
        .Cells(reportRow, 4).Value = "=SUM(D6:D" & lastDataRow & ")"
        .Cells(reportRow, 4).Font.Bold = True
        .Cells(reportRow, 5).Value = "=MAX(E6:E" & lastDataRow & ")"
        .Cells(reportRow, 5).Font.Bold = True
        .Cells(reportRow, 7).Value = "=SUM(G6:G" & lastDataRow & ")"
        .Cells(reportRow, 7).NumberFormat = "0.0000"
        .Cells(reportRow, 7).Font.Bold = True
        .Cells(reportRow, 7).Font.Color = RGB(0, 220, 120)
        .Cells(reportRow, 8).Value = "Всего циклов: " & totalCycles
        .Cells(reportRow, 8).Font.Bold = True
        .Rows(reportRow).Interior.Color = RGB(10, 30, 55)
        .Rows(reportRow).Font.Color = RGB(200, 220, 240)
        .Cells(reportRow, 7).Font.Color = RGB(0, 220, 120)
        .Cells(reportRow, 1).Font.Color = RGB(0, 180, 220)
        .Rows(reportRow).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Rows(reportRow).Borders(xlEdgeTop).Color = RGB(0, 180, 220)
        .Rows(reportRow).Borders(xlEdgeTop).Weight = xlMedium
    End With
End Sub

'-------------------------------------------------------------
' Форматирование листа отчёта
'-------------------------------------------------------------
Sub FormatReportSheet(wsReport As Worksheet)
    With wsReport
        .Columns(1).ColumnWidth = 7
        .Columns(2).ColumnWidth = 20
        .Columns(3).ColumnWidth = 20
        .Columns(4).ColumnWidth = 12
        .Columns(5).ColumnWidth = 13
        .Columns(6).ColumnWidth = 13
        .Columns(7).ColumnWidth = 13
        .Columns(8).ColumnWidth = 26
        .Columns(9).ColumnWidth = 8
        .Columns(10).ColumnWidth = 24
        .Columns(1).HorizontalAlignment = xlCenter
        .Columns(4).HorizontalAlignment = xlCenter
        .Columns(7).HorizontalAlignment = xlCenter
        .Columns(9).HorizontalAlignment = xlCenter
        .Cells.Interior.Color = RGB(8, 20, 38)
        .Cells.Font.Color = RGB(180, 200, 220)
        .Rows(1).Interior.Color = RGB(5, 15, 30)
        .Rows(2).Interior.Color = RGB(5, 15, 30)
        .Rows(3).Interior.Color = RGB(5, 15, 30)
        .Rows(5).Interior.Color = RGB(15, 40, 65)
        .Cells(1, 1).Font.Color = RGB(0, 180, 220)
        .Cells(1, 1).Font.Size = 14
    End With
End Sub

'-------------------------------------------------------------
' График температуры и F0
'-------------------------------------------------------------
Sub BuildTemperatureChart(wb As Workbook, wsData As Worksheet, lastRow As Long)
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name = "График" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "График"

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=10, Top:=10, Width:=900, Height:=480)

    Dim cht As Chart
    Set cht = co.Chart

    ' Температура продукта = столбец 5 (E), F0 накопл. = столбец 18 (R)
    Dim rngTemp As Range
    Set rngTemp = wsData.Range(wsData.Cells(2, 5), wsData.Cells(lastRow, 5))
    Dim rngF0 As Range
    Set rngF0 = wsData.Range(wsData.Cells(2, 18), wsData.Cells(lastRow, 18))

    cht.ChartType = xlLine
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Dim s1 As Series
    Set s1 = cht.SeriesCollection.NewSeries
    s1.Name = "T продукта (C)"
    s1.Values = rngTemp
    s1.Format.Line.ForeColor.RGB = RGB(0, 180, 220)
    s1.Format.Line.Weight = 2
    s1.AxisGroup = xlPrimary

    Dim s2 As Series
    Set s2 = cht.SeriesCollection.NewSeries
    s2.Name = "F0 накопл. (мин)"
    s2.Values = rngF0
    s2.Format.Line.ForeColor.RGB = RGB(255, 165, 0)
    s2.Format.Line.Weight = 2.5
    s2.AxisGroup = xlSecondary

    Dim s3 As Series
    Set s3 = cht.SeriesCollection.NewSeries
    s3.Name = "Tref = 121.1C"
    Dim trefArr() As Double
    ReDim trefArr(1 To lastRow - 1)
    Dim k As Long
    For k = 1 To lastRow - 1
        trefArr(k) = T_REF
    Next k
    s3.Values = trefArr
    s3.Format.Line.ForeColor.RGB = RGB(200, 40, 40)
    s3.Format.Line.DashStyle = msoLineDash
    s3.Format.Line.Weight = 1.5
    s3.AxisGroup = xlPrimary

    With cht
        .HasTitle = True
        .ChartTitle.Text = "Температурный профиль автоклава и накопленный F0"
        .ChartTitle.Font.Size = 13
        .ChartTitle.Font.Bold = True
        .PlotArea.Interior.Color = RGB(12, 25, 42)
        .PlotArea.Border.LineStyle = xlNone
        .ChartArea.Interior.Color = RGB(8, 20, 38)
        .ChartArea.Border.LineStyle = xlNone

        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Температура (C)"
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
        End With

        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "F0 (мин)"
        End With

        .HasLegend = True
    End With
End Sub

Function SheetExistsInWb(wb As Workbook, sheetName As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Sheets(sheetName)
    SheetExistsInWb = Not s Is Nothing
    On Error GoTo 0
End Function`;

const MOCK_CYCLES = [
  { num: 1, tMax: 122.4, tMin: 18.2, dur: 87, f0: 12.847, status: "norm" },
  { num: 2, tMax: 121.8, tMin: 19.1, dur: 92, f0: 9.312, status: "norm" },
  { num: 3, tMax: 119.3, tMin: 18.7, dur: 78, f0: 2.184, status: "warn" },
  { num: 4, tMax: 122.1, tMin: 18.4, dur: 95, f0: 14.521, status: "norm" },
];

export default function Index() {
  const [copied, setCopied] = useState(false);

  const copyCode = () => {
    navigator.clipboard.writeText(VBA_CODE);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  return (
    <div className="min-h-screen bg-[#060d16] text-[#b0c4d4] font-['IBM_Plex_Sans',sans-serif]">
      {/* Header */}
      <header className="border-b border-[#142030] bg-[#040b14]/95 backdrop-blur sticky top-0 z-50">
        <div className="max-w-6xl mx-auto px-6 h-13 flex items-center justify-between py-3">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded bg-[#0ab4dc]/10 border border-[#0ab4dc]/25 flex items-center justify-center">
              <Icon name="Thermometer" size={15} className="text-[#0ab4dc]" />
            </div>
            <div>
              <div className="font-bold text-[#c8dce8] tracking-widest text-sm">
                АВТОКЛАВ F0
              </div>
              <div className="text-[9px] text-[#2a4060] tracking-widest uppercase font-['IBM_Plex_Mono']">
                Система расчёта стерилизации
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-1 rounded-sm border border-[#0f5040]/60 bg-[#061a14] text-[10px] text-[#22c55e] font-['IBM_Plex_Mono'] tracking-widest">
              Tref 121.1°C
            </span>
            <span className="px-2.5 py-1 rounded-sm border border-[#5a2a08]/60 bg-[#160a04] text-[10px] text-[#f97316] font-['IBM_Plex_Mono'] tracking-widest">
              z = 10°C
            </span>
            <span className="px-2.5 py-1 rounded-sm border border-[#1e4070]/60 bg-[#0a1830] text-[10px] text-[#60a0d0] font-['IBM_Plex_Mono'] tracking-widest">
              F0 ≥ 6 мин
            </span>
          </div>
        </div>
      </header>

      <div className="max-w-6xl mx-auto px-6 py-8 space-y-5">
        {/* Hero */}
        <div className="rounded border border-[#142030] bg-[#080f1c] p-8 relative overflow-hidden">
          <div className="absolute top-0 right-0 w-80 h-80 bg-[#0ab4dc]/3 rounded-full blur-3xl pointer-events-none" />
          <div className="absolute bottom-0 left-40 w-60 h-60 bg-[#f97316]/2 rounded-full blur-3xl pointer-events-none" />
          <div className="relative">
            <div className="flex items-center gap-3 mb-5">
              <div className="h-px w-8 bg-[#0ab4dc]/40" />
              <span className="text-[10px] text-[#2a5070] font-['IBM_Plex_Mono'] uppercase tracking-[0.25em]">
                Excel VBA Macro · ISO 11138
              </span>
            </div>
            <h1 className="text-4xl font-bold text-[#d0e8f4] mb-3 leading-tight">
              Расчёт <span className="text-[#0ab4dc]">F0</span> для автоклава
            </h1>
            <p className="text-[#4a6880] max-w-2xl leading-relaxed mb-6 text-sm">
              VBA-макрос с полным циклом: импорт CSV → определение циклов по
              температуре продукта (столбец E) → расчёт стерилизационного
              эффекта F0 методом трапеций → итоговый протокол по каждому циклу
              + температурный график.
            </p>
            <div className="flex flex-wrap gap-2">
              {[
                { label: "Парсинг CSV", color: "#0ab4dc" },
                { label: "Автоопределение циклов", color: "#22c55e" },
                { label: "Формула Биглоу", color: "#a78bfa" },
                { label: "Интеграция трапеций", color: "#f97316" },
                { label: "F0_Report", color: "#0ab4dc" },
                { label: "График T и F0", color: "#22c55e" },
              ].map((f) => (
                <span
                  key={f.label}
                  className="px-3 py-1 rounded-sm border border-[#142030] bg-[#0a1524] text-[11px]"
                  style={{ color: f.color + "cc" }}
                >
                  {f.label}
                </span>
              ))}
            </div>
          </div>
        </div>

        {/* Параметры + Формула */}
        <div className="grid md:grid-cols-2 gap-5">
          {/* Параметры */}
          <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
            <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
              <Icon name="Settings2" size={14} className="text-[#0ab4dc]" />
              <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
                Параметры стерилизации
              </span>
            </div>
            <div className="space-y-3">
              {[
                {
                  key: "T_REF",
                  val: "121.1°C",
                  desc: "Эталонная температура",
                  c: "#0ab4dc",
                },
                {
                  key: "Z_FACTOR",
                  val: "10°C",
                  desc: "Z-фактор",
                  c: "#f97316",
                },
                {
                  key: "T_START",
                  val: "30°C",
                  desc: "Порог начала цикла",
                  c: "#22c55e",
                },
                {
                  key: "T_MIN_STERIL",
                  val: "100°C",
                  desc: "Минимум для счёта F0",
                  c: "#a78bfa",
                },
                {
                  key: "F0_NORM",
                  val: "≥ 6 мин",
                  desc: "Норма стерилизации",
                  c: "#22c55e",
                },
                {
                  key: "COL_TEMP_PROD",
                  val: "Col E (5)",
                  desc: "Температура продукта в CSV",
                  c: "#0ab4dc",
                },
                {
                  key: "COL_F0_OUT",
                  val: "Col R (18)",
                  desc: "Запись накопленного F0",
                  c: "#f97316",
                },
              ].map((p) => (
                <div
                  key={p.key}
                  className="flex items-center justify-between py-1.5"
                >
                  <div className="flex items-center gap-2">
                    <div
                      className="w-1.5 h-1.5 rounded-full"
                      style={{ backgroundColor: p.c }}
                    />
                    <span className="text-[11px] font-['IBM_Plex_Mono'] text-[#3a5870]">
                      {p.key}
                    </span>
                    <span className="text-[11px] text-[#3a5870]">
                      — {p.desc}
                    </span>
                  </div>
                  <span
                    className="text-sm font-bold font-['IBM_Plex_Mono']"
                    style={{ color: p.c }}
                  >
                    {p.val}
                  </span>
                </div>
              ))}
            </div>
          </div>

          {/* Формула */}
          <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
            <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
              <Icon name="FlaskConical" size={14} className="text-[#a78bfa]" />
              <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
                Формула F0 (Биглоу)
              </span>
            </div>
            <div className="bg-[#04090f] border border-[#1a2d44] rounded p-4 font-['IBM_Plex_Mono'] text-sm mb-3 space-y-2">
              <div className="text-[10px] text-[#2a4060] uppercase tracking-widest">
                Летальность
              </div>
              <div className="text-[#c0d8f0]">
                L(T) = 10
                <sup className="text-[#0ab4dc] text-[10px]">
                  ((T − 121.1) / 10)
                </sup>
              </div>
              <div className="text-[10px] text-[#2a4060] uppercase tracking-widest mt-2">
                Стерилизационный эффект
              </div>
              <div className="text-[#c0d8f0]">
                F0 = Σ L(Tᵢ) × Δtᵢ{" "}
                <span className="text-[#2a4060]">[мин]</span>
              </div>
              <div className="text-[10px] text-[#2a4060] uppercase tracking-widest mt-2">
                Где
              </div>
              <div className="text-[11px] text-[#4a6880] space-y-0.5">
                <div>T — температура продукта (столбец E)</div>
                <div>Δt — интервал из столбца A (мин)</div>
              </div>
            </div>
            <div className="space-y-2">
              {[
                { label: "F0 ≥ 6", desc: "НОРМА", color: "#22c55e" },
                { label: "3 ≤ F0 < 6", desc: "НА ПРЕДЕЛЕ", color: "#f97316" },
                { label: "F0 < 3", desc: "НЕДОСТАТОЧНО", color: "#ef4444" },
              ].map((r) => (
                <div
                  key={r.label}
                  className="flex items-center justify-between px-3 py-1.5 rounded-sm"
                  style={{
                    backgroundColor: r.color + "10",
                    borderLeft: `2px solid ${r.color}60`,
                  }}
                >
                  <span
                    className="text-xs font-['IBM_Plex_Mono'] font-bold"
                    style={{ color: r.color }}
                  >
                    {r.label}
                  </span>
                  <span
                    className="text-[10px] font-semibold tracking-wider"
                    style={{ color: r.color + "aa" }}
                  >
                    {r.desc}
                  </span>
                </div>
              ))}
            </div>
          </div>
        </div>

        {/* Алгоритм циклов */}
        <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
            <Icon name="GitBranch" size={14} className="text-[#22c55e]" />
            <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
              Алгоритм определения циклов
            </span>
          </div>
          <div className="grid md:grid-cols-4 gap-3">
            {[
              {
                step: "01",
                label: "Простой",
                desc: "T = 18–30°C, автоклав без загрузки — игнорируется",
                color: "#3a5870",
                icon: "Minus",
              },
              {
                step: "02",
                label: "Старт цикла",
                desc: "T поднимается выше 30°C и продолжает расти",
                color: "#22c55e",
                icon: "TrendingUp",
              },
              {
                step: "03",
                label: "Стерилизация",
                desc: "T ≥ 100°C — идёт расчёт F0, накопление летальности",
                color: "#0ab4dc",
                icon: "Zap",
              },
              {
                step: "04",
                label: "Конец цикла",
                desc: "После пика T < 30°C — цикл закрывается, F0 фиксируется",
                color: "#f97316",
                icon: "TrendingDown",
              },
            ].map((s) => (
              <div
                key={s.step}
                className="p-4 rounded-sm border bg-[#04090f]"
                style={{ borderColor: s.color + "25" }}
              >
                <div className="flex items-center gap-2 mb-2">
                  <span
                    className="text-[10px] font-['IBM_Plex_Mono']"
                    style={{ color: s.color + "80" }}
                  >
                    {s.step}
                  </span>
                  <Icon name={s.icon} size={13} style={{ color: s.color }} />
                  <span
                    className="text-xs font-semibold"
                    style={{ color: s.color }}
                  >
                    {s.label}
                  </span>
                </div>
                <p className="text-[11px] text-[#3a5870] leading-relaxed">
                  {s.desc}
                </p>
              </div>
            ))}
          </div>
          <div className="mt-3 px-3 py-2 bg-[#04090f] border border-[#142030] rounded-sm text-[11px] text-[#3a5870]">
            Несколько циклов за день корректно разделяются. Каждый цикл
            получает отдельную строку в F0_Report с временными метками
            (начало/конец) и всеми параметрами.
          </div>
        </div>

        {/* Демо таблица */}
        <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
            <Icon name="Table" size={14} className="text-[#22c55e]" />
            <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
              Пример F0_Report
            </span>
            <span className="ml-2 text-[10px] text-[#1e3050] bg-[#0a1828] px-2 py-0.5 rounded-sm border border-[#1e3050]">
              демо-данные
            </span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs font-['IBM_Plex_Mono']">
              <thead>
                <tr className="border-b-2 border-[#0ab4dc]/30">
                  {[
                    "Цикл",
                    "T макс.",
                    "T мин.",
                    "Длит.",
                    "F0 (мин)",
                    "Результат",
                  ].map((h) => (
                    <th
                      key={h}
                      className="text-left py-2 px-3 text-[#0ab4dc] font-semibold text-[10px] uppercase tracking-widest"
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {MOCK_CYCLES.map((c, idx) => (
                  <tr
                    key={c.num}
                    className={`border-b border-[#0f1e2e] ${idx % 2 === 0 ? "bg-[#060d18]" : "bg-[#080f1c]"}`}
                  >
                    <td className="py-2.5 px-3 text-[#3a5870] font-bold">
                      #{c.num}
                    </td>
                    <td className="py-2.5 px-3 text-[#0ab4dc]">{c.tMax}°C</td>
                    <td className="py-2.5 px-3 text-[#2a4860]">{c.tMin}°C</td>
                    <td className="py-2.5 px-3 text-[#4a6880]">{c.dur} мин</td>
                    <td
                      className={`py-2.5 px-3 font-bold ${c.status === "norm" ? "text-[#22c55e]" : "text-[#f97316]"}`}
                    >
                      {c.f0.toFixed(4)}
                    </td>
                    <td className="py-2.5 px-3">
                      {c.status === "norm" ? (
                        <span className="text-[#22c55e] bg-[#06200e] px-2 py-0.5 rounded-sm border border-[#0f5030]/40">
                          ✓ НОРМА
                        </span>
                      ) : (
                        <span className="text-[#f97316] bg-[#1c1004] px-2 py-0.5 rounded-sm border border-[#603010]/40">
                          ⚠ ПРЕДЕЛ
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
                <tr className="border-t-2 border-[#0ab4dc]/40 bg-[#030710]">
                  <td className="py-2.5 px-3 text-[#0ab4dc] font-bold">
                    ИТОГО
                  </td>
                  <td className="py-2.5 px-3 text-[#0ab4dc] font-bold">
                    122.4°C
                  </td>
                  <td className="py-2.5 px-3" />
                  <td className="py-2.5 px-3 text-[#c0d8e8] font-bold">
                    352 мин
                  </td>
                  <td className="py-2.5 px-3 text-[#22c55e] font-bold">
                    38.8640
                  </td>
                  <td className="py-2.5 px-3 text-[#3a5870]">Циклов: 4</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Маппинг столбцов CSV */}
        <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
            <Icon name="Columns" size={14} className="text-[#a78bfa]" />
            <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
              Структура CSV — маппинг столбцов
            </span>
          </div>
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2">
            {[
              { col: "A (1)", name: "ДАТА", note: "Дата записи", key: true },
              { col: "B (2)", name: "ВРЕМЯ", note: "Время записи", key: true },
              { col: "C (3)", name: "МИЛЛИСЕКУНДЫ", note: "", key: false },
              { col: "D (4)", name: "ТЕМП. СРЕДЫ", note: "", key: false },
              { col: "E (5)", name: "ТЕМП. ПРОДУКТА", note: "← для расчёта F0", key: true },
              { col: "F (6)", name: "ДАВЛЕНИЕ", note: "", key: false },
              { col: "G (7)", name: "УРОВЕНЬ ВОДЫ", note: "", key: false },
              { col: "H (8)", name: "Q recirculación", note: "", key: false },
              { col: "I (9)", name: "ДАВЛ. ВОЗДУХА", note: "", key: false },
              { col: "J (10)", name: "SP presión", note: "", key: false },
              { col: "K (11)", name: "ЗАД. ТЕМП.", note: "", key: false },
              { col: "L (12)", name: "КЛ. ДАВЛ. +", note: "", key: false },
              { col: "M (13)", name: "КЛ. ДАВЛ. −", note: "", key: false },
              { col: "N (14)", name: "КЛ. НАГРЕВА", note: "", key: false },
              { col: "O (15)", name: "КЛ. ОХЛАЖД.", note: "", key: false },
              { col: "P (16)", name: "КЛ. ЗАПОЛН.", note: "", key: false },
              { col: "Q (17)", name: "НОМ. ЗАМЕСА", note: "", key: false },
              { col: "R (18)", name: "F0_накоп.", note: "← добавляет макрос", key: true },
            ].map((s) => (
              <div
                key={s.col}
                className={`flex items-start gap-2 p-2.5 rounded-sm border ${
                  s.key
                    ? "border-[#0ab4dc]/30 bg-[#04121e]"
                    : "border-[#142030] bg-[#04090f]"
                }`}
              >
                <span className={`text-[10px] font-['IBM_Plex_Mono'] font-bold shrink-0 ${s.key ? "text-[#0ab4dc]" : "text-[#2a4060]"}`}>
                  {s.col}
                </span>
                <div>
                  <div className={`text-[10px] font-semibold leading-tight ${s.key ? "text-[#80b8d0]" : "text-[#3a5870]"}`}>
                    {s.name}
                  </div>
                  {s.note && (
                    <div className="text-[9px] text-[#0ab4dc]/60 mt-0.5">{s.note}</div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Шаги */}
        <div className="rounded border border-[#142030] bg-[#080f1c] p-5">
          <div className="flex items-center gap-2 mb-4 pb-3 border-b border-[#142030]">
            <Icon name="ListOrdered" size={14} className="text-[#f97316]" />
            <span className="text-xs font-semibold text-[#a0c0d4] uppercase tracking-wider">
              Установка и запуск
            </span>
          </div>
          <div className="grid md:grid-cols-5 gap-3">
            {[
              {
                n: "1",
                t: "Открыть Excel",
                d: "Открываете рабочую книгу (.xlsx или .xlsm)",
              },
              {
                n: "2",
                t: "Alt + F11",
                d: "Открываете редактор VBA (Visual Basic for Applications)",
              },
              {
                n: "3",
                t: "Вставить модуль",
                d: "Insert → Module → вставляете скопированный код",
              },
              {
                n: "4",
                t: "Запуск макроса",
                d: "F5 или вкладка Разработчик → Макросы → Autoclave_ProcessCSV",
              },
              {
                n: "5",
                t: "Результат",
                d: "Выбираете CSV → получаете Data, F0_Report и График",
              },
            ].map((s) => (
              <div
                key={s.n}
                className="p-3 rounded-sm border border-[#142030] bg-[#04090f]"
              >
                <div className="text-2xl font-bold text-[#1a3050] font-['IBM_Plex_Mono'] mb-1.5">
                  {s.n}
                </div>
                <div className="text-[11px] font-semibold text-[#6a90a8] mb-1">
                  {s.t}
                </div>
                <div className="text-[10px] text-[#2a4060] leading-relaxed">
                  {s.d}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* VBA Код */}
        <div className="rounded border border-[#142030] bg-[#080f1c] overflow-hidden">
          <div className="flex items-center justify-between px-5 py-3.5 border-b border-[#142030] bg-[#030710]">
            <div className="flex items-center gap-3">
              <Icon name="Code2" size={14} className="text-[#0ab4dc]" />
              <span className="text-sm font-semibold text-[#a0c0d4]">
                Module1 — Полный VBA-код
              </span>
              <span className="text-[10px] text-[#2a4060] bg-[#0a1828] px-2 py-0.5 rounded-sm border border-[#1e3050] font-['IBM_Plex_Mono']">
                Excel VBA
              </span>
            </div>
            <button
              onClick={copyCode}
              className={`flex items-center gap-1.5 px-4 py-2 rounded-sm text-xs font-medium transition-all border ${
                copied
                  ? "bg-[#061a0e] border-[#22c55e]/40 text-[#22c55e]"
                  : "bg-[#0a1828] border-[#0ab4dc]/25 text-[#0ab4dc] hover:bg-[#0c1e34] hover:border-[#0ab4dc]/40"
              }`}
            >
              <Icon name={copied ? "Check" : "Copy"} size={12} />
              {copied ? "Скопировано!" : "Копировать код"}
            </button>
          </div>
          <textarea
            readOnly
            value={VBA_CODE}
            className="w-full h-[480px] bg-[#03060e] text-[#6a9ab8] font-['IBM_Plex_Mono'] text-[11px] leading-relaxed p-5 resize-none outline-none border-0"
            style={{ scrollbarColor: "#142030 transparent" }}
          />
        </div>

        {/* Настройка */}
        <div className="rounded border border-[#3a2010]/60 bg-[#0e0804] p-5">
          <div className="flex items-start gap-3">
            <Icon
              name="AlertTriangle"
              size={16}
              className="text-[#f97316] shrink-0 mt-0.5"
            />
            <div>
              <h3 className="text-sm font-semibold text-[#d48040] mb-3">
                Настройка под ваш CSV-файл
              </h3>
              <div className="grid md:grid-cols-2 gap-3 text-[11px]">
                {[
                  {
                    key: "COL_TEMP_PROD = 5",
                    note: "Номер столбца с температурой продукта (E = 5). Если у вас другой — меняйте.",
                  },
                  {
                    key: "T_START = 30",
                    note: "Порог начала цикла. Если автоклав простаивает при другой температуре — укажите свою.",
                  },
                  {
                    key: "Разделитель CSV",
                    note: 'Определяется автоматически: ";" или ",". Ручная правка не нужна.',
                  },
                  {
                    key: 'Значения в "кавычках"',
                    note: 'Автоматически очищаются. Формат "2026/05/20" и "14.91910" — обрабатываются корректно.',
                  },
                  {
                    key: "Дата YYYY/MM/DD",
                    note: "Нестандартный формат автоклава конвертируется через DateSerial — Excel принимает как дату.",
                  },
                  {
                    key: "Δt из миллисекунд",
                    note: "Если интервал по времени = 0 (записи за 1 сек) — используется столбец C (миллисекунды).",
                  },
                ].map((tip) => (
                  <div
                    key={tip.key}
                    className="flex items-start gap-2 p-2.5 rounded-sm bg-[#08060e] border border-[#2a1808]"
                  >
                    <code className="text-[#f97316] bg-[#1e0e04] px-1.5 py-0.5 rounded text-[10px] shrink-0 font-['IBM_Plex_Mono']">
                      {tip.key}
                    </code>
                    <span className="text-[#5a3820] leading-relaxed">
                      {tip.note}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="text-center py-4 text-[#1e3050] text-[10px] font-['IBM_Plex_Mono'] tracking-widest">
          AUTOCLAVE F0 CALCULATOR · Tref=121.1°C · z=10°C · БИГЛОУ · VBA EXCEL
        </div>
      </div>
    </div>
  );
}