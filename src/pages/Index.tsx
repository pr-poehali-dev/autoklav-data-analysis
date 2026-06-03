import { useState, useRef } from "react";
import Icon from "@/components/ui/icon";

const VBA_CODE = `'=============================================================
' АВТОКЛАВ — Расчёт стерилизационного эффекта F0
' Параметры: Tref = 121.1 C, z = 10 C
' Алгоритм: автоопределение циклов + расчёт F0 по формуле
'=============================================================

Option Explicit

' Глобальные константы стерилизации (режим СТЕРИЛИЗАЦИИ — тушёнка по ГОСТ)
' Tref=121.1°C, z=10 — стандарт для Clostridium botulinum
' Стерилизующий эффект считается по датчику в центре продукта (столбец E)
Const T_REF As Double = 121.1   ' Эталонная температура (C)
Const Z_FACTOR As Double = 10#   ' Z-фактор (C)
Const T_START As Double = 30#    ' Порог начала цикла (C)
' Порог накопления F0. По данным автоклава фактор стерилизации начинает
' накапливаться примерно с 90°C (вклад < 100°C мал, но даёт плавную кривую).
Const T_MIN_STERIL As Double = 90#   ' Минимум для счёта F0
' Порог "пика" — цикл считается стерилизационным если достиг этой T
Const T_PEAK_STERIL As Double = 100# ' Подтверждение стерилизации

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
        Title:="Выберите ОСНОВНОЙ CSV-файл с данными автоклава")
    If filePath = "False" Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wb = ActiveWorkbook

    ' Извлекаем имя файла без пути
    Dim csvFileName As String
    Dim csvFolder As String
    csvFileName = Mid(filePath, InStrRev(filePath, "\\") + 1)
    csvFolder   = Left(filePath, InStrRev(filePath, "\\"))

    ' ----------------------------------------------------------------
    ' Импортируем основной файл
    ' ----------------------------------------------------------------
    Call ImportAndParseCSV(wb, filePath, wsData)

    ' ----------------------------------------------------------------
    ' ПРОВЕРКА 1: замес начался в ПРЕДЫДУЩИХ сутках?
    ' (файл начинается с T > 40°C при времени ~00:00)
    ' ----------------------------------------------------------------
    Dim prevFilePath As String
    prevFilePath = ""

    If NeedsPreviousFile(wsData) Then
        Application.ScreenUpdating = True

        ' Подсказываем имя предыдущего файла (дата −1 день)
        Dim suggestPrev As String
        suggestPrev = GetNeighborFileName(csvFileName, -1)
        Dim hintPrev As String
        If suggestPrev <> "" Then
            hintPrev = Chr(13) & Chr(13) & "РЕКОМЕНДУЕМЫЙ ФАЙЛ:  " & suggestPrev
            If Len(Dir(csvFolder & suggestPrev)) > 0 Then
                hintPrev = hintPrev & Chr(13) & "(найден в этой папке)"
            Else
                hintPrev = hintPrev & Chr(13) & "(в папке не найден — выберите вручную)"
            End If
        End If

        Dim ans As Integer
        ans = MsgBox("Данные начинаются при температуре продукта > 40°C в начале суток." & Chr(13) & _
                     "Похоже, замес начался в ПРЕДЫДУЩЕМ файле." & hintPrev & Chr(13) & Chr(13) & _
                     "Загрузить предыдущий CSV-файл для корректного расчёта F0 и времени цикла?", _
                     vbYesNo + vbQuestion, "Переход суток — нужен ПРЕДЫДУЩИЙ файл")

        If ans = vbYes Then
            ' Если рекомендуемый файл найден — предлагаем сразу его
            Dim defPrevPath As String
            defPrevPath = ""
            If suggestPrev <> "" And Len(Dir(csvFolder & suggestPrev)) > 0 Then
                defPrevPath = csvFolder & suggestPrev
            End If

            prevFilePath = Application.GetOpenFilename( _
                FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
                Title:="Выберите ПРЕДЫДУЩИЙ файл: " & suggestPrev)
            If prevFilePath = "False" Then prevFilePath = ""
        End If
        Application.ScreenUpdating = False
    End If

    ' ----------------------------------------------------------------
    ' ПРОВЕРКА 2: замес уходит в СЛЕДУЮЩИЕ сутки?
    ' (файл заканчивается T > 40°C при времени ~23:59)
    ' ----------------------------------------------------------------
    Dim nextFilePath As String
    nextFilePath = ""

    If NeedsNextFile(wsData) Then
        Application.ScreenUpdating = True

        ' Подсказываем имя следующего файла (дата +1 день)
        Dim suggestNext As String
        suggestNext = GetNeighborFileName(csvFileName, 1)
        Dim hintNext As String
        If suggestNext <> "" Then
            hintNext = Chr(13) & Chr(13) & "РЕКОМЕНДУЕМЫЙ ФАЙЛ:  " & suggestNext
            If Len(Dir(csvFolder & suggestNext)) > 0 Then
                hintNext = hintNext & Chr(13) & "(найден в этой папке)"
            Else
                hintNext = hintNext & Chr(13) & "(в папке не найден — выберите вручную)"
            End If
        End If

        Dim ansN As Integer
        ansN = MsgBox("Данные заканчиваются при температуре продукта > 40°C в конце суток." & Chr(13) & _
                     "Похоже, замес продолжается в СЛЕДУЮЩЕМ файле." & hintNext & Chr(13) & Chr(13) & _
                     "Загрузить следующий CSV-файл для корректного расчёта F0?", _
                     vbYesNo + vbQuestion, "Переход суток — нужен СЛЕДУЮЩИЙ файл")

        If ansN = vbYes Then
            nextFilePath = Application.GetOpenFilename( _
                FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
                Title:="Выберите СЛЕДУЮЩИЙ файл: " & suggestNext)
            If nextFilePath = "False" Then nextFilePath = ""
        End If
        Application.ScreenUpdating = False
    End If

    ' ----------------------------------------------------------------
    ' Если выбран предыдущий файл — вставляем его данные ПЕРЕД основными
    ' ----------------------------------------------------------------
    If prevFilePath <> "" Then
        Call PrependPreviousCSV(wb, wsData, prevFilePath)
        Dim prevName As String
        prevName = Mid(prevFilePath, InStrRev(prevFilePath, "\\") + 1)
        csvFileName = prevName & " + " & csvFileName
    End If

    ' ----------------------------------------------------------------
    ' Если выбран следующий файл — дописываем его данные ПОСЛЕ основных
    ' ----------------------------------------------------------------
    If nextFilePath <> "" Then
        Call AppendNextCSV(wb, wsData, nextFilePath)
        Dim nextName As String
        nextName = Mid(nextFilePath, InStrRev(nextFilePath, "\\") + 1)
        csvFileName = csvFileName & " + " & nextName
    End If

    Call PrepareReportSheet(wb, wsReport, csvFileName)

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Call DetectCyclesAndCalculateF0(wsData, wsReport, lastRow)
    Call FormatReportSheet(wsReport)
    Call BuildTemperatureChart(wb, wsData, lastRow, csvFileName)

    ' ----------------------------------------------------------------
    ' Сохраняем только нужные листы (Data, F0_Report, График) в новый .xlsm
    ' ----------------------------------------------------------------
    Dim saveName As String
    saveName = csvFileName
    If LCase(Right(saveName, 4)) = ".csv" Then saveName = Left(saveName, Len(saveName) - 4)
    saveName = Replace(Replace(saveName, "+", "_"), " ", "_")

    Dim savePath As String
    savePath = csvFolder & saveName & ".xlsm"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Копируем три листа в новую книгу
    Dim savedMsg As String
    On Error Resume Next
    Dim wbNew As Workbook

    ' Собираем листы для копирования
    Dim sheetsToCopy(2) As String
    sheetsToCopy(0) = "Data"
    sheetsToCopy(1) = "F0_Report"
    sheetsToCopy(2) = "График"

    ' Проверяем что все листы существуют
    Dim allExist As Boolean
    allExist = True
    Dim si As Integer
    For si = 0 To 2
        If Not SheetExistsInWb(wb, sheetsToCopy(si)) Then allExist = False
    Next si

    If allExist Then
        ' Копируем все три листа вместе в новую книгу
        wb.Sheets(Array("Data", "F0_Report", "График")).Copy
        Set wbNew = ActiveWorkbook

        ' Сохраняем новую книгу как xlsm
        Application.DisplayAlerts = False
        wbNew.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        wbNew.Close SaveChanges:=False

        If Len(Dir(savePath)) > 0 Then
            savedMsg = Chr(13) & Chr(13) & "Файл сохранён:" & Chr(13) & savePath
        Else
            savedMsg = Chr(13) & Chr(13) & "Не удалось сохранить — проверьте права на папку."
        End If
    Else
        savedMsg = Chr(13) & Chr(13) & "Не все листы созданы — сохранение пропущено."
    End If
    On Error GoTo 0

    MsgBox "Расчёт завершён! Результаты на листе 'F0_Report'." & savedMsg, _
           vbInformation, "Автоклав F0 — Готово"
End Sub

'-------------------------------------------------------------
' Проверяет: начинается ли файл с "горячей" температуры в 00:00
' Признак перехода суток — замес начался вчера
'-------------------------------------------------------------
Function NeedsPreviousFile(wsData As Worksheet) As Boolean
    NeedsPreviousFile = False
    Dim lastR As Long
    lastR = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Exit Function

    ' Смотрим первые 10 строк данных
    Dim r As Long
    For r = 2 To IIf(lastR < 11, lastR, 11)
        Dim tVal As Variant
        tVal = wsData.Cells(r, 5).Value  ' столбец E — температура продукта
        If Not IsNumeric(tVal) Then GoTo NextCheck

        Dim tempCheck As Double
        tempCheck = CDbl(tVal)

        ' Берём время из столбца B
        Dim timeRaw As Variant
        timeRaw = wsData.Cells(r, 2).Value
        Dim timeDbl As Double
        timeDbl = 0

        If IsNumeric(timeRaw) Then
            timeDbl = CDbl(timeRaw)
        ElseIf InStr(CStr(timeRaw), ":") > 0 Then
            On Error Resume Next
            timeDbl = CDbl(TimeValue(Trim(CStr(timeRaw))))
            On Error GoTo 0
        End If

        ' Время до 5 минут от полуночи (< 0.0035 в дробях Excel = ~5 мин)
        If tempCheck > 40 And timeDbl < 0.0035 Then
            NeedsPreviousFile = True
            Exit Function
        End If
NextCheck:
    Next r
End Function

'-------------------------------------------------------------
' Проверяет: заканчивается ли файл "горячей" температурой в ~23:59
' Признак того, что замес продолжается в СЛЕДУЮЩЕМ файле (следующие сутки)
'-------------------------------------------------------------
Function NeedsNextFile(wsData As Worksheet) As Boolean
    NeedsNextFile = False
    Dim lastR As Long
    lastR = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Exit Function

    ' Смотрим последние 10 строк данных
    Dim r As Long
    Dim startR As Long
    startR = lastR - 10
    If startR < 2 Then startR = 2

    For r = lastR To startR Step -1
        Dim tVal As Variant
        tVal = wsData.Cells(r, 5).Value  ' столбец E — температура продукта
        If Not IsNumeric(tVal) Then GoTo NextCheck2

        Dim tempCheck As Double
        tempCheck = CDbl(tVal)

        Dim timeRaw As Variant
        timeRaw = wsData.Cells(r, 2).Value
        Dim timeDbl As Double
        timeDbl = 0
        If IsNumeric(timeRaw) Then
            timeDbl = CDbl(timeRaw)
        ElseIf InStr(CStr(timeRaw), ":") > 0 Then
            On Error Resume Next
            timeDbl = CDbl(TimeValue(Trim(CStr(timeRaw))))
            On Error GoTo 0
        End If

        ' Время после 23:55 (> 0.9965 в дробях Excel) и температура горячая
        If tempCheck > 40 And timeDbl > 0.9965 Then
            NeedsNextFile = True
            Exit Function
        End If
NextCheck2:
    Next r
End Function

'-------------------------------------------------------------
' Вычисляет имя соседнего файла по дате (±1 день).
' Имена файлов в формате YYYYMMDD.csv (напр. 20260520.csv).
' offsetDays = -1 для предыдущего, +1 для следующего.
' Возвращает имя файла или "" если не удалось распознать.
'-------------------------------------------------------------
Function GetNeighborFileName(currentFileName As String, offsetDays As Integer) As String
    GetNeighborFileName = ""

    ' Отделяем имя без расширения
    Dim baseName As String
    Dim ext As String
    Dim dotPos As Integer
    dotPos = InStrRev(currentFileName, ".")
    If dotPos > 0 Then
        baseName = Left(currentFileName, dotPos - 1)
        ext = Mid(currentFileName, dotPos)  ' включая точку
    Else
        baseName = currentFileName
        ext = ".csv"
    End If

    ' Ищем в имени 8 цифр подряд = дата YYYYMMDD
    Dim i As Integer
    Dim digits As String
    digits = ""
    Dim datePos As Integer
    datePos = 0
    For i = 1 To Len(baseName)
        Dim ch As String
        ch = Mid(baseName, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
            If Len(digits) = 8 Then
                datePos = i - 7
                Exit For
            End If
        Else
            digits = ""
        End If
    Next i

    If Len(digits) <> 8 Then Exit Function  ' дата не найдена

    ' Парсим YYYYMMDD
    Dim yr As Integer, mo As Integer, dy As Integer
    yr = CInt(Left(digits, 4))
    mo = CInt(Mid(digits, 5, 2))
    dy = CInt(Mid(digits, 7, 2))

    On Error GoTo BadDate
    Dim baseDate As Date
    baseDate = DateSerial(yr, mo, dy)
    Dim newDate As Date
    newDate = baseDate + offsetDays

    ' Формируем новую дату YYYYMMDD
    Dim newDigits As String
    newDigits = Format(Year(newDate), "0000") & _
                Format(Month(newDate), "00") & _
                Format(Day(newDate), "00")

    ' Подставляем обратно в имя
    Dim prefix As String, suffix As String
    prefix = Left(baseName, datePos - 1)
    suffix = Mid(baseName, datePos + 8)

    GetNeighborFileName = prefix & newDigits & suffix & ext
    Exit Function
BadDate:
    GetNeighborFileName = ""
End Function

'-------------------------------------------------------------
' Вставляет данные предыдущего CSV В НАЧАЛО листа Data (перед основными)
' Дата предыдущего файла остаётся как есть — она на сутки раньше
'-------------------------------------------------------------
Sub PrependPreviousCSV(wb As Workbook, wsData As Worksheet, prevFilePath As String)
    ' Читаем предыдущий CSV во временный массив
    Dim fileNum As Integer
    Dim lineText As String
    Dim fields() As String
    Dim prevRows() As String
    Dim prevCount As Long
    prevCount = 0

    fileNum = FreeFile
    Open prevFilePath For Input As #fileNum
    ReDim prevRows(1 To 50000)

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) = 0 Then GoTo SkipLine
        If Left(lineText, 1) = "#" Then GoTo SkipLine
        prevCount = prevCount + 1
        prevRows(prevCount) = lineText
SkipLine:
    Loop
    Close #fileNum

    If prevCount = 0 Then Exit Sub

    ' Сдвигаем существующие данные вниз (кроме строки заголовков — строка 1)
    Dim existingLastRow As Long
    existingLastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' Вставляем пустые строки после заголовка для данных предыдущего файла
    ' Первая строка prevRows может быть заголовком CSV — пропускаем
    Dim startIdx As Long
    startIdx = 1
    Dim firstFields() As String
    If InStr(prevRows(1), ";") > 0 Then
        firstFields = Split(prevRows(1), ";")
    Else
        firstFields = Split(prevRows(1), ",")
    End If
    ' Чистим кавычки первого поля для проверки
    Dim fp As String
    fp = Trim(firstFields(0))
    If Len(fp) >= 2 And Left(fp, 1) = Chr(34) Then fp = Mid(fp, 2, Len(fp) - 2)
    ' Если первая строка — заголовок (не дата), пропускаем её
    If Not (fp Like "####/##/##") And Not IsDate(fp) Then startIdx = 2

    Dim insertRows As Long
    insertRows = prevCount - startIdx + 1
    If insertRows <= 0 Then Exit Sub

    wsData.Rows("2:" & (insertRows + 1)).Insert Shift:=xlDown

    ' Записываем строки предыдущего файла начиная со строки 2
    Dim writeRow As Long
    writeRow = 2
    Dim pi As Long
    Dim pFields() As String
    Dim pk As Integer

    For pi = startIdx To prevCount
        lineText = prevRows(pi)
        If InStr(lineText, ";") > 0 Then
            pFields = Split(lineText, ";")
        Else
            pFields = Split(lineText, ",")
        End If

        ' Очищаем кавычки
        For pk = 0 To UBound(pFields)
            pFields(pk) = Trim(pFields(pk))
            If Len(pFields(pk)) >= 2 Then
                If Left(pFields(pk), 1) = Chr(34) And Right(pFields(pk), 1) = Chr(34) Then
                    pFields(pk) = Mid(pFields(pk), 2, Len(pFields(pk)) - 2)
                End If
            End If
        Next pk

        Dim pTotalCols As Integer
        pTotalCols = UBound(pFields) + 1
        If pTotalCols > 17 Then pTotalCols = 17

        Dim pi2 As Integer
        For pi2 = 0 To pTotalCols - 1
            Dim pCell As String
            pCell = Trim(pFields(pi2))

            Select Case pi2
                Case 0 ' Дата YYYY/MM/DD
                    If InStr(pCell, "/") > 0 Then
                        Dim dp() As String
                        dp = Split(pCell, "/")
                        If UBound(dp) = 2 Then
                            wsData.Cells(writeRow, 1).Value = DateSerial(CInt(dp(0)), CInt(dp(1)), CInt(dp(2)))
                        Else
                            wsData.Cells(writeRow, 1).Value = pCell
                        End If
                    ElseIf IsDate(pCell) Then
                        wsData.Cells(writeRow, 1).Value = CDate(pCell)
                    Else
                        wsData.Cells(writeRow, 1).Value = pCell
                    End If

                Case 1 ' Время HH:MM:SS
                    If InStr(pCell, ":") > 0 Then
                        On Error Resume Next
                        wsData.Cells(writeRow, 2).Value = TimeValue(pCell)
                        On Error GoTo 0
                    Else
                        wsData.Cells(writeRow, 2).Value = pCell
                    End If

                Case Else
                    pCell = Replace(pCell, ",", ".")
                    If IsNumeric(pCell) Then
                        wsData.Cells(writeRow, pi2 + 1).Value = CDbl(pCell)
                    Else
                        wsData.Cells(writeRow, pi2 + 1).Value = pCell
                    End If
            End Select
        Next pi2

        writeRow = writeRow + 1
    Next pi

    ' Форматируем добавленные строки
    wsData.Columns(1).NumberFormat = "dd.mm.yyyy"
    wsData.Columns(2).NumberFormat = "hh:mm:ss"
End Sub

'-------------------------------------------------------------
' Дописывает данные следующего CSV В КОНЕЦ листа Data (после основных)
' Дата следующего файла на сутки позже — замес продолжается в нём
'-------------------------------------------------------------
Sub AppendNextCSV(wb As Workbook, wsData As Worksheet, nextFilePath As String)
    Dim fileNum As Integer
    Dim lineText As String
    Dim nextRows() As String
    Dim nextCount As Long
    nextCount = 0

    fileNum = FreeFile
    Open nextFilePath For Input As #fileNum
    ReDim nextRows(1 To 50000)

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) = 0 Then GoTo SkipLine2
        If Left(lineText, 1) = "#" Then GoTo SkipLine2
        nextCount = nextCount + 1
        nextRows(nextCount) = lineText
SkipLine2:
    Loop
    Close #fileNum

    If nextCount = 0 Then Exit Sub

    ' Определяем с какой строки начинать (пропускаем заголовок CSV)
    Dim startIdx As Long
    startIdx = 1
    Dim firstFields() As String
    If InStr(nextRows(1), ";") > 0 Then
        firstFields = Split(nextRows(1), ";")
    Else
        firstFields = Split(nextRows(1), ",")
    End If
    Dim fp As String
    fp = Trim(firstFields(0))
    If Len(fp) >= 2 And Left(fp, 1) = Chr(34) Then fp = Mid(fp, 2, Len(fp) - 2)
    If Not (fp Like "####/##/##") And Not IsDate(fp) Then startIdx = 2

    ' Пишем данные в конец листа Data
    Dim writeRow As Long
    writeRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1

    Dim pi As Long
    Dim pFields() As String
    Dim pk As Integer

    For pi = startIdx To nextCount
        lineText = nextRows(pi)
        If InStr(lineText, ";") > 0 Then
            pFields = Split(lineText, ";")
        Else
            pFields = Split(lineText, ",")
        End If

        For pk = 0 To UBound(pFields)
            pFields(pk) = Trim(pFields(pk))
            If Len(pFields(pk)) >= 2 Then
                If Left(pFields(pk), 1) = Chr(34) And Right(pFields(pk), 1) = Chr(34) Then
                    pFields(pk) = Mid(pFields(pk), 2, Len(pFields(pk)) - 2)
                End If
            End If
        Next pk

        Dim pTotalCols As Integer
        pTotalCols = UBound(pFields) + 1
        If pTotalCols > 17 Then pTotalCols = 17

        Dim pi2 As Integer
        For pi2 = 0 To pTotalCols - 1
            Dim pCell As String
            pCell = Trim(pFields(pi2))

            Select Case pi2
                Case 0 ' Дата YYYY/MM/DD
                    If InStr(pCell, "/") > 0 Then
                        Dim dp() As String
                        dp = Split(pCell, "/")
                        If UBound(dp) = 2 Then
                            wsData.Cells(writeRow, 1).Value = DateSerial(CInt(dp(0)), CInt(dp(1)), CInt(dp(2)))
                        Else
                            wsData.Cells(writeRow, 1).Value = pCell
                        End If
                    ElseIf IsDate(pCell) Then
                        wsData.Cells(writeRow, 1).Value = CDate(pCell)
                    Else
                        wsData.Cells(writeRow, 1).Value = pCell
                    End If

                Case 1 ' Время HH:MM:SS
                    If InStr(pCell, ":") > 0 Then
                        On Error Resume Next
                        wsData.Cells(writeRow, 2).Value = TimeValue(pCell)
                        On Error GoTo 0
                    Else
                        wsData.Cells(writeRow, 2).Value = pCell
                    End If

                Case Else
                    pCell = Replace(pCell, ",", ".")
                    If IsNumeric(pCell) Then
                        wsData.Cells(writeRow, pi2 + 1).Value = CDbl(pCell)
                    Else
                        wsData.Cells(writeRow, pi2 + 1).Value = pCell
                    End If
            End Select
        Next pi2

        writeRow = writeRow + 1
    Next pi

    wsData.Columns(1).NumberFormat = "dd.mm.yyyy"
    wsData.Columns(2).NumberFormat = "hh:mm:ss"
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
    headers(11) = "ЗАДАННАЯ ТЕМПЕРАТУРА"
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
Sub PrepareReportSheet(wb As Workbook, ByRef wsReport As Worksheet, csvFileName As String)
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name = "F0_Report" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    Set wsReport = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsReport.Name = "F0_Report"

    With wsReport
        ' Белый фон всего листа — сначала, чтобы не перекрыть заголовки
        .Cells.Interior.Color = RGB(255, 255, 255)
        .Cells.Font.Color = RGB(30, 30, 30)

        .Cells(1, 1).Value = "ПРОТОКОЛ СТЕРИЛИЗАЦИИ — РАСЧЁТ F0"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 100, 180)
        .Range("A1:J1").Merge

        .Cells(2, 1).Value = "Файл: " & csvFileName & "   |   Дата формирования: " & Format(Now, "dd.mm.yyyy hh:mm")
        .Cells(2, 1).Font.Color = RGB(100, 120, 140)
        .Cells(2, 1).Font.Bold = True
        .Range("A2:J2").Merge

        .Cells(3, 1).Value = "Стерилизация: Tref = 121.1C  |  z-фактор = 10C  |  СЭ по датчику в центре продукта  |  Норма F0 >= 5.5"
        .Cells(3, 1).Font.Color = RGB(100, 120, 140)
        .Range("A3:J3").Merge

        .Rows(4).RowHeight = 6

        Dim colHeaders(1 To 10) As String
        colHeaders(1) = "Цикл"
        colHeaders(2) = "Начало"
        colHeaders(3) = "Конец"
        colHeaders(4) = "Длит. (мин)"
        colHeaders(5) = "T макс. (C)"
        colHeaders(6) = "T мин. при стерил. (C)"
        colHeaders(7) = "F0 (мин)"
        colHeaders(8) = "Результат"
        colHeaders(9) = "Строк"
        colHeaders(10) = "Примечание"

        Dim c As Integer
        For c = 1 To 10
            .Cells(5, c).Value = colHeaders(c)
            .Cells(5, c).Font.Bold = True
            .Cells(5, c).Font.Color = RGB(255, 255, 255)
            .Cells(5, c).Interior.Color = RGB(30, 80, 130)
            .Cells(5, c).HorizontalAlignment = xlCenter
            .Cells(5, c).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Cells(5, c).Borders(xlEdgeBottom).Color = RGB(0, 120, 200)
        Next c
    End With
End Sub

'-------------------------------------------------------------
' Вспомогательная: возвращает Δt в минутах между строками r1 и r2
'-------------------------------------------------------------
Function CalcDeltaT(ws As Worksheet, r1 As Long, r2 As Long) As Double
    CalcDeltaT = 0
    ' Источник 1: столбец C (миллисекунды суток)
    Dim ms1v As Variant, ms2v As Variant
    ms1v = ws.Cells(r1, 3).Value
    ms2v = ws.Cells(r2, 3).Value
    If IsNumeric(ms1v) And IsNumeric(ms2v) Then
        Dim diffSec As Double
        diffSec = (CDbl(ms2v) - CDbl(ms1v)) / 1000#
        If diffSec < 0 Then  ' переход суток
            Dim da1 As Variant, da2 As Variant
            da1 = ws.Cells(r1, 1).Value : da2 = ws.Cells(r2, 1).Value
            If IsNumeric(da1) And IsNumeric(da2) Then
                diffSec = diffSec + (CLng(CDbl(da2)) - CLng(CDbl(da1))) * 86400#
            Else
                diffSec = diffSec + 86400#
            End If
        End If
        If diffSec >= 0.5 And diffSec <= 1800 Then
            CalcDeltaT = diffSec / 60#
            Exit Function
        End If
    End If
    ' Источник 2: столбцы A+B (дата + время как числа Excel)
    Dim d1v As Variant, t1v As Variant, d2v As Variant, t2v As Variant
    d1v = ws.Cells(r1, 1).Value : t1v = ws.Cells(r1, 2).Value
    d2v = ws.Cells(r2, 1).Value : t2v = ws.Cells(r2, 2).Value
    If IsNumeric(d1v) And IsNumeric(t1v) And IsNumeric(d2v) And IsNumeric(t2v) Then
        Dim dtMin As Double
        dtMin = ((CDbl(d2v) + CDbl(t2v)) - (CDbl(d1v) + CDbl(t1v))) * 1440#
        If dtMin < 0 Then dtMin = dtMin + 1440#  ' переход суток
        If dtMin >= 0.008 And dtMin <= 30 Then
            CalcDeltaT = dtMin
            Exit Function
        End If
    End If
    ' Источник 3: время как строка HH:MM:SS
    If IsNumeric(d1v) And IsNumeric(d2v) Then
        Dim ts1 As String, ts2 As String
        ts1 = Trim(CStr(t1v)) : ts2 = Trim(CStr(t2v))
        If InStr(ts1, ":") > 0 And InStr(ts2, ":") > 0 Then
            On Error Resume Next
            Dim tv1 As Double, tv2 As Double
            tv1 = CDbl(TimeValue(ts1)) : tv2 = CDbl(TimeValue(ts2))
            If Err.Number = 0 Then
                Dim tdiff As Double
                tdiff = (tv2 - tv1)
                If tdiff < 0 Then tdiff = tdiff + (CLng(CDbl(d2v)) - CLng(CDbl(d1v))) + 1
                tdiff = tdiff * 1440#
                If tdiff >= 0.008 And tdiff <= 30 Then CalcDeltaT = tdiff
            End If
            On Error GoTo 0
        End If
    End If
End Function

'-------------------------------------------------------------
' Определение циклов стерилизации и расчёт F0
' Алгоритм ДВУХПРОХОДНЫЙ:
'   Проход 1: определяем границы циклов и MAX(столбец K) → Tref
'   Проход 2: считаем F0 = Σ 10^((T-Tref)/z) * Δt  (прямая формула Бигелоу)
'-------------------------------------------------------------
Sub DetectCyclesAndCalculateF0(wsData As Worksheet, wsReport As Worksheet, lastRow As Long)
    Const COL_DATE As Integer = 1
    Const COL_TIME As Integer = 2
    Const COL_TEMP_PROD As Integer = 5
    Const COL_PRESSURE As Integer = 6   ' F — Давление
    Const COL_WATER As Integer = 7      ' G — Уровень воды
    Const COL_TREF As Integer = 11
    Const COL_F0 As Integer = 18

    ' Пороги начала/конца цикла по состоянию автоклава
    Const WATER_ON As Double = 9#       ' Уровень воды для старта цикла
    Const PRESS_ON As Double = 0.7      ' Давление для старта (>0.7 бар)
    Const WATER_OFF As Double = 9#      ' Слив воды — конец цикла

    ' ================================================================
    ' ПРОХОД 1: находим границы циклов ПО УРОВНЮ ВОДЫ + ДАВЛЕНИЮ
    '   Старт: вода > 9 И давление > 0.7 (автоклав наполнен и под давлением)
    '   Конец: вода опустилась < 9 (слив) — программа завершена
    ' ================================================================
    Dim cycleCount As Integer
    cycleCount = 0
    ReDim cycleStarts(1 To 500) As Long
    ReDim cycleEnds_(1 To 500) As Long
    ReDim cycleTref(1 To 500) As Double

    Dim inCyc As Boolean
    Dim cyStart As Long
    Dim lowCount As Integer  ' счётчик подряд "сухих" строк для надёжного конца
    inCyc = False : lowCount = 0

    Dim p As Long
    For p = 2 To lastRow
        ' Считываем уровень воды и давление
        Dim wRv As Variant, prRv As Variant
        wRv = wsData.Cells(p, COL_WATER).Value
        prRv = wsData.Cells(p, COL_PRESSURE).Value

        Dim waterLvl As Double, pressLvl As Double
        waterLvl = IIf(IsNumeric(wRv), CDbl(wRv), 0)
        pressLvl = IIf(IsNumeric(prRv), CDbl(prRv), 0)

        If Not inCyc Then
            ' Старт цикла: вода набрана И есть давление
            If waterLvl > WATER_ON And pressLvl > PRESS_ON Then
                inCyc = True
                cyStart = p
                lowCount = 0
            End If
        Else
            ' Конец цикла: вода слита (несколько строк подряд для надёжности)
            If waterLvl < WATER_OFF Then
                lowCount = lowCount + 1
            Else
                lowCount = 0
            End If

            Dim cycEnd1 As Boolean
            cycEnd1 = (lowCount >= 3) Or (p = lastRow)
            If cycEnd1 Then
                ' Конец цикла — на строке где вода ушла (откатываем назад)
                Dim realEnd As Long
                realEnd = IIf(lowCount >= 3, p - lowCount + 1, p)
                If realEnd < cyStart Then realEnd = cyStart

                If cycleCount < 500 Then
                    cycleCount = cycleCount + 1
                    cycleStarts(cycleCount) = cyStart
                    cycleEnds_(cycleCount) = realEnd
                    cycleTref(cycleCount) = T_REF  ' Tref=121.1 для стерилизации
                End If
                inCyc = False : lowCount = 0
            End If
        End If
P1Next:
    Next p

    ' ================================================================
    ' ПРОХОД 2: считаем F0 по каждому циклу с правильным Tref
    ' Формула Бигелоу: F0 = Σ 10^((T - Tref) / z) * Δt
    ' ================================================================
    Dim reportRow As Integer
    reportRow = 6

    Dim ci As Integer
    For ci = 1 To cycleCount
        Dim rStart As Long, rEnd As Long, tRefC As Double
        rStart = cycleStarts(ci)
        rEnd   = cycleEnds_(ci)
        tRefC  = cycleTref(ci)

        Dim f0C As Double, tMaxC As Double, tMinC As Double
        Dim peakC As Boolean
        f0C = 0 : tMaxC = -999 : tMinC = 999 : peakC = False

        Dim ri As Long
        For ri = rStart To rEnd
            Dim tpC As Variant
            tpC = wsData.Cells(ri, COL_TEMP_PROD).Value
            If Not IsNumeric(tpC) Then GoTo P2Next
            Dim tC As Double : tC = CDbl(tpC)

            If tC > tMaxC Then tMaxC = tC
            ' Пик стерилизации — подтверждение, что цикл достиг 100°C
            If tC >= T_PEAK_STERIL Then peakC = True

            ' F0 накапливается с 90°C — плавная кривая как на автоклаве
            If tC >= T_MIN_STERIL Then
                If tC < tMinC Then tMinC = tC

                Dim dtC As Double : dtC = 0
                If ri > rStart Then dtC = CalcDeltaT(wsData, ri - 1, ri)

                If dtC > 0 Then
                    ' Прямая формула Бигелоу: L = 10^((T - Tref)/z), F0 += L * Δt
                    f0C = f0C + (10# ^ ((tC - tRefC) / Z_FACTOR)) * dtC
                End If
            End If
            ' Записываем накопленный F0 в столбец R
            wsData.Cells(ri, COL_F0).Value = Round(f0C, 4)
P2Next:
        Next ri

        ' Результат для отчёта
        Dim cycleNum As Integer : cycleNum = ci
        Dim durationMin As Double
        durationMin = GetDateTimeAsDouble(wsData, rStart)
        Dim dtEnd2 As Double
        dtEnd2 = GetDateTimeAsDouble(wsData, rEnd)
        If dtEnd2 > durationMin Then
            durationMin = (dtEnd2 - durationMin) * 1440#
        Else
            durationMin = 0
        End If

        Dim tMin As Double, tMax As Double
        tMax = tMaxC
        tMin = IIf(tMinC < 999, tMinC, 0)

        Dim f0Cycle As Double : f0Cycle = f0C
        Dim tRefCycle As Double : tRefCycle = tRefC

        Dim result As String, resultColor As Long, noteText As String
        ' Норма F0 для СТЕРИЛИЗАЦИИ (Tref=121.1°C, тушёнка по ГОСТ):
        ' минимум 4.0-5.5, рекомендуемо 5.5-8.0 усл. минут
        Dim f0Norm As Double
        f0Norm = 5.5

        Dim trefStr As String
        trefStr = "Tref=121.1C, z=10"

        If Not peakC Then
            result = "— Без стерилизации"
            resultColor = RGB(100, 120, 140)
            noteText = "Пик T < 100C — только прогрев"
        ElseIf f0Cycle >= f0Norm Then
            result = "OK НОРМА (F0 >= " & f0Norm & ")"
            resultColor = RGB(0, 160, 80)
            noteText = trefStr
        ElseIf f0Cycle >= f0Norm / 2 Then
            result = "! ПРЕДЕЛ (F0 >= " & f0Norm / 2 & ")"
            resultColor = RGB(220, 140, 0)
            noteText = trefStr
        Else
            result = "X НЕДОСТАТОЧНО (F0 < " & f0Norm / 2 & ")"
            resultColor = RGB(200, 40, 40)
            noteText = trefStr
        End If

        Dim startDateStr As String, endDateStr As String
        startDateStr = FormatDateTime_FromRow(wsData, rStart)
        endDateStr   = FormatDateTime_FromRow(wsData, rEnd)

        With wsReport
            .Cells(reportRow, 1).Value = ci
            .Cells(reportRow, 2).Value = startDateStr
            .Cells(reportRow, 3).Value = endDateStr
            .Cells(reportRow, 4).Value = Round(durationMin, 1)
            .Cells(reportRow, 5).Value = Round(tMax, 2)
            .Cells(reportRow, 6).Value = IIf(tMin < 999, Round(tMin, 2), "—")
            .Cells(reportRow, 7).Value = Round(f0Cycle, 4)
            .Cells(reportRow, 8).Value = result
            .Cells(reportRow, 9).Value = rEnd - rStart + 1
            .Cells(reportRow, 10).Value = noteText

            .Cells(reportRow, 7).NumberFormat = "0.0000"
            .Cells(reportRow, 8).Font.Color = resultColor
            .Cells(reportRow, 8).Font.Bold = True

            If ci Mod 2 = 0 Then
                .Rows(reportRow).Interior.Color = RGB(240, 246, 252)
            Else
                .Rows(reportRow).Interior.Color = RGB(255, 255, 255)
            End If
            .Rows(reportRow).Font.Color = RGB(30, 30, 30)
            .Cells(reportRow, 8).Font.Color = resultColor
        End With

        reportRow = reportRow + 1
    Next ci

    Call AddSummaryRow(wsReport, reportRow, cycleCount)
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
        .Cells(reportRow, 4).Value = ""
        .Cells(reportRow, 5).Value = "=MAX(E6:E" & lastDataRow & ")"
        .Cells(reportRow, 5).Font.Bold = True
        ' F0 не суммируется — у каждой программы своё значение
        .Cells(reportRow, 7).Value = ""
        .Cells(reportRow, 8).Value = "Всего циклов: " & totalCycles
        .Cells(reportRow, 8).Font.Bold = True
        .Rows(reportRow).Interior.Color = RGB(220, 235, 248)
        .Rows(reportRow).Font.Color = RGB(20, 60, 100)
        .Cells(reportRow, 1).Font.Color = RGB(0, 100, 180)
        .Rows(reportRow).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Rows(reportRow).Borders(xlEdgeTop).Color = RGB(0, 100, 180)
        .Rows(reportRow).Borders(xlEdgeTop).Weight = xlMedium
    End With
End Sub

'-------------------------------------------------------------
' Форматирование листа отчёта
'-------------------------------------------------------------
Sub FormatReportSheet(wsReport As Worksheet)
    With wsReport
        .Columns(1).ColumnWidth = 7
        .Columns(2).ColumnWidth = 22
        .Columns(3).ColumnWidth = 22
        .Columns(4).ColumnWidth = 13
        .Columns(5).ColumnWidth = 13
        .Columns(6).ColumnWidth = 24
        .Columns(7).ColumnWidth = 13
        .Columns(8).ColumnWidth = 26
        .Columns(9).ColumnWidth = 8
        .Columns(10).ColumnWidth = 20
        .Columns(1).HorizontalAlignment = xlCenter
        .Columns(4).HorizontalAlignment = xlCenter
        .Columns(7).HorizontalAlignment = xlCenter
        .Columns(9).HorizontalAlignment = xlCenter
        ' Белый фон, чёрный шрифт
        .Cells.Interior.Color = RGB(255, 255, 255)
        .Cells.Font.Color = RGB(0, 0, 0)
        ' Заголовок
        .Cells(1, 1).Font.Color = RGB(0, 100, 180)
        .Cells(1, 1).Font.Size = 14
        .Rows(2).Font.Color = RGB(80, 80, 80)
        .Rows(3).Font.Color = RGB(80, 80, 80)
    End With
End Sub

'-------------------------------------------------------------
' График температуры и F0
'-------------------------------------------------------------
'-------------------------------------------------------------
' Строит один график для одного цикла (строки rStart..rEnd)
'-------------------------------------------------------------
Sub BuildOneCycleChart(ws As Worksheet, wsData As Worksheet, _
    rStart As Long, rEnd As Long, cycleIdx As Integer, _
    tRefC As Double, topOffset As Long)

    ' Размер под печать A4: ширина ~ до колонки N, один график на лист
    Const CHART_W As Long = 720   ' до колонки N (помещается при печати A4)
    Const CHART_H As Long = 680   ' почти весь лист A4 по высоте

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add( _
        Left:=5, Top:=topOffset, Width:=CHART_W, Height:=CHART_H)

    Dim cht As Chart
    Set cht = co.Chart
    cht.ChartType = xlLine
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Dim nRows As Long : nRows = rEnd - rStart + 1

    ' --- Метки времени на оси X (формат HH:MM) ---
    Dim timeLabels() As String
    ReDim timeLabels(1 To nRows)
    Dim ri As Long
    For ri = 1 To nRows
        timeLabels(ri) = FormatDateTime_FromRow(wsData, rStart + ri - 1)
        ' Оставляем только время HH:MM
        Dim fullLabel As String : fullLabel = timeLabels(ri)
        Dim spPos As Integer : spPos = InStr(fullLabel, " ")
        If spPos > 0 Then
            Dim timePart2 As String : timePart2 = Mid(fullLabel, spPos + 1)
            ' Обрезаем до HH:MM (убираем секунды)
            Dim colPos As Integer : colPos = InStr(timePart2, ":")
            If colPos > 0 Then
                colPos = InStr(colPos + 1, timePart2, ":")
                If colPos > 0 Then timePart2 = Left(timePart2, colPos - 1)
            End If
            timeLabels(ri) = timePart2
        End If
    Next ri

    ' --- Данные серий ---
    Dim arrEnv() As Double, arrProd() As Double, arrF0() As Double, arrTref() As Double
    ReDim arrEnv(1 To nRows)
    ReDim arrProd(1 To nRows)
    ReDim arrF0(1 To nRows)
    ReDim arrTref(1 To nRows)

    Dim f0MaxVal As Double : f0MaxVal = 0

    For ri = 1 To nRows
        Dim rowIdx As Long : rowIdx = rStart + ri - 1
        Dim vEnv As Variant, vProd As Variant, vF0 As Variant
        vEnv  = wsData.Cells(rowIdx, 4).Value   ' D — темп.среды
        vProd = wsData.Cells(rowIdx, 5).Value   ' E — темп.продукта
        vF0   = wsData.Cells(rowIdx, 18).Value  ' R — F0
        arrEnv(ri)  = IIf(IsNumeric(vEnv), CDbl(vEnv), 0)
        arrProd(ri) = IIf(IsNumeric(vProd), CDbl(vProd), 0)
        arrF0(ri)   = IIf(IsNumeric(vF0), CDbl(vF0), 0)
        arrTref(ri) = tRefC
        If arrF0(ri) > f0MaxVal Then f0MaxVal = arrF0(ri)
    Next ri

    ' Округляем максимум F0 вверх для красивой шкалы
    Dim f0AxisMax As Double
    If f0MaxVal <= 0 Then
        f0AxisMax = 1
    ElseIf f0MaxVal <= 5 Then
        f0AxisMax = 5
    ElseIf f0MaxVal <= 10 Then
        f0AxisMax = 10
    ElseIf f0MaxVal <= 20 Then
        f0AxisMax = 20
    ElseIf f0MaxVal <= 50 Then
        f0AxisMax = 50
    Else
        f0AxisMax = CDbl(CLng(f0MaxVal * 1.2 / 10 + 1) * 10)
    End If

    ' --- Серия 1: T° продукта — бирюзовая, сглаженная ---
    Dim s1 As Series
    Set s1 = cht.SeriesCollection.NewSeries
    s1.Name = "T° продукта"
    s1.Values = arrProd
    s1.XValues = timeLabels
    s1.Format.Line.ForeColor.RGB = RGB(0, 210, 200)
    s1.Format.Line.Weight = 2.5
    s1.MarkerStyle = xlMarkerStyleNone
    s1.Smooth = True
    s1.AxisGroup = xlPrimary

    ' --- Серия 2: Температура среды — жёлтая ---
    Dim s2 As Series
    Set s2 = cht.SeriesCollection.NewSeries
    s2.Name = "Температура"
    s2.Values = arrEnv
    s2.XValues = timeLabels
    s2.Format.Line.ForeColor.RGB = RGB(220, 190, 0)
    s2.Format.Line.Weight = 2
    s2.MarkerStyle = xlMarkerStyleNone
    s2.Smooth = False  ' среда реагирует быстро — не сглаживаем
    s2.AxisGroup = xlPrimary

    ' --- Серия 3: F0 накопленный — оранжевая, вторая ось, сглаженная ---
    Dim s3 As Series
    Set s3 = cht.SeriesCollection.NewSeries
    s3.Name = "F0 накопл."
    s3.Values = arrF0
    s3.XValues = timeLabels
    s3.Format.Line.ForeColor.RGB = RGB(255, 100, 0)
    s3.Format.Line.Weight = 2.5
    s3.MarkerStyle = xlMarkerStyleNone
    s3.Smooth = True   ' сглаженная кривая
    s3.AxisGroup = xlSecondary

    ' --- Серия 4: Tref — красная пунктирная ---
    Dim s4 As Series
    Set s4 = cht.SeriesCollection.NewSeries
    s4.Name = "Tref=" & Format(tRefC, "0") & "C"
    s4.Values = arrTref
    s4.XValues = timeLabels
    s4.Format.Line.ForeColor.RGB = RGB(220, 50, 50)
    s4.Format.Line.DashStyle = msoLineDash
    s4.Format.Line.Weight = 1.5
    s4.MarkerStyle = xlMarkerStyleNone
    s4.AxisGroup = xlPrimary

    With cht
        .HasTitle = True
        .ChartTitle.Text = "Цикл " & cycleIdx & "  |  Tref = " & Format(tRefC, "0") & "°C"
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(20, 20, 20)

        .PlotArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Border.LineStyle = xlContinuous
        .PlotArea.Border.Color = RGB(180, 200, 220)
        .ChartArea.Interior.Color = RGB(250, 252, 255)
        .ChartArea.Border.Color = RGB(180, 200, 220)

        ' Ось Y (температура)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = False
            .MinimumScale = 0
            .MaximumScale = 150
            .MajorUnit = 25
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(210, 220, 230)
            .MajorGridlines.Format.Line.DashStyle = msoLineDash
            .TickLabels.Font.Color = RGB(50, 50, 50)
            .TickLabels.Font.Size = 9
        End With

        ' Ось F0 (вторичная) — фиксируем шкалу от 0 до max чтобы не прыгала
        With .Axes(xlValue, xlSecondary)
            .HasTitle = False
            .MinimumScale = 0
            .MaximumScale = f0AxisMax
            .TickLabels.Font.Color = RGB(200, 90, 0)
            .TickLabels.Font.Size = 8
        End With

        ' Ось X — время HH:MM
        ' Цель: ~20-30 меток на графике независимо от количества строк
        With .Axes(xlCategory)
            .HasTitle = False
            .TickLabels.Font.Size = 7
            .TickLabels.Font.Color = RGB(50, 50, 50)
            Dim tickStep As Long
            If nRows > 3000 Then
                tickStep = CLng(nRows / 25)      ' ~25 меток
            ElseIf nRows > 600 Then
                tickStep = CLng(nRows / 20)      ' ~20 меток
            ElseIf nRows > 120 Then
                tickStep = CLng(nRows / 15)      ' ~15 меток
            ElseIf nRows > 30 Then
                tickStep = CLng(nRows / 10)      ' ~10 меток
            Else
                tickStep = 1
            End If
            If tickStep < 1 Then tickStep = 1
            .TickLabelSpacing = tickStep
            .TickMarkSpacing = tickStep
        End With

        .HasLegend = True
        .Legend.Interior.Color = RGB(255, 255, 255)
        .Legend.Font.Color = RGB(30, 30, 30)
        .Legend.Font.Size = 9
        .Legend.Position = xlLegendPositionTop
    End With
End Sub

'-------------------------------------------------------------
' Строит графики по всем циклам на листе График
'-------------------------------------------------------------
Sub BuildTemperatureChart(wb As Workbook, wsData As Worksheet, lastRow As Long, csvFileName As String)
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If ws.Name = "График" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "График"
    ws.Cells.Interior.Color = RGB(245, 248, 252)

    ' Заголовок листа
    ws.Cells(1, 1).Value = "Графики температуры и F0 — " & csvFileName
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = RGB(0, 80, 160)

    ' Сканируем Data — находим циклы по УРОВНЮ ВОДЫ + ДАВЛЕНИЮ (как в Проходе 1)
    Const WATER_ON As Double = 9#
    Const PRESS_ON As Double = 0.7
    Const WATER_OFF As Double = 9#

    Dim inCyc As Boolean : inCyc = False
    Dim cyStart As Long, cyEnd As Long
    Dim lowCount As Integer : lowCount = 0
    Dim cycIdx As Integer : cycIdx = 0
    Dim topOffset As Long : topOffset = 30

    Dim p As Long
    For p = 2 To lastRow
        Dim wRv As Variant, prRv As Variant
        wRv = wsData.Cells(p, 7).Value   ' G — уровень воды
        prRv = wsData.Cells(p, 6).Value  ' F — давление
        Dim waterLvl As Double, pressLvl As Double
        waterLvl = IIf(IsNumeric(wRv), CDbl(wRv), 0)
        pressLvl = IIf(IsNumeric(prRv), CDbl(prRv), 0)

        If Not inCyc Then
            If waterLvl > WATER_ON And pressLvl > PRESS_ON Then
                inCyc = True : cyStart = p : lowCount = 0
            End If
        Else
            If waterLvl < WATER_OFF Then
                lowCount = lowCount + 1
            Else
                lowCount = 0
            End If
            Dim cycEnd2 As Boolean
            cycEnd2 = (lowCount >= 3) Or (p = lastRow)
            If cycEnd2 Then
                cyEnd = IIf(lowCount >= 3, p - lowCount + 1, p)
                If cyEnd < cyStart Then cyEnd = cyStart

                Dim tRefCy As Double
                tRefCy = T_REF

                cycIdx = cycIdx + 1
                Call BuildOneCycleChart(ws, wsData, cyStart, cyEnd, cycIdx, tRefCy, topOffset)
                topOffset = topOffset + 700  ' следующий график на новом листе A4

                inCyc = False : lowCount = 0
            End If
        End If
ChartNext:
    Next p

    ' Настройка печати: A4 портрет, вписать по ширине в 1 страницу,
    ' каждый график (700px) попадает на отдельный лист
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False  ' высота не ограничена — много страниц
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
    End With
End Sub

Function SheetExistsInWb(wb As Workbook, sheetName As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = wb.Sheets(sheetName)
    SheetExistsInWb = Not s Is Nothing
    On Error GoTo 0
End Function

'-------------------------------------------------------------
' Вспомогательная: форматирует дату+время из строки листа Data
' Время строится из миллисекунд (столбец C) — самый точный источник
'-------------------------------------------------------------
Function FormatDateTime_FromRow(ws As Worksheet, rowIdx As Long) As String
    Dim dateVal As Variant
    Dim msVal As Variant
    Dim timeVal As Variant

    dateVal = ws.Cells(rowIdx, 1).Value  ' столбец A — дата
    msVal   = ws.Cells(rowIdx, 3).Value  ' столбец C — миллисекунды от начала суток
    timeVal = ws.Cells(rowIdx, 2).Value  ' столбец B — время (запасной вариант)

    ' Форматируем дату
    Dim datePart As String
    If IsNumeric(dateVal) And CLng(dateVal) > 0 Then
        datePart = Format(CDate(CLng(dateVal)), "dd.mm.yyyy")
    ElseIf IsDate(dateVal) Then
        datePart = Format(CDate(dateVal), "dd.mm.yyyy")
    Else
        datePart = CStr(dateVal)
    End If

    ' Форматируем время — приоритет: столбец B (строка HH:MM:SS или дробь Excel)
    Dim timePart As String
    timePart = ""

    ' Вариант 1: время как строка "HH:MM:SS" (столбец B)
    Dim tvStr As String
    tvStr = Trim(CStr(timeVal))
    If InStr(tvStr, ":") > 0 Then
        ' Убираем возможные кавычки
        If Left(tvStr, 1) = Chr(34) Then tvStr = Mid(tvStr, 2, Len(tvStr) - 2)
        On Error Resume Next
        Dim tvCDate As Date
        tvCDate = TimeValue(tvStr)
        If Err.Number = 0 Then
            timePart = Format(tvCDate, "hh:mm:ss")
        End If
        On Error GoTo 0
    End If

    ' Вариант 2: время как дробь Excel (0..1)
    If timePart = "" And IsNumeric(timeVal) Then
        Dim tvDbl As Double
        tvDbl = CDbl(timeVal)
        If tvDbl >= 0 And tvDbl < 1 Then
            timePart = Format(CDate(tvDbl), "hh:mm:ss")
        End If
    End If

    ' Вариант 3: из миллисекунд суток (столбец C), если там большое число (> 1000)
    If timePart = "" And IsNumeric(msVal) Then
        Dim msLng As Long
        msLng = CLng(CDbl(msVal))
        If msLng > 1000 Then
            Dim totalSec As Long
            totalSec = msLng \\ 1000
            Dim hh As Integer, mm As Integer, ss As Integer
            hh = totalSec \\ 3600
            mm = (totalSec Mod 3600) \\ 60
            ss = totalSec Mod 60
            timePart = Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(ss, "00")
        End If
    End If

    If timePart = "" Then timePart = CStr(timeVal)

    FormatDateTime_FromRow = datePart & " " & timePart
End Function

'-------------------------------------------------------------
' Вспомогательная: возвращает дату+время как число Excel (для Δt)
'-------------------------------------------------------------
Function GetDateTimeAsDouble(ws As Worksheet, rowIdx As Long) As Double
    Dim dateVal As Variant
    Dim timeVal As Variant
    dateVal = ws.Cells(rowIdx, 1).Value
    timeVal = ws.Cells(rowIdx, 2).Value

    ' Получаем числовое значение даты
    Dim dateDbl As Double
    dateDbl = 0
    If IsNumeric(dateVal) And CLng(CDbl(dateVal)) > 0 Then
        dateDbl = CDbl(CLng(CDbl(dateVal)))
    ElseIf IsDate(dateVal) Then
        dateDbl = CDbl(CDate(dateVal))
    End If

    ' Получаем числовое значение времени (дробь 0..1)
    Dim timeDbl As Double
    timeDbl = 0
    If IsNumeric(timeVal) Then
        timeDbl = CDbl(timeVal)
    Else
        Dim tvs As String
        tvs = Trim(CStr(timeVal))
        If Left(tvs, 1) = Chr(34) Then tvs = Mid(tvs, 2, Len(tvs) - 2)
        If InStr(tvs, ":") > 0 Then
            On Error Resume Next
            timeDbl = CDbl(TimeValue(tvs))
            On Error GoTo 0
        End If
    End If

    If dateDbl > 0 Then
        GetDateTimeAsDouble = dateDbl + timeDbl
    Else
        GetDateTimeAsDouble = 0
    End If
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