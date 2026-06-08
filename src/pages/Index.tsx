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
' Порог накопления F0. По реальным данным автоклава фактор стерилизации
' начинает накапливаться с 91°C (соответствует поведению прибора).
Const T_MIN_STERIL As Double = 91#   ' Минимум для счёта F0
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
    Dim suggestPrev As String
    Dim hintPrev As String
    Dim msgText As String
    Dim ans As Integer
    Dim defPrevPath As String
    Dim pickAgain As Boolean
    Dim chosenName As String
    Dim warnAns As Integer
    prevFilePath = ""

    If NeedsPreviousFile(wsData) Then
        Application.ScreenUpdating = True

        ' Подсказываем имя предыдущего файла (дата −1 день)
        suggestPrev = GetNeighborFileName(csvFileName, -1)
        If suggestPrev <> "" Then
            hintPrev = Chr(13) & Chr(13) & "РЕКОМЕНДУЕМЫЙ ФАЙЛ:  " & suggestPrev
            If Len(Dir(csvFolder & suggestPrev)) > 0 Then
                hintPrev = hintPrev & Chr(13) & "(найден в этой папке)"
            Else
                hintPrev = hintPrev & Chr(13) & "(в папке не найден — выберите вручную)"
            End If
        End If

        ' Формируем сообщение — имя файла крупным блоком через пустые строки
        msgText = "Данные начинаются при температуре продукта > 40" & Chr(176) & "C в начале суток." & Chr(13) & _
                  "Похоже, замес начался в ПРЕДЫДУЩЕМ файле." & Chr(13) & Chr(13)
        If suggestPrev <> "" Then
            msgText = msgText & "РЕКОМЕНДУЕМЫЙ ФАЙЛ:" & Chr(13) & Chr(13) & _
                      "  >>> " & suggestPrev & " <<<" & Chr(13) & Chr(13)
            If Len(Dir(csvFolder & suggestPrev)) > 0 Then
                msgText = msgText & "(найден в этой папке)" & Chr(13) & Chr(13)
            Else
                msgText = msgText & "(в папке не найден — выберите вручную)" & Chr(13) & Chr(13)
            End If
        End If
        msgText = msgText & "Загрузить предыдущий CSV-файл для корректного расчёта F0 и времени цикла?"

        ans = MsgBox(msgText, vbYesNo + vbQuestion, "Переход суток — нужен ПРЕДЫДУЩИЙ файл")

        If ans = vbYes Then
            defPrevPath = ""
            pickAgain = True

            Do While pickAgain
                pickAgain = False
                prevFilePath = Application.GetOpenFilename( _
                    FileFilter:="CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
                    Title:="Выберите ПРЕДЫДУЩИЙ файл (рекомендуется: " & suggestPrev & ")")
                If prevFilePath = "False" Then
                    prevFilePath = ""
                Else
                    ' Валидация: проверяем что выбран правильный файл
                    chosenName = Mid(prevFilePath, InStrRev(prevFilePath, "\\") + 1)
                    If suggestPrev <> "" And LCase(chosenName) <> LCase(suggestPrev) Then
                        warnAns = MsgBox("Выбран файл:  " & chosenName & Chr(13) & Chr(13) & _
                                         "Рекомендуется:  " & suggestPrev & Chr(13) & Chr(13) & _
                                         "Вы выбрали другой файл — данные могут быть некорректны." & Chr(13) & Chr(13) & _
                                         "Использовать выбранный файл?", _
                                         vbYesNo + vbExclamation, "Проверка файла")
                        If warnAns = vbNo Then
                            pickAgain = True
                            prevFilePath = ""
                        End If
                    End If
                End If
            Loop
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
    ' Если выбран предыдущий файл — вставляем только хвост последнего цикла
    ' (строки от T>=40°C до конца предыдущего файла)
    ' ----------------------------------------------------------------
    Dim prevInsertedRows As Long
    prevInsertedRows = 0
    If prevFilePath <> "" Then
        prevInsertedRows = PrependPreviousCSV(wb, wsData, prevFilePath)
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

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' Определяем Tref из данных (столбец K) для строки заголовка отчёта
    Dim trefInfoStr As String
    Dim trefScan As Long, tKscan As Double : tKscan = 0
    For trefScan = 2 To lastRow
        Dim tkSv As Variant : tkSv = wsData.Cells(trefScan, 11).Value
        If IsNumeric(tkSv) Then
            Dim tkSd As Double : tkSd = CDbl(tkSv)
            If tkSd >= 100# And tkSd <= 130# And tkSd > tKscan Then tKscan = tkSd
        End If
    Next trefScan
    Dim trefNorm As Double
    trefNorm = 8#
    trefInfoStr = "Стерилизация: Tref = 121.1C (ГОСТ)  |  z-фактор = 10C  |  СЭ по датчику в центре продукта  |  Норма F0 >= 8"

    Call PrepareReportSheet(wb, wsReport, csvFileName, trefInfoStr)
    Call DetectCyclesAndCalculateF0(wsData, wsReport, lastRow, prevInsertedRows)
    Call FormatReportSheet(wsReport)
    Call BuildTemperatureChart(wb, wsData, lastRow, csvFileName, prevInsertedRows)

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
' Возвращает количество строк вставленных из предыдущего файла (0 если ничего не вставлено)
Function PrependPreviousCSV(wb As Workbook, wsData As Worksheet, prevFilePath As String) As Long
    PrependPreviousCSV = 0

    ' Читаем предыдущий CSV во временный массив
    Dim fileNum As Integer
    Dim lineText As String
    Dim prevRows() As String
    Dim prevCount As Long
    prevCount = 0

    fileNum = FreeFile
    Open prevFilePath For Input As #fileNum
    ReDim prevRows(1 To 100000)

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)
        If Len(lineText) = 0 Then GoTo SkipLine2
        If Left(lineText, 1) = "#" Then GoTo SkipLine2
        prevCount = prevCount + 1
        prevRows(prevCount) = lineText
SkipLine2:
    Loop
    Close #fileNum

    If prevCount = 0 Then Exit Function

    ' Определяем startIdx: пропускаем заголовок CSV если есть
    Dim startIdx As Long
    startIdx = 1
    Dim firstFields() As String
    If InStr(prevRows(1), ";") > 0 Then
        firstFields = Split(prevRows(1), ";")
    Else
        firstFields = Split(prevRows(1), ",")
    End If
    Dim fp As String
    fp = Trim(firstFields(0))
    If Len(fp) >= 2 And Left(fp, 1) = Chr(34) Then fp = Mid(fp, 2, Len(fp) - 2)
    If Not (fp Like "####/##/##") And Not IsDate(fp) Then startIdx = 2

    ' ----------------------------------------------------------------
    ' КЛЮЧЕВОЕ: из предыдущего файла берём только хвост —
    ' строки начиная с первой где T среды (col 4 = D) >= 40°C
    ' И при этом давление (col 6 = F) активно (>= 600 мБар) ИЛИ
    ' идёт непрерывный нагрев до такого давления.
    ' Алгоритм: находим последний цикл (последняя строка с P>=600),
    ' затем идём назад до T среды < 40°C — это и есть наш хвост.
    ' ----------------------------------------------------------------

    ' Шаг 1: парсим все строки предыдущего файла в простые массивы
    Dim pCount As Long
    pCount = prevCount - startIdx + 1
    If pCount <= 0 Then Exit Function

    Dim pTenv() As Double   ' T среды (col 4)
    Dim pPress() As Double  ' давление (col 6)
    Dim pDate() As Double   ' дата (Excel days)
    Dim pTime() As Double   ' время (Excel fraction)
    ReDim pTenv(1 To pCount)
    ReDim pPress(1 To pCount)
    ReDim pDate(1 To pCount)
    ReDim pTime(1 To pCount)

    Dim pi As Long
    Dim idx As Long
    idx = 0
    For pi = startIdx To prevCount
        lineText = prevRows(pi)
        Dim pf() As String
        If InStr(lineText, ";") > 0 Then
            pf = Split(lineText, ";")
        Else
            pf = Split(lineText, ",")
        End If

        ' Очищаем кавычки всех полей
        Dim pk As Integer
        For pk = 0 To UBound(pf)
            pf(pk) = Trim(pf(pk))
            If Len(pf(pk)) >= 2 Then
                If Left(pf(pk), 1) = Chr(34) And Right(pf(pk), 1) = Chr(34) Then
                    pf(pk) = Mid(pf(pk), 2, Len(pf(pk)) - 2)
                End If
            End If
        Next pk

        idx = idx + 1

        ' Дата (col 0)
        If UBound(pf) >= 0 Then
            Dim dCell As String : dCell = Trim(pf(0))
            If InStr(dCell, "/") > 0 Then
                Dim dp2() As String : dp2 = Split(dCell, "/")
                If UBound(dp2) = 2 Then
                    On Error Resume Next
                    pDate(idx) = CDbl(DateSerial(CInt(dp2(0)), CInt(dp2(1)), CInt(dp2(2))))
                    On Error GoTo 0
                End If
            ElseIf IsDate(dCell) Then
                On Error Resume Next
                pDate(idx) = CDbl(CDate(dCell))
                On Error GoTo 0
            End If
        End If

        ' Время (col 1)
        If UBound(pf) >= 1 Then
            Dim tCell As String : tCell = Trim(pf(1))
            If InStr(tCell, ":") > 0 Then
                On Error Resume Next
                pTime(idx) = CDbl(TimeValue(tCell))
                On Error GoTo 0
            End If
        End If

        ' T среды (col 3 = 0-based → pf(3))
        If UBound(pf) >= 3 Then
            Dim envCell As String : envCell = Replace(Trim(pf(3)), ",", ".")
            If IsNumeric(envCell) Then pTenv(idx) = CDbl(envCell)
        End If

        ' Давление (col 5 = 0-based → pf(5))
        If UBound(pf) >= 5 Then
            Dim prCell As String : prCell = Replace(Trim(pf(5)), ",", ".")
            If IsNumeric(prCell) Then pPress(idx) = CDbl(prCell)
        End If
    Next pi

    ' Шаг 2: находим последнюю строку с давлением >= 600 мБар (активный цикл)
    Dim lastActiveRow As Long
    lastActiveRow = 0
    Dim ri As Long
    For ri = pCount To 1 Step -1
        If pPress(ri) >= 600# Then
            lastActiveRow = ri
            Exit For
        End If
    Next ri

    ' Если в предыдущем файле нет активного цикла — ничего не вставляем
    If lastActiveRow = 0 Then Exit Function

    ' Шаг 3: от lastActiveRow идём назад до T среды < 40°C — это начало нагрева цикла
    Dim cycTailStart As Long
    cycTailStart = lastActiveRow
    For ri = lastActiveRow - 1 To 1 Step -1
        If pTenv(ri) >= 40# Then
            cycTailStart = ri
        Else
            Exit For
        End If
    Next ri
    ' Берём ещё ~5 минут запаса ДО начала нагрева (CSV ~10 сек/строка → 30 строк = 5 мин)
    ' чтобы на графике была видна точка начала нагрева с небольшим отступом слева
    Const PREPEND_EXTRA As Long = 30
    cycTailStart = cycTailStart - PREPEND_EXTRA
    If cycTailStart < 1 Then cycTailStart = 1

    ' Шаг 4: от lastActiveRow идём вперёд — берём ещё строки пока P > 0 или T > 40
    ' (давление могло упасть раньше чем конец файла — берём хвост охлаждения)
    Dim cycTailEnd As Long
    cycTailEnd = pCount  ' берём до самого конца предыдущего файла

    ' Шаг 5: вставляем строки cycTailStart..cycTailEnd перед данными основного файла
    Dim insertRows As Long
    insertRows = cycTailEnd - cycTailStart + 1
    If insertRows <= 0 Then Exit Function

    wsData.Rows("2:" & (insertRows + 1)).Insert Shift:=xlDown

    ' Записываем данные
    Dim writeRow As Long
    writeRow = 2
    For ri = cycTailStart To cycTailEnd
        ' Восстанавливаем из предыдущего файла — повторно парсим нужную строку
        lineText = prevRows(startIdx - 1 + ri)
        Dim pFields2() As String
        If InStr(lineText, ";") > 0 Then
            pFields2 = Split(lineText, ";")
        Else
            pFields2 = Split(lineText, ",")
        End If
        Dim pk2 As Integer
        For pk2 = 0 To UBound(pFields2)
            pFields2(pk2) = Trim(pFields2(pk2))
            If Len(pFields2(pk2)) >= 2 Then
                If Left(pFields2(pk2), 1) = Chr(34) And Right(pFields2(pk2), 1) = Chr(34) Then
                    pFields2(pk2) = Mid(pFields2(pk2), 2, Len(pFields2(pk2)) - 2)
                End If
            End If
        Next pk2

        Dim pTotalCols As Integer
        pTotalCols = UBound(pFields2) + 1
        If pTotalCols > 17 Then pTotalCols = 17

        Dim pi2 As Integer
        For pi2 = 0 To pTotalCols - 1
            Dim pCell As String
            pCell = Trim(pFields2(pi2))

            Select Case pi2
                Case 0
                    If InStr(pCell, "/") > 0 Then
                        Dim dpw() As String : dpw = Split(pCell, "/")
                        If UBound(dpw) = 2 Then
                            wsData.Cells(writeRow, 1).Value = DateSerial(CInt(dpw(0)), CInt(dpw(1)), CInt(dpw(2)))
                        Else
                            wsData.Cells(writeRow, 1).Value = pCell
                        End If
                    ElseIf IsDate(pCell) Then
                        wsData.Cells(writeRow, 1).Value = CDate(pCell)
                    Else
                        wsData.Cells(writeRow, 1).Value = pCell
                    End If
                Case 1
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
    Next ri

    wsData.Columns(1).NumberFormat = "dd.mm.yyyy"
    wsData.Columns(2).NumberFormat = "hh:mm:ss"

    PrependPreviousCSV = insertRows
End Function

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
Sub PrepareReportSheet(wb As Workbook, ByRef wsReport As Worksheet, csvFileName As String, trefInfo As String)
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

        .Cells(1, 1).Value = "ПРОТОКОЛ СТЕРИЛИЗАЦИИ"
        .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(0, 100, 180)
        .Range("A1:H1").Merge

        .Cells(2, 1).Value = "Файл: " & csvFileName
        .Cells(2, 1).Font.Color = RGB(100, 120, 140)
        .Cells(2, 1).Font.Bold = True
        .Range("A2:H2").Merge

        .Rows(3).RowHeight = 4

        ' ===== ТАБЛИЦА 1: основные данные (8 столбцов, умещается на А4) =====
        Dim h1(1 To 7) As String
        h1(1) = "Цикл" : h1(2) = "Начало" : h1(3) = "Конец"
        h1(4) = "Длительность" : h1(5) = "T макс. (C)"
        h1(6) = "F0 (мин)" : h1(7) = "Результат"

        Dim c As Integer
        For c = 1 To 7
            .Cells(5, c).Value = h1(c)
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
'-------------------------------------------------------------
'-------------------------------------------------------------
' Форматирует секунды в строку "Xч YYм" или "YYм"
'-------------------------------------------------------------
Function FormatMinSec(totalSec As Double) As String
    If totalSec <= 0 Then
        FormatMinSec = "—"
        Exit Function
    End If
    Dim totalMin As Long : totalMin = CLng(Int(totalSec / 60))
    Dim hrs As Long : hrs = Int(totalMin / 60)
    Dim mins As Long : mins = totalMin - hrs * 60
    If hrs > 0 Then
        FormatMinSec = CStr(hrs) & "ч " & Format(mins, "00") & "м"
    Else
        FormatMinSec = CStr(mins) & "м"
    End If
End Function

'-------------------------------------------------------------
' Вспомогательная: абсолютное время строки в секундах от 01.01.1900
' Используется для надёжного расчёта Δt через границы суток и файлов
'-------------------------------------------------------------
Function RowAbsSeconds(ws As Worksheet, r As Long) As Double
    RowAbsSeconds = 0

    Dim dv As Variant, tv As Variant
    dv = ws.Cells(r, 1).Value  ' столбец A — дата
    tv = ws.Cells(r, 2).Value  ' столбец B — время

    ' Дата как число Excel (целая часть = дни от 01.01.1900)
    Dim dateDays As Double
    dateDays = 0
    If IsNumeric(dv) And CLng(CDbl(dv)) > 0 Then
        dateDays = CLng(CDbl(dv))
    ElseIf IsDate(dv) Then
        dateDays = CDbl(CDate(dv))
    End If

    ' Время как дробь Excel (0..1) или строка HH:MM:SS
    Dim timeFrac As Double
    timeFrac = 0
    If IsNumeric(tv) Then
        timeFrac = CDbl(tv)
    ElseIf InStr(CStr(tv), ":") > 0 Then
        On Error Resume Next
        timeFrac = CDbl(TimeValue(Trim(CStr(tv))))
        On Error GoTo 0
    End If

    If dateDays > 0 Then
        RowAbsSeconds = (dateDays + timeFrac) * 86400#
    End If
End Function

Function CalcDeltaT(ws As Worksheet, r1 As Long, r2 As Long) As Double
    CalcDeltaT = 0

    ' --------------------------------------------------------
    ' Основной метод: абсолютные секунды через дату+время A+B
    ' Работает через границы суток и объединённые файлы
    ' --------------------------------------------------------
    Dim abs1 As Double, abs2 As Double
    abs1 = RowAbsSeconds(ws, r1)
    abs2 = RowAbsSeconds(ws, r2)

    If abs1 > 0 And abs2 > 0 Then
        Dim dtSec As Double
        dtSec = abs2 - abs1
        ' Допустимый диапазон: 1 секунда .. 30 минут
        If dtSec >= 1# And dtSec <= 1800# Then
            CalcDeltaT = dtSec / 60#
            Exit Function
        End If
    End If

    ' --------------------------------------------------------
    ' Резерв: столбец C как мс СУТОК (большие значения > 1000)
    ' --------------------------------------------------------
    Dim ms1v As Variant, ms2v As Variant
    ms1v = ws.Cells(r1, 3).Value
    ms2v = ws.Cells(r2, 3).Value
    If IsNumeric(ms1v) And IsNumeric(ms2v) Then
        Dim msV1 As Double, msV2 As Double
        msV1 = CDbl(ms1v) : msV2 = CDbl(ms2v)
        If msV1 > 1000 Then  ' это абсолютные мс суток
            Dim diffSec As Double
            diffSec = (msV2 - msV1) / 1000#
            If diffSec < 0 Then  ' переход суток
                Dim da1 As Variant, da2 As Variant
                da1 = ws.Cells(r1, 1).Value : da2 = ws.Cells(r2, 1).Value
                If IsNumeric(da1) And IsNumeric(da2) Then
                    diffSec = diffSec + (CLng(CDbl(da2)) - CLng(CDbl(da1))) * 86400#
                Else
                    diffSec = diffSec + 86400#
                End If
            End If
            If diffSec >= 1# And diffSec <= 1800# Then
                CalcDeltaT = diffSec / 60#
                Exit Function
            End If
        End If
    End If
End Function

'-------------------------------------------------------------
' Определение циклов стерилизации и расчёт F0
' Алгоритм ДВУХПРОХОДНЫЙ:
'   Проход 1: определяем границы циклов и MAX(столбец K) → Tref
'   Проход 2: считаем F0 = Σ 10^((T-Tref)/z) * Δt  (прямая формула Бигелоу)
'-------------------------------------------------------------
Sub DetectCyclesAndCalculateF0(wsData As Worksheet, wsReport As Worksheet, lastRow As Long, prevInsertedRows As Long)
    Const COL_DATE As Integer = 1
    Const COL_TIME As Integer = 2
    Const COL_TEMP_PROD As Integer = 5
    Const COL_PRESSURE As Integer = 6   ' F — Давление
    Const COL_WATER As Integer = 7      ' G — Уровень воды
    Const COL_TREF As Integer = 11
    Const COL_F0 As Integer = 18

    ' Пороги начала/конца цикла — ТОЛЬКО по давлению (вода может не сбрасываться)
    ' Старт: давление >= 6.0 бар
    ' Конец: давление < 6.0 бар — 5 строк подряд (защита от кратких колебаний)
    Const CYCLE_THRESH As Double = 6#   ' порог давления
    Const END_COUNT As Integer = 5      ' строк подряд ниже порога = конец цикла
    ' Минимальная длина цикла — интервал ~10 сек, строк ~6/мин
    ' Цикл менее 5 минут = менее ~30 строк — явно ложное срабатывание
    Const MIN_CYCLE_ROWS As Integer = 30

    ' ================================================================
    ' ПРОХОД 1: находим границы циклов ТОЛЬКО ПО ДАВЛЕНИЮ
    '   Старт: давление >= 6.0 бар
    '   Конец: давление < 6.0 бар — 5 строк подряд
    ' ================================================================
    Dim cycleCount As Integer
    cycleCount = 0
    Dim cycleStarts(1 To 500) As Long
    Dim cycleEnds_(1 To 500) As Long
    Dim cycleTref(1 To 500) As Double

    Dim inCyc As Boolean
    Dim cyStart As Long
    Dim endCount As Integer
    Dim tKmaxP1 As Double
    Dim p As Long
    Dim wRv As Variant, prRv As Variant
    Dim pressLvl As Double
    Dim cycleActive As Boolean
    Dim cycEnd1 As Boolean
    Dim realEnd As Long
    Dim tkRvP1 As Variant
    Dim tkVP1 As Double
    inCyc = False : endCount = 0 : tKmaxP1 = 0

    For p = 2 To lastRow
        wRv  = wsData.Cells(p, COL_WATER).Value
        prRv = wsData.Cells(p, COL_PRESSURE).Value

        pressLvl = IIf(IsNumeric(prRv), CDbl(prRv), 0)

        ' Признак "активного" состояния — только по давлению
        cycleActive = (pressLvl >= CYCLE_THRESH)

        If Not inCyc Then
            If cycleActive Then
                inCyc = True
                cyStart = p
                endCount = 0
                tKmaxP1 = 0
            End If
        Else
            ' Накапливаем MAX заданной температуры из столбца K
            tkRvP1 = wsData.Cells(p, COL_TREF).Value
            If IsNumeric(tkRvP1) Then
                tkVP1 = CDbl(tkRvP1)
                If tkVP1 >= 100# And tkVP1 <= 130# And tkVP1 > tKmaxP1 Then
                    tKmaxP1 = tkVP1
                End If
            End If

            ' Считаем строки где автоклав "неактивен" (оба параметра упали)
            If Not cycleActive Then
                endCount = endCount + 1
            Else
                endCount = 0
            End If

            cycEnd1 = (endCount >= END_COUNT) Or (p = lastRow)

            If cycEnd1 Then
                If endCount >= END_COUNT Then
                    realEnd = p - endCount + 1  ' первая строка падения
                Else
                    realEnd = p
                End If
                If realEnd < cyStart Then realEnd = cyStart

                ' Пропускаем слишком короткие циклы — ложные срабатывания
                If (realEnd - cyStart + 1) < MIN_CYCLE_ROWS Then
                    inCyc = False : endCount = 0 : tKmaxP1 = 0
                    GoTo P1Next
                End If

                ' Пропускаем цикл целиком из prevDay — он войдёт в следующий цикл основного файла
                If (prevInsertedRows > 0) And (realEnd <= prevInsertedRows + 1) Then
                    inCyc = False : endCount = 0 : tKmaxP1 = 0
                    GoTo P1Next
                End If

                If cycleCount < 500 Then
                    cycleCount = cycleCount + 1
                    cycleStarts(cycleCount) = cyStart
                    cycleEnds_(cycleCount) = realEnd
                    ' Tref всегда 121.1°C по ГОСТ (Clostridium botulinum)
                    cycleTref(cycleCount) = T_REF
                End If
                inCyc = False : endCount = 0 : tKmaxP1 = 0
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

    ' Массив для хранения примечаний (для таблицы 2) — только прошедшие фильтр 60 мин
    ReDim noteArr(1 To 500) As String
    ReDim resultArr(1 To 500) As String
    ReDim resultColorArr(1 To 500) As Long
    Dim validCycleCount As Integer : validCycleCount = 0

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

        ' Результат для отчёта — длительность через абсолютные секунды (работает через сутки)
        Dim cycleNum As Integer : cycleNum = ci
        Dim durationMin As Double
        Dim absStart As Double, absEnd As Double
        absStart = RowAbsSeconds(wsData, rStart)
        absEnd   = RowAbsSeconds(wsData, rEnd)
        If absStart > 0 And absEnd > 0 And absEnd > absStart Then
            durationMin = (absEnd - absStart) / 60#
        Else
            durationMin = 0
        End If

        ' Цикл короче 60 минут — пропускаем, не добавляем в отчёт
        If durationMin < 60# Then GoTo NextCycle

        Dim tMin As Double, tMax As Double
        tMax = tMaxC

        ' Максимальная температура не превысила 50°C — только прогрев, не добавляем в отчёт
        If tMax <= 50# Then GoTo NextCycle
        tMin = IIf(tMinC < 999, tMinC, 0)

        Dim f0Cycle As Double : f0Cycle = f0C
        Dim tRefCycle As Double : tRefCycle = tRefC

        Dim result As String, resultBgColor As Long, resultFontColor As Long, noteText As String

        ' Норма F0 по ГОСТ — единая для всех программ
        Dim f0Norm As Double : f0Norm = 8#
        Dim trefStr As String : trefStr = "Tref=121.1C (ГОСТ), z=10"

        If Not peakC Then
            result = "— Без стерилизации"
            resultBgColor = RGB(220, 225, 230)
            resultFontColor = RGB(60, 60, 60)
            noteText = "Пик T < 100C — только прогрев"
        ElseIf f0Cycle >= f0Norm Then
            result = "OK НОРМА (F0 >= " & f0Norm & ")"
            resultBgColor = RGB(0, 185, 90)      ' зелёный фон
            resultFontColor = RGB(0, 0, 0)        ' чёрные буквы
            noteText = trefStr
        ElseIf f0Cycle >= f0Norm / 2 Then
            result = "! ПРЕДЕЛ (F0 >= " & f0Norm / 2 & ")"
            resultBgColor = RGB(255, 220, 0)      ' жёлтый фон
            resultFontColor = RGB(0, 0, 0)        ' чёрные буквы
            noteText = trefStr
        Else
            result = "X НЕДОСТАТОЧНО (F0 < " & f0Norm / 2 & ")"
            resultBgColor = RGB(210, 30, 30)      ' красный фон
            resultFontColor = RGB(255, 255, 255)  ' белые буквы
            noteText = trefStr
        End If

        Dim startDateStr As String, endDateStr As String
        startDateStr = FormatDateTime_FromRow(wsData, rStart)
        endDateStr   = FormatDateTime_FromRow(wsData, rEnd)

        ' Форматируем длительность как "X ч YY м"
        Dim durHours As Long, durMins As Long
        durHours = CLng(Int(durationMin / 60))
        durMins  = CLng(Int(durationMin - durHours * 60))
        Dim durStr As String
        If durHours > 0 Then
            durStr = durHours & " ч " & Format(durMins, "00") & " м"
        Else
            durStr = durMins & " м"
        End If

        With wsReport
            .Cells(reportRow, 1).Value = ci
            .Cells(reportRow, 2).Value = startDateStr
            .Cells(reportRow, 3).Value = endDateStr
            .Cells(reportRow, 4).Value = durStr
            .Cells(reportRow, 5).Value = Round(tMax, 2)
            .Cells(reportRow, 6).Value = Round(f0Cycle, 4)
            .Cells(reportRow, 7).Value = result

            .Cells(reportRow, 6).NumberFormat = "0.0000"

            If ci Mod 2 = 0 Then
                .Range(.Cells(reportRow, 1), .Cells(reportRow, 6)).Interior.Color = RGB(240, 246, 252)
            Else
                .Range(.Cells(reportRow, 1), .Cells(reportRow, 6)).Interior.Color = RGB(255, 255, 255)
            End If
            .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Font.Color = RGB(30, 30, 30)

            ' Ячейка результата — цветной фон + контрастный шрифт
            .Cells(reportRow, 7).Interior.Color = resultBgColor
            .Cells(reportRow, 7).Font.Color = resultFontColor
            .Cells(reportRow, 7).Font.Bold = True
        End With

        ' Сохраняем для таблицы 2 (только циклы прошедшие фильтр 60 мин)
        validCycleCount = validCycleCount + 1
        noteArr(validCycleCount) = noteText
        resultArr(validCycleCount) = result
        resultColorArr(validCycleCount) = resultBgColor

        reportRow = reportRow + 1
NextCycle:
    Next ci

    Call AddSummaryRow(wsReport, reportRow, validCycleCount, validCycleCount, noteArr, resultArr, resultColorArr)
End Sub

'-------------------------------------------------------------
' Итоговая строка + Таблица 2 (примечания/программа)
'-------------------------------------------------------------
Sub AddSummaryRow(wsReport As Worksheet, reportRow As Integer, totalCycles As Integer, _
    cycleCount As Integer, noteArr() As String, resultArr() As String, resultColorArr() As Long)
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
        .Cells(reportRow, 7).Value = "Всего циклов: " & totalCycles
        .Cells(reportRow, 7).Font.Bold = True
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Interior.Color = RGB(220, 235, 248)
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Font.Color = RGB(20, 60, 100)
        .Cells(reportRow, 1).Font.Color = RGB(0, 100, 180)
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Borders(xlEdgeTop).Color = RGB(0, 100, 180)
        .Range(.Cells(reportRow, 1), .Cells(reportRow, 7)).Borders(xlEdgeTop).Weight = xlMedium
    End With


End Sub

'-------------------------------------------------------------
' Форматирование листа отчёта
'-------------------------------------------------------------
Sub FormatReportSheet(wsReport As Worksheet)
    With wsReport
        ' Ширины под таблицу 1 (8 колонок, умещается на А4 с узкими полями)
        .Columns(1).ColumnWidth = 6    ' Цикл
        .Columns(2).ColumnWidth = 19   ' Начало
        .Columns(3).ColumnWidth = 19   ' Конец
        .Columns(4).ColumnWidth = 12   ' Длительность
        .Columns(5).ColumnWidth = 11   ' T макс
        .Columns(6).ColumnWidth = 11   ' F0
        .Columns(7).ColumnWidth = 28   ' Результат
        .Columns(1).HorizontalAlignment = xlCenter
        .Columns(4).HorizontalAlignment = xlCenter
        .Columns(5).HorizontalAlignment = xlCenter
        .Columns(6).HorizontalAlignment = xlCenter
        ' Белый фон, чёрный шрифт
        .Cells.Interior.Color = RGB(255, 255, 255)
        .Cells.Font.Color = RGB(0, 0, 0)
        ' Заголовок
        .Cells(1, 1).Font.Color = RGB(0, 100, 180)
        .Cells(1, 1).Font.Size = 14
        .Range("A1:G1").Merge
        .Range("A2:G2").Merge
        .Range("A3:G3").Merge
        .Rows(2).Font.Color = RGB(80, 80, 80)
        .Rows(3).Font.Color = RGB(80, 80, 80)
        ' Поля печати — узкие для А4
        With .PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.4)
            .RightMargin = Application.InchesToPoints(0.4)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .Zoom = False
        End With
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
    tRefC As Double, topOffset As Long, csvFileName As String)

    ' Размер под печать A4 (portrait 96dpi ≈ 794×1123px, поля ~60px)
    ' Один график = одна страница: высота не более ~490pt чтобы не залезать на следующий лист
    Const CHART_W As Long = 790   ' ширина графика (pt)
    Const CHART_H As Long = 510   ' высота графика (pt) — увеличена на 20pt для двух шкал времени внизу

    ' Добавляем ~30 минут строк после конца цикла чтобы было видно спуск температуры
    ' CSV пишется каждые ~10 сек → 30 мин = ~180 строк
    Const EXTRA_ROWS As Long = 180
    Dim rEndExt As Long
    rEndExt = rEnd + EXTRA_ROWS
    If rEndExt > wsData.UsedRange.Rows.Count Then rEndExt = wsData.UsedRange.Rows.Count

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add( _
        Left:=5, Top:=topOffset, Width:=CHART_W, Height:=CHART_H)

    Dim cht As Chart
    Set cht = co.Chart
    cht.ChartType = xlLine
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Dim nRows As Long : nRows = rEndExt - rStart + 1

    ' ================================================================
    ' Расчёт фаз цикла по температуре среды (столбец D):
    '   Нагрев     — от rStart до первого момента T >= tRefC - 2
    '   Удержание  — пока T >= tRefC - 2
    '   Охлаждение — после последнего момента T >= tRefC - 2 до rEnd
    ' Время фаз вычисляется как разность абсолютных секунд строк
    ' ================================================================
    Dim phaseHeatSec As Double  : phaseHeatSec = 0
    Dim phaseHoldSec As Double  : phaseHoldSec = 0
    Dim phaseCoolSec As Double  : phaseCoolSec = 0

    Dim tHoldMin As Double : tHoldMin = tRefC - 2#   ' порог удержания

    ' Если tRefC завышен для данного цикла — найдём реальный максимум T среды
    ' и опустим порог до tEnvMax - 2°C, чтобы фазы всегда определялись
    Dim tEnvMaxPh As Double : tEnvMaxPh = 0
    Dim phRi As Long
    For phRi = rStart To rEnd
        Dim tEnvChk As Variant : tEnvChk = wsData.Cells(phRi, 4).Value
        If IsNumeric(tEnvChk) Then
            If CDbl(tEnvChk) > tEnvMaxPh Then tEnvMaxPh = CDbl(tEnvChk)
        End If
    Next phRi
    If tEnvMaxPh > 0 And tEnvMaxPh < tHoldMin Then
        tHoldMin = tEnvMaxPh - 2#
        If tHoldMin < 40# Then tHoldMin = 40#
    End If

    ' Находим первую строку где T среды >= tHoldMin (конец нагрева)
    Dim rHoldStart As Long : rHoldStart = 0
    Dim rHoldEnd   As Long : rHoldEnd   = 0
    For phRi = rStart To rEnd
        Dim tEnvPh As Variant : tEnvPh = wsData.Cells(phRi, 4).Value
        If IsNumeric(tEnvPh) Then
            If CDbl(tEnvPh) >= tHoldMin Then
                If rHoldStart = 0 Then rHoldStart = phRi
                rHoldEnd = phRi
            End If
        End If
    Next phRi

    ' Считаем секунды каждой фазы
    Dim secStart As Double : secStart = RowAbsSeconds(wsData, rStart)
    Dim secEnd   As Double : secEnd   = RowAbsSeconds(wsData, rEnd)
    If rHoldStart > 0 Then
        ' Нагрев: rStart → rHoldStart
        Dim secHoldS As Double : secHoldS = RowAbsSeconds(wsData, rHoldStart)
        Dim secHoldE As Double : secHoldE = RowAbsSeconds(wsData, rHoldEnd)
        If secHoldS > secStart Then phaseHeatSec = secHoldS - secStart
        If secHoldE > secHoldS Then phaseHoldSec = secHoldE - secHoldS
        If secEnd   > secHoldE Then phaseCoolSec = secEnd   - secHoldE
    Else
        ' Фаза удержания не найдена — весь цикл считаем нагревом
        phaseHeatSec = secEnd - secStart
    End If

    ' Функция форматирования секунд → "Xч YYм"
    Dim heatStr As String, holdStr As String, coolStr As String
    heatStr = FormatMinSec(phaseHeatSec)
    holdStr = FormatMinSec(phaseHoldSec)
    coolStr = FormatMinSec(phaseCoolSec)

    ' --- Метки времени на оси X — только длительность цикла (серые) ---
    ' Реальное время рисуется отдельно синими TextBox внутри Chart.Shapes
    Dim timeLabels() As String
    ReDim timeLabels(1 To nRows)
    Dim ri As Long
    Dim secBase As Double : secBase = RowAbsSeconds(wsData, rStart)
    For ri = 1 To nRows
        Dim secCur   As Double : secCur   = RowAbsSeconds(wsData, rStart + ri - 1)
        Dim secElapsed As Double : secElapsed = secCur - secBase
        If secElapsed < 0 Then secElapsed = 0
        Dim elMin As Long : elMin = CLng(Int(secElapsed / 60))
        Dim elHr  As Long : elHr  = CLng(Int(elMin / 60))
        Dim elMn  As Long : elMn  = elMin Mod 60
        timeLabels(ri) = Format(elHr, "00") & ":" & Format(elMn, "00")
    Next ri

    ' --- Данные серий ---
    ' arrF0 — Variant чтобы хранить Empty (разрыв линии до начала накопления)
    Dim arrEnv() As Double, arrProd() As Double, arrTref() As Double
    Dim arrF0() As Variant
    Dim arrPressure() As Variant
    ReDim arrEnv(1 To nRows)
    ReDim arrProd(1 To nRows)
    ReDim arrF0(1 To nRows)
    ReDim arrTref(1 To nRows)
    ReDim arrPressure(1 To nRows)

    Dim f0MaxVal As Double : f0MaxVal = 0
    Dim pressMaxVal As Double : pressMaxVal = 0
    Dim envMaxVal As Double, prodMaxVal As Double

    ' Проход 1: заполняем env/prod/tref/pressure, собираем сырые F0 во временный массив
    Dim arrF0raw() As Double
    ReDim arrF0raw(1 To nRows)
    For ri = 1 To nRows
        Dim rowIdx As Long : rowIdx = rStart + ri - 1
        Dim vEnv As Variant, vProd As Variant, vF0 As Variant, vPress As Variant
        vEnv   = wsData.Cells(rowIdx, 4).Value   ' D — темп.среды
        vProd  = wsData.Cells(rowIdx, 5).Value   ' E — темп.продукта
        vF0    = wsData.Cells(rowIdx, 18).Value  ' R — F0 (накопленный)
        vPress = wsData.Cells(rowIdx, 6).Value   ' F — давление
        arrEnv(ri)      = IIf(IsNumeric(vEnv), CDbl(vEnv), 0)
        arrProd(ri)     = IIf(IsNumeric(vProd), CDbl(vProd), 0)
        arrTref(ri)     = tRefC
        arrF0raw(ri)    = IIf(IsNumeric(vF0), CDbl(vF0), 0)
        arrPressure(ri) = IIf(IsNumeric(vPress), CDbl(vPress), 0)
        If arrF0raw(ri) > f0MaxVal Then f0MaxVal = arrF0raw(ri)
        If IsNumeric(arrPressure(ri)) And CDbl(arrPressure(ri)) > pressMaxVal Then pressMaxVal = CDbl(arrPressure(ri))
        If arrEnv(ri)  > envMaxVal  Then envMaxVal  = arrEnv(ri)
        If arrProd(ri) > prodMaxVal Then prodMaxVal = arrProd(ri)
    Next ri

    ' Проход 1б: обрезаем давление — ищем резкий сброс (падение >50% за 10 точек подряд)
    ' После такого сброса — Empty (линия обрывается, хвост не рисуется)
    Dim pressMax2 As Double : pressMax2 = 0
    Dim pressDropRi2 As Long : pressDropRi2 = 0
    For ri = 1 To nRows
        If IsNumeric(arrPressure(ri)) And CDbl(arrPressure(ri)) > pressMax2 Then
            pressMax2 = CDbl(arrPressure(ri))
        End If
    Next ri
    If pressMax2 > 0 Then
        ' Скользящий поиск: ищем точку где за следующие 10 точек давление упало > 50%
        Dim wnd As Long : wnd = 10
        For ri = 1 To nRows - wnd
            Dim pCur As Double : pCur = IIf(IsNumeric(arrPressure(ri)), CDbl(arrPressure(ri)), 0)
            Dim pFwd As Double : pFwd = IIf(IsNumeric(arrPressure(ri + wnd)), CDbl(arrPressure(ri + wnd)), 0)
            If pCur > pressMax2 * 0.1 And pFwd < pCur * 0.5 Then
                pressDropRi2 = ri + 1
                Exit For
            End If
        Next ri
        If pressDropRi2 > 0 Then
            For ri = pressDropRi2 To nRows
                arrPressure(ri) = Empty
            Next ri
        End If
    End If

    ' Проход 2: находим первую точку где F0 реально начал расти (> 0)
    ' Это точнее чем T>=90, т.к. F0 считается с первого момента накопления
    Dim f0StartRi As Long : f0StartRi = 0
    For ri = 1 To nRows
        If arrF0raw(ri) > 0.0001 Then
            f0StartRi = ri
            Exit For
        End If
    Next ri

    ' Заполняем arrF0: реальные накопленные значения F0 из данных
    ' - до f0StartRi — Empty (линия не рисуется)
    ' - после — реальное накопленное значение (растёт, потом выравнивается)
    For ri = 1 To nRows
        If f0StartRi > 0 And ri >= f0StartRi Then
            arrF0(ri) = arrF0raw(ri)
        Else
            arrF0(ri) = Empty
        End If
    Next ri

    ' Масштаб F0: максимум = реальный макс + 5% отступ, без лишнего пространства
    Dim f0AxisMax As Double
    If f0MaxVal <= 0 Then
        f0AxisMax = 1
    Else
        f0AxisMax = f0MaxVal * 1.05
        ' Округляем до 1 знака для читаемой шкалы
        f0AxisMax = CDbl(Int(f0AxisMax * 10 + 1)) / 10
        If f0AxisMax < 1 Then f0AxisMax = 1
    End If

    ' --- Серия 1: Температура среды — синяя ---
    Dim s1 As Series
    Set s1 = cht.SeriesCollection.NewSeries
    s1.Name = "1. T среды (°C)"
    s1.Values = arrEnv
    s1.XValues = timeLabels
    s1.Format.Line.ForeColor.RGB = RGB(30, 80, 200)
    s1.Format.Line.Weight = 2
    s1.MarkerStyle = xlMarkerStyleNone
    s1.Smooth = False
    s1.AxisGroup = xlPrimary

    ' --- Серия 2: T° продукта — красная, сглаженная ---
    Dim s2 As Series
    Set s2 = cht.SeriesCollection.NewSeries
    s2.Name = "2. T продукта (°C)"
    s2.Values = arrProd
    s2.XValues = timeLabels
    s2.Format.Line.ForeColor.RGB = RGB(210, 30, 30)
    s2.Format.Line.Weight = 2.5
    s2.MarkerStyle = xlMarkerStyleNone
    s2.Smooth = True
    s2.AxisGroup = xlPrimary

    ' --- Серия 3: Температура задан. программы — тёмно-жёлтая пунктирная ---
    Dim s3 As Series
    Set s3 = cht.SeriesCollection.NewSeries
    s3.Name = "3. Температура задан. программы " & Format(tRefC, "0") & "°C"
    s3.Values = arrTref
    s3.XValues = timeLabels
    s3.Format.Line.ForeColor.RGB = RGB(180, 140, 0)
    s3.Format.Line.DashStyle = msoLineDash
    s3.Format.Line.Weight = 1.5
    s3.MarkerStyle = xlMarkerStyleNone
    s3.AxisGroup = xlPrimary

    ' --- Серия 4: F0 — кривая накопления стерилизационного эффекта ---
    ' Масштабируем F0 в нижнюю зону графика (0..F0_BAND°C на левой оси)
    ' Используем ФИКСИРОВАННЫЙ максимум F0_MAX_NORM (30 мин) — чтобы маленькие
    ' значения (0.7 мин) визуально отличались от больших (19.5 мин).
    ' Если реальный F0 превышает F0_MAX_NORM — масштабируем по факту (не обрезаем).
    Const F0_BAND As Double = 20#
    Const F0_MAX_NORM As Double = 30#   ' эталонный максимум шкалы (мин)
    Dim arrF0disp() As Variant
    ReDim arrF0disp(1 To nRows)
    Dim f0Scale As Double
    Dim f0ScaleBase As Double
    f0ScaleBase = IIf(f0MaxVal > F0_MAX_NORM, f0MaxVal, F0_MAX_NORM)
    f0Scale = F0_BAND / f0ScaleBase

    ' Находим последнюю точку где F0 > 0 — после неё ставим Empty
    Dim f0LastRi As Long : f0LastRi = 0
    For ri = nRows To 1 Step -1
        If Not IsEmpty(arrF0(ri)) Then
            If CDbl(arrF0(ri)) > 0.0001 Then
                f0LastRi = ri
                Exit For
            End If
        End If
    Next ri

    ' Находим точку начала плато F0 (где рост < 0.1% от максимума) — для подписи
    Dim f0PlatoRi As Long : f0PlatoRi = f0LastRi
    If f0MaxVal > 0 Then
        Dim prevF0 As Double : prevF0 = 0
        For ri = f0StartRi To f0LastRi
            If Not IsEmpty(arrF0(ri)) Then
                Dim curF0 As Double : curF0 = CDbl(arrF0(ri))
                If curF0 >= f0MaxVal * 0.995 And f0PlatoRi = f0LastRi Then
                    f0PlatoRi = ri
                End If
            End If
        Next ri
    End If

    For ri = 1 To nRows
        If IsEmpty(arrF0(ri)) Or ri > f0LastRi Then
            arrF0disp(ri) = Empty
        Else
            arrF0disp(ri) = CDbl(arrF0(ri)) * f0Scale
        End If
    Next ri

    Dim s4 As Series
    Set s4 = cht.SeriesCollection.NewSeries
    s4.Name = "4. F0 стерил. эффект (мин)"
    s4.Values = arrF0disp
    s4.XValues = timeLabels
    s4.Format.Line.ForeColor.RGB = RGB(100, 200, 80)   ' светло-зелёный
    s4.Format.Line.Weight = 2.5
    s4.MarkerStyle = xlMarkerStyleNone
    s4.Smooth = True
    s4.AxisGroup = xlPrimary

    ' --- Серия 5: Давление — коричневая, вторичная ось Y (бар) ---
    Dim s5 As Series
    Set s5 = cht.SeriesCollection.NewSeries
    s5.Name = "5. Давление (бар)"
    s5.Values = arrPressure
    s5.XValues = timeLabels
    s5.Format.Line.ForeColor.RGB = RGB(130, 70, 20)    ' тёмно-коричневый
    s5.Format.Line.Weight = 1.5
    s5.MarkerStyle = xlMarkerStyleNone
    s5.Smooth = False
    s5.AxisGroup = xlSecondary   ' вторичная ось Y (правая)

    Dim tickStep  As Long
    Dim pa        As PlotArea
    Dim paILeft   As Double
    Dim paIWidth  As Double
    Dim paITop    As Double
    Dim paIHeight As Double
    Dim ckStepPt  As Double
    Dim ckTki     As Long
    Dim ckTvB     As Variant
    Dim ckAbsSec  As Double
    Dim ckTotalSec As Double
    Dim ckAllMin  As Long
    Dim ckHH      As Long
    Dim ckMM      As Long
    Dim ckLabel   As String
    Dim ckXpos    As Double
    Dim ckYpos    As Double
    Dim ckW       As Double
    Dim ckTb      As Shape
    Dim ckTV2        As Date
    Dim ckLastAdded  As Long
    Dim grayYpos     As Double
    Dim blueYpos     As Double
    Dim xPt          As Double
    Dim xPtLast      As Double
    Dim xPtPrev      As Double
    Dim grayLabel    As String
    Dim pressAxisMax As Double
    Dim heatMidRi    As Long
    Dim sr1 As Series, sr2 As Series, sr4 As Series, sr5 As Series
    Dim ptCount1 As Long, ptCount2 As Long, ptCount4 As Long, ptCount5 As Long
    Dim lbl1Pt As Long, lbl2Pt As Long, lbl5Pt As Long
    Dim pt1 As Point, pt2 As Point, pt4start As Point, pt5 As Point
    Dim f0MidPt As Long
    Dim platoPoint As Point
    Dim coolMidRi  As Long
    Dim sr1r As Series, sr2r As Series, sr5r As Series
    Dim ptCnt1r As Long, ptCnt2r As Long, ptCnt5r As Long
    Dim lbl1r As Long, lbl2r As Long, lbl5r As Long
    Dim pt1r As Point, pt2r As Point, pt5r As Point
    Dim cycDateVal As Variant
    Dim cycDateStr As String
    Dim tbDate As Shape
    Dim totalCycleSec As Double
    Dim totalStr As String
    Dim phaseText As String
    Dim tbLeft As Double
    Dim tbTop  As Double
    Dim tb As Shape
    Dim holdStart2 As Integer
    Dim coolStart2 As Integer
    Dim totalStart As Integer
    Dim tbLabels As Shape
    Dim tbValues As Shape
    Dim tbFileNames As Shape
    Dim chartTitleStr As String
    Dim ctLeft As Double
    Dim ctTop  As Double
    Dim sr3 As Series, pt3 As Point, ptCount3 As Long, lbl3Pt As Long
    Dim tbStats As Shape, tbStatsLabels As Shape, tbStatsValues As Shape, tbF0 As Shape
    Dim f0TbX As Double
    Dim statsLeft As Double, statsTop As Double

    ' Вычисляем дату заранее — нужна для заголовка графика
    cycDateVal = wsData.Cells(rStart, 1).Value
    If IsDate(cycDateVal) Then
        cycDateStr = Format(CDate(cycDateVal), "DD.MM.YYYY")
    ElseIf IsNumeric(cycDateVal) And CLng(CDbl(cycDateVal)) > 1000 Then
        cycDateStr = Format(CDate(CDbl(cycDateVal)), "DD.MM.YYYY")
    Else
        cycDateStr = CStr(cycDateVal)
    End If

    ' Заголовок: дата + Термограмма цикл N
    chartTitleStr = cycDateStr & "     Термограмма  цикл " & cycleIdx

    With cht
        .HasTitle = True
        .ChartTitle.Text = chartTitleStr
        .ChartTitle.Font.Size = 11
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(20, 20, 20)

        .PlotArea.Interior.Color = RGB(255, 255, 255)
        .PlotArea.Border.LineStyle = xlContinuous
        .PlotArea.Border.Color = RGB(180, 200, 220)
        .ChartArea.Interior.Color = RGB(250, 252, 255)
        .ChartArea.Border.Color = RGB(180, 200, 220)

        ' Ось Y левая — температура (°C)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Температура (°C)"
            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Color = RGB(30, 80, 200)
            .MinimumScale = 0
            .MaximumScale = 150
            .MajorUnit = 25
            ' Горизонтальная сетка — мелкий пунктир
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 210, 225)
            .MajorGridlines.Format.Line.DashStyle = msoLineSysDash
            .MajorGridlines.Format.Line.Weight = 0.5
            .HasMinorGridlines = True
            .MinorGridlines.Format.Line.ForeColor.RGB = RGB(230, 235, 242)
            .MinorGridlines.Format.Line.DashStyle = msoLineSysDot
            .MinorGridlines.Format.Line.Weight = 0.25
            .MinorUnit = 5
            .TickLabels.Font.Color = RGB(30, 80, 200)
            .TickLabels.Font.Size = 9
        End With

        ' Правая ось — давление (бар), коричневая
        On Error Resume Next
        .HasAxis(xlValue, xlSecondary) = True
        On Error GoTo 0
        With .Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Давление (мБар)"
            .AxisTitle.Font.Size = 8
            .AxisTitle.Font.Color = RGB(130, 70, 20)
            If pressMaxVal > 0 Then
                pressAxisMax = Application.WorksheetFunction.RoundUp(pressMaxVal * 1.15, 0)
                If pressAxisMax < 4 Then pressAxisMax = 4
            Else
                pressAxisMax = 6
            End If
            .MinimumScale = 0
            .MaximumScale = pressAxisMax
            .MajorUnit = Application.WorksheetFunction.RoundUp(pressAxisMax / 6, 0)
            .HasMajorGridlines = False
            .TickLabels.Font.Color = RGB(130, 70, 20)
            .TickLabels.Font.Size = 9
        End With

        ' === Ось X: скрываем встроенные метки, рисуем ОБЕ шкалы вручную через TextBox ===
        ' Так обе строки всегда точно под линиями сетки и не наезжают друг на друга
        If nRows > 3000 Then
            tickStep = CLng(nRows / 25)
        ElseIf nRows > 600 Then
            tickStep = CLng(nRows / 20)
        ElseIf nRows > 120 Then
            tickStep = CLng(nRows / 15)
        ElseIf nRows > 30 Then
            tickStep = CLng(nRows / 10)
        Else
            tickStep = 1
        End If
        If tickStep < 1 Then tickStep = 1

        With .Axes(xlCategory)
            .HasTitle = False
            ' Скрываем встроенные метки оси X — рисуем их вручную ниже
            .TickLabels.Font.Color = RGB(255, 255, 255)  ' белый = невидимый
            .TickLabels.Font.Size = 1
            .TickLabelSpacing = tickStep
            .TickMarkSpacing = tickStep
            .HasMajorGridlines = True
            .MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 210, 225)
            .MajorGridlines.Format.Line.DashStyle = msoLineSysDash
            .MajorGridlines.Format.Line.Weight = 0.5
            .MajorTickMark = xlTickMarkCross
        End With

        ' Уменьшаем область построения снизу — освобождаем ~45pt для двух шкал времени
        Set pa = .PlotArea
        On Error Resume Next
        pa.Height = pa.Height - 30
        On Error GoTo 0

        paILeft  = pa.InsideLeft
        paIWidth = pa.InsideWidth
        paITop   = pa.InsideTop
        paIHeight = pa.InsideHeight

        If nRows > 1 Then
            ckStepPt = paIWidth / (nRows - 1) * tickStep
        Else
            ckStepPt = paIWidth
        End If
        ckW = ckStepPt - 1
        If ckW < 32 Then ckW = 32
        If ckW > 52 Then ckW = 52

        ' Шкала 1 (синяя, реальное время): сразу под линиями графика
        blueYpos = paITop + paIHeight + 3
        ' Шкала 2 (серая, длительность цикла): на 18pt ниже синей — с зазором между ними
        grayYpos = paITop + paIHeight + 21

        ckLastAdded = 0
        For ckTki = 1 To nRows Step tickStep
            xPt = paILeft + (ckTki - 1) * (paIWidth / (nRows - 1)) - ckW / 2

            ' --- Серая метка: длительность от начала цикла (00:00, 00:11, ...) ---
            grayLabel = timeLabels(ckTki)
            Set ckTb = .Shapes.AddTextbox(msoTextOrientationHorizontal, xPt, grayYpos, ckW, 12)
            With ckTb
                .Line.Visible = msoFalse : .Fill.Visible = msoFalse
                With .TextFrame2.TextRange
                    .Text = grayLabel
                    .Font.Size = 7
                    .ParagraphFormat.Alignment = msoAlignCenter
                    .Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
                    .Font.Bold = False
                End With
                .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
            End With

            ' --- Синяя метка: реальное время суток (через абс.секунды, чтобы правильно через полночь) ---
            ckAbsSec = RowAbsSeconds(wsData, rStart + ckTki - 1)
            ckLabel = ""
            If ckAbsSec > 0 Then
                ckTotalSec = Int(ckAbsSec) - Int(Int(ckAbsSec) / 86400) * 86400
                ckHH    = CLng(Int(ckTotalSec / 3600))
                ckMM    = CLng(Int((ckTotalSec - ckHH * 3600) / 60))
                ckLabel = Format(ckHH, "00") & ":" & Format(ckMM, "00")
            End If
            If ckLabel <> "" Then
                Set ckTb = .Shapes.AddTextbox(msoTextOrientationHorizontal, xPt, blueYpos, ckW, 12)
                With ckTb
                    .Line.Visible = msoFalse : .Fill.Visible = msoFalse
                    With .TextFrame2.TextRange
                        .Text = ckLabel
                        .Font.Size = 7
                        .ParagraphFormat.Alignment = msoAlignCenter
                        .Font.Fill.ForeColor.RGB = RGB(0, 80, 200)
                        .Font.Bold = False
                    End With
                    .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                    .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
                End With
            End If
            ckLastAdded = ckTki
        Next ckTki

        ' Последняя точка — если не попала в Step
        If ckLastAdded < nRows Then
            xPtLast = paILeft + paIWidth - ckW / 2
            ' Защита от слипания: пропускаем если < 80% ширины метки от предыдущей
            xPtPrev = paILeft + (ckLastAdded - 1) * (paIWidth / (nRows - 1)) - ckW / 2
            If xPtLast - xPtPrev >= ckW * 0.8 Then
                ' Серая последняя
                Set ckTb = .Shapes.AddTextbox(msoTextOrientationHorizontal, xPtLast, grayYpos, ckW, 12)
                With ckTb
                    .Line.Visible = msoFalse : .Fill.Visible = msoFalse
                    With .TextFrame2.TextRange
                        .Text = timeLabels(nRows)
                        .Font.Size = 7
                        .ParagraphFormat.Alignment = msoAlignCenter
                        .Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
                        .Font.Bold = False
                    End With
                    .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                    .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
                End With
                ' Синяя последняя (через абс.секунды — правильно через полночь)
                ckAbsSec = RowAbsSeconds(wsData, rStart + nRows - 1)
                ckLabel = ""
                If ckAbsSec > 0 Then
                    ckTotalSec = Int(ckAbsSec) - Int(Int(ckAbsSec) / 86400) * 86400
                    ckHH    = CLng(Int(ckTotalSec / 3600))
                    ckMM    = CLng(Int((ckTotalSec - ckHH * 3600) / 60))
                    ckLabel = Format(ckHH, "00") & ":" & Format(ckMM, "00")
                End If
                If ckLabel <> "" Then
                    Set ckTb = .Shapes.AddTextbox(msoTextOrientationHorizontal, xPtLast, blueYpos, ckW, 12)
                    With ckTb
                        .Line.Visible = msoFalse : .Fill.Visible = msoFalse
                        With .TextFrame2.TextRange
                            .Text = ckLabel
                            .Font.Size = 7
                            .ParagraphFormat.Alignment = msoAlignCenter
                            .Font.Fill.ForeColor.RGB = RGB(0, 80, 200)
                            .Font.Bold = False
                        End With
                        .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                        .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
                    End With
                End If
            End If
        End If

        ' Легенда справа — с номерами линий читается и на ч/б распечатке
        .HasLegend = True
        .Legend.Interior.Color = RGB(255, 255, 255)
        .Legend.Font.Color = RGB(30, 30, 30)
        .Legend.Font.Size = 8
        .Legend.Position = xlLegendPositionRight

        ' ===== Метки-цифры на линиях =====
        ' Принцип расстановки: каждая цифра ставится на ДРУГОЙ точке (разнесены по X),
        ' и с той стороны линии (Above/Below/Left/Right), куда линия НЕ идёт — чтобы
        ' метка не накладывалась на соседние линии.
        '
        ' Зона нагрева:
        '   "5" (давление, коричневая)  — самая ранняя точка (20% нагрева), справа
        '   "1" (T среды, синяя)        — 40% нагрева, выше  (синяя линия ниже красной)
        '   "3" (Tref, пунктир жёлтый)  — ранняя фиксированная точка, выше
        '   "2" (T продукта, красная)   — 65% нагрева, ниже  (красная выше синей)
        '   "4" (F0, зелёная)           — первая ненулевая точка F0, ниже

        ' Середина нагрева (индекс точки) — rHoldStart относительно rStart
        If rHoldStart > rStart Then
            heatMidRi = CLng((rHoldStart - rStart) / 2)
        Else
            heatMidRi = CLng(nRows / 4)
        End If
        If heatMidRi < 1 Then heatMidRi = 1

        ' --- Серия 5 (Давление, коричневая): цифра "5" очень рано (20% нагрева), справа от линии ---
        On Error Resume Next
        If .SeriesCollection.Count >= 5 Then
            Set sr5 = .SeriesCollection(5)
            ptCount5 = sr5.Points.Count
            If ptCount5 > 0 Then
                lbl5Pt = CLng(heatMidRi * 0.4)
                If lbl5Pt < 1 Then lbl5Pt = 1
                If lbl5Pt > ptCount5 Then lbl5Pt = ptCount5
                sr5.HasDataLabels = False
                Set pt5 = sr5.Points(lbl5Pt)
                pt5.HasDataLabel = True
                With pt5.DataLabel
                    .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                    .NumberFormat = "@" : .Characters.Text = "5"
                    .Font.Size = 18 : .Font.Bold = True
                    .Font.Color = RGB(130, 70, 20)
                    .Position = xlLabelPositionRight
                End With
            End If
        End If
        On Error GoTo 0

        ' --- Серия 1 (T среды, синяя): цифра "1" на 40% нагрева, выше линии ---
        On Error Resume Next
        Set sr1 = .SeriesCollection(1)
        ptCount1 = sr1.Points.Count
        If ptCount1 > 0 Then
            lbl1Pt = CLng(heatMidRi * 0.8)
            If lbl1Pt < 1 Then lbl1Pt = 1
            If lbl1Pt > ptCount1 Then lbl1Pt = ptCount1
            sr1.HasDataLabels = False
            Set pt1 = sr1.Points(lbl1Pt)
            pt1.HasDataLabel = True
            With pt1.DataLabel
                .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                .NumberFormat = "@" : .Characters.Text = "1"
                .Font.Size = 18 : .Font.Bold = True
                .Font.Color = RGB(30, 80, 200)
                .Position = xlLabelPositionAbove
            End With
        End If
        On Error GoTo 0

        ' --- Серия 3 (пунктирная, жёлтая): цифра "3" у самого левого края, выше ---
        On Error Resume Next
        Set sr3 = .SeriesCollection(3)
        ptCount3 = sr3.Points.Count
        If ptCount3 > 0 Then
            sr3.HasDataLabels = False
            lbl3Pt = IIf(ptCount3 >= 8, 8, IIf(ptCount3 >= 4, 4, 1))
            Set pt3 = sr3.Points(lbl3Pt)
            pt3.HasDataLabel = True
            With pt3.DataLabel
                .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                .NumberFormat = "@" : .Characters.Text = "3"
                .Font.Size = 18 : .Font.Bold = True
                .Font.Color = RGB(160, 120, 0)
                .Position = xlLabelPositionAbove
            End With
        End If
        On Error GoTo 0

        ' --- Серия 2 (T продукта, красная): цифра "2" на 65% нагрева, ниже линии ---
        On Error Resume Next
        Set sr2 = .SeriesCollection(2)
        ptCount2 = sr2.Points.Count
        If ptCount2 > 0 Then
            lbl2Pt = CLng(heatMidRi * 1.3)
            If lbl2Pt < 1 Then lbl2Pt = 1
            If lbl2Pt > ptCount2 Then lbl2Pt = ptCount2
            sr2.HasDataLabels = False
            Set pt2 = sr2.Points(lbl2Pt)
            pt2.HasDataLabel = True
            With pt2.DataLabel
                .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                .NumberFormat = "@" : .Characters.Text = "2"
                .Font.Size = 18 : .Font.Bold = True
                .Font.Color = RGB(210, 30, 30)
                .Position = xlLabelPositionBelow
            End With
        End If
        On Error GoTo 0

        ' --- Серия 4 (F0, зелёная): цифра "4" у первой точки F0, ниже линии ---
        On Error Resume Next
        Set sr4 = .SeriesCollection(4)
        ptCount4 = sr4.Points.Count
        If ptCount4 > 0 Then
            sr4.HasDataLabels = False
            If f0StartRi >= 1 And f0StartRi <= ptCount4 Then
                Set pt4start = sr4.Points(f0StartRi)
                pt4start.HasDataLabel = True
                With pt4start.DataLabel
                    .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                    .NumberFormat = "@" : .Characters.Text = "4"
                    .Font.Size = 18 : .Font.Bold = True
                    .Font.Color = RGB(40, 140, 40)
                    .Position = xlLabelPositionBelow
                End With
            End If
        End If
        On Error GoTo 0

        ' === Дублирующие цифры серий справа — в зоне охлаждения ===
        ' Каждая цифра ставится на РАЗНУЮ точку по X, чтобы не накладываться.
        ' Зона охлаждения: rHoldEnd → rEnd
        '   "1" (синяя)     — 15% охлаждения, слева (линия ещё горизонтальная, только начала падать)
        '   "2" (красная)   — 45% охлаждения, ниже  (красная выше синей при спуске)
        '   "5" (давление)  — 70% охлаждения, ниже  (давление уже упало, линия низко)
        Dim coolLen As Long
        If rHoldEnd > 0 And rEnd > rHoldEnd Then
            coolLen = rEnd - rHoldEnd
            coolMidRi = (rHoldEnd - rStart) + CLng(coolLen * 0.5)
        Else
            coolLen = CLng(nRows * 0.25)
            coolMidRi = CLng(nRows * 0.85)
        End If
        If coolMidRi < 1 Then coolMidRi = 1
        If coolMidRi > nRows Then coolMidRi = nRows

        On Error Resume Next
        ' --- Цифра "1" справа (15% охлаждения) — слева от синей линии ---
        Set sr1r = .SeriesCollection(1)
        ptCnt1r = sr1r.Points.Count
        If ptCnt1r > 0 Then
            If rHoldEnd > 0 And rEnd > rHoldEnd Then
                lbl1r = (rHoldEnd - rStart) + CLng(coolLen * 0.15)
            Else
                lbl1r = CLng(nRows * 0.78)
            End If
            If lbl1r < 1 Then lbl1r = 1
            If lbl1r > ptCnt1r Then lbl1r = ptCnt1r
            Set pt1r = sr1r.Points(lbl1r)
            pt1r.HasDataLabel = True
            With pt1r.DataLabel
                .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                .NumberFormat = "@" : .Characters.Text = "1"
                .Font.Size = 18 : .Font.Bold = True
                .Font.Color = RGB(30, 80, 200)
                .Position = xlLabelPositionLeft
            End With
        End If
        ' --- Цифра "2" справа (45% охлаждения) — ниже красной линии ---
        Set sr2r = .SeriesCollection(2)
        ptCnt2r = sr2r.Points.Count
        If ptCnt2r > 0 Then
            If rHoldEnd > 0 And rEnd > rHoldEnd Then
                lbl2r = (rHoldEnd - rStart) + CLng(coolLen * 0.45)
            Else
                lbl2r = coolMidRi
            End If
            If lbl2r < 1 Then lbl2r = 1
            If lbl2r > ptCnt2r Then lbl2r = ptCnt2r
            Set pt2r = sr2r.Points(lbl2r)
            pt2r.HasDataLabel = True
            With pt2r.DataLabel
                .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                .NumberFormat = "@" : .Characters.Text = "2"
                .Font.Size = 18 : .Font.Bold = True
                .Font.Color = RGB(210, 30, 30)
                .Position = xlLabelPositionBelow
            End With
        End If
        ' --- Цифра "5" (давление) справа (70% охлаждения) — ниже коричневой линии ---
        If .SeriesCollection.Count >= 5 Then
            Set sr5r = .SeriesCollection(5)
            ptCnt5r = sr5r.Points.Count
            If ptCnt5r > 0 Then
                If rHoldEnd > 0 And rEnd > rHoldEnd Then
                    lbl5r = (rHoldEnd - rStart) + CLng(coolLen * 0.7)
                Else
                    lbl5r = CLng(nRows * 0.92)
                End If
                If lbl5r < 1 Then lbl5r = 1
                If lbl5r > ptCnt5r Then lbl5r = ptCnt5r
                Set pt5r = sr5r.Points(lbl5r)
                pt5r.HasDataLabel = True
                With pt5r.DataLabel
                    .ShowValue = False : .ShowSeriesName = False : .ShowLegendKey = False
                    .NumberFormat = "@" : .Characters.Text = "5"
                    .Font.Size = 18 : .Font.Bold = True
                    .Font.Color = RGB(130, 70, 20)
                    .Position = xlLabelPositionBelow
                End With
            End If
        End If
        On Error GoTo 0

        ' ====== ВСЁ ВНУТРИ cht.Shapes — гарантированно печатается ======
        ' Имена файлов — правый угол заголовка, на одном уровне с "Термограмма цикл N"
        ' Позиция внутри графика: верх ChartArea, правее центра
        Set tbFileNames = cht.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, 395, 2, 380, 18)
        With tbFileNames
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            With .TextFrame2.TextRange
                .Text = csvFileName
                .Font.Size = 9
                .Font.Bold = False
                .ParagraphFormat.Alignment = msoAlignRight
                .Font.Fill.ForeColor.RGB = RGB(40, 40, 140)
            End With
            .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 4
            .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
        End With

        ' Табличка максимумов — над легендой справа
        statsLeft = 637
        statsTop  = 28

        ' Рамка фон (4 строки: T среды, T продукта, давление, F0)
        Set tbStats = cht.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, statsLeft, statsTop, 143, 4)
        With tbStats
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(180, 180, 190)
            .Fill.ForeColor.RGB = RGB(250, 252, 255)
            .Fill.Solid
            .TextFrame2.TextRange.Text = ""
            .Height = 84
        End With

        ' Метки
        Set tbStatsLabels = cht.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, statsLeft + 4, statsTop + 3, 85, 78)
        With tbStatsLabels
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            With .TextFrame2.TextRange
                .Font.Name = "Calibri"
                .Font.Size = 8
                .Font.Bold = True
                .Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
                .Text = "Макс. Т среды:" & Chr(10) & "Макс. Т продукта:" & Chr(10) & _
                        "Макс. давление:" & Chr(10) & "F0 стерил. эффект:"
            End With
            .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
            .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
        End With

        ' Значения
        Set tbStatsValues = cht.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, statsLeft + 89, statsTop + 3, 52, 78)
        With tbStatsValues
            .Line.Visible = msoFalse
            .Fill.Visible = msoFalse
            With .TextFrame2.TextRange
                .Font.Name = "Calibri"
                .Font.Size = 8
                .Font.Bold = False
                .Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
                .Text = Format(envMaxVal, "0.0") & " °C" & Chr(10) & _
                        Format(prodMaxVal, "0.0") & " °C" & Chr(10) & _
                        Format(pressMaxVal, "0") & " мБар" & Chr(10) & _
                        Format(f0MaxVal, "0.0") & " мин"
            End With
            .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
            .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
        End With

        ' TextBox "F0=X.X мин" — правее плато F0, над осью X, с белым фоном чтобы не перекрывать линии
        If f0MaxVal > 0 And f0LastRi > 0 And nRows > 1 Then
            f0TbX = paILeft + (f0LastRi - 1) * (paIWidth / (nRows - 1)) + 6
            If f0TbX > paILeft + paIWidth - 65 Then f0TbX = paILeft + paIWidth - 65
            If f0TbX < paILeft Then f0TbX = paILeft
            Set tbF0 = cht.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, f0TbX, paITop + paIHeight - 28, 65, 16)
            With tbF0
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = RGB(180, 220, 180)
                .Line.Weight = 0.5
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = RGB(245, 255, 245)
                .Fill.Solid
                With .TextFrame2.TextRange
                    .Font.Name = "Calibri"
                    .Font.Size = 8
                    .Font.Bold = True
                    .Font.Fill.ForeColor.RGB = RGB(30, 130, 30)
                    .Text = "F0=" & Format(f0MaxVal, "0.0") & " мин"
                End With
                .TextFrame.MarginLeft = 2 : .TextFrame.MarginRight = 2
                .TextFrame.MarginTop = 1  : .TextFrame.MarginBottom = 1
            End With
        End If

        ' Табличка фаз — под легендой (легенда справа ~160pt от правого края)
        ' Координаты внутри ChartArea: X от правого края минус ширина таблички
        If phaseHeatSec > 0 Or phaseHoldSec > 0 Or phaseCoolSec > 0 Then
            totalCycleSec = phaseHeatSec + phaseHoldSec + phaseCoolSec
            totalStr = FormatMinSec(totalCycleSec)

            ' Позиция внутри cht: справа (легенда ~155pt шириной), под легендой
            ctLeft = 635
            ctTop  = 360

            ' Рамка — фон
            Set tb = cht.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, ctLeft, ctTop, 145, 4)
            With tb
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = RGB(180, 180, 190)
                .Fill.ForeColor.RGB = RGB(250, 252, 255)
                .Fill.Solid
                .TextFrame2.TextRange.Text = ""
                If phaseCoolSec > 0 Then
                    .Height = 96
                Else
                    .Height = 82
                End If
            End With

            ' Левый блок — метки (жирные)
            Set tbLabels = cht.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, ctLeft + 4, ctTop + 3, 72, 76)
            With tbLabels
                .Line.Visible = msoFalse
                .Fill.Visible = msoFalse
                With .TextFrame2.TextRange
                    .Font.Name = "Calibri"
                    .Font.Size = 8
                    .Font.Bold = True
                    .Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
                    If phaseCoolSec > 0 Then
                        .Text = "Нагрев:" & Chr(10) & "Удержание:" & Chr(10) & "Охлаждение:" & Chr(10) & Chr(10) & "Общее время:"
                    Else
                        .Text = "Нагрев:" & Chr(10) & "Удержание:" & Chr(10) & Chr(10) & "Общее время:"
                    End If
                End With
                .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
            End With

            ' Правый блок — значения
            Set tbValues = cht.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, ctLeft + 76, ctTop + 3, 65, 76)
            With tbValues
                .Line.Visible = msoFalse
                .Fill.Visible = msoFalse
                With .TextFrame2.TextRange
                    .Font.Name = "Calibri"
                    .Font.Size = 8
                    .Font.Bold = False
                    .Font.Fill.ForeColor.RGB = RGB(20, 20, 20)
                    If phaseCoolSec > 0 Then
                        .Text = heatStr & Chr(10) & holdStr & Chr(10) & coolStr & Chr(10) & Chr(10) & totalStr
                    Else
                        .Text = heatStr & Chr(10) & holdStr & Chr(10) & Chr(10) & totalStr
                    End If
                End With
                .TextFrame.MarginLeft = 0 : .TextFrame.MarginRight = 0
                .TextFrame.MarginTop = 0  : .TextFrame.MarginBottom = 0
            End With
        End If

    End With
End Sub

'-------------------------------------------------------------
' Строит графики по всем циклам на листе График
'-------------------------------------------------------------
Sub BuildTemperatureChart(wb As Workbook, wsData As Worksheet, lastRow As Long, csvFileName As String, prevInsertedRows As Long)
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets("График").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "График"
    ws.Cells.Interior.Color = RGB(245, 248, 252)

    ' Заголовок листа
    ' Строка-заголовок листа убрана — вся информация теперь в заголовке каждого графика


    ' Та же логика обнаружения циклов что и в DetectCyclesAndCalculateF0
    ' Цикл активен: давление >= 6.0 (вода не учитывается)
    Const CY_THRESH As Double = 6#
    Const CY_END_CNT As Integer = 5

    Dim inCyc As Boolean : inCyc = False
    Dim cyStart As Long, cyEnd As Long
    Dim endCntCy As Integer : endCntCy = 0
    Dim tKmaxCy As Double : tKmaxCy = 0
    Dim tProdMaxCy As Double : tProdMaxCy = 0  ' MAX T продукта за цикл
    Dim cycIdx As Integer : cycIdx = 0
    Dim topOffset As Long : topOffset = 30  ' отступ сверху — место для имён файлов над первым графиком

    Dim p As Long
    For p = 2 To lastRow
        Dim wRv As Variant, prRv As Variant
        wRv  = wsData.Cells(p, 7).Value  ' G — уровень воды
        prRv = wsData.Cells(p, 6).Value  ' F — давление
        Dim pressLvlCy As Double
        pressLvlCy = IIf(IsNumeric(prRv), CDbl(prRv), 0)

        Dim cycActCy As Boolean
        cycActCy = (pressLvlCy >= CY_THRESH)

        If Not inCyc Then
            If cycActCy Then
                inCyc = True : cyStart = p : endCntCy = 0 : tKmaxCy = 0 : tProdMaxCy = 0
            End If
        Else
            ' MAX заданной температуры (столбец K) для Tref графика
            Dim tkRvCy As Variant : tkRvCy = wsData.Cells(p, 11).Value
            If IsNumeric(tkRvCy) Then
                Dim tkVCy As Double : tkVCy = CDbl(tkRvCy)
                If tkVCy >= 100# And tkVCy <= 130# And tkVCy > tKmaxCy Then tKmaxCy = tkVCy
            End If
            ' MAX T продукта (столбец E)
            Dim tpRvCy As Variant : tpRvCy = wsData.Cells(p, 5).Value
            If IsNumeric(tpRvCy) Then
                Dim tpVCy As Double : tpVCy = CDbl(tpRvCy)
                If tpVCy > tProdMaxCy Then tProdMaxCy = tpVCy
            End If

            If Not cycActCy Then
                endCntCy = endCntCy + 1
            Else
                endCntCy = 0
            End If

            Dim cycEnd2 As Boolean
            cycEnd2 = (endCntCy >= CY_END_CNT) Or (p = lastRow)
            If cycEnd2 Then
                cyEnd = IIf(endCntCy >= CY_END_CNT, p - endCntCy + 1, p)
                If cyEnd < cyStart Then cyEnd = cyStart

                Dim tRefCy As Double
                If tKmaxCy >= 118# Then
                    tRefCy = 120#
                ElseIf tKmaxCy >= 100# Then
                    tRefCy = 115#
                Else
                    tRefCy = T_REF
                End If

                ' Не строить график если T продукта не поднималась до 100°C — прогрев без стерилизации
                ' Исключение: если цикл начался в prevDay-данных — tProdMaxCy могло не набраться,
                ' поэтому дополнительно проверяем максимум T продукта по всему диапазону cyStart..cyEnd
                Dim tProdCheck As Double : tProdCheck = tProdMaxCy
                If tProdCheck < T_PEAK_STERIL And prevInsertedRows > 0 And cyStart <= prevInsertedRows + 1 Then
                    ' Цикл начался в предыдущем файле — пересчитываем макс T продукта по всему диапазону
                    Dim scanRow As Long
                    For scanRow = 2 To cyEnd
                        Dim scanTv As Variant : scanTv = wsData.Cells(scanRow, 5).Value
                        If IsNumeric(scanTv) Then
                            If CDbl(scanTv) > tProdCheck Then tProdCheck = CDbl(scanTv)
                        End If
                    Next scanRow
                End If

                If tProdCheck >= T_PEAK_STERIL Then
                    ' Пропускаем цикл который целиком принадлежит prevDay-данным:
                    ' он будет объединён со следующим циклом через backRow-поиск
                    Dim isFullyInPrev As Boolean
                    isFullyInPrev = (prevInsertedRows > 0) And (cyEnd <= prevInsertedRows + 1)
                    If Not isFullyInPrev Then
                        cycIdx = cycIdx + 1
                        ' Захватываем нагрев ДО начала давления — ищем строки назад где T среды >= 40°C
                        ' (если prevInsertedRows > 0 — backRow уйдёт в предыдущий файл автоматически)
                        Dim cyStartExt As Long : cyStartExt = cyStart
                        Dim backRow As Long
                        For backRow = cyStart - 1 To 2 Step -1
                            Dim bTv As Variant : bTv = wsData.Cells(backRow, 4).Value
                            If IsNumeric(bTv) Then
                                If CDbl(bTv) < 40# Then Exit For
                            Else
                                Exit For
                            End If
                            cyStartExt = backRow
                        Next backRow
                        ' Если цикл начался в prevDay — cyStartExt не должен уходить раньше строки 2
                        If cyStartExt < 2 Then cyStartExt = 2
                        Call BuildOneCycleChart(ws, wsData, cyStartExt, cyEnd, cycIdx, tRefCy, topOffset, csvFileName)
                        topOffset = topOffset + 530  ' CHART_H(510) + 20pt отступ между графиками
                    End If
                End If

                inCyc = False : endCntCy = 0 : tKmaxCy = 0 : tProdMaxCy = 0
            End If
        End If
ChartNext:
    Next p

    ' Настройка печати: A4 портрет, масштаб 85% — всё влезает включая таблицу фаз под графиком
    With ws.PageSetup
        .Orientation        = xlPortrait
        .PaperSize          = xlPaperA4
        .Zoom               = 85
        .LeftMargin         = Application.CentimetersToPoints(0.8)
        .RightMargin        = Application.CentimetersToPoints(0.8)
        .TopMargin          = Application.CentimetersToPoints(0.8)
        .BottomMargin       = Application.CentimetersToPoints(0.8)
        .CenterHorizontally = True
        .PrintGridlines     = False
        .PrintHeadings      = False
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
    On Error Resume Next
    If IsNumeric(dateVal) And CLng(CDbl(dateVal)) > 0 Then
        dateDbl = CDbl(CLng(CDbl(dateVal)))
    ElseIf IsDate(dateVal) Then
        dateDbl = CDbl(CDate(dateVal))
    End If
    On Error GoTo 0

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

  const downloadBas = () => {
    // Кодируем в Windows-1251 — иначе кириллица в VBA редакторе Excel становится ????????
    const win1251 = new TextEncoder();
    // Таблица перекодировки Unicode → Windows-1251
    const charMap: Record<number, number> = {
      0x0410:0xC0,0x0411:0xC1,0x0412:0xC2,0x0413:0xC3,0x0414:0xC4,0x0415:0xC5,
      0x0416:0xC6,0x0417:0xC7,0x0418:0xC8,0x0419:0xC9,0x041A:0xCA,0x041B:0xCB,
      0x041C:0xCC,0x041D:0xCD,0x041E:0xCE,0x041F:0xCF,0x0420:0xD0,0x0421:0xD1,
      0x0422:0xD2,0x0423:0xD3,0x0424:0xD4,0x0425:0xD5,0x0426:0xD6,0x0427:0xD7,
      0x0428:0xD8,0x0429:0xD9,0x042A:0xDA,0x042B:0xDB,0x042C:0xDC,0x042D:0xDD,
      0x042E:0xDE,0x042F:0xDF,
      0x0430:0xE0,0x0431:0xE1,0x0432:0xE2,0x0433:0xE3,0x0434:0xE4,0x0435:0xE5,
      0x0436:0xE6,0x0437:0xE7,0x0438:0xE8,0x0439:0xE9,0x043A:0xEA,0x043B:0xEB,
      0x043C:0xEC,0x043D:0xED,0x043E:0xEE,0x043F:0xEF,0x0440:0xF0,0x0441:0xF1,
      0x0442:0xF2,0x0443:0xF3,0x0444:0xF4,0x0445:0xF5,0x0446:0xF6,0x0447:0xF7,
      0x0448:0xF8,0x0449:0xF9,0x044A:0xFA,0x044B:0xFB,0x044C:0xFC,0x044D:0xFD,
      0x044E:0xFE,0x044F:0xFF,
      0x0401:0xA8,0x0451:0xB8,0x2014:0x97,0x2013:0x96,0x00AB:0xAB,0x00BB:0xBB,
      0x00B0:0xB0,0x00B1:0xB1,0x00B2:0xB2,0x00B3:0xB3,0x00B5:0xB5,0x00B7:0xB7,
      0x00A0:0xA0,0x00A9:0xA9,0x00AE:0xAE,0x2116:0xB9,0x20AC:0x88,
    };
    const bytes = new Uint8Array(VBA_CODE.length * 2);
    let idx = 0;
    for (let i = 0; i < VBA_CODE.length; i++) {
      const cp = VBA_CODE.charCodeAt(i);
      if (cp < 0x80) {
        bytes[idx++] = cp;
      } else if (charMap[cp] !== undefined) {
        bytes[idx++] = charMap[cp];
      } else {
        bytes[idx++] = 0x3F; // '?' для неизвестных символов
      }
    }
    const blob = new Blob([bytes.subarray(0, idx)], { type: "text/plain;charset=windows-1251" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "AutoclavF0_Module1.bas";
    a.click();
    URL.revokeObjectURL(url);
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
            <div className="flex items-center gap-2">
              <button
                onClick={downloadBas}
                className="flex items-center gap-1.5 px-4 py-2 rounded-sm text-xs font-medium transition-all border bg-[#0a1828] border-[#22c55e]/30 text-[#22c55e] hover:bg-[#061a0e] hover:border-[#22c55e]/50"
                title="Скачать .bas файл в кодировке Windows-1251 — для импорта в редактор VBA"
              >
                <Icon name="Download" size={12} />
                Скачать .bas
              </button>
              <button
                onClick={copyCode}
                title="⚠️ Только для просмотра — для вставки в Excel используйте .bas файл"
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
          </div>
          <div className="px-5 py-1.5 bg-[#0a1410] border-b border-[#1e3050] text-[10px] text-[#f59e0b] flex items-center gap-1.5">
            <Icon name="AlertTriangle" size={11} />
            Для вставки в Excel — только через «Скачать .bas», затем File → Import. Копирование текста даёт ???? из-за кодировки.
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