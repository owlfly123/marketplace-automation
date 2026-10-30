Sub SplitTo300RowsFiles()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim startRow As Long, endRow As Long
    Dim chunkSize As Long
    Dim partNum As Long
    Dim wbNew As Workbook
    Dim basePath As String
    Dim fileName As String
    
    ' ==== НАСТРОЙКИ ====
    chunkSize = 300                 ' размер куска (строк в одном файле)
    Set ws = ThisWorkbook.Sheets(1) ' используем первый лист в книге
    ' ====================
    
    ' Последняя строка по столбцу A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Последний столбец в первой строке
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow = 0 Then
        MsgBox "Данных на листе нет."
        Exit Sub
    End If
    
    ' Папка, где лежит текущий файл
    basePath = ThisWorkbook.Path
    If basePath = "" Then
        basePath = Application.DefaultFilePath
    End If
    
    partNum = 1
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Идём по строкам с шагом chunkSize
    For startRow = 1 To lastRow Step chunkSize
        endRow = Application.WorksheetFunction.Min(startRow + chunkSize - 1, lastRow)
        
        ' Создаём новую книгу
        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        
        ' Копируем блок строк в новую книгу, начиная с A1
        ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, lastCol)).Copy _
            Destination:=wbNew.Sheets(1).Range("A1")
        
        ' Формируем имя файла: ДД.ММ.ГГГГ_N.xlsx
        fileName = basePath & "\" & Format(Date, "dd.mm.yyyy") & "_" & partNum & ".xlsx"
        
        ' Сохраняем новый файл
        wbNew.SaveAs fileName, FileFormat:=xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        
        partNum = partNum + 1
    Next startRow
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Готово. Создано файлов: " & (partNum - 1)
End Sub
