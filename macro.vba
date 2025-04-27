Sub ExportToExcelUltraFast()
    ' 🛠 Export zmien a komentárov z Wordu do Excelu s korekciou Parent ID
    ' 🛠 Export changes and comments from Word to Excel with Parent ID correction

    ' 🔵 PARAMETRE NASTAVENIA / PARAMETERS AND SETTINGS
    Const FastMode As Boolean = False            ' True = Fast mode (no page number) / Rýchly režim (bez čísla strany)
    Const StatusUpdateFrequency As Long = 1   ' Update status bar every X processed items / Aktualizácia stavového riadku každých X položiek
    Const MaxBackwardSearch As Long = 20         ' Max rows to search backward for Parent Comment / Maximálny počet riadkov na spätné hľadanie Parent ID

    Dim doc As Document
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim data() As Variant
    Dim rowCount As Long
    Dim rev As Revision
    Dim cmt As Comment
    Dim totalItems As Long
    Dim startTime As Double
    Dim commentMap As Object
    Dim currentCommentID As Long
    Dim filePath As String
    Dim fileName As String
    Dim folderPath As String
    Dim i As Long, j As Long, jStart As Long

    ' 🔵 KONTROLA A PRÍPRAVA DOKUMENTU / DOCUMENT CHECK AND PREPARATION
    If Documents.Count = 0 Then
        MsgBox "No document is open. Please open a document and run the macro again." & vbCrLf & _
               "Nie je otvorený žiadny dokument. Otvorte dokument a spustite makro znova.", vbExclamation, "No document"
        Exit Sub
    End If

    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "No active document found." & vbCrLf & _
               "Nebolo nájdené aktívne okno dokumentu.", vbExclamation, "No active document"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Starting processing... / Začínam spracovanie..."

    startTime = Timer
    Set commentMap = CreateObject("Scripting.Dictionary")

    ' 🔵 OTVORENIE EXCELU / OPENING EXCEL
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)

    ' 🔵 HLAVIČKA TABUĽKY / EXCEL HEADER
    xlSheet.Cells(1, 1).Value = "Author / Autor"
    xlSheet.Cells(1, 2).Value = "Date / Dátum"
    xlSheet.Cells(1, 3).Value = "Type / Typ"
    xlSheet.Cells(1, 4).Value = "Content / Obsah"
    xlSheet.Cells(1, 5).Value = "Chapter / Kapitola"
    xlSheet.Cells(1, 6).Value = "Paragraph/Image / Odstavec/Obrázok"
    xlSheet.Cells(1, 7).Value = "Page / Strana"
    xlSheet.Cells(1, 8).Value = "Comment ID"
    xlSheet.Cells(1, 9).Value = "Parent Comment ID"

    ' 🔵 PRÍPRAVA DÁT / DATA PREPARATION
    totalItems = doc.Revisions.Count + doc.Comments.Count
    ReDim data(1 To totalItems, 1 To 9)

    rowCount = 1
    currentCommentID = 1

    ' 🔵 SPRACOVANIE ZMIEN / PROCESSING REVISIONS
    For Each rev In doc.Revisions
        data(rowCount, 1) = rev.Author
        data(rowCount, 2) = rev.Date
        data(rowCount, 3) = "Change / Zmena"
        data(rowCount, 4) = CleanText(rev.Range.Text)
        data(rowCount, 5) = GetNearestHeading(rev.Range)
        data(rowCount, 6) = GetNearestParagraphOrImage(rev.Range)
        If FastMode Then
            data(rowCount, 7) = ""
        Else
            data(rowCount, 7) = rev.Range.Information(wdActiveEndPageNumber)
        End If
        data(rowCount, 8) = ""
        data(rowCount, 9) = ""

        If rowCount Mod StatusUpdateFrequency = 0 Then
            Application.StatusBar = "Processing revisions: " & rowCount & " / " & totalItems
        End If

        rowCount = rowCount + 1
    Next rev

    ' 🔵 SPRACOVANIE KOMENTÁROV A ODPOVEDÍ / PROCESSING COMMENTS AND REPLIES
    For Each cmt In doc.Comments
        data(rowCount, 1) = cmt.Author
        data(rowCount, 2) = cmt.Date

        If cmt.Ancestor Is Nothing Then
            data(rowCount, 3) = "Comment / Komentár"
            data(rowCount, 9) = ""
        Else
            data(rowCount, 3) = "Reply / Reakcia"
            data(rowCount, 9) = "Unknown"
        End If

        data(rowCount, 4) = CleanText(cmt.Range.Text)
        data(rowCount, 5) = GetNearestHeading(cmt.Scope)
        data(rowCount, 6) = GetNearestParagraphOrImage(cmt.Scope)
        If FastMode Then
            data(rowCount, 7) = ""
        Else
            data(rowCount, 7) = cmt.Scope.Information(wdActiveEndPageNumber)
        End If
        data(rowCount, 8) = currentCommentID
        commentMap.Add cmt, currentCommentID

        currentCommentID = currentCommentID + 1

        If rowCount Mod StatusUpdateFrequency = 0 Then
            Application.StatusBar = "Processing comments: " & rowCount & " / " & totalItems
        End If

        rowCount = rowCount + 1
    Next cmt

    ' 🔵 OPRAVA PARENT ID / CORRECTING PARENT COMMENT ID
    For i = 1 To UBound(data)
        If data(i, 9) = "Unknown" Then
            If (i - MaxBackwardSearch) > 1 Then
                jStart = i - MaxBackwardSearch
            Else
                jStart = 1
            End If
            
            For j = i - 1 To jStart Step -1
                If data(j, 3) = "Comment / Komentár" Then
                    If data(j, 5) = data(i, 5) And data(j, 6) = data(i, 6) Then
                        data(i, 9) = data(j, 8)
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    ' 🔵 ZAPIS DÁT DO EXCELU / EXPORTING DATA TO EXCEL
    xlSheet.Range("A2").Resize(UBound(data), UBound(data, 2)).Value = data
    xlSheet.Columns.AutoFit

    ' 🔵 ULOŽENIE SÚBORU / SAVING THE FILE
    On Error Resume Next
    folderPath = doc.Path
    If folderPath = "" Then folderPath = xlApp.GetSaveAsFilename("Exported_Changes_", "Excel Files (*.xlsx), *.xlsx")
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    fileName = "Exported_Changes_" & Format(Now, "yyyymmdd_HHmm") & ".xlsx"
    xlBook.SaveAs folderPath & fileName, 51
    On Error GoTo 0

    xlApp.Visible = True

    ' 🔵 OBNOVENIE WORDU / RESTORING WORD
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""

    MsgBox "Export completed successfully!" & vbCrLf & _
           "Time elapsed: " & Format((Timer - startTime) / 60, "0.00") & " minutes." & vbCrLf & _
           "Saved as: " & folderPath & fileName, vbInformation, "Done"
End Sub

' --- 🔵 PODPORNÉ FUNKCIE / SUPPORT FUNCTIONS ---

Function GetNearestHeading(rng As Range) As String
    Dim para As Paragraph
    Dim heading As String
    heading = "Unknown Chapter / Neznáma kapitola"
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If para.OutlineLevel <= 3 Then
            heading = CleanText(para.Range.Text)
        End If
    Next para
    GetNearestHeading = heading
End Function

Function GetNearestParagraphOrImage(rng As Range) As String
    Dim para As Paragraph
    Dim shape As InlineShape
    Dim nearestText As String
    nearestText = "Unknown Paragraph/Image / Neznámy odstavec/obrázok"
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If Len(CleanText(para.Range.Text)) > 10 Then
            nearestText = CleanText(para.Range.Text)
            Exit For
        End If
    Next para
    For Each shape In rng.Document.InlineShapes
        If shape.Range.Start > rng.Start Then Exit For
        If shape.AlternativeText = "" Then
            nearestText = "Image / Obrázok"
        Else
            nearestText = "Image: " & shape.AlternativeText
        End If
        Exit For
    Next shape
    GetNearestParagraphOrImage = nearestText
End Function

Function CleanText(txt As String) As String
    CleanText = Trim(Replace(Replace(txt, Chr(13), " "), Chr(10), " "))
End Function
