Sub ExportToExcelOptimized()
    Dim doc As Document
    Dim rev As Revision
    Dim cmt As Comment
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim row As Integer
    Dim nearestHeading As String
    Dim nearestPara As String
    Dim pageNum As Integer
    
    ' Nastavenie dokumentu
    Set doc = ActiveDocument
    
    ' Zlepšenie výkonu
    Application.ScreenUpdating = False
    Application.StatusBar = "Spracovanie dokumentu..."
    
    ' Otvorenie Excelu
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlSheet = xlApp.Workbooks.Add.Sheets(1)
    
    ' Záhlavie tabuľky
    xlSheet.Cells(1, 1).Value = "Autor"
    xlSheet.Cells(1, 2).Value = "Dátum"
    xlSheet.Cells(1, 3).Value = "Typ"
    xlSheet.Cells(1, 4).Value = "Obsah"
    xlSheet.Cells(1, 5).Value = "Kapitola"
    xlSheet.Cells(1, 6).Value = "Odstavec/Obrázok"
    xlSheet.Cells(1, 7).Value = "Strana"
    
    row = 2
    
    ' Export revízií (zmeny)
    For Each rev In doc.Revisions
        nearestHeading = GetNearestHeading(rev.Range)
        nearestPara = GetNearestParagraphOrImage(rev.Range)
        pageNum = rev.Range.Information(wdActiveEndPageNumber)
        
        xlSheet.Cells(row, 1).Value = rev.Author
        xlSheet.Cells(row, 2).Value = rev.Date
        xlSheet.Cells(row, 3).Value = "Zmena"
        xlSheet.Cells(row, 4).Value = Trim(rev.Range.Text)
        xlSheet.Cells(row, 5).Value = nearestHeading
        xlSheet.Cells(row, 6).Value = nearestPara
        xlSheet.Cells(row, 7).Value = pageNum
        row = row + 1
    Next rev
    
    ' Export komentárov
    For Each cmt In doc.Comments
        nearestHeading = GetNearestHeading(cmt.Scope)
        nearestPara = GetNearestParagraphOrImage(cmt.Scope)
        pageNum = cmt.Scope.Information(wdActiveEndPageNumber)
        
        xlSheet.Cells(row, 1).Value = cmt.Author
        xlSheet.Cells(row, 2).Value = cmt.Date
        xlSheet.Cells(row, 3).Value = "Komentár"
        xlSheet.Cells(row, 4).Value = Trim(cmt.Range.Text)
        xlSheet.Cells(row, 5).Value = nearestHeading
        xlSheet.Cells(row, 6).Value = nearestPara
        xlSheet.Cells(row, 7).Value = pageNum
        row = row + 1
    Next cmt
    
    ' Obnovenie obrazovky
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    MsgBox "Export dokončený", vbInformation, "Hotovo"
End Sub

' Funkcia na získanie najbližšej kapitoly
Function GetNearestHeading(rng As Range) As String
    Dim para As Paragraph
    Dim heading As String
    
    heading = "Neznáma kapitola"
    
    ' Prehľadávanie odsekov od miesta revízie smerom nahor
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If para.OutlineLevel <= 3 Then
            heading = Trim(para.Range.Text)
        End If
    Next para
    
    GetNearestHeading = heading
End Function

' Funkcia na získanie najbližšieho relevantného odstavca alebo obrázka
Function GetNearestParagraphOrImage(rng As Range) As String
    Dim para As Paragraph
    Dim shape As InlineShape
    Dim nearestText As String
    Dim pageNumber As Integer
    
    nearestText = "Neznámy odstavec/obrázok"
    
    ' Hľadanie najbližšieho odstavca s dostatočne dlhým textom
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If Len(Trim(para.Range.Text)) > 10 Then
            nearestText = Trim(para.Range.Text)
            Exit For
        End If
    Next para
    
    ' Skontrolovanie prítomnosti obrázkov
    For Each shape In rng.Document.InlineShapes
        If shape.Range.Start > rng.Start Then Exit For
        pageNumber = shape.Range.Information(wdActiveEndPageNumber)
        
        If shape.AlternativeText = "" Then
            nearestText = "Obrázok"
        Else
            nearestText = "Obrázok: " & shape.AlternativeText
        End If
        
        Exit For
    Next shape
    
    GetNearestParagraphOrImage = nearestText
End Function
