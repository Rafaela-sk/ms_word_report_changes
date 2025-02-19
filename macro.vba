Sub ExportToExcelOptimized()
    Dim doc As Document
    Dim rev As Revision
    Dim cmt As Comment
    Dim xlApp As Object
    Dim xlSheet As Object
    Dim row As Integer
    Dim nearestHeading As String
    Dim nearestPara As String
    
    ' Nastavenie dokumentu
    Set doc = ActiveDocument
    
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
    
    row = 2
    
    ' Export revízií
    For Each rev In doc.Revisions
        nearestHeading = GetNearestHeading(rev.Range)
        nearestPara = GetNearestParagraphOrImage(rev.Range)
        
        xlSheet.Cells(row, 1).Value = rev.Author
        xlSheet.Cells(row, 2).Value = rev.Date
        xlSheet.Cells(row, 3).Value = "Zmena"
        xlSheet.Cells(row, 4).Value = Trim(rev.Range.Text)
        xlSheet.Cells(row, 5).Value = nearestHeading
        xlSheet.Cells(row, 6).Value = nearestPara
        row = row + 1
    Next rev
    
    ' Export komentárov
    For Each cmt In doc.Comments
        nearestHeading = GetNearestHeading(cmt.Scope)
        nearestPara = GetNearestParagraphOrImage(cmt.Scope)
        
        xlSheet.Cells(row, 1).Value = cmt.Author
        xlSheet.Cells(row, 2).Value = cmt.Date
        xlSheet.Cells(row, 3).Value = "Komentár"
        xlSheet.Cells(row, 4).Value = Trim(cmt.Range.Text)
        xlSheet.Cells(row, 5).Value = nearestHeading
        xlSheet.Cells(row, 6).Value = nearestPara
        row = row + 1
    Next cmt
    
    MsgBox "Export dokončený", vbInformation, "Hotovo"
End Sub

' Funkcia na získanie najbližšej kapitoly s číslovaním
Function GetNearestHeading(rng As Range) As String
    Dim para As Paragraph
    Dim heading As String
    
    heading = "Neznáma kapitola"
    
    ' Hľadanie najbližšieho nadpisu smerom nahor
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If para.OutlineLevel <= 3 Then
            heading = Trim(para.Range.ListFormat.ListString & " " & para.Range.Text)
        End If
    Next para
    
    GetNearestHeading = heading
End Function

' Funkcia na získanie obsahu odstavca alebo správneho čísla strany obrázka
Function GetNearestParagraphOrImage(rng As Range) As String
    Dim para As Paragraph
    Dim shape As InlineShape
    Dim nearestText As String
    Dim pageNumber As Integer
    
    nearestText = "Neznámy odstavec/obrázok"
    
    ' Hľadanie najbližšieho relevantného odstavca
    For Each para In rng.Document.Paragraphs
        If para.Range.Start > rng.Start Then Exit For
        If Len(Trim(para.Range.Text)) > 10 Then
            nearestText = Trim(para.Range.Text)
            Exit For
        End If
    Next para
    
    ' Skontrolovanie prítomnosti obrázkov v texte
    For Each shape In rng.Document.InlineShapes
        ' *** Oprava: Každý obrázok dostane správne číslo strany ***
        If shape.Range.Start > rng.Start Then Exit For
        
        ' Určenie strany, na ktorej sa obrázok nachádza
        pageNumber = shape.Range.Information(wdActiveEndPageNumber)
        
        ' Ak obrázok nemá popis, uvedieme len stranu
        If shape.AlternativeText = "" Then
            nearestText = "Obrázok na strane " & pageNumber
        Else
            nearestText = "Obrázok: " & shape.AlternativeText & " (strana " & pageNumber & ")"
        End If
        
        ' Dôležité: Uistíme sa, že sa nezachová číslo strany z predchádzajúcich iterácií
        Exit For
    Next shape
    
    GetNearestParagraphOrImage = nearestText
End Function
