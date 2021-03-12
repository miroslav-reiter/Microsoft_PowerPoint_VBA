' Zmena jazyka všetkých slajdov PowerPoint

Option Explicit

Public Sub zmenJazykKontrolyGramatiky()
    Dim i As Integer
    Dim j As Integer
    Dim pocetSlajdov As Integer
    Dim pocetTvarov As Integer
    
    pocetSlajdov = ActivePresentation.Slides.Count
    For i = 1 To pocetSlajdov
        pocetTvarov = ActivePresentation.Slides(i).Shapes.Count
        For j = 1 To pocetTvarov
            If (ActivePresentation.Slides(i).Shapes(j).HasTextFrame) Then
				' Zoznam jazykovych kodov - https://msdn.microsoft.com/en-us/library/aa432635.aspx?f=255&MSPPError=-2147217396
                ActivePresentation.Slides(i).Shapes(j).TextFrame.TextRange.LanguageID = msoLanguageIDSlovak
            End If
        Next j
    Next i
    
    MsgBox "Vsetko som prekonvertoval na slovenský jazyk", vbInformation, "Zmena jazyka gramatiky"
    
End Sub
