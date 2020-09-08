Attribute VB_Name = "shuffleSlides"
Sub shuffleSlides()

FirstSlide = 2
LastSlide = ActivePresentation.Slides.Count

Randomize
'generate random number between 2 to 7'
RSN = Int((LastSlide - FirstSlide + 1) * Rnd + FirstSlide)

For i = FirstSlide To LastSlide
ActivePresentation.Slides(i).MoveTo (RSN)
Next i
End Sub
