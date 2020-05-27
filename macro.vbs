Sub ITIP()

'Copies Excel data to Powerpoint file
Dim powerpointApp As PowerPoint.Application
Dim currentSlide As PowerPoint.Slide
Dim presentation As PowerPoint.presentation
Dim presentationPath As String
Dim excelChart As Excel.ChartObject

'Check if Powerpoint is open
On Error Resume Next
Set powerpointApp = GetObject(, "PowerPoint.Application")
On Error GoTo 0

'Open Powerpoint if it is not already
If powerpointApp Is Nothing Then
    Set powerpointApp = New PowerPoint.Application
End If

'Open slideshow
presentationPath = "c:/code/excel/slideshow.pptx"
Set presentation = powerpointApp.Presentations.Open(presentationPath)

'Copy current excel table into powerpoint
ActiveSheet.ListObjects("Table1").Range.Copy
presentation.Slides(presentation.Slides.Count).Shapes.PasteSpecial ppPasteBitmap

End Sub
