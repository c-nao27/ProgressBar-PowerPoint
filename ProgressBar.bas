Sub ProgressBar()
    On Error Resume Next
    With ActivePresentation
        
        Dim objectName As String: objectName = "ProgressBar"
        Dim barHeight As Integer: barHeight = 2
        Dim offsetVertical As Integer: offsetVertical = 2
        Dim barWidth As Integer: barWidth = 0
        Dim offsetHorizontal As Integer: offsetHorizontal = 0
        Dim barColor As Long: barColor = RGB(255, 0, 0)
        
        For i = 1 To .Slides.Count
            .Slides(i).Shapes(objectName).Delete
            
            Set bar = .Slides(i).Shapes.AddShape( _
                Type:=msoShapeRectangle, _
                Left:=offsetHorizontal, _
                Top:=.PageSetup.SlideHeight - offsetVertical, _
                Width:=i * (.PageSetup.SlideWidth - barWidth) / .Slides.Count, _
                height:=barHeight)
                
            bar.Fill.ForeColor.RGB = barColor
            bar.Line.ForeColor.RGB = barColor
            bar.Name = objectName
        Next i:

    End With
End Sub
