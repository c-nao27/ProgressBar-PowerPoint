Sub ProgressBar()
    On Error Resume Next
    With ActivePresentation
        
        Dim barHeight As Integer: barHeight = 2
        Dim offsetTop As Integer: offsetTop = 2
        Dim barWidth As Integer: barWidth = 0
        Dim offsetLeft As Integer: offsetLeft = 0
        Dim barColor As Long: barColor = RGB(255, 0, 0)
        
        For i = 1 To .Slides.Count
            .Slides(i).Shapes("ProgressBar").Delete
            
            Set bar = .Slides(i).Shapes.AddShape( _
                Type:=msoShapeRectangle, _
                Left:=offsetLeft, _
                Top:=.PageSetup.SlideHeight - offsetTop, _
                Width:=i * (.PageSetup.SlideWidth - barWidth) / .Slides.Count, _
                height:=barHeight)
                
            bar.Fill.ForeColor.RGB = barColor
            bar.Name = "ProgressBar"
        Next i:

    End With
End Sub
