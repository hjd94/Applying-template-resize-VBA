Sub graphing()
'
    ' apply template, reselect revelent data and resize graph
'

'
    a = Application.Worksheets.Count

    For i = 1 To a
    Worksheets(i).Activate
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.ApplyChartTemplate ( _
            "path" _
        )
    ActiveChart.SetSourceData Source:=ActiveSheet.Range("a3:a5,c3:c5")
    ActiveSheet.Shapes("Chart 1").ScaleWidth 1.30, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Chart 1").ScaleHeight 1.39, msoFalse, _
        msoScaleFromTopLeft
    Next
    
End Sub
