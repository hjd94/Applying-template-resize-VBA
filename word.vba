Sub apply_template()
'
    ' This code select chart and applies a template. It then selects revelent the new data and resizes the graph
'

  Dim objShape As InlineShape
  For Each objShape In ActiveDocument.InlineShapes  
    objShape.Chart.ApplyChartTemplate ("PATH")
    objShape.Select
    objShape.LockAspectRatio = False
    objShape.Width = 435
    objShape.Height = 95
  Next
End Sub
