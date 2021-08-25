' carbon image: https://carbon.now.sh/?bg=rgba%2894%2C129%2C172%2C1%29&t=nord&wt=none&l=vb&ds=true&dsyoff=20px&dsblur=68px&wc=true&wa=true&pv=29px&ph=30px&ln=true&fl=1&fm=Source+Code+Pro&fs=14px&lh=133%25&si=false&es=2x&wm=false&code=Sub%2520sankeyCluster%28%29%250A%2520%2520Dim%2520i%2520As%2520Long%252C%2520clstClr%2520As%2520Long%252C%2520clust%2520As%2520String%250A%2520%2520Dim%2520src%2520As%2520Range%252C%2520tgt%2520As%2520Range%252C%2520wgt%2520As%2520Long%252C%2520maxWgt%2520As%2520Long%250A%2520%2520Dim%2520bgnX%2520As%2520Single%252C%2520bgnY%2520As%2520Single%252C%2520endX%2520As%2520Single%252C%2520endY%2520As%2520Single%250A%2520%2520%250A%2520%2520maxWgt%2520%253D%2520Application.WorksheetFunction.Max%28Sheet4.Range%28%2522%2524C%253A%2524C%2522%29%29%250A%2520%2520Sheet12.Rows%28%25224%253A17%2522%29.RowHeight%2520%253D%2520maxWgt%2520%252B%25205%250A%2520%2520%250A%2520%2520For%2520i%2520%253D%25202%2520To%252023%250A%2520%2520%2520%2520wgt%2520%253D%2520Sheet4.Range%28%2522C%2522%2520%2526%2520i%29.Value%250A%2520%2520%2520%2520Set%2520src%2520%253D%2520Sheet12.Range%28Sheet4.Range%28%2522E%2522%2520%2526%2520i%29.Value%29%250A%2520%2520%2520%2520Set%2520tgt%2520%253D%2520Sheet12.Range%28Sheet4.Range%28%2522F%2522%2520%2526%2520i%29.Value%29%250A%2520%2520%2520%2520%250A%2520%2520%2520%2520bgnX%2520%253D%2520src.Left%250A%2520%2520%2520%2520bgnY%2520%253D%2520src.Top%2520%252B%2520%28src.Height%2520%252F%25202%29%250A%2520%2520%2520%2520endX%2520%253D%2520tgt.Left%250A%2520%2520%2520%2520endY%2520%253D%2520tgt.Top%2520%252B%2520%28tgt.Height%2520%252F%25202%29%250A%2520%2520%2520%2520%250A%2520%2520%2520%2520clust%2520%253D%2520Sheet4.Range%28%2522G%2522%2520%2526%2520i%29.Value%250A%2520%2520%2520%2520Select%2520Case%2520clust%250A%2520%2520%2520%2520%2520%2520Case%2520%2522black%2522%253A%2520%2520clstClr%2520%253D%2520VBA.RGB%280%252C%25200%252C%25200%29%250A%2520%2520%2520%2520%2520%2520Case%2520%2522red%2522%253A%2520%2520%2520%2520clstClr%2520%253D%2520VBA.RGB%28255%252C%25200%252C%25200%29%250A%2520%2520%2520%2520%2520%2520Case%2520%2522blue%2522%253A%2520%2520%2520clstClr%2520%253D%2520VBA.RGB%280%252C%25200%252C%2520255%29%250A%2520%2520%2520%2520%2520%2520Case%2520%2522green%2522%253A%2520%2520clstClr%2520%253D%2520VBA.RGB%280%252C%2520255%252C%25200%29%250A%2520%2520%2520%2520%2520%2520Case%2520%2522orange%2522%253A%2520clstClr%2520%253D%2520VBA.RGB%28255%252C%2520192%252C%25200%29%250A%2520%2520%2520%2520%2520%2520Case%2520%2522purple%2522%253A%2520clstClr%2520%253D%2520VBA.RGB%28112%252C%252048%252C%2520160%29%250A%2520%2520%2520%2520End%2520Select%250A%2520%2520%2520%2520%250A%2520%2520%2520%2520%27.AddConnector%28msoConnectorCurve%252C%2520%2522BeginX%2522%252C%2520%2522BeginY%2522%252C%2520%2522EndX%2522%252C%2520%2522EndY%2522%29%250A%2520%2520%2520%2520Sheet12.Shapes.AddConnector%28msoConnectorCurve%252C%2520bgnX%252C%2520bgnY%252C%2520endX%252C%2520endY%29.Select%250A%2520%2520%2520%2520%250A%2520%2520%2520%2520With%2520Selection.ShapeRange.Line%250A%2520%2520%2520%2520%2520%2520.Visible%2520%253D%2520msoTrue%250A%2520%2520%2520%2520%2520%2520.ForeColor.RGB%2520%253D%2520clstClr%250A%2520%2520%2520%2520%2520%2520.Weight%2520%253D%2520wgt%250A%2520%2520%2520%2520%2520%2520.Transparency%2520%253D%25200.5%250A%2520%2520%2520%2520End%2520With%250A%2520%2520Next%2520i%250A%2520%2520Sheet12.Range%28%2522A1%2522%29.Select%250AEnd%2520Sub

Sub sankeyCluster()
  Dim i As Long, clstClr As Long, clust As String
  Dim src As Range, tgt As Range, wgt As Long, maxWgt As Long
  Dim bgnX As Single, bgnY As Single, endX As Single, endY As Single
  
  maxWgt = Application.WorksheetFunction.Max(Sheet4.Range("$C:$C"))
  Rows("4:17").RowHeight = 40
  
  For i = 2 To 23
    wgt = Sheet4.Range("C" & i).Value
    Set src = Sheet12.Range(Sheet4.Range("E" & i).Value)
    Set tgt = Sheet12.Range(Sheet4.Range("F" & i).Value)
    
    bgnX = src.Left
    bgnY = src.Top + (src.Height / 2)
    endX = tgt.Left
    endY = tgt.Top + (tgt.Height / 2)
    
    clust = Sheet4.Range("G" & i).Value
    Select Case clust
      Case "black":  clstClr = VBA.RGB(0, 0, 0)
      Case "red":    clstClr = VBA.RGB(255, 0, 0)
      Case "blue":   clstClr = VBA.RGB(0, 0, 255)
      Case "green":  clstClr = VBA.RGB(0, 255, 0)
      Case "orange": clstClr = VBA.RGB(255, 192, 0)
      Case "purple": clstClr = VBA.RGB(112, 48, 160)
    End Select
    
    '.AddConnector(msoConnectorCurve, "BeginX", "BeginY", "EndX", "EndY")
    Sheet12.Shapes.AddConnector(msoConnectorCurve, bgnX, bgnY, endX, endY).Select
    
    With Selection.ShapeRange.Line
      .Visible = msoTrue
      .ForeColor.RGB = clstClr
      .Weight = wgt
      .Transparency = 0.5
    End With
  Next i
  Sheet12.Range("A1").Select
End Sub
