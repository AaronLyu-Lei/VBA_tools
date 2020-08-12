Attribute VB_Name = "Org_Chart"


Sub Organization_Chart()
    Dim saLayOut As SmartArtLayout
    Dim sa As SmartArt
    Dim LastRow As Long
    
    Range("J9").Select
    LastRow = Range("B" & Rows.Count).End(xlUp).Row
    
    Set saLayOut = Application.SmartArtLayouts( _
        "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1")
    Set oshp = ActiveSheet.Shapes.AddSmartArt(saLayOut)
    
    'add nodes to smartart
    If LastRow > 4 Then
        For i = 1 To LastRow - 5
            With oshp
                .Select
                .SmartArt.AllNodes.Add
            End With
        Next i
    Else
        GoTo error
    End If
    
    'put company name and equity ratio to smartart's nodes
    For i = 1 To LastRow
        With oshp
        .Select
            Set sa = .SmartArt
            If Range("E" & i).Value = "" Then
                sa.AllNodes(i).TextFrame2.TextRange.Text = Range("B" & i).Value
            Else
                sa.AllNodes(i).TextFrame2.TextRange.Text = Range("E" & i) & vbNewLine & Range("B" & i).Value
            End If
                
            'At the beginning,set every company to level1
            Do Until sa.AllNodes(i).Level = 1
                sa.AllNodes(i).Promote
            Loop
            'Reset the company to its correct level
            j = VBA.CLng(Trim(Range("D" & i).Value)) - 1
            
            If j = 0 Then
                sa.AllNodes(i).Shapes.Fill.ForeColor.RGB = 15123099
                sa.AllNodes(i).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
            Else
                For n = 1 To j
                    'set the node's level according to its class column D
                    sa.AllNodes(i).Demote
                    sa.AllNodes(i).Shapes.Fill.ForeColor.RGB = 6567712 + j * 1500000
                    sa.AllNodes(i).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0
                Next n
            End If
        End With
    Next i
    
    'adjust the shape
    With oshp
         .ScaleWidth 15, msoFalse, _
        msoScaleFromTopLeft
          .ScaleHeight 10, msoFalse, _
        msoScaleFromTopLeft
    End With
    
    'zoom the spreadsheet
     ActiveWindow.Zoom = 50
     
     Exit Sub

error:
     MsgBox "至少应当有5家单位"
     Exit Sub
    
End Sub


