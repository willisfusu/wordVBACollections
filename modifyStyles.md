```vba
Sub setshapstyle ()
    On Error Resume Next
    Dim myshape As InlineShape
    Dim figstyle As Style, styname As String
    styname = "Figure"
    Set figstyle = ActiveDocument.Styles(styname)
    
    Application.ScreenUpdating = False
    
    For Each myshape In ActiveDocument.InlineShapes
        With myshape
            If .Type = wdInlineShapePicture Then
                .Range.Style = figstyle
            End If
        End With
    Next
    Application.ScreenUpdating = True
End Sub
```