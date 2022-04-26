```vba
Sub picsize()
    Dim shap As InlineShape
    Application.ScreenUpdating = False
    For Each shap In ActiveDocument.InlineShapes
        If shap.Type = wdInlineShapePicture Then
            If shap.Width >= CentimetersToPoints(14.5) Then
                shap.LockAspectRatio = msoCTrue
                shap.Width = CentimetersToPoints(14)
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub

```