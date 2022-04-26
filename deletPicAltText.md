```vba
Function delPictureByAltText()
    Dim shape As Variant

    Application.ScreenUpdating = False
    For Each shape In ThisDocument.InlineShapes
        shape.AlternativeText = ""
    Next
    Application.ScreenUpdating = True
End Function

```