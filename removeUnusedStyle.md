```vb
Sub organizerdelete()

On Error GoTo ErrorHandler
Dim oStyle As Style, i&i = 0
For Each oStyle In ActiveDocument.Styles
        With ActiveDocument.Content.Find
            .ClearFormatting
            .MatchWildcards = False
            .Style = CVar(oStyle.NameLocal)
            .Execute FindText:="", Format:=True
            If Not .Found Then
                Application.OrganizerDelete _
                Source:=ActiveDocument.Path & "\" & ActiveDocument.Name, _
                Name:=oStyle.NameLocal, Object:=wdOrganizerObjectStyles
                i = i + 1
            End If
        End With
Next oStyle
MsgBox "共删除" & i & "未使用样式"

Exit Sub '退出过程

'发生错误时处理
ErrorHandler:
    i = i - 1 '发生一次错误则减1
    Resume Next

End Sub
```