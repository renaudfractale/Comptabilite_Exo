Public Class VarsGlobals
    Public Shared BD As Class_BD
    Public Shared textBoox_Log As RichTextBox
    Public Shared Sub Init()
        textBoox_Log = New RichTextBox
        BD = New Class_BD
    End Sub
End Class
