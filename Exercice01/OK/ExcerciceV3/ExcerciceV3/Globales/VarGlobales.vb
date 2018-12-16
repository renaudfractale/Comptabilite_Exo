Public Class VarGlobales
    Public Shared textBox_Log As RichTextBox
    Public Shared DataShowIHM As DataShow
    Public Shared ListeComptes As List(Of Class_line)
    Public Shared Sub Init()
        textBox_Log = New RichTextBox
        DataShowIHM = New DataShow
        ListeComptes = New List(Of Class_line)
    End Sub
End Class
