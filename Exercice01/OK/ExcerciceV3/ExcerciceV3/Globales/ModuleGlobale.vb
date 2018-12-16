Module ModuleGlobale
    Public Sub Log_application(txt As String)
        Dim DateStr As String = Now.ToString("yyyy/MM/dd hh:mm:ss ffff : ")
        VarGlobales.textBox_Log.IsReadOnly = False
        VarGlobales.textBox_Log.AppendText(DateStr + txt + vbCr)
        VarGlobales.textBox_Log.ScrollToEnd()
        VarGlobales.textBox_Log.IsReadOnly = True
        Debug.WriteLine(DateStr + txt)
    End Sub
End Module
