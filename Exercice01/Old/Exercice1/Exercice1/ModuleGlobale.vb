Module ModuleGlobale
    Public Sub Log_application(txt As String)
        Dim DateStr As String = Now.ToString("yyyy/MM/dd hh:mm:ss ffff : ")
        VarsGlobals.textBoox_Log.IsReadOnly = False
        VarsGlobals.textBoox_Log.AppendText(DateStr + txt + vbCr)
        VarsGlobals.textBoox_Log.ScrollToEnd()
        VarsGlobals.textBoox_Log.IsReadOnly = True
        Debug.WriteLine(DateStr + txt)
    End Sub
End Module
