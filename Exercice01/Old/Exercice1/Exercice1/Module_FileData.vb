Imports System.IO

Public Module Module_FileData

    Public Function OpenFileTxt(path As String) As List(Of Class_line)
        ModuleGlobale.Log_application("Start creation liste init")
        Dim Liste As New List(Of Class_line)
        ModuleGlobale.Log_application("Start creation liste init")
        ModuleGlobale.Log_application("Start Ouverture du fichier : " + path)
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(path)
        ModuleGlobale.Log_application("End Ouverture du fichier : " + path)
        Dim txt As String

        Dim i As Integer = 0
        Do
            txt = reader.ReadLine
            If txt IsNot Nothing And i > 0 Then
                ModuleGlobale.Log_application("Start Analyse de la ligne : " + txt)
                Liste.Add(New Class_line(txt))
                ModuleGlobale.Log_application("End Analyse de la ligne : " + txt)
            End If
            i = i + 1
        Loop Until txt Is Nothing

        ModuleGlobale.Log_application("Start Fermeture du fichier : " + path)
        reader.Close()
        ModuleGlobale.Log_application("End Fermeture du fichier : " + path)

        Return Liste
    End Function
End Module
