Imports System.Data
Imports System.IO

Module Module_Tranfert
    Public Function File2Datatable(FileName As String) As System.Data.DataTable
        Dim SR As StreamReader = New StreamReader(FileName, True)
        Dim line As String = SR.ReadLine()
        Dim strArray As String() = line.Split(";"c)
        Dim dt As DataTable = New DataTable()
        Dim row As DataRow

        For Each s As String In strArray
            dt.Columns.Add(s.Replace(" ", ""))
        Next

        Do
            line = SR.ReadLine
            If Not line = String.Empty Then
                row = dt.NewRow()
                row.ItemArray = line.Replace(" ", "").Split(";"c)
                dt.Rows.Add(row)
            Else
                Exit Do
            End If
        Loop

        Return dt
    End Function


    Public Function Datatable2ListLines(dt As DataTable) As List(Of Class_line)
        Dim Liste = New List(Of Class_line)
        Dim lastCompte = ""
        For Each rowData As DataRow In dt.Rows
            Dim lineCompte As New Class_line(rowData, lastCompte)
            If lineCompte.Compte.Substring(0, 2) = "60" Then
                lastCompte = lineCompte.Compte
            ElseIf lineCompte.Compte.Substring(0, 2) = "44" Or lineCompte.Compte.Substring(0, 2) = "40" Then
                lastCompte = ""
            End If
            Liste.Add(lineCompte)
        Next

        Return Liste
    End Function

    Public Sub ListLines2File(FileName As String, Tableau As List(Of Class_line))
        Dim Filetxt = File.CreateText(FileName)
        Dim newlien As New Class_line()
        Filetxt.WriteLine(newlien.Tilte)
        For Each Line In VarGlobales.ListeComptes
            Filetxt.WriteLine(Line.ToTxt)
        Next
        Filetxt.Close()
    End Sub

End Module
