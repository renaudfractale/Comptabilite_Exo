Imports System.Data
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Class MainWindow
    Private Sub IsNumber_Input()
        Dim regex As Regex = New Regex("[^0-9]+")
        Dim txt As String = TextBox_Input_ventilation.Text
        Dim State As Boolean = regex.IsMatch(txt)

        If Not State And txt.Length > 0 Then
            Dim i As Integer = CInt(txt)
            If i >= 0 And i <= 100 Then
                TextBox_Input_ventilation.Background = Brushes.LawnGreen
            Else
                TextBox_Input_ventilation.Background = Brushes.PaleVioletRed
            End If
        Else
            TextBox_Input_ventilation.Background = Brushes.PaleVioletRed
        End If
    End Sub

    Private Sub TextBox_Input_ventilation_CHANGED(sender As Object, e As TextChangedEventArgs) Handles TextBox_Input_ventilation.TextChanged
        IsNumber_Input()
    End Sub

    Private Sub Button_Get_Path_File_Click(sender As Object, e As RoutedEventArgs) Handles Button_Get_Path_File.Click
        Dim fd As OpenFileDialog = New OpenFileDialog With {
            .Title = "Selection le fichier Source",
            .InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            .Filter = "xtx fichier (*.txt)|*.txt",
            .Multiselect = False
        }

        If fd.ShowDialog() Then
            TextBox_Path_File_Inpute.IsReadOnly = False
            TextBox_Path_File_Inpute.Text = fd.FileName
            TextBox_Path_File_Inpute.IsReadOnly = True
            Log_application("Fichier source selectionné : " + TextBox_Path_File_Inpute.Text)


            Dim SR As StreamReader = New StreamReader(TextBox_Path_File_Inpute.Text, True)
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


            If Not VarGlobales.DataShowIHM.IsVisible Then
                Try
                    VarGlobales.DataShowIHM.Show()
                Catch ex As Exception
                    VarGlobales.DataShowIHM = New DataShow
                    VarGlobales.DataShowIHM.Show()
                End Try

            End If
            VarGlobales.DataShowIHM.GetData(dt)
            ''TODO : A transformer en object

            Dim lastCompte = ""

            VarGlobales.ListeComptes = New List(Of Class_line)

            For Each rowData As DataRow In dt.Rows
                Dim lineCompte As New Class_line(rowData, lastCompte)
                If lineCompte.Compte.Substring(0, 2) = "60" Then
                    lastCompte = lineCompte.Compte
                ElseIf lineCompte.Compte.Substring(0, 2) = "44" Or lineCompte.Compte.Substring(0, 2) = "40" Then
                    lastCompte = ""
                End If
                VarGlobales.ListeComptes.Add(lineCompte)
            Next



            Button_Output.IsEnabled = True
        End If
    End Sub
    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        VarGlobales.Init()
        TextBox_Path_File_Inpute.IsReadOnly = True
        TextBox_Path_File_OutPut.IsReadOnly = True
        RichTextBox_Log.IsReadOnly = True
        VarGlobales.textBox_Log = RichTextBox_Log
        Button_Get_Path_File.IsEnabled = True
        Button_Output.IsEnabled = False
        TextBox_Input_ventilation.Text = "4"
        Me.IsNumber_Input()

    End Sub

    Private Sub Button_Output_Click(sender As Object, e As RoutedEventArgs) Handles Button_Output.Click

        If TextBox_Input_ventilation.Background.Equals(Brushes.PaleVioletRed) Then
            ModuleGlobale.Log_application("     Pourcenatge non valide --> ventilation abandonée")
            Exit Sub
        End If


        'Pour chaque Pièces
        Dim Liste_Piece = (From line In VarGlobales.ListeComptes
                           Select line.Piece
                           Distinct).ToList
        For Each Piece In Liste_Piece

            ModuleGlobale.Log_application("Pour la Piece '" + Piece)

            'On cherche le compte de TVA de la pièce //44//
            Dim a = VarGlobales.ListeComptes(0).Compte.PadLeft(2)
            Dim CompteTVAs = (From line In VarGlobales.ListeComptes
                              Where line.Piece = Piece And line.Compte.Substring(0, 2) = "44"
                              Select line.Compte).ToList
            If CompteTVAs.Count <> 1 Then
                ModuleGlobale.Log_application("     Pour la Piece '" + Piece + "' il y a plusieurs ou aucun compte de TVA --> ventilation abandonée")
                Continue For
            End If

            Dim CompteTVA = CompteTVAs.First

            'On calcul le pourcenatge de ventilation
            Dim pourcentage = CInt(TextBox_Input_ventilation.Text)

            Dim TVAValue = (From line In VarGlobales.ListeComptes
                            Where line.Piece = Piece And line.Compte = CompteTVA
                            Select line
                            Distinct).First

            Dim Remise As Integer = CInt(Math.Round(TVAValue.DebitSa * (pourcentage / 100), 0))

            ModuleGlobale.Log_application("     Pourcenatge valide Remise de la TVA = " + Remise.ToString)
            ModuleGlobale.Log_application("     Ancienne TVA = " + TVAValue.DebitSa.ToString)
            TVAValue.DebitSa = TVAValue.DebitSa - Remise
            ModuleGlobale.Log_application("     Nouvelle TVA = " + TVAValue.DebitSa.ToString)


            'On liste les comptes 60***** de la pièce
            Dim Liste_60XXXAndAllInfos = (From line In VarGlobales.ListeComptes
                                          Where line.Piece = Piece And line.Compte.Substring(0, 2) = "60"
                                          Select line
                                          Distinct).ToList

            'On fais la somme de DebitSa des compte 60
            Dim Somme As Integer = 0
            For Each Line In Liste_60XXXAndAllInfos
                Somme = Somme + Line.DebitSa
            Next
            Dim VentilationListe As New List(Of Integer)
            For Each Line In Liste_60XXXAndAllInfos
                VentilationListe.Add(CInt(Math.Round(Remise * (Line.DebitSa / Somme))))
            Next


            For i As Integer = 0 To Liste_60XXXAndAllInfos.Count - 1
                Dim Line = Liste_60XXXAndAllInfos(i)
                Dim Ventilation = VentilationListe(i)

                Dim OldSolde60 As Integer = Line.DebitSa
                Dim NewSolde60 As Integer = 0

                ModuleGlobale.Log_application("     Pour le Compte " + Line.Compte + "de la pièce " + Line.Piece)
                ModuleGlobale.Log_application("        On une ventilation de " + Ventilation.ToString)

                Dim Liste_AnalytiqueAndAllInfos = (From lineData In VarGlobales.ListeComptes
                                                   Where lineData.Piece = Piece And lineData.DernierCompte = Line.Compte
                                                   Select lineData
                                                   Distinct).ToList
                Dim SommeAnalityque As Integer = 0
                For Each CompteAna In Liste_AnalytiqueAndAllInfos
                    ModuleGlobale.Log_application("         On a le compte Analityque : " + CompteAna.Compte)
                    SommeAnalityque = SommeAnalityque + CompteAna.DebitSa
                Next

                For Each CompteAna In Liste_AnalytiqueAndAllInfos
                    Dim newSolde = CInt(Math.Round(CompteAna.DebitSa + (Ventilation * (CompteAna.DebitSa / SommeAnalityque)), 0))
                    ModuleGlobale.Log_application("         On a pour compte Analityque : " + CompteAna.Compte + " Avec comme ancien solde : " + CompteAna.DebitSa.ToString)
                    CompteAna.DebitSa = newSolde
                    NewSolde60 = NewSolde60 + newSolde
                    ModuleGlobale.Log_application("         On a pour compte Analityque : " + CompteAna.Compte + " Avec comme nouveau solde : " + CompteAna.DebitSa.ToString)
                Next
                ModuleGlobale.Log_application("     Ancien Solde  " + Line.Compte + " = " + Line.DebitSa.ToString)
                Line.DebitSa = NewSolde60
                ModuleGlobale.Log_application("     Nouveau Solde  " + Line.Compte + " = " + Line.DebitSa.ToString)

            Next

        Next

        Dim Filetxt = File.CreateText(TextBox_Path_File_Inpute.Text.Replace(".txt", "V2.txt"))
        Dim newlien As New Class_line()
        Filetxt.WriteLine(newlien.Tilte)
        For Each Line In VarGlobales.ListeComptes
            Filetxt.WriteLine(Line.ToTxt)
        Next
        Filetxt.Close()

        Process.Start(TextBox_Path_File_Inpute.Text.Replace(".txt", "V2.txt"))

    End Sub
End Class
