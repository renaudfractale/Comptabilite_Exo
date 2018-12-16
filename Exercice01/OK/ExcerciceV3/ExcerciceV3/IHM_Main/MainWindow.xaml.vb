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

            'Tranformation en dataTable le fichier Txt
            Dim dt = File2Datatable(TextBox_Path_File_Inpute.Text)
            'Tranformation en liste object  la dataTable
            VarGlobales.ListeComptes = Datatable2ListLines(dt)


            ' Affiche Datatable
            If Not VarGlobales.DataShowIHM.IsVisible Then
                Try
                    VarGlobales.DataShowIHM.Show()
                Catch ex As Exception
                    VarGlobales.DataShowIHM = New DataShow
                    VarGlobales.DataShowIHM.Show()
                End Try
            End If
            VarGlobales.DataShowIHM.GetData(dt)




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


        VentilationTVA(CInt(TextBox_Input_ventilation.Text), VarGlobales.ListeComptes)


        Dim FileNameNew = TextBox_Path_File_Inpute.Text.Replace(".txt", "V2.txt")
        ListLines2File(FileNameNew, VarGlobales.ListeComptes)

        'Tranformation en dataTable le fichier Txt
        Dim dt = File2Datatable(FileNameNew)
        ' Affiche Datatable
        If Not VarGlobales.DataShowIHM.IsVisible Then
            Try
                VarGlobales.DataShowIHM.Show()
            Catch ex As Exception
                VarGlobales.DataShowIHM = New DataShow
                VarGlobales.DataShowIHM.Show()
            End Try
        End If
        VarGlobales.DataShowIHM.GetData(dt)


        Process.Start(FileNameNew)

    End Sub
End Class
