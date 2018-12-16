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

    Private Sub Button_Init_Click(sender As Object, e As RoutedEventArgs) Handles Button_Init.Click
        VarsGlobals.BD.Open()
        Button_Clear.IsEnabled = True
        Button_Close.IsEnabled = True
        Button_Init.IsEnabled = False
        Button_Get_Path_File.IsEnabled = True
    End Sub

    Private Sub Button_Clear_Click(sender As Object, e As RoutedEventArgs) Handles Button_Clear.Click
        VarsGlobals.BD.Clear()
        Button_Clear.IsEnabled = True
        Button_Close.IsEnabled = True
        Button_Init.IsEnabled = False
        Button_Get_Path_File.IsEnabled = True
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

            Dim Liste = OpenFileTxt(TextBox_Path_File_Inpute.Text)
            For Each item In Liste
                VarsGlobals.BD.Add_InsertLine(item)
            Next

        End If
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        VarsGlobals.Init()
        TextBox_Path_File_Inpute.IsReadOnly = True
        TextBox_Path_File_OutPut.IsReadOnly = True
        RichTextBox_Log.IsReadOnly = True
        VarsGlobals.textBoox_Log = RichTextBox_Log
        Button_Clear.IsEnabled = False
        Button_Close.IsEnabled = False
        Button_Init.IsEnabled = True
        Button_Get_Path_File.IsEnabled = False
    End Sub



    Private Sub Button_Close_Click(sender As Object, e As RoutedEventArgs) Handles Button_Close.Click
        VarsGlobals.BD.Close()
        Button_Clear.IsEnabled = False
        Button_Close.IsEnabled = False
        Button_Init.IsEnabled = True
        Button_Get_Path_File.IsEnabled = False
    End Sub
End Class
