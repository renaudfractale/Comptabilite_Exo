Imports System.Data

Public Class DataShow
    Public Sub GetData(Data As DataTable)
        dataGridView.ItemsSource = Data.DefaultView
    End Sub

End Class
