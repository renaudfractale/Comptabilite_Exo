Public Class Class_Data(Of T)
    Public Property Key As Integer = 0
    Public Property IdKey As Integer = 0
    Public Property Colonne As String = ""
    Public Property Value As T
    Public Property DateEvent As Date = Date.MinValue

    Public Sub New(lKey As Integer, lIdKey As Integer,
                   lColonne As String, lValue As T,
                   Optional lDateEvent As Date = #1970/01/01#)
        Key = lKey
        IdKey = lIdKey
        Colonne = lColonne
        Value = lValue
        DateEvent = lDateEvent
        If DateEvent = #1970/01/01# Then
            DateEvent = Now
        End If
    End Sub

End Class
