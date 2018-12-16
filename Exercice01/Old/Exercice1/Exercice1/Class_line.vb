Public Class Class_line
    Public Property TypeCompta As Integer = 0
    Public Property IdJournal As Integer = 0
    Public Property Compte As String = ""
    Public Property DateCompte As New List(Of Class_Data(Of Date))
    Public Property DebitSa As New List(Of Class_Data(Of Integer))
    Public Property CreditSa As New List(Of Class_Data(Of Integer))
    Public Property DateCompteID As Integer = 0
    Public Property DebitSaID As Integer = 0
    Public Property CreditSaID As Integer = 0

    Public Property Piece As String = ""
    Public Property Key As Integer = -1
    Public Sub New()

    End Sub
    'Public Sub New(lTypeCompta As Integer, lIdJournal As Integer,
    '         lCompte As String, lDateCompte As Date, lDebitSa As Integer,
    '               lCreditSa As Integer, lPiece As String, Optional lKey As Integer = -1)
    '    TypeCompta = lTypeCompta
    '    IdJournal = lIdJournal
    '    Compte = lCompte
    '    DateCompte = lDateCompte
    '    DebitSa = lDebitSa
    '    CreditSa = lCreditSa
    '    Piece = lPiece
    '    Key = lKey
    'End Sub

    Public Sub New(txt As String)
        Dim Liste As List(Of String) = txt.Replace(" ", "").Split(";"c).ToList
        If Liste.Count >= 7 Then
            TypeCompta = CintRH(Liste.Item(0))
            IdJournal = CintRH(Liste.Item(1))
            Compte = Liste.Item(2)
            DateCompteID = 0
            DateCompte.Add(New Class_Data(Of Date)(0, Key,
                                                   "datecomptable",
                                                   CDate(Liste.Item(3))))
            DebitSaID = 0
            DebitSa.Add(New Class_Data(Of Integer)(0, Key,
                                                   "debitsa",
                                                   CintRH(Liste.Item(4))))

            CreditSaID = 0
            CreditSa.Add(New Class_Data(Of Integer)(0, Key,
                                                   "creditsa",
                                                   CintRH(Liste.Item(5))))
            Piece = Liste.Item(6)
        End If

    End Sub

    Private Function CintRH(txt As String) As Integer
        Dim i As Integer = 0
        If txt <> "" Then
            i = CInt(txt)
        End If
        Return i
    End Function



End Class
