Imports System.Data

Public Class Class_line
    Public Property TypeCompta As String = ""
    Public Property IDJournal As String = ""
    Public Property Compte As String = ""
    Public Property DateComptable As String = ""
    Public Property DebitSa As Integer = 0
    Public Property CreditSa As String = ""
    Public Property Piece As String = ""
    Public Property DatePiece As String = ""
    Public Property PieceTiers As String = ""
    Public Property DateEcheance As String = ""
    Public Property CodeOperation As String = ""
    Public Property Libelle As String = ""
    Public Property LibPiece As String = ""
    Public Property DeviseSa As String = ""
    Public Property Representant As String = ""
    Public Property Quantite As String = ""
    Public Property CompteGene As String = ""
    Public Property PlanComptable_Libelle As String = ""
    Public Property StatutLigne As String = ""
    Public Property DernierCompte As String = ""

    Public Sub New()

    End Sub
    Public Sub New(Row As DataRow, lastCompte As String)

        TypeCompta = Row("TypeCompta").ToString
        IDJournal = Row("IDJournal").ToString
        Compte = Row("Compte").ToString
        If Compte.Substring(0, 2) = "60" Or Compte.Substring(0, 2) = "44" Or Compte.Substring(0, 2) = "40" Then
            'on ne fait rein
        Else
            DernierCompte = lastCompte
        End If

        DateComptable = Row("DateComptable").ToString
        Dim DebitSaStr = Row("DebitSa").ToString
        If DebitSaStr.Length = 0 Then
            DebitSa = 0
        Else
            DebitSa = CInt(DebitSaStr)
        End If
        CreditSa = Row("CreditSa").ToString
        Piece = Row("Piece").ToString
        DatePiece = Row("DatePiece").ToString
        PieceTiers = Row("PieceTiers").ToString
        DateEcheance = Row("DateEcheance").ToString
        CodeOperation = Row("CodeOperation").ToString
        Libelle = Row("Libelle").ToString
        LibPiece = Row("LibPiece").ToString
        DeviseSa = Row("DeviseSa").ToString
        Representant = Row("Representant").ToString
        Quantite = Row("Quantite").ToString
        CompteGene = Row("CompteGene").ToString
        PlanComptable_Libelle = Row("PlanComptable_Libelle").ToString
        StatutLigne = Row("StatutLigne").ToString
    End Sub

    Public Function ToTxt() As String
        Return TypeCompta + ";" + IDJournal + ";" + Compte + ";" + DateComptable + ";" + DebitSa.ToString + ";" + CreditSa + ";" + Piece + ";" + DatePiece + ";" + PieceTiers + ";" + DateEcheance + ";" + CodeOperation + ";" + Libelle + ";" + LibPiece + ";" + DeviseSa + ";" + Representant + ";" + Quantite + ";" + CompteGene + ";" + PlanComptable_Libelle + ";" + StatutLigne
    End Function

    Public Function Tilte() As String
        Return "TypeCompta;IDJournal;Compte        ;DateComptable;DebitSa     ;CreditSa    ;Piece      ;DatePiece ;PieceTiers ;DateEcheance;CodeOperation;Libelle             ;LibPiece                 ;DeviseSa;Representant;Quantite  ;CompteGene ;PlanComptable_Libelle ;StatutLigne"
    End Function
End Class
