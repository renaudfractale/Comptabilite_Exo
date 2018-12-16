Module Module_RegleMetier
    Public Sub VentilationTVA(TVAvalueSeleced As Integer, ByRef Tableau As List(Of Class_line))
        'Pour chaque Pièces
        Dim Liste_Piece = (From line In Tableau
                           Select line.Piece
                           Distinct).ToList
        For Each Piece In Liste_Piece

            ModuleGlobale.Log_application("Pour la Piece '" + Piece)

            'On cherche le compte de TVA de la pièce //44//
            Dim a = Tableau(0).Compte.PadLeft(2)
            Dim CompteTVAs = (From line In Tableau
                              Where line.Piece = Piece And line.Compte.Substring(0, 2) = "44"
                              Select line.Compte).ToList
            If CompteTVAs.Count <> 1 Then
                ModuleGlobale.Log_application("     Pour la Piece '" + Piece + "' il y a plusieurs ou aucun compte de TVA --> ventilation abandonée")
                Continue For
            End If

            Dim CompteTVA = CompteTVAs.First

            'On calcul le pourcenatge de ventilation
            Dim pourcentage = TVAvalueSeleced

            Dim TVAValue = (From line In Tableau
                            Where line.Piece = Piece And line.Compte = CompteTVA
                            Select line
                            Distinct).First

            Dim Remise As Integer = CInt(Math.Round(TVAValue.DebitSa * (pourcentage / 100), 0))

            ModuleGlobale.Log_application("     Pourcenatge valide Remise de la TVA = " + Remise.ToString)
            ModuleGlobale.Log_application("     Ancienne TVA = " + TVAValue.DebitSa.ToString)
            TVAValue.DebitSa = TVAValue.DebitSa - Remise
            ModuleGlobale.Log_application("     Nouvelle TVA = " + TVAValue.DebitSa.ToString)


            'On liste les comptes 60***** de la pièce
            Dim Liste_60XXXAndAllInfos = (From line In Tableau
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

                Dim Liste_AnalytiqueAndAllInfos = (From lineData In Tableau
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

    End Sub
End Module
