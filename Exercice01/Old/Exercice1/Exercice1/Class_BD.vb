Imports System.Data.SQLite


Public Class Class_BD
    Private mBD As SQLiteConnection
    Private mBdIsOpen As Boolean = False

    Private Function GetPathBD() As String
        Return My.Application.Info.DirectoryPath + "\Data\MyDatabase.sqlite"
    End Function
    Private Sub Init()
        Dim PathRoot As String = System.IO.Path.GetDirectoryName(GetPathBD)
        If Not My.Computer.FileSystem.DirectoryExists(PathRoot) Then
            Log_application("Start Creation de repertoire de la base de données")
            My.Computer.FileSystem.CreateDirectory(PathRoot)
            Log_application("End Creation de repertoire de la base de données")
        End If
        Log_application("Dossier de la base de données : " + PathRoot)
        Dim IsNewFile As Boolean = False
        If Not My.Computer.FileSystem.FileExists(GetPathBD) Then
            Log_application("Start Creation du fichier de la base de données")
            SQLiteConnection.CreateFile(GetPathBD)
            Log_application("End Creation du fichier de la base de données")
            IsNewFile = True
        End If
        Log_application("Chemin de la base de données : " + GetPathBD())
        ' Récupération du chemin vers notre fichier de base de données

        'Instanciation de notre connexion
        Log_application("Start Ouverture du fichier de la base de données")
        mBD = New SQLiteConnection("Data Source=" + GetPathBD() + ";Version=3;")
        mBD.Open()
        Log_application("End Ouverture du fichier de la base de données")
        If IsNewFile Then
            Log_application("Start Creation des Tables")
            Dim sql_1 = "create table MainTable (id INTEGER PRIMARY KEY, typecompta int, idjournal int, compte TEXT, datecomptable int, debitsa int, creditsa int, piece TEXT)"
            Dim command_1 = New SQLiteCommand(sql_1, mBD)
            command_1.ExecuteNonQuery()

            Dim sql_2 = "create table intTable (id INTEGER PRIMARY KEY, idkey int, colonne TEXT, value int, datesevent TEXT)"
            Dim command_2 = New SQLiteCommand(sql_2, mBD)
            command_2.ExecuteNonQuery()

            Dim sql_3 = "create table dateTable (id INTEGER PRIMARY KEY, idkey int, colonne TEXT, value TEXT, datesevent TEXT)"
            Dim command_3 = New SQLiteCommand(sql_3, mBD)
            command_3.ExecuteNonQuery()

            Log_application("End Creation des Tables")
        End If
        mBdIsOpen = True
        Log_application("Base de Données marquée comme Ouverte")
    End Sub

    Private Function MakeKey(line As Class_line) As String
        Return line.TypeCompta.ToString + ";" + line.IdJournal.ToString + ";" + line.Compte + ";" + line.Piece
    End Function


    Private Function GetAllTable() As List(Of Class_line)

        Dim Liste As New List(Of Class_line)
        If mBdIsOpen Then
            Dim req = "select * from MainTable;"
            Dim mySQLcommand = mBD.CreateCommand()
            mySQLcommand.CommandText = req
            Dim mySQLreader = mySQLcommand.ExecuteReader()
            While (mySQLreader.Read())
                Dim line As New Class_line()
                line.Key = CInt(mySQLreader("id"))
                line.TypeCompta = CInt(mySQLreader("typecompta"))
                line.IdJournal = CInt(mySQLreader("idjournal"))
                line.Compte = mySQLreader("compte").ToString
                line.DateCompteID = CInt(mySQLreader("datecomptable"))
                line.DebitSaID = CInt(mySQLreader("debitsa"))
                line.CreditSaID = CInt(mySQLreader("creditsa"))
                line.Piece = mySQLreader("piece").ToString

                Dim req_DateCompteID = "select * from dateTable where idkey = " + line.Key.ToString + " AND colonne = 'datecomptable';"
                Dim mySQLcommand_DateCompteID = mBD.CreateCommand()
                mySQLcommand_DateCompteID.CommandText = req_DateCompteID
                Dim mySQLreader_DateCompteID = mySQLcommand_DateCompteID.ExecuteReader()
                While (mySQLreader_DateCompteID.Read())
                    line.DateCompte.Add(New Class_Data(Of Date)(CInt(mySQLreader_DateCompteID("id")),
                                                                CInt(mySQLreader_DateCompteID("idkey")),
                                                                mySQLreader_DateCompteID("colonne").ToString(),
                                                                CDate(mySQLreader_DateCompteID("value").ToString()),
                                                                CDate(mySQLreader_DateCompteID("datesevent").ToString())))
                End While

                Dim req_DebitSaID = "select * from intTable where idkey = " + line.Key.ToString + " AND colonne = 'debitsa';"
                Dim mySQLcommand_DebitSaID = mBD.CreateCommand()
                mySQLcommand_DebitSaID.CommandText = req_DebitSaID
                Dim mySQLreader_DebitSaID = mySQLcommand_DebitSaID.ExecuteReader()
                While (mySQLreader_DebitSaID.Read())
                    line.DebitSa.Add(New Class_Data(Of Integer)(CInt(mySQLreader_DebitSaID("id")),
                                                                CInt(mySQLreader_DebitSaID("idkey")),
                                                                mySQLreader_DebitSaID("colonne").ToString(),
                                                                CInt(mySQLreader_DebitSaID("value")),
                                                                CDate(mySQLreader_DebitSaID("datesevent").ToString())))
                End While

                Dim req_CreditSaID = "select * from intTable where idkey = " + line.Key.ToString + " AND colonne = 'creditsa';"
                Dim mySQLcommand_CreditSaID = mBD.CreateCommand()
                mySQLcommand_CreditSaID.CommandText = req_CreditSaID
                Dim mySQLreader_CreditSaID = mySQLcommand_CreditSaID.ExecuteReader()
                While (mySQLreader_CreditSaID.Read())
                    line.CreditSa.Add(New Class_Data(Of Integer)(CInt(mySQLreader_CreditSaID("id")),
                                                                CInt(mySQLreader_CreditSaID("idkey")),
                                                                mySQLreader_CreditSaID("colonne").ToString(),
                                                                CInt(mySQLreader_CreditSaID("value")),
                                                                CDate(mySQLreader_CreditSaID("datesevent").ToString())))
                End While

                Liste.Add(line)
            End While
        End If

        Return Liste
    End Function

    Private Function GetKeysDico() As Dictionary(Of String, Class_line)
        Dim Dico As New Dictionary(Of String, Class_line)
        Dim Liste = Me.GetAllTable()
        For Each item In Liste
            Dim Key = MakeKey(item)

            If Dico.ContainsKey(Key) Then
                'Exixste
                Dico.Item(Key) = (item)
            Else
                'New Key
                Dico.Add(Key, item)
            End If

        Next
        Return Dico
    End Function

    Public Sub Add_InsertLine(line As Class_line)
        Dim Key As String = Me.MakeKey(line)
        Dim Dico = GetKeysDico()
        If mBdIsOpen Then
            If Dico.ContainsKey(Key) Then
                'Dim item = Dico(Key)
                ''Exixste
                'Dim req = "UPDATE MainTable SET datecomptable = '" + line.DateCompte.ToString("dd/MM/yyyy") + "',debitsa = " + line.DebitSa.ToString() + ", creditsa = " + line.CreditSa.ToString() + "  WHERE id = " + item.Key.ToString + ";"
                'Dim mySQLcommand = mBD.CreateCommand()
                'mySQLcommand.CommandText = req
                'Dim mySQLreader = mySQLcommand.ExecuteReader()
            Else
                'New Key



                Dim req = "INSERT INTO MainTable VALUES (null, " + line.TypeCompta.ToString _
                    + ", " + line.IdJournal.ToString _
                    + ", '" + line.Compte + "'" _
                    + ", 0" _
                    + ", 0" _
                    + ", 0" _
                    + ", '" + line.Piece + "'" _
                    + ");"

                Dim mySQLcommand = mBD.CreateCommand()
                mySQLcommand.CommandText = req
                mySQLcommand.ExecuteReader()

                'Afir
                Dim req_GetIDKeyMaster = "select * from MainTable where idkey = " + line.Key.ToString + " AND colonne = 'datecomptable';"
                Dim mySQLcommand_id = mBD.CreateCommand()
                mySQLcommand_id.CommandText = req_GetIDKeyMaster
                Dim mySQLreader_Id = mySQLcommand_id.ExecuteReader()
                If mySQLreader_Id.Read() Then
                    Log_application("Keymaster = " + CInt(mySQLreader_Id("id")).ToString + vbCrLf + req)
                End If
            End If
        End If
    End Sub

    Public Sub Clear()
        Log_application(" ########  Effacement de la Base de données ########")
        Me.Close()
        If Not mBdIsOpen And My.Computer.FileSystem.FileExists(GetPathBD) Then
            Log_application(" //////////////  Fichier Existant : Suppresion du fichier  \\\\\\\\\\")
            My.Computer.FileSystem.DeleteFile(GetPathBD())
        End If
        Log_application(" **********  Effacement de la Base de données **********")
    End Sub

    Public Sub Close()
        Log_application(" ########  Fermeture de la Base de données ########")
        If mBdIsOpen Then
            Log_application(" //////////////  Fichier Ouvert : Fermeture du fichier  \\\\\\\\\\")
            mBD.Close()
            mBdIsOpen = False
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Else
            Log_application(" //////////////  Fichier Fermé : On ne fait rien  \\\\\\\\\\")

        End If
        Log_application(" **********  Fermeture de la Base de données **********")
    End Sub

    Public Sub Open()
        Log_application(" ########  Ouverture de la Base de données ########")
        If Not mBdIsOpen Then
            Log_application(" //////////////  Fichier Fermé : Ouverture du fichier  \\\\\\\\\\")
            Me.Init()
        Else
            Log_application(" //////////////  Fichier Ouvert : On ne fait rien  \\\\\\\\\\")
        End If
        Log_application(" **********  Ouverture de la Base de données **********")
    End Sub

End Class
