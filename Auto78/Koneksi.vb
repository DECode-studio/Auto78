Imports MySql.Data.MySqlClient
Imports System.Data.OleDb

Module Koneksi
    Public db As New MySql.Data.MySqlClient.MySqlConnection
    Public sql As String
    Public cmd As MySqlCommand
    Public Adapter As MySqlDataAdapter
    Public rs As MySqlDataReader
    Public DataSet As DataSet
    Public dt As DataTable
    Public Jenis_User As String
    Public Nama_User As String

    Public Sub OpenConnection()
        sql = "server=localhost;uid=root;pwd=;database=db_auto78"
        Try
            db.ConnectionString = sql
            db.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public CONN As OleDbConnection
    Public COMMANDD As OleDbCommand
    Public DS As New DataSet
    Public DA As OleDbDataAdapter
    Public DR As OleDbDataReader
    Public database As String

    Public Sub Connection()

        Try
            Dim dir As String = "DBAse/db_auto78.mdb"
            Dim password As String = ""

            database = "provider=microsoft.jet.oledb.4.0;data source=" & dir & _
                ";Jet OLEDB:Database Password=" & password & ";persist security info=false;"

            CONN = New OleDbConnection(database)

            If CONN.State = ConnectionState.Closed Then
                CONN.Open()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Module
