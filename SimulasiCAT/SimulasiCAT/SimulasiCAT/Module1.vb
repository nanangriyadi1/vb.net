Imports System.Data.OleDb
Imports System.Data


Module Module1

    Public Conn As OleDbConnection
    Public DA As OleDbDataAdapter
    Public DS As DataSet
    Public CMD As OleDbCommand
    Public DR As OleDbDataReader
    Public kn As String = ("provider=microsoft.jet.oledb.4.0;data source=simulasi.mdb")


    Public Sub Koneksi()
        Try
            Conn = New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=simulasi.mdb")
            Conn.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            End
        End Try
    End Sub

End Module


