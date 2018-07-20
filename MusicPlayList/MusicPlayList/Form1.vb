Imports System
Imports System.IO
Imports System.Data
Imports System.Diagnostics
Imports System.Reflection
Imports System.IO.Compression
Imports System.Threading
Imports System.Runtime.InteropServices
Imports Oracle.DataAccess.Client

Public Class Form1
    'Dim sw As StreamWriter = New StreamWriter("d:/aaa.txt")
    Dim cnn As New OleDb.OleDbConnection
    Dim cmd As OleDb.OleDbCommand
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sql As String
        Dim sql1 As String
        Dim filename As String
        Dim lan As String
        Dim yea As String
        oracle_execute_non_query("ten", "cbs", "cbs", "truncate table z_du")

        DirSearch("D:\Music")

        Dim oradb As String = "Data Source=ten;User Id=cbs;Password=cbs;"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        sql = "select yea,lan,'D:\Music\Playlist - System\'||yea||' '||lan||' ('||count(1)||').m3u' from (select yea,lan,pat from (SELECT pkgsmgbcommon.delimitedtext(linedata,'\',7) yea,       pkgsmgbcommon.delimitedtext(linedata,'\',4) lan,       linedata pat  FROM Z_DU where pkgsmgbcommon.delimitedtext(linedata,'\',3) = 'Film Songs')) group by yea,lan"
        Dim cmd4 As New OracleCommand(sql, conn)
        Dim dr As OracleDataReader = cmd4.ExecuteReader()
        While dr.Read()
            yea = dr.Item(0).ToString
            lan = dr.Item(1).ToString
            filename = dr.Item(2).ToString
            If filename <> "" Then
                Dim sw As StreamWriter = New StreamWriter(filename)
                sql1 = "select replace(replace(pat,'\','/'),'E:/','/storage/emulated/0/') from (SELECT pkgsmgbcommon.delimitedtext(linedata,'\',7) yea,       pkgsmgbcommon.delimitedtext(linedata,'\',4) lan,       linedata pat  FROM Z_DU where pkgsmgbcommon.delimitedtext(linedata,'\',3) = 'Film Songs') where yea = " & yea & " and lan= '" & lan & "' "
                Dim cmd5 As New OracleCommand(sql1, conn)
                Dim dr1 As OracleDataReader = cmd5.ExecuteReader()
                While dr1.Read()
                    sw.WriteLine(dr1.Item(0).ToString)
                End While
                dr1.Close()
                sw.Close()
            End If
        End While
        dr.Close()

        sql = "select yea,'D:\Music\Playlist - System\'||yea||' ('||count(1)||').m3u' from (select yea,lan,pat from (SELECT pkgsmgbcommon.delimitedtext(linedata,'\',7) yea,       pkgsmgbcommon.delimitedtext(linedata,'\',4) lan,       linedata pat  FROM Z_DU where pkgsmgbcommon.delimitedtext(linedata,'\',3) = 'Film Songs')) group by yea"
        Dim cmd6 As New OracleCommand(sql, conn)
        Dim dr6 As OracleDataReader = cmd6.ExecuteReader()
        While dr6.Read()
            yea = dr6.Item(0).ToString
            filename = dr6.Item(1).ToString
            If filename <> "" Then
                Dim sw As StreamWriter = New StreamWriter(filename)
                sql1 = "select replace(replace(pat,'\','/'),'E:/','/storage/emulated/0/') from (SELECT pkgsmgbcommon.delimitedtext(linedata,'\',7) yea,       pkgsmgbcommon.delimitedtext(linedata,'\',4) lan,       linedata pat  FROM Z_DU where pkgsmgbcommon.delimitedtext(linedata,'\',3) = 'Film Songs') where yea = " & yea
                Dim cmd7 As New OracleCommand(sql1, conn)
                Dim dr7 As OracleDataReader = cmd7.ExecuteReader()
                While dr7.Read()
                    sw.WriteLine(dr7.Item(0).ToString)
                End While
                dr7.Close()
                sw.Close()
            End If
        End While
        dr6.Close()

        conn.Close()


        MsgBox("Process completed successfully")


    End Sub
    Sub oracle_execute_non_query(ByVal database As String, ByVal user As String, ByVal password As String, ByVal query As String)

        Dim oradb As String = "Data Source=ten;User Id=cbs;Password=cbs;"
        Dim conn As New OracleConnection(oradb)
        conn.Open()

        Dim cmd1 As New OracleCommand(query, conn)
        cmd1.ExecuteNonQuery()

        conn.Close()
        conn.Dispose()

    End Sub
    Public Sub DirSearch(ByVal sDir As String)
        Try
            For Each d As String In Directory.GetDirectories(sDir)
                For Each f As String In Directory.GetFiles(d)
                    oracle_execute_non_query("ten", "cbs", "cbs", "insert into z_du (linedata) values ('" & f & "')")
                Next

                DirSearch(d)
            Next
        Catch excpt As System.Exception
            Console.WriteLine(excpt.Message)
        End Try
    End Sub

End Class
