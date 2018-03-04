﻿Imports MySql.Data.MySqlClient

Public Class MySqlHelper 'Conexion de acceso a datos MySQLClient MySQLSever
    Private objectLibrary As New Library
    Public Function strstrconnection() As MySqlConnection
        Return New MySqlConnection("server=localhost; uid=root;pwd=123456; database=autodatasystem")
    End Function
    Public strcon As MySqlConnection = strstrconnection()
    Public result As String
    Public cmd As New MySqlCommand
    Public da As New MySqlDataAdapter
    Public dt As New DataTable

    Public Function MySqlHelperExecuteNonQuery(ByVal sql As String) As Integer
        Try
            strcon.Open()
            With cmd
                .Connection = strcon
                .CommandText = sql
                MySqlHelperExecuteNonQuery = cmd.ExecuteNonQuery
            End With
        Catch ex As Exception
            MySqlHelperExecuteNonQuery = 0
            objectLibrary.WriteErrorLog(ex.Message & vbCrLf & " SQL = " & sql & vbCrLf)
            strcon.Close()
            cmd.Dispose()
        End Try
        strcon.Close()
        cmd.Dispose()
    End Function
    Public Function MySqlHelperExecuteReader(ByVal sqlReader As String) As MySqlDataReader
        Try
            strcon.Open()
            With cmd
                .Connection = strcon
                .CommandText = sqlReader
            End With
            MySqlHelperExecuteReader = cmd.ExecuteReader()
        Catch ex As Exception
            objectLibrary.WriteErrorLog(ex.Message & vbCrLf & " SQL = " & sqlReader & vbCrLf)
            strcon.Close()
            MySqlHelperExecuteReader = Nothing
        End Try
    End Function
End Class
