Imports System.Data.SqlClient

Public Class clnExecProcedure
    Public Function retornacolunas(SQL As String, strcon As String) As DataTable
        Dim conexao As SqlConnection = Nothing
        Dim ds As New DataSet()

        Try
            conexao = New SqlConnection(strcon)
            conexao.Open()
            Dim da As New SqlDataAdapter(SQL, conexao)
            da.SelectCommand.CommandTimeout = TimeSpan.FromMinutes(180).Seconds
            da.Fill(ds)
        Catch mex As SqlException
            Return Nothing
        Finally
            If conexao IsNot Nothing Then
                If conexao.State = ConnectionState.Open Then
                    conexao.Close()
                End If
            End If
            conexao.Dispose()
            ds.Dispose()
        End Try
        Try
            Return ds.Tables(0)
        Catch
            Return Nothing
        End Try
    End Function

    Public Function retornacolunasDS(SQL As String, strcon As String) As DataSet
        Dim conexao As SqlConnection = Nothing
        Dim ds As New DataSet()

        Try
            conexao = New SqlConnection(strcon)
            conexao.Open()
            Dim da As New SqlDataAdapter(SQL, conexao)
            da.SelectCommand.CommandTimeout = TimeSpan.FromMinutes(180).Seconds
            da.Fill(ds)
        Catch mex As SqlException
            Return Nothing
        Finally
            If conexao IsNot Nothing Then
                If conexao.State = ConnectionState.Open Then
                    conexao.Close()
                End If
            End If
            conexao.Dispose()
            ds.Dispose()
        End Try
        Try
            Return ds
        Catch
            Return Nothing
        End Try
    End Function
End Class
