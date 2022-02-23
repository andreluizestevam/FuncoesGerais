Imports System.Data.SqlClient
'1
'sp_SISLOG_LOG_i

Namespace SISLOG

    Public Class clnSISLOG

#Region "Métodos"
        Public Shared Sub IncluirLog(ByVal prg_ID As Decimal)
            Try

                Dim prmParametro As SqlClient.SqlParameter() = New SqlClient.SqlParameter(0) {}
                prmParametro(0) = New SqlClient.SqlParameter("@prg_ID", prg_ID)

                FuncoesGerais.SqlHelper.ExecuteNonQuery(RetornaStringConexao, CommandType.StoredProcedure, "sp_SISLOG_LOG_i", prmParametro)

            Catch ex As Exception
                Throw
            End Try
        End Sub
        Private Shared Function RetornaStringConexao() As String

            Try
                Return clnServidor.fpuObterStrConn(FuncoesGerais.Sistema.GCMONLINEv300, FuncoesGerais.TipoConexao.Producao)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function
#End Region

    End Class

End Namespace

