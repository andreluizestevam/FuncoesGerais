Imports System.Data.SqlClient
Namespace SISCOP
    Public Class clnSISCOP

#Region "Atributos"

        Private _Categoria As Decimal
        Private _SubCategoria As Decimal
        Private _Descricao As String
        Private _gcmonline_usr_id As Decimal

        ''' <summary>
        ''' ID do usuário que está na transação. Somente utilizado quando o erro for em aplicações ASP.NET
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>Somente utilizado quando o erro for em aplicações ASP.NET</remarks>
        Public Property gcmonline_usr_id() As Decimal
            Get
                Return _gcmonline_usr_id
            End Get
            Set(ByVal value As Decimal)
                _gcmonline_usr_id = value
            End Set
        End Property

        Public Property Categoria() As Decimal
            Get
                Return _Categoria
            End Get
            Set(ByVal value As Decimal)
                _Categoria = value
            End Set
        End Property

        Public Property SubCategoria() As Decimal
            Get
                Return _SubCategoria
            End Get
            Set(ByVal value As Decimal)
                _SubCategoria = value
            End Set
        End Property

        Public Property Descricao() As String
            Get
                Return _Descricao
            End Get
            Set(ByVal value As String)
                _Descricao = value
            End Set
        End Property

#End Region

#Region "Métodos"
        Private Shared Function RetornaStringConexao() As String

            Try
                Return clnServidor.fpuObterStrConn(FuncoesGerais.Sistema.SISCOP, FuncoesGerais.TipoConexao.Producao)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function
        
        Public Shared Function RetornaNomeUsuario() As String

            Dim strSQL As String

            Dim prmParametro As SqlClient.SqlParameter() = New SqlClient.SqlParameter(0) {}
            prmParametro(0) = New SqlClient.SqlParameter("@gcmonline_usr_id", RetornaIDUsuario(RetornaUsuarioLogin))

            strSQL = "select GCMONLINE_V300.dbo.scRetornaNomeFuncionario (@gcmonline_usr_id)"

            Dim drUsuario As SqlClient.SqlDataReader
            Dim strUsuarioNome As String

            Try
                drUsuario = FuncoesGerais.SqlHelper.ExecuteReader(RetornaStringConexao, CommandType.Text, strSQL, prmParametro)
                If drUsuario.Read Then
                    strUsuarioNome = drUsuario.Item(0).ToString.Trim
                End If

                drUsuario.Close()

                Return strUsuarioNome

            Catch ex As Exception
                Throw New Exception(ex.Message)
                drUsuario.Close()
            End Try

        End Function
        Private Shared Function RetornaIDUsuario(ByVal strUsuario As String) As Decimal

            Dim strSQL As String

            Dim prmParametro As SqlClient.SqlParameter() = New SqlClient.SqlParameter(0) {}
            prmParametro(0) = New SqlClient.SqlParameter("@gcmonline_usr_usuario", strUsuario)

            strSQL = "select GCMONLINE_V300.dbo.scRetornaIDFuncionario (@gcmonline_usr_usuario)"

            Dim drUsuario As SqlClient.SqlDataReader
            Dim decUsuario As Decimal

            Try
                drUsuario = FuncoesGerais.SqlHelper.ExecuteReader(RetornaStringConexao, CommandType.Text, strSQL, prmParametro)
                If drUsuario.Read Then
                    decUsuario = Convert.ToDecimal(drUsuario.Item(0).ToString)
                End If

                drUsuario.Close()

                Return decUsuario

            Catch ex As Exception
                Throw New Exception(ex.Message)
                drUsuario.Close()
            End Try

        End Function
        Public Shared Function RetornaUsuarioLogin() As String
            Dim strUsuario As String
            strUsuario = System.Security.Principal.WindowsIdentity.GetCurrent.Name

            Dim intTotal As Integer = strUsuario.Length
            Dim intCount As Integer

            For intCount = intTotal To 1 Step -1

                If Microsoft.VisualBasic.Mid(strUsuario, intCount, 1) = "\" Then
                    strUsuario = Microsoft.VisualBasic.Replace(Microsoft.VisualBasic.Replace(Microsoft.VisualBasic.Replace(Microsoft.VisualBasic.Mid(strUsuario, intCount, intTotal), "\", ""), " ", "_"), ".dll", "")
                End If
            Next

            Return strUsuario

        End Function
        Public Function Gravar() As Decimal
            Try
                Dim prmParametro As SqlParameter() = New SqlParameter(6) {}

                prmParametro(0) = New SqlParameter("@cat_ID", SqlDbType.Decimal)
                prmParametro(0).Value = _Categoria

                prmParametro(1) = New SqlParameter("@sca_ID", SqlDbType.Decimal)
                prmParametro(1).Value = _SubCategoria

                prmParametro(2) = New SqlParameter("@sol_Solicitante", SqlDbType.Decimal)
                prmParametro(2).Value = RetornaIDUsuario(RetornaUsuarioLogin)

                prmParametro(3) = New SqlParameter("@sol_DataSolicitacao", SqlDbType.DateTime)
                prmParametro(3).Value = DateTime.Now.ToShortDateString

                prmParametro(4) = New SqlParameter("@sol_Descricao", SqlDbType.NVarChar, 3000)
                prmParametro(4).Value = _Descricao

                prmParametro(5) = New SqlParameter("@sol_UsuarioTransacao", SqlDbType.Decimal)
                prmParametro(5).Value = RetornaIDUsuario(RetornaUsuarioLogin)

                prmParametro(6) = New SqlParameter("@set_ID", SqlDbType.Decimal)
                prmParametro(6).Value = 29

                Return Convert.ToDecimal(SqlHelper.ExecuteScalar(RetornaStringConexao, CommandType.StoredProcedure, "sp_SISCOP_Solicitacao_i", prmParametro))

            Catch ex As Exception
                Throw
            End Try
        End Function
        Public Function GravarWeb() As Decimal
            Try
                Dim prmParametro As SqlParameter() = New SqlParameter(6) {}

                prmParametro(0) = New SqlParameter("@cat_ID", SqlDbType.Decimal)
                prmParametro(0).Value = _Categoria

                prmParametro(1) = New SqlParameter("@sca_ID", SqlDbType.Decimal)
                prmParametro(1).Value = _SubCategoria

                prmParametro(2) = New SqlParameter("@sol_Solicitante", SqlDbType.Decimal)
                prmParametro(2).Value = _gcmonline_usr_id

                prmParametro(3) = New SqlParameter("@sol_DataSolicitacao", SqlDbType.DateTime)
                prmParametro(3).Value = DateTime.Now.ToShortDateString

                prmParametro(4) = New SqlParameter("@sol_Descricao", SqlDbType.NVarChar, 3000)
                prmParametro(4).Value = _Descricao

                prmParametro(5) = New SqlParameter("@sol_UsuarioTransacao", SqlDbType.Decimal)
                prmParametro(5).Value = _gcmonline_usr_id

                prmParametro(6) = New SqlParameter("@set_ID", SqlDbType.Decimal)
                prmParametro(6).Value = 29

                Return Convert.ToDecimal(SqlHelper.ExecuteScalar(RetornaStringConexao, CommandType.StoredProcedure, "sp_SISCOP_Solicitacao_i", prmParametro))

            Catch ex As Exception
                Throw
            End Try
        End Function       
#End Region

    End Class
End Namespace