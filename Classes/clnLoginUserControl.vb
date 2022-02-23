Imports System.Data.SqlClient
Imports System.Web
Imports System.Text

Public Class clnLoginUserControl

    Private Shared Function RetornaStringConexao() As String
        Try
            Return clnServidor.fpuObterStrConn(FuncoesGerais.Sistema.GCMONLINEv300, FuncoesGerais.TipoConexao.Producao)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Private Shared Function retornacolunas(SQL As String) As DataTable
        Dim conexao As SqlConnection = Nothing
        Dim ds As New DataSet()
        Dim strcon As String = RetornaStringConexao()

        Try
            conexao = New SqlConnection(strcon)
            conexao.Open()
            Dim da As New SqlDataAdapter(SQL, conexao)
            da.SelectCommand.CommandTimeout = TimeSpan.FromMinutes(9999).Seconds
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
        If ds.Tables.Count > 0 Then
            Return ds.Tables(0)
        End If
    End Function
    Public Shared Function LoginUserControl(ByVal strUsuario As String) As Boolean

        Dim retorno As Boolean = False
        Dim dtNow As DateTime = DateTime.Now
        Try
            Dim dt As New System.Data.DataTable()
            Try
                'dt = retornacolunas("select gcmonline_usr_flag_onLine from gcm_gcmonline_usuario where gcmonline_usr_usuario = '" + strUsuario + "' and gcmonline_usr_sts = 'AA'")
                dt = retornacolunas("update [GCMONLINE_V300].[dbo].[gcm_gcmonline_usuario] set gcmonline_usr_dta_ultimoacesso = GETDATE() ,gcmonline_usr_flag_onLine = 1 where gcmonline_usr_usuario = '" + strUsuario + "'")
            Catch ex As Exception

            End Try
            'If Convert.ToBoolean(dt.Rows(0)("gcmonline_usr_flag_onLine")) = False Then
            'dt = retornacolunas("update gcm_gcmonline_usuario set gcmonline_usr_dta_ultimoacesso = '" + CStr(dtNow) + "' ,gcmonline_usr_flag_onLine = 1 where gcmonline_usr_usuario = '" + strUsuario + "'")
            'End If
            'dt.Dispose()
            Return retorno
        Catch ex As Exception
            Return retorno
        End Try

        'Try
        'Dim dtNow As DateTime = DateTime.Now
        'Dim dt As New System.Data.DataTable()
        ''dt = retornacolunas("select ((DATEPART(HOUR, GETDATE()) - DATEPART(HOUR, gcmonline_usu_data))*60) + ((DATEPART(minutE, GETDATE()) - DATEPART(minutE, gcmonline_usu_data))) as diferenca from gcm_gcmonline_usuariologado where gcmonline_usr_username = '" + strUsuario + "'")
        'dt = retornacolunas("select gcmonline_usr_flag_onLine from gcm_gcmonline_usuario where gcmonline_usr_usuario = '" + strUsuario + "' and gcmonline_usr_sts = 'AA'")

        ''Se estiver logado é necessário retornar false para não entrar no IF
        'If Convert.ToBoolean(dt.Rows(0)("gcmonline_usr_flag_onLine")) Then
        '        dt = retornacolunas("update gcm_gcmonline_usuario set gcmonline_usr_dta_ultimoacesso = '" + CStr(dtNow) + "' ,gcmonline_usr_flag_onLine = 1 where gcmonline_usr_usuario = '" + strUsuario + "'")
        'retorno = False
        'Else
        'retorno = True
        'End If

        'Catch ex As Exception
        '        retorno = False
        'End Try

        Return retorno
    End Function
    Public Function SystemOwner(ByVal cat_id As Decimal) As DataTable

        Dim SQL As String
        SQL = ""
        SQL += "select cat.cat_id, cat.cat_grupoproprietario, usu.gcmonline_usr_id, usu.gcmonline_usr_descricao , 30 as tap_id, '_KEY USER' as tap_descricao, 20 as tap_Peso "
        SQL += "from tbl_CAT_Categoria cat "
        SQL += "inner join [GCMONLINE_V300].[dbo].[gcm_gcmonline_grupo] gru on gru.gcmonline_gru_id = cat.cat_GrupoProprietario "
        SQL += "inner join [GCMONLINE_V300].[dbo].[gcm_gcmonline_usuario] usu on usu.gcmonline_usr_id = gru.gcmonline_usr_id and usu.gcmonline_usr_sts = 'AA' "
        SQL += "where cat.cat_id = " + Convert.ToString(cat_id)

        Return retornacolunas(SQL)

    End Function

    Public Overloads Sub AutenticarUsuario(ByVal IDUsuario As Decimal)
        Dim ctx As System.Web.HttpContext
        ctx = HttpContext.Current
        ctx.Session("dsUsuario") = RetornarUsuario(IDUsuario)
        AdicionarCookieUsuario()
    End Sub
    Public Function RetornarUsuario(ByVal IDUsuario As Decimal) As DataTable

        Try
            'Dim Conn As New SqlConnection(FuncoesGerais.clnServidor.fpuObterStrConn(FuncoesGerais.Sistema.GCMONLINEv300, FuncoesGerais.TipoConexao.Producao))
            Dim Conn As New SqlConnection(RetornaStringConexao())

            Dim arrParms() As SqlParameter = New SqlParameter(1) {}

            arrParms(0) = New SqlParameter("@gcmonline_usr_usuario", SqlDbType.Decimal)
            arrParms(0).Value = IDUsuario

            arrParms(1) = New SqlParameter("@TipoPesquisa", SqlDbType.Int)
            arrParms(1).Value = 3

            Return SqlHelper.ExecuteDataset(Conn, CommandType.StoredProcedure, "sp_GCMOnLine_RetornarUsuario", arrParms).Tables(0)

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Sub AdicionarCookieUsuario()
        Try
            Dim dtNow As DateTime = DateTime.Now
            Dim tsMinute As New TimeSpan(0, 1, 0, 0)

            Dim ctx As System.Web.HttpContext
            ctx = HttpContext.Current
            If ctx.Request.Cookies.Item("IDUsuario") Is Nothing Then
                ctx.Response.Cookies.Add(New Web.HttpCookie("IDUsuario", CType(ctx.Session("dsUsuario"), DataTable).Rows(0).Item("gcmonline_usr_usuario").ToString))
            Else
                ctx.Response.Cookies.Item("IDUsuario").Value = CType(ctx.Session("dsUsuario"), DataTable).Rows(0).Item("gcmonline_usr_usuario").ToString
            End If
            ctx.Request.Cookies.Item("IDUsuario").Expires = dtNow.AddHours(8) + tsMinute
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Function CriptografiaBO(ByVal entrada As String) As String

        Dim tripledescryptoserviceprovider As System.Security.Cryptography.TripleDESCryptoServiceProvider = New System.Security.Cryptography.TripleDESCryptoServiceProvider()
        Dim md5cryptoserviceprovider As System.Security.Cryptography.MD5CryptoServiceProvider = New System.Security.Cryptography.MD5CryptoServiceProvider()

        Dim myKey As String = "15a7ml3/40912197"
        tripledescryptoserviceprovider.Key = md5cryptoserviceprovider.ComputeHash(ASCIIEncoding.ASCII.GetBytes(myKey))
        tripledescryptoserviceprovider.Mode = System.Security.Cryptography.CipherMode.ECB
        Dim desdencrypt As System.Security.Cryptography.ICryptoTransform = tripledescryptoserviceprovider.CreateDecryptor
        Dim buff() As Byte = Convert.FromBase64String(entrada)

        Return ASCIIEncoding.ASCII.GetString(desdencrypt.TransformFinalBlock(buff, 0, buff.Length))

    End Function
End Class
