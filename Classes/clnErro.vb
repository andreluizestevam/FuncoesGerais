Namespace GerenciaErro
    Public Class clnErro
        Public Shared Sub GerarRelatorioErro(ByVal frmFormulario As Form, ByVal ex As Exception, ByVal Categoria As Decimal, ByVal SubCategoria As Decimal)

            Dim frmErro As New frmRelatorioErro(Categoria, SubCategoria)

            frmErro.txtDescricao.Text = "*******************************************" + Environment.NewLine
            frmErro.txtDescricao.Text += "LOGIN:" + SISCOP.clnSISCOP.RetornaUsuarioLogin + Environment.NewLine
            frmErro.txtDescricao.Text += "******************************************" + Environment.NewLine
            frmErro.txtDescricao.Text += "SISTEMA: " + ex.Source + Environment.NewLine
            frmErro.txtDescricao.Text += "******************************************" + Environment.NewLine + Environment.NewLine
            frmErro.txtDescricao.Text += "Formulário: " + frmFormulario.Name + Environment.NewLine
            frmErro.txtDescricao.Text += "Formulário Descrição: " + frmFormulario.Text + Environment.NewLine + Environment.NewLine
            frmErro.txtDescricao.Text += "Mensagem: " + ex.Message + Environment.NewLine + Environment.NewLine
            frmErro.txtDescricao.Text += "Trace: " + ex.StackTrace + Environment.NewLine + Environment.NewLine
            frmErro.txtDescricao.Text += "HelpLink: " + ex.HelpLink + Environment.NewLine + Environment.NewLine
            frmErro.ShowDialog()
        End Sub
        Public Shared Sub GerarRelatorioErro(ByVal srvServico As System.ServiceProcess.ServiceBase, ByVal ex As Exception, ByVal Categoria As Decimal, ByVal SubCategoria As Decimal)

            Dim strErro As String
            strErro = "*******************************************" + Environment.NewLine
            strErro += "LOGIN:" + SISCOP.clnSISCOP.RetornaUsuarioLogin + Environment.NewLine
            strErro += "******************************************" + Environment.NewLine
            strErro += "SISTEMA: " + ex.Source + Environment.NewLine
            strErro += "******************************************" + Environment.NewLine + Environment.NewLine
            strErro += "Serviço: " + srvServico.ServiceName + Environment.NewLine
            strErro += "Mensagem: " + ex.Message + Environment.NewLine + Environment.NewLine
            strErro += "Trace: " + ex.StackTrace + Environment.NewLine + Environment.NewLine
            strErro += "HelpLink: " + ex.HelpLink + Environment.NewLine + Environment.NewLine

            Dim clsSISCOP As New SISCOP.clnSISCOP
            clsSISCOP.Categoria = Categoria
            clsSISCOP.Descricao = strErro
            clsSISCOP.SubCategoria = SubCategoria
            clsSISCOP.Gravar()

        End Sub
        Public Shared Sub GerarRelatorioErroWeb(ByVal strPagina As String, ByVal ex As Exception, ByVal Categoria As Decimal, ByVal SubCategoria As Decimal)

            Web.HttpContext.Current.Session.Add("ssErroDetalhado", ex)
            Web.HttpContext.Current.Response.Redirect("~/GCMONLINE_WEB/erro/frmErroDetalhado.aspx?cat_ID=" + Categoria.ToString + "&sca_ID=" + SubCategoria.ToString + "&strPagina=" + strPagina)

        End Sub
    End Class
End Namespace
