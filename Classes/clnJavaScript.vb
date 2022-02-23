Imports System.Threading

Public Enum Carregar
    Sim = 1
    Não = 2
End Enum
Public Enum Mensagem
    Processando = 1
    Carregando = 2
End Enum

Public Class clnJavaScript
#Region "Java Script"

    Public Shared Sub jsBarraStatus(ByVal page As System.Web.UI.Page)

        Dim sb As New System.Text.StringBuilder
        sb.Append("<script>window.status='GCMONLINE';</script>")

        page.RegisterStartupScript("jsStatusBar", sb.ToString)

    End Sub
    Public Shared Function jsBarraStatus() As String
        Dim sb As New System.Text.StringBuilder
        sb.Append("<script>window.status='GCMONLINE';</script>")

        Return sb.ToString
    End Function

    Public Shared Sub jsSair(ByVal imgSair As System.Web.UI.WebControls.Image)

        imgSair.Attributes.Add("onClick", "javascript:window.close();")

    End Sub

    Public Shared Sub jsSair(ByVal btnSair As System.Web.UI.WebControls.ImageButton)

        btnSair.Attributes.Add("onClick", "javascript:window.close(); return false;")

    End Sub

    Public Shared Sub jsPopUp(ByVal img As System.Web.UI.WebControls.Image, ByVal strPagina As String, ByVal intLargura As Integer, ByVal intAltura As Integer, ByVal blnRetorno As Boolean, ByVal ParamArray arrParam() As System.Web.UI.WebControls.ListItem)

        'Exemplo de chamada:
        'BPadrao.jsPopUp(wcToolBar.btnTransferencia, _
        '     "Medico/SIAVIMSOTRANSFERIRMEDICO.aspx", _
        '    480, _
        '    180, _
        '    False, _
        '    New ListItem("simvim_tr_id", BPadrao.CriptografarTexto(objBUsuario.Territorio)), _
        '    New ListItem("brick", BPadrao.CriptografarTexto(drwConsultorio("simvim_cons_brick").ToString)) _
        '    )

        Dim sb As New System.Text.StringBuilder

        sb.Append("javascript:OpenModal('")
        sb.Append(strPagina)

        If Not (arrParam Is Nothing) AndAlso (arrParam.Length > 0) Then

            sb.Append("?")

            Dim i As Integer

            For i = 0 To arrParam.Length - 1

                If Not i.ToString.Equals("0") Then
                    sb.Append("&")
                End If
                sb.Append(arrParam(i).Text)
                sb.Append("=")
                sb.Append(arrParam(i).Value)

            Next

        End If

        sb.Append("',")
        sb.Append(intLargura)
        sb.Append(",")
        sb.Append(intAltura)
        sb.Append(");")

        If Not blnRetorno Then
            sb.Append(" return false;")
        End If

        img.Attributes.Add("onClick", sb.ToString)

    End Sub
    Public Shared Sub jsPopUp(ByVal lnk As System.Web.UI.WebControls.LinkButton, ByVal strPagina As String, ByVal intLargura As Integer, ByVal intAltura As Integer, ByVal blnRetorno As Boolean, ByVal ParamArray arrParam() As System.Web.UI.WebControls.ListItem)

        'Exemplo de chamada:
        'BPadrao.jsPopUp(wcToolBar.btnTransferencia, _
        '     "Medico/SIAVIMSOTRANSFERIRMEDICO.aspx", _
        '    480, _
        '    180, _
        '    False, _
        '    New ListItem("simvim_tr_id", BPadrao.CriptografarTexto(objBUsuario.Territorio)), _
        '    New ListItem("brick", BPadrao.CriptografarTexto(drwConsultorio("simvim_cons_brick").ToString)) _
        '    )

        Dim sb As New System.Text.StringBuilder

        sb.Append("javascript:OpenModal('")
        sb.Append(strPagina)

        If Not (arrParam Is Nothing) AndAlso (arrParam.Length > 0) Then

            sb.Append("?")

            Dim i As Integer

            For i = 0 To arrParam.Length - 1

                If Not i.ToString.Equals("0") Then
                    sb.Append("&")
                End If
                sb.Append(arrParam(i).Text)
                sb.Append("=")
                sb.Append(arrParam(i).Value)

            Next

        End If

        sb.Append("',")
        sb.Append(intLargura)
        sb.Append(",")
        sb.Append(intAltura)
        sb.Append(");")

        If Not blnRetorno Then
            sb.Append(" return false;")
        End If

        lnk.Attributes.Add("onClick", sb.ToString)

    End Sub
    Public Shared Sub jsConfirm(ByVal btn As System.Web.UI.WebControls.ImageButton, ByVal strMensagem As String)

        Dim sb As New System.Text.StringBuilder

        sb.Append("javascript:return confirm('")
        sb.Append(strMensagem)
        sb.Append("');")

        btn.Attributes.Add("onClick", sb.ToString)

    End Sub

    Public Shared Sub jsAlertStartUp(ByVal page As System.Web.UI.Page, ByVal strMensagem As String)

        Dim sb As New System.Text.StringBuilder
        sb.Append("<script>alert('")
        sb.Append(strMensagem)
        sb.Append("');</script>")

        page.RegisterStartupScript("jsAlert", sb.ToString)

    End Sub

    Public Shared Sub jsClose(ByVal page As System.Web.UI.Page)

        page.RegisterStartupScript("jsClose", "<script>window.close();</script>")

    End Sub

    Public Shared Sub jsAlertImagem(ByVal img As System.Web.UI.WebControls.Image, ByVal strMensagem As String)

        Dim sb As New System.Text.StringBuilder

        sb.Append("javascript:alert('")
        sb.Append(strMensagem)
        sb.Append("');")

        img.Attributes.Add("onClick", sb.ToString)

    End Sub

    Public Shared Sub jsAlertBotao(ByVal btn As System.Web.UI.WebControls.ImageButton, ByVal strMensagem As String, Optional ByVal blnRetorno As Boolean = True)

        Dim sb As New System.Text.StringBuilder

        sb.Append("javascript:alert('")
        sb.Append(strMensagem)
        sb.Append("');")

        If Not blnRetorno Then
            sb.Append(" return false;")
        End If

        btn.Attributes.Add("onClick", sb.ToString)

    End Sub
#End Region
    Public Overloads Shared Sub GerarLoadNovo()

        Dim oResponse As System.Web.HttpResponse = System.Web.HttpContext.Current.Response

        Dim sb As New System.Text.StringBuilder


        sb.Append("<div id='mydiv' style='position:absolute;font-family:Trebuchet MS;font-size:13px;background-color:#fffffa;border:solid 2px green;width:250px;height:23px;padding:5px;text-align:center;'>")
        sb.Append("<table cellpadding='0' cellspacing='0' width='100%'>")
        sb.Append("<tr align='center'><td><img alt='' src='../imagens/Carregando.gif' height='22px' style='vertical-align:text-bottom' />&nbsp&nbsp<font style='font-family:Trebuchet MS;font-size:13px;'>Working on your request...</font></td></tr>")
        sb.Append("</table>")
        sb.Append("</div>")

        sb.Append("<script language='javascript'>;")
        sb.Append("var top;var left;")
        sb.Append("function DefineTela(){var x = document.body.clientWidth;var y = document.body.clientHeight;left = ((x/2)- 135);top = ((y/2)- 50);StartShowWait();}")
        sb.Append("function StartShowWait(){mydiv.style.top = top;mydiv.style.left = left;mydiv.style.visibility = 'visible';}")
        sb.Append("function HideWait(){mydiv.style.visibility ='hidden';window.clearInterval();}")
        sb.Append("DefineTela();</script>")

        oResponse.Write(sb.ToString)

        oResponse.Flush()

        Thread.Sleep(1)
    End Sub

    Public Shared Sub EncerrarLoadNovo()

        Dim oResponse As System.Web.HttpResponse = System.Web.HttpContext.Current.Response

        oResponse.Write("<script>HideWait();</script>")

    End Sub

    Public Overloads Shared Sub MensagemCarregando(ByVal opcao As Carregar, ByVal opcaoMensagem As Mensagem)

        If opcao = Carregar.Sim Then
            Dim oResponse As System.Web.HttpResponse = System.Web.HttpContext.Current.Response

            Dim sb As New System.Text.StringBuilder

            sb.Append("<div id='mydiv' style='DISPLAY: inline; FONT-WEIGHT: bold; Z-INDEX: 1; LEFT: 5px; WIDTH: 100px; COLOR: white; FONT-SIZE: 9px; FONT-FAMILY: Verdana; POSITION: absolute; TOP: 5px; HEIGHT: 15px; BACKGROUND-COLOR: red'>")
            sb.Append(opcaoMensagem.ToString)
            sb.Append("</div>")
            sb.Append("<script>mydiv.innerText = '';</script>")

            sb.Append("<script language=javascript>;")
            sb.Append("var dots = 0;var dotmax = 6;function ShowWait()")
            sb.Append("{var output; output = ' " & opcaoMensagem.ToString & "';dots++;if(dots>=dotmax)dots=1;")
            sb.Append("for(var x = 0;x < dots;x++){output += '.';}mydiv.innerText =  output;}")
            sb.Append("function StartShowWait(){mydiv.style.visibility = 'visible';window.setInterval('ShowWait()',1000);}")
            sb.Append("function HideWait(){mydiv.style.visibility ='hidden';window.clearInterval();}")
            sb.Append("StartShowWait();</script>")

            oResponse.Write(sb.ToString)

            oResponse.Flush()

            Thread.Sleep(1)
        Else
            Dim oResponse As System.Web.HttpResponse = System.Web.HttpContext.Current.Response

            oResponse.Write("<script>HideWait();</script>")
        End If
    End Sub
End Class
