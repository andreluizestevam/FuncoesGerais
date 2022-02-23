Imports System
Imports System.Windows.Forms
Imports System.Web.UI.WebControls
Imports System.Threading

Public Enum Aguarda
    Sim = 1
    Não = 2
End Enum

Namespace FormularioWin
    Public Class clnFormulario

        Public Shared Function gfInstanciaFormulario(ByVal strCaminhoAplicacao As String, ByVal strNomeFormulario As String) As System.Windows.Forms.Form

            Try
                Dim strDLL As String

                strDLL = strCaminhoAplicacao & "\WORKFLOW_WIN.DLL"

                Dim DLL As System.Reflection.Assembly
                DLL = System.Reflection.Assembly.LoadFrom(strDLL)

                Return DirectCast(DLL.CreateInstance("WORKFLOW_WIN." & strNomeFormulario), System.Windows.Forms.Form)

            Catch ex As Exception

            End Try

        End Function


        Public Shared Sub gfLimpaCamposFormulario(ByVal contin As Control.ControlCollection)

            Dim foundcontrol As Control

            For Each foundcontrol In contin
                Select Case foundcontrol.GetType.ToString
                    Case "System.Windows.Forms.TextBox"
                        gfLimpaCampoUnico(foundcontrol)
                    Case "System.Windows.Forms.RichTextBox"
                        gfLimpaCampoUnico(foundcontrol)
                    Case "System.Windows.Forms."
                        gfLimpaCampoUnico(foundcontrol)
                    Case "System.Windows.Forms.ComboBox"
                        gfLimpaCampoUnico(foundcontrol)
                    Case "System.Windows.Forms.ListBox"
                        gfLimpaCampoUnico(foundcontrol)
                    Case "System.Windows.Forms.RadioButton"
                        gfLimpaCampoUnico(foundcontrol)
                    Case Else

                End Select
                If foundcontrol.Controls.Count <> 0 Then
                    gfLimpaCamposFormulario(foundcontrol.Controls)
                End If
            Next foundcontrol

        End Sub
        Private Shared Sub gfLimpaCampoUnico(ByVal Controle As System.Windows.Forms.Control)

            If TypeOf Controle Is System.Windows.Forms.RadioButton Then
                Dim optTemp As New System.Windows.Forms.RadioButton
                optTemp = DirectCast(Controle, Windows.Forms.RadioButton)
                optTemp.Checked = False
            ElseIf TypeOf Controle Is System.Windows.Forms.CheckBox Then
                Dim chkTemp As New System.Windows.Forms.CheckBox
                chkTemp = DirectCast(Controle, Windows.Forms.CheckBox)
                chkTemp.Checked = False
            Else
                Controle.Text = String.Empty
                Controle.BackColor = System.Drawing.Color.White
            End If

        End Sub
        Public Shared Sub gfMensagemAguarde(ByVal strAguarda As Aguarda, ByVal frmForm As Form)

            Select Case strAguarda
                Case Aguarda.Sim
                    Dim t As New Thread(AddressOf psAguarde)  'Creates the new thread
                    t.Start() 'Begins the execution of the new thread
                    frmForm.Cursor = System.Windows.Forms.Cursors.WaitCursor()
                Case Aguarda.Não
                    frmForm.Cursor = System.Windows.Forms.Cursors.Default
                    frmForm.Focus()
            End Select

        End Sub
        Private Shared Sub psAguarde()
            Dim form As New frmAguarde
            form.Cursor = System.Windows.Forms.Cursors.WaitCursor()
            form.Show()
            form.Cursor = System.Windows.Forms.Cursors.Default()
            form.Refresh()

            form.Hide()
            form.Dispose()
        End Sub
        Public Overloads Shared Function spuMessageBox(ByVal strTitulo As String, ByVal strCorpo As String, ByVal btnBotao As MessageBoxButtons) As DialogResult
            Return MessageBox.Show(strCorpo, strTitulo, btnBotao)
        End Function
        Public Overloads Shared Function spuMessageBox(ByVal strTitulo As String, ByVal strCorpo As String, ByVal btnBotao As MessageBoxButtons, ByVal strTipoIcone As MessageBoxIcon) As DialogResult
            Return MessageBox.Show(strCorpo, strTitulo, btnBotao, strTipoIcone)
        End Function
        Public Shared Function spuMensagemErro(ByVal strMensagem As String, ByVal exErro As Exception) As DialogResult
            Return MessageBox.Show("Ocorreu o Seguinte Erro: " & _
                        Environment.NewLine & strMensagem & Environment.NewLine & Environment.NewLine & _
                        "Erro Sistema:" & Environment.NewLine & exErro.Source & exErro.Message & _
                        Environment.NewLine & "Local: " & exErro.Source, Application.ProductName.ToString(), _
                        MessageBoxButtons.OK, MessageBoxIcon.Error, _
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Function

    End Class
End Namespace

Namespace FormularioWeb
    Public Class clnFormulario

        Public Shared Sub gfHabilitaCamposFormulario(ByVal controlP As System.Web.UI.Control)

            Dim ctl As System.Web.UI.Control
            For Each ctl In controlP.Controls
                If TypeOf ctl Is System.Web.UI.WebControls.TextBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.TextBox).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.DropDownList Then
                    DirectCast(ctl, System.Web.UI.WebControls.DropDownList).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.Button Then
                    DirectCast(ctl, System.Web.UI.WebControls.Button).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.CheckBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.CheckBox).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.CheckBoxList Then
                    DirectCast(ctl, System.Web.UI.WebControls.CheckBoxList).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.DataGrid Then
                    DirectCast(ctl, System.Web.UI.WebControls.DataGrid).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.HyperLink Then
                    DirectCast(ctl, System.Web.UI.WebControls.HyperLink).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.ImageButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.ImageButton).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.LinkButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.LinkButton).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.ListBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.ListBox).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.RadioButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.RadioButton).Enabled = True
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.RadioButtonList Then
                    DirectCast(ctl, System.Web.UI.WebControls.RadioButtonList).Enabled = True
                ElseIf ctl.Controls.Count > 0 Then
                    gfHabilitaCamposFormulario(ctl)
                End If
            Next

        End Sub

        Public Shared Sub gfDesabilitaCamposFormulario(ByVal controlP As System.Web.UI.Control)

            Dim ctl As System.Web.UI.Control
            For Each ctl In controlP.Controls
                If TypeOf ctl Is System.Web.UI.WebControls.TextBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.TextBox).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.DropDownList Then
                    DirectCast(ctl, System.Web.UI.WebControls.DropDownList).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.Button Then
                    DirectCast(ctl, System.Web.UI.WebControls.Button).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.CheckBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.CheckBox).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.CheckBoxList Then
                    DirectCast(ctl, System.Web.UI.WebControls.CheckBoxList).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.DataGrid Then
                    DirectCast(ctl, System.Web.UI.WebControls.DataGrid).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.HyperLink Then
                    DirectCast(ctl, System.Web.UI.WebControls.HyperLink).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.ImageButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.ImageButton).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.LinkButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.LinkButton).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.ListBox Then
                    DirectCast(ctl, System.Web.UI.WebControls.ListBox).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.RadioButton Then
                    DirectCast(ctl, System.Web.UI.WebControls.RadioButton).Enabled = False
                ElseIf TypeOf ctl Is System.Web.UI.WebControls.RadioButtonList Then
                    DirectCast(ctl, System.Web.UI.WebControls.RadioButtonList).Enabled = False
                ElseIf ctl.Controls.Count > 0 Then
                    gfDesabilitaCamposFormulario(ctl)
                End If
            Next

        End Sub

    End Class
End Namespace
