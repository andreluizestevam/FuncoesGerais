Imports System
Imports System.Windows.Forms
Imports System.Web.UI.WebControls
Imports System.Globalization

Public Enum OpcaoConsiste
    VerNulo = 1
    VerNumerico = 2
    VerData = 3
    verCep = 4
    verLen = 5
    verMoeda = 6
    verMes = 7
    verCNPJ = 8
    verCPF = 9
    verHora = 10
End Enum

Namespace Validacao

    Public Class clnValidacao

        Public Shared Function gfRetornaDiaSemana(ByVal intCodigoSemana As Integer) As String
            '============ Array para pegar o dia da semana
            Dim aArray As New ArrayList
            aArray.Add("Domingo")
            aArray.Add("Segunda-Feira")
            aArray.Add("Terça-Feira")
            aArray.Add("Quarta-Feira")
            aArray.Add("Quinta-Feira")
            aArray.Add("Sexta-Feira")
            aArray.Add("Sábado")

            Return aArray.Item(intCodigoSemana).ToString()

        End Function

        Public Shared Function gfIsNumeric(ByVal strValor As String) As Boolean
            Try
                Return Microsoft.VisualBasic.IsNumeric(strValor)
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function gfIsDate(ByVal strData As String) As Boolean
            Try
                Return Microsoft.VisualBasic.IsDate(strData)
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function gfIsDateTime(ByVal strData As String) As Boolean
            Try
                Dim datData As System.DateTime
                datData = DateTime.Parse(strData)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        Public Shared Function gfIsCNPJ(ByVal strCNPJ As String) As Boolean
            Try
                gfIsCNPJ = False

                Dim a, j, d1, d2 As Double
                Dim i As Integer

                If strCNPJ.Length <> 14 Then
                    Return False
                    Exit Function
                End If

                If Not gfIsNumeric(strCNPJ) Then
                    Return False
                    Exit Function
                End If

                a = 0
                i = 0
                d1 = 0
                d2 = 0
                j = 5
                For i = 0 To 11 Step 1
                    a = a + (Convert.ToInt16(strCNPJ.Substring(i, 1)) * j)
                    If j > 2 Then
                        j = j - 1
                    Else
                        j = 9
                    End If
                Next i
                a = a Mod 11

                If a > 1 Then
                    d1 = 11 - a
                Else
                    d1 = 0
                End If

                a = 0
                i = 0
                j = 6
                For i = 0 To 12 Step 1
                    a = a + Convert.ToInt16(strCNPJ.Substring(i, 1)) * j
                    If j > 2 Then
                        j = j - 1
                    Else
                        j = 9
                    End If
                Next i
                a = a Mod 11

                If a > 1 Then
                    d2 = 11 - a
                Else
                    d2 = 0
                End If

                If d1 = Convert.ToInt16(strCNPJ.Substring(12, 1)) And d2 = Convert.ToInt16(strCNPJ.Substring(13, 1)) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Function

        Public Shared Function gfIsCPF(ByVal strCPFCliente As String) As Boolean

            Try
                If Not gfIsNumeric(strCPFCliente) Then
                    Return False
                End If

                '--Declaração das Variáveis
                Dim strCPFOriginal As String = strCPFCliente
                Dim strCPF As String = strCPFOriginal.Substring(0, 9)
                Dim strCPFTemp As String
                Dim intSoma As Integer
                Dim intResto As Integer
                Dim strDigito As String
                Dim intMultiplicador As Integer = 10
                Const constIntMultiplicador As Integer = 11
                Dim i As Integer
                '--------------------------

                For i = 0 To strCPF.ToString.Length - 1
                    intSoma += CInt(strCPF.ToString.Chars(i).ToString) * intMultiplicador
                    intMultiplicador -= 1
                Next

                If (intSoma Mod constIntMultiplicador) < 2 Then
                    intResto = 0
                Else
                    intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
                End If

                strDigito = Convert.ToString(intResto)
                intSoma = 0

                strCPFTemp = strCPF & strDigito
                intMultiplicador = 11

                For i = 0 To strCPFTemp.Length - 1
                    intSoma += CInt(strCPFTemp.Chars(i).ToString) * intMultiplicador
                    intMultiplicador -= 1
                Next

                If (intSoma Mod constIntMultiplicador) < 2 Then
                    intResto = 0
                Else
                    intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
                End If

                strDigito &= intResto

                If strDigito = strCPFOriginal.Substring(9) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Function

    End Class

    Namespace ValidacaoWin
        Public Class clnValidacao
            Public Overloads Shared Function gfConsisteCampos(ByVal Controle As System.Windows.Forms.Control, ByVal errProvider As System.Windows.Forms.ErrorProvider, ByVal tipoOpcao As OpcaoConsiste, ByVal strTitulo As String, Optional ByVal intCaracterLen As Integer = 0) As Boolean

                Select Case tipoOpcao

                    Case OpcaoConsiste.VerNulo

                        If TypeOf Controle Is Windows.Forms.TextBox Then

                            If Controle.Text = String.Empty Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Preenchido!!!")
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf Controle Is Windows.Forms.MaskedTextBox Then
                            If Controle.Text.Trim.Replace("/", "").Replace(",", "").Replace("-", "").Replace(".", "").Trim = String.Empty Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Preenchido!!!")
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf Controle Is Windows.Forms.ComboBox Then

                            If Controle.Text = String.Empty Or Controle.Text = "Selecione" Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!")
                                'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf Controle Is RichTextBox Then

                            If Controle.Text = String.Empty Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!")
                                'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf Controle Is Windows.Forms.ListBox Then

                            If Controle.Text = String.Empty Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!")
                                'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf Controle Is Windows.Forms.DateTimePicker Then

                            If Controle.Text = String.Empty Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!")
                                'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & " deve ser Selecionado!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        End If
                    Case OpcaoConsiste.VerNumerico

                        If Not Controle.Text = "" Then
                            If Not Validacao.clnValidacao.gfIsNumeric(Controle.Text) Then
                                errProvider.SetError(Controle, "O Campo: " & Controle.Tag.ToString() & " deve ser Numérico!!!")
                                'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & " deve ser Numérico!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Controle.Focus()
                                Return False
                            Else
                                errProvider.SetError(Controle, "")
                                Controle.BackColor = System.Drawing.Color.White
                                Return True
                            End If
                        Else
                            Return True
                        End If

                    Case OpcaoConsiste.VerData
                        If Microsoft.VisualBasic.IsDate(Controle.Text) = False Then
                            errProvider.SetError(Controle, "Campo: " & Controle.Tag.ToString() & " Formato Inválido!!!")
                            'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & "!!! Formato Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            errProvider.SetError(Controle, "")
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verHora
                        Dim regTime As New System.Text.RegularExpressions.Regex("^(([0-9])|([0-1][0-9])|([2][0-3])):(([0-9])|([0-5][0-9]))$")
                        Dim vldTime As System.Text.RegularExpressions.Match = regTime.Match(Controle.Text)

                        If Not vldTime.Success Then
                            errProvider.SetError(Controle, "Campo: " & Controle.Tag.ToString() & " Formato Inválido!!!")
                            'MessageBox.Show("O Campo: " & Controle.Tag.ToString() & "!!! Formato Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            errProvider.SetError(Controle, "")
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If
                    Case OpcaoConsiste.verCep
                        'If Not gfIsNumeric(Left(Controle.Text, 5)) Or Mid(Controle.Text, 6, 1) <> "-" Or Not gfIsNumeric(Right(Controle.Text, 2)) Then
                        If Not Validacao.clnValidacao.gfIsNumeric(Controle.Text.Substring(0, 5)) Or Controle.Text.Substring(6, 1) <> "-" Or Not Validacao.clnValidacao.gfIsNumeric(Controle.Text.Substring(Controle.Text.Length - 2, 2)) Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! Formato Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            errProvider.SetError(Controle, "")
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verLen
                        If Controle.Text.Length <> intCaracterLen Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! Formato Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            errProvider.SetError(Controle, "")
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verMoeda
                        If Microsoft.VisualBasic.IsNumeric(Controle.Text) = False Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! Formato Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verMes
                        If CType(Controle.Text, Double) > 12 Or CType(Controle.Text, Double) < 1 Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! Mês Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verCNPJ

                        Dim blnResultado As Boolean
                        Dim strCNPJ As String
                        strCNPJ = Controle.Text
                        strCNPJ = strCNPJ.Replace(".", "")
                        strCNPJ = strCNPJ.Replace("/", "")
                        strCNPJ = strCNPJ.Replace("-", "")

                        If Not Validacao.clnValidacao.gfIsNumeric(strCNPJ) Or strCNPJ.Length <> 14 Then
                            blnResultado = False


                        End If

                        Dim Conta As Integer, Soma As Long, Passo As Integer
                        Dim Digito1 As Integer, Digito2 As Integer, Flag As Integer

                        strCNPJ = strCNPJ.Trim

                        For Passo = 5 To 6
                            Soma = 0
                            Flag = Passo

                            For Conta = 1 To Passo + 7
                                Soma = CType(Soma + (Microsoft.VisualBasic.Val(strCNPJ.Substring(Conta, 1)) * Flag), Integer)
                                If Flag > 2 Then
                                    Flag = Flag - 1
                                Else
                                    Flag = 9
                                End If
                            Next

                            Soma = Soma Mod 11

                            If Passo = 5 Then
                                If Soma > 1 Then
                                    Digito1 = CType(11 - Soma, Integer)
                                Else
                                    Digito1 = 0
                                End If
                            End If
                            If Passo = 6 Then
                                If Soma > 1 Then
                                    Digito2 = CType(11 - Soma, Integer)
                                Else
                                    Digito2 = 0
                                End If
                            End If
                        Next

                        If (Digito1 = Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCNPJ, 13, 1)) And Digito2 = Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCNPJ, 14, 1))) Then
                            blnResultado = True
                        Else
                            blnResultado = False
                        End If


                        If blnResultado = False Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! CNPJ Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verCPF

                        Dim blnResultado As Boolean
                        Dim strCPF As String
                        strCPF = Controle.Text
                        strCPF = strCPF.Replace(".", "")
                        strCPF = strCPF.Replace("/", "")
                        strCPF = strCPF.Replace("-", "")

                        If Not Validacao.clnValidacao.gfIsNumeric(strCPF) Or strCPF.Length <> 11 Then
                            blnResultado = False
                        Else

                            Dim Conta As Integer, Soma As Integer, Resto As Integer, Passo As Integer

                            strCPF = strCPF.Trim

                            For Passo = 11 To 12
                                Soma = 0
                                For Conta = 1 To Passo - 2
                                    Soma = CType(Soma + Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCPF, Conta, 1)) * (Passo - Conta), Integer)
                                Next

                                Resto = CType(11 - (Soma - (Convert.ToInt32(Soma / 11) * 11)), Integer)

                                If Resto = 10 Or Resto = 11 Then Resto = 0

                                If Resto <> Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCPF, Passo - 1, 1)) Then
                                    blnResultado = False
                                Else
                                    blnResultado = True
                                End If
                            Next

                        End If

                        If blnResultado = False Then
                            MessageBox.Show("Campo: " & Controle.Tag.ToString() & "!!! CPF Inválido!!!", strTitulo, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                            Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Controle.Focus()
                            Return False
                        Else
                            Controle.BackColor = System.Drawing.Color.White
                            Return True
                        End If

                End Select

            End Function

        End Class
    End Namespace

    Namespace ValidacaoWeb
        Public Class clnValidacao
            Public Overloads Shared Function gfConsisteCampos(ByVal controle As System.Web.UI.Control, ByVal tipoOpcao As OpcaoConsiste, Optional ByVal intCaracterLen As Integer = 0) As Boolean
                Select Case tipoOpcao

                    Case OpcaoConsiste.VerNulo
                        If TypeOf controle Is Web.UI.WebControls.TextBox Then

                            If DirectCast(controle, Web.UI.WebControls.TextBox).Text = String.Empty Then
                                DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Return False
                            Else
                                DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                                Return True
                            End If

                        ElseIf TypeOf controle Is Web.UI.WebControls.DropDownList Then

                            If DirectCast(controle, DropDownList).SelectedItem.Text <= String.Empty Or DirectCast(controle, DropDownList).SelectedItem.Text = "Selecione" Then
                                DirectCast(controle, DropDownList).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Return False
                            Else
                                DirectCast(controle, DropDownList).BackColor = System.Drawing.Color.FromName("#f0f0e0")
                                Return True
                            End If


                        ElseIf TypeOf controle Is Web.UI.WebControls.ListBox Then

                            If DirectCast(controle, Web.UI.WebControls.ListBox).SelectedItem.Text = String.Empty Then
                                DirectCast(controle, Web.UI.WebControls.ListBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Return False
                            Else
                                DirectCast(controle, Web.UI.WebControls.ListBox).BackColor = System.Drawing.Color.White
                                Return True
                            End If


                        End If

                    Case OpcaoConsiste.VerNumerico

                        If Not DirectCast(controle, Web.UI.WebControls.TextBox).Text = String.Empty Then
                            If Not Validacao.clnValidacao.gfIsNumeric(DirectCast(controle, Web.UI.WebControls.TextBox).Text) Then
                                DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                                Return False
                            Else
                                DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                                Return True
                            End If
                        Else
                            Return True
                        End If

                    Case OpcaoConsiste.VerData
                        If Validacao.clnValidacao.gfIsDate(DirectCast(controle, Web.UI.WebControls.TextBox).Text) = False Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verHora
                        Dim regTime As New System.Text.RegularExpressions.Regex("^(([0-9])|([0-1][0-9])|([2][0-3])):(([0-9])|([0-5][0-9]))$")
                        Dim vldTime As System.Text.RegularExpressions.Match = regTime.Match(DirectCast(controle, Web.UI.WebControls.TextBox).Text)

                        If Not vldTime.Success Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verCep
                        'If Not gfIsNumeric(Left(Controle.Text, 5)) Or Mid(Controle.Text, 6, 1) <> "-" Or Not gfIsNumeric(Right(Controle.Text, 2)) Then
                        If Not Validacao.clnValidacao.gfIsNumeric(DirectCast(controle, Web.UI.WebControls.TextBox).Text.Substring(0, 5)) Or DirectCast(controle, Web.UI.WebControls.TextBox).Text.Substring(6, 1) <> "-" Or Not Validacao.clnValidacao.gfIsNumeric(DirectCast(controle, Web.UI.WebControls.TextBox).Text.Substring(DirectCast(controle, Web.UI.WebControls.TextBox).Text.Length - 2, 2)) Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verLen
                        If DirectCast(controle, Web.UI.WebControls.TextBox).Text.Length <> intCaracterLen Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verMoeda
                        If Validacao.clnValidacao.gfIsNumeric(DirectCast(controle, Web.UI.WebControls.TextBox).Text) = False Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verMes
                        If CType(DirectCast(controle, Web.UI.WebControls.TextBox).Text, Double) > 12 Or CType(DirectCast(controle, Web.UI.WebControls.TextBox).Text, Double) < 1 Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verCNPJ

                        Dim blnResultado As Boolean
                        Dim strCNPJ As String
                        strCNPJ = DirectCast(controle, Web.UI.WebControls.TextBox).Text
                        strCNPJ = strCNPJ.Replace(".", "")
                        strCNPJ = strCNPJ.Replace("/", "")
                        strCNPJ = strCNPJ.Replace("-", "")

                        If Not Validacao.clnValidacao.gfIsNumeric(strCNPJ) Or strCNPJ.Length <> 14 Then
                            blnResultado = False
                        End If

                        Dim Conta As Integer, Soma As Long, Passo As Integer
                        Dim Digito1 As Integer, Digito2 As Integer, Flag As Integer

                        strCNPJ = strCNPJ.Trim

                        For Passo = 5 To 6
                            Soma = 0
                            Flag = Passo

                            For Conta = 1 To Passo + 7
                                Soma = CType(Soma + (Microsoft.VisualBasic.Val(strCNPJ.Substring(conta, 1)) * Flag), Integer)
                                If Flag > 2 Then
                                    Flag = Flag - 1
                                Else
                                    Flag = 9
                                End If
                            Next

                            Soma = Soma Mod 11

                            If Passo = 5 Then
                                If soma > 1 Then
                                    Digito1 = CType(11 - soma, Integer)
                                Else
                                    Digito1 = 0
                                End If
                            End If
                            If Passo = 6 Then
                                If soma > 1 Then
                                    Digito2 = CType(11 - soma, Integer)
                                Else
                                    Digito2 = 0
                                End If
                            End If
                        Next

                        If (Digito1 = Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCNPJ, 13, 1)) And Digito2 = Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCNPJ, 14, 1))) Then
                            blnResultado = True
                        Else
                            blnResultado = False
                        End If


                        If blnResultado = False Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                    Case OpcaoConsiste.verCPF

                        Dim blnResultado As Boolean
                        Dim strCPF As String
                        strCPF = DirectCast(controle, Web.UI.WebControls.TextBox).Text
                        strCPF = strCPF.Replace(".", "")
                        strCPF = strCPF.Replace("/", "")
                        strCPF = strCPF.Replace("-", "")

                        If Not Validacao.clnValidacao.gfIsNumeric(strCPF) Or strCPF.Length <> 11 Then
                            blnResultado = False
                        Else

                            Dim Conta As Integer, Soma As Integer, Resto As Integer, Passo As Integer

                            strCPF = strCPF.Trim

                            For Passo = 11 To 12
                                Soma = 0
                                For Conta = 1 To Passo - 2
                                    Soma = CType(Soma + Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCPF, Conta, 1)) * (Passo - Conta), Integer)
                                Next

                                Resto = CType(11 - (Soma - (Convert.ToInt32(Soma / 11) * 11)), Integer)

                                If Resto = 10 Or Resto = 11 Then Resto = 0

                                If Resto <> Microsoft.VisualBasic.Val(Microsoft.VisualBasic.Mid(strCPF, Passo - 1, 1)) Then
                                    blnResultado = False
                                Else
                                    blnResultado = True
                                End If
                            Next

                        End If

                        If blnResultado = False Then
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            Return False
                        Else
                            DirectCast(controle, Web.UI.WebControls.TextBox).BackColor = System.Drawing.Color.White
                            Return True
                        End If

                End Select

            End Function

        End Class
    End Namespace

End Namespace



