
Public Enum OpcaoFormata
    ForCep = 1
    ForData = 2
    forCNPJ = 3
    forEngagement = 4
    forCPF = 5
    forInscEstadual = 6
    forTelefone = 7
    forHora = 8
End Enum

Namespace Formatacao
    Public Class clnFormatacao
        Public Shared Function gfFormataValor(ByVal strValor As String) As String
            Return strValor.Replace(".", "").Replace(",", ".")
        End Function
    End Class
    Namespace FormatacaoWin
        Public Class clnFormatacao
            Public Shared Sub PsFormataCampo(ByVal Opcao As OpcaoFormata, ByVal txtTexto As Windows.Forms.TextBox, Optional ByVal strKey As String = "")

                If strKey = CChar(Microsoft.VisualBasic.vbBack) Then Exit Sub

                Select Case Opcao

                    Case OpcaoFormata.ForData

                        If txtTexto.Text.Length.Equals(2) Or txtTexto.Text.Length.Equals(5) Then
                            txtTexto.Text = txtTexto.Text & "/"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.forHora
                        If txtTexto.Text.Length.Equals(2) Then
                            txtTexto.Text = txtTexto.Text & ":"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.ForCep
                        If txtTexto.Text.Length.Equals(5) Then
                            txtTexto.Text = txtTexto.Text & "-"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.forCNPJ

                        If txtTexto.Text.Length.Equals(2) Or txtTexto.Text.Length.Equals(6) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                        If txtTexto.Text.Length.Equals(10) Then
                            txtTexto.Text = txtTexto.Text & "/"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                        If txtTexto.Text.Length.Equals(15) Then
                            txtTexto.Text = txtTexto.Text & "-"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.forEngagement

                        If txtTexto.Text.Length.Equals(4) Then
                            txtTexto.Text = txtTexto.Text & "-"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.forCPF

                        If txtTexto.Text.Length.Equals(3) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        ElseIf txtTexto.Text.Length.Equals(7) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        ElseIf txtTexto.Text.Length.Equals(11) Then
                            txtTexto.Text = txtTexto.Text & "-"
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If


                    Case OpcaoFormata.forInscEstadual

                        If txtTexto.Text.Length.Equals(3) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        ElseIf txtTexto.Text.Length.Equals(6) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        ElseIf txtTexto.Text.Length.Equals(9) Then
                            txtTexto.Text = txtTexto.Text & "."
                            txtTexto.SelectionStart = txtTexto.Text.Length + 1
                        End If

                    Case OpcaoFormata.forTelefone

                        'If Len(txtTexto.Text) = 0 Then
                        '    txtTexto.Text = txtTexto.Text & "("
                        '    txtTexto.SelectionStart = Len(txtTexto.Text) + 1
                        'ElseIf Len(txtTexto.Text) = 3 Then
                        '    txtTexto.Text = txtTexto.Text & ")"
                        '    txtTexto.SelectionStart = Len(txtTexto.Text) + 1
                        'ElseIf Len(txtTexto.Text) = 8 Then
                        '    txtTexto.Text = txtTexto.Text & "-"
                        '    txtTexto.SelectionStart = Len(txtTexto.Text) + 1
                        'End If


                End Select

            End Sub

        End Class
    End Namespace

    Namespace FormatacaoWeb
        Public Class clnFormatacao

        End Class
    End Namespace

End Namespace