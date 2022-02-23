Imports System.Windows.Forms

Public Class clnDialog

    Public Overloads Shared Function SaveFileDialog(ByVal strTitulo As String, _
                                                    Optional ByVal blnOverwritePrompt As Boolean = False, _
                                                    Optional ByVal strFiltroExtensao As String = "", _
                                                    Optional ByVal strExtensaoDefault As String = "", _
                                                    Optional ByVal blnChecarDiretorioExiste As Boolean = False, _
                                                    Optional ByVal strDiretorioInicial As String = "" _
                                                    ) As String

        Dim Dialogo As New SaveFileDialog

        Dialogo.CheckPathExists = blnChecarDiretorioExiste
        If strTitulo <> String.Empty Then
            Dialogo.Title = strTitulo
        Else
            Dialogo.Title = "Informe o caminho e o nome do arquivo a ser exportado"
        End If
        If strDiretorioInicial.Trim <> String.Empty Then
            Dialogo.InitialDirectory = strDiretorioInicial
        End If
        If strFiltroExtensao.Trim <> String.Empty Then
            Dialogo.Filter = strFiltroExtensao
        End If
        If strExtensaoDefault.Trim <> String.Empty Then
            Dialogo.DefaultExt = strExtensaoDefault
        End If
        Dialogo.OverwritePrompt = blnOverwritePrompt

        Dialogo.ShowDialog()
        Return Dialogo.FileName

    End Function
    Public Overloads Shared Function OpenFileDialog(ByVal strTitulo As String, _
                                                        Optional ByVal strFiltroExtensao As String = "", _
                                                        Optional ByVal strDiretorioInicial As String = "" _
                                                        ) As String

        Dim Dialogo As New OpenFileDialog

        Dialogo.CheckPathExists = True
        If strTitulo <> String.Empty Then
            Dialogo.Title = strTitulo
        Else
            Dialogo.Title = "Informe o caminho do arquivo"
        End If
        If strDiretorioInicial.Trim <> String.Empty Then
            Dialogo.InitialDirectory = strDiretorioInicial
        Else
            Dialogo.InitialDirectory = "C:\"
        End If
        If strFiltroExtensao.Trim <> String.Empty Then
            Dialogo.Filter = strFiltroExtensao
        End If

        Dialogo.ShowDialog()
        Return Dialogo.FileName

    End Function
    Public Overloads Shared Function FolderBrowserDialog(ByVal strTitulo As String) As String

        Dim Dialogo As New FolderBrowserDialog

        Dialogo.ShowNewFolderButton = True

        If strTitulo <> String.Empty Then
            Dialogo.Description = strTitulo
        Else
            Dialogo.Description = "Informe o caminho"
        End If
        Dialogo.RootFolder = Environment.SpecialFolder.MyComputer

        If Dialogo.ShowDialog() = DialogResult.OK Then
            Return Dialogo.SelectedPath
        Else
            Return ""
        End If

    End Function

End Class

