Imports Microsoft.Practices.EnterpriseLibrary.Security.Cryptography
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient

Public Class clnCriptografia
    Private Const strTipoCriptografia As String = "RijndaelManaged"

    Public Overloads Shared Function fpuCriptografar(ByVal Valor As String) As String
        Try
            Return Cryptographer.EncryptSymmetric(strTipoCriptografia, Valor)
        Catch ex As Exception
            Throw New Exception("Erro ao criptografar valor. " & ex.Message)
        End Try
    End Function
    Public Overloads Shared Function fpuDescriptografar(ByVal Valor As String) As String
        Try
            Return Cryptographer.DecryptSymmetric(strTipoCriptografia, Valor)
        Catch ex As Exception
            Throw New Exception("Erro ao descriptografar valor. " & ex.Message)
        End Try
    End Function

    Public Overloads Shared Function fpuCriptografar(ByVal strTexto As String, ByVal Chave As String) As String

        Try

            Dim oCripto As New DESCryptoServiceProvider

            Dim strVetor As Byte() = System.Text.Encoding.UTF8.GetBytes(System.Configuration.ConfigurationSettings.AppSettings("strVetor").ToString)
            Dim strChave As Byte() = System.Text.Encoding.UTF8.GetBytes(Chave)

            Dim strBytes As Byte()

            oCripto.IV = strVetor
            oCripto.Key = strChave

            Dim enc As ICryptoTransform

            enc = oCripto.CreateEncryptor

            strBytes = enc.TransformFinalBlock(System.Text.Encoding.UTF8.GetBytes(strTexto), 0, strTexto.Length)

            Return Convert.ToBase64String(strBytes)

        Catch ex As Exception

            Throw New Exception("Erro ao criptografar valor. " & ex.Message)

        End Try

    End Function

    Public Overloads Shared Function fpuDescriptografar(ByVal strTexto As String, ByVal Chave As String, Optional ByVal url As Boolean = True) As String

        Try

            If url Then
                strTexto = Replace(strTexto, " ", "+")
            End If

            Dim oCripto As New DESCryptoServiceProvider

            Dim strVetor As Byte() = System.Text.Encoding.UTF8.GetBytes(System.Configuration.ConfigurationSettings.AppSettings("strVetor").ToString)
            Dim strChave As Byte() = System.Text.Encoding.UTF8.GetBytes(Chave)
            Dim strBytes As Byte()

            oCripto.IV = strVetor
            oCripto.Key = strChave

            Dim enc As ICryptoTransform

            enc = oCripto.CreateDecryptor

            Dim Buffer() As Byte = Convert.FromBase64String(strTexto)

            strBytes = enc.TransformFinalBlock(Buffer, 0, Buffer.Length)

            Return System.Text.Encoding.UTF8.GetString(strBytes)

        Catch ex As Exception

            Throw New Exception("Erro ao descriptografar valor. " & ex.Message)

        End Try

    End Function
End Class

