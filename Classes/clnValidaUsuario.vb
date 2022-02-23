Imports System.Security.Cryptography
Imports System.Text
Imports System.Web

Public Class clnValidaUsuario

    Public Function Codificar(entrada As String) As String
        Dim tripledescryptoserviceprovider As New TripleDESCryptoServiceProvider()
        Dim md5cryptoserviceprovider As New MD5CryptoServiceProvider()

        Try
            If entrada.Trim() <> "" Then
                Dim myKey As String = "15a7ml3/40912197"
                tripledescryptoserviceprovider.Key = md5cryptoserviceprovider.ComputeHash(ASCIIEncoding.ASCII.GetBytes(myKey))
                tripledescryptoserviceprovider.Mode = CipherMode.ECB
                Dim desdencrypt As ICryptoTransform = tripledescryptoserviceprovider.CreateEncryptor()
                Dim MyASCIIEncoding As New ASCIIEncoding()
                Dim buff As Byte() = Encoding.ASCII.GetBytes(entrada)

                Return Convert.ToBase64String(desdencrypt.TransformFinalBlock(buff, 0, buff.Length))
            Else
                Return ""
            End If
        Catch exception As Exception
            Throw exception
        Finally
            tripledescryptoserviceprovider = Nothing
            md5cryptoserviceprovider = Nothing
        End Try
    End Function

    Public Function Decodificar(entrada As String) As String
        Dim tripledescryptoserviceprovider As New TripleDESCryptoServiceProvider()
        Dim md5cryptoserviceprovider As New MD5CryptoServiceProvider()

        Try
            If entrada.Trim() <> "" Then
                Dim myKey As String = "15a7ml3/40912197"
                tripledescryptoserviceprovider.Key = md5cryptoserviceprovider.ComputeHash(ASCIIEncoding.ASCII.GetBytes(myKey))
                tripledescryptoserviceprovider.Mode = CipherMode.ECB
                Dim desdencrypt As ICryptoTransform = tripledescryptoserviceprovider.CreateDecryptor()
                Dim buff As Byte() = Convert.FromBase64String(entrada)

                Return ASCIIEncoding.ASCII.GetString(desdencrypt.TransformFinalBlock(buff, 0, buff.Length))
            Else
                Return ""
            End If
        Catch exception As Exception
            Throw exception
        Finally
            tripledescryptoserviceprovider = Nothing
            md5cryptoserviceprovider = Nothing
        End Try
    End Function

    Public Function RecuperaCookie(nomecookie As String) As String
        Dim temp As String = ""
        Try
            Dim cookie As HttpCookie = Request.Cookies(nomecookie)
            temp = Decodificar(cookie.Value.ToString())
        Catch
            temp = "0"
        End Try

        Return temp
    End Function
End Class

