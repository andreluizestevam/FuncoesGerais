

'Não esqueça de fazer o Web Reference para o Query Web Service
'http://consultacep.correios.com.br/office2003/Query.asmx
'Neste exemplo o Web Reference Name escolhido foi wsCorreios

'Duas formas de utilizar:

'Forma chamando método de busca
'   Dim objBuscaCep As New cBuscaCEP
'   objBuscaCep.localizarEndereco(strResposta)

'Forma passando no construtor
'   With New cBuscaCEP(strResposta)
'   txtCep.Text = .Cep
'   txtBairro.Text = .Bairro
'   txtCidade.Text = .Cidade
'   txtEstado.Text = .Estado
'   txtLogradouro.Text = .Logradouro
'End With

Imports System
Imports System.Xml
Imports System.Text.RegularExpressions
Namespace Cep


Public Class clnCep

#Region "Atributos"

        Private Shared _ws As wsCorreios.QueryProcessor

        Private _cep As String
        Private _endereco As String
        Private _logradouro As String
        Private _bairro As String
        Private _cidade As String
        Private _estado As String

#End Region

#Region "Propriedades / Acessores"

        Public ReadOnly Property Cep() As String
            Get
                Return _cep
            End Get
        End Property

        Public ReadOnly Property Logradouro() As String
            Get
                Return _logradouro
            End Get
        End Property

        Public ReadOnly Property Bairro() As String
            Get
                Return _bairro
            End Get
        End Property

        Public ReadOnly Property Cidade() As String
            Get
                Return _cidade
            End Get
        End Property

        Public ReadOnly Property Estado() As String
            Get
                Return _estado
            End Get
        End Property

        Private Sub setCep(ByVal Value As String) 'FW1.1 Ñ Permite Get WriteOnly Set ReadOnly
            If Not Value.Length = 9 Then          'com mesmo nome
                Throw New ApplicationException("Cep Inválido!")
            End If
            _cep = Value
        End Sub

#End Region

#Region "Construtores"

        Public Sub New()
            If _ws Is Nothing Then
                _ws = New wsCorreios.QueryProcessor
            End If
        End Sub

        Public Sub New(ByVal cep As String)
            Me.New()
            Me.setCep(cep)
            Call localizarEndereco()
        End Sub

#End Region

#Region "Metodos"

        Private Function consultarWeb() As String
            Dim MemStream As New System.IO.MemoryStream
            Dim XMLwriter As New XmlTextWriter(MemStream, System.Text.Encoding.UTF8)

            With XMLwriter
                .WriteStartDocument()

                .WriteStartElement("QueryPacket", ns:="urn:Microsoft.Search.Query")
                .WriteStartElement("Query")

                .WriteStartElement("Context")
                .WriteStartElement("QueryText")
                .WriteString(_cep)
                .WriteEndElement()
                .WriteEndElement()

                .WriteStartElement("OfficeContext", _
                                   ns:="urn:Microsoft.Search.Query.Office.Context")
                .WriteStartElement("ApplicationContext")
                .WriteStartElement("Name")
                .WriteString("Microsoft Office")
                .WriteEndElement()
                .WriteEndElement()
                .WriteEndElement()

                .WriteEndElement()
                .WriteEndElement()

                .WriteEndDocument()
            End With

            XMLwriter.Flush()
            MemStream.Flush()

            MemStream.Position = 0
            Dim stReader As New IO.StreamReader(MemStream)

            XMLwriter = Nothing
            MemStream = Nothing

            Dim queryText As String = stReader.ReadToEnd.ToString()

            Return _ws.Query(queryText)
        End Function

        Private Function removeTags(ByVal Str As String) As String
            Return Regex.Replace(Str, "(?:<!--.*-->.*?)|(?:<[^>]*>.*?)", "").Trim.ToString
        End Function

        Public Sub localizarEndereco(ByVal cep As String)
            Me.setCep(cep)
            localizarEndereco()
        End Sub

        Private Sub localizarEndereco()
            Dim strResultado As String = consultarWeb()

            If New Regex("Data").IsMatch(strResultado) Then
                Dim regExProcessa As New Regex("(?:<P.*?>(.*?)</P.*?>.*?)")

                With regExProcessa.Matches(strResultado)
                    _logradouro = _
                        removeTags(.Item(1).Value)
                    _bairro = _
                        removeTags(.Item(2).Value)
                    _estado = _
                        removeTags(.Item(3).Value.Split(Convert.ToChar("-"))(0).Trim)
                    _cidade = _
                        removeTags(.Item(3).Value.Split(Convert.ToChar("-"))(1).Trim)
                End With
            Else
                Throw New ApplicationException("CEP não encontrado!")
            End If
        End Sub

#End Region

End Class
End Namespace