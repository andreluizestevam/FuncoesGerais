Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Web

Public Enum Sistema
    GCMONLINEv300 = 0
    SIFP = 1
    GCMONLINE = 3
    WINNER = 6
    SISINCO = 7
    ENQUETE = 8
    SAED = 9
    SISCEM = 11
    SISCOP = 12
    SISFORECAST = 13
    SGO = 14
    SAGI = 15
    SISAD = 16
    SISGEM = 18
    SISPAUD = 19
    SISVENC = 21
    SISAC = 22
    SISLAUDO = 23
    COMISSOES = 24
    CERTIFICADO = 25
    SFF = 26
    AMBIENTE = 27
End Enum

Public Enum TipoConexao
    Producao = 0
    Desenvolvimento = 1
    Prototipo = 2
End Enum

Public Enum FormatoData
    DiaMesAno = 1
    MesDiaAno = 2
    MesDiaAnoHora = 3
End Enum

Public Enum eTipoBuscaMovimento
    Produto = 0
    Cliente = 1
    Marca = 2
    ClasseTerapeutica = 3
End Enum

Public Enum eTipoUnidadeMovimento
    Quantidade = 0
    Real = 1
    Dolar = 2
End Enum

Public Class clnServidor
    Public Shared Function PfDataServidor(ByVal strFormatoData As FormatoData) As String

        Dim strData As String
        Dim DrData As SqlDataReader

        Dim strSql As String = "SELECT DataHora=getdate()"
        DrData = SqlHelper.ExecuteReader(fpuObterStrConn(Sistema.GCMONLINEv300, TipoConexao.Producao), CommandType.Text, strSql)
        If DrData.Read() = True Then
            strData = DrData("DataHora").ToString()
        End If

        Select Case strFormatoData
            Case FormatoData.DiaMesAno
                Return strData.Format("dd/MM/yyyy")
                'Return CStr(Format(CDate(strData), "dd/MM/yyyy"))
            Case FormatoData.MesDiaAno
                Return strData.Format("MM/dd/yyyy")
                'Return CStr(Format(CDate(strData), "MM/dd/yyyy"))
            Case FormatoData.MesDiaAnoHora
                Return strData.Format("MM/dd/yyyy hh:mm:ss")
                'Return CStr(Format(CDate(strData), "MM/dd/yyyy hh:mm:ss"))
        End Select
        DrData.Close()

    End Function
    Public Shared Function PfDataServidor() As DateTime

        Dim strData As DateTime
        Dim DrData As SqlDataReader

        Dim strSql As String = "SELECT DataHora=getdate()"
        DrData = SqlHelper.ExecuteReader(fpuObterStrConn(Sistema.GCMONLINEv300, TipoConexao.Producao), CommandType.Text, strSql)
        If DrData.Read() = True Then
            strData = Convert.ToDateTime(DrData("DataHora").ToString())
        End If

        DrData.Close()
        Return strData

    End Function

    Public Overloads Shared Function fpuObterStrConn() As String

        Dim strValorWebConfig As String = ConfigurationSettings.AppSettings("connectionString")
        Return clnCriptografia.fpuDescriptografar(strValorWebConfig)

    End Function

    Public Overloads Shared Function fpuObterStrConn(ByVal enuSistema As Sistema, ByVal enuOpcao As TipoConexao) As String
        Dim producao As Boolean = True

        If producao Then
            If enuSistema = Sistema.GCMONLINEv300 Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=GCMOnline_V300;user id=USR#PCD; pwd=pcd@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.GCMONLINE Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=GCMONLINE;user id=USR#PCD; pwd=pcd@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.ENQUETE Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=ENQUETE;user id=USR#PCD; pwd=pcd@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SGO Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=SFC;user id=USR#SFC;password=sfc@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISCEM Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=SISCEM ;user id=USR#SISCEM; pwd=siscem@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISCOP Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=SISCOP ;user id=USR#SISCOP; pwd=siscop@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.WINNER Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=WINNER;user id=USR#WINNER;password=winner@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISVENC Then
                Return "Data Source=sdb00.uniaoquimica.com.br;initial Catalog=SISVENC;user id=USR#SISVENC; password=sisvenc@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISGEM Then
                Return "Data Source=sdb00.uniaoquimica.com.br;initial catalog=SISGEM;user id=USR#SISGEM; password=sisgem@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISAC Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=SISAC;user id=USR#SISAC; pwd=sisac@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.COMISSOES Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=COMISSOES;user id=USR#COMISSOES; pwd=comissoes@2014$; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SFF Then
                Return "Data Source=sdb00.uniaoquimica.com.br;Initial Catalog=SFF;user id=USR#SFF; pwd=sff@2014$; Min Pool Size=0; Max Pool Size=10000;"
                ' alterar tambem a DTSX
            ElseIf enuSistema = Sistema.AMBIENTE Then
                Return "Produção"
            End If
        Else
            If enuSistema = Sistema.GCMONLINEv300 Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=GCMOnline_V300;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.GCMONLINE Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=GCMONLINE;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.ENQUETE Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=ENQUETE;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SGO Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=SFC;user id=adm_aestevam;password=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISCEM Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=SISCEM ;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISCOP Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=SISCOP ;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.WINNER Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=WINNER;user id=adm_aestevam;password=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISVENC Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;initial Catalog=SISVENC;user id=adm_aestevam; password=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISGEM Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;initial catalog=SISGEM;user id=adm_aestevam; password=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SISAC Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=SISAC;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.COMISSOES Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=COMISSOES;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.SFF Then
                Return "Data Source=sqdb00.uniaoquimica.com.br;Initial Catalog=SFF;user id=adm_aestevam; pwd=uq2015; Min Pool Size=0; Max Pool Size=10000;"
            ElseIf enuSistema = Sistema.AMBIENTE Then
                Return "QAS"
            End If            
        End If
    End Function

    Public Overloads Shared Function RecuperaStringConexao(ByVal strcon As String) As String
        Dim temp As String
        temp = HttpContext.Current.Server.MapPath(".")

        Dim reader As XmlTextReader = New XmlTextReader("c:/PROJETO/FuncoesGerais/StrCon.xml")
        Do While (reader.Read())
            Select Case reader.NodeType
                Case XmlNodeType.Element 'Exibir o início do elemento.
                    Console.Write("<" + reader.Name)
                    If reader.HasAttributes Then 'Se existirem atributos
                        While reader.MoveToNextAttribute()
                            'Exibir o nome e o valor do atributo.
                            Console.Write(" {0}='{1}'", reader.Name, reader.Value)
                        End While
                    End If
                    Console.WriteLine(">")
                Case XmlNodeType.Text 'Exibir o texto em cada elemento.
                    Console.WriteLine(reader.Value)
                Case XmlNodeType.EndElement 'Exibir o fim do elemento.
                    Console.Write("</" + reader.Name)
                    Console.WriteLine(">")
            End Select
        Loop
        Return "123"

    End Function

End Class
