'Imports GCMONLINE_INT
'Public Class clnRemoto

'    Private Shared strIp As String = "RADIO" '"DMZ03" '"172.16.4.4" 'System.Configuration.ConfigurationSettings.AppSettings("IP")
'    'Private Shared strIp As String = "172.16.8.32" '"DMZ03" '"172.16.4.4" 'System.Configuration.ConfigurationSettings.AppSettings("IP")

'    Public Shared Function RetornaICliente() As ICliente
'        Return DirectCast(Activator.GetObject(GetType(ICliente), "tcp://" & strIp & ":52001/GCMONLINEServico_Cliente"), ICliente)
'    End Function

'    Public Shared Function RetornaICusto() As ICusto
'        Return DirectCast(Activator.GetObject(GetType(ICusto), "tcp://" & strIp & ":52002/GCMONLINEServico_Custo"), ICusto)
'    End Function

'    Public Shared Function RetornaIEstruturaEmpresa() As IEstruturaEmpresa
'        Return DirectCast(Activator.GetObject(GetType(IProduto), "tcp://" & strIp & ":52003/GCMONLINEServico_EstruturaEmpresa"), IEstruturaEmpresa)
'    End Function

'    Public Shared Function RetornaIProduto() As IProduto
'        Return DirectCast(Activator.GetObject(GetType(IProduto), "tcp://" & strIp & ":52004/GCMONLINEServico_Produto"), IProduto)
'    End Function

'    Public Shared Function RetornaISeguranca() As ISeguranca
'        Return DirectCast(Activator.GetObject(GetType(ISeguranca), "tcp://" & strIp & ":52005/GCMONLINEServico_Seguranca"), ISeguranca)
'    End Function

'    Public Shared Function RetornaIUsuario() As IUsuario
'        Return DirectCast(Activator.GetObject(GetType(IUsuario), "tcp://" & strIp & ":52006/GCMONLINEServico_Usuario"), IUsuario)
'    End Function

'    Public Shared Function RetornaIConexao() As IConexao
'        Return DirectCast(Activator.GetObject(GetType(IConexao), "tcp://" & strIp & ":52007/GCMOnlineServico_Conexao"), IConexao)
'    End Function

'    Public Shared Function RetornaIComercial() As IComercial
'        Return DirectCast(Activator.GetObject(GetType(IComercial), "tcp://" & strIp & ":52008/GCMOnlineServico_Comercial"), IComercial)
'    End Function
'End Class
