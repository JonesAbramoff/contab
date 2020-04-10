VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassiFoodPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public codPedido As String
Public idPedidoCurto As Long
Public togo As Boolean
Public dataEntrega As Date
Public vlrPratos As Double
Public vlrTaxa As Double
Public vlrDesconto As Double
Public vlrTotal As Double
Public obsPedido As String
Public condicaoPgto As String
Public pagamentos As New Collection
Public telefones As New Collection
Public vlrTroco As Double
Public codCampanha As String
Public idCliente As String
Public nome As String
Public email As String
Public clienteNovo As Boolean
Public referenciaGuia As String
Public referenciaXy As String
Public tipoLogradouro As String
Public logradouro As String
Public bairro As String
Public logradouroNum As String
Public complemento As String
Public referencia As String
Public cidade As String
Public estado As String
Public pais As String
Public cep As String
Public locale As String
Public codFornecedor As String
Public codEmpresa As String
Public nomeFornecedor As String
Public status As String
Public dataAlteracaoStatus As Date
Public nomeAtendente As String
Public dataPedidoComanda As Date
Public dataPrevista As Date
Public agendado As String
Public colItens As New Collection

