VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBloqueioLiberacaoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlCodPedido As Long
Private mvariSeqBloqueio As Integer
Private mvarsNomeReduzidoCliente As String
Private mvardtDataEmissao As Date
Private mvardValorPedido As Double
Private mvariTipoDeBloqueio As Integer
Private mvarsNomeReduzidoTipoBloqueio As String
Private mvarsUsuario As String
Private mvardtDataBloqueio As Date
Private mvariMarcado As Integer
Private mvarsObservacao As String

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iMarcado(ByVal vData As Integer)
    mvariMarcado = vData
End Property

Public Property Get iMarcado() As Integer
    iMarcado = mvariMarcado
End Property

Public Property Let dtDataBloqueio(ByVal vData As Date)
    mvardtDataBloqueio = vData
End Property

Public Property Get dtDataBloqueio() As Date
    dtDataBloqueio = mvardtDataBloqueio
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let sNomeReduzidoTipoBloqueio(ByVal vData As String)
    mvarsNomeReduzidoTipoBloqueio = vData
End Property

Public Property Get sNomeReduzidoTipoBloqueio() As String
    sNomeReduzidoTipoBloqueio = mvarsNomeReduzidoTipoBloqueio
End Property

Public Property Let iTipoDeBloqueio(ByVal vData As Integer)
    mvariTipoDeBloqueio = vData
End Property

Public Property Get iTipoDeBloqueio() As Integer
    iTipoDeBloqueio = mvariTipoDeBloqueio
End Property

Public Property Let dValorPedido(ByVal vData As Double)
    mvardValorPedido = vData
End Property

Public Property Get dValorPedido() As Double
    dValorPedido = mvardValorPedido
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sNomeReduzidoCliente(ByVal vData As String)
    mvarsNomeReduzidoCliente = vData
End Property

Public Property Get sNomeReduzidoCliente() As String
    sNomeReduzidoCliente = mvarsNomeReduzidoCliente
End Property

Public Property Let iSeqBloqueio(ByVal vData As Integer)
    mvariSeqBloqueio = vData
End Property

Public Property Get iSeqBloqueio() As Integer
    iSeqBloqueio = mvariSeqBloqueio
End Property

Public Property Let lCodPedido(ByVal vData As Long)
    mvarlCodPedido = vData
End Property

Public Property Get lCodPedido() As Long
    lCodPedido = mvarlCodPedido
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



