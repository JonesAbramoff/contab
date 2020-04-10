VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMovEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvariTipoMov As Integer
Private mvardtData As Date
Private mvarcolItens As New ColItensMovEstoque
Private mvarlCliente As Long
Private mvarlFornecedor As Long
Private mvarsDocOrigem As String
Private mvarlNumIntDoc As Long
Private mvariOperacao As Integer 'quando preenchido serve para indicar se se trata de uma exclusão de um movimento
Private mvardtHora As Date
Private mvarlRequisitante As Long
Private mvarsObservacao As String
Private mvarobjProjetoInfo As Object
Private mvarobjApontamento As Object
Private mvariFilialCli As Integer
Private mvariFilialForn As Integer

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let iFilialCli(ByVal vData As Integer)
    mvariFilialCli = vData
End Property

Public Property Get iFilialCli() As Integer
    iFilialCli = mvariFilialCli
End Property

Public Property Set objApontamento(ByVal vData As Object)
    Set mvarobjApontamento = vData
End Property

Public Property Get objApontamento() As Object
    Set objApontamento = mvarobjApontamento
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lRequisitante(ByVal vData As Long)
    mvarlRequisitante = vData
End Property

Public Property Get lRequisitante() As Long
    lRequisitante = mvarlRequisitante
End Property

Public Property Let dtHora(ByVal vData As Date)
    mvardtHora = vData
End Property

Public Property Get dtHora() As Date
    dtHora = mvardtHora
End Property

Public Property Let iOperacao(ByVal vData As Integer)
    mvariOperacao = vData
End Property

Public Property Get iOperacao() As Integer
    iOperacao = mvariOperacao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sDocOrigem(ByVal vData As String)
    mvarsDocOrigem = vData
End Property

Public Property Get sDocOrigem() As String
    sDocOrigem = mvarsDocOrigem
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Set colItens(ByVal vData As Object)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As ColItensMovEstoque
    Set colItens = mvarcolItens
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iTipoMov(ByVal vData As Integer)
    mvariTipoMov = vData
End Property

Public Property Get iTipoMov() As Integer
    iTipoMov = mvariTipoMov
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property




