VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensSolicSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public gobjInfoUsu As Object

Private mvarlNumIntDoc As Long
Private mvarlNumIntSolicSRV As Long
Private mvarsProduto As String
Private mvarsProdutoDesc As String
Private mvardtDataVenda As Date
Private mvarsServico As String
Private mvarsServicoDesc As String
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvarsSolicitacao As String
Private mvarlGarantia As Long
Private mvariFilialEmpresa As Integer
Private mvarsContrato As String
Private mvariStatusItem As Integer
Private mvariTipo As Integer
Private mvarcolDet As New Collection
Private mvarsReparo As String
Private mvardtDataBaixa As Date

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let sReparo(ByVal vData As String)
    mvarsReparo = vData
End Property

Public Property Get sReparo() As String
    sReparo = mvarsReparo
End Property

Public Property Get colDet() As Collection
    Set colDet = mvarcolDet
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iStatusItem(ByVal vData As Integer)
    mvariStatusItem = vData
End Property

Public Property Get iStatusItem() As Integer
    iStatusItem = mvariStatusItem
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lGarantia(ByVal vData As Long)
    mvarlGarantia = vData
End Property

Public Property Get lGarantia() As Long
    lGarantia = mvarlGarantia
End Property

Public Property Let sSolicitacao(ByVal vData As String)
    mvarsSolicitacao = vData
End Property

Public Property Get sSolicitacao() As String
    sSolicitacao = mvarsSolicitacao
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sServicoDesc(ByVal vData As String)
    mvarsServicoDesc = vData
End Property

Public Property Get sServicoDesc() As String
    sServicoDesc = mvarsServicoDesc
End Property

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
End Property

Public Property Let dtDataVenda(ByVal vData As Date)
    mvardtDataVenda = vData
End Property

Public Property Get dtDataVenda() As Date
    dtDataVenda = mvardtDataVenda
End Property

Public Property Let sProdutoDesc(ByVal vData As String)
    mvarsProdutoDesc = vData
End Property

Public Property Get sProdutoDesc() As String
    sProdutoDesc = mvarsProdutoDesc
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntSolicSRV(ByVal vData As Long)
    mvarlNumIntSolicSRV = vData
End Property

Public Property Get lNumIntSolicSRV() As Long
    lNumIntSolicSRV = mvarlNumIntSolicSRV
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

