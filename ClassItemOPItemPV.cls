VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOPItemPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocItemOP As Long
Private mvarlNumIntDocItemPV As Long
Private mvarsCodigoOP As String
Private mvarlCodigoPV As Long
Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvardtDataEmissao As Date
Private mvardtDataEntrega As Date
Private mvarsDescProduto As String
Private mvarsNomeCliente As String
Private mvarsNomeFilialCliente As String
Private mvarsNomeFilialEmpresa As String
Private mvariPrioridade As Integer
Private mvardQuantidadeProd As Double

Public Property Let dQuantidadeProd(ByVal vData As Double)
    mvardQuantidadeProd = vData
End Property

Public Property Get dQuantidadeProd() As Double
    dQuantidadeProd = mvardQuantidadeProd
End Property

Public Property Let iPrioridade(ByVal vData As Integer)
    mvariPrioridade = vData
End Property

Public Property Get iPrioridade() As Integer
    iPrioridade = mvariPrioridade
End Property

Public Property Let sNomeFilialEmpresa(ByVal vData As String)
    mvarsNomeFilialEmpresa = vData
End Property

Public Property Get sNomeFilialEmpresa() As String
    sNomeFilialEmpresa = mvarsNomeFilialEmpresa
End Property

Public Property Let sNomeFilialCliente(ByVal vData As String)
    mvarsNomeFilialCliente = vData
End Property

Public Property Get sNomeFilialCliente() As String
    sNomeFilialCliente = mvarsNomeFilialCliente
End Property

Public Property Let sNomeCliente(ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let sDescProduto(ByVal vData As String)
    mvarsDescProduto = vData
End Property

Public Property Get sDescProduto() As String
    sDescProduto = mvarsDescProduto
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let lNumIntDocItemOP(ByVal vData As Long)
    mvarlNumIntDocItemOP = vData
End Property

Public Property Get lNumIntDocItemOP() As Long
    lNumIntDocItemOP = mvarlNumIntDocItemOP
End Property

Public Property Let lNumIntDocItemPV(ByVal vData As Long)
    mvarlNumIntDocItemPV = vData
End Property

Public Property Get lNumIntDocItemPV() As Long
    lNumIntDocItemPV = mvarlNumIntDocItemPV
End Property

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

Public Property Let lCodigoPV(ByVal vData As Long)
    mvarlCodigoPV = vData
End Property

Public Property Get lCodigoPV() As Long
    lCodigoPV = mvarlCodigoPV
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
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
