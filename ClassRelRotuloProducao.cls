VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRotuloProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeq As Integer
Private mvarlNumIntRastreamentoLote As Long
Private mvardPesoLiquido As Double
Private mvardPesoBruto As Double
Private mvarlCodigo As Long
Private mvariFilialOP As Integer
Private mvarsLote As String
Private mvarsProduto As String
Private mvariQtdeEmb As Integer
'################################
'Inserido por Wagner
Private mvariImprimir As Integer
Private mvardtDataFabricacao As Date
Private mvardtDataValidade As Date

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let dtDataFabricacao(ByVal vData As Date)
    mvardtDataFabricacao = vData
End Property

Public Property Get dtDataFabricacao() As Date
    dtDataFabricacao = mvardtDataFabricacao
End Property

Public Property Let iImprimir(ByVal vData As Integer)
    mvariImprimir = vData
End Property

Public Property Get iImprimir() As Integer
    iImprimir = mvariImprimir
End Property
'################################

Public Property Let iQtdeEmb(ByVal vData As Integer)
    mvariQtdeEmb = vData
End Property

Public Property Get iQtdeEmb() As Integer
    iQtdeEmb = mvariQtdeEmb
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntRastreamentoLote(ByVal vData As Long)
    mvarlNumIntRastreamentoLote = vData
End Property

Public Property Get lNumIntRastreamentoLote() As Long
    lNumIntRastreamentoLote = mvarlNumIntRastreamentoLote
End Property

Public Property Let dPesoLiquido(ByVal vData As Double)
    mvardPesoLiquido = vData
End Property

Public Property Get dPesoLiquido() As Double
    dPesoLiquido = mvardPesoLiquido
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property
