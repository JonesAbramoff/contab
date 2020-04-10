VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassApropriacaoInsumosProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvarlNumIntReqProd As Long
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvariTipoInsumo As Integer
Private mvarcolRastreamentoMovto As Collection 'local copy

Public iAutomatico As Integer

Public Property Set colRastreamentoMovto(ByVal vData As Collection)
    Set mvarcolRastreamentoMovto = vData
End Property

Public Property Get colRastreamentoMovto() As Collection
    Set colRastreamentoMovto = mvarcolRastreamentoMovto
End Property

Public Property Let iTipoInsumo(ByVal vData As Integer)
    mvariTipoInsumo = vData
End Property

Public Property Get iTipoInsumo() As Integer
    iTipoInsumo = mvariTipoInsumo
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntReqProd(ByVal vData As Long)
    mvarlNumIntReqProd = vData
End Property

Public Property Get lNumIntReqProd() As Long
    lNumIntReqProd = mvarlNumIntReqProd
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



Private Sub Class_Initialize()
    iAutomatico = APROPINSUMOSPROD_AUTOMATICO
End Sub
