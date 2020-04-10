VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOPCargaInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntItemOP As Long
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvardKgCarga As Double
Private mvariSeq As Integer
Private mvarlFornecedor As Long
Private mvarsLote As String

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property
Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let dKgCarga(ByVal vData As Double)
    mvardKgCarga = vData
End Property

Public Property Get dKgCarga() As Double
    dKgCarga = mvardKgCarga
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntItemOP(ByVal vData As Long)
    mvarlNumIntItemOP = vData
End Property

Public Property Get lNumIntItemOP() As Long
    lNumIntItemOP = mvarlNumIntItemOP
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



