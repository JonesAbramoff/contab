VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCliEmissoresTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCliente As Long
Private mvariSeq As Integer
Private mvarlFornEmissor As Long
Private mvardPercComissao As Double
Private mvarcolCliEmissoresExcTRV As New Collection
Private mvardPercCI As Double
Private mvarsNumCartao As String
Private mvarsCPF As String
Private mvarlCargo As Long

Public Property Let lCargo(ByVal vData As Long)
    mvarlCargo = vData
End Property

Public Property Get lCargo() As Long
    lCargo = mvarlCargo
End Property

Public Property Let sCPF(ByVal vData As String)
    mvarsCPF = vData
End Property

Public Property Get sCPF() As String
    sCPF = mvarsCPF
End Property

Public Property Let sNumCartao(ByVal vData As String)
    mvarsNumCartao = vData
End Property

Public Property Get sNumCartao() As String
    sNumCartao = mvarsNumCartao
End Property

Public Property Let dPercCI(ByVal vData As Double)
    mvardPercCI = vData
End Property

Public Property Get dPercCI() As Double
    dPercCI = mvardPercCI
End Property

Public Property Set colCliEmissoresExcTRV(ByVal vData As Collection)
    Set mvarcolCliEmissoresExcTRV = vData
End Property

Public Property Get colCliEmissoresExcTRV() As Collection
    Set colCliEmissoresExcTRV = mvarcolCliEmissoresExcTRV
End Property

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let lFornEmissor(ByVal vData As Long)
    mvarlFornEmissor = vData
End Property

Public Property Get lFornEmissor() As Long
    lFornEmissor = mvarlFornEmissor
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property


