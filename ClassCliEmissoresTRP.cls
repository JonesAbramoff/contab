VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCliEmissoresTRP"
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
Private mvarcolCliEmissoresExcTRP As New Collection
Private mvarlCodigo As Long
Private mvarlSuperior As Long
Private mvarobjUsuarioWeb As New ClassTRPUsuarios
Private mvarsHierarquia As String

Public Property Let sHierarquia(ByVal vData As String)
    mvarsHierarquia = vData
End Property

Public Property Get sHierarquia() As String
    sHierarquia = mvarsHierarquia
End Property

Public Property Set objUsuarioWeb(ByVal vData As ClassTRPUsuarios)
    Set mvarobjUsuarioWeb = vData
End Property

Public Property Get objUsuarioWeb() As ClassTRPUsuarios
    Set objUsuarioWeb = mvarobjUsuarioWeb
End Property

Public Property Let lSuperior(ByVal vData As Long)
    mvarlSuperior = vData
End Property

Public Property Get lSuperior() As Long
    lSuperior = mvarlSuperior
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Set colCliEmissoresExcTRP(ByVal vData As Collection)
    Set mvarcolCliEmissoresExcTRP = vData
End Property

Public Property Get colCliEmissoresExcTRP() As Collection
    Set colCliEmissoresExcTRP = mvarcolCliEmissoresExcTRP
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


