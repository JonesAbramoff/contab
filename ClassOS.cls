VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsCodigo As String
Private mvardtDataEmissao As Date
Private mvarcolItens As New Collection
Private mvariNumItens As Integer
Private mvariNumItensBaixados As Integer
Private mvarlCodigoNumerico As Long
Private mvariStatus As Integer
Private mvarlCodSolSRV As Long
Private mvarsProduto As String
Private mvarsLote As String
Private mvariFilialOP As Integer
Public mvarobjTela As Object
Private mvarlCodPedSRV As Long
Private mvariFilialPedSRV As Integer
Private mvarsServico As String
Private mvarlTipo As Long

Public Property Let lTipo(ByVal vData As Long)
    mvarlTipo = vData
End Property

Public Property Get lTipo() As Long
    lTipo = mvarlTipo
End Property

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
End Property

Public Property Let iFilialPedSRV(ByVal vData As Integer)
    mvariFilialPedSRV = vData
End Property

Public Property Get iFilialPedSRV() As Integer
    iFilialPedSRV = mvariFilialPedSRV
End Property

Public Property Let lCodPedSRV(ByVal vData As Long)
    mvarlCodPedSRV = vData
End Property

Public Property Get lCodPedSRV() As Long
    lCodPedSRV = mvarlCodPedSRV
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
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

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lCodSolSRV(ByVal vData As Long)
    mvarlCodSolSRV = vData
End Property

Public Property Get lCodSolSRV() As Long
    lCodSolSRV = mvarlCodSolSRV
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lCodigoNumerico(ByVal vData As Long)
    mvarlCodigoNumerico = vData
End Property

Public Property Get lCodigoNumerico() As Long
    lCodigoNumerico = mvarlCodigoNumerico
End Property

Public Property Let iNumItensBaixados(ByVal vData As Integer)
    mvariNumItensBaixados = vData
End Property

Public Property Get iNumItensBaixados() As Integer
    iNumItensBaixados = mvariNumItensBaixados
End Property

Public Property Let iNumItens(ByVal vData As Integer)
    mvariNumItens = vData
End Property

Public Property Get iNumItens() As Integer
    iNumItens = mvariNumItens
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property





