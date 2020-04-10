VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOSProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarsCodigo As String
Private mvariItem As Integer
Private mvarsProduto As String
Private mvarsSiglaUM As String
Private mvardQuantidade As Double
Private mvardQuantidadeProd As Double
Private mvariStatus As Integer

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property
Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
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

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property


Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property


Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dQuantidadeProd(ByVal vData As Double)
    mvardQuantidadeProd = vData
End Property

Public Property Get dQuantidadeProd() As Double
    dQuantidadeProd = mvardQuantidadeProd
End Property


Public Property Let sSiglaUM(ByVal vData As String)
    mvarsSiglaUM = vData
End Property

Public Property Get sSiglaUM() As String
    sSiglaUM = mvarsSiglaUM
End Property
