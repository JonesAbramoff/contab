VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProjeto As String
Private mvarsEtapa As String
Private mvarcolPRJCustoReceitaReal As New Collection
Private mvarlNumIntDocPRJ As Long
Private mvarlNumIntDocEtapa As Long
Private mvariTipoOrigem As Integer
Private mvarlNumIntDocOrigem As Long
Private mvarlNumIntDoc As Long
Private mvarsCodigoOP As String
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sCodigoOP(ByVal vData As String)
    mvarsCodigoOP = vData
End Property

Public Property Get sCodigoOP() As String
    sCodigoOP = mvarsCodigoOP
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoOrigem(ByVal vData As Integer)
    mvariTipoOrigem = vData
End Property

Public Property Get iTipoOrigem() As Integer
    iTipoOrigem = mvariTipoOrigem
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Set colPRJCustoReceitaReal(ByVal vData As Collection)
    Set mvarcolPRJCustoReceitaReal = vData
End Property

Public Property Get colPRJCustoReceitaReal() As Collection
    Set colPRJCustoReceitaReal = mvarcolPRJCustoReceitaReal
End Property

Public Property Let sEtapa(ByVal vData As String)
    mvarsEtapa = vData
End Property

Public Property Get sEtapa() As String
    sEtapa = mvarsEtapa
End Property

Public Property Let sProjeto(ByVal vData As String)
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property
