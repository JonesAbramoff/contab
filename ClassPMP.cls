VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarlCodGeracao As Long
Private mvardtDataGeracao As Date
Private mvarcolItens As New Collection
Private mvariAlterado As Integer
Private mvariSimulacao As Integer
Private mvarsVersao As String
Private mvariUltimoProxOP As Integer

Public Property Let iUltimoProxOP(ByVal vData As Integer)
    mvariUltimoProxOP = vData
End Property

Public Property Get iUltimoProxOP() As Integer
    iUltimoProxOP = mvariUltimoProxOP
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let iSimulacao(ByVal vData As Integer)
    mvariSimulacao = vData
End Property

Public Property Get iSimulacao() As Integer
    iSimulacao = mvariSimulacao
End Property

Public Property Let iAlterado(ByVal vData As Integer)
    mvariAlterado = vData
End Property

Public Property Get iAlterado() As Integer
    iAlterado = mvariAlterado
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lCodGeracao(ByVal vData As Long)
    mvarlCodGeracao = vData
End Property

Public Property Get lCodGeracao() As Long
    lCodGeracao = mvarlCodGeracao
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

