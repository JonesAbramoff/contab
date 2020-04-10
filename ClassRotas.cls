VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsCodigo As String
Private mvarlChave As Long
Private mvariFilialEmpresa As Integer
Private mvarsDescricao As String
Private mvariAtivo As Integer
Private mvarcolVend As New Collection
Private mvarcolPontos As New Collection

'Início Auxiliar
Private mvarsChave As String
Private mvarsFilial As String
Private mvariFilial As Integer

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let sFilial(ByVal vData As String)
    mvarsFilial = vData
End Property

Public Property Get sFilial() As String
    sFilial = mvarsFilial
End Property

Public Property Let sChave(ByVal vData As String)
    mvarsChave = vData
End Property

Public Property Get sChave() As String
    sChave = mvarsChave
End Property
'Fim Auxiliar

Public Property Set colPontos(ByVal vData As Object)
    Set mvarcolPontos = vData
End Property

Public Property Get colPontos() As Collection
    Set colPontos = mvarcolPontos
End Property

Public Property Set colVend(ByVal vData As Object)
    Set mvarcolVend = vData
End Property

Public Property Get colVend() As Collection
    Set colVend = mvarcolVend
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sCodigo(ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
    sCodigo = mvarsCodigo
End Property

Public Property Let lChave(ByVal vData As Long)
    mvarlChave = vData
End Property

Public Property Get lChave() As Long
    lChave = mvarlChave
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property
