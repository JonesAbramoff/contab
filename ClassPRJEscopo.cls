VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJEscopo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsDescricao As String
Private mvarsExpectativa As String
Private mvarsFatoresSucesso As String
Private mvarsRestricoes As String
Private mvarsPremissas As String
Private mvarsExclusoesEspecificas As String

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sExpectativa(ByVal vData As String)
    mvarsExpectativa = vData
End Property

Public Property Get sExpectativa() As String
    sExpectativa = mvarsExpectativa
End Property

Public Property Let sFatoresSucesso(ByVal vData As String)
    mvarsFatoresSucesso = vData
End Property

Public Property Get sFatoresSucesso() As String
    sFatoresSucesso = mvarsFatoresSucesso
End Property

Public Property Let sRestricoes(ByVal vData As String)
    mvarsRestricoes = vData
End Property

Public Property Get sRestricoes() As String
    sRestricoes = mvarsRestricoes
End Property

Public Property Let sPremissas(ByVal vData As String)
    mvarsPremissas = vData
End Property

Public Property Get sPremissas() As String
    sPremissas = mvarsPremissas
End Property

Public Property Let sExclusoesEspecificas(ByVal vData As String)
    mvarsExclusoesEspecificas = vData
End Property

Public Property Get sExclusoesEspecificas() As String
    sExclusoesEspecificas = mvarsExclusoesEspecificas
End Property
