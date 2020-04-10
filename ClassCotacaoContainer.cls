VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvariTipoContainer As Integer
Private mvariQuantidade As Integer

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Let iTipoContainer(ByVal vData As Integer)
    mvariTipoContainer = vData
End Property

Public Property Get iTipoContainer() As Integer
    iTipoContainer = mvariTipoContainer
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


