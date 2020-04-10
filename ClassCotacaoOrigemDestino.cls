VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoOrigemDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocOrigem As Long
Private mvarsServico As String
Private mvarsOrigem As String
Private mvarsDestino As String

Public Property Let sDestino(ByVal vData As String)
    mvarsDestino = vData
End Property

Public Property Get sDestino() As String
    sDestino = mvarsDestino
End Property

Public Property Let sOrigem(ByVal vData As String)
    mvarsOrigem = vData
End Property

Public Property Get sOrigem() As String
    sOrigem = mvarsOrigem
End Property

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
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



