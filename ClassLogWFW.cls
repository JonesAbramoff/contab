VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLogWFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsDoc As String
Private mvarsMsg As String
Private mvardtData As Date
Private mvardHora As Double
Private mvarsUsuario As String
Private mvariTransacao As Integer

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let dHora(ByVal vData As Double)
    mvardHora = vData
End Property

Public Property Get dHora() As Double
    dHora = mvardHora
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sMsg(ByVal vData As String)
    mvarsMsg = vData
End Property

Public Property Get sMsg() As String
    sMsg = mvarsMsg
End Property

Public Property Let sDoc(ByVal vData As String)
    mvarsDoc = vData
End Property

Public Property Get sDoc() As String
    sDoc = mvarsDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property




