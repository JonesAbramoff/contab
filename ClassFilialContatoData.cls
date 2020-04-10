VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilialContatoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvardtData As Date
Private mvariLigar As Integer
Private mvariLigacaoEfetuada As Integer
Private mvarsHistorico As String
Private mvarsCodUsuario As String

Public Property Let sCodUsuario(ByVal vData As String)
    mvarsCodUsuario = vData
End Property

Public Property Get sCodUsuario() As String
    sCodUsuario = mvarsCodUsuario
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iLigacaoEfetuada(ByVal vData As Integer)
    mvariLigacaoEfetuada = vData
End Property

Public Property Get iLigacaoEfetuada() As Integer
    iLigacaoEfetuada = mvariLigacaoEfetuada
End Property

Public Property Let iLigar(ByVal vData As Integer)
    mvariLigar = vData
End Property

Public Property Get iLigar() As Integer
    iLigar = mvariLigar
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property
