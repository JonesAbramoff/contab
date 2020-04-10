VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeFedLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlLote As Long
Private mvardtData As Date
Private mvardHora As Double
Private mvarsStatus As String
Private mvarlNumIntNF As Long
Private mvarsUsuario As String

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property
Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property

Public Property Let sStatus(ByVal vData As String)
    mvarsStatus = vData
End Property

Public Property Get sStatus() As String
    sStatus = mvarsStatus
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

Public Property Let lLote(ByVal vData As Long)
    mvarlLote = vData
End Property

Public Property Get lLote() As Long
    lLote = mvarlLote
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



