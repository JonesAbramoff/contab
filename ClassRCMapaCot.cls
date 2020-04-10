VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRCMapaCot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsRequisitante As String
Private mvardtPrazo As Date
Private mvardQuantidade As Double
Private mvarsCCL As String
Private mvarsUrgente As String
Private mvarsObs As String

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sRequisitante(ByVal vData As String)
    mvarsRequisitante = vData
End Property

Public Property Get sRequisitante() As String
    sRequisitante = mvarsRequisitante
End Property

Public Property Let dtPrazo(ByVal vData As Date)
    mvardtPrazo = vData
End Property

Public Property Get dtPrazo() As Date
    dtPrazo = mvardtPrazo
End Property

Public Property Let sCCL(ByVal vData As String)
    mvarsCCL = vData
End Property

Public Property Get sCCL() As String
    sCCL = mvarsCCL
End Property

Public Property Let sObs(ByVal vData As String)
    mvarsObs = vData
End Property

Public Property Get sObs() As String
    sObs = mvarsObs
End Property

Public Property Let sUrgente(ByVal vData As String)
    mvarsUrgente = vData
End Property

Public Property Get sUrgente() As String
    sUrgente = mvarsUrgente
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property
