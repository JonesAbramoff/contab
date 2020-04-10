VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassApontamentoProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvarlNumIntDocPO As Long
Private mvardtData As Date
Private mvardPercConcluido As Double
Private mvardQuantidade As Double
Private mvariConcluido As Integer
Private mvarsObservacao As String

Public Property Let lNumIntDocPO(ByVal vData As Long)
    mvarlNumIntDocPO = vData
End Property

Public Property Get lNumIntDocPO() As Long
    lNumIntDocPO = mvarlNumIntDocPO
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dPercConcluido(ByVal vData As Double)
    mvardPercConcluido = vData
End Property

Public Property Get dPercConcluido() As Double
    dPercConcluido = mvardPercConcluido
End Property

Public Property Let iConcluido(ByVal vData As Integer)
    mvariConcluido = vData
End Property

Public Property Get iConcluido() As Integer
    iConcluido = mvariConcluido
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property
