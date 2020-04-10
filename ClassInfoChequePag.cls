VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoChequePag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "auxiliar para a geracao de cheques automaticos e avulsos"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariSeqCheque As Integer
Private mvarsFavorecido As String
Private mvardValor As Double
Private mvarlNumRealCheque As Long
Private mvariImprimir As Integer 'zero ou 1
Private mvariPortador As Integer

'#########################################
'Inserido por Wagner
Private mvarsObservacao As String
Private mvarsVerso As String

Public Property Let sVerso(ByVal vData As String)
    mvarsVerso = vData
End Property

Public Property Get sVerso() As String
    sVerso = mvarsVerso
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property
'#########################################

Public Property Let iPortador(ByVal vData As Integer)
    mvariPortador = vData
End Property

Public Property Get iPortador() As Integer
    iPortador = mvariPortador
End Property

Public Property Let lNumRealCheque(ByVal vData As Long)
    mvarlNumRealCheque = vData
End Property

Public Property Get lNumRealCheque() As Long
    lNumRealCheque = mvarlNumRealCheque
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sFavorecido(ByVal vData As String)
    mvarsFavorecido = vData
End Property

Public Property Get sFavorecido() As String
    sFavorecido = mvarsFavorecido
End Property

Public Property Let iSeqCheque(ByVal vData As Integer)
    mvariSeqCheque = vData
End Property

Public Property Get iSeqCheque() As Integer
    iSeqCheque = mvariSeqCheque
End Property

Public Property Let iImprimir(ByVal vData As Integer)
    mvariImprimir = vData
End Property

Public Property Get iImprimir() As Integer
    iImprimir = mvariImprimir
End Property



