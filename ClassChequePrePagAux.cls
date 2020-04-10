VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequePrePagAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariChequeAtual As Integer
Private mvarcolCheques As New Collection

Public Property Get colCheques() As Collection
    Set colCheques = mvarcolCheques
End Property

Public Property Set colCheques(ByVal vData As Collection)
    Set mvarcolCheques = vData
End Property

Public Property Let iChequeAtual(ByVal vData As Integer)
    mvariChequeAtual = vData
End Property

Public Property Get iChequeAtual() As Integer
    iChequeAtual = mvariChequeAtual
End Property
