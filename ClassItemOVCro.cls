VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOVCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsConcorrente As String
Private mvariCondPagto As Integer
Private mvariEscolhido As Integer

Public Property Let iEscolhido(ByVal vData As Integer)
    mvariEscolhido = vData
End Property

Public Property Get iEscolhido() As Integer
    iEscolhido = mvariEscolhido
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property

Public Property Let sConcorrente(ByVal vData As String)
    mvarsConcorrente = vData
End Property

Public Property Get sConcorrente() As String
    sConcorrente = mvarsConcorrente
End Property


