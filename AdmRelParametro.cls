VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmRelParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsParametro As String
Private mvarsParamValor As String

Public Property Let sParamValor(ByVal vData As String)
    mvarsParamValor = vData
End Property

Public Property Get sParamValor() As String
    sParamValor = mvarsParamValor
End Property

Public Property Let sParametro(ByVal vData As String)
    mvarsParametro = vData
End Property

Public Property Get sParametro() As String
    sParametro = mvarsParametro
End Property
