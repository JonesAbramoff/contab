VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmCodigoValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvardValor As Double
Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property


Public Property Get dValor() As Double
    dValor = mvardValor
End Property



Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property


Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



