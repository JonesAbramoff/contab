VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarCodigo As Integer
Private mvarData As Date
Private mvarValor As Double

Public Property Let dValor(ByVal vData As Double)
    mvarValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvarValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvarData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvarData
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvarCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvarCodigo
End Property



