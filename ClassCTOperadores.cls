VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCTOperadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodTipoMO As Integer
Private mvarlNumIntDocCT As Long
Private mvariQuantidade As Integer

Public Property Let iQuantidade(ByVal vData As Integer)
    mvariQuantidade = vData
End Property

Public Property Get iQuantidade() As Integer
    iQuantidade = mvariQuantidade
End Property

Public Property Let lNumIntDocCT(ByVal vData As Long)
    mvarlNumIntDocCT = vData
End Property

Public Property Get lNumIntDocCT() As Long
    lNumIntDocCT = mvarlNumIntDocCT
End Property

Public Property Let iCodTipoMO(ByVal vData As Integer)
    mvariCodTipoMO = vData
End Property

Public Property Get iCodTipoMO() As Integer
    iCodTipoMO = mvariCodTipoMO
End Property

