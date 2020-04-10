VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoArqICMSProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumIntLivRegESCadProd As Long

Public Property Let lNumIntLivRegESCadProd(ByVal vData As Long)
    mvarlNumIntLivRegESCadProd = vData
End Property

Public Property Get lNumIntLivRegESCadProd() As Long
    lNumIntLivRegESCadProd = mvarlNumIntLivRegESCadProd
End Property
