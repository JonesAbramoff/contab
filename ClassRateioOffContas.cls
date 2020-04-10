VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRateioOffContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsContaInicio As String
Private mvarsContaFim As String
Private mvarlCodigo As Long
Private mvariItem As Integer

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let lCodigo(ByVal vData As Long)
Attribute lCodigo.VB_Description = "Codigo do RateioOff"
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let sContaFim(ByVal vData As String)
    mvarsContaFim = vData
End Property

Public Property Get sContaFim() As String
    sContaFim = mvarsContaFim
End Property

Public Property Let sContaInicio(ByVal vData As String)
    mvarsContaInicio = vData
End Property

Public Property Get sContaInicio() As String
    sContaInicio = mvarsContaInicio
End Property



