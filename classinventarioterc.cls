VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInventarioTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClassInventarioTercProd"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariTipoTerc As Integer
Private mvarlCodTerc As Long
Private mvariFilialTerc As Integer
Private mvardtData As Date
Private mvarcolInventarioTercProd As New Collection
Private mvariCodEscaninho As Integer

Public Property Let iCodEscaninho(ByVal vData As Integer)
    mvariCodEscaninho = vData
End Property

Public Property Get iCodEscaninho() As Integer
    iCodEscaninho = mvariCodEscaninho
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialTerc(ByVal vData As Integer)
    mvariFilialTerc = vData
End Property

Public Property Get iFilialTerc() As Integer
    iFilialTerc = mvariFilialTerc
End Property

Public Property Let lCodTerc(ByVal vData As Long)
    mvarlCodTerc = vData
End Property

Public Property Get lCodTerc() As Long
    lCodTerc = mvarlCodTerc
End Property

Public Property Let iTipoTerc(ByVal vData As Integer)
    mvariTipoTerc = vData
End Property

Public Property Get iTipoTerc() As Integer
    iTipoTerc = mvariTipoTerc
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Set colInventarioTercProd(ByVal vData As Object)
    Set mvarcolInventarioTercProd = vData
End Property

Public Property Get colInventarioTercProd() As Collection
    Set colInventarioTercProd = mvarcolInventarioTercProd
End Property
