VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaArqFisMestre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodECF As Integer
Private mvariFilialEmpresa As Integer
Private mvardtData As Date
Private mvarsNumSerieECF As String
Private mvarlCOOIni As Long
Private mvarlCOOFim As Long
Private mvarlCRZ As Long
Private mvardGrandeTotal As Double
Private mvarlCRO As Long
Private mvardVendaBruta As Double
Private mvariNumEquip As Integer

Public Property Let iNumEquip(ByVal vData As Integer)
    mvariNumEquip = vData
End Property

Public Property Get iNumEquip() As Integer
    iNumEquip = mvariNumEquip
End Property

Public Property Let dVendaBruta(ByVal vData As Double)
    mvardVendaBruta = vData
End Property

Public Property Get dVendaBruta() As Double
    dVendaBruta = mvardVendaBruta
End Property

Public Property Let lCRO(ByVal vData As Long)
    mvarlCRO = vData
End Property

Public Property Get lCRO() As Long
    lCRO = mvarlCRO
End Property

Public Property Let lCRZ(ByVal vData As Long)
    mvarlCRZ = vData
End Property

Public Property Get lCRZ() As Long
    lCRZ = mvarlCRZ
End Property

Public Property Let dGrandeTotal(ByVal vData As Double)
    mvardGrandeTotal = vData
End Property

Public Property Get dGrandeTotal() As Double
    dGrandeTotal = mvardGrandeTotal
End Property

Public Property Let lCOOFim(ByVal vData As Long)
    mvarlCOOFim = vData
End Property

Public Property Get lCOOFim() As Long
    lCOOFim = mvarlCOOFim
End Property

Public Property Let lCOOIni(ByVal vData As Long)
    mvarlCOOIni = vData
End Property

Public Property Get lCOOIni() As Long
    lCOOIni = mvarlCOOIni
End Property

Public Property Let sNumSerieECF(ByVal vData As String)
    mvarsNumSerieECF = vData
End Property

Public Property Get sNumSerieECF() As String
    sNumSerieECF = mvarsNumSerieECF
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodECF(ByVal vData As Integer)
    mvariCodECF = vData
End Property

Public Property Get iCodECF() As Integer
    iCodECF = mvariCodECF
End Property



