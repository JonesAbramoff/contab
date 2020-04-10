VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVendedorTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCargo As Integer
Private mvariSuperior As Integer
Private mvarcolVNDComissaoTRV As New Collection
Private mvarcolVNDReducaoTRV As New Collection
Private mvarcolVNDRegiaoTRV As New Collection
Private mvardPercCallCenter As Double

Public Property Let dPercCallCenter(ByVal vData As Double)
    mvardPercCallCenter = vData
End Property

Public Property Get dPercCallCenter() As Double
    dPercCallCenter = mvardPercCallCenter
End Property

Public Property Set colVNDRegiaoTRV(ByVal vData As Collection)
    Set mvarcolVNDRegiaoTRV = vData
End Property

Public Property Get colVNDRegiaoTRV() As Collection
    Set colVNDRegiaoTRV = mvarcolVNDRegiaoTRV
End Property

Public Property Set colVNDReducaoTRV(ByVal vData As Collection)
    Set mvarcolVNDReducaoTRV = vData
End Property

Public Property Get colVNDReducaoTRV() As Collection
    Set colVNDReducaoTRV = mvarcolVNDReducaoTRV
End Property

Public Property Set colVNDComissaoTRV(ByVal vData As Collection)
    Set mvarcolVNDComissaoTRV = vData
End Property

Public Property Get colVNDComissaoTRV() As Collection
    Set colVNDComissaoTRV = mvarcolVNDComissaoTRV
End Property

Public Property Let iSuperior(ByVal vData As Integer)
    mvariSuperior = vData
End Property

Public Property Get iSuperior() As Integer
    iSuperior = mvariSuperior
End Property

Public Property Let iCargo(ByVal vData As Integer)
    mvariCargo = vData
End Property

Public Property Get iCargo() As Integer
    iCargo = mvariCargo
End Property

