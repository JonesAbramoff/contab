VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassVendedorTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCargo As Integer
Private mvariSuperior As Integer
Private mvarcolVNDComissaoTRP As New Collection
Private mvarcolVNDReducaoTRP As New Collection
Private mvarcolVNDRegiaoTRP As New Collection
Private mvarobjUsuarioWeb As New ClassTRPUsuarios

Public Property Set objUsuarioWeb(ByVal vData As ClassTRPUsuarios)
    Set mvarobjUsuarioWeb = vData
End Property

Public Property Get objUsuarioWeb() As ClassTRPUsuarios
    Set objUsuarioWeb = mvarobjUsuarioWeb
End Property

Public Property Set colVNDRegiaoTRP(ByVal vData As Collection)
    Set mvarcolVNDRegiaoTRP = vData
End Property

Public Property Get colVNDRegiaoTRP() As Collection
    Set colVNDRegiaoTRP = mvarcolVNDRegiaoTRP
End Property

Public Property Set colVNDReducaoTRP(ByVal vData As Collection)
    Set mvarcolVNDReducaoTRP = vData
End Property

Public Property Get colVNDReducaoTRP() As Collection
    Set colVNDReducaoTRP = mvarcolVNDReducaoTRP
End Property

Public Property Set colVNDComissaoTRP(ByVal vData As Collection)
    Set mvarcolVNDComissaoTRP = vData
End Property

Public Property Get colVNDComissaoTRP() As Collection
    Set colVNDComissaoTRP = mvarcolVNDComissaoTRP
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

