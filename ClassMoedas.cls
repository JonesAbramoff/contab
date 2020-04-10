VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMoedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodigo As Integer
Private mvarsNome As String
Private mvarsSimbolo As String
Private mvariPercentual As Integer
Private mvariCodBacen As Integer

Public Property Let iCodBacen(ByVal vData As Integer)
    mvariCodBacen = vData
End Property

Public Property Get iCodBacen() As Integer
    iCodBacen = mvariCodBacen
End Property

Public Property Let iPercentual(ByVal vData As Integer)
    mvariPercentual = vData
End Property

Public Property Get iPercentual() As Integer
    iPercentual = mvariPercentual
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sSimbolo(ByVal vData As String)
    mvarsSimbolo = vData
End Property

Public Property Get sSimbolo() As String
    sSimbolo = mvarsSimbolo
End Property
