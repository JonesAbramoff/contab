VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCamposGenericosValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodCampo As Long
Private mvarlCodValor As Long
Private mvarsValor As String
Private mvariPadrao As Integer
Private mvarsComplemento1 As String
Private mvarsComplemento2 As String
Private mvarsComplemento3 As String
Private mvarsComplemento4 As String
Private mvarsComplemento5 As String

Public Property Let lCodCampo(ByVal vData As Long)
    mvarlCodCampo = vData
End Property

Public Property Get lCodCampo() As Long
    lCodCampo = mvarlCodCampo
End Property

Public Property Let lCodValor(ByVal vData As Long)
    mvarlCodValor = vData
End Property

Public Property Get lCodValor() As Long
    lCodValor = mvarlCodValor
End Property

Public Property Let sValor(ByVal vData As String)
    mvarsValor = vData
End Property

Public Property Get sValor() As String
    sValor = mvarsValor
End Property

Public Property Let iPadrao(ByVal vData As Integer)
    mvariPadrao = vData
End Property

Public Property Get iPadrao() As Integer
    iPadrao = mvariPadrao
End Property

Public Property Let sComplemento1(ByVal vData As String)
    mvarsComplemento1 = vData
End Property

Public Property Get sComplemento1() As String
    sComplemento1 = mvarsComplemento1
End Property

Public Property Let sComplemento2(ByVal vData As String)
    mvarsComplemento2 = vData
End Property

Public Property Get sComplemento2() As String
    sComplemento2 = mvarsComplemento2
End Property

Public Property Let sComplemento3(ByVal vData As String)
    mvarsComplemento3 = vData
End Property

Public Property Get sComplemento3() As String
    sComplemento3 = mvarsComplemento3
End Property

Public Property Let sComplemento4(ByVal vData As String)
    mvarsComplemento4 = vData
End Property

Public Property Get sComplemento4() As String
    sComplemento4 = mvarsComplemento4
End Property

Public Property Let sComplemento5(ByVal vData As String)
    mvarsComplemento5 = vData
End Property

Public Property Get sComplemento5() As String
    sComplemento5 = mvarsComplemento5
End Property


