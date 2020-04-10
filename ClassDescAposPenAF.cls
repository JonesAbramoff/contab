VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDescAposPenAF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlMatricula As Long
Private mvariDV As Integer
Private mvariMes As Integer
Private mvariAno As Integer
Private mvardValor As Double

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property

Public Property Let iMes(ByVal vData As Integer)
    mvariMes = vData
End Property

Public Property Get iMes() As Integer
    iMes = mvariMes
End Property

Public Property Let lMatricula(ByVal vData As Long)
    mvarlMatricula = vData
End Property

Public Property Get lMatricula() As Long
    lMatricula = mvarlMatricula
End Property

Public Property Let iDV(ByVal vData As Integer)
    mvariDV = vData
End Property

Public Property Get iDV() As Integer
    iDV = mvariDV
End Property
