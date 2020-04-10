VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldMesEstAlm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAlmoxarifado As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvardQuantInicialConsig3 As Double
Private mvardValorInicialConsig3 As Double
Private mvardQuantInicialDemo3 As Double
Private mvardValorInicialDemo3 As Double
Private mvardQuantInicialConserto3 As Double
Private mvardValorInicialConserto3 As Double
Private mvardQuantInicialOutros3 As Double
Private mvardValorInicialOutros3 As Double
Private mvardQuantInicialBenef3 As Double
Private mvardValorInicialBenef3 As Double
Private mvardSaldoQuantConsig3(1 To 12) As Double
Private mvardSaldoValorConsig3(1 To 12) As Double
Private mvardSaldoQuantDemo3(1 To 12) As Double
Private mvardSaldoValorDemo3(1 To 12) As Double
Private mvardSaldoQuantConserto3(1 To 12) As Double
Private mvardSaldoValorConserto3(1 To 12) As Double
Private mvardSaldoQuantOutros3(1 To 12) As Double
Private mvardSaldoValorOutros3(1 To 12) As Double
Private mvardSaldoQuantBenef3(1 To 12) As Double
Private mvardSaldoValorBenef3(1 To 12) As Double

Public Property Let dSaldoValorBenef3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorBenef3(iIndice) = vData
End Property

Public Property Get dSaldoValorBenef3(ByVal iIndice As Integer) As Double
    dSaldoValorBenef3 = mvardSaldoValorBenef3(iIndice)
End Property

Public Property Let dSaldoQuantBenef3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantBenef3(iIndice) = vData
End Property

Public Property Get dSaldoQuantBenef3(ByVal iIndice As Integer) As Double
    dSaldoQuantBenef3 = mvardSaldoQuantBenef3(iIndice)
End Property

Public Property Let dSaldoValorOutros3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorOutros3(iIndice) = vData
End Property

Public Property Get dSaldoValorOutros3(ByVal iIndice As Integer) As Double
    dSaldoValorOutros3 = mvardSaldoValorOutros3(iIndice)
End Property

Public Property Let dSaldoQuantOutros3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantOutros3(iIndice) = vData
End Property

Public Property Get dSaldoQuantOutros3(ByVal iIndice As Integer) As Double
    dSaldoQuantOutros3 = mvardSaldoQuantOutros3(iIndice)
End Property

Public Property Let dSaldoValorConserto3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorConserto3(iIndice) = vData
End Property

Public Property Get dSaldoValorConserto3(ByVal iIndice As Integer) As Double
    dSaldoValorConserto3 = mvardSaldoValorConserto3(iIndice)
End Property

Public Property Let dSaldoQuantConserto3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantConserto3(iIndice) = vData
End Property

Public Property Get dSaldoQuantConserto3(ByVal iIndice As Integer) As Double
    dSaldoQuantConserto3 = mvardSaldoQuantConserto3(iIndice)
End Property

Public Property Let dSaldoValorDemo3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorDemo3(iIndice) = vData
End Property

Public Property Get dSaldoValorDemo3(ByVal iIndice As Integer) As Double
    dSaldoValorDemo3 = mvardSaldoValorDemo3(iIndice)
End Property

Public Property Let dSaldoQuantDemo3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantDemo3(iIndice) = vData
End Property

Public Property Get dSaldoQuantDemo3(ByVal iIndice As Integer) As Double
    dSaldoQuantDemo3 = mvardSaldoQuantDemo3(iIndice)
End Property

Public Property Let dSaldoValorConsig3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorConsig3(iIndice) = vData
End Property

Public Property Get dSaldoValorConsig3(ByVal iIndice As Integer) As Double
    dSaldoValorConsig3 = mvardSaldoValorConsig3(iIndice)
End Property

Public Property Let dSaldoQuantConsig3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantConsig3(iIndice) = vData
End Property

Public Property Get dSaldoQuantConsig3(ByVal iIndice As Integer) As Double
    dSaldoQuantConsig3 = mvardSaldoQuantConsig3(iIndice)
End Property

Public Property Let dValorInicialBenef3(ByVal vData As Double)
    mvardValorInicialBenef3 = vData
End Property

Public Property Get dValorInicialBenef3() As Double
    dValorInicialBenef3 = mvardValorInicialBenef3
End Property

Public Property Let dQuantInicialBenef3(ByVal vData As Double)
    mvardQuantInicialBenef3 = vData
End Property

Public Property Get dQuantInicialBenef3() As Double
    dQuantInicialBenef3 = mvardQuantInicialBenef3
End Property

Public Property Let dValorInicialOutros3(ByVal vData As Double)
    mvardValorInicialOutros3 = vData
End Property

Public Property Get dValorInicialOutros3() As Double
    dValorInicialOutros3 = mvardValorInicialOutros3
End Property

Public Property Let dQuantInicialOutros3(ByVal vData As Double)
    mvardQuantInicialOutros3 = vData
End Property

Public Property Get dQuantInicialOutros3() As Double
    dQuantInicialOutros3 = mvardQuantInicialOutros3
End Property

Public Property Let dValorInicialConserto3(ByVal vData As Double)
    mvardValorInicialConserto3 = vData
End Property

Public Property Get dValorInicialConserto3() As Double
    dValorInicialConserto3 = mvardValorInicialConserto3
End Property

Public Property Let dQuantInicialConserto3(ByVal vData As Double)
    mvardQuantInicialConserto3 = vData
End Property

Public Property Get dQuantInicialConserto3() As Double
    dQuantInicialConserto3 = mvardQuantInicialConserto3
End Property

Public Property Let dValorInicialDemo3(ByVal vData As Double)
    mvardValorInicialDemo3 = vData
End Property

Public Property Get dValorInicialDemo3() As Double
    dValorInicialDemo3 = mvardValorInicialDemo3
End Property

Public Property Let dQuantInicialDemo3(ByVal vData As Double)
    mvardQuantInicialDemo3 = vData
End Property

Public Property Get dQuantInicialDemo3() As Double
    dQuantInicialDemo3 = mvardQuantInicialDemo3
End Property

Public Property Let dValorInicialConsig3(ByVal vData As Double)
    mvardValorInicialConsig3 = vData
End Property

Public Property Get dValorInicialConsig3() As Double
    dValorInicialConsig3 = mvardValorInicialConsig3
End Property

Public Property Let dQuantInicialConsig3(ByVal vData As Double)
    mvardQuantInicialConsig3 = vData
End Property

Public Property Get dQuantInicialConsig3() As Double
    dQuantInicialConsig3 = mvardQuantInicialConsig3
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let iAno(ByVal vData As Integer)
    mvariAno = vData
End Property

Public Property Get iAno() As Integer
    iAno = mvariAno
End Property





