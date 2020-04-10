VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldMesEstAlm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAlmoxarifado As Integer
Private mvariAno As Integer
Private mvarsProduto As String
Private mvardQuantInicialConsig As Double
Private mvardValorInicialConsig As Double
Private mvardQuantInicialDemo As Double
Private mvardValorInicialDemo As Double
Private mvardQuantInicialConserto As Double
Private mvardValorInicialConserto As Double
Private mvardQuantInicialOutros As Double
Private mvardValorInicialOutros As Double
Private mvardQuantInicialBenef As Double
Private mvardValorInicialBenef As Double
Private mvardSaldoQuantConsig(1 To 12) As Double
Private mvardSaldoValorConsig(1 To 12) As Double
Private mvardSaldoQuantDemo(1 To 12) As Double
Private mvardSaldoValorDemo(1 To 12) As Double
Private mvardSaldoQuantConserto(1 To 12) As Double
Private mvardSaldoValorConserto(1 To 12) As Double
Private mvardSaldoQuantOutros(1 To 12) As Double
Private mvardSaldoValorOutros(1 To 12) As Double
Private mvardSaldoQuantBenef(1 To 12) As Double
Private mvardSaldoValorBenef(1 To 12) As Double

Public Property Let dSaldoValorBenef(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorBenef(iIndice) = vData
End Property

Public Property Get dSaldoValorBenef(ByVal iIndice As Integer) As Double
    dSaldoValorBenef = mvardSaldoValorBenef(iIndice)
End Property

Public Property Let dSaldoQuantBenef(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantBenef(iIndice) = vData
End Property

Public Property Get dSaldoQuantBenef(ByVal iIndice As Integer) As Double
    dSaldoQuantBenef = mvardSaldoQuantBenef(iIndice)
End Property

Public Property Let dSaldoValorOutros(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorOutros(iIndice) = vData
End Property

Public Property Get dSaldoValorOutros(ByVal iIndice As Integer) As Double
    dSaldoValorOutros = mvardSaldoValorOutros(iIndice)
End Property

Public Property Let dSaldoQuantOutros(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantOutros(iIndice) = vData
End Property

Public Property Get dSaldoQuantOutros(ByVal iIndice As Integer) As Double
    dSaldoQuantOutros = mvardSaldoQuantOutros(iIndice)
End Property

Public Property Let dSaldoValorConserto(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorConserto(iIndice) = vData
End Property

Public Property Get dSaldoValorConserto(ByVal iIndice As Integer) As Double
    dSaldoValorConserto = mvardSaldoValorConserto(iIndice)
End Property

Public Property Let dSaldoQuantConserto(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantConserto(iIndice) = vData
End Property

Public Property Get dSaldoQuantConserto(ByVal iIndice As Integer) As Double
    dSaldoQuantConserto = mvardSaldoQuantConserto(iIndice)
End Property

Public Property Let dSaldoValorDemo(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorDemo(iIndice) = vData
End Property

Public Property Get dSaldoValorDemo(ByVal iIndice As Integer) As Double
    dSaldoValorDemo = mvardSaldoValorDemo(iIndice)
End Property

Public Property Let dSaldoQuantDemo(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantDemo(iIndice) = vData
End Property

Public Property Get dSaldoQuantDemo(ByVal iIndice As Integer) As Double
    dSaldoQuantDemo = mvardSaldoQuantDemo(iIndice)
End Property

Public Property Let dSaldoValorConsig(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorConsig(iIndice) = vData
End Property

Public Property Get dSaldoValorConsig(ByVal iIndice As Integer) As Double
    dSaldoValorConsig = mvardSaldoValorConsig(iIndice)
End Property

Public Property Let dSaldoQuantConsig(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantConsig(iIndice) = vData
End Property

Public Property Get dSaldoQuantConsig(ByVal iIndice As Integer) As Double
    dSaldoQuantConsig = mvardSaldoQuantConsig(iIndice)
End Property

Public Property Let dValorInicialBenef(ByVal vData As Double)
    mvardValorInicialBenef = vData
End Property

Public Property Get dValorInicialBenef() As Double
    dValorInicialBenef = mvardValorInicialBenef
End Property

Public Property Let dQuantInicialBenef(ByVal vData As Double)
    mvardQuantInicialBenef = vData
End Property

Public Property Get dQuantInicialBenef() As Double
    dQuantInicialBenef = mvardQuantInicialBenef
End Property

Public Property Let dValorInicialOutros(ByVal vData As Double)
    mvardValorInicialOutros = vData
End Property

Public Property Get dValorInicialOutros() As Double
    dValorInicialOutros = mvardValorInicialOutros
End Property

Public Property Let dQuantInicialOutros(ByVal vData As Double)
    mvardQuantInicialOutros = vData
End Property

Public Property Get dQuantInicialOutros() As Double
    dQuantInicialOutros = mvardQuantInicialOutros
End Property

Public Property Let dValorInicialConserto(ByVal vData As Double)
    mvardValorInicialConserto = vData
End Property

Public Property Get dValorInicialConserto() As Double
    dValorInicialConserto = mvardValorInicialConserto
End Property

Public Property Let dQuantInicialConserto(ByVal vData As Double)
    mvardQuantInicialConserto = vData
End Property

Public Property Get dQuantInicialConserto() As Double
    dQuantInicialConserto = mvardQuantInicialConserto
End Property

Public Property Let dValorInicialDemo(ByVal vData As Double)
    mvardValorInicialDemo = vData
End Property

Public Property Get dValorInicialDemo() As Double
    dValorInicialDemo = mvardValorInicialDemo
End Property

Public Property Let dQuantInicialDemo(ByVal vData As Double)
    mvardQuantInicialDemo = vData
End Property

Public Property Get dQuantInicialDemo() As Double
    dQuantInicialDemo = mvardQuantInicialDemo
End Property

Public Property Let dValorInicialConsig(ByVal vData As Double)
    mvardValorInicialConsig = vData
End Property

Public Property Get dValorInicialConsig() As Double
    dValorInicialConsig = mvardValorInicialConsig
End Property

Public Property Let dQuantInicialConsig(ByVal vData As Double)
    mvardQuantInicialConsig = vData
End Property

Public Property Get dQuantInicialConsig() As Double
    dQuantInicialConsig = mvardQuantInicialConsig
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



