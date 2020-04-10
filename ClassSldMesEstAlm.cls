VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldMesEstAlm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAno As Integer
Private mvariAlmoxarifado As Integer
Private mvarsProduto As String
Private mvardQuantInicial As Double
Private mvardValorInicial As Double
Private mvardQuantInicialCusto As Double
Private mvardValorInicialCusto As Double
Private mvardQuantEnt(1 To 12) As Double
Private mvardValorEnt(1 To 12) As Double
Private mvardQuantSai(1 To 12) As Double
Private mvardQuantCons(1 To 12) As Double
Private mvardQuantVend(1 To 12) As Double
Private mvardQuantVendConsig3(1 To 12) As Double
Private mvardValorSai(1 To 12) As Double
Private mvariStatus(1 To 12) As Integer
Private mvardQuantComp(1 To 12) As Double
Private mvardValorComp(1 To 12) As Double
Private mvardValorVend(1 To 12) As Double
Private mvardValorVendConsig3(1 To 12) As Double
Private mvardValorCons(1 To 12) As Double
Private mvardSaldoQuantCusto(1 To 12) As Double
Private mvardSaldoValorCusto(1 To 12) As Double
Private mvardSaldoQuantRecIndl(1 To 12) As Double
Private mvardSaldoValorRecIndl(1 To 12) As Double

Public Property Let dSaldoValorRecIndl(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorRecIndl(iIndice) = vData
End Property

Public Property Get dSaldoValorRecIndl(ByVal iIndice As Integer) As Double
    dSaldoValorRecIndl = mvardSaldoValorRecIndl(iIndice)
End Property

Public Property Let dSaldoQuantRecIndl(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantRecIndl(iIndice) = vData
End Property

Public Property Get dSaldoQuantRecIndl(ByVal iIndice As Integer) As Double
    dSaldoQuantRecIndl = mvardSaldoQuantRecIndl(iIndice)
End Property

Public Property Let dSaldoValorCusto(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoValorCusto(iIndice) = vData
End Property

Public Property Get dSaldoValorCusto(ByVal iIndice As Integer) As Double
    dSaldoValorCusto = mvardSaldoValorCusto(iIndice)
End Property

Public Property Let dSaldoQuantCusto(ByVal iIndice As Integer, ByVal vData As Double)
    mvardSaldoQuantCusto(iIndice) = vData
End Property

Public Property Get dSaldoQuantCusto(ByVal iIndice As Integer) As Double
    dSaldoQuantCusto = mvardSaldoQuantCusto(iIndice)
End Property

Public Property Let dValorCons(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorCons(iIndice) = vData
End Property

Public Property Get dValorCons(ByVal iIndice As Integer) As Double
    dValorCons = mvardValorCons(iIndice)
End Property

Public Property Let dValorVendConsig3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorVendConsig3(iIndice) = vData
End Property

Public Property Get dValorVendConsig3(ByVal iIndice As Integer) As Double
    dValorVendConsig3 = mvardValorVendConsig3(iIndice)
End Property

Public Property Let dValorVend(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorVend(iIndice) = vData
End Property

Public Property Get dValorVend(ByVal iIndice As Integer) As Double
    dValorVend = mvardValorVend(iIndice)
End Property

Public Property Let dValorComp(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorComp(iIndice) = vData
End Property

Public Property Get dValorComp(ByVal iIndice As Integer) As Double
    dValorComp = mvardValorComp(iIndice)
End Property

Public Property Let dQuantComp(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantComp(iIndice) = vData
End Property

Public Property Get dQuantComp(ByVal iIndice As Integer) As Double
    dQuantComp = mvardQuantComp(iIndice)
End Property

Public Property Let iStatus(ByVal iIndice As Integer, ByVal vData As Integer)
    mvariStatus(iIndice) = vData
End Property

Public Property Get iStatus(ByVal iIndice As Integer) As Integer
    iStatus = mvariStatus(iIndice)
End Property

Public Property Let dValorSai(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorSai(iIndice) = vData
End Property

Public Property Get dValorSai(ByVal iIndice As Integer) As Double
    dValorSai = mvardValorSai(iIndice)
End Property

Public Property Let dQuantVendConsig3(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantVendConsig3(iIndice) = vData
End Property

Public Property Get dQuantVendConsig3(ByVal iIndice As Integer) As Double
    dQuantVendConsig3 = mvardQuantVendConsig3(iIndice)
End Property

Public Property Let dQuantVend(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantVend(iIndice) = vData
End Property

Public Property Get dQuantVend(ByVal iIndice As Integer) As Double
    dQuantVend = mvardQuantVend(iIndice)
End Property

Public Property Let dQuantCons(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantCons(iIndice) = vData
End Property

Public Property Get dQuantCons(ByVal iIndice As Integer) As Double
    dQuantCons = mvardQuantCons(iIndice)
End Property

Public Property Let dQuantSai(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantSai(iIndice) = vData
End Property

Public Property Get dQuantSai(ByVal iIndice As Integer) As Double
    dQuantSai = mvardQuantSai(iIndice)
End Property

Public Property Let dValorEnt(ByVal iIndice As Integer, ByVal vData As Double)
    mvardValorEnt(iIndice) = vData
End Property

Public Property Get dValorEnt(ByVal iIndice As Integer) As Double
    dValorEnt = mvardValorEnt(iIndice)
End Property

Public Property Let dQuantEnt(ByVal iIndice As Integer, ByVal vData As Double)
    mvardQuantEnt(iIndice) = vData
End Property

Public Property Get dQuantEnt(ByVal iIndice As Integer) As Double
    dQuantEnt = mvardQuantEnt(iIndice)
End Property

Public Property Let dValorInicialCusto(ByVal vData As Double)
    mvardValorInicialCusto = vData
End Property

Public Property Get dValorInicialCusto() As Double
    dValorInicialCusto = mvardValorInicialCusto
End Property

Public Property Let dQuantInicialCusto(ByVal vData As Double)
    mvardQuantInicialCusto = vData
End Property

Public Property Get dQuantInicialCusto() As Double
    dQuantInicialCusto = mvardQuantInicialCusto
End Property

Public Property Let dValorInicial(ByVal vData As Double)
    mvardValorInicial = vData
End Property

Public Property Get dValorInicial() As Double
    dValorInicial = mvardValorInicial
End Property

Public Property Let dQuantInicial(ByVal vData As Double)
    mvardQuantInicial = vData
End Property

Public Property Get dQuantInicial() As Double
    dQuantInicial = mvardQuantInicial
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

