VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAlmoxInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariAlmoxarifado As Integer
Private mvardQuantEntrada As Double
Private mvardQuantSaida As Double
Private mvardQuantCons As Double
Private mvardValorCons As Double
Private mvardValorEntrada As Double
Private mvardValorSaida As Double
Private mvardSaldoValorConsig As Double
Private mvardSaldoValorDemo As Double
Private mvardSaldoValorConserto As Double
Private mvardSaldoValorOutros As Double
Private mvardSaldoValorBenef As Double
Private mvardQuantVenda As Double
Private mvardValorVenda As Double
Private mvardSaldoValorCusto As Double
Private mvardSaldoValorConsig3 As Double
Private mvardSaldoValorDemo3 As Double
Private mvardSaldoValorConserto3 As Double
Private mvardSaldoValorOutros3 As Double
Private mvardSaldoValorBenef3 As Double

Public Property Let dSaldoValorBenef3(ByVal vData As Double)
    mvardSaldoValorBenef3 = vData
End Property

Public Property Get dSaldoValorBenef3() As Double
    dSaldoValorBenef3 = mvardSaldoValorBenef3
End Property

Public Property Let dSaldoValorOutros3(ByVal vData As Double)
    mvardSaldoValorOutros3 = vData
End Property

Public Property Get dSaldoValorOutros3() As Double
    dSaldoValorOutros3 = mvardSaldoValorOutros3
End Property

Public Property Let dSaldoValorConserto3(ByVal vData As Double)
    mvardSaldoValorConserto3 = vData
End Property

Public Property Get dSaldoValorConserto3() As Double
    dSaldoValorConserto3 = mvardSaldoValorConserto3
End Property

Public Property Let dSaldoValorDemo3(ByVal vData As Double)
    mvardSaldoValorDemo3 = vData
End Property

Public Property Get dSaldoValorDemo3() As Double
    dSaldoValorDemo3 = mvardSaldoValorDemo3
End Property

Public Property Let dSaldoValorConsig3(ByVal vData As Double)
    mvardSaldoValorConsig3 = vData
End Property

Public Property Get dSaldoValorConsig3() As Double
    dSaldoValorConsig3 = mvardSaldoValorConsig3
End Property

Public Property Let dSaldoValorCusto(ByVal vData As Double)
    mvardSaldoValorCusto = vData
End Property

Public Property Get dSaldoValorCusto() As Double
    dSaldoValorCusto = mvardSaldoValorCusto
End Property

Public Property Let dValorVenda(ByVal vData As Double)
    mvardValorVenda = vData
End Property

Public Property Get dValorVenda() As Double
    dValorVenda = mvardValorVenda
End Property

Public Property Let dQuantVenda(ByVal vData As Double)
    mvardQuantVenda = vData
End Property

Public Property Get dQuantVenda() As Double
    dQuantVenda = mvardQuantVenda
End Property

Public Property Let dSaldoValorBenef(ByVal vData As Double)
    mvardSaldoValorBenef = vData
End Property

Public Property Get dSaldoValorBenef() As Double
    dSaldoValorBenef = mvardSaldoValorBenef
End Property

Public Property Let dSaldoValorOutros(ByVal vData As Double)
    mvardSaldoValorOutros = vData
End Property

Public Property Get dSaldoValorOutros() As Double
    dSaldoValorOutros = mvardSaldoValorOutros
End Property

Public Property Let dSaldoValorConserto(ByVal vData As Double)
    mvardSaldoValorConserto = vData
End Property

Public Property Get dSaldoValorConserto() As Double
    dSaldoValorConserto = mvardSaldoValorConserto
End Property

Public Property Let dSaldoValorDemo(ByVal vData As Double)
    mvardSaldoValorDemo = vData
End Property

Public Property Get dSaldoValorDemo() As Double
    dSaldoValorDemo = mvardSaldoValorDemo
End Property

Public Property Let dSaldoValorConsig(ByVal vData As Double)
    mvardSaldoValorConsig = vData
End Property

Public Property Get dSaldoValorConsig() As Double
    dSaldoValorConsig = mvardSaldoValorConsig
End Property

Public Property Let dValorSaida(ByVal vData As Double)
    mvardValorSaida = vData
End Property

Public Property Get dValorSaida() As Double
    dValorSaida = mvardValorSaida
End Property

Public Property Let dValorEntrada(ByVal vData As Double)
    mvardValorEntrada = vData
End Property

Public Property Get dValorEntrada() As Double
    dValorEntrada = mvardValorEntrada
End Property

Public Property Let dValorCons(ByVal vData As Double)
    mvardValorCons = vData
End Property

Public Property Get dValorCons() As Double
    dValorCons = mvardValorCons
End Property

Public Property Let dQuantCons(ByVal vData As Double)
    mvardQuantCons = vData
End Property

Public Property Get dQuantCons() As Double
    dQuantCons = mvardQuantCons
End Property

Public Property Let dQuantSaida(ByVal vData As Double)
    mvardQuantSaida = vData
End Property

Public Property Get dQuantSaida() As Double
    dQuantSaida = mvardQuantSaida
End Property

Public Property Let dQuantEntrada(ByVal vData As Double)
    mvardQuantEntrada = vData
End Property

Public Property Get dQuantEntrada() As Double
    dQuantEntrada = mvardQuantEntrada
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property
