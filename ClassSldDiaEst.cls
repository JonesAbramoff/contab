VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSldDiaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardtData As Date
Private mvardQuantEntrada As Double
Private mvardQuantSaida As Double
Private mvardQuantCons As Double
Private mvardQuantVend As Double
Private mvardQuantVendConsig3 As Double
Private mvardValorEntrada As Double
Private mvardValorSaida As Double
Private mvardQuantComp As Double
Private mvardValorComp As Double
Private mvardValorVend As Double
Private mvardValorVendConsig3 As Double
Private mvardValorCons As Double
Private mvardQuantEntCusto As Double
Private mvardValorEntCusto As Double
Private mvardQuantSaiCusto As Double
Private mvardValorSaiCusto As Double
Private mvardQuantEntConsig3 As Double
Private mvardValorEntConsig3 As Double
Private mvardQuantSaiConsig3 As Double
Private mvardValorSaiConsig3 As Double
Private mvardQuantEntDemo3 As Double
Private mvardValorEntDemo3 As Double
Private mvardQuantSaiDemo3 As Double
Private mvardValorSaiDemo3 As Double
Private mvardQuantEntConserto3 As Double
Private mvardValorEntConserto3 As Double
Private mvardQuantSaiConserto3 As Double
Private mvardValorSaiConserto3 As Double
Private mvardQuantEntOutros3 As Double
Private mvardValorEntOutros3 As Double
Private mvardQuantSaiOutros3 As Double
Private mvardValorSaiOutros3 As Double
Private mvardQuantEntBenef3 As Double
Private mvardValorEntBenef3 As Double
Private mvardQuantSaiBenef3 As Double
Private mvardValorSaiBenef3 As Double
Private mvardQuantEntConsig As Double
Private mvardValorEntConsig As Double
Private mvardQuantSaiConsig As Double
Private mvardValorSaiConsig As Double
Private mvardQuantEntDemo As Double
Private mvardValorEntDemo As Double
Private mvardQuantSaiDemo As Double
Private mvardValorSaiDemo As Double
Private mvardQuantEntConserto As Double
Private mvardValorEntConserto As Double
Private mvardQuantSaiConserto As Double
Private mvardValorSaiConserto As Double
Private mvardQuantEntOutros As Double
Private mvardValorEntOutros As Double
Private mvardQuantSaiOutros As Double
Private mvardValorSaiOutros As Double
Private mvardQuantEntBenef As Double
Private mvardValorEntBenef As Double
Private mvardQuantSaiBenef As Double
Private mvardValorSaiBenef As Double
Private mvardQuantEntRecIndl As Double
Private mvardValorEntRecIndl As Double
Private mvardQuantSaiRecIndl As Double
Private mvardValorSaiRecIndl As Double
'Private mvardQuantEntDispNossa As Double
'Private mvardQuantEntIndOutras As Double
'Private mvardQuantEntDefeituosa As Double
'Private mvardQuantEntOP As Double
'Private mvardQuantEntReservada As Double
'Private mvardQuantEntReservadaConsig As Double
'Private mvardQuantSaiDispNossa As Double
'Private mvardQuantSaiIndOutras As Double
'Private mvardQuantSaiDefeituosa As Double
'Private mvardQuantSaiOP As Double
'Private mvardQuantSaiReservada As Double
'Private mvardQuantSaiReservadaConsig As Double
'
'Public Property Let dQuantSaiReservadaConsig(ByVal vData As Double)
'    mvardQuantSaiReservadaConsig = vData
'End Property
'
'Public Property Get dQuantSaiReservadaConsig() As Double
'    dQuantSaiReservadaConsig = mvardQuantSaiReservadaConsig
'End Property
'
'Public Property Let dQuantSaiReservada(ByVal vData As Double)
'    mvardQuantSaiReservada = vData
'End Property
'
'Public Property Get dQuantSaiReservada() As Double
'    dQuantSaiReservada = mvardQuantSaiReservada
'End Property
'
'Public Property Let dQuantSaiOP(ByVal vData As Double)
'    mvardQuantSaiOP = vData
'End Property
'
'Public Property Get dQuantSaiOP() As Double
'    dQuantSaiOP = mvardQuantSaiOP
'End Property
'
'Public Property Let dQuantSaiOP(ByVal vData As Double)
'    mvardQuantSaiOP = vData
'End Property
'
'Public Property Get dQuantSaiDefeituosa() As Double
'    dQuantSaiDefeituosa = mvardQuantSaiDefeituosa
'End Property
'
'Public Property Let dQuantSaiIndOutras(ByVal vData As Double)
'    mvardQuantSaiIndOutras = vData
'End Property
'
'Public Property Get dQuantSaiIndOutras() As Double
'    dQuantSaiIndOutras = mvardQuantSaiIndOutras
'End Property
'
'Public Property Let dQuantSaiDispNossa(ByVal vData As Double)
'    mvardQuantSaiDispNossa = vData
'End Property
'
'Public Property Get dQuantSaiDispNossa() As Double
'    dQuantSaiDispNossa = mvardQuantSaiDispNossa
'End Property
'
'Public Property Let dQuantEntReservadaConsig(ByVal vData As Double)
'    mvardQuantEntReservadaConsig = vData
'End Property
'
'Public Property Get dQuantEntReservadaConsig() As Double
'    dQuantEntReservadaConsig = mvardQuantEntReservadaConsig
'End Property
'
'Public Property Let dQuantEntReservada(ByVal vData As Double)
'    mvardQuantEntReservada = vData
'End Property
'
'Public Property Get dQuantEntReservada() As Double
'    dQuantEntReservada = mvardQuantEntReservada
'End Property
'
'Public Property Let dQuantEntOP(ByVal vData As Double)
'    mvardQuantEntOP = vData
'End Property
'
'Public Property Get dQuantEntOP() As Double
'    dQuantEntOP = mvardQuantEntOP
'End Property
'
'Public Property Let dQuantEntOP(ByVal vData As Double)
'    mvardQuantEntOP = vData
'End Property
'
'Public Property Get dQuantEntDefeituosa() As Double
'    dQuantEntDefeituosa = mvardQuantEntDefeituosa
'End Property
'
'Public Property Let dQuantEntIndOutras(ByVal vData As Double)
'    mvardQuantEntIndOutras = vData
'End Property
'
'Public Property Get dQuantEntIndOutras() As Double
'    dQuantEntIndOutras = mvardQuantEntIndOutras
'End Property
'
'Public Property Let dQuantEntDispNossa(ByVal vData As Double)
'    mvardQuantEntDispNossa = vData
'End Property
'
'Public Property Get dQuantEntDispNossa() As Double
'    dQuantEntDispNossa = mvardQuantEntDispNossa
'End Property

Public Property Let dValorSaiRecIndl(ByVal vData As Double)
    mvardValorSaiRecIndl = vData
End Property

Public Property Get dValorSaiRecIndl() As Double
    dValorSaiRecIndl = mvardValorSaiRecIndl
End Property

Public Property Let dQuantSaiRecIndl(ByVal vData As Double)
    mvardQuantSaiRecIndl = vData
End Property

Public Property Get dQuantSaiRecIndl() As Double
    dQuantSaiRecIndl = mvardQuantSaiRecIndl
End Property

Public Property Let dValorEntRecIndl(ByVal vData As Double)
    mvardValorEntRecIndl = vData
End Property

Public Property Get dValorEntRecIndl() As Double
    dValorEntRecIndl = mvardValorEntRecIndl
End Property

Public Property Let dQuantEntRecIndl(ByVal vData As Double)
    mvardQuantEntRecIndl = vData
End Property

Public Property Get dQuantEntRecIndl() As Double
    dQuantEntRecIndl = mvardQuantEntRecIndl
End Property

Public Property Let dValorSaiBenef(ByVal vData As Double)
    mvardValorSaiBenef = vData
End Property

Public Property Get dValorSaiBenef() As Double
    dValorSaiBenef = mvardValorSaiBenef
End Property

Public Property Let dQuantSaiBenef(ByVal vData As Double)
    mvardQuantSaiBenef = vData
End Property

Public Property Get dQuantSaiBenef() As Double
    dQuantSaiBenef = mvardQuantSaiBenef
End Property

Public Property Let dValorEntBenef(ByVal vData As Double)
    mvardValorEntBenef = vData
End Property

Public Property Get dValorEntBenef() As Double
    dValorEntBenef = mvardValorEntBenef
End Property

Public Property Let dQuantEntBenef(ByVal vData As Double)
    mvardQuantEntBenef = vData
End Property

Public Property Get dQuantEntBenef() As Double
    dQuantEntBenef = mvardQuantEntBenef
End Property

Public Property Let dValorSaiOutros(ByVal vData As Double)
    mvardValorSaiOutros = vData
End Property

Public Property Get dValorSaiOutros() As Double
    dValorSaiOutros = mvardValorSaiOutros
End Property

Public Property Let dQuantSaiOutros(ByVal vData As Double)
    mvardQuantSaiOutros = vData
End Property

Public Property Get dQuantSaiOutros() As Double
    dQuantSaiOutros = mvardQuantSaiOutros
End Property

Public Property Let dValorEntOutros(ByVal vData As Double)
    mvardValorEntOutros = vData
End Property

Public Property Get dValorEntOutros() As Double
    dValorEntOutros = mvardValorEntOutros
End Property

Public Property Let dQuantEntOutros(ByVal vData As Double)
    mvardQuantEntOutros = vData
End Property

Public Property Get dQuantEntOutros() As Double
    dQuantEntOutros = mvardQuantEntOutros
End Property

Public Property Let dValorSaiConserto(ByVal vData As Double)
    mvardValorSaiConserto = vData
End Property

Public Property Get dValorSaiConserto() As Double
    dValorSaiConserto = mvardValorSaiConserto
End Property

Public Property Let dQuantSaiConserto(ByVal vData As Double)
    mvardQuantSaiConserto = vData
End Property

Public Property Get dQuantSaiConserto() As Double
    dQuantSaiConserto = mvardQuantSaiConserto
End Property

Public Property Let dValorEntConserto(ByVal vData As Double)
    mvardValorEntConserto = vData
End Property

Public Property Get dValorEntConserto() As Double
    dValorEntConserto = mvardValorEntConserto
End Property

Public Property Let dQuantEntConserto(ByVal vData As Double)
    mvardQuantEntConserto = vData
End Property

Public Property Get dQuantEntConserto() As Double
    dQuantEntConserto = mvardQuantEntConserto
End Property

Public Property Let dValorSaiDemo(ByVal vData As Double)
    mvardValorSaiDemo = vData
End Property

Public Property Get dValorSaiDemo() As Double
    dValorSaiDemo = mvardValorSaiDemo
End Property

Public Property Let dQuantSaiDemo(ByVal vData As Double)
    mvardQuantSaiDemo = vData
End Property

Public Property Get dQuantSaiDemo() As Double
    dQuantSaiDemo = mvardQuantSaiDemo
End Property

Public Property Let dValorEntDemo(ByVal vData As Double)
    mvardValorEntDemo = vData
End Property

Public Property Get dValorEntDemo() As Double
    dValorEntDemo = mvardValorEntDemo
End Property

Public Property Let dQuantEntDemo(ByVal vData As Double)
    mvardQuantEntDemo = vData
End Property

Public Property Get dQuantEntDemo() As Double
    dQuantEntDemo = mvardQuantEntDemo
End Property

Public Property Let dValorSaiConsig(ByVal vData As Double)
    mvardValorSaiConsig = vData
End Property

Public Property Get dValorSaiConsig() As Double
    dValorSaiConsig = mvardValorSaiConsig
End Property

Public Property Let dQuantSaiConsig(ByVal vData As Double)
    mvardQuantSaiConsig = vData
End Property

Public Property Get dQuantSaiConsig() As Double
    dQuantSaiConsig = mvardQuantSaiConsig
End Property

Public Property Let dValorEntConsig(ByVal vData As Double)
    mvardValorEntConsig = vData
End Property

Public Property Get dValorEntConsig() As Double
    dValorEntConsig = mvardValorEntConsig
End Property

Public Property Let dQuantEntConsig(ByVal vData As Double)
    mvardQuantEntConsig = vData
End Property

Public Property Get dQuantEntConsig() As Double
    dQuantEntConsig = mvardQuantEntConsig
End Property

Public Property Let dValorSaiBenef3(ByVal vData As Double)
    mvardValorSaiBenef3 = vData
End Property

Public Property Get dValorSaiBenef3() As Double
    dValorSaiBenef3 = mvardValorSaiBenef3
End Property

Public Property Let dQuantSaiBenef3(ByVal vData As Double)
    mvardQuantSaiBenef3 = vData
End Property

Public Property Get dQuantSaiBenef3() As Double
    dQuantSaiBenef3 = mvardQuantSaiBenef3
End Property

Public Property Let dValorEntBenef3(ByVal vData As Double)
    mvardValorEntBenef3 = vData
End Property

Public Property Get dValorEntBenef3() As Double
    dValorEntBenef3 = mvardValorEntBenef3
End Property

Public Property Let dQuantEntBenef3(ByVal vData As Double)
    mvardQuantEntBenef3 = vData
End Property

Public Property Get dQuantEntBenef3() As Double
    dQuantEntBenef3 = mvardQuantEntBenef3
End Property

Public Property Let dValorSaiOutros3(ByVal vData As Double)
    mvardValorSaiOutros3 = vData
End Property

Public Property Get dValorSaiOutros3() As Double
    dValorSaiOutros3 = mvardValorSaiOutros3
End Property

Public Property Let dQuantSaiOutros3(ByVal vData As Double)
    mvardQuantSaiOutros3 = vData
End Property

Public Property Get dQuantSaiOutros3() As Double
    dQuantSaiOutros3 = mvardQuantSaiOutros3
End Property

Public Property Let dValorEntOutros3(ByVal vData As Double)
    mvardValorEntOutros3 = vData
End Property

Public Property Get dValorEntOutros3() As Double
    dValorEntOutros3 = mvardValorEntOutros3
End Property

Public Property Let dQuantEntOutros3(ByVal vData As Double)
    mvardQuantEntOutros3 = vData
End Property

Public Property Get dQuantEntOutros3() As Double
    dQuantEntOutros3 = mvardQuantEntOutros3
End Property

Public Property Let dValorSaiConserto3(ByVal vData As Double)
    mvardValorSaiConserto3 = vData
End Property

Public Property Get dValorSaiConserto3() As Double
    dValorSaiConserto3 = mvardValorSaiConserto3
End Property

Public Property Let dQuantSaiConserto3(ByVal vData As Double)
    mvardQuantSaiConserto3 = vData
End Property

Public Property Get dQuantSaiConserto3() As Double
    dQuantSaiConserto3 = mvardQuantSaiConserto3
End Property

Public Property Let dValorEntConserto3(ByVal vData As Double)
    mvardValorEntConserto3 = vData
End Property

Public Property Get dValorEntConserto3() As Double
    dValorEntConserto3 = mvardValorEntConserto3
End Property

Public Property Let dQuantEntConserto3(ByVal vData As Double)
    mvardQuantEntConserto3 = vData
End Property

Public Property Get dQuantEntConserto3() As Double
    dQuantEntConserto3 = mvardQuantEntConserto3
End Property

Public Property Let dValorSaiDemo3(ByVal vData As Double)
    mvardValorSaiDemo3 = vData
End Property

Public Property Get dValorSaiDemo3() As Double
    dValorSaiDemo3 = mvardValorSaiDemo3
End Property

Public Property Let dQuantSaiDemo3(ByVal vData As Double)
    mvardQuantSaiDemo3 = vData
End Property

Public Property Get dQuantSaiDemo3() As Double
    dQuantSaiDemo3 = mvardQuantSaiDemo3
End Property

Public Property Let dValorEntDemo3(ByVal vData As Double)
    mvardValorEntDemo3 = vData
End Property

Public Property Get dValorEntDemo3() As Double
    dValorEntDemo3 = mvardValorEntDemo3
End Property

Public Property Let dQuantEntDemo3(ByVal vData As Double)
    mvardQuantEntDemo3 = vData
End Property

Public Property Get dQuantEntDemo3() As Double
    dQuantEntDemo3 = mvardQuantEntDemo3
End Property

Public Property Let dValorSaiConsig3(ByVal vData As Double)
    mvardValorSaiConsig3 = vData
End Property

Public Property Get dValorSaiConsig3() As Double
    dValorSaiConsig3 = mvardValorSaiConsig3
End Property

Public Property Let dQuantSaiConsig3(ByVal vData As Double)
    mvardQuantSaiConsig3 = vData
End Property

Public Property Get dQuantSaiConsig3() As Double
    dQuantSaiConsig3 = mvardQuantSaiConsig3
End Property

Public Property Let dValorEntConsig3(ByVal vData As Double)
    mvardValorEntConsig3 = vData
End Property

Public Property Get dValorEntConsig3() As Double
    dValorEntConsig3 = mvardValorEntConsig3
End Property

Public Property Let dQuantEntConsig3(ByVal vData As Double)
    mvardQuantEntConsig3 = vData
End Property

Public Property Get dQuantEntConsig3() As Double
    dQuantEntConsig3 = mvardQuantEntConsig3
End Property

Public Property Let dValorSaiCusto(ByVal vData As Double)
    mvardValorSaiCusto = vData
End Property

Public Property Get dValorSaiCusto() As Double
    dValorSaiCusto = mvardValorSaiCusto
End Property

Public Property Let dQuantSaiCusto(ByVal vData As Double)
    mvardQuantSaiCusto = vData
End Property

Public Property Get dQuantSaiCusto() As Double
    dQuantSaiCusto = mvardQuantSaiCusto
End Property

Public Property Let dValorEntCusto(ByVal vData As Double)
    mvardValorEntCusto = vData
End Property

Public Property Get dValorEntCusto() As Double
    dValorEntCusto = mvardValorEntCusto
End Property

Public Property Let dQuantEntCusto(ByVal vData As Double)
    mvardQuantEntCusto = vData
End Property

Public Property Get dQuantEntCusto() As Double
    dQuantEntCusto = mvardQuantEntCusto
End Property

Public Property Let dValorCons(ByVal vData As Double)
    mvardValorCons = vData
End Property

Public Property Get dValorCons() As Double
    dValorCons = mvardValorCons
End Property

Public Property Let dValorVendConsig3(ByVal vData As Double)
    mvardValorVendConsig3 = vData
End Property

Public Property Get dValorVendConsig3() As Double
    dValorVendConsig3 = mvardValorVendConsig3
End Property

Public Property Let dValorVend(ByVal vData As Double)
    mvardValorVend = vData
End Property

Public Property Get dValorVend() As Double
    dValorVend = mvardValorVend
End Property

Public Property Let dValorComp(ByVal vData As Double)
    mvardValorComp = vData
End Property

Public Property Get dValorComp() As Double
    dValorComp = mvardValorComp
End Property

Public Property Let dQuantComp(ByVal vData As Double)
    mvardQuantComp = vData
End Property

Public Property Get dQuantComp() As Double
    dQuantComp = mvardQuantComp
End Property

Public Property Let dValorSaida(ByVal vData As Double)
Attribute dValorSaida.VB_Description = "Valor que saiu do estoque"
    mvardValorSaida = vData
End Property

Public Property Get dValorSaida() As Double
    dValorSaida = mvardValorSaida
End Property

Public Property Let dValorEntrada(ByVal vData As Double)
Attribute dValorEntrada.VB_Description = "Valor que entrou no estoque"
    mvardValorEntrada = vData
End Property

Public Property Get dValorEntrada() As Double
    dValorEntrada = mvardValorEntrada
End Property

Public Property Let dQuantVendConsig3(ByVal vData As Double)
    mvardQuantVendConsig3 = vData
End Property

Public Property Get dQuantVendConsig3() As Double
    dQuantVendConsig3 = mvardQuantVendConsig3
End Property

Public Property Let dQuantVend(ByVal vData As Double)
    mvardQuantVend = vData
End Property

Public Property Get dQuantVend() As Double
    dQuantVend = mvardQuantVend
End Property

Public Property Let dQuantCons(ByVal vData As Double)
Attribute dQuantCons.VB_Description = "Quantidade consumida/vendida"
    mvardQuantCons = vData
End Property

Public Property Get dQuantCons() As Double
    dQuantCons = mvardQuantCons
End Property

Public Property Let dQuantSaida(ByVal vData As Double)
Attribute dQuantSaida.VB_Description = "Quantidade que saiu do estoque"
    mvardQuantSaida = vData
End Property

Public Property Get dQuantSaida() As Double
    dQuantSaida = mvardQuantSaida
End Property

Public Property Let dQuantEntrada(ByVal vData As Double)
Attribute dQuantEntrada.VB_Description = "Quantidade que entrou no estoque"
    mvardQuantEntrada = vData
End Property

Public Property Get dQuantEntrada() As Double
    dQuantEntrada = mvardQuantEntrada
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data do movimento no estoque"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sProduto(ByVal vData As String)
Attribute sProduto.VB_Description = "Codigo do Produto"
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial da Empresa"
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



