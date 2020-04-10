VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEstoqueProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'Existe QuantDisponivel que é calculada
Private mvarsProduto As String
Private mvariAlmoxarifado As Integer
Private mvarsAlmoxarifadoNomeReduzido As String
Private mvardQuantDispNossa As Double
Private mvardSaldo As Double
Private mvarsLocalizacaoFisica As String
Private mvardQuantReservada As Double
Private mvardQuantReservadaConsig As Double
Private mvardtDataInventario As Date
Private mvardQuantidadeInicial As Double
Private mvardSaldoInicial As Double
Private mvardtDataInicial As Date
Private mvardQuantEmpenhada As Double
Private mvardQuantRecIndl As Double
Private mvardQuantInd As Double
Private mvardQuantDefeituosa As Double
Private mvardQuantConsig As Double
Private mvardQuantConsig3 As Double
Private mvardQuantDemo As Double
Private mvardQuantDemo3 As Double
Private mvardQuantConserto3 As Double
Private mvardQuantConserto As Double
Private mvardQuantOutras As Double
Private mvardQuantOutras3 As Double
Private mvarsContaContabil As String
Private mvardQuantOP As Double
Private mvardQuantBenef As Double
Private mvardQuantBenef3 As Double
Private mvardQuantPedido As Double
Private mvardValorConsig As Double
Private mvardValorConsig3 As Double
Private mvardValorDemo As Double
Private mvardValorDemo3 As Double
Private mvardValorConserto3 As Double
Private mvardValorConserto As Double
Private mvardValorOutras As Double
Private mvardValorOutras3 As Double
Private mvardValorBenef As Double
Private mvardValorBenef3 As Double

'Raphael 21/07/00
Private mvardQuantInicialConsig3 As Double
Private mvardQuantInicialConsig As Double
Private mvardQuantInicialDemo3 As Double
Private mvardQuantInicialDemo As Double
Private mvardQuantInicialConserto3 As Double
Private mvardQuantInicialConserto As Double
Private mvardQuantInicialOutras3 As Double
Private mvardQuantInicialOutras As Double
Private mvardQuantInicialBenef As Double
Private mvardQuantInicialBenef3 As Double
Private mvardValorInicialConsig3 As Double
Private mvardValorInicialConsig As Double
Private mvardValorInicialDemo3 As Double
Private mvardValorInicialDemo As Double
Private mvardValorInicialConserto3 As Double
Private mvardValorInicialConserto As Double
Private mvardValorInicialOutras3 As Double
Private mvardValorInicialOutras As Double
Private mvardValorInicialBenef As Double
Private mvardValorInicialBenef3 As Double
Private mvariFilialEmpresa As Integer

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dValorInicialBenef3(ByVal vData As Double)
    mvardValorInicialBenef3 = vData
End Property

Public Property Get dValorInicialBenef3() As Double
    dValorInicialBenef3 = mvardValorInicialBenef3
End Property

Public Property Let dValorInicialBenef(ByVal vData As Double)
    mvardValorInicialBenef = vData
End Property

Public Property Get dValorInicialBenef() As Double
    dValorInicialBenef = mvardValorInicialBenef
End Property

Public Property Let dValorInicialOutras(ByVal vData As Double)
    mvardValorInicialOutras = vData
End Property

Public Property Get dValorInicialOutras() As Double
    dValorInicialOutras = mvardValorInicialOutras
End Property

Public Property Let dValorInicialOutras3(ByVal vData As Double)
    mvardValorInicialOutras3 = vData
End Property

Public Property Get dValorInicialOutras3() As Double
    dValorInicialOutras3 = mvardValorInicialOutras3
End Property

Public Property Let dValorInicialConserto(ByVal vData As Double)
    mvardValorInicialConserto = vData
End Property

Public Property Get dValorInicialConserto() As Double
    dValorInicialConserto = mvardValorInicialConserto
End Property

Public Property Let dValorInicialConserto3(ByVal vData As Double)
    mvardValorInicialConserto3 = vData
End Property

Public Property Get dValorInicialConserto3() As Double
    dValorInicialConserto3 = mvardValorInicialConserto3
End Property

Public Property Let dValorInicialDemo(ByVal vData As Double)
    mvardValorInicialDemo = vData
End Property

Public Property Get dValorInicialDemo() As Double
    dValorInicialDemo = mvardValorInicialDemo
End Property

Public Property Let dValorInicialDemo3(ByVal vData As Double)
    mvardValorInicialDemo3 = vData
End Property

Public Property Get dValorInicialDemo3() As Double
    dValorInicialDemo3 = mvardValorInicialDemo3
End Property

Public Property Let dValorInicialConsig(ByVal vData As Double)
    mvardValorInicialConsig = vData
End Property

Public Property Get dValorInicialConsig() As Double
    dValorInicialConsig = mvardValorInicialConsig
End Property

Public Property Let dValorInicialConsig3(ByVal vData As Double)
    mvardValorInicialConsig3 = vData
End Property

Public Property Get dValorInicialConsig3() As Double
    dValorInicialConsig3 = mvardValorInicialConsig3
End Property

Public Property Let dQuantInicialBenef3(ByVal vData As Double)
    mvardQuantInicialBenef3 = vData
End Property

Public Property Get dQuantInicialBenef3() As Double
    dQuantInicialBenef3 = mvardQuantInicialBenef3
End Property

Public Property Let dQuantInicialBenef(ByVal vData As Double)
    mvardQuantInicialBenef = vData
End Property

Public Property Get dQuantInicialBenef() As Double
    dQuantInicialBenef = mvardQuantInicialBenef
End Property

Public Property Let dQuantInicialOutras(ByVal vData As Double)
    mvardQuantInicialOutras = vData
End Property

Public Property Get dQuantInicialOutras() As Double
    dQuantInicialOutras = mvardQuantInicialOutras
End Property

Public Property Let dQuantInicialOutras3(ByVal vData As Double)
    mvardQuantInicialOutras3 = vData
End Property

Public Property Get dQuantInicialOutras3() As Double
    dQuantInicialOutras3 = mvardQuantInicialOutras3
End Property

Public Property Let dQuantInicialConserto(ByVal vData As Double)
    mvardQuantInicialConserto = vData
End Property

Public Property Get dQuantInicialConserto() As Double
    dQuantInicialConserto = mvardQuantInicialConserto
End Property

Public Property Let dQuantInicialConserto3(ByVal vData As Double)
    mvardQuantInicialConserto3 = vData
End Property

Public Property Get dQuantInicialConserto3() As Double
    dQuantInicialConserto3 = mvardQuantInicialConserto3
End Property

Public Property Let dQuantInicialDemo(ByVal vData As Double)
    mvardQuantInicialDemo = vData
End Property

Public Property Get dQuantInicialDemo() As Double
    dQuantInicialDemo = mvardQuantInicialDemo
End Property

Public Property Let dQuantInicialDemo3(ByVal vData As Double)
    mvardQuantInicialDemo3 = vData
End Property

Public Property Get dQuantInicialDemo3() As Double
    dQuantInicialDemo3 = mvardQuantInicialDemo3
End Property

Public Property Let dQuantInicialConsig(ByVal vData As Double)
    mvardQuantInicialConsig = vData
End Property

Public Property Get dQuantInicialConsig() As Double
    dQuantInicialConsig = mvardQuantInicialConsig
End Property

Public Property Let dQuantInicialConsig3(ByVal vData As Double)
    mvardQuantInicialConsig3 = vData
End Property

Public Property Get dQuantInicialConsig3() As Double
    dQuantInicialConsig3 = mvardQuantInicialConsig3
End Property

Public Property Let dValorBenef3(ByVal vData As Double)
    mvardValorBenef3 = vData
End Property

Public Property Get dValorBenef3() As Double
    dValorBenef3 = mvardValorBenef3
End Property

Public Property Let dValorBenef(ByVal vData As Double)
    mvardValorBenef = vData
End Property

Public Property Get dValorBenef() As Double
    dValorBenef = mvardValorBenef
End Property

Public Property Let dValorOutras3(ByVal vData As Double)
    mvardValorOutras3 = vData
End Property

Public Property Get dValorOutras3() As Double
    dValorOutras3 = mvardValorOutras3
End Property

Public Property Let dValorOutras(ByVal vData As Double)
    mvardValorOutras = vData
End Property

Public Property Get dValorOutras() As Double
    dValorOutras = mvardValorOutras
End Property

Public Property Let dValorConserto(ByVal vData As Double)
    mvardValorConserto = vData
End Property

Public Property Get dValorConserto() As Double
    dValorConserto = mvardValorConserto
End Property

Public Property Let dValorConserto3(ByVal vData As Double)
    mvardValorConserto3 = vData
End Property

Public Property Get dValorConserto3() As Double
    dValorConserto3 = mvardValorConserto3
End Property

Public Property Let dValorDemo3(ByVal vData As Double)
    mvardValorDemo3 = vData
End Property

Public Property Get dValorDemo3() As Double
    dValorDemo3 = mvardValorDemo3
End Property

Public Property Let dValorDemo(ByVal vData As Double)
    mvardValorDemo = vData
End Property

Public Property Get dValorDemo() As Double
    dValorDemo = mvardValorDemo
End Property

Public Property Let dValorConsig3(ByVal vData As Double)
    mvardValorConsig3 = vData
End Property

Public Property Get dValorConsig3() As Double
    dValorConsig3 = mvardValorConsig3
End Property

Public Property Let dValorConsig(ByVal vData As Double)
    mvardValorConsig = vData
End Property

Public Property Get dValorConsig() As Double
    dValorConsig = mvardValorConsig
End Property

Public Property Let dQuantPedido(ByVal vData As Double)
Attribute dQuantPedido.VB_Description = "Quantidade do Produto em Pedido"
    mvardQuantPedido = vData
End Property

Public Property Get dQuantPedido() As Double
    dQuantPedido = mvardQuantPedido
End Property

Public Property Let dQuantBenef3(ByVal vData As Double)
Attribute dQuantBenef3.VB_Description = "Quantidade de Terceiros em beneficiamento"
    mvardQuantBenef3 = vData
End Property

Public Property Get dQuantBenef3() As Double
    dQuantBenef3 = mvardQuantBenef3
End Property

Public Property Let dQuantBenef(ByVal vData As Double)
Attribute dQuantBenef.VB_Description = "Quantidade Nossa em beneficiamento"
    mvardQuantBenef = vData
End Property

Public Property Get dQuantBenef() As Double
    dQuantBenef = mvardQuantBenef
End Property

Public Property Let dQuantOP(ByVal vData As Double)
    mvardQuantOP = vData
End Property

Public Property Get dQuantOP() As Double
    dQuantOP = mvardQuantOP
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Get dQuantDisponivel() As Double
    dQuantDisponivel = mvardQuantDispNossa + mvardQuantConsig3
End Property

Public Property Let dQuantOutras3(ByVal vData As Double)
    mvardQuantOutras3 = vData
End Property

Public Property Get dQuantOutras3() As Double
    dQuantOutras3 = mvardQuantOutras3
End Property

Public Property Let dQuantOutras(ByVal vData As Double)
    mvardQuantOutras = vData
End Property

Public Property Get dQuantOutras() As Double
    dQuantOutras = mvardQuantOutras
End Property

Public Property Let dQuantConserto(ByVal vData As Double)
    mvardQuantConserto = vData
End Property

Public Property Get dQuantConserto() As Double
    dQuantConserto = mvardQuantConserto
End Property

Public Property Let dQuantConserto3(ByVal vData As Double)
    mvardQuantConserto3 = vData
End Property

Public Property Get dQuantConserto3() As Double
    dQuantConserto3 = mvardQuantConserto3
End Property

Public Property Let dQuantDemo3(ByVal vData As Double)
    mvardQuantDemo3 = vData
End Property

Public Property Get dQuantDemo3() As Double
    dQuantDemo3 = mvardQuantDemo3
End Property

Public Property Let dQuantDemo(ByVal vData As Double)
    mvardQuantDemo = vData
End Property

Public Property Get dQuantDemo() As Double
    dQuantDemo = mvardQuantDemo
End Property

Public Property Let dQuantConsig3(ByVal vData As Double)
    mvardQuantConsig3 = vData
End Property

Public Property Get dQuantConsig3() As Double
    dQuantConsig3 = mvardQuantConsig3
End Property

Public Property Let dQuantConsig(ByVal vData As Double)
    mvardQuantConsig = vData
End Property

Public Property Get dQuantConsig() As Double
    dQuantConsig = mvardQuantConsig
End Property

Public Property Let dQuantDefeituosa(ByVal vData As Double)
    mvardQuantDefeituosa = vData
End Property

Public Property Get dQuantDefeituosa() As Double
    dQuantDefeituosa = mvardQuantDefeituosa
End Property

Public Property Let dQuantInd(ByVal vData As Double)
    mvardQuantInd = vData
End Property

Public Property Get dQuantInd() As Double
    dQuantInd = mvardQuantInd
End Property

Public Property Let dQuantRecIndl(ByVal vData As Double)
    mvardQuantRecIndl = vData
End Property

Public Property Get dQuantRecIndl() As Double
    dQuantRecIndl = mvardQuantRecIndl
End Property

Public Property Let dQuantEmpenhada(ByVal vData As Double)
    mvardQuantEmpenhada = vData
End Property

Public Property Get dQuantEmpenhada() As Double
    dQuantEmpenhada = mvardQuantEmpenhada
End Property

Public Property Let dSaldoInicial(ByVal vData As Double)
    mvardSaldoInicial = vData
End Property

Public Property Get dSaldoInicial() As Double
    dSaldoInicial = mvardSaldoInicial
End Property

Public Property Let dQuantidadeInicial(ByVal vData As Double)
    mvardQuantidadeInicial = vData
End Property

Public Property Get dQuantidadeInicial() As Double
    dQuantidadeInicial = mvardQuantidadeInicial
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let dtDataInventario(ByVal vData As Date)
    mvardtDataInventario = vData
End Property

Public Property Get dtDataInventario() As Date
    dtDataInventario = mvardtDataInventario
End Property

Public Property Let dQuantReservadaConsig(ByVal vData As Double)
    mvardQuantReservadaConsig = vData
End Property

Public Property Get dQuantReservadaConsig() As Double
    dQuantReservadaConsig = mvardQuantReservadaConsig
End Property

Public Property Let dQuantReservada(ByVal vData As Double)
    mvardQuantReservada = vData
End Property

Public Property Get dQuantReservada() As Double
    dQuantReservada = mvardQuantReservada
End Property

Public Property Let sLocalizacaoFisica(ByVal vData As String)
    mvarsLocalizacaoFisica = vData
End Property

Public Property Get sLocalizacaoFisica() As String
    sLocalizacaoFisica = mvarsLocalizacaoFisica
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let dQuantDispNossa(ByVal vData As Double)
    mvardQuantDispNossa = vData
End Property

Public Property Get dQuantDispNossa() As Double
    dQuantDispNossa = mvardQuantDispNossa
End Property

Public Property Let sAlmoxarifadoNomeReduzido(ByVal vData As String)
    mvarsAlmoxarifadoNomeReduzido = vData
End Property

Public Property Get sAlmoxarifadoNomeReduzido() As String
    sAlmoxarifadoNomeReduzido = mvarsAlmoxarifadoNomeReduzido
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property



