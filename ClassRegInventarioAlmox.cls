VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRegInventarioAlmox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsProduto As String
Private mvardtData As Date
Private mvariAlmoxarifado As Integer
Private mvarlNumIntDocLivFechado As Long
Private mvardQuantidadeUMEstoque As Double
Private mvardValorUnitario As Double
Private mvardQtdeNossaEmTerc As Double
Private mvardQtdeDeTercConosco As Double
Private mvardQuantConserto As Double
Private mvardQuantDemo As Double
Private mvardQuantConsig As Double
Private mvardQuantBenef As Double
Private mvardQuantOutras As Double
Private mvardCustoConserto As Double
Private mvardCustoDemo As Double
Private mvardCustoConsig As Double
Private mvardCustoBenef As Double
Private mvardCustoOutras As Double
Private mvardQuantConserto3 As Double
Private mvardQuantDemo3 As Double
Private mvardQuantConsig3 As Double
Private mvardQuantBenef3 As Double
Private mvardQuantOutras3 As Double
Private mvardCustoConserto3 As Double
Private mvardCustoDemo3 As Double
Private mvardCustoConsig3 As Double
Private mvardCustoBenef3 As Double
Private mvardCustoOutras3 As Double
Private mvardValorEstoque As Double
Private mvardValorBenef As Double
Private mvardValorBenef3 As Double
Private mvardValorConserto As Double
Private mvardValorConserto3 As Double
Private mvardValorConsig As Double
Private mvardValorConsig3 As Double
Private mvardValorDemo As Double
Private mvardValorDemo3 As Double
Private mvardValorOutras As Double
Private mvardValorOutras3 As Double
Private mvardSaldoQuantCusto As Double
Private mvardSaldoValorCusto As Double

Public Property Let dSaldoValorCusto(ByVal vData As Double)
    mvardSaldoValorCusto = vData
End Property

Public Property Get dSaldoValorCusto() As Double
    dSaldoValorCusto = mvardSaldoValorCusto
End Property

Public Property Let dSaldoQuantCusto(ByVal vData As Double)
    mvardSaldoQuantCusto = vData
End Property

Public Property Get dSaldoQuantCusto() As Double
    dSaldoQuantCusto = mvardSaldoQuantCusto
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

Public Property Let dValorConserto3(ByVal vData As Double)
    mvardValorConserto3 = vData
End Property

Public Property Get dValorConserto3() As Double
    dValorConserto3 = mvardValorConserto3
End Property

Public Property Let dValorConserto(ByVal vData As Double)
    mvardValorConserto = vData
End Property

Public Property Get dValorConserto() As Double
    dValorConserto = mvardValorConserto
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

Public Property Let dValorEstoque(ByVal vData As Double)
    mvardValorEstoque = vData
End Property

Public Property Get dValorEstoque() As Double
    dValorEstoque = mvardValorEstoque
End Property

Public Property Let dCustoOutras3(ByVal vData As Double)
    mvardCustoOutras3 = vData
End Property

Public Property Get dCustoOutras3() As Double
    dCustoOutras3 = mvardCustoOutras3
End Property

Public Property Let dCustoBenef3(ByVal vData As Double)
    mvardCustoBenef3 = vData
End Property

Public Property Get dCustoBenef3() As Double
    dCustoBenef3 = mvardCustoBenef3
End Property

Public Property Let dCustoConsig3(ByVal vData As Double)
    mvardCustoConsig3 = vData
End Property

Public Property Get dCustoConsig3() As Double
    dCustoConsig3 = mvardCustoConsig3
End Property

Public Property Let dCustoDemo3(ByVal vData As Double)
    mvardCustoDemo3 = vData
End Property

Public Property Get dCustoDemo3() As Double
    dCustoDemo3 = mvardCustoDemo3
End Property

Public Property Let dCustoConserto3(ByVal vData As Double)
    mvardCustoConserto3 = vData
End Property

Public Property Get dCustoConserto3() As Double
    dCustoConserto3 = mvardCustoConserto3
End Property

Public Property Let dQuantOutras3(ByVal vData As Double)
    mvardQuantOutras3 = vData
End Property

Public Property Get dQuantOutras3() As Double
    dQuantOutras3 = mvardQuantOutras3
End Property

Public Property Let dQuantBenef3(ByVal vData As Double)
    mvardQuantBenef3 = vData
End Property

Public Property Get dQuantBenef3() As Double
    dQuantBenef3 = mvardQuantBenef3
End Property

Public Property Let dQuantConsig3(ByVal vData As Double)
    mvardQuantConsig3 = vData
End Property

Public Property Get dQuantConsig3() As Double
    dQuantConsig3 = mvardQuantConsig3
End Property

Public Property Let dQuantDemo3(ByVal vData As Double)
    mvardQuantDemo3 = vData
End Property

Public Property Get dQuantDemo3() As Double
    dQuantDemo3 = mvardQuantDemo3
End Property

Public Property Let dQuantConserto3(ByVal vData As Double)
    mvardQuantConserto3 = vData
End Property

Public Property Get dQuantConserto3() As Double
    dQuantConserto3 = mvardQuantConserto3
End Property

Public Property Let dCustoOutras(ByVal vData As Double)
    mvardCustoOutras = vData
End Property

Public Property Get dCustoOutras() As Double
    dCustoOutras = mvardCustoOutras
End Property

Public Property Let dCustoBenef(ByVal vData As Double)
    mvardCustoBenef = vData
End Property

Public Property Get dCustoBenef() As Double
    dCustoBenef = mvardCustoBenef
End Property

Public Property Let dCustoConsig(ByVal vData As Double)
    mvardCustoConsig = vData
End Property

Public Property Get dCustoConsig() As Double
    dCustoConsig = mvardCustoConsig
End Property

Public Property Let dCustoDemo(ByVal vData As Double)
    mvardCustoDemo = vData
End Property

Public Property Get dCustoDemo() As Double
    dCustoDemo = mvardCustoDemo
End Property

Public Property Let dCustoConserto(ByVal vData As Double)
    mvardCustoConserto = vData
End Property

Public Property Get dCustoConserto() As Double
    dCustoConserto = mvardCustoConserto
End Property

Public Property Let dQuantOutras(ByVal vData As Double)
    mvardQuantOutras = vData
End Property

Public Property Get dQuantOutras() As Double
    dQuantOutras = mvardQuantOutras
End Property

Public Property Let dQuantBenef(ByVal vData As Double)
    mvardQuantBenef = vData
End Property

Public Property Get dQuantBenef() As Double
    dQuantBenef = mvardQuantBenef
End Property

Public Property Let dQuantConsig(ByVal vData As Double)
    mvardQuantConsig = vData
End Property

Public Property Get dQuantConsig() As Double
    dQuantConsig = mvardQuantConsig
End Property

Public Property Let dQuantDemo(ByVal vData As Double)
    mvardQuantDemo = vData
End Property

Public Property Get dQuantDemo() As Double
    dQuantDemo = mvardQuantDemo
End Property

Public Property Let dQuantConserto(ByVal vData As Double)
    mvardQuantConserto = vData
End Property

Public Property Get dQuantConserto() As Double
    dQuantConserto = mvardQuantConserto
End Property

Public Property Let dQtdeDeTercConosco(ByVal vData As Double)
Attribute dQtdeDeTercConosco.VB_Description = "Quantidade de terceiros em nosso poder"
    mvardQtdeDeTercConosco = vData
End Property

Public Property Get dQtdeDeTercConosco() As Double
    dQtdeDeTercConosco = mvardQtdeDeTercConosco
End Property

Public Property Let dQtdeNossaEmTerc(ByVal vData As Double)
Attribute dQtdeNossaEmTerc.VB_Description = "Quantidade nossa em poder de terceiros"
    mvardQtdeNossaEmTerc = vData
End Property

Public Property Get dQtdeNossaEmTerc() As Double
    dQtdeNossaEmTerc = mvardQtdeNossaEmTerc
End Property

Public Property Let dValorUnitario(ByVal vData As Double)
Attribute dValorUnitario.VB_Description = "Valor Unitário (Custo)"
    mvardValorUnitario = vData
End Property

Public Property Get dValorUnitario() As Double
    dValorUnitario = mvardValorUnitario
End Property

Public Property Let dQuantidadeUMEstoque(ByVal vData As Double)
Attribute dQuantidadeUMEstoque.VB_Description = "Quantidade do Produto na unidade de medida De Estoque"
    mvardQuantidadeUMEstoque = vData
End Property

Public Property Get dQuantidadeUMEstoque() As Double
    dQuantidadeUMEstoque = mvardQuantidadeUMEstoque
End Property

Public Property Let lNumIntDocLivFechado(ByVal vData As Long)
Attribute lNumIntDocLivFechado.VB_Description = "Número Intero do Livro Fechado"
    mvarlNumIntDocLivFechado = vData
End Property

Public Property Get lNumIntDocLivFechado() As Long
    lNumIntDocLivFechado = mvarlNumIntDocLivFechado
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
Attribute iAlmoxarifado.VB_Description = "Almoxarifado"
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let dtData(ByVal vData As Date)
Attribute dtData.VB_Description = "Data do Balanço"
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let sProduto(ByVal vData As String)
Attribute sProduto.VB_Description = "Identificador do Produto"
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property
