VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTipoDeProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariTipo As Integer 'local copy
Private mvarsDescricao As String 'local copy
Private mvariClasseUM As Integer 'local copy
Private mvarsSiglaUMEstoque As String 'local copy
Private mvarsSiglaUMCompra As String 'local copy
Private mvarsSiglaUMVenda As String 'local copy
Private mvariControleEstoque As Integer 'local copy
Private mvariPrazoValidade As Integer 'local copy
Private mvariFaturamento As Integer 'local copy
Private mvariCompras As Integer 'local copy
Private mvariPCP As Integer 'local copy
Private mvariKitBasico As Integer 'local copy
Private mvariKitInt As Integer 'local copy
Private mvarsIPICodigo As String 'local copy
Private mvarsIPICodDIPI As String 'local copy
Private mvardIPIAliquota As Double 'local copy
Private mvardISSAliquota As Double 'local copy
Private mvarsISSCodigo As String 'local copy
Private mvariIRIncide As Integer 'local copy
Private mvariICMSAgregaCusto As Integer 'local copy
Private mvariIPIAgregaCusto As Integer 'local copy
Private mvariFreteAgregaCusto As Integer 'local copy
Private mvariApropriacaoCusto As Integer 'local copy
Private mvarsContaContabil As String 'local copy
Private mvariIntRessup As Integer 'local copy
Private mvardTempoRessupMax As Double 'local copy
Private mvardConsumoMedioMax As Double 'local copy
Private mvardResiduo As Double 'local copy
Private mvariMesesConsumoMedio As Integer 'local copy
Private mvariConsideraQuantCotAnt As Integer 'local copy
Private mvardPercentMaisQuantCotAnt As Double 'local copy
Private mvardPercentMenosQuantCotAnt As Double 'local copy
Private mvariTemFaixaReceb As Integer 'local copy
Private mvardPercentMaisReceb As Double 'local copy
Private mvardPercentMenosReceb As Double 'local copy
Private mvariRecebForaFaixa As Integer 'local copy
Private mvarsSigla As String 'local copy
Private mvarcolCategoriaItem As New Collection  'local copy
'local variable(s) to hold property value(s)
Private mvariNatureza As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarsContaProducao As String 'local copy
Private mvariKitVendaComp As Integer
Private mvariExTIPI As Integer
Private mvariProdutoEspecifico As Integer
Private mvarsGenero As String
Private mvarsISSQN As String
Private mvarsSiglaUMTrib As String
Private mvariOrigem As Integer
Private mvarsCEST As String

Public Property Let sCEST(ByVal vData As String)
   mvarsCEST = vData
End Property

Public Property Get sCEST() As String
    sCEST = mvarsCEST
End Property

Public Property Let iOrigem(ByVal vData As Integer)
   mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let sSiglaUMTrib(ByVal vData As String)
   mvarsSiglaUMTrib = vData
End Property

Public Property Get sSiglaUMTrib() As String
    sSiglaUMTrib = mvarsSiglaUMTrib
End Property

Public Property Let sISSQN(ByVal vData As String)
   mvarsISSQN = vData
End Property

Public Property Get sISSQN() As String
    sISSQN = mvarsISSQN
End Property

Public Property Let sGenero(ByVal vData As String)
   mvarsGenero = vData
End Property

Public Property Get sGenero() As String
    sGenero = mvarsGenero
End Property

Public Property Let iProdutoEspecifico(ByVal vData As Integer)
   mvariProdutoEspecifico = vData
End Property

Public Property Get iProdutoEspecifico() As Integer
    iProdutoEspecifico = mvariProdutoEspecifico
End Property

Public Property Let iExTIPI(ByVal vData As Integer)
   mvariExTIPI = vData
End Property

Public Property Get iExTIPI() As Integer
    iExTIPI = mvariExTIPI
End Property

Public Property Let iKitVendaComp(ByVal vData As Integer)
   mvariKitVendaComp = vData
End Property

Public Property Get iKitVendaComp() As Integer
    iKitVendaComp = mvariKitVendaComp
End Property

Public Property Let sContaProducao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaProducao = 5
    mvarsContaProducao = vData
End Property


Public Property Get sContaProducao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaProducao
    sContaProducao = mvarsContaProducao
End Property



Public Property Let iNatureza(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iNatureza = 5
    mvariNatureza = vData
End Property


Public Property Get iNatureza() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iNatureza
    iNatureza = mvariNatureza
End Property



Public Property Set colCategoriaItem(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colCategoriaItem = Form1
    Set mvarcolCategoriaItem = vData
End Property


Public Property Get colCategoriaItem() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colCategoriaItem
    Set colCategoriaItem = mvarcolCategoriaItem
End Property



Public Property Let sSigla(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSigla = 5
    mvarsSigla = vData
End Property


Public Property Get sSigla() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSigla
    sSigla = mvarsSigla
End Property



Public Property Let iRecebForaFaixa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iRecebForaFaixa = 5
    mvariRecebForaFaixa = vData
End Property


Public Property Get iRecebForaFaixa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iRecebForaFaixa
    iRecebForaFaixa = mvariRecebForaFaixa
End Property



Public Property Let dPercentMenosReceb(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMenosReceb = 5
    mvardPercentMenosReceb = vData
End Property


Public Property Get dPercentMenosReceb() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentMenosReceb
    dPercentMenosReceb = mvardPercentMenosReceb
End Property



Public Property Let dPercentMaisReceb(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMaisReceb = 5
    mvardPercentMaisReceb = vData
End Property


Public Property Get dPercentMaisReceb() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentMaisReceb
    dPercentMaisReceb = mvardPercentMaisReceb
End Property



Public Property Let iTemFaixaReceb(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTemFaixaReceb = 5
    mvariTemFaixaReceb = vData
End Property


Public Property Get iTemFaixaReceb() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTemFaixaReceb
    iTemFaixaReceb = mvariTemFaixaReceb
End Property



Public Property Let dPercentMenosQuantCotAnt(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMenosQuantCotAnt = 5
    mvardPercentMenosQuantCotAnt = vData
End Property


Public Property Get dPercentMenosQuantCotAnt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentMenosQuantCotAnt
    dPercentMenosQuantCotAnt = mvardPercentMenosQuantCotAnt
End Property



Public Property Let dPercentMaisQuantCotAnt(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMaisQuantCotAnt = 5
    mvardPercentMaisQuantCotAnt = vData
End Property


Public Property Get dPercentMaisQuantCotAnt() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentMaisQuantCotAnt
    dPercentMaisQuantCotAnt = mvardPercentMaisQuantCotAnt
End Property



Public Property Let iConsideraQuantCotAnt(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iConsideraQuantCotAnt = 5
    mvariConsideraQuantCotAnt = vData
End Property


Public Property Get iConsideraQuantCotAnt() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iConsideraQuantCotAnt
    iConsideraQuantCotAnt = mvariConsideraQuantCotAnt
End Property



Public Property Let iMesesConsumoMedio(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMesesConsumoMedio = 5
    mvariMesesConsumoMedio = vData
End Property


Public Property Get iMesesConsumoMedio() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMesesConsumoMedio
    iMesesConsumoMedio = mvariMesesConsumoMedio
End Property



Public Property Let dResiduo(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dResiduo = 5
    mvardResiduo = vData
End Property


Public Property Get dResiduo() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dResiduo
    dResiduo = mvardResiduo
End Property



Public Property Let dConsumoMedioMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dConsumoMedioMax = 5
    mvardConsumoMedioMax = vData
End Property


Public Property Get dConsumoMedioMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dConsumoMedioMax
    dConsumoMedioMax = mvardConsumoMedioMax
End Property



Public Property Let dTempoRessupMax(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTempoRessupMax = 5
    mvardTempoRessupMax = vData
End Property


Public Property Get dTempoRessupMax() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTempoRessupMax
    dTempoRessupMax = mvardTempoRessupMax
End Property



Public Property Let iIntRessup(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIntRessup = 5
    mvariIntRessup = vData
End Property


Public Property Get iIntRessup() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIntRessup
    iIntRessup = mvariIntRessup
End Property



Public Property Let sContaContabil(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
End Property



Public Property Let iApropriacaoCusto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iApropriacaoCusto = 5
    mvariApropriacaoCusto = vData
End Property


Public Property Get iApropriacaoCusto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iApropriacaoCusto
    iApropriacaoCusto = mvariApropriacaoCusto
End Property



Public Property Let iFreteAgregaCusto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFreteAgregaCusto = 5
    mvariFreteAgregaCusto = vData
End Property


Public Property Get iFreteAgregaCusto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFreteAgregaCusto
    iFreteAgregaCusto = mvariFreteAgregaCusto
End Property



Public Property Let iIPIAgregaCusto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIPIAgregaCusto = 5
    mvariIPIAgregaCusto = vData
End Property


Public Property Get iIPIAgregaCusto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIPIAgregaCusto
    iIPIAgregaCusto = mvariIPIAgregaCusto
End Property



Public Property Let iICMSAgregaCusto(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iICMSAgregaCusto = 5
    mvariICMSAgregaCusto = vData
End Property


Public Property Get iICMSAgregaCusto() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iICMSAgregaCusto
    iICMSAgregaCusto = mvariICMSAgregaCusto
End Property



Public Property Let iIRIncide(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIRIncide = 5
    mvariIRIncide = vData
End Property


Public Property Get iIRIncide() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIRIncide
    iIRIncide = mvariIRIncide
End Property



Public Property Let sISSCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sISSCodigo = 5
    mvarsISSCodigo = vData
End Property


Public Property Get sISSCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sISSCodigo
    sISSCodigo = mvarsISSCodigo
End Property



Public Property Let dISSAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dISSAliquota = 5
    mvardISSAliquota = vData
End Property


Public Property Get dISSAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dISSAliquota
    dISSAliquota = mvardISSAliquota
End Property



Public Property Let dIPIAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dIPIAliquota = 5
    mvardIPIAliquota = vData
End Property


Public Property Get dIPIAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dIPIAliquota
    dIPIAliquota = mvardIPIAliquota
End Property



Public Property Let sIPICodDIPI(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sIPICodDIPI = 5
    mvarsIPICodDIPI = vData
End Property


Public Property Get sIPICodDIPI() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sIPICodDIPI
    sIPICodDIPI = mvarsIPICodDIPI
End Property



Public Property Let sIPICodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sIPICodigo = 5
    mvarsIPICodigo = vData
End Property


Public Property Get sIPICodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sIPICodigo
    sIPICodigo = mvarsIPICodigo
End Property



Public Property Let iKitInt(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iKitInt = 5
    mvariKitInt = vData
End Property


Public Property Get iKitInt() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iKitInt
    iKitInt = mvariKitInt
End Property



Public Property Let iKitBasico(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iKitBasico = 5
    mvariKitBasico = vData
End Property


Public Property Get iKitBasico() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iKitBasico
    iKitBasico = mvariKitBasico
End Property



Public Property Let iPCP(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPCP = 5
    mvariPCP = vData
End Property


Public Property Get iPCP() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPCP
    iPCP = mvariPCP
End Property



Public Property Let iCompras(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCompras = 5
    mvariCompras = vData
End Property


Public Property Get iCompras() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCompras
    iCompras = mvariCompras
End Property



Public Property Let iFaturamento(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFaturamento = 5
    mvariFaturamento = vData
End Property


Public Property Get iFaturamento() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFaturamento
    iFaturamento = mvariFaturamento
End Property



Public Property Let iPrazoValidade(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPrazoValidade = 5
    mvariPrazoValidade = vData
End Property


Public Property Get iPrazoValidade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPrazoValidade
    iPrazoValidade = mvariPrazoValidade
End Property



Public Property Let iControleEstoque(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iControleEstoque = 5
    mvariControleEstoque = vData
End Property


Public Property Get iControleEstoque() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iControleEstoque
    iControleEstoque = mvariControleEstoque
End Property



Public Property Let sSiglaUMVenda(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaUMVenda = 5
    mvarsSiglaUMVenda = vData
End Property


Public Property Get sSiglaUMVenda() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaUMVenda
    sSiglaUMVenda = mvarsSiglaUMVenda
End Property



Public Property Let sSiglaUMCompra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaUMCompra = 5
    mvarsSiglaUMCompra = vData
End Property


Public Property Get sSiglaUMCompra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaUMCompra
    sSiglaUMCompra = mvarsSiglaUMCompra
End Property



Public Property Let sSiglaUMEstoque(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaUMEstoque = 5
    mvarsSiglaUMEstoque = vData
End Property


Public Property Get sSiglaUMEstoque() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaUMEstoque
    sSiglaUMEstoque = mvarsSiglaUMEstoque
End Property



Public Property Let iClasseUM(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iClasseUM = 5
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iClasseUM
    iClasseUM = mvariClasseUM
End Property



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let iTipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipo = 5
    mvariTipo = vData
End Property


Public Property Get iTipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipo
    iTipo = mvariTipo
End Property



