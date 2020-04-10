Attribute VB_Name = "ADCAPI1N"
Option Explicit

Type typeProdutoFilial
    iFilialEmpresa As Integer
    sProduto As String
    iAlmoxarifado As Integer
    lFornecedor As Long
    iVisibilidadeAlmoxarifados As Integer
    dEstoqueSeguranca As Double
    dEstoqueMaximo As Double
    dPontoPedido As Double
    sClasseABC As String
    dLoteEconomico As Double
    iIntRessup As Integer
    iTabelaPreco As Integer
    dQuantPedida As Double
End Type

Type typeProduto
    sCodigo As String
    iTipo As Integer
    sDescricao As String
    sNomeReduzido As String
    sModelo As String
    iGerencial As Integer
    iNivel As Integer
    sSubstituto1 As String
    sSubstituto2 As String
    iPrazoValidade As Integer
    sCodigoBarras As String
    iEtiquetasCodBarras As Integer
    dPesoLiq As Double
    dPesoBruto As Double
    dComprimento As Double
    dEspessura As Double
    dLargura As Double
    sCor As String
    sObsFisica As String
    iClasseUM As Integer
    sSiglaUMEstoque As String
    sSiglaUMCompra As String
    sSiglaUMVenda As String
    iAtivo As Integer
    iFaturamento As Integer
    iCompras As Integer
    iPCP As Integer
    iKitBasico As Integer
    iKitInt As Integer
    dIPIAliquota As Double
    sIPICodigo As String
    sIPICodDIPI As String
'    dISSAliquota As Double
'    sISSCodigo As String
'    iIRIncide As Integer
    iControleEstoque As Integer
    iICMSAgregaCusto As Integer
    iIPIAgregaCusto As Integer
    iFreteAgregaCusto As Integer
    iApropriacaoCusto As Integer
    sContaContabil As String
    sContaContabilProducao As String
    dResiduo As Double
    iNatureza As Integer
    dCustoReposicao As Double
    iOrigemMercadoria As Integer
    iTabelaPreco As Integer
End Type

Type typeUsuario
    sCodUsuario As String
    iLote As Integer
End Type


