Attribute VB_Name = "GlobalCRFATInpal"
Option Explicit

'ALTERAÇÕES INPAL
'1 - 08/2000 Marcio - Incluido os campos em Cliente iPadraoTaxaFin, iPadraoJuros, dTaxaFin, dJuros
'2 - 09/2000 Marcio - Incluido os campos ValorOriginal e MotivoDiferença em ParcelasPagar

'TIPOS PARA AJUDA DE CUSTO
Public Const AJUDACUSTO_MINIMA = 0
Public Const AJUDACUSTO_FIXA = 1

'Padrão
Public Const VALOR_PADRAO = 1
Public Const VALOR_NAO_PADRAO = 0

'Mnemônicos para cálculo de comissões
Public Const MNEMONICO_COMISSOES_PRECO_TABELA_INPAL = "PrecoTabelaInpal"
Public Const MNEMONICO_COMISSOES_PRECO_BASE_INPAL = "PrecoBaseInpal"

'FORMULAS
Public Const FORMULA_PERCCOMISSAO = "IF(" & MNEMONICO_COMISSOES_PRECO_VENDA & CARACTER_TODAS_LINHAS_GRID & "<" & MNEMONICO_COMISSOES_PRECO_TABELA_INPAL & " OU " & MNEMONICO_COMISSOES_PRECO_TABELA_INPAL & "=0"

'PESOS DAS REGRAS DE COMISSÕES
Public Const PESO_COMISSOES_REGIAO = 100
Public Const PESO_COMISSOES_CLIENTE = 1000
Public Const PESO_COMISSOES_FILIALCLI = 10000
Public Const PESO_COMISSOES_ITEMCATPRODUTO = 10

'NÚMERO MÁXIMO DE REGRAS NO GRID DA TELA COMISSOESINPALPLAN
Public Const NUM_MAX_REGRAS_COMISSOES_INPAL = 999

'1 - 08/2000 Marcio - Incluido os campos em Cliente iPadraoTaxaFin, iPadraoJuros, dTaxaFin, dJuros
Type typeClienteUsu
    iPadraoTaxaFin As Integer 'campo incluido p/INPAL
    iPadraoJuros As Integer     'campo incluido p/INPAL
    dTaxaFinanceira As Double       'campo incluido p/INPAL
    dJuros As Double                'campo incluido p/INPAL
End Type

'Incluído por Luiz em 31/01/02
'Esse type só existe para a Inpal, pois as funções que o utilizam também são
'específicas para a Inpal
Type typeArqComissoes
    iFilialEmp As Integer
    lNumNotaFiscal As Long
    sSerie As String
    iItemNF As Integer
    dtDataEmissao As Date
    iCondPagto As Integer
    sProduto As String
    sNomeProduto As String
    sLinha As String
    sGrupo As String
    sSubGrupo As String
    sNomeGrupo As String
    sClassFiscal As String
    lCodCliente As Long
    iFilialCliente As Integer
    sRazaoSocCliente As String
    iRegiao As Integer
    iCodVendedor As Integer
    sNomeVendedor As String
    sNatOpInt As String
    sNatOp As String
    sVenda As String
    dQuant As Double
    sUM As String
    dPesoSemConv As Double
    dPrecoUn As Double
    dDespFinanceira As Double
    dAliquotaIPI As Double
    iStatusNota As Integer
    dDesconto As Double
    dPrecoAVista As Double
End Type


' *** TYPES CRIADOS PARA TELA DE REGRAS DE COMISSÕES DA INPAL ***
'Type de Planilhas
Type typeComissoesInpalPlan
    lCodigo As Long
    iVendedor As Integer
    iTecnico As Integer
    dPercComissaoEmissao As Double
    dPercComissaoBaixa As Double
    iComissaoSobreTotal As Integer
    iComissaoFrete As Integer
    iComissaoDesp As Integer
    iComissaoIPI As Integer
    iComissaoSeguro As Integer
    dAjudaCusto As Double
    iTipoAjudaCusto As Integer
End Type

'Type de Regras
Type typeComissoesInpalRegras
    lNumIntDoc As Long
    lCodPlanilha As Long
    iRegiaoVenda As Integer
    lCliente As Long
    iFilialCliente As Integer
    sCategoriaProduto As String
    sItemCatProduto As String
    dPercTabelaA As Double
    dPercTabelaB As Double
End Type
' ***************************************************************

'*** 11/04/02 - INÍCIO Luiz G.F.Nogueira ***
'Type utilizado para manipular a tabela TabelasDePreco que foi customizada para Inpal
'DIFERENÇA: inclusão do campo AliquotaICMS
Type typeTabelasDePreco
    iCodigo As Integer
    sDescricao As String
    dAliquotaICMS As Double
End Type
'*** 11/04/02 - FIM Luiz G.F.Nogueira ***

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRegTick
    lTick_Indice As Long
    sTick_Sequencial As String
    sTick_PlacaCarreta As String
    sTick_PlacaVeiculo As String
    sTick_CodTransportadora As String
    sTick_CodEmissor As String
    sTick_CodItem As String
    sTick_SttFim As String
    sTick_Status As String
    sTick_RecExp As String
    dTick_SeqRecExp As Double
    sTick_DescrStt As String
    sTick_DescrRecExp As String
    dTick_PesoDosagem As Double
    sTick_RazSocTrans As String
    sTick_RazSocEmissor As String
    sTick_DescricaoItem As String
    dTick_PesoInicial As Double
    sTick_OpPesoInicial As String
    dtTick_DtHrPesoInicial As Date
    sTick_BalPesoInicial As String
    dTick_PesoFinal As Double
    sTick_OpPesoFinal As String
    dtTick_DtHrPesoFinal As Date
    sTick_BalPesoFinal As String
    dTick_PesoLiquido As Double
    dTick_FatCorrecao As Double
    dTick_FatConversao As Double
    dTick_LiquidoCorrigido As Double
    sTick_UnidadeAposConversao As String
    sTick_DescDocTot As String
    dTick_PesoTotDoc As Double
    dTick_DifOrigemRealPc As Double
    dTick_DifOrigemRealKg As Double
    dTick_PesoLiquido1 As Double
    dTick_PesoLiquido2 As Double
    dTick_PesoLiquido3 As Double
    dTick_PesoLiquido4 As Double
    dTick_PesoLiquido5 As Double
    dTick_FatCorrecao1 As Double
    dTick_FatCorrecao2 As Double
    dTick_FatCorrecao3 As Double
    dTick_FatCorrecao4 As Double
    dTick_FatCorrecao5 As Double
    dTick_FatCorrecao6 As Double
    sTick_Observacao1 As String
    sTick_Observacao2 As String
    sTick_SenhaPreEntrada As String
    dtTick_DtHrPreEntrada As Date
    sTick_OpPreEntrada As String
    sTick_MemoObs As String
    dTick_PesoCavInic As Double
    sTick_PlacaCavInic As String
    dTick_PesoCavFinal As Double
    sTick_PlacaCavFinal As String
    lTick_QtdeTara1 As Long
    dTick_PesoTara1 As Double
    dTick_TotTara1 As Double
    lTick_QtdeTara2 As Long
    dTick_PesoTara2 As Double
    dTick_TotTara2 As Double
    dTick_TotTaras As Double
    dTick_PesoM3 As Double
    iTick_HabCavalo As Integer
    lTick_Compartimentos As Long
    dtTick_DtHrPosPesa As Date
    sTick_TmpVeicEmpr As String
    sTick_NumTransp As String
    sTick_CampoUsu1 As String
    sTick_CampoUsu2 As String
    sTick_CampoUsu3 As String
    sTick_CampoUsu4 As String
    sTick_CampoUsu5 As String
    sTick_CampoUsu6 As String
    sTick_CampoUsu7 As String
    sTick_CampoUsu8 As String
    sTick_CampoUsu9 As String
    sTick_CampoUsu10 As String
    dTick_PesoLiqCorrUsu As Double
    lTick_NumOcupante As Long
    sCgc As String
End Type

Public Function UnidadeDeMedida_IgnorarNaVenda(ByVal sUnidadeMed As String) As Boolean

    Select Case sUnidadeMed
        Case "PAR", "RL", "PC", "DIV", "TB", "SERVI"
            UnidadeDeMedida_IgnorarNaVenda = True
        Case Else
            UnidadeDeMedida_IgnorarNaVenda = False
    End Select
    
End Function
