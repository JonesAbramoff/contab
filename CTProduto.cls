VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjAnotacao As ClassAnotacoes 'Wagner
Public gobjIdiomaTela As ClassIdiomaTela
Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'#####################################
'Inserido por Wagner 15/03/2006
'Acrescentado campos de Próximo Número de Série e Parte Numerica da Série
'e alterado o tabIndex
'#####################################

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim iClasseUMAlterada As Integer
Dim iClasseUMAnterior As Integer
Dim iTipoAlterado As Integer

'temporaria --> tirar de global loja e botar em globalmat
Const STRING_CONCATENACAO = 255

'temporario --> tirar de global loja e botar em globalmat
Type typeLog

    lNumIntDoc As Long
    iOperacao As Integer
    sLog1 As String
    sLog2 As String
    sLog3 As String
    sLog4 As String
    dData As Date
    dHora As Double
    
End Type


'***ALTERACAO POR TULIO EM 27/05***
'Variavel Global
Dim sProdutoAnterior As String
'***FIM ALTERACAO POR TULIO EM 27/05***

Private iProdutoAlterado As Integer
Dim iFabricante_ListIndex_Padrao As Integer

Dim objGridCategoria As AdmGrid
Dim iGrid_Categoria_Col As Integer
Dim iGrid_Valor_Col As Integer

Public objGridTabelaPreco As AdmGrid
Public iGrid_Tabela_Col
Public iGrid_DescricaoTabela_Col As Integer
Public iGrid_ValorEmpresa_Col As Integer
Public iGrid_ValorFilial_Col As Integer
Public iGrid_DataPreco_Col As Integer

Private WithEvents objEventoTabelaPrecoItem As AdmEvento
Attribute objEventoTabelaPrecoItem.VB_VarHelpID = -1
Private WithEvents objEventoTipoDeProduto As AdmEvento
Attribute objEventoTipoDeProduto.VB_VarHelpID = -1
Private WithEvents objEventoClasseUM As AdmEvento
Attribute objEventoClasseUM.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoProdutoSubst1 As AdmEvento
Attribute objEventoProdutoSubst1.VB_VarHelpID = -1
Private WithEvents objEventoProdutoSubst2 As AdmEvento
Attribute objEventoProdutoSubst2.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoContaProducao As AdmEvento
Attribute objEventoContaProducao.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoClasFiscIPI As AdmEvento
Attribute objEventoClasFiscIPI.VB_VarHelpID = -1
Private WithEvents objEventoGenero As AdmEvento
Attribute objEventoGenero.VB_VarHelpID = -1
Private WithEvents objEventoISSQN As AdmEvento
Attribute objEventoISSQN.VB_VarHelpID = -1
Private WithEvents objEventoCodServNFe As AdmEvento
Attribute objEventoCodServNFe.VB_VarHelpID = -1
Private WithEvents objEventoCEST As AdmEvento
Attribute objEventoCEST.VB_VarHelpID = -1
'''Private WithEvents objEventoEmbalagem As AdmEvento - 05/09/01 Marcelo

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Categoria = 2
Private Const TAB_Complemento = 3
Private Const TAB_CaracFisicas = 4
Private Const TAB_Precos = 5
Private Const TAB_UM = 6
Private Const TAB_Tributacao = 7

'--- inicio dos properties get dos controles da tela

Public Property Get cProdANVISA() As Object
     Set cProdANVISA = objUserControl.Controls("cProdANVISA")
End Property

Public Property Get ProdEmEscalaRelev() As Object
     Set ProdEmEscalaRelev = objUserControl.Controls("ProdEmEscalaRelev")
End Property

Public Property Get CEST() As Object
     Set CEST = objUserControl.Controls("CEST")
End Property

Public Property Get FCI() As Object
     Set FCI = objUserControl.Controls("FCI")
End Property

Public Property Get LabelFCI() As Object
     Set LabelFCI = objUserControl.Controls("LabelFCI")
End Property

Public Property Get FatorAjuste() As Object
     Set FatorAjuste = objUserControl.Controls("FatorAjuste")
End Property

Public Property Get LabelFatorAjuste() As Object
     Set LabelFatorAjuste = objUserControl.Controls("LabelFatorAjuste")
End Property

Public Property Get Fabricante() As Object
     Set Fabricante = objUserControl.Controls("Fabricante")
End Property

Public Property Get LabelGenero() As Object
     Set LabelGenero = objUserControl.Controls("LabelGenero")
End Property

Public Property Get Genero() As Object
     Set Genero = objUserControl.Controls("Genero")
End Property

Public Property Get DescGenero() As Object
     Set DescGenero = objUserControl.Controls("DescGenero")
End Property

Public Property Get LabelISSQN() As Object
     Set LabelISSQN = objUserControl.Controls("LabelISSQN")
End Property

Public Property Get ISSQN() As Object
     Set ISSQN = objUserControl.Controls("ISSQN")
End Property

Public Property Get DescISSQN() As Object
     Set DescISSQN = objUserControl.Controls("DescISSQN")
End Property

Public Property Get NomeUMTrib() As Object
     Set NomeUMTrib = objUserControl.Controls("NomeUMTrib")
End Property

Public Property Get CodServNFe() As Object
     Set CodServNFe = objUserControl.Controls("CodServNFe")
End Property

Public Property Get LabelCodServNFe() As Object
     Set LabelCodServNFe = objUserControl.Controls("LabelCodServNFe")
End Property

Public Property Get ExTIPI() As Object
     Set ExTIPI = objUserControl.Controls("ExTIPI")
End Property

Public Property Get SiglaUMTrib() As Object
     Set SiglaUMTrib = objUserControl.Controls("SiglaUMTrib")
End Property

Public Property Get ProdutoEspecifico() As Object
     Set ProdutoEspecifico = objUserControl.Controls("ProdutoEspecifico")
End Property

Public Property Get PrecoMaxConsumidor() As Object
     Set PrecoMaxConsumidor = objUserControl.Controls("PrecoMaxConsumidor")
End Property
Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get NivelGerencial() As Object
     Set NivelGerencial = objUserControl.Controls("NivelGerencial")
End Property

Public Property Get NivelFinal() As Object
     Set NivelFinal = objUserControl.Controls("NivelFinal")
End Property

Public Property Get BotaoVisualizar() As Object
     Set BotaoVisualizar = objUserControl.Controls("BotaoVisualizar")
End Property

Public Property Get BotaoProcurar() As Object
     Set BotaoProcurar = objUserControl.Controls("BotaoProcurar")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get NaturezaProduto() As Object
     Set NaturezaProduto = objUserControl.Controls("NaturezaProduto")
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objUserControl.Controls("NomeReduzido")
End Property

Public Property Get Modelo() As Object
     Set Modelo = objUserControl.Controls("Modelo")
End Property

Public Property Get Ativo() As Object
     Set Ativo = objUserControl.Controls("Ativo")
End Property

Public Property Get TipoProduto() As Object
     Set TipoProduto = objUserControl.Controls("TipoProduto")
End Property

Public Property Get Descricao() As Object
     Set Descricao = objUserControl.Controls("Descricao")
End Property

Public Property Get Referencia() As Object
     Set Referencia = objUserControl.Controls("Referencia")
End Property

Public Property Get NomeFigura() As Object
     Set NomeFigura = objUserControl.Controls("NomeFigura")
End Property

Public Property Get CommonDialog() As Object
     Set CommonDialog = objUserControl.Controls("CommonDialog")
End Property

Public Property Get Figura() As Object
     Set Figura = objUserControl.Controls("Figura")
End Property

Public Property Get DescTipoProduto() As Object
     Set DescTipoProduto = objUserControl.Controls("DescTipoProduto")
End Property

Public Property Get LblTipoProduto() As Object
     Set LblTipoProduto = objUserControl.Controls("LblTipoProduto")
End Property

Public Property Get LabelCodigo() As Object
     Set LabelCodigo = objUserControl.Controls("LabelCodigo")
End Property

Public Property Get LabelNomeReduzido() As Object
     Set LabelNomeReduzido = objUserControl.Controls("LabelNomeReduzido")
End Property

Public Property Get BotaoCorTonTP() As Object
     Set BotaoCorTonTP = objUserControl.Controls("BotaoCorTonTP")
End Property

Public Property Get FrameGrade() As Object
     Set FrameGrade = objUserControl.Controls("FrameGrade")
End Property

Public Property Get Grades() As Object
     Set Grades = objUserControl.Controls("Grades")
End Property

Public Property Get BotaoCriarGrade() As Object
     Set BotaoCriarGrade = objUserControl.Controls("BotaoCriarGrade")
End Property

Public Property Get LabelGrade() As Object
     Set LabelGrade = objUserControl.Controls("LabelGrade")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get ComboCategoriaProduto() As Object
     Set ComboCategoriaProduto = objUserControl.Controls("ComboCategoriaProduto")
End Property

Public Property Get ComboCategoriaProdutoItem() As Object
     Set ComboCategoriaProdutoItem = objUserControl.Controls("ComboCategoriaProdutoItem")
End Property

Public Property Get GridCategoria() As Object
     Set GridCategoria = objUserControl.Controls("GridCategoria")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get ListaCaracteristicas() As Object
     Set ListaCaracteristicas = objUserControl.Controls("ListaCaracteristicas")
End Property

Public Property Get Produzido() As Object
     Set Produzido = objUserControl.Controls("Produzido")
End Property

Public Property Get Comprado() As Object
     Set Comprado = objUserControl.Controls("Comprado")
End Property

Public Property Get Embalagem() As Object
     Set Embalagem = objUserControl.Controls("Embalagem")
End Property

Public Property Get Rastro() As Object
     Set Rastro = objUserControl.Controls("Rastro")
End Property

Public Property Get ObsFisica() As Object
     Set ObsFisica = objUserControl.Controls("ObsFisica")
End Property

Public Property Get Cor() As Object
     Set Cor = objUserControl.Controls("Cor")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get Comprimento() As Object
     Set Comprimento = objUserControl.Controls("Comprimento")
End Property

Public Property Get Largura() As Object
     Set Largura = objUserControl.Controls("Largura")
End Property

Public Property Get Espessura() As Object
     Set Espessura = objUserControl.Controls("Espessura")
End Property

Public Property Get PesoEspecifico() As Object
     Set PesoEspecifico = objUserControl.Controls("PesoEspecifico")
End Property

Public Property Get HorasMaquina() As Object
     Set HorasMaquina = objUserControl.Controls("HorasMaquina")
End Property

Public Property Get DescricaoEmbalagem() As Object
     Set DescricaoEmbalagem = objUserControl.Controls("DescricaoEmbalagem")
End Property

Public Property Get LabelEmbalagem() As Object
     Set LabelEmbalagem = objUserControl.Controls("LabelEmbalagem")
End Property

Public Property Get LabelMinutos() As Object
     Set LabelMinutos = objUserControl.Controls("LabelMinutos")
End Property

Public Property Get LabelHorasMaq() As Object
     Set LabelHorasMaq = objUserControl.Controls("LabelHorasMaq")
End Property

Public Property Get LabelRastro() As Object
     Set LabelRastro = objUserControl.Controls("LabelRastro")
End Property

Public Property Get LabelPesoEspKg() As Object
     Set LabelPesoEspKg = objUserControl.Controls("LabelPesoEspKg")
End Property

Public Property Get LabelPesoEspecifico() As Object
     Set LabelPesoEspecifico = objUserControl.Controls("LabelPesoEspecifico")
End Property

Public Property Get Frame20() As Object
     Set Frame20 = objUserControl.Controls("Frame20")
End Property

Public Property Get SerieProx() As Object
     Set SerieProx = objUserControl.Controls("SerieProx")
End Property

Public Property Get SerieNum() As Object
     Set SerieNum = objUserControl.Controls("SerieNum")
End Property

Public Property Get SeriePartNum() As Object
     Set SeriePartNum = objUserControl.Controls("SeriePartNum")
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
End Property

Public Property Get CodigoBarras() As Object
     Set CodigoBarras = objUserControl.Controls("CodigoBarras")
End Property

Public Property Get BotaoProdutoCodBarras() As Object
     Set BotaoProdutoCodBarras = objUserControl.Controls("BotaoProdutoCodBarras")
End Property

Public Property Get EtiquetasCodBarras() As Object
     Set EtiquetasCodBarras = objUserControl.Controls("EtiquetasCodBarras")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get Substituto1() As Object
     Set Substituto1 = objUserControl.Controls("Substituto1")
End Property

Public Property Get Substituto2() As Object
     Set Substituto2 = objUserControl.Controls("Substituto2")
End Property

Public Property Get DescSubst2() As Object
     Set DescSubst2 = objUserControl.Controls("DescSubst2")
End Property

Public Property Get DescSubst1() As Object
     Set DescSubst1 = objUserControl.Controls("DescSubst1")
End Property

Public Property Get LblSubst2() As Object
     Set LblSubst2 = objUserControl.Controls("LblSubst2")
End Property

Public Property Get LblSubst1() As Object
     Set LblSubst1 = objUserControl.Controls("LblSubst1")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get QuantPedido() As Object
     Set QuantPedido = objUserControl.Controls("QuantPedido")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get ApropriacaoComp() As Object
     Set ApropriacaoComp = objUserControl.Controls("ApropriacaoComp")
End Property

Public Property Get ApropriacaoProd() As Object
     Set ApropriacaoProd = objUserControl.Controls("ApropriacaoProd")
End Property

Public Property Get CustoReposicao() As Object
     Set CustoReposicao = objUserControl.Controls("CustoReposicao")
End Property

Public Property Get PrazoValidade() As Object
     Set PrazoValidade = objUserControl.Controls("PrazoValidade")
End Property

Public Property Get Residuo() As Object
     Set Residuo = objUserControl.Controls("Residuo")
End Property

Public Property Get TempoProducao() As Object
     Set TempoProducao = objUserControl.Controls("TempoProducao")
End Property

Public Property Get Frame11() As Object
     Set Frame11 = objUserControl.Controls("Frame11")
End Property

Public Property Get OrigemMercadoria() As Object
     Set OrigemMercadoria = objUserControl.Controls("OrigemMercadoria")
End Property

Public Property Get Frame19() As Object
     Set Frame19 = objUserControl.Controls("Frame19")
End Property

Public Property Get INSSPercBase() As Object
     Set INSSPercBase = objUserControl.Controls("INSSPercBase")
End Property

Public Property Get Frame18() As Object
     Set Frame18 = objUserControl.Controls("Frame18")
End Property

Public Property Get UsaBalanca() As Object
     Set UsaBalanca = objUserControl.Controls("UsaBalanca")
End Property

Public Property Get comboAliquota() As Object
     Set comboAliquota = objUserControl.Controls("comboAliquota")
End Property

Public Property Get SituacaoTributaria() As Object
     Set SituacaoTributaria = objUserControl.Controls("SituacaoTributaria")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get IncideIPI() As Object
     Set IncideIPI = objUserControl.Controls("IncideIPI")
End Property

Public Property Get AliquotaIPI() As Object
     Set AliquotaIPI = objUserControl.Controls("AliquotaIPI")
End Property

Public Property Get CodigoIPI() As Object
     Set CodigoIPI = objUserControl.Controls("CodigoIPI")
End Property

Public Property Get ClasFiscIPI() As Object
     Set ClasFiscIPI = objUserControl.Controls("ClasFiscIPI")
End Property

Public Property Get Frame15() As Object
     Set Frame15 = objUserControl.Controls("Frame15")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get ContaProducao() As Object
     Set ContaProducao = objUserControl.Controls("ContaProducao")
End Property

Public Property Get LabelContaProducao() As Object
     Set LabelContaProducao = objUserControl.Controls("LabelContaProducao")
End Property

Public Property Get ContaContabilLabel() As Object
     Set ContaContabilLabel = objUserControl.Controls("ContaContabilLabel")
End Property

Public Property Get LabelClassificacaoFiscal() As Object
     Set LabelClassificacaoFiscal = objUserControl.Controls("LabelClassificacaoFiscal")
End Property

Public Property Get BotaoTeste() As Object
     Set BotaoTeste = objUserControl.Controls("BotaoTeste")
End Property

Public Property Get Frame12() As Object
     Set Frame12 = objUserControl.Controls("Frame12")
End Property

Public Property Get NaoTemFaixaReceb() As Object
     Set NaoTemFaixaReceb = objUserControl.Controls("NaoTemFaixaReceb")
End Property

Public Property Get Frame13() As Object
     Set Frame13 = objUserControl.Controls("Frame13")
End Property

Public Property Get RecebForaFaixa() As Object
     Set RecebForaFaixa = objUserControl.Controls("RecebForaFaixa")
End Property

Public Property Get Frame14() As Object
     Set Frame14 = objUserControl.Controls("Frame14")
End Property

Public Property Get PercentMaisReceb() As Object
     Set PercentMaisReceb = objUserControl.Controls("PercentMaisReceb")
End Property

Public Property Get PercentMenosReceb() As Object
     Set PercentMenosReceb = objUserControl.Controls("PercentMenosReceb")
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objUserControl.Controls("Frame10")
End Property

Public Property Get ConsideraQuantCotacaoAnterior() As Object
     Set ConsideraQuantCotacaoAnterior = objUserControl.Controls("ConsideraQuantCotacaoAnterior")
End Property

Public Property Get PercentMaisQuantCotacaoAnterior() As Object
     Set PercentMaisQuantCotacaoAnterior = objUserControl.Controls("PercentMaisQuantCotacaoAnterior")
End Property

Public Property Get PercentMenosQuantCotacaoAnterior() As Object
     Set PercentMenosQuantCotacaoAnterior = objUserControl.Controls("PercentMenosQuantCotacaoAnterior")
End Property

Public Property Get Frame16() As Object
     Set Frame16 = objUserControl.Controls("Frame16")
End Property

Public Property Get Frame17() As Object
     Set Frame17 = objUserControl.Controls("Frame17")
End Property

Public Property Get SiglaUMEstoque() As Object
     Set SiglaUMEstoque = objUserControl.Controls("SiglaUMEstoque")
End Property

Public Property Get SiglaUMCompra() As Object
     Set SiglaUMCompra = objUserControl.Controls("SiglaUMCompra")
End Property

Public Property Get SiglaUMVenda() As Object
     Set SiglaUMVenda = objUserControl.Controls("SiglaUMVenda")
End Property

Public Property Get LblUMEstoque() As Object
     Set LblUMEstoque = objUserControl.Controls("LblUMEstoque")
End Property

Public Property Get NomeUMEstoque() As Object
     Set NomeUMEstoque = objUserControl.Controls("NomeUMEstoque")
End Property

Public Property Get LblUMCompras() As Object
     Set LblUMCompras = objUserControl.Controls("LblUMCompras")
End Property

Public Property Get NomeUMCompra() As Object
     Set NomeUMCompra = objUserControl.Controls("NomeUMCompra")
End Property

Public Property Get LblUMVendas() As Object
     Set LblUMVendas = objUserControl.Controls("LblUMVendas")
End Property

Public Property Get NomeUMVenda() As Object
     Set NomeUMVenda = objUserControl.Controls("NomeUMVenda")
End Property

Public Property Get ClasseUM() As Object
     Set ClasseUM = objUserControl.Controls("ClasseUM")
End Property

Public Property Get DescricaoClasseUM() As Object
     Set DescricaoClasseUM = objUserControl.Controls("DescricaoClasseUM")
End Property

Public Property Get LblClasseUM() As Object
     Set LblClasseUM = objUserControl.Controls("LblClasseUM")
End Property

Public Property Get DescricaoTabela() As Object
     Set DescricaoTabela = objUserControl.Controls("DescricaoTabela")
End Property

Public Property Get Tabela() As Object
     Set Tabela = objUserControl.Controls("Tabela")
End Property

Public Property Get BotaoTabelaPreco() As Object
     Set BotaoTabelaPreco = objUserControl.Controls("BotaoTabelaPreco")
End Property

Public Property Get ValorFilial() As Object
     Set ValorFilial = objUserControl.Controls("ValorFilial")
End Property

Public Property Get ValorEmpresa() As Object
     Set ValorEmpresa = objUserControl.Controls("ValorEmpresa")
End Property

Public Property Get DataPreco() As Object
     Set DataPreco = objUserControl.Controls("DataPreco")
End Property

Public Property Get GridTabelaPreco() As Object
     Set GridTabelaPreco = objUserControl.Controls("GridTabelaPreco")
End Property

Public Property Get DescrUM() As Object
     Set DescrUM = objUserControl.Controls("DescrUM")
End Property

Public Property Get BotaoEmbalagem() As Object
     Set BotaoEmbalagem = objUserControl.Controls("BotaoEmbalagem")
End Property

Public Property Get BotaoFornecedores() As Object
     Set BotaoFornecedores = objUserControl.Controls("BotaoFornecedores")
End Property

Public Property Get BotaoEstoque() As Object
     Set BotaoEstoque = objUserControl.Controls("BotaoEstoque")
End Property

Public Property Get BotaoCustos() As Object
     Set BotaoCustos = objUserControl.Controls("BotaoCustos")
End Property

Public Property Get BotaoControleEstoque() As Object
     Set BotaoControleEstoque = objUserControl.Controls("BotaoControleEstoque")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get FrmComissao() As Object
     Set FrmComissao = objUserControl.Controls("FrmComissao")
End Property

Public Property Get FrmFabricante() As Object
     Set FrmFabricante = objUserControl.Controls("FrmFabricante")
End Property

Public Property Get PercComissao() As Object
     Set PercComissao = objUserControl.Controls("PercComissao")
End Property

Public Property Get MetaComissao() As Object
     Set MetaComissao = objUserControl.Controls("MetaComissao")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoCorTonTP_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoCorTonTP_Click

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 178508

        If iPreenchido = PRODUTO_VAZIO Then gError 178509

        objProduto.sCodigo = sProduto

        Call Chama_Tela("CorTonTPint", objProduto)
        
    Else
    
        gError 178507
        
    End If
    
    Exit Sub
    
Erro_BotaoCorTonTP_Click:

    Select Case gErr

        Case 178507
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 178508

        Case 178509
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178510)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoCriarGrade_Click()

    Call Chama_Tela("Grade")

End Sub

'***ALTERACAO POR TULIO EM 28/05***
Public Sub BotaoProdutoCodBarras_Click()
    
Dim objProduto As New ClassProduto
Dim iProdPreenchido As Integer
Dim sProdutoFormatado As String
Dim lErro As Long

On Error GoTo Erro_BotaoProdutoCodBarras_Click

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProdutoFormatado, iProdPreenchido)
        If lErro <> SUCESSO Then gError 101702

        'copia para o obj o codigo ja formatado do produto
        objProduto.sCodigo = sProdutoFormatado
        
    End If
    
    'obtem da combo os codigos de barra
    Call ObterCodigosBarra(objProduto.colCodBarras)
    
    'chama a tela de produto de forma modal passando objproduto como parametro
    Call Chama_Tela_Modal("ProdutoCodBarras", objProduto)

    'recarrega a combo com os codigos de barras retornados da tela
    'em objProduto.colCodBarras
    Call Carrega_CodigoBarras_Produto(objProduto.colCodBarras)
    
    Exit Sub
    
Erro_BotaoProdutoCodBarras_Click:

    Select Case gErr
    
        Case 101702
        
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165461)
            
    End Select
    
    Exit Sub

End Sub
'***FIM ALTERACAO POR TULIO EM 28/05***

Public Sub Codigo_GotFocus()

Dim iProdPreenchido As Integer
Dim sProdutoFormatado As String
Dim lErro As Integer

On Error GoTo Erro_Codigo_GotFocus

    'copia para o produto anterior o produto da tela
    sProdutoAnterior = Codigo.Text
        
    Exit Sub
    
Erro_Codigo_GotFocus:

    Select Case gErr
        
        Case 101703
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165462)
            
    End Select
    
    Exit Sub

End Sub

Public Sub comboAliquota_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboAliquota_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

'''Private Sub KitVendaComp_Change()
'''
'''    iAlterado = REGISTRO_ALTERADO
'''
'''End Sub
'''
'''Private Sub KitVendaComp_Click()
'''
'''    iAlterado = REGISTRO_ALTERADO
'''
'''End Sub

Public Sub BotaoEmbalagem_Click()

'''05/09/01 Marcelo chamada para a tela ProdutoXEmbalagem
Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEmbalagem_Click

    'Verifica se existe alguma mudança e se deseja salvá-la
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código do Produto está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
    
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 93581
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objProduto = New ClassProduto
            objProduto.sCodigo = sProduto
        
        End If
        
    End If

    'Chama a Tela se Custos
    Call Chama_Tela("ProdutoEmbalagem", objProduto)

    Exit Sub

Erro_BotaoEmbalagem_Click:

    Select Case gErr

        Case 93581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165463)

    End Select

    Exit Sub

End Sub



Public Sub BotaoProcurar_Click()

On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog.CancelError = True
    ' Set flags
    CommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog.Filter = "All Files (*.*)|*.*|Bitmap Files" & _
    "(*.bmp)|*.bmp|Jpg Files (*.jpg)|*.jpg"
    ' Specify default filter
    CommonDialog.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog.ShowOpen

    ' Display name of selected file
     NomeFigura.Text = CommonDialog.FileName
     
     Call BotaoVisualizar_Click
     
    Exit Sub
    
 Call BotaoVisualizar_Click

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
    Exit Sub


End Sub

Public Sub BotaoVisualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVisualizar_Click

    'verifica se a figura foi preenchida
    If Len(Trim(NomeFigura.Text)) > 0 Then
    
        '?????? fazer um código muito melhor
        'verifica se o arquivo é do tipo imagem
        If GetAttr(NomeFigura.Text) = vbArchive Or GetAttr(NomeFigura.Text) = vbArchive + vbReadOnly Or GetAttr(NomeFigura.Text) = vbNormal Or GetAttr(NomeFigura.Text) = vbNormal + vbReadOnly Or GetAttr(NomeFigura.Text) = 8224 Then
            'coloca a figura na tela
            Figura.Picture = LoadPicture(NomeFigura.Text)
        Else
            gError 81211
            
        End If
    Else
        Figura.Picture = LoadPicture
    
    End If
    
    Exit Sub
    
Erro_BotaoVisualizar_Click:

    Select Case gErr
    
        Case 53
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_ENCONTRADO", gErr, NomeFigura.Text)
            
        Case 81210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
        
        Case 81211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_INVALIDO", gErr, NomeFigura.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165464)

    End Select

    Exit Sub

End Sub

'05/09/01 Marcelo

'Private Sub Embalagem_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim objEmbalagem As ClassEmbalagem
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_Embalagem_Validate
'
'    'verifica se o código da embalagem foi preenchido
'    If Len(Trim(Embalagem.Text)) > 0 Then
'
'        lErro = Valor_Positivo_Critica(Embalagem.Text)
'        If lErro <> SUCESSO Then gError 81200
'
'        Set objEmbalagem = New ClassEmbalagem
'
'        objEmbalagem.iCodigo = StrParaInt(Embalagem.Text)
'
'        'Le a embalagem com o código informado
'        '''30/08/01 Marcelo mudança da chamada da funcão
'        lErro = CF("Embalagem_Le", objEmbalagem)
'
'        If lErro <> SUCESSO And lErro <> 82763 Then gError 81201
'        If lErro = 82763 Then gError 81209
'
'        'Coloca a descrição na tela
'        DescricaoEmbalagem.Caption = objEmbalagem.sDescricao
'
'    Else
'        iTipoAlterado = 0
'
'    End If
'
'  Exit Sub
'
'Erro_Embalagem_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'            Case 81200, 81201
'
'            Case 81209
'                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_EMBALAGEM", objEmbalagem.iCodigo)
'
'                If vbMsgRes = vbYes Then
'                    Call Chama_Tela("Embalagem", objEmbalagem)
'
'                End If
'
'        Case 81203
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EMBALAGEM_NAO_ENCONTRADA", gErr, objEmbalagem.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165465)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'05/09/01 Marcelo

'Private Sub LabelEmbalagem_Click()
'Dim colSelecao As New Collection
'Dim objEmbalagem As New ClassEmbalagem
'Dim lErro As Long
'
'On Error GoTo Erro_LabelEmbalagem_Click
'
'    'verifica se o código da embalagem foi preenchido
'    If Len(Trim(Embalagem.Text)) > 0 Then objEmbalagem.iCodigo = StrParaInt(Embalagem.Text)
'
'    Call Chama_Tela("EmbalagensLista", colSelecao, objEmbalagem, objEventoEmbalagem)
'
'    Exit Sub
'
'Erro_LabelEmbalagem_Click:
'
'    Select Case gErr
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165466)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub NomeFigura_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 71929

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 71930

    'Mostra os dados do Produto na tela
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError 71931

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 71929, 71931

        Case 71930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165467)

    End Select

    Exit Sub

End Sub

Public Sub LabelCodigo_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o produto foi preenchido
    If Len(Codigo.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Codigo.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71928

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case 71928

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165468)

    End Select

    Exit Sub

End Sub

Public Sub LabelNomeReduzido_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNomeReduzido_Click

    objProduto.sNomeReduzido = NomeReduzido.Text

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelNomeReduzido_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165469)

    End Select

    Exit Sub

End Sub


Public Sub ConsideraQuantCotacaoAnterior_Click()

Dim lErro As Long

On Error GoTo Erro_ConsideraQuantCotacaoAnterior_Click

    iAlterado = REGISTRO_ALTERADO

    If ConsideraQuantCotacaoAnterior.Value = vbChecked Then

        'Habilita os controles
        PercentMaisQuantCotacaoAnterior.Enabled = False
        PercentMenosQuantCotacaoAnterior.Enabled = False

    Else

        'Desabilita os controles
        PercentMaisQuantCotacaoAnterior.Enabled = True
        PercentMenosQuantCotacaoAnterior.Enabled = True

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_ConsideraQuantCotacaoAnterior_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165470)

    End Select

    Exit Sub
        
End Sub

Public Sub CustoReposicao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CustoReposicao_LostFocus()

'Dim lErro As Long
'
'On Error GoTo Erro_CustoReposicao_LostFocus
'
'    If Len(Trim(CustoReposicao.ClipText)) > 0 Then
'
'        lErro = Valor_NaoNegativo_Critica(CustoReposicao.Text)
'        If lErro <> SUCESSO Then gError 64322
'
'    End If
'
'    Exit Sub
'
'Erro_CustoReposicao_LostFocus:
'
'    Select Case gErr
'
'        Case 64322
'            CustoReposicao.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165471)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub CustoReposicao_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_CustoReposicao_Validate

    If Len(Trim(CustoReposicao.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoReposicao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    Exit Sub
    
Erro_CustoReposicao_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165471)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoControleEstoque_Click()

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    'Verifica se o Produto já foi salvo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'Passa o codigo para o formato de BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 64323

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objProduto = New ClassProduto
    
            objProduto.sCodigo = sProduto
            
        End If
        
    End If

    'Chama a Tela de Estoque
    Call Chama_Tela("Estoque", objProduto)

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case Err

        Case 64323

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165472)

    End Select

    Exit Sub

End Sub

Public Sub Comprado_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If Comprado.Value = True Then
        ApropriacaoComp.Visible = True
        ApropriacaoComp.ListIndex = 0
        ApropriacaoProd.Visible = False
    End If

End Sub

Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ContaContabil_LostFocus()

'Dim lErro As Long
'Dim sContaFormatada As String
'Dim objPlanoConta As New ClassPlanoConta
'Dim vbMsgRes As VbMsgBoxResult
'Dim sContaMascarada As String
'
'On Error GoTo Erro_ContaContabil_LostFocus
'
'    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
'    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
'    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 64324
'
'    If lErro = SUCESSO Then
'
'        sContaFormatada = objPlanoConta.sConta
'
'        'mascara a conta
'        sContaMascarada = String(STRING_CONTA, 0)
'
'        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
'        If lErro <> SUCESSO Then Error 64325
'
'        ContaContabil.PromptInclude = False
'        ContaContabil.Text = sContaMascarada
'        ContaContabil.PromptInclude = True
'
'
'    'Se não encontrou a conta simples
'    ElseIf lErro = 44096 Or lErro = 44098 Then
'
'        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
'        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
'        If lErro <> SUCESSO And lErro <> 5700 Then Error 64326
'
'        'Conta não cadastrada
'        If lErro = 5700 Then Error 64327
'
'    End If
'
'    Exit Sub
'
'Erro_ContaContabil_LostFocus:
'
'    Select Case Err
'
'        Case 64324, 64326
'            ContaContabil.SetFocus
'
'        Case 64325
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
'            ContaContabil.SetFocus
'
'        Case 64327
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
'            ContaContabil.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165473)
'
'    End Select
'
'    Exit Sub
    
End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 64324

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 64325

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 64326

        'Conta não cadastrada
        If lErro = 5700 Then Error 64327

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case Err

        Case 64324, 64326
            'ContaContabil.SetFocus
    
        Case 64325
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            'ContaContabil.SetFocus
            
        Case 64327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaContabil.Text)
            'ContaContabil.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165473)
    
    End Select

    Exit Sub
    
End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 64328

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

        Case 64328
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165474)

    End Select

    Exit Sub
    
End Sub

Public Sub ContaProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ContaProducao_LostFocus()
'
'Dim lErro As Long
'Dim sContaFormatada As String
'Dim objPlanoConta As New ClassPlanoConta
'Dim vbMsgRes As VbMsgBoxResult
'Dim sContaMascarada As String
'
'On Error GoTo Erro_ContaProducao_LostFocus
'
'    'Verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
'    lErro = CF("ContaSimples_Critica_Modulo", ContaProducao.Text, ContaProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
'    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 64329
'
'    If lErro = SUCESSO Then
'
'        sContaFormatada = objPlanoConta.sConta
'
'        'Mascara a conta
'        sContaMascarada = String(STRING_CONTA, 0)
'
'        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
'        If lErro <> SUCESSO Then Error 64330
'
'        ContaProducao.PromptInclude = False
'        ContaProducao.Text = sContaMascarada
'        ContaProducao.PromptInclude = True
'
'    'Se não encontrou a conta simples
'    ElseIf lErro = 44096 Or lErro = 44098 Then
'
'        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
'        lErro = CF("Conta_Critica", ContaProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
'        If lErro <> SUCESSO And lErro <> 5700 Then Error 64331
'
'        'Conta não cadastrada
'        If lErro = 5700 Then Error 64332
'
'    End If
'
'    Exit Sub
'
'Erro_ContaProducao_LostFocus:
'
'    Select Case Err
'
'        Case 64329, 64331
'            ContaProducao.SetFocus
'
'        Case 64330
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
'            ContaProducao.SetFocus
'
'        Case 64332
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaProducao.Text)
'            ContaProducao.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165475)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub ContaProducao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaProducao_Validate

    'Verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaProducao.Text, ContaProducao.ClipText, objPlanoConta, MODULO_ESTOQUE)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 64329

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'Mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then Error 64330

        ContaProducao.PromptInclude = False
        ContaProducao.Text = sContaMascarada
        ContaProducao.PromptInclude = True

    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaProducao.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 5700 Then Error 64331

        'Conta não cadastrada
        If lErro = 5700 Then Error 64332

    End If

    Exit Sub

Erro_ContaProducao_Validate:

    Cancel = True

    Select Case Err

        Case 64329, 64331
            'ContaProducao.SetFocus
    
        Case 64330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            'ContaProducao.SetFocus
            
        Case 64332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, ContaProducao.Text)
            'ContaProducao.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165475)
    
    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim lErro As Long
Dim sMascaraConta As String
Dim objTela As Object

On Error GoTo Erro_Form_Load

    If gobjMAT.iProdFatorAjuste <> 0 Then
        LabelFatorAjuste.Visible = True
        FatorAjuste.Visible = True
    End If

    Referencia.MaxLength = STRING_PRODUTO_REFERENCIA
    NomeReduzido.MaxLength = STRING_PRODUTO_NOME_REDUZIDO
    Descricao.MaxLength = STRING_PRODUTO_DESCRICAO_TELA
    Modelo.MaxLength = STRING_PRODUTO_MODELO
    'Inicializa o Frame atual com 1
    iFrameAtual = 1

    Set objEventoClasseUM = New AdmEvento
    Set objEventoTipoDeProduto = New AdmEvento
    Set objEventoProdutoSubst1 = New AdmEvento
    Set objEventoProdutoSubst2 = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoContaProducao = New AdmEvento
    Set objEventoProduto = New AdmEvento
    'Set objEventoEmbalagem = New AdmEvento 05/09/01 Marcelo
    Set objEventoTabelaPrecoItem = New AdmEvento
    Set objEventoClasFiscIPI = New AdmEvento
    Set objEventoGenero = New AdmEvento
    Set objEventoISSQN = New AdmEvento
    Set objEventoCodServNFe = New AdmEvento
    Set objEventoCEST = New AdmEvento
    
    Set gobjIdiomaTela = New ClassIdiomaTela
    Call gobjIdiomaTela.Iniciar(Me.Name)

    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 64333

    ContaContabil.Mask = sMascaraConta
    
    ContaProducao.Mask = sMascaraConta
    
    If gobjEST.iProdExibeComissao = MARCADO Then
        FrmComissao.Visible = True
        FrmFabricante.Visible = False
    End If

    'Inicialiaza o Grid de Categoria
    Set objGridCategoria = New AdmGrid
    
    lErro = Inicializa_Grid_Categoria(objGridCategoria)
    If lErro <> SUCESSO Then gError 64334
    
    'Inicializa o Grid TabelaPreco
    Set objGridTabelaPreco = New AdmGrid

    lErro = Inicializa_Grid_TabelaPreco(objGridTabelaPreco)
    If lErro <> SUCESSO Then gError 64335
    
    'Carrega a combobox de Categoria Produto
    lErro = Carrega_ComboCategoriaProduto()
    If lErro <> SUCESSO Then gError 64336
    
    'Carrega a combobox de Situação Tributária
    lErro = Carrega_ComboSituacaoTributaria
    If lErro <> SUCESSO Then gError 81193
    
    lErro = Carrega_ComboGrade()
    If lErro <> SUCESSO Then gError 86359
    
    lErro = Carrega_Fabricante(Fabricante)
    If lErro <> SUCESSO Then gError 86359
    
    CodigoBarras.Clear

    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Codigo)
    If lErro <> SUCESSO Then gError 64338

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Substituto1)
    If lErro <> SUCESSO Then gError 64339

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Substituto2)
    If lErro <> SUCESSO Then gError 64340
    
    'desabilta gerencial se a mascara de produto contiver apenas um nivel
    Call DesabilitaGerencialUmNivel
    
    ApropriacaoComp.ListIndex = 0

    CustoReposicao.Format = FORMATO_CUSTO
    
'    'Verifica se o Módulo de Compras está inativo
'    If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then
'
'        'Desabilita o TabCompras
'        Frame2(7).Enabled = False
'
'    End If
    
    OrigemMercadoria.ListIndex = 0
    
    'Verifica se é Versao Light
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        Rastro.left = POSICAO_FORA_TELA
'        Rastro.TabStop = False
'        LabelRastro.left = POSICAO_FORA_TELA
'        LabelRastro.Visible = False
'        HorasMaquina.left = POSICAO_FORA_TELA
'        HorasMaquina.TabStop = False
'        LabelHorasMaq.left = False
'        LabelHorasMaq.Visible = False
'        LabelMinutos.left = False
'        LabelMinutos.Visible = False
'        PesoEspecifico.left = POSICAO_FORA_TELA
'        PesoEspecifico.TabStop = False
'        LabelPesoEspecifico.left = False
'        LabelPesoEspecifico.Visible = False
'        LabelPesoEspKg.left = False
'        LabelPesoEspKg.Visible = False
'
'    End If
    
    Set objTela = Me
    
    lErro = CF("Produto_Customiza_Form_Load", objTela)
    If lErro <> SUCESSO Then gError 178505
    
    Rastro.ListIndex = 0
    
    lErro = CF2(Me, "Prod_Form_Load")
    If lErro <> SUCESSO Then gError 178505
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 64333, 64334, 64335, 64336, 64337, 64338, 64339, 64340, 81193, 86359, 178505

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165476)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Categoria(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Categoria

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Categoria")
    objGridInt.colColuna.Add ("Item")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ComboCategoriaProduto.Name)
    objGridInt.colCampo.Add (ComboCategoriaProdutoItem.Name)

    'Colunas do Grid
    iGrid_Categoria_Col = 1
    iGrid_Valor_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCategoria

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridCategoria.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Categoria = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_TabelaPreco(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tabela")
    objGridInt.colColuna.Add ("Descrição Tabela")
    objGridInt.colColuna.Add ("Valor Empresa")
    objGridInt.colColuna.Add ("Valor Filial")
    objGridInt.colColuna.Add ("A Partir De")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Tabela.Name)
    objGridInt.colCampo.Add (DescricaoTabela.Name)
    objGridInt.colCampo.Add (ValorEmpresa.Name)
    objGridInt.colCampo.Add (ValorFilial.Name)
    objGridInt.colCampo.Add (DataPreco.Name)

    If giFilialEmpresa = EMPRESA_TODA Then
        ValorEmpresa.Enabled = True
        ValorFilial.Enabled = False
    Else
        ValorEmpresa.Enabled = False
        ValorFilial.Enabled = True
    End If

    'Colunas do Grid
    iGrid_Tabela_Col = 1
    iGrid_DescricaoTabela_Col = 2
    iGrid_ValorEmpresa_Col = 3
    iGrid_ValorFilial_Col = 4
    iGrid_DataPreco_Col = 5

    'Grid do GridInterno
    objGridInt.objGrid = GridTabelaPreco
    
    objGridInt.objGrid.Rows = 101
           
    'Todas as linhas do grid
    'objGridInt.objGrid.Rows = 30

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 12

    'Largura da primeira coluna
    GridTabelaPreco.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Impede a inclusao de linhas no grid por parte do usuario
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_TabelaPreco = SUCESSO

    Exit Function

End Function

Private Function Carrega_ComboCategoriaProduto() As Long
'Carrega as Categorias na Combobox

Dim lErro As Long
Dim colCategorias As New Collection
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Carrega_ComboCategoriaProduto

    'Lê o código e a descrição de todas as categorias
    lErro = CF("CategoriasProduto_Le_Todas", colCategorias)
    If lErro <> SUCESSO And lErro <> 22542 Then Error 64341

    For Each objCategoriaProduto In colCategorias

        'Insere na combo CategoriaProduto
        ComboCategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    Carrega_ComboCategoriaProduto = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProduto:

    Carrega_ComboCategoriaProduto = Err

    Select Case Err

        Case 64341

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165477)

    End Select

    Exit Function

End Function

Public Sub AliquotaIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AliquotaIPI_LostFocus()

'Dim lErro As Long
'Dim dAliquota As Double
'
'On Error GoTo Erro_AliquotaIPI_LostFocus
'
'    'Verifica se esta preenchida
'    If Len(Trim(AliquotaIPI.Text)) = 0 Then Exit Sub
'
'    'Critica
'    lErro = Porcentagem_Critica(AliquotaIPI.Text)
'    If lErro <> SUCESSO Then Error 64342
'
'    If CDbl(AliquotaIPI.Text) = 100# Then Error 64343
'
'    Exit Sub
'
'Erro_AliquotaIPI_LostFocus:
'
'    Select Case Err
'
'        Case 64342
'            AliquotaIPI.SetFocus
'
'        Case 64343
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_ALIQUOTA_INVALIDO", Err)
'            AliquotaIPI.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165478)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub AliquotaIPI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dAliquota As Double

On Error GoTo Erro_AliquotaIPI_Validate

    'Verifica se esta preenchida
    If Len(Trim(AliquotaIPI.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(AliquotaIPI.Text)
    If lErro <> SUCESSO Then Error 64342

    If CDbl(AliquotaIPI.Text) = 100# Then Error 64343

    Exit Sub

Erro_AliquotaIPI_Validate:

    Cancel = True

    Select Case Err

        Case 64342
            'AliquotaIPI.SetFocus

        Case 64343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_ALIQUOTA_INVALIDO", Err)
            'AliquotaIPI.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165478)

    End Select

    Exit Sub

End Sub

Public Sub Ativo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim iTemFilho As Integer
'iAviso por sAviso - Alteracao Daniel - 25/09/2001
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) = 0 Then Error 64344

    'Passa o código para o formato do BD
    lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then Error 64345

    If iPreenchido = PRODUTO_PREENCHIDO Then
        
        objProduto.sCodigo = sProduto
    
        'Lê o Produto com o Código passado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 64346
        
        'Se não encontrar --> Erro
        If lErro = 28030 Then Error 64347
    
        'Verifica se o Produto tem Filhos
        lErro = CF("Produto_Tem_Filho", sProduto, iTemFilho)
        If lErro <> SUCESSO Then Error 64348
    
        'Se tiver
        If iTemFilho = True Then
            'carrega sAviso  com a mensagem de um Produto Gerencial com filhos
            sAviso = "AVISO_EXCLUSAO_PRODUTO_GERENCIAL_COM_FILHOS"
        'senão tiver
        Else
            'Verifica se é Gerencial ou Final
            'iAviso por sAviso - Alteracao Daniel - 25/09/2001
            If objProduto.iGerencial = GERENCIAL Then
                sAviso = "AVISO_EXCLUSAO_PRODUTO_GERENCIAL"
            Else
                sAviso = "AVISO_EXCLUSAO_PRODUTO_FINAL"
            End If
        End If
            
        'Pergunta se o usuário confirma a exclusão
        'iAviso por sAviso - Alteracao Daniel - 25/09/2001
        vbMsgRes = Rotina_Aviso(vbYesNo, sAviso)
            
        'se o usuário confirmar a exclusão
        If vbMsgRes = vbYes Then
            
            'Exclui o Produto
            lErro = CF("Produto_Exclui", objProduto)
            If lErro <> SUCESSO Then Error 64349
        
'            Call Exclui_Arvore_Produto(TvwProduto.Nodes, objProduto)
    
            'Fecha comando de setas se estiver aberto
            lErro = ComandoSeta_Fechar(Me.Name)
        
            Call Limpa_Tela_Produto
    
        End If
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
        
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 64344
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 64345, 64346, 64348, 64349

        Case 64347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165479)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 64350

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    'Limpa a Tela
    Call Limpa_Tela_Produto

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 64350

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165480)

    End Select

    Exit Sub

End Sub

Private Function Carrega_ComboSituacaoTributaria() As Long
'Carrega as situações tributárias na Combobox
Dim lErro As Long
Dim colTiposTribICMS As New Collection
Dim objTiposTribICMS As New ClassTipoTribICMS

On Error GoTo Erro_Carrega_ComboSituacaoTributaria

    'Lê o tipo e a descrição da tabela TiposTribICMS
    lErro = TiposTribICMSLoja_Le(colTiposTribICMS)
    If lErro <> SUCESSO Then gError 81194
  
    For Each objTiposTribICMS In colTiposTribICMS

       'Insere na combo Situação tributária
       SituacaoTributaria.AddItem "ICMS " & objTiposTribICMS.sDescricao
       SituacaoTributaria.ItemData(SituacaoTributaria.NewIndex) = objTiposTribICMS.iTipo
       
    Next

    SituacaoTributaria.AddItem TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO_DESC
    SituacaoTributaria.ItemData(SituacaoTributaria.NewIndex) = TIPOTRIBISS_NAO_TRIBUTADO

    SituacaoTributaria.AddItem TIPOTRIBISS_SITUACAOTRIBECF_ISENTA_DESC
    SituacaoTributaria.ItemData(SituacaoTributaria.NewIndex) = TIPOTRIBISS_ISENTA

    SituacaoTributaria.AddItem TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST_DESC
    SituacaoTributaria.ItemData(SituacaoTributaria.NewIndex) = TIPOTRIBISS_TRIB_SUBST

    SituacaoTributaria.AddItem TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL_DESC
    SituacaoTributaria.ItemData(SituacaoTributaria.NewIndex) = TIPOTRIBISS_INTEGRAL

    Carrega_ComboSituacaoTributaria = SUCESSO

    Exit Function

Erro_Carrega_ComboSituacaoTributaria:

    Carrega_ComboSituacaoTributaria = gErr

    Select Case gErr

        Case 81194

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165481)

    End Select

    Exit Function

End Function

Private Function Carrega_ComboAliquota() As Long
'Carrega as Aliquotas na Combobox

Dim lErro As Long
Dim colAliquotaICMS As New Collection
Dim objAliquotaICMS As New ClassAliquotaICMS

On Error GoTo Erro_Carrega_ComboAliquota

    comboAliquota.Clear

    If SituacaoTributaria.ListIndex <> -1 Then

        'Inicializando combo AliquotaICMS com
        'as aliquotas lidas da tabela
        lErro = CF("AliquotaICMS_Le_Todas", colAliquotaICMS)
        If lErro <> SUCESSO Then gError 98388
        
        For Each objAliquotaICMS In colAliquotaICMS
        
            If objAliquotaICMS.iISS = 1 And _
            SituacaoTributaria.ItemData(SituacaoTributaria.ListIndex) = TIPOTRIBISS_INTEGRAL Then
            
                comboAliquota.AddItem objAliquotaICMS.sSigla & SEPARADOR & FormatPercent(objAliquotaICMS.dAliquota)
                
            ElseIf objAliquotaICMS.iISS = 0 And _
            SituacaoTributaria.ItemData(SituacaoTributaria.ListIndex) = TIPOTRIBICMS_INTEGRAL Then
            
                comboAliquota.AddItem objAliquotaICMS.sSigla & SEPARADOR & FormatPercent(objAliquotaICMS.dAliquota)
                
            End If
        Next

    End If

    Carrega_ComboAliquota = SUCESSO

    Exit Function

Erro_Carrega_ComboAliquota:

    Carrega_ComboAliquota = gErr

    Select Case gErr

        Case 98388

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165482)

    End Select

    Exit Function

End Function

Public Sub ClasFiscIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ClasFiscIPI_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ClasFiscIPI, iAlterado)
    
End Sub

Public Sub ClasseUM_Change()

    iAlterado = REGISTRO_ALTERADO
    iClasseUMAlterada = REGISTRO_ALTERADO

End Sub

Public Sub ClasseUM_GotFocus()

Dim iClasseUMAux As Integer

    iClasseUMAux = iClasseUMAlterada
    Call MaskEdBox_TrataGotFocus(ClasseUM, iAlterado)
    iClasseUMAlterada = iClasseUMAux
    
End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = 1
    
End Sub

Public Sub CodigoBarras_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodigoIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Comprimento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Comprimento_LostFocus()

'Dim lErro As Long
'
'On Error GoTo Erro_Comprimento_LostFocus
'
'    'Verifica se foi preenchido
'    If Len(Trim(Comprimento.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(Comprimento.Text)
'    If lErro <> SUCESSO Then Error 64351
'
'    'Coloca na Tela o valor já formatado
'    Comprimento.Text = Format(Comprimento.Text, Comprimento.Format)
'
'    Exit Sub
'
'Erro_Comprimento_LostFocus:
'
'    Select Case Err
'
'        Case 64351
'            Comprimento.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165483)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub Comprimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Comprimento_Validate

    'Verifica se foi preenchido
    If Len(Trim(Comprimento.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Comprimento.Text)
    If lErro <> SUCESSO Then Error 64351

    'Coloca na Tela o valor já formatado
    Comprimento.Text = Format(Comprimento.Text, Comprimento.Format)

    Exit Sub

Erro_Comprimento_Validate:

    Cancel = True

    Select Case Err

        Case 64351
            'Comprimento.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165483)

    End Select

    Exit Sub

End Sub

Public Sub Descricao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Espessura_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Espessura_LostFocus()

'Dim lErro As Long
'
'On Error GoTo Erro_Espessura_LostFocus
'
'    'Verifica se está preenchido
'    If Len(Trim(Espessura.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(Espessura.Text)
'    If lErro <> SUCESSO Then Error 64352
'
'    'Coloca o valor na tela já formatado
'    Espessura.Text = Format(Espessura.Text, Espessura.Format)
'
'    Exit Sub
'
'Erro_Espessura_LostFocus:
'
'    Select Case Err
'
'        Case 64352
'            Espessura.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165484)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub Espessura_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Espessura_Validate

    'Verifica se está preenchido
    If Len(Trim(Espessura.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Espessura.Text)
    If lErro <> SUCESSO Then Error 64352

    'Coloca o valor na tela já formatado
    Espessura.Text = Format(Espessura.Text, Espessura.Format)

    Exit Sub

Erro_Espessura_Validate:

    Cancel = True

    Select Case Err

        Case 64352
            'Espessura.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165484)

    End Select

    Exit Sub

End Sub

Public Sub EtiquetasCodBarras_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EtiquetasCodBarras_GotFocus()

    Call MaskEdBox_TrataGotFocus(EtiquetasCodBarras, iAlterado)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set gobjAnotacao = Nothing
    Set gobjIdiomaTela = Nothing
    
    Set objEventoTabelaPrecoItem = Nothing
    Set objEventoTipoDeProduto = Nothing
    Set objEventoClasseUM = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoProdutoSubst1 = Nothing
    Set objEventoProdutoSubst2 = Nothing
    Set objEventoContaContabil = Nothing
    Set objEventoContaProducao = Nothing
    Set objEventoProduto = Nothing
    Set objEventoClasFiscIPI = Nothing
    Set objEventoGenero = Nothing
    Set objEventoISSQN = Nothing
    Set objEventoCodServNFe = Nothing
    Set objEventoCEST = Nothing
    
    Set objGridCategoria = Nothing
    Set objGridTabelaPreco = Nothing

    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Call CF2(Me, "Prod_Form_UnLoad")

    Set gobjInfoUsu = Nothing

End Sub

Public Sub HorasMaquina_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub HorasMaquina_GotFocus()

    Call MaskEdBox_TrataGotFocus(HorasMaquina, iAlterado)
    
End Sub

Public Sub IncideIPI_Click()

    'Verifica se está selecionado
    If IncideIPI.Value = vbUnchecked Then
        'Senão estiver desativa e limpa os campos ligados ao IPI
        AliquotaIPI.Text = ""
        AliquotaIPI.Enabled = False
        CodigoIPI.Text = ""
        CodigoIPI.Enabled = False
        
    'Se estiver selecionado
    Else
        'tiva os campos relacionados ao IPI
        AliquotaIPI.Enabled = True
        CodigoIPI.Enabled = True
        
    End If

    Exit Sub

End Sub

Public Sub LabelContaProducao_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_LabelContaProducao_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaProducao.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 64353

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaProducao)

    Exit Sub

Erro_LabelContaProducao_Click:

    Select Case Err

        Case 64353
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165485)

    End Select

    Exit Sub

End Sub

Public Sub Largura_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Largura_LostFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Largura_LostFocus
'
'    'Verifica se foi preenchida
'    If Len(Trim(Largura.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(Largura.Text)
'    If lErro <> SUCESSO Then Error 64354
'
'    'Coloca o valor já formatado na Tela
'    Largura.Text = Format(Largura.Text, Largura.Format)
'
'    Exit Sub
'
'Erro_Largura_LostFocus:
'
'    Select Case Err
'
'        Case 64354
'            Largura.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165486)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub Largura_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Largura_Validate

    'Verifica se foi preenchida
    If Len(Trim(Largura.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(Largura.Text)
    If lErro <> SUCESSO Then Error 64354

    'Coloca o valor já formatado na Tela
    Largura.Text = Format(Largura.Text, Largura.Format)

    Exit Sub

Erro_Largura_Validate:

    Cancel = True

    Select Case Err

        Case 64354
           ' Largura.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165486)

    End Select

    Exit Sub

End Sub

Public Sub LblSubst1_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LblSubst1_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Substituto1.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Substituto1.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71932

        objProduto.sCodigo = sProdutoFormatado

    End If

    'Chama a tela de browse se Produto
    Call Chama_Tela("ProdutosSubstLista", colSelecao, objProduto, objEventoProdutoSubst1)

    Exit Sub
    
Erro_LblSubst1_Click:

    Select Case gErr

        Case 71932

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165487)

    End Select

    Exit Sub


End Sub

Public Sub LblSubst2_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_LblSubst2_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Substituto2.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Substituto2.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 71933

        objProduto.sCodigo = sProdutoFormatado

    End If

    'Chama a tela de browse se Produto
    Call Chama_Tela("ProdutosSubstLista", colSelecao, objProduto, objEventoProdutoSubst2)

    Exit Sub
    
Erro_LblSubst2_Click:

    Select Case gErr

        Case 71933

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165488)

    End Select

    Exit Sub

End Sub

Public Sub ListaCaracteristicas_Click()

    If ListaCaracteristicas.Selected(2) = True Or ListaCaracteristicas.Selected(3) = True Or Produzido.Value = True Then
        ListaCaracteristicas.Selected(1) = True
    End If
    
    'Se for Kit de Venda é Faturável
    If ListaCaracteristicas.Selected(4) Then
        ListaCaracteristicas.Selected(0) = True
    End If
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Modelo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NaoTemFaixaReceb_Click()

Dim lErro As Long

On Error GoTo Erro_NaoTemFaixaReceb_Click

    'Verifica valor na checkbox
    If NaoTemFaixaReceb.Value = False Then

        'Habilita os controles
        PercentMaisReceb.Enabled = True
        PercentMenosReceb.Enabled = True
        RecebForaFaixa(0).Enabled = True
        RecebForaFaixa(1).Enabled = True

    Else

        'Desabilita os controles
        PercentMaisReceb.Enabled = False
        PercentMenosReceb.Enabled = False
        RecebForaFaixa(0).Enabled = False
        RecebForaFaixa(1).Enabled = False

    End If
        
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_NaoTemFaixaReceb_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165489)

    End Select

    Exit Sub

End Sub

Public Sub NaturezaProduto_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub NivelFinal_LostFocus()

'Dim lErro As Long
'Dim sProduto As String
'Dim iPreenchido As Integer
'Dim iTemFilho As Integer
'
'On Error GoTo Erro_NivelFinal_LostFocus
'
'    'Verifica se o Código está preenchido
'    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
'
'    'Passa para o formato do BD
'    lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
'    If lErro <> SUCESSO Then Error 64355
'
'    If iPreenchido <> PRODUTO_PREENCHIDO Then Exit Sub
'
'    'Verifica se o Produto tem filhos
'    lErro = CF("Produto_Tem_Filho", sProduto, iTemFilho)
'    If lErro <> SUCESSO Then Error 64356
'
'    'Se tiver erro
'    If iTemFilho = True Then Error 64357
'
'    Exit Sub
'
'Erro_NivelFinal_LostFocus:
'
'    Select Case Err
'
'        Case 64355, 64356
'
'        Case 64357
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FINAL_COM_FILHOS", Err, Codigo.Text)
'            'NivelGerencial.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165490)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub NivelGerencial_Click()

    iAlterado = REGISTRO_ALTERADO
    
    LabelGrade.Enabled = True
    Grades.Enabled = True
    
End Sub

Public Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        'Comentado por Wagner 07/03/2007
        'A Harmonia usa o nome reduzido numérico
        'If Not IniciaLetra(NomeReduzido.Text) Then Error 64358

    End If
    
    Exit Sub

Erro_NomeReduzido_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 64358
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165491)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoClasseUM_evSelecao(obj1 As Object)

Dim objClasseUM As New ClassClasseUM
Dim bCancel As Boolean

    Set objClasseUM = obj1

    'Preenche Text da ClasseUM
    ClasseUM.Text = CStr(objClasseUM.iClasse)
    Call ClasseUM_Validate(bCancel)

    Me.Show

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabil.Text = ""
    Else
        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 64359

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 64359
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165492)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaProducao_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaProducao_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaProducao.Text = ""
    Else
        ContaProducao.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 64360

        ContaProducao.Text = sContaEnxuta
        ContaProducao.PromptInclude = True
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaProducao_evSelecao:

    Select Case Err

        Case 64360
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165493)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoSubst1_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProdutoSubst1_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 64361

    'Coloca na tela o Produto selecionado
    Substituto1.PromptInclude = False
    Substituto1.Text = sProduto
    Substituto1.PromptInclude = True
    DescSubst1.Caption = objProduto.sDescricao
    Substituto1.SetFocus

    Me.Show

    Exit Sub

Erro_objEventoProdutoSubst1_evSelecao:

    Select Case Err

        Case 64361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165494)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoSubst2_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProdutoSubst2_evSelecao

    Set objProduto = obj1

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then Error 64362

    'Coloca na tela o produto selecionado
    Substituto2.PromptInclude = False
    Substituto2.Text = sProduto
    Substituto2.PromptInclude = True
    DescSubst2.Caption = objProduto.sDescricao
    Substituto2.SetFocus

    Me.Show

    Exit Sub

Erro_objEventoProdutoSubst2_evSelecao:

    Select Case Err

        Case 64362
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165495)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTipoDeProduto_evSelecao(obj1 As Object)

Dim objTipoProduto As ClassTipoDeProduto
Dim bCancel As Boolean

    Set objTipoProduto = obj1

    'coloca na tela o Tipo de Produto Selecionado e dispara o evento LostFocus
    TipoProduto.Text = objTipoProduto.iTipo
    Call TipoProduto_Validate(bCancel)

    Me.Show

End Sub

Public Sub Opcao_Click()

    If Opcao.SelectedItem.Index <> iFrameAtual Then
        
        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
                
'        If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then
'
'            If Opcao.SelectedItem.Index = 7 Then
'
'                Frame2(7).Enabled = False
'                Frame2(iFrameAtual).Visible = True
'                Opcao.Tabs.Item(iFrameAtual).Selected = True
'                Exit Sub
'            End If
'        End If

        'Esconde o frame atual, mostra o novo
        Frame2(Opcao.SelectedItem.Index).Visible = True
        Frame2(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_PRODUTO_DADOS_PRINCIPAIS
            
            Case TAB_Categoria
                Parent.HelpContextID = IDH_PRODUTO_CATEGORIA
            
            Case TAB_Complemento
                Parent.HelpContextID = IDH_PRODUTO_COMPLEMENTO
            
            Case TAB_CaracFisicas
                Parent.HelpContextID = IDH_PRODUTO_CARACTERISTICAS_FISICAS
            
            Case TAB_Precos
                Parent.HelpContextID = IDH_PRODUTO_PRECOS
        
            Case TAB_UM
                Parent.HelpContextID = IDH_PRODUTO_UNIDADES_MEDIDA
            
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_PRODUTO_TRIBUTACAO
    
        End Select
        
    End If

End Sub

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 64363
        
        If lErro = SUCESSO Then
            
            lErro = Traz_Produto_Tela(objProduto)
            If lErro <> SUCESSO Then Error 64364
            
        Else
            Codigo.PromptInclude = False
            Codigo.Text = objProduto.sCodigo
            Codigo.PromptInclude = True
        End If
    Else
        Call Limpa_Tela_Produto
    End If
    
    iProdutoAlterado = 0
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 64363, 64364
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165496)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub BotaoCustos_Click()

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoCustos_Click

    'Verifica se existe alguma mudança e se deseja salvá-la
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código do Produto está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
    
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 64365
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objProduto = New ClassProduto
            objProduto.sCodigo = sProduto
        
        End If
        
    End If

    'Chama a Tela se Custos
    Call Chama_Tela("Custos", objProduto)

    Exit Sub

Erro_BotaoCustos_Click:

    Select Case Err

        Case 64365

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165497)

    End Select

    Exit Sub

End Sub

Public Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As ClassEstoqueProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    'Verifica se o Produto já foi salvo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'Passa o codigo para o formato de BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 64366

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objEstoqueProduto = New ClassEstoqueProduto
    
            'Carrega em objEstoqueProduto
            objEstoqueProduto.sProduto = sProduto
            
        End If
        
    End If

    'Chama a Tela de EstoqueProduto
    Call Chama_Tela("EstoqueProduto", objEstoqueProduto)

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case Err

        Case 64366

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165498)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFornecedores_Click()

Dim lErro As Long
Dim objFornecedorProdutoFF As ClassFornecedorProdutoFF
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoFornecedores_Click

    'Verifica se o produto foi salvo
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código do Produto está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then

        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 64367

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            Set objFornecedorProdutoFF = New ClassFornecedorProdutoFF
    
            'Carrega o Produto em  objFornecedorProduto
            objFornecedorProdutoFF.sProduto = sProduto
            
        End If
        
    End If

    'Chama a tela FornecedorProdutoFF
    Call Chama_Tela("FornFilialProduto", objFornecedorProdutoFF)

    Exit Sub

Erro_BotaoFornecedores_Click:

    Select Case Err

        Case 64367

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165499)

    End Select

    Exit Sub

End Sub

Public Sub BotaoTabelaPreco_Click()

Dim colSelecao As New Collection
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem

    'Chama a Tela TabelaPrecoItemLista para consulta
    Call Chama_Tela("TabelaPrecoItemLista", colSelecao, objTabelaPrecoItem, objEventoTabelaPrecoItem)

End Sub

Public Sub OrigemMercadoria_Click()

    iAlterado = REGISTRO_ALTERADO
    OrigemMercadoria.ToolTipText = OrigemMercadoria.Text
    
End Sub

Public Sub INSSPercBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub INSSPercBase_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_INSSPercBase_Validate

    'Verifica se esta preenchida
    If Len(Trim(INSSPercBase.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(INSSPercBase.Text)
    If lErro <> SUCESSO Then gError 89105

    Exit Sub

Erro_INSSPercBase_Validate:

    Cancel = True

    Select Case gErr

        Case 89105

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165500)

    End Select

    Exit Sub

End Sub

Public Sub PesoBruto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoBruto_LostFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_PesoBruto_LostFocus
'
'    'Verifica se está preenchido
'    If Len(Trim(PesoBruto.Text)) = 0 Then Exit Sub
'
'    'Criticao valor
'    lErro = Valor_Positivo_Critica(PesoBruto.Text)
'    If lErro <> SUCESSO Then Error 64368
'
'    'Coloca o valor formatado na Tela
'    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format) 'Alterado por Wagner
'
'    Exit Sub
'
'Erro_PesoBruto_LostFocus:
'
'    Select Case Err
'
'        Case 64368
'            PesoBruto.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165501)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoBruto_Validate

    'Verifica se está preenchido
    If Len(Trim(PesoBruto.Text)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_Positivo_Critica(PesoBruto.Text)
    If lErro <> SUCESSO Then Error 64368

    'Coloca o valor formatado na Tela
    PesoBruto.Text = Format(PesoBruto.Text, PesoBruto.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoBruto_Validate:

    Cancel = True

    Select Case Err

        Case 64368
            'PesoBruto.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165501)

    End Select

    Exit Sub

End Sub

Public Sub PesoEspecifico_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub PesoEspecifico_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoEspecifico_Validate

    'Verifica se está preenchido
    If Len(Trim(PesoEspecifico.Text)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_Positivo_Critica(PesoEspecifico.Text)
    If lErro <> SUCESSO Then gError 76344

    'Coloca o valor formatado na Tela
    PesoEspecifico.Text = Format(PesoEspecifico.Text, PesoEspecifico.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoEspecifico_Validate:

    Cancel = True
    
    Select Case gErr

        Case 76344
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165502)

    End Select

    Exit Sub

End Sub

Public Sub PesoLiquido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PesoLiquido_LostFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_PesoLiquido_LostFocus
'
'    'Verifica se foi preenchido
'    If Len(Trim(PesoLiquido.Text)) = 0 Then Exit Sub
'
'    'Critica o valor
'    lErro = Valor_Positivo_Critica(PesoLiquido.Text)
'    If lErro <> SUCESSO Then Error 64369
'
'    'Coloca o valor formatado na Tela
'    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner
'
'    Exit Sub
'
'Erro_PesoLiquido_LostFocus:
'
'    Select Case Err
'
'        Case 64369
'            PesoLiquido.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165503)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub PesoLiquido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoLiquido_Validate

    'Verifica se foi preenchido
    If Len(Trim(PesoLiquido.Text)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(PesoLiquido.Text)
    If lErro <> SUCESSO Then Error 64369

    'Coloca o valor formatado na Tela
    PesoLiquido.Text = Format(PesoLiquido.Text, PesoLiquido.Format) 'Alterado por Wagner

    Exit Sub

Erro_PesoLiquido_Validate:

    Cancel = True

    Select Case Err

        Case 64369
            'PesoLiquido.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165503)

    End Select

    Exit Sub

End Sub

Public Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrazoValidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrazoValidade, iAlterado)
    
End Sub

Public Sub Produzido_Click()

    iAlterado = REGISTRO_ALTERADO

    If Produzido.Value = True Then
        ListaCaracteristicas.Selected(1) = True
        ApropriacaoProd.ListIndex = 1
        ApropriacaoProd.Visible = True
        ApropriacaoComp.Visible = False
    End If

End Sub

Public Sub Referencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Residuo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Residuo_LostFocus()
'
'Dim lErro As Long
'
'On Error GoTo Erro_Residuo_LostFocus
'
'    'Se estiver preenchido
'    If Len(Trim(Residuo.ClipText)) > 0 Then
'
'        'Faz a crítica do Resíduo
'        lErro = Porcentagem_Critica(Residuo.Text)
'        If lErro <> SUCESSO Then Error 64370
'
'    End If
'
'    Exit Sub
'
'Erro_Residuo_LostFocus:
'
'    Select Case Err
'
'        Case 64370
'            Residuo.SetFocus
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165504)
'
'    End Select
'
'    Exit Sub

End Sub

Public Sub Residuo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Residuo_Validate

    'Se estiver preenchido
    If Len(Trim(Residuo.ClipText)) > 0 Then

        'Faz a crítica do Resíduo
        lErro = Porcentagem_Critica(Residuo.Text)
        If lErro <> SUCESSO Then Error 64370

    End If

    Exit Sub

Erro_Residuo_Validate:

    Cancel = True

    Select Case Err

        Case 64370
            'Residuo.SetFocus
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165504)

    End Select

    Exit Sub

End Sub

Public Sub SiglaUMCompra_Click()

Dim lErro As Long
Dim sSiglaUMCompra As String

On Error GoTo Erro_SiglaUMCompra_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMCompra.ListIndex = -1 Then Exit Sub
    
    sSiglaUMCompra = SiglaUMCompra.Text
        
    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMCompra, "NomeUMCompra")
    If lErro <> SUCESSO Then Error 64371
        
    Exit Sub

Erro_SiglaUMCompra_Click:

    Select Case Err

        Case 64371
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165505)

    End Select

    Exit Sub

End Sub

Public Sub SiglaUMVenda_Click()

Dim lErro As Long
Dim sSiglaUMVenda As String

On Error GoTo Erro_SiglaUMVenda_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMVenda.ListIndex = -1 Then Exit Sub

    sSiglaUMVenda = SiglaUMVenda.Text
    
    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMVenda, "NomeUMVenda")
    If lErro <> SUCESSO Then Error 64372
    
    Exit Sub

Erro_SiglaUMVenda_Click:

    Select Case Err

        Case 64372
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165506)
    
    End Select

    Exit Sub

End Sub

Public Sub SiglaUMEstoque_Click()

Dim lErro As Long
Dim sSiglaUMEstoque As String

On Error GoTo Erro_SiglaUMEstoque_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não tiver nada selecionado --> Sai
    If SiglaUMEstoque.ListIndex = -1 Then Exit Sub

    sSiglaUMEstoque = SiglaUMEstoque.Text
    
    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMEstoque, "NomeUMEstoque")
    If lErro <> SUCESSO Then Error 64373
    
    Exit Sub

Erro_SiglaUMEstoque_Click:

    Select Case Err

        Case 64373
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165507)

    End Select

    Exit Sub

End Sub

Public Sub SituacaoTributaria_Click()

Dim lErro As Long

On Error GoTo Erro_SituacaoTributaria_Click

    If SituacaoTributaria.ListIndex <> -1 Then

        iAlterado = REGISTRO_ALTERADO
    
        'Carrega a combobox de Aliquota
        lErro = Carrega_ComboAliquota
        If lErro <> SUCESSO Then gError 133670

    End If

    Exit Sub

Erro_SituacaoTributaria_Click:

    Select Case gErr

        Case 133670
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165508)
    
    End Select

    Exit Sub

End Sub

Public Sub Substituto1_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoSubst As String
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Substituto1_Validate

    'Verifica se está Preenchido
    If Len(Trim(Substituto1.ClipText)) = 0 Then
        DescSubst1.Caption = ""
        Exit Sub
    End If

    'Passa para o formato do BD
    lErro = CF("Produto_Formata", Substituto1.Text, sProdutoSubst, iPreenchido)
    If lErro <> SUCESSO Then Error 64374

    'Verifica se é igual ao Código
    If Len(Trim(Codigo.ClipText)) > 0 Then
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64375
        If sProduto = sProdutoSubst Then Error 64376
    End If

    'Verifica se é igual ao Substituto2
    If Len(Trim(Substituto2.ClipText)) > 0 Then
        lErro = CF("Produto_Formata", Substituto2.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64377
        If sProduto = sProdutoSubst Then Error 64378
    End If

    objProduto.sCodigo = sProdutoSubst

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 64379
    
    'Se não encontrar --> Erro
    If lErro = 28030 Then Error 64380

    'se o produto for gerencial ==> erro
    If objProduto.iGerencial = GERENCIAL Then Error 64381

    'Coloca a Descrição do Produto na Tela
    DescSubst1.Caption = objProduto.sDescricao

    Exit Sub

Erro_Substituto1_Validate:

    Cancel = True

    Select Case Err

        Case 64374, 64375, 64377, 64379

        Case 64376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SUBSTITUTO_IGUAL_PRODUTO", Err, Substituto1.Text)

        Case 64378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SUBSTITUTO1_IGUAL_SUBSTITUTO2", Err, Substituto1.Text)

        Case 64380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, Substituto1.Text)

        Case 64381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SUBSTITUTO_GERENCIAL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165509)

    End Select

    Exit Sub

End Sub

Public Sub Substituto2_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoSubst As String
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Substituto2_Validate

    'verifica se foi preenchido
    If Len(Trim(Substituto2.ClipText)) = 0 Then
        DescSubst2.Caption = ""
        Exit Sub
    End If

    'Passa para o formato do BD
    lErro = CF("Produto_Formata", Substituto2.Text, sProdutoSubst, iPreenchido)
    If lErro <> SUCESSO Then Error 64382

    'Verifica se é igual ao Código
    If Len(Trim(Codigo.ClipText)) > 0 Then
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64383
        If sProduto = sProdutoSubst Then Error 64384
    End If

    'Verifica se é igual ao Substituto2
    If Len(Trim(Substituto1.ClipText)) > 0 Then
        lErro = CF("Produto_Formata", Substituto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64385
        If sProduto = sProdutoSubst Then Error 64386
    End If
    
    objProduto.sCodigo = sProdutoSubst

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error 64387
    
    'Se não encontrar --> Erro
    If lErro = 28030 Then Error 64388

    'se o produto for gerencial ==> erro
    If objProduto.iGerencial = GERENCIAL Then Error 64389

    'Coloca a descrição na Tela
    DescSubst2.Caption = objProduto.sDescricao

    Exit Sub

Erro_Substituto2_Validate:

    Cancel = True

    Select Case Err

        Case 64382, 64383, 64385, 64387

        Case 64384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SUBSTITUTO_IGUAL_PRODUTO", Err, Substituto2.Text)

        Case 64386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SUBSTITUTO1_IGUAL_SUBSTITUTO2", Err, Substituto2.Text)

        Case 64388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, Substituto2.Text)

        Case 64389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SUBSTITUTO_GERENCIAL", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165510)

    End Select

    Exit Sub

End Sub

Public Sub TempoProducao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub TempoProducao_GotFocus()

    Call MaskEdBox_TrataGotFocus(TempoProducao, iAlterado)
        
End Sub

Public Sub TipoProduto_Change()

    iAlterado = REGISTRO_ALTERADO
    iTipoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoProduto_GotFocus()

Dim iTipoProdutoAux As Integer

    iTipoProdutoAux = iTipoAlterado
    Call MaskEdBox_TrataGotFocus(TipoProduto, iAlterado)
    iTipoAlterado = iTipoProdutoAux

End Sub

Public Sub TipoProduto_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoProduto As New ClassTipoDeProduto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoProduto_Validate

    'Verifica se o Tipo foi alterado
    If iTipoAlterado = 0 Then Exit Sub
    
    'Limpa a lista de características
    For iIndice = 0 To ListaCaracteristicas.ListCount - 1
        ListaCaracteristicas.Selected(iIndice) = False
    Next

    'Verifica se o Tipo está preenchido
    If Len(Trim(TipoProduto.Text)) = 0 Then
        DescTipoProduto.Caption = ""
        iTipoAlterado = 0
        Exit Sub

    End If

    'Critica o valor
    lErro = Inteiro_Critica(TipoProduto.Text)
    If lErro <> SUCESSO Then gError 64390

    objTipoProduto.iTipo = CInt(TipoProduto.Text)

    'Lê o tipo
    lErro = CF("TipoDeProduto_Le", objTipoProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then gError 64391
    
    'Se não encontrar --> Erro
    If lErro = 22531 Then gError 64392
        
    lErro = Exibe_Dados_TipoProduto(objTipoProduto)
    If lErro <> SUCESSO And lErro <> 31273 Then gError 64393
    
    iTipoAlterado = 0
    
    Exit Sub

Erro_TipoProduto_Validate:

    Cancel = True

    Select Case gErr

        Case 64390, 64391, 64393

        Case 64392
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOPRODUTO", objTipoProduto.iTipo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoProduto", objTipoProduto)

            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165511)

    End Select

    Exit Sub

End Sub

Function Exibe_Dados_TipoProduto(objTipoProduto As ClassTipoDeProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_Exibe_Dados_TipoProduto

    'Coloca a Descrição na Tela
    DescTipoProduto.Caption = objTipoProduto.sDescricao

    'Carrega a lista de características
    If objTipoProduto.iFaturamento = 1 Then ListaCaracteristicas.Selected(0) = True
    If objTipoProduto.iPCP = 1 Then ListaCaracteristicas.Selected(1) = True
    If objTipoProduto.iKitBasico = 1 Then ListaCaracteristicas.Selected(2) = True
    If objTipoProduto.iKitInt = 1 Then ListaCaracteristicas.Selected(3) = True
    If objTipoProduto.iKitVendaComp = 1 Then ListaCaracteristicas.Selected(4) = True
    
    If objTipoProduto.iApropriacaoCusto = APROPR_CUSTO_MEDIO Then
        Comprado.Value = True
        'Preenche a Combo Apropriação
        For iIndice = 0 To ApropriacaoComp.ListCount
            If ApropriacaoComp.ItemData(iIndice) = objTipoProduto.iApropriacaoCusto Then
                ApropriacaoComp.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        Produzido.Value = True
        'Preenche a Combo Apropriação
        For iIndice = 0 To ApropriacaoProd.ListCount - 1
            If ApropriacaoProd.ItemData(iIndice) = objTipoProduto.iApropriacaoCusto Then
                ApropriacaoProd.ListIndex = iIndice
                Exit For
            End If
        Next
    
    End If
    
    'Exibe a Natureza
    For iIndice = 0 To (NaturezaProduto.ListCount - 1)
    
        If NaturezaProduto.ItemData(iIndice) = objTipoProduto.iNatureza Then
            NaturezaProduto.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Ler as categorias e colocar no grid
    lErro = CF("TipoDeProduto_Le_Categorias", objTipoProduto, objTipoProduto.colCategoriaItem)
    If lErro <> SUCESSO Then Error 64394

    Call Grid_Limpa(objGridCategoria)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objTipoProduto.colCategoriaItem.Count

        'Insere no Grid Categoria
        GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = objTipoProduto.colCategoriaItem.Item(iIndice).sCategoria
        GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = objTipoProduto.colCategoriaItem.Item(iIndice).sItem

    Next

    objGridCategoria.iLinhasExistentes = objTipoProduto.colCategoriaItem.Count
    
    If objTipoProduto.iClasseUM <> 0 Then
        ClasseUM.Text = CStr(objTipoProduto.iClasseUM)
    Else
        ClasseUM.Text = ""
    End If
    
    Call ClasseUM_Validate(bCancel)
    
    For iIndice = 0 To SiglaUMCompra.ListCount - 1
        If SiglaUMCompra.List(iIndice) = objTipoProduto.sSiglaUMCompra Then
            SiglaUMCompra.ListIndex = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 0 To SiglaUMVenda.ListCount - 1
        If SiglaUMVenda.List(iIndice) = objTipoProduto.sSiglaUMVenda Then
            SiglaUMVenda.ListIndex = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 0 To SiglaUMEstoque.ListCount - 1
        If SiglaUMEstoque.List(iIndice) = objTipoProduto.sSiglaUMEstoque Then
            SiglaUMEstoque.ListIndex = iIndice
            Exit For
        End If
    Next
    
    For iIndice = 0 To SiglaUMTrib.ListCount - 1
        If SiglaUMTrib.List(iIndice) = objTipoProduto.sSiglaUMTrib Then
            SiglaUMTrib.ListIndex = iIndice
            Exit For
        End If
    Next
    
    If objTipoProduto.dResiduo <> -1 Then
        Residuo.Text = objTipoProduto.dResiduo * 100
    Else
        Residuo.Text = ""
    End If
    
    objProduto.dIPIAliquota = objTipoProduto.dIPIAliquota
    objProduto.sIPICodDIPI = objTipoProduto.sIPICodDIPI
    objProduto.sIPICodigo = objTipoProduto.sIPICodigo
    objProduto.sContaContabil = objTipoProduto.sContaContabil
    objProduto.sContaContabilProducao = objTipoProduto.sContaProducao
    objProduto.iExTIPI = objTipoProduto.iExTIPI
    objProduto.iOrigemMercadoria = objTipoProduto.iOrigem
    objProduto.sGenero = objTipoProduto.sGenero
    objProduto.sISSQN = objTipoProduto.sISSQN
    objProduto.iProdutoEspecifico = objTipoProduto.iProdutoEspecifico
    objProduto.sCEST = objTipoProduto.sCEST
    objProduto.iProdEmEscalaRelev = -1
    
    'Traz os dados do tipo para o Tab de Tributação
    lErro = Traz_TabTributacao_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64395
    
    Call Preenche_COMConfiguracoes(objTipoProduto)
    
    Exibe_Dados_TipoProduto = SUCESSO

    Exit Function
    
Erro_Exibe_Dados_TipoProduto:

    Exibe_Dados_TipoProduto = Err

    Select Case Err
    
        Case 64394, 64395
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165512)
            
    End Select
     
    Exit Function
    
End Function

Public Sub ClasseUM_Validate(Cancel As Boolean)

Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection

On Error GoTo Erro_ClasseUM_Validate

    'Verifica se a ClasseUM foi alterada
    If iClasseUMAlterada = 0 Then Exit Sub

    'Se estiver Preenchida
    If Len(Trim(ClasseUM.Text)) > 0 Then

        'Critica o valor
        lErro = Inteiro_Critica(ClasseUM.Text)
        If lErro <> SUCESSO Then gError 64396

        objClasseUM.iClasse = CInt(ClasseUM.Text)

    End If

    'Limpa o conteúdo das Combos e a Descrição da Classe
    SiglaUMEstoque.Clear
    SiglaUMCompra.Clear
    SiglaUMVenda.Clear
    SiglaUMTrib.Clear
    NomeUMEstoque.Caption = ""
    NomeUMVenda.Caption = ""
    NomeUMCompra.Caption = ""
    NomeUMTrib.Caption = ""
    DescricaoClasseUM.Caption = ""

    If objClasseUM.iClasse = 0 Then
        iClasseUMAlterada = 0
        iClasseUMAnterior = 0
        Exit Sub
    End If

    'Verificar se é uma classe cadastrada em ClasseUM
    lErro = CF("ClasseUM_Le", objClasseUM)
    If lErro <> SUCESSO And lErro <> 22537 Then gError 64397

    If lErro = 22537 Then gError 64398

    'Coloca na Tela a Descrição
    DescricaoClasseUM.Caption = objClasseUM.sDescricao

    lErro = Carrega_CombosUM(objClasseUM)
    If lErro <> SUCESSO Then gError 64399

    iClasseUMAlterada = 0
    iClasseUMAnterior = objClasseUM.iClasse

    Exit Sub

Erro_ClasseUM_Validate:

    Cancel = True

    Select Case gErr

        Case 64396, 64397, 64399

        Case 64398
        
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLASSEUM", objClasseUM.iClasse)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("ClasseUM", objClasseUM)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165513)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)
'se nao for produto do 1o nivel garantir que exista "pai" e este seja sintetico.
'Ex.Nao pode editar 1.1.2 se nao existir 1.1

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoSubst As String
Dim sProdutoAntFormatado As String
Dim sProduto As String
Dim objProd As New ClassProduto

On Error GoTo Erro_Codigo_Validate

    If iProdutoAlterado <> 0 Then
    
        If Len(Codigo.ClipText) > 0 Then
    
            sProduto = Codigo.Text
    
            lErro = CF("Trata_Segmento_Produto", sProduto)
            If lErro <> SUCESSO Then gError 199356
    
            Codigo.Text = sProduto
    
            'critica o formato da Produto
            lErro = CF("Produto_Formata", Codigo.Text, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 64400
    
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
                'Verifica se o Produto tem um Produto "Pai" já está cadastrado
                lErro = CF("Produto_Critica_ProdutoPai", sProdutoFormatado, MODULO_ESTOQUE)
                If lErro <> SUCESSO Then gError 64401
    
                'verifica se é igual ao substituto1
                If Len(Trim(Substituto1.ClipText)) > 0 Then
                    lErro = CF("Produto_Formata", Substituto1.Text, sProdutoSubst, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 64402
                    
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then If sProdutoFormatado = sProdutoSubst Then gError 64403
                    
                End If
    
                'verifica se é igual ao substituto2
                If Len(Trim(Substituto2.ClipText)) > 0 Then
                    lErro = CF("Produto_Formata", Substituto2.Text, sProdutoSubst, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 64404
                    
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then If sProdutoFormatado = sProdutoSubst Then gError 64405
                    
                End If
    
            End If
    
        End If
    
        'critica o formato da Produto
        lErro = CF("Produto_Formata", sProdutoAnterior, sProdutoAntFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 110081
    
        '***ALTERACAO POR TULIO EM 28/05***
        'Compara sProdutoAnterior com o produto que esta na tela
        'obs, sProdutoAnterior ja esta formatado.. (ver Codigo_GotFocus())
        If sProdutoAntFormatado <> sProdutoFormatado Then
    
            'se eles forem diferentes, limpa a combo de codigo de barras
            CodigoBarras.Clear
        
        End If
        '***FIM ALTERACAO POR TULIO EM 28/05***
        
        objProd.sCodigo = sProdutoFormatado

        lErro = CF("Produto_Le", objProd)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 64400
    
        If lErro = SUCESSO Then
            'Mostra os dados do Produto na tela
            lErro = Traz_Produto_Tela(objProd)
            If lErro <> SUCESSO Then Error 64400
        End If

        iProdutoAlterado = 0
        
    End If
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 64400, 64401, 64402, 64404, 110081, 199356

        Case 64403, 64405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SUBSTITUTO_IGUAL_PRODUTO", gErr, Codigo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165514)

    End Select

    Exit Sub

End Sub

Public Sub LblTipoProduto_Click()

Dim objTipoDeProduto As ClassTipoDeProduto
Dim colSelecao As Collection

    Call Chama_Tela("TipoProdutoLista", colSelecao, objTipoDeProduto, objEventoTipoDeProduto)

End Sub

Public Sub LblClasseUM_Click()

Dim objClasseUM As ClassUnidadeDeMedida
Dim colSelecao As Collection

    Call Chama_Tela("ClasseUMLista", colSelecao, objClasseUM, objEventoClasseUM)

End Sub

Public Sub NivelFinal_Click()
    
Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim iTemFilho As Integer

On Error GoTo Erro_NivelFinal_Click

    iAlterado = REGISTRO_ALTERADO
    Grades.Text = ""
    LabelGrade.Enabled = False
    Grades.Enabled = False
    
    'Verifica se o Código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Passa para o formato do BD
    lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then Error 64355

    If iPreenchido <> PRODUTO_PREENCHIDO Then Exit Sub

    'Verifica se o Produto tem filhos
    lErro = CF("Produto_Tem_Filho", sProduto, iTemFilho)
    If lErro <> SUCESSO Then Error 64356

    'Se tiver erro
    If iTemFilho = True Then Error 64357

    Exit Sub

Erro_NivelFinal_Click:

    Select Case Err

        Case 64355, 64356

        Case 64357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FINAL_COM_FILHOS", Err, Codigo.Text)
            'NivelGerencial.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165490)

    End Select

    Exit Sub
    

End Sub

Public Sub Substituto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Substituto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub GridCategoria_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_GotFocus()

    Call Grid_Recebe_Foco(objGridCategoria)

End Sub

Public Sub GridCategoria_EnterCell()

    Call Grid_Entrada_Celula(objGridCategoria, iAlterado)

End Sub

Public Sub GridCategoria_LeaveCell()

    Call Saida_Celula(objGridCategoria)

End Sub

Public Sub GridCategoria_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCategoria)

End Sub

Public Sub GridCategoria_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCategoria, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCategoria, iAlterado)
    End If

End Sub

Public Sub GridCategoria_LostFocus()

    Call Grid_Libera_Foco(objGridCategoria)

End Sub

Public Sub GridCategoria_RowColChange()

    Call Grid_RowColChange(objGridCategoria)

End Sub

Public Sub GridCategoria_Scroll()

    Call Grid_Scroll(objGridCategoria)

End Sub

'Raphael - 01/00: grid de Tabela de Preço deixa de ser editavel
'Private Sub GridTabelaPreco_Click()
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Click(objGridTabelaPreco, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridTabelaPreco, iAlterado)
'    End If
'
'End Sub
'
'Private Sub GridTabelaPreco_GotFocus()
'
'    Call Grid_Recebe_Foco(objGridTabelaPreco)
'
'End Sub
'
'Private Sub GridTabelaPreco_EnterCell()
'
'    Call Grid_Entrada_Celula(objGridTabelaPreco, iAlterado)
'
'End Sub
'
'Private Sub GridTabelaPreco_LeaveCell()
'
'    Call Saida_Celula(objGridTabelaPreco)
'
'End Sub
'
'Private Sub GridTabelaPreco_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyDelete Then Exit Sub
'
'    Call Grid_Trata_Tecla1(KeyCode, objGridTabelaPreco)
'
'End Sub
'
'Private Sub GridTabelaPreco_KeyPress(KeyAscii As Integer)
'
'Dim iExecutaEntradaCelula As Integer
'
'    Call Grid_Trata_Tecla(KeyAscii, objGridTabelaPreco, iExecutaEntradaCelula)
'
'    If iExecutaEntradaCelula = 1 Then
'        Call Grid_Entrada_Celula(objGridTabelaPreco, iAlterado)
'    End If
'
'End Sub
'
'Private Sub GridTabelaPreco_LostFocus()
'
'    Call Grid_Libera_Foco(objGridTabelaPreco)
'
'End Sub
'
'Private Sub GridTabelaPreco_RowColChange()
'
'    Call Grid_RowColChange(objGridTabelaPreco)
'
'End Sub
'
'Private Sub GridTabelaPreco_Scroll()
'
'    Call Grid_Scroll(objGridTabelaPreco)
'
'End Sub
'
Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        If objGridInt.objGrid.Name = GridCategoria.Name Then

            'Verifica qual a coluna do Grid
            Select Case GridCategoria.Col

                Case iGrid_Categoria_Col

                    lErro = Saida_Celula_Categoria(objGridInt)
                    If lErro <> SUCESSO Then Error 64406

                Case iGrid_Valor_Col

                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then Error 64407

            End Select

'Raphael - 01/00: grid de Tabela de Preço deixa de ser editavel
'        'Verifica se o Grid é o GridTabelaPreco
'        ElseIf objGridInt.objGrid.Name = GridTabelaPreco.Name Then
'
'            'Verifica qual a coluna do Grid
'            Select Case GridTabelaPreco.Col
'
'                Case iGrid_DataPreco_Col
'
'                    lErro = Saida_Celula_DataPreco(objGridInt)
'                    If lErro <> SUCESSO Then Error 26998
'
'                Case iGrid_ValorEmpresa_Col, iGrid_ValorFilial_Col
'
'                    lErro = Saida_Celula_ValorProduto(objGridInt)
'                    If lErro <> SUCESSO Then Error 26999
'
'            End Select

        Else

            lErro = CF2(Me, "Prod_Saida_Celula", objGridInt)
            If lErro <> SUCESSO Then Error 64407

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 64408

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 64406, 64407

        Case 64408
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165515)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Categoria(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto

On Error GoTo Erro_Saida_Celula_Categoria

    Set objGridInt.objControle = ComboCategoriaProduto

    If Len(Trim(ComboCategoriaProduto.Text)) > 0 Then

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProduto)
        If lErro <> SUCESSO Then
        
            'Preenche o objeto com a Categoria
             objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

             'Lê Categoria De Produto no BD
             lErro = CF("CategoriaProduto_Le", objCategoriaProduto)
             If lErro <> SUCESSO And lErro <> 22540 Then Error 64409

             If lErro <> SUCESSO Then Error 64410  'Categoria não está cadastrada

        End If

        'Verifica se já existe a categoria no Grid
        For iIndice = 1 To objGridCategoria.iLinhasExistentes

            If iIndice <> GridCategoria.Row Then If GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = ComboCategoriaProduto.Text Then Error 64411

        Next

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    Else
        
        GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Valor_Col) = ""
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 64412

    Saida_Celula_Categoria = SUCESSO

    Exit Function

Erro_Saida_Celula_Categoria:

    Saida_Celula_Categoria = Err

    Select Case Err

        Case 64409, 64412
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 64410  'Categoria não está cadastrada

            'pergunta se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTO")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 64411
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_JA_SELECIONADA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165516)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Item do grid que está deixando de ser a corrente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colItens As New Collection

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem

    If Len(Trim(ComboCategoriaProdutoItem.Text)) > 0 Then

        'se o campo de categoria estiver vazio ==> erro
        If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) = 0 Then Error 64413

        'Tenta selecionar na combo
        lErro = Combo_Item_Igual(ComboCategoriaProdutoItem)
        If lErro <> SUCESSO Then

            'Preenche o objeto com a Categoria
            objCategoriaProdutoItem.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)
            objCategoriaProdutoItem.sItem = ComboCategoriaProdutoItem.Text

            'Lê Categoria De Produto no BD
            lErro = CF("CategoriaProduto_Le_Item", objCategoriaProdutoItem)
            If lErro <> SUCESSO And lErro <> 22603 Then Error 64414

            If lErro <> SUCESSO Then Error 64415 'Item da Categoria não está cadastrado

        End If

        If GridCategoria.Row - GridCategoria.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 64416

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 64413
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_CATEGORIA_NAO_PREENCHIDA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 64414, 64416
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 64415 'Item da Categoria não está cadastrado
            'Se não for perguntar se deseja criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CATEGORIAPRODUTOITEM")

            If vbMsgRes = vbYes Then

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Preenche o objeto com a Categoria
                objCategoriaProduto.sCategoria = ComboCategoriaProduto.Text

                'Chama a Tela "CategoriaProduto"
                Call Chama_Tela("CategoriaProduto", objCategoriaProduto)

            Else

                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165517)

    End Select

    Exit Function

End Function

'Private Sub TvwProduto_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim lErro As Long
'Dim sCodigo As String
'Dim objProduto As New ClassProduto
'
'On Error GoTo Erro_TvwProduto_NodeClick
'
'    'Armazena key do nó clicado sem caracter inicial
'    sCodigo = Right(Node.Key, Len(Node.Key) - 1)
'
'    objProduto.sCodigo = sCodigo
'
'    'Lê Produto
'    lErro = CF("Produto_Le",objProduto)
'    If lErro <> SUCESSO And lErro <> 28030 Then Error 64417
'
'    'Mostra os dados do Produto na tela
'    lErro = Traz_Produto_Tela(objProduto)
'    If lErro <> SUCESSO Then Error 64418
'
'    'Fecha comando de setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Exit Sub
'
'Erro_TvwProduto_NodeClick:
'
'    Select Case Err
'
'        Case 64417, 64418
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165518)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long
'Mostra os dados do Produto na tela

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_Traz_Produto_Tela

    'Limpa a Tela
    Call Limpa_Tela_Produto

    lErro = CF2(Me, "Prod_Traz_Produto_Tela", objProduto)
    If lErro <> SUCESSO Then Error 64419

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then Error 64419

    'Coloca o Codigo na tela
    Codigo.PromptInclude = False
    Codigo.Text = sProdutoEnxuto
    Codigo.PromptInclude = True
    sProdutoAnterior = Codigo.Text
    iProdutoAlterado = 0
    
    'Coloca os demais dados do Produto na tela
    Descricao.Text = objProduto.sDescricao
    NomeReduzido.Text = objProduto.sNomeReduzido
    Modelo.Text = objProduto.sModelo
    Referencia.Text = objProduto.sReferencia
    NomeFigura.Text = objProduto.sFigura

    For iIndice = 0 To NaturezaProduto.ListCount - 1
    
        If NaturezaProduto.ItemData(iIndice) = objProduto.iNatureza Then
            NaturezaProduto.ListIndex = iIndice
            Exit For
        End If
    Next
        
    If objProduto.iAtivo = PRODUTO_ATIVO Then
        Ativo.Value = vbChecked
    Else
        Ativo.Value = vbUnchecked
    End If

    If objProduto.iGerencial = GERENCIAL Then
        NivelGerencial.Value = MARCADO
    Else
        NivelFinal.Value = MARCADO
    End If

    objTipoProduto.iTipo = objProduto.iTipo
    
    'Lê o tipo
    lErro = CF("TipoDeProduto_Le", objTipoProduto)
    If lErro <> SUCESSO And lErro <> 22531 Then Error 64420
    
    'Se encontrar --> trazer tipo para a tela
    If lErro = SUCESSO Then
    
        TipoProduto.Text = CStr(objTipoProduto.iTipo)
        DescTipoProduto.Caption = objTipoProduto.sDescricao
        iTipoAlterado = 0
        
    End If
    
    If gobjEST.iProdExibeComissao = MARCADO Then
        PercComissao.Text = CStr(100 * objProduto.dPercComissao)
        MetaComissao.Text = Formata_Estoque(objProduto.dMetaComissao)
    End If
    
    'Traz os dados do Tab com título de Categoria
    lErro = Traz_TabCategoria_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64421
    
    'Traz os dados do Tab com título de Complemento
    lErro = Traz_TabComplemento_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64422

    'Traz os dados do Tab com título de Características
    lErro = Traz_TabCaracteristicas_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64423

    'Traz os dados do Tab com título de Preços
    lErro = CF2(Me, "Prod_Traz_TabPrecos_Tela", objProduto)
    If lErro <> SUCESSO Then Error 64424
    
    'Traz os dados do Tab com título de Unidades de Medidas
    lErro = Traz_TabUnidadesMedida_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64425

    'Traz os dados do Tab com título de Tributação
    lErro = Traz_TabTributacao_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64426
        
    'Traz os dados do Tab de Compras
    lErro = Traz_TabCompras_Tela(objProduto)
    If lErro <> SUCESSO Then Error 64427
    
    Call gobjIdiomaTela.Reiniciar
    
    lErro = gobjIdiomaTela.Associar("Produtos", 0, objProduto.sCodigo, 0, 0)
    If lErro <> SUCESSO Then gError 64427
    
    lErro = CF("IdiomaTextos_Le", gobjIdiomaTela)
    If lErro <> SUCESSO Then gError 64427
    
    Call Combo_Seleciona_ItemData(Fabricante, objProduto.lFabricante)
    
    Call BotaoVisualizar_Click
    
    iAlterado = 0

    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err

        Case 64419, 64420, 64421, 64422, 64423, 64424, 64425, 64426, 64427
                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165519)

    End Select

    Exit Function

End Function

Private Function Traz_TabCategoria_Tela(objProduto As ClassProduto) As Long
'traz os dados do tab de categoria do BD para a tela

Dim iIndice As Integer

On Error GoTo Erro_Traz_TabCategoria_Tela

    'Carrega a lista de características
    If objProduto.iFaturamento = 1 Then ListaCaracteristicas.Selected(0) = True
    If objProduto.iPCP = 1 Then ListaCaracteristicas.Selected(1) = True
    If objProduto.iKitBasico = 1 Then ListaCaracteristicas.Selected(2) = True
    If objProduto.iKitInt = 1 Then ListaCaracteristicas.Selected(3) = True
    If objProduto.iKitVendaComp = 1 Then ListaCaracteristicas.Selected(4) = True
    
    Grades.Text = objProduto.sGrade
    
    Call Grid_Limpa(objGridCategoria)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objProduto.colCategoriaItem.Count

        'Insere no Grid Categoria
        GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = objProduto.colCategoriaItem.Item(iIndice).sCategoria
        GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = objProduto.colCategoriaItem.Item(iIndice).sItem

    Next

    objGridCategoria.iLinhasExistentes = objProduto.colCategoriaItem.Count
    
    Traz_TabCategoria_Tela = SUCESSO

    Exit Function
    
Erro_Traz_TabCategoria_Tela:

    Traz_TabCategoria_Tela = Err

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165520)
            
    End Select
     
    Exit Function
    
End Function

Private Function Traz_TabComplemento_Tela(objProduto As ClassProduto) As Long
'traz os dados do tab de complemento do BD para a tela

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objProdutoSubst As New ClassProduto
Dim objCategoriaItem As ClassCategoriaProduto
Dim objProdutoFilial As ClassProdutoFilial
Dim colProdutoFilial As New Collection
Dim dQuantSoma As Double
Dim iSituacao As Integer
Dim iIndex As Integer
Dim sSigla As String

On Error GoTo Erro_Traz_TabComplemento_Tela

    'Se tiver um produto substituto
    If Len(Trim(objProduto.sSubstituto1)) > 0 Then

        objProdutoSubst.sCodigo = objProduto.sSubstituto1

        'Lê o produto Substituto1
        lErro = CF("Produto_Le", objProdutoSubst)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 64428
        
        'Se não achou o Produto --> Erro
        If lErro = 28030 Then gError 64429

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sSubstituto1, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 64430
        
        'Coloca o produto na tela
        Substituto1.PromptInclude = False
        Substituto1.Text = sProdutoEnxuto
        Substituto1.PromptInclude = True
        DescSubst1.Caption = objProdutoSubst.sDescricao

    End If

    'verifica se tem um Substituto2
    If Len(Trim(objProduto.sSubstituto2)) > 0 Then
    
        objProdutoSubst.sCodigo = objProduto.sSubstituto2
        
        'Tenta ler o produto
        lErro = CF("Produto_Le", objProdutoSubst)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 64431
        
        'não achou --> Erro
        If lErro = 28030 Then gError 64432

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sSubstituto2, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 64433
        
        'Coloca o Produto na tela
        Substituto2.PromptInclude = False
        Substituto2.Text = sProdutoEnxuto
        Substituto2.PromptInclude = True
        DescSubst2.Caption = objProdutoSubst.sDescricao
        
    End If

    'Carrega a colCategoriaItem do objProduto
    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
    If lErro <> SUCESSO Then gError 64434

    'Carrega o Grid Categoria
    For iIndice = 1 To objProduto.colCategoriaItem.Count
        GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col) = objProduto.colCategoriaItem.Item(iIndice).sCategoria
        GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col) = objProduto.colCategoriaItem.Item(iIndice).sItem
    Next
    
    objGridCategoria.iLinhasExistentes = objProduto.colCategoriaItem.Count
       
    'Preenche os demais dados da tela
    PrazoValidade.Text = objProduto.iPrazoValidade
    '***ALTERACAO POR TULIO EM 27/05***
    'Comentei o codigo abaixo, pois agora o tratamento ira mudar..
    'CodigoBarras.Text = objProduto.sCodigoBarras
        
    'Faz a leitura dos codigos de barra
    lErro = CF("CodigosBarra_Le_Produto", objProduto)
    If lErro <> SUCESSO Then gError 101730
    
    'Chama a funcao responsavel por carregar a combo de codigo de barras utilizando
    'os codigos da colecao do produto
    Call Carrega_CodigoBarras_Produto(objProduto.colCodBarras)
    
    '***FIM ALTERACAO POR TULIO EM 27/05***
    
    EtiquetasCodBarras.Text = objProduto.iEtiquetasCodBarras
    
    If objProduto.dResiduo <> -1 Then
        Residuo.Text = objProduto.dResiduo * 100
    Else
        Residuo.Text = ""
    End If
    
    If objProduto.iCompras = PRODUTO_COMPRAVEL Then
    
        'Preenche a Combo Apropriação
        Comprado.Value = True
        
        For iIndice = 0 To ApropriacaoComp.ListCount
            If ApropriacaoComp.ItemData(iIndice) = objProduto.iApropriacaoCusto Then
                ApropriacaoComp.ListIndex = iIndice
                Exit For
            End If
        Next
        
    Else
    
        Produzido.Value = True
        'Preenche a Combo Apropriação
        For iIndice = 0 To ApropriacaoProd.ListCount - 1
        
            If ApropriacaoProd.ItemData(iIndice) = objProduto.iApropriacaoCusto Then
                ApropriacaoProd.ListIndex = iIndice
                Exit For
            End If
            
        Next
        
        TempoProducao.Text = objProduto.iTempoProducao
        
    
    End If
    
    'Preenche o custo Reposicao
    CustoReposicao.Text = CStr(objProduto.dCustoReposicao)
        
    'Traz Quantidade em Pedido
    If giFilialEmpresa = EMPRESA_TODA Then
            
        'Passa o Produto e obtem a coleção com o Produto para Todas as Filiais
        lErro = CF("ProdutoFiliais_Le", objProduto.sCodigo, colProdutoFilial)
        If lErro <> SUCESSO And lErro <> 61127 Then gError 64435
        
        For Each objProdutoFilial In colProdutoFilial
        
            dQuantSoma = dQuantSoma + objProdutoFilial.dQuantPedida
        
        Next
        
        QuantPedido.Caption = Formata_Estoque(dQuantSoma)
        
    Else
        
        Set objProdutoFilial = New ClassProdutoFilial
        
        objProdutoFilial.sProduto = objProduto.sCodigo
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Passa Produto e Filial e obtem QuantPedida
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 64436
        
        QuantPedido.Caption = Formata_Estoque(objProdutoFilial.dQuantPedida)
        iSituacao = -1
        Select Case objProdutoFilial.sSituacaoTribECF
            
            Case TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO
                iSituacao = TIPOTRIBICMS_NAO_TRIBUTADO
                        
            Case TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA
                iSituacao = TIPOTRIBICMS_ISENTA
                    
            Case TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST
                iSituacao = TIPOTRIBICMS_TRIB_SUBST
            
            Case TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL
                iSituacao = TIPOTRIBICMS_INTEGRAL
                 
            Case TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO
                iSituacao = TIPOTRIBISS_NAO_TRIBUTADO
                        
            Case TIPOTRIBISS_SITUACAOTRIBECF_ISENTA
                iSituacao = TIPOTRIBISS_ISENTA
                    
            Case TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST
                iSituacao = TIPOTRIBISS_TRIB_SUBST
            
            Case TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL
                iSituacao = TIPOTRIBISS_INTEGRAL
            
        End Select
            
        For iIndice = 0 To SituacaoTributaria.ListCount - 1
        
            If SituacaoTributaria.ItemData(iIndice) = iSituacao Then
            
                SituacaoTributaria.ListIndex = iIndice
                
                Exit For
                
            End If
            
        Next
        'Modificado por cyntia para transformá-lo numa combo
        'inicio - Daniel a pedido de Mario
        'Busca na combo a aliquota do obj
        For iIndex = 0 To comboAliquota.ListCount - 1
        
            sSigla = SCodigo_Extrai(comboAliquota.List(iIndex))
        
            'Quando achar, seleciona este item na combo
            If sSigla = objProdutoFilial.sICMSAliquota Then
                comboAliquota.ListIndex = iIndex
                Exit For
            End If
            
        Next
        'fim
        
    End If
    
    '#############################################
    'Inserido por Wagner 08/03/2006
    SerieProx.Text = objProduto.sSerieProx
    
    If objProduto.iSerieParteNum <> 0 Then
        SerieNum.Text = objProduto.iSerieParteNum
    End If
    
    lErro = Traz_Serie_ParteNumerica_Tela()
    If lErro <> SUCESSO Then gError 141789
    '#############################################
    
    Traz_TabComplemento_Tela = SUCESSO

    Exit Function

Erro_Traz_TabComplemento_Tela:

    Traz_TabComplemento_Tela = gErr

    Select Case gErr

        Case 64428, 64430, 64431, 64433, 64434, 64435, 64436, 101730, 141789

        Case 64429, 64432
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SUBSTITUTO_INEXISTENTE", gErr, objProdutoSubst.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165521)

    End Select

    Exit Function

End Function

Private Function Traz_TabCaracteristicas_Tela(objProduto As ClassProduto) As Long
'traz os dados do tab de caracteristicas do BD para a tela

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_TabCaracteristicas_Tela

    'Coloca as carcterísticas físicas do Produto na tela
    If objProduto.dPesoLiq > 0 Then PesoLiquido.Text = Format(objProduto.dPesoLiq, PesoLiquido.Format) 'Alterado por Wagner
    If objProduto.dPesoBruto > 0 Then PesoBruto.Text = Format(objProduto.dPesoBruto, PesoBruto.Format) 'Alterado por Wagner
    If objProduto.dComprimento > 0 Then Comprimento.Text = Format(objProduto.dComprimento, Comprimento.Format)
    If objProduto.dLargura > 0 Then Largura.Text = Format(objProduto.dLargura, Largura.Format)
    If objProduto.dEspessura > 0 Then Espessura.Text = Format(objProduto.dEspessura, Espessura.Format)
    Cor.Text = objProduto.sCor
    ObsFisica.Text = objProduto.sObsFisica
    
    '05/09/01 Marcelo
    'Verifica se objProduto.iEmbalagem está preenchido
'    If objProduto.iEmbalagem > 0 Then
'        'Atribui o valor de objProduto.iEmbalagem a Embalagem.Text
'        Embalagem.Text = objProduto.iEmbalagem
'        Call Embalagem_Validate(bSGECancelDummy)
'    End If
        
    If objProduto.dPesoEspecifico > 0 Then PesoEspecifico.Text = Format(objProduto.dPesoEspecifico, PesoEspecifico.Format) 'Alterado por Wagner
    If objProduto.lHorasMaquina > 0 Then HorasMaquina.Text = objProduto.lHorasMaquina
    
    For iIndice = 0 To Rastro.ListCount
        If Rastro.ItemData(iIndice) = objProduto.iRastro Then
            Rastro.ListIndex = iIndice
            Exit For
        End If
    Next

    lErro = CF2(Me, "Prod_TrazCarac_Tela", objProduto)
    If lErro <> SUCESSO Then gError 184270

    Traz_TabCaracteristicas_Tela = SUCESSO

    Exit Function

Erro_Traz_TabCaracteristicas_Tela:

    Traz_TabCaracteristicas_Tela = gErr

    Select Case gErr

        Case 184270
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165522)

    End Select

    Exit Function

End Function

Function Prod_Traz_TabPrecos_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
'traz os dados do tab tabela de preços do BD para a tela
'???? existe customizacao para a artmill

Dim lErro As Long
Dim iIndice As Integer
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim colTabelaPrecoItem As New Collection
Dim colTabPrecoEmpresa As New Collection
Dim iLinha As Integer

On Error GoTo Erro_Prod_Traz_TabPrecos_Tela

    'Limpa o Grid de Tabelas de Preço
    Call Grid_Limpa(objGridTabelaPreco)

    'Lê os preços do Produtos nas Tabela
    lErro = CF("Produto_Le_TabelaPrecoItem", objProduto, colTabelaPrecoItem, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 64437
    
    'Redimenciona o Grid
    If colTabelaPrecoItem.Count >= objGridTabelaPreco.objGrid.Rows Then
        objGridTabelaPreco.objGrid.Rows = colTabelaPrecoItem.Count + 1
        objGridTabelaPreco.iLinhasExistentes = colTabelaPrecoItem.Count
        Call Grid_Inicializa(objGridTabelaPreco)
'    Else
'        objGridTabelaPreco.objGrid.Rows = 11
'        objGridTabelaPreco.iLinhasExistentes = colTabelaPrecoItem.Count
'        Call Grid_Inicializa(objGridTabelaPreco)
    End If
    
    For Each objTabelaPrecoItem In colTabelaPrecoItem
        
        iIndice = iIndice + 1
        
        GridTabelaPreco.TextMatrix(iIndice, iGrid_Tabela_Col) = objTabelaPrecoItem.iCodTabela
        GridTabelaPreco.TextMatrix(iIndice, iGrid_DescricaoTabela_Col) = objTabelaPrecoItem.sDescricaoTabela
        
        'Coloca o Preço na tela
        If giFilialEmpresa <> EMPRESA_TODA Then
            GridTabelaPreco.TextMatrix(iIndice, iGrid_ValorFilial_Col) = Format(objTabelaPrecoItem.dPreco, gobjFAT.sFormatoPrecoUnitario)
        End If

        If objTabelaPrecoItem.dtDataVigencia <> DATA_NULA Then GridTabelaPreco.TextMatrix(iIndice, iGrid_DataPreco_Col) = Format(objTabelaPrecoItem.dtDataVigencia, "dd/mm/yyyy")
        
    Next
    
    'No caso de não ser empresa toda mostra na tela o preço e empresa toda
    If giFilialEmpresa <> EMPRESA_TODA Then

        'Lê os preços para empresa toda
        lErro = CF("Produto_Le_TabelaPrecoItem", objProduto, colTabPrecoEmpresa, EMPRESA_TODA)
        If lErro <> SUCESSO Then Error 64438

        For Each objTabelaPrecoItem In colTabPrecoEmpresa
        
            'Descobre a linha no Grid da Tabela em questão
            For iIndice = 1 To objGridTabelaPreco.iLinhasExistentes
                If GridTabelaPreco.TextMatrix(iIndice, iGrid_Tabela_Col) = objTabelaPrecoItem.iCodTabela Then
                    iLinha = iIndice
                    Exit For
                End If
            Next
            
            'Coloca o Preço no Grid
            If iLinha <= objGridTabelaPreco.iLinhasExistentes And iLinha > 0 Then
                GridTabelaPreco.TextMatrix(iLinha, iGrid_ValorEmpresa_Col) = Format(objTabelaPrecoItem.dPreco, gobjFAT.sFormatoPrecoUnitario)
            End If
        Next

    End If
    
    DescrUM.Caption = objProduto.sSiglaUMVenda
    
    Prod_Traz_TabPrecos_Tela = SUCESSO

    Exit Function

Erro_Prod_Traz_TabPrecos_Tela:

    Prod_Traz_TabPrecos_Tela = Err

    Select Case Err

        Case 64437, 64438

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165523)

    End Select

    Exit Function

End Function

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Traz_TabUnidadesMedida_Tela(objProduto As ClassProduto) As Long
'traz os dados do tab de unidades de medida do BD para a tela

Dim lErro As Long
Dim objClasseUM As New ClassClasseUM
Dim iIndice As Integer
Dim bCancel As Boolean

    ClasseUM.Text = objProduto.iClasseUM
    Call ClasseUM_Validate(bCancel)

    'Seleciona nas combos as U.M. contidas no objProduto
    If SiglaUMCompra.ListCount > 0 Then
    
        'U.M. de Estoque
        For iIndice = 0 To SiglaUMEstoque.ListCount - 1
            If SiglaUMEstoque.List(iIndice) = objProduto.sSiglaUMEstoque Then
                SiglaUMEstoque.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'U.M. de Compra
        For iIndice = 0 To SiglaUMCompra.ListCount - 1
            If SiglaUMCompra.List(iIndice) = objProduto.sSiglaUMCompra Then
                SiglaUMCompra.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'U.M. de Venda
        For iIndice = 0 To SiglaUMVenda.ListCount - 1
            If SiglaUMVenda.List(iIndice) = objProduto.sSiglaUMVenda Then
                SiglaUMVenda.ListIndex = iIndice
                Exit For
            End If
        Next

        'U.M. de Trib
        For iIndice = 0 To SiglaUMTrib.ListCount - 1
            If SiglaUMTrib.List(iIndice) = objProduto.sSiglaUMTrib Then
                SiglaUMTrib.ListIndex = iIndice
                Exit For
            End If
        Next
    End If

    Traz_TabUnidadesMedida_Tela = SUCESSO

End Function

Private Function Traz_TabTributacao_Tela(objProduto As ClassProduto) As Long
'Preenche o Tab Tributação com os dados do BD

Dim lErro As Long
Dim sContaMascarada As String
Dim iIndex As Integer

On Error GoTo Erro_Traz_TabTributacao_Tela

    'Relacionados ao IPI
    If objProduto.dIPIAliquota > 0 Or Len(Trim(objProduto.sIPICodDIPI)) > 0 Or objProduto.iIPIIncide <> 0 Then
    
        IncideIPI.Value = vbChecked
        AliquotaIPI.Enabled = True
        AliquotaIPI.Text = objProduto.dIPIAliquota * 100
        CodigoIPI.Enabled = True
        CodigoIPI.Text = objProduto.sIPICodDIPI
        
    Else
        IncideIPI = vbUnchecked
        AliquotaIPI.Text = ""
        AliquotaIPI.Enabled = False
        CodigoIPI.Text = ""
        CodigoIPI.Enabled = False
    End If
    
    ClasFiscIPI.PromptInclude = False
    ClasFiscIPI.Text = objProduto.sIPICodigo
    ClasFiscIPI.PromptInclude = True
        
    CEST.PromptInclude = False
    CEST.Text = objProduto.sCEST
    CEST.PromptInclude = True
    
    cProdANVISA.Text = objProduto.scProdANVISA
    
'    If objProduto.iProdEmEscalaRelev = MARCADO Then
'        ProdEmEscalaRelev.Value = vbChecked
'    ElseIf objProduto.iProdEmEscalaRelev = DESMARCADO Then
'        ProdEmEscalaRelev.Value = vbUnchecked
'    Else
'        ProdEmEscalaRelev.Value = vbGrayed
'    End If

    Call Combo_Seleciona_ItemData(ProdEmEscalaRelev, objProduto.iProdEmEscalaRelev)
    
    If objProduto.sContaContabil <> "" Then
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabil, sContaMascarada)
        If lErro <> SUCESSO Then Error 64439
    Else
        sContaMascarada = ""
    End If

    ContaContabil.PromptInclude = False
    ContaContabil.Text = sContaMascarada
    ContaContabil.PromptInclude = True

    If objProduto.sContaContabilProducao <> "" Then
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objProduto.sContaContabilProducao, sContaMascarada)
        If lErro <> SUCESSO Then Error 64440
    Else
        sContaMascarada = ""
    End If

    ContaProducao.PromptInclude = False
    ContaProducao.Text = sContaMascarada
    ContaProducao.PromptInclude = True

    'Preenche a Origem na Tela
    Call Combo_Seleciona_ItemData(OrigemMercadoria, objProduto.iOrigemMercadoria)
    
    If objProduto.dINSSPercBase > 0 Then
        INSSPercBase.Text = objProduto.dINSSPercBase * 100
    Else
        INSSPercBase.Text = ""
    End If
    
    If gobjMAT.iProdFatorAjuste <> 0 Then
    
        If objProduto.dFatorAjuste > 0 Then
            FatorAjuste.Text = objProduto.dFatorAjuste * 100
        Else
            FatorAjuste.Text = ""
        End If
    
    End If
    
    FCI.Text = objProduto.sFCI
    
    If objProduto.iUsaBalanca = USA_BALANCA Then
        UsaBalanca.Value = vbChecked
    Else
        UsaBalanca.Value = vbUnchecked
    End If
    
    'Busca na combo o kit de venda do obj
''''    If objProduto.iKitVendaComp <> -1 Then
''''
''''        For iIndex = 0 To KitVendaComp.ListCount - 1
''''            'Quando achar, seleciona este item na combo
''''            If KitVendaComp.ItemData(iIndex) = objProduto.iKitVendaComp Then
''''                KitVendaComp.ListIndex = iIndex
''''                Exit For
''''            End If
''''        Next
''''
''''    End If

    If objProduto.dPrecoMaxConsumidor <> 0 Then
        PrecoMaxConsumidor.Text = Format(objProduto.dPrecoMaxConsumidor, "STANDARD")
    End If
    If objProduto.iExTIPI <> 0 Then
        ExTIPI.PromptInclude = False
        ExTIPI.Text = CStr(objProduto.iExTIPI)
        ExTIPI.PromptInclude = True
    End If
    Genero.PromptInclude = False
    Genero.Text = Format(objProduto.sGenero, Genero.Format)
    Genero.PromptInclude = True
    Call Genero_Validate(bSGECancelDummy)
    ISSQN.PromptInclude = False
    ISSQN.Text = Format(objProduto.sISSQN, ISSQN.Format)
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)
    ProdutoEspecifico.ListIndex = objProduto.iProdutoEspecifico
    
    Traz_TabTributacao_Tela = SUCESSO

    Exit Function

Erro_Traz_TabTributacao_Tela:

    Traz_TabTributacao_Tela = Err

    Select Case Err

        Case 64439, 64440

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165524)

    End Select

    Exit Function

End Function

Private Function Carrega_CombosUM(objClasseUM As ClassClasseUM) As Long
'Carrega as combos de Unidades de Medida de acordo com a ClasseUM passada

Dim lErro As Long
Dim colSiglas As New Collection
Dim iIndice As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_Carrega_CombosUM

    'Lê as U.M. da Classe passada
    lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
    If lErro <> SUCESSO And lErro <> 22539 Then Error 64441
    
    'Carrega as combos
    If lErro = SUCESSO Then

        For iIndice = 1 To colSiglas.Count
            SiglaUMCompra.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMEstoque.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMVenda.AddItem colSiglas.Item(iIndice).sSigla
            SiglaUMTrib.AddItem colSiglas.Item(iIndice).sSigla
            If colSiglas.Item(iIndice).sSigla = objClasseUM.sSiglaUMBase Then iIndice2 = iIndice
        Next

        SiglaUMCompra.ListIndex = iIndice2 - 1
        SiglaUMEstoque.ListIndex = iIndice2 - 1
        SiglaUMVenda.ListIndex = iIndice2 - 1
        SiglaUMTrib.ListIndex = iIndice2 - 1

    End If

    Carrega_CombosUM = SUCESSO

    Exit Function

Erro_Carrega_CombosUM:

    Carrega_CombosUM = Err

    Select Case Err

        Case 64441

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165525)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Produto()

Dim iIndice As Integer

    'Chama Limpa_Tela
    Call Limpa_Tela(Me)

    'Limpa os Campos não limpos no limpa tela
    Ativo.Value = vbChecked
    IncideIPI.Value = False
    DescTipoProduto.Caption = ""
    DescSubst1.Caption = ""
    DescSubst2.Caption = ""
    
    DescricaoClasseUM.Caption = ""
    SiglaUMCompra.Clear
    SiglaUMEstoque.Clear
    SiglaUMVenda.Clear
    NomeUMCompra.Caption = ""
    NomeUMEstoque.Caption = ""
    NomeUMVenda.Caption = ""
    NaturezaProduto.ListIndex = -1
    Rastro.ListIndex = 0
    Referencia.Text = ""
    NomeFigura.Text = ""
    '''Embalagem.Text = "" 05/09/01 Marcelo
    SituacaoTributaria.ListIndex = -1
    comboAliquota.ListIndex = -1
    Figura.Picture = LoadPicture
    '''DescricaoEmbalagem.Caption = "" 05/09/01 Marcelo
    Grades.Text = ""
    DescISSQN.Caption = ""
    DescGenero.Caption = ""
    CodServNFe.Caption = ""
    NomeUMTrib.Caption = ""
    SiglaUMTrib.Clear
    ProdutoEspecifico.ListIndex = -1
    Fabricante.ListIndex = iFabricante_ListIndex_Padrao
    
    'ProdEmEscalaRelev.Value = vbGrayed
    Call Combo_Seleciona_ItemData(ProdEmEscalaRelev, -1)
    
    '***ALTERACAO POR TULIO EM 28/05***
    'limpa a combo de codigo de barras
    CodigoBarras.Clear
    '***FIM ALTERACAO POR TULIO EM 28/05***
    
    ConsideraQuantCotacaoAnterior.Value = vbUnchecked
    NaoTemFaixaReceb.Value = vbChecked
    
    'Limpa os Grids
    Call Grid_Limpa(objGridCategoria)
    Call Grid_Limpa(objGridTabelaPreco)

    'Limpa a lista de características
    For iIndice = 0 To ListaCaracteristicas.ListCount - 1
        ListaCaracteristicas.Selected(iIndice) = False
    Next
    
    If Produzido.Value = True Then ListaCaracteristicas.Selected(1) = True
    OrigemMercadoria.ListIndex = 0
    
    iClasseUMAlterada = 0
    iClasseUMAnterior = 0
    iTipoAlterado = 0
    
    SeriePartNum.Caption = ""
    
    Call CF2(Me, "Prod_Limpa_Tela")
    
    Set gobjAnotacao = Nothing
    
    Call gobjIdiomaTela.Reiniciar
    
    iProdutoAlterado = 0
    iAlterado = 0

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 64442

    'Limpa a Tela
    Call Limpa_Tela_Produto

    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 64442

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 165526)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colTabelaPrecoItem As New Collection
Dim iIndice As Integer
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim objCategoriaItem As ClassProdutoCategoria
Dim objProdutoFilial As New ClassProdutoFilial
Dim iControleEstoque As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios da Tela estão preenchidos
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 64443
    If Len(Trim(Descricao.Text)) = 0 Then gError 64444
    If Len(Trim(NomeReduzido.Text)) = 0 Then gError 64445
    If Len(Trim(ClasseUM.Text)) = 0 Then gError 64446
    If Len(Trim(SiglaUMCompra.Text)) = 0 Then gError 64447
    If Len(Trim(SiglaUMEstoque.Text)) = 0 Then gError 64448
    If Len(Trim(SiglaUMVenda.Text)) = 0 Then gError 64449
    If Len(Trim(SiglaUMTrib.Text)) = 0 Then gError 64449
    If Len(Trim(NaturezaProduto.List(NaturezaProduto.ListIndex))) = 0 Then gError 64450
    If Len(Trim(TipoProduto.ClipText)) = 0 Then gError 62671
    
    '#######################################################
    'Inserido por Wagner 08/03/2006
    If Len(Trim(SerieProx.Text)) <> 0 And StrParaInt(SerieNum.Text) = 0 Then gError 141797
    If Len(Trim(SerieProx.Text)) = 0 And StrParaInt(SerieNum.Text) <> 0 Then gError 141798
    '#######################################################
        
    'Verifica se o tipo é produzido
    If Produzido.Value = True Then
    
        'Verifica se foi informada a apropriação
        If ApropriacaoProd.ListIndex = -1 Then gError 64451
        
    Else
    
        'Verifica se foi informada a apropriação
        If ApropriacaoComp.ListIndex = -1 Then gError 64452
        
    End If
    
'    'Exige o preenchimento de pelo menos um dos dois campos Código de barras ou Referência para Produtos não gerenciais e de vendas
'    If (ListaCaracteristicas.Selected(0) = True) And (NivelGerencial = False) Then
'
'        If Len(Trim(CodigoBarras)) = 0 And Len(Trim(Referencia)) = 0 Then gError 81212
'
        'Exige o preenchimento da alíquota ICMS/ISS para produtos em que o tipo de tributação for Integral ou Subst Tributária
        If SituacaoTributaria.ListIndex > -1 Then
            If SituacaoTributaria.ItemData(SituacaoTributaria.ListIndex) = TIPOTRIBICMS_INTEGRAL _
             Or SituacaoTributaria.ItemData(SituacaoTributaria.ListIndex) = TIPOTRIBISS_INTEGRAL Then
    
                If comboAliquota.ListIndex = -1 Then gError 81213
    
            End If
        End If
'    End If
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objProduto, colTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 64453
    
    If objProduto.iKitVendaComp = MARCADO And objProduto.iGerencial = DESMARCADO Then gError 177577

    lErro = Trata_Alteracao(objProduto, objProduto.sCodigo)
    If lErro <> SUCESSO Then Error 32306

    lErro = CF2(Me, "Prod_Gravar_Registro", objProduto)
    If lErro <> SUCESSO Then Error 32306
    
    'Verifica se alguma data de Preço não foi preenchida
    For Each objTabelaPrecoItem In colTabelaPrecoItem
        iIndice = iIndice + 1
        If objTabelaPrecoItem.dPreco > 0 Then If objTabelaPrecoItem.dtDataVigencia = DATA_NULA Then gError 64454
    Next

    For Each objCategoriaItem In objProduto.colCategoriaItem
        If Len(Trim(objCategoriaItem.sCategoria)) = 0 Then gError 64455
        If Len(Trim(objCategoriaItem.sItem)) = 0 Then gError 64456
    Next
        
    'Grava o Produto
    lErro = CF("Produto_Grava", objProduto, colTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 64457
    
    iIndice = 0
 
'    'Atualiza as mudanças da Arvore de Produtos
'    lErro = Atualiza_Arvore_Produto(objProduto)
'    If lErro <> SUCESSO Then gError 64458

    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PRODUTO
            gobjAnotacao.sID = objProduto.sCodigo
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 141373
            
        End If
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 32306, 141373

        Case 62671
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_TIPOPRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 64443
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 64444
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_PRODUTO_NAO_INFORMADA", gErr)

        Case 64445
            Call Rotina_Erro(vbOKOnly, "ERRO_NOMEREDUZIDO_PRODUTO_NAO_INFORMADO", gErr)

        Case 64446
            Call Rotina_Erro(vbOKOnly, "ERRO_CLASSEUM_NAO_INFORMADA", gErr)

        Case 64447
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_COMPRA_NAO_INFORMADA", gErr)

        Case 64448
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_ESTOQUE_NAO_INFORMADA", gErr)

        Case 64449
            Call Rotina_Erro(vbOKOnly, "ERRO_UM_VENDA_NAO_INFORMADA", gErr)

        Case 64450
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZA_PRODUTO_NAO_PREENCHIDA", gErr)
                 
        Case 64451, 64452
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_APROPRIACAO_NAO_PREENCHIDA", gErr)
            
        Case 64453, 64457, 64458 'Tratados nas rotinas chamadas

        Case 64454
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_PRECO_NAO_PREENCHIDA", gErr, objTabelaPrecoItem.iCodTabela)

        Case 64455
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", gErr)

        Case 64456
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCH_CAMPOS_IDENTIFICACAO", gErr)
            
        Case 81212
            Call Rotina_Erro(vbOKOnly, "ERRO_CODBARRAS_OU_REFERENCIA_PREENCH_OBRIGATORIOS", gErr)
            
        Case 81213
            Call Rotina_Erro(vbOKOnly, "ERRO_PREENCH_ICMS", gErr)

        Case 141797
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_NAO_PREENCHIDO", gErr)
            SerieProx.SetFocus
        
        Case 141798
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIENUM_NAO_PREENCHIDO", gErr)
            SerieNum.SetFocus
            
        Case 177577
            Call Rotina_Erro(vbOKOnly, "ERRO_KITVENDA_NAO_GERENCIAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165527)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objProduto As ClassProduto, colTabelaPrecoItem As Collection) As Long

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim colTabelas As New Collection
Dim iNivel As Integer
Dim sVerifica As String

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64459

        If iPreenchido = PRODUTO_VAZIO Then Error 64460

        objProduto.sCodigo = sProduto
        
        'Obtem o nível do Produto
        lErro = Mascara_Produto_ObterNivel(sProduto, iNivel)
        If lErro <> SUCESSO Then Error 64461

        objProduto.iNivel = iNivel
        
    End If
    
    'Recolhe os demais dados
    objProduto.sDescricao = Descricao.Text
    objProduto.sNomeReduzido = NomeReduzido.Text
    objProduto.sModelo = Modelo.Text
    objProduto.sFigura = NomeFigura.Text
    objProduto.sReferencia = Referencia.Text
    objProduto.sGrade = Grades.Text
   
    If Len(Trim(NaturezaProduto.List(NaturezaProduto.ListIndex))) > 0 Then objProduto.iNatureza = NaturezaProduto.ItemData(NaturezaProduto.ListIndex)
    
    If Ativo.Value = vbUnchecked Then objProduto.iAtivo = PRODUTO_INATIVO
    
    If Len(Trim(TipoProduto.Text)) > 0 Then objProduto.iTipo = CInt(TipoProduto.Text)
    
    If NivelGerencial.Value = True Then
        objProduto.iGerencial = GERENCIAL
    Else
        objProduto.iGerencial = NAO_GERENCIAL
    End If

    'Move o que está em Lista de Características para o objeto
    If ListaCaracteristicas.Selected(0) = True Then objProduto.iFaturamento = 1
    If ListaCaracteristicas.Selected(1) = True Then objProduto.iPCP = 1
    If ListaCaracteristicas.Selected(2) = True Then objProduto.iKitBasico = 1
    If ListaCaracteristicas.Selected(3) = True Then objProduto.iKitInt = 1
    If ListaCaracteristicas.Selected(4) = True Then objProduto.iKitVendaComp = 1
    
    If NivelGerencial.Enabled = False And objProduto.iKitVendaComp = 1 Then objProduto.iGerencial = GERENCIAL
    
    If Fabricante.ListIndex <> -1 Then objProduto.lFabricante = Fabricante.ItemData(Fabricante.ListIndex)

    If Produzido.Value = True Then
    
        objProduto.iCompras = PRODUTO_PRODUZIVEL
        'Move o que esta selecionado na Combobox Apropriação para o objeto
        If ApropriacaoProd.ListIndex <> -1 Then objProduto.iApropriacaoCusto = ApropriacaoProd.ItemData(ApropriacaoProd.ListIndex)
        
    Else
    
        objProduto.iCompras = PRODUTO_COMPRAVEL
        If ApropriacaoComp.ListIndex <> -1 Then objProduto.iApropriacaoCusto = ApropriacaoComp.ItemData(ApropriacaoComp.ListIndex)
        
    End If
    
    If gobjEST.iProdExibeComissao = MARCADO Then
        If Len(Trim(PercComissao.Text)) > 0 Then objProduto.dPercComissao = CDbl(PercComissao.Text) / 100
        objProduto.dMetaComissao = StrParaDbl(MetaComissao.Text)
    End If
            
    lErro = Move_TabComplemento_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 64462

    lErro = Move_TabCaracteristicas_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 64463
    
    lErro = Move_TabUnidadeMed_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 64464

    lErro = Move_TabTributacao_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 64465
    
    lErro = Move_TabCompras_Memoria(objProduto)
    If lErro <> SUCESSO Then Error 64466
    
    Set objProduto.objIdiomaTela = gobjIdiomaTela
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
    
        Case 64459, 64462, 64463, 64464, 64465, 64466
        
        Case 64460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 64461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_PRODUTO_OBTERNIVEL", Err, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165528)

    End Select

    Exit Function

End Function

Private Function Move_TabComplemento_Memoria(objProduto As ClassProduto)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objProdutoCategoria As ClassProdutoCategoria

On Error GoTo Erro_Move_TabComplemento_Memoria

    'Verifica se algum produto Substituto foi preenchido
    If Len(Trim(Substituto1.ClipText)) > 0 Then
    
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Substituto1.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64467
        
        'Guarda em objProduto
        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sSubstituto1 = sProduto

    End If
    
    'Verifica se algum produto Substituto foi preenchido
    If Len(Trim(Substituto2.ClipText)) > 0 Then
    
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Substituto2.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64468
        
        'Guarda em objProduto
        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sSubstituto2 = sProduto

    End If

    objProduto.sCodigoBarras = CodigoBarras.Text
    
    If Len(Trim(PrazoValidade.Text)) > 0 Then objProduto.iPrazoValidade = CInt(PrazoValidade.Text)
    If Len(Trim(EtiquetasCodBarras.Text)) > 0 Then objProduto.iEtiquetasCodBarras = CInt(EtiquetasCodBarras.Text)
    
    If Len(Trim(Residuo.Text)) > 0 Then
        objProduto.dResiduo = PercentParaDbl(Residuo.Text & "%")
    Else
        objProduto.dResiduo = -1
    End If
    
    If Len(Trim(TempoProducao.Text)) > 0 Then
        objProduto.iTempoProducao = StrParaInt(TempoProducao.Text)
    End If
    
    'ir preenchendo uma coleção com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridCategoria.iLinhasExistentes
    
        'Verifica se a Categoria foi preenchida
        If Len(Trim(GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col))) <> 0 Then

            Set objProdutoCategoria = New ClassProdutoCategoria

            If Len(Codigo.Text) > 0 Then objProdutoCategoria.sProduto = Codigo.Text
            
            objProdutoCategoria.sCategoria = GridCategoria.TextMatrix(iIndice, iGrid_Categoria_Col)
            objProdutoCategoria.sItem = GridCategoria.TextMatrix(iIndice, iGrid_Valor_Col)

            objProduto.colCategoriaItem.Add objProdutoCategoria

        End If
        
    Next
    
    objProduto.dPrecoMaxConsumidor = StrParaDbl(PrecoMaxConsumidor.Text)
    objProduto.iProdutoEspecifico = Codigo_Extrai(ProdutoEspecifico.Text)
    objProduto.sISSQN = Trim(ISSQN.Text)
    objProduto.sGenero = Trim(Genero.Text)
    objProduto.iExTIPI = StrParaInt(ExTIPI.ClipText)
    
    'Move o custo
    If Len(Trim(CustoReposicao.Text)) > 0 Then objProduto.dCustoReposicao = CDbl(CustoReposicao.FormattedText)
    
    '***ALTERACAO POR TULIO EM 27/05***
    'passa para a colecao de codigo de barras que existe em objproduto
    'os codigos de barras da combo de codigo de barras
    Call ObterCodigosBarra(objProduto.colCodBarras)

    '***FIM ALTERACAO POR TULIO EM 27/05***
    
    '####################################################
    'Inserido por Wagner 08/03/2006
    objProduto.sSerieProx = SerieProx.Text
    objProduto.iSerieParteNum = StrParaInt(SerieNum.Text)
    '####################################################
    
    Move_TabComplemento_Memoria = SUCESSO

    Exit Function

Erro_Move_TabComplemento_Memoria:

    Move_TabComplemento_Memoria = Err

    Select Case Err

        Case 64467, 64468

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165529)

    End Select

    Exit Function

End Function

Private Function Move_TabCaracteristicas_Memoria(objProduto As ClassProduto) As Long
'Recolhe os dados do Tab de Características Físicas

    If Len(Trim(PesoBruto.Text)) > 0 Then objProduto.dPesoBruto = CDbl(PesoBruto.Text)
    If Len(Trim(PesoLiquido.Text)) > 0 Then objProduto.dPesoLiq = CDbl(PesoLiquido.Text)
    
    '''objProduto.iEmbalagem = StrParaInt(Embalagem.Text) 05/09/01 Marcelo
    objProduto.sObsFisica = ObsFisica.Text
    objProduto.sCor = Cor.Text
    
    If Len(Trim(PesoEspecifico.Text)) > 0 Then objProduto.dPesoEspecifico = CDbl(PesoEspecifico.Text)
    If Len(Trim(HorasMaquina.Text)) > 0 Then objProduto.lHorasMaquina = StrParaLong(HorasMaquina.Text)
    'Move o que esta selecionado na Combobox Rastro para o objeto
    If Rastro.ListIndex <> -1 Then objProduto.iRastro = Rastro.ItemData(Rastro.ListIndex)
                  
    Call CF2(Me, "Prod_MoveCaracMemoria", objProduto)
    
    Move_TabCaracteristicas_Memoria = SUCESSO

End Function

Private Function Move_GridTabelaPreco_Memoria(colTabelaPrecoItem As Collection, objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim iColunaGrid As Integer

    If giFilialEmpresa = EMPRESA_TODA Then
        iColunaGrid = iGrid_ValorEmpresa_Col
    Else
        iColunaGrid = iGrid_ValorFilial_Col
    End If

    'Guarda os dados do Grid de Tabelas de Preço em colTabelaPrecoItem
    For iIndice = 1 To objGridTabelaPreco.iLinhasExistentes

        Set objTabelaPrecoItem = New ClassTabelaPrecoItem

        If Len(Trim(GridTabelaPreco.TextMatrix(iIndice, iColunaGrid))) > 0 Then
        
            objTabelaPrecoItem.dPreco = CDbl(GridTabelaPreco.TextMatrix(iIndice, iColunaGrid))
            objTabelaPrecoItem.iCodTabela = GridTabelaPreco.TextMatrix(iIndice, iGrid_Tabela_Col)
            objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
            objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
            
            If Len(Trim(GridTabelaPreco.TextMatrix(iIndice, iGrid_DataPreco_Col))) > 0 Then
                objTabelaPrecoItem.dtDataVigencia = CDate(GridTabelaPreco.TextMatrix(iIndice, iGrid_DataPreco_Col))
            Else
                objTabelaPrecoItem.dtDataVigencia = DATA_NULA
            End If

            colTabelaPrecoItem.Add objTabelaPrecoItem
            
        End If

    Next

    Move_GridTabelaPreco_Memoria = SUCESSO

End Function

Private Function Move_TabUnidadeMed_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long

On Error GoTo Erro_Move_TabUnidadeMed_Memoria
   
    'Guarda a Classe e  as Unidades de Medidas selecionadas
    If Len(Trim(ClasseUM.Text)) > 0 Then objProduto.iClasseUM = CInt(ClasseUM)
    
    objProduto.sSiglaUMCompra = SiglaUMCompra.Text
    objProduto.sSiglaUMEstoque = SiglaUMEstoque.Text
    objProduto.sSiglaUMVenda = SiglaUMVenda.Text
    objProduto.sSiglaUMTrib = SiglaUMTrib.Text

    Move_TabUnidadeMed_Memoria = SUCESSO

    Exit Function

Erro_Move_TabUnidadeMed_Memoria:

    Move_TabUnidadeMed_Memoria = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165530)

    End Select

    Exit Function

End Function

Private Function Move_TabTributacao_Memoria(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_TabTributacao_Memoria

    'Verifica se o IncideIPI esta selecionado
    If IncideIPI.Value = vbChecked Then
    
        objProduto.iIPIIncide = 1
        
        'Recolhe os dados relacionadosao IPI
        If Len(Trim(AliquotaIPI.Text)) > 0 Then objProduto.dIPIAliquota = CDbl(AliquotaIPI / 100)
        
        objProduto.sIPICodDIPI = CodigoIPI.Text
        
    Else
    
        objProduto.iIPIIncide = 0
    
    End If
    
    objProduto.sIPICodigo = Trim(ClasFiscIPI.Text)
    objProduto.sCEST = Trim(CEST.Text)
    
    objProduto.scProdANVISA = Trim(cProdANVISA.Text)
    
'    If ProdEmEscalaRelev.Value = vbChecked Then
'        objProduto.iProdEmEscalaRelev = MARCADO
'    ElseIf ProdEmEscalaRelev.Value = vbUnchecked Then
'        objProduto.iProdEmEscalaRelev = DESMARCADO
'    Else
'        objProduto.iProdEmEscalaRelev = -1
'    End If
    objProduto.iProdEmEscalaRelev = ProdEmEscalaRelev.ItemData(ProdEmEscalaRelev.ListIndex)
    
    'Recolhe os dados relacionados ao ICMS
    If comboAliquota.ListIndex <> -1 Then objProduto.sICMSAliquota = SCodigo_Extrai(comboAliquota.List(comboAliquota.ListIndex))
        
    'Recolhe os dados relacionados ao kit de venda
'''    If KitVendaComp.ListIndex <> -1 Then
'''        objProduto.iKitVendaComp = KitVendaComp.ItemData(KitVendaComp.ListIndex)
'''    End If
    
    'Verifica se a Conta Contábil foi informada
    If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 64469
        
        objProduto.sContaContabil = sContaFormatada
        
    End If

    'Verifica se a Conta Produção foi informada
    If Len(Trim(ContaProducao.ClipText)) > 0 Then
    
        'Guarda a conta corrente
        lErro = CF("Conta_Formata", ContaProducao.Text, sContaFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then Error 64470
        
        objProduto.sContaContabilProducao = sContaFormatada
        
    End If
    
    objProduto.iOrigemMercadoria = OrigemMercadoria.ItemData(OrigemMercadoria.ListIndex)
    
    If SituacaoTributaria.ListIndex > -1 Then
        
        Select Case SituacaoTributaria.ItemData(SituacaoTributaria.ListIndex)
            
            Case TIPOTRIBICMS_NAO_TRIBUTADO
                objProduto.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_NAO_TRIBUTADO
                            
            Case TIPOTRIBICMS_ISENTA
                objProduto.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_ISENTA
             
             Case TIPOTRIBICMS_TRIB_SUBST
                objProduto.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_TRIB_SUBST
            
            Case TIPOTRIBICMS_INTEGRAL
                objProduto.sSituacaoTribECF = TIPOTRIBICMS_SITUACAOTRIBECF_INTEGRAL
        
            Case TIPOTRIBISS_NAO_TRIBUTADO
                objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_NAO_TRIBUTADO
                            
            Case TIPOTRIBISS_ISENTA
                objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_ISENTA
             
             Case TIPOTRIBISS_TRIB_SUBST
                objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_TRIB_SUBST
            
            Case TIPOTRIBISS_INTEGRAL
                objProduto.sSituacaoTribECF = TIPOTRIBISS_SITUACAOTRIBECF_INTEGRAL
        
        End Select
        
    End If
        
    'Recolhe os dados relacionados ao INSS
    If Len(Trim(INSSPercBase.Text)) > 0 Then objProduto.dINSSPercBase = CDbl(INSSPercBase.Text / 100)
    
    If gobjMAT.iProdFatorAjuste <> 0 Then
        If Len(Trim(FatorAjuste.Text)) > 0 Then objProduto.dFatorAjuste = CDbl(FatorAjuste.Text / 100)
    End If
    
    objProduto.sFCI = Trim(FCI.Text)
    
    'Verifica se o IncideIPI esta selecionado
    If UsaBalanca.Value = vbChecked Then
        objProduto.iUsaBalanca = USA_BALANCA
    Else
        objProduto.iUsaBalanca = NAO_USA_BALANCA
    End If
        
    Move_TabTributacao_Memoria = SUCESSO

    Exit Function

Erro_Move_TabTributacao_Memoria:

    Move_TabTributacao_Memoria = Err

    Select Case Err

        Case 64469, 64470

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165531)

    End Select

    Exit Function

End Function

Public Sub ComboCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProduto
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ComboCategoriaProdutoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboCategoriaProdutoItem_GotFocus()

Dim lErro As Long

On Error GoTo Erro_ComboCategoriaProdutoItem_GotFocus

    'Preenche com os ítens relacionados a Categoria correspondente
    Call Trata_ComboCategoriaProdutoItem

    Call Grid_Campo_Recebe_Foco(objGridCategoria)

    Exit Sub

Erro_ComboCategoriaProdutoItem_GotFocus:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165532)

    End Select

    Exit Sub

End Sub

Public Sub ComboCategoriaProdutoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCategoria)

End Sub

Public Sub ComboCategoriaProdutoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCategoria.objControle = ComboCategoriaProdutoItem
    lErro = Grid_Campo_Libera_Foco(objGridCategoria)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Trata_ComboCategoriaProdutoItem()

Dim lErro As Long
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim iIndice As Integer
Dim sValor As String

On Error GoTo Erro_Trata_ComboCategoriaProdutoItem

    sValor = ComboCategoriaProdutoItem.Text

    ComboCategoriaProdutoItem.Clear

    ComboCategoriaProdutoItem.Text = sValor

    'Se alguém estiver selecionado
    If Len(GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)) > 0 Then

        'Preencher a Combo de Itens desta Categoria
        objCategoriaProduto.sCategoria = GridCategoria.TextMatrix(GridCategoria.Row, iGrid_Categoria_Col)

        lErro = Carrega_ComboCategoriaProdutoItem(objCategoriaProduto)
        If lErro <> SUCESSO Then Error 64471

        For iIndice = 0 To ComboCategoriaProdutoItem.ListCount - 1
            If ComboCategoriaProdutoItem.List(iIndice) = GridCategoria.Text Then
                ComboCategoriaProdutoItem.ListIndex = iIndice
                Exit For
            End If
        Next

    End If

    Exit Sub

Erro_Trata_ComboCategoriaProdutoItem:

    Select Case Err

        Case 64471

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165533)

    End Select
    
    Exit Sub

End Sub

Private Function Carrega_ComboCategoriaProdutoItem(objCategoriaProduto As ClassCategoriaProduto) As Long
'Carrega o Item da Categoria na Combobox

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_Carrega_ComboCategoriaProdutoItem

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colItensCategoria)
    If lErro <> SUCESSO Then Error 64472

    'Insere na combo CategoriaProdutoItem
    For Each objCategoriaProdutoItem In colItensCategoria

        'Insere na combo CategoriaProduto
        ComboCategoriaProdutoItem.AddItem objCategoriaProdutoItem.sItem

    Next

    Carrega_ComboCategoriaProdutoItem = SUCESSO

    Exit Function

Erro_Carrega_ComboCategoriaProdutoItem:

    Carrega_ComboCategoriaProdutoItem = Err

    Select Case Err

        Case 64472

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165534)

    End Select

    Exit Function

End Function

'Raphael - 01/00: Estes campos deixam de ser editaveis no grid de Tabela de Preço
'Private Function Saida_Celula_DataPreco(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_Saida_Celula_DataPreco
'
'    Set objGridInt.objControle = DataPreco
'
'    'Verifica se a Data esta preenchida
'    If Len(Trim(DataPreco.ClipText)) > 0 Then
'
'        'Critica a data
'        lErro = Data_Critica(DataPreco.Text)
'        If lErro <> SUCESSO Then Error 18420
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then Error 18425
'
'    Saida_Celula_DataPreco = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_DataPreco:
'
'    Saida_Celula_DataPreco = Err
'
'    Select Case Err
'
'        Case 18420, 18425
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165535)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Saida_Celula_ValorProduto(objGridInt As AdmGrid) As Long
''Faz a crítica da celula Valor do grid que está deixando de ser a corrente
'
'Dim lErro As Long
'Dim dColunaSoma As Double
'
'On Error GoTo Erro_Saida_Celula_ValorProduto
'
'    If GridTabelaPreco.Col = iGrid_ValorEmpresa_Col Then
'        Set objGridInt.objControle = ValorEmpresa
'    Else
'        Set objGridInt.objControle = ValorFilial
'    End If
'
'    'Verifica se valor está preenchido
'    If Len(objGridInt.objControle.ClipText) > 0 Then
'
'        'Critica se valor é positivo
'        lErro = Valor_Positivo_Critica(objGridInt.objControle.Text)
'        If lErro <> SUCESSO Then Error 18426
'
'    End If
'
'    lErro = Grid_Abandona_Celula(objGridInt)
'    If lErro <> SUCESSO Then Error 18427
'
'    Saida_Celula_ValorProduto = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_ValorProduto:
'
'    Saida_Celula_ValorProduto = Err
'
'    Select Case Err
'
'        Case 18426, 18427
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165536)
'
'    End Select
'
'    Exit Function
'
'End Function

'Raphael - 01/00: Estes campos deixam de ser editaveis no grid de Tabela de Preço
'Private Sub DataPreco_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub DataPreco_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridTabelaPreco)
'
'End Sub
'
'Private Sub DataPreco_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTabelaPreco)
'
'End Sub
'
'Private Sub DataPreco_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridTabelaPreco.objControle = DataPreco
'    lErro = Grid_Campo_Libera_Foco(objGridTabelaPreco)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Private Sub ValorFilial_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ValorFilial_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridTabelaPreco)
'
'End Sub
'
'Private Sub ValorFilial_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTabelaPreco)
'
'End Sub
'
'Private Sub ValorFilial_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridTabelaPreco.objControle = ValorFilial
'    lErro = Grid_Campo_Libera_Foco(objGridTabelaPreco)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub
'
'Private Sub ValorEmpresa_Change()
'
'    iAlterado = REGISTRO_ALTERADO
'
'End Sub
'
'Private Sub ValorEmpresa_GotFocus()
'
'    Call Grid_Campo_Recebe_Foco(objGridTabelaPreco)
'
'End Sub
'
'Private Sub ValorEmpresa_KeyPress(KeyAscii As Integer)
'
'    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTabelaPreco)
'
'End Sub
'
'Private Sub ValorEmpresa_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'
'    Set objGridTabelaPreco.objControle = ValorEmpresa
'    lErro = Grid_Campo_Libera_Foco(objGridTabelaPreco)
'    If lErro <> SUCESSO Then Cancel = True
'
'End Sub

'Function Alterar_Arvore_Produto(colNodes As Nodes, objProduto As ClassProduto) As Long
''Faz a alteração na árvore do Produto passado
'
'Dim objNode As Node
'Dim sProduto As String
'Dim sProdutoMascarado As String
'Dim lErro As Long
'Dim iAchou As Integer
'
'On Error GoTo Erro_Alterar_Arvore_Produto
'
'    sProduto = "X" & objProduto.sCodigo
'
'    iAchou = 0
'
'    For Each objNode In colNodes
'
'        If objNode.Key = sProduto Then
'
'            sProdutoMascarado = String(STRING_PRODUTO, 0)
'
'            'coloca o Produto no formato que é exibido na tela
'            lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
'            If lErro <> SUCESSO Then Error 64473
'
'            objNode.Text = sProdutoMascarado & SEPARADOR & objProduto.sDescricao
'
'            iAchou = 1
'
'            Exit For
'
'        End If
'
'    Next
'
'    'se não achou o Produto na árvore
'    If iAchou = 0 Then Error 64474
'
'    Alterar_Arvore_Produto = SUCESSO
'
'    Exit Function
'
'Erro_Alterar_Arvore_Produto:
'
'    Alterar_Arvore_Produto = Err
'
'    Select Case Err
'
'        Case 64473
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)
'
'        Case 64474
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165537)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Inserir_Arvore_Produto(colNodes As Nodes, objProduto As ClassProduto) As Long
''insere a Produto na lista de Produtos
'
'Dim objNode As Node
'Dim lErro As Long
'Dim sProdutoMascarado As String
'Dim sProduto As String
'Dim sProdutoPai As String
'Dim sProdutoAvo As String
'Dim iAchou As Integer
'
'On Error GoTo Erro_Inserir_Arvore_Produto
'
'    'verifica se o ramo da arvore a que se refere o nó em questão já está carregado na arvore
'    'se estiver, pode inserir. Se não estiver, não pode inserir.
'    'os niveis 1 e 2 estão sempre na arvore.
'    'se o nivel do produto for maior do que 2 ==> testar se o avo indica a carga dos netos. Se não indicar, não inserir o no.
'    If objProduto.iNivel > 2 Then
'
'        sProdutoAvo = String(STRING_CONTA, 0)
'
'        lErro = Mascara_RetornaProdutoNoNivel(objProduto.iNivel - 2, objProduto.sCodigo, sProdutoAvo)
'        If lErro <> SUCESSO Then Error 64475
'
'        sProduto = "X" & sProdutoAvo
'
'        iAchou = 0
'
'        For Each objNode In colNodes
'            If objNode.Key = sProduto Then
'                iAchou = 1
'                'se os netos do avo do elemento em questão não estão na arvore ==> não pode inserir o elemento na arvore
'                If objNode.Tag <> NETOS_NA_ARVORE Then Error 64476
'                Exit For
'            End If
'        Next
'
'        'o avo do nó em questão ainda não está carregado ==> não pode inserir o elemento na árvore
'        If iAchou = 0 Then Error 64477
'
'    End If
'
'    sProdutoMascarado = String(STRING_PRODUTO, 0)
'
'    'coloca o Produto no formato que é exibido na tela
'    lErro = Mascara_MascararProduto(objProduto.sCodigo, sProdutoMascarado)
'    If lErro <> SUCESSO Then Error 64478
'
'    sProduto = "X" & objProduto.sCodigo
'
'    sProdutoPai = String(STRING_PRODUTO, 0)
'
'    'retorna a Produto "pai" da Produto em questão, se houver
'    lErro = Mascara_RetornaProdutoPai(objProduto.sCodigo, sProdutoPai)
'    If lErro <> SUCESSO Then Error 64479
'
'    'se o Produto possui um Produto "pai"
'    If Len(Trim(sProdutoPai)) > 0 Then
'
'        sProdutoPai = "X" & sProdutoPai
'
'        Set objNode = colNodes.Add(colNodes.Item(sProdutoPai), tvwChild, sProduto, sProdutoMascarado & SEPARADOR & objProduto.sDescricao)
'        colNodes.Item(sProdutoPai).Sorted = True
'
'    Else
'        'se o Produto não possui Produto "pai"
'        Set objNode = colNodes.Add(, , sProduto, sProdutoMascarado & SEPARADOR & objProduto.sDescricao)
'        TvwProduto.Sorted = True
'    End If
'
'    Inserir_Arvore_Produto = SUCESSO
'
'    Exit Function
'
'Erro_Inserir_Arvore_Produto:
'
'    Inserir_Arvore_Produto = Err
'
'    Select Case Err
'
'        Case 64475, 64476, 64477
'
'        Case 64478
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProduto.sCodigo)
'
'        Case 64479
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOPAI", Err, objProduto.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165538)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Exclui_Arvore_Produto(colNodes As Nodes, objProduto As ClassProduto) As Long
''Exclui o Produto da árvore
'
'Dim objNode As Node
'Dim sProduto As String
'
'    sProduto = "X" & objProduto.sCodigo
'
'    'Procura pelo produto e o retira da árvore
'    For Each objNode In colNodes
'
'        If objNode.Key = sProduto Then
'            colNodes.Remove (objNode.Index)
'            Exit For
'        End If
'
'    Next
'
'    Exclui_Arvore_Produto = SUCESSO
'
'End Function

'Function Atualiza_Arvore_Produto(objProduto As ClassProduto) As Long
''Atualiza a árvore de Produtos após a rotina de gravação
'
'Dim lErro As Long
'
'On Error GoTo Erro_Atualiza_Arvore_Produto
'
'    'Tenta alterar o Produto
'    lErro = Alterar_Arvore_Produto(TvwProduto.Nodes, objProduto)
'    If lErro <> SUCESSO And lErro <> 64474 Then Error 64480
'
'    'Se ele não existir na árvore
'    If lErro = 64474 Then
'        'Insere o Produto na árvore
'        lErro = Inserir_Arvore_Produto(TvwProduto.Nodes, objProduto)
'        If lErro <> SUCESSO Then Error 64481
'
'    End If
'
'    Atualiza_Arvore_Produto = SUCESSO
'
'    Exit Function
'
'Erro_Atualiza_Arvore_Produto:
'
'    Atualiza_Arvore_Produto = Err
'
'    Select Case Err
'
'        Case 64480, 64481
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165539)
'
'    End Select
'
'    Exit Function
'
'End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colTabelaPrecoItem As New Collection

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Produtos"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objProduto, colTabelaPrecoItem)
    If lErro <> SUCESSO Then Error 64482

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProduto.sCodigo, STRING_PRODUTO, "Codigo"
    colCampoValor.Add "Tipo", objProduto.iTipo, 0, "Tipo"
    colCampoValor.Add "Descricao", objProduto.sDescricao, STRING_PRODUTO_DESCRICAO, "Descricao"
    colCampoValor.Add "NomeReduzido", objProduto.sNomeReduzido, STRING_PRODUTO_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "Modelo", objProduto.sModelo, STRING_PRODUTO_MODELO, "Modelo"
    colCampoValor.Add "Gerencial", objProduto.iGerencial, 0, "Gerencial"
    colCampoValor.Add "Nivel", objProduto.iNivel, 0, "Nivel"
    colCampoValor.Add "Substituto1", objProduto.sSubstituto1, STRING_PRODUTO_SUBSTITUTO1, "Substituto1"
    colCampoValor.Add "Substituto2", objProduto.sSubstituto2, STRING_PRODUTO_SUBSTITUTO2, "Substituto2"
    colCampoValor.Add "PrazoValidade", objProduto.iPrazoValidade, 0, "PrazoValidade"
    colCampoValor.Add "CodigoBarras", objProduto.sCodigoBarras, STRING_PRODUTO_CODIGO_BARRAS, "CodigoBarras"
    colCampoValor.Add "EtiquetasCodBarras", objProduto.iEtiquetasCodBarras, 0, "EtiquetasCodBarras"
    colCampoValor.Add "PesoLiq", objProduto.dPesoLiq, 0, "PesoLiq"
    colCampoValor.Add "PesoBruto", objProduto.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "Comprimento", objProduto.dComprimento, 0, "Comprimento"
    colCampoValor.Add "Espessura", objProduto.dEspessura, 0, "Espessura"
    colCampoValor.Add "Largura", objProduto.dLargura, 0, "Largura"
    colCampoValor.Add "Cor", objProduto.sCor, STRING_PRODUTO_COR, "Cor"
    colCampoValor.Add "ObsFisica", objProduto.sObsFisica, STRING_PRODUTO_OBS_FISICA, "ObsFisica"
    colCampoValor.Add "ClasseUM", objProduto.iClasseUM, 0, "ClasseUM"
    colCampoValor.Add "SiglaUMCompra", objProduto.sSiglaUMCompra, STRING_PRODUTO_SIGLAUMCOMPRA, "SiglaUMCompra"
    colCampoValor.Add "SiglaUMEstoque", objProduto.sSiglaUMEstoque, STRING_PRODUTO_SIGLAUMESTOQUE, "SiglaUMEstoque"
    colCampoValor.Add "SiglaUMVenda", objProduto.sSiglaUMVenda, STRING_PRODUTO_SIGLAUMVENDA, "SiglaUMVenda"
    colCampoValor.Add "Ativo", objProduto.iAtivo, 0, "Ativo"
    colCampoValor.Add "Faturamento", objProduto.iFaturamento, 0, "Faturamento"
    colCampoValor.Add "Compras", objProduto.iCompras, 0, "Compras"
    colCampoValor.Add "PCP", objProduto.iPCP, 0, "PCP"
    colCampoValor.Add "KitBasico", objProduto.iKitBasico, 0, "KitBasico"
    colCampoValor.Add "KitInt", objProduto.iKitInt, 0, "KitInt"
    colCampoValor.Add "IPIAliquota", objProduto.dIPIAliquota, 0, "IPIAliquota"
    colCampoValor.Add "IPICodigo", objProduto.sIPICodigo, STRING_PRODUTO_IPI_CODIGO, "IPICodigo"
    colCampoValor.Add "IPICodDIPI", objProduto.sIPICodDIPI, STRING_PRODUTO_IPI_COD_DIPI, "IPICodDIPI"
    colCampoValor.Add "Apropriacao", objProduto.iApropriacaoCusto, 0, "Apropriacao"
    colCampoValor.Add "ContaContabil", objProduto.sContaContabil, STRING_CONTA, "ContaContabil"
    colCampoValor.Add "Natureza", objProduto.iNatureza, 0, "Natureza"
    colCampoValor.Add "ContaContabilProducao", objProduto.sContaContabilProducao, STRING_CONTA, "ContaContabilProducao"
    colCampoValor.Add "PercentMaisQuantCotAnt", objProduto.dPercentMaisQuantCotAnt, 0, "PercentMaisQuantCotAnt"
    colCampoValor.Add "PercentMenosQuantCotAnt", objProduto.dPercentMenosQuantCotAnt, 0, "PercentMenosQuantCotAnt"
    colCampoValor.Add "ConsideraQuantCotAnt", objProduto.iConsideraQuantCotAnt, 0, "ConsideraQuantCotAnt"
    colCampoValor.Add "TemFaixaReceb", objProduto.iTemFaixaReceb, 0, "TemFaixaReceb"
    colCampoValor.Add "PercentMaisReceb", objProduto.dPercentMaisReceb, 0, "PercentMaisReceb"
    colCampoValor.Add "PercentMenosReceb", objProduto.dPercentMenosReceb, 0, "PercentMenosReceb"
    colCampoValor.Add "RecebForaFaixa", objProduto.iRecebForaFaixa, 0, "RecebForaFaixa"
    colCampoValor.Add "CreditoICMS", objProduto.iCreditoICMS, 0, "CreditoICMS"
    colCampoValor.Add "CreditoIPI", objProduto.iCreditoIPI, 0, "CreditoIPI"
    colCampoValor.Add "Residuo", objProduto.dResiduo, 0, "Residuo"
    colCampoValor.Add "CustoReposicao", objProduto.dCustoReposicao, 0, "CustoReposicao"
    colCampoValor.Add "OrigemMercadoria", objProduto.iOrigemMercadoria, 0, "OrigemMercadoria"
    colCampoValor.Add "TempoProducao", objProduto.iTempoProducao, 0, "TempoProducao"
    colCampoValor.Add "Rastro", objProduto.iRastro, 0, "Rastro"
    colCampoValor.Add "HorasMaquina", objProduto.lHorasMaquina, 0, "HorasMaquina"
    colCampoValor.Add "PesoEspecifico", objProduto.dPesoEspecifico, 0, "PesoEspecifico"
    '''colCampoValor.Add "Embalagem", objProduto.iEmbalagem, 0, "Embalagem" 05/09/01 Marcelo
    colCampoValor.Add "Figura", objProduto.sFigura, STRING_NOME_ARQ_COMPLETO, "Figura"
    colCampoValor.Add "Referencia", objProduto.sReferencia, STRING_PRODUTO_REFERENCIA, "Referencia"
    colCampoValor.Add "INSSPercBase", objProduto.dINSSPercBase, 0, "INSSPercBase"
''''''    colCampoValor.Add "KitVendaComp", objProduto.iKitVendaComp, 0, "KitVendaComp"
    colCampoValor.Add "Grade", objProduto.sGrade, STRING_GRADE_CODIGO, "Grade"
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 64482

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165540)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iControleEstoque As Integer

On Error GoTo Erro_Tela_Preenche

    objProduto.sCodigo = colCampoValor.Item("Codigo").vValor

    If Len(Trim(objProduto.sCodigo)) <> 0 Then

        'Carrega objProduto com os dados passados em colCampoValor
        objProduto.iTipo = colCampoValor.Item("Tipo").vValor
        objProduto.sDescricao = colCampoValor.Item("Descricao").vValor
        objProduto.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objProduto.sModelo = colCampoValor.Item("Modelo").vValor
        objProduto.iGerencial = colCampoValor.Item("Gerencial").vValor
        objProduto.iNivel = colCampoValor.Item("Nivel").vValor
        objProduto.sSubstituto1 = colCampoValor.Item("Substituto1").vValor
        objProduto.sSubstituto2 = colCampoValor.Item("Substituto2").vValor
        objProduto.iPrazoValidade = colCampoValor.Item("PrazoValidade").vValor
        objProduto.sCodigoBarras = colCampoValor.Item("CodigoBarras").vValor
        objProduto.iEtiquetasCodBarras = colCampoValor.Item("EtiquetasCodBarras").vValor
        objProduto.dPesoLiq = colCampoValor.Item("PesoLiq").vValor
        objProduto.dPesoBruto = colCampoValor.Item("PesoBruto").vValor
        objProduto.dComprimento = colCampoValor.Item("Comprimento").vValor
        objProduto.dEspessura = colCampoValor.Item("Espessura").vValor
        objProduto.dLargura = colCampoValor.Item("Largura").vValor
        objProduto.sCor = colCampoValor.Item("Cor").vValor
        objProduto.sObsFisica = colCampoValor.Item("ObsFisica").vValor
        objProduto.iClasseUM = colCampoValor.Item("ClasseUM").vValor
        objProduto.sSiglaUMCompra = colCampoValor.Item("SiglaUMCompra").vValor
        objProduto.sSiglaUMEstoque = colCampoValor.Item("SiglaUMEstoque").vValor
        objProduto.sSiglaUMVenda = colCampoValor.Item("SiglaUMVenda").vValor
        objProduto.iAtivo = colCampoValor.Item("Ativo").vValor
        objProduto.iFaturamento = colCampoValor.Item("Faturamento").vValor
        objProduto.iCompras = colCampoValor.Item("Compras").vValor
        objProduto.iPCP = colCampoValor.Item("PCP").vValor
        objProduto.iKitBasico = colCampoValor.Item("KitBasico").vValor
        objProduto.iKitInt = colCampoValor.Item("KitInt").vValor
        objProduto.dIPIAliquota = colCampoValor.Item("IPIAliquota").vValor
        objProduto.sIPICodigo = colCampoValor.Item("IPICodigo").vValor
        objProduto.sIPICodDIPI = colCampoValor.Item("IPICodDIPI").vValor
        objProduto.iApropriacaoCusto = colCampoValor.Item("Apropriacao").vValor
        objProduto.sContaContabil = colCampoValor.Item("ContaContabil").vValor
        objProduto.iNatureza = colCampoValor.Item("Natureza").vValor
        
        objProduto.sContaContabilProducao = colCampoValor.Item("ContaContabilProducao").vValor
        objProduto.dPercentMaisQuantCotAnt = colCampoValor.Item("PercentMaisQuantCotAnt").vValor
        objProduto.dPercentMenosQuantCotAnt = colCampoValor.Item("PercentMenosQuantCotAnt").vValor
        objProduto.iConsideraQuantCotAnt = colCampoValor.Item("ConsideraQuantCotAnt").vValor
        objProduto.iTemFaixaReceb = colCampoValor.Item("TemFaixaReceb").vValor
        objProduto.dPercentMaisReceb = colCampoValor.Item("PercentMaisReceb").vValor
        objProduto.dPercentMenosReceb = colCampoValor.Item("PercentMenosReceb").vValor
        objProduto.iRecebForaFaixa = colCampoValor.Item("RecebForaFaixa").vValor
        objProduto.iCreditoICMS = colCampoValor.Item("CreditoICMS").vValor
        objProduto.iCreditoIPI = colCampoValor.Item("CreditoIPI").vValor
        objProduto.dResiduo = colCampoValor.Item("Residuo").vValor
        objProduto.iNatureza = colCampoValor.Item("Natureza").vValor
        objProduto.dCustoReposicao = colCampoValor.Item("CustoReposicao").vValor
        objProduto.iOrigemMercadoria = colCampoValor.Item("OrigemMercadoria").vValor
        objProduto.iTempoProducao = colCampoValor.Item("TempoProducao").vValor
        objProduto.iRastro = colCampoValor.Item("Rastro").vValor
        objProduto.lHorasMaquina = colCampoValor.Item("HorasMaquina").vValor
        objProduto.dPesoEspecifico = colCampoValor.Item("PesoEspecifico").vValor
        objProduto.sReferencia = colCampoValor.Item("Referencia").vValor
        objProduto.sFigura = colCampoValor.Item("Figura").vValor
        '''objProduto.iEmbalagem = colCampoValor.Item("Embalagem").vValor 05/09/01 Marcelo
        objProduto.dINSSPercBase = colCampoValor.Item("INSSPercBase").vValor
''''        objProduto.iKitVendaComp = colCampoValor.Item("KitVendaComp").vValor
        objProduto.sGrade = colCampoValor.Item("Grade").vValor
   
   
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 133752
   
        'Se não achou o Produto --> erro
        If lErro = 28030 Then gError 133753
   
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError 64483
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 64483, 133752

        Case 133753
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165541)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Function SiglaUM_Exibe(sSiglaUM As String, sNome As String) As Long

Dim lErro As Long
Dim objUM As New ClassUnidadeDeMedida

On Error GoTo Erro_SiglaUM_Exibe

    'Se não selecionou nada --> Sai
    If SiglaUMCompra.ListIndex = -1 Then Exit Function
    
    objUM.iClasse = CInt(ClasseUM.Text)
    objUM.sSigla = sSiglaUM

    'Lê a Sigla de Unidade de Medida
    lErro = CF("UM_Le", objUM)
    If lErro <> SUCESSO And lErro <> 23775 Then Error 64484

    'Se não encontrar --> erro
    If lErro = 23775 Then Error 64485
    
    If sNome = "NomeUMEstoque" Then
        NomeUMEstoque.Caption = objUM.sNome
    ElseIf sNome = "NomeUMVenda" Then
        NomeUMVenda.Caption = objUM.sNome
    ElseIf sNome = "NomeUMTrib" Then
        NomeUMTrib.Caption = objUM.sNome
    Else
        NomeUMCompra.Caption = objUM.sNome
    End If
    
    SiglaUM_Exibe = SUCESSO
    
    Exit Function
        
Erro_SiglaUM_Exibe:

    SiglaUM_Exibe = Err
    
    Select Case Err
    
        Case 64484

        Case 64485
            lErro = Rotina_Erro(vbOKOnly, "ERRO_UNIDADE_MEDIDA_NAO_CADASTRADA", Err, objUM.iClasse, objUM.sSiglaUMBase)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165542)
    End Select
        
    Exit Function
            
End Function

'Private Sub TvwProduto_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwProduto_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Produto_Netos",objNode, TvwProduto.Nodes)
'        If lErro <> SUCESSO Then Error 64486
'
'    End If
'
'    Exit Sub
'
'Erro_TvwProduto_Expand:
'
'    Select Case Err
'
'        Case 64486
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165543)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PRODUTO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Produto"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Produto"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******


'**** fim do trecho a ser copiado *****

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is Tabela Then
            Call BotaoTabelaPreco_Click
        ElseIf Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeReduzido_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        ElseIf Me.ActiveControl Is ContaProducao Then
            Call LabelContaProducao_Click
        ElseIf Me.ActiveControl Is Substituto1 Then
            Call LblSubst1_Click
        ElseIf Me.ActiveControl Is Substituto2 Then
            Call LblSubst2_Click
        ElseIf Me.ActiveControl Is TipoProduto Then
            Call LblTipoProduto_Click
        ElseIf Me.ActiveControl Is ClasseUM Then
            Call LblClasseUM_Click
        ElseIf Me.ActiveControl Is ClasFiscIPI Then
            Call LabelClassificacaoFiscal_Click
        End If
                
    End If
    
    If KeyCode = KEYCODE_IDIOMAS Then
        Call gobjIdiomaTela.AbreTela(Me.ActiveControl)
    End If

End Sub

'Parte da Versao 2 da tela Produto
Private Function Traz_TabCompras_Tela(objProduto As ClassProduto) As Long
'Traz os dados do tab Compras do BD para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_TabCompras_Tela

    If objProduto.iConsideraQuantCotAnt = 1 Then
        ConsideraQuantCotacaoAnterior.Value = vbUnchecked
    Else
        ConsideraQuantCotacaoAnterior.Value = vbChecked
    End If
    
    PercentMaisQuantCotacaoAnterior.Text = objProduto.dPercentMaisQuantCotAnt * 100
    PercentMenosQuantCotacaoAnterior.Text = objProduto.dPercentMenosQuantCotAnt * 100
    
    If objProduto.iTemFaixaReceb = 1 Then
        NaoTemFaixaReceb.Value = vbChecked
    Else
        NaoTemFaixaReceb.Value = vbUnchecked
    End If
        
    PercentMaisReceb.Text = objProduto.dPercentMaisReceb * 100
    PercentMenosReceb.Text = objProduto.dPercentMenosReceb * 100
                    
    RecebForaFaixa(objProduto.iRecebForaFaixa).Value = True
    
    Traz_TabCompras_Tela = SUCESSO

    Exit Function

Erro_Traz_TabCompras_Tela:

    Traz_TabCompras_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165544)

    End Select

    Exit Function

End Function

Private Function Move_TabCompras_Memoria(objProduto As ClassProduto) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Move_TabCompras_Memoria
    
    If ConsideraQuantCotacaoAnterior.Value = vbChecked Then
        objProduto.iConsideraQuantCotAnt = 0
    Else
        objProduto.iConsideraQuantCotAnt = 1
    End If
    
    If Len(Trim(PercentMaisQuantCotacaoAnterior.Text)) > 0 Then objProduto.dPercentMaisQuantCotAnt = CDbl(PercentMaisQuantCotacaoAnterior.Text) / 100
    If Len(Trim(PercentMenosQuantCotacaoAnterior.Text)) > 0 Then objProduto.dPercentMenosQuantCotAnt = CDbl(PercentMenosQuantCotacaoAnterior.Text) / 100
    
    If NaoTemFaixaReceb.Value = vbChecked Then
        objProduto.iTemFaixaReceb = 1
    Else
        objProduto.iTemFaixaReceb = 0
    End If
        
    If Len(Trim(PercentMaisReceb.Text)) > 0 Then objProduto.dPercentMaisReceb = CDbl(PercentMaisReceb.Text) / 100
    If Len(Trim(PercentMenosReceb.Text)) > 0 Then objProduto.dPercentMenosReceb = CDbl(PercentMenosReceb.Text) / 100
                    
    For iIndice = 0 To RecebForaFaixa.Count - 1
        
        If RecebForaFaixa(iIndice).Value = True Then
            objProduto.iRecebForaFaixa = iIndice
        End If
    
    Next
    
    Move_TabCompras_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_TabCompras_Memoria:

    Move_TabCompras_Memoria = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165545)

    End Select

    Exit Function
        
End Function

Public Sub PercentMaisQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMaisQuantCotacaoAnterior_Validate

    'Verifica se esta preenchida
    If Len(Trim(PercentMaisQuantCotacaoAnterior.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMaisQuantCotacaoAnterior.Text)
    If lErro <> SUCESSO Then Error 54106
    
    'Testa se percentagem é 100
    If StrParaDbl(PercentMaisQuantCotacaoAnterior.Text) = 100# Then Error 54107

    'Coloca na tela
    PercentMaisQuantCotacaoAnterior.Text = Format(PercentMaisQuantCotacaoAnterior.Text, "Fixed")

    Exit Sub

Erro_PercentMaisQuantCotacaoAnterior_Validate:

    Cancel = True


    Select Case Err

        Case 54106 'Erro criticado na rotina de chamada
                
        Case 54107
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165546)

    End Select

    Exit Sub

End Sub

Public Sub PercentMaisReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMaisReceb_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMaisReceb_Validate

    'Verifica se está preenchida
    If Len(Trim(PercentMaisReceb.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMaisReceb.Text)
    If lErro <> SUCESSO Then Error 54108

    'Testa se percentagem é 100
    If StrParaDbl(PercentMaisReceb.Text) = 100# Then Error 54109

    'Coloca na tela
    PercentMaisReceb.Text = Format(PercentMaisReceb.Text, "Fixed")

    Exit Sub

Erro_PercentMaisReceb_Validate:

    Cancel = True


    Select Case Err

        Case 54108  'Erro criticado na rotina chamada
            
        Case 54109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 165547)

    End Select

    Exit Sub

End Sub

Public Sub PercentMenosQuantCotacaoAnterior_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMenosQuantCotacaoAnterior_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMenosQuantCotacaoAnterior_Validate

    'Verifica se está preenchida
    If Len(Trim(PercentMenosQuantCotacaoAnterior.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMenosQuantCotacaoAnterior.Text)
    If lErro <> SUCESSO Then Error 54102
      
    'Testa se percentagem é 100
    If StrParaDbl(PercentMenosQuantCotacaoAnterior.Text) = 100# Then Error 54103

    'Coloca na tela
    PercentMenosQuantCotacaoAnterior.Text = Format(PercentMenosQuantCotacaoAnterior.Text, "Fixed")

    Exit Sub

Erro_PercentMenosQuantCotacaoAnterior_Validate:

    Cancel = True


    Select Case Err

        Case 54102 'Erro criticado na rotina de chamada
            
        Case 54103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165548)

    End Select

    Exit Sub

End Sub

Public Sub PercentMenosReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentMenosReceb_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercentMenosReceb_Validate

    'Verifica se está preenchido
    If Len(Trim(PercentMenosReceb.Text)) = 0 Then Exit Sub

    'Critica se é percentagem
    lErro = Porcentagem_Critica(PercentMenosReceb.Text)
    If lErro <> SUCESSO Then Error 54110

    'Testa se percentagem é 100
    If StrParaDbl(PercentMenosReceb.Text) = 100# Then Error 54111

    'Coloca na tela
    PercentMenosReceb.Text = Format(PercentMenosReceb.Text, "Fixed")

    Exit Sub

Erro_PercentMenosReceb_Validate:

    Cancel = True
    
    Select Case Err

        Case 54110 'Erro criticado na rotina de chamada
            
        Case 54111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_IGUAL_100", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165549)

    End Select

    Exit Sub

End Sub

Private Sub Preenche_COMConfiguracoes(objTipoProduto As ClassTipoDeProduto)

    If objTipoProduto.iConsideraQuantCotAnt = PRODUTO_CONSIDERA_QUANT_COTACAO_ANTERIOR Then
        ConsideraQuantCotacaoAnterior.Value = vbChecked
    Else
        ConsideraQuantCotacaoAnterior.Value = vbUnchecked
    End If
    PercentMaisQuantCotacaoAnterior = Format(objTipoProduto.dPercentMaisQuantCotAnt * 100, "Fixed")
    PercentMenosQuantCotacaoAnterior = Format(objTipoProduto.dPercentMenosQuantCotAnt * 100, "Fixed")
    If objTipoProduto.iTemFaixaReceb = 0 Then
        NaoTemFaixaReceb.Value = vbUnchecked
    Else
        NaoTemFaixaReceb.Value = vbChecked
    End If
    PercentMaisReceb.Text = Format(objTipoProduto.dPercentMaisReceb * 100, "Fixed")
    PercentMenosReceb.Text = Format(objTipoProduto.dPercentMenosReceb * 100, "Fixed")
    RecebForaFaixa(objTipoProduto.iRecebForaFaixa).Value = True
    
End Sub































'Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label28, Source, X, Y)
'End Sub

'Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label28, Button, Shift, X, Y)
'End Sub





'Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(LabelProduto, Source, X, Y)
'End Sub
'
'Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(LabelProduto, Button, Shift, X, Y)
'End Sub































































Function TiposTribICMSLoja_Le(colTiposTribICMS As Collection) As Long
'Lê todos os tipos de tributação ICMS e guarda em colTiposTribICMS

Dim lErro As Long
Dim objTiposTribICMS As ClassTipoTribICMS
Dim lComando As Long
Dim iTipo As Integer
Dim sDescricao As String

On Error GoTo Erro_TiposTribICMSLoja_Le

    'abre o comando
    lComando = Comando_Abrir
    If lComando = 0 Then gError 81195
    
    sDescricao = String(STRING_TIPOSTRIBICMS_DESCRICAO, 0)
    
    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao FROM TiposTribICMS WHERE SituacaoTribECF <> ? ", iTipo, sDescricao, "")
    If lErro <> AD_SQL_SUCESSO Then gError 81196

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81197

    Do While lErro = AD_SQL_SUCESSO
    
        'Guarda na coleção o tipo e a descrição
        Set objTiposTribICMS = New ClassTipoTribICMS

        objTiposTribICMS.iTipo = iTipo
        objTiposTribICMS.sDescricao = sDescricao
    
            colTiposTribICMS.Add objTiposTribICMS

        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81198

    Loop

    Call Comando_Fechar(lComando)

    TiposTribICMSLoja_Le = SUCESSO

    Exit Function

Erro_TiposTribICMSLoja_Le:

    TiposTribICMSLoja_Le = gErr
    
    Select Case gErr
    
        Case 81195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 81196, 81197, 81198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165550)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

'05/09/01 Marcelo

'Private Sub objEventoEmbalagem_evSelecao(obj1 As Object)
'
'Dim objEmbalagem  As New ClassEmbalagem
'
'    Set objEmbalagem = obj1
'
'    'Preenche codigo e descrição da Embalagem na tela
'    Embalagem.Text = CStr(objEmbalagem.iCodigo)
'    DescricaoEmbalagem.Caption = objEmbalagem.sDescricao
'
'    Me.Show
'
'End Sub

'***ALTERACAO POR TULIO EM 27/05***
'CRIADA A FUNCAO
Private Sub Carrega_CodigoBarras_Produto(ByVal colCodigoBarras As Collection)
'Carrega a combo de codigo de barras com uma colecao
'colCodigoBarras eh parametro de INPUT que traz os codigos de barra a serem
'carregados na combo

Dim vCodBarras As Variant

On Error GoTo Erro_Carrega_CodigoBarras_Produto

    'limpa a combo
    CodigoBarras.Clear
    
    For Each vCodBarras In colCodigoBarras

        'Insere na combo CodBarras o codigo de barras apontado pela "string"
        'vCodBarras
        CodigoBarras.AddItem CStr(vCodBarras)

    Next

    If CodigoBarras.ListCount = 1 Then CodigoBarras.ListIndex = 0
    
    Exit Sub

Erro_Carrega_CodigoBarras_Produto:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165551)
    
    End Select
    
    Exit Sub

End Sub

'CRIADA A FUNCAO
Private Sub ObterCodigosBarra(ByVal colCodigoBarras As Collection)
'copia para a colecao de codigo de barras do objProduto (colCodBarras referencia ela)
'os codigos de barra discriminados na combo...
'colCodigoBarras eh parametro de OUTPUT que retornara a colecao com os codigos de
'barras preenchidos...

Dim iIndex As Integer
Dim sCodigoBarra As String

On Error GoTo Erro_ObterCodigosBarra

    'varre a combo de codigo de barras
    For iIndex = 0 To CodigoBarras.ListCount - 1
        
        'copia o codigo de barra de indice "iIndex" para sCodigoBarra
        sCodigoBarra = CodigoBarras.List(iIndex)
    
        'adiciona o dito cujo na colecao
        colCodigoBarras.Add sCodigoBarra
        
    Next

    Exit Sub

Erro_ObterCodigosBarra:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165552)
    
    End Select
    
    Exit Sub

End Sub

'***FIM ALTERACAO POR TULIO EM 27/05***

Private Sub Command1_Click()
'Função de Teste

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objLog As New ClassLog

On Error GoTo Erro_Command1_Click
 
    lErro = Log_Le(objLog)
    If lErro <> SUCESSO And lErro <> 101741 Then gError 101742
    
    lErro = Produto_Desmembra_Log(objProduto, objLog)
    '??If lErro <> SUCESSO And lErro = 104195 Then gError 104196

    Exit Sub
    
Erro_Command1_Click:
    
    Select Case gErr
                                                                    
        Case 101742
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165553)
         
        End Select
         
    Exit Sub
    
End Sub

Function Log_Le(objLog As ClassLog) As Long
'Le o log, feito so para teste

Dim lErro As Long
Dim tLog As typeLog
Dim lComando As Long

On Error GoTo Erro_Log_Le

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104197
    
    'Inicializa o Buffer da Variáveis String
    tLog.sLog1 = String(STRING_CONCATENACAO, 0)
    tLog.sLog2 = String(STRING_CONCATENACAO, 0)
    tLog.sLog3 = String(STRING_CONCATENACAO, 0)
    tLog.sLog4 = String(STRING_CONCATENACAO, 0)

    'Seleciona código e nome dos meios de pagamentos da tabela AdmMeioPagto
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc, Operacao, Log1, Log2, Log3, Log4 , Data , Hora FROM Log ", tLog.lNumIntDoc, tLog.iOperacao, tLog.sLog1, tLog.sLog2, tLog.sLog3, tLog.sLog4, tLog.dData, tLog.dData)
    If lErro <> SUCESSO Then gError 104198

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 104199


    If lErro = AD_SQL_SUCESSO Then

        'Carrega o objLog com as Infromações de bonco de dados
        objLog.lNumIntDoc = tLog.lNumIntDoc
        objLog.iOperacao = tLog.iOperacao
        objLog.sLog = tLog.sLog1 & tLog.sLog2 & tLog.sLog3 & tLog.sLog4
        objLog.dtData = tLog.dData
        objLog.dHora = tLog.dHora

    End If

    If lErro = AD_SQL_SEM_DADOS Then gError 101741
    
    Log_Le = SUCESSO

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

Erro_Log_Le:

    Log_Le = gErr

   Select Case gErr

    Case gErr

        Case 104198, 104199
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_LOG", gErr)
    
        Case 104202
            Call Rotina_Erro(vbOKOnly, "ERRO_LOG_NAO_EXISTENTE", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165554)

        End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Produto_Desmembra_Log(objProduto As ClassProduto, objLog As ClassLog) As Long
'Função que pega as informações do banco de Dados e Carrega no Obj
'objProduto eh parametro de OUTPUT q retorna um obj com os dados do log
'objLog eh parametro de INPUT que traz os dados do log

Dim lErro As Long
Dim iPosCTRL As Integer
Dim iPosSHIFT As Integer
Dim iPosFimProp As Integer
Dim iPosIniProp As Integer
Dim iPosFim As Integer
Dim sProduto As String
Dim sCodBarras As String
Dim sCategoriaItem As String
Dim iIndice As Integer

On Error GoTo Erro_Produto_Desmembra_Log

    'iPosCTRL Guarda a posição do Primeiro Control (delimita fim das propriedades
    'de produto e inicio dos codigos de barra
    iPosCTRL = InStr(1, objLog.sLog, Chr(vbKeyControl))

    'iPosFIM Guarda o Final da String
    iPosFim = InStr(1, objLog.sLog, Chr(vbKeyEnd))

    'iPosShift Guarda a posicao do primeiro shift (delimita fim dos codigos de barra e
    'inicio das propriedades de varios itens de categoria
    iPosSHIFT = InStr(iPosCTRL + 1, objLog.sLog, Chr(vbKeyShift))

    'String que Guarda as Propriedades do Produto
    sProduto = Mid(objLog.sLog, 1, iPosCTRL - 1)

    'String que Guarda os Codigos de barra
    sCodBarras = Mid(objLog.sLog, iPosCTRL + 1, iPosSHIFT - 1)
   
    'String que Guarda os Itens de Categoria
    sCategoriaItem = Mid(objLog.sLog, iPosSHIFT + 1, iPosFim - 1)

    'instancia uma nova area de memoria para objproduto
    Set objProduto = New ClassProduto

    'seta a primeira posicao da string de produto
    iPosIniProp = 1
    
    'Procura o Primeiro Esc dentro da String sProduto e Armazena a Posição
    'ou seja, se posiciona para carregar a primeira propriedade do produto
    iPosFimProp = InStr(iPosIniProp, sProduto, Chr(vbKeyEscape))
        
    'variavel que contem o indice do "esc" corrente
    iIndice = 1

    'enquanto achar esc
    Do While iPosIniProp <> 0
        
        'Verifica qual eh o indice corrente pois, atraves dele, sabe
        'qual propriedade deve ser carregada no objProduto
        Select Case iIndice
            Case 1: objProduto.dComprimento = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 2: objProduto.dCustoReposicao = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 3: objProduto.dEspessura = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 4: objProduto.dINSSPercBase = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 5: objProduto.dIPIAliquota = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 6: objProduto.dLargura = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 7: objProduto.dPercentMaisQuantCotAnt = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 8: objProduto.dPercentMaisReceb = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 9: objProduto.dPercentMenosQuantCotAnt = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 10: objProduto.dPercentMenosReceb = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 11: objProduto.dPesoBruto = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 12: objProduto.dPesoEspecifico = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 13: objProduto.dPesoLiq = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 14: objProduto.dResiduo = StrParaDbl(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 15: objProduto.iApropriacaoCusto = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 16: objProduto.iAtivo = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 17: objProduto.iClasseUM = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 18: objProduto.iCompras = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 19: objProduto.iConsideraQuantCotAnt = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 20: objProduto.iControleEstoque = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 21: objProduto.iCreditoICMS = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 22: objProduto.iCreditoIPI = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 23: objProduto.iEmbalagem = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 24: objProduto.iEtiquetasCodBarras = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 25: objProduto.iFaturamento = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 26: objProduto.iFreteAgregaCusto = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 27: objProduto.iGerencial = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 28: objProduto.iICMSAgregaCusto = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 29: objProduto.iIPIAgregaCusto = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 30: objProduto.iKitBasico = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 31: objProduto.iKitInt = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 32: objProduto.iKitVendaComp = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 33: objProduto.iNatureza = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 34: objProduto.iNivel = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 35: objProduto.iOrigemMercadoria = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 36: objProduto.iPCP = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 37: objProduto.iPrazoValidade = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 38: objProduto.iRastro = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 39: objProduto.iRecebForaFaixa = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 40: objProduto.iTabelaPreco = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 41: objProduto.iTemFaixaReceb = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 42: objProduto.iTempoProducao = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 43: objProduto.iTipo = StrParaInt(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 44: objProduto.lHorasMaquina = StrParaLong(Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp))
            Case 45: objProduto.sCodigo = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 46: objProduto.sCodigoBarras = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 47: objProduto.sContaContabil = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 48: objProduto.sContaContabilProducao = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 49: objProduto.sCor = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 50: objProduto.sDescricao = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 51: objProduto.sFigura = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 52: objProduto.sICMSAliquota = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 53: objProduto.sIPICodDIPI = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 54: objProduto.sIPICodigo = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 55: objProduto.sModelo = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 56: objProduto.sNomeReduzido = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 57: objProduto.sObsFisica = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 58: objProduto.sReferencia = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 59: objProduto.sSiglaUMCompra = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 60: objProduto.sSiglaUMEstoque = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 61: objProduto.sSiglaUMVenda = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 62: objProduto.sSituacaoTribECF = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 63: objProduto.sSubstituto1 = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 64: objProduto.sSubstituto2 = Mid(sProduto, iPosIniProp, iPosFimProp - iPosIniProp)
            Case 65: Exit Do
        End Select
        
        'incrementa o indice
        iIndice = iIndice + 1
        
        'Atualiza as Posições
        iPosIniProp = iPosFimProp + 1
        iPosFimProp = InStr(iPosFimProp + 1, sProduto, Chr(vbKeyEscape))

    Loop
    
    'coloca a informacao contida em sCodBarras na colecao de codigo de barras do objproduto
    Call Carrega_ColCodBarras_Log(objProduto, sCodBarras)
    
    'coloca a informacao contida em sCategoriaItem na colecao de codigo de barras do objproduto
    Call Carrega_ColCategoriaItem_Log(objProduto, sCategoriaItem)
    
    Produto_Desmembra_Log = SUCESSO

    Exit Function

Erro_Produto_Desmembra_Log:

    Produto_Desmembra_Log = -1

    Select Case gErr

    Case gErr

        Case -1
            Call Rotina_Erro(vbOKOnly, "ERRO_NO_CARREGAMENTO_DO_PRODUTO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165555)

        End Select


    Exit Function

End Function

Public Sub Carrega_ColCodBarras_Log(objProduto As ClassProduto, sCodBarras As String)
'objproduto eh parametro de OUTPUT q ira retornar a colecao de codigo de barras carregada
'sCodBarras eh parametro de INPUT q sera desembrado para carregar a colecao de codigo de barras

Dim iPosIni As Integer
Dim iPosFim As Integer

    'instancia a colecao de codigos de barra
    Set objProduto.colCodBarras = New Collection
    
    'inicaliza a posicao inicial com 1
    iPosIni = 1
    
    'Procura o Primeiro Esc dentro da String sCodBarras e Armazena a Posição
    'ou seja, se posiciona para carregar o primeiro codigo de barras
    iPosFim = InStr(iPosIni, sCodBarras, Chr(vbKeyEscape))
    
    'enquanto achar esc
    Do While iPosFim <> 0
        
        'carrega o codigo de barra no obj
        objProduto.colCodBarras.Add Mid(sCodBarras, iPosIni, iPosFim - iPosIni)
        
        'Atualiza as Posições
        iPosIni = iPosFim + 1
        iPosFim = InStr(iPosFim + 1, sCodBarras, Chr(vbKeyEscape))

    Loop

End Sub

Public Sub Carrega_ColCategoriaItem_Log(objProduto As ClassProduto, sCategoriaItem As String)
'objproduto eh parametro de OUTPUT q ira retornar a colecao de itens de categoria carregada
'sCategoriaItem eh parametro de INPUT q sera desembrado para carregar a colecao de itens de categoria

Dim iPosIni As Integer
Dim iPosFim As Integer
Dim iPosIniShift As Integer
Dim iPosFimShift As Integer
Dim iIndice As Integer
Dim objCategoriaItem As ClassCategoriaProduto

    'instancia a colecao de itens de categoria
    Set objProduto.colCategoriaItem = New Collection
    
    'inicializa a posicao inicial da categoria como um todo
    iPosIniShift = 1
    
    'inicializa a posicao final da categoria como um todo
    iPosFimShift = InStr(iPosIniShift, sCategoriaItem, Chr(vbKeyShift))
    
    'inicaliza a posicao inicial com 1
    iPosIni = 1
    
    'Procura o Primeiro Esc dentro da String sCodBarras e Armazena a Posição
    'ou seja, se posiciona para carregar o primeiro codigo de barras
    iPosFim = InStr(iPosIni, sCategoriaItem, Chr(vbKeyEscape))
    
    'variavel que contem o indice do "esc" corrente
    iIndice = 1

    'enquanto achar shift
    Do While iPosFimShift <> 0
    
        'enquanto achar esc
        Do While iPosIni <> 0
            
            'instancia um novo objeto da classe categoriaproduto
            Set objCategoriaItem = New ClassCategoriaProduto
            
            'Verifica qual eh o indice corrente pois, atraves dele, sabe
            'qual propriedade deve ser carregada no objCategoriaItem
            Select Case iIndice
                Case 1: objCategoriaItem.sCategoria = Mid(sCategoriaItem, iPosIni, iPosFim - iPosIni)
                Case 2: objCategoriaItem.sDescricao = Mid(sCategoriaItem, iPosIni, iPosFim - iPosIni)
                Case 3: objCategoriaItem.sSigla = Mid(sCategoriaItem, iPosIni, iPosFim - iPosIni)
                Case 4: Exit Do
            End Select
            
            'atualiza o indice
            iIndice = iIndice + 1
            
            'Atualiza as Posições
            iPosIni = iPosFim + 1
            iPosFim = InStr(iPosFim + 1, sCategoriaItem, Chr(vbKeyEscape))
    
        Loop

        'inicializa o indice novamente
        iIndice = 1
        
        'atualiza as posicoes
        iPosIniShift = iPosFimShift + 1
        iPosFimShift = InStr(iPosIniShift, sCategoriaItem, Chr(vbKeyShift))
        iPosIni = iPosIniShift
        iPosFim = InStr(iPosIni, sCategoriaItem, Chr(vbKeyEscape))

    Loop
    
End Sub


Function Carrega_ComboGrade() As Long

Dim lErro As Long
Dim objGrade As ClassGrade
Dim colGrade As New Collection

On Error GoTo Erro_Carrega_ComboGrade

    'Lê todas as Grades de Produto
    lErro = CF("Grades_Le_Todas", colGrade)
    If lErro <> SUCESSO Then gError 86259

    'Adiciona as Grades lidas na List
    For Each objGrade In colGrade
        Grades.AddItem objGrade.sCodigo
    Next

    Carrega_ComboGrade = SUCESSO

    Exit Function

Erro_Carrega_ComboGrade:

    Carrega_ComboGrade = gErr

    Select Case gErr

        Case 86259

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165556)

    End Select

    Exit Function

End Function

Public Sub ClasFiscIPI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim vbMsgRes As VbMsgBoxResult
Dim sCEST As String

On Error GoTo Erro_ClasFiscIPI_Validate

    'Verifica se o cmpo classificação fiscal está preenchido
    If Len(Trim(ClasFiscIPI.Text)) = 0 Then Exit Sub
    
    objClassificacaoFiscal.sCodigo = ClasFiscIPI.ClipText

    'Verifica se existe a Classificação Fiscal informada
    lErro = CF("ClassificacaoFiscal_Le", objClassificacaoFiscal)
    If lErro <> SUCESSO And lErro <> 123494 Then gError 125018
    
    'Se não existe, então pergunta se deseja criar
    If lErro = 123494 Then gError 125019
    
    lErro = CF("NCM_Le_CEST", objClassificacaoFiscal.sCodigo, sCEST)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 125018
    
    If lErro = SUCESSO Then
        CEST.PromptInclude = False
        CEST.Text = sCEST
        CEST.PromptInclude = True
        
        Call CEST_Validate(bSGECancelDummy)
    End If
    
    Exit Sub
    
Erro_ClasFiscIPI_Validate:

    Cancel = True

    Select Case gErr
    
        Case 125018
        
        Case 125019
        
            'Pergunta se deseja criar a Classificação Fiscal
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLASSIFICACAOFISCAL")
            
            If vbMsgRes = vbYes Then
            
                'Chama a tela para cadastrar uma nova Classificação Fiscal
                Call Chama_Tela("ClassificacaoFiscal", objClassificacaoFiscal)
                
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165557)

    End Select

    Exit Sub

End Sub

Public Sub LabelClassificacaoFiscal_Click()

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_LabelClassificacaoFiscal_Click

    'Preenche na memória o Código passado
    If Len(Trim(ClasFiscIPI.ClipText)) > 0 Then objClassificacaoFiscal.sCodigo = ClasFiscIPI.ClipText

    Call Chama_Tela("ClassificacaoFiscalLista", colSelecao, objClassificacaoFiscal, objEventoClasFiscIPI)

    Exit Sub
    
Erro_LabelClassificacaoFiscal_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165558)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClasFiscIPI_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objClassificacaoFiscal As New ClassClassificacaoFiscal
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoClasFiscIPI_evSelecao
    
    Set objClassificacaoFiscal = obj1

    'lErro = CF("ClassificacaoFiscal_Le", objClassificacaoFiscal)
    'If lErro <> SUCESSO And lErro <> 123494 Then gError 125020

    'If lErro = 123494 Then gError 125021

    'Preenche o Cliente com o Cliente selecionado
    ClasFiscIPI.PromptInclude = False
    ClasFiscIPI.Text = objClassificacaoFiscal.sCodigo
    ClasFiscIPI.PromptInclude = True
    Call ClasFiscIPI_Validate(bSGECancelDummy)

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoClasFiscIPI_evSelecao:

    Select Case gErr
    
        Case 125020
        
        Case 125021
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CLASSIFICACAOFISCAL_NAO_EXISTENTE", gErr, objClassificacaoFiscal.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165559)

    End Select

    Exit Sub

End Sub

Public Sub BotaoTeste_Click()

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoTeste_Click

    'Verifica se existe alguma mudança e se deseja salvá-la
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código do Produto está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
    
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 93581
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objProduto = New ClassProduto
            objProduto.sCodigo = sProduto
        
        End If
        
    End If

    'Chama a Tela se Custos
    Call Chama_Tela("ProdutoTeste", objProduto)

    Exit Sub

Erro_BotaoTeste_Click:

    Select Case gErr

        Case 93581

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165560)

    End Select

    Exit Sub

End Sub

'###########################################################
'Inserido por Wagner 08/03/2006
Private Function Traz_Serie_ParteNumerica_Tela() As Long

Dim lErro As Long
Dim iParteNum As Integer
Dim sSerieProx As String
Dim sParteNumerica As String

On Error GoTo Erro_Traz_Serie_ParteNumerica_Tela

    iParteNum = StrParaInt(SerieNum.Text)
    sSerieProx = SerieProx.Text
    
    If iParteNum <> 0 And Len(Trim(sSerieProx)) > 0 Then
    
        If Len(sSerieProx) < iParteNum Then gError 141791
        
        sParteNumerica = right(sSerieProx, iParteNum)
        If Not (IsNumeric(sParteNumerica)) Then gError 141792
    
        SeriePartNum.Caption = right(sSerieProx, iParteNum)
    
    End If

    Traz_Serie_ParteNumerica_Tela = SUCESSO

    Exit Function

Erro_Traz_Serie_ParteNumerica_Tela:

    Traz_Serie_ParteNumerica_Tela = gErr

    Select Case gErr

        Case 141791
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_MENOR_PARTENUMERICA", gErr, Len(sSerieProx), iParteNum)
        
        Case 141792
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIEPROX_PARTENUMERICA_NAO_NUMERICA", gErr, sParteNumerica)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141790)

    End Select

    Exit Function

End Function

Public Sub SerieProx_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieNum_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SerieProx_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieProx_Validate

    lErro = Traz_Serie_ParteNumerica_Tela
    If lErro <> SUCESSO Then gError 141794
    
    Exit Sub

Erro_SerieProx_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 141794
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141795)
    
    End Select
    
    Exit Sub

End Sub

Public Sub SerieNum_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieNum_Validate

    If StrParaInt(SerieNum.Text) <> 0 Then
    
        lErro = Valor_Positivo_Critica(SerieNum.Text)
        If lErro <> SUCESSO Then gError 141799
    
        lErro = Traz_Serie_ParteNumerica_Tela
        If lErro <> SUCESSO Then gError 141796
    
    End If
    
    Exit Sub

Erro_SerieNum_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 141796, 141799
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141797)
    
    End Select
    
    Exit Sub

End Sub
'###########################################################

Public Function BotaoCorTonTP1() As Object

    Set BotaoCorTonTP1 = BotaoCorTonTP
    
End Function

Public Function Referencia1() As Object

    Set Referencia1 = Referencia
    
End Function

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String
Dim sProdutoFormatado As String
Dim iProdPreenchido As Integer
Dim sProduto As String

On Error GoTo Erro_BotaoProxNum_Click

    lErro = CF2(Me, "Prod_BotaoProxNum_Click")
    If lErro <> SUCESSO Then gError 199310

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 199310
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199311)
    
    End Select

    Exit Sub

End Sub
Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

'customizacao
Function Prod_Form_Load(ByVal objCT As Object) As Long
    Prod_Form_Load = SUCESSO
End Function

Function Prod_Traz_Produto_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
    Prod_Traz_Produto_Tela = SUCESSO
End Function

Function Prod_TrazCarac_Tela(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
    Prod_TrazCarac_Tela = SUCESSO
End Function

Function Prod_Gravar_Registro(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
    Prod_Gravar_Registro = SUCESSO
End Function

Function Prod_MoveCaracMemoria(ByVal objCT As Object, ByVal objProduto As ClassProduto) As Long
    If Len(Trim(Comprimento.Text)) > 0 Then objProduto.dComprimento = CDbl(Comprimento.Text)
    If Len(Trim(Largura.Text)) > 0 Then objProduto.dLargura = CDbl(Largura.Text)
    If Len(Trim(Espessura.Text)) > 0 Then objProduto.dEspessura = CDbl(Espessura.Text)
    Prod_MoveCaracMemoria = SUCESSO
End Function

Function Prod_Limpa_Tela(ByVal objCT As Object) As Long
        
    Prod_Limpa_Tela = SUCESSO
    
End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PRODUTO
    'Verifica se o Código foi preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 64459

        If iPreenchido = PRODUTO_VAZIO Then Error 64460
        
        objAnotacao.sID = sProduto
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case 64459, 64460
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157853)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157854)
     
    End Select
     
    Exit Sub

End Sub

Function Prod_Form_Unload(ByVal objCT As Object) As Long

    Prod_Form_Unload = SUCESSO
    
End Function

Function Prod_BotaoProxNum_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim sCodigo As String
Dim sProdutoFormatado As String
Dim iProdPreenchido As Integer
Dim sProduto As String

On Error GoTo Erro_Prod_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("Produto_Automatico", sCodigo)
    If lErro <> SUCESSO Then gError 187403
    
    lErro = CF("Produto_Formata", sCodigo, sProdutoFormatado, iProdPreenchido)
     If lErro <> SUCESSO Then gError 187404
    
    lErro = Mascara_MascararProduto(sProdutoFormatado, sCodigo)
    If lErro <> SUCESSO Then gError 187405

    Codigo.PromptInclude = False
    Codigo.Text = sProdutoFormatado
    Codigo.PromptInclude = True

    Prod_BotaoProxNum_Click = SUCESSO

    Exit Function

Erro_Prod_BotaoProxNum_Click:

    Prod_BotaoProxNum_Click = gErr

    Select Case gErr

        Case 187403, 187404, 187405
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 187404)
    
    End Select

    Exit Function

End Function

Public Function Prod_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    Prod_Saida_Celula = SUCESSO
End Function

Public Function Prod_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    Prod_Rotina_Grid_Enable = SUCESSO
End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)
    Call CF2(Me, "Prod_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada)
End Sub

Public Sub BotaoProxCodBarras_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim sCodBarras As String

On Error GoTo Erro_BotaoProxCodBarras_Click

    lErro = CF2(Me, "Prod_BotaoProxCodBarras_Click")
    If lErro <> SUCESSO Then gError 199872

    Exit Sub

Erro_BotaoProxCodBarras_Click:

    Select Case gErr

        Case 199872
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199873)
    
    End Select

    Exit Sub
    
End Sub

Public Function Prod_BotaoProxCodBarras_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim lCodigo As Long
Dim sCodBarras As String

On Error GoTo Erro_Prod_BotaoProxCodBarras_Click

    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt_Trans", "ESTConfig", "NUM_PROX_CODBARRAS", lCodigo)
    If lErro <> SUCESSO Then gError 199436

    sCodBarras = "2" & Format(lCodigo, "00000000000")

    lErro = CF("DV_EAN13", sCodBarras)
    If lErro <> SUCESSO Then gError 199869

    objCT.CodigoBarras.AddItem sCodBarras
    objCT.CodigoBarras.ListIndex = objCT.CodigoBarras.NewIndex
        
    Exit Function

Erro_Prod_BotaoProxCodBarras_Click:

    Select Case gErr

        Case 199436, 199869
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199437)
    
    End Select

    Exit Function
    
End Function

Public Sub PrecoMaxConsumidor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExTIPI_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SiglaUMTrib_Click()

Dim lErro As Long
Dim sSiglaUMTrib As String

On Error GoTo Erro_SiglaUMTrib_Click

    iAlterado = REGISTRO_ALTERADO

    'Se não selecionou nada --> Sai
    If SiglaUMTrib.ListIndex = -1 Then Exit Sub
    
    sSiglaUMTrib = SiglaUMTrib.Text
        
    'Verifica se Existe e Exibe De acordo com o Parametro
    lErro = SiglaUM_Exibe(sSiglaUMTrib, "NomeUMTrib")
    If lErro <> SUCESSO Then gError 64371
        
    Exit Sub

Erro_SiglaUMTrib_Click:

    Select Case gErr

        Case 64371
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165505)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoEspecifico_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Genero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSQN_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PrecoMaxConsumidor_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_PrecoMaxConsumidor_Validate

    If Len(Trim(PrecoMaxConsumidor.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(PrecoMaxConsumidor.Text)
        If lErro <> SUCESSO Then gError 64322
        
    End If
    
    Exit Sub
    
Erro_PrecoMaxConsumidor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 64322
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165471)
    
    End Select
    
    Exit Sub

End Sub

Public Sub ISSQN_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objISSQN As New ClassISSQN
    
On Error GoTo Erro_ISSQN_Validate

    If Len(Trim(ISSQN.ClipText)) > 0 Then
    
        objISSQN.sCodigo = ISSQN.ClipText
        
        lErro = CF("ISSQN_Le", objISSQN)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202773
        If lErro <> SUCESSO Then gError 202774
        
        DescISSQN.Caption = objISSQN.sDescricao
        CodServNFe.Caption = CStr(objISSQN.lCodServNFe)
        
    Else
    
        DescISSQN.Caption = ""
        CodServNFe.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_ISSQN_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202773
        
        Case 202774
            Call Rotina_Erro(vbOKOnly, "ERRO_ISSQN_NAO_CADASTRADO", gErr, objISSQN.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202775)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Genero_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProdutoGenero As New ClassProdutoGenero
    
On Error GoTo Erro_Genero_Validate

    If Len(Trim(Genero.ClipText)) > 0 Then
    
        objProdutoGenero.sCodigo = Genero.Text
        
        lErro = CF("ProdutoGenero_Le", objProdutoGenero)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 202776
        If lErro <> SUCESSO Then gError 202777
        
        DescGenero.Caption = objProdutoGenero.sDescricao
        
    Else
        DescGenero.Caption = ""
    End If
    
    Exit Sub
    
Erro_Genero_Validate:

    Cancel = True

    Select Case gErr
    
        Case 202776
        
        Case 202777
            Call Rotina_Erro(vbOKOnly, "ERRO_GENERO_NAO_CADASTRADO", gErr, objProdutoGenero.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202778)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoISSQN_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objISSQN As ClassISSQN
    
On Error GoTo Erro_objEventoISSQN_evSelecao
    
    Set objISSQN = obj1

    ISSQN.PromptInclude = False
    ISSQN.Text = objISSQN.sCodigo
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoISSQN_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165559)

    End Select

    Exit Sub

End Sub

Private Sub objEventoGenero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProdutoGenero As ClassProdutoGenero
    
On Error GoTo Erro_objEventoGenero_evSelecao
    
    Set objProdutoGenero = obj1

    Genero.PromptInclude = False
    Genero.Text = objProdutoGenero.sCodigo
    Genero.PromptInclude = True
    Call Genero_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoGenero_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165559)

    End Select

    Exit Sub

End Sub

Public Sub LabelGenero_Click()

Dim objProdutoGenero As New ClassProdutoGenero
Dim colSelecao As Collection

    objProdutoGenero.sCodigo = Genero.ClipText

    Call Chama_Tela("ProdutoGeneroLista", colSelecao, objProdutoGenero, objEventoGenero)

End Sub

Public Sub LabelISSQN_Click()

Dim objISSQN As New ClassISSQN
Dim colSelecao As Collection

    objISSQN.sCodigo = ISSQN.ClipText

    Call Chama_Tela("ISSQNLista", colSelecao, objISSQN, objEventoISSQN, , "Codigo")

End Sub

Public Sub LabelCodServNFe_Click()

Dim objISSQN As New ClassISSQN
Dim colSelecao As Collection

    objISSQN.lCodServNFe = StrParaLong(CodServNFe.Caption)

    Call Chama_Tela("ISSQNLista", colSelecao, objISSQN, objEventoISSQN, , "CodServNFe")

End Sub

Public Sub BotaoSRV_Click()

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoSRV_Click

    'Verifica se existe alguma mudança e se deseja salvá-la
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Exit Sub

    'Verifica se o Código do Produto está preenchido
    If Len(Trim(Codigo.ClipText)) > 0 Then
    
        lErro = CF("Produto_Formata", Codigo.Text, sProduto, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 64365
        
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            Set objProduto = New ClassProduto
            objProduto.sCodigo = sProduto
        
        End If
        
    End If

    'Chama a Tela se Custos
    Call Chama_Tela("ProdutoServico", objProduto)

    Exit Sub

Erro_BotaoSRV_Click:

    Select Case gErr

        Case 64365

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165497)

    End Select

    Exit Sub

End Sub

Public Sub Fabricante_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fabricante_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Fabricante_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Fabricante_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_PROD_FABR, Fabricante, "AVISO_CRIAR_PROD_FABRICANTE")
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Exit Sub

Erro_Fabricante_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211534)

    End Select

End Sub

Private Function Carrega_Fabricante(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Fabricante

Dim lErro As Long

On Error GoTo Erro_Carrega_Fabricante

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_PROD_FABR, objComboBox)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    iFabricante_ListIndex_Padrao = objComboBox.ListIndex

    Carrega_Fabricante = SUCESSO

    Exit Function

Erro_Carrega_Fabricante:

    Carrega_Fabricante = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211535)

    End Select

    Exit Function

End Function

Private Sub DesabilitaGerencialUmNivel()

Dim lErro As Long
Dim objSegmento As New ClassSegmento
Dim colSegmentos As New Collection

On Error GoTo Erro_DesabilitaGerencialUmNivel

    'Le quantosa segmentos tem o produto
    objSegmento.sCodigo = SEGMENTO_PRODUTO

    lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmentos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If colSegmentos.Count = 1 Then
    
        NivelFinal.Value = MARCADO
        NivelGerencial.Enabled = False

    End If

    Exit Sub

Erro_DesabilitaGerencialUmNivel:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208411)

    End Select

    Exit Sub

End Sub

Public Sub FCI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FatorAjuste_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FatorAjuste_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FatorAjuste_Validate

    'Verifica se esta preenchida
    If Len(Trim(FatorAjuste.Text)) = 0 Then Exit Sub

    'Critica
    lErro = Porcentagem_Critica(FatorAjuste.Text)
    If lErro <> SUCESSO Then gError 89105

    Exit Sub

Erro_FatorAjuste_Validate:

    Cancel = True

    Select Case gErr

        Case 89105

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165500)

    End Select

    Exit Sub

End Sub

Sub FCI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer, sChar As String

On Error GoTo Erro_FCI_Validate
    
    If Len(Trim(FCI.Text)) > 0 Then
    
        If Len(Trim(FCI.Text)) <> 36 Then gError 201573
        
        For iIndice = 1 To Len(Trim(FCI.Text))
            sChar = Mid(FCI.Text, iIndice, 1)
            If sChar <> "0" And sChar <> "1" And sChar <> "2" And sChar <> "3" And _
            sChar <> "4" And sChar <> "5" And sChar <> "6" And sChar <> "7" And _
            sChar <> "8" And sChar <> "9" And sChar <> "A" And sChar <> "B" And _
            sChar <> "C" And sChar <> "D" And sChar <> "E" And sChar <> "F" And _
            sChar <> "-" Then gError 201573
        Next
        
    End If

    Exit Sub
    
Erro_FCI_Validate:

    Cancel = True
    
    Select Case gErr

        Case 201573
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMATO_FCI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201572)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDescSimilar_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoDescSimilar_Click

    'Verifica se o produto foi preenchido
    If Len(Descricao.Text) <> 0 Then

        objProduto.sDescricao = Descricao.Text

    End If
    
    colSelecao.Add "%" & objProduto.sDescricao & "%"

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto, "Descricao LIKE ?", "Descricao")

    Exit Sub

Erro_BotaoDescSimilar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 165468)

    End Select

    Exit Sub
    
End Sub

Public Sub PercComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComissao_Validate

    'Verifica se está preenchido
    If Len(Trim(PercComissao.Text)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Porcentagem_Critica(PercComissao.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    PercComissao.Text = Format(PercComissao.Text, "Fixed")

    Exit Sub

Erro_PercComissao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165501)

    End Select

    Exit Sub

End Sub

Public Sub MetaComissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MetaComissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_MetaComissao_Validate

    'Verifica se está preenchido
    If Len(Trim(MetaComissao.Text)) = 0 Then Exit Sub

    'Criticao valor
    lErro = Valor_Positivo_Critica(MetaComissao.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor formatado na Tela
    MetaComissao.Text = Format(MetaComissao.Text, MetaComissao.Format)

    Exit Sub

Erro_MetaComissao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 165501)

    End Select

    Exit Sub

End Sub

Public Sub LabelCEST_Click()

Dim lErro As Long
Dim objCEST As New ClassCEST
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCEST_Click

    'Preenche na memória o Código passado
    If Len(Trim(CEST.ClipText)) > 0 Then objCEST.sCodigo = CEST.ClipText
    
    If Len(Trim(ClasFiscIPI.ClipText)) > 0 Then
    
        colSelecao.Add ClasFiscIPI.ClipText

        Call Chama_Tela("CESTLista", colSelecao, objCEST, objEventoCEST, "? LIKE (NCM + '%') AND NCM <> ''")
        
    Else
    
        Call Chama_Tela("CESTLista", colSelecao, objCEST, objEventoCEST, "NCM = ''")
    
    End If

    Exit Sub
    
Erro_LabelCEST_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216098)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCEST_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCEST As New ClassCEST
Dim bCancel As Boolean
    
On Error GoTo Erro_objEventoCEST_evSelecao
    
    Set objCEST = obj1

    'Preenche o Cliente com o Cliente selecionado
    CEST.PromptInclude = False
    CEST.Text = objCEST.sCodigo
    CEST.PromptInclude = True
    
    Call CEST_Validate(bSGECancelDummy)

    Me.Show

    iAlterado = 0
    
    Exit Sub

Erro_objEventoCEST_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216099)

    End Select

    Exit Sub

End Sub

Public Sub CEST_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CEST_GotFocus()
    Call MaskEdBox_TrataGotFocus(CEST, iAlterado)
End Sub

Public Sub CEST_Validate(Cancel As Boolean)
Dim lErro As Long
Dim objCEST As New ClassCEST
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CEST_Validate

    'Verifica se o cmpo classificação fiscal está preenchido
    If Len(Trim(CEST.ClipText)) = 0 Then Exit Sub
    
    objCEST.sCodigo = CEST.ClipText
    objCEST.sNCM = ClasFiscIPI.ClipText

    'Verifica se existe a Classificação Fiscal informada
    lErro = CF("CEST_Le", objCEST)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    'Se não existe, então pergunta se deseja criar
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216096
    
    Exit Sub
    
Erro_CEST_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 216096
            Call Rotina_Erro(vbOKOnly, "ERRO_CEST_NCM_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216097)

    End Select

    Exit Sub
    
End Sub

Public Sub cProdANVISA_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub cProdANVISA_GotFocus()
    Call MaskEdBox_TrataGotFocus(cProdANVISA, iAlterado)
End Sub

Public Sub cProdANVISA_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_cProdANVISA_Validate

    'Verifica se o cmpo classificação fiscal está preenchido
    If Len(Trim(cProdANVISA.ClipText)) = 0 Then Exit Sub
    
    
    Exit Sub
    
Erro_cProdANVISA_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216097)

    End Select

    Exit Sub
    
End Sub
