VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Mnemônicos para formação da mensagem da NF
'NF
Private Const MN_EXISTE_PRODUTO As String = "Existe_Produto"
Private Const MN_ENTRADA As String = "Entrada"
Private Const MN_SAIDA As String = "Saida"
Private Const MN_TIPOTRIB As String = "TipoTrib"
Private Const MN_TIPONF As String = "TipoNF"
Private Const MN_FILEMPCGC As String = "FilEmpCGC"
Private Const MN_FILEMP As String = "FilEmp"
Private Const MN_EXISTE_CATPRODITEM As String = "Existe_CatProdItem"
Private Const MN_EXISTE_CATPROD As String = "Existe_CatProd"
Private Const MN_EXISTE_TIPOPROD As String = "Existe_TipoProd"
Private Const MN_CODBACENDEST As String = "CodBacenDest"
Private Const MN_UFDEST As String = "UFDest"
Private Const MN_ENDREFDEST As String = "EndRefDest"
Private Const MN_REGRA As String = "Regra"
Private Const MN_BASECREDICMS As String = "BaseCredICMS"
Private Const MN_ALIQUOTACREDICMS As String = "AliquotaCredICMS"
Private Const MN_VALORCREDICMS As String = "ValorCredICMS"
Private Const MN_MSGPV As String = "MsgPV"
Private Const MN_MSGFILCLI As String = "MsgFilCli"
Private Const MN_ENDENTREGA As String = "EndEntrega"
Private Const MN_ENDTRANSREDESP As String = "EndTranspRedesp"
Private Const MN_FILCLIFORNINSCSUF As String = "FilCliFornInscSuf"
Private Const MN_NUMPEDIDOTERC As String = "NumPedidoTerc"
Private Const MN_TIPOCLIFORN As String = "TipoCliForn"
Private Const MN_CFOP As String = "CFOP"
Private Const MN_CLIFORNNOMERED As String = "CliFornNomeRed"
Private Const MN_CLIFORN As String = "CliForn"
Private Const MN_FILCLIFORNINSCEST As String = "FilCliFornInscEst"
Private Const MN_FILCLIFORNINSCMUN As String = "FilCliFornInscMun"
Private Const MN_POSSUI_ICMSST As String = "Possui_ICMSST"
Private Const MN_POSSUI_ICMSRED As String = "Possui_ICMSRED"
Private Const MN_POSSUI_CATCLIFORITEM As String = "Possui_CatCliForItem"
Private Const MN_POSSUI_CATCLIFORN As String = "Possui_CatCliForn"
Private Const MN_INTERNA As String = "Interna"
Private Const MN_EXTERNA As String = "Externa"
Private Const MN_SERIE As String = "Serie"
Private Const MN_VLRDESCONTONF As String = "VlrDescontoNF"
Private Const MN_VLRTOTALNF As String = "VlrTotalNF"
Private Const MN_VLRFRETENF As String = "VlrFreteNF"
Private Const MN_VLRSEGURONF As String = "VlrSeguroNF"
Private Const MN_VLRDESPESASNF As String = "VlrDespesasNF"
Private Const MN_VLRPRODUTONF As String = "VlrProdutoNF"
Private Const MN_NUMPEDIDOVENDA As String = "NumPedidoVenda"

Private Const MN_DEVOLUCAO As String = "Devolucao"
Private Const MN_TIPONFORIG As String = "TipoNFOrig"
Private Const MN_NUMNFORIG As String = "NumNFOrig"
Private Const MN_NUMNFORIGDET As String = "NumNFOrigDet"
Private Const MN_NUMNFORIGDETBENEF As String = "NumNFOrigDetBenef"
Private Const MN_NUMNFORIGVINC As String = "NumNFOrigVinc"
Private Const MN_CFOPNFORIG As String = "CFOPNFOrig"
Private Const MN_EMISSAONFORIG As String = "EmissaoNFOrig"
Private Const MN_CLIFORNNFORIG As String = "CliFornNFOrig"
Private Const MN_TOTTRIB As String = "TotTrib"
Private Const MN_PERCTOTTRIB As String = "PercTotTrib"
Private Const MN_TOTTRIBFED As String = "TotTribFed"
Private Const MN_PERCTOTTRIBFED As String = "PercTotTribFed"
Private Const MN_TOTTRIBEST As String = "TotTribEst"
Private Const MN_PERCTOTTRIBEST As String = "PercTotTribEst"
Private Const MN_TOTTRIBMUNIC As String = "TotTribMunic"
Private Const MN_PERCTOTTRIBMUNIC As String = "PercTotTribMunic"

'Exclusivos do ItemNF
Private Const MN_PRODUTO As String = "Produto"
Private Const MN_TIPOPROD As String = "TipoProd"
Private Const MN_POSSUI_CATPROD As String = "Possui_CatProd"
Private Const MN_POSSUI_CATPRODITEM As String = "Possui_CatProdItem"
Private Const MN_CFOPITEM As String = "CFOPItem"
Private Const MN_NCM As String = "NCM"
Private Const MN_CATPRODITEM As String = "CatProdItem"
Private Const MN_CATPRODITEMDESC As String = "CatProdItemDesc"
Private Const MN_LOTE As String = "Lote"
Private Const MN_QTDLOTE As String = "QtdLote"
Private Const MN_LOTEOBS As String = "LoteOBS"
Private Const MN_POSSUI_RASTREAMENTO As String = "Possui_Rastreamento"
Private Const MN_NUMLOTES As String = "NumLotes"
Private Const MN_FABRICACAOLOTE As String = "FabricacaoLote"
Private Const MN_VALIDADELOTE As String = "ValidadeLote"
Private Const MN_TODOSLOTES As String = "TodosLotes"
Private Const MN_TODOSPEDCLI As String = "TodosPedCli"
Private Const MN_INFOADIC_TXTAD As String = "InfoAdicTexto"
Private Const MN_PROD_OBS As String = "Obs_Prod"
Private Const MN_TOTTRIBITEM As String = "TotTribItem"
Private Const MN_PERCTOTTRIBITEM As String = "PercTotTribItem"
Private Const MN_INFOADIC_PEDCOMPRAITEM As String = "PedCompraDoItem"
Private Const MN_INFOADIC_ITEMPEDCOMPRAITEM As String = "ItemPedCompraDoItem"
Private Const MN_FCI As String = "FCI"
Private Const MN_TOTTRIBFEDITEM As String = "TotTribFedItem"
Private Const MN_PERCTOTTRIBFEDITEM As String = "PercTotTribFedItem"
Private Const MN_TOTTRIBESTITEM As String = "TotTribEstItem"
Private Const MN_PERCTOTTRIBESTITEM As String = "PercTotTribEstItem"
Private Const MN_TOTTRIBMUNICITEM As String = "TotTribMunicItem"
Private Const MN_PERCTOTTRIBMUNICITEM As String = "PercTotTribMunicItem"

Private Const MN_VLRICMS As String = "VlrICMS"
Private Const MN_VLRICMSINTERESTFCP As String = "VlrICMSInterestFCP"
Private Const MN_VLRICMSINTERESTUFDEST As String = "VlrICMSInterestDest"
Private Const MN_VLRICMSINTERESTUFREMET As String = "VlrICMSInterestRemet"

Private Const MN_VLRICMSST As String = "VlrICMSST"
Private Const MN_BASEICMSST As String = "BaseICMSST"
Private Const MN_VLRICMSSTCOBRANT As String = "VlrICMSSTCobrAnt"
Private Const MN_BASEICMSSTCOBRANT As String = "BaseICMSSTCobrAnt"

Private Const MN_BASEICMSSTNFORIG As String = "BaseICMSSTNFOrig"
Private Const MN_VLRICMSSTNFORIG As String = "VlrICMSSTNFOrig"


Private gcolRetRegras As Collection
Private gcolRegras As Collection

Private mvariTodosOsPedidos As Integer
Private mvariFilialPedidoDe As Integer
Private mvariFilialPedidoAte As Integer
Private mvarlPedidosDe As Long
Private mvarlPedidosAte As Long
Private mvarlClientesDe As Long
Private mvarlClientesAte As Long
Private mvardtEmissaoDe As Date
Private mvardtEmissaoAte As Date
Private mvardtEntregaDe As Date
Private mvardtEntregaAte As Date
Private mvarsOrdenacao As String
'Private mvarlGeracaoNFId As Long
Private mvarcolNFiscalInfo As New Collection
Private mvarcolAlmoxFilial As Collection 'elementos objAlmoxarifado com codigo e nome reduzido da filialfaturamento
Private mvariTipoNFiscal As Integer
Private mvariAlmoxarifadoPadrao As Integer
Private mvariTransacaoGerNFiscal As Integer
Private mvarobjNFiscal As ClassNFiscal
Private mvarobjPedidoVenda As ClassPedidoDeVenda
Private mvarobjItemNF As ClassItemNF
Private mvarobjProduto As ClassProduto
Private mvarsSerie As String
Private mvardtDataSaida As Date
Private mvariImprime As Integer
Private mvarobjEvolucao As Object
Private mvarobjTelaAtualizacao As Object 'tela de acompanhamento da atualizacao de arquivos. Necessaria para a contabilizacao.
Private mvardtContabil As Date
Private mvariLoteContabil As Integer
Private mvarlViagem As Long
Private mvariVendedorInicial As Integer
Private mvariVendedorFinal As Integer
Private mvarsRegioes As String

Public Property Let sRegioes(ByVal vData As String)
    mvarsRegioes = vData
End Property

Public Property Get sRegioes() As String
    sRegioes = mvarsRegioes
End Property

Public Property Let iVendedorFinal(ByVal vData As Integer)
    mvariVendedorFinal = vData
End Property

Public Property Get iVendedorFinal() As Integer
    iVendedorFinal = mvariVendedorFinal
End Property

Public Property Let iVendedorInicial(ByVal vData As Integer)
    mvariVendedorInicial = vData
End Property

Public Property Get iVendedorInicial() As Integer
    iVendedorInicial = mvariVendedorInicial
End Property

Public Property Let lViagem(ByVal vData As Long)
    mvarlViagem = vData
End Property

Public Property Get lViagem() As Long
    lViagem = mvarlViagem
End Property

Public Property Let iLoteContabil(ByVal vData As Integer)
    mvariLoteContabil = vData
End Property

Public Property Get iLoteContabil() As Integer
    iLoteContabil = mvariLoteContabil
End Property

Public Property Let dtContabil(ByVal vData As Date)
    mvardtContabil = vData
End Property

Public Property Get dtContabil() As Date
    dtContabil = mvardtContabil
End Property

Public Property Let objTelaAtualizacao(ByVal vData As Object)
    Set mvarobjTelaAtualizacao = vData
End Property

Public Property Get objTelaAtualizacao() As Object
    Set objTelaAtualizacao = mvarobjTelaAtualizacao
End Property

Public Property Let objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Get objEvolucao() As Object
    Set objEvolucao = mvarobjEvolucao
End Property

Public Property Set objEvolucao(ByVal vData As Object)
    Set mvarobjEvolucao = vData
End Property

Public Property Let iImprime(ByVal vData As Integer)
    mvariImprime = vData
End Property

Public Property Get iImprime() As Integer
    iImprime = mvariImprime
End Property

Public Property Let dtDataSaida(ByVal vData As Date)
Attribute dtDataSaida.VB_Description = "Data da Saida da Mercadoria relativa a Nota Fiscal"
    mvardtDataSaida = vData
End Property

Public Property Get dtDataSaida() As Date
    dtDataSaida = mvardtDataSaida
End Property

Public Property Let sSerie(ByVal vData As String)
Attribute sSerie.VB_Description = "Serie a qual a Nota Fiscal Pertence."
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Set objNFiscal(ByVal vData As Object)
    Set mvarobjNFiscal = vData
End Property

Public Property Get objNFiscal() As ClassNFiscal
    Set objNFiscal = mvarobjNFiscal
End Property

Public Property Set objPedidoVenda(ByVal vData As Object)
    Set mvarobjPedidoVenda = vData
End Property

Public Property Get objPedidoVenda() As ClassPedidoDeVenda
    Set objPedidoVenda = mvarobjPedidoVenda
End Property

Public Property Set objItemNF(ByVal vData As Object)
    Set mvarobjItemNF = vData
End Property

Public Property Get objItemNF() As ClassItemNF
    Set objItemNF = mvarobjItemNF
End Property

Public Property Set objProduto(ByVal vData As Object)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property

Public Property Let iTransacaoGerNFiscal(ByVal vData As Integer)
    mvariTransacaoGerNFiscal = vData
End Property

Public Property Get iTransacaoGerNFiscal() As Integer
    iTransacaoGerNFiscal = mvariTransacaoGerNFiscal
End Property

Public Property Let iAlmoxarifadoPadrao(ByVal vData As Integer)
    mvariAlmoxarifadoPadrao = vData
End Property

Public Property Get iAlmoxarifadoPadrao() As Integer
    iAlmoxarifadoPadrao = mvariAlmoxarifadoPadrao
End Property

Public Property Let iTipoNFiscal(ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal() As Integer
    iTipoNFiscal = mvariTipoNFiscal
End Property

Public Property Set colAlmoxFilial(ByVal vData As Object)
    Set mvarcolAlmoxFilial = vData
End Property

Public Property Get colAlmoxFilial() As Collection
    Set colAlmoxFilial = mvarcolAlmoxFilial
End Property

Public Property Set colNFiscalInfo(ByVal vData As Object)
    Set mvarcolNFiscalInfo = vData
End Property

Public Property Get colNFiscalInfo() As Collection
    Set colNFiscalInfo = mvarcolNFiscalInfo
End Property

'Public Property Let lGeracaoNFId(ByVal vData As Long)
'    mvarlGeracaoNFId = vData
'End Property
'
'Public Property Get lGeracaoNFId() As Long
'    lGeracaoNFId = mvarlGeracaoNFId
'End Property

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Let dtEntregaAte(ByVal vData As Date)
    mvardtEntregaAte = vData
End Property

Public Property Get dtEntregaAte() As Date
    dtEntregaAte = mvardtEntregaAte
End Property

Public Property Let dtEntregaDe(ByVal vData As Date)
    mvardtEntregaDe = vData
End Property

Public Property Get dtEntregaDe() As Date
    dtEntregaDe = mvardtEntregaDe
End Property

Public Property Let dtEmissaoAte(ByVal vData As Date)
    mvardtEmissaoAte = vData
End Property

Public Property Get dtEmissaoAte() As Date
    dtEmissaoAte = mvardtEmissaoAte
End Property

Public Property Let dtEmissaoDe(ByVal vData As Date)
    mvardtEmissaoDe = vData
End Property

Public Property Get dtEmissaoDe() As Date
    dtEmissaoDe = mvardtEmissaoDe
End Property

Public Property Let lClientesAte(ByVal vData As Long)
    mvarlClientesAte = vData
End Property

Public Property Get lClientesAte() As Long
    lClientesAte = mvarlClientesAte
End Property

Public Property Let lClientesDe(ByVal vData As Long)
    mvarlClientesDe = vData
End Property

Public Property Get lClientesDe() As Long
    lClientesDe = mvarlClientesDe
End Property

Public Property Let lPedidosAte(ByVal vData As Long)
    mvarlPedidosAte = vData
End Property

Public Property Get lPedidosAte() As Long
    lPedidosAte = mvarlPedidosAte
End Property

Public Property Let lPedidosDe(ByVal vData As Long)
    mvarlPedidosDe = vData
End Property

Public Property Get lPedidosDe() As Long
    lPedidosDe = mvarlPedidosDe
End Property

Public Property Let iFilialPedidoAte(ByVal vData As Integer)
    mvariFilialPedidoAte = vData
End Property

Public Property Get iFilialPedidoAte() As Integer
    iFilialPedidoAte = mvariFilialPedidoAte
End Property

Public Property Let iFilialPedidoDe(ByVal vData As Integer)
    mvariFilialPedidoDe = vData
End Property

Public Property Get iFilialPedidoDe() As Integer
    iFilialPedidoDe = mvariFilialPedidoDe
End Property

Public Property Let iTodosOsPedidos(ByVal vData As Integer)
    mvariTodosOsPedidos = vData
End Property

Public Property Get iTodosOsPedidos() As Integer
    iTodosOsPedidos = mvariTodosOsPedidos
End Property

Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long

    GeraContabilizacao = mvarobjTelaAtualizacao.GeraContabilizacao(objContabAutomatica, vParams)
    
End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, bAchou As Boolean, objFilialEmpresa As AdmFiliais
Dim iLinha As Integer, sCat As String, sCatItem As String
Dim objProdutoAux As ClassProduto, objProdCat As ClassProdutoCategoria, objItemNFAux As ClassItemNF
Dim objCli As ClassCliente, objFilCli As ClassFilialCliente, objTransp As ClassTransportadora
Dim objForn As ClassFornecedor, objFilForn As ClassFilialFornecedor
Dim objEnd As ClassEndereco, objPais As ClassPais, sAux As String, dValorBase As Double
Dim dAcumSemSubst As Double, dAcumComSubst As Double, dValorNaoFat As Double, dAliquota As Double, dValor As Double
Dim objDASAliquotas As ClassDASAliquotas, objPV As ClassPedidoDeVenda, objMensagem As ClassMensagem
Dim colCatItem As Collection, objFilCliCat As ClassFilialCliCategoria, objFilFornCat As ClassFilialFornCategoria
Dim objRastro As ClassRastroItemNF, iItem As Integer, objLote As ClassRastreamentoLote
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objDataEntrega As ClassDataEntrega, iFilialEmpresaAux As Integer
Dim objRegra As ClassRegrasMsg, objRegraAux As ClassRegrasMsg, iRegra As Integer
Dim objNFiscalOriginal As ClassNFiscal, X As ClassFAT, colAux As Collection, vValor As Variant
Dim objCatProdItem As ClassCategoriaProdutoItem, dValorBaseAux As Double, iAliqDif As Integer
Dim sCGCTranspRedesp As String, objItemNFOrig As ClassItemNF

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
 
        Case MN_EXISTE_PRODUTO
            bAchou = False
            For Each objItemNFAux In objNFiscal.colItens
                If UCase(objItemNFAux.sProduto) = UCase(objMnemonicoValor.vParam(1)) Then
                    bAchou = True
                    Exit For
                End If
            Next
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_ENTRADA
        
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM

            If objTipoDocInfo.iTipo = 1 Or objTipoDocInfo.iTipo = 3 Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_SAIDA
            
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            If objTipoDocInfo.iTipo = 1 Or objTipoDocInfo.iTipo = 3 Then
                objMnemonicoValor.colValor.Add False
            Else
                objMnemonicoValor.colValor.Add True
            End If
            
        Case MN_TIPOTRIB
            objMnemonicoValor.colValor.Add objNFiscal.objTributacao.iTipoTributacao

        Case MN_TIPONF
            objMnemonicoValor.colValor.Add objNFiscal.iTipoNFiscal
            
        Case MN_NUMPEDIDOVENDA
            objMnemonicoValor.colValor.Add objNFiscal.lNumPedidoVenda
            
        Case MN_REGRA
            iLinha = objMnemonicoValor.vParam(1)
            'If iLinha > gcolRetRegras.Count Then gError 208272 'Não pode referenciar uma regra posterior
            'objMnemonicoValor.colValor.Add gcolRetRegras.Item(iLinha)
            iRegra = gcolRetRegras.Count + 1
            bAchou = False
            iItem = 0
            Set objRegra = gcolRegras.Item(iRegra)
            For Each objRegraAux In gcolRegras
                iItem = iItem + 1
                If iItem = iRegra Then Exit For
                If objRegraAux.iTipoDoc = objRegra.iTipoDoc And objRegraAux.iTipoMsg = objRegra.iTipoMsg And objRegraAux.iSeq = iLinha Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then gError 208272
            objMnemonicoValor.colValor.Add gcolRetRegras.Item(iItem)
            
        Case MN_FILEMPCGC
        
            Set objFilialEmpresa = New AdmFiliais
        
            objFilialEmpresa.iCodFilial = objNFiscal.iFilialEmpresa
            
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            objMnemonicoValor.colValor.Add objFilialEmpresa.sCgc
            
        Case MN_FILEMP
            objMnemonicoValor.colValor.Add objNFiscal.iFilialEmpresa
            
        Case MN_EXISTE_CATPRODITEM, MN_EXISTE_CATPROD
        
            bAchou = False
            iLinha = 0
            sCat = objMnemonicoValor.vParam(1)
            For Each objItemNFAux In objNFiscal.colItens
                
                Set objProdutoAux = New ClassProduto
                
                objProdutoAux.sCodigo = objItemNFAux.sProduto
                
                'Carrega a colCategoriaItem do objProdutoAux
                lErro = CF("Produto_Le_Categorias", objProdutoAux, objProdutoAux.colCategoriaItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objProdCat In objProdutoAux.colCategoriaItem
                    If UCase(objProdCat.sCategoria) = UCase(sCat) Then
                        If MN_EXISTE_CATPROD = objMnemonicoValor.sMnemonico Then
                            bAchou = True
                            Exit For
                        End If
                        sCatItem = objMnemonicoValor.vParam(2)
                        If UCase(objProdCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    End If
                Next
            Next
            objMnemonicoValor.colValor.Add bAchou
            
            
        Case MN_EXISTE_TIPOPROD
        
            bAchou = False
            For Each objItemNFAux In objNFiscal.colItens
                
                Set objProdutoAux = New ClassProduto
                
                objProdutoAux.sCodigo = objItemNFAux.sProduto

                lErro = CF("Produto_Le", objProdutoAux)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                
                If objProdutoAux.iTipo = objMnemonicoValor.vParam(1) Then
                    bAchou = True
                    Exit For
                End If
                
            Next
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_CODBACENDEST, MN_UFDEST, MN_ENDREFDEST
        
            Set objEnd = New ClassEndereco
       
            If objNFiscal.lCliente > 0 And objNFiscal.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = objNFiscal.lCliente
                objFilCli.iCodFilial = objNFiscal.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                            
                objEnd.lCodigo = objFilCli.lEndereco
            
            ElseIf objNFiscal.lFornecedor > 0 And objNFiscal.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                
                objFilForn.lCodFornecedor = objNFiscal.lFornecedor
                objFilForn.iCodFilial = objNFiscal.iFilialForn
            
                lErro = CF("FilialFornecedor_Le", objFilForn)
                If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
                        
                objEnd.lCodigo = objFilForn.lEndereco
            
            End If
                            
            If objEnd.lCodigo > 0 Then
            
                lErro = CF("Endereco_Le", objEnd)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                If objMnemonicoValor.sMnemonico = MN_UFDEST Then
                    objMnemonicoValor.colValor.Add objEnd.sSiglaEstado
                ElseIf objMnemonicoValor.sMnemonico = MN_ENDREFDEST Then
                    objMnemonicoValor.colValor.Add objEnd.sReferencia
                Else
                    Set objPais = New ClassPais
                    
                    objPais.iCodigo = objEnd.iCodigoPais
                    
                    lErro = CF("Paises_Le", objPais)
                    If lErro <> SUCESSO And lErro <> 47876 Then gError ERRO_SEM_MENSAGEM
                    
                    objMnemonicoValor.colValor.Add objPais.iCodBacen
                End If
                
            Else
                            
                Select Case objMnemonicoValor.sMnemonico
                
                    Case MN_CODBACENDEST
                        objMnemonicoValor.colValor.Add 1058 'codbacen do Brasil
                    
                    Case MN_UFDEST
                        objMnemonicoValor.colValor.Add "SP"
                    
                    Case MN_ENDREFDEST
                        objMnemonicoValor.colValor.Add ""
                    
                End Select
                
            End If
            
        Case MN_BASECREDICMS, MN_ALIQUOTACREDICMS, MN_VALORCREDICMS
        
            bAchou = False
            iLinha = 0
            dValorBase = 0
            dValorBaseAux = 0
            dAliquota = 0
            iAliqDif = 0
            dValor = 0
            
            For Each objItemNFAux In objNFiscal.colItens
                If objItemNFAux.objTributacao.dICMSvCredSN > 0 Then
                    'Se utilizar o total do item como base deixar os valores inconsitentes utiliza o valor calculado
                    If Abs((objItemNFAux.objTributacao.dPrecoTotal * objItemNFAux.objTributacao.dICMSpCredSN) - objItemNFAux.objTributacao.dICMSvCredSN) > DELTA_VALORMONETARIO Then
                        dValorBaseAux = Arredonda_Moeda(objItemNFAux.objTributacao.dICMSvCredSN / objItemNFAux.objTributacao.dICMSpCredSN)
                    Else
                        dValorBaseAux = objItemNFAux.objTributacao.dPrecoTotal
                    End If
                    dValorBase = dValorBase + dValorBaseAux
                    dValor = dValor + objItemNFAux.objTributacao.dICMSvCredSN
                    If dAliquota = 0 Then
                        dAliquota = objItemNFAux.objTributacao.dICMSpCredSN
                    ElseIf Abs(dAliquota - objItemNFAux.objTributacao.dICMSpCredSN) > DELTA_VALORMONETARIO2 Then
                    'Se existe mais de uma alíquota vai exibir a média
                        iAliqDif = MARCADO
                    End If
                End If
            Next
                       
            If objMnemonicoValor.sMnemonico = MN_BASECREDICMS Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            ElseIf objMnemonicoValor.sMnemonico = MN_ALIQUOTACREDICMS Then
                If iAliqDif = MARCADO Then
                    objMnemonicoValor.colValor.Add Format(dValor / dValorBase, "PERCENT")
                Else
                    objMnemonicoValor.colValor.Add Format(dAliquota, "PERCENT")
                End If
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
        
        Case MN_MSGPV
            If Not (objPedidoVenda Is Nothing) Then
                objMnemonicoValor.colValor.Add objPedidoVenda.sMensagemNota
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_MSGFILCLI
            
            sAux = ""
            
            If objNFiscal.lCliente > 0 And objNFiscal.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = objNFiscal.lCliente
                objFilCli.iCodFilial = objNFiscal.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
    
                If objFilCli.iCodMensagem > 0 Then
                
                    Set objMensagem = New ClassMensagem
                
                    objMensagem.iCodigo = objFilCli.iCodMensagem
                    
                    lErro = CF("Mensagem_Le", objMensagem)
                    If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro = SUCESSO Then sAux = objMensagem.sDescricao
                    
                Else
                
                    Set objCli = New ClassCliente
                    
                    objCli.lCodigo = objNFiscal.lCliente
            
                    lErro = CF("Cliente_Le", objCli)
                    If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                   
                   If objCli.iCodMensagem > 0 Then
                   
                       Set objMensagem = New ClassMensagem
                   
                       objMensagem.iCodigo = objCli.iCodMensagem
                       
                       lErro = CF("Mensagem_Le", objMensagem)
                       If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
               
                       If lErro = SUCESSO Then sAux = objMensagem.sDescricao
                       
                    End If
            
                End If
                
            End If

            objMnemonicoValor.colValor.Add sAux
            
        Case MN_ENDENTREGA
            
            sAux = ""
                       
            'If objNFiscal.lCliente > 0 And objNFiscal.iFilialCli > 0 Then
            
                If Not (objNFiscal.objInfoAdic Is Nothing) Then
                    If Not (objNFiscal.objInfoAdic.objRetEnt Is Nothing) Then
                        If Not (objNFiscal.objInfoAdic.objRetEnt.objEnderecoEnt Is Nothing) Then
                            Set objEnd = objNFiscal.objInfoAdic.objRetEnt.objEnderecoEnt
                            If Len(Trim(objEnd.sEndereco)) = 0 Then
                                objEnd.sEndereco = objEnd.sTipoLogradouro
                                objEnd.sEndereco = objEnd.sEndereco & IIf(Len(Trim(objEnd.sEndereco)) <> 0, " ", "") & objEnd.sLogradouro
                                objEnd.sEndereco = objEnd.sEndereco & IIf(objEnd.lNumero <> 0, ", " & CStr(objEnd.lNumero), "")
                                objEnd.sEndereco = objEnd.sEndereco & IIf(Len(Trim(objEnd.sComplemento)) <> 0, " - " & objEnd.sComplemento, "")
                            End If
                            If Len(Trim(objEnd.sEndereco)) > 0 Then
                                sAux = objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
                            End If
                        End If
                    End If
                End If
            
'                Set objFilCli = New ClassFilialCliente
'                Set objEnd = New ClassEndereco
'
'                objFilCli.lCodCliente = objNFiscal.lCliente
'                objFilCli.iCodFilial = objNFiscal.iFilialCli
'
'                lErro = CF("FilialCliente_Le", objFilCli)
'                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
'
'                objEnd.lCodigo = objFilCli.lEnderecoEntrega
'
'                lErro = CF("Endereco_Le", objEnd)
'                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'                If Len(Trim(objEnd.sEndereco)) > 0 Then
'                    sAux = objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
'                End If
                
            'End If
            
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_ENDTRANSREDESP
        
            sAux = ""
            If objNFiscal.iCodTranspRedesp <> 0 Then
        
                Set objTransp = New ClassTransportadora
                Set objEnd = New ClassEndereco
        
                objTransp.iCodigo = objNFiscal.iCodTranspRedesp
            
                lErro = CF("Transportadora_Le", objTransp)
                If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
            
                objEnd.lCodigo = objTransp.lEndereco
 
                lErro = CF("Endereco_Le", objEnd)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                 Select Case Len(Trim(objTransp.sCgc))
                
                     Case STRING_CPF
                         sCGCTranspRedesp = " CPF: " & Format(objTransp.sCgc, "000\.000\.000-00; ; ; ")
                     Case STRING_CGC
                         sCGCTranspRedesp = " CNPJ: " & Format(objTransp.sCgc, "00\.000\.000\/0000-00; ; ; ")
                     Case Else
                         sCGCTranspRedesp = ""
                 End Select
            
                If Len(Trim(objEnd.sEndereco)) > 0 Then
                    sAux = objTransp.sNomeReduzido & sCGCTranspRedesp & " - " & objEnd.sTelefone1 & " - " & objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
                End If
                
            End If
            
            objMnemonicoValor.colValor.Add sAux
                
        Case MN_FILCLIFORNINSCSUF, MN_FILCLIFORNINSCEST, MN_FILCLIFORNINSCMUN

            sAux = ""
        
            If objNFiscal.lCliente > 0 And objNFiscal.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = objNFiscal.lCliente
                objFilCli.iCodFilial = objNFiscal.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                            
                If objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCSUF Then
                    If Len(Trim(objFilCli.sInscricaoSuframa)) <> 0 Then sAux = Format(objFilCli.sInscricaoSuframa, "00\.0000\-00\-0")
                ElseIf objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCEST Then
                    sAux = objFilCli.sInscricaoEstadual
                Else
                    sAux = objFilCli.sInscricaoMunicipal
                End If

            ElseIf objNFiscal.lFornecedor > 0 And objNFiscal.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                
                objFilForn.lCodFornecedor = objNFiscal.lFornecedor
                objFilForn.iCodFilial = objNFiscal.iFilialForn
            
                lErro = CF("FilialFornecedor_Le", objFilForn)
                If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
            
                If objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCSUF Then
                    If Len(Trim(objFilForn.sInscricaoSuframa)) <> 0 Then sAux = Format(objFilForn.sInscricaoSuframa, "00\.0000\-00\-0")
                ElseIf objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCEST Then
                    sAux = objFilForn.sInscricaoEstadual
                Else
                    sAux = objFilForn.sInscricaoMunicipal
                End If
            
            End If
            
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_NUMPEDIDOTERC
            objMnemonicoValor.colValor.Add objNFiscal.sNumPedidoTerc
            
        Case MN_CLIFORN
            If objNFiscal.lCliente > 0 Then
                objMnemonicoValor.colValor.Add objNFiscal.lCliente
            Else
                objMnemonicoValor.colValor.Add objNFiscal.lFornecedor
            End If
                
        Case MN_CFOP
            objMnemonicoValor.colValor.Add objNFiscal.objTributacao.sNaturezaOp
             
        Case MN_TOTTRIB
            If objNFiscal.objTributacao.dTotTrib > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTrib, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIB
            If objNFiscal.objTributacao.dTotTrib > 0 And objNFiscal.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTrib / objNFiscal.dValorTotal, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBFED
            If objNFiscal.objTributacao.dTotTribFed > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribFed, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBFED
            If objNFiscal.objTributacao.dTotTribFed > 0 And objNFiscal.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribFed / objNFiscal.dValorTotal, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBEST
            If objNFiscal.objTributacao.dTotTribEst > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribEst, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBEST
            If objNFiscal.objTributacao.dTotTribEst > 0 And objNFiscal.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribEst / objNFiscal.dValorTotal, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBMUNIC
            If objNFiscal.objTributacao.dTotTribMunic > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribMunic, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBMUNIC
            If objNFiscal.objTributacao.dTotTribMunic > 0 And objNFiscal.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dTotTribMunic / objNFiscal.dValorTotal, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TIPOCLIFORN, MN_CLIFORNNOMERED

            bAchou = False
            If objNFiscal.lCliente > 0 Then
            
                Set objCli = New ClassCliente
                
                objCli.lCodigo = objNFiscal.lCliente
            
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add objCli.iTipo
                Else
                    objMnemonicoValor.colValor.Add objCli.sNomeReduzido
                End If
                
                bAchou = True
                
            ElseIf objNFiscal.lFornecedor > 0 Then
            
                Set objForn = New ClassFornecedor
                
                objForn.lCodigo = objNFiscal.lFornecedor
            
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add objForn.iTipo
                Else
                    objMnemonicoValor.colValor.Add objForn.sNomeReduzido
                End If
            
                bAchou = True
            
            End If
            
            If bAchou = False Then
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add 0
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            End If
            
        Case MN_POSSUI_CATCLIFORITEM, MN_POSSUI_CATCLIFORN
            
            bAchou = False
            sCat = objMnemonicoValor.vParam(1)
            
            If objNFiscal.lCliente > 0 And objNFiscal.iFilialCli > 0 Then
                Set objFilCli = New ClassFilialCliente
                Set colCatItem = New Collection
                
                objFilCli.lCodCliente = objNFiscal.lCliente
                objFilCli.iCodFilial = objNFiscal.iFilialCli
                
                lErro = CF("FilialCliente_Le_Categorias", objFilCli, colCatItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objFilCliCat In colCatItem
                    
                    If UCase(objFilCliCat.sCategoria) = UCase(sCat) Then
                        If MN_POSSUI_CATCLIFORN = objMnemonicoValor.sMnemonico Then
                            bAchou = True
                            Exit For
                        End If
                        sCatItem = objMnemonicoValor.vParam(2)
                        If UCase(objFilCliCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    End If
                Next
            
            ElseIf objNFiscal.lFornecedor > 0 And objNFiscal.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                Set colCatItem = New Collection
                
                objFilForn.lCodFornecedor = objNFiscal.lFornecedor
                objFilForn.iCodFilial = objNFiscal.iFilialForn
                
                lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilForn, colCatItem)
                If lErro <> SUCESSO And lErro <> 127085 Then gError ERRO_SEM_MENSAGEM
                
                For Each objFilFornCat In colCatItem
                    
                    If UCase(objFilFornCat.sCategoria) = UCase(sCat) Then
                        If MN_POSSUI_CATCLIFORN = objMnemonicoValor.sMnemonico Then
                            bAchou = True
                            Exit For
                        End If
                        sCatItem = objMnemonicoValor.vParam(2)
                        If UCase(objFilFornCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    End If
                Next
            End If

            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_POSSUI_ICMSST
            If objNFiscal.objTributacao.dICMSSubstValor > 0 Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_POSSUI_ICMSRED
        
            bAchou = False
            
            For Each objItemNFAux In objNFiscal.colItens
                
                If objItemNFAux.objTributacaoItemNF.dICMSPercRedBase <> 0 Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Next
            
            If bAchou Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
        
        
        Case MN_INTERNA
        
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            If objTipoDocInfo.iTipo = 3 Then
                objMnemonicoValor.colValor.Add False
            Else
                objMnemonicoValor.colValor.Add True
            End If
            
        Case MN_EXTERNA
                
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            If objTipoDocInfo.iTipo = 3 Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_SERIE
            objMnemonicoValor.colValor.Add objNFiscal.sSerie
            
        Case MN_VLRDESCONTONF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorDesconto, "STANDARD")
                
        Case MN_VLRTOTALNF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorTotal, "STANDARD")
                
        Case MN_VLRFRETENF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorFrete, "STANDARD")
                
        Case MN_VLRSEGURONF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorSeguro, "STANDARD")
                
        Case MN_VLRDESPESASNF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorOutrasDespesas, "STANDARD")
                
        Case MN_VLRPRODUTONF
            objMnemonicoValor.colValor.Add Format(objNFiscal.dValorProdutos, "STANDARD")
                
        'Exclusivas dos itens
        Case MN_PRODUTO
            objMnemonicoValor.colValor.Add objItemNF.sProduto
            
        Case MN_INFOADIC_TXTAD
            objMnemonicoValor.colValor.Add left(Trim(objItemNF.objInfoAdicDocItem.sMsg), 255)
            
        Case MN_INFOADIC_PEDCOMPRAITEM
            objMnemonicoValor.colValor.Add objItemNF.objInfoAdicDocItem.sNumPedidoCompra
        
        Case MN_INFOADIC_ITEMPEDCOMPRAITEM
            If objItemNF.objInfoAdicDocItem.lItemPedCompra <> 0 Then
                objMnemonicoValor.colValor.Add CStr(objItemNF.objInfoAdicDocItem.lItemPedCompra)
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TIPOPROD
            objMnemonicoValor.colValor.Add objProduto.iTipo
        
        Case MN_PROD_OBS
            objMnemonicoValor.colValor.Add objProduto.sObsFisica
        
        Case MN_NCM
            objMnemonicoValor.colValor.Add objItemNF.objTributacao.sIPICodProduto
        
        Case MN_CFOPITEM
            objMnemonicoValor.colValor.Add objItemNF.objTributacao.sNaturezaOp
        
        Case MN_TOTTRIBITEM
            If objItemNF.objTributacao.dTotTrib > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTrib, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                        
        Case MN_PERCTOTTRIBITEM
            If objItemNF.objTributacao.dTotTrib > 0 And objItemNF.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTrib / objItemNF.dValorTotal, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBFEDITEM
            If objItemNF.objTributacao.dTotTribAliqFed > 0 And objItemNF.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqFed * objItemNF.dValorTotal, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                        
        Case MN_PERCTOTTRIBFEDITEM
            If objItemNF.objTributacao.dTotTribAliqFed > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqFed, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBESTITEM
            If objItemNF.objTributacao.dTotTribAliqEst > 0 And objItemNF.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqEst * objItemNF.dValorTotal, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                        
        Case MN_PERCTOTTRIBESTITEM
            If objItemNF.objTributacao.dTotTribAliqEst > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqEst, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBMUNICITEM
            If objItemNF.objTributacao.dTotTribAliqMunic > 0 And objItemNF.dValorTotal > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqMunic * objItemNF.dValorTotal, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                        
        Case MN_PERCTOTTRIBMUNICITEM
            If objItemNF.objTributacao.dTotTribAliqMunic > 0 Then
                objMnemonicoValor.colValor.Add Format(objItemNF.objTributacao.dTotTribAliqMunic, "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_FCI
            objMnemonicoValor.colValor.Add objItemNF.objTributacao.sFCI
            
        Case MN_POSSUI_CATPROD, MN_POSSUI_CATPRODITEM
        
            bAchou = False
            sCat = objMnemonicoValor.vParam(1)
            For Each objProdCat In objProduto.colCategoriaItem
                If UCase(objProdCat.sCategoria) = UCase(sCat) Then
                    If MN_POSSUI_CATPROD = objMnemonicoValor.sMnemonico Then
                        bAchou = True
                        Exit For
                    End If
                    sCatItem = objMnemonicoValor.vParam(2)
                    If UCase(objProdCat.sItem) = UCase(sCatItem) Then
                        bAchou = True
                        Exit For
                    End If
                End If
            Next
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_CATPRODITEM, MN_CATPRODITEMDESC
            sAux = ""
            sCat = objMnemonicoValor.vParam(1)
            For Each objProdCat In objProduto.colCategoriaItem
                If UCase(objProdCat.sCategoria) = UCase(sCat) Then
                    sAux = objProdCat.sItem
                    Exit For
                End If
            Next
            If objMnemonicoValor.sMnemonico = MN_CATPRODITEMDESC And sAux <> "" Then
                Set objCatProdItem = New ClassCategoriaProdutoItem
                
                objCatProdItem.sCategoria = sCat
                objCatProdItem.sItem = sAux
                
                lErro = CF("CategoriaProduto_Le_Item", objCatProdItem)
                If lErro <> SUCESSO And lErro <> 22603 Then gError ERRO_SEM_MENSAGEM
                
                sAux = objCatProdItem.sDescricao
            End If
            objMnemonicoValor.colValor.Add sAux

        Case MN_NUMLOTES
            objMnemonicoValor.colValor.Add objItemNF.colRastreamento.Count

        Case MN_POSSUI_RASTREAMENTO
            If objItemNF.colRastreamento.Count > 0 Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_LOTE
            sAux = ""
            iItem = objMnemonicoValor.vParam(1)
            If objItemNF.colRastreamento.Count >= iItem Then
                Set objRastro = objItemNF.colRastreamento.Item(iItem)
                sAux = objRastro.sLote
            End If
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_QTDLOTE
            sAux = ""
            iItem = objMnemonicoValor.vParam(1)
            If objItemNF.colRastreamento.Count >= iItem Then
                Set objRastro = objItemNF.colRastreamento.Item(iItem)
                sAux = Formata_Estoque(objRastro.dLoteQdtAlocada)
            End If
            objMnemonicoValor.colValor.Add sAux

        Case MN_LOTEOBS, MN_FABRICACAOLOTE, MN_VALIDADELOTE
            sAux = ""
            iItem = objMnemonicoValor.vParam(1)
            If objItemNF.colRastreamento.Count >= iItem Then
                Set objRastro = objItemNF.colRastreamento.Item(iItem)
                Set objLote = New ClassRastreamentoLote
                objLote.sCodigo = objRastro.sLote
                objLote.sProduto = objItemNF.sProduto
                objLote.iFilialOP = objRastro.iLoteFilialOP
                lErro = CF("RastreamentoLote_Le", objLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError ERRO_SEM_MENSAGEM

                If objMnemonicoValor.sMnemonico = MN_FABRICACAOLOTE Then
                    If objLote.dtDataFabricacao <> DATA_NULA Then
                        sAux = Format(objLote.dtDataFabricacao, "dd/mm/yyyy")
                    Else
                        sAux = ""
                    End If
                ElseIf objMnemonicoValor.sMnemonico = MN_VALIDADELOTE Then
                    If objLote.dtDataValidade <> DATA_NULA Then
                        sAux = Format(objLote.dtDataValidade, "dd/mm/yyyy")
                    Else
                        sAux = ""
                    End If
                Else
                    sAux = objLote.sObservacao
                End If
            
            End If
            objMnemonicoValor.colValor.Add sAux
        
        Case MN_TODOSLOTES
            sAux = ""
            For Each objRastro In objItemNF.colRastreamento
                sAux = sAux & objRastro.sLote & "/"
            Next
            If Len(sAux) <> 0 Then sAux = left(sAux, Len(sAux) - 1)
            objMnemonicoValor.colValor.Add sAux
        
        Case MN_TODOSPEDCLI
            sAux = ""
            Set colAux = New Collection
            For Each objDataEntrega In objItemNF.colDataEntrega
                bAchou = False
                For Each vValor In colAux
                    If UCase(Trim(vValor)) = UCase(Trim(objDataEntrega.sPedidoCliente)) Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    sAux = sAux & Trim(objDataEntrega.sPedidoCliente) & "/"
                    colAux.Add objDataEntrega.sPedidoCliente
                End If
            Next
            If Len(sAux) <> 0 Then sAux = left(sAux, Len(sAux) - 1)
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_DEVOLUCAO
        
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM

            If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalEntDev" Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_NUMNFORIGDETBENEF
        
            objMnemonicoValor.colValor.Add ""
            
        Case MN_TIPONFORIG, MN_NUMNFORIG, MN_NUMNFORIGDET, MN_CFOPNFORIG, MN_EMISSAONFORIG, MN_CLIFORNNFORIG, MN_NUMNFORIGVINC
        
            If objNFiscal.lNumIntNotaOriginal > 0 Then
        
                Set objNFiscalOriginal = New ClassNFiscal
                objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
                
                'Lâ Nota Fiscal Original
                lErro = CF("NFiscal_Le", objNFiscalOriginal)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPONFORIG Then
                    objMnemonicoValor.colValor.Add objNFiscalOriginal.iTipoNFiscal
                ElseIf objMnemonicoValor.sMnemonico = MN_CLIFORNNFORIG Then
                    If objNFiscalOriginal.lCliente > 0 Then
                        objMnemonicoValor.colValor.Add objNFiscalOriginal.lCliente
                    Else
                        objMnemonicoValor.colValor.Add objNFiscalOriginal.lFornecedor
                    End If
                ElseIf objMnemonicoValor.sMnemonico = MN_NUMNFORIG Or objMnemonicoValor.sMnemonico = MN_NUMNFORIGDET Then
                    objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal)
                ElseIf objMnemonicoValor.sMnemonico = MN_CFOPNFORIG Then
                    objMnemonicoValor.colValor.Add objNFiscalOriginal.sNaturezaOp
                ElseIf objMnemonicoValor.sMnemonico = MN_EMISSAONFORIG Then
                    objMnemonicoValor.colValor.Add Format(objNFiscalOriginal.dtDataEmissao, "dd/mm/yyyy")
                ElseIf objMnemonicoValor.sMnemonico = MN_NUMNFORIGVINC Then
                    Set X = New ClassFAT
                    Call X.Inicializa
                    If objNFiscalOriginal.iTipoNFiscal = DOCINFO_NFISSRGOV Then
                        objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal + 1)
                    ElseIf objNFiscalOriginal.lCliente = X.lCodCliGov And objNFiscalOriginal.lCliente <> 0 Then
                        objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal - 1)
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                End If
                
            Else
        
                If objMnemonicoValor.sMnemonico = MN_TIPONFORIG Or objMnemonicoValor.sMnemonico = MN_CLIFORNNFORIG Then
                    objMnemonicoValor.colValor.Add 0
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            End If
            
        Case MN_VLRICMS
            objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSValor, "STANDARD")
            
        Case MN_VLRICMSINTERESTFCP
            objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSInterestVlrFCPUFDest, "STANDARD")
            
        Case MN_VLRICMSINTERESTUFDEST
             objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSInterestVlrUFDest, "STANDARD")
           
        Case MN_VLRICMSINTERESTUFREMET
            objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSInterestVlrUFRemet, "STANDARD")
            
        Case MN_VLRICMSST
            objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSSubstValor, "STANDARD")
            
        Case MN_BASEICMSST
            objMnemonicoValor.colValor.Add Format(objNFiscal.objTributacao.dICMSSubstBase, "STANDARD")
        
        Case MN_VLRICMSSTCOBRANT, MN_BASEICMSSTCOBRANT
        
            dValorBase = 0
            dValor = 0
            For Each objItemNFAux In objNFiscal.colItens
                dValorBase = dValorBase + objItemNFAux.objTributacao.dICMSSTCobrAntBase
                dValor = dValor + objItemNFAux.objTributacao.dICMSSTCobrAntValor
            Next
            
            If objMnemonicoValor.sMnemonico = MN_BASEICMSSTCOBRANT Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
            
        Case MN_BASEICMSSTNFORIG, MN_VLRICMSSTNFORIG
            dValorBase = 0
            dValor = 0
        
            For Each objItemNFAux In objNFiscal.colItens
            
                If objItemNFAux.lNumIntDocOrig <> 0 Then
                
                    Set objItemNFOrig = New ClassItemNF
                    objItemNFOrig.lNumIntDoc = objItemNFAux.lNumIntDocOrig
                    
                    lErro = CF("ItemNFiscal_Le", objItemNFOrig)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                       
                    Set objItemNFOrig.objTributacao = New ClassTributacaoDocItem
                    objItemNFOrig.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
                    objItemNFOrig.objTributacao.lNumIntDocItem = objItemNFOrig.lNumIntDoc
                    lErro = CF("TributacaoDocItem_Le", objItemNFOrig.objTributacao)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                    dValorBase = dValorBase + (objItemNFOrig.objTributacao.dICMSSubstBase / objItemNFOrig.dQuantidade) * objItemNFAux.dQuantidade
                    dValor = dValor + (objItemNFOrig.objTributacao.dICMSSubstValor / objItemNFOrig.dQuantidade) * objItemNFAux.dQuantidade
                
                End If
                
            Next
        
            If objMnemonicoValor.sMnemonico = MN_BASEICMSSTNFORIG Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
        Case Else
             gError 208271

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 208271
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 208272
            Call Rotina_Erro(vbOKOnly, "ERRO_REGRASMSG_MN_REGRA_LIN_POS", gErr)

        Case 208273
            Call Rotina_Erro(vbOKOnly, "ERRO_DASALIQUOTAS_NAO_CADASTRADO", gErr, objDASAliquotas.iAno, objDASAliquotas.iMes)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208274)

    End Select
    
    Resume Next

    Exit Function

End Function

Public Function Inicializa_Calculo_RegrasMsg() As Long
    Set gcolRetRegras = New Collection
End Function

Public Function Inicializa_Calculo_RegrasMsg2(ByVal colRegras As Collection) As Long
    Set gcolRetRegras = New Collection
    Set gcolRegras = colRegras
End Function

Public Function Retorno_Calculo_RegrasMsg(ByVal bFlag As Boolean) As Long
    gcolRetRegras.Add bFlag
End Function

Public Function Finaliza_Calculo_RegrasMsg() As Long
    Set gcolRetRegras = Nothing
    Set gcolRegras = Nothing
End Function

Private Sub Class_Terminate()
    Set objNFiscal = Nothing
    Set objPedidoVenda = Nothing
    Set objItemNF = Nothing
    Set objProduto = Nothing
End Sub
