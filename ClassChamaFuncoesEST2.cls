VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesEST"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ESTConfig_Grava_Configs(colConfigs As ColESTConfig) As Long
'Grava na tabela ESTConfig uma coleção de registros de configuração

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ESTConfig_Grava_Configs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CONFIG_GRAVA_CONFIGS", X)
    If lErro <> SUCESSO Then Error 33661
    
    
    ESTConfig_Grava_Configs = X.ESTConfig_Grava_Configs(colConfigs)

    Exit Function
    
Erro_ESTConfig_Grava_Configs:
    
    ESTConfig_Grava_Configs = Err
    
    Select Case Err

        Case 33661
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145548)
        
    End Select
    
    Exit Function

End Function

Function ESTConfig_Le_Configs(colConfigs As ColESTConfig) As Long
'Recupera da tabela ESTConfig uma coleção de registros

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ESTConfig_Le_Configs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CONFIG_LE_CONFIGS", X)
    If lErro <> SUCESSO Then Error 33662
    
    
    ESTConfig_Le_Configs = X.ESTConfig_Le_Configs(colConfigs)

    Exit Function
    
Erro_ESTConfig_Le_Configs:
    
    ESTConfig_Le_Configs = Err
    
    Select Case Err

        Case 33662
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145549)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Atualizacao_InvLote(sNomeArqParam As String, iID_Atualizacao As Integer, dtDataContabil As Date, iLoteContabil As Integer) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Atualizacao_InvLote
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ROTINA_ATUALIZACAO_INVLOTE", X)
    If lErro <> SUCESSO Then Error 36710
        
    Rotina_Atualizacao_InvLote = X.Rotina_Atualizacao_InvLote(sNomeArqParam, iID_Atualizacao, dtDataContabil, iLoteContabil)

    Exit Function
    
Erro_Rotina_Atualizacao_InvLote:
    
    Rotina_Atualizacao_InvLote = Err
    
    Select Case Err

        Case 36710
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145550)
        
    End Select
    
    Exit Function

End Function

Function Rotina_CustoMedioProducao_Calcula(sNomeArqParam As String, iAno As Integer, iMes As Integer) As Long
'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_CustoMedioProducao_Calcula
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ROTINA_CUSTOMEDIOPRODUCAO_CALCULA", X)
    If lErro <> SUCESSO Then Error 45208
    
    
    Rotina_CustoMedioProducao_Calcula = X.Rotina_CustoMedioProducao_Calcula(sNomeArqParam, iAno, iMes)

    Exit Function
    
Erro_Rotina_CustoMedioProducao_Calcula:
    
    Rotina_CustoMedioProducao_Calcula = Err
    
    Select Case Err

        Case 45208
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145551)
        
    End Select
    
    Exit Function

End Function

Function Almoxarifado_Automatico(iCodigo As Integer) As Long
'Obtem o próximo número da Almoxarifado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifado_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ALMOXARIFADO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 45654
    
    
    Almoxarifado_Automatico = X.Almoxarifado_Automatico(iCodigo)
    
    Exit Function
    
Erro_Almoxarifado_Automatico:

    Almoxarifado_Automatico = Err
    
    Select Case Err

        Case 45654
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145552)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntradaDevolucao_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz a gravação de Uma Nota Fiscal de Entrada-Devolução

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaDevolucao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADADEVOLUCAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45717
    
    
    NFiscalEntradaDevolucao_Grava = X.NFiscalEntradaDevolucao_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalEntradaDevolucao_Grava:

    NFiscalEntradaDevolucao_Grava = Err
    
    Select Case Err

        Case 45717
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145553)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntradaDevolucao_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Entrada-Devolução no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaDevolucao_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADADEVOLUCAO_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 45698
    
    
    NFiscalEntradaDevolucao_Grava_BD = X.NFiscalEntradaDevolucao_Grava_BD(objNFiscal)
    
    Exit Function
    
Erro_NFiscalEntradaDevolucao_Grava_BD:

    NFiscalEntradaDevolucao_Grava_BD = Err
    
    Select Case Err

        Case 45698
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145554)
        
    End Select
    
    Exit Function
    
End Function

Function ItensNFiscalEntDev_Lock_Gravacao(objNFiscal As ClassNFiscal) As Long
'Faz a validação e os "lock's" dos dados ligados aos ítens da Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensNFiscalEntDev_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSNFISCALENTDEV_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then Error 45699
    
    
    ItensNFiscalEntDev_Lock_Gravacao = X.ItensNFiscalEntDev_Lock_Gravacao(objNFiscal)
    
    Exit Function
    
Erro_ItensNFiscalEntDev_Lock_Gravacao:

    ItensNFiscalEntDev_Lock_Gravacao = Err
    
    Select Case Err

        Case 45699
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145555)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntDevRem_Baixa_Recebimento(objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntDevRem_Baixa_Recebimento
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTDEVREM_BAIXA_RECEBIMENTO", X)
    If lErro <> SUCESSO Then Error 45700
    
    
    NFiscalEntDevRem_Baixa_Recebimento = X.NFiscalEntDevRem_Baixa_Recebimento(objNFiscal)
    
    Exit Function
    
Erro_NFiscalEntDevRem_Baixa_Recebimento:

    NFiscalEntDevRem_Baixa_Recebimento = Err
    
    Select Case Err

        Case 45700
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145556)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_NumFornCli(objNFiscal As ClassNFiscal) As Long
'Le a Nota Fiscal usando Serie, NumNotaFiscal, Cliente, FilialCli, Fornecedor, FilialForn, TipoNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_NumFornCli
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCAL_LE_NUMFORNCLI", X)
    If lErro <> SUCESSO Then Error 45701
    
    
    NFiscal_Le_NumFornCli = X.NFiscal_Le_NumFornCli(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Le_NumFornCli:

    NFiscal_Le_NumFornCli = Err
    
    Select Case Err

        Case 45701
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145557)
        
    End Select
    
    Exit Function
    
End Function

'mario 2/7/99 Remover NFiscalBaixada comando
'Function ItemNFiscalBaixadaSaida_Le_NumNFItem(objItemNF As ClassItemNF) As Long
''Lê o Ítem de Uma Nota Fiscal de Saida Baixada através do Número da Nota Fiscal, Série e Numero do Ítem passados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ItemNFiscalBaixadaSaida_Le_NumNFItem
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFISCALBAIXADASAIDA_LE_NUMNFITEM",X)
'    If lErro <> SUCESSO Then Error 45702
'
'
'    ItemNFiscalBaixadaSaida_Le_NumNFItem = X.ItemNFiscalBaixadaSaida_Le_NumNFItem(objItemNF)
'
'    Exit Function
'
'Erro_ItemNFiscalBaixadaSaida_Le_NumNFItem:
'
'    ItemNFiscalBaixadaSaida_Le_NumNFItem = Err
'
'    Select Case Err
'
'        Case 45702
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145558)
'
'    End Select
'
'    Exit Function
'
'End Function

Function ItemNFiscalSaida_Le_NumNFItem(objItemNF As ClassItemNF) As Long
'Lê o Ítem de Uma Nota Fiscal de Saida através do Número da Nota Fiscal, Série e Numero do Ítem passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFiscalSaida_Le_NumNFItem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFISCALSAIDA_LE_NUMNFITEM", X)
    If lErro <> SUCESSO Then Error 45703
    
    
    ItemNFiscalSaida_Le_NumNFItem = X.ItemNFiscalSaida_Le_NumNFItem(objItemNF)
    
    Exit Function
    
Erro_ItemNFiscalSaida_Le_NumNFItem:

    ItemNFiscalSaida_Le_NumNFItem = Err
    
    Select Case Err

        Case 45703
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145559)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntrada_Verifica_Existencia(objNFiscal As ClassNFiscal, ByVal iTipo As Integer) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Verifica_Existencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADA_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then Error 45704
    
    
    NFiscalEntrada_Verifica_Existencia = X.NFiscalEntrada_Verifica_Existencia(objNFiscal, iTipo)
    
    Exit Function
    
Erro_NFiscalEntrada_Verifica_Existencia:

    NFiscalEntrada_Verifica_Existencia = Err
    
    Select Case Err

        Case 45704
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145560)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDocInfo_Le_SiglaNFOriginal(iTipoNFiscal As Integer, objTipoDocInfo As ClassTipoDocInfo) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDocInfo_Le_SiglaNFOriginal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_TIPOSDOCINFO_LE_SIGLANFORIGINAL", X)
    If lErro <> SUCESSO Then Error 45705
    
    
    TiposDocInfo_Le_SiglaNFOriginal = X.TiposDocInfo_Le_SiglaNFOriginal(iTipoNFiscal, objTipoDocInfo)
    
    Exit Function
    
Erro_TiposDocInfo_Le_SiglaNFOriginal:

    TiposDocInfo_Le_SiglaNFOriginal = Err
    
    Select Case Err

        Case 45705
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145561)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalFatEntrada_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntrada_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADA_GRAVA", X)
    If lErro <> SUCESSO Then Error 45707
    
    
    NFiscalFatEntrada_Grava = X.NFiscalFatEntrada_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalFatEntrada_Grava:

    NFiscalFatEntrada_Grava = Err
    
    Select Case Err

        Case 45707
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145562)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalFatEntrada_Lock_Gravacao(objNFiscal As ClassNFiscal) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntrada_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then Error 45708
    
    
    NFiscalFatEntrada_Lock_Gravacao = X.NFiscalFatEntrada_Lock_Gravacao(objNFiscal)
    
    Exit Function
    
Erro_NFiscalFatEntrada_Lock_Gravacao:

    NFiscalFatEntrada_Lock_Gravacao = Err
    
    Select Case Err

        Case 45708
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145563)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalFatEntrada_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Insere uma Nova Nota Fiscal Fatura de Entrada no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntrada_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADA_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 45709
    
    
    NFiscalFatEntrada_Grava_BD = X.NFiscalFatEntrada_Grava_BD(objNFiscal)
    
    Exit Function
    
Erro_NFiscalFatEntrada_Grava_BD:

    NFiscalFatEntrada_Grava_BD = Err
    
    Select Case Err

        Case 45709
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145564)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalFatEntrada_Baixa_Recebimento(objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntrada_Baixa_Recebimento
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADA_BAIXA_RECEBIMENTO", X)
    If lErro <> SUCESSO Then Error 45710
    
    
    NFiscalFatEntrada_Baixa_Recebimento = X.NFiscalFatEntrada_Baixa_Recebimento(objNFiscal)
    
    Exit Function
    
Erro_NFiscalFatEntrada_Baixa_Recebimento:

    NFiscalFatEntrada_Baixa_Recebimento = Err
    
    Select Case Err

        Case 45710
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145565)
        
    End Select
    
    Exit Function
    
End Function

Function MovEstoque_ReqProducao_Grava(objMovEstoque As ClassMovEstoque, objContabil As ClassContabil) As Long
'Grava o Movimento de Estoque de Requisição de Produção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_ReqProducao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_MOVESTOQUE_REQPRODUCAO_GRAVA", X)
    If lErro <> SUCESSO Then Error 45718
    
    
    MovEstoque_ReqProducao_Grava = X.MovEstoque_ReqProducao_Grava(objMovEstoque, objContabil)
    
    Exit Function
    
Erro_MovEstoque_ReqProducao_Grava:

    MovEstoque_ReqProducao_Grava = Err
    
    Select Case Err

        Case 45718
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145566)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutoKit_Le_Raiz(objProdutoKit As ClassProdutoKit) As Long
'Le a tabela de Kits de Produtos a partir do produto e da versão passados em objProdutoKit e nível = KIT_NIVEL_RAIZ

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoKit_Le_Raiz
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_PRODUTOKIT_LE_RAIZ", X)
    If lErro <> SUCESSO Then Error 45719
    
    
    ProdutoKit_Le_Raiz = X.ProdutoKit_Le_Raiz(objProdutoKit)
    
    Exit Function
    
Erro_ProdutoKit_Le_Raiz:

    ProdutoKit_Le_Raiz = Err
    
    Select Case Err

        Case 45719
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145567)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutosKit_Le_PrimeiroNivel(objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir do produto e da versão passados em objKit e Nivel = KIT_NIVEL_RAIZ + 1

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutosKit_Le_PrimeiroNivel
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_PRODUTOKIT_LE_PRIMEIRONIVEL", X)
    If lErro <> SUCESSO Then Error 45720
    
    
    ProdutosKit_Le_PrimeiroNivel = X.ProdutosKit_Le_PrimeiroNivel(objKit)
    
    Exit Function
    
Erro_ProdutosKit_Le_PrimeiroNivel:

    ProdutosKit_Le_PrimeiroNivel = Err
    
    Select Case Err

        Case 45720
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145568)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialC_Grava(objNFiscal As ClassNFiscal) As Long
'Grava a Nota Fiscal de Recebimento de Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialC_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALC_GRAVA", X)
    If lErro <> SUCESSO Then Error 45721
    
    
    RecebMaterialC_Grava = X.RecebMaterialC_Grava(objNFiscal)
    
    Exit Function
    
Erro_RecebMaterialC_Grava:

    RecebMaterialC_Grava = Err
    
    Select Case Err

        Case 45721
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145569)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialC_Exclui(objNFiscal As ClassNFiscal) As Long
'Exclui a Nota Fiscal de Recebimento de material de Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialC_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALC_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45722
    
    
    RecebMaterialC_Exclui = X.RecebMaterialC_Exclui(objNFiscal)
    
    Exit Function
    
Erro_RecebMaterialC_Exclui:

    RecebMaterialC_Exclui = Err
    
    Select Case Err

        Case 45722
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145570)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalInternaSaidaNaoDev_Le_Numero(objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_SAIDA
'          NFiscal.Status<>STATUS_EXCLUIDO
'          NomeTelaNFiscal<>"NFiscalDev"
'          FilialEmpresa = giFilialEmpresa
'          Complementar = DOCINFO_NORMAL

'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalInternaSaidaNaoDev_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALINTERNASAIDANAODEV_LE_NUMERO", X)
    If lErro <> SUCESSO Then Error 45723
    
    
    NFiscalInternaSaidaNaoDev_Le_Numero = X.NFiscalInternaSaidaNaoDev_Le_Numero(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalInternaSaidaNaoDev_Le_Numero:

    NFiscalInternaSaidaNaoDev_Le_Numero = Err
    
    Select Case Err

        Case 45723
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145571)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialC_Pesquisa_NotaFiscal(objNFiscal As ClassNFiscal, lComando1 As Long) As Long
'Pesquisa a Nota Fiscal de Recebimento de Material

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialC_Pesquisa_NotaFiscal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALC_PESQUISA_NOTAFISCAL", X)
    If lErro <> SUCESSO Then Error 45724
    
    
    RecebMaterialC_Pesquisa_NotaFiscal = X.RecebMaterialC_Pesquisa_NotaFiscal(objNFiscal, lComando1)
    
    Exit Function
    
Erro_RecebMaterialC_Pesquisa_NotaFiscal:

    RecebMaterialC_Pesquisa_NotaFiscal = Err
    
    Select Case Err

        Case 45724
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145572)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialF_Grava(objNFiscal As ClassNFiscal) As Long
'Grava a Nota Fiscal de Recebimento de Material de Fornecedor

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialF_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALF_GRAVA", X)
    If lErro <> SUCESSO Then Error 45725
    
    
    RecebMaterialF_Grava = X.RecebMaterialF_Grava(objNFiscal)
    
    Exit Function
    
Erro_RecebMaterialF_Grava:

    RecebMaterialF_Grava = Err
    
    Select Case Err

        Case 45725
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145573)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialF_Exclui(objNFiscal As ClassNFiscal) As Long
'Exclui a Nota Fiscal de Recebimento de material de Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialF_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALF_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45726
    
    
    RecebMaterialF_Exclui = X.RecebMaterialF_Exclui(objNFiscal)
    
    Exit Function
    
Erro_RecebMaterialF_Exclui:

    RecebMaterialF_Exclui = Err
    
    Select Case Err

        Case 45726
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145574)
        
    End Select
    
    Exit Function
    
End Function

Function RecebMaterialF_Pesquisa_NotaFiscal(objNFiscal As ClassNFiscal, lComando1 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialF_Pesquisa_NotaFiscal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALF_PESQUISA_NOTAFISCAL", X)
    If lErro <> SUCESSO Then Error 45727
    
    
    RecebMaterialF_Pesquisa_NotaFiscal = X.RecebMaterialF_Pesquisa_NotaFiscal(objNFiscal, lComando1)
    
    Exit Function
    
Erro_RecebMaterialF_Pesquisa_NotaFiscal:

    RecebMaterialF_Pesquisa_NotaFiscal = Err
    
    Select Case Err

        Case 45727
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145575)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntradaRemessa_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Entrada-Remessa no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaRemessa_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADAREMESSA_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 45742
    
    
    NFiscalEntradaRemessa_Grava_BD = X.NFiscalEntradaRemessa_Grava_BD(objNFiscal)
    
    Exit Function
    
Erro_NFiscalEntradaRemessa_Grava_BD:

    NFiscalEntradaRemessa_Grava_BD = Err
    
    Select Case Err

        Case 45742
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145576)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntradaRemessa_Lock_Gravacao(objNFiscal As ClassNFiscal, lComando As Long) As Long
'Faz o lock dos campos da NFiszcal Entrada Remessa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaRemessa_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADAREMESSA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then Error 45743
    
    
    NFiscalEntradaRemessa_Lock_Gravacao = X.NFiscalEntradaRemessa_Lock_Gravacao(objNFiscal, lComando)
    
    Exit Function
    
Erro_NFiscalEntradaRemessa_Lock_Gravacao:

    NFiscalEntradaRemessa_Lock_Gravacao = Err
    
    Select Case Err

        Case 45743
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145577)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalEntradaRemessa_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz as validações necessárias a gravação de Uma Nota Fiscal de Entrada-Remessa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaRemessa_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADAREMESSA_GRAVA", X)
    If lErro <> SUCESSO Then Error 45744
    
    
    NFiscalEntradaRemessa_Grava = X.NFiscalEntradaRemessa_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalEntradaRemessa_Grava:

    NFiscalEntradaRemessa_Grava = Err
    
    Select Case Err

        Case 45744
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145578)
        
    End Select
    
    Exit Function
    
End Function

Function Custos_Grava(iMes As Integer, dSaldoReposicao As Double, objSldMesEst As ClassSldMesEst, objTela As Object) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Custos_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CUSTOS_GRAVA", X)
    If lErro <> SUCESSO Then Error 45765
    
    
    Custos_Grava = X.Custos_Grava(iMes, dSaldoReposicao, objSldMesEst, objTela)
    
    Exit Function
    
Erro_Custos_Grava:

    Custos_Grava = Err
    
    Select Case Err

        Case 45765
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145579)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Grava(objReserva As ClassReserva) As Long
'Grava uma Reserva

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_GRAVA", X)
    If lErro <> SUCESSO Then Error 45766
    
    
    Reserva_Grava = X.Reserva_Grava(objReserva)
    
    Exit Function
    
Erro_Reserva_Grava:

    Reserva_Grava = Err
    
    Select Case Err

        Case 45766
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145580)
        
    End Select
    
    Exit Function
    
End Function

Function Reservas_Le_Codigo(colCodigo As Collection) As Long
'Lê o código de todas as reservas no Bd

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reservas_Le_Codigo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVAS_LE_CODIGO", X)
    If lErro <> SUCESSO Then Error 45767
    
    
    Reservas_Le_Codigo = X.Reservas_Le_Codigo(colCodigo)
    
    Exit Function
    
Erro_Reservas_Le_Codigo:

    Reservas_Le_Codigo = Err
    
    Select Case Err

        Case 45767
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145581)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Le(objReserva As ClassReserva) As Long
'Lê a reserva a partir do código
'Filtra por FilialEmpresa=giFilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_LE", X)
    If lErro <> SUCESSO Then Error 45768
    
    
    Reserva_Le = X.Reserva_Le(objReserva)
    
    Exit Function
    
Erro_Reserva_Le:

    Reserva_Le = Err
    
    Select Case Err

        Case 45768
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145582)
        
    End Select
    
    Exit Function
    
End Function

Function Bloqueio_ReservaGrava(objReserva As ClassReserva) As Long
'Grava o bloqueio com os dados passados em objReserva

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Bloqueio_ReservaGrava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIO_RESERVAGRAVA", X)
    If lErro <> SUCESSO Then Error 45770
    
    
    Bloqueio_ReservaGrava = X.Bloqueio_ReservaGrava(objReserva)
    
    Exit Function
    
Erro_Bloqueio_ReservaGrava:

    Bloqueio_ReservaGrava = Err
    
    Select Case Err

        Case 45770
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145583)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Exclui(objReserva As ClassReserva) As Long
'Exclui do BD a Reserva passada em objReserva

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 45771
    
    
    Reserva_Exclui = X.Reserva_Exclui(objReserva)
    
    Exit Function
    
Erro_Reserva_Exclui:

    Reserva_Exclui = Err
    
    Select Case Err

        Case 45771
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145584)
        
    End Select
    
    Exit Function
    
End Function

Function Bloqueio_ReservaExclui(objReserva As ClassReserva) As Long
'CHAMADA DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Bloqueio_ReservaExclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIO_RESERVAEXCLUI", X)
    If lErro <> SUCESSO Then Error 45772
    
    
    Bloqueio_ReservaExclui = X.Bloqueio_ReservaExclui(objReserva)
    
    Exit Function
    
Erro_Bloqueio_ReservaExclui:

    Bloqueio_ReservaExclui = Err
    
    Select Case Err

        Case 45772
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145585)
        
    End Select
    
    Exit Function
    
End Function

Function BloqueioPV_Grava(objBloqueioPV As ClassBloqueioPV) As Long
'Grava o bloqueio do pedido de venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioPV_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_BLOQUEIOPV_GRAVA", X)
    If lErro <> SUCESSO Then Error 45775
    
    
    BloqueioPV_Grava = X.BloqueioPV_Grava(objBloqueioPV)
    
    Exit Function
    
Erro_BloqueioPV_Grava:

    BloqueioPV_Grava = Err
    
    Select Case Err

        Case 45775
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145586)
        
    End Select
    
    Exit Function
    
End Function


Function ItemPedido_Le_Produto(objItemPedido As ClassItemPedido) As Long
'Lê no Banco de Dados o Item correspondente ao produto e CodPedido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedido_Le_Produto
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMPEDIDO_LE_PRODUTO", X)
    If lErro <> SUCESSO Then Error 45777
    
    
    ItemPedido_Le_Produto = X.ItemPedido_Le_Produto(objItemPedido)
    
    Exit Function
    
Erro_ItemPedido_Le_Produto:

    ItemPedido_Le_Produto = Err
    
    Select Case Err

        Case 45777
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145587)
        
    End Select
    
    Exit Function
    
End Function

Function Almoxarifados_Le_Cod_NomeRed(colCodNome As AdmColCodigoNome) As Long
'Lê o código e nome reduzido dos almoxarifados e retorna na coleção

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifados_Le_Cod_NomeRed
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ALMOXARIFADOS_LE_COD_NOMERED", X)
    If lErro <> SUCESSO Then Error 45778
    
    
    Almoxarifados_Le_Cod_NomeRed = X.Almoxarifados_Le_Cod_NomeRed(colCodNome)
    
    Exit Function
    
Erro_Almoxarifados_Le_Cod_NomeRed:

    Almoxarifados_Le_Cod_NomeRed = Err
    
    Select Case Err

        Case 45778
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145588)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Exclui_TipoPedidoCod(objReserva As ClassReserva, lComando As Long, lComando1 As Long, lComando2 As Long, lComando3 As Long) As Long
'Exclui o Tipo de Pedido associado à Reserva

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Exclui_TipoPedidoCod
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_EXCLUI_TIPOPEDIDOCOD", X)
    If lErro <> SUCESSO Then Error 45780
    
    
    Reserva_Exclui_TipoPedidoCod = X.Reserva_Exclui_TipoPedidoCod(objReserva, lComando, lComando1, lComando2, lComando3)
    
    Exit Function
    
Erro_Reserva_Exclui_TipoPedidoCod:

    Reserva_Exclui_TipoPedidoCod = Err
    
    Select Case Err

        Case 45780
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145589)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Grava_TipoPedidoCod(objReserva As ClassReserva, lComando1 As Long, lComando2 As Long) As Long
'Faz as atualizações à Nível de Item de Pedido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Grava_TipoPedidoCod
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_GRAVA_TIPOPEDIDOCOD", X)
    If lErro <> SUCESSO Then Error 45781
    
    
    Reserva_Grava_TipoPedidoCod = X.Reserva_Grava_TipoPedidoCod(objReserva, lComando1, lComando2)
    
    Exit Function
    
Erro_Reserva_Grava_TipoPedidoCod:

    Reserva_Grava_TipoPedidoCod = Err
    
    Select Case Err

        Case 45781
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145590)
        
    End Select
    
    Exit Function
    
End Function

Function Reserva_Grava_MovtoEstoque(objReserva As ClassReserva, dDifQtdeReserva As Double) As Long
'chamar movestoque_grava pela diferenca entre o que estava no bd
'e a nova qtde de reserva, se diferentes

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Grava_MovtoEstoque
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_RESERVA_GRAVA_MOVTOESTOQUE", X)
    If lErro <> SUCESSO Then Error 45782
    
    
    Reserva_Grava_MovtoEstoque = X.Reserva_Grava_MovtoEstoque(objReserva, dDifQtdeReserva)
    
    Exit Function
    
Erro_Reserva_Grava_MovtoEstoque:

    Reserva_Grava_MovtoEstoque = Err
    
    Select Case Err

        Case 45782
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145591)
        
    End Select
    
    Exit Function
    
End Function

Function GeracaoOP_SelecionarItens(objGeracaoOP As ClassGeracaoOP) As Long
'pesquisa no bd os itens candidatos preenchendo objGeracaoOP.colItens a partir das outras informacoes em

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoOP_SelecionarItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_GERACAOOP_SELECIONARITENS", X)
    If lErro <> SUCESSO Then Error 45796
    
    
    GeracaoOP_SelecionarItens = X.GeracaoOP_SelecionarItens(objGeracaoOP)
    
    Exit Function
    
Erro_GeracaoOP_SelecionarItens:

    GeracaoOP_SelecionarItens = Err
    
    Select Case Err

        Case 45796
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145592)
        
    End Select
    
    Exit Function
    
End Function

Function ESTConfig_Grava_Configs_Trans(colConfigs As ColESTConfig) As Long
'Grava na tabela ESTConfig uma coleção de registros de configuração
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ESTConfig_Grava_Configs_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_CONFIG_GRAVA_CONFIGS", X)
    If lErro <> SUCESSO Then Error 41818
    
    
    ESTConfig_Grava_Configs_Trans = X.ESTConfig_Grava_Configs_Trans(colConfigs)

    Exit Function
    
Erro_ESTConfig_Grava_Configs_Trans:
    
    ESTConfig_Grava_Configs_Trans = Err
    
    Select Case Err

        Case 41818
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145593)
        
    End Select
    
    Exit Function

End Function

Function EST_Instalacao_Filial(iFilialEmpresa As Integer, sIntervaloProducao As String) As Long
'faz as inicializacoes necessarias à criacao de uma nova filial especificas do modulo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_EST_Instalacao_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_INSTALACAO_FILIAL", X)
    If lErro <> SUCESSO Then Error 41936
    
    
    EST_Instalacao_Filial = X.EST_Instalacao_Filial(iFilialEmpresa, sIntervaloProducao)

    Exit Function
    
Erro_EST_Instalacao_Filial:
    
    EST_Instalacao_Filial = Err
    
    Select Case Err

        Case 41936
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145594)
        
    End Select
    
    Exit Function

End Function


Function NFiscalEntrada_Lock_Gravacao(objNFiscal As ClassNFiscal) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTRADA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then Error 19690
    
    
    NFiscalEntrada_Lock_Gravacao = X.NFiscalEntrada_Lock_Gravacao(objNFiscal)

    Exit Function
    
Erro_NFiscalEntrada_Lock_Gravacao:
    
    NFiscalEntrada_Lock_Gravacao = Err
    
    Select Case Err

        Case 19690
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145595)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTRADA_GRAVA", X)
    If lErro <> SUCESSO Then Error 19685
    
    
    NFiscalEntrada_Grava = X.NFiscalEntrada_Grava(objNFiscal, objContabil)

    Exit Function
    
Erro_NFiscalEntrada_Grava:
    
    NFiscalEntrada_Grava = Err
    
    Select Case Err

        Case 19685
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145596)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTRADA_GRAVA_BD", X)
    If lErro <> SUCESSO Then Error 19687
    
    
    NFiscalEntrada_Grava_BD = X.NFiscalEntrada_Grava_BD(objNFiscal)

    Exit Function
    
Erro_NFiscalEntrada_Grava_BD:
    
    NFiscalEntrada_Grava_BD = Err
    
    Select Case Err

        Case 19687
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145597)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Exclui(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Exclui a Nota Fiscal do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTRADA_EXCLUI", X)
    If lErro <> SUCESSO Then Error 19686
    
    
    NFiscalEntrada_Exclui = X.NFiscalEntrada_Exclui(objNFiscal, objContabil)

    Exit Function
    
Erro_NFiscalEntrada_Exclui:
    
    NFiscalEntrada_Exclui = Err
    
    Select Case Err

        Case 19686
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145598)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Baixa_Recebimento(objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Baixa_Recebimento
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTRADA_BAIXA_RECEBIMENTO", X)
    If lErro <> SUCESSO Then Error 19688
    
    
    NFiscalEntrada_Baixa_Recebimento = X.NFiscalEntrada_Baixa_Recebimento(objNFiscal)

    Exit Function
    
Erro_NFiscalEntrada_Baixa_Recebimento:
    
    NFiscalEntrada_Baixa_Recebimento = Err
    
    Select Case Err

        Case 19688
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145599)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntNum_Automatico(sSerie As String, lNumAuto As Long) As Long
'Retorna o número do próximo NumNFiscal disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntNum_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALENTNUM_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 19689
    
    
    NFiscalEntNum_Automatico = X.NFiscalEntNum_Automatico(sSerie, lNumAuto)

    Exit Function
    
Erro_NFiscalEntNum_Automatico:
    
    NFiscalEntNum_Automatico = Err
    
    Select Case Err

        Case 19689
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145600)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_NFiscal(objMovEstoque As ClassMovEstoque, sTela As String, objNFiscal As ClassNFiscal) As Long
'Esta função retorna o nome da tela, os dados da Nota Fiscal
'relacionada com o Movimento Estoque passado no parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_NomeTela_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("EST_MOVESTOQUE_LE_NOMETELA_NFISCAL", X)
    If lErro <> SUCESSO Then Error 55404
    
    MovEstoque_Le_NomeTela_NFiscal = X.MovEstoque_Le_NomeTela_NFiscal(objMovEstoque, sTela, objNFiscal)

    Exit Function
    
Erro_MovEstoque_Le_NomeTela_NFiscal:
    
    MovEstoque_Le_NomeTela_NFiscal = Err
    
    Select Case Err

        Case 55404
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145601)
        
    End Select
    
    Exit Function

End Function

Public Function MovEstoque_Le_NomeTela_ItemNFiscal(objMovEstoque As ClassMovEstoque, sTela As String, objNFiscal As ClassNFiscal) As Long
'Esta função retorna o nome da tela, os dados da Nota Fiscal dos Itens relacionado
'relacionada com o Movimento Estoque passado no parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_NomeTela_ItemNFiscal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_MOVESTOQUE_LE_NOMETELA_ITEMNFISCAL", X)
    If lErro <> SUCESSO Then Error 55405
    
    
    MovEstoque_Le_NomeTela_ItemNFiscal = X.MovEstoque_Le_NomeTela_ItemNFiscal(objMovEstoque, sTela, objNFiscal)

    Exit Function
    
Erro_MovEstoque_Le_NomeTela_ItemNFiscal:
    
    MovEstoque_Le_NomeTela_ItemNFiscal = Err
    
    Select Case Err

        Case 55405
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145602)
        
    End Select
    
    Exit Function

End Function

Function OrdemDeProducao_Le_SemItens(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao sem carregar os seus itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_Le_SemItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ORDEMDEPRODUCAO_LE_SEMITENS", X)
    If lErro <> SUCESSO Then Error 54737
    
    
    OrdemDeProducao_Le_SemItens = X.OrdemDeProducao_Le_SemItens(objOrdemDeProducao)

    Exit Function
    
Erro_OrdemDeProducao_Le_SemItens:
    
    OrdemDeProducao_Le_SemItens = Err
    
    Select Case Err

        Case 54737
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145603)
        
    End Select
    
    Exit Function

End Function

Function OPBaixada_Le_SemItens(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'carrega do bd a ordem de producao baixada sem carregar os seus itens

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OPBaixada_Le_SemItens
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_OPBAIXADA_LE_SEMITENS", X)
    If lErro <> SUCESSO Then Error 54738
    
    
    OPBaixada_Le_SemItens = X.OPBaixada_Le_SemItens(objOrdemDeProducao)

    Exit Function
    
Erro_OPBaixada_Le_SemItens:
    
    OPBaixada_Le_SemItens = Err
    
    Select Case Err

        Case 54738
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145604)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Grava_MovimentosEst(objNFiscal As ClassNFiscal, iTipoMovtoEstoque As Integer, sDocOrigem As String) As Long
'Gera um Movimento de Estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Grava_MovimentosEst
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFISCALENTRADA_GRAVA_MOVIMENTOSEST", X)
    If lErro <> SUCESSO Then Error 54500
    
    
    NFiscalEntrada_Grava_MovimentosEst = X.NFiscalEntrada_Grava_MovimentosEst(objNFiscal, iTipoMovtoEstoque, sDocOrigem)
    
    Exit Function
    
Erro_NFiscalEntrada_Grava_MovimentosEst:

    NFiscalEntrada_Grava_MovimentosEst = Err
    
    Select Case Err

        Case 54500
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145605)
        
    End Select
    
    Exit Function
    
End Function

Function ProdutosKit_Le_PrimeiroNivel_Est(objKit As ClassKit) As Long
'Preenche objKit.colComponentes a partir dos produtos inventariados e da versão passados em objKit e Nivel = KIT_NIVEL_RAIZ + 1

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutosKit_Le_PrimeiroNivel_Est
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_PRODUTOKIT_LE_PRIMEIRONIVEL_EST", X)
    If lErro <> SUCESSO Then Error 55479
    
    
    ProdutosKit_Le_PrimeiroNivel_Est = X.ProdutosKit_Le_PrimeiroNivel_Est(objKit)
    
    Exit Function
    
Erro_ProdutosKit_Le_PrimeiroNivel_Est:

    ProdutosKit_Le_PrimeiroNivel_Est = Err
    
    Select Case Err

        Case 55479
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145606)
        
    End Select
    
    Exit Function
    
End Function

Function OrdemDeProducao_TestaExistencia(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Pesquisa no bd a ordem de producao a partir de FilialEmpresa, Codigo e Data de Emissão

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OrdemDeProducao_TestaExistencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_ORDEMDEPRODUCAO_TESTAEXISTENCIA", X)
    If lErro <> SUCESSO Then Error 59526
    
    
    OrdemDeProducao_TestaExistencia = X.OrdemDeProducao_TestaExistencia(objOrdemDeProducao)
    
    Exit Function
    
Erro_OrdemDeProducao_TestaExistencia:

    OrdemDeProducao_TestaExistencia = Err
    
    Select Case Err

        Case 59526
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145607)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalRec_Interna_TestaExistencia(objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalRec_Interna_TestaExistencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALREC_INTERNA_TESTAEXISTENCIA", X)
    If lErro <> SUCESSO Then Error 59527
    
    
    NFiscalRec_Interna_TestaExistencia = X.NFiscalRec_Interna_TestaExistencia(objNFiscal)
    
    Exit Function
    
Erro_NFiscalRec_Interna_TestaExistencia:

    NFiscalRec_Interna_TestaExistencia = Err
    
    Select Case Err

        Case 59527
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145608)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Interna_TestaExistencia(objNFiscal As ClassNFiscal) As Long
'Verifica se foi cadastrada uma Nota Fiscal no BD a partir de FilialEmpresa, Serie, Número, Tipo, Data Emissão

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Interna_TestaExistencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCAL_INTERNA_TESTAEXISTENCIA", X)
    If lErro <> SUCESSO Then Error 59528
    
    
    NFiscal_Interna_TestaExistencia = X.NFiscal_Interna_TestaExistencia(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Interna_TestaExistencia:

    NFiscal_Interna_TestaExistencia = Err
    
    Select Case Err

        Case 59528
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145609)
        
    End Select
    
    Exit Function
    
End Function

Function Recebimento_Testa_Baixado(objNFiscal As ClassNFiscal) As Long
'Testa se o Recebimento passado está Baixado se estiver dá mensagem

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Recebimento_Testa_Baixado
       
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTO_TESTA_BAIXADO", X)
    If lErro <> SUCESSO Then Error 59867
    
    Recebimento_Testa_Baixado = X.Recebimento_Testa_Baixado(objNFiscal)
    
    Exit Function
    
Erro_Recebimento_Testa_Baixado:

    Recebimento_Testa_Baixado = Err
    
    Select Case Err

        Case 59867
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145610)
        
    End Select
    
    Exit Function
    
End Function

Function Recebimento_Automatico(objNFiscal As ClassNFiscal) As Long
'Preenche o Código do Recebimento de objNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Recebimento_Automatico
       
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTO_AUTOMATICO", X)
    If lErro <> SUCESSO Then Error 59868
    
    Recebimento_Automatico = X.Recebimento_Automatico(objNFiscal)
    
    Exit Function
    
Erro_Recebimento_Automatico:

    Recebimento_Automatico = Err
    
    Select Case Err

        Case 59868
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145611)
        
    End Select
    
    Exit Function
    
End Function

Function RecebimentoF_Verifica_Existencia(objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebimentoF_Verifica_Existencia
       
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTOF_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then Error 59869
    
    RecebimentoF_Verifica_Existencia = X.RecebimentoF_Verifica_Existencia(objNFiscal)
    
    Exit Function
    
Erro_RecebimentoF_Verifica_Existencia:

    RecebimentoF_Verifica_Existencia = Err
    
    Select Case Err

        Case 59869
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145612)
        
    End Select
    
    Exit Function
    
End Function

Function RecebimentoC_Verifica_Existencia(objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebimentoC_Verifica_Existencia
       
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTOC_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then Error 59870
    
    RecebimentoC_Verifica_Existencia = X.RecebimentoC_Verifica_Existencia(objNFiscal)
    
    Exit Function
    
Erro_RecebimentoC_Verifica_Existencia:

    RecebimentoC_Verifica_Existencia = Err
    
    Select Case Err

        Case 59870
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145613)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_Recebimento(objNFiscal As ClassNFiscal) As Long
'Le a Nota de Recebimento partir do Código do Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Recebimento
       
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCAL_LE_RECEBIMENTO", X)
    If lErro <> SUCESSO Then Error 65088
    
    NFiscal_Le_Recebimento = X.NFiscal_Le_Recebimento(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Le_Recebimento:

    NFiscal_Le_Recebimento = Err
    
    Select Case Err

        Case 65088
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145614)
        
    End Select
    
    Exit Function
    
End Function

Function NotaFiscalEntrada_Cancelar(objNFiscal As ClassNFiscal) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalEntrada_Cancelar
       
    lErro = Rotina_ObterObjetoFuncao("EST_NOTAFISCALENTRADA_CANCELAR", X)
    If lErro <> SUCESSO Then Error 65192
    
    NotaFiscalEntrada_Cancelar = X.NotaFiscalEntrada_Cancelar(objNFiscal)
    
    Exit Function
    
Erro_NotaFiscalEntrada_Cancelar:

    NotaFiscalEntrada_Cancelar = Err
    
    Select Case Err

        Case 65192
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145615)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalInternaEntrada_Le_Numero(objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo]
'Busca pela Serie e Número passados
'Filtros : TiposDocInfo.Tipo=DOCINFO_NF_INT_ENTRADA
'          NFiscal.Status<>STATUS_EXCLUIDO
'Ordena por DataEmissao, DECENDING
'Lê todos os Dados da Tabela NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalInternaEntrada_Le_Numero
       
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALINTERNAENTRADA_LE_NUMERO", X)
    If lErro <> SUCESSO Then Error 65193
    
    NFiscalInternaEntrada_Le_Numero = X.NFiscalInternaEntrada_Le_Numero(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalInternaEntrada_Le_Numero:

    NFiscalInternaEntrada_Le_Numero = Err
    
    Select Case Err

        Case 65193
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145616)
        
    End Select
    
    Exit Function
    
End Function

Function Recebimento_Desfaz_Baixa(objNFiscal As ClassNFiscal) As Long
'Desfaz a baixa do Recebimento vinculado a nota fiscal que foi cancelada e passada no obj

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Recebimento_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTO_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then gError 71179

    Recebimento_Desfaz_Baixa = X.Recebimento_Desfaz_Baixa(objNFiscal)

    Exit Function
    
Erro_Recebimento_Desfaz_Baixa:
    
    Recebimento_Desfaz_Baixa = gErr
    
    Select Case gErr

        Case 71179
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145617)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntradaCancela_Contabilidade(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'Gera um lançamento extorno para a Nota Fiscal ou se ainda não foi contabilizado Exclui

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaCancela_Contabilidade
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADACANCELA_CONTABILIDADE", X)
    If lErro <> SUCESSO Then gError 71181

    NFiscalEntradaCancela_Contabilidade = X.NFiscalEntradaCancela_Contabilidade(objNFiscal, objTipoDocInfo)

    Exit Function
    
Erro_NFiscalEntradaCancela_Contabilidade:
    
    NFiscalEntradaCancela_Contabilidade = gErr
    
    Select Case gErr

        Case 71181
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145618)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntrada_Calcula_Custo(objItemNF As ClassItemNF) As Long
'calcula o custo do item da NF em função do credito do ICMS e do IPI

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Calcula_Custo
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADA_CALCULA_CUSTO", X)
    If lErro <> SUCESSO Then gError 71182

    NFiscalEntrada_Calcula_Custo = X.NFiscalEntrada_Calcula_Custo(objItemNF)

    Exit Function
    
Erro_NFiscalEntrada_Calcula_Custo:
    
    NFiscalEntrada_Calcula_Custo = gErr
    
    Select Case gErr

        Case 71182
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145619)
        
    End Select
    
    Exit Function

End Function


Function NFiscalEntrada_Verifica_Existencia2(objNFiscal As ClassNFiscal, objTipoDocInfo As ClassTipoDocInfo) As Long
'pesquisa a existencia da nota fiscal com os dados passados como parametro ou avisa de uma nota com uma data proxima
'OBS: Da aviso que vai gravar os dados contabeis

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntrada_Verifica_Existencia2
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADA_VERIFICA_EXISTENCIA2", X)
    If lErro <> SUCESSO Then gError 71185

    NFiscalEntrada_Verifica_Existencia2 = X.NFiscalEntrada_Verifica_Existencia2(objNFiscal, objTipoDocInfo)

    Exit Function
    
Erro_NFiscalEntrada_Verifica_Existencia2:
    
    NFiscalEntrada_Verifica_Existencia2 = gErr
    
    Select Case gErr

        Case 71185
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145620)
        
    End Select
    
    Exit Function

End Function

'*** VERSÃO 2

Function NotaFiscalEntradaCom_Cancelar(objNFiscal As ClassNFiscal) As Long
'Faz o cancelamento de uma nota fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalEntradaCom_Cancelar
    
    lErro = Rotina_ObterObjetoFuncao("EST_NOTAFISCALENTRADACOM_CANCELAR", X)
    If lErro <> SUCESSO Then gError 71186

    NotaFiscalEntradaCom_Cancelar = X.NotaFiscalEntradaCom_Cancelar(objNFiscal)

    Exit Function
    
Erro_NotaFiscalEntradaCom_Cancelar:
    
    NotaFiscalEntradaCom_Cancelar = gErr
    
    Select Case gErr

        Case 71186
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145621)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemPC_Exclui_RelacionadaItemNF(lNumIntItemNF As Long) As Long
'Exlui Registros de ItemNFItemPC associados ao ItemNF de número interno passado em lNumIntItemNF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemPC_Exclui_RelacionadaItemNF
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMPC_EXCLUI_RELACIONADAITEMN", X)
    If lErro <> SUCESSO Then gError 71187

    ItemNFItemPC_Exclui_RelacionadaItemNF = X.ItemNFItemPC_Exclui_RelacionadaItemNF(lNumIntItemNF)

    Exit Function
    
Erro_ItemNFItemPC_Exclui_RelacionadaItemNF:
    
    ItemNFItemPC_Exclui_RelacionadaItemNF = gErr
    
    Select Case gErr

        Case 71187
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145622)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Desfaz_Baixa(objNFiscal As ClassNFiscal, lNumIntItemNF As Long, objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Refaz os Pedidos de Compras já Baixados e os ItensPC com
'Status Recebido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRA_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then gError 71188

    PedidoCompra_Desfaz_Baixa = X.PedidoCompra_Desfaz_Baixa(objNFiscal, lNumIntItemNF, objFilFornFilEmp)

    Exit Function
    
Erro_PedidoCompra_Desfaz_Baixa:
    
    PedidoCompra_Desfaz_Baixa = gErr
    
    Select Case gErr

        Case 71188
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145623)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemRC_Exclui_RelacionadaItemNF(lNumIntItemNF As Long) As Long
'Exlui Registros de ItemNFItemRC associados ao ItemNF de número interno passado em lNumIntItemNF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemRC_Exclui_RelacionadaItemNF
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMRC_EXCLUI_RELACIONADAITEMN", X)
    If lErro <> SUCESSO Then gError 71189

    ItemNFItemRC_Exclui_RelacionadaItemNF = X.ItemNFItemRC_Exclui_RelacionadaItemNF(lNumIntItemNF)

    Exit Function
    
Erro_ItemNFItemRC_Exclui_RelacionadaItemNF:
    
    ItemNFItemRC_Exclui_RelacionadaItemNF = gErr
    
    Select Case gErr

        Case 71189
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145624)
        
    End Select
    
    Exit Function

End Function

Function RequisicaoCompra_Desfaz_Baixa(lNumIntNF As Long, lNumIntItemNF As Long) As Long
'Refaz os Requisições de Compras já Baixadas e os ItensRC com
'Status Recebido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RequisicaoCompra_Desfaz_Baixa
    
    lErro = Rotina_ObterObjetoFuncao("EST_REQUISICAOCOMPRA_DESFAZ_BAIXA", X)
    If lErro <> SUCESSO Then gError 71190

    RequisicaoCompra_Desfaz_Baixa = X.RequisicaoCompra_Desfaz_Baixa(lNumIntNF, lNumIntItemNF)

    Exit Function
    
Erro_RequisicaoCompra_Desfaz_Baixa:
    
    RequisicaoCompra_Desfaz_Baixa = gErr
    
    Select Case gErr

        Case 71190
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145625)
        
    End Select
    
    Exit Function

End Function

Function OPGeradora_Le_OPGerada(objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_OPGeradora_Le_OPGerada
    
    lErro = Rotina_ObterObjetoFuncao("EST_OPGERADORA_LE_OPGERADA", X)
    If lErro <> SUCESSO Then gError 71191

    OPGeradora_Le_OPGerada = X.OPGeradora_Le_OPGerada(objOrdemDeProducao)

    Exit Function
    
Erro_OPGeradora_Le_OPGerada:
    
    OPGeradora_Le_OPGerada = gErr
    
    Select Case gErr

        Case 71191
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145626)
        
    End Select
    
    Exit Function

End Function

Function ItensPC_Le_Codigo(objPedidoCompra As ClassPedidoCompras) As Long
'Lê os itens de um Pedido de Compra a partir do código e de FilialEmpresa do Pedido de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPC_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPC_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 71192

    ItensPC_Le_Codigo = X.ItensPC_Le_Codigo(objPedidoCompra)

    Exit Function
    
Erro_ItensPC_Le_Codigo:
    
    ItensPC_Le_Codigo = gErr
    
    Select Case gErr

        Case 71192
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145627)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Le_NFiscalReceb(objNFiscal As ClassNFiscal, colItemPedCompraInfo As Collection) As Long
'Lê os itens de Pedido de compras não baixados relacionados a um Recebimento e os devolve em colItemPedCompraInfo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Le_NFiscalReceb
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPEDCOMPRA_LE_NFISCALRECEB", X)
    If lErro <> SUCESSO Then gError 71193

    ItensPedCompra_Le_NFiscalReceb = X.ItensPedCompra_Le_NFiscalReceb(objNFiscal, colItemPedCompraInfo)

    Exit Function
    
Erro_ItensPedCompra_Le_NFiscalReceb:
    
    ItensPedCompra_Le_NFiscalReceb = gErr
    
    Select Case gErr

        Case 71193
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145628)
        
    End Select
    
    Exit Function

End Function

Function PedidosCompraEnv_Le_Recebimento(colPedidoCompra As Collection, lCodigoFornecedor As Long, iFilial As Integer, iFilialCompra As Integer) As Long
'Lê os Pedidos de Compra Enviados que possuem a Filial, o Fornecedor e a FilialCompra passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidosCompraEnv_Le_Recebimento
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOSCOMPRAENV_LE_RECEBIMENTO", X)
    If lErro <> SUCESSO Then gError 71194

    PedidosCompraEnv_Le_Recebimento = X.PedidosCompraEnv_Le_Recebimento(colPedidoCompra, lCodigoFornecedor, iFilial, iFilialCompra)

    Exit Function
    
Erro_PedidosCompraEnv_Le_Recebimento:
    
    PedidosCompraEnv_Le_Recebimento = gErr
    
    Select Case gErr

        Case 71194
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145629)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_Recebimento(objNFiscal As ClassNFiscal, colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra não baixados Relacionados aos itens da Nota fiscal passada por objNFical

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_Recebimento
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRA_LE_RECEBIMENTO", X)
    If lErro <> SUCESSO Then gError 71195

    PedidoCompra_Le_Recebimento = X.PedidoCompra_Le_Recebimento(objNFiscal, colPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Le_Recebimento:
    
    PedidoCompra_Le_Recebimento = gErr
    
    Select Case gErr

        Case 71195
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145630)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompras_Le_Valores(colPedidoCompras As Collection) As Long
'Lê valores dos Pedido de compras passados na coleção colPedidoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompras_Le_Valores
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRAS_LE_VALORES", X)
    If lErro <> SUCESSO Then gError 71196

    PedidoCompras_Le_Valores = X.PedidoCompras_Le_Valores(colPedidoCompras)

    Exit Function
    
Erro_PedidoCompras_Le_Valores:
    
    PedidoCompras_Le_Valores = gErr
    
    Select Case gErr

        Case 71196
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145631)
        
    End Select
    
    Exit Function

End Function

Function RecebMaterialFCom_Grava(objNFiscal As ClassNFiscal, colCodPedCompras As Collection) As Long
'Grava a Nota Fiscal de Recebimento de Material de Fornecedor de compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialFCom_Grava
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALFCOM_GRAVA", X)
    If lErro <> SUCESSO Then gError 71197

    RecebMaterialFCom_Grava = X.RecebMaterialFCom_Grava(objNFiscal, colCodPedCompras)

    Exit Function
    
Erro_RecebMaterialFCom_Grava:
    
    RecebMaterialFCom_Grava = gErr
    
    Select Case gErr

        Case 71197
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145632)
        
    End Select
    
    Exit Function

End Function

Function RecebMaterialFCom_Exclui(objNFiscal As ClassNFiscal) As Long
'Exclui o Recebimento de Material de Fornecedor de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebMaterialFCom_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBMATERIALFCOM_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71198

    RecebMaterialFCom_Exclui = X.RecebMaterialFCom_Exclui(objNFiscal)

    Exit Function
    
Erro_RecebMaterialFCom_Exclui:
    
    RecebMaterialFCom_Exclui = gErr
    
    Select Case gErr

        Case 71198
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145633)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemPC_Exclui(lNumIntItemNFiscal As Long) As Long
'Exclui os Registros da tabela ItemNFItemPC relacionados a itens
'de Nota fiscal passados por lNumIntItemNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemPC_Exclui
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMPC_EXCLUI", X)
    If lErro <> SUCESSO Then gError 71199

    ItemNFItemPC_Exclui = X.ItemNFItemPC_Exclui(lNumIntItemNFiscal)

    Exit Function
    
Erro_ItemNFItemPC_Exclui:
    
    ItemNFItemPC_Exclui = gErr
    
    Select Case gErr

        Case 71199
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145634)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemPC_IncluiBD(objItemNF As ClassItemNF) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemPC_IncluiBD
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMPC_INCLUIBD", X)
    If lErro <> SUCESSO Then gError 71200

    ItemNFItemPC_IncluiBD = X.ItemNFItemPC_IncluiBD(objItemNF)

    Exit Function
    
Erro_ItemNFItemPC_IncluiBD:
    
    ItemNFItemPC_IncluiBD = gErr
    
    Select Case gErr

        Case 71200
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145635)
        
    End Select
    
    Exit Function

End Function

Function RecebimentoFCom_Verifica_Existencia(objNFiscal As ClassNFiscal) As Long
'Testa se existe outro recebimento com os memos dados de Número de Nota

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RecebimentoFCom_Verifica_Existencia
    
    lErro = Rotina_ObterObjetoFuncao("EST_RECEBIMENTOFCOM_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then gError 71201

    RecebimentoFCom_Verifica_Existencia = X.RecebimentoFCom_Verifica_Existencia(objNFiscal)

    Exit Function
    
Erro_RecebimentoFCom_Verifica_Existencia:
    
    RecebimentoFCom_Verifica_Existencia = gErr
    
    Select Case gErr

        Case 71201
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145636)
        
    End Select
    
    Exit Function

End Function

Function ItensPedidosCompra_VerificaQuantidade(colPedidos As Collection) As Long
'Remove da coleção passada os Pedidos de compras que não possuem pelo menos um
'ItemPC com quantidade à Receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedidosCompra_VerificaQuantidade
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPEDIDOSCOMPRA_VERIFICAQUANTIDAD", X)
    If lErro <> SUCESSO Then gError 71202

    ItensPedidosCompra_VerificaQuantidade = X.ItensPedidosCompra_VerificaQuantidade(colPedidos)

    Exit Function
    
Erro_ItensPedidosCompra_VerificaQuantidade:
    
    ItensPedidosCompra_VerificaQuantidade = gErr
    
    Select Case gErr

        Case 71202
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145637)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntradaCom_Lock_Gravacao(objNFiscal As ClassNFiscal, colPedCompras As Collection, colReqCompras As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaCom_Lock_Gravacao
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADACOM_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 71203

    NFiscalEntradaCom_Lock_Gravacao = X.NFiscalEntradaCom_Lock_Gravacao(objNFiscal, colPedCompras, colReqCompras)

    Exit Function
    
Erro_NFiscalEntradaCom_Lock_Gravacao:
    
    NFiscalEntradaCom_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 71203
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145638)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntradaCom_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil, colPedCompras As Collection, colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaCom_Grava
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADACOM_GRAVA", X)
    If lErro <> SUCESSO Then gError 71204

    NFiscalEntradaCom_Grava = X.NFiscalEntradaCom_Grava(objNFiscal, objContabil, colPedCompras, colReqCompras)

    Exit Function
    
Erro_NFiscalEntradaCom_Grava:
    
    NFiscalEntradaCom_Grava = gErr
    
    Select Case gErr

        Case 71204
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145639)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntradaCom_Grava_BD(objNFiscal As ClassNFiscal, colPedCompras As Collection, colReqCompras As Collection) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaCom_Grava_BD
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADACOM_GRAVA_BD", X)
    If lErro <> SUCESSO Then gError 71205

    NFiscalEntradaCom_Grava_BD = X.NFiscalEntradaCom_Grava_BD(objNFiscal, colPedCompras, colReqCompras)

    Exit Function
    
Erro_NFiscalEntradaCom_Grava_BD:
    
    NFiscalEntradaCom_Grava_BD = gErr
    
    Select Case gErr

        Case 71205
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145640)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemPC_Inclui(objNFiscal As ClassNFiscal) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemPC_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMPC_INCLUI", X)
    If lErro <> SUCESSO Then gError 71205

    ItemNFItemPC_Inclui = X.ItemNFItemPC_Inclui(objNFiscal)

    Exit Function
    
Erro_ItemNFItemPC_Inclui:
    
    ItemNFItemPC_Inclui = gErr
    
    Select Case gErr

        Case 71205
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145641)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Atualiza_Quantidades(objNFiscal As ClassNFiscal, colPedCompras As Collection) As Long
'Atualiza a Quantidade Recebida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Atualiza_Quantidades
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPEDCOMPRA_ATUALIZA_QUANTIDADES", X)
    If lErro <> SUCESSO Then gError 71206

    ItensPedCompra_Atualiza_Quantidades = X.ItensPedCompra_Atualiza_Quantidades(objNFiscal, colPedCompras)

    Exit Function
    
Erro_ItensPedCompra_Atualiza_Quantidades:
    
    ItensPedCompra_Atualiza_Quantidades = gErr
    
    Select Case gErr

        Case 71206
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145642)
        
    End Select
    
    Exit Function

End Function

Function ItemNFItemRC_Inclui(objNFiscal As ClassNFiscal) As Long
'Insere um registro na tabela ItemNFItemPC para cada item da coleção colItemNFItemPC passada por objItemNF

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFItemRC_Inclui
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMRC_INCLUI", X)
    If lErro <> SUCESSO Then gError 71207

    ItemNFItemRC_Inclui = X.ItemNFItemRC_Inclui(objNFiscal)

    Exit Function
    
Erro_ItemNFItemRC_Inclui:
    
    ItemNFItemRC_Inclui = gErr
    
    Select Case gErr

        Case 71207
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145643)
        
    End Select
    
    Exit Function

End Function

Function ItensReqCompra_Atualiza_Quantidades(objNFiscal As ClassNFiscal) As Long
'Atualiza Quantidades dos Itens de Requisição de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensReqCompra_Atualiza_Quantidades
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMNFITEMRC_INCLUI", X)
    If lErro <> SUCESSO Then gError 71208

    ItensReqCompra_Atualiza_Quantidades = X.ItensReqCompra_Atualiza_Quantidades(objNFiscal)

    Exit Function
    
Erro_ItensReqCompra_Atualiza_Quantidades:
    
    ItensReqCompra_Atualiza_Quantidades = gErr
    
    Select Case gErr

        Case 71208
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145644)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntCom_BaixaPedidos(colPedCompras As Collection, objNFiscal As ClassNFiscal) As Long
'Baixa os Pedidos de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntCom_BaixaPedidos
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTCOM_BAIXAPEDIDOS", X)
    If lErro <> SUCESSO Then gError 71209

    NFiscalEntCom_BaixaPedidos = X.NFiscalEntCom_BaixaPedidos(colPedCompras, objNFiscal)

    Exit Function
    
Erro_NFiscalEntCom_BaixaPedidos:
    
    NFiscalEntCom_BaixaPedidos = gErr
    
    Select Case gErr

        Case 71209
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145645)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntCom_BaixaRequisicoes(colReqCompras As Collection, objNFiscal As ClassNFiscal) As Long
'Baixa as Requisições de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntCom_BaixaRequisicoes
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTCOM_BAIXAREQUISICOES", X)
    If lErro <> SUCESSO Then gError 71210

    NFiscalEntCom_BaixaRequisicoes = X.NFiscalEntCom_BaixaRequisicoes(colReqCompras, objNFiscal)

    Exit Function
    
Erro_NFiscalEntCom_BaixaRequisicoes:
    
    NFiscalEntCom_BaixaRequisicoes = gErr
    
    Select Case gErr

        Case 71210
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145646)
        
    End Select
    
    Exit Function

End Function

Function NFiscalEntradaCom_Baixa_Recebimento(objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalEntradaCom_Baixa_Recebimento
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALENTRADACOM_BAIXA_RECEBIMENTO", X)
    If lErro <> SUCESSO Then gError 71211

    NFiscalEntradaCom_Baixa_Recebimento = X.NFiscalEntradaCom_Baixa_Recebimento(objNFiscal)

    Exit Function
    
Erro_NFiscalEntradaCom_Baixa_Recebimento:
    
    NFiscalEntradaCom_Baixa_Recebimento = gErr
    
    Select Case gErr

        Case 71211
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145647)
        
    End Select
    
    Exit Function

End Function

Function ItensPCTodos_Le_Codigo(objPedidoCompra As ClassPedidoCompras) As Long
'Lê os itens de um Pedido de Compra (baixado ou não) a partir do código e de FilialEmpresa do Pedido de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPCTodos_Le_Codigo
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPCTODOS_LE_CODIGO", X)
    If lErro <> SUCESSO Then gError 71212

    ItensPCTodos_Le_Codigo = X.ItensPCTodos_Le_Codigo(objPedidoCompra)

    Exit Function
    
Erro_ItensPCTodos_Le_Codigo:
    
    ItensPCTodos_Le_Codigo = gErr
    
    Select Case gErr

        Case 71212
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145648)
        
    End Select
    
    Exit Function

End Function

Function ReqComprasTodas_Le_PedidoCompra(objPedidoCompras As ClassPedidoCompras, colReqCompras As Collection) As Long
'Lê as Requisições de compras (baixadas ou não) associadas ao Pedidos de Compra passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReqComprasTodas_Le_PedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("EST_REQCOMPRASTODAS_LE_PEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71213

    ReqComprasTodas_Le_PedidoCompra = X.ReqComprasTodas_Le_PedidoCompra(objPedidoCompras, colReqCompras)

    Exit Function
    
Erro_ReqComprasTodas_Le_PedidoCompra:
    
    ReqComprasTodas_Le_PedidoCompra = gErr
    
    Select Case gErr

        Case 71213
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145649)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_ReqComprasVinculada(objPedidoCompras As ClassPedidoCompras, colReqCompras As Collection) As Long
'Lê as Requisições de compras associadas ao Pedidos de Compra passado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_ReqComprasVinculada
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRA_LE_REQCOMPRASVINCULADA", X)
    If lErro <> SUCESSO Then gError 71214

    PedidoCompra_Le_ReqComprasVinculada = X.PedidoCompra_Le_ReqComprasVinculada(objPedidoCompras, colReqCompras)

    Exit Function
    
Erro_PedidoCompra_Le_ReqComprasVinculada:
    
    PedidoCompra_Le_ReqComprasVinculada = gErr
    
    Select Case gErr

        Case 71214
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145650)
        
    End Select
    
    Exit Function

End Function

Function ItensRequisicaoTodos_Le_PedidoCompra(objRequisicaoCompras As ClassRequisicaoCompras, colItemRCInfo As Collection, lNumIntDocNF As Long) As Long
'Lê todos os Itens de Requisição de Compras (baixados ou não) associadas ao Pedido de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensRequisicaoTodos_Le_PedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSREQUISICAOTODOS_LE_PEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71215

    ItensRequisicaoTodos_Le_PedidoCompra = X.ItensRequisicaoTodos_Le_PedidoCompra(objRequisicaoCompras, colItemRCInfo, lNumIntDocNF)

    Exit Function
    
Erro_ItensRequisicaoTodos_Le_PedidoCompra:
    
    ItensRequisicaoTodos_Le_PedidoCompra = gErr
    
    Select Case gErr

        Case 71215
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145651)
        
    End Select
    
    Exit Function

End Function

Function ItensRequisicao_Le_PedidoCompra_NF(objRequisicaoCompras As ClassRequisicaoCompras, colItemRCInfo As Collection, lNumIntNF As Long) As Long
'Lê os Itens de Requisição de Compras não baixados associadas ao Pedido de Compras associados a um item de nota fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensRequisicao_Le_PedidoCompra_NF
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSREQUISICAO_LE_PEDIDOCOMPRA_NF", X)
    If lErro <> SUCESSO Then gError 71216

    ItensRequisicao_Le_PedidoCompra_NF = X.ItensRequisicao_Le_PedidoCompra_NF(objRequisicaoCompras, colItemRCInfo, lNumIntNF)

    Exit Function
    
Erro_ItensRequisicao_Le_PedidoCompra_NF:
    
    ItensRequisicao_Le_PedidoCompra_NF = gErr
    
    Select Case gErr

        Case 71216
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145652)
        
    End Select
    
    Exit Function

End Function

Function ItensRequisicao_Le_PedidoCompra(objRequisicaoCompras As ClassRequisicaoCompras, colItemRCInfo As Collection) As Long
'Lê os Itens de Requisição de Compras não baixados associadas ao Pedido de Compras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensRequisicao_Le_PedidoCompra
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSREQUISICAO_LE_PEDIDOCOMPRA", X)
    If lErro <> SUCESSO Then gError 71217

    ItensRequisicao_Le_PedidoCompra = X.ItensRequisicao_Le_PedidoCompra(objRequisicaoCompras, colItemRCInfo)

    Exit Function
    
Erro_ItensRequisicao_Le_PedidoCompra:
    
    ItensRequisicao_Le_PedidoCompra = gErr
    
    Select Case gErr

        Case 71217
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145653)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompraTodos_Le_NFiscal(objNFiscal As ClassNFiscal, colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra Relacionados aos itens da Nota fiscal passada por objNFical

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompraTodos_Le_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRATODOS_LE_NFISCAL", X)
    If lErro <> SUCESSO Then gError 71218

    PedidoCompraTodos_Le_NFiscal = X.PedidoCompraTodos_Le_NFiscal(objNFiscal, colPedidoCompra)

    Exit Function
    
Erro_PedidoCompraTodos_Le_NFiscal:
    
    PedidoCompraTodos_Le_NFiscal = gErr
    
    Select Case gErr

        Case 71218
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145654)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_NFiscal(objNFiscal As ClassNFiscal, colPedidoCompra As Collection) As Long
'Lê os Pedidos de Compra não baixados Relacionados aos itens da Nota fiscal passada por objNFical

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRA_LE_NFISCAL", X)
    If lErro <> SUCESSO Then gError 71219

    PedidoCompra_Le_NFiscal = X.PedidoCompra_Le_NFiscal(objNFiscal, colPedidoCompra)

    Exit Function
    
Erro_PedidoCompra_Le_NFiscal:
    
    PedidoCompra_Le_NFiscal = gErr
    
    Select Case gErr

        Case 71219
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145655)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompraTodos_Le_NFiscal(objNFiscal As ClassNFiscal, colItemPedCompraInfo As Collection) As Long
'Lê todos os itens de Pedido de compras (baixados ou não) relacionados a nota fiscal e os devolve em colItemPedCompraInfo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompraTodos_Le_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPEDCOMPRATODOS_LE_NFISCAL", X)
    If lErro <> SUCESSO Then gError 71220

    ItensPedCompraTodos_Le_NFiscal = X.ItensPedCompraTodos_Le_NFiscal(objNFiscal, colItemPedCompraInfo)

    Exit Function
    
Erro_ItensPedCompraTodos_Le_NFiscal:
    
    ItensPedCompraTodos_Le_NFiscal = gErr
    
    Select Case gErr

        Case 71220
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145656)
        
    End Select
    
    Exit Function

End Function

Function ItensPedCompra_Le_NFiscal(objNFiscal As ClassNFiscal, colItemPedCompraInfo As Collection) As Long
'Lê os itens de Pedido de compras não baixados relacionados a nota fiscal e os devolve em colItemPedCompraInfo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedCompra_Le_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITENSPEDCOMPRA_LE_NFISCAL", X)
    If lErro <> SUCESSO Then gError 71221

    ItensPedCompra_Le_NFiscal = X.ItensPedCompra_Le_NFiscal(objNFiscal, colItemPedCompraInfo)

    Exit Function
    
Erro_ItensPedCompra_Le_NFiscal:
    
    ItensPedCompra_Le_NFiscal = gErr
    
    Select Case gErr

        Case 71221
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145657)
        
    End Select
    
    Exit Function

End Function

Function ItemPedidoCompra_Le_QuantSobra(objItemPedCompraInfo As ClassItemPedCompraInfo, iFilialEmpresa As Integer, dQuantidadeSobra As Double) As Long
'Lê a Soma das Quantidades de Itens para o Pedido de Compra desvinculado a Requisição de Compra
'iFilialEmpresa que originou o Pedido de Compra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemPedidoCompra_Le_QuantSobra
    
    lErro = Rotina_ObterObjetoFuncao("EST_ITEMPEDIDOCOMPRA_LE_QUANTSOBRA", X)
    If lErro <> SUCESSO Then gError 71222

    ItemPedidoCompra_Le_QuantSobra = X.ItemPedidoCompra_Le_QuantSobra(objItemPedCompraInfo, iFilialEmpresa, dQuantidadeSobra)

    Exit Function
    
Erro_ItemPedidoCompra_Le_QuantSobra:
    
    ItemPedidoCompra_Le_QuantSobra = gErr
    
    Select Case gErr

        Case 71222
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145658)
        
    End Select
    
    Exit Function

End Function

Function ReqCompras_Le_PedidoCompras(objPedidoCompra As ClassPedidoCompras, lReqCompra As Long, colPedCompras As Collection) As Long
'Lê os Pedidos de Compras que possuem associação com uma determinada Requisição passada em lReqCompra

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReqCompras_Le_PedidoCompras
    
    lErro = Rotina_ObterObjetoFuncao("EST_REQCOMPRAS_LE_PEDIDOCOMPRAS", X)
    If lErro <> SUCESSO Then gError 71223

    ReqCompras_Le_PedidoCompras = X.ReqCompras_Le_PedidoCompras(objPedidoCompra, lReqCompra, colPedCompras)

    Exit Function
    
Erro_ReqCompras_Le_PedidoCompras:
    
    ReqCompras_Le_PedidoCompras = gErr
    
    Select Case gErr

        Case 71223
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145659)
        
    End Select
    
    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil, colPedCompras As Collection, colReqCompras As Collection) As Long
'Faz a Gravação de Uma Nota Fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntradaCom_Grava
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADACOM_GRAVA", X)
    If lErro <> SUCESSO Then gError 71224

    NFiscalFatEntradaCom_Grava = X.NFiscalFatEntradaCom_Grava(objNFiscal, objContabil, colPedCompras, colReqCompras)

    Exit Function
    
Erro_NFiscalFatEntradaCom_Grava:
    
    NFiscalFatEntradaCom_Grava = gErr
    
    Select Case gErr

        Case 71224
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145660)
        
    End Select
    
    Exit Function

End Function

Function NFiscalFatEntradaCom_Grava_BD(objNFiscal As ClassNFiscal, colPedCompras As Collection, colReqCompras As Collection) As Long
'Insere uma Nova Nota Fiscal de Entrada no Banco de Dados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntradaCom_Grava_BD
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADACOM_GRAVA_BD", X)
    If lErro <> SUCESSO Then gError 71225

    NFiscalFatEntradaCom_Grava_BD = X.NFiscalFatEntradaCom_Grava_BD(objNFiscal, colPedCompras, colReqCompras)

    Exit Function
    
Erro_NFiscalFatEntradaCom_Grava_BD:
    
    NFiscalFatEntradaCom_Grava_BD = gErr
    
    Select Case gErr

        Case 71225
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145661)
        
    End Select
    
    Exit Function

End Function

Function NFiscalFatEntradaCom_Baixa_Recebimento(objNFiscal As ClassNFiscal) As Long
'Faz a baixa de um Recebimento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatEntradaCom_Baixa_Recebimento
    
    lErro = Rotina_ObterObjetoFuncao("EST_NFISCALFATENTRADACOM_BAIXA_RECEBIMEN", X)
    If lErro <> SUCESSO Then gError 71226

    NFiscalFatEntradaCom_Baixa_Recebimento = X.NFiscalFatEntradaCom_Baixa_Recebimento(objNFiscal)

    Exit Function
    
Erro_NFiscalFatEntradaCom_Baixa_Recebimento:
    
    NFiscalFatEntradaCom_Baixa_Recebimento = gErr
    
    Select Case gErr

        Case 71226
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145662)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Reprocessamento_MovEstoque(sNomeArqParam As String, ByVal iFilialEmpresa As Integer, dtDataInicio As Date, iOrdemReproc As Integer) As Long
'reprocessa os movimentos de estoque a partir da data passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Rotina_Reprocessamento_MovEstoque
    
    lErro = Rotina_ObterObjetoFuncao("EST_ROTINA_REPROCESSAMENTO_MOVESTOQUE", X)
    If lErro <> SUCESSO Then gError 83783

    Rotina_Reprocessamento_MovEstoque = X.Rotina_Reprocessamento_MovEstoque(sNomeArqParam, iFilialEmpresa, dtDataInicio, iOrdemReproc)

    Exit Function
    
Erro_Rotina_Reprocessamento_MovEstoque:
    
    Rotina_Reprocessamento_MovEstoque = gErr
    
    Select Case gErr

        Case 83783
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145663)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_Produto(objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o MovEstoque a partir do Código, Filial Empresa e Produto.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_Produto
    
    lErro = Rotina_ObterObjetoFuncao("EST_MOVESTOQUE_LE_PRODUTO", X)
    If lErro <> SUCESSO Then gError 88005

    MovEstoque_Le_Produto = X.MovEstoque_Le_Produto(objItemMovEstoque)

    Exit Function
    
Erro_MovEstoque_Le_Produto:
    
    MovEstoque_Le_Produto = gErr
    
    Select Case gErr

        Case 88005
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145664)
        
    End Select
    
    Exit Function

End Function

Function MovEstoque_Le_NumIntDoc(objItemMovEstoque As ClassItemMovEstoque) As Long
'Lê o MovEstoque a partir do NumIntDoc.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Le_NumIntDoc
    
    lErro = Rotina_ObterObjetoFuncao("EST_MOVESTOQUE_LE_NUMINTDOC", X)
    If lErro <> SUCESSO Then gError 88006

    MovEstoque_Le_NumIntDoc = X.MovEstoque_Le_NumIntDoc(objItemMovEstoque)

    Exit Function
    
Erro_MovEstoque_Le_NumIntDoc:
    
    MovEstoque_Le_NumIntDoc = gErr
    
    Select Case gErr

        Case 88006
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145665)
        
    End Select
    
    Exit Function

End Function

Function ApropriacaoInsumo_Le_NumIntDocOrigem(lNumIntDocOrigem As Long, colApropriacaoInsumo As Collection) As Long
'Lê a tabela de RastreamentoMovto através do Movimento de Estoque

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ApropriacaoInsumo_Le_NumIntDocOrigem
    
    lErro = Rotina_ObterObjetoFuncao("EST_APROPRIACAOINSUMO_LE_NUMINTDOCORIGEM", X)
    If lErro <> SUCESSO Then gError 88007

    ApropriacaoInsumo_Le_NumIntDocOrigem = X.ApropriacaoInsumo_Le_NumIntDocOrigem(lNumIntDocOrigem, colApropriacaoInsumo)

    Exit Function
    
Erro_ApropriacaoInsumo_Le_NumIntDocOrigem:
    
    ApropriacaoInsumo_Le_NumIntDocOrigem = gErr
    
    Select Case gErr

        Case 88007
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145666)
        
    End Select
    
    Exit Function

End Function

Function ReqProducao_TestaExistencia(sCodigoOP As String) As Long
'Pesquisa no bd a ordem de producao a partir de FilialEmpresa, Codigo e Data de Emissão

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReqProducao_TestaExistencia
    
    lErro = Rotina_ObterObjetoFuncao("EST_REQPRODUCAO_TESTAEXISTENCIA", X)
    If lErro <> SUCESSO Then gError 88008

    ReqProducao_TestaExistencia = X.ReqProducao_TestaExistencia(sCodigoOP)

    Exit Function
    
Erro_ReqProducao_TestaExistencia:
    
    ReqProducao_TestaExistencia = gErr
    
    Select Case gErr

        Case 88008
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145667)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscalEntrada_Excluir(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz a exclusao de uma nota fiscal de Entrada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalEntrada_Excluir
       
    lErro = Rotina_ObterObjetoFuncao("EST_NOTAFISCALENTRADA_EXCLUIR", X)
    If lErro <> SUCESSO Then gError 89054
    
    NotaFiscalEntrada_Excluir = X.NotaFiscalEntrada_Excluir(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NotaFiscalEntrada_Excluir:

    NotaFiscalEntrada_Excluir = gErr
    
    Select Case gErr

        Case 89054
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145668)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoCompras_Le_Valores1(colPedidoCompras As Collection) As Long
'Lê valores dos Pedido de compras passados na coleção colPedidoCompras

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompras_Le_Valores1
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRAS_LE_VALORES1", X)
    If lErro <> SUCESSO Then gError 89230

    PedidoCompras_Le_Valores1 = X.PedidoCompras_Le_Valores1(colPedidoCompras)

    Exit Function
    
Erro_PedidoCompras_Le_Valores1:
    
    PedidoCompras_Le_Valores1 = gErr
    
    Select Case gErr

        Case 89230
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145669)
        
    End Select
    
    Exit Function

End Function

Function PedidoCompra_Le_EnvComQuantReceber(colPedidos As Collection, lCodigoFornecedor As Long, iFilial As Integer, iFilialCompra As Integer) As Long
'Le os pedidos de compra enviados para o fornecedor em questão com quantidade a receber

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoCompra_Le_EnvComQuantReceber
    
    lErro = Rotina_ObterObjetoFuncao("EST_PEDIDOCOMPRAS_LE_ENVCOMQUANTRECEBER", X)
    If lErro <> SUCESSO Then gError 89273

    PedidoCompra_Le_EnvComQuantReceber = X.PedidoCompra_Le_EnvComQuantReceber(colPedidos, lCodigoFornecedor, iFilial, iFilialCompra)

    Exit Function
    
Erro_PedidoCompra_Le_EnvComQuantReceber:
    
    PedidoCompra_Le_EnvComQuantReceber = gErr
    
    Select Case gErr

        Case 89273
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145670)
        
    End Select
    
    Exit Function

End Function

Public Function Aglutina_Pedidos(colPedidoCompra As Collection, colPedidos As Collection, colPedidosRec As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Aglutina_Pedidos
    
    lErro = Rotina_ObterObjetoFuncao("EST_AGLUTINA_PEDIDOS", X)
    If lErro <> SUCESSO Then gError 89274

    Aglutina_Pedidos = X.Aglutina_Pedidos(colPedidoCompra, colPedidos, colPedidosRec)

    Exit Function
    
Erro_Aglutina_Pedidos:
    
    Aglutina_Pedidos = gErr
    
    Select Case gErr

        Case 89274
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145671)
        
    End Select
    
    Exit Function

End Function

