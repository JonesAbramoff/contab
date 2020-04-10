VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChamaFuncoesFAT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function FATConfig_Le_Configs(colConfigs As ColFATConfig) As Long
'recupera da tabela FATConfig uma colecao de registros
Dim X As Object, lErro As Long
    
On Error GoTo Erro_FATConfig_Le_Configs
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CONFIG_LE_CONFIGS", X)
    If lErro <> SUCESSO Then gError 19326
    
    
    FATConfig_Le_Configs = X.FATConfig_Le_Configs(colConfigs)

    Exit Function
    
Erro_FATConfig_Le_Configs:
    
    FATConfig_Le_Configs = gErr
    
    Select Case gErr

        Case 19326
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145672)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Grava_Comissoes(objNFiscal As ClassNFiscal) As Long
'Atualiza os dados de comissões de NFiscal, com as novas comissões passadas em objNFiscal

Dim X As Object, lErro As Long

On Error GoTo Erro_NFiscal_Grava_Comissoes

    
    

    lErro = Rotina_ObterObjetoFuncao("CRFAT_NFSREC_GRAVA_COMISSOES", X)
    If lErro <> SUCESSO Then gError 24565

    
    NFiscal_Grava_Comissoes = X.NFiscal_Grava_Comissoes(objNFiscal)

    Exit Function

Erro_NFiscal_Grava_Comissoes:

    NFiscal_Grava_Comissoes = gErr

    Select Case gErr

        Case 24565
            

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145673)

    End Select

    Exit Function

End Function

Function LiberacaoDeBloqueios_ObterBloqueios(objLiberacaoBloqueios As ClassLiberacaoBloqueios, colTiposBloqueio As Collection) As Long
'Preenche colBloqueioLiberacaoInfo em objLiberacaoBloqueios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoDeBloqueios_ObterBloqueios
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAODEBLOQUEIOS_OBTERBLOQUEIOS", X)
    If lErro <> SUCESSO Then gError 24362
    
    
    LiberacaoDeBloqueios_ObterBloqueios = X.LiberacaoDeBloqueios_ObterBloqueios(objLiberacaoBloqueios, colTiposBloqueio)

    Exit Function
    
Erro_LiberacaoDeBloqueios_ObterBloqueios:
    
    LiberacaoDeBloqueios_ObterBloqueios = gErr
    
    Select Case gErr

        Case 24362
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145674)
        
    End Select
    
    Exit Function

End Function

Function GeracaoFatura_ObterNFs(objGeracaoFatura As ClassGeracaoFatura) As Long
'preenche a colecao em objGeracaoFatura com as NFsRec que podem ser faturadas e atendem
'aos criterios passados em objGeracaoFatura com restricoes de No. NF, data de emissao,...

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoFatura_ObterNFs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_GERACAOFATURA_OBTERNFS", X)
    If lErro <> SUCESSO Then gError 24416
    
    
    GeracaoFatura_ObterNFs = X.GeracaoFatura_ObterNFs(objGeracaoFatura)

    Exit Function
    
Erro_GeracaoFatura_ObterNFs:
    
    GeracaoFatura_ObterNFs = gErr
    
    Select Case gErr

        Case 24416
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145675)
        
    End Select
    
    Exit Function

End Function

Function GeracaoNFiscal_ObterPedidos(objGeracaoNFiscal As ClassGeracaoNFiscal) As Long
'preenche a colecao em objGeracaoNFiscal com os pedidos que podem ser faturados e atendem
'aos criterios passados em objGeracaoNFiscal com restricoes de No. de pedido, data de emissao,...

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoNFiscal_ObterPedidos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_GERACAONFISCAL_OBTERPEDIDOS", X)
    If lErro <> SUCESSO Then gError 24473
    
    
    GeracaoNFiscal_ObterPedidos = X.GeracaoNFiscal_ObterPedidos(objGeracaoNFiscal)

    Exit Function
    
Erro_GeracaoNFiscal_ObterPedidos:
    
    GeracaoNFiscal_ObterPedidos = gErr
    
    Select Case gErr

        Case 24473
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145676)
        
    End Select
    
    Exit Function

End Function


Function PedidoDeVenda_Grava_Comissoes(objPedidoDeVenda As ClassPedidoDeVenda) As Long
'grava no bd as comissoes de um PV (tabela ComissoesPedVendas)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Grava_Comissoes
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_GRAVA_COMISSOES", X)
    If lErro <> SUCESSO Then gError 24561
    
    
    PedidoDeVenda_Grava_Comissoes = X.PedidoDeVenda_Grava_Comissoes(objPedidoDeVenda)

    Exit Function
    
Erro_PedidoDeVenda_Grava_Comissoes:
    
    PedidoDeVenda_Grava_Comissoes = gErr
    
    Select Case gErr

        Case 24561
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145677)
        
    End Select
    
    Exit Function

End Function

Function ComissoesPV_Le(objPedidoDeVenda As ClassPedidoDeVenda) As Long
'carrega em colComissoes a colecao com as comissoes do PV passado em objPedidoDeVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ComissoesPV_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_COMISSOESPV_LE", X)
    If lErro <> SUCESSO Then gError 24562
    
    
    ComissoesPV_Le = X.ComissoesPV_Le(objPedidoDeVenda)

    Exit Function
    
Erro_ComissoesPV_Le:
    
    ComissoesPV_Le = gErr
    
    Select Case gErr

        Case 24562
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145678)
        
    End Select
    
    Exit Function

End Function

Function Series_Le(colSerie As colSerie) As Long
'Lê todos as Series contidas na tabela Serie e as coloca na coleçao colSerie

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Series_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIES_LE", X)
    If lErro <> SUCESSO Then gError 24563
    
    
    Series_Le = X.Series_Le(colSerie)

    Exit Function
    
Erro_Series_Le:
    
    Series_Le = gErr
    
    Select Case gErr

        Case 24563
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145679)
        
    End Select
    
    Exit Function

End Function

Function Serie_Le(objSerie As ClassSerie) As Long
'Lê a Serie a partir do código em objSerie.
'Devolve os dados em objSerie.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_LE", X)
    If lErro <> SUCESSO Then gError 24564
    
    
    Serie_Le = X.Serie_Le(objSerie)

    Exit Function
    
Erro_Serie_Le:
    
    Serie_Le = gErr
    
    Select Case gErr

        Case 24564
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145680)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE", X)
    If lErro <> SUCESSO Then gError 24995
    
    
    PedidoDeVenda_Le = X.PedidoDeVenda_Le(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le:
    
    PedidoDeVenda_Le = gErr
    
    Select Case gErr

        Case 24995
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145681)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le_Completo(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado (incluindo tributacao, itens de pedido, parcelas, Comissoes, bloqueios)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_Completo
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_COMPLETO", X)
    If lErro <> SUCESSO Then gError 24996
    
    
    PedidoDeVenda_Le_Completo = X.PedidoDeVenda_Le_Completo(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le_Completo:
    
    PedidoDeVenda_Le_Completo = gErr
    
    Select Case gErr

        Case 24996
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145682)
        
    End Select
    
    Exit Function

End Function

Function ReservaItem_Le(objItemPedido As ClassItemPedido, objReservaBD As ClassReserva) As Long
'Lê a Resrva com o Produto, FilialEmpresa e Código de Pedido do Ítem passado e com o Almoxarifado Passado no objReserva

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ReservaItem_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_RESERVAITEM_LE", X)
    If lErro <> SUCESSO Then gError 24997
    
    
    ReservaItem_Le = X.ReservaItem_Le(objItemPedido, objReservaBD)

    Exit Function
    
Erro_ReservaItem_Le:
    
    ReservaItem_Le = gErr
    
    Select Case gErr

        Case 24997
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145683)
        
    End Select
    
    Exit Function

End Function

Function CanalVenda_Le(objCanal As ClassCanalVenda) As Long
'Lê os dados do Canal, cujo codigo foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CanalVenda_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CANALVENDA_LE", X)
    If lErro <> SUCESSO Then gError 24998
    
    
    CanalVenda_Le = X.CanalVenda_Le(objCanal)

    Exit Function
    
Erro_CanalVenda_Le:
    
    CanalVenda_Le = gErr
    
    Select Case gErr

        Case 24998
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145684)
        
    End Select
    
    Exit Function

End Function

Function CanalVenda_Automatico(iCodigo As Integer) As Long
'retorna o número do proximo Canal disponivel
'??? trocar p/leitura nao ser feita em CPRConfig
'??? usar funcao padrao p/este tipo de codigo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CanalVenda_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CANALVENDA_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 24999
    
    
    CanalVenda_Automatico = X.CanalVenda_Automatico(iCodigo)

    Exit Function
    
Erro_CanalVenda_Automatico:
    
    CanalVenda_Automatico = gErr
    
    Select Case gErr

        Case 24999
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145685)
        
    End Select
    
    Exit Function

End Function

Function CanalVenda_Grava(objCanal As ClassCanalVenda) As Long
'Atualiza ou insere um novo registro na tabela CanalVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CanalVenda_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CANALVENDA_GRAVA", X)
    If lErro <> SUCESSO Then gError 29500
    
    
    CanalVenda_Grava = X.CanalVenda_Grava(objCanal)

    Exit Function
    
Erro_CanalVenda_Grava:
    
    CanalVenda_Grava = gErr
    
    Select Case gErr

        Case 29500
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145686)
        
    End Select
    
    Exit Function

End Function

Function CanalVenda_Exclui(objCanal As ClassCanalVenda) As Long
'Exclui o Canal do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CanalVenda_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CANALVENDA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 29501
    
    
    CanalVenda_Exclui = X.CanalVenda_Exclui(objCanal)

    Exit Function
    
Erro_CanalVenda_Exclui:
    
    CanalVenda_Exclui = gErr
    
    Select Case gErr

        Case 29501
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145687)
        
    End Select
    
    Exit Function

End Function


Function FATConfig_Grava_Configs(colConfigs As ColFATConfig) As Long
'grava na tabela FATConfig uma colecao de registros de configuracao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FATConfig_Grava_Configs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CONFIG_GRAVA_CONFIGS", X)
    If lErro <> SUCESSO Then gError 19327
    
    
    FATConfig_Grava_Configs = X.FATConfig_Grava_Configs(colConfigs)

    Exit Function
    
Erro_FATConfig_Grava_Configs:
    
    FATConfig_Grava_Configs = gErr
    
    Select Case gErr

        Case 19327
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145688)
        
    End Select
    
    Exit Function

End Function

Function TipoDeBloqueio_Automatico(iCodigo As Integer) As Long
'retorna o número do proximo Tipo disponivel
'??? substituir por chamada a funcao generic

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeBloqueio_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODEBLOQUEIO_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 19665
    
    
    TipoDeBloqueio_Automatico = X.TipoDeBloqueio_Automatico(iCodigo)

    Exit Function
    
Erro_TipoDeBloqueio_Automatico:
    
    TipoDeBloqueio_Automatico = gErr
    
    Select Case gErr

        Case 19665
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145689)
        
    End Select
    
    Exit Function

End Function

Function TipoDeBloqueio_Le(objTipo As ClassTipoDeBloqueio) As Long
'Lê os dados do Tipo, cujo codigo foi passado como parâmetro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeBloqueio_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODEBLOQUEIO_LE", X)
    If lErro <> SUCESSO Then gError 19666
    
    
    TipoDeBloqueio_Le = X.TipoDeBloqueio_Le(objTipo)

    Exit Function
    
Erro_TipoDeBloqueio_Le:
    
    TipoDeBloqueio_Le = gErr
    
    Select Case gErr

        Case 19666
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145690)
        
    End Select
    
    Exit Function

End Function

Function TipoDeBloqueio_Exclui(objTipo As ClassTipoDeBloqueio) As Long
'Exclui o tipo do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeBloqueio_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODEBLOQUEIO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 19667
    
    
    TipoDeBloqueio_Exclui = X.TipoDeBloqueio_Exclui(objTipo)

    Exit Function
    
Erro_TipoDeBloqueio_Exclui:
    
    TipoDeBloqueio_Exclui = gErr
    
    Select Case gErr

        Case 19667
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145691)
        
    End Select
    
    Exit Function

End Function

Function TipoDeBloqueio_Grava(objTipo As ClassTipoDeBloqueio) As Long
'Atualiza ou insere um novo registro na tabela Tipos de Bloqueio

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDeBloqueio_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODEBLOQUEIO_GRAVA", X)
    If lErro <> SUCESSO Then gError 19668
    
    
    TipoDeBloqueio_Grava = X.TipoDeBloqueio_Grava(objTipo)

    Exit Function
    
Erro_TipoDeBloqueio_Grava:
    
    TipoDeBloqueio_Grava = gErr
    
    Select Case gErr

        Case 19668
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145692)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Le_Tributacao(objNFiscal As ClassNFiscal) As Long
'le a parte de tributacao associada a uma NF e a incorpora a objNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Tributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_TRIBUTACAO", X)
    If lErro <> SUCESSO Then gError 19691
    
    
    NFiscal_Le_Tributacao = X.NFiscal_Le_Tributacao(objNFiscal)

    Exit Function
    
Erro_NFiscal_Le_Tributacao:
    
    NFiscal_Le_Tributacao = gErr
    
    Select Case gErr

        Case 19691
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145693)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Le_TributacaoCompl(objNFiscal As ClassNFiscal, objTribComplNF As ClassTribComplNF, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_TributacaoCompl
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_TRIBUTACAOCOMPL", X)
    If lErro <> SUCESSO Then gError 19692
    
    
    NFiscal_Le_TributacaoCompl = X.NFiscal_Le_TributacaoCompl(objNFiscal, objTribComplNF, lComando)

    Exit Function
    
Erro_NFiscal_Le_TributacaoCompl:
    
    NFiscal_Le_TributacaoCompl = gErr
    
    Select Case gErr

        Case 19692
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145694)
        
    End Select
    
    Exit Function

End Function

Function NFiscalItens_Le_Trib(objNFiscal As ClassNFiscal) As Long
'Lê os dados de tributacao dos ítens da Nota Fiscal passada em objNFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalItens_Le_Trib
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALITENS_LE_TRIB", X)
    If lErro <> SUCESSO Then gError 19693
    
    
    NFiscalItens_Le_Trib = X.NFiscalItens_Le_Trib(objNFiscal)

    Exit Function
    
Erro_NFiscalItens_Le_Trib:
    
    NFiscalItens_Le_Trib = gErr
    
    Select Case gErr

        Case 19693
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145695)
        
    End Select
    
    Exit Function

End Function

'Function NFiscalBaixada_Le_Num_Forn(objNFiscal As ClassNFiscal) As Long
''Lê a Nota Fiscal Baixada a partir de TipoNFiscal, Fornecedor, Filial, Serie e NFiscal.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscalBaixada_Le_Num_Forn
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALBAIXADA_LE_NUM_FORN",X)
'    If lErro <> SUCESSO Then gError 19695
'
'
'    NFiscalBaixada_Le_Num_Forn = X.NFiscalBaixada_Le_Num_Forn(objNFiscal)
'
'    Exit Function
'
'Erro_NFiscalBaixada_Le_Num_Forn:
'
'    NFiscalBaixada_Le_Num_Forn = gErr
'
'    Select Case gErr
'
'        Case 19695
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145696)
'
'    End Select
'
'    Exit Function
'
'End Function


Function PrevVenda_Le(objPrevVenda As ClassPrevVenda) As Long
'Lê a Previsão de Venda a partir da chave fornecida em objPrevVenda. Retorna os dados em objPrevVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PrevVenda_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PREVVENDA_LE", X)
    If lErro <> SUCESSO Then gError 19765
    
    
    PrevVenda_Le = X.PrevVenda_Le(objPrevVenda)

    Exit Function
    
Erro_PrevVenda_Le:
    
    PrevVenda_Le = gErr
    
    Select Case gErr

        Case 19765
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145697)
        
    End Select
    
    Exit Function

End Function

Function PrevVenda_Grava(objPrevVenda As ClassPrevVenda) As Long
'Grava a Previsão de Venda no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PrevVenda_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PREVVENDA_GRAVA", X)
    If lErro <> SUCESSO Then gError 19766
    
    
    PrevVenda_Grava = X.PrevVenda_Grava(objPrevVenda)

    Exit Function
    
Erro_PrevVenda_Grava:
    
    PrevVenda_Grava = gErr
    
    Select Case gErr

        Case 19766
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145698)
        
    End Select
    
    Exit Function

End Function

Function PrevVenda_Exclui(objPrevVenda As ClassPrevVenda) As Long
'Exclui a previsão de venda do BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PrevVenda_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PREVVENDA_EXCLUI", X)
    If lErro <> SUCESSO Then gError 19767
    
    
    PrevVenda_Exclui = X.PrevVenda_Exclui(objPrevVenda)

    Exit Function
    
Erro_PrevVenda_Exclui:
    
    PrevVenda_Exclui = gErr
    
    Select Case gErr

        Case 19767
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145699)
        
    End Select
    
    Exit Function

End Function

Function ItensPedido_Le_Todos(lComando As Long, objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê no BD a coleção de ítens do Pedido de Venda e devolve em colItensPedido dentro de objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedido_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITENSPEDIDO_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 36659
    
    
    ItensPedido_Le_Todos = X.ItensPedido_Le_Todos(lComando, objPedidoVenda)

    Exit Function
    
Erro_ItensPedido_Le_Todos:
    
    ItensPedido_Le_Todos = gErr
    
    Select Case gErr

        Case 36659
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145700)
        
    End Select
    
    Exit Function

End Function

Function Bloqueio_Libera(colBloqueioPV As Collection) As Long
'libera os bloqueios passados na colecao colBloqueioPV

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Bloqueio_Libera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIO_LIBERA", X)
    If lErro <> SUCESSO Then gError 36707
    
    
    Bloqueio_Libera = X.Bloqueio_Libera(colBloqueioPV)

    Exit Function
    
Erro_Bloqueio_Libera:
    
    Bloqueio_Libera = gErr
    
    Select Case gErr

        Case 36707
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145701)
        
    End Select
    
    Exit Function

End Function

Function BloqueioPV_Insere(lComando As Long, objBloqueioPV As ClassBloqueioPV) As Long
'insere um bloqueio de pedido de venda
'IMPORTANTE: Tem que ser chamada dentro de TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueioPV_Insere
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOPV_INSERE", X)
    If lErro <> SUCESSO Then gError 36708
    
    
    BloqueioPV_Insere = X.BloqueioPV_Insere(lComando, objBloqueioPV)

    Exit Function
    
Erro_BloqueioPV_Insere:
    
    BloqueioPV_Insere = gErr
    
    Select Case gErr

        Case 36708
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145702)
        
    End Select
    
    Exit Function

End Function


Function LiberacaoCredito_Le(objLiberacaoCredito As ClassLiberacaoCredito) As Long
'Lê a liberacao de credito a partir do código do usuario.
'Devolve os dados em objLiberacaoCredito.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoCredito_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAOCREDITO_LE", X)
    If lErro <> SUCESSO Then gError 36969
    
    
    LiberacaoCredito_Le = X.LiberacaoCredito_Le(objLiberacaoCredito)

    Exit Function
    
Erro_LiberacaoCredito_Le:
    
    LiberacaoCredito_Le = gErr
    
    Select Case gErr

        Case 36969
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145703)
        
    End Select
    
    Exit Function

End Function

Function ValorLiberadoCredito_Le(objValorLiberadoCredito As ClassValorLiberadoCredito) As Long
'Lê a estatistica de liberação de credito de um usuario em um determinado ano
'Devolve os dados em objValorLiberadoCredito.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ValorLiberadoCredito_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_VALORLIBERADOCREDITO_LE", X)
    If lErro <> SUCESSO Then gError 36974
    
    
    ValorLiberadoCredito_Le = X.ValorLiberadoCredito_Le(objValorLiberadoCredito)

    Exit Function
    
Erro_ValorLiberadoCredito_Le:
    
    ValorLiberadoCredito_Le = gErr
    
    Select Case gErr

        Case 36974
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145704)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le_Lock(lComando As Long, objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_LOCK", X)
    If lErro <> SUCESSO Then gError 36990
    
    
    PedidoDeVenda_Le_Lock = X.PedidoDeVenda_Le_Lock(lComando, objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le_Lock:
    
    PedidoDeVenda_Le_Lock = gErr
    
    Select Case gErr

        Case 36990
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145705)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le_ComBloqueio(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda a partir do Código passado, bloqueando-o

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_ComBloqueio
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_COMBLOQUEIO", X)
    If lErro <> SUCESSO Then gError 44182
    
    
    PedidoDeVenda_Le_ComBloqueio = X.PedidoDeVenda_Le_ComBloqueio(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le_ComBloqueio:
    
    PedidoDeVenda_Le_ComBloqueio = gErr
    
    Select Case gErr

        Case 44182
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145706)
        
    End Select
    
    Exit Function

End Function

Function SldMesFat_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaFat As ClassSldDiaFat) As Long
'Atualiza a tabela de SldMesFat utilizando os valores passados atraves de objSldDiaFat.
'IMPORTANTE: TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldMesFat_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SLDMESFAT_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 44308
    
    
    SldMesFat_Atualiza = X.SldMesFat_Atualiza(lComando1, lComando2, objSldDiaFat)

    Exit Function
    
Erro_SldMesFat_Atualiza:
    
    SldMesFat_Atualiza = gErr
    
    Select Case gErr

        Case 44308
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145707)
        
    End Select
    
    Exit Function

End Function

Function SldDiaFat_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, objSldDiaFat As ClassSldDiaFat) As Long
'Atualiza a tabela de SldDiaFat utilizando os valores passados atraves de objSldDiaFat.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_SldDiaFat_Atualiza
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SLDDIAFAT_ATUALIZA", X)
    If lErro <> SUCESSO Then gError 44309
    
    
    SldDiaFat_Atualiza = X.SldDiaFat_Atualiza(lComando1, lComando2, objSldDiaFat)

    Exit Function
    
Erro_SldDiaFat_Atualiza:
    
    SldDiaFat_Atualiza = gErr
    
    Select Case gErr

        Case 44309
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145708)
        
    End Select
    
    Exit Function

End Function

'mario
Function GeracaoNFiscal_GerarNFs(objGeracaoNFiscal As ClassGeracaoNFiscal) As Long
'gera as notas fiscais a partir dos pedidos selecionados.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoNFiscal_GerarNFs
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_GERACAONFISCAL_GERARNFS", X)
    If lErro <> SUCESSO Then gError 44339
    
    
    GeracaoNFiscal_GerarNFs = X.GeracaoNFiscal_GerarNFs(objGeracaoNFiscal)

    Exit Function
    
Erro_GeracaoNFiscal_GerarNFs:
    
    GeracaoNFiscal_GerarNFs = gErr
    
    Select Case gErr

        Case 44339
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145709)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscalSaidaNum_Automatico(sSerie As String, lNumAuto As Long) As Long
'Retorna o número do próximo NumNFiscal disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalSaidaNum_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALSAIDANUM_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 44348
    
    
    NFiscalSaidaNum_Automatico = X.NFiscalSaidaNum_Automatico(sSerie, lNumAuto)

    Exit Function
    
Erro_NFiscalSaidaNum_Automatico:
    
    NFiscalSaidaNum_Automatico = gErr
    
    Select Case gErr

        Case 44348
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145710)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscal_Grava_Trans(alComando() As Long, objNFiscal As ClassNFiscal) As Long
'grava a nota fiscal
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Grava_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_GRAVA_TRANS", X)
    If lErro <> SUCESSO Then gError 44359
    
    
    NFiscal_Grava_Trans = X.NFiscal_Grava_Trans(alComando(), objNFiscal)

    Exit Function
    
Erro_NFiscal_Grava_Trans:
    
    NFiscal_Grava_Trans = gErr
    
    Select Case gErr

        Case 44359
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145711)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscal_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'grava uma nota fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_GRAVA", X)
    If lErro <> SUCESSO Then gError 44361
    
    
    NFiscal_Grava = X.NFiscal_Grava(objNFiscal, objContabil)

    Exit Function
    
Erro_NFiscal_Grava:
    
    NFiscal_Grava = gErr
    
    Select Case gErr

        Case 44361
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145712)
        
    End Select
    
    Exit Function

End Function

'mario
'Function NFiscalBaixada_Le_Num_Cli(objNFiscal As ClassNFiscal) As Long
''Lê a Nota Fiscal Baixada a partir de TipoNFiscal, Cliente, Filial, Série e NFiscal.
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscalBaixada_Le_Num_Cli
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALBAIXADA_LE_NUM_CLI",X)
'    If lErro <> SUCESSO Then gError 44362
'
'
'    NFiscalBaixada_Le_Num_Cli = X.NFiscalBaixada_Le_Num_Cli(objNFiscal)
'
'    Exit Function
'
'Erro_NFiscalBaixada_Le_Num_Cli:
'
'    NFiscalBaixada_Le_Num_Cli = gErr
'
'    Select Case gErr
'
'        Case 44362
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145713)
'
'    End Select
'
'    Exit Function
'
'End Function

'mario
Function NFiscal_Le_Num_Cli(objNFiscal As ClassNFiscal, ByVal iTipo As Integer) As Long
'Lê a Nota Fiscal a partir de TipoNFiscal, Cliente, Filial, Serie e NFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Num_Cli
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_NUM_CLI", X)
    If lErro <> SUCESSO Then gError 44363
    
    
    NFiscal_Le_Num_Cli = X.NFiscal_Le_Num_Cli(objNFiscal, iTipo)

    Exit Function
    
Erro_NFiscal_Le_Num_Cli:
    
    NFiscal_Le_Num_Cli = gErr
    
    Select Case gErr

        Case 44363
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145714)
        
    End Select
    
    Exit Function

End Function

'mario
Function AlocacoesNF_Le(objNFiscal As ClassNFiscal) As Long
'Lê as Alocações dos itens da Nota Fiscal passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_AlocacoesNF_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ALOCACOESNF_LE", X)
    If lErro <> SUCESSO Then gError 44364
    
    
    AlocacoesNF_Le = X.AlocacoesNF_Le(objNFiscal)

    Exit Function
    
Erro_AlocacoesNF_Le:
    
    AlocacoesNF_Le = gErr
    
    Select Case gErr

        Case 44364
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145715)
        
    End Select
    
    Exit Function

End Function

'''mario
''Function TiposDocInfo_Le(colTipoDocInfo As colTipoDocInfo) As Long
'''Lê todos os registros de TiposDocInfo
''
''Dim X As Object, lErro As Long
''
''On Error GoTo Erro_TiposDocInfo_Le
''
''
''
''
''    lErro = Rotina_ObterObjetoFuncao("FAT_TIPOSDOCINFO_LE",X)
''    If lErro <> SUCESSO Then gError 44366
''
''
''    TiposDocInfo_Le = X.TiposDocInfo_Le(colTipoDocInfo)
''
''    Exit Function
''
''Erro_TiposDocInfo_Le:
''
''    TiposDocInfo_Le = gErr
''
''    Select Case gErr
''
''        Case 44366
''
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145716)
''
''    End Select
''
''    Exit Function
''
''End Function

'mario
Function TiposDocInfo_Le_NomeTela(objTipoDocInfo As ClassTipoDocInfo) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDocInfo_Le_NomeTela
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPOSDOCINFO_LE_NOMETELA", X)
    If lErro <> SUCESSO Then gError 44367
    
    
    TiposDocInfo_Le_NomeTela = X.TiposDocInfo_Le_NomeTela(objTipoDocInfo)

    Exit Function
    
Erro_TiposDocInfo_Le_NomeTela:
    
    TiposDocInfo_Le_NomeTela = gErr
    
    Select Case gErr

        Case 44367
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145717)
        
    End Select
    
    Exit Function

End Function

'mario
Function TabelaDePreco_Lock(iTabelaPreco As Integer) As Long
'Faz lock da Tabela de Preço passada como parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaDePreco_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELADEPRECO_LOCK", X)
    If lErro <> SUCESSO Then gError 44368
    
    
    TabelaDePreco_Lock = X.TabelaDePreco_Lock(iTabelaPreco)

    Exit Function
    
Erro_TabelaDePreco_Lock:
    
    TabelaDePreco_Lock = gErr
    
    Select Case gErr

        Case 44368
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145718)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscal_Le_NFOriginal(objNFOriginal As ClassNFiscal) As Long
'Lê a Nota Fiscal Original

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_NFOriginal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_NFORIGINAL", X)
    If lErro <> SUCESSO Then gError 44369
    
    
    NFiscal_Le_NFOriginal = X.NFiscal_Le_NFOriginal(objNFOriginal)

    Exit Function
    
Erro_NFiscal_Le_NFOriginal:
    
    NFiscal_Le_NFOriginal = gErr
    
    Select Case gErr

        Case 44369
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145719)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscal_Testa_Existencia(lComando As Long, lComando1 As Long, objNFiscal As ClassNFiscal) As Long
'verifical se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Testa_Existencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_TESTA_EXISTENCIA", X)
    If lErro <> SUCESSO Then gError 44371
    
    
    NFiscal_Testa_Existencia = X.NFiscal_Testa_Existencia(lComando, lComando1, objNFiscal)

    Exit Function
    
Erro_NFiscal_Testa_Existencia:
    
    NFiscal_Testa_Existencia = gErr
    
    Select Case gErr

        Case 44371
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145720)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscal_TabelaPreco_Lock(lComando As Long, objNFiscal As ClassNFiscal) As Long
'executa o lock da tabela de preços

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_TabelaPreco_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_TABELAPRECO_LOCK", X)
    If lErro <> SUCESSO Then gError 44372
    
    
    NFiscal_TabelaPreco_Lock = X.NFiscal_TabelaPreco_Lock(lComando, objNFiscal)

    Exit Function
    
Erro_NFiscal_TabelaPreco_Lock:
    
    NFiscal_TabelaPreco_Lock = gErr
    
    Select Case gErr

        Case 44372
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145721)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscalFatura_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Grava NFiscal Fatura

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatura_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALFATURA_GRAVA", X)
    If lErro <> SUCESSO Then gError 44373
    
    
    NFiscalFatura_Grava = X.NFiscalFatura_Grava(objNFiscal, objContabil)

    Exit Function
    
Erro_NFiscalFatura_Grava:
    
    NFiscalFatura_Grava = gErr
    
    Select Case gErr

        Case 44373
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145722)
        
    End Select
    
    Exit Function

End Function

'mario
Function NFiscalFatura_Grava_Trans(alComando() As Long, objNFiscal As ClassNFiscal) As Long
'Grava NFiscal Fatura
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatura_Grava_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALFATURA_GRAVA_TRANS", X)
    If lErro <> SUCESSO Then gError 44374
    
    
    NFiscalFatura_Grava_Trans = X.NFiscalFatura_Grava_Trans(alComando(), objNFiscal)

    Exit Function
    
Erro_NFiscalFatura_Grava_Trans:
    
    NFiscalFatura_Grava_Trans = gErr
    
    Select Case gErr

        Case 44374
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145723)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Utilizar(lComandoSel As Long, lComandoDel As Long, objMovEstoque As ClassMovEstoque, objItemPV As ClassItemPedido, objAlmoxarifado As ClassAlmoxarifado, dQtdeUtilizada As Double) As Long
'Utiliza a reserva. Ou seja serão gerados movimentos de utilização e possivelmente cancelamento do que não foi utilizado. Além disso a reserva é excluida
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Utilizar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_RESERVA_UTILIZAR", X)
    If lErro <> SUCESSO Then gError 44342
    
    
    Reserva_Utilizar = X.Reserva_Utilizar(lComandoSel, lComandoDel, objMovEstoque, objItemPV, objAlmoxarifado, dQtdeUtilizada)

    Exit Function
    
Erro_Reserva_Utilizar:
    
    Reserva_Utilizar = gErr
    
    Select Case gErr

        Case 44342
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145724)
        
    End Select
    
    Exit Function

End Function

Function Reserva_Cancelar(lComandoSel As Long, lComandoDel As Long, objMovEstoque As ClassMovEstoque, objItemPV As ClassItemPedido, objAlmoxarifado As ClassAlmoxarifado) As Long
'exclui uma reserva da tabela de reservas e gera movto
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Reserva_Cancelar
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_RESERVA_CANCELAR", X)
    If lErro <> SUCESSO Then gError 44343
    
    
    Reserva_Cancelar = X.Reserva_Cancelar(lComandoSel, lComandoDel, objMovEstoque, objItemPV, objAlmoxarifado)

    Exit Function
    
Erro_Reserva_Cancelar:
    
    Reserva_Cancelar = gErr
    
    Select Case gErr

        Case 44343
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145725)
        
    End Select
    
    Exit Function

End Function

'mario
Function MovEstoque_Gera(alComando() As Long, alComando1() As Long, objNFiscal As ClassNFiscal) As Long
'gera os movimentos de estoque, se necessários

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoque_Gera
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_MOVESTOQUE_GERA", X)
    If lErro <> SUCESSO Then gError 44356
    
    
    MovEstoque_Gera = X.MovEstoque_Gera(alComando(), alComando1(), objNFiscal)

    Exit Function
    
Erro_MovEstoque_Gera:
    
    MovEstoque_Gera = gErr
    
    Select Case gErr

        Case 44356
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 145726)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPV_Grava1(objPedidoVenda As ClassPedidoDeVenda, objBloqueioPV As ClassBloqueioPV) As Long
'Insere no BD os Bloqueios do Pedido de Venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Grava1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOSPV_GRAVA1", X)
    If lErro <> SUCESSO Then gError 44395
    
    
    BloqueiosPV_Grava1 = X.BloqueiosPV_Grava1(objPedidoVenda, objBloqueioPV)

    Exit Function
    
Erro_BloqueiosPV_Grava1:
    
    BloqueiosPV_Grava1 = gErr
    
    Select Case gErr

        Case 44395
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145727)
        
    End Select
    
    Exit Function

End Function

Function NFiscalFatura_Le_NFOriginal(objNFOriginal As ClassNFiscal) As Long
'Lê a Nota Fiscal Original

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFatura_Le_NFOriginal
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALFATURA_LE_NFORIGINAL", X)
    If lErro <> SUCESSO Then gError 45348
    
    
    NFiscalFatura_Le_NFOriginal = X.NFiscalFatura_Le_NFOriginal(objNFOriginal)

    Exit Function
    
Erro_NFiscalFatura_Le_NFOriginal:
    
    NFiscalFatura_Le_NFOriginal = gErr
    
    Select Case gErr

        Case 45348
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145728)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Testa_Credito(objNFiscal As ClassNFiscal) As Long
' verifica se o cliente tem crédito. Se não tiver pede autorização. Se possui o credito, retorna SUCESSO.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Testa_Credito
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_TESTA_CREDITO", X)
    If lErro <> SUCESSO Then gError 44458
    
    
    NFiscal_Testa_Credito = X.NFiscal_Testa_Credito(objNFiscal)

    Exit Function
    
Erro_NFiscal_Testa_Credito:
    
    NFiscal_Testa_Credito = gErr
    
    Select Case gErr

        Case 44458
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145729)
        
    End Select
    
    Exit Function

End Function

Function ValorLiberadoCredito_Lock(objValorLiberadoCredito As ClassValorLiberadoCredito) As Long
'Lê e loca a estatistica de liberação de credito de um usuario em um determinado ano
'Devolve os dados em objValorLiberadoCredito.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ValorLiberadoCredito_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_VALORLIBERADOCREDITO_LOCK", X)
    If lErro <> SUCESSO Then gError 44472
    
    
    ValorLiberadoCredito_Lock = X.ValorLiberadoCredito_Lock(objValorLiberadoCredito)

    Exit Function
    
Erro_ValorLiberadoCredito_Lock:
    
    ValorLiberadoCredito_Lock = gErr
    
    Select Case gErr

        Case 44472
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145730)
        
    End Select
    
    Exit Function

End Function

Function ValorLiberadoCredito_Grava(sCodUsuario As String, dValor As Double, dtData As Date) As Long
'Atualiza a estatistica de liberação de credito de um usuario (sCodUsuario) em um determinado mes/ano (dtData) com o valor (dValor).
'se não existir a estitistica para o usuario/ano, cria.
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ValorLiberadoCredito_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_VALORLIBERADOCREDITO_GRAVA", X)
    If lErro <> SUCESSO Then gError 44473
    
    
    ValorLiberadoCredito_Grava = X.ValorLiberadoCredito_Grava(sCodUsuario, dValor, dtData)

    Exit Function
    
Erro_ValorLiberadoCredito_Grava:
    
    ValorLiberadoCredito_Grava = gErr
    
    Select Case gErr

        Case 44473
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145731)
        
    End Select
    
    Exit Function

End Function

Function LiberacaoCredito_Lock(objLiberacaoCredito As ClassLiberacaoCredito) As Long
'Lê e loca a liberacao de credito a partir do código do usuario.
'Devolve os dados em objLiberacaoCredito.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoCredito_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAOCREDITO_LOCK", X)
    If lErro <> SUCESSO Then gError 44481
    
    
    LiberacaoCredito_Lock = X.LiberacaoCredito_Lock(objLiberacaoCredito)

    Exit Function
    
Erro_LiberacaoCredito_Lock:
    
    LiberacaoCredito_Lock = gErr
    
    Select Case gErr

        Case 44481
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145732)
        
    End Select
    
    Exit Function

End Function

Function BloqueiosPV_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê os Bloqueios associados ao objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BloqueiosPV_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BLOQUEIOSPV_LE", X)
    If lErro <> SUCESSO Then gError 45519
    
    
    BloqueiosPV_Le = X.BloqueiosPV_Le(objPedidoVenda)

    Exit Function
    
Erro_BloqueiosPV_Le:
    
    BloqueiosPV_Le = gErr
    
    Select Case gErr

        Case 45519
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145733)
        
    End Select
    
    Exit Function

End Function

Function LiberacaoCredito_Le_Todas(colLiberacaoCredito As Collection) As Long
'Guarda em colLiberacaoCredito todas as alçadas cadastradas

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoCredito_Le_Todas
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAOCREDITO_LE_TODAS", X)
    If lErro <> SUCESSO Then gError 45569
    
    
    LiberacaoCredito_Le_Todas = X.LiberacaoCredito_Le_Todas(colLiberacaoCredito)

    Exit Function
    
Erro_LiberacaoCredito_Le_Todas:
    
    LiberacaoCredito_Le_Todas = gErr
    
    Select Case gErr

        Case 45569
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145734)
        
    End Select
    
    Exit Function

End Function

Function LiberacaoCredito_Grava(objLiberacaoCredito As ClassLiberacaoCredito) As Long
'Grava uma alçada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoCredito_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAOCREDITO_GRAVA", X)
    If lErro <> SUCESSO Then gError 45570
    
    
    LiberacaoCredito_Grava = X.LiberacaoCredito_Grava(objLiberacaoCredito)

    Exit Function
    
Erro_LiberacaoCredito_Grava:
    
    LiberacaoCredito_Grava = gErr
    
    Select Case gErr

        Case 45570
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145735)
        
    End Select
    
    Exit Function

End Function

Function LiberacaoCredito_Exclui(objLiberacaoCredito As ClassLiberacaoCredito) As Long
'Exclui a alçada passada como parametro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacaoCredito_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACAOCREDITO_EXCLUI", X)
    If lErro <> SUCESSO Then gError 45571
    
    
    LiberacaoCredito_Exclui = X.LiberacaoCredito_Exclui(objLiberacaoCredito)

    Exit Function
    
Erro_LiberacaoCredito_Exclui:
    
    LiberacaoCredito_Exclui = gErr
    
    Select Case gErr

        Case 45571
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145736)
        
    End Select
    
    Exit Function

End Function

Function Almoxarifados_Le_Quantidade(lAlmoxarifados As Long) As Long
'Lê o número de Almoxarifados existentes no Sistema

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Almoxarifados_Le_Quantidade
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ALMOXARIFADOS_LE_QUANTIDADE", X)
    If lErro <> SUCESSO Then gError 45573
    
    
    Almoxarifados_Le_Quantidade = X.Almoxarifados_Le_Quantidade(lAlmoxarifados)

    Exit Function
    
Erro_Almoxarifados_Le_Quantidade:
    
    Almoxarifados_Le_Quantidade = gErr
    
    Select Case gErr

        Case 45573
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145737)
        
    End Select
    
    Exit Function

End Function

Function TabelaPrecoItem_AtualizarPrecosPercentual(sProdInicial As String, sProdFinal As String, dPercentual As Double, colTabelas As Collection) As Long
'colTabelas deve conter strings formados pelo identificador da tabela (um integer)

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoItem_AtualizarPrecosPercentual
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELAPRECOITEM_ATUALIZARPRECOSPERC", X)
    If lErro <> SUCESSO Then gError 45574
    
    
    TabelaPrecoItem_AtualizarPrecosPercentual = X.TabelaPrecoItem_AtualizarPrecosPercentual(sProdInicial, sProdFinal, dPercentual, colTabelas)

    Exit Function
    
Erro_TabelaPrecoItem_AtualizarPrecosPercentual:
    
    TabelaPrecoItem_AtualizarPrecosPercentual = gErr
    
    Select Case gErr

        Case 45574
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145738)
        
    End Select
    
    Exit Function

End Function

Function AlmoxarifadosFilial_Le_Quantidade(iFilialEmpresa As Integer, lNumAlmoxarifados As Long) As Long
'Retorna o número de Almoxarifados da FilialEmpresa

Dim X As Object, lErro As Long
    
On Error GoTo Erro_AlmoxarifadosFilial_Le_Quantidade
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ALMOXARIFADOSFILIAL_LE_QUANTIDADE", X)
    If lErro <> SUCESSO Then gError 45575
    
    
    AlmoxarifadosFilial_Le_Quantidade = X.AlmoxarifadosFilial_Le_Quantidade(iFilialEmpresa, lNumAlmoxarifados)

    Exit Function
    
Erro_AlmoxarifadosFilial_Le_Quantidade:
    
    AlmoxarifadosFilial_Le_Quantidade = gErr
    
    Select Case gErr

        Case 45575
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145739)
        
    End Select
    
    Exit Function

End Function


Function NFiscalInternaEntNaoDev_Le_Numero(objNFiscalOriginal As ClassNFiscal) As Long
'Busca Nota Fiscal, em [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo], com a Série e o Número Passados
'Nota Interna de Entrada e Com o Tipo com o Complementar NOrmal com o NOme da Tela NFiscal <> "NFiscalDev"

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalInternaEntNaoDev_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALINTERNAENTNAODEV_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 45601
    
    
    NFiscalInternaEntNaoDev_Le_Numero = X.NFiscalInternaEntNaoDev_Le_Numero(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalInternaEntNaoDev_Le_Numero:

    NFiscalInternaEntNaoDev_Le_Numero = gErr
    
    Select Case gErr

        Case 45601
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145740)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalExternaNaoDev_Le_Numero(objNFiscalOriginal As ClassNFiscal) As Long
'Pesquisa Nota Fiscal, nas Tabelas [NFiscal e TiposDocInfo] UNION [NotaFiscalBaixadas e TiposDocInfo],
'pela Serie, Número, Fornecedor, FilialForn, Cliente, FilialCli passados
'Procuta por Notas co o Tipo sendo NOta Fiscal Externa, com NOme da telas Fiscal <> "NFiscalDev" e complementar Normal.
'Lê todos os Dados da Tabela NFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalExternaNaoDev_Le_Numero
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALEXTERNANAODEV_LE_NUMERO", X)
    If lErro <> SUCESSO Then gError 45602
    
    
    NFiscalExternaNaoDev_Le_Numero = X.NFiscalExternaNaoDev_Le_Numero(objNFiscalOriginal)
    
    Exit Function
    
Erro_NFiscalExternaNaoDev_Le_Numero:

    NFiscalExternaNaoDev_Le_Numero = gErr
    
    Select Case gErr

        Case 45602
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145741)
        
    End Select
    
    Exit Function
    
End Function


Function ItemNFEntradaNaoDev_Le_NumNFItem(objNFiscal As ClassNFiscal, objItemNF As ClassItemNF) As Long
'Lê o Ítem de Uma Nota Fiscal de Entrada através do Número da Nota Fiscal, Série e Numero do Ítem passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItemNFEntradaNaoDev_Le_NumNFItem
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITEMNFENTRADANAODEV_LE_NUMNFITEM", X)
    If lErro <> SUCESSO Then gError 45606
    
    
    ItemNFEntradaNaoDev_Le_NumNFItem = X.ItemNFEntradaNaoDev_Le_NumNFItem(objNFiscal, objItemNF)
    
    Exit Function
    
Erro_ItemNFEntradaNaoDev_Le_NumNFItem:

    ItemNFEntradaNaoDev_Le_NumNFItem = gErr
    
    Select Case gErr

        Case 45606
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145742)
        
    End Select
    
    Exit Function
    
End Function

'mario 2/7/99 Remover NFiscalBaixada
'Function ItemNFBaixadaEntradaNaoDev_Le_NumNFItem(objNFiscal As ClassNFiscal, objItemNF As ClassItemNF) As Long
''Lê o Ítem de Uma Nota Fiscal de Entrada através do Número da Nota Fiscal, Série e Numero do Ítem passados
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ItemNFBaixadaEntradaNaoDev_Le_NumNFItem
'
'
'
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_ITEMNFBAIXADAENTNAODEV_LE_NUMNFITEM",X)
'    If lErro <> SUCESSO Then gError 45607
'
'
'    ItemNFBaixadaEntradaNaoDev_Le_NumNFItem = X.ItemNFBaixadaEntradaNaoDev_Le_NumNFItem(objNFiscal, objItemNF)
'
'    Exit Function
'
'Erro_ItemNFBaixadaEntradaNaoDev_Le_NumNFItem:
'
'    ItemNFBaixadaEntradaNaoDev_Le_NumNFItem = gErr
'
'    Select Case gErr
'
'        Case 45607
'
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145743)
'
'    End Select
'
'    Exit Function
'
'End Function

Function ItensNFiscalDev_Lock_Gravacao(objNFiscal As ClassNFiscal) As Long
'Faz a validação e os "lock's" dos dados ligados aos ítens da Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensNFiscalDev_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITENSNFISCALDEV_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 45608
    
    
    ItensNFiscalDev_Lock_Gravacao = X.ItensNFiscalDev_Lock_Gravacao(objNFiscal)
    
    Exit Function
    
Erro_ItensNFiscalDev_Lock_Gravacao:

    ItensNFiscalDev_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 45608
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145744)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalDev_Verifica_Existencia(objNFiscal As ClassNFiscal, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalDev_Verifica_Existencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALDEV_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then gError 45609
    
    
    NFiscalDev_Verifica_Existencia = X.NFiscalDev_Verifica_Existencia(objNFiscal, lComando)
    
    Exit Function
    
Erro_NFiscalDev_Verifica_Existencia:

    NFiscalDev_Verifica_Existencia = gErr
    
    Select Case gErr

        Case 45609
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145745)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalDev_Lock_Gravacao(objNFiscal As ClassNFiscal, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalDev_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALDEV_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 45610
    
    
    NFiscalDev_Lock_Gravacao = X.NFiscalDev_Lock_Gravacao(objNFiscal, lComando)
    
    Exit Function
    
Erro_NFiscalDev_Lock_Gravacao:

    NFiscalDev_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 45610
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145746)
        
    End Select
    
    Exit Function
    
End Function


Function NFiscalDevolucao_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Grava uma NFiscal Devolução

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalDevolucao_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALDEVOLUCAO_GRAVA", X)
    If lErro <> SUCESSO Then gError 45612
    
    
    NFiscalDevolucao_Grava = X.NFiscalDevolucao_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalDevolucao_Grava:

    NFiscalDevolucao_Grava = gErr
    
    Select Case gErr

        Case 45612
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145747)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalDevolucao_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Saida-Devolução no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalDevolucao_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALDEVOLUCAO_GRAVA_BD", X)
    If lErro <> SUCESSO Then gError 45613
    
    
    NFiscalDevolucao_Grava_BD = X.NFiscalDevolucao_Grava_BD(objNFiscal)
    
    Exit Function
    
Erro_NFiscalDevolucao_Grava_BD:

    NFiscalDevolucao_Grava_BD = gErr
    
    Select Case gErr

        Case 45613
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145748)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscal_Le_Pedido(objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal com o TipoNFiscal, Número, Serie, DataEmissao, NumPedidoVenda e FilialPedido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Pedido
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_PEDIDO", X)
    If lErro <> SUCESSO Then gError 45614
    
    
    NFiscal_Le_Pedido = X.NFiscal_Le_Pedido(objNFiscal)
    
    Exit Function
    
Erro_NFiscal_Le_Pedido:

    NFiscal_Le_Pedido = gErr
    
    Select Case gErr

        Case 45614
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145749)
        
    End Select
    
    Exit Function
    
End Function

Function PedidoVendaBaixado_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê o Pedido de Venda Baixado com o Código e a Filial passados

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoVendaBaixado_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDOVENDABAIXADO_LE", X)
    If lErro <> SUCESSO Then gError 45615
    
    
    PedidoVendaBaixado_Le = X.PedidoVendaBaixado_Le(objPedidoVenda)
    
    Exit Function
    
Erro_PedidoVendaBaixado_Le:

    PedidoVendaBaixado_Le = gErr
    
    Select Case gErr

        Case 45615
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145750)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalFaturaPedido_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal fatura gerada a partir de um pedido de venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalFaturaPedido_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALFATURAPEDIDO_GRAVA", X)
    If lErro <> SUCESSO Then gError 45616
    
    
    NFiscalFaturaPedido_Grava = X.NFiscalFaturaPedido_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalFaturaPedido_Grava:

    NFiscalFaturaPedido_Grava = gErr
    
    Select Case gErr

        Case 45616
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145751)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalPedido_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz a gravação de uma nota fiscal gerada a partir de um pedido de venda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalPedido_Grava
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALPEDIDO_GRAVA", X)
    If lErro <> SUCESSO Then gError 45617
    
    
    NFiscalPedido_Grava = X.NFiscalPedido_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalPedido_Grava:

    NFiscalPedido_Grava = gErr
    
    Select Case gErr

        Case 45617
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145752)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalRemessa_Grava_BD(objNFiscal As ClassNFiscal) As Long
'Faz a Inserção de uma Nota Fiscal de Saida-Remessa no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalRemessa_Grava_BD
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALREMESSA_GRAVA_BD", X)
    If lErro <> SUCESSO Then gError 45618
    
    
    NFiscalRemessa_Grava_BD = X.NFiscalRemessa_Grava_BD(objNFiscal)
    
    Exit Function
    
Erro_NFiscalRemessa_Grava_BD:

    NFiscalRemessa_Grava_BD = gErr
    
    Select Case gErr

        Case 45618
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145753)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalNum_Automatico(sSerie As String, lNumAuto As Long) As Long
'Retorna o número do próximo NumNFiscal disponivel

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalNum_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALNUM_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 45619
    
    
    NFiscalNum_Automatico = X.NFiscalNum_Automatico(sSerie, lNumAuto)
    
    Exit Function
    
Erro_NFiscalNum_Automatico:

    NFiscalNum_Automatico = gErr
    
    Select Case gErr

        Case 45619
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145754)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalRemessa_Verifica_Existencia(objNFiscal As ClassNFiscal, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalRemessa_Verifica_Existencia
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALREMESSA_VERIFICA_EXISTENCIA", X)
    If lErro <> SUCESSO Then gError 45620
    
    
    NFiscalRemessa_Verifica_Existencia = X.NFiscalRemessa_Verifica_Existencia(objNFiscal, lComando)
    
    Exit Function
    
Erro_NFiscalRemessa_Verifica_Existencia:

    NFiscalRemessa_Verifica_Existencia = gErr
    
    Select Case gErr

        Case 45620
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145755)
        
    End Select
    
    Exit Function
    
End Function

Function ItensNFiscalRem_Lock_Gravacao(objNFiscal As ClassNFiscal) As Long
'Faz a validação e os "lock's" dos dados ligados aos ítens da Nota Fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensNFiscalRem_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITENSNFISCALREM_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 45621
    
    
    ItensNFiscalRem_Lock_Gravacao = X.ItensNFiscalRem_Lock_Gravacao(objNFiscal)
    
    Exit Function
    
Erro_ItensNFiscalRem_Lock_Gravacao:

    ItensNFiscalRem_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 45621
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145756)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalRemessa_Lock_Gravacao(objNFiscal As ClassNFiscal, lComando As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalRemessa_Lock_Gravacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALREMESSA_LOCK_GRAVACAO", X)
    If lErro <> SUCESSO Then gError 45622
    
    
    NFiscalRemessa_Lock_Gravacao = X.NFiscalRemessa_Lock_Gravacao(objNFiscal, lComando)
    
    Exit Function
    
Erro_NFiscalRemessa_Lock_Gravacao:

    NFiscalRemessa_Lock_Gravacao = gErr
    
    Select Case gErr

        Case 45622
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145757)
        
    End Select
    
    Exit Function
    
End Function

Function NFiscalRemessa_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscalRemessa_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCALREMESSA_GRAVA", X)
    If lErro <> SUCESSO Then gError 45623
    
    
    NFiscalRemessa_Grava = X.NFiscalRemessa_Grava(objNFiscal, objContabil)
    
    Exit Function
    
Erro_NFiscalRemessa_Grava:

    NFiscalRemessa_Grava = gErr
    
    Select Case gErr

        Case 45623
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145758)
        
    End Select
    
    Exit Function
    
End Function

Function Serie_Exclui(objSerie As ClassSerie) As Long
'Exclui a Serie

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Exclui
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_EXCLUI", X)
    If lErro <> SUCESSO Then gError 45624
    
    
    Serie_Exclui = X.Serie_Exclui(objSerie)
    
    Exit Function
    
Erro_Serie_Exclui:

    Serie_Exclui = gErr
    
    Select Case gErr

        Case 45624
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145759)
        
    End Select
    
    Exit Function
    
End Function

Function Serie_Grava(objSerie As ClassSerie, bPadrao As Boolean) As Long
'Grava a Serie no BD

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_GRAVA", X)
    If lErro <> SUCESSO Then gError 45625
    
    
    Serie_Grava = X.Serie_Grava(objSerie, bPadrao)
    
    Exit Function
    
Erro_Serie_Grava:

    Serie_Grava = gErr
    
    Select Case gErr

        Case 45625
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145760)
        
    End Select
    
    Exit Function
    
End Function

Function TiposDeBloqueio_Le_Todos(colTipoDeBloqueio As Collection) As Long
'Preenche colTipoDeBloqueio com os tipos de Bloqueio existentes na tabela TiposDeBloqueio.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TiposDeBloqueio_Le_Todos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPOSDEBLOQUEIO_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 45626
    
    
    TiposDeBloqueio_Le_Todos = X.TiposDeBloqueio_Le_Todos(colTipoDeBloqueio)
    
    Exit Function
    
Erro_TiposDeBloqueio_Le_Todos:

    TiposDeBloqueio_Le_Todos = gErr
    
    Select Case gErr

        Case 45626
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145761)
        
    End Select
    
    Exit Function
    
End Function

Function FATConfig_Grava_Configs_Trans(colConfigs As ColFATConfig) As Long
'grava na tabela FATConfig uma colecao de registros de configuracao
'TEM QUE SER CHAMADA EM TRANSACAO

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FATConfig_Grava_Configs_Trans
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CONFIG_GRAVA_CONFIGS", X)
    If lErro <> SUCESSO Then gError 41820
    
    
    FATConfig_Grava_Configs_Trans = X.FATConfig_Grava_Configs_Trans(colConfigs)

    Exit Function
    
Erro_FATConfig_Grava_Configs_Trans:
    
    FATConfig_Grava_Configs_Trans = gErr
    
    Select Case gErr

        Case 41820
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145762)
        
    End Select
    
    Exit Function

End Function

Function FAT_Instalacao_Filial(iFilialEmpresa As Integer) As Long
'faz as inicializacoes necessarias à criacao de uma nova filial especificas do modulo

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FAT_Instalacao_Filial
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_INSTALACAO_FILIAL", X)
    If lErro <> SUCESSO Then gError 41937
    
    
    FAT_Instalacao_Filial = X.FAT_Instalacao_Filial(iFilialEmpresa)

    Exit Function
    
Erro_FAT_Instalacao_Filial:
    
    FAT_Instalacao_Filial = gErr
    
    Select Case gErr

        Case 41937
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145763)
        
    End Select
    
    Exit Function

End Function

Function FATConfig_Grava_Configs1(objFAT As ClassFAT) As Long
'grava nas tabelas FATConfig e CRFATConfig uma colecao de registros de configuracao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FATConfig_Grava_Configs1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CONFIG_GRAVA_CONFIGS1", X)
    If lErro <> SUCESSO Then gError 55011
    
    
    FATConfig_Grava_Configs1 = X.FATConfig_Grava_Configs1(objFAT)

    Exit Function
    
Erro_FATConfig_Grava_Configs1:
    
    FATConfig_Grava_Configs1 = gErr
    
    Select Case gErr

        Case 55011
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145764)
        
    End Select
    
    Exit Function

End Function

Function FilialCliente_Grava_Estatistica(objNFiscal As ClassNFiscal) As Long
'Tem que ser chamado dentro de Transação
'Atualiza as estatísticas do Cliente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialCliente_Grava_Estatistica
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_FILIALCLIENTE_GRAVA_ESTATISTICA", X)
    If lErro <> SUCESSO Then gError 54672
    
    
    FilialCliente_Grava_Estatistica = X.FilialCliente_Grava_Estatistica(objNFiscal)

    Exit Function
    
Erro_FilialCliente_Grava_Estatistica:
    
    FilialCliente_Grava_Estatistica = gErr
    
    Select Case gErr

        Case 54672
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145765)
        
    End Select
    
    Exit Function

End Function

Function FilCliFilEmp_Le_Lock(lComando As Long, objFilCliFilEmp As ClassFilCliFilEmp) As Long
'le e lock a tabela FilCliFilEmp
'Chaves iFilialEmpresa, lCodCliente, iCodFilial do objFilCliFilEmp

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilCliFilEmp_Le_Lock
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_FILCLIFILEMP_LE_LOCK", X)
    If lErro <> SUCESSO Then gError 54673
    
    
    FilCliFilEmp_Le_Lock = X.FilCliFilEmp_Le_Lock(lComando, objFilCliFilEmp)

    Exit Function
    
Erro_FilCliFilEmp_Le_Lock:
    
    FilCliFilEmp_Le_Lock = gErr
    
    Select Case gErr

        Case 54673
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145766)
        
    End Select
    
    Exit Function

End Function

Function ProdutoVendido_Grava_Estatisticas(objNFiscal As ClassNFiscal, iCancelamento As Integer) As Long
'Grava a estatística dos produtos vendidos

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ProdutoVendido_Grava_Estatisticas
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PRODUTOVENDIDO_GRAVA_ESTATISTICAS", X)
    If lErro <> SUCESSO Then gError 54674
    
    ProdutoVendido_Grava_Estatisticas = X.ProdutoVendido_Grava_Estatisticas(objNFiscal, iCancelamento)

    Exit Function
    
Erro_ProdutoVendido_Grava_Estatisticas:
    
    ProdutoVendido_Grava_Estatisticas = gErr
    
    Select Case gErr

        Case 54674
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145767)
        
    End Select
    
    Exit Function

End Function

Function GeracaoFatura_Automatico(lNumInt1 As Long) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoFatura_Automatico
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_GERACAOFATURA_AUTOMATICO", X)
    If lErro <> SUCESSO Then gError 54682
    
    
    GeracaoFatura_Automatico = X.GeracaoFatura_Automatico(lNumInt1)

    Exit Function
    
Erro_GeracaoFatura_Automatico:
    
    GeracaoFatura_Automatico = gErr
    
    Select Case gErr

        Case 54682
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145768)
        
    End Select
    
    Exit Function

End Function

Function GeracaoFatura_GerarFatura(objGeracaoFatura As ClassGeracaoFatura, colNFiscalMarcado As Collection, colcolComissao As colcolComissao, colcolDesconto As colcolDesconto, objContabil As ClassContabil) As Long
'faz todas as atualizacoes no bd necessarias ao registro da fatura que cobra um grupo de NFs

Dim X As Object, lErro As Long
    
On Error GoTo Erro_GeracaoFatura_GerarFatura
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_GERACAOFATURA_GERARFATURA", X)
    If lErro <> SUCESSO Then gError 54683
    
    
    GeracaoFatura_GerarFatura = X.GeracaoFatura_GerarFatura(objGeracaoFatura, colNFiscalMarcado, colcolComissao, colcolDesconto, objContabil)

    Exit Function
    
Erro_GeracaoFatura_GerarFatura:
    
    GeracaoFatura_GerarFatura = gErr
    
    Select Case gErr

        Case 54683
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145769)
        
    End Select
    
    Exit Function

End Function

Function RelFatPrazoPag_Grava(lCodigo As Long, ColPrazos As Collection) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_RelFatPrazoPag_Grava
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_RELFATPRAZOPAG_GRAVA", X)
    If lErro <> SUCESSO Then gError 54688
    
    
    RelFatPrazoPag_Grava = X.RelFatPrazoPag_Grava(lCodigo, ColPrazos)

    Exit Function
    
Erro_RelFatPrazoPag_Grava:
    
    RelFatPrazoPag_Grava = gErr
    
    Select Case gErr

        Case 54688
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145770)
        
    End Select
    
    Exit Function

End Function

Function FilialFaturamento_Le(iFilialEmpresa As Integer, iFilialFaturamento As Integer) As Long
'Lê o código da filial de faturamento pada a filialempresa passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_FilialFaturamento_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_FILIALFATURAMENTO_LE", X)
    If lErro <> SUCESSO Then gError 54731
    
    
    FilialFaturamento_Le = X.FilialFaturamento_Le(iFilialEmpresa, iFilialFaturamento)

    Exit Function
    
Erro_FilialFaturamento_Le:
    
    FilialFaturamento_Le = gErr
    
    Select Case gErr

        Case 54731
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145771)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Le_NumeroSerie(objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal com o Número, Série e Filial passados em objNFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_NumeroSerie
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_NUMEROSERIE", X)
    If lErro <> SUCESSO Then gError 54732
    
    
    NFiscal_Le_NumeroSerie = X.NFiscal_Le_NumeroSerie(objNFiscal)

    Exit Function
    
Erro_NFiscal_Le_NumeroSerie:
    
    NFiscal_Le_NumeroSerie = gErr
    
    Select Case gErr

        Case 54732
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145772)
        
    End Select
    
    Exit Function

End Function

Function MovEstoqueItensBenef_Le(objNFiscal As ClassNFiscal) As Long
'Le os Ítens do MovimentodeEstoque para o Beneficiamento

Dim X As Object, lErro As Long
    
On Error GoTo Erro_MovEstoqueItensBenef_Le
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_MOVESTOQUEITENSBENEF_LE", X)
    If lErro <> SUCESSO Then gError 54733
    
    
    MovEstoqueItensBenef_Le = X.MovEstoqueItensBenef_Le(objNFiscal)

    Exit Function
    
Erro_MovEstoqueItensBenef_Le:
    
    MovEstoqueItensBenef_Le = gErr
    
    Select Case gErr

        Case 54733
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145773)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le_Tributacao(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_Tributacao
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_TRIBUTACAO", X)
    If lErro <> SUCESSO Then gError 54909
    
    
    PedidoDeVenda_Le_Tributacao = X.PedidoDeVenda_Le_Tributacao(objPedidoVenda)

    Exit Function
    
Erro_PedidoDeVenda_Le_Tributacao:
    
    PedidoDeVenda_Le_Tributacao = gErr
    
    Select Case gErr

        Case 54909
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145774)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedidos_ObterPedidos(objBaixaPedido As ClassBaixaPedidos) As Long
'Preenche a colecao colPVInfo em objBaixaPedido com os pedidos que podem ser faturados e atendem
'aos criterios passados em objBaixaPedido com restricoes de No. de pedido, data de emissao,...

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedidos_ObterPedidos
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BAIXAPEDIDOS_OBTERPEDIDOS", X)
    If lErro <> SUCESSO Then gError 59540
    
    
    BaixaPedidos_ObterPedidos = X.BaixaPedidos_ObterPedidos(objBaixaPedido)

    Exit Function
    
Erro_BaixaPedidos_ObterPedidos:
    
    BaixaPedidos_ObterPedidos = gErr
    
    Select Case gErr

        Case 59540
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145775)
        
    End Select
    
    Exit Function

End Function

Sub BaixaPedidos_ObterPedidos1(objBaixaPedido As ClassBaixaPedidos, sSelect As String)
'Montar SELECT dinamicamente

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedidos_ObterPedidos1
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BAIXAPEDIDOS_OBTERPEDIDOS1", X)
    If lErro <> SUCESSO Then gError 59541
    
    
    Call X.BaixaPedidos_ObterPedidos1(objBaixaPedido, sSelect)
    
    Exit Sub
    
Erro_BaixaPedidos_ObterPedidos1:
    
    Select Case gErr

        Case 59541
            
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145776)
        
    End Select
    
    Exit Sub

End Sub

Function BaixaPedidos_ObterPedidos2(lComando As Long, sSelect As String, vItensPedidoDeVendaDataEntrega As Variant, vClientesNomeReduzido As Variant, vPedidosDeVendaCodigo As Variant, vPedidosDeVendaCliente As Variant, vPedidosDeVendaFilial As Variant, vPedidosDeVendaDataEmissao As Variant, vPedidosDeVendaValorTotal As Variant) As Long
'Preparar parte fixa do SELECT

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedidos_ObterPedidos2
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BAIXAPEDIDOS_OBTERPEDIDOS2", X)
    If lErro <> SUCESSO Then gError 59542
    
    
    BaixaPedidos_ObterPedidos2 = X.BaixaPedidos_ObterPedidos2(lComando, sSelect, vItensPedidoDeVendaDataEntrega, vClientesNomeReduzido, vPedidosDeVendaCodigo, vPedidosDeVendaCliente, vPedidosDeVendaFilial, vPedidosDeVendaDataEmissao, vPedidosDeVendaValorTotal)

    Exit Function
    
Erro_BaixaPedidos_ObterPedidos2:
    
    BaixaPedidos_ObterPedidos2 = gErr
    
    Select Case gErr

        Case 59542
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145777)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedidos_ObterPedidos3(lComando As Long, objBaixaPedido As ClassBaixaPedidos) As Long
'Filtra por giFilialEmpresa
'Complementa a passagem dos parametros que variam de acordo com a selecao do usuario

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedidos_ObterPedidos3
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BAIXAPEDIDOS_OBTERPEDIDOS3", X)
    If lErro <> SUCESSO Then gError 59543
    
    
    BaixaPedidos_ObterPedidos3 = X.BaixaPedidos_ObterPedidos3(lComando, objBaixaPedido)

    Exit Function
    
Erro_BaixaPedidos_ObterPedidos3:
    
    BaixaPedidos_ObterPedidos3 = gErr
    
    Select Case gErr

        Case 59543
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145778)
        
    End Select
    
    Exit Function

End Function

Function BaixaPedidos_ObterPedidos4(lComando As Long, objBaixaPedido As ClassBaixaPedidos, dtItensPedidoDeVendaDataEntrega As Date, sClientesNomeReduzido As String, lPedidosDeVendaCodigo As Long, lPedidosDeVendaCliente As Long, iPedidosDeVendaFilial As Integer, dtPedidosDeVendaDataEmissao As Date, dPedidosDeVendaValorTotal As Double) As Long
'Processar os registros lidos preenchendo a colecao

Dim X As Object, lErro As Long
    
On Error GoTo Erro_BaixaPedidos_ObterPedidos4
    
    
    
    
    lErro = Rotina_ObterObjetoFuncao("FAT_BAIXAPEDIDOS_OBTERPEDIDOS4", X)
    If lErro <> SUCESSO Then gError 59544
    
    
    BaixaPedidos_ObterPedidos4 = X.BaixaPedidos_ObterPedidos4(lComando, objBaixaPedido, dtItensPedidoDeVendaDataEntrega, sClientesNomeReduzido, lPedidosDeVendaCodigo, lPedidosDeVendaCliente, iPedidosDeVendaFilial, dtPedidosDeVendaDataEmissao, dPedidosDeVendaValorTotal)

    Exit Function
    
Erro_BaixaPedidos_ObterPedidos4:
    
    BaixaPedidos_ObterPedidos4 = gErr
    
    Select Case gErr

        Case 59544
            
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145779)
        
    End Select
    
    Exit Function

End Function

Function TabelaPrecoPadrao_Le(objProdutoFilial As ClassProdutoFilial, dtDataFinal As Date, dPrecoTabela As Double) As Long
'Lê último preço da tabela padrão vigente para par produto-filialEmpresa até a DataFinal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TabelaPrecoPadrao_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TABELAPRECOPADRAO_LE", X)
    If lErro <> SUCESSO Then gError 59865
    
    TabelaPrecoPadrao_Le = X.TabelaPrecoPadrao_Le(objProdutoFilial, dtDataFinal, dPrecoTabela)

    Exit Function
    
Erro_TabelaPrecoPadrao_Le:
    
    TabelaPrecoPadrao_Le = gErr
    
    Select Case gErr

        Case 59865
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145780)
        
    End Select
    
    Exit Function

End Function

Function LiberacoesCredito_Filial_Le(colUsuarios As Collection, colLiberacaoCredito As Collection) As Long
'Guarda em colLiberacaoCredito todas as alçadas cadastradas
'Para os Usuarios passados na coleção de Usuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_LiberacoesCredito_Filial_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_LIBERACOESCREDITO_FILIAL_LE", X)
    If lErro <> SUCESSO Then gError 59873
    
    LiberacoesCredito_Filial_Le = X.LiberacoesCredito_Filial_Le(colUsuarios, colLiberacaoCredito)

    Exit Function
    
Erro_LiberacoesCredito_Filial_Le:
    
    LiberacoesCredito_Filial_Le = gErr
    
    Select Case gErr

        Case 59873
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145781)
        
    End Select
    
    Exit Function

End Function

Function UsuariosFilialEmpresa_Le_Todos(colUsuarios As Collection) As Long
'Le todos os usuários desta Filial Empresa na tabela de UsuarioFilialEmpresa e coloca os dados em colUsuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_UsuariosFilialEmpresa_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("FAT_USUARIOSFILIALEMPRESA_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 59874
    
    UsuariosFilialEmpresa_Le_Todos = X.UsuariosFilialEmpresa_Le_Todos(colUsuarios)

    Exit Function
    
Erro_UsuariosFilialEmpresa_Le_Todos:
    
    UsuariosFilialEmpresa_Le_Todos = gErr
    
    Select Case gErr

        Case 59874
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145782)
        
    End Select
    
    Exit Function

End Function

Function Usuarios_Com_LiberacaoCredito_Le(colTodosUsuarios As Collection, dValorLiberacao As Double, colUsuariosComLiberacao As Collection) As Long
'Monta uma colecao de Usuarios que tenham Liberacao de Crédito por Operacao maior que o passado
'E que nao tenham ultrapassados o LImite de Credito Mensal
'Para os Usuarios passados na coleção de Usuarios

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Usuarios_Com_LiberacaoCredito_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_USUARIOS_COM_LIBERACAOCREDITO_LE", X)
    If lErro <> SUCESSO Then gError 59875
    
    Usuarios_Com_LiberacaoCredito_Le = X.Usuarios_Com_LiberacaoCredito_Le(colTodosUsuarios, dValorLiberacao, colUsuariosComLiberacao)

    Exit Function
    
Erro_Usuarios_Com_LiberacaoCredito_Le:
    
    Usuarios_Com_LiberacaoCredito_Le = gErr
    
    Select Case gErr

        Case 59875
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145783)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscalPedido_Le_ItemPV(lCodigoPedido As Long, colItemPedidoNF As Collection) As Long
'Lê todas as Notas Fiscais com relação ao Pedido passado e Ordena por ItemPedido

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalPedido_Le_ItemPV
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NOTAFISCALPEDIDO_LE_ITEMPV", X)
    If lErro <> SUCESSO Then gError 59876
    
    NotaFiscalPedido_Le_ItemPV = X.NotaFiscalPedido_Le_ItemPV(lCodigoPedido, colItemPedidoNF)

    Exit Function
    
Erro_NotaFiscalPedido_Le_ItemPV:
    
    NotaFiscalPedido_Le_ItemPV = gErr
    
    Select Case gErr

        Case 59876
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145784)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscalPedido_Le_Nota(lCodigoPedido As Long, colItemPedidoNF As Collection) As Long
'Lê todas as Notas Fiscais com relação ao Pedido passado e Ordena por NotaFiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalPedido_Le_Nota
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NOTAFISCALPEDIDO_LE_NOTA", X)
    If lErro <> SUCESSO Then gError 59877
    
    NotaFiscalPedido_Le_Nota = X.NotaFiscalPedido_Le_Nota(lCodigoPedido, colItemPedidoNF)

    Exit Function
    
Erro_NotaFiscalPedido_Le_Nota:
    
    NotaFiscalPedido_Le_Nota = gErr
    
    Select Case gErr

        Case 59877
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145785)
        
    End Select
    
    Exit Function

End Function

Function ItensPedido_Le_Inclusive_Baixados(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê no BD a coleção de ítens do Pedido de Venda e devolve em colItensPedido dentro de objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ItensPedido_Le_Inclusive_Baixados
    
    lErro = Rotina_ObterObjetoFuncao("FAT_ITENSPEDIDO_LE_INCLUSIVE_BAIXADOS", X)
    If lErro <> SUCESSO Then gError 59878
    
    ItensPedido_Le_Inclusive_Baixados = X.ItensPedido_Le_Inclusive_Baixados(objPedidoVenda)

    Exit Function
    
Erro_ItensPedido_Le_Inclusive_Baixados:
    
    ItensPedido_Le_Inclusive_Baixados = gErr
    
    Select Case gErr

        Case 59878
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145786)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Le_Pedido_TipoDocInfo(objNFiscal As ClassNFiscal, sTela As String) As Long
'Lê a Nota Fiscal com o Número, Serie, DataEmissao, NumPedidoVenda e FilialPedido e Retorna a Tela pelo Tipo doc Info

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_Pedido_TipoDocInfo
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_PEDIDO_TIPODOCINFO", X)
    If lErro <> SUCESSO Then gError 59882
    
    NFiscal_Le_Pedido_TipoDocInfo = X.NFiscal_Le_Pedido_TipoDocInfo(objNFiscal, sTela)

    Exit Function
    
Erro_NFiscal_Le_Pedido_TipoDocInfo:
    
    NFiscal_Le_Pedido_TipoDocInfo = gErr
    
    Select Case gErr

        Case 59882
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145787)
        
    End Select
    
    Exit Function

End Function

Function ParcelasPV_Le(objPedidoVenda As ClassPedidoDeVenda) As Long
'Lê as parcelas associadas ao objPedidoVenda

Dim X As Object, lErro As Long
    
On Error GoTo Erro_ParcelasPV_Le
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PARCELASPV_LE", X)
    If lErro <> SUCESSO Then gError 59883
    
    ParcelasPV_Le = X.ParcelasPV_Le(objPedidoVenda)

    Exit Function
    
Erro_ParcelasPV_Le:
    
    ParcelasPV_Le = gErr
    
    Select Case gErr

        Case 59883
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145788)
        
    End Select
    
    Exit Function

End Function

Function PedidoDeVenda_Le_Todos_Completo(objPedidoDeVenda As ClassPedidoDeVenda, iStatus As Integer) As Long
'Le o Pedido Inclusive se ele estiver Baixado
'????? Trocar algumas coisas quando for implementado a Tributacao Baixado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_PedidoDeVenda_Le_Todos_Completo
    
    lErro = Rotina_ObterObjetoFuncao("FAT_PEDIDODEVENDA_LE_TODOS_COMPLETO", X)
    If lErro <> SUCESSO Then gError 59884
    
    PedidoDeVenda_Le_Todos_Completo = X.PedidoDeVenda_Le_Todos_Completo(objPedidoDeVenda, iStatus)

    Exit Function
    
Erro_PedidoDeVenda_Le_Todos_Completo:
    
    PedidoDeVenda_Le_Todos_Completo = gErr
    
    Select Case gErr

        Case 59884
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145789)
        
    End Select
    
    Exit Function

End Function

Function TipoDocInfo_Le_NomeTela_NFiscal(objNFiscal As ClassNFiscal, sTela As String) As Long
'Lê o Nome da Tela Nota Fiscal com a Série e o Número passado em objNFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_TipoDocInfo_Le_NomeTela_NFiscal
    
    lErro = Rotina_ObterObjetoFuncao("FAT_TIPODOCINFO_LE_NOMETELA_NFISCAL", X)
    If lErro <> SUCESSO Then gError 59885
    
    TipoDocInfo_Le_NomeTela_NFiscal = X.TipoDocInfo_Le_NomeTela_NFiscal(objNFiscal, sTela)

    Exit Function
    
Erro_TipoDocInfo_Le_NomeTela_NFiscal:
    
    TipoDocInfo_Le_NomeTela_NFiscal = gErr
    
    Select Case gErr

        Case 59885
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145790)
        
    End Select
    
    Exit Function

End Function

Function NF_NFFatura_Le_NumeroSerie(objNFiscal As ClassNFiscal) As Long
'Lê a Nota Fiscal com o Número, Série e Filial passados em objNFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NF_NFFatura_Le_NumeroSerie
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NF_NFFATURA_LE_NUMEROSERIE", X)
    If lErro <> SUCESSO Then gError 59886
    
    NF_NFFatura_Le_NumeroSerie = X.NF_NFFatura_Le_NumeroSerie(objNFiscal)

    Exit Function
    
Erro_NF_NFFatura_Le_NumeroSerie:
    
    NF_NFFatura_Le_NumeroSerie = gErr
    
    Select Case gErr

        Case 59886
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145791)
        
    End Select
    
    Exit Function

End Function

Function Serie_Le_Padrao(sSeriePadrao As String) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Le_Padrao
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_LE_PADRAO", X)
    If lErro <> SUCESSO Then gError 64116
    
    Serie_Le_Padrao = X.Serie_Le_Padrao(sSeriePadrao)

    Exit Function
    
Erro_Serie_Le_Padrao:
    
    Serie_Le_Padrao = gErr
    
    Select Case gErr

        Case 64116
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145792)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscalSaida_Cancelar(objNFiscal As ClassNFiscal) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalSaida_Cancelar
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NOTAFISCALSAIDA_CANCELAR", X)
    If lErro <> SUCESSO Then gError 65101
    
    NotaFiscalSaida_Cancelar = X.NotaFiscalSaida_Cancelar(objNFiscal)

    Exit Function
    
Erro_NotaFiscalSaida_Cancelar:
    
    NotaFiscalSaida_Cancelar = gErr
    
    Select Case gErr

        Case 65101
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145793)
        
    End Select
    
    Exit Function

End Function

Function CancelamentoNFiscal_Atualiza_Pedido(objNFiscal As ClassNFiscal) As Long
'Atualiza o Pedido de Venda após o cancelamento da NF passada em objNFiscal.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CancelamentoNFiscal_Atualiza_Pedido
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CANCELAMENTONFISCAL_ATUALIZA_PEDIDO", X)
    If lErro <> SUCESSO Then gError 65102
    
    CancelamentoNFiscal_Atualiza_Pedido = X.CancelamentoNFiscal_Atualiza_Pedido(objNFiscal)

    Exit Function
    
Erro_CancelamentoNFiscal_Atualiza_Pedido:
    
    CancelamentoNFiscal_Atualiza_Pedido = gErr
    
    Select Case gErr

        Case 65102
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145794)
        
    End Select
    
    Exit Function

End Function

Function Consultas_Le_Todos(sNomeTela As String, colConsultas As Collection) As Long
'Le Todas as Consultas para a Tela Passada

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Consultas_Le_Todos
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CONSULTAS_LE_TODOS", X)
    If lErro <> SUCESSO Then gError 65103
    
    Consultas_Le_Todos = X.Consultas_Le_Todos(sNomeTela, colConsultas)

    Exit Function
    
Erro_Consultas_Le_Todos:
    
    Consultas_Le_Todos = gErr
    
    Select Case gErr

        Case 65103
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145795)
        
    End Select
    
    Exit Function

End Function

Function CliFilEmp_Le_Rel(objFilCliFilEmp As ClassFilCliFilEmp, lComando As Long) As Long
'Le em FilialClienteFilEmp dados estatísticos de compras do Cliente e FilialCliente para Relatório.
'IMPORTANTE : O lComando somente é aberto se vier zerado
'IMPORTANTE : O lComando não é fechado, só é fechado em caso de Erro

Dim X As Object, lErro As Long
    
On Error GoTo Erro_CliFilEmp_Le_Rel
    
    lErro = Rotina_ObterObjetoFuncao("FAT_CLIFILEMP_LE_REL", X)
    If lErro <> SUCESSO Then gError 65109
    
    CliFilEmp_Le_Rel = X.CliFilEmp_Le_Rel(objFilCliFilEmp, lComando)

    Exit Function
    
Erro_CliFilEmp_Le_Rel:
    
    CliFilEmp_Le_Rel = gErr
    
    Select Case gErr

        Case 65109
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145796)
        
    End Select
    
    Exit Function

End Function

Function Serie_Unlock_ImpressaoNF(objSerie As ClassSerie) As Long
'Faz Unlock no Relatorio de Notas Fiscais

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Unlock_ImpressaoNF
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_UNLOCK_IMPRESSAONF", X)
    If lErro <> SUCESSO Then gError 65110
    
    Serie_Unlock_ImpressaoNF = X.Serie_Unlock_ImpressaoNF(objSerie)

    Exit Function
    
Erro_Serie_Unlock_ImpressaoNF:
    
    Serie_Unlock_ImpressaoNF = gErr
    
    Select Case gErr

        Case 65110
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145797)
        
    End Select
    
    Exit Function

End Function

Function Serie_Unlock_Atualiza_ImpressaoNF(objSerie As ClassSerie) As Long
'Faz Unlock no Relatório de NF e Atualiza o Número da última Nota impressa.

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Unlock_Atualiza_ImpressaoNF
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_UNLOCK_ATUALIZA_IMPRESSAONF", X)
    If lErro <> SUCESSO Then gError 65111
    
    Serie_Unlock_Atualiza_ImpressaoNF = X.Serie_Unlock_Atualiza_ImpressaoNF(objSerie)

    Exit Function
    
Erro_Serie_Unlock_Atualiza_ImpressaoNF:
    
    Serie_Unlock_Atualiza_ImpressaoNF = gErr
    
    Select Case gErr

        Case 65111
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145798)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_Le_UltimaCadastrada(objNFiscal As ClassNFiscal) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Le_UltimaCadastrada
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_LE_ULTIMACADASTRADA", X)
    If lErro <> SUCESSO Then gError 65112
    
    NFiscal_Le_UltimaCadastrada = X.NFiscal_Le_UltimaCadastrada(objNFiscal)

    Exit Function
    
Erro_NFiscal_Le_UltimaCadastrada:
    
    NFiscal_Le_UltimaCadastrada = gErr
    
    Select Case gErr

        Case 65112
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145799)
        
    End Select
    
    Exit Function

End Function

Function Serie_Lock_ImpressaoNFiscal(objSerie As ClassSerie) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Lock_ImpressaoNFiscal
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_LOCK_IMPRESSAONFISCAL", X)
    If lErro <> SUCESSO Then gError 65113
    
    Serie_Lock_ImpressaoNFiscal = X.Serie_Lock_ImpressaoNFiscal(objSerie)

    Exit Function
    
Erro_Serie_Lock_ImpressaoNFiscal:
    
    Serie_Lock_ImpressaoNFiscal = gErr
    
    Select Case gErr

        Case 65113
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145800)
        
    End Select
    
    Exit Function

End Function

Function Serie_Altera_Imprimindo(objSerie As ClassSerie) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Altera_Imprimindo
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_ALTERA_IMPRIMINDO", X)
    If lErro <> SUCESSO Then gError 65114
    
    Serie_Altera_Imprimindo = X.Serie_Altera_Imprimindo(objSerie)

    Exit Function
    
Erro_Serie_Altera_Imprimindo:
    
    Serie_Altera_Imprimindo = gErr
    
    Select Case gErr

        Case 65114
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145801)
        
    End Select
    
    Exit Function

End Function

Function Serie_Altera_Nao_Imprimindo(objSerie As ClassSerie) As Long

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Serie_Altera_Nao_Imprimindo
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIE_ALTERA_NAO_IMPRIMINDO", X)
    If lErro <> SUCESSO Then gError 65115
    
    Serie_Altera_Nao_Imprimindo = X.Serie_Altera_Nao_Imprimindo(objSerie)

    Exit Function
    
Erro_Serie_Altera_Nao_Imprimindo:
    
    Serie_Altera_Nao_Imprimindo = gErr
    
    Select Case gErr

        Case 65115
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145802)
        
    End Select
    
    Exit Function

End Function

Function Series_Le_TipoFormulario(colSerie As colSerie, iTipoFormulario As Integer) As Long
'Lê todos as Séries que tenham o mesmo tipoFormulario passado contidas na tabela Serie e as coloca na coleçao colSerie

Dim X As Object, lErro As Long
    
On Error GoTo Erro_Series_Le_TipoFormulario
    
    lErro = Rotina_ObterObjetoFuncao("FAT_SERIES_LE_TIPOFORMULARIO", X)
    If lErro <> SUCESSO Then gError 65116
    
    Series_Le_TipoFormulario = X.Series_Le_TipoFormulario(colSerie, iTipoFormulario)

    Exit Function
    
Erro_Series_Le_TipoFormulario:
    
    Series_Le_TipoFormulario = gErr
    
    Select Case gErr

        Case 65116
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145803)
        
    End Select
    
    Exit Function

End Function

Function NFiscal_ObterDebCredICMS(lNumIntNF As Long, dValorBase As Double, dValor As Double, dValorBaseSubst As Double, dValorSubst As Double, lComandoTribNF As Long, lComandoTribItemNF As Long) As Long
'obtem informacoes sobre valores de credito ou debito de icms de uma nota fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_ObterDebCredICMS
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_OBTERDEBCREDICMS", X)
    If lErro <> SUCESSO Then gError 65475
    
    NFiscal_ObterDebCredICMS = X.NFiscal_ObterDebCredICMS(lNumIntNF, dValorBase, dValor, dValorBaseSubst, dValorSubst, lComandoTribNF, lComandoTribItemNF)

    Exit Function
    
Erro_NFiscal_ObterDebCredICMS:
    
    NFiscal_ObterDebCredICMS = gErr
    
    Select Case gErr

        Case 65475
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145804)
        
    End Select
    
    Exit Function

End Function

Function NF_Obter_NumeroItens(lNumIntDoc As Long, iNumItens As Integer, lComando As Long) As Long
'OBS: Função para o cálculo de Número de Itens de Nota Fiscal para Relatorio
'Se o Comando vier zerado vai ser aberto e não será fechado

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NF_Obter_NumeroItens
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NF_OBTER_NUMEROITENS", X)
    If lErro <> SUCESSO Then gError 65476
    
    NF_Obter_NumeroItens = X.NF_Obter_NumeroItens(lNumIntDoc, iNumItens, lComando)

    Exit Function
    
Erro_NF_Obter_NumeroItens:
    
    NF_Obter_NumeroItens = gErr
    
    Select Case gErr

        Case 65476
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145805)
        
    End Select
    
    Exit Function

End Function

Function NotaFiscalSaida_Excluir(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'Faz o cancelamento de uma nota fiscal de Saida

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NotaFiscalSaida_Excluir
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NOTAFISCALSAIDA_EXCLUIR", X)
    If lErro <> SUCESSO Then gError 83936
    
    NotaFiscalSaida_Excluir = X.NotaFiscalSaida_Excluir(objNFiscal, objContabil)

    Exit Function
    
Erro_NotaFiscalSaida_Excluir:
    
    NotaFiscalSaida_Excluir = gErr
    
    Select Case gErr

        Case 83936
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145806)
        
    End Select
    
    Exit Function

End Function

'Function NFiscal_Altera_MovEstoque(objNFiscal As ClassNFiscal) As Long
''trata da alteração de data dos movimentos de estoque associados à nota fiscal
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_NFiscal_Altera_MovEstoque
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_ALTERA_MOVESTOQUE", X)
'    If lErro <> SUCESSO Then gError 89790
'
'    NFiscal_Altera_MovEstoque = X.NFiscal_Altera_MovEstoque(objNFiscal)
'
'    Exit Function
'
'Erro_NFiscal_Altera_MovEstoque:
'
'    NFiscal_Altera_MovEstoque = gErr
'
'    Select Case gErr
'
'        Case 89790
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145807)
'
'    End Select
'
'    Exit Function
'
'End Function

'Function ProdutoVendido_Grava_EstatisticasDev(objNFiscal As ClassNFiscal, iCancelamento As Integer) As Long
''Grava a estatística dos produtos devolvidos
'
'Dim X As Object, lErro As Long
'
'On Error GoTo Erro_ProdutoVendido_Grava_EstatisticasDev
'
'    lErro = Rotina_ObterObjetoFuncao("FAT_PRODUTOVENDIDO_GRAVA_ESTATISTICASDEV", X)
'    If lErro <> SUCESSO Then Error 54669
'
'    ProdutoVendido_Grava_EstatisticasDev = X.ProdutoVendido_Grava_EstatisticasDev(objNFiscal, iCancelamento)
'
'    Exit Function
'
'Erro_ProdutoVendido_Grava_EstatisticasDev:
'
'    ProdutoVendido_Grava_EstatisticasDev = Err
'
'    Select Case Err
'
'        Case 54669
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 145808)
'
'    End Select
'
'    Exit Function
'
'End Function

Function NFiscal_Alteracao(objNFiscal As ClassNFiscal) As Long
'trata da alteração de data dos movimentos de estoque associados à nota fiscal

Dim X As Object, lErro As Long
    
On Error GoTo Erro_NFiscal_Alteracao
    
    lErro = Rotina_ObterObjetoFuncao("FAT_NFISCAL_ALTERACAO", X)
    If lErro <> SUCESSO Then gError 89791
    
    NFiscal_Alteracao = X.NFiscal_Alteracao(objNFiscal)

    Exit Function
    
Erro_NFiscal_Alteracao:
    
    NFiscal_Alteracao = gErr
    
    Select Case gErr

        Case 89791
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 145809)
        
    End Select
    
    Exit Function

End Function

