Attribute VB_Name = "GlobalFATArtmill"
Option Explicit

Public Const FATOR_PROPORCAO_100 = 1
Public Const FATOR_PROPORCAO_75 = 0.75
Public Const FATOR_PROPORCAO_50 = 0.5
Public Const FATOR_PROPORCAO_25 = 0.25
Public Const FATOR_PROPORCAO_0 = 0

Public Const DELTA_FILIALREAL_OFICIAL = 50

Function NFiscal_Grava_Clone(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal sNomeFuncGravacao As String, lNumNFOficial As Long) As Long

Dim lErro As Long
Dim objNFiscalOficial As New ClassNFiscal
Dim bNFInterna As Boolean, bClonar As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer
Dim dFatorValor As Double
Dim objSerie As New ClassSerie

On Error GoTo Erro_NFiscal_Grava_Clone

    lNumNFOficial = 0
    
    'Se nf nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa) <> objNFiscal.iFilialEmpresa Then
    
        bClonar = True
        
        objNFiscalOficial.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
        
        Set objNFiscalOficial.objMovEstoque = New ClassMovEstoque
        
        objNFiscalOficial.sSerie = objNFiscal.sSerie
            
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
            
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 500033
        
        'Se não encontrou o Tipo de Documento --> erro
        If lErro <> SUCESSO Then gError 500034
    
        bNFInterna = (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS)
        
        If bNFInterna Then
        
            If objNFiscal.lNumNotaFiscal = 0 Then
            
                'verificar se a serie existe na filial oficial, lockar a serie e pegar o proximo numero de nf
                lErro = NFiscalNumAuto2(objNFiscalOficial)
                If lErro <> SUCESSO And lErro <> 32285 Then gError 60443
            
                'se nao existir a série na filial oficial
                If lErro <> SUCESSO Then bClonar = False
            
            Else
            
                objSerie.iFilialEmpresa = objNFiscalOficial.iFilialEmpresa
                objSerie.sSerie = objNFiscalOficial.sSerie
            
                lErro = CF("Serie_Le", objSerie)
                If lErro <> SUCESSO And lErro <> 22202 Then gError 207650
                
                'se a serie nao existir na filial oficial ==> nao clonar
                If lErro <> SUCESSO Then bClonar = False
        
            
            
            End If
            
                        
            
        Else 'se for nf externa
        
'            lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo, True)
'            If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 500035
'
'            'Se for uma nota nova
'            If lErro = SUCESSO Then bClonar = True
        
            bClonar = False
        
        
        End If
        
        If bClonar Then
        
            Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
        
            colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, objNFiscalOficial.iFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
            
            'Lê as configurações em ESTConfig
            lErro = CF("ESTConfig_Le_Configs", colConfig)
            If lErro <> SUCESSO Then gError 126846
            
            iAceitaEstoqueNegativo = gobjMAT.iAceitaEstoqueNegativo
            
            gobjMAT.iAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
        
            'clonar o objeto nfiscal
            lErro = NFiscal_Clonar(objNFiscal, objNFiscalOficial)
            If lErro <> SUCESSO Then gError 500002
            
            'chamar a funcao de gravacao para o clone SEM CTB
            lErro = CF(sNomeFuncGravacao, objNFiscalOficial, Nothing)
            If lErro <> SUCESSO Then gError 500003
            
            lNumNFOficial = objNFiscalOficial.lNumNotaFiscal
        
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
            
        End If
    
    End If
    
    NFiscal_Grava_Clone = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Grava_Clone:

    NFiscal_Grava_Clone = gErr
     
    Select Case gErr
          
        Case 126846, 500033, 500035
          
        Case 500002, 500003
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
            
        Case 500034
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161625)
     
    End Select
     
    Exit Function

End Function

Function FilialEmpresa_ConvFRFO(ByVal iFilialEmpresa As Integer) As Integer

    If iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL Then
        FilialEmpresa_ConvFRFO = iFilialEmpresa
    Else
        FilialEmpresa_ConvFRFO = iFilialEmpresa - DELTA_FILIALREAL_OFICIAL
    End If
    
End Function

Function Almoxarifado_ConvFRFO(ByVal iAlmoxarifadoFilialReal As Integer) As Integer
'retorna almoxarifado da filial oficial correspondente

    If iAlmoxarifadoFilialReal <= DELTA_FILIALREAL_OFICIAL Then
        Almoxarifado_ConvFRFO = iAlmoxarifadoFilialReal
    Else
        Almoxarifado_ConvFRFO = iAlmoxarifadoFilialReal - DELTA_FILIALREAL_OFICIAL
    End If
    
End Function

Function NFiscal_Converte_Tipo(ByVal objNFiscal As ClassNFiscal) As Integer
'Devolve o tipo de nfiscal analogo ao da nf passada mas que nao tenha vinculo com pedido de compras ou de vendas
'Bastará criar um select case das nfs vinculadas a pv e pc. No futuro posso colocar na tabela tiposdocinfo.

Dim iNovoTipo As Integer
    
    Select Case objNFiscal.iTipoNFiscal
    
        Case 185
            iNovoTipo = 108
        Case 186
            iNovoTipo = 109
        Case 187
            iNovoTipo = 188
        Case 189
            iNovoTipo = 190
        Case 156
            iNovoTipo = 149
        Case 103
            iNovoTipo = 75
        Case 102
            iNovoTipo = 74
        Case 101
            iNovoTipo = 73
        Case 100
            iNovoTipo = 72
        Case 81
            iNovoTipo = 67
        Case 80
            iNovoTipo = 66
        Case 92
            iNovoTipo = 63
        Case 93
            iNovoTipo = 62
        Case 94
            iNovoTipo = 61
        Case 97
            iNovoTipo = 60
        Case 95
            iNovoTipo = 59
        Case 96
            iNovoTipo = 58
        Case 85
            iNovoTipo = 57
        Case 84
            iNovoTipo = 56
        Case 99
            iNovoTipo = 52
        Case 98
            iNovoTipo = 48
        Case 153
            iNovoTipo = 47
        Case 155
            iNovoTipo = 46
        Case 154
            iNovoTipo = 43
        Case 150
            iNovoTipo = 36
        Case 151
            iNovoTipo = 34
        Case 152
            iNovoTipo = 33
        Case 83
            iNovoTipo = 27
        Case 86
            iNovoTipo = 26
        Case 88
            iNovoTipo = 25
        Case 87
            iNovoTipo = 19
        Case 82
            iNovoTipo = 11
        Case 91
            iNovoTipo = 10
        Case 90
            iNovoTipo = 9
        Case 89
            iNovoTipo = 3
        
        Case Else
            iNovoTipo = objNFiscal.iTipoNFiscal
    
    End Select
    
    NFiscal_Converte_Tipo = iNovoTipo
     
End Function

Private Function ItemNFiscal_ClonarAlocacoes(ByVal objItemOrig As ClassItemNF, ByVal objItemNovo As ClassItemNF) As Long

Dim objItemAlocOrig As ClassItemNFAlocacao, objItemAlocNovo As ClassItemNFAlocacao
Dim objItemRomaneioGradeOrig As ClassItemRomaneioGrade
Dim objItemRomaneioGradeNovo As ClassItemRomaneioGrade
Dim objReservaItemOrig As ClassReservaItem
Dim objReservaItemNovo As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long



    For Each objItemAlocOrig In objItemOrig.ColAlocacoes

        Set objItemAlocNovo = objItemNovo.ColAlocacoes.Add(Almoxarifado_ConvFRFO(objItemAlocOrig.iAlmoxarifado), "", objItemAlocOrig.dQuantidade)
        objItemAlocNovo.sUnidadeMed = objItemAlocOrig.sUnidadeMed
        objItemAlocNovo.iTransferencia = objItemAlocOrig.iTransferencia
    
    Next
    
    
    
    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeOrig In objItemOrig.colItensRomaneioGrade
                    
        Set objItemRomaneioGradeNovo = New ClassItemRomaneioGrade
            
        objItemRomaneioGradeNovo.sProduto = objItemRomaneioGradeOrig.sProduto
        objItemRomaneioGradeNovo.dQuantOP = objItemRomaneioGradeOrig.dQuantOP
        objItemRomaneioGradeNovo.dQuantSC = objItemRomaneioGradeOrig.dQuantSC
        objItemRomaneioGradeNovo.dQuantPV = objItemRomaneioGradeOrig.dQuantPV
        objItemRomaneioGradeNovo.sDescricao = objItemRomaneioGradeOrig.sDescricao
        objItemRomaneioGradeNovo.sUMEstoque = objItemRomaneioGradeOrig.sUMEstoque
        objItemRomaneioGradeNovo.dQuantidade = objItemRomaneioGradeOrig.dQuantidade
        objItemRomaneioGradeNovo.dQuantAFaturar = objItemRomaneioGradeOrig.dQuantAFaturar
        objItemRomaneioGradeNovo.dQuantReservada = objItemRomaneioGradeOrig.dQuantReservada
        objItemRomaneioGradeNovo.lNumIntItemPV = objItemRomaneioGradeOrig.lNumIntItemPV
        objItemRomaneioGradeNovo.lNumIntDoc = objItemRomaneioGradeOrig.lNumIntDoc
                            
        objItemNovo.colItensRomaneioGrade.Add objItemRomaneioGradeNovo
    
        'Transfere as informações de Localização
        Set objItemRomaneioGradeNovo.colLocalizacao = New Collection
            
        For Each objReservaItemOrig In objItemRomaneioGradeOrig.colLocalizacao
            
            objAlmoxarifado.iCodigo = Almoxarifado_ConvFRFO(objReservaItemOrig.iAlmoxarifado)
                        
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> 25056 And lErro <> SUCESSO Then gError 94331
            If lErro = 25056 Then gError 94332
                
'            iFilialEmpresa = giFilialEmpresa
'            lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa)
'            If lErro <> SUCESSO Then gError 94331
            
        '    If objAlmoxarifado.iFilialEmpresa = FilialEmpresa_ConvFRFO(objReservaItemOrig.iFilialEmpresa) Then
            
                Set objReservaItemNovo = New ClassReservaItem
                
                objReservaItemNovo.dQuantidade = objReservaItemOrig.dQuantidade
                objReservaItemNovo.dtDataValidade = objReservaItemOrig.dtDataValidade
                objReservaItemNovo.iAlmoxarifado = Almoxarifado_ConvFRFO(objReservaItemOrig.iAlmoxarifado)
                objReservaItemNovo.iFilialEmpresa = objAlmoxarifado.iFilialEmpresa
                objReservaItemNovo.lNumIntDoc = objReservaItemOrig.lNumIntDoc
                objReservaItemNovo.sAlmoxarifado = objReservaItemOrig.sAlmoxarifado
                objReservaItemNovo.sResponsavel = objReservaItemOrig.sResponsavel
                
                objItemRomaneioGradeNovo.colLocalizacao.Add objReservaItemNovo
         '   End If
        Next
    
    Next
    
    
End Function

Private Function NFiscal_ClonarItens(ByVal objNFiscalOriginal As ClassNFiscal, ByVal objNFiscalClone As ClassNFiscal, ByVal dFatorValor As Double) As Long
'O valor dos itens tem que ser proporcionalizados.

Dim objItemOrig As ClassItemNF, objItemNovo As ClassItemNF
Dim dFatoUMTrib As Double, lErro As Long
Dim objProduto As ClassProduto
Dim objNFiscalOrigAux1 As ClassNFiscal
Dim objNFiscalOrigAux51 As ClassNFiscal
Dim objItemNFOrigAux1 As ClassItemNF
Dim objItemNFOrigAux51 As ClassItemNF
Dim objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFiscal_ClonarItens

    For Each objItemOrig In objNFiscalOriginal.ColItensNF
    
        Set objItemNovo = New ClassItemNF
        
        With objItemNovo
        
            'desnecessario
            '.lNumIntNF As Long
            
            .iItem = objItemOrig.iItem
            .sProduto = objItemOrig.sProduto
            .sUnidadeMed = objItemOrig.sUnidadeMed
            .dQuantidade = objItemOrig.dQuantidade
            
            If dFatorValor <> 1 Then
                .dPrecoUnitario = Round(objItemOrig.dPrecoUnitario * dFatorValor, 2)
                .dValorDesconto = Round(objItemOrig.dValorDesconto * dFatorValor, 2)
                .dCusto = Round(objItemOrig.dCusto * dFatorValor, 2) '??????
            Else
                .dPrecoUnitario = objItemOrig.dPrecoUnitario
                .dValorDesconto = objItemOrig.dValorDesconto
                .dCusto = objItemOrig.dCusto
            End If
            
            .dPercDesc = objItemOrig.dPercDesc
            .dtDataEntrega = objItemOrig.dtDataEntrega
            .sDescricaoItem = objItemOrig.sDescricaoItem
            .dValorAbatComissao = 0 '??? objItemOrig.dValorAbatComissao
            
            '.lNumIntPedVenda =objItemOrig.
            '.lNumIntItemPedVenda =objItemOrig.
            
            'desnecessario
            '.lNumIntDoc =objItemOrig.
            
            'em desuso
            '.lNumIntTrib =objItemOrig.
            
            'copiar alocacoes
            Call ItemNFiscal_ClonarAlocacoes(objItemOrig, objItemNovo)
            Call ItemNFiscal_ClonarInfoAdicDocItem(objItemOrig.objInfoAdicDocItem, objItemNovo.objInfoAdicDocItem)
            
            'converter
            .iAlmoxarifado = Almoxarifado_ConvFRFO(objItemOrig.iAlmoxarifado)
            
            'se for uma nota de remessa pedido e estiver gerando uma nota de remessa
            'as alocacoes tem que estar nos itens da nota
            'só deverá ter uma alocacao por item
            If objNFiscalOriginal.iTipoNFiscal >= 150 And objNFiscalOriginal.iTipoNFiscal <= 156 Then
'                If objItemNovo.ColAlocacoes.Count > 1 Then gError 105152
                If objItemNovo.ColAlocacoes.Count < 1 Then gError 126992
                .iAlmoxarifado = objItemNovo.ColAlocacoes.Item(1).iAlmoxarifado
            End If
            
            .sAlmoxarifadoNomeRed = objItemOrig.sAlmoxarifadoNomeRed
            
            .iStatus = objItemOrig.iStatus
            
            If objItemOrig.lNumIntDocOrig <> 0 Then

                Set objTipoDocInfo = New ClassTipoDocInfo
                Set objNFiscalOrigAux1 = New ClassNFiscal
                Set objNFiscalOrigAux51 = New ClassNFiscal
                Set objItemNFOrigAux1 = New ClassItemNF
                Set objItemNFOrigAux51 = New ClassItemNF

                'Le a NF original da 51
                objItemNFOrigAux51.lNumIntDoc = objItemOrig.lNumIntDocOrig

                lErro = CF("ItemNFiscal_Le", objItemNFOrigAux51)
                If lErro <> SUCESSO And lErro <> 35225 Then gError ERRO_SEM_MENSAGEM

                objNFiscalOrigAux51.lNumIntDoc = objItemNFOrigAux51.lNumIntNF

                lErro = CF("NFiscal_Le", objNFiscalOrigAux51)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM

                objTipoDocInfo.iCodigo = objNFiscalOrigAux51.iTipoNFiscal
                    
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
                
                'Só pega a original se a NF Original for interna, porque senão provavelmente não terá na filial 1
                'Pelo menos foi o que foi visto no BD
                If objTipoDocInfo.iEmitente = DOCINFO_EMPRESA Then
                
                    'Le a NF Original da 1
                    objNFiscalOrigAux1.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscalOrigAux51.iFilialEmpresa)
                    If ISSerieEletronica(objNFiscalOrigAux51.sSerie) Then
                        objNFiscalOrigAux1.sSerie = "1-e"
                    Else
                        objNFiscalOrigAux1.sSerie = "1"
                    End If
                    objNFiscalOrigAux1.lNumNotaFiscal = objNFiscalOrigAux51.lNumNotaFiscal
                    objNFiscalOrigAux1.iTipoNFiscal = NFiscal_Converte_Tipo(objNFiscalOrigAux51)
                    objNFiscalOrigAux1.dtDataEmissao = objNFiscalOrigAux51.dtDataEmissao
                    objNFiscalOrigAux1.lCliente = objNFiscalOrigAux51.lCliente
                    objNFiscalOrigAux1.iFilialCli = objNFiscalOrigAux51.iFilialCli
                    objNFiscalOrigAux1.lFornecedor = objNFiscalOrigAux51.lFornecedor
                    objNFiscalOrigAux1.iFilialForn = objNFiscalOrigAux51.iFilialForn
    
                    'Verifica se a existe nota fiscal está cadastrada e pega o numintdoc
                    lErro = CF("NFiscal_Le_1", objNFiscalOrigAux1)
                    If lErro <> SUCESSO And lErro <> 83971 Then gError ERRO_SEM_MENSAGEM
    
                    If lErro <> SUCESSO Then gError 208901
    
                    'Lê os itens da nota fiscal
                    lErro = CF("NFiscalItens_Le", objNFiscalOrigAux1)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                    .sSerieNFOrig = objNFiscalOrigAux1.sSerie
                    .lNumNFOrig = objNFiscalOrigAux1.lNumIntNotaOriginal
                    .iItemNFOrig = objItemNFOrigAux51.iItem
    
                    For Each objItemNFOrigAux1 In objNFiscalOrigAux1.colItens
                        If objItemNFOrigAux1.iItem = objItemNFOrigAux51.iItem Then
                            Exit For
                        End If
                    Next
    
                    .lNumIntDocOrig = objItemNFOrigAux1.lNumIntDoc
                    
                Else
                
                    .lNumIntDocOrig = objItemOrig.lNumIntDocOrig
                    
                End If

            End If
            
            Call .objTributacaoItemNF.Copia(objItemOrig.objTributacaoItemNF)
            
        End With
        
        If dFatorValor <> 1 Then
            If (objItemOrig.objTributacaoItemNF.sUMTrib = objItemOrig.sUnidadeMed) Then
                objItemNovo.objTributacaoItemNF.dValorUnitTrib = objItemNovo.dPrecoUnitario
            Else
            
                Set objProduto = New ClassProduto
    
                objProduto.sCodigo = objItemOrig.sProduto
                
                'ler dados do produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 201193
        
                'Faz a conversão da UM da tela para a UM de estoque
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemOrig.sUnidadeMed, objItemOrig.objTributacaoItemNF.sUMTrib, dFatoUMTrib)
                If lErro <> SUCESSO Then gError 201193
                If dFatoUMTrib = 0 Then dFatoUMTrib = 1
                objItemNovo.objTributacaoItemNF.dValorUnitTrib = Round(objItemNovo.dPrecoUnitario / dFatoUMTrib, 2)
            End If
        End If
            
        With objItemNovo
            .iControleEstoque = objItemOrig.iControleEstoque
            .sUMEstoque = objItemOrig.sUMEstoque
            .iClasseUM = objItemOrig.iClasseUM
            .sUMVenda = objItemOrig.sUMVenda
            .dQuantUMVenda = objItemOrig.dQuantUMVenda
            .sCcl = objItemOrig.sCcl
            .iApropriacaoProd = objItemOrig.iApropriacaoProd
            
            '.colItemNFItemPC As New Collection
            '.colItemNFItemRC As New Collection
            
'            .dCusto = Round(objItemOrig.dCusto * dFatorValor, 2) '??????
            
            '.colRastreamento As New Collection
    
        End With
        
        Call objNFiscalClone.ColItensNF.Add1(objItemNovo)
    
    Next

    NFiscal_ClonarItens = SUCESSO
    
    Exit Function
    
Erro_NFiscal_ClonarItens:
    
    NFiscal_ClonarItens = gErr
    
    Select Case gErr
    
        Case 201193
    
        Case 105152
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCALREMESSA_ALOCACAO", gErr, objItemOrig.iItem)
    
        Case 126992
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCALREMESSA_ALOCACAO1", gErr, objItemOrig.iItem)
        
        Case 208901
            Call Rotina_Erro(vbOKOnly, "ERRO_NFORIG_NAO_CADASTRADA", gErr, objNFiscalOrigAux1.lNumNotaFiscal, objNFiscalOrigAux1.sSerie, objNFiscalOrigAux1.iFilialEmpresa)
        
        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161629)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161626)
     
    End Select
     
    Exit Function
    
    
End Function

Private Function NFiscal_ClonarParcPag(ByVal objNFiscalOriginal As ClassNFiscal, ByVal objNFiscalClone As ClassNFiscal) As Long
'copiar as datas de vcto e proporcionalizar os valores, deixando a sobra, se houver na ultima parcela.

Dim lErro As Long, objParcelaPagar As ClassParcelaPagar
Dim dSobraValor As Double, dValorParcela As Double, dTotalNFOrig As Double, dTotalNFClone As Double

On Error GoTo Erro_NFiscal_ClonarParcPag

    If objNFiscalOriginal.ColParcelaPagar.Count <> 0 Then
    
        dTotalNFOrig = Round(objNFiscalOriginal.dValorTotal - (objNFiscalOriginal.objTributacaoNF.dIRRFValor + objNFiscalOriginal.objTributacaoNF.dPISRetido + objNFiscalOriginal.objTributacaoNF.dCOFINSRetido + objNFiscalOriginal.objTributacaoNF.dCSLLRetido), 2)
        dTotalNFClone = Round(objNFiscalClone.dValorTotal - (objNFiscalClone.objTributacaoNF.dIRRFValor + objNFiscalClone.objTributacaoNF.dPISRetido + objNFiscalClone.objTributacaoNF.dCOFINSRetido + objNFiscalOriginal.objTributacaoNF.dCSLLRetido), 2)
        dSobraValor = dTotalNFClone
        
        For Each objParcelaPagar In objNFiscalOriginal.ColParcelaPagar
            
            dValorParcela = Round(objParcelaPagar.dValor * dTotalNFClone / dTotalNFOrig, 2)
            dSobraValor = dSobraValor - dValorParcela
            
            With objParcelaPagar
                Call objNFiscalClone.ColParcelaPagar.Add(.lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, dValorParcela, dValorParcela, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, "")
            End With
            
        Next
            
        If Abs(dSobraValor) > DELTA_VALORMONETARIO Then
            Set objParcelaPagar = objNFiscalClone.ColParcelaPagar.Item(objNFiscalClone.ColParcelaPagar.Count)
            objParcelaPagar.dSaldo = Round(objParcelaPagar.dSaldo + dSobraValor, 2)
            objParcelaPagar.dValor = Round(objParcelaPagar.dValor + dSobraValor, 2)
        End If
        
    End If
    
    NFiscal_ClonarParcPag = SUCESSO
    
    Exit Function
    
Erro_NFiscal_ClonarParcPag:

    NFiscal_ClonarParcPag = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161627)
     
    End Select
     
    Exit Function

End Function

Private Function NFiscal_ClonarParcRec(ByVal objNFiscalOriginal As ClassNFiscal, ByVal objNFiscalClone As ClassNFiscal) As Long
'copiar as datas de vcto e proporcionalizar os valores, deixando a sobra, se houver na ultima parcela.

Dim lErro As Long, objParcelaReceber As ClassParcelaReceber
Dim dSobraValor As Double, dValorParcela As Double, dTotalNFOrig As Double, dTotalNFClone As Double

On Error GoTo Erro_NFiscal_ClonarParcRec

    If objNFiscalOriginal.ColParcelaReceber.Count <> 0 Then
    
        dTotalNFOrig = Round(objNFiscalOriginal.dValorTotal - (objNFiscalOriginal.objTributacaoNF.dIRRFValor + objNFiscalOriginal.objTributacaoNF.dPISRetido + objNFiscalOriginal.objTributacaoNF.dCOFINSRetido + objNFiscalOriginal.objTributacaoNF.dCSLLRetido), 2)
        dTotalNFClone = Round(objNFiscalClone.dValorTotal - (objNFiscalClone.objTributacaoNF.dIRRFValor + objNFiscalClone.objTributacaoNF.dPISRetido + objNFiscalClone.objTributacaoNF.dCOFINSRetido + objNFiscalOriginal.objTributacaoNF.dCSLLRetido), 2)
        dSobraValor = dTotalNFClone
        
        For Each objParcelaReceber In objNFiscalOriginal.ColParcelaReceber
            
            dValorParcela = Round(objParcelaReceber.dValor * dTotalNFClone / dTotalNFOrig, 2)
            dSobraValor = dSobraValor - dValorParcela
            
            With objParcelaReceber
                Call objNFiscalClone.ColParcelaReceber.Add(.lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, dValorParcela, dValorParcela, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, 0, "", dValorParcela)
            End With
            
        Next
            
        If Abs(dSobraValor) > DELTA_VALORMONETARIO Then
            Set objParcelaReceber = objNFiscalClone.ColParcelaReceber.Item(objNFiscalClone.ColParcelaReceber.Count)
            objParcelaReceber.dSaldo = Round(objParcelaReceber.dSaldo + dSobraValor, 2)
            objParcelaReceber.dValor = Round(objParcelaReceber.dValor + dSobraValor, 2)
        End If
    
    End If
    
    NFiscal_ClonarParcRec = SUCESSO
    
    Exit Function
    
Erro_NFiscal_ClonarParcRec:

    NFiscal_ClonarParcRec = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161628)
     
    End Select
     
    Exit Function

End Function

Function NFiscal_Clonar(ByVal objNFiscalOriginal As ClassNFiscal, ByVal objNFiscalClone As ClassNFiscal) As Long
'copia dados correspondentes à ClassNFiscal, ignorando alguns campos e modificando outros
'Na versao oficial não haverá informacao de rastreamento.
'nfs associadas a pv, pc ou orcamento perderão este vinculo e terao seu tipo substituido.
'O valor dos itens e complementos tem que ser proporcionalizados.

Dim lErro As Long, dFatorValor As Double
Dim objItemNovo As ClassItemNF
Dim objNFiscalOrigAux1 As New ClassNFiscal
Dim objNFiscalOrigAux51 As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_NFiscal_Clonar

    With objNFiscalClone
    
        '.lNumOrcamentoVenda As Long
        '.iFilialOrcamento As Integer
        
        '.lNumIntDoc As Long
        '.iFilialEmpresa As Integer
        '.sSerie As String
        .iNaoVerificaCredito = NAO_VERIFICA_CREDITO_CLIENTE
        .lNumNotaFiscal = objNFiscalOriginal.lNumNotaFiscal
        
        .lCliente = objNFiscalOriginal.lCliente
        .iFilialCli = objNFiscalOriginal.iFilialCli
        .iFilialEntrega = objNFiscalOriginal.iFilialEntrega
        .lFornecedor = objNFiscalOriginal.lFornecedor
        .iFilialForn = objNFiscalOriginal.iFilialForn
        .dtDataEmissao = objNFiscalOriginal.dtDataEmissao
        .dtDataSaida = objNFiscalOriginal.dtDataSaida
        '.lNumPedidoVenda As Long
        .sNumPedidoTerc = objNFiscalOriginal.sNumPedidoTerc
        
        lErro = CF("NFiscal_ObtemFatorValor", objNFiscalOriginal.iFilialEmpresa, objNFiscalOriginal.iTipoNFiscal, objNFiscalOriginal.sSerie, dFatorValor)
        If lErro <> SUCESSO Then gError 500006
        
        lErro = NFiscal_ClonarItens(objNFiscalOriginal, objNFiscalClone, dFatorValor)
        If lErro <> SUCESSO Then gError 105153
        
        If dFatorValor <> 0 Then
        
            For Each objItemNovo In objNFiscalClone.ColItensNF
        
                If objItemNovo.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
            
                    .dValorProdutos = .dValorProdutos + ((objItemNovo.dQuantidade * objItemNovo.dPrecoUnitario) - objItemNovo.dValorDesconto)
                    
                End If
        
            Next
        
            '.dValorProdutos = Round(objNFiscalOriginal.dValorProdutos * dFatorValor, 2)
            '.dValorProdutos = PrecoTotal_Calcula(objNFiscalOriginal, dFatorValor)
            .dValorFrete = Round(objNFiscalOriginal.dValorFrete * dFatorValor, 2)
            .dValorSeguro = Round(objNFiscalOriginal.dValorSeguro * dFatorValor, 2)
            .dValorOutrasDespesas = Round(objNFiscalOriginal.dValorOutrasDespesas * dFatorValor, 2)
            .dValorDesconto = Round(objNFiscalOriginal.dValorDesconto * dFatorValor, 2)
            .dValorProdutos = .dValorProdutos - .dValorDesconto
        
        End If
        
        .iCodTransportadora = objNFiscalOriginal.iCodTransportadora
        .iCodTranspRedesp = objNFiscalOriginal.iCodTranspRedesp
        .sMensagemNota = objNFiscalOriginal.sMensagemNota
        .sMensagemCorpoNota = objNFiscalOriginal.sMensagemCorpoNota
        .iTabelaPreco = objNFiscalOriginal.iTabelaPreco
        
        .iTipoNFiscal = NFiscal_Converte_Tipo(objNFiscalOriginal)
        
        .sNaturezaOp = objNFiscalOriginal.sNaturezaOp
        .dPesoLiq = objNFiscalOriginal.dPesoLiq
        .dPesoBruto = objNFiscalOriginal.dPesoBruto
        .dtDataVencimento = objNFiscalOriginal.dtDataVencimento
        
        'em desuso
        '.lNumIntTrib As Long
        
        .sPlaca = objNFiscalOriginal.sPlaca
        .sPlacaUF = objNFiscalOriginal.sPlacaUF
        .lVolumeQuant = objNFiscalOriginal.lVolumeQuant
        .lVolumeEspecie = objNFiscalOriginal.lVolumeEspecie
        .lVolumeMarca = objNFiscalOriginal.lVolumeMarca
        .iCanal = objNFiscalOriginal.iCanal
        
        If objNFiscalOriginal.lNumIntNotaOriginal <> 0 Then
        
            'Le a NF original da 51
            objNFiscalOrigAux51.lNumIntDoc = objNFiscalOriginal.lNumIntNotaOriginal
            
            lErro = CF("NFiscal_Le", objNFiscalOrigAux51)
            If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
            
            objTipoDocInfo.iCodigo = objNFiscalOrigAux51.iTipoNFiscal
                
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            'Só pega a original se a NF Original for interna, porque senão provavelmente não terá na filial 1
            'Pelo menos foi o que foi visto no BD
            If objTipoDocInfo.iEmitente = DOCINFO_EMPRESA Then
            
                'Le a NF Original da 1
                objNFiscalOrigAux1.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscalOrigAux51.iFilialEmpresa)
                If ISSerieEletronica(objNFiscalOrigAux51.sSerie) Then
                    objNFiscalOrigAux1.sSerie = "1-e"
                Else
                    objNFiscalOrigAux1.sSerie = "1"
                End If
                objNFiscalOrigAux1.lNumNotaFiscal = objNFiscalOrigAux51.lNumNotaFiscal
                objNFiscalOrigAux1.iTipoNFiscal = NFiscal_Converte_Tipo(objNFiscalOrigAux51)
                objNFiscalOrigAux1.dtDataEmissao = objNFiscalOrigAux51.dtDataEmissao
                objNFiscalOrigAux1.lCliente = objNFiscalOrigAux51.lCliente
                objNFiscalOrigAux1.iFilialCli = objNFiscalOrigAux51.iFilialCli
                objNFiscalOrigAux1.lFornecedor = objNFiscalOrigAux51.lFornecedor
                objNFiscalOrigAux1.iFilialForn = objNFiscalOrigAux51.iFilialForn
                
                'Verifica se a existe nota fiscal está cadastrada e pega o numintdoc
                lErro = CF("NFiscal_Le_1", objNFiscalOrigAux1)
                If lErro <> SUCESSO And lErro <> 83971 Then gError ERRO_SEM_MENSAGEM
                
                If lErro <> SUCESSO Then gError 208901
            
                .lNumIntNotaOriginal = objNFiscalOrigAux1.lNumIntDoc
                
            Else
                .lNumIntNotaOriginal = objNFiscalOriginal.lNumIntNotaOriginal
            End If
        
        End If
        
        'serao ignoradas
        '.colComissoesNF As New Collection
        
        .dtDataEntrada = objNFiscalOriginal.dtDataEntrada
        
        'calcular em funcao dos outros valores
        .dValorTotal = Round(.dValorProdutos + objNFiscalOriginal.objTributacaoNF.dICMSSubstValor + objNFiscalOriginal.objTributacaoNF.dIPIValor + .dValorFrete + .dValorSeguro + .dValorOutrasDespesas + IIf(objNFiscalOriginal.objTributacaoNF.iISSIncluso <> 0, 0, objNFiscalOriginal.objTributacaoNF.dISSValor), 2)
        
        'vao ser preenchidos depois, na propria gravacao
        '.iClasseDocCPR As Integer
        '.lNumIntDocCPR As Long
        
        .iStatus = objNFiscalOriginal.iStatus
        
        .sCodUsuario = objNFiscalOriginal.sCodUsuario
        '.iFilialPedido As Integer
        .lClienteBenef = objNFiscalOriginal.lClienteBenef
        .iFilialCliBenef = objNFiscalOriginal.iFilialCliBenef
        .lFornecedorBenef = objNFiscalOriginal.lFornecedorBenef
        .iFilialFornBenef = objNFiscalOriginal.iFilialFornBenef
        
        'vai ser criado depois, na propria gravacao
        '.objMovEstoque As ClassMovEstoque
        
        .iCondicaoPagto = objNFiscalOriginal.iCondicaoPagto
        .sVolumeNumero = objNFiscalOriginal.sVolumeNumero
        .iFreteRespons = objNFiscalOriginal.iFreteRespons
        .dtDataRegistro = objNFiscalOriginal.dtDataRegistro
        .dtDataReferencia = objNFiscalOriginal.dtDataReferencia
        
        '.lNumRecebimento As Long
        
        '??? deve ser lixo mas por via das duvidas
        .iTipoDocInfo = .iTipoNFiscal
        
        .sObservacao = objNFiscalOriginal.sObservacao
        .sCodUsuarioCancel = objNFiscalOriginal.sCodUsuarioCancel
        .sMotivoCancel = objNFiscalOriginal.sMotivoCancel
        
        '.objConhecimentoFrete As New ClassConhecimentoFrete
        '.objRastreamento As Object
        
        '.sNomeTelaNFiscal As String
        
        .dtHoraEntrada = objNFiscalOriginal.dtHoraEntrada
        .dtHoraSaida = objNFiscalOriginal.dtHoraSaida
        
        'os campos abaixo só interessam p/conhecimentos de frete
        '.sDestino As String
        '.sOrigem As String
        '.dValorContainer As Double
        '.dValorMercadoria As Double
        
        '.colComprovServ As New Collection
    
        Call objNFiscalClone.objTributacao.Copia(objNFiscalOriginal.objTributacao)
        
        Call NFiscal_ClonarParcPag(objNFiscalOriginal, objNFiscalClone)
        
        Call NFiscal_ClonarParcRec(objNFiscalOriginal, objNFiscalClone)
        
        Call NFiscal_ClonarInfoAdic(objNFiscalOriginal.objInfoAdic, objNFiscalClone.objInfoAdic)
        
    End With
    
    NFiscal_Clonar = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Clonar:

    NFiscal_Clonar = gErr
     
    Select Case gErr
          
        Case 105153, 500006
        
        Case 208901
            Call Rotina_Erro(vbOKOnly, "ERRO_NFORIG_NAO_CADASTRADA", gErr, objNFiscalOrigAux1.lNumNotaFiscal, objNFiscalOrigAux1.sSerie, objNFiscalOrigAux1.iFilialEmpresa)
        
        Case ERRO_SEM_MENSAGEM
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161629)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161629)
     
    End Select
     
    Exit Function

End Function

Function NFiscalNumAuto2(ByVal objNFiscal As ClassNFiscal) As Long
'Lê o Proximo número na tabela de Série e Coloca no objNFiscal
'Faz Lock Exclusive e atualiza o Número na Tabela de Série

Dim lErro As Long
Dim tSerie As typeSerie
Dim lComando As Long
Dim lComando2 As Long

On Error GoTo Erro_NFiscalNumAuto2

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 60436

'    lComando1 = Comando_Abrir()
'    If lComando1 = 0 Then Error 60437

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 60437

    lErro = Comando_Executar(lComando, "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", tSerie.lProxNumNFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 60438

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60439
    
    If lErro <> AD_SQL_SUCESSO Then Error 32285
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        objNFiscal.sSerie = "1-e"
    Else
        objNFiscal.sSerie = "1"
    End If
        
    lErro = Comando_ExecutarPos(lComando2, "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, tSerie.lProxNumNFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 60438

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 60439
    
    lErro = Comando_LockExclusive(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 60440
    
    objNFiscal.lNumNotaFiscal = tSerie.lProxNumNFiscal
    
'    tSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal + 1
'
'    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumNFiscal = ?", lComando, tSerie.lProxNumNFiscal)
'    If lErro <> AD_SQL_SUCESSO Then Error 60441
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    NFiscalNumAuto2 = SUCESSO
    
    Exit Function
    
Erro_NFiscalNumAuto2:

    NFiscalNumAuto2 = Err
    
    Select Case Err
    
        Case 32285
        
        Case 60436, 60437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 60438, 60439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", Err, objNFiscal.sSerie)

        Case 60440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", Err)
        
        Case 60441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", Err, objNFiscal.sSerie)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161630)
            
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function
    
End Function

Public Function PrecoTotal_Calcula(objNFiscalOriginal As ClassNFiscal, dFator As Double)

Dim dPrecoTotalReal As Double
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim dValorDescontoGlobal As Double
Dim objItemOrig As ClassItemNF
Dim dSubTotal As Double

    dSubTotal = 0

    For Each objItemOrig In objNFiscalOriginal.ColItensNF
                
        'copia os valores para uma area de memoria de trabalho
        dPrecoUnitario = Round(objItemOrig.dPrecoUnitario * dFator, 2)
        dPercentDesc = objItemOrig.dPercDesc
        dQuantidade = objItemOrig.dQuantidade
        dDesconto = Round(objItemOrig.dValorDesconto * dFator, 2)
        
        'Calcula o Valor Real (com os descontos a nivel de item)
        Call ValorReal_Calcula(dQuantidade, dPrecoUnitario, dPercentDesc, dDesconto, dPrecoTotalReal)
    
        'acumula o somatorio
        dSubTotal = dSubTotal + dPrecoTotalReal

    Next
    
    'obtem o valor do desconto aplicando o fator no original
    dValorDescontoGlobal = Round(objNFiscalOriginal.dValorDesconto * dFator, 2)

    'tulio180303 arredondamento
    dSubTotal = StrParaDbl(Format(CStr(dSubTotal - dValorDescontoGlobal), "standard"))
    
    PrecoTotal_Calcula = dSubTotal
    
    Exit Function

End Function

Private Sub ValorReal_Calcula(dQuantidade As Double, dValorUnitario As Double, dPercentDesc As Double, dDesconto As Double, dValorReal As Double)
'Calcula o Valor Real

Dim dValorTotal As Double
Dim dPercDesc1 As Double
Dim dPercDesc2 As Double

    dValorTotal = dValorUnitario * dQuantidade

    'Se o Percentual Desconto estiver preenchido
    If dPercentDesc > 0 Then

        'Testa se o desconto está preenchido
        If dDesconto = 0 Then
            dPercDesc2 = 0
        Else
            'Calcula o Percentual em cima dos valores passados
            dPercDesc2 = dDesconto / dValorTotal
            dPercDesc2 = CDbl(Format(dPercDesc2, "0.0000"))
        End If
        'se os percentuais (passado e calulado) forem diferentes calcula-se o desconto
        If dPercentDesc <> dPercDesc2 Then dDesconto = dPercentDesc * dValorTotal

    End If

    dValorReal = dValorTotal - dDesconto

End Sub

Function NFiscal_Exclui_Clone(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil) As Long

Dim lErro As Long
Dim objNFiscalOficial As New ClassNFiscal
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer

On Error GoTo Erro_NFiscal_Exclui_Clone

    'Se nf nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa) <> objNFiscal.iFilialEmpresa Then
    
        objNFiscalOficial.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
        
        If ISSerieEletronica(objNFiscal.sSerie) Then
            objNFiscalOficial.sSerie = "1-e"
        Else
            objNFiscalOficial.sSerie = "1"
        End If
    
        objNFiscalOficial.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
        
        objNFiscalOficial.iTipoNFiscal = NFiscal_Converte_Tipo(objNFiscal)
        
        objNFiscalOficial.dtDataEmissao = objNFiscal.dtDataEmissao
        
        objNFiscalOficial.lCliente = objNFiscal.lCliente
        
        objNFiscalOficial.iFilialCli = objNFiscal.iFilialCli
        
        objNFiscalOficial.lFornecedor = objNFiscal.lFornecedor
        
        objNFiscalOficial.iFilialForn = objNFiscal.iFilialForn
            
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_1", objNFiscalOficial)
        If lErro <> SUCESSO And lErro <> 83971 Then gError 126970
        
        If lErro = SUCESSO Then
            
            'Lê os itens da nota fiscal
            lErro = CF("NFiscalItens_Le", objNFiscalOficial)
            If lErro <> SUCESSO Then gError 126971
                
            Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
        
            colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, objNFiscalOficial.iFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
            
            'Lê as configurações em ESTConfig
            lErro = CF("ESTConfig_Le_Configs", colConfig)
            If lErro <> SUCESSO Then gError 126984
            
            iAceitaEstoqueNegativo = gobjMAT.iAceitaEstoqueNegativo
            
            gobjMAT.iAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
                
            lErro = CF("NotaFiscalSaida_Excluir_EmTrans", objNFiscalOficial, objContabil)
            If lErro <> SUCESSO Then gError 126972
            
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
            
        End If
            
    End If
    
    NFiscal_Exclui_Clone = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Exclui_Clone:

    NFiscal_Exclui_Clone = gErr
     
    Select Case gErr
          
        Case 126970, 126971, 126984
        
        Case 126972
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161631)
     
    End Select
     
    Exit Function

End Function

Function NFiscal_Cancela_Clone(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long

Dim lErro As Long
Dim objNFiscalOficial As New ClassNFiscal
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer

On Error GoTo Erro_NFiscal_Cancela_Clone

    'Se nf nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa) <> objNFiscal.iFilialEmpresa Then
    
        objNFiscalOficial.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
        
        If ISSerieEletronica(objNFiscal.sSerie) Then
            objNFiscalOficial.sSerie = "1-e"
        Else
            objNFiscalOficial.sSerie = "1"
        End If
        
        objNFiscalOficial.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
            
        objNFiscalOficial.sMotivoCancel = objNFiscal.sMotivoCancel
            
        'Lê a nota fiscal de saída
        lErro = CF("NFiscalInternaSaida_Le_Numero2", objNFiscalOficial)
        If lErro <> SUCESSO And lErro <> 62144 Then gError 126974
        
        If lErro = SUCESSO Then
            
            'Verifica se a nota já está cancelada
            If objNFiscalOficial.iStatus = STATUS_CANCELADO Then gError 126975
            
            'Lê os itens da nota fiscal
            lErro = CF("NFiscalItens_Le", objNFiscalOficial)
            If lErro <> SUCESSO Then gError 126976
                
            Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
        
            colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, objNFiscalOficial.iFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
            
            'Lê as configurações em ESTConfig
            lErro = CF("ESTConfig_Le_Configs", colConfig)
            If lErro <> SUCESSO Then gError 126985
            
            iAceitaEstoqueNegativo = gobjMAT.iAceitaEstoqueNegativo
            
            gobjMAT.iAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
                
            lErro = CF("NotaFiscalSaida_Cancelar_EmTrans", objNFiscalOficial, dtDataCancelamento)
            If lErro <> SUCESSO Then gError 126977
            
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
            
        End If
            
    End If
    
    NFiscal_Cancela_Clone = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Cancela_Clone:

    NFiscal_Cancela_Clone = gErr
     
    Select Case gErr
          
        Case 126974, 126976
        
        Case 126977
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
          
        Case 126975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, objNFiscalOficial.sSerie, objNFiscalOficial.lNumNotaFiscal)
        
        Case 126985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161632)
     
    End Select
     
    Exit Function

End Function


Function NFiscal_Altera_Clone(ByVal objNFiscal As ClassNFiscal, ByVal objContabil As ClassContabil, ByVal sNomeFuncGravacao As String, lNumNFOficial As Long) As Long

Dim lErro As Long
Dim objNFiscalOficial As New ClassNFiscal
Dim bNFInterna As Boolean, bClonar As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer
Dim dFatorValor As Double
Dim objSerie As New ClassSerie

On Error GoTo Erro_NFiscal_Altera_Clone

    lNumNFOficial = 0
    
    'Se nf nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa) <> objNFiscal.iFilialEmpresa Then
    
'        bClonar = False
        bClonar = True
        
        objNFiscalOficial.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
        
        objNFiscalOficial.sSerie = objNFiscal.sSerie
            
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
            
        'Lê o Tipo de Documento
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 207662
        
        'Se não encontrou o Tipo de Documento --> erro
        If lErro <> SUCESSO Then gError 207663
    
        bNFInterna = (objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIS)
        
        If bNFInterna Then
        
            
            objSerie.iFilialEmpresa = objNFiscalOficial.iFilialEmpresa
            objSerie.sSerie = objNFiscalOficial.sSerie
        
            lErro = CF("Serie_Le", objSerie)
            If lErro <> SUCESSO And lErro <> 22202 Then gError 207664
            
            'se a serie nao existir na filial oficial ==> nao clonar
            If lErro <> SUCESSO Then bClonar = False
            
        Else 'se for nf externa
        
            lErro = CF("NFiscalEntrada_Verifica_Existencia2", objNFiscal, objTipoDocInfo, True)
            If lErro <> SUCESSO And lErro <> 61414 And lErro <> 89723 Then gError 207665
            
'            'Se for uma nota nova
'            If lErro = SUCESSO Then bClonar = True
        
            lErro = CF("NFiscal_ObtemFatorValor", objNFiscal.iFilialEmpresa, objNFiscal.iTipoNFiscal, objNFiscal.sSerie, dFatorValor)
            If lErro <> SUCESSO Then gError 207666
            
            If dFatorValor = 0 Then bClonar = False
            
        End If
        
        If bClonar Then
        
            'clonar o objeto nfiscal
            lErro = NFiscal_Clonar(objNFiscal, objNFiscalOficial)
            If lErro <> SUCESSO Then gError 207667
            
            Set objNFiscalOficial.objContabil = objContabil
            
            'chamar a funcao de gravacao para o clone SEM CTB
            lErro = CF(sNomeFuncGravacao, objNFiscalOficial, Nothing)
            If lErro <> SUCESSO Then gError 207668
            
            lNumNFOficial = objNFiscalOficial.lNumNotaFiscal
        
            
        End If
    
    End If
    
    NFiscal_Altera_Clone = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Altera_Clone:

    NFiscal_Altera_Clone = gErr
     
    Select Case gErr
          
        Case 207662, 207664 To 207668
          
        Case 207663
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207669)
     
    End Select
     
    Exit Function

End Function

Function NFiscal_Ent_Cancela_Clone(ByVal objNFiscal As ClassNFiscal, ByVal dtDataCancelamento As Date) As Long

Dim lErro As Long
Dim objNFiscalOficial As New ClassNFiscal
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer

On Error GoTo Erro_NFiscal_Ent_Cancela_Clone

    'Se nf nao é de filial oficial entao
    If FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa) <> objNFiscal.iFilialEmpresa Then
    
        objNFiscalOficial.iFilialEmpresa = FilialEmpresa_ConvFRFO(objNFiscal.iFilialEmpresa)
        
        If ISSerieEletronica(objNFiscal.sSerie) Then
            objNFiscalOficial.sSerie = "1-e"
        Else
            objNFiscalOficial.sSerie = "1"
        End If
        
        objNFiscalOficial.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
            
        objNFiscalOficial.sMotivoCancel = objNFiscal.sMotivoCancel
            
        'Lê a nota fiscal de saída
        lErro = CF("NFiscalInternaEntrada_Le_Numero", objNFiscalOficial)
        If lErro <> SUCESSO And lErro <> 62144 Then gError 126974
        
        If lErro = SUCESSO Then
            
            'Verifica se a nota já está cancelada
            If objNFiscalOficial.iStatus = STATUS_CANCELADO Then gError 126975
            
            'Lê os itens da nota fiscal
            lErro = CF("NFiscalItens_Le", objNFiscalOficial)
            If lErro <> SUCESSO Then gError 126976
                
            Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
        
            colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, objNFiscalOficial.iFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
            
            'Lê as configurações em ESTConfig
            lErro = CF("ESTConfig_Le_Configs", colConfig)
            If lErro <> SUCESSO Then gError 126985
            
            iAceitaEstoqueNegativo = gobjMAT.iAceitaEstoqueNegativo
            
            gobjMAT.iAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
                
            lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans", objNFiscalOficial, dtDataCancelamento)
            If lErro <> SUCESSO Then gError 126977
            
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
            
        End If
            
    End If
    
    NFiscal_Ent_Cancela_Clone = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Ent_Cancela_Clone:

    NFiscal_Ent_Cancela_Clone = gErr
     
    Select Case gErr
          
        Case 126974, 126976
        
        Case 126977
            gobjMAT.iAceitaEstoqueNegativo = iAceitaEstoqueNegativo
          
        Case 126975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_CANCELADA", gErr, objNFiscalOficial.sSerie, objNFiscalOficial.lNumNotaFiscal)
        
        Case 126985
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161640)
     
    End Select
     
    Exit Function

End Function

Private Function ItemNFiscal_ClonarInfoAdicDocItem(ByVal objInfoAdicDocItemOrig As ClassInfoAdicDocItem, ByVal objInfoAdicDocItemNovo As ClassInfoAdicDocItem) As Long

    objInfoAdicDocItemNovo.dtDataLimiteFaturamento = objInfoAdicDocItemOrig.dtDataLimiteFaturamento
    objInfoAdicDocItemNovo.iIncluiValorTotal = objInfoAdicDocItemOrig.iIncluiValorTotal
    objInfoAdicDocItemNovo.iItem = objInfoAdicDocItemOrig.iItem
    objInfoAdicDocItemNovo.iTipoDoc = objInfoAdicDocItemOrig.iTipoDoc
    objInfoAdicDocItemNovo.lItemPedCompra = objInfoAdicDocItemOrig.lItemPedCompra
    'objInfoAdicDocItemNovo.lnumIntDocItem = objInfoAdicDocItemOrig.lnumIntDocItem
    objInfoAdicDocItemNovo.sDescProd = objInfoAdicDocItemOrig.sDescProd
    objInfoAdicDocItemNovo.sMsg = objInfoAdicDocItemOrig.sMsg
    objInfoAdicDocItemNovo.sNumPedidoCompra = objInfoAdicDocItemOrig.sNumPedidoCompra
    objInfoAdicDocItemNovo.sProduto = objInfoAdicDocItemOrig.sProduto
    
End Function

Private Function NFiscal_ClonarInfoAdic(ByVal objInfoAdicOriginal As ClassInfoAdic, ByVal objInfoAdicClone As ClassInfoAdic) As Long


    objInfoAdicClone.iTipoDoc = objInfoAdicOriginal.iTipoDoc
    'objInfoAdicClone.lNumIntDoc = objInfoAdicOriginal.lNumIntDoc
    
    If Not objInfoAdicOriginal.objCompra Is Nothing Then
        Set objInfoAdicClone.objCompra = New ClassInfoAdicCompra
        objInfoAdicClone.objCompra.iTipoDoc = objInfoAdicOriginal.objCompra.iTipoDoc
        'objInfoAdicClone.objCompra.lNumIntDoc = objInfoAdicOriginal.objCompra.lNumIntDoc
        objInfoAdicClone.objCompra.sContrato = objInfoAdicOriginal.objCompra.sContrato
        objInfoAdicClone.objCompra.sNotaEmpenho = objInfoAdicOriginal.objCompra.sNotaEmpenho
        objInfoAdicClone.objCompra.sPedido = objInfoAdicOriginal.objCompra.sPedido
    
    End If
    
    If Not objInfoAdicOriginal.objExportacao Is Nothing Then
        Set objInfoAdicClone.objExportacao = New ClassInfoAdicExportacao
        objInfoAdicClone.objExportacao.iTipoDoc = objInfoAdicOriginal.objExportacao.iTipoDoc
        'objInfoAdicClone.objExportacao.lNumIntDoc = objInfoAdicOriginal.objExportacao.lNumIntDoc
        objInfoAdicClone.objExportacao.sLocalEmbarque = objInfoAdicOriginal.objExportacao.sLocalEmbarque
        objInfoAdicClone.objExportacao.sUFEmbarque = objInfoAdicOriginal.objExportacao.sUFEmbarque
    
    End If
    
    If Not objInfoAdicOriginal.objRetEnt Is Nothing Then
        Set objInfoAdicClone.objRetEnt = New ClassRetiradaEntrega
        objInfoAdicClone.objRetEnt.iFilialCliEnt = objInfoAdicOriginal.objRetEnt.iFilialCliEnt
        objInfoAdicClone.objRetEnt.iFilialCliRet = objInfoAdicOriginal.objRetEnt.iFilialCliRet
        objInfoAdicClone.objRetEnt.iFilialFornEnt = objInfoAdicOriginal.objRetEnt.iFilialFornEnt
        objInfoAdicClone.objRetEnt.iFilialFornRet = objInfoAdicOriginal.objRetEnt.iFilialFornRet
        objInfoAdicClone.objRetEnt.iTipoDoc = objInfoAdicOriginal.objRetEnt.iTipoDoc
        objInfoAdicClone.objRetEnt.lClienteEnt = objInfoAdicOriginal.objRetEnt.lClienteEnt
        objInfoAdicClone.objRetEnt.lClienteRet = objInfoAdicOriginal.objRetEnt.lClienteRet
        objInfoAdicClone.objRetEnt.lEnderecoEnt = objInfoAdicOriginal.objRetEnt.lEnderecoEnt
        objInfoAdicClone.objRetEnt.lEnderecoRet = objInfoAdicOriginal.objRetEnt.lEnderecoRet
        objInfoAdicClone.objRetEnt.lFornecedorEnt = objInfoAdicOriginal.objRetEnt.lFornecedorEnt
        objInfoAdicClone.objRetEnt.lFornecedorRet = objInfoAdicOriginal.objRetEnt.lFornecedorRet
        'objInfoAdicClone.objRetEnt.lNumIntDoc = objInfoAdicOriginal.objRetEnt.lNumIntDoc
        objInfoAdicClone.objRetEnt.sCNPJCPFEnt = objInfoAdicOriginal.objRetEnt.sCNPJCPFEnt
        objInfoAdicClone.objRetEnt.sCNPJCPFRet = objInfoAdicOriginal.objRetEnt.sCNPJCPFRet
        Set objInfoAdicClone.objRetEnt.objEnderecoEnt = New ClassEndereco
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sEndereco = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sEndereco
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sBairro = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sBairro
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sCidade = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sCidade
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sSiglaEstado = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sSiglaEstado
        objInfoAdicClone.objRetEnt.objEnderecoEnt.iCodigoPais = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.iCodigoPais
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sCEP = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sCEP
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sTelefone1 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sTelefone1
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sTelefone2 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sTelefone2
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sEmail = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sEmail
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sFax = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sFax
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sContato = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sContato
        'objInfoAdicClone.objRetEnt.objEnderecoEnt.lCodigo = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.lCodigo
        
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sReferencia = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sReferencia
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sLogradouro = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sLogradouro
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sComplemento = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sComplemento
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sTipoLogradouro = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sTipoLogradouro
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sEmail2 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sEmail2
        objInfoAdicClone.objRetEnt.objEnderecoEnt.lNumero = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.lNumero
        objInfoAdicClone.objRetEnt.objEnderecoEnt.iTelDDD1 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.iTelDDD1
        objInfoAdicClone.objRetEnt.objEnderecoEnt.iTelDDD2 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.iTelDDD2
        objInfoAdicClone.objRetEnt.objEnderecoEnt.iFaxDDD = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.iFaxDDD
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sTelNumero1 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sTelNumero1
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sTelNumero2 = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sTelNumero2
        objInfoAdicClone.objRetEnt.objEnderecoEnt.sFaxNumero = objInfoAdicOriginal.objRetEnt.objEnderecoEnt.sFaxNumero
        Set objInfoAdicClone.objRetEnt.objEnderecoRet = New ClassEndereco
        objInfoAdicClone.objRetEnt.objEnderecoRet.sEndereco = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sEndereco
        objInfoAdicClone.objRetEnt.objEnderecoRet.sBairro = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sBairro
        objInfoAdicClone.objRetEnt.objEnderecoRet.sCidade = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sCidade
        objInfoAdicClone.objRetEnt.objEnderecoRet.sSiglaEstado = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sSiglaEstado
        objInfoAdicClone.objRetEnt.objEnderecoRet.iCodigoPais = objInfoAdicOriginal.objRetEnt.objEnderecoRet.iCodigoPais
        objInfoAdicClone.objRetEnt.objEnderecoRet.sCEP = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sCEP
        objInfoAdicClone.objRetEnt.objEnderecoRet.sTelefone1 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sTelefone1
        objInfoAdicClone.objRetEnt.objEnderecoRet.sTelefone2 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sTelefone2
        objInfoAdicClone.objRetEnt.objEnderecoRet.sEmail = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sEmail
        objInfoAdicClone.objRetEnt.objEnderecoRet.sFax = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sFax
        objInfoAdicClone.objRetEnt.objEnderecoRet.sContato = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sContato
        'objInfoAdicClone.objRetEnt.objEnderecoRet.lCodigo = objInfoAdicOriginal.objRetEnt.objEnderecoRet.lCodigo
        
        objInfoAdicClone.objRetEnt.objEnderecoRet.sReferencia = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sReferencia
        objInfoAdicClone.objRetEnt.objEnderecoRet.sLogradouro = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sLogradouro
        objInfoAdicClone.objRetEnt.objEnderecoRet.sComplemento = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sComplemento
        objInfoAdicClone.objRetEnt.objEnderecoRet.sTipoLogradouro = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sTipoLogradouro
        objInfoAdicClone.objRetEnt.objEnderecoRet.sEmail2 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sEmail2
        objInfoAdicClone.objRetEnt.objEnderecoRet.lNumero = objInfoAdicOriginal.objRetEnt.objEnderecoRet.lNumero
        objInfoAdicClone.objRetEnt.objEnderecoRet.iTelDDD1 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.iTelDDD1
        objInfoAdicClone.objRetEnt.objEnderecoRet.iTelDDD2 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.iTelDDD2
        objInfoAdicClone.objRetEnt.objEnderecoRet.iFaxDDD = objInfoAdicOriginal.objRetEnt.objEnderecoRet.iFaxDDD
        objInfoAdicClone.objRetEnt.objEnderecoRet.sTelNumero1 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sTelNumero1
        objInfoAdicClone.objRetEnt.objEnderecoRet.sTelNumero2 = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sTelNumero2
        objInfoAdicClone.objRetEnt.objEnderecoRet.sFaxNumero = objInfoAdicOriginal.objRetEnt.objEnderecoRet.sFaxNumero

    End If

End Function



