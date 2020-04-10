VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRotCalcNecesProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'        Qual o saldo "livre" do produto produzido ?
'           + Saldo disponivel em estoque (propria ou de 3o, na empresa)
            '- qtde empenhada
            ' - saldo a faturar de item de pv em aberto
            ' + saldo de op destinada a estoque

'
'    Qual o saldo "livre" de produtos comprados ?
'           + Saldo disponivel em estoque (propria ou de 3o, na empresa)
            '- qtde empenhada
            ' - saldo a faturar de item de pv em aberto

            '+ saldo de requisicoes nao baixadas desvinculado de pedidos de compra e nao destinadas a ops
            '+ saldo de pedidos de compra nao vinculado a requisicoes vinculadas a ops


'no futuro pode incluir saldo de previsao de vendas

'a analise de empenhos indiretos pode demandar percorrer itens "de baixo para cima"
    'segundo sua utilizacao, buscando saldos negativos (disponivel < empenho)
    'para evitar esta analise as ops devem ser geradas em profundidade
    'ou os empenhos indiretos serem gravados juntamente com os diretos.
        
Public Function Producao_Calcula_Necessidades(ByVal objCalcNecesProd As ClassCalcNecesProd) As Long
'calcula necessidades de producao e/ou compras a partir de necessidades informadas

Dim lErro As Long, objNecesProdInfo As ClassNecesProdInfo, dFatorConversao As Double, objProduto As New ClassProduto, dQuantidadeProd As Double
Dim objKit As New ClassKit

On Error GoTo Erro_Producao_Calcula_Necessidades

    'percorre a colecao de necessidades informadas
    For Each objNecesProdInfo In objCalcNecesProd.colNecesInfProd
    
        objProduto.sCodigo = objNecesProdInfo.sProduto
        
        'Lê o produto para descobrir as unidades de medidas associadas
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 124205
    
        With objNecesProdInfo
            .iCompras = objProduto.iCompras
            .iClasseUM = objProduto.iClasseUM
            .sUMEstoque = objProduto.sSiglaUMEstoque
            .sUMCompra = objProduto.sSiglaUMCompra
        End With
        
        If objNecesProdInfo.sUMNecesInfo = "" Then
            objNecesProdInfo.sUMNecesInfo = objProduto.sSiglaUMEstoque
        End If
        
        'Descobre o fator de conversao da UM da necessidade informada p/UM de estoque do produto (raiz)
        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objNecesProdInfo.sUMNecesInfo, objProduto.sSiglaUMEstoque, dFatorConversao)
        If lErro <> SUCESSO Then gError 124206

        dQuantidadeProd = objNecesProdInfo.dQuantNecesInfo * dFatorConversao
        objNecesProdInfo.dQuantNecesCalcEst = dQuantidadeProd
        
        If objNecesProdInfo.sVersao = "" Then

            objKit.sProdutoRaiz = objNecesProdInfo.sProduto

            lErro = CF("Kit_Le_Padrao", objKit)
            If lErro <> SUCESSO And lErro <> 106304 Then gError 124207

            'Se não existe versão padrão p/ o Kit -> Erro.
            If lErro <> SUCESSO Then gError 124281
            
            objNecesProdInfo.sVersao = objKit.sVersao

        End If
        
        If objNecesProdInfo.objItemOP Is Nothing Then
        
            lErro = Producao_Calcula_Necessidades2(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0)
            If lErro <> SUCESSO Then gError 124208
        
        Else
        
            If objNecesProdInfo.objItemOP.colProdutoKitInfo.Count = 0 Then
            
                lErro = Producao_Calcula_Necessidades2(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0)
                If lErro <> SUCESSO Then gError 124208
                
            Else
            
                lErro = Producao_Calcula_Necessidades4(objCalcNecesProd, objNecesProdInfo, dQuantidadeProd, 0, objNecesProdInfo.objItemOP)
                If lErro <> SUCESSO Then gError 124208
            
            End If
        
        End If
            
        lErro = Producao_Calcula_Necessidades3(objCalcNecesProd, objNecesProdInfo)
        If lErro <> SUCESSO Then gError 124208
            
    Next
    
    'percorre a colecao de necessidades de compra e as coloca na unidade de compras
    For Each objNecesProdInfo In objCalcNecesProd.colNecesCalcProd
    
        objNecesProdInfo.dQuantNecesCalcEst = objNecesProdInfo.dQuantNecesCalc
        
        If objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL Then
        
            'Descobre o fator de conversao da UM de estoque para a de Compra
            lErro = CF("UM_Conversao_Trans", objNecesProdInfo.iClasseUM, objNecesProdInfo.sUMEstoque, objNecesProdInfo.sUMCompra, dFatorConversao)
            If lErro <> SUCESSO Then gError 124215
        
            'armazena a necessidade de compra na unidade de medida de compra
            objNecesProdInfo.dQuantNecesCalc = objNecesProdInfo.dQuantNecesCalcEst * dFatorConversao
                    
        End If
        
    Next
    
    Producao_Calcula_Necessidades = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades:

    Producao_Calcula_Necessidades = gErr
     
    Select Case gErr
          
        Case 124205 To 124208, 124215
        
        Case 124281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153379)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades2(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo, ByVal dQuantidadeProd As Double, ByVal iNiveis As Integer) As Long

Dim lErro As Long, objNecesProdInfoFilho As ClassNecesProdInfo
Dim objKit As New ClassKit, objNecesProdInfoAux As ClassNecesProdInfo
Dim objProdutoKit As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double, bAchou As Boolean
Dim dQuantidadePai As Double, dQuantidadeFilho As Double
Dim dFatorConversaoProdutoRaiz As Double
Dim dQtdSaldoLivre As Double

'##############################################
'Inserido por Wagner
Dim dTotalEmpenhado As Double
Dim dTotalEmpenhoAtual As Double
Dim dSaldoInicialLivre As Double
Dim bAchouPrimeiro As Boolean
Dim bAchouPrimeiro2 As Boolean
Dim objNecesProdInfoPrimeiro As ClassNecesProdInfo
Dim objNecesProdInfoAux2 As ClassNecesProdInfo
Dim objNecesProdInfoAux3 As ClassNecesProdInfo
'##############################################

On Error GoTo Erro_Producao_Calcula_Necessidades2

    'Se a quantidade já chegou a 0.1 ou já desceu na recursividade n instâncias
    If (dQuantidadeProd <= 0.0001 And iNiveis > 0) Or iNiveis >= 20 Then
    
        Producao_Calcula_Necessidades2 = SUCESSO
        Exit Function

    End If

    If iNiveis > 0 Then
    
        bAchou = False
        bAchouPrimeiro = False 'Inserido por Wagner
        For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd
        
            '############################################
            'Inserido por Wagner
            If Not bAchouPrimeiro And objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto Then
                bAchouPrimeiro = True
                Set objNecesProdInfoPrimeiro = objNecesProdInfoAux
            End If
            '############################################
            
            If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto And objNecesProdInfoAux.sVersao = objNecesProdInfo.sVersao Then
                
                bAchou = True
                
                If objNecesProdInfo.dtData1 <> DATA_NULA Then
                    If objNecesProdInfoAux.dtData1 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData1 = objNecesProdInfo.dtData1
                End If
                
                If objNecesProdInfo.dtData2 <> DATA_NULA Then
                    If objNecesProdInfoAux.dtData2 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData2 = objNecesProdInfo.dtData2
                End If
                
                Exit For
                
            End If
                          
        Next
        
        If bAchou = False Then
            
            If objNecesProdInfo.sVersao = "" And objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then

                objKit.sProdutoRaiz = objNecesProdInfo.sProduto

                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235

                'Se não existe versão padrão p/ o Kit -> Erro.
                If lErro <> SUCESSO Then gError 124282

                objNecesProdInfo.sVersao = objKit.sVersao

            End If
            
            Set objNecesProdInfoAux = New ClassNecesProdInfo
            
            With objNecesProdInfo
                objNecesProdInfoAux.sProduto = .sProduto
                objNecesProdInfoAux.sVersao = .sVersao
                objNecesProdInfoAux.iCompras = .iCompras
                objNecesProdInfoAux.iClasseUM = .iClasseUM
                objNecesProdInfoAux.sUMNecesInfo = ""
                objNecesProdInfoAux.sUMEstoque = .sUMEstoque
                objNecesProdInfoAux.sUMCompra = .sUMCompra
                objNecesProdInfoAux.dtData1 = .dtData1
                objNecesProdInfoAux.dtData2 = .dtData2
                objNecesProdInfoAux.lNumIntItemOP = .lNumIntItemOP
            End With
            
            'obter saldo "livre" de produto
            '############################################
            'Alterado por Wagner
            If Not bAchouPrimeiro Then
                lErro = Producao_Calcula_Necessidades3(objCalcNecesProd, objNecesProdInfoAux)
                If lErro <> SUCESSO Then gError 124209
            End If
            '############################################
            
            'incluir o produto na colecao
            objCalcNecesProd.colNecesCalcProd.Add objNecesProdInfoAux
            
        End If
                
        objNecesProdInfoAux.dQuantNecesCalc = objNecesProdInfoAux.dQuantNecesCalc + dQuantidadeProd
        
        '############################################
        'Alterado por Wagner
        If bAchouPrimeiro Then
            dQtdSaldoLivre = objNecesProdInfoPrimeiro.dQuantSaldoLivre
        Else
            dQtdSaldoLivre = objNecesProdInfoAux.dQuantSaldoLivre
        End If
        '############################################
        
        'se o saldo livre for >= necessario
        If dQtdSaldoLivre >= dQuantidadeProd Then
            
            'abater o necessario do saldo livre
            dQtdSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
            dQuantidadeProd = 0
                    
        Else
                        
            'abater da necessidade o saldo livre
            dQuantidadeProd = dQuantidadeProd - dQtdSaldoLivre
            dQtdSaldoLivre = 0
            
        End If
        
        '############################################
        'Alterado por Wagner/ Corrigido em 30/08/05
        'Se já existe entrada para esse produto
        If bAchouPrimeiro Then
            
            objNecesProdInfoAux.dSaldoDisponivel = objNecesProdInfoPrimeiro.dSaldoDisponivel
            objNecesProdInfoPrimeiro.dQuantSaldoLivre = dQtdSaldoLivre
        
            'Para cada necessidade já inserida
            For Each objNecesProdInfoAux3 In objCalcNecesProd.colNecesCalcProd
        
                bAchouPrimeiro2 = False
                dTotalEmpenhoAtual = 0
                dSaldoInicialLivre = 0
                dTotalEmpenhado = 0
        
                'Para cada outra necessidade
                For Each objNecesProdInfoAux2 In objCalcNecesProd.colNecesCalcProd
                
                    'Se é o mesmo produto
                    If objNecesProdInfoAux3.sProduto = objNecesProdInfoAux2.sProduto Then
                        
                        '############################################
                        'Inserido por Wagner
                        'Obtém o Empenho e Saldo Livre Inicial
                        If Not bAchouPrimeiro2 Then
                            bAchouPrimeiro2 = True
                            dTotalEmpenhoAtual = objNecesProdInfoAux2.dSaldoEmpenhosDiretos
                            dSaldoInicialLivre = objNecesProdInfoAux2.dQuantSaldoLivreInicial
                        End If
                        '############################################
                        
                        'Se já chegou no item de mesma versão, então sai da função
                        If objNecesProdInfoAux3.sVersao = objNecesProdInfoAux2.sVersao Then Exit For
                
                        'Acumula a quantidade que vai ser usada nos itens anteriores para ser
                        'somada ao empenho
                        dTotalEmpenhado = dTotalEmpenhado + objNecesProdInfoAux2.dQuantNecesCalc
                    
                    End If
                
                Next
        
                objNecesProdInfoAux3.dSaldoEmpenhosDiretos = dTotalEmpenhoAtual + dTotalEmpenhado
                           
                objNecesProdInfoAux3.dQuantSaldoLivreInicial = dSaldoInicialLivre - dTotalEmpenhado
            
                If objNecesProdInfoAux3.dQuantSaldoLivreInicial < 0 Then objNecesProdInfoAux3.dQuantSaldoLivreInicial = 0
       
            Next
                            
        Else
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre
        End If
        '############################################
        
    End If
    
    If dQuantidadeProd > 0 Then
        
        If objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
        
            objKit.sVersao = objNecesProdInfo.sVersao
            objKit.sProdutoRaiz = objNecesProdInfo.sProduto
        
            'tem que ler o proprio cara e os seus filhos
            lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
            If lErro <> SUCESSO And lErro <> 103203 Then gError 124210
            
            'Se o produto for um Kit, então leva a recursividade em diante
            If lErro = SUCESSO Then
            
                'percorre os elementos componentes do Kit Lido
                For Each objProdutoKit In objKit.colComponentes
                
                    objProduto.sCodigo = objProdutoKit.sProduto
                                
                    'Lê o produto para descobrir as unidades de medidas associadas
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 124211
                        
                    'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProdutoKit.sUnidadeMed, objProduto.sSiglaUMEstoque, dFatorConversao)
                    If lErro <> SUCESSO Then gError 124212
                
                    If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                        
                        dQuantidadePai = (dQuantidadeProd / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                        
                    Else
                        
                        Set objNecesProdInfoFilho = New ClassNecesProdInfo
                
                        '##############################################
                        'Inserido por Wagner 06/10/2005
                        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                            
                            If objProdutoKit.sVersaoKitComp = "" Then
                
                                objKit.sProdutoRaiz = objProdutoKit.sProduto
                
                                lErro = CF("Kit_Le_Padrao", objKit)
                                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
                
                                'Se não existe versão padrão p/ o Kit -> Erro.
                                If lErro <> SUCESSO Then gError 124282
                
                                objNecesProdInfoFilho.sVersao = objKit.sVersao
                
                            Else
                        
                                objNecesProdInfoFilho.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                
                            End If
                            
                        End If
                        '###############################################

                        With objNecesProdInfoFilho
                            .sProduto = objProdutoKit.sProduto
                            '.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                            .iCompras = objProduto.iCompras
                            .iClasseUM = objProduto.iClasseUM
                            .sUMEstoque = objProduto.sSiglaUMEstoque
                            .sUMCompra = objProduto.sSiglaUMCompra
                            If objNecesProdInfo.dtData1 <> DATA_NULA Then
                                .dtData2 = objNecesProdInfo.dtData1
                                .dtData1 = .dtData2 - objProduto.iTempoProducao
                            End If
                            .lNumIntItemOP = objNecesProdInfo.lNumIntItemOP
                        End With
                
                        'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                        dQuantidadeFilho = dQuantidadePai * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                          
                        'Calcula de maneira recursiva a quantidade de componentes necessários.
                        lErro = Producao_Calcula_Necessidades2(objCalcNecesProd, objNecesProdInfoFilho, dQuantidadeFilho, iNiveis + 1)
                        If lErro <> SUCESSO Then gError 124213
                
                    End If
                
                Next
            
            End If
    
        End If
        
    End If
    
    Producao_Calcula_Necessidades2 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades2:

    Producao_Calcula_Necessidades2 = gErr
     
    Select Case gErr
          
        Case 124209 To 124213, 124235
        
        Case 124282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153380)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades3(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo) As Long
'obtem do bd informacoes sobre saldo de produto

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objEmpenho As New ClassEmpenho
Dim dEmpenho As Double, dSaldoEmpenhos As Double
Dim objNecesProdInfoAux As ClassNecesProdInfo

On Error GoTo Erro_Producao_Calcula_Necessidades3

    '??? falta tratar outros flags referentes a qtde em pedidos de compra, requisicao,..
    
    If (objNecesProdInfo.iCompras = PRODUTO_COMPRAVEL And objCalcNecesProd.iSemSaldoProdComprado = 0) Or (objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL And objCalcNecesProd.iSemSaldoProdProduzido = 0) Then
    
        
        objEstoqueProduto.sProduto = objNecesProdInfo.sProduto
        dEmpenho = 0
        
        If objCalcNecesProd.iIgnoraEst = DESMARCADO Then
            'Lê as quantidades do produtos nos almoxs da filial
            lErro = CF("EstoqueProduto_Le_Todos_Almoxarifados_Filial", objEstoqueProduto, objCalcNecesProd.iFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 52979 Then gError 124214

            For Each objNecesProdInfoAux In objCalcNecesProd.colNecesInfProd
            
                objEmpenho.sProduto = objNecesProdInfo.sProduto
                objEmpenho.iFilialEmpresa = objCalcNecesProd.iFilialEmpresa
                objEmpenho.lNumIntDocItemOP = objNecesProdInfoAux.lNumIntItemOP
                
                lErro = CF("Empenho_Le_1", objEmpenho)
                If lErro <> SUCESSO And lErro <> 178592 Then gError 178594
                
                dEmpenho = dEmpenho + objEmpenho.dQuantidade
                           
            Next
            
        End If
        
        With objEstoqueProduto
        
            objNecesProdInfo.dQuantSaldoLivreInicial = .dQuantDisponivel + .dQuantRecIndl + .dQuantBenef3 + .dQuantConsig3
            
            If gobjMAT.iOPTipoCalcNecessidade = OP_TIPO_CALC_NECESS_USA_QUANTOP Then
                objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial + .dQuantOP
            End If
            
            objNecesProdInfo.dSaldoDisponivel = .dQuantDisponivel
            objNecesProdInfo.dSaldoEmpenhosDiretos = .dQuantEmpenhada - dEmpenho
            objNecesProdInfo.dSaldoReservado = .dQuantReservada
            
            'se for p/abater empenhos diretos
            If objCalcNecesProd.iSemSaldoEmpenhosDiretos = 0 Then
                If gobjMAT.iOPTipoCalcNecessidade <> OP_TIPO_CALC_NECESS_QUANTOP_ABATE_EMPENHO Then
                    objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial - objNecesProdInfo.dSaldoEmpenhosDiretos
                Else
                    'Só considero os empenhos que excedem a produção
                    dSaldoEmpenhos = objNecesProdInfo.dSaldoEmpenhosDiretos - .dQuantOP
                    If dSaldoEmpenhos < QTDE_ESTOQUE_DELTA Then dSaldoEmpenhos = 0 'Tudo que está empenhado já está em produção
                    objNecesProdInfo.dQuantSaldoLivreInicial = objNecesProdInfo.dQuantSaldoLivreInicial - dSaldoEmpenhos
                End If
            End If
            
        End With
    
    Else
    
        objNecesProdInfo.dQuantSaldoLivreInicial = 0
        
    End If
    
    If objNecesProdInfo.dQuantSaldoLivreInicial < 0 Then objNecesProdInfo.dQuantSaldoLivreInicial = 0
    objNecesProdInfo.dQuantSaldoLivre = objNecesProdInfo.dQuantSaldoLivreInicial
    
    Producao_Calcula_Necessidades3 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades3:

    Producao_Calcula_Necessidades3 = gErr
     
    Select Case gErr
          
        Case 124214, 178594
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153381)
     
    End Select
     
    Exit Function

End Function

    'Para cada produto filho direto:

            'verifica se o produto já está na colecao de saldos de produtos
    
            'se nao estiver
                
                'obter saldo "livre" lendo o bd e colocando-o na colecao.
                
            'se estiver
            
                'obter saldo "livre" da colecao
                
            'se o saldo livre for >= necessario
                
                'abater o necessario do saldo livre
                'nao vai produzir nada
                        
            'senao
                            
                'abater da necessidade o saldo livre
                'saldo livre = 0
                        
                'se produto produzido
                    '=> existe uma necessidade que vai demandar producao
                        'recursividade
            
                'senao (produto comprado)
                
                    'verifica se produto já está na colecao de produtos a comprar
    
                    'se estiver aumenta a qtde a comprar
                    
                    'senao
                        'inclui na colecao
                

Public Function Rel_Producao_Calcula_Necessidades(ByVal objCalcNecesProd As ClassCalcNecesProd) As Long
'imprime relatorio com necessidades de producao de subprodutos e compra de itens correspondentes
'à necessidades informadas

Dim lErro As Long, lNumIntRel As Long, lTransacao As Long, alComando(1 To 1) As Long
Dim objRelatorio As New AdmRelatorio, objNecesProdInfo As ClassNecesProdInfo, iIndice As Integer
Dim lSeq As Long

On Error GoTo Erro_Rel_Producao_Calcula_Necessidades

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 124216
    Next

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 124217

    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "PCPConfig", "NUM_PROX_REL_NECESSIDADEMATERIAL", lNumIntRel)
    If lErro <> SUCESSO Then gError 124218
    
    'percorre a colecao de necessidades informadas
    For Each objNecesProdInfo In objCalcNecesProd.colNecesInfProd
    
        lSeq = lSeq + 1
        
        With objNecesProdInfo
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelNecessidadeProd (NumIntRel, Tipo, Seq, Produto, Quantidade, UM, QuantidadeUMEst, SaldoLivreInicial, Disponivel, Empenhada, Versao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                lNumIntRel, 1, lSeq, .sProduto, .dQuantNecesInfo, .sUMNecesInfo, .dQuantNecesCalcEst, .dQuantSaldoLivreInicial, .dSaldoDisponivel, .dSaldoEmpenhosDiretos, .sVersao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 124219
        
    Next
    
    lSeq = 0
    
    'percorre a colecao de necessidades calculadas
    For Each objNecesProdInfo In objCalcNecesProd.colNecesCalcProd
    
        lSeq = lSeq + 1
        
        With objNecesProdInfo
            If .iCompras = PRODUTO_PRODUZIVEL Then
                lErro = Comando_Executar(alComando(1), "INSERT INTO RelNecessidadeProd (NumIntRel, Tipo, Seq, Produto, Quantidade, UM, QuantidadeUMEst, SaldoLivreInicial, Disponivel, Empenhada, Versao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, 2, lSeq, .sProduto, .dQuantNecesCalc, .sUMEstoque, .dQuantNecesCalc, .dQuantSaldoLivreInicial, .dSaldoDisponivel, .dSaldoEmpenhosDiretos, .sVersao)
            Else
                lErro = Comando_Executar(alComando(1), "INSERT INTO RelNecessidadeProd (NumIntRel, Tipo, Seq, Produto, Quantidade, UM, QuantidadeUMEst, SaldoLivreInicial, Disponivel, Empenhada, Versao) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, 3, lSeq, .sProduto, .dQuantNecesCalc, .sUMCompra, .dQuantNecesCalcEst, .dQuantSaldoLivreInicial, .dSaldoDisponivel, .dSaldoEmpenhosDiretos, .sVersao)
            End If
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 124220
        
    Next
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 124221

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    lErro = objRelatorio.ExecutarDireto("Necessidade de Produção e Compra", "", 0, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 124222
    
    Rel_Producao_Calcula_Necessidades = SUCESSO
     
    Exit Function
    
Erro_Rel_Producao_Calcula_Necessidades:

    Rel_Producao_Calcula_Necessidades = gErr
     
    Select Case gErr
          '??? colocar tratamento correto dos erros
        Case 124216 To 124222
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153382)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153383)
     
    End Select
     
    Exit Function

End Function

Private Function Producao_Calcula_Necessidades4(ByVal objCalcNecesProd As ClassCalcNecesProd, ByVal objNecesProdInfo As ClassNecesProdInfo, ByVal dQuantidadeProd As Double, ByVal iNiveis As Integer, ByVal objItemOP As ClassItemOP) As Long

Dim lErro As Long, objNecesProdInfoFilho As ClassNecesProdInfo
Dim objKit As New ClassKit, objNecesProdInfoAux As ClassNecesProdInfo
Dim objProdutoKit As New ClassProdutoKit
Dim objProduto As New ClassProduto
Dim dFatorConversao As Double, bAchou As Boolean
Dim dQuantidadePai As Double, dQuantidadeFilho As Double
Dim dFatorConversaoProdutoRaiz As Double
Dim dQtdSaldoLivre As Double

'##############################################
'Inserido por Wagner
Dim dTotalEmpenhado As Double
Dim dTotalEmpenhoAtual As Double
Dim dSaldoInicialLivre As Double
Dim bAchouPrimeiro As Boolean
Dim bAchouPrimeiro2 As Boolean
Dim objNecesProdInfoPrimeiro As ClassNecesProdInfo
Dim objNecesProdInfoAux2 As ClassNecesProdInfo
Dim objNecesProdInfoAux3 As ClassNecesProdInfo
Dim objProdutoKitInfo As ClassProdutoKitInfo
'##############################################

On Error GoTo Erro_Producao_Calcula_Necessidades4

    'Se a quantidade já chegou a 0.1 ou já desceu na recursividade n instâncias
    If (dQuantidadeProd <= 0.0001 And iNiveis > 0) Or iNiveis >= 20 Then
    
        Producao_Calcula_Necessidades4 = SUCESSO
        Exit Function

    End If

    If iNiveis > 0 Then
    
        '#############################################################
        'Inserido por Wagner 16/08/2006
        'Só inclui se ele tiver sido marcado para gerar requisição
        For Each objProdutoKitInfo In objItemOP.colProdutoKitInfo
        
            If objProdutoKitInfo.objProdutoKit.sProdutoRaiz = objNecesProdInfo.sProdutoRaiz And _
            objProdutoKitInfo.objProdutoKit.sVersao = objNecesProdInfo.sVersaoPai And _
            objProdutoKitInfo.objProdutoKit.iNivel = objNecesProdInfo.iNivel And _
            objProdutoKitInfo.objProdutoKit.iSeq = objNecesProdInfo.iSeq Then
                Exit For
            End If
        Next
        '#############################################################
    
        bAchou = False
        bAchouPrimeiro = False 'Inserido por Wagner
            
        For Each objNecesProdInfoAux In objCalcNecesProd.colNecesCalcProd
        
            '############################################
            'Inserido por Wagner
            If Not bAchouPrimeiro And objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto Then
                
                bAchouPrimeiro = True
                Set objNecesProdInfoPrimeiro = objNecesProdInfoAux
            End If
            '############################################
            
            If objNecesProdInfoAux.sProduto = objNecesProdInfo.sProduto And objNecesProdInfoAux.sVersao = objNecesProdInfo.sVersao Then
                
                bAchou = True
                
                If objNecesProdInfo.dtData1 <> DATA_NULA Then
                    If objNecesProdInfoAux.dtData1 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData1 = objNecesProdInfo.dtData1
                End If
                
                If objNecesProdInfo.dtData2 <> DATA_NULA Then
                    If objNecesProdInfoAux.dtData2 > objNecesProdInfo.dtData1 Then objNecesProdInfoAux.dtData2 = objNecesProdInfo.dtData2
                End If
                
                Exit For
                
            End If
                          
        Next
        
        If bAchou = False Then
            
            If objNecesProdInfo.sVersao = "" And objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then

                objKit.sProdutoRaiz = objNecesProdInfo.sProduto

                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235

                'Se não existe versão padrão p/ o Kit -> Erro.
                If lErro <> SUCESSO Then gError 124282

                objNecesProdInfo.sVersao = objKit.sVersao

            End If
            
            Set objNecesProdInfoAux = New ClassNecesProdInfo
            
            With objNecesProdInfo
                objNecesProdInfoAux.sProduto = .sProduto
                objNecesProdInfoAux.sVersao = .sVersao
                objNecesProdInfoAux.iCompras = .iCompras
                objNecesProdInfoAux.iClasseUM = .iClasseUM
                objNecesProdInfoAux.sUMNecesInfo = ""
                objNecesProdInfoAux.sUMEstoque = .sUMEstoque
                objNecesProdInfoAux.sUMCompra = .sUMCompra
                objNecesProdInfoAux.dtData1 = .dtData1
                objNecesProdInfoAux.dtData2 = .dtData2
                objNecesProdInfoAux.lNumIntItemOP = .lNumIntItemOP
            End With
            
            'obter saldo "livre" de produto
            '############################################
            'Alterado por Wagner
            If Not bAchouPrimeiro Then
                lErro = Producao_Calcula_Necessidades3(objCalcNecesProd, objNecesProdInfoAux)
                If lErro <> SUCESSO Then gError 124209
            End If
            '############################################
            
            '#############################################################
            'Alterado por Wagner 16/08/2006
            'Só inclui se ele tiver sido marcado para gerar requisição de compra ou for produzível
            If objProdutoKitInfo.iSelecionado = MARCADO Or objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                'incluir o produto na colecao
                objCalcNecesProd.colNecesCalcProd.Add objNecesProdInfoAux
            End If
            '#############################################################
            
        End If
                
        
        '#############################################################
        'Alterado por Wagner 16/08/2006
        'Só altera se ele tiver sido marcado para gerar requisição de compra ou for produzível
        If objProdutoKitInfo.iSelecionado = MARCADO Or objProdutoKitInfo.objProduto.iCompras = PRODUTO_PRODUZIVEL Then
            objNecesProdInfoAux.dQuantNecesCalc = objNecesProdInfoAux.dQuantNecesCalc + dQuantidadeProd
        End If
        
        '############################################
        'Alterado por Wagner
        If bAchouPrimeiro Then
            dQtdSaldoLivre = objNecesProdInfoPrimeiro.dQuantSaldoLivre
        Else
            dQtdSaldoLivre = objNecesProdInfoAux.dQuantSaldoLivre
        End If
        '############################################
        
        'se o saldo livre for >= necessario
        If dQtdSaldoLivre >= dQuantidadeProd Then
            
            'abater o necessario do saldo livre
            dQtdSaldoLivre = dQtdSaldoLivre - dQuantidadeProd
            dQuantidadeProd = 0
                    
        Else
                        
            'abater da necessidade o saldo livre
            dQuantidadeProd = dQuantidadeProd - dQtdSaldoLivre
            dQtdSaldoLivre = 0
            
        End If
        
        '############################################
        'Alterado por Wagner/ Corrigido em 30/08/05
        'Se já existe entrada para esse produto
        If bAchouPrimeiro Then
            
            objNecesProdInfoAux.dSaldoDisponivel = objNecesProdInfoPrimeiro.dSaldoDisponivel
            objNecesProdInfoPrimeiro.dQuantSaldoLivre = dQtdSaldoLivre
        
            'Para cada necessidade já inserida
            For Each objNecesProdInfoAux3 In objCalcNecesProd.colNecesCalcProd
        
                bAchouPrimeiro2 = False
                dTotalEmpenhoAtual = 0
                dSaldoInicialLivre = 0
                dTotalEmpenhado = 0
        
                'Para cada outra necessidade
                For Each objNecesProdInfoAux2 In objCalcNecesProd.colNecesCalcProd
                
                    'Se é o mesmo produto
                    If objNecesProdInfoAux3.sProduto = objNecesProdInfoAux2.sProduto Then
                        
                        '############################################
                        'Inserido por Wagner
                        'Obtém o Empenho e Saldo Livre Inicial
                        If Not bAchouPrimeiro2 Then
                            bAchouPrimeiro2 = True
                            dTotalEmpenhoAtual = objNecesProdInfoAux2.dSaldoEmpenhosDiretos
                            dSaldoInicialLivre = objNecesProdInfoAux2.dQuantSaldoLivreInicial
                        End If
                        '############################################
                        
                        'Se já chegou no item de mesma versão, então sai da função
                        If objNecesProdInfoAux3.sVersao = objNecesProdInfoAux2.sVersao Then Exit For
                
                        'Acumula a quantidade que vai ser usada nos itens anteriores para ser
                        'somada ao empenho
                        dTotalEmpenhado = dTotalEmpenhado + objNecesProdInfoAux2.dQuantNecesCalc
                    
                    End If
                
                Next
        
                objNecesProdInfoAux3.dSaldoEmpenhosDiretos = dTotalEmpenhoAtual + dTotalEmpenhado
                           
                objNecesProdInfoAux3.dQuantSaldoLivreInicial = dSaldoInicialLivre - dTotalEmpenhado
            
                If objNecesProdInfoAux3.dQuantSaldoLivreInicial < 0 Then objNecesProdInfoAux3.dQuantSaldoLivreInicial = 0
       
            Next
                            
        Else
            objNecesProdInfoAux.dQuantSaldoLivre = dQtdSaldoLivre
        End If
        '############################################
        
    End If
    
    If dQuantidadeProd > 0 Then
        
        'Alterado por Wagner 16/08/2006
        'Pode ter que gerar requisições de compra de insumos de kit de produtos comprados
'        If objNecesProdInfo.iCompras = PRODUTO_PRODUZIVEL Then
        
            objKit.sVersao = objNecesProdInfo.sVersao
            objKit.sProdutoRaiz = objNecesProdInfo.sProduto
        
            'tem que ler o proprio cara e os seus filhos
            lErro = CF("ProdutoKit_Le_PrimeiroNivel_Versao", objKit)
            If lErro <> SUCESSO And lErro <> 103203 Then gError 124210
            
            'Se o produto for um Kit, então leva a recursividade em diante
            If lErro = SUCESSO Then
            
                'percorre os elementos componentes do Kit Lido
                For Each objProdutoKit In objKit.colComponentes
                
                    objProduto.sCodigo = objProdutoKit.sProduto
                                
                    'Lê o produto para descobrir as unidades de medidas associadas
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 124211
                        
                    'Descobre o fator de conversao da UM recebida p/UM de estoque do produto
                    lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, objProdutoKit.sUnidadeMed, dFatorConversao)
                    If lErro <> SUCESSO Then gError 124212
                
                    If objProdutoKit.iNivel = KIT_NIVEL_RAIZ Then
                        
                        dQuantidadePai = (dQuantidadeProd / objProdutoKit.dQuantidade * dFatorConversao) / (1 - objProdutoKit.dPercentualPerda)
                        
                    Else
                        
                        Set objNecesProdInfoFilho = New ClassNecesProdInfo
                
                        '##############################################
                        'Inserido por Wagner 06/10/2005
                        If objProduto.iCompras = PRODUTO_PRODUZIVEL Then
                            
                            If objProdutoKit.sVersaoKitComp = "" Then
                
                                objKit.sProdutoRaiz = objProdutoKit.sProduto
                
                                lErro = CF("Kit_Le_Padrao", objKit)
                                If lErro <> SUCESSO And lErro <> 106304 Then gError 124235
                
                                'Se não existe versão padrão p/ o Kit -> Erro.
                                If lErro <> SUCESSO Then gError 124282
                
                                objNecesProdInfoFilho.sVersao = objKit.sVersao
                
                            Else
                        
                                objNecesProdInfoFilho.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                
                            End If
                            
                        End If
                        '###############################################

                        With objNecesProdInfoFilho
                            .sProduto = objProdutoKit.sProduto
                            '.sVersao = objProdutoKit.sVersaoKitComp 'Alterado por Wagner
                            .iCompras = objProduto.iCompras
                            .iClasseUM = objProduto.iClasseUM
                            .sUMEstoque = objProduto.sSiglaUMEstoque
                            .sUMCompra = objProduto.sSiglaUMCompra
                            If objNecesProdInfo.dtData1 <> DATA_NULA Then
                                .dtData2 = objNecesProdInfo.dtData1
                                .dtData1 = .dtData2 - objProduto.iTempoProducao
                            End If
                            .lNumIntItemOP = objNecesProdInfo.lNumIntItemOP
                            .sVersaoPai = objProdutoKit.sVersao
                            .sProdutoRaiz = objProdutoKit.sProdutoRaiz
                            .iSeq = objProdutoKit.iSeq
                            .iNivel = objProdutoKit.iNivel
                        End With
                                        
                
                
                        'Converte a quantidade na UM de venda para a quantidade equivalente na UM de estoque.
                        dQuantidadeFilho = dQuantidadePai * (objProdutoKit.dQuantidade / (1 - objProdutoKit.dPercentualPerda) * dFatorConversao)
                          
                        'Calcula de maneira recursiva a quantidade de componentes necessários.
                        lErro = Producao_Calcula_Necessidades4(objCalcNecesProd, objNecesProdInfoFilho, dQuantidadeFilho, iNiveis + 1, objItemOP)
                        If lErro <> SUCESSO Then gError 124213
                
                    End If
                
                Next
            
            End If
    
'        End If
        
    End If
    
    Producao_Calcula_Necessidades4 = SUCESSO
     
    Exit Function
    
Erro_Producao_Calcula_Necessidades4:

    Producao_Calcula_Necessidades4 = gErr
     
    Select Case gErr
          
        Case 124209 To 124213, 124235
        
        Case 124282
            lErro = Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PADRAO", gErr, objKit.sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153380)
     
    End Select
     
    Exit Function

End Function
