VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMnemoComisCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Indica que é para obter o endereço da filialcliente
Const FILIALCLIENTE_OBTEM_ENDERECO = 1

'Variaveis Globais, usadas para armazenar dados de acessos ao BD
Private sUltEstadoDestino As String
Private objUltICMSAliquota As Object
Private colTabelaPrecoItem As New Collection

Function Calcula_Mnemonico_ComissoesCust(ByVal objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux, ByVal objMnemonicoValor As ClassMnemonicoValor) As Long
'Verifica qual eh o mnemonico passado e chama a funcao apropriada para calcula-los
'objMnemonicoValor RECEBE(Input) o mnemônico que será calculado e RETORNA(Output) o valor desse mnemônico

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Calcula_Mnemonico_ComissoesCust

    'Faz uma selecao pelo mnemonico
    Select Case objMnemonicoValor.sMnemonico
    
        'Se for RegiaoVenda
        Case MNEMONICO_COMISSOES_PRECO_TABELA_INPAL
            
            'Calcula o valor do mnemonico RegiaoVenda
            lErro = Calcula_MnemonicoComissoes_PrecoTabelaInpal(objMnemonicoComissCalcAux.objTela, objMnemonicoValor)
            If lErro <> SUCESSO Then gError 102015
        
        'Se for PrecoBase
        Case MNEMONICO_COMISSOES_PRECO_BASE_INPAL
        '*** incluido tratamento para telas de nf simples e nf simples a partir de pedido (o dessa ultima se encontra encapsulado na funcao de obter Preco Base)... (nao tem condpagto)
        
            'se a tela for de nota fiscal simples
            If UCase(TypeName(objMnemonicoComissCalcAux.objTela)) = UCase(NOME_CLASSE_CTNFISCAL) Then
            
                'Calcula o valor do mnemonico Produto como preco unitario...
                lErro = objMnemonicoComissCalcAux.objTela.objMnemonicoComissCalc.Calcula_MnemonicoComissoes_PrecoVenda(objMnemonicoValor)
                If lErro <> SUCESSO Then gError 102017
            
            Else
                                                
                'Calcula o valor do mnemonico PrecoBase
                lErro = Calcula_MnemonicoComissoes_PrecoBaseInpal(objMnemonicoComissCalcAux.objTela, objMnemonicoValor)
                If lErro <> SUCESSO Then gError 101780
            
            End If
            
        Case Else
        
            Calcula_Mnemonico_ComissoesCust = MNEMONICOCOMISSOES_NAO_ENCONTRADO
            
            Exit Function
    
    End Select
        
    Calcula_Mnemonico_ComissoesCust = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico_ComissoesCust:
    
    Calcula_Mnemonico_ComissoesCust = gErr
    
    Select Case gErr
    
        Case 101757, 101770, 101771, 101772, 101757, 102015
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

Private Function Calcula_MnemonicoComissoes_PrecoTabelaInpal(ByVal objTela As Object, objMnemonicoValor As ClassMnemonicoValor) As Long
'Devolve uma coleção com o preço de tabela (conforme o modelo da Inpal) para cada produto do grid de Itens
'objMnemonicoValor RETORNA(Output) os preços de tabela de cada produto

Dim iLinha As Long
Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim objTabelaPrecoItem As ClassTabelaPrecoItem
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Calcula_MnemonicoComissoes_PrecoTabelaInpal

    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Guarda o produto da linha atual
        sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'Se o produto está preenchido
        If Len(Trim(sProduto)) > 0 Then
        
            'Cria um novo objTabelaPrecoItem na memória
            Set objTabelaPrecoItem = New ClassTabelaPrecoItem
            
            'formata o produto para que o mesmo possa ser usado como filtro
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 101639
        
            'coloca o produto formatado no obj
            objTabelaPrecoItem.sCodProduto = sProdutoFormatado
        
            'Chama as funções que obtêm o preço de tabela da Inpal
            lErro = Obtem_PrecoTabelaInpal(objTela, objTabelaPrecoItem)
            If lErro <> SUCESSO Then gError 102014
            
            'Guarda o preço encontrado na coleção
            objMnemonicoValor.colValor.Add objTabelaPrecoItem.dPreco
        
        'Senão
        Else
        
            'Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
            
        End If
    
    Next
    
    Calcula_MnemonicoComissoes_PrecoTabelaInpal = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_PrecoTabelaInpal:

    Calcula_MnemonicoComissoes_PrecoTabelaInpal = gErr
    
    Select Case gErr
    
        Case 102014, 101639
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function

End Function

Private Function Obtem_PrecoTabelaInpal(ByVal objTela As Object, objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Obtém os parâmetros necessário para encontrar o preço do produto
'objTabelaPrecoItem RECEBE(Input) o código do produto que terá o preço obtio
'objTabelaPrecoItem RETORNA(Output) os dados lidos

Dim lErro As Long
Dim dtDataVigencia As Date
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As New ClassEndereco
Dim objICMSAliquota As Object
Dim objTabelaPreco As New ClassTabelaPreco
Dim objTabelaPrecoItemAux As New ClassTabelaPrecoItem
Dim bAchou As Boolean

On Error GoTo Erro_Obtem_PrecoTabelaInpal

    'Obtém a alíquota de ICMS para a operação
    lErro = Obtem_ICMSAliquota(objTela, objICMSAliquota)
    If lErro <> SUCESSO Then gError 102010

    Select Case TypeName(objICMSAliquota)
    
        'Se for a classe ClassEstado
        Case NOME_CLASSE_ESTADO
            
            objTabelaPreco.dAliquotaICMS = objICMSAliquota.dICMSAliquotaInterna
            
        'Se for a classe ClassICMSAliqExterna
        Case NOME_CLASSE_ICMSALIQEXTERNA
                    
            objTabelaPreco.dAliquotaICMS = objICMSAliquota.dAliquota
    
    End Select
    
    'Obtém a tabela de preço que está vinculada à alíquota encontrada acima
    lErro = CF("TabelaPreco_Le_AliqICMS", objTabelaPreco)
    If lErro <> SUCESSO And lErro <> 102001 Then gError 102002
    
    'Obtém a data de vigência máxima para opreço que será obtido
    lErro = Obtem_TabelaPreco_DataVigencia(objTela, objTela.Name, dtDataVigencia)
    If lErro <> SUCESSO Then gError 102011
    
    'Guarda no objTabelaPrecoItem os dados necessários para obter o preço de venda do produto
    objTabelaPrecoItem.iCodTabela = objTabelaPreco.iCodigo
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    objTabelaPrecoItem.dtDataVigencia = dtDataVigencia
    
    'Procura o preço do produto na coleção
    For Each objTabelaPrecoItemAux In colTabelaPrecoItem
    
        'Se todos os parâmetros permanecem iguais
        If (objTabelaPrecoItem.iCodTabela = objTabelaPrecoItemAux.iCodTabela) And (objTabelaPrecoItem.sCodProduto = objTabelaPrecoItemAux.sCodProduto) And _
        (objTabelaPrecoItem.iFilialEmpresa = objTabelaPrecoItemAux.iFilialEmpresa) And (objTabelaPrecoItem.dtDataVigencia = objTabelaPrecoItemAux.dtDataVigencia) Then
        
            'Instancia objTabelaPrecoItem com os dados da última leitura feita para tais parâmetros
            Set objTabelaPrecoItem = objTabelaPrecoItemAux
            
            'Indica que já encontrou o produto na coleção
            bAchou = True
            
            'Finaliza a procura do item na coleção, pois o mesmo já foi encontrado
            Exit For
        
        End If
    
    Next
        
    'Se não encontrou os dados na coleção
    If Not bAchou Then
    
        'Obtém no BD o preço de venda do produto de acordo com os parâmetros passados
        lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem, dtDataVigencia)
        If lErro <> SUCESSO And lErro <> 28014 Then gError 102012
        
        'Guarda o obj na coleção global
        colTabelaPrecoItem.Add objTabelaPrecoItem
    
    End If
    
    Obtem_PrecoTabelaInpal = SUCESSO
    
    Exit Function
    
Erro_Obtem_PrecoTabelaInpal:

    Obtem_PrecoTabelaInpal = gErr
    
    Select Case gErr
    
        Case 102002, 102010, 102011, 102012
        
        Case 102013
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE4", gErr, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

Private Function Obtem_ICMSAliquota(ByVal objTela As Object, objICMSAliquota As Object) As Long

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais
Dim objFilialCliente As New ClassFilialCliente
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Obtem_ICMSAliquota

    'Obtém os dados da filialempresa ativa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then gError 102007
    
    'Obtém os dados da filialcliente selecionada na tela
    lErro = objTela.objMnemonicoComissCalc.Obtem_FilialCliente_Dados(objFilialCliente, FILIALCLIENTE_OBTEM_ENDERECO, objEndereco)
    If lErro <> SUCESSO Then gError 94990
    
    'Se o estado destino mudou
    If objEndereco.sSiglaEstado <> sUltEstadoDestino Then
    
        'Se o estado da filial empresa é diferente do estado da filial cliente, ou seja,
        'se o estado origem é diferente do estado destino...
        If objFilialEmpresa.objEndereco.sSiglaEstado <> objEndereco.sSiglaEstado Then
        
            '... significa que a alíquota de icms é entre operações externas, ou seja, interestaduais,
            'portanto, será feita uma leitura na tabela ICMSAliqExterna
            
            'Instancia objICMSAliquota como sendo da classe ICMSAliqExterna
            Set objICMSAliquota = New ClassICMSAliqExterna
            
            'Guarda no obj as siglas dos estados de origem e destino
            objICMSAliquota.sSiglaOrig = objEndereco.sSiglaEstado
            objICMSAliquota.sSiglaDest = objFilialEmpresa.objEndereco.sSiglaEstado
            
            'Lê a alíquota de ICMS em ICMSALiqExternas para os estados origem e destino passados
            lErro = CF("ICMSAliqExternas_Le_Origem_Destino", objICMSAliquota)
            If lErro <> SUCESSO And lErro <> 94994 Then gError 94995
            
            'Se não encontrou => erro
            If lErro = 94994 Then gError 94996
            
            'Atualiza a sigla do último estado destino lido
            sUltEstadoDestino = objICMSAliquota.sSiglaDest
        
        'Senão, ou seja, se origem e destino são iguais
        Else
        
            '... significa que a alíquota de icms é para uma operação intraestadual,
            'portanto, será feita uma leitura na tabela Estados
            
            'Instancia objICMSAliquota como sendo da classe Estado
            Set objICMSAliquota = New ClassEstado
            
            'Guarda no obj a sigla do Estado que será lido
            objICMSAliquota.sSigla = objFilialEmpresa.objEndereco.sSiglaEstado
            
            'Lê a alíquota de ICMS para o Estado passado
            lErro = CF("Estado_Le", objICMSAliquota)
            If lErro <> SUCESSO And lErro <> 28485 Then gError 102008
            
            'Se não encontrou o estado => erro
            If lErro = 28485 Then gError 102009
            
            'Atualiza a sigla do último estado destino lido
            sUltEstadoDestino = objICMSAliquota.sSigla
            
        End If

        'Instancia objUltICMSAliquota com os dados da última alíquota lida
        Set objUltICMSAliquota = objICMSAliquota
    
    'Senão, ou seja, se é o mesmo estado
    Else
    
        'Instancia objICMSAliquota com os dados da última alíquota lida
        Set objICMSAliquota = objUltICMSAliquota
        
    End If
    
    Obtem_ICMSAliquota = SUCESSO
    
    Exit Function
    
Erro_Obtem_ICMSAliquota:
    
    Obtem_ICMSAliquota = gErr
    
    Select Case gErr
    
        Case 94990, 94995, 102007, 102008
        
        Case 94996
            Call Rotina_Erro(vbOKOnly, "ERRO_ALIQUOTA_ICMS_NAO_ENCONTRADA", gErr, objICMSAliquota.sSiglaOrig, objICMSAliquota.sSiglaDest)
    
        Case 102009
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", gErr, objICMSAliquota.sSigla)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function
    
End Function

Private Function Obtem_TabelaPreco_DataVigencia(ByVal objTela As Object, ByVal sNomeTela As String, dtDataVigencia As Date) As Long

Dim lErro As Long
Dim objPedidoVenda As ClassPedidoDeVenda

On Error GoTo Erro_Obtem_TabelaPreco_DataVigencia

    'Faz uma seleção pelo nome da tela passado como parâmetro
    Select Case sNomeTela
    
        'Se é uma tela de pedido de venda, ou de nota fiscal não vinculada a pedido
        Case NOME_TELA_PEDIDOVENDA, NOME_TELA_NFISCAL, NOME_TELA_NFISCALFATURA, NOME_TELA_COMISSOESCALCULA
        
            'A data de vigência é a data de emissão do documento
            dtDataVigencia = StrParaDate(objTela.Controls("DataEmissao").Text)
        
        'Senão, ou seja, se é uma nota fiscal vinculada a pedido
        Case NOME_TELA_NFISCALPEDIDO, NOME_TELA_NFISCALFATURAPEDIDO
        
            'Instancia objPedidoVenda
            Set objPedidoVenda = New ClassPedidoDeVenda
            
            'A data de vigência é a data de emissão do documento que está vinculado à nota
            lErro = objTela.objMnemonicoComissCalc.Obtem_PedidoVenda_Dados(objPedidoVenda)
            If lErro <> SUCESSO Then gError 102005
            
            'Guarda a data de emissão do pedido como data de vigência
            dtDataVigencia = objPedidoVenda.dtDataEmissao
        
        Case Else
        
            gError 102006
    
    End Select
            
        
    Obtem_TabelaPreco_DataVigencia = SUCESSO
    
    Exit Function
    
Erro_Obtem_TabelaPreco_DataVigencia:

    Obtem_TabelaPreco_DataVigencia = gErr
    
    Select Case gErr
    
        Case 102005
        
        Case 102006
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVIGENCIA_FALTA_TRATAMENTO", gErr, sNomeTela)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function
    
End Function

'Alterado por Luiz em 30/05/03
Private Function Calcula_MnemonicoComissoes_PrecoBaseInpal(ByVal objTela As Object, objMnemonicoValor As ClassMnemonicoValor, Optional iCondPagto As Integer) As Long
'Devolve uma coleção com o preço base para cada produto do grid de Itens
'objMnemonicoValor RETORNA(Output) os preços base de cada produto

Dim iLinha As Long
Dim lErro As Long
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dPrecoBase As Double
Dim X As New ClassFatGlobal
Dim sCliente As String
Dim objProduto As New ClassProduto
Dim sUM As String
Dim dFator As Double

On Error GoTo Erro_Calcula_MnemonicoComissoes_PrecoBaseInpal

    'Para cada linha no grid de itens
    For iLinha = 1 To objTela.objGridItens.iLinhasExistentes
    
        'Guarda o produto da linha atual
        sProduto = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_Produto_Col)
        
        'Se o produto está preenchido
        If Len(Trim(sProduto)) > 0 Then
        
            'formata o produto para que o mesmo possa ser usado como filtro
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 101759
        
            'Guarda o código do produto no obj
            objProduto.sCodigo = sProdutoFormatado
            
            'Obtém os dados do produto
            lErro = objTela.objMnemonicoComissCalc.Produto_Obtem_Dados(objProduto)
            If lErro <> SUCESSO Then gError 102233
            
            'Guarda a U.M. do produto da linha atual
            sUM = objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_UnidadeMed_Col)
            
            'Encontra o fator de conversão da UM que está no grid para a UM base de venda
            lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError 102234
        
            'guarda o preco unitario q eh parametro de input e output da funcao q sera chamada na prox. linha
            dPrecoBase = StrParaDbl(objTela.Controls("GridItens").TextMatrix(iLinha, objTela.iGrid_PrecoUnitario_Col)) / dFator
        
            If TypeName(objTela.Controls("Cliente")) = "Label" Then
                sCliente = objTela.Controls("Cliente").Caption
            Else
                sCliente = objTela.Controls("Cliente").Text
            End If
            
            'obter precobase do produto
            lErro = Obtem_PrecoBaseInpal(objTela, sProdutoFormatado, sCliente, dPrecoBase)
            If lErro <> SUCESSO Then gError 101758
                    
            'formata o preco
            dPrecoBase = Format(dPrecoBase, X.gobjFAT.sFormatoPrecoUnitario)
            
            '??? Péssimo... onde vc está guardando alguma informação em dPrecoBase ?
            'Guarda o preço encontrado na coleção
            objMnemonicoValor.colValor.Add dPrecoBase
                       
        'Senão
        Else
        
            'Guarda uma string vazia na coleção, indicando que o mnemônico não foi calculado
            objMnemonicoValor.colValor.Add 0
            
        End If
    
    Next
    
    Calcula_MnemonicoComissoes_PrecoBaseInpal = SUCESSO
    
    Exit Function
    
Erro_Calcula_MnemonicoComissoes_PrecoBaseInpal:

    Calcula_MnemonicoComissoes_PrecoBaseInpal = gErr
    
    Select Case gErr
    
        Case 101639, 101758, 101759, 102014, 102233, 102234
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Function

End Function

Private Sub Class_Terminate()

On Error GoTo Erro_Class_Terminate

    'Libera a memória que estava alocada para os objetos e coleções globais à classe
    Set objUltICMSAliquota = Nothing
    Set colTabelaPrecoItem = Nothing

    Exit Sub
    
Erro_Class_Terminate:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Exit Sub

End Sub

Private Function Obtem_PrecoBaseInpal(ByVal objTela As Object, ByVal sProduto As String, ByVal sClienteNomeRed, dPrecoBase As Double) As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim dPrecoBaseAux As Double, dDias As Double, dTaxaFinDia As Double
Dim iCondPagto As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Obtem_PrecoBaseInpal

    'trabalhar com var. auxiliar....
    dPrecoBaseAux = dPrecoBase

    'obtem os dados do cliente
    lErro = objTela.objMnemonicoComissCalc.Cliente_Obtem_Dados(objCliente)
    If lErro <> SUCESSO Then gError 101760
                   
    'se a tela for de nfiscalpedido
    If UCase(TypeName(objTela)) = UCase(NOME_CLASSE_CTNFISCALPEDIDO) Then

        'move a chave do pedido para a memoria
        objPedidoVenda.lCodigo = objTela.glCodPedido
        objPedidoVenda.iFilialEmpresa = objTela.giFilialPedido
    
        'Busca o pedido na tabela de Pedidos de Venda
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then gError 101775
            
        'obtem os dados da condpagto
        lErro = objTela.objMnemonicoComissCalc.CondPagto_Obtem_Dados(objCondicaoPagto, objPedidoVenda.iCondicaoPagto)
        If lErro <> SUCESSO Then gError 101776

    Else
    
        'obtem os dados da condpagto
        lErro = objTela.objMnemonicoComissCalc.CondPagto_Obtem_Dados(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 101765
        
    End If
        
    'Calcula a média ponderada de dias para a taxa financeira
    Call objTela.objMnemonicoComissCalc.CalculaDias_CondPagamento(objCondicaoPagto, dDias)
        
'    'Calcula a taxa financeira diária
'    'Call objTela.objMnemonicoComissCalc.Calcula_TaxaFinanceira(dDias, StrParaDbl(objTela.Controls("PercAcrescFin").Text) / 100, dTaxaFinDia)
'    Call objTela.objMnemonicoComissCalc.Calcula_TaxaFinanceira(dDias, gobjCRFAT.dTaxaFinanceira, dTaxaFinDia)
                            
    'Recalcula o Preço Unitário a partir do Preço Base
    dPrecoBaseAux = dPrecoBaseAux / ((1 + gobjCRFAT.dTaxaFinanceira) ^ dDias)
        
    'retorna o valor de dprecobase
    dPrecoBase = dPrecoBaseAux
    
    'somente para versoes de debug... tulio 02/08/02
    'Call MsgBox(dPrecoBase, vbOKOnly)
    
    Obtem_PrecoBaseInpal = SUCESSO
    
    Exit Function
    
Erro_Obtem_PrecoBaseInpal:
    
    Obtem_PrecoBaseInpal = gErr
    
    Select Case gErr
    
        Case 101760, 101765, 101775, 101776
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select
    
    Exit Function
    
End Function



