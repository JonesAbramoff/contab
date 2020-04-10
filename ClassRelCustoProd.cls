VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelCustoProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'criada para atender às chamadas de EXT_FUNC nos relatorios atraves de pesquisa, principalmente, nas tabela EstoqueMes
    
Private mvarobjEstoqueMes As ClassEstoqueMes
Private mvarcolEstoqueMesProduto As New Collection

Private sUltProdAdicaoMP As String
Private sUltCodOPAdicaoMP As String
Private dUltAdicaoMP As Double

Private sUltProdQuantReal As String
Private sUltCodOPQuantReal As String
Private dUltQuantReal As Double
Private dUltQuantPropProd As Double

Public Function ObterDifProducao(iFilialEmpresa As Integer, dtData As Date, dCustoMatPrim As Double, dDIF As Double) As Long
'Para obter o valor de gastos indiretos

Dim lErro As Long

On Error GoTo Erro_ObterDifProducao
    
    lErro = AtualizaEstoqueMes(iFilialEmpresa, dtData)
    If lErro <> SUCESSO Then gError 93753
        
    If mvarobjEstoqueMes.dCustoMatPrim = 0 Then
    
        dDIF = 0
    
    Else
      
        dDIF = Round(mvarobjEstoqueMes.dGastosIndiretos * dCustoMatPrim / mvarobjEstoqueMes.dCustoMatPrim, 4)
    
    End If
        
    ObterDifProducao = SUCESSO
    
    Exit Function
    
Erro_ObterDifProducao:

    ObterDifProducao = gErr
    
    Select Case gErr

        Case 93753
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153069)

    End Select

    Exit Function
    
End Function

Public Function ObterDdfProducao(iFilialEmpresa As Integer, dtData As Date, lHorasMaquina As Long, dDDF As Double) As Long
'Para obter o valor de gastos diretos

Dim lErro As Long

On Error GoTo Erro_ObterDdfProducao
    
    lErro = AtualizaEstoqueMes(iFilialEmpresa, dtData)
    If lErro <> SUCESSO Then gError 93756
        
    If mvarobjEstoqueMes.lHorasMaquina = 0 Then
    
        dDDF = 0
    
    Else
    
        dDDF = Round(mvarobjEstoqueMes.dGastosDiretos * lHorasMaquina / mvarobjEstoqueMes.lHorasMaquina, 4)
    
    End If
    
    ObterDdfProducao = SUCESSO
    
    Exit Function
    
Erro_ObterDdfProducao:

    ObterDdfProducao = gErr
    
    Select Case gErr

        Case 93756
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153070)

    End Select

    Exit Function
    
End Function

Private Function AtualizaEstoqueMes(iFilialEmpresa As Integer, dtData As Date) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizaEstoqueMes

    If mvarobjEstoqueMes Is Nothing Then Set mvarobjEstoqueMes = New ClassEstoqueMes
    
    If mvarobjEstoqueMes.iFilialEmpresa <> iFilialEmpresa Or mvarobjEstoqueMes.iAno <> Year(dtData) Or mvarobjEstoqueMes.iMes <> Month(dtData) Then
    
        mvarobjEstoqueMes.iFilialEmpresa = iFilialEmpresa
        mvarobjEstoqueMes.iAno = Year(dtData)
        mvarobjEstoqueMes.iMes = Month(dtData)
        
        lErro = CF("EstoqueMes_Le", mvarobjEstoqueMes)
        If lErro <> SUCESSO And lErro <> 36513 Then gError 93754
       
    End If

    AtualizaEstoqueMes = SUCESSO
    
    Exit Function
    
Erro_AtualizaEstoqueMes:

    AtualizaEstoqueMes = gErr
    
    Select Case gErr

        Case 93754
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153071)

    End Select

    Exit Function

End Function

Private Function AtualizaEstoqueMesProduto(ByVal iAno As Integer, ByVal iMes As Integer) As Long

Dim lErro As Long
Dim objEstoqueMesProduto As ClassEstoqueMesProduto

On Error GoTo Erro_AtualizaEstoqueMesProduto

    If mvarcolEstoqueMesProduto.Count > 0 Then
    
        Set objEstoqueMesProduto = mvarcolEstoqueMesProduto.Item(1)
        
        If objEstoqueMesProduto.iAno <> iAno Or objEstoqueMesProduto.iMes <> iMes Then
    
            Set mvarcolEstoqueMesProduto = New Collection
    
            'preenche uma colecao com os produtos que tiveram gastos informados e que portanto não terão seu calculo feito com os demais produtos
            lErro = CF("EstoqueMesProduto_Le", iAno, iMes, mvarcolEstoqueMesProduto)
            If lErro <> SUCESSO Then gError 92925
        
        End If
        
    Else
    
        'preenche uma colecao com os produtos que tiveram gastos informados e que portanto não terão seu calculo feito com os demais produtos
        lErro = CF("EstoqueMesProduto_Le", iAno, iMes, mvarcolEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 92926
        
    End If

    AtualizaEstoqueMesProduto = SUCESSO
    
    Exit Function
    
Erro_AtualizaEstoqueMesProduto:

    AtualizaEstoqueMesProduto = gErr
    
    Select Case gErr

        Case 92925, 92926
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153072)

    End Select

    Exit Function

End Function
        
Public Function ObterProduto_IndicacaoGastoInformado(iProdutoGastoInformado As Integer, iFilialEmpresa As Integer, sProduto As String, iAno As Integer, iMes As Integer) As Long
'Para obter o valor de gastos indiretos

Dim lErro As Long
Dim objEstoqueMesProduto As ClassEstoqueMesProduto

On Error GoTo Erro_ObterProduto_IndicacaoGastoInformado
    
    lErro = AtualizaEstoqueMesProduto(iAno, iMes)
    If lErro <> SUCESSO Then gError 92927
        
    iProdutoGastoInformado = 0
        
    For Each objEstoqueMesProduto In mvarcolEstoqueMesProduto
        If objEstoqueMesProduto.iFilialEmpresa = iFilialEmpresa And objEstoqueMesProduto.sProduto = sProduto Then
            iProdutoGastoInformado = 1
            Exit For
        End If
    Next
             
    ObterProduto_IndicacaoGastoInformado = SUCESSO
    
    Exit Function
    
Erro_ObterProduto_IndicacaoGastoInformado:

    ObterProduto_IndicacaoGastoInformado = gErr
    
    Select Case gErr

        Case 92927
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153073)

    End Select

    Exit Function
    
End Function

Public Function ObterProduto_ValorGasto(dValor As Double, iFilialEmpresa As Integer, sProduto As String, iAno As Integer, iMes As Integer, dQuantidade As Double, sSiglaUM As String) As Long
'Para obter o valor de gastos indiretos

Dim lErro As Long
Dim objEstoqueMesProduto As ClassEstoqueMesProduto
Dim iAchou As Integer
Dim objProduto As New ClassProduto
Dim dFator As Double

On Error GoTo Erro_ObterProduto_ValorGasto
    
    lErro = AtualizaEstoqueMesProduto(iAno, iMes)
    If lErro <> SUCESSO Then gError 92927
        
    For Each objEstoqueMesProduto In mvarcolEstoqueMesProduto
        If objEstoqueMesProduto.iFilialEmpresa = iFilialEmpresa And objEstoqueMesProduto.sProduto = sProduto Then
            iAchou = 1
            Exit For
        End If
    Next
             
    If iAchou = 0 Then
    
        dValor = 0
        
    Else
    
        objProduto.sCodigo = sProduto
    
        'Lê os dados do produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 92928
        
        'Se não encontrou o produto => erro
        If lErro = 28030 Then gError 92929
    
        'Converte a unidade de medida da quantidade pedida
        lErro = CF("UM_Conversao", objProduto.iClasseUM, sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 92930

        If objEstoqueMesProduto.dQuantidade > 0 Then
            dValor = (objEstoqueMesProduto.dGasto / objEstoqueMesProduto.dQuantidade) * dQuantidade * dFator
        Else
            dValor = 0
        End If
    End If
             
    ObterProduto_ValorGasto = SUCESSO
    
    Exit Function
    
Erro_ObterProduto_ValorGasto:

    ObterProduto_ValorGasto = gErr
    
    Select Case gErr

        Case 92927, 92928, 92930
        
        Case 92929
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153074)

    End Select

    Exit Function
    
End Function

Public Function Calcula_Participacao_Componente_Kit(ByVal sProdutoRaiz As String, ByVal sVersao As String, ByVal sComponente As String) As String
'Calcula o percentual de participação do componente para ProdutoRaiz/Versão e retorna.

Dim lErro As Long
Dim colItens As New Collection
Dim dQuantidadeKit As Double
Dim iIndice As Integer
Dim iPosComponente As Integer
Dim objProdutoKitProdutos As New ClassProdutoKitProdutos
Dim objProdutoKitProdutosTEMP As New ClassProdutoKitProdutos

On Error GoTo Erro_Calcula_Participacao_Componente_Kit

    objProdutoKitProdutos.sProdutoRaiz = sProdutoRaiz
    objProdutoKitProdutos.sVersao = sVersao
    
    'Obtem uma Colecao com os itens do kit para empenho (a embalagem do ProdutoRaiz será o ultimo elemento da colecao)
    lErro = CF("OrdemProducao_Le_Col_Empenho2", objProdutoKitProdutos, colItens, False)
    If lErro <> SUCESSO And lErro <> 106393 Then gError 108560
    
    'Se nao Encontrou => Erro
    If lErro = 106393 Then gError 108561
        
    For Each objProdutoKitProdutosTEMP In colItens
    
        iIndice = iIndice + 1
        
        'Guarda a posicao do componente na colecao
        If colItens.Item(iIndice).sProduto = sComponente Then iPosComponente = iIndice
    
        'Faz o Somatório das Quantidades do kit
        dQuantidadeKit = dQuantidadeKit + colItens.Item(iIndice).dQuantidade
        
    Next
            
    If iPosComponente <> 0 Then
        Calcula_Participacao_Componente_Kit = CStr(Round(100 * colItens.Item(iPosComponente).dQuantidade / dQuantidadeKit, 5))
    Else
        Calcula_Participacao_Componente_Kit = "0,00000"
    End If
    
    Exit Function
    
Erro_Calcula_Participacao_Componente_Kit:

    Calcula_Participacao_Componente_Kit = ""
    
    Select Case gErr

        Case 108560
        
        Case 108561
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, sProdutoRaiz)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153075)

    End Select

End Function

Public Function Calcula_Rendimento_ProdutoVersao(ByVal sProduto As String, ByVal sVersao As String) As String
'Calcula a média dos últimos 3 tempos de processo do produto/Versão na Entrada de Material Produzido
'Retorna em Minutos/Máquina

Dim lErro As Long
Dim sSQL As String
Dim lComando As Long
Dim lHorasMaquina As Long
Dim iMaxEntradas As Integer
Dim lTotalHorasMaquina As Long

On Error GoTo Erro_Calcula_Rendimento_ProdutoVersao

    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108565
    
    'Guarda a Expressao SQL
    sSQL = "SELECT SUM(HorasMaquina) FROM ItensOrdemProducaoBaixadas ItensOrdemProducao, MovimentoEstoque, OrdensDeProducaoBaixadas OrdensDeProducao WHERE " & _
                "OrdensDeProducao.Codigo = ItensOrdemProducao.Codigo AND OrdensDeProducao.FilialEmpresa = ItensOrdemProducao.FilialEmpresa " & _
                "AND ItensOrdemProducao.NumIntDoc = MovimentoEstoque.NumIntDocOrigem and MovimentoEstoque.TipoMov IN (?,?) AND " & _
                "ItensOrdemProducao.Produto=? AND ItensOrdemProducao.Versao=? GROUP by MovimentoEstoque.Data, MovimentoEstoque.Codigo ORDER BY MovimentoEstoque.Data DESC"
    
    'Prepara a Execução do Select
    lErro = Comando_Executar(lComando, sSQL, lHorasMaquina, MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3, sProduto, sVersao)
    If lErro <> SUCESSO Then gError 108566
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108567
    
    'Se não encontrou => Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 108568
    
    'Enquanto tiver dados e não exceder o limite de 3 entradas
    Do While lErro <> AD_SQL_SEM_DADOS And iMaxEntradas < 3
    
        'incrementa o Número de registros encontrados
        iMaxEntradas = iMaxEntradas + 1
        
        'Acumula o total de Horas
        lTotalHorasMaquina = lTotalHorasMaquina + lHorasMaquina
        
        'Busca o Próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108569

    Loop
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Calcula_Rendimento_ProdutoVersao = CStr(Round(lTotalHorasMaquina / iMaxEntradas, 2))
    
    Exit Function
    
Erro_Calcula_Rendimento_ProdutoVersao:

    Calcula_Rendimento_ProdutoVersao = ""
    
    Select Case gErr

        Case 108565
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 108566, 108567, 108569
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE", gErr)
        
        Case 108568 'Sem Dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153076)

    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Public Function Obter_AdicaoMP(ByVal sProduto As String, ByVal sCodOP As String) As String
'Calcula a quantidade real necessária de insumos de um kit.

Dim lErro As Long
Dim dAdicao As Double
Dim iIndice As Integer
Dim bEhComponente As Boolean
Dim bEhAgua As Boolean, bAchou As Boolean
Dim colItensOP As New Collection
Dim colItensMovEst As New Collection
Dim objProdKitProdOP As New ClassProdutoKitProdutos
Dim objProdKitProdCol As New ClassProdutoKitProdutos
Dim objProdKitProdMatReq As New ClassProdutoKitProdutos

On Error GoTo Erro_Obter_AdicaoMP

    If sProduto = sUltProdAdicaoMP And sCodOP = sUltCodOPAdicaoMP Then
    
        dAdicao = dUltAdicaoMP
    
    Else
    
        'Monta uma colecao com os insumos teoricos (previstos) para producao do Produto/OP
        lErro = CF("Obter_AdicaoMP_ProdOP", sProduto, sCodOP, colItensOP)
        If lErro <> SUCESSO And lErro <> 108632 Then gError 108610
        
        If lErro = 108632 Then gError 108635
        
        'Monta uma colecao com o Material Requisitado (abatido do devolvido)
        lErro = CF("Obter_AdicaoMP_MatReq", sProduto, sCodOP, colItensMovEst)
        If lErro <> SUCESSO And lErro <> 108603 Then gError 108611
        
        If lErro = 108603 Then gError 108636
        
        bEhComponente = False
            
        'Verifica na Coleção que contém os componentes do Kit se o Produto passado é componente dele mesmo
        For Each objProdKitProdCol In colItensOP
            
            'Se o Produto faz parte dele mesmo => pode contar como adição de MP
            If objProdKitProdCol.sProduto = sProduto Then
                bEhComponente = True
                Exit For
            End If
        
        Next
                
        'Para cada elemento da Colecao de Material Requisitado
        For Each objProdKitProdMatReq In colItensMovEst
            
            'Verifica se o produto é Agua
            lErro = CF("Verifica_Produto_Agua", objProdKitProdMatReq.sProduto, bEhAgua)
            If lErro <> SUCESSO Then gError 108630
            
            'Se o produto não é Agua e é componente dele mesmo (kit)
            If bEhAgua = False Then
            
                'se o produto requisitado nao é o produzido ou ele é componente de si mesmo
                If objProdKitProdMatReq.sProduto <> sProduto Or bEhComponente = True Then
                
                    bAchou = False
                
                    'Para cada componente do Kit
                    For Each objProdKitProdOP In colItensOP
        
                        'Se o Produto for igual
                        If objProdKitProdOP.sProduto = objProdKitProdMatReq.sProduto Then
                        
                            'se tem diferença
                            If Abs(objProdKitProdMatReq.dQuantidade - objProdKitProdOP.dQuantidade) > QTDE_ESTOQUE_DELTA Then
                            
                                'Acumula a Quantidade que foi adicionada
                                dAdicao = dAdicao + (objProdKitProdMatReq.dQuantidade - objProdKitProdOP.dQuantidade)
                            
                            End If
                            
                            bAchou = True
                            Exit For
                            
                        End If
                    
                    Next
                    
                    If bAchou = False Then
                    
                        dAdicao = dAdicao + objProdKitProdMatReq.dQuantidade
                        
                    End If
                
                End If
                
            End If
        
        Next

        'guarda os valores p/nao ter que recalcular
        sUltProdAdicaoMP = sProduto
        sUltCodOPAdicaoMP = sCodOP
        dUltAdicaoMP = dAdicao
    
    End If
    
    Obter_AdicaoMP = Replace(Round(dAdicao, 2), ",", ".")
    
    Exit Function
    
Erro_Obter_AdicaoMP:

    Select Case gErr

        Case 108630
        
        Case 108635, 108636
            'Obter_AdicaoMP = "Não Produzido"
        
        Case 108610, 108611
            Obter_AdicaoMP = "UM INCOMPATÍVEL"
        
        Case 108561
            Call Rotina_Erro(vbOKOnly, "ERRO_KIT_SEM_PRIMEIRO_NIVEL", gErr, sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153077)

    End Select
    
End Function

'??? se houver versao em pcpselect: verificar a inicializacao de sSiglaUM, verificar outras modificacoes
Public Function Obter_QuantidadeReal(ByVal sProduto As String, ByVal sOPBaixada As String) As Double
'Retorna o Somatório da quantidade de material produzido para o produto/op baixada

Dim lErro As Long
Dim dQuantidadeReal As Double, dQuantidadePropProd As Double

On Error GoTo Erro_Obter_QuantidadeReal
    
    If sProduto = sUltProdQuantReal And sOPBaixada = sUltCodOPQuantReal Then
    
        dQuantidadeReal = dUltQuantReal
    
    Else
    
        lErro = Obter_QuantidadeRealAux(sProduto, sOPBaixada, dQuantidadeReal, dQuantidadePropProd)
        If lErro <> SUCESSO Then gError 106867
    
        'guarda os valores p/nao ter que recalcular
        sUltProdQuantReal = sProduto
        sUltCodOPQuantReal = sOPBaixada
        dUltQuantReal = dQuantidadeReal
        dUltQuantPropProd = dQuantidadePropProd
    
    End If
    
    Obter_QuantidadeReal = dQuantidadeReal
    
    Exit Function
    
Erro_Obter_QuantidadeReal:

    Obter_QuantidadeReal = 0
    
    Select Case gErr
    
        Case 106867
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153078)
    
    End Select
    
End Function

Public Function Excluir_PrevVendaPrevConsumo_Aux(ByVal lNumIntDoc As Long) As Long
'Exclui da tabela PrevVendaPrevConsumo os registros com o numintdoc passado

Dim lErro As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim alComando(1) As Long
Dim lNumIntDocTEMP As Long

On Error GoTo Erro_Excluir_PrevVendaPrevConsumo_Aux

    'Abertura da Transacao
    lTransacao = Transacao_Abrir
    If lTransacao = 0 Then gError 108677
    
    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 108672
    Next
    
    'Prepara para a execução do Select
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM PrevVendaPrevConsumo WHERE NumIntDoc = ?", 0, lNumIntDocTEMP, lNumIntDoc)
    If lErro <> SUCESSO Then gError 108673
    
    'Busca o Primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108674
    
    'Enquanto encontrar registro...
    Do While lErro <> AD_SQL_SEM_DADOS
    
        'Apaga da tabela prevvendaprevconsumo o registro com o numintdoc passado
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM PrevVendaPrevConsumo", alComando(0))
        If lErro <> SUCESSO Then gError 108675
        
        'Busca o próximo registro da tabela
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108676

    Loop
    
    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 108678
        
    Excluir_PrevVendaPrevConsumo_Aux = SUCESSO
    
    Exit Function
    
Erro_Excluir_PrevVendaPrevConsumo_Aux:

    Excluir_PrevVendaPrevConsumo_Aux = gErr

    Select Case gErr
    
        Case 108672
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 108673
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_PREVVENDAPREVCONSUMO", gErr)
        
        Case 108674, 108676
        
        Case 108675
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PREVVENDAPREVCONSUMO", gErr)
            
        Case 108677
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 108678
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153079)
    
    End Select

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Retorna os dados
    Call Transacao_Rollback

End Function

'??? se houver versao em pcpselect: verificar a inicializacao de sSiglaUM, verificar outras modificacoes
Public Function Obter_QuantidadeRealAux(ByVal sProduto As String, ByVal sOPBaixada As String, dQuantidadeReal As Double, dQuantidadePropProd As Double) As Long
'Retorna o Somatório da quantidade de material produzido para o produto/op baixada

Dim lErro As Long
Dim sSiglaUM As String, sUMDest As String
Dim lComando As Long, lComando2 As Long
Dim dFatorConv As Double, iNivel As Integer
Dim dQuantidade As Double
Dim objProduto As New ClassProduto, iTipoMov As Integer

On Error GoTo Erro_Obter_QuantidadeRealAux
    
    sUMDest = "KG"
    
    'Abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 108600
    
    'Abertura do Comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 108600
    
    sSiglaUM = String(STRING_UM_SIGLA, 0)
    
    'Prepara a Execucao do Comando SQL
    lErro = Comando_Executar(lComando, "SELECT TipoMov, Quantidade, SiglaUM FROM MatReqProdOP WHERE ProdutoRaiz=? AND FilialEmpresa=? AND Codigo=? AND Produto = ?", iTipoMov, dQuantidade, sSiglaUM, sProduto, giFilialEmpresa, sOPBaixada, sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 108601
    
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108602
    
    'Se Encontrou
    If lErro = AD_SQL_SUCESSO Then
    
        'Enquanto tiver dados => Preenche a colecao
        While lErro = AD_SQL_SUCESSO
            
            If sSiglaUM <> sUMDest Then
            
                'Le o Produto
                objProduto.sCodigo = sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 108604
                
                'Se nao encontrou => Erro
                If lErro = 28030 Then gError 108605
            
                'Transforma produto em kg
                'Faz a Conversao de unidades de medida
                lErro = CF("UM_Conversao_Rel", objProduto.iClasseUM, sSiglaUM, sUMDest, dFatorConv)
                If lErro <> SUCESSO Then gError 108609
                
                dQuantidade = dQuantidade * dFatorConv
                
            End If
            
            'Adiciona à quantidade já existente
            Select Case iTipoMov
            
                Case MOV_EST_PRODUCAO, MOV_EST_PRODUCAO_BENEF3
                    dQuantidadeReal = dQuantidadeReal + dQuantidade
                    
                Case Else 'requisicao do proprio produto produzido
                        
                    '??? isto nao deveria ser testado aqui por questao de performance
                    lErro = Comando_Executar(lComando2, "SELECT Nivel FROM ItensOrdemProducaoBaixadas, ProdutoKit WHERE ItensOrdemProducaoBaixadas.FilialEmpresa = ? AND ItensOrdemProducaoBaixadas.Codigo = ? AND ItensOrdemProducaoBaixadas.Produto = ? AND ItensOrdemProducaoBaixadas.Produto = ProdutoKit.ProdutoRaiz AND ItensOrdemProducaoBaixadas.Versao = ProdutoKit.Versao AND ProdutoKit.ProdutoRaiz = ProdutoKit.Produto AND ProdutoKit.Nivel <> 0", _
                        iNivel, giFilialEmpresa, sOPBaixada, sProduto)
                    If lErro <> AD_SQL_SUCESSO Then gError 106531
                        
                    lErro = Comando_BuscarProximo(lComando2)
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106532
                    
                    'se o produto nao é componente dele mesmo
                    If lErro = AD_SQL_SEM_DADOS Then
                        dQuantidadeReal = dQuantidadeReal - dQuantidade
                        dQuantidadePropProd = dQuantidadePropProd + dQuantidade
                    End If
                        
            End Select
            
            'Busca o Próximo registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 108612
        
        Wend
    
    End If
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)

    Obter_QuantidadeRealAux = SUCESSO
    
    Exit Function
    
Erro_Obter_QuantidadeRealAux:

    Obter_QuantidadeRealAux = gErr
    
    Select Case gErr
    
        Case 108600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 108603, 106531, 106532
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSOP", gErr)
            
        Case 108605
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr)
            
        Case 108601, 108602, 108604, 108609, 108612
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153080)
    
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
End Function

'??? se houver versao em pcpselect: verificar a inicializacao de sSiglaUM, verificar outras modificacoes
Public Function Obter_QuantidadePropProd(ByVal sProduto As String, ByVal sOPBaixada As String) As Double
'Retorna o Somatório da quantidade de material produzido para o produto/op baixada

Dim lErro As Long
Dim dQuantidadeReal As Double, dQuantidadePropProd As Double

On Error GoTo Erro_Obter_QuantidadePropProd
    
    If sProduto = sUltProdQuantReal And sOPBaixada = sUltCodOPQuantReal Then
    
        dQuantidadePropProd = dUltQuantPropProd
    
    Else
    
        lErro = Obter_QuantidadeRealAux(sProduto, sOPBaixada, dQuantidadeReal, dQuantidadePropProd)
        If lErro <> SUCESSO Then gError 106867
    
        'guarda os valores p/nao ter que recalcular
        sUltProdQuantReal = sProduto
        sUltCodOPQuantReal = sOPBaixada
        dUltQuantReal = dQuantidadeReal
        dUltQuantPropProd = dQuantidadePropProd
        
    End If
    
    Obter_QuantidadePropProd = dQuantidadePropProd
    
    Exit Function
    
Erro_Obter_QuantidadePropProd:

    Obter_QuantidadePropProd = 0
    
    Select Case gErr
    
        Case 106867
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153081)
    
    End Select
    
End Function


