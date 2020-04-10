Attribute VB_Name = "ReprocMovEst"
Option Explicit

'Número de selects que são feitos em cada fase do reprocessamento
Const NUM_SELECT_REPROC_DESFAZ = 1
Const NUM_SELECT_REPROC_REFAZ = 3
Const NUM_SELECT_REPROC_TESTAINTEGRIDADE = 1

'Identificador do select a ser executado
Const REPROCESSAMENTO_SELECT_DESFAZ1 = 31
Const REPROCESSAMENTO_SELECT_REFAZ1 = 41
Const REPROCESSAMENTO_SELECT_REFAZ2 = 42
Const REPROCESSAMENTO_SELECT_REFAZ3 = 43
Const REPROCESSAMENTO_SELECT_TESTAINTEGRIDADE1 = 71
Const REPROCESSAMENTO_SELECT_TESTAINTEGRIDADE2 = 72

'Caracteres que identificam o tipo de uma variável (ex.: d -> double; s-> string, etc.)
Const ID_VARIAVEL_INTEGER As String = "i"
Const ID_VARIAVEL_LONG As String = "l"
Const ID_VARIAVEL_DOUBLE As String = "d"
Const ID_VARIAVEL_DATE As String = "dt"
Const ID_VARIAVEL_STRING As String = "s"

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function Rotina_Reprocessamento_MovEstoque_Int(objReprocessamentoEst As ClassReprocessamentoEST) As Long
'reprocessa os movimentos de estoque entre a data inicio e a data do movimento mais recente

' *** Função alterada em 02/08/2001 por Luiz Gustavo de Freitas Nogueira ***
' *** A alteração foi feita para gravar a data em que se iniciou o último reprocessamento efetuado ***

' *** Função alterada em 07/08/2001 por Luiz Gustavo de Freitas Nogueira ***
' *** Alteração 1: a função passou a receber um obj como parâmetro ***
' *** Alteração 2: a fase desfaz do reprocessamento pode ser pulada, caso tenha sido configurado assim na tela que o dispara ***

Dim lTransacao As Long
Dim lErro As Long
Dim objMATConfig As New ClassMATConfig
Dim alComando(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim iIndice As Integer
Dim objEstoqueMes As New ClassEstoqueMes
Dim vbMsgRes As VbMsgBoxResult
Dim colExercicio As New Collection
Dim alComando1(1 To 4) As Long
Dim lTotalMovEstoque As Long
Dim lErro1 As Long
Dim lNumIntDoc As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Rotina_Reprocessamento_MovEstoque_Int


    'Inicia a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 83576
    
    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83577
    Next

    'libera comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 83641
    Next
    
    If objReprocessamentoEst.iAcertaEstProd = REPROCESSAMENTO_ACERTA_ESTPROD Then
        lErro = Acerta_EstoqueProduto_QuantDispNossa()
        If lErro <> SUCESSO Then gError 1
    Else
        
        ' *** Incluído em 02/08/2001 por Luiz Gustavo de Freitas Nogueira ***
        '*** Grava em MATConfig a data de início do último Reprocessamento feito ***
            
        'Se não foi passado parâmetro de produto e data final, e não for teste de integridade
        If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) = 0 And (objReprocessamentoEst.dtDataFim = DATA_NULA) Then
            
            'Guarda no obj os parâmetros que serão utilizados para efetuar a gravação
            objMATConfig.sCodigo = DATAINICIO_ULTIMO_REPROC
            objMATConfig.iFilialEmpresa = EMPRESA_TODA
            objMATConfig.sDescricao = DATAINICIO_ULTIMO_REPROC_DESCR
            objMATConfig.iTipo = CONFIG_TIPO_DATA
            objMATConfig.sConteudo = objReprocessamentoEst.dtDataInicio
            
            'Grava o registro em MATConfig
            lErro = CF("MATConfig_Grava_Trans", objMATConfig)
            If lErro <> SUCESSO Then gError 90584
        
        End If
        '*** Fim do trecho incluído em 02/08/2001 ***
        
        For Each objFiliais In gcolFiliais
        
            'verificar se os meses que envolvem o reprocessamentos estão fechados, se estiverem perguntar o que fazer.
            objEstoqueMes.iFilialEmpresa = objFiliais.iCodFilial
            objEstoqueMes.iAno = Year(objReprocessamentoEst.dtDataInicio)
            objEstoqueMes.iMes = Month(objReprocessamentoEst.dtDataInicio)
        
            'verifica se o mes relativo a data inicio dos movimentos que serão reprocessados está fechado
            lErro = CF("EstoqueMes_Le_Lock", alComando1(1), objEstoqueMes)
            If lErro <> SUCESSO And lErro <> 41774 Then gError 83763
        
            'faz o lock para impedir qualquer movimentação do estoque durate o reprocessamento
            objMATConfig.iFilialEmpresa = objFiliais.iCodFilial
            objMATConfig.sCodigo = DATA_REPROCESSAMENTO
        
            lErro = CF("MATConfig_Le_Lock1", objMATConfig, alComando1(3))
            If lErro <> SUCESSO And lErro <> 83776 Then gError 83772
        
            If lErro = SUCESSO Then
            
                'verifica se a data de reprocessamento é menor que a data de inicio do reprocessamento
                If CDate(objMATConfig.sConteudo) < objReprocessamentoEst.dtDataInicio Then gError 83777
                
                'Se não foi passado parâmetro de produto e data final, e não for teste de integridade
                If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) = 0 And (objReprocessamentoEst.dtDataFim = DATA_NULA) Then
                
                    lErro = Comando_ExecutarPos(alComando1(4), "DELETE FROM MATConfig", alComando1(3))
                    If lErro <> AD_SQL_SUCESSO Then gError 83780
                
                End If
                
            End If
    
        Next
    
        'faz o lock para impedir qualquer movimentação do estoque durate o reprocessamento
        objMATConfig.iFilialEmpresa = EMPRESA_TODA
        objMATConfig.sCodigo = NUM_PROX_ITEM_MOV_ESTOQUE
    
        lErro = CF("MATConfig_Le_Lock", objMATConfig, alComando(11))
        If lErro <> SUCESSO Then gError 83544
    
        lNumIntDoc = CLng(objMATConfig.sConteudo)
        
        'le os movimentos de estoque em ordem descendente de data e numintdoc até a data passada como parametro
        lErro = Comando_Executar(alComando1(2), "SELECT Count(*) FROM MovimentoEstoque WHERE Data >= ?", lTotalMovEstoque, objReprocessamentoEst.dtDataInicio)
        If lErro <> AD_SQL_SUCESSO Then gError 83562
    
        lErro = Comando_BuscarPrimeiro(alComando1(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83563
    
        'Tela acompanhamento Batch inicializa dValorTotal
        TelaAcompanhaBatchEST.dValorTotal = 2 * lTotalMovEstoque
    
        'desativa os locks dos comandos a seguir
        lErro = Conexao_DesativarLocks(DESATIVAR_LOCKS)
        If lErro <> SUCESSO Then gError 89449
        
        'Se é para pular a fase desfaz do reprocessamento
        'Significa que é necessário zerar os saldos das tabelas envolvidas no reprocessamento
        If objReprocessamentoEst.iPulaFaseDesfaz = REPROCESSAMENTO_PULA_DESFAZ Then
        
            lErro = Rotina_Reproc_Zera_Saldos(objReprocessamentoEst)
            If lErro <> SUCESSO Then gError 90603
            
        'Senão
        Else
        
            'Executa a fase desfaz
            lErro = Rotina_Reproc_Desfaz(alComando(), objReprocessamentoEst.iFilialEmpresa, objReprocessamentoEst.dtDataInicio, colExercicio, objReprocessamentoEst)
            If lErro <> SUCESSO Then gError 83766
            
        End If
        
        lErro = Rotina_Reproc_Refaz(alComando(), colExercicio, lNumIntDoc, objReprocessamentoEst)
        If lErro <> SUCESSO Then gError 83767
    
        'reativa os locks
        lErro = Conexao_DesativarLocks(REATIVAR_LOCKS)
        If lErro <> SUCESSO Then gError 89450
    
        lErro = Comando_ExecutarPos(alComando(12), "UPDATE MATConfig SET Conteudo=?", alComando(11), CStr(lNumIntDoc))
        If lErro <> AD_SQL_SUCESSO Then gError 83654

    End If

    'libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'libera comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    If objReprocessamentoEst.iPulaFaseDesfaz = REPROCESSAMENTO_PULA_DESFAZ Then
        lErro = CF("Reprocessamento_Ajuste_InvCliForn", objReprocessamentoEst.sProdutoCodigo)
        If lErro <> SUCESSO Then gError 83767
    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 83768
    
    'Alteracao Daniel em 07/05/02
    Call Rotina_Aviso(vbOKOnly, "AVISO_REPROCESSAMENTO_EXECUTADO_SUCESSO")
    
    Rotina_Reprocessamento_MovEstoque_Int = SUCESSO

    Exit Function

Erro_Rotina_Reprocessamento_MovEstoque_Int:

    Rotina_Reprocessamento_MovEstoque_Int = gErr

    Select Case gErr

        Case 83544, 83763, 83765, 83766, 83767, 83772, 90584, 90603
        
        Case 83576
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
          
       Case 83577, 83641
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 83764
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
             
        Case 83768
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 83777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_REPROC_MENOR_DATA_INICIO", gErr, CDate(objMATConfig.sConteudo), objReprocessamentoEst.dtDataInicio)
             
        Case 83780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MATCONFIG", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
             
        Case 89449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
             
        Case 89450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECI4DO_PELO_VB", gErr, Error)

    End Select

    'reativa os locks
    Call Conexao_DesativarLocks(REATIVAR_LOCKS)

    'Rollback
    Call Transacao_Rollback

   'Fechamento comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'libera comandos
    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
    
    Exit Function

End Function

Private Sub Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque As typeItemMovEstoque, objItemMovEst As ClassItemMovEstoque)
    
    objItemMovEst.lNumIntDoc = tItemMovEstoque.lNumIntDoc
    objItemMovEst.iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    objItemMovEst.lCodigo = tItemMovEstoque.lCodigo
    objItemMovEst.dCusto = tItemMovEstoque.dCusto
    objItemMovEst.iApropriacao = tItemMovEstoque.iApropriacao
    objItemMovEst.sProduto = tItemMovEstoque.sProduto
    objItemMovEst.sSiglaUM = tItemMovEstoque.sSiglaUM
    objItemMovEst.dQuantidade = tItemMovEstoque.dQuantidade
    objItemMovEst.iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
    objItemMovEst.iTipoMov = tItemMovEstoque.iTipoMov
    objItemMovEst.lNumIntDocOrigem = tItemMovEstoque.lNumIntDocOrigem
    objItemMovEst.iTipoNumIntDocOrigem = tItemMovEstoque.iTipoNumIntDocOrigem
    objItemMovEst.dtData = tItemMovEstoque.dtData
    objItemMovEst.sCcl = tItemMovEstoque.sCcl
    objItemMovEst.lNumIntDocEst = tItemMovEstoque.lNumIntDocEst
    objItemMovEst.lCliente = tItemMovEstoque.lCliente
    objItemMovEst.lFornecedor = tItemMovEstoque.lFornecedor
    objItemMovEst.sOPCodigo = tItemMovEstoque.sOPCodigo
    objItemMovEst.sDocOrigem = tItemMovEstoque.sDocOrigem
    objItemMovEst.sContaContabilEst = tItemMovEstoque.sContaContabilEst
    objItemMovEst.sContaContabilAplic = tItemMovEstoque.sContaContabilAplic
    objItemMovEst.lHorasMaquina = tItemMovEstoque.lHorasMaquina
    objItemMovEst.dtDataInicioProducao = tItemMovEstoque.dtDataInicioProducao

End Sub

Private Function Rotina_Reproc_Desfaz(alComando() As Long, ByVal iFilialEmpresa As Integer, ByVal dtDataInicio As Date, colExercicio As Collection, objReprocessamentoEst As ClassReprocessamentoEST) As Long
'reaplica os movimentos de estoque cadastrados, desfazendo na ordem inversa em que foram cadastrados os seus efeitos

Dim lErro As Long
Dim tItemMovEstoque As typeItemMovEstoque
Dim tItemMovEstoqueVar As typeItemMovEstoqueVariant
Dim objItemMovEst As New ClassItemMovEstoque
Dim alComando1(1 To 3) As Long
Dim iIndice As Integer
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim iOrigemLcto As Integer
Dim lNumIntDocOrigemCTB As Long
Dim iOperacao As Integer
Dim iIndiceSelect As Integer
Dim asComandoSelect(NUM_SELECT_REPROC_DESFAZ) As String

On Error GoTo Erro_Rotina_Reproc_Desfaz

    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 83600
    Next

    lErro = Rotina_Reproc_MontaSelect(objReprocessamentoEst, asComandoSelect, REPROCESSAMENTO_DESFAZ)
    If lErro <> SUCESSO Then gError 94500
    
    lErro = Rotina_Reproc_ExecutaSelect_Comum(alComando1(1), objReprocessamentoEst, asComandoSelect(1), tItemMovEstoque, tItemMovEstoqueVar, REPROCESSAMENTO_SELECT_DESFAZ1)
    If lErro <> SUCESSO Then gError 94501
    
    Do While lErro = AD_SQL_SUCESSO
    
        'Atualiza tela de acompanhamento do Batch
        lErro = Rotina_Reproc_AtualizaTelaBatch()
        If lErro <> SUCESSO Then gError 83771
    
        Set objItemMovEst = New ClassItemMovEstoque
    
        'move os dados de tItemMovEstoque para objItemMovEst
        Call Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque, objItemMovEst)
    
        'desfaz o movimento de estoque
        lErro = CF("Estoque_Reprocessamento", alComando(), objItemMovEst, REPROCESSAMENTO_DESFAZ, Nothing)
        If lErro <> SUCESSO Then gError 83564
        
        objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
        
        'ler os dados referentes ao tipo de movimento
        lErro = CF("TiposMovEst_Le1", alComando1(3), objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 83602
        
        'se for um ajuste do custo standard
        If objTipoMovEstoque.sEntradaOuSaida = TIPOMOV_EST_AJUSTE_CUSTO_STANDARD Then
        
            'devolve o documento que originou o movimento de estoque.
            'Utilizado para descobrir os lançamentos contábeis associados e reprocessá-los.
            lErro = Retorna_Origem_Estoque_Contab(objItemMovEst.iTipoNumIntDocOrigem, objItemMovEst.lNumIntDocOrigem, objItemMovEst.lNumIntDoc, iOrigemLcto, lNumIntDocOrigemCTB)
            If lErro <> SUCESSO Then gError 83603
    
            'exclui as contabilizações associados ao ajuste do custo standard
            lErro = CF("Rotina_Reproc_Exclui_Lanc", objItemMovEst.iFilialEmpresa, iOrigemLcto, lNumIntDocOrigemCTB, colExercicio)
            If lErro <> SUCESSO Then gError 83604
        
            'exclui o movimento de estoque
            lErro = Comando_ExecutarPos(alComando1(2), "DELETE FROM MovimentoEstoque", alComando1(1))
            If lErro <> AD_SQL_SUCESSO Then gError 83601
        
        End If
        
        lErro = Comando_BuscarProximo(alComando1(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83565
        
        Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
        
    Loop

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Rotina_Reproc_Desfaz = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Desfaz:

    Rotina_Reproc_Desfaz = gErr
    
    Select Case gErr
    
        Case 83562, 83563, 83565
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, iFilialEmpresa)
    
        Case 83564, 83602, 83603, 83604, 83771, 94500, 94501

        Case 83600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 83601
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MOVIMENTOESTOQUE", gErr, objItemMovEst.iFilialEmpresa, objItemMovEst.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173766)

    End Select

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Exit Function

End Function

Private Function Rotina_Reproc_Refaz(alComando() As Long, colExercicio As Collection, lNumIntDoc As Long, objReprocessamentoEst As ClassReprocessamentoEST) As Long
'reaplica os movimentos de estoque cadastrados, refazendo na ordem em que foram lançados os seus efeitos

' *** Função alterada em 09/08/01 por Luiz G.F. Nogueira ***
' *** Alteração 1: a função passou a receber um obj no lugar de alguns parâmetros ***
' *** Alteração 2: o select dos movimentos que serão reprocessados é feito dinamicamente,
' pois existem parâmetros nesses select (Produto, DataFinal), que são opcionais. Assim,
' foram criadas funções para montar e para executar o select ***
'*************************************************************

Dim lErro As Long
Dim tItemMovEstoque As typeItemMovEstoque
Dim tItemMovEstoqueVar As typeItemMovEstoqueVariant
Dim lComando As Long
Dim dtDataCstStd As Date 'contem a ultimo data em que foi processado o ajuste do custo std
Dim iMesCstStd As Integer 'contem o ultimo mes em que foi processado o ajuste do custo std
Dim iAnoCstStd As Integer 'contem o ultimo ano em que foi processado o ajuste do custo std
Dim objTela As Object
Dim colItemMovEstoque As Collection
Dim alComando1(1 To 4) As Long
Dim iIndice As Integer
Dim iAnoUltFechamento As Integer 'contem o ultimo ano em que foi fechado o estoque
Dim objTipoMovEstoque As New ClassTipoMovEst
Dim lCodigo As Long
Dim objItemMovEst As ClassItemMovEstoque
Dim asSelect(1 To NUM_SELECT_REPROC_REFAZ) As String
Dim iIndiceSelect As Integer
Dim iOperacao As Integer
Dim iFilialEmpresa As Integer
Dim colApropriacaoInsumo As Collection
Dim colEstoqueMesProduto As Collection
Dim dQtdCalc As Double, dCustoCalc As Double

'*** Para depurar o reprocessamento, usando o BatchEst como .dll, o trecho abaixo deve estar descomentado
'Dim colProdutosErro As New Collection
'Dim iPulaProdutoErro As Integer
'Dim iIndiceLuiz As Integer
'Dim lCodigoLuiz As Long
'Dim vProduto As Variant
' ***
Dim objFiliaisReprocAno As AdmFiliais

On Error GoTo Erro_Rotina_Reproc_Refaz

    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 83663
    Next

    dtDataCstStd = DATA_NULA
    
    'cria uma nova instancia da tela de Custos sem exibi-la. Utilizado na contabilização do ajuste do custo standard
    'comentado por mario em 31/10/01 pois no batch não pode instanciar tela. Tem que achar outra solução.
'    lErro = Chama_Tela_Nova_Instancia1("Custos", objTela)
'    If lErro <> SUCESSO Then gError 83650
    
    'Monta os três selects que poderão ser utilizados
    'Os selects 1 e 3 serão utilizados quando os movimentos forem ordenados por entrada / saída
    'O select 2 será utilizado quando os movimentos forem ordenados por Data, Hora e NumIntDoc
    lErro = Rotina_Reproc_MontaSelect(objReprocessamentoEst, asSelect, REPROCESSAMENTO_REFAZ)
    If lErro <> SUCESSO Then gError 90649
    
    'Verifica se serão reprocessados primeiro os movimentos de entrada
    If objReprocessamentoEst.iOrdemReproc = REPROCESSAMENTO_ORDENA_ENTRADAS Then
    
        ' *** Select 1 ***
        'O select aqui executado, lê o NumIntDoc dos movimentos de estoque, ordenando primeiro as entradas e depois as saídas
        'Esse select deveria ser feito através um ExecutarPos, mas como ele deve ser feito em cima de duas tabelas
        'Foi substituído por um Executar, e mais embaixo (Select 3), no início do looping é feito um select em MovimentoEstoque, usando o NumIntDoc que foi lido aqui

        'Executa o select
        lErro = Rotina_Reproc_ExecutaSelect_Refaz1(alComando1(3), objReprocessamentoEst, asSelect(1), tItemMovEstoque, tItemMovEstoqueVar)
        If lErro <> SUCESSO Then gError 90702
        
    
    'Senão
    'significa que os movimentos devem ser reprocessados por ordem de hora em que foram registrados
    Else
        
        ' *** Select 2 ***
        'Lê os movimentos de estoque ordenados por Data, Hora e NumIntDoc
        
        'Executa o select
        lErro = Rotina_Reproc_ExecutaSelect_Comum(alComando1(1), objReprocessamentoEst, asSelect(2), tItemMovEstoque, tItemMovEstoqueVar, REPROCESSAMENTO_SELECT_REFAZ2)
        If lErro <> SUCESSO Then gError 90703
    
    End If
    
    lCodigo = tItemMovEstoque.lCodigo
    iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
    Set colItemMovEstoque = New Collection
    
    'iAnoUltFechamento = Year(tItemMovEstoque.dtData)
    iAnoUltFechamento = Year(objReprocessamentoEst.dtDataInicio)

    Do While lErro = AD_SQL_SUCESSO
    
        'Se serão reprocessados primeiro os movimentos de entrada
        If objReprocessamentoEst.iOrdemReproc = REPROCESSAMENTO_ORDENA_ENTRADAS Then
            
            ' *** Select 3 ***
            'Lê o movimento de estoque com o NumIntDoc obtido no select feito acima (Select 1)
            
            'Executa o select
            lErro = Rotina_Reproc_ExecutaSelect_Comum(alComando1(4), objReprocessamentoEst, asSelect(3), tItemMovEstoque, tItemMovEstoqueVar, REPROCESSAMENTO_SELECT_REFAZ3)
            If lErro <> SUCESSO Then gError 90704
            
            If iAnoUltFechamento = 0 Then iAnoUltFechamento = Year(tItemMovEstoque.dtData)
            
            '***
'Retirar o código comentado abaixo
'            'Inicializa as variáveis strings que serão utilizadas no select
'            tItemMovEstoque.sProduto = String(STRING_PRODUTO, 0)
'            tItemMovEstoque.sSiglaUM = String(STRING_UM_SIGLA, 0)
'            tItemMovEstoque.sCcl = String(STRING_CCL, 0)
'            tItemMovEstoque.sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
'            tItemMovEstoque.sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
'            tItemMovEstoque.sContaContabilAplic = String(STRING_CONTA, 0)
'            tItemMovEstoque.sContaContabilEst = String(STRING_CONTA, 0)
'
'            lErro = Comando_ExecutarPos(alComando1(4), "SELECT NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao FROM MovimentoEstoque WHERE NumIntDoc = ? ", 0, _
'            tItemMovEstoque.lNumIntDoc, tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.dHora, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, tItemMovEstoque.lHorasMaquina, tItemMovEstoque.dtDataInicioProducao, tItemMovEstoque.lNumIntDoc)
'            If lErro <> AD_SQL_SUCESSO Then gError 90597
'
'            lErro = Comando_BuscarPrimeiro(alComando1(4))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90598
'Retirar o código comentado acima

        End If
        
        '*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar descomentado
'        iPulaProdutoErro = 0
'
'        For Each vProduto In colProdutosErro
'            If tItemMovEstoque.sProduto = vProduto Then
'                iPulaProdutoErro = 1
'                Exit For
'            End If
'        Next
'
'        If iPulaProdutoErro = 0 Then '(descomentar também o if que está mais abaixo)
        ' ***
        
            'Atualiza tela de acompanhamento do Batch
            lErro = Rotina_Reproc_AtualizaTelaBatch()
            If lErro <> SUCESSO Then gError 83770
        
            objTipoMovEstoque.iCodigo = tItemMovEstoque.iTipoMov
            
            'ler os dados referentes ao tipo de movimento
            lErro = CF("TiposMovEst_Le1", alComando(10), objTipoMovEstoque)
            If lErro <> SUCESSO Then gError 20368
        
            If objTipoMovEstoque.iAtualizaMovEstoque <> TIPOMOV_EST_ESTORNOMOV Then
        
                'se passou para um outro conjunto de movimentos de estoque (com outro codigo) processa o conjunto com o codigo anterior
                If lCodigo <> tItemMovEstoque.lCodigo Or iFilialEmpresa <> tItemMovEstoque.iFilialEmpresa Then
        
                    'faz o reprocessamento do conjunto de movimentos de estoque com o mesmo codigo previamento armazenados em colItemMovEstoque e a contabilidade associada
                    lErro = Executa_Reprocessamento_MovEstoque(alComando, colExercicio, colItemMovEstoque, colEstoqueMesProduto, Len(Trim(objReprocessamentoEst.sProdutoCodigo)) <> 0)
                    If lErro <> SUCESSO Then gError 83662
                    
                    'faz o reprocessamento dos estornos
                    lErro = Executa_Reprocessamento_Estornos(alComando, colExercicio, colItemMovEstoque, colEstoqueMesProduto, Len(Trim(objReprocessamentoEst.sProdutoCodigo)) <> 0)
                    If lErro <> SUCESSO Then gError 83752
                        
                    'começa uma nova coleção com um novo codigo
                    Set colItemMovEstoque = New Collection
                    
                    lCodigo = tItemMovEstoque.lCodigo
                    iFilialEmpresa = tItemMovEstoque.iFilialEmpresa
        
                End If
        
                'quando qualquer movimento passa para um mes/ano maior do que o ultimo mes/ano em que o custo standard foi ajustado ==> verifica os produtos que precisam ter o custostandard ajustado e processa-os
                If Month(tItemMovEstoque.dtData) > iMesCstStd Or Year(tItemMovEstoque.dtData) > iAnoCstStd Then

                    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
                    
                        dtDataCstStd = tItemMovEstoque.dtData
                                                                        
                        'If iAnoCstStd = 0 Then iAnoCstStd = Year(dtDataCstStd)
                        If iAnoCstStd = 0 Then iAnoCstStd = Year(objReprocessamentoEst.dtDataInicio)
                                                    
                        Do While iAnoCstStd < Year(tItemMovEstoque.dtData)
                    
                            For Each objFiliaisReprocAno In gcolFiliais
                                
                                If objFiliaisReprocAno.iCodFilial <> EMPRESA_TODA Then
                            
                                    'processa a transferencia dos saldos iniciais
                                    lErro = CF("Rotina_FechamentoAno_Reproc", iAnoCstStd, objFiliaisReprocAno.iCodFilial, objReprocessamentoEst.sProdutoCodigo)
                                    If lErro <> SUCESSO Then gError 83634
                                    
                                End If
                                
                            Next
                            
                            iAnoCstStd = iAnoCstStd + 1
                                            
                        Loop
                        
                    Else
                        'Processa os ajustes do Custos Standard de todos os produtos custeados por este dos meses que sofreram alteração desda a DataCstStd até a data do movimento que está sendo processado
                        lErro = Processa_Ajuste_Custo_Standard(alComando, objReprocessamentoEst.dtDataInicio, dtDataCstStd, tItemMovEstoque.dtData, objTela, lNumIntDoc, iAnoUltFechamento, colExercicio, colEstoqueMesProduto)
                        If lErro <> SUCESSO Then gError 83634
                    End If
                    
                    iMesCstStd = Month(dtDataCstStd)
                    iAnoCstStd = Year(dtDataCstStd)

                End If
                
                Set objItemMovEst = New ClassItemMovEstoque
            
                'move os dados de tItemMovEstoque para objItemMovEst
                Call Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque, objItemMovEst)
                
'                Public Const MOV_EST_ACRES_INVENT_DISPONIVEL_NOSSA_FX_ZERA = 345
'                Public Const MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE_FX_ZERA = 346
'                Public Const MOV_EST_DECR_INVENT_DISPONIVEL_NOSSA_FX_ZERA = 347
'                Public Const MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE_FX_ZERA = 348
                If objItemMovEst.iTipoMov = MOV_EST_ACRES_INVENT_DISPONIVEL_NOSSA_FX_ZERA Or _
                    objItemMovEst.iTipoMov = MOV_EST_DECR_INVENT_DISPONIVEL_NOSSA_FX_ZERA Then
                
                     'Tem que acertar as quantidades
                    lErro = CF("Prod_Obtem_Qtde_Custo_Est_Data", objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado, objItemMovEst.dtData, "", 0, dQtdCalc, dCustoCalc)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                           
                    If dQtdCalc <= 0 Then
                        objItemMovEst.iTipoMov = MOV_EST_ACRES_INVENT_DISPONIVEL_NOSSA_FX_ZERA
                    Else
                        objItemMovEst.iTipoMov = MOV_EST_DECR_INVENT_DISPONIVEL_NOSSA_FX_ZERA
                    End If
                    
                    objItemMovEst.dQuantidade = Abs(dQtdCalc)
                
                End If
                
                Set colApropriacaoInsumo = New Collection
                
                If objItemMovEst.iTipoMov = MOV_EST_PRODUCAO Or objItemMovEst.iTipoMov = MOV_EST_PRODUCAO_BENEF3 Then
                
                    'Le as Apriações do Item
                    lErro = CF("ApropriacaoInsumo_Le_NumIntDocOrigem", tItemMovEstoque.lNumIntDoc, colApropriacaoInsumo)
                    If lErro <> SUCESSO Then gError 92672
                                
                End If
                                
                Set objItemMovEst.colApropriacaoInsumo = colApropriacaoInsumo
                
                colItemMovEstoque.Add objItemMovEst
                
            End If
        
        'End If
        
'*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar descomentado
'Continua_Pos_Erro:
' ***

        'Se serão reprocessados primeiro os movimentos de entrada
        If objReprocessamentoEst.iOrdemReproc = REPROCESSAMENTO_ORDENA_ENTRADAS Then
        
            'Busca o NumIntDoc do próximo movimento a ser reprocessado
            'A leitura do movimento é feita através de um select posicionado no início do loop
            lErro = Comando_BuscarProximo(alComando1(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83569
            
            Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
            
            'Fecha o comando que será reutilizado para ler o movimento de estoque utilizando o NumIntDoc lido
            Call Comando_Fechar(alComando1(4))
            
            'Reabre o comando que será reutilizado para ler o movimento de estoque utilizando o NumIntDoc lido
            alComando1(4) = Comando_Abrir()
            If alComando1(4) = 0 Then gError 90600
            
        'Senão
        Else
        
            'Busca o próximo movimento a ser reprocessado
            lErro = Comando_BuscarProximo(alComando1(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83569
            
            Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
            
        End If
        
    Loop

    'faz o reprocessamento dos movimentos de estoque com o mesmo codigo
    lErro = Executa_Reprocessamento_MovEstoque(alComando, colExercicio, colItemMovEstoque, colEstoqueMesProduto, Len(Trim(objReprocessamentoEst.sProdutoCodigo)) <> 0)
    If lErro <> SUCESSO Then gError 83753

    'faz o reprocessamento dos estornos
    lErro = Executa_Reprocessamento_Estornos(alComando, colExercicio, colItemMovEstoque, colEstoqueMesProduto, Len(Trim(objReprocessamentoEst.sProdutoCodigo)) <> 0)
    If lErro <> SUCESSO Then gError 83754

    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) = 0 Then
        'Processa os ajustes do Custos Standard de todos os produtos custeados por este, dos meses que sofreram alteração desde a DataCstStd até a data atual
        lErro = Processa_Ajuste_Custo_Standard(alComando, objReprocessamentoEst.dtDataInicio, dtDataCstStd, gdtDataHoje, objTela, lNumIntDoc, iAnoUltFechamento, colExercicio, colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 83660
    
    Else
    
        Do While iAnoCstStd < Year(gdtDataHoje)
    
            For Each objFiliaisReprocAno In gcolFiliais
                
                If objFiliaisReprocAno.iCodFilial <> EMPRESA_TODA Then
            
                    'processa a transferencia dos saldos iniciais
                    lErro = CF("Rotina_FechamentoAno_Reproc", iAnoCstStd, objFiliaisReprocAno.iCodFilial, objReprocessamentoEst.sProdutoCodigo)
                    If lErro <> SUCESSO Then gError 83634
                    
                End If
                
            Next
            
            iAnoCstStd = iAnoCstStd + 1
                            
        Loop
    
    End If

    If Not objTela Is Nothing Then
        objTela.Unload objTela
    End If

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Rotina_Reproc_Refaz = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Refaz:

    Rotina_Reproc_Refaz = gErr

    Select Case gErr
    
        Case 83566, 83567, 83569, 90597 To 90599, 90649
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
    
        Case 20368, 83634, 83660, 83662, 83663, 83752, 83753, 83754, 83770, 90600, 90702 To 90704
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173767)

    End Select

    Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_REPROCESSAMENTO", gl_UltimoErro, tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.lNumIntDoc, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.dtData)

    '*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar descomentado
'    If gl_UltimoErro <> 83753 And gl_UltimoErro <> 83754 And gl_UltimoErro <> 83660 Then
'        colProdutosErro.Add tItemMovEstoque.sProduto
'        Resume Continua_Pos_Erro
'    End If
     '***
            
'*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar comentado
    If Not objTela Is Nothing Then
        objTela.Unload objTela
    End If

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next
' ***

    Exit Function

End Function

Private Function Executa_Reprocessamento_MovEstoque(alComando() As Long, colExercicio As Collection, colItemMovEstoque As Collection, colEstoqueMesProduto As Collection, ByVal bReprocessandoProdIsolado As Boolean) As Long
'faz o reprocessamento dos movimentos de estoque com o mesmo codigo

Dim lErro As Long
Dim objItemMovEst As New ClassItemMovEstoque, objItemMovEstAux As New ClassItemMovEstoque
Dim iOrigemLcto As Integer
Dim lNumIntDocOrigemCTB As Long
Dim objItemMovEst1 As ClassItemMovEstoque
Dim iFilialEmpresa As Integer


On Error GoTo Erro_Executa_Reprocessamento_MovEstoque

    'reprocessa cada item do movimento com o mesmo codigo
    For Each objItemMovEst In colItemMovEstoque

        lErro = CF("Estoque_Reprocessamento", alComando(), objItemMovEst, REPROCESSAMENTO_REFAZ, colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 83568
    
    Next
    
    'se o módulo de contabilidade estiver ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        Set objItemMovEst = colItemMovEstoque(1)
        
        If bReprocessandoProdIsolado Then
        
            objItemMovEstAux.iFilialEmpresa = objItemMovEst.iFilialEmpresa
            objItemMovEstAux.lCodigo = objItemMovEst.lCodigo
            
            lErro = CF("MovEstoqueItem_Le_Primeiro", objItemMovEstAux)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set objItemMovEst = objItemMovEstAux
            
        End If
        
        'devolve o documento que originou o movimento de estoque.
        'Utilizado para descobrir os lançamentos contábeis associados e reprocessá-los.
        lErro = Retorna_Origem_Estoque_Contab(objItemMovEst.iTipoNumIntDocOrigem, objItemMovEst.lNumIntDocOrigem, objItemMovEst.lNumIntDoc, iOrigemLcto, lNumIntDocOrigemCTB)
        If lErro <> SUCESSO Then gError 83591
    
        iFilialEmpresa = objItemMovEst.iFilialEmpresa
    
        'reprocessa a contabilização do movimento de estoque
        lErro = CF("Rotina_Reprocessamento_DocOrigem", iOrigemLcto, lNumIntDocOrigemCTB, colExercicio, objItemMovEst.iFilialEmpresa)
        If lErro <> SUCESSO Then gError 83592

    End If
    
    Executa_Reprocessamento_MovEstoque = SUCESSO
    
    Exit Function
    
Erro_Executa_Reprocessamento_MovEstoque:

    Executa_Reprocessamento_MovEstoque = gErr
    
    Select Case gErr
    
        Case 83598, 83591, 83592, 83568, ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173768)

    End Select

    Exit Function

End Function

Function Retorna_Origem_Estoque_Contab(ByVal iTipoNumIntDocOrigem As Integer, ByVal lNumIntDocOrigemEst As Long, ByVal lNumIntDocItemMovEst As Long, iOrigemLcto As Integer, lNumIntDocOrigemCTB As Long) As Long
'devolve o documento que originou o movimento de estoque.
'Utilizado para descobrir os lançamentos contábeis associados e reprocessá-los.

Dim objItemNFiscal As New ClassItemNF
Dim lErro As Long
Dim lNumIntNF As Long
Dim lNumIntGrade As Long

On Error GoTo Erro_Retorna_Origem_Estoque_Contab

    Select Case iTipoNumIntDocOrigem
    
        Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
        
            'o numintdocorigem do movimento é o item da nota fiscal
            objItemNFiscal.lNumIntDoc = lNumIntDocOrigemEst
            
            'a partir do item consegue-se o numintdoc da nota que é a origem da contabilização
            lErro = CF("ItemNFiscal_Le", objItemNFiscal)
            If lErro <> SUCESSO And lErro <> 35225 Then gError 83586
            
            'se o item da nota fiscal não estiver cadastrado ==> erro
            If lErro = 35225 Then gError 83587
            
            lNumIntDocOrigemCTB = objItemNFiscal.lNumIntNF
            iOrigemLcto = TRANSACAOCTBORIGEM_NFISCAL
            
        Case MOVEST_TIPONUMINTDOCORIGEM_NFISCAL
            
            'a origem da contabilidade coincide com a origem do movimento do estoque, ou seja, é o numintdoc da nota fiscal
            lNumIntDocOrigemCTB = lNumIntDocOrigemEst
            iOrigemLcto = TRANSACAOCTBORIGEM_NFISCAL
            
        Case 0, MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO, MOVEST_TIPONUMINTDOCORIGEM_MOVESTOQUE, MOVEST_TIPONUMINTDOCORIGEM_ITEMOP
        
            'a origem é o proprio numero do movimento de estoque (o primeiro)
            lNumIntDocOrigemCTB = lNumIntDocItemMovEst
            iOrigemLcto = TRANSACAOCTBORIGEM_MOVIMENTOESTOQUE
            
        Case MOVEST_TIPONUMINTDOCORIGEM_CUPOMFISCAL
            lNumIntDocOrigemCTB = 0
            iOrigemLcto = 0

        Case MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCALGRADE

            'o numintdocorigem do movimento é o item da grade
            lNumIntGrade = lNumIntDocOrigemEst
            
            'a partir do item consegue-se o numintdoc da nota que é a origem da contabilização
            lErro = CF("ItensNFGrade_Le1", lNumIntNF, lNumIntGrade)
            If lErro <> SUCESSO And lErro <> 133854 Then gError 133855
            
            'se o item da nota fiscal não estiver cadastrado ==> erro
            If lErro = 133854 Then gError 133856
            
            lNumIntDocOrigemCTB = lNumIntNF
            iOrigemLcto = TRANSACAOCTBORIGEM_NFISCAL

        Case Else
            gError 83590
        
    End Select
    
    Retorna_Origem_Estoque_Contab = SUCESSO
    
    Exit Function
    
Erro_Retorna_Origem_Estoque_Contab:

    Retorna_Origem_Estoque_Contab = gErr

    Select Case gErr
    
        Case 83586, 133855

        Case 83587
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NFORIGINAL_NAO_CADASTRADO2", gErr, objItemNFiscal.lNumIntDoc)

        Case 83590
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPONUMINTDOCORIGEM_NAO_TRATADO", gErr, iTipoNumIntDocOrigem)

        Case 133856
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNFGRADE_NAO_CADASTRADO", gErr, lNumIntGrade)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173769)

    End Select

    Exit Function

End Function

Private Function Processa_Ajuste_Custo_Standard(alComando() As Long, ByVal dtDataInicio As Date, dtDataUltimoAjuste As Date, ByVal dtDataMovEst As Date, objTela As Object, lNumIntDoc As Long, iAnoUltFechamento As Integer, colExercicio As Collection, colEstoqueMesProduto As Collection) As Long
'processa o ajuste do custo standard do dia 1 de cada mes entre a data que já foi calculado o custo standard (dtDataUltimoAjuste) e a data do movimento de estoque
'processa a apuração do custo de produção para os meses que estão terminando entre a data que já foi calculado o custo standard e a data do movimento de estoque
'dtDataCstStd deverá conter a data do ultimo ajuste do custo standard do produto

Dim dtData As Date
Dim tProduto As typeProduto
Dim lComando As Long
Dim alComando1(1 To 2) As Long
Dim objSldMesEst As ClassSldMesEst
Dim dtDataAjusteAtual As Date
Dim dCustoStd As Double
Dim iMesAtual As Integer
Dim lCodigo As Long
Dim objMovEstoque As New ClassMovEstoque
Dim colAlmoxarifados As New Collection
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objContabAutomatica As ClassContabAutomatica
Dim iIndice As Integer
Dim lErro As Long
Dim objFiliais As AdmFiliais
Dim iFilialEmpresa As Integer
Dim iAno As Integer
Dim iMes As Integer
Dim colEstoqueMes As Collection
Dim iAnoUltFechamentoAux As Integer 'Inserido por Wagner

On Error GoTo Erro_Processa_Ajuste_Custo_Standard

    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 83643
    Next

    'se a data do ultimo ajuste do custo standard for nula (ou seja ainda não começou o processamento do ajuste do custo standard)
    If dtDataUltimoAjuste = DATA_NULA Then
    
        'se a data de inicio for diferente do dia 1 de um mes ==> coloca o dia 1 como a data do ultimo ajuste do custo standard
        If Day(dtDataInicio) <> 1 Then
    
            dtDataUltimoAjuste = CDate("1/" & CStr(Month(dtDataInicio)) & "/" & Year(dtDataInicio))
            
        Else
        
            'coloca a data do dia 1 do mes anterior como a data do ultimo ajuste
            dtData = dtDataInicio - 1
            dtDataUltimoAjuste = CDate("1/" & CStr(Month(dtData)) & "/" & Year(dtData))
            
            
        End If
    
    End If
    
    'enquanto o mes/ano do ultimo ajuste for menor que o mes/ano do movimento de estoque em questão
    Do While Year(dtDataUltimoAjuste) < Year(dtDataMovEst) Or (Year(dtDataUltimoAjuste) = Year(dtDataMovEst) And Month(dtDataUltimoAjuste) < Month(dtDataMovEst))

        'se o mes/ano do ultimo ajuste for diferente do mes/ano data datainicio -1 ==> pode processar a apuraçãoo do custo de produção
        'quando o mes/ano do ultimo ajuste coincide com o mes/ano da datainicio - 1
        'é sinal que o reprocessamento do custo de producao do mes/ano contido na dataultimoajuste não é para ser feito já que o processamento se inicia no dia 01 do mes seguinte (dtDataInicio).
        If Month(dtDataUltimoAjuste) <> Month(dtDataInicio - 1) Or Year(dtDataUltimoAjuste) <> Year(dtDataInicio - 1) Then
        
            'se o ano não mudou processa a transferencia dos custos de um mes para o outro
            If Year(dtDataUltimoAjuste) = Year(dtDataMovEst) And Month(dtDataUltimoAjuste) < Month(dtDataMovEst) Then
        
                For Each objFiliais In gcolFiliais
                
                    iFilialEmpresa = objFiliais.iCodFilial
        
                    If iFilialEmpresa <> EMPRESA_TODA Then
        
                        'Rotina que fecha o mes passado como parametro para a rotina de reprocessamento
                        lErro = CF("Rotina_FechamentoMes_Reproc", iFilialEmpresa, Year(dtDataUltimoAjuste), Month(dtDataUltimoAjuste))
                        If lErro <> SUCESSO Then gError 89213
                    
                    End If
                    
                Next
        
            End If
        
        
'            Mario. 24/9/01. Retirado pois o novo calculo do custo de produção segue a rotina de calculo do custo medio
'            'calcula o custo médio de produção para mes/ano passados e valora os movimentos de estoque
'            lErro = Rotina_CustoMedioProducao_Reproc(iFilialEmpresa, Year(dtDataUltimoAjuste), Month(dtDataUltimoAjuste))
'            If lErro <> SUCESSO Then gError 83761
            
            For Each objFiliais In gcolFiliais
                
                iFilialEmpresa = objFiliais.iCodFilial
                
                If iFilialEmpresa <> EMPRESA_TODA Then
            
                    'Ajustar a contabilidade do mes estoque em questão
                    lErro = CF("Rotina_Reprocessamento_CProd", iFilialEmpresa, Month(dtDataUltimoAjuste), Year(dtDataUltimoAjuste), colExercicio)
                    If lErro <> SUCESSO Then gError 83762
                    
                End If
                
            Next
                
        End If


        dtDataAjusteAtual = DateAdd("m", 1, dtDataUltimoAjuste)

        iAno = Year(dtDataAjusteAtual)
        iMes = Month(dtDataAjusteAtual)

        Set colEstoqueMes = New Collection

        'faz lock em estoqueMes, le os Gastos Diretos e Indiretos de todas as filiais e coloca estes valores em colEstoqueMes.
        lErro = CF("Rotina_CMP_EstoqueMes_CriticaLock", alComando(2), iAno, iMes, colEstoqueMes)
        If lErro <> SUCESSO Then gError 92547
        
        Set colEstoqueMesProduto = New Collection
        
        'preenche uma colecao com os produtos que tiveram gastos informados e que portanto não terão seu calculo feito com os demais produtos
        lErro = CF("EstoqueMesProduto_Le", iAno, iMes, colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 92890
        
        'Apura o total de horas maquina e custo das matérias primas para cada filial e coloca em colEstoqueMes
        lErro = CF("MovEstoque_Le_HorasMaq_CustoMPrim", colEstoqueMes, iMes, iAno, colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 92548
        
        'Grava o total de horas maquina e custo das matérias primas para cada filial
        lErro = CF("EstoqueMes_Grava", colEstoqueMes, True)
        If lErro <> SUCESSO Then gError 92549

        'Grava a quantidade total produzida no mes de cada produto que teve o gasto especificado
        lErro = CF("EstoqueMesProduto_Grava", colEstoqueMesProduto)
        If lErro <> SUCESSO Then gError 92902

        If iAnoUltFechamento < Year(dtDataAjusteAtual) Then
            
            iAnoUltFechamentoAux = iAnoUltFechamento
            
            For Each objFiliais In gcolFiliais
                
                iFilialEmpresa = objFiliais.iCodFilial
                
                If iFilialEmpresa <> EMPRESA_TODA Then
            
                    iAnoUltFechamentoAux = iAnoUltFechamento
                    
                    'processa a transferencia dos saldos iniciais de um iAnoAtual até iAnoUltMov
                    lErro = Processa_FechamentoAno_Reproc(iAnoUltFechamentoAux, Year(dtDataAjusteAtual), iFilialEmpresa)
                    If lErro <> SUCESSO Then gError 83745
                    
                End If
                
            Next
            
            iAnoUltFechamento = iAnoUltFechamentoAux
        
        End If

        tProduto.sCodigo = String(STRING_PRODUTO, 0)
        tProduto.sDescricao = String(STRING_PRODUTO_DESCRICAO, 0)
        tProduto.sSiglaUMEstoque = String(STRING_PRODUTO_SIGLAUMESTOQUE, 0)

        'le todos os produtos que possuem apropriacao pelo custo standard
        lErro = Comando_Executar(alComando1(1), "SELECT Codigo, Descricao, SiglaUMEstoque FROM Produtos WHERE Apropriacao = ? ", tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, APROPR_CUSTO_STANDARD)
        If lErro <> AD_SQL_SUCESSO Then gError 83644

        lErro = Comando_BuscarPrimeiro(alComando1(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83645

        Do While lErro = AD_SQL_SUCESSO
        
            For Each objFiliais In gcolFiliais
                
                iFilialEmpresa = objFiliais.iCodFilial
                
                If iFilialEmpresa <> EMPRESA_TODA Then
        
                    Set objSldMesEst = New ClassSldMesEst
                
                    objSldMesEst.sProduto = tProduto.sCodigo
                    objSldMesEst.iAno = Year(dtDataUltimoAjuste)
                    objSldMesEst.iFilialEmpresa = iFilialEmpresa
                
                    'Le na tabela SldMesEst os custos relativos ao ajuste do mes anterior
                    lErro = CF("SldMesEst_Le_Custos", Month(dtDataUltimoAjuste), objSldMesEst)
                    If lErro <> SUCESSO And lErro <> 41755 Then gError 83646
        
                    dCustoStd = objSldMesEst.dCustoStandard(Month(dtDataUltimoAjuste))
        
                    'se existia um valor no mes anterior
                    If lErro <> 41755 And dCustoStd <> 0 Then
        
                        Set objSldMesEst = New ClassSldMesEst
        
                        objSldMesEst.sProduto = tProduto.sCodigo
                        objSldMesEst.iAno = Year(dtDataAjusteAtual)
                        objSldMesEst.iFilialEmpresa = iFilialEmpresa
                        iMesAtual = Month(dtDataAjusteAtual)
                    
                        'Le na tabela SldMesEst os custos relativos ao ajuste do mes que se está processando
                        lErro = CF("SldMesEst_Le_Custos", iMesAtual, objSldMesEst)
                        If lErro <> SUCESSO And lErro <> 41755 Then gError 83647
            
                        If lErro = 41755 Then gError 83648
                        
                        If dCustoStd <> objSldMesEst.dCustoStandard(iMesAtual) Then
                
                            lErro = CF("MovEstoque_Automatico_EmTransacao", iFilialEmpresa, lCodigo)
                            If lErro <> SUCESSO Then gError 83649
                        
                            objMovEstoque.dtData = dtDataAjusteAtual
                            objMovEstoque.iFilialEmpresa = giFilialEmpresa
                            objMovEstoque.lCodigo = lCodigo
                            
                            'Lê todos os Almoxarifados do Produto para todas as filiais
                            lErro = CF("EstoqueProduto_Le_Almoxarifados", tProduto.sCodigo, colAlmoxarifados)
                            If lErro <> SUCESSO Then gError 83651
                            
                            For Each objAlmoxarifado In colAlmoxarifados
                            
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_CONSIG_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_DEMO_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_CONSERTO_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_OUTROS_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                objMovEstoque.colItens.Add 0, MOV_EST_AJUSTE_CUSTO_STD_BENEF_NOSSO, objSldMesEst.dCustoStandard(iMesAtual) - dCustoStd, APROPR_CUSTO_INFORMADO, tProduto.sCodigo, tProduto.sDescricao, tProduto.sSiglaUMEstoque, 0, objAlmoxarifado.iCodigo, "", 0, "", 0, "", "", "", "", 0, Nothing, Nothing, DATA_NULA
                                    
                            Next
                
                            'grava os movimentos de estoque de reajuste do custo standard
                            lErro = CF("Estoque_ReprocAjusteStd", alComando, objMovEstoque, lNumIntDoc)
                            If lErro <> SUCESSO Then gError 83655
                
                            If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
                            
                                Set objContabAutomatica = New ClassContabAutomatica
                            
                                lErro = objContabAutomatica.Inicializa_Contab(objTela, MODULO_ESTOQUE, "CES", objMovEstoque.dtData, objMovEstoque.dtData)
                                If lErro <> SUCESSO Then gError 83656
                                
                                lErro = objContabAutomatica.GeraContabilizacao(objMovEstoque)
                                If lErro <> SUCESSO Then gError 83657
                            
                                lErro = objContabAutomatica.Finaliza_Contab()
                                If lErro <> SUCESSO Then gError 83658
                                
                            End If
            
                        End If
                        
                    End If
                    
                End If
                
            Next
            
            lErro = Comando_BuscarProximo(alComando1(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83659

        Loop

        'pega o primeiro dia do proximo mes
        dtDataUltimoAjuste = dtDataAjusteAtual
        
    Loop

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

     Processa_Ajuste_Custo_Standard = SUCESSO
    
    Exit Function
    
Erro_Processa_Ajuste_Custo_Standard:

    Processa_Ajuste_Custo_Standard = gErr

    Select Case gErr
    
        Case 83643
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 83644, 83645, 83659
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS1", gErr)
    
        Case 83646, 83647, 83649, 83651, 83655, 83656, 83657, 83658, 83745, 83761, 83762, 89213, 92890, 92902

        Case 83648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOS_INEXISTENTES", gErr, objSldMesEst.iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173770)

    End Select

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Exit Function
    
End Function

Private Function Processa_FechamentoAno_Reproc(iAnoUltFechamento As Integer, ByVal iAnoUltMov As Integer, ByVal iFilialEmpresa As Integer) As Long
'processa a transferencia dos saldos iniciais de iAnotUltFechamento até iAnoUltMov

Dim lErro As Long

On Error GoTo Erro_Processa_FechamentoAno_Reproc

    Do While iAnoUltFechamento < iAnoUltMov

        'Rotina que transfere os saldos de um ano para outro para fins de reprocessamento.
        'iAnoAtual é o ano que está terminando.
        lErro = CF("Rotina_FechamentoAno_Reproc", iAnoUltFechamento, iFilialEmpresa)
        If lErro <> SUCESSO Then gError 83744

        iAnoUltFechamento = iAnoUltFechamento + 1

    Loop

    Processa_FechamentoAno_Reproc = SUCESSO
    
    Exit Function
    
Erro_Processa_FechamentoAno_Reproc:

    Processa_FechamentoAno_Reproc = gErr

    Select Case gErr
    
        Case 83744
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173771)

    End Select

    Exit Function
    
End Function

Private Function Executa_Reprocessamento_Estornos(alComando() As Long, colExercicio As Collection, colItemMovEstoque As Collection, colEstoqueMesProduto As Collection, ByVal bReprocessandoProdIsolado As Boolean) As Long
'processa os estornos (se houverem) para a coleção de itens de movimento de estoque que possuem o mesmo codigo

Dim lErro As Long
Dim colEstornos As New Collection
Dim colItemMovEstoqueEstorno As Collection
Dim objItemMovEst1 As ClassItemMovEstoque
Dim objItemMovEst As ClassItemMovEstoque
Dim iAchou As Integer
Dim bDesordenado As Boolean
Dim colEstornosOrdenados As New Collection
Dim lNumIntDoc As Long
Dim iIndice As Integer
Dim iIndiceMenorNumIntDoc As Integer

On Error GoTo Erro_Executa_Reprocessamento_Estornos

    'tenta achar os estornos associados a cada item
    For Each objItemMovEst In colItemMovEstoque
        
        If objItemMovEst.lNumIntDocEst <> 0 Then
        
            Set objItemMovEst1 = New ClassItemMovEstoque
            
            objItemMovEst1.lNumIntDoc = objItemMovEst.lNumIntDocEst
        
            'Carrega os dados do movimento com o NumIntDoc passado como parametro
            lErro = CF("MovimentoEstoque_Le_NumIntDoc", alComando(88), objItemMovEst1)
            If lErro <> SUCESSO And lErro <> 83748 Then gError 83750
        
            If lErro = 83748 Then gError 83751
            
            iAchou = 0
            
            'coloca os movimentos de estoque de estorno que possuem o mesmo codigo juntos.
            'isto permitirá processar todos os movimentos de estorno com o mesmo codigo juntos e depois sua contabilização.
            For Each colItemMovEstoqueEstorno In colEstornos
                If colItemMovEstoqueEstorno.Item(1).lCodigo = objItemMovEst1.lCodigo Then
                    colItemMovEstoqueEstorno.Add objItemMovEst1
                    iAchou = 1
                    Exit For
                End If
            Next
                
            'se ainda não havia nenhuma coleção de movimentos com o codigo em questão ==> cria uma nova coleção para este codigo
            If iAchou = 0 Then
                Set colItemMovEstoqueEstorno = New Collection
                colItemMovEstoqueEstorno.Add objItemMovEst1
                colEstornos.Add colItemMovEstoqueEstorno
            End If
            
        End If
    Next
            
    'se existem estornos
    If colEstornos.Count > 0 Then
    
        If colEstornos.Count = 1 Then
            Set colEstornosOrdenados = colEstornos
        Else
        
            bDesordenado = True
                    
            'coloca as coleções por ordem de numintdoc que é sua ordem de processamento
            Do While True
            
                lNumIntDoc = colEstornos.Item(1).Item(1).lNumIntDoc
                
                For iIndice = 1 To colEstornos.Count
                    
                    Set colItemMovEstoqueEstorno = colEstornos.Item(iIndice)
                    If colItemMovEstoqueEstorno.Item(1).lNumIntDoc <= lNumIntDoc Then
                        iIndiceMenorNumIntDoc = iIndice
                    End If
        
                Next
                
                colEstornosOrdenados.Add colEstornos.Item(iIndiceMenorNumIntDoc)
                colEstornos.Remove (iIndiceMenorNumIntDoc)
    
                If colEstornos.Count = 0 Then
                    Exit Do
                End If
    
            Loop
    
        End If
    
    End If

    'reprocessa os estornos do codigo em questão na mesma ordem em que foram processados originalmente
    For Each colItemMovEstoqueEstorno In colEstornosOrdenados
    
        'faz o reprocessamento dos movimentos de estoque com o mesmo codigo
        lErro = Executa_Reprocessamento_MovEstoque(alComando, colExercicio, colItemMovEstoqueEstorno, colEstoqueMesProduto, bReprocessandoProdIsolado)
        If lErro <> SUCESSO Then gError 83755
    
    Next
    
    Executa_Reprocessamento_Estornos = SUCESSO
    
    Exit Function

Erro_Executa_Reprocessamento_Estornos:

    Executa_Reprocessamento_Estornos = gErr

    Select Case gErr
    
        Case 83750, 83755
    
        Case 83751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_NAO_CADASTRADO", gErr, objItemMovEst1.lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173772)

    End Select

    Exit Function

End Function

Private Function Rotina_Reproc_AtualizaTelaBatch()
'Atualiza tela de acompanhamento do Batch

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Rotina_Reproc_AtualizaTelaBatch

    '*** Para depurar, usando o BatchEst como .dll, o trecho abaixo deve estar comentado
    'Atualiza tela de acompanhamento do Batch

    lErro = DoEvents()

    TelaAcompanhaBatchEST.dValorAtual = TelaAcompanhaBatchEST.dValorAtual + 1
    TelaAcompanhaBatchEST.TotReg.Caption = CStr(TelaAcompanhaBatchEST.dValorAtual)
    TelaAcompanhaBatchEST.ProgressBar1.Value = CInt((TelaAcompanhaBatchEST.dValorAtual / TelaAcompanhaBatchEST.dValorTotal) * 100)

    If TelaAcompanhaBatchEST.iCancelaBatch = CANCELA_BATCH Then

        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_CANCELAR_REPROC_MOVESTOQUE")

        If vbMsgBox = vbYes Then gError 83769

        TelaAcompanhaBatchEST.iCancelaBatch = 0

    End If
    '***
    
    Rotina_Reproc_AtualizaTelaBatch = SUCESSO
    
    Exit Function

Erro_Rotina_Reproc_AtualizaTelaBatch:

    Rotina_Reproc_AtualizaTelaBatch = gErr

    Select Case gErr

        Case 83769

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173773)

    End Select
       
    Exit Function

End Function

Private Function MATConfig_Grava_DataReproc(ByVal dtData As Date, ByVal iFilialEmpresa As Integer) As Long
'Insere ou Atualiza registro em MATConfig com a data a partir da qual deve ser feito o reprocessamento de estoque.

Dim sConteudo As String
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim dtDataBD As Date

On Error GoTo Erro_MATConfig_Grava_DataReproc

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 83554
    Next

    sConteudo = String(STRING_CONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM MATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, DATA_REPROCESSAMENTO, iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 83555
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83556
    
    If lErro = AD_SQL_SUCESSO Then
    
        dtDataBD = CDate(sConteudo)
    
        'se a data passada como parametro for menor do que a data do banco de dados ==> atualiza a data
        If dtData < dtDataBD Then
    
            sConteudo = CStr(dtData)
    
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE MATConfig SET Conteudo = ?", alComando(1), sConteudo)
            If lErro <> AD_SQL_SUCESSO Then gError 83557
    
        End If
    
    Else
    
        sConteudo = CStr(dtData)
    
        lErro = Comando_Executar(alComando(2), "INSERT INTO MATConfig (Codigo, FilialEmpresa, Descricao, Tipo, Conteudo) VALUES (?,?,?,?,?)", DATA_REPROCESSAMENTO, iFilialEmpresa, DATA_REPROCESSAMENTO_DESCR, CONFIG_TIPO_DATA, sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 83558
        
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    MATConfig_Grava_DataReproc = SUCESSO
    
    Exit Function

Erro_MATConfig_Grava_DataReproc:

    MATConfig_Grava_DataReproc = gErr
    
    Select Case gErr
    
        Case 83554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 83555, 83556
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MATCONFIG1", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case 83557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_MATCONFIG1", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case 83558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MATCONFIG", gErr, DATA_REPROCESSAMENTO, iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173774)
        
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function Rotina_Reproc_Zera_Saldos(objReprocessamentoEst As ClassReprocessamentoEST) As Long
'Essa função deve ser chamada EM TRANSAÇÃO

Dim lErro As Long

On Error GoTo Erro_Rotina_Reproc_Zera_Saldos

    lErro = CF("Rotina_Reproc_Zera_SldMesEst", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90640
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst1", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90641
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst2", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90642
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90643
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm1", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90644
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEstAlm2", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90645
    
    lErro = CF("Rotina_Reproc_Calc_SldIni", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90645
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEst", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90646
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEstAlm", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90647
    
    lErro = CF("Rotina_Reproc_Zera_EstoqueProduto", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90648
    
    lErro = CF("Rotina_Reproc_Zera_SldDiaEstTerc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90646
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst1Terc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90641
    
    lErro = CF("Rotina_Reproc_Zera_SldMesEst2Terc", objReprocessamentoEst)
    If lErro <> SUCESSO Then gError 90642
   
    Rotina_Reproc_Zera_Saldos = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_Zera_Saldos:

    Rotina_Reproc_Zera_Saldos = gErr
    
    Select Case gErr
    
        Case 90640 To 90648
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173775)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Reproc_MontaSelect(objReprocessamentoEst As ClassReprocessamentoEST, asComandoSelect() As String, iOperacao As Integer) As Long
'Monta o select que será feito para efetuar a fase refaz do reprocessamento

Dim asSelect() As String
Dim asFrom() As String
Dim asWhere() As String
Dim asOrderBy() As String
Dim iNumSelect As String
Dim iIndice As Integer

On Error GoTo Erro_Rotina_Reproc_MontaSelect

    'Obtém o número de selects que deverão ser montados
    iNumSelect = UBound(asComandoSelect)
    
    'Redimensiona os arrays conforme o número de selects que serão montados
    ReDim asSelect(iNumSelect) As String
    ReDim asFrom(iNumSelect) As String
    ReDim asWhere(iNumSelect) As String
    ReDim asOrderBy(iNumSelect) As String
    
    'Para cada select da fase DESFAZ
    For iIndice = 1 To iNumSelect
        'Começa a montagem do select
        asSelect(iIndice) = "SELECT "
        asFrom(iIndice) = "FROM "
        asWhere(iIndice) = "WHERE "
        asOrderBy(iIndice) = "ORDER BY "
    Next
    
    Select Case iOperacao
    
        'Se estiver executando a fase desfaz do reprocessamento
        'Gera apenas o select que será utilizado no desfaz
        Case REPROCESSAMENTO_DESFAZ
        
            asSelect(1) = asSelect(1) & "NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao "
            asFrom(1) = asFrom(1) & "MovimentoEstoque "
            asWhere(1) = asWhere(1) & "Data >= ? "
            asOrderBy(1) = asOrderBy(1) & "NumIntDoc DESC "
            
            'Se foi incluído um filtro de produtos => inclui o filtro na cláusula WHERE
            If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then asWhere(1) = asWhere(1) & "AND Produto = ? "
            
            'Se foi incluído um filtro de data final => inclui o filtro na cláusula WHERE
            If objReprocessamentoEst.dtDataFim <> DATA_NULA Then asWhere(1) = asWhere(1) & "AND Data <= ? "

            If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then asWhere(1) = asWhere(1) & "AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"

        'Senão
        'ou seja, se for a fase refaz, gera os selects que podem ser utilizados nessa fase
        Case REPROCESSAMENTO_REFAZ
        
            'Verifica se serão reprocessados primeiro os movimentos de entrada
            If objReprocessamentoEst.iOrdemReproc = REPROCESSAMENTO_ORDENA_ENTRADAS Then
            
                'Monta um select para ler o NumIntDoc dos movimentos de estoque, ordenando primeiro as entradas e depois as saídas
                asSelect(1) = asSelect(1) & "MovimentoEstoque.NumIntDoc, MovimentoEstoque.Codigo, MovimentoEstoque.FilialEmpresa, MovimentoEstoque.Data "
                asFrom(1) = asFrom(1) & "MovimentoEstoque, TiposMovimentoEstoque, TiposOrdemCusto "
                asWhere(1) = asWhere(1) & "MovimentoEstoque.Data >= ? AND MovimentoEstoque.TipoMov = TiposMovimentoEstoque.Codigo AND TiposMovimentoEstoque.OrdemCusto = TiposOrdemCusto.Codigo "
                asOrderBy(1) = asOrderBy(1) & "MovimentoEstoque.Data, TiposOrdemCusto.Ordem, NumIntDoc"
                
                'Se foi incluído um filtro de produtos => inclui o filtro na cláusula WHERE
                If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then asWhere(1) = asWhere(1) & "AND MovimentoEstoque.Produto = ? "
                
                'Se foi incluído um filtro de data final => inclui o filtro na cláusula WHERE
                If objReprocessamentoEst.dtDataFim <> DATA_NULA Then asWhere(1) = asWhere(1) & "AND MovimentoEstoque.Data <= ? "
                
                If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then asWhere(1) = asWhere(1) & "AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
                
                'Monta um select para ler o movimento de estoque com o NumIntDoc obtido no select feito acima
                asSelect(3) = asSelect(3) & "NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao "
                asFrom(3) = asFrom(3) & "MovimentoEstoque "
                asWhere(3) = asWhere(3) & "NumIntDoc = ? "
                asOrderBy(3) = ""
            
            'Senão
            'significa que os movimentos devem ser reprocessados por ordem de hora em que foram registrados
            Else
                
                asSelect(2) = asSelect(2) & "NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao "
                asFrom(2) = asFrom(2) & "MovimentoEstoque "
                asWhere(2) = asWhere(2) & "Data >= ? "
                asOrderBy(2) = asOrderBy(2) & "Data, Hora, NumIntDoc"
                
                'Se foi incluído um filtro de produtos => inclui o filtro na cláusula WHERE
                If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then asWhere(2) = asWhere(2) & "AND Produto = ? "
                
                'Se foi incluído um filtro de data final => inclui o filtro na cláusula WHERE
                If objReprocessamentoEst.dtDataFim <> DATA_NULA Then asWhere(2) = asWhere(2) & "AND Data <= ? "
                
                If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then asWhere(2) = asWhere(2) & "AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
                
            End If
        
        'Se for um teste de integridade do reprocessamento
        Case REPROCESSAMENTO_TESTA_INTEGRIDADE
    
                asSelect(1) = asSelect(1) & "Count(*) "
                asSelect(2) = asSelect(2) & "NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao "
                
                For iIndice = 1 To iNumSelect
                
                    asFrom(iIndice) = asFrom(iIndice) & "MovimentoEstoque "
                    asWhere(iIndice) = asWhere(iIndice) & "Data >= ? "
                    asOrderBy(iIndice) = asOrderBy(iIndice) & "Produto, Data, Almoxarifado, Hora"
                
                    'Se foi incluído um filtro de produtos => inclui o filtro na cláusula WHERE
                    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then asWhere(iIndice) = asWhere(iIndice) & "AND Produto = ? "
                    
                    'Se foi incluído um filtro de data final => inclui o filtro na cláusula WHERE
                    If objReprocessamentoEst.dtDataFim <> DATA_NULA Then asWhere(iIndice) = asWhere(iIndice) & "AND Data <= ? "
                
                    If Len(Trim(objReprocessamentoEst.sFilialEmpGrupo)) <> 0 Then asWhere(iIndice) = asWhere(iIndice) & "AND FilialEmpresa IN (" & objReprocessamentoEst.sFilialEmpGrupo & ")"
                
                Next
                
    End Select
    
    'Concatena as partes SELECT, FROM, WHERE e ORDERBY para cada select
    For iIndice = 1 To iNumSelect
        asComandoSelect(iIndice) = asSelect(iIndice) & asFrom(iIndice) & asWhere(iIndice) & asOrderBy(iIndice)
    Next
    
    Rotina_Reproc_MontaSelect = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_MontaSelect:

    Rotina_Reproc_MontaSelect = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173776)
    
    End Select
    
    Exit Function
    
End Function

Function Rotina_Reproc_ExecutaSelect(lComando As Long, objReprocessamentoEst As ClassReprocessamentoEST, sSelect As String, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, iIndiceSelect As Integer) As Long
'Executa o select que será utilizado pela fase refaz do reprocessamento

Dim lErro As Long

On Error GoTo Erro_Rotina_Reproc_ExecutaSelect

    Select Case iIndiceSelect
    
        Case 1
    
            lErro = Rotina_Reproc_ExecutaSelect_Refaz1(lComando, objReprocessamentoEst, sSelect, tItemMovEstoque, tItemMovEstoqueVar)
            If lErro <> SUCESSO Then gError 90700
        
        Case 2, 3, 4
                
            lErro = Rotina_Reproc_ExecutaSelect_Comum(lComando, objReprocessamentoEst, sSelect, tItemMovEstoque, tItemMovEstoqueVar, iIndiceSelect)
            If lErro <> SUCESSO Then gError 90701
            
    End Select
    
    Rotina_Reproc_ExecutaSelect = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_ExecutaSelect:

    Rotina_Reproc_ExecutaSelect = gErr
    
    Select Case gErr
    
        Case 90700, 90701
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173777)
    
    End Select
    
    Exit Function
    
End Function

Function Rotina_Reproc_ExecutaSelect_Refaz1(lComando As Long, objReprocessamentoEst As ClassReprocessamentoEST, sSelect As String, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant) As Long
'Executa o select que será utilizado pela fase refaz do reprocessamento

Dim lErro As Long
Dim iRetorno As Integer
'Dim vlNumIntDoc As Variant, vlCodigo  As Variant
'Dim viFilialEmpresa As Variant,
Dim vdtDataInicio As Variant
Dim vsProdutoCodigo As Variant, vdtDataFim  As Variant

On Error GoTo Erro_Rotina_Reproc_ExecutaSelect_Refaz1

    'Prepara o comando select para ser executado
    iRetorno = Comando_PrepararInt(lComando, sSelect)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90650
    
    'Prepara a variável que receberá o NumIntDoc para ser executada
    tItemMovEstoqueVar.vlNumIntDoc = CLng(tItemMovEstoqueVar.vlNumIntDoc)
    iRetorno = Comando_BindVarInt(lComando, tItemMovEstoqueVar.vlNumIntDoc)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90651
     
    'Prepara a variável que receberá o Codigo para ser executada
    tItemMovEstoqueVar.vlCodigo = CLng(tItemMovEstoqueVar.vlCodigo)
    iRetorno = Comando_BindVarInt(lComando, tItemMovEstoqueVar.vlCodigo)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90652
     
    'Prepara a variável que receberá a filialempresa para ser executada
    tItemMovEstoqueVar.viFilialEmpresa = CInt(tItemMovEstoqueVar.viFilialEmpresa)
    iRetorno = Comando_BindVarInt(lComando, tItemMovEstoqueVar.viFilialEmpresa)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90653
    
    'Prepara a variável que receberá a data do movto para ser executada
    tItemMovEstoqueVar.vdtData = DATA_NULA
    iRetorno = Comando_BindVarInt(lComando, tItemMovEstoqueVar.vdtData)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90653
    '***
    
    'Parâmetro DataInicial
    'Passa o valor do parâmetro para uma variável Variant que será executada
    vdtDataInicio = objReprocessamentoEst.dtDataInicio
    
    'Prepara a variável que está passando o parâmetro para ser executada
    iRetorno = Comando_BindVarInt(lComando, vdtDataInicio)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90654
    '***
    
    'Parâmetro Produto
    'Se foi passado um produto como parâmetro
    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
         
        'Passa o valor do parâmetro para uma variável Variant que será executada
        vsProdutoCodigo = objReprocessamentoEst.sProdutoCodigo
        
        'Prepara a variável que está passando o parâmetro para ser executada
        iRetorno = Comando_BindVarInt(lComando, vsProdutoCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90655
        
    End If
    '***
    
    'Parâmetro DataFinal
    'Se foi passada uma data final como parâmetro
    If objReprocessamentoEst.dtDataFim <> DATA_NULA Then
         
         vdtDataFim = objReprocessamentoEst.dtDataFim
         
         'Prepara a variável que está passando o parâmetro para ser executada
         iRetorno = Comando_BindVarInt(lComando, vdtDataFim)
         If (iRetorno <> AD_SQL_SUCESSO) Then gError 90656
         
    End If
    '***

    'Executa o comando
    iRetorno = Comando_ExecutarInt(lComando)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90657

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90658
          
    tItemMovEstoque.lNumIntDoc = tItemMovEstoqueVar.vlNumIntDoc
    tItemMovEstoque.lCodigo = tItemMovEstoqueVar.vlCodigo
    tItemMovEstoque.iFilialEmpresa = tItemMovEstoqueVar.viFilialEmpresa
    tItemMovEstoque.dtData = tItemMovEstoqueVar.vdtData
    
    Rotina_Reproc_ExecutaSelect_Refaz1 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_ExecutaSelect_Refaz1:

    Rotina_Reproc_ExecutaSelect_Refaz1 = gErr
    
    Select Case gErr
    
        Case 90650 To 90658
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173778)
    
    End Select
    
    Exit Function
    
End Function

Function Rotina_Reproc_ExecutaSelect_Comum(lComando As Long, objReprocessamentoEst As ClassReprocessamentoEST, sSelect As String, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, iSelect As Integer) As Long
'Executa o select que será utilizado pela fase refaz do reprocessamento

Dim lErro As Long
Dim iRetorno As Integer

'Dim colCampos As New Collection
'Dim vlNumIntDoc As Variant, viFilialEmpresa As Variant, vlCodigo As Variant, vdCusto As Variant, viApropriacao As Variant
'Dim vsProduto As Variant, vsSiglaUM As Variant, vdQuantidade As Variant, viAlmoxarifado As Variant, viTipoMov As Variant
'Dim vlNumIntDocOrigem As Variant, viTipoNumIntDocOrigem As Variant, vdtData As Variant, vdHora As Variant, vsCcl As Variant
'Dim vlNumIntDocEst As Variant, vlCliente As Variant, vlFornecedor As Variant, vsOPCodigo As Variant, vsDocOrigem As Variant
'Dim vsContaContabilEst As Variant, vsContaContabilAplic As Variant, vlHorasMaquina As Variant, vdtDataInicioProducao As Variant

On Error GoTo Erro_Rotina_Reproc_ExecutaSelect_Comum
                
    'Inicializa as variáveis strings que serão utilizadas no select
    tItemMovEstoqueVar.vsProduto = String(STRING_PRODUTO, 0)
    tItemMovEstoqueVar.vsSiglaUM = String(STRING_UM_SIGLA, 0)
    tItemMovEstoqueVar.vsCcl = String(STRING_CCL, 0)
    tItemMovEstoqueVar.vsOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
    tItemMovEstoqueVar.vsDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
    tItemMovEstoqueVar.vsContaContabilAplic = String(STRING_CONTA, 0)
    tItemMovEstoqueVar.vsContaContabilEst = String(STRING_CONTA, 0)
    
'    le os movimentos de estoque em ordem crescente de data, hora e numintdoc desde a data passada como parametro
'    lErro = Comando_ExecutarPos(lComando, sSelect, 0, tItemMovEstoque.lNumIntDoc, tItemMovEstoque.iFilialEmpresa, tItemMovEstoque.lCodigo, tItemMovEstoque.dCusto, tItemMovEstoque.iApropriacao, tItemMovEstoque.sProduto, tItemMovEstoque.sSiglaUM, tItemMovEstoque.dQuantidade, tItemMovEstoque.iAlmoxarifado, tItemMovEstoque.iTipoMov, tItemMovEstoque.lNumIntDocOrigem, tItemMovEstoque.iTipoNumIntDocOrigem, tItemMovEstoque.dtData, tItemMovEstoque.dHora, tItemMovEstoque.sCcl, tItemMovEstoque.lNumIntDocEst, tItemMovEstoque.lCliente, tItemMovEstoque.lFornecedor, tItemMovEstoque.sOPCodigo, tItemMovEstoque.sDocOrigem, tItemMovEstoque.sContaContabilEst, tItemMovEstoque.sContaContabilAplic, tItemMovEstoque.lHorasMaquina, tItemMovEstoque.dtDataInicioProducao, objReprocessamentoEST.iFilialEmpresa, objReprocessamentoEST.dtDataInicio)
'    If lErro <> AD_SQL_SUCESSO Then gError 83566
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 83567
    
    'Prepara o comando select para ser executado
    iRetorno = Comando_PrepararInt(lComando, sSelect)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90659
    
    With tItemMovEstoqueVar
        'Prepara a variável que receberá o NumIntDoc para ser executada
        .vlNumIntDoc = CLng(.vlNumIntDoc)
        iRetorno = Comando_BindVarInt(lComando, .vlNumIntDoc)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90660
         
        'Prepara a variável que receberá a FilialEmpresa para ser executada
        .viFilialEmpresa = CInt(.viFilialEmpresa)
        iRetorno = Comando_BindVarInt(lComando, .viFilialEmpresa)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90661
         
        'Prepara a variável que receberá o Código para ser executada
        .vlCodigo = CLng(.vlCodigo)
        iRetorno = Comando_BindVarInt(lComando, .vlCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90662
    
        'Prepara a variável que receberá o Custo para ser executada
        .vdCusto = CDbl(.vdCusto)
        iRetorno = Comando_BindVarInt(lComando, .vdCusto)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90663
    
        'Prepara a variável que receberá a Apropriação para ser executada
        .viApropriacao = CInt(.viApropriacao)
        iRetorno = Comando_BindVarInt(lComando, .viApropriacao)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90664
    
        'Prepara a variável que receberá o Produto para ser executada
        '.vsProduto = CStr(.vsProduto)
        iRetorno = Comando_BindVarInt(lComando, .vsProduto)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90665
    
        'Prepara a variável que receberá a SiglaUM para ser executada
        '.vsSiglaUM = CStr(.vsSiglaUM)
        iRetorno = Comando_BindVarInt(lComando, .vsSiglaUM)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90666
        
        'Prepara a variável que receberá a Quantidade para ser executada
        .vdQuantidade = CDbl(.vdQuantidade)
        iRetorno = Comando_BindVarInt(lComando, .vdQuantidade)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90667
        
        'Prepara a variável que receberá o Almoxarifado para ser executada
        .viAlmoxarifado = CInt(.viAlmoxarifado)
        iRetorno = Comando_BindVarInt(lComando, .viAlmoxarifado)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90668
        
        'Prepara a variável que receberá o Tipo de Movimento para ser executada
        .viTipoMov = CInt(.viTipoMov)
        iRetorno = Comando_BindVarInt(lComando, .viTipoMov)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90669
        
        'Prepara a variável que receberá o NumIntDocOrigem para ser executada
        .vlNumIntDocOrigem = CLng(.vlNumIntDocOrigem)
        iRetorno = Comando_BindVarInt(lComando, .vlNumIntDocOrigem)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90670
        
        'Prepara a variável que receberá o TipoNumIntDocOrigem para ser executada
        .viTipoNumIntDocOrigem = CInt(.viTipoNumIntDocOrigem)
        iRetorno = Comando_BindVarInt(lComando, .viTipoNumIntDocOrigem)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90671
        
        'Prepara a variável que receberá a Data do movimento para ser executada
        .vdtData = CDate(.vdtData)
        iRetorno = Comando_BindVarInt(lComando, .vdtData)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90672
        
        'Prepara a variável que receberá a Hora do movimento para ser executada
        .vdHora = CDbl(.vdHora)
        iRetorno = Comando_BindVarInt(lComando, .vdHora)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90673
        
        'Prepara a variável que receberá o Ccl para ser executada
        '.vsCcl = CStr(.vsCcl)
        iRetorno = Comando_BindVarInt(lComando, .vsCcl)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90674
        
        'Prepara a variável que receberá o NumIntDocEst para ser executada
        .vlNumIntDocEst = CLng(.vlNumIntDocEst)
        iRetorno = Comando_BindVarInt(lComando, .vlNumIntDocEst)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90675
    
        'Prepara a variável que receberá o Cliente para ser executada
        .vlCliente = CLng(.vlCliente)
        iRetorno = Comando_BindVarInt(lComando, .vlCliente)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90676
    
        'Prepara a variável que receberá o Fornecedor para ser executada
        .vlFornecedor = CLng(.vlFornecedor)
        iRetorno = Comando_BindVarInt(lComando, .vlFornecedor)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90677
    
        'Prepara a variável que receberá o Códsigo da OP para ser executada
        '.vsOPCodigo = CStr(.vsOPCodigo)
        iRetorno = Comando_BindVarInt(lComando, .vsOPCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90678
    
        'Prepara a variável que receberá o DocOrigem para ser executada
        '.vsDocOrigem = CStr(.vsDocOrigem)
        iRetorno = Comando_BindVarInt(lComando, .vsDocOrigem)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90679
    
        'Prepara a variável que receberá a ContaContabilEst para ser executada
        '.vsContaContabilEst = CStr(.vsContaContabilEst)
        iRetorno = Comando_BindVarInt(lComando, .vsContaContabilEst)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90680
    
        'Prepara a variável que receberá a ContaContabilAplic para ser executada
        '.vsContaContabilAplic = CStr(.vsContaContabilAplic)
        iRetorno = Comando_BindVarInt(lComando, .vsContaContabilAplic)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90681
    
        'Prepara a variável que receberá a quant. de HorasMáquina para ser executada
        .vlHorasMaquina = CLng(.vlHorasMaquina)
        iRetorno = Comando_BindVarInt(lComando, .vlHorasMaquina)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90682
    
        'Prepara a variável que receberá a DataInicioProducao para ser executada
        .vdtDataInicioProducao = CDate(.vdtDataInicioProducao)
        iRetorno = Comando_BindVarInt(lComando, .vdtDataInicioProducao)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90683
    
    End With
    
    Select Case iSelect
    
        Case REPROCESSAMENTO_SELECT_REFAZ2, REPROCESSAMENTO_SELECT_DESFAZ1, REPROCESSAMENTO_SELECT_TESTAINTEGRIDADE2
        
            lErro = Rotina_Reproc_ExecutaSelect_Comum2(lComando, tItemMovEstoque, tItemMovEstoqueVar, objReprocessamentoEst)
            If lErro <> SUCESSO Then gError 90698
        
        Case REPROCESSAMENTO_SELECT_REFAZ3
        
            lErro = Rotina_Reproc_ExecutaSelect_Refaz3(lComando, tItemMovEstoque, tItemMovEstoqueVar, objReprocessamentoEst)
            If lErro <> SUCESSO Then gError 90699
        
'        Case REPROCESSAMENTO_DESFAZ
'
'            lErro = Rotina_Reproc_ExecutaSelect_Desfaz(lComando, tItemMovEstoque, tItemMovEstoqueVar, objReprocessamentoEst)
'            If lErro <> SUCESSO Then gError 94508
    
    End Select
    
    Rotina_Reproc_ExecutaSelect_Comum = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_ExecutaSelect_Comum:

    Rotina_Reproc_ExecutaSelect_Comum = gErr
    
    Select Case gErr
        
        Case 90698, 90699, 94508
        
        Case 90659 To 90683
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173779)
    
    End Select
    
    Exit Function
    
End Function

Function Rotina_Reproc_ExecutaSelect_Comum2(lComando As Long, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, objReprocessamentoEst As ClassReprocessamentoEST) As Long
    
Dim lErro As Long
Dim iRetorno As Integer
Dim viFilialEmpresaMov As Variant, vdtDataInicio As Variant, vsProdutoCodigo As Variant, vdtDataFim As Variant

On Error GoTo Erro_Rotina_Reproc_ExecutaSelect_Comum2

'    'Parâmetro FilialEmpresa
'    'Passa o valor do parâmetro para uma variável Variant que será executada
'    viFilialEmpresaMov = objReprocessamentoEst.iFilialEmpresa
'
'    'Prepara a variável que está passando o parâmetro para ser executada
'    iRetorno = Comando_BindVarInt(lComando, viFilialEmpresaMov)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90685
'    '***
    
    'Parâmetro DataInicial
    'Passa o valor do parâmetro para uma variável Variant que será executada
    vdtDataInicio = objReprocessamentoEst.dtDataInicio
    
    'Prepara a variável que está passando o parâmetro para ser executada
    iRetorno = Comando_BindVarInt(lComando, vdtDataInicio)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90686
    '***

    'Parâmetro Produto
    'Se foi passado um produto como parâmetro
    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
         
        vsProdutoCodigo = objReprocessamentoEst.sProdutoCodigo
        
        'Prepara a variável que está passando o parâmetro para ser executada
        iRetorno = Comando_BindVarInt(lComando, vsProdutoCodigo)
        If (iRetorno <> AD_SQL_SUCESSO) Then gError 90687
        
    End If
    '***

    'Parâmetro DataFinal
    'Se foi passada uma data final como parâmetro
    If objReprocessamentoEst.dtDataFim <> DATA_NULA Then
         
         vdtDataFim = objReprocessamentoEst.dtDataFim
         
         'Prepara a variável que está passando o parâmetro para ser executada
         iRetorno = Comando_BindVarInt(lComando, vdtDataFim)
         If (iRetorno <> AD_SQL_SUCESSO) Then gError 90688
         
    End If

    'Executa o comando
    iRetorno = Comando_ExecutarInt(lComando)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90689

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90690
    
    Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
    
    Rotina_Reproc_ExecutaSelect_Comum2 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_ExecutaSelect_Comum2:

    Rotina_Reproc_ExecutaSelect_Comum2 = gErr
    
    Select Case gErr
    
        Case 90685 To 90690
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173780)
        
    End Select
    
    Exit Function

End Function

Function Rotina_Reproc_ExecutaSelect_Refaz3(lComando As Long, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, objReprocessamentoEst As ClassReprocessamentoEST) As Long
    
Dim lErro As Long
Dim iRetorno As Integer
Dim vlNumIntDocMov As Variant

On Error GoTo Erro_Rotina_Reproc_ExecutaSelect_Refaz3

    'Parâmetro NumIntDoc
    'Passa o valor do parâmetro para uma variável Variant que será executada
    vlNumIntDocMov = tItemMovEstoque.lNumIntDoc
    
    'Prepara a variável que está passando o parâmetro para ser executada
    iRetorno = Comando_BindVarInt(lComando, vlNumIntDocMov)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90695
    '***

    'Executa o comando
    iRetorno = Comando_ExecutarInt(lComando)
    If (iRetorno <> AD_SQL_SUCESSO) Then gError 90696

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 90697
    
    Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
    
    Rotina_Reproc_ExecutaSelect_Refaz3 = SUCESSO
    
    Exit Function
    
Erro_Rotina_Reproc_ExecutaSelect_Refaz3:

    Rotina_Reproc_ExecutaSelect_Refaz3 = gErr
    
    Select Case gErr
    
        Case 90695 To 90697
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173781)
        
    End Select
    
    Exit Function
    
End Function

Sub Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant)

On Error GoTo Erro_Move_tItemMovEstoqueVariant_tItemMovEstoque

    With tItemMovEstoque

        .lNumIntDoc = tItemMovEstoqueVar.vlNumIntDoc
        .iFilialEmpresa = tItemMovEstoqueVar.viFilialEmpresa
        .lCodigo = tItemMovEstoqueVar.vlCodigo
        .dCusto = tItemMovEstoqueVar.vdCusto
        .iApropriacao = tItemMovEstoqueVar.viApropriacao
        .sProduto = tItemMovEstoqueVar.vsProduto
        .sSiglaUM = tItemMovEstoqueVar.vsSiglaUM
        .dQuantidade = tItemMovEstoqueVar.vdQuantidade
        .iAlmoxarifado = tItemMovEstoqueVar.viAlmoxarifado
        .iTipoMov = tItemMovEstoqueVar.viTipoMov
        .lNumIntDocOrigem = tItemMovEstoqueVar.vlNumIntDocOrigem
        .iTipoNumIntDocOrigem = tItemMovEstoqueVar.viTipoNumIntDocOrigem
        .dtData = tItemMovEstoqueVar.vdtData
        .dHora = tItemMovEstoqueVar.vdHora
        .sCcl = tItemMovEstoqueVar.vsCcl
        .lNumIntDocEst = tItemMovEstoqueVar.vlNumIntDocEst
        .lCliente = tItemMovEstoqueVar.vlCliente
        .lFornecedor = tItemMovEstoqueVar.vlFornecedor
        .sOPCodigo = tItemMovEstoqueVar.vsOPCodigo
        .sDocOrigem = tItemMovEstoqueVar.vsDocOrigem
        .sContaContabilEst = tItemMovEstoqueVar.vsContaContabilEst
        .sContaContabilAplic = tItemMovEstoqueVar.vsContaContabilAplic
        .lHorasMaquina = tItemMovEstoqueVar.vlHorasMaquina
        .dtDataInicioProducao = tItemMovEstoqueVar.vdtDataInicioProducao

    End With

    Exit Sub
    
Erro_Move_tItemMovEstoqueVariant_tItemMovEstoque:

    Select Case gErr
    
    End Select
    
    Exit Sub

End Sub

'Function Rotina_Reproc_ExecutaSelect4(lComando As Long, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, objReprocessamentoEst As ClassReprocessamentoEST) As Long
'
'Dim lErro As Long
'Dim iRetorno As Integer
'Dim viFilialEmpresaMov As Variant, vdtDataInicio As Variant, vsProdutoCodigo As Variant, vdtDataFim As Variant
'
'On Error GoTo Erro_Rotina_Reproc_ExecutaSelect4
'
'    'Parâmetro FilialEmpresa
'    'Passa o valor do parâmetro para uma variável Variant que será executada
'    viFilialEmpresaMov = objReprocessamentoEst.iFilialEmpresa
'
'    'Prepara a variável que está passando o parâmetro para ser executada
'    iRetorno = Comando_BindVarInt(lComando, viFilialEmpresaMov)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError 94502
'    '***
'
'    'Parâmetro DataInicial
'    'Passa o valor do parâmetro para uma variável Variant que será executada
'    vdtDataInicio = objReprocessamentoEst.dtDataInicio
'
'    'Prepara a variável que está passando o parâmetro para ser executada
'    iRetorno = Comando_BindVarInt(lComando, vdtDataInicio)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError 94503
'    '***
'
'    'Parâmetro Produto
'    'Se foi passado um produto como parâmetro
'    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
'
'        vsProdutoCodigo = objReprocessamentoEst.sProdutoCodigo
'
'        'Prepara a variável que está passando o parâmetro para ser executada
'        iRetorno = Comando_BindVarInt(lComando, vsProdutoCodigo)
'        If (iRetorno <> AD_SQL_SUCESSO) Then gError 94504
'
'    End If
'    '***
'
'    'Parâmetro DataFinal
'    'Se foi passada uma data final como parâmetro
'    If objReprocessamentoEst.dtDataFim <> DATA_NULA Then
'
'         vdtDataFim = objReprocessamentoEst.dtDataFim
'
'         'Prepara a variável que está passando o parâmetro para ser executada
'         iRetorno = Comando_BindVarInt(lComando, vdtDataFim)
'         If (iRetorno <> AD_SQL_SUCESSO) Then gError 94505
'
'    End If
'
'    'Executa o comando
'    iRetorno = Comando_ExecutarInt(lComando)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError 94506
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94507
'
'    Call Move_tItemMovEstoqueVariant_tItemMovEstoque(tItemMovEstoque, tItemMovEstoqueVar)
'
'    Rotina_Reproc_ExecutaSelect4 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_Reproc_ExecutaSelect4:
'
'    Rotina_Reproc_ExecutaSelect4 = gErr
'
'    Select Case gErr
'
'        Case 94502 To 94507
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173782)
'
'    End Select
'
'    Exit Function
'
'End Function

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Function Rotina_Reproc_TestaIntegridade_Int(objReprocessamentoEst As ClassReprocessamentoEST) As Long
''Testa a integridade do último reprocessamento executado no BD
'
'Dim lTransacao As Long
'Dim lErro As Long
'Dim alComando(3) As Long
'Dim iIndice As Integer
'Dim lTotalNumMovEst As Long
'Dim asComandoSelect(NUM_SELECT_REPROC_TESTAINTEGRIDADE) As String
'Dim objMATConfig As New ClassMATConfig
'
'On Error GoTo Erro_Rotina_Reproc_TestaIntegridade_Int
'
'    'Inicia a transação
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError xxx
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError xxx
'    Next
'
'    'faz o lock para impedir qualquer movimentação do estoque durate o reprocessamento
'    objMATConfig.iFilialEmpresa = objReprocessamentoEst.iFilialEmpresa
'    objMATConfig.sCodigo = DATA_REPROCESSAMENTO
'
'    lErro = CF("MATConfig_Le_Lock1",objMATConfig, alComando(0))
'    If lErro <> SUCESSO And lErro <> 83776 Then gError xxx
'
'    'faz o lock para impedir qualquer movimentação do estoque durate o reprocessamento
'    objMATConfig.iFilialEmpresa = EMPRESA_TODA
'    objMATConfig.sCodigo = NUM_PROX_ITEM_MOV_ESTOQUE
'
'    lErro = CF("MATConfig_Le_Lock",objMATConfig, alComando(1))
'    If lErro <> SUCESSO Then gError xxx
'
'    lNumIntDoc = CLng(objMATConfig.sConteudo)
'
'    lErro = Rotina_Reproc_MontaSelect(objReprocessamentoEst, asComandoSelect, REPROCESSAMENTO_TESTA_INTEGRIDADE)
'    If lErro <> SUCESSO Then gError xxx
'
'    lErro = Rotina_Reproc_ExecutaSelect_TestaIntegridade1(alComando(2), objReprocessamentoEst, asComandoSelect(1), tItemMovEstoque, tItemMovEstoqueVar, lTotalNumMovEst)
'    If lErro <> SUCESSO Then gError xxx
'
'    'Tela acompanhamento Batch inicializa dValorTotal
'    TelaAcompanhaBatchEST.dValorTotal = 2 * lTotalNumMovEst
'
'    '??? chamar função
'
'    'libera comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'Confirma a transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError xxx
'
'    Rotina_Reproc_TestaIntegridade_Int = SUCESSO
'
'    Exit Function
'
'Rotina_Reproc_TestaIntegridade_Int:
'
'    Rotina_Reproc_TestaIntegridade_Int = gErr
'
'    Select Case gErr
'
'        Case 83544, 83763, 83765, 83766, 83767, 83772, 90584
'
'        Case 83576
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
'
'       Case 83577, 83641
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 83764
'            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_INEXISTENTE", gErr, objEstoqueMes.iFilialEmpresa, objEstoqueMes.iAno, objEstoqueMes.iMes)
'
'        Case 83768
'            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
'
'        Case 83777
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_REPROC_MENOR_DATA_INICIO", gErr, CDate(objMATConfig.sConteudo), objReprocessamentoEst.dtDataInicio)
'
'        Case 83780
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MATCONFIG", gErr, objMATConfig.sCodigo, objMATConfig.iFilialEmpresa)
'
'        Case 89449
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESATIVACAO_LOCKS", gErr)
'
'        Case 89450
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_REATIVACAO_LOCKS", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173783)
'
'    End Select
'
'    'reativa os locks
'    Call Conexao_DesativarLocks(REATIVAR_LOCKS)
'
'    'Rollback
'    Call Transacao_Rollback
'
'   'Fechamento comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    'libera comandos
'    For iIndice = LBound(alComando1) To UBound(alComando1)
'        Call Comando_Fechar(alComando1(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Function Rotina_Reproc_ExecutaSelect_TestaIntegridade1(lComando As Long, objReprocessamentoEst As ClassReprocessamentoEST, sSelect As String, tItemMovEstoque As typeItemMovEstoque, tItemMovEstoqueVar As typeItemMovEstoqueVariant, vlTotalNumMovEstoque As Variant) As Long
''Executa o select que será utilizado pela fase refaz do reprocessamento
'
'Dim lErro As Long
'Dim iRetorno As Integer
''Dim vlNumIntDoc As Variant, vlCodigo  As Variant
'Dim viFilialEmpresa As Variant, vdtDataInicio As Variant
'Dim vsProdutoCodigo As Variant, vdtDataFim  As Variant
'
'On Error GoTo Erro_Rotina_Reproc_ExecutaSelect_TestaIntegridade1
'
'    'Prepara o comando select para ser executado
'    iRetorno = Comando_PrepararInt(lComando, sSelect)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'
'    'Prepara a variável que receberá o NumIntDoc para ser executada
'    iRetorno = Comando_BindVarInt(lComando, vlTotalNumMovEstoque)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'
'    'Parâmetro FilialEmpresa
'    'Passa o valor do parâmetro para uma variável Variant que será executada
'    viFilialEmpresa = objReprocessamentoEst.iFilialEmpresa
'
'    'Prepara a variável que está passando o parâmetro para ser executada
'    iRetorno = Comando_BindVarInt(lComando, viFilialEmpresa)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'    '***
'
'    'Parâmetro DataInicial
'    'Passa o valor do parâmetro para uma variável Variant que será executada
'    vdtDataInicio = objReprocessamentoEst.dtDataInicio
'
'    'Prepara a variável que está passando o parâmetro para ser executada
'    iRetorno = Comando_BindVarInt(lComando, vdtDataInicio)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'    '***
'
'    'Parâmetro Produto
'    'Se foi passado um produto como parâmetro
'    If Len(Trim(objReprocessamentoEst.sProdutoCodigo)) > 0 Then
'
'        'Passa o valor do parâmetro para uma variável Variant que será executada
'        vsProdutoCodigo = objReprocessamentoEst.sProdutoCodigo
'
'        'Prepara a variável que está passando o parâmetro para ser executada
'        iRetorno = Comando_BindVarInt(lComando, vsProdutoCodigo)
'        If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'
'    End If
'    '***
'
'    'Parâmetro DataFinal
'    'Se foi passada uma data final como parâmetro
'    If objReprocessamentoEst.dtDataFim <> DATA_NULA Then
'
'         vdtDataFim = objReprocessamentoEst.dtDataFim
'
'         'Prepara a variável que está passando o parâmetro para ser executada
'         iRetorno = Comando_BindVarInt(lComando, vdtDataFim)
'         If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'
'    End If
'    '***
'
'    'Executa o comando
'    iRetorno = Comando_ExecutarInt(lComando)
'    If (iRetorno <> AD_SQL_SUCESSO) Then gError xxx
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError xxx
'
'    Rotina_Reproc_ExecutaSelect_TestaIntegridade1 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_Reproc_ExecutaSelect_TestaIntegridade1:
'
'    Rotina_Reproc_ExecutaSelect_TestaIntegridade1 = gErr
'
'    Select Case gErr
'
'        Case 90650 To 90658
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MOVIMENTOESTOQUE2", gErr, objReprocessamentoEst.iFilialEmpresa)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173784)
'
'    End Select
'
'    Exit Function
'
'End Function
'

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Function Rotina_Reproc_TestaIntegridade_Int1(objReprocessamentoEst As ClassReprocessamentoEST) As Long
'
'Dim iIndice As Integer
'Dim lErro As Long
'Dim alComando(1) As Long
'Dim tItemMovEstoque As typeItemMovEstoque
'Dim tItemMovEstoqueVar As typeItemMovEstoqueVariant
'Dim sProduto As String
'Dim iAlmoxarifado As Integer
'Dim dtDataMovimento As Date
'Dim dtDataInicio As Date
'Dim dtDataFim As Date
'Dim colSldMesEstAlmAcumulado As Collection
'Dim objItemMovEstoque As ClassItemMovEstoque
'Dim objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm
'Dim objSldDiaEstAcumulado As ClassSldDiaEst
'Dim objSldMesEstAlmAcumulado As ClassSldMesEstAlm
'Dim objSldMesEstAcumulado As ClassSldMesEst
'Dim iOcorreuErro As Integer
'
'On Error GoTo Erro_Rotina_Reproc_TestaIntegridade_Int1
'
'    'Abre comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError xxx
'    Next
'
'    lErro = Rotina_Reproc_ExecutaSelect_Comum(alComando(1), objReprocessamentoEst, sSelect, tItemMovEstoque, tItemMovEstoqueVar, REPROCESSAMENTO_SELECT_TESTAINTEGRIDADE2)
'    If lErro <> SUCESSO Then gError xxx
'
'    Set objItemMovEstoque = New ClassItemMovEstoque
'
'    Call Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque, objItemMovEstoque)
'
'    sProduto = tItemMovEstoque.sProduto
'    iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
'    dtDataMovimento = tItemMovEstoque.dtData
'
'    'Set ColItensMovEstoque = New Collection
'
'    Do While lErro = SUCESSO
'
'        With objItemMovEstoque
'
'            Set objItemMovEstoque = New ClassItemMovEstoque
'
'            Call Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque, objItemMovEstoque)
'
'            If .sProduto = sProduto And .dtData = dtDataMovimento And .iAlmoxarifado = iAlmoxarifado Then
'
'                lErro = Estoque_Acumula_ItemMovEst_SldDiaAlm(objItemMovEstoque, objSldDiaEstAlmAcumulado)
'                If lErro <> SUCESSO Then gError xxx
'
'            Else
'
'                objSldDiaEstAlmAcumulado.sProduto = sProduto
'                objSldDiaEstAlmAcumulado.dtData = dtDataMovimento
'                objSldDiaEstAlmAcumulado.iAlmoxarifado = iAlmoxarifado
'
'                lErro = Estoque_Compara_SldDiaEstAlm(objSldDiaEstAlmAcumulado)
'                If lErro <> SUCESSO And lErro <> 94547 Then gError xxx
'
'                lErro = Estoque_Acumula_SldDia(objSldDiaEstAlmAcumulado, objSldDiaEstAcumulado)
'                If lErro <> SUCESSO Then gError xxx
'
'                If .dtData <> dtDataMovimento Or .sProduto <> sProduto Then
'
'                    lErro = Estoque_Compara_SldDiaEst(objSldDiaEstAcumulado)
'                    If lErro <> SUCESSO And lErro <> 94547 Then gError xxx
'
'                    Set objSldDiaEstAcumulado = New ClassSldDiaEst
'
'                    If .sProduto <> sProduto Then
'
'                        dtDataInicio = StrParaDate("01/" & Month(objReprocessamentoEst.dtDataInicio) & "/" & Year(objReprocessamentoEst.dtDataInicio))
'                        dtDataFim = StrParaDate("01/" & Month(objReprocessamentoEst.dtDataFim) & "/" & Year(objReprocessamentoEst.dtDataFim)) - 1
'
'                        lErro = Estoque_Acumula_SldDiaEstAlm(dtDataInicio, dtDataFim, colSldMesEstAlmAcumulado)
'                        If lErro <> SUCESSO Then gError xxx
'
'                        '??? comparar com saldomesestalm
'
'                        lErro = Estoque_Acumula_SldMesEstAlm(colSldMesEstAlmAcumulado, objSldMesEstAcumulado)
'                        If lErro <> SUCESSO Then gError xxx
'
'                    End If
'
'                End If
'
'                Set objSldDiaEstAlmAcumulado = New ClassSldDiaEstAlm
'
'                lErro = Estoque_Acumula_ItemMovEst_SldDiaAlm(objItemMovEstoque, objSldDiaEstAlmAcumulado)
'                If lErro <> SUCESSO Then gError xxx
'
'            End If
'
'            sProduto = tItemMovEstoque.sProduto
'            iAlmoxarifado = tItemMovEstoque.iAlmoxarifado
'            dtDataMovimento = tItemMovEstoque.dtData
'
'            lErro = Comando_BuscarProximo(alComando(1))
'            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError xxx
'
'    Loop
'
'
'
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Comando_Fechar (alComando(iIndice))
'    Next
'
'    Rotina_Reproc_TestaIntegridade_Int1 = SUCESSO
'
'    Exit Function
'
'Erro_Rotina_Reproc_TestaIntegridade_Int1:
'
'    Rotina_Reproc_TestaIntegridade_Int1 = gErr
'
'    Select Case gErr
'
'        Case Else
'
'    End Select
'
'    'Fecha os comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Comando_Fechar (alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

'Criada por: Luiz G.F.Nogueira
'em: 19/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function Estoque_Acumula_ItemMovEst_SldDiaAlm(objItemMovEstoque As ClassItemMovEstoque, objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm) As Long
'Recebe um item de movimento de estoque e acumula seus valores em um objSldDiaEstAlm

Dim lErro As Long
Dim lComando As Long
Dim objTipoMovEstoque As ClassTipoMovEst

On Error GoTo Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94533

    'Instancia o objTipoMovEstoque
    Set objTipoMovEstoque = New ClassTipoMovEst
    'Set objEstoqueProduto = New ClassEstoqueProduto
    
    'Lê o tipo de movimento de estoque do item que será processado
    lErro = CF("TiposMovEst_Le1", lComando, objTipoMovEstoque)
    If lErro <> SUCESSO Then gError 94534
    
    'Acumula uma parte dos valores a serem acumulados
    lErro = Estoque_Acumula_ItemMovEst_SldDiaAlm1(objItemMovEstoque, objTipoMovEstoque, objSldDiaEstAlmAcumulado)
    If lErro <> SUCESSO Then gError 94535
    
    'Acumula outra parte dos valores a serem acumulados
    lErro = Estoque_Acumula_ItemMovEst_SldDiaAlm2(objItemMovEstoque, objTipoMovEstoque, objSldDiaEstAlmAcumulado)
    If lErro <> SUCESSO Then gError 94536
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Estoque_Acumula_ItemMovEst_SldDiaAlm = SUCESSO
    
    Exit Function
    
Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm:

    Estoque_Acumula_ItemMovEst_SldDiaAlm = gErr
    
    Select Case gErr
    
        Case 94533
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94534 To 94536
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173785)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Criada por: Luiz G.F.Nogueira
'em: 19/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Private Function Estoque_Acumula_ItemMovEst_SldDiaAlm1(objItemMovEstoque As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm) As Long
'Chama a função Estoque_AtualizaItemMov3, que seta as variáveis com os valores e os sinais devidos
'Depois acumula em um objSldDiaEstAlm os valores devolvido por Estoque_AtualizaItemMov3

Dim lErro As Long
Dim objSldDiaEstAlm As ClassSldDiaEstAlm
Dim objEstoqueProduto As New ClassEstoqueProduto 'esse obj não serve para nada, apenas é declarado para ser passado para Estoque_AtualizaItemMov3

On Error GoTo Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm1

        'Guarda em objEstoqueProduto e em objSldDiaEst os valores e os sinais com os quais deverão ser acumulados
        lErro = CF("Estoque_AtualizaItemMov3", objItemMovEstoque, objTipoMovEstoque, objEstoqueProduto, objSldDiaEstAlm, REPROCESSAMENTO_REFAZ)
        If lErro <> SUCESSO Then gError 94527
        
        'Acumula os valores no objSldDiaEstAlmAcumulado
        With objSldDiaEstAlmAcumulado
            
            .dQuantEntRecIndl = .dQuantEntRecIndl + objSldDiaEstAlm.dQuantEntRecIndl
            .dValorEntRecIndl = .dValorEntRecIndl + objSldDiaEstAlm.dValorEntRecIndl
            .dQuantSaiRecIndl = .dQuantSaiRecIndl + objSldDiaEstAlm.dQuantSaiRecIndl
            .dValorSaiRecIndl = .dValorSaiRecIndl + objSldDiaEstAlm.dValorSaiRecIndl
            .dQuantEntrada = .dQuantEntrada + objSldDiaEstAlm.dQuantEntrada
            .dValorEntrada = .dValorEntrada + objSldDiaEstAlm.dValorEntrada
            .dValorEntrada = .dValorEntrada + objSldDiaEstAlm.dQuantSaida
            .dValorSaida = .dValorSaida + objSldDiaEstAlm.dValorSaida
            .dQuantCons = .dQuantCons + objSldDiaEstAlm.dQuantCons
            .dValorCons = .dValorCons + objSldDiaEstAlm.dValorCons
            .dQuantVend = .dQuantVend + objSldDiaEstAlm.dQuantVend
            .dValorVend = .dValorVend + objSldDiaEstAlm.dValorVend
            .dQuantVendConsig3 = .dQuantVendConsig3 + objSldDiaEstAlm.dQuantVendConsig3
            .dValorVendConsig3 = .dValorVendConsig3 + objSldDiaEstAlm.dValorVendConsig3
            .dQuantComp = .dQuantComp + objSldDiaEstAlm.dQuantComp
            .dValorComp = .dValorComp + objSldDiaEstAlm.dValorComp
        
        End With
        
        Estoque_Acumula_ItemMovEst_SldDiaAlm1 = SUCESSO
        
        Exit Function

Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm1:

    Estoque_Acumula_ItemMovEst_SldDiaAlm1 = gErr
    
    Select Case gErr
        
        Case 94527
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173786)
    
    End Select
    
    Exit Function

End Function

'Criada por: Luiz G.F.Nogueira
'em: 19/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Private Function Estoque_Acumula_ItemMovEst_SldDiaAlm2(objItemMovEstoque As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm) As Long
'Chama a função Estoque_AtualizaItemMov4, que seta as variáveis com os valores e os sinais devidos
'Depois acumula em um objSldDiaEstAlm os valores devolvido por Estoque_AtualizaItemMov4

Dim lErro As Long
Dim objSldDiaEstAlm As ClassSldDiaEstAlm
Dim objEstoqueProduto As New ClassEstoqueProduto 'esse obj não serve para nada, apenas é declarado para ser passado para Estoque_AtualizaItemMov3

On Error GoTo Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm2

        'Guarda em objEstoqueProduto e em objSldDiaEstAlm os valores e os sinais com os quais deverão ser acumulados
        lErro = CF("Estoque_AtualizaItemMov4", objItemMovEstoque, objTipoMovEstoque, objEstoqueProduto, objSldDiaEstAlm)
        If lErro <> SUCESSO Then gError 94528
        
        'Acumula os valores no objSldDiaEstAlmAcumulado
        With objSldDiaEstAlmAcumulado
            
            .dQuantEntCusto = .dQuantEntCusto + objSldDiaEstAlm.dQuantEntCusto
            .dValorEntCusto = .dValorEntCusto + objSldDiaEstAlm.dValorEntCusto
            .dQuantSaiCusto = .dQuantSaiCusto + objSldDiaEstAlm.dQuantSaiCusto
            .dValorSaiCusto = .dValorSaiCusto + objSldDiaEstAlm.dValorSaiCusto
            .dQuantEntConsig3 = .dQuantEntConsig3 + objSldDiaEstAlm.dQuantEntConsig3
            .dValorEntConsig3 = .dValorEntConsig3 + objSldDiaEstAlm.dValorEntConsig3
            .dQuantSaiConsig3 = .dQuantSaiConsig3 + objSldDiaEstAlm.dQuantSaiConsig3
            .dValorSaiConsig3 = .dValorSaiConsig3 + objSldDiaEstAlm.dValorSaiConsig3
            .dQuantEntDemo3 = .dQuantEntDemo3 + objSldDiaEstAlm.dQuantEntDemo3
            .dValorEntDemo3 = .dValorEntDemo3 + objSldDiaEstAlm.dValorEntDemo3
            .dQuantSaiDemo3 = .dQuantSaiDemo3 + objSldDiaEstAlm.dQuantSaiDemo3
            .dValorSaiDemo3 = .dValorSaiDemo3 + objSldDiaEstAlm.dValorSaiDemo3
            .dQuantEntConserto3 = .dQuantEntConserto3 + objSldDiaEstAlm.dQuantEntConserto3
            .dValorEntConserto3 = .dValorEntConserto3 + objSldDiaEstAlm.dValorEntConserto3
            .dQuantSaiConserto3 = .dQuantSaiConserto3 + objSldDiaEstAlm.dQuantSaiConserto3
            .dValorSaiConserto3 = .dValorSaiConserto3 + objSldDiaEstAlm.dValorSaiConserto3
            .dQuantEntOutros3 = .dQuantEntOutros3 + objSldDiaEstAlm.dQuantEntOutros3
            .dValorEntOutros3 = .dValorEntOutros3 + objSldDiaEstAlm.dValorEntOutros3
            .dQuantSaiOutros3 = .dQuantSaiOutros3 + objSldDiaEstAlm.dQuantSaiOutros3
            .dValorSaiOutros3 = .dValorSaiOutros3 + objSldDiaEstAlm.dValorSaiOutros3
            .dQuantEntBenef3 = .dQuantEntBenef3 + objSldDiaEstAlm.dQuantEntBenef3
            .dValorEntBenef3 = .dValorEntBenef3 + objSldDiaEstAlm.dValorEntBenef3
            .dQuantSaiBenef3 = .dQuantSaiBenef3 + objSldDiaEstAlm.dQuantSaiBenef3
            .dValorSaiBenef3 = .dValorSaiBenef3 + objSldDiaEstAlm.dValorSaiBenef3
            .dQuantEntConsig = .dQuantEntConsig + objSldDiaEstAlm.dQuantEntConsig
            .dValorEntConsig = .dValorEntConsig + objSldDiaEstAlm.dValorEntConsig
            .dQuantSaiConsig = .dQuantSaiConsig + objSldDiaEstAlm.dQuantSaiConsig
            .dValorSaiConsig = .dValorSaiConsig + objSldDiaEstAlm.dValorSaiConsig
            .dQuantEntDemo = .dQuantEntDemo + objSldDiaEstAlm.dQuantEntDemo
            .dValorEntDemo = .dValorEntDemo + objSldDiaEstAlm.dValorEntDemo
            .dQuantSaiDemo = .dQuantSaiDemo + objSldDiaEstAlm.dQuantSaiDemo
            .dValorSaiDemo = .dValorSaiDemo + objSldDiaEstAlm.dValorSaiDemo
            .dQuantEntConserto = .dQuantEntConserto + objSldDiaEstAlm.dQuantEntConserto
            .dValorEntConserto = .dValorEntConserto + objSldDiaEstAlm.dValorEntConserto
            .dQuantSaiConserto = .dQuantSaiConserto + objSldDiaEstAlm.dQuantSaiConserto
            .dValorSaiConserto = .dValorSaiConserto + objSldDiaEstAlm.dValorSaiConserto
            .dQuantEntOutros = .dQuantEntOutros + objSldDiaEstAlm.dQuantEntOutros
            .dValorEntOutros = .dValorEntOutros + objSldDiaEstAlm.dValorEntOutros
            .dQuantSaiOutros = .dQuantSaiOutros + objSldDiaEstAlm.dQuantSaiOutros
            .dValorSaiOutros = .dValorSaiOutros + objSldDiaEstAlm.dValorSaiOutros
            .dQuantEntBenef = .dQuantEntBenef + objSldDiaEstAlm.dQuantEntBenef
            .dValorEntBenef = .dValorEntBenef + objSldDiaEstAlm.dValorEntBenef
            .dQuantSaiBenef = .dQuantSaiBenef + objSldDiaEstAlm.dQuantSaiBenef
            .dValorSaiBenef = .dValorSaiBenef + objSldDiaEstAlm.dValorSaiBenef

        End With
        
        Estoque_Acumula_ItemMovEst_SldDiaAlm2 = SUCESSO
        
        Exit Function

Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm2:

    Estoque_Acumula_ItemMovEst_SldDiaAlm2 = gErr
    
    Select Case gErr
        
        Case 94528
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173787)
    
    End Select
    
    Exit Function

End Function

Function Estoque_Acumula_SldDia(objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm, objSldDiaEstAcumulado As ClassSldDiaEst) As Long

On Error GoTo Erro_Estoque_Acumula_SldDia
    
    With objSldDiaEstAlmAcumulado
    
        .dQuantEntrada = .dQuantEntrada + objSldDiaEstAcumulado.dQuantEntrada
        .dQuantSaida = .dQuantSaida + objSldDiaEstAcumulado.dQuantSaida
        .dQuantCons = .dQuantCons + objSldDiaEstAcumulado.dQuantCons
        .dQuantVend = .dQuantVend + objSldDiaEstAcumulado.dQuantVend
        .dQuantVendConsig3 = .dQuantVendConsig3 + objSldDiaEstAcumulado.dQuantVendConsig3
        .dValorEntrada = .dValorEntrada + objSldDiaEstAcumulado.dValorEntrada
        .dValorSaida = .dValorSaida + objSldDiaEstAcumulado.dValorSaida
        .dQuantComp = .dQuantComp + objSldDiaEstAcumulado.dQuantComp
        .dValorComp = .dValorComp + objSldDiaEstAcumulado.dValorComp
        .dValorVend = .dValorVend + objSldDiaEstAcumulado.dValorVend
        .dValorCons = .dValorCons + objSldDiaEstAcumulado.dValorCons
        .dValorVendConsig3 = .dValorVendConsig3 + objSldDiaEstAcumulado.dValorVendConsig3
        .dQuantEntCusto = .dQuantEntCusto + objSldDiaEstAcumulado.dQuantEntCusto
        .dValorEntCusto = .dValorEntCusto + objSldDiaEstAcumulado.dValorEntCusto
        .dQuantSaiCusto = .dQuantSaiCusto + objSldDiaEstAcumulado.dQuantSaiCusto
        .dValorSaiCusto = .dValorSaiCusto + objSldDiaEstAcumulado.dValorSaiCusto
        .dQuantEntConsig3 = .dQuantEntConsig3 + objSldDiaEstAcumulado.dQuantEntConsig3
        .dValorEntConsig3 = .dValorEntConsig3 + objSldDiaEstAcumulado.dValorEntConsig3
        .dQuantSaiConsig3 = .dQuantSaiConsig3 + objSldDiaEstAcumulado.dQuantSaiConsig3
        .dValorSaiConsig3 = .dValorSaiConsig3 + objSldDiaEstAcumulado.dValorSaiConsig3
        .dQuantEntDemo3 = .dQuantEntDemo3 + objSldDiaEstAcumulado.dQuantEntDemo3
        .dValorEntDemo3 = .dValorEntDemo3 + objSldDiaEstAcumulado.dValorEntDemo3
        .dQuantSaiDemo3 = .dQuantSaiDemo3 + objSldDiaEstAcumulado.dQuantSaiDemo3
        .dValorSaiDemo3 = .dValorSaiDemo3 + objSldDiaEstAcumulado.dValorSaiDemo3
        .dQuantEntConserto3 = .dQuantEntConserto3 + objSldDiaEstAcumulado.dQuantEntConserto3
        .dValorEntConserto3 = .dValorEntConserto3 + objSldDiaEstAcumulado.dValorEntConserto3
        .dQuantSaiConserto3 = .dQuantSaiConserto3 + objSldDiaEstAcumulado.dQuantSaiConserto3
        .dValorSaiConserto3 = .dValorSaiConserto3 + objSldDiaEstAcumulado.dValorSaiConserto3
        .dQuantEntOutros3 = .dQuantEntOutros3 + objSldDiaEstAcumulado.dQuantEntOutros3
        .dValorEntOutros3 = .dValorEntOutros3 + objSldDiaEstAcumulado.dValorEntOutros3
        .dQuantSaiOutros3 = .dQuantSaiOutros3 + objSldDiaEstAcumulado.dQuantSaiOutros3
        .dValorSaiOutros3 = .dValorSaiOutros3 + objSldDiaEstAcumulado.dValorSaiOutros3
        .dQuantEntBenef3 = .dQuantEntBenef3 + objSldDiaEstAcumulado.dQuantEntBenef3
        .dValorEntBenef3 = .dValorEntBenef3 + objSldDiaEstAcumulado.dValorEntBenef3
        .dQuantSaiBenef3 = .dQuantSaiBenef3 + objSldDiaEstAcumulado.dQuantSaiBenef3
        .dValorSaiBenef3 = .dValorSaiBenef3 + objSldDiaEstAcumulado.dValorSaiBenef3
        .dQuantEntConsig = .dQuantEntConsig + objSldDiaEstAcumulado.dQuantEntConsig
        .dValorEntConsig = .dValorEntConsig + objSldDiaEstAcumulado.dValorEntConsig
        .dQuantSaiConsig = .dQuantSaiConsig + objSldDiaEstAcumulado.dQuantSaiConsig
        .dValorSaiConsig = .dValorSaiConsig + objSldDiaEstAcumulado.dValorSaiConsig
        .dQuantEntDemo = .dQuantEntDemo + objSldDiaEstAcumulado.dQuantEntDemo
        .dValorEntDemo = .dValorEntDemo + objSldDiaEstAcumulado.dValorEntDemo
        .dQuantSaiDemo = .dQuantSaiDemo + objSldDiaEstAcumulado.dQuantSaiDemo
        .dValorSaiDemo = .dValorSaiDemo + objSldDiaEstAcumulado.dValorSaiDemo
        .dQuantEntConserto = .dQuantEntConserto + objSldDiaEstAcumulado.dQuantEntConserto
        .dValorEntConserto = .dValorEntConserto + objSldDiaEstAcumulado.dValorEntConserto
        .dQuantSaiConserto = .dQuantSaiConserto + objSldDiaEstAcumulado.dQuantSaiConserto
        .dValorSaiConserto = .dValorSaiConserto + objSldDiaEstAcumulado.dValorSaiConserto
        .dQuantEntOutros = .dQuantEntOutros + objSldDiaEstAcumulado.dQuantEntOutros
        .dValorEntOutros = .dValorEntOutros + objSldDiaEstAcumulado.dValorEntOutros
        .dQuantSaiOutros = .dQuantSaiOutros + objSldDiaEstAcumulado.dQuantSaiOutros
        .dValorSaiOutros = .dValorSaiOutros + objSldDiaEstAcumulado.dValorSaiOutros
        .dQuantEntBenef = .dQuantEntBenef + objSldDiaEstAcumulado.dQuantEntBenef
        .dValorEntBenef = .dValorEntBenef + objSldDiaEstAcumulado.dValorEntBenef
        .dQuantSaiBenef = .dQuantSaiBenef + objSldDiaEstAcumulado.dQuantSaiBenef
        .dValorSaiBenef = .dValorSaiBenef + objSldDiaEstAcumulado.dValorSaiBenef
        .dQuantEntRecIndl = .dQuantEntRecIndl + objSldDiaEstAcumulado.dQuantEntRecIndl
        .dValorEntRecIndl = .dValorEntRecIndl + objSldDiaEstAcumulado.dValorEntRecIndl
        .dQuantSaiRecIndl = .dQuantSaiRecIndl + objSldDiaEstAcumulado.dQuantSaiRecIndl
        .dValorSaiRecIndl = .dValorSaiRecIndl + objSldDiaEstAcumulado.dValorSaiRecIndl
    
    End With
        
    Estoque_Acumula_SldDia = SUCESSO
    
    Exit Function
    
Erro_Estoque_Acumula_SldDia:

    Estoque_Acumula_SldDia = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173788)
            
    End Select
    
    Exit Function
    
End Function

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Function Estoque_Compara_SldMesEstAlm(sProduto As String, dtDataInicio As Date, dtDataFim As Date) As Long
'
'Dim iIndice As Integer
'Dim lErro As Long
'Dim lErro1 As Long
'Dim lComando As Long
'Dim sSelect As String
'Dim tSldDiaEstAlm As typeSldDiaEstAlm
'Dim objSldDiaEstAlm As New ClassSldDiaEstAlm
'Dim objSldMesEstAlm As ClassSldMesEstAlm
'Dim objSldMesEstAlm2 As ClassSldMesEstAlm2
'Dim iAlmoxarifado As Integer
'Dim objSldMesEstAlmAcumulado As ClassSldMesEstAlm
'
'Dim colPropertiesSldMesEstAlm As New Collection
'Dim colPropertiesSldMesEstAlm1 As New Collection
'Dim colPropertiesSldMesEstAlm2 As New Collection
'Dim objCampo As New AdmCampos
'
'On Error GoTo Erro_Estoque_Acumula_SldDiaEstAlm
'
'    'Abre o comando
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError xxx
'
'    'Lê em SldDiaEstAlmFilial os dados que serão comparados
'    lErro1 = SldDiaEstAlmFilial_Le1(alComando(0), sProduto, dtDataInicio, dtDataFim, iMes, iAno, objSldMesEstAlmAcumulado, objSldMesEstAlm1Acumulado, objSldMesEstAlm2Acumulado)
'    If lErro1 <> SUCESSO Then gError xxx
'
'    'Lê em SldMestEstAlm os dados que serão comparados
'    lErro = SldMesEstAlm_Le1(sProduto, iMes, iAno, iAlmoxarifado, objSldMesEstAlm)
'    If lErro <> SUCESSO Then gError xxx
'
'    'Lê em SldMestEstAlm1 os dados que serão comparados
'    lErro = SldMesEstAlm1_Le1(sProduto, iMes, iAno, iAlmoxarifado, objSldMesEstAlm)
'    If lErro <> SUCESSO Then gError xxx
'
'    'Lê em SldMestEstAlm2 os dados que serão comparados
'    lErro = SldMesEstAlm2_Le1(sProduto, iMes, iAno, iAlmoxarifado, objSldMesEstAlm2)
'    If lErro <> SUCESSO Then gError xxx
'
'    'Obtém uma coleção com as properties que compõem a tabela SldMesEstAlm
'    lErro = Obtem_Properties_ObjetoTabela("SldMesEstAlm", colPropertiesSldMesEstAlm)
'    If lErro <> SUCESSO Then gError xxx
'
'    'Obtém uma coleção com as properties que compõem a tabela SldMesEstAlm1
'    lErro = Obtem_Properties_ObjetoTabela("SldMesEstAlm1", colPropertiesSldMesEstAlm1)
'    If lErro <> SUCESSO Then gError 94545
'
'    'Obtém uma coleção com as properties que compõem a tabela SldMesEstAlm2
'    lErro = Obtem_Properties_ObjetoTabela("SldDiaEstAlm", colProperties)
'    If lErro <> SUCESSO Then gError 94545
'
'    Do While lErro1 = SUCESSO
'
'        For Each objCampo In colPropertiesSldMesEstAlm
'            If Abs(CallByName(objSldMesEstAlm, objCampo.sDescricao, VbGet) - CallByName(objSldMesEstAlmAcumulado, objCampo.sDescricao, VbGet)) > QTDE_ESTOQUE_DELTA Then gError xxx
'        Next
'
'        For Each objCampo In colPropertiesSldMesEstAlm1
'            If Abs(CallByName(objSldMesEstAlm1, objCampo.sDescricao, VbGet) - CallByName(objSldMesEstAlmAcumulado1, objCampo.sDescricao, VbGet)) > QTDE_ESTOQUE_DELTA Then gError xxx
'        Next
'
'        For Each objCampo In colPropertiesSldMesEstAlm2
'            If Abs(CallByName(objSldMesEstAlm2, objCampo.sDescricao, VbGet) - CallByName(objSldMesEstAlmAcumulado2, objCampo.sDescricao, VbGet)) > QTDE_ESTOQUE_DELTA Then gError xxx
'        Next
'
'        lErro1 = Comando_BuscarProximo(alComando(0))
'        If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError xxx
'
'    Loop
'
'
''    'Guarda na variável o Almoxarifado do último registro processado
''    'Essa variável será utilizada para verificar se foi alterado o almoxarifado
''    'Pois, quando for um novo almoxarifado, deve ser instanciado um novo obj
''    iAlmoxarifado = objSldDiaEstAlm.iAlmoxarifado
''
''    'Enquanto houverem registros a serem processados
''    Do While lErro = AD_SQL_SUCESSO
''
''        'Se o almoxarifado atual (objSldDiaEstAlm.iAlmoxarifado) for diferente do almoxarifado do último registro (iAlmoxarifado)
''        If iAlmoxarifado <> objSldDiaEstAlm.iAlmoxarifado Then
''
''            'Guarda no obj o almoxarifado que está sendo adicionado na coleção
''            objSldMesEstAlmAcumulado.iAlmoxarifado = iAlmoxarifado
''
''            'Guarda na coleção os totais para o almoxarifado que acabou de ser processado
''            colSldMesEstAlmAcumulado.Add objSldMesEstAlmAcumulado
''
''            'Instancia um novo objSldMesEstAlmAcumulado que será utilizado para processar o novo almoxarifado
''            Set objSldMesEstAlmAcumulado = New ClassSldMesEstAlm
''
''        End If
''
''        'Acumula os valores em objSldMesEstAlmAcumulado
''        With objSldMesEstAlmAcumulado
''
''            .dQuantComp(1) = .dQuantComp + tSldDiaEstAlm.dQuantComp
''            .dQuantCons(1) = .dQuantCons(1) + tSldDiaEstAlm.dQuantCons
''            .dQuantEnt(1) = .dQuantEnt(1) + tSldDiaEstAlm.dQuantEnt
''            .dQuantSai(1) = .dQuantSai(1) + tSldDiaEstAlm.dQuantSai
''            .dQuantVend(1) = .dQuantVend(1) + tSldDiaEstAlm.dQuantVend
''            .dQuantVendConsig3(1) = .dQuantVendConsig3(1) + tSldDiaEstAlm.dQuantVendConsig3
''            .dSaldoQuantCusto(1) = .dSaldoQuantCusto(1) + tSldDiaEstAlm.dSaldoQuantCusto
''            .dSaldoQuantRecIndl(1) = .dSaldoQuantRecIndl(1) + tSldDiaEstAlm.dSaldoQuantRecIndl
''            .dSaldoValorCusto(1) = .dSaldoValorCusto(1) + tSldDiaEstAlm.dSaldoValorCusto
''            .dSaldoValorRecIndl(1) = .dSaldoValorRecIndl(1) + tSldDiaEstAlm.dSaldoValorRecIndl
''            .dValorComp(1) = .dValorComp(1) + tSldDiaEstAlm.dValorComp
''            .dValorCons(1) = .dValorCons(1) + tSldDiaEstAlm.dValorCons
''            .dValorEnt(1) = .dValorEnt(1) + tSldDiaEstAlm.dValorEnt
''            .dValorSai(1) = .dValorSai(1) + tSldDiaEstAlm.dValorSai
''            .dValorVend(1) = .dValorVend(1) + tSldDiaEstAlm.dValorVend
''            .dValorVendConsig3(1) = .dValorVendConsig3(1) + tSldDiaEstAlm.dValorVendConsig3
''
''        End With
''
''        'Guarda na variável o Almoxarifado do último registro processado
''        'Essa variável será utilizada para verificar se foi alterado o almoxarifado
''        'Pois, quando for um novo almoxarifado, deve ser instanciado um novo obj
''        iAlmoxarifado = objSldDiaEstAlm.iAlmoxarifado
''
''        'Busca o próximo registro a ser processado
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94532
''
''        'Transfere para o obj os dados lidos do BD
''        Call Move_tSldDiaEstAlm_objSldDiaEstAlm
''
''    Loop
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Estoque_Acumula_SldDiaEstAlm = SUCESSO
'
'    Exit Function
'
'Erro_Estoque_Acumula_SldDiaEstAlm:
'
'    Estoque_Acumula_SldDiaEstAlm = gErr
'
'    Select Case gErr
'
'        Case 94529
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173789)
'
'    End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Private Sub Move_tSldDiaEstAlm_objSldDiaEstAlm(tSldDiaEstAlm As typeSldDiaEstAlm, objSldDiaEstAlm As ClassSldDiaEstAlm)
        
On Error GoTo Erro_Move_tSldDiaEstAlm_objSldDiaEstAlm

    With tSldDiaEstAlm
        
        objSldDiaEstAlm.sProduto = .sProduto
        objSldDiaEstAlm.dtData = .dtData
        objSldDiaEstAlm.iAlmoxarifado = .iAlmoxarifado
        objSldDiaEstAlm.dQuantEntrada = .dQuantEntrada
        objSldDiaEstAlm.dQuantSaida = .dQuantSaida
        objSldDiaEstAlm.dQuantCons = .dQuantCons
        objSldDiaEstAlm.dQuantVend = .dQuantVend
        objSldDiaEstAlm.dQuantVendConsig3 = .dQuantVendConsig3
        objSldDiaEstAlm.dValorEntrada = .dValorEntrada
        objSldDiaEstAlm.dValorSaida = .dValorSaida
        objSldDiaEstAlm.dQuantComp = .dQuantComp
        objSldDiaEstAlm.dValorComp = .dValorComp
        objSldDiaEstAlm.dValorVend = .dValorVend
        objSldDiaEstAlm.dValorCons = .dValorCons
        objSldDiaEstAlm.dQuantEntCusto = .dQuantEntCusto
        objSldDiaEstAlm.dValorEntCusto = .dValorEntCusto
        objSldDiaEstAlm.dQuantSaiCusto = .dQuantSaiCusto
        objSldDiaEstAlm.dValorSaiCusto = .dValorSaiCusto
        objSldDiaEstAlm.dValorVendConsig3 = .dValorVendConsig3
        objSldDiaEstAlm.dQuantEntConsig3 = .dQuantEntConsig3
        objSldDiaEstAlm.dValorEntConsig3 = .dValorEntConsig3
        objSldDiaEstAlm.dQuantSaiConsig3 = .dQuantSaiConsig3
        objSldDiaEstAlm.dValorSaiConsig3 = .dValorSaiConsig3
        objSldDiaEstAlm.dQuantEntDemo3 = .dQuantEntDemo3
        objSldDiaEstAlm.dValorEntDemo3 = .dValorEntDemo3
        objSldDiaEstAlm.dQuantSaiDemo3 = .dQuantSaiDemo3
        objSldDiaEstAlm.dValorSaiDemo3 = .dValorSaiDemo3
        objSldDiaEstAlm.dQuantEntConserto3 = .dQuantEntConserto3
        objSldDiaEstAlm.dValorEntConserto3 = .dValorEntConserto3
        objSldDiaEstAlm.dQuantSaiConserto3 = .dQuantSaiConserto3
        objSldDiaEstAlm.dValorSaiConserto3 = .dValorSaiConserto3
        objSldDiaEstAlm.dQuantEntOutros3 = .dQuantEntOutros3
        objSldDiaEstAlm.dValorEntOutros3 = .dValorEntOutros3
        objSldDiaEstAlm.dQuantSaiOutros3 = .dQuantSaiOutros3
        objSldDiaEstAlm.dValorSaiOutros3 = .dValorSaiOutros3
        objSldDiaEstAlm.dQuantEntBenef3 = .dQuantEntBenef3
        objSldDiaEstAlm.dValorEntBenef3 = .dValorEntBenef3
        objSldDiaEstAlm.dQuantSaiBenef3 = .dQuantSaiBenef3
        objSldDiaEstAlm.dValorSaiBenef3 = .dValorSaiBenef3
        objSldDiaEstAlm.dQuantEntConsig = .dQuantEntConsig
        objSldDiaEstAlm.dValorEntConsig = .dValorEntConsig
        objSldDiaEstAlm.dQuantSaiConsig = .dQuantSaiConsig
        objSldDiaEstAlm.dValorSaiConsig = .dValorSaiConsig
        objSldDiaEstAlm.dQuantEntDemo = .dQuantEntDemo
        objSldDiaEstAlm.dValorEntDemo = .dValorEntDemo
        objSldDiaEstAlm.dQuantSaiDemo = .dQuantSaiDemo
        objSldDiaEstAlm.dValorSaiDemo = .dValorSaiDemo
        objSldDiaEstAlm.dQuantEntConserto = .dQuantEntConserto
        objSldDiaEstAlm.dValorEntConserto = .dValorEntConserto
        objSldDiaEstAlm.dQuantSaiConserto = .dQuantSaiConserto
        objSldDiaEstAlm.dValorSaiConserto = .dValorSaiConserto
        objSldDiaEstAlm.dQuantEntOutros = .dQuantEntOutros
        objSldDiaEstAlm.dValorEntOutros = .dValorEntOutros
        objSldDiaEstAlm.dQuantSaiOutros = .dQuantSaiOutros
        objSldDiaEstAlm.dValorSaiOutros = .dValorSaiOutros
        objSldDiaEstAlm.dQuantEntBenef = .dQuantEntBenef
        objSldDiaEstAlm.dValorEntBenef = .dValorEntBenef
        objSldDiaEstAlm.dQuantSaiBenef = .dQuantSaiBenef
        objSldDiaEstAlm.dValorSaiBenef = .dValorSaiBenef
        objSldDiaEstAlm.dQuantEntRecIndl = .dQuantEntRecIndl
        objSldDiaEstAlm.dValorEntRecIndl = .dValorEntRecIndl
        objSldDiaEstAlm.dQuantSaiRecIndl = .dQuantSaiRecIndl
        objSldDiaEstAlm.dValorSaiRecIndl = .dValorSaiRecIndl
    
    End With
    
    Exit Sub

Erro_Move_tSldDiaEstAlm_objSldDiaEstAlm:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173790)
    
    End Select
    
    Exit Sub

End Sub

'Criada por: Luiz G.F.Nogueira
'em: 19/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function Estoque_Acumula_SldMesEstAlm(colSldMesEstAlmAcumulado As Collection, objSldMesEstAcumulado As ClassSldMesEst) As Long
'Recebe uma coleção de obj's SldMesEstAlm contendo os valores para cada almoxarifado e
'acumula esses valores em um único objSldMesEst
'A coleção deve conter um obj para cada almoxarifado sendo que todos os obj's estão relacionados a um único produto

Dim objSldMesEstAlm As ClassSldMesEstAlm

On Error GoTo Erro_Estoque_Acumula_SldMesEstAlm

    'Para cada obj da coleção
    For Each objSldMesEstAlm In colSldMesEstAlmAcumulado
        
        'Acumula os valores
        With objSldMesEstAcumulado
            
            .dQuantComp(1) = .dQuantComp(1) + objSldMesEstAlm.dQuantComp(1)
            .dQuantCons(1) = .dQuantCons(1) + objSldMesEstAlm.dQuantCons(1)
            .dQuantEnt(1) = .dQuantEnt(1) + objSldMesEstAlm.dQuantEnt(1)
            .dQuantSai(1) = .dQuantSai(1) + objSldMesEstAlm.dQuantSai(1)
            .dQuantVend(1) = .dQuantVend(1) + objSldMesEstAlm.dQuantVend(1)
            .dQuantVendConsig3(1) = .dQuantVendConsig3(1) + objSldMesEstAlm.dQuantVendConsig3(1)
            .dSaldoQuantCusto(1) = .dSaldoQuantCusto(1) + objSldMesEstAlm.dSaldoQuantCusto(1)
            .dSaldoQuantRecIndl(1) = .dSaldoQuantRecIndl(1) + objSldMesEstAlm.dSaldoQuantRecIndl(1)
            .dSaldoValorCusto(1) = .dSaldoValorCusto(1) + objSldMesEstAlm.dSaldoValorCusto(1)
            .dSaldoValorRecIndl(1) = .dSaldoValorRecIndl(1) + objSldMesEstAlm.dSaldoValorRecIndl(1)
            .dValorComp(1) = .dValorComp(1) + objSldMesEstAlm.dValorComp(1)
            .dValorCons(1) = .dValorCons(1) + objSldMesEstAlm.dValorCons(1)
            .dValorEnt(1) = .dValorEnt(1) + objSldMesEstAlm.dValorEnt(1)
            .dValorSai(1) = .dValorSai(1) + objSldMesEstAlm.dValorSai(1)
            .dValorVend(1) = .dValorVend(1) + objSldMesEstAlm.dValorVend(1)
            .dValorVendConsig3(1) = .dValorVendConsig3(1) + objSldMesEstAlm.dValorVendConsig3(1)
        
        End With
    
    Next
            
    Estoque_Acumula_SldMesEstAlm = SUCESSO
    
    Exit Function

Erro_Estoque_Acumula_SldMesEstAlm:

    Estoque_Acumula_SldMesEstAlm = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173791)
    
    End Select
    
    Exit Function
    
End Function

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Private Function Estoque_Acumula_ItemMovEst_SldDiaAlm1(objItemMovEstoque As ClassItemMovEstoque, objTipoMovEstoque As ClassTipoMovEst, objEstoqueProduto As ClassEstoqueProduto) As Long
''Chama a função Estoque_AtualizaItemMov2, que seta as variáveis com os valores e os sinais
''Depois acumula os valores devolvido por Estoque_AtualizaItemMov2
'
'Dim lErro As Long
'
'On Error GoTo Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm1
'
'        'Guarda em objEstoqueProduto os valores com os sinais que deverão ser acumulados
'        lErro = CF("Estoque_AtualizaItemMov2", objItemMovEstoque, objTipoMovEstoque, objEstoqueProduto)
'        If lErro <> SUCESSO Then gError 94526
'
'        'Acumula os valores no obj
'        With objEstoqueProdutoAcumulado
'
'            .dQuantBenef = .dQuantBenef + objEstoqueProduto.dQuantBenef
'            .dQuantBenef3 = .dQuantBenef3 + objEstoqueProduto.dQuantBenef3
'            .dQuantConserto = .dQuantConserto + objEstoqueProduto.dQuantConserto
'            .dQuantConserto3 = .dQuantConserto3 + objEstoqueProduto.dQuantConserto3
'            .dQuantConsig = .dQuantConsig + objEstoqueProduto.dQuantConsig
'            .dQuantConsig3 = .dQuantConsig3 + objEstoqueProduto.dQuantConsig3
'            .dQuantDemo = .dQuantDemo + objEstoqueProduto.dQuantDemo
'            .dQuantDemo3 = .dQuantDemo3 + objEstoqueProduto.dQuantDemo3
'            .dQuantOutras = .dQuantOutras + objEstoqueProduto.dQuantOutras
'            .dQuantOutras3 = .dQuantOutras3 + objEstoqueProduto.dQuantOutras3
'
'        End With
'
'        Estoque_Acumula_ItemMovEst_SldDiaAlm1 = SUCESSO
'
'        Exit Function
'
'Erro_Estoque_Acumula_ItemMovEst_SldDiaAlm1:
'
'    Estoque_Acumula_ItemMovEst_SldDiaAlm1 = gErr
'
'    Select Case gErr
'
'        Case 94526
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173792)
'
'    End Select
'
'    Exit Function
'
'End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???

Function Estoque_Compara_SldDiaEstAlm(objSldDiaEstAlmAcumulado As ClassSldDiaEstAlm) As Long
'Lê os valores em SldDiaEstAlm para o produto, data, almoxarifado passados no objSldDiaEstAlmAcumulado
'Compara os valores lidos com os valores contidos em objSldDiaEstAlmAcumulado

Dim lErro As Long
Dim objSldDiaEstAlm As New ClassSldDiaEstAlm
Dim objCampo As AdmCampos
Dim colProperties As New Collection
Dim dvalor1 As Double
Dim dvalor2 As Double

On Error GoTo Erro_Estoque_Compara_SldDiaEstAlm

    'Faz a leitura dos dados na tabela SldDiaEstAlm
    lErro = SldDiaEstAlm_Le(objSldDiaEstAlm)
    If lErro <> SUCESSO Then gError 94546
    
    'Obtém uma coleção com as properties que compõem a tabela SldDiaEstAlm
    lErro = Obtem_Properties_ObjetoTabela("SldDiaEstAlm", colProperties)
    If lErro <> SUCESSO Then gError 94545
    
    'Para cada campo (property) da coleção
    For Each objCampo In colProperties
    
        'Se o campo deve ser comparado
        If objCampo.iTestaIntegridade = CAMPO_TESTA_INTEGRIDADE Then
            
            'Se o valor absoluto da diferença entre o campo no obj acumulado (objSldDiaEstAlmAcumulado) e no obj lido do bd(objSldDiaEstAlm) é maior que um valor delta => erro
            'ex.: supondo que objCampo.sNome seja dQuantEntrada, a linha de código abaixo seria o mesmo que
            'If Abs(objSldDiaEstAlm.dQuantEntrada - objSldDiaEstAlmAcumulado.dQuantEntrada) > QTDE_ESTOQUE_DELTA Then gError xxx
            If Abs(CallByName(objSldDiaEstAlm, objCampo.sNome, VbGet) - CallByName(objSldDiaEstAlmAcumulado, objCampo.sNome, VbGet)) > QTDE_ESTOQUE_DELTA Then gError 94547

        End If
    
    Next
        
'        If Abs(objSldDiaEstAlm.dQuantEntrada - .dQuantEntrada) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaida - .dQuantSaida) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantCons - .dQuantCons) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantVend - .dQuantVend) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantVendConsig3 - .dQuantVendConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntrada - .dValorEntrada) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaida - .dValorSaida) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantComp - .dQuantComp) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorComp - .dValorComp) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorVend - .dValorVend) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorCons - .dValorCons) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntCusto - .dQuantEntCusto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntCusto - .dValorEntCusto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiCusto - .dQuantSaiCusto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiCusto - .dValorSaiCusto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorVendConsig3 - .dValorVendConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntConsig3 - .dQuantEntConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntConsig3 - .dValorEntConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiConsig3 - .dQuantSaiConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiConsig3 - .dValorSaiConsig3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntDemo3 - .dQuantEntDemo3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntDemo3 - .dValorEntDemo3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiDemo3 - .dQuantSaiDemo3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiDemo3 - .dValorSaiDemo3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntConserto3 - .dQuantEntConserto3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntConserto3 - .dValorEntConserto3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiConserto3 - .dQuantSaiConserto3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiConserto3 - .dValorSaiConserto3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntOutros3 - .dQuantEntOutros3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntOutros3 - .dValorEntOutros3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiOutros3 - .dQuantSaiOutros3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiOutros3 - .dValorSaiOutros3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntBenef3 - .dQuantEntBenef3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntBenef3 - .dValorEntBenef3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiBenef3 - .dQuantSaiBenef3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiBenef3 - .dValorSaiBenef3) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntConsig - .dQuantEntConsig) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntConsig - .dValorEntConsig) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiConsig - .dQuantSaiConsig) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiConsig - .dValorSaiConsig) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntDemo - .dQuantEntDemo) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntDemo - .dValorEntDemo) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiDemo - .dQuantSaiDemo) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiDemo - .dValorSaiDemo) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntConserto - .dQuantEntConserto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntConserto - .dValorEntConserto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiConserto - .dQuantSaiConserto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiConserto - .dValorSaiConserto) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntOutros - .dQuantEntOutros) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntOutros - .dValorEntOutros) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiOutros - .dQuantSaiOutros) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiOutros - .dValorSaiOutros) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntBenef - .dQuantEntBenef) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntBenef - .dValorEntBenef) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiBenef - .dQuantSaiBenef) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiBenef - .dValorSaiBenef) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantEntRecIndl - .dQuantEntRecIndl) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorEntRecIndl - .dValorEntRecIndl) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dQuantSaiRecIndl - .dQuantSaiRecIndl) > QTDE_ESTOQUE_DELTA Then gError xxx
'        If Abs(objSldDiaEstAlm.dValorSaiRecIndl - .dValorSaiRecIndl) > QTDE_ESTOQUE_DELTA Then gError xxx
    
    Estoque_Compara_SldDiaEstAlm = SUCESSO
    
    Exit Function

Erro_Estoque_Compara_SldDiaEstAlm:

    Estoque_Compara_SldDiaEstAlm = gErr
    
    Select Case gErr
    
        Case 94545, 94546
        
        Case 94547
            
            dvalor1 = CallByName(objSldDiaEstAlm, objCampo.sNome, VbGet)
            dvalor2 = CallByName(objSldDiaEstAlmAcumulado, objCampo.sNome, VbGet)
            Call Rotina_Erro(vbOKOnly, "ERRO_SLDDIAESTALM_CAMPO_INCONSISTENTE", gErr, objCampo.sNome, dvalor1, dvalor2, objSldDiaEstAlm.sProduto, objSldDiaEstAlm.dtData, objSldDiaEstAlm.iAlmoxarifado)
            'O valor do campo %s (%s) não está consistente com o valor acumulado (%s) na tabela MovimentoEstoque. Produto: %s, Data: %s, Almoxarifado: %s.
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173793)
            
    End Select
    
    Exit Function

End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function SldDiaEstAlm_Le(objSldDiaEstAlm As ClassSldDiaEstAlm) As Long
'Lê em SldDiaEstAlm os valores para o Produto, Data, Almoxarifado passados em objSldDiaEstAlm
    
Dim lErro As Long
Dim lComando As Long
Dim tSldDiaEstAlm As typeSldDiaEstAlm
Dim sSelect As String

On Error GoTo Erro_SldDiaEstAlm_Le
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94536
    
    'Monta o select que será feito em cima da tabela SldDiaEstAlm
    sSelect = "SELECT Produto, Data, Almoxarifado, QuantEntrada , QuantSaida, QuantCons, QuantVend, QuantVendConsig3, ValorEntrada, ValorSaida, QuantComp, ValorComp, ValorVend, ValorCons, QuantEntCusto, ValorEntCusto, QuantSaiCusto, ValorSaiCusto, ValorVendConsig3, QuantEntConsig3, ValorEntConsig3, QuantSaiConsig3, ValorSaiConsig3 , QuantEntDemo3, ValorEntDemo3, QuantSaiDemo3, ValorSaiDemo3, QuantEntConserto3, ValorEntConserto3, QuantSaiConserto3, ValorSaiConserto3, QuantEntOutros3, ValorEntOutros3, QuantSaiOutros3, ValorSaiOutros3, QuantEntBenef3, ValorEntBenef3, QuantSaiBenef3, ValorSaiBenef3, QuantEntConsig, ValorEntConsig, QuantSaiConsig , ValorSaiConsig, QuantEntDemo, ValorEntDemo, QuantSaiDemo, ValorSaiDemo, QuantEntConserto, ValorEntConserto, QuantSaiConserto, ValorSaiConserto, QuantEntOutros, ValorEntOutros, QuantSaiOutros, ValorSaiOutros, QuantEntBenef, ValorEntBenef, QuantSaiBenef, ValorSaiBenef, QuantEntRecIndl, ValorEntRecIndl , QuantSaiRecIndl, ValorSaiRecIndl FROM SldDiaEstAlm"
    sSelect = sSelect & "WHERE Produto = ? AND Data = ? AND Almoxarifado = ?"
    
    With tSldDiaEstAlm
    
        'Faz a leitura no BD
        lErro = Comando_Executar(lComando, sSelect, .sProduto, .dtData, .iAlmoxarifado, .dQuantEntrada, .dQuantSaida, .dQuantCons, .dQuantVend, .dQuantVendConsig3, .dValorEntrada, .dValorSaida, .dQuantComp, .dValorComp, .dValorVend, .dValorCons, .dQuantEntCusto, .dValorEntCusto, .dQuantSaiCusto, .dValorSaiCusto, .dValorVendConsig3, .dQuantEntConsig3, .dValorEntConsig3, .dQuantSaiConsig3, .dValorSaiConsig3, .dQuantEntDemo3, .dValorEntDemo3, .dQuantSaiDemo3, .dValorSaiDemo3, .dQuantEntConserto3, .dValorEntConserto3, .dQuantSaiConserto3, .dValorSaiConserto3, .dQuantEntOutros3, .dValorEntOutros3, .dQuantSaiOutros3, .dValorSaiOutros3, _
        .dQuantEntBenef3, .dValorEntBenef3, .dQuantSaiBenef3, .dValorSaiBenef3, .dQuantEntConsig, .dValorEntConsig, .dQuantSaiConsig, .dValorSaiConsig, .dQuantEntDemo, .dValorEntDemo, .dQuantSaiDemo, .dValorSaiDemo, .dQuantEntConserto, .dValorEntConserto, .dQuantSaiConserto, .dValorSaiConserto, .dQuantEntOutros, .dValorEntOutros, .dQuantSaiOutros, .dValorSaiOutros, .dQuantEntBenef, .dValorEntBenef, .dQuantSaiBenef, .dValorSaiBenef, .dQuantEntRecIndl, _
        .dValorEntRecIndl, .dQuantSaiRecIndl, .dValorSaiRecIndl, objSldDiaEstAlm.sProduto, objSldDiaEstAlm.dtData, objSldDiaEstAlm.iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 94537
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94538
        
        'Se não encontrou => erro (a ser tratado na rotina chamadora)
        If lErro = AD_SQL_SEM_DADOS Then gError 94539
        
    End With
        
    'Transfere para o obj os dados lidos do BD
    Call Move_tSldDiaEstAlm_objSldDiaEstAlm(tSldDiaEstAlm, objSldDiaEstAlm)

    SldDiaEstAlm_Le = SUCESSO
    
    Exit Function
    
Erro_SldDiaEstAlm_Le:

    SldDiaEstAlm_Le = gErr
    
    Select Case gErr
    
        Case 94536
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94537, 94538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM", gErr, objSldDiaEstAlm.iAlmoxarifado, objSldDiaEstAlm.sProduto, objSldDiaEstAlm.dtData)
            
        Case 94539
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173794)
            
    End Select
    
    Exit Function

End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function Campos_Le_Arquivo(sNomeTabela As String, colCampos As Collection) As Long
'Lê os campos para o arquivo passado em sNomeTabela e os devolve em uma coleção

Dim lComando As Long
Dim lErro As Long
Dim tCampos As typeCampos
Dim objCampos As AdmCampos
    
On Error GoTo Erro_Campos_Le_Arquivo

    'Inicializa as strings que serão lidas do BD
    tCampos.sDescricao = String(STRING_DESCRICAO_CAMPO, 0)
    tCampos.sFormatacao = String(STRING_FORMATACAO_CAMPO, 0)
    tCampos.sNome = String(STRING_NOME_CAMPO, 0)
    tCampos.sNomeArq = String(STRING_NOME_TABELA, 0)
    tCampos.sTituloEntradaDados = String(STRING_TITULO_ENTRADA_DADOS_CAMPO, 0)
    tCampos.sTituloGrid = String(STRING_TITULO_GRID_CAMPO, 0)
    tCampos.sValDefault = String(STRING_VALOR_DEFAULT_CAMPO, 0)
    tCampos.sValidacao = String(STRING_VALIDACAO_CAMPO, 0)

    'Abre o comando
    lComando = Comando_AbrirExt(GL_lConexaoDic)
    If lComando = 0 Then gError 94539
    
    'Lê no BD todos os campos de sNomeTabela
    lErro = Comando_Executar(lComando, "SELECT NomeArq, Nome, Descricao, Obrigatorio, Imexivel, Ativo, ValDefault, Validacao, Formatacao, Tipo, Tamanho, Precisao, Decimais, TamExibicao, TituloEntradaDados, TituloGrid, Subtipo, Alinhamento, TestaIntegridade FROM Campos WHERE NomeArq = ?", tCampos.sNomeArq, tCampos.sNome, tCampos.sDescricao, tCampos.iObrigatorio, tCampos.iImexivel, tCampos.iAtivo, tCampos.sValDefault, tCampos.sValidacao, tCampos.sFormatacao, tCampos.iTipo, tCampos.iTamanho, tCampos.iPrecisao, tCampos.iDecimais, tCampos.iTamExibicao, tCampos.sTituloEntradaDados, tCampos.sTituloGrid, tCampos.iSubTipo, tCampos.iAlinhamento, tCampos.iTestaIntegridade, sNomeTabela)
    If lErro <> AD_SQL_SUCESSO Then gError 94543
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94541

    'Se não encontrou ==>> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 94542

    'Para cada campo encontrado
    Do While lErro = AD_SQL_SUCESSO
    
        'Instancia um novo obj
        Set objCampos = New AdmCampos
        
        'Transfere os dados lido para o obj
        objCampos.iAtivo = tCampos.iAtivo
        objCampos.iDecimais = tCampos.iDecimais
        objCampos.iImexivel = tCampos.iImexivel
        objCampos.iObrigatorio = tCampos.iObrigatorio
        objCampos.iPrecisao = tCampos.iPrecisao
        objCampos.iTamanho = tCampos.iTamanho
        objCampos.iTamExibicao = tCampos.iTamExibicao
        objCampos.iTipo = tCampos.iTipo
        objCampos.sDescricao = tCampos.sDescricao
        objCampos.sFormatacao = tCampos.sFormatacao
        objCampos.sTituloEntradaDados = tCampos.sTituloEntradaDados
        objCampos.sTituloGrid = tCampos.sTituloGrid
        objCampos.sValDefault = tCampos.sValDefault
        objCampos.sValidacao = tCampos.sValidacao
        objCampos.iSubTipo = tCampos.iSubTipo
        objCampos.iAlinhamento = tCampos.iAlinhamento
        objCampos.iTestaIntegridade = tCampos.iTestaIntegridade
        
        'Guarda o obj na coleção
        colCampos.Add objCampos
    
    Loop

    Call Comando_Fechar(lComando)
    
    Campos_Le_Arquivo = SUCESSO
    
    Exit Function
    
Erro_Campos_Le_Arquivo:

    Campos_Le_Arquivo = gErr

    Select Case gErr
    
        Case 94539
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 94541, 94543
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CAMPOS", gErr)
        
        Case 94542
            'SEM DADOS. Erro a ser tratado na rotina chamadora!
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173795)
        
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function Obtem_Properties_ObjetoTabela(sNomeTabela As String, colProperties As Collection)
'Lê na tabela Campos os nomes dos campos que compõem a tabela passada em sNomeTabela
'Adiciona ao nome do campo a letra que identifica o tipo da variável, deixando-o semelhante a uma propertie (ex.: Produto -> sProduto)
'Devolve uma coleção de objCampos, onde cada objCampos.sNome contém o nome de uma propertie o objeto relacionado à tabela

Dim lErro As Long
Dim objCampos As AdmCampos
Dim colCampos As Collection

On Error GoTo Erro_Obtem_Properties_ObjetoTabela
    
    'Lê na tabela campos os nomes dos campos que compõem sNomeTabela
    lErro = Campos_Le_Arquivo(sNomeTabela, colCampos)
    If lErro <> SUCESSO Then gError 94544
    
    'Para cada campo da tabela
    For Each objCampos In colCampos
    
        'Verifica o tipo do campo
        Select Case objCampos.iTipo
        
            'Se for smallint(integer)
            Case ADM_TIPO_SMALLINT
                
                'Adiciona ao nome do campo o caracter "i"
                objCampos.sNome = ID_VARIAVEL_INTEGER & objCampos.sNome
                
            'Se for integer(long)
            Case ADM_TIPO_INTEGER
                
                'Adiciona ao nome do campo o caracter "l"
                objCampos.sNome = ID_VARIAVEL_LONG & objCampos.sNome
            
            'Se for double
            Case ADM_TIPO_DOUBLE
                
                'Adiciona ao nome do campo o caracter "d"
                objCampos.sNome = ID_VARIAVEL_DOUBLE & objCampos.sNome
            
            'Se for varchar(string)
            Case ADM_TIPO_VARCHAR
                
                'Adiciona ao nome do campo o caracter "s"
                objCampos.sNome = ID_VARIAVEL_STRING & objCampos.sNome
            
            'Se for date
            Case ADM_TIPO_DATE
                
                'Adiciona ao nome do campo o caracter "dt"
                objCampos.sNome = ID_VARIAVEL_DATE & objCampos.sNome
            
            'Se não for um dos tipos acima ==>> erro
            Case Else
                gError 94545
                
        End Select
        
        'Adiciona o obj à coleção de campos
        colProperties.Add objCampos
    
    Next
    
    Obtem_Properties_ObjetoTabela = SUCESSO
    
    Exit Function
    
Erro_Obtem_Properties_ObjetoTabela:

    Obtem_Properties_ObjetoTabela = gErr
    
    Select Case gErr
    
        Case 94544
        
        Case 94545
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CAMPO_INVALIDO", Err, objCampos.iTipo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173796)
            
    End Select
    
End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function SldDiaEstAlmFilial_Le1(lComando As Long, sProduto As String, dtDataInicio As Date, dtDataFim As Date, iMes As Integer, iAno As Integer, objSldMesEstAlmAcumulado As ClassSldMesEstAlm, objSldMesEstAlm1Acumulado As ClassSldMesEstAlm1, objSldMesEstAlm2Acumulado As ClassSldMesEstAlm2) As Long
'Lê em SldDiaEstAlmFilial os campos envolvendo quantidades e valores, no formato das tabelas SldMesEstAlm, SldMesEstAlm1 e SldMesEstAlm2,
'ou seja, os campos que nessas tabelas são representados apenas pelo saldo, são lidos aqui da mesma forma.
'Ex.: Em SldDiaEstAlm existem os campos QuantEntConsig3 e QuantSaiConsig3,
'já em SldMesEstAlm existe apenas o campos SaldoQuantEntConsig3,
'portanto, nesse caso, a leitura é feita pegando apenas a diferença entre os campos de SldDiaEstAlm
'Filtro: sProduto, dtDataInicio e dtDataFim
'A função devolve 3 obj's, pois os campos de SldDiaEstAlm possuem correspondentes em SldMesEstAlm, SldMesEstAlm1 e SldMesEstAlm2
'****** TEM QUE PASSAR COMO PARAMETRO A FILIAL EM VEZ DE USAR GIFILIALEMPRESA POIS O REPROCESSAMENTO ESTA TRABALHANDO COM TODAS AS FILIAIS


Dim lErro As Long
Dim sSelect As String
Dim tSldMesEstAlm As typeSldMesEstAlm
Dim tSldMesEstAlm1 As typeSldMesEstAlm1
Dim tSldMesEstAlm2 As typeSldMesEstAlm2

On Error GoTo Erro_SldDiaEstAlmFilial_Le1

    'Monta o select que será feito em cima da tabela SldDiaEstAlm
    sSelect = "SELECT MONTH(Data), YEAR(Data), SUM(QuantEntrada) , SUM(QuantSaida), SUM(QuantCons), SUM(QuantVend), SUM(QuantVendConsig3), SUM(ValorEntrada), SUM(ValorSaida), SUM(QuantComp), SUM(ValorComp), SUM(ValorVend), SUM(ValorCons), SUM(ValorVendConsig3), SUM(QuantEntConsig3 - QuantSaidConsig3), SUM(ValorEntConsig3 - ValorSaiConsig3), SUM(QuantEntDemo3 - QuantSaiDemo3), SUM(ValorEntDemo3 - ValorSaiDemo3), SUM(QuantEntConserto3 - QuantSaiConserto3), SUM(ValorEntConserto3 - ValorSaiConserto3), SUM(QuantEntOutros3 - QuantSaiOutros3), SUM(ValorEntOutros3 - ValorSaiOutros3), SUM(QuantEntBenef3 - QuantSaiBenef3), SUM(ValorEntBenef3 - ValorSaiBenef3), SUM(QuantEntConsig - QuantSaiConsig), SUM(ValorEntConsig - ValorSaiConsig), SUM(QuantEntDemo - QuantSaiDemo), SUM(ValorEntDemo - ValorSaiDemo), SUM(QuantEntConserto - QuantSaiConserto)," & _
    " SUM(ValorEntConserto - ValorSaidConserto), SUM(QuantEntOutros - QuantSaiOutros), SUM(ValorEntOutros - ValorSaiOutros), SUM(QuantEntBenef - QuantSaiBenef), SUM(ValorEntBenef - ValorSaiBenef), SUM(QuantEntRecIndl - QuantSaiRecIndl), SUM(ValorEntRecIndl - ValorSaiRecIndl), SUM(QuantEntCusto - QuantSaiCusto), SUM(ValorEntCusto - ValorSaiCusto) FROM SldDiaEstAlmFilial WHERE Produto = ? AND Data >= ? AND Data <= ? AND FilialEmpresa = ? GROUP BY Almoxarifado, MONTH(Data), YEAR(Data)"

    'Faz a leitura no BD
    lErro = Comando_Executar(lComando, sSelect, iMes, iAno, tSldMesEstAlm.adQuantEnt(1), tSldMesEstAlm.adQuantSai(1), tSldMesEstAlm.adQuantCons(1), tSldMesEstAlm.adQuantVend(1), tSldMesEstAlm.adQuantVendConsig3(1), tSldMesEstAlm.adValorEnt(1), tSldMesEstAlm.adValorSai(1), tSldMesEstAlm.adQuantComp(1), tSldMesEstAlm.adValorComp(1), tSldMesEstAlm.adValorVend(1), tSldMesEstAlm.adValorCons(1), tSldMesEstAlm.adValorVendConsig3(1), tSldMesEstAlm1.adSaldoQuantConsig3(1), tSldMesEstAlm1.adSaldoValorConsig3(1), tSldMesEstAlm1.adSaldoQuantDemo3(1), tSldMesEstAlm1.adSaldoValorDemo3(1), tSldMesEstAlm1.adSaldoQuantConserto3(1), tSldMesEstAlm1.adSaldoValorConserto3(1), tSldMesEstAlm1.adSaldoQuantOutros3(1), tSldMesEstAlm1.adSaldoValorOutros3(1), _
    tSldMesEstAlm1.adSaldoQuantBenef3(1), tSldMesEstAlm1.adSaldoValorBenef3(1), tSldMesEstAlm2.adSaldoQuantConsig(1), tSldMesEstAlm2.adSaldoValorConsig(1), tSldMesEstAlm2.adSaldoQuantDemo(1), tSldMesEstAlm2.adSaldoValorDemo(1), tSldMesEstAlm2.adSaldoQuantConserto(1), tSldMesEstAlm2.adSaldoValorConserto(1), tSldMesEstAlm2.adSaldoQuantOutros(1), tSldMesEstAlm2.adSaldoValorOutros(1), tSldMesEstAlm2.adSaldoQuantBenef(1), tSldMesEstAlm2.adSaldoValorBenef(1), tSldMesEstAlm.adSaldoQuantRecIndl(1), tSldMesEstAlm.adSaldoValorRecIndl(1), tSldMesEstAlm.adSaldoQuantCusto(1), tSldMesEstAlm.adSaldoValorCusto(1), sProduto, dtDataInicio, dtDataFim, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 94530
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94531
    
    'Se não encontrou => erro (a ser tratado na rotina chamadora)
    If lErro = AD_SQL_SEM_DADOS Then gError 94535
        
    'Transfere para o objSldMesEstAlmAcumulado os dados lidos do BD
    With objSldMesEstAlmAcumulado
        
        .dQuantEnt(iMes) = tSldMesEstAlm.adQuantEnt(1)
        .dQuantSai(iMes) = tSldMesEstAlm.adQuantSai(1)
        .dQuantCons(iMes) = tSldMesEstAlm.adQuantCons(1)
        .dQuantVend(iMes) = tSldMesEstAlm.adQuantVend(1)
        .dQuantVendConsig3(iMes) = tSldMesEstAlm.adQuantVendConsig3(1)
        .dValorEnt(iMes) = tSldMesEstAlm.adValorEnt(1)
        .dValorSai(iMes) = tSldMesEstAlm.adValorSai(1)
        .dQuantComp(iMes) = tSldMesEstAlm.adQuantComp(1)
        .dValorComp(iMes) = tSldMesEstAlm.adValorComp(1)
        .dValorVend(iMes) = tSldMesEstAlm.adValorVend(1)
        .dValorCons(iMes) = tSldMesEstAlm.adValorCons(1)
        .dValorVendConsig3(iMes) = tSldMesEstAlm.adValorVendConsig3(1)
        .dSaldoQuantRecIndl(iMes) = tSldMesEstAlm.adSaldoQuantRecIndl(1)
        .dSaldoValorRecIndl(iMes) = tSldMesEstAlm.adSaldoValorRecIndl(1)
        .dSaldoQuantCusto(iMes) = tSldMesEstAlm.adSaldoQuantCusto(1)
        .dSaldoValorCusto(iMes) = tSldMesEstAlm.adSaldoValorCusto(1)
        
    End With
        
    With objSldMesEstAlm1Acumulado
    
        'Transfere os dados do type para o obj
        .dSaldoQuantConsig3(iMes) = tSldMesEstAlm1.adSaldoQuantConsig3(iMes)
        .dSaldoValorConsig3(iMes) = tSldMesEstAlm1.adSaldoValorConsig3(iMes)
        .dSaldoQuantDemo3(iMes) = tSldMesEstAlm1.adSaldoQuantDemo3(iMes)
        .dSaldoValorDemo3(iMes) = tSldMesEstAlm1.adSaldoValorDemo3(iMes)
        .dSaldoQuantConserto3(iMes) = tSldMesEstAlm1.adSaldoQuantConserto3(iMes)
        .dSaldoValorConserto3(iMes) = tSldMesEstAlm1.adSaldoValorConserto3(iMes)
        .dSaldoQuantOutros3(iMes) = tSldMesEstAlm1.adSaldoQuantOutros3(iMes)
        .dSaldoValorOutros3(iMes) = tSldMesEstAlm1.adSaldoValorOutros3(iMes)
        .dSaldoQuantBenef3(iMes) = tSldMesEstAlm1.adSaldoQuantBenef3(iMes)
        .dSaldoValorBenef3(iMes) = tSldMesEstAlm1.adSaldoValorBenef3(iMes)
    
    End With
    
    With objSldMesEstAlm2Acumulado
    
        'Transfere os dados do type para o obj
        .dSaldoQuantConsig(iMes) = tSldMesEstAlm2.adSaldoQuantConsig(iMes)
        .dSaldoValorConsig(iMes) = tSldMesEstAlm2.adSaldoValorConsig(iMes)
        .dSaldoQuantDemo(iMes) = tSldMesEstAlm2.adSaldoQuantDemo(iMes)
        .dSaldoValorDemo(iMes) = tSldMesEstAlm2.adSaldoValorDemo(iMes)
        .dSaldoQuantConserto(iMes) = tSldMesEstAlm2.adSaldoQuantConserto(iMes)
        .dSaldoValorConserto(iMes) = tSldMesEstAlm2.adSaldoValorConserto(iMes)
        .dSaldoQuantOutros(iMes) = tSldMesEstAlm2.adSaldoQuantOutros(iMes)
        .dSaldoValorOutros(iMes) = tSldMesEstAlm2.adSaldoValorOutros(iMes)
        .dSaldoQuantBenef(iMes) = tSldMesEstAlm2.adSaldoQuantBenef(iMes)
        .dSaldoValorBenef(iMes) = tSldMesEstAlm2.adSaldoValorBenef(iMes)
    
    End With
    
    SldDiaEstAlmFilial_Le1 = SUCESSO
    
    Exit Function
    
Erro_SldDiaEstAlmFilial_Le1:
    
    SldDiaEstAlmFilial_Le1 = gErr
    
    Select Case gErr
    
        Case 94530, 94531
            '??? mudar constante de erro para uma com msg mais apropriada
            'Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAESTALM1", gErr, iAlmoxarifado, objSldDiaEstAlm.iAlmoxarifado)
        
        Case 94535
            'Sem dados. Deve ser tratado na rotina chamadora!
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173797)
    
    End Select
    
    Exit Function
    
End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
Function SldMesEstAlm_Le1(sProduto As String, iMes As Integer, iAno As Integer, iAlmoxarifado As Integer, objSldMesEstAlm As ClassSldMesEstAlm) As Long
'Lê na tabela SldMesEstAlm os campos envolvendo quantidades e valores para o mês passado como parâmetro
'Filtro: iAlmoxarifado, sProduto, iAno
'Os campos com os valores iniciais NÃO SÃO LIDOS

Dim lComando As Long
Dim lErro As Long
Dim sSelect As String
Dim tSldMesEstAlm As typeSldMesEstAlm

On Error GoTo Erro_SldMesEstAlm_Le1

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94552
    
    'Monta o select que será feito em cima da tabela SldDiaEstAlm
    sSelect = "SELECT Almoxarifado, Produto, Ano, QuantEnt" & CStr(iMes) & ",ValorEnt" & CStr(iMes) & ", QuantSai" & CStr(iMes) & ", QuantCons" & CStr(iMes) & ", QuantVend" & CStr(iMes) & ", QuantVendConsig3" & CStr(iMes) & ", ValorSai" & CStr(iMes) & ", QuantComp" & CStr(iMes) & ", ValorComp" & CStr(iMes) & ", ValorVend" & CStr(iMes) & ", ValorVendConsig3" & CStr(iMes) & ", ValorCons" & CStr(iMes) & ",SaldoQuantRecIndl" & CStr(iMes) & ",SaldoValorRecIndl" & CStr(iMes) & ",SaldoQuantCusto" & CStr(iMes) & ",SaldoValorCusto" & CStr(iMes) & "FROM SldMesEstAlm WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
    
    With tSldMesEstAlm
        
        'Lê na tabela SldMesEstAlm os campos envolvendo valores e quantidade com exceção dos INICIAIS, de acordo com o filtro passado.
        lErro = Comando_Executar(lComando, sSelect, .iAlmoxarifado, .sProduto, .iAno, .adQuantEnt(iMes), .adValorEnt(iMes), .adQuantSai(iMes), .adQuantCons(iMes), .adQuantVend(iMes), .adQuantVendConsig3(iMes), .adValorSai(iMes), .adQuantComp(iMes), .adValorVend(iMes), .adValorVendConsig3(iMes), .adValorCons(iMes), .adSaldoQuantRecIndl(iMes), .adSaldoValorRecIndl(iMes), .adSaldoQuantCusto(iMes), .adSaldoValorCusto(iMes), iAlmoxarifado, sProduto, iAno)
        If lErro <> AD_SQL_SUCESSO Then gError 94553
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94554
        
        If lErro = AD_SQL_SEM_DADOS Then gError 94555
        
    End With
    
    With tSldMesEstAlm
    
        'Transfere os dados do type para o obj
        objSldMesEstAlm.sProduto = .sProduto
        objSldMesEstAlm.iAno = .iAno
        objSldMesEstAlm.iAlmoxarifado = .iAlmoxarifado
        objSldMesEstAlm.dQuantEnt(iMes) = .adQuantEnt(iMes)
        objSldMesEstAlm.dValorEnt(iMes) = .adValorEnt(iMes)
        objSldMesEstAlm.dQuantSai(iMes) = .adQuantSai(iMes)
        objSldMesEstAlm.dQuantCons(iMes) = .adQuantCons(iMes)
        objSldMesEstAlm.dQuantVend(iMes) = .adQuantVend(iMes)
        objSldMesEstAlm.dQuantVendConsig3(iMes) = .adQuantVendConsig3(iMes)
        objSldMesEstAlm.dValorSai(iMes) = .adValorSai(iMes)
        objSldMesEstAlm.dQuantComp(iMes) = .adQuantComp(iMes)
        objSldMesEstAlm.dValorComp(iMes) = .adValorComp(iMes)
        objSldMesEstAlm.dValorVend(iMes) = .adValorVend(iMes)
        objSldMesEstAlm.dValorVendConsig3(iMes) = .adValorVendConsig3(iMes)
        objSldMesEstAlm.dValorCons(iMes) = .adValorCons(iMes)
        objSldMesEstAlm.dSaldoQuantRecIndl(iMes) = .adSaldoQuantRecIndl(iMes)
        objSldMesEstAlm.dSaldoValorRecIndl(iMes) = .adSaldoValorRecIndl(iMes)
        objSldMesEstAlm.dSaldoQuantCusto(iMes) = .adSaldoQuantCusto(iMes)
        objSldMesEstAlm.dSaldoValorCusto(iMes) = .adSaldoValorCusto(iMes)
    
    End With

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    SldMesEstAlm_Le1 = SUCESSO
    
    Exit Function

Erro_SldMesEstAlm_Le1:
    
    SldMesEstAlm_Le1 = gErr
    
    Select Case gErr
    
        Case 94552
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94553, 94554
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno, iAlmoxarifado, sProduto)
        
        Case 94555
            'SEM DADOS. Erro tratado na rotina chamadora
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173798)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???

Function SldMesEstAlm1_Le1(sProduto As String, iMes As Integer, iAno As Integer, iAlmoxarifado As Integer, objSldMesEstAlm1 As ClassSldMesEstAlm1) As Long
'Lê na tabela SldMesEstAlm1 os valores dos campos de saldo para o mês passado como parâmetro.
'Filtro: iAlmoxarifado, sProduto, iAno
'Os campos com os valores iniciais NÃO SÃO LIDOS

Dim lComando As Long
Dim lErro As Long
Dim sSelect As String
Dim tSldMesEstAlm1 As typeSldMesEstAlm1

On Error GoTo Erro_SldMesEstAlm1_Le1

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 94548
    
    'Monta o select que será feito em cima da tabela SldDiaEstAlm
    sSelect = "SELECT Almoxarifado, Produto, Ano, SaldoQuantConsig3" & CStr(iMes) & ",SaldoValorConsig3" & CStr(iMes) & ",SaldoQuantDemo3" & CStr(iMes) & ", SaldoValorDemo3" & CStr(iMes) & ", SaldoQuantConserto3" & CStr(iMes) & ", SaldoValorConserto3" & CStr(iMes) & ", SaldoQuantOutros3" & CStr(iMes) & ", SaldoValorOutros3" & CStr(iMes) & ", SaldoQuantBenef3" & CStr(iMes) & ", SaldoValorBenef3" & CStr(iMes) & "FROM SldMesEstAlm1 WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
        
    With tSldMesEstAlm1
        
        lErro = Comando_Executar(lComando, sSelect, .iAlmoxarifado, .sProduto, .iAno, .adSaldoQuantConsig3(iMes), .adSaldoValorConsig3(iMes), .adSaldoQuantDemo3(iMes), .adSaldoValorDemo3(iMes), .adSaldoQuantConserto3(iMes), .adSaldoValorConserto3(iMes), .adSaldoQuantOutros3(iMes), .adSaldoValorOutros3(iMes), .adSaldoQuantBenef3(iMes), .adSaldoValorBenef3(iMes), sProduto, iAno, iAlmoxarifado)
        If lErro <> AD_SQL_SUCESSO Then gError 94549
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 94550
        
        If lErro = AD_SQL_SEM_DADOS Then gError 94551
        
        'Transfere os dados do type para o obj
        objSldMesEstAlm1.dSaldoQuantConsig3(iMes) = .adSaldoQuantConsig3(iMes)
        objSldMesEstAlm1.dSaldoValorConsig3(iMes) = .adSaldoValorConsig3(iMes)
        objSldMesEstAlm1.dSaldoQuantDemo3(iMes) = .adSaldoQuantDemo3(iMes)
        objSldMesEstAlm1.dSaldoValorDemo3(iMes) = .adSaldoValorDemo3(iMes)
        objSldMesEstAlm1.dSaldoQuantConserto3(iMes) = .adSaldoQuantConserto3(iMes)
        objSldMesEstAlm1.dSaldoValorConserto3(iMes) = .adSaldoValorConserto3(iMes)
        objSldMesEstAlm1.dSaldoQuantOutros3(iMes) = .adSaldoQuantOutros3(iMes)
        objSldMesEstAlm1.dSaldoValorOutros3(iMes) = .adSaldoValorOutros3(iMes)
        objSldMesEstAlm1.dSaldoQuantBenef3(iMes) = .adSaldoQuantBenef3(iMes)
        objSldMesEstAlm1.dSaldoValorBenef3(iMes) = .adSaldoValorBenef3(iMes)
    
    End With
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    SldMesEstAlm1_Le1 = SUCESSO
    
    Exit Function

Erro_SldMesEstAlm1_Le1:
    
    SldMesEstAlm1_Le1 = gErr
    
    Select Case gErr
    
        Case 94548
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 94549, 94550
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno, iAlmoxarifado, sProduto)
        
        Case 94551
            'SEM DADOS. Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173799)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

'??? NÃO APAGAR ESSA FUNÇÃO, POIS ELA SERÁ DESCOMENTADA !!!
'Criada por: Luiz G.F.Nogueira
'em: 20/09/01
'Transferida para em:
'Inserida no Rotinas, RotinasModulo e GrupoRotinas: não
'Pendências: sim
'??? Transferir para Rotinas???
'Function SldMesEstAlm2_Le1(sProduto As String, iMes As Integer, iAno As Integer, iAlmoxarifado As Integer, objSldMesEstAlm2 As ClassSldMesEstAlm2) As Long
''Lê na tabela SldMesEstAlm2 os valores dos campos de saldo para o mês passado como parâmetro.
''Filtro: iAlmoxarifado, sProduto, iAno
''Os campos com os valores iniciais NÃO SÃO LIDOS
'
'Dim lComando As Long
'Dim lErro As Long
'Dim sSelect As String
'Dim tSldMesEstAlm2 As typeSldMesEstAlm2
'
'On Error GoTo Erro_SldMesEstAlm2_Le1
'
'    lComando = Comando_Abrir()
'    If lComando = 0 Then gError xxx
'
'    'Monta o select que será feito em cima da tabela SldDiaEstAlm
'    sSelect = "SELECT Almoxarifado, Produto, Ano, SaldoQuantConsig" & CStr(iMes) & ",SaldoValorConsig" & CStr(iMes) & ",SaldoQuantDemo" & CStr(iMes) & ", SaldoValorDemo" & CStr(iMes) & ", SaldoQuantConserto" & CStr(iMes) & ", SaldoValorConserto" & CStr(iMes) & ", SaldoQuantOutros" & CStr(iMes) & ", SaldoValorOutros" & CStr(iMes) & ", SaldoQuantBenef" & CStr(iMes) & ", SaldoValorBenef" & CStr(iMes) & "FROM SldMesEstAlm2 WHERE Almoxarifado = ? AND Produto = ? AND Ano = ?"
'
'    With tSldMesEstAlm2
'
'        lErro = Comando_Executar(lComando, sSelect, .iAlmoxarifado, .sProduto, .iAno, .adSaldoQuantConsig(iMes), .adSaldoValorConsig(iMes), .adSaldoQuantDemo(iMes), .adSaldoValorDemo(iMes), .adSaldoQuantConserto(iMes), .adSaldoValorConserto(iMes), .adSaldoQuantOutros(iMes), .adSaldoValorOutros(iMes), .adSaldoQuantBenef(iMes), .adSaldoValorBenef(iMes), sProduto, iAno, iAlmoxarifado)
'        If lErro <> AD_SQL_SUCESSO Then gError xxx
'
'        lErro = Comando_BuscarPrimeiro(lComando)
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError xxx
'
'        If lErro = AD_SQL_SEM_DADOS Then gError xxx
'
'    End With
'
'    With objSldMesEstAlm2
'
'        'Transfere os dados do type para o obj
'        .dSaldoQuantConsig(iMes) = tSldMesEstAlm2.adSaldoQuantConsig(iMes)
'        .dSaldoValorConsig(iMes) = tSldMesEstAlm2.adSaldoValorConsig(iMes)
'        .dSaldoQuantDemo(iMes) = tSldMesEstAlm2.adSaldoQuantDemo(iMes)
'        .dSaldoValorDemo(iMes) = tSldMesEstAlm2.adSaldoValorDemo(iMes)
'        .dSaldoQuantConserto(iMes) = tSldMesEstAlm2.adSaldoQuantConserto(iMes)
'        .dSaldoValorConserto(iMes) = tSldMesEstAlm2.adSaldoValorConserto(iMes)
'        .dSaldoQuantOutros(iMes) = tSldMesEstAlm2.adSaldoQuantOutros(iMes)
'        .dSaldoValorOutros(iMes) = tSldMesEstAlm2.adSaldoValorOutros(iMes)
'        .dSaldoQuantBenef(iMes) = tSldMesEstAlm2.adSaldoQuantBenef(iMes)
'        .dSaldoValorBenef(iMes) = tSldMesEstAlm2.adSaldoValorBenef(iMes)
'
'    End With
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    SldMesEstAlm2_Le1 = SUCESSO
'
'    Exit Function
'
'Erro_SldMesEstAlm2_Le1:
'
'    SldMesEstAlm2_Le1 = gErr
'
'    Select Case gErr
'
'        Case 94548
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 94549, 94550
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESESTALM11", gErr, iAno, iAlmoxarifado, sProduto)
'
'        Case 94551
'            'SEM DADOS. Erro tratado na rotina chamadora
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173800)
'
'    End Select
'
'    'Fecha o comando
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

Function Acerta_EstoqueProduto_QuantDispNossa() As Long

Dim tItemMovEstoque As typeItemMovEstoque
Dim alComando(1 To 5) As Long
Dim objTipoMovEstoque As ClassTipoMovEst
Dim iIndice As Integer
Dim lErro As Long
Dim objSldDiaEst As New ClassSldDiaEst
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objItemMovEst As ClassItemMovEstoque
Dim dQuantDispNossa As Double

On Error GoTo Erro_Acerta_EstoqueProduto_QuantDispNossa
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1
    Next
    
    With tItemMovEstoque

        .sProduto = String(STRING_PRODUTO, 0)
        .sSiglaUM = String(STRING_UM_SIGLA, 0)
        .sCcl = String(STRING_CCL, 0)
        .sOPCodigo = String(STRING_ORDEM_DE_PRODUCAO, 0)
        .sDocOrigem = String(STRING_MOVESTOQUE_DOCORIGEM, 0)
        .sContaContabilAplic = String(STRING_CONTA, 0)
        .sContaContabilEst = String(STRING_CONTA, 0)

        lErro = Comando_Executar(alComando(1), "SELECT NumIntDoc, FilialEmpresa, Codigo, Custo, Apropriacao, Produto, SiglaUM, Quantidade, Almoxarifado, TipoMov, NumIntDocOrigem, TipoNumIntDocOrigem, Data, Hora, Ccl, NumIntDocEst, Cliente, Fornecedor, CodigoOP, DocOrigem, ContaContabilEst, ContaContabilAplic, HorasMaquina, DataInicioProducao FROM MovimentoEstoque ORDER BY Produto, Almoxarifado, Data, Hora", _
           .lNumIntDoc, .iFilialEmpresa, .lCodigo, .dCusto, .iApropriacao, .sProduto, .sSiglaUM, .dQuantidade, .iAlmoxarifado, .iTipoMov, .lNumIntDocOrigem, .iTipoNumIntDocOrigem, .dtData, .dHora, .sCcl, .lNumIntDocEst, .lCliente, .lFornecedor, .sOPCodigo, .sDocOrigem, .sContaContabilEst, .sContaContabilAplic, .lHorasMaquina, .dtDataInicioProducao)
        If lErro <> AD_SQL_SUCESSO Then gError 1

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1

        Set objTipoMovEstoque = New ClassTipoMovEst
        
        Set objItemMovEst = New ClassItemMovEstoque
        
        objItemMovEst.sProduto = .sProduto
        objItemMovEst.iAlmoxarifado = .iAlmoxarifado
    
    End With
    
    Do While lErro = AD_SQL_SUCESSO
    
        If objItemMovEst.sProduto <> tItemMovEstoque.sProduto Or objItemMovEst.iAlmoxarifado <> tItemMovEstoque.iAlmoxarifado Then

            lErro = Acerta_EstoqueProduto(dQuantDispNossa, objItemMovEst)
            If lErro <> SUCESSO Then gError 1

            dQuantDispNossa = 0
    
        End If
    
        Set objItemMovEst = New ClassItemMovEstoque
        
        Set objEstoqueProduto = New ClassEstoqueProduto
    
        'move os dados de tItemMovEstoque para objItemMovEst
        Call Move_tItemMovEstoque_objItemMovEst(tItemMovEstoque, objItemMovEst)
    
        objTipoMovEstoque.iCodigo = objItemMovEst.iTipoMov
        
        'Lê o tipo de movimento de estoque do item que será processado
        lErro = CF("TiposMovEst_Le1", alComando(2), objTipoMovEstoque)
        If lErro <> SUCESSO Then gError 94534
    
        'transforma a quantidade do movimento na quantidade de estoque
        lErro = Estoque_Transforma_UM_1(alComando(), objItemMovEst)
        If lErro <> SUCESSO Then gError 36108
    
        If objTipoMovEstoque.iAtualizaMovEstoque = TIPOMOV_EST_ESTORNOMOV Then
            objItemMovEst.dQuantidadeEst = -objItemMovEst.dQuantidadeEst
        End If
    
        'Guarda em objEstoqueProduto e em objSldDiaEst os valores e os sinais com os quais deverão ser acumulados
        lErro = CF("Estoque_AtualizaItemMov3", objItemMovEst, objTipoMovEstoque, objEstoqueProduto, objSldDiaEst, REPROCESSAMENTO_REFAZ)
        If lErro <> SUCESSO Then gError 1
    
        dQuantDispNossa = dQuantDispNossa + objEstoqueProduto.dQuantDispNossa
    
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1
    
    Loop
    
    lErro = Acerta_EstoqueProduto(dQuantDispNossa, objItemMovEst)
    If lErro <> SUCESSO Then gError 1
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ACERTO_ESTOQUEPRODUTO_EXECUTADO_SUCESSO")
    
    Acerta_EstoqueProduto_QuantDispNossa = SUCESSO
    
    Exit Function
    
Erro_Acerta_EstoqueProduto_QuantDispNossa:

    Acerta_EstoqueProduto_QuantDispNossa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173801)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Function Acerta_EstoqueProduto(ByVal dQuantDispNossa As Double, ByVal objItemMovEst As ClassItemMovEstoque) As Long

Dim dQuantDispNossa1 As Double
Dim alComando(1 To 2) As Long
Dim lErro As Long
Dim iIndice As Integer
Dim dQuantInicial As Double
Dim dQuantReservada As Double

On Error GoTo Erro_Acerta_EstoqueProduto
    
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 1
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantDispNossa, QuantidadeInicial, QuantReservada FROM EstoqueProduto WHERE Produto = ? AND Almoxarifado = ?", 0, _
       dQuantDispNossa1, dQuantInicial, dQuantReservada, objItemMovEst.sProduto, objItemMovEst.iAlmoxarifado)
    If lErro <> AD_SQL_SUCESSO Then gError 1

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 1
    
    If dQuantDispNossa1 <> dQuantDispNossa + dQuantInicial - dQuantReservada Then
    
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE EstoqueProduto SET QuantDispNossa= ?", alComando(1), dQuantDispNossa + dQuantInicial - dQuantReservada)
        If lErro <> AD_SQL_SUCESSO Then gError 83654
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next
    
    Acerta_EstoqueProduto = SUCESSO
    
    Exit Function
    
Erro_Acerta_EstoqueProduto:

    Acerta_EstoqueProduto = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173802)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Comando_Fechar (alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Estoque_Transforma_UM_1(alComando() As Long, ByVal objItemMovEst As ClassItemMovEstoque) As Long
' transforma a quantidade do movimento na quantidade de estoque

Dim lErro As Long
Dim objUnidadeMedida As New ClassUnidadeDeMedida
Dim dQuantidade As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Estoque_Transforma_UM_1

    objProduto.sCodigo = objItemMovEst.sProduto
    
    'le o atributo controleestoque do produto em questão. Os atributos siglaUM e classeUM se deve a necessidade converter a unidade do movimento para unidade de estoque
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 36036

    objItemMovEst.sSiglaUMEst = objProduto.sSiglaUMEstoque
    objItemMovEst.iClasseUM = objProduto.iClasseUM

    If objItemMovEst.sSiglaUM = objItemMovEst.sSiglaUMEst Then
    
        objItemMovEst.dQuantidadeEst = objItemMovEst.dQuantidade
    
    Else
    
        objUnidadeMedida.iClasse = objItemMovEst.iClasseUM
        objUnidadeMedida.sSigla = objItemMovEst.sSiglaUM
    
        lErro = CF("UM_Le1", alComando(3), objUnidadeMedida)
        If lErro <> SUCESSO Then Error 36112
    
        dQuantidade = objUnidadeMedida.dQuantidade

        objUnidadeMedida.sSigla = objItemMovEst.sSiglaUMEst
    
        lErro = CF("UM_Le1", alComando(5), objUnidadeMedida)
        If lErro <> SUCESSO Then Error 36113
    
        objItemMovEst.dQuantidadeEst = (objItemMovEst.dQuantidade * dQuantidade) / objUnidadeMedida.dQuantidade
        
    End If
    
    Estoque_Transforma_UM_1 = SUCESSO
    
    Exit Function
    
Erro_Estoque_Transforma_UM_1:

    Estoque_Transforma_UM_1 = Err
    
    Select Case Err
    
        Case 36112, 36113
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 173803)
    
    End Select
    
    Exit Function

End Function


