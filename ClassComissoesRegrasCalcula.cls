VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComissoesRegrasCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'Indica que uma fórmula não retornou resultado
Const ERRO_FORMULA_FUNCAO = -9999

'Utilizado para instanciar objComissao
Private sClasseobjComissao As String 'Variável que guarda a classe que será utilizada para instanciar objComissao
Const INSTANCIA_OBJCOMISSAO_PV As String = "GlobaisPV.ClassComissaoPedVendas"
Const INSTANCIA_OBJCOMISSAO_NF As String = "GlobaisCRFAT.ClassComissaoNF"
'**************************************

'Variável que aponta para o espaço de memória que aloca as fórmulas compiladas
Private glExpComp As Long

'Indica se todas as regras já foram compiladas
Private bRegrasCompiladas As Boolean

'Guarda o código do vendedor direto
Private iVendedorDireto As Integer

'Coleções que serão utilizadas para calcular as comissões
Private colcolComissoesRegras As Collection
Private colMnemonicosComissoes As Collection
Private colGridTransacao As ClassColGridTransacao

Public Function Comissoes_Calcula_Regras_Int(ByVal objTela As Object, bCalculou As Boolean, colComissoes As Collection) As Long
'Calcula as comissões baseando-se nas regras de comissões
'objTela RECEBE(Input) as informações da tela que chamou o cálculo de comissões
'bCalculou RETORNA(Output) uma flag indicando se as comissões foram calculadas ou não
'colComissoes RETORNA(Output) as comissões calculadas

Dim lErro As Long
Dim objComissoesRegras As ClassComissoesRegras
Dim objExeExp As New ClassExeExp
Dim colComissoesRegrasTodas As Collection

On Error GoTo Erro_Comissoes_Calcula_Regras_Int

    'Exibe o ponteiro ampulheta
    If Not GL_objMDIForm Is Nothing Then
        GL_objMDIForm.MousePointer = vbHourglass
   
        'Se não houve alteração nos campos que são utilizados para calcular as comissões => sai da função
        If Not objTela.iComissoesAlterada = REGISTRO_ALTERADO Then
        
            'Exibe o ponteiro padrão
            GL_objMDIForm.MousePointer = vbDefault
        
            Comissoes_Calcula_Regras_Int = SUCESSO
            
            Exit Function
    
        End If
    
    End If
    
    'Se as Regras de comissões ainda não foram colocadas na coleção
    If colcolComissoesRegras Is Nothing Then

        'Instancia a coleção que armazenará as regras
        Set colComissoesRegrasTodas = New Collection

        'Lê todas as regras de comissões em ComissoesRegras
        lErro = CF("ComissoesRegras_Le_Todas", colComissoesRegrasTodas)
        If lErro <> SUCESSO And lErro <> 94916 Then gError 94912

        'Se não encontrou => erro
        If lErro = 94916 Then gError 94918
        
        'Preenche colcolComissoesRegras
        Call Move_colComissoesRegrasTodas_colcolComissoesRegras(colComissoesRegrasTodas)
        
    End If

    'Se os Mnemônicos utilizados para o cálculo de comissões ainda não foram colocados na coleção
    If colMnemonicosComissoes Is Nothing Then

        'Instancia a coleção que armazenará os mnemônicos
        Set colMnemonicosComissoes = New Collection

        'Lê os Mnemônicos utilizados para cálculo de comissões
        lErro = CF("MnemonicoComissoes_Todos_Le_Todos", colMnemonicosComissoes)
        If lErro <> SUCESSO And lErro <> 94944 Then gError 94919

        'Se não encontrou mnemônicos em nenhuma das duas tabelas => erro
        If lErro = 94944 Then gError 94928

    End If

    'Alterado por Luiz em 30/04/02 -> esse IF existia para otimizar o código, evitando que
    'essa coleção fosse preenchida várias vezes. Entretanto, a função que preenche a coleção, também atualiza
    'o número de linhas dos grids, por isso foi necessário comentar esse teste e passar a preencher
    'a coleção todas as vezes que se dispara o cálculo das comissões
    
    'Se a coleção de grids da tela ainda não foi preenchida
    'If colGridTransacao Is Nothing Then

        'instancia a coleção que armazenará os grids
        Set colGridTransacao = New ClassColGridTransacao

        'Lê os grids que compõem a tela
        lErro = CF("Transacao_Grids", objTela, colGridTransacao)
        If lErro <> SUCESSO Then gError 94929

    'End If
    '********************************************************************************************

    'Guarda na variável global o tipo do objComissao que será instanciado posteriormente
    'Essa informação é guardada globalmente porque esse objeto será instanciado várias vezes
    sClasseobjComissao = INSTANCIA_OBJCOMISSAO_NF
    
    If Not GL_objMDIForm Is Nothing Then

    Select Case objTela.Name
    
        'Se for a tela de Pedido de Venda
        Case NOME_TELA_PEDIDOVENDA, NOME_TELA_COMISSOESCALCULA
        
            'Guarda o endereço da classe de comissões utilizada dentro de ClassPedidoDeVenda
            sClasseobjComissao = INSTANCIA_OBJCOMISSAO_PV
        
        'Se for uma tela de Nota Fiscal
        Case NOME_TELA_NFISCAL, NOME_TELA_NFISCALFATURA, NOME_TELA_NFISCALPEDIDO, NOME_TELA_NFISCALFATURAPEDIDO, NOME_TELA_CONHECIMENTOFRETEFATURA, NOME_TELA_NFISCALFATURAPEDSRV, NOME_TELA_NFISCALSRV
            
            'Guarda o endereço da classe de comissões utilizada dentro de ClassNFiscal
            sClasseobjComissao = INSTANCIA_OBJCOMISSAO_NF
            
        'Se for outro tipo de tela
        Case Else
            
            'Erro pois a tela não foi tratada
            gError 94951
        
    End Select

    End If
    
    'Guarda a coleção de mnemônicos dentro de objExeExp
    Set objExeExp.colMnemonico = colMnemonicosComissoes

    'Guarda as informações da tela em objExeExp
    Set objExeExp.objTransacao = objTela
    
    'Guarda a coleção de grid's dentro de objExeExp
    Set objExeExp.colGridTransacao = colGridTransacao

    'Instancia a coleção de valores de mnemônicos (a coleção será usada na hora
    'de executar as regras)
    Set objExeExp.colMnemonicoValor = New ClassColMnemonicoValor

    'Instancia a coleção de lancamentos (a coleção será usada para retornar os valores
    'encontrados pela execução das regras)
    Set objExeExp.colLancamentos = New ClassColLancamentos

    'Processa o cálculo das regras de comissões
    lErro = Processa_Comissoes_Regras(objExeExp, VENDEDOR_TODOS, colComissoes)
    If lErro <> SUCESSO Then gError 94971

    'Aqui ainda tem que agrupar os itens encontrados em colComissoes
    lErro = Processa_colComissoes(colComissoes)
    If lErro <> SUCESSO Then gError 102027
    
    'Indica que o cálculo foi efetuado
    bCalculou = True
    
    If Not GL_objMDIForm Is Nothing Then
    
        'Exibe o ponteiro padrão
        GL_objMDIForm.MousePointer = vbDefault
    
    End If
    
    Comissoes_Calcula_Regras_Int = SUCESSO

    Exit Function

Erro_Comissoes_Calcula_Regras_Int:

    Comissoes_Calcula_Regras_Int = gErr

    Select Case gErr

        Case 94912, 94919, 94929, 94971, 102027

        Case 94951
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_TELA_NAO_TRATADA", gErr, objTela.Name)
        
        Case 94918
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOESREGRAS_INEXISTENTES", gErr)

        Case 94928
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOSCOMISSOESTODOS_INEXISTENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147569)

    End Select

    If Not GL_objMDIForm Is Nothing Then

        'Exibe o ponteiro padrão
        GL_objMDIForm.MousePointer = vbDefault
    
    End If
    
    Exit Function

End Function

Private Sub Move_colComissoesRegrasTodas_colcolComissoesRegras(colComissoesRegras As Collection)
'Separa as regras diretas das indiretas guardando-as em coleções diferentes
'E adicionando essas coleções à coleção de coleções
'colComissoesRegras RECEBE(Input) as regras que serão separadas

Dim objComissoesRegras As New ClassComissoesRegras
Dim colComissoesRegrasDiretas As New Collection
Dim colComissoesRegrasIndiretas As New Collection

On Error GoTo Erro_Move_colComissoesRegrasTodas_colcolComissoesRegras

    'Instancia a coleção de coleções de regras
    Set colcolComissoesRegras = New Collection
    
    'Para cada regra lida do BD
    For Each objComissoesRegras In colComissoesRegras
        
        'Se a regra for para comissão direta
        If objComissoesRegras.iVendedorIndireto = VENDEDOR_DIRETO Then
            
            'Guarda a regra na coleção de regras diretas
            colComissoesRegrasDiretas.Add objComissoesRegras
        
        'Se a regra for para comissão indireta
        ElseIf objComissoesRegras.iVendedorIndireto = VENDEDOR_INDIRETO Then

            'Guarda a regra na coleção de regras indiretas
            colComissoesRegrasIndiretas.Add objComissoesRegras
        
        End If
    
    Next
    
    'Adiciona à coleção de coleções a coleção com todas as regras
    colcolComissoesRegras.Add colComissoesRegras, VENDEDOR_TODOS_STRING
    
    'Adiciona à coleção de coleções a coleção com as regras diretas
    colcolComissoesRegras.Add colComissoesRegrasDiretas, VENDEDOR_DIRETO_STRING
    
    'Adiciona à coleção de coleções a coleção com as regras indiretas
    colcolComissoesRegras.Add colComissoesRegrasIndiretas, VENDEDOR_INDIRETO_STRING
    
    Exit Sub
    
Erro_Move_colComissoesRegrasTodas_colcolComissoesRegras:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 147570)

    End Select

    Exit Sub

End Sub

Private Function Processa_Comissoes_Regras(ByVal objExeExp As ClassExeExp, ByVal iComissoesTipo As Integer, colComissoes As Collection) As Long
'Inicia o processo de execução das regras para cálculo de comissões
'objExeExp RECEBE(Input) os mnemônicos e os grids necessários para executar as fórmulas
'colComissoes RETORNA(Output) as comissões calculadas para serem exibidas no grid

Dim lErro As Long

On Error GoTo Erro_Processa_Comissoes_Regras

    'Se ainda não compilou todas as regras
    If Not bRegrasCompiladas Then
    
        'Se ainda não foi alocada memória para armazenar as fórmulas compiladas
        If glExpComp = 0 Then
    
            'Inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil_Comissoes(glExpComp)
            If lErro <> SUCESSO Then gError 94930
        End If

        'Monta as fórmulas e compila-as
        lErro = Prepara_ComissoesRegras_Execucao(objExeExp, iComissoesTipo)
        If lErro <> SUCESSO Then gError 94959

    End If

    'Descobre a quantidade total de linhas envolvidas no cálculo
    lErro = Descobre_Numero_Linhas_Grids(objExeExp, AddressOf MaqExp_Grid_Linhas, glExpComp)
    If lErro <> SUCESSO Then gError 102029

    'Executa as fórmulas já compiladas
    lErro = Executa_ComissoesRegras(objExeExp, iComissoesTipo, colComissoes)
    If lErro <> SUCESSO Then gError 94960

    'Se é para calcular comissões diretas e indiretas
    If iComissoesTipo = VENDEDOR_TODOS Then

        'Processa novamente as regras de comissões, indicando, agora
        'que é para processar apenas as regras para comissões indiretas
        lErro = Processa_Comissoes_Regras(objExeExp, VENDEDOR_INDIRETO, colComissoes)
        If lErro <> SUCESSO Then gError 94972

    End If

    Processa_Comissoes_Regras = SUCESSO

    Exit Function

Erro_Processa_Comissoes_Regras:

    Processa_Comissoes_Regras = gErr

    Select Case gErr

        Case 94930, 94959, 94960, 94972

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147571)

    End Select

    Exit Function

End Function

Private Function Prepara_ComissoesRegras_Execucao(ByVal objExeExp As ClassExeExp, ByVal iComissoesTipo As Integer) As Long
'Monta as regras para execução e compila-as
'objExeExp RETORNA(Output) as regras compiladas
'colComissoesRegras RECEBE(Input) as regras de comissões que serão montadas e compiladas
'iComissoesTipo RECEBE(Input) uma flag que indica quais as regras que é para montar e compilar

'ATENÇÃO:
'Aqui apenas a compilação é feita. A execução que retornará o resutlado,
'só será feita mais tarde, para todas as regras de uma vez só

Dim lErro As Long
Dim iComissoesTipoAux As Integer
Dim sVendedor As String
Dim sFormula As String
Dim objComissoesRegras As ClassComissoesRegras

On Error GoTo Erro_Prepara_ComissoesRegras_Execucao

    '*** MONTAGEM E COMPILAÇÃO DAS FÓRMULAS ***
    'Se vai calcular as comissões diretas
    If (iComissoesTipo <> VENDEDOR_INDIRETO) Then

        'Monta a primeira parte da fórmula que será executada para encontrar o vendedor
        'No caso das comissões diretas a fórmula do vendedor será sempre o mnemônico
        'que calcula o vendedor relacionado à filialcliente na tela em questão
        lErro = Monta_Formula_Expressao(sVendedor, FORMULA_FUNCAO_NUMTEXTO, MNEMONICO_COMISSOES_VENDEDOR)
        If lErro <> SUCESSO Then gError 94955

        'Indica que vai calcular as comissões diretas
        iComissoesTipoAux = VENDEDOR_DIRETO
    
    Else
    
        'Indica que vai calcular as comissões indiretas
        iComissoesTipoAux = VENDEDOR_INDIRETO

    End If

    'Para cada regra na coleção que contém todas as regras
    For Each objComissoesRegras In colcolComissoesRegras.Item(VENDEDOR_TODOS_STRING)

        'Se a regra é do tipo de comissão que se está calculando
        If objComissoesRegras.iVendedorIndireto = iComissoesTipoAux Then

            '*** TRATAMENTO PARA O VENDEDOR ***
            'Se está calculando uma comissão indireta
            If iComissoesTipoAux = VENDEDOR_INDIRETO Then

                'Monta a fórmula que será executada para o vendedor
                'No caso das comissões indiretas o vendedor já é guardado no BD e o seu código
                'e é passado aqui apenas para compor a fórmula
                lErro = Monta_Formula_Expressao(sVendedor, FORMULA_FUNCAO_NUMTEXTO, objComissoesRegras.iVendedor)
                If lErro <> SUCESSO Then gError 94958

            End If

            'Compila a fórmula que será executada para obter o vendedor
            lErro = CompilaExe_Formula_Contabil_Comissoes(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sVendedor, CAMPO_CONTA, glExpComp)
            If lErro <> SUCESSO Then gError 94949
            '*** FIM DO TRATAMENTO PARA O VENDEDOR ***

            '*** TRATAMENTO PARA VALORBASE ***
            'se foi especificada uma formula para o valor base ==> compila-a
            If Len(objComissoesRegras.sValorBase) > 0 Then

                'Guarda a regra de ValorBase em sFormula
                sFormula = objComissoesRegras.sValorBase
                
                'Monta a fórmula que será executada para o valor base
                'A fórmula do valor base será calculada caso a condição gravada no BD seja verdadeira
                'Se a condição for falsa, retornará erro -9999
                lErro = Monta_Formula_Expressao(sFormula, FORMULA_FUNCAO_IF, objComissoesRegras.sRegra, objComissoesRegras.sValorBase, ERRO_FORMULA_FUNCAO)
                If lErro <> SUCESSO Then gError 94956

                'Compila a fórmula
                lErro = CompilaExe_Formula_Contabil_Comissoes(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, CAMPO_CREDITO, glExpComp)
                If lErro <> SUCESSO Then gError 94950

            End If
            '*** FIM DO TRATAMENTO PARA VALORBASE ***

            '*** TRATAMENTO PARA PERCCOMISSAO ***
            'se foi especificada uma formula para o percentual de comissão ==>
            If Len(objComissoesRegras.sPercComissao) > 0 Then

                'Guarda a regra de Percentual em sFormula
                sFormula = objComissoesRegras.sPercComissao
                
                'Monta a fórmula que será executada para o percentual de comissão
                'A fórmula do percentual de comissão será calculada caso a condição gravada no BD seja verdadeira
                'Se a condição for falsa, retornará erro -9999
                lErro = Monta_Formula_Expressao(sFormula, FORMULA_FUNCAO_IF, objComissoesRegras.sRegra, objComissoesRegras.sPercComissao, ERRO_FORMULA_FUNCAO)
                If lErro <> SUCESSO Then gError 94957

                'Compila a fórmula
                lErro = CompilaExe_Formula_Contabil_Comissoes(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, CAMPO_DEBITO, glExpComp)
                If lErro <> SUCESSO Then gError 94931

            End If
            '*** FIM DO TRATAMENTO PARA PERCCOMISSAO ***

        End If

    Next
    '*** FIM DA MONTAGEM E COMPILAÇÃO DAS FÓRMULAS ***

    'Se for para montar e compilar todas as regras
    If iComissoesTipo = VENDEDOR_TODOS Then

        'Chama novamente essa mesma função, indicando que é para montar
        'e compilar apenas as regras indiretas
        lErro = Prepara_ComissoesRegras_Execucao(objExeExp, VENDEDOR_INDIRETO)
        If lErro <> SUCESSO Then gError 94973

    End If

    'Indica que todas as regras já foram compiladas
    bRegrasCompiladas = True
    
    Prepara_ComissoesRegras_Execucao = SUCESSO

    Exit Function

Erro_Prepara_ComissoesRegras_Execucao:

    Prepara_ComissoesRegras_Execucao = gErr

    Select Case gErr

        Case 94931, 94949, 94950, 94955 To 94958, 94973, 102029

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147572)

    End Select

    Exit Function

End Function

Private Function Monta_Formula_Expressao(sFormula As String, sNomeFormula As String, ParamArray avParams()) As Long
'Recebe o nome de uma fórmula e um conjunto de parâmetros e devolve a fórmula com os
'parâmetros entre parêmteses
'sFormula RETORNA(Output) a fórmula montada
'sNomeFormula RECEBE(Input) o nome da fórmula
'avParams RECEBE(Input) os parâmetros que serão acrescentados à fórmula

Dim iIndice As Integer
Dim sParametros As String
Dim vValorAux As Variant

On Error GoTo Erro_Monta_Formula_Expressao

    'Se a fórmula não foi informada => erro
    If Len(Trim(sNomeFormula)) = 0 Then gError 94954

    'Abre parênteses após o texto da fórmula
    sFormula = sNomeFormula & "("

    'Se o índice inferior do array de parâmetros é menor que o índice superior
    'Significa que foram passados parâmetros
    If LBound(avParams) <= UBound(avParams) Then

        'Para cada parâmetro passado
        For iIndice = LBound(avParams) To UBound(avParams)

            'Monta uma string concatenando todos os parâmetros separados por vírgulas
            sParametros = sParametros & avParams(iIndice) & SEPARADOR_FORMULA_FUNCAO

        Next

        'Acrescenta os parâemtros ao texto da fórmula
        sFormula = sFormula & Mid(sParametros, 1, Len(sParametros) - 1)

    End If

    'Fecha parênteses
    sFormula = sFormula & ")"

    Monta_Formula_Expressao = SUCESSO

    Exit Function

Erro_Monta_Formula_Expressao:

    Monta_Formula_Expressao = gErr

    Select Case gErr

        Case 94954
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMULA_NAO_INFORMADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147573)

    End Select

    Exit Function

End Function

Private Function Executa_ComissoesRegras(ByVal objExeExp As ClassExeExp, ByVal iComissoesTipo As Integer, ByVal colComissoes As Collection) As Long
'Verifica o tipo das comissões que serão calculadas e chama a função apropriada
'objExeExp RECEBE(Input) as fórmulas compiladas que serão repassadas para a função que efetivamente fará a execução
'iComissoesTipo RECEBE(Input) o tipo de comissões que serão calculadas
'colComissoes RETORNA(Output) as comissões calculadas

Dim lErro As Long

On Error GoTo Erro_Executa_ComissoesRegras
    
    'Se é para calcular as comissões indiretas
    If iComissoesTipo = VENDEDOR_INDIRETO Then
    
        'Executa as regras para comissões indiretas, separando-as por vendedor
        lErro = Executa_ComissoesRegras_Indiretas(objExeExp, colComissoes)
        If lErro <> SUCESSO Then gError 94975
        
    'Senão
    Else
    
        'Executa todas as regras para comissões diretas
        lErro = Executa_ComissoesRegras1(objExeExp, colcolComissoesRegras.Item(VENDEDOR_DIRETO_STRING), 0, colComissoes)
        If lErro <> SUCESSO Then gError 94976
        
        'Se a execução retornou pelo uma comissão direta => guarda o vendedor em uma variável global
        If colComissoes.Count > 0 Then iVendedorDireto = colComissoes.Item(1).iCodVendedor
        
    End If
    
    Executa_ComissoesRegras = SUCESSO

    Exit Function

Erro_Executa_ComissoesRegras:

    Executa_ComissoesRegras = gErr

    Select Case gErr

        Case 94975, 94976
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147574)

    End Select

    Exit Function

End Function

Private Function Executa_ComissoesRegras_Indiretas(ByVal objExeExp As ClassExeExp, colComissoes As Collection)
'Separa as regras indiretas em blocos que contém o mesmo vendedor e processa esses blocos
'Até que todas as regras indiretas tenham sido processadas
'objExeExp RECEBE(Input) as fórmulas compiladas que serão repassadas para a função que efetivamente fará a execução
'colComissoes RETORNA(Output) as comissões calculadas

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim sVendedorAnterior As String
Dim objComissoesRegras As ClassComissoesRegras
Dim colComissoesRegrasCalcular As New Collection

On Error GoTo Erro_Executa_ComissoesRegras_Indiretas

    'Se não existem regras indiretas => sai da função
    If colcolComissoesRegras.Item(VENDEDOR_INDIRETO_STRING).Count = 0 Then Exit Function
    
    'Guarda o vendedor da primeira regra (como não há regra anterior, na primeira passagem
    'o vendedor anterior é o mesmo vendedor atual obrigatoriamente)
    sVendedorAnterior = colcolComissoesRegras.Item(VENDEDOR_INDIRETO_STRING).Item(1).iVendedor
    
    'Guarda o índice da primeira regra indireta na coleção que guarda todas as regras
    'iIndice é passado para Executa_ComissoesRegras1, para indicar a posição da primeira
    'regra que será executada
    iIndice = colcolComissoesRegras(VENDEDOR_DIRETO_STRING).Count
    
    'Inicializa o índice auxiliar com o mesmo valor de iIndice
    'iIndiceAux vai sendo incrementada para descobrir o iIndice do próximo conjunto
    'de regras que será executado
    iIndiceAux = iIndice
    
    'Para cada regra na coleção de regras indiretas
    For Each objComissoesRegras In colcolComissoesRegras.Item(VENDEDOR_INDIRETO_STRING)
    
        'Se o vendedor é diferente do vendedor direto => significa que é uma regra indireta válida
        'e que deve ser processada
        If (objComissoesRegras.iVendedor <> iVendedorDireto) Then
        
            'Se o vendedor da regra atual é o mesmo da regra anterior
            If (objComissoesRegras.iVendedor = sVendedorAnterior) Then
            
                'Guarda a regra na coleção de regras que serão calculadas
                colComissoesRegrasCalcular.Add objComissoesRegras
            
            'Senão, ou seja, se mudou o vendedor
            Else
            
                'Executa as regras que estão guardadas na coleção
                lErro = Executa_ComissoesRegras1(objExeExp, colComissoesRegrasCalcular, iIndice, colComissoes)
                If lErro <> SUCESSO Then gError 94974
                
                'Instancia uma nova coleção para armazenar as regras do novo vendedor
                Set colComissoesRegrasCalcular = New Collection
                
                'Guarda o novo índice
                iIndice = iIndiceAux
                
                'Guarda a regra na nova coleção de regras a serem calculadas
                colComissoesRegrasCalcular.Add objComissoesRegras
            
            End If
        
        End If
    
        'Guarda o vendedor da regra atual para comparar com o vendedor da próxima regra
        sVendedorAnterior = objComissoesRegras.iVendedor
        
        'Incrementa o índice auxiliar usado para descobrir o índice inicial do próximo
        'conjunto de regras que será executado
        iIndiceAux = iIndiceAux + 1
    
    Next

    
    'Se ainda tem pelo menos uma regra a ser calculada
    If colComissoesRegrasCalcular.Count > 0 Then
    
        'Executa as regras que estão guardadas na coleção
        lErro = Executa_ComissoesRegras1(objExeExp, colComissoesRegrasCalcular, iIndice, colComissoes)
        If lErro <> SUCESSO Then gError 102032
    
    End If
    
    Executa_ComissoesRegras_Indiretas = SUCESSO
    
    Exit Function
    
Erro_Executa_ComissoesRegras_Indiretas:

    Executa_ComissoesRegras_Indiretas = gErr
    
    Select Case gErr

        Case 94974, 102032
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147575)

    End Select

    Exit Function
    
End Function

Private Function Executa_ComissoesRegras1(ByVal objExeExp As ClassExeExp, ByVal colComissoesRegrasCalcular As Collection, ByVal iIndice As Integer, ByVal colComissoes As Collection)
'Faz a execução efetiva das fórmulas compiladas anteriormente e retorna uma coleção com os resultados obtidos
'objExeExp RECEBE(Input) as fórmulas compiladas
'colComissoesRegrasCalcular RECEBE(Input) as regras que serão executadas
'iIndice RECEBE(Input) o índice da primeira regra a ser executada
'colComissoes RETORNA(Output) as comissões encontradas a partir da execução das regras

Dim lErro As Long
Dim iLinha As Integer
Dim objComissao As Object
Dim objComissoesRegras As ClassComissoesRegras
Dim objLancamento As ClassLancamentos
Dim colLinhasNaoCalculadas As Collection
Dim iIndiceAux As Integer
Dim iIndiceProxResult As Integer
Dim bForcaInclusao As Boolean

On Error GoTo Erro_Executa_ComissoesRegras1

    '*** EXECUÇÃO DAS REGRAS ***
    'Para cada regra na coleção
    For Each objComissoesRegras In colComissoesRegrasCalcular

        'Guarda o índice do próximo resultado que deve ser processado (o primeiro item incluído em colLancamentos pela próxima execução)
        'Esse índice é igual ao count da coleção de Lançamentos (que guarda o resultado das execuções) mais 1
        iIndiceProxResult = objExeExp.colLancamentos.Count + 1
        
        'Alterado por Tulio em 07/02/03 => alterada função passada como parâmetro de MaqExp_Devolve_Valor_Total_Mnemonico para MaqExp_Devolve_Valor_Total_MnemonicoComiss
        'Executa a regra que foi compilada anteriormente
        lErro = Executa_Formula_Contabil_Comissoes(objExeExp, AddressOf MaqExp_Grid_Linhas, glExpComp, AddressOf MaqExp_Armazena_Campo, AddressOf MaqExp_Devolve_Valor_MnemonicoComiss, AddressOf MaqExp_Devolve_Valor_Total_MnemonicoComiss, iIndice)
        If lErro <> SUCESSO Then gError 94934

        'Se não executou = > erro
        'If Not objExeExp.colLancamentos Is Nothing Then
            
        '*** INICIALIZAÇÃO DA COLEÇÃO QUE GUARDA AS LINHAS NÃO CALCULADAS ***
        'A inicialização dessa coleção é feita uma única vez para cada conjunto de regras executadas,
        'e tem que ser feita depois da execução da primeira regra do conjunto, pois é quando temos a
        'informação do número total de linhas a serem calculadas (objExeExp.colLancamentos.Count - iIndiceProxResult + 1)
        
        'Se a coleção de linhas não calculadas não foi instanciada
        If colLinhasNaoCalculadas Is Nothing Then

            'Instancia a coleção
            Set colLinhasNaoCalculadas = New Collection

            'Para cada elemento em colLancamentos que ainda não foi processado
            For iLinha = 1 To (objExeExp.colLancamentos.Count - iIndiceProxResult + 1)
                
                'Adiciona um índice referente à linha que deve ser calculada
                colLinhasNaoCalculadas.Add iLinha
               
            Next

        End If
        '*** FIM DA INICIALIZAÇÃO DA COLEÇÃO DE LINHAS NÃO CALCULADAS ***

        'Limpa a variável que controla a linha que está sendo calculada
        iLinha = 0

        'Para cada linha que ainda deve ser calculada
        For iIndiceAux = colLinhasNaoCalculadas.Count To 1 Step -1

            'Calcula a linha que terá o resultado processado. Como as linhas são processada em ordem inversa,
            'ou seja, da última para a primeira, para encontrar o número da linha é preciso pegar o índice
            'do primeiro resultado gerado pela última execução, subtrair de um, para garantir que ele esteja no
            'intervalo e adicionar o número dessa linha que está guardado na coleção
            iLinha = iIndiceProxResult - 1 + colLinhasNaoCalculadas.Item(iIndiceAux)
            
            'Instancia o objLancamento que contém os dados referentes à linha
            Set objLancamento = objExeExp.colLancamentos(iLinha)

            'Se  valor base(Credito) ou percentual comissão(Debito)
            If (objLancamento.dCredito <> ERRO_FORMULA_FUNCAO) Or (objLancamento.dDebito <> ERRO_FORMULA_FUNCAO) Then

                'Instancia um novo objComissao
                Set objComissao = CreateObject(sClasseobjComissao)
                
                objComissao.iLinha = iLinha - iIndiceProxResult + 1
                
                'Guarda em objComissao os dados obtidos pela execução da regra
                objComissao.iCodVendedor = StrParaInt(objLancamento.sConta)
                objComissao.dValorBase = objLancamento.dCredito
                objComissao.dPercentual = objLancamento.dDebito
                
                'o percentual de comissão de emissão não precisa ser executado,
                'pois já é armazenado o valor desejado
                objComissao.dPercentualEmissao = objComissoesRegras.dPercComissaoEmiss

                'Adiciona o obj a coleção
                lErro = CF("Executa_ComissaoRegra1_Cust", bForcaInclusao)
                If lErro <> SUCESSO Then gError 94934
                
                If objComissao.dPercentual <> 0 Or bForcaInclusao Then colComissoes.Add objComissao

                'Remove a linha da coleção de linhas não calculadas
                colLinhasNaoCalculadas.Remove iIndiceAux

            End If
            
        Next
        
        'Se todas as linhas já foram calculadas => sai do loop
        If colLinhasNaoCalculadas.Count = 0 Then Exit For
    
        'Incrementa o índice da regra a ser executada
        iIndice = iIndice + 1
        
        'End If

    Next
    
    Executa_ComissoesRegras1 = SUCESSO
    
    Exit Function
    
Erro_Executa_ComissoesRegras1:

    Executa_ComissoesRegras1 = gErr
    
    Select Case gErr

        Case 94934, 102024, 102028
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147576)

    End Select

    Exit Function

End Function

Private Function Processa_colComissoes(colComissoes As Collection) As Long
'Agrupa os itens em colComissoes para o mesmo vendedor, pois o cálculo executado
'anteriormente, pode retornar várias linhas para o mesmo vendedor, que deverá constar uma única vez no grid
'colComissoes RECEBE(Input) as comissões que serão agrupadas e
'RETORNA (Output) as comissões já agrupadas

Dim lErro As Long
Dim iVendedorAnterior As Integer
Dim dValorComissao As Double
Dim dValorComissaoEmiss As Double
Dim objComissao As Object
Dim objComissaoAux As Object
Dim colComissoesAux As New Collection
Dim sClassObjeto As String

On Error GoTo Erro_Processa_colComissoes

    'Se não foi calculada pelo menos uma comissão, sai da função, pois não tem  o que agrupar
    If colComissoes.Count = 0 Then Exit Function

    'Instancia um novo objComissaoAux
    Set objComissaoAux = CreateObject(sClasseobjComissao)
    
    'Guarda o código do vendededor da primeira comissão em colComissoes
    objComissaoAux.iCodVendedor = colComissoes.Item(1).iCodVendedor

    'Para cada comissão na coleção
    For Each objComissao In colComissoes

        'Se o vendedor da comissão atual é o mesmo vendedor da comissão anterior
        If objComissao.iCodVendedor = objComissaoAux.iCodVendedor Then

            'Acumula os valores para esse vendedor
            lErro = Comissao_Acumula(objComissao, objComissaoAux)
            If lErro <> SUCESSO Then gError 102035

        'Senão, ou seja, se mudou o vendedor
        Else

            'Calcula os percentuais e guarda o obj na coleção
            lErro = Move_Comissao_Colecao(objComissaoAux, colComissoesAux)
            If lErro <> SUCESSO Then gError 102033
            
            'Instancia um novo objComissaoAux para armazenar os dados do novo vendedor
            Set objComissaoAux = CreateObject(sClasseobjComissao)
            
            'Guarda o código do novo vendedor no obj
            objComissaoAux.iCodVendedor = objComissao.iCodVendedor

            'Acumula os valores o novo vendedor
            lErro = Comissao_Acumula(objComissao, objComissaoAux)
            If lErro <> SUCESSO Then gError 102036

        End If

    Next

    'Calcula os percentuais e guarda o obj na coleção
    lErro = Move_Comissao_Colecao(objComissaoAux, colComissoesAux)
    If lErro <> SUCESSO Then gError 102034
    
    'Guarda em colComissoes as comissões que foram agrupadas em colComissoesAux
    Set colComissoes = colComissoesAux
    
    Processa_colComissoes = SUCESSO
    
    Exit Function
    
Erro_Processa_colComissoes:

    Processa_colComissoes = gErr
    
    Select Case gErr

        Case 102025, 102026, 102033 To 102036
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147577)

    End Select

    Exit Function
    
End Function

Private Function Comissao_Acumula(objComissao As Object, objComissaoAux As Object) As Long
'Acumula ao objComissaoAux os seguintes valores: ValorBase, ValorComissao, ValorComissaoEmissao e ValorComissaoBaixa
'objComissao RECEBE(Input) os valores que serão acumulados
'objComissaoAux RETORNA(Output) os valores acumulados

Dim dValorComissao As Double
Dim objComissaoPorItem As ClassComissaoPorItem

On Error GoTo Erro_Comissao_Acumula

    'Encontra o valor dessa comissão
    dValorComissao = (objComissao.dValorBase * objComissao.dPercentual)
    
    'Encontra e acumula o valor que será pago na emissão
    objComissaoAux.dValorEmissao = objComissaoAux.dValorEmissao + (dValorComissao * objComissao.dPercentualEmissao)
    
    'Soma o valor base de comissão ao valor base de comissões para esse vendedor
    objComissaoAux.dValorBase = objComissaoAux.dValorBase + objComissao.dValorBase
    
    'Soma o valor da comissão ao valor de comissões para esse vendedor
    objComissaoAux.dValor = objComissaoAux.dValor + dValorComissao

    Set objComissaoPorItem = New ClassComissaoPorItem

    With objComissaoPorItem
    
        .iLinha = objComissao.iLinha
        .dValorBase = objComissao.dValorBase
        .dPercentual = objComissao.dPercentual
        .dValor = dValorComissao
        .dPercentualEmissao = objComissao.dPercentualEmissao
        .dValorEmissao = dValorComissao * objComissao.dPercentualEmissao
    
    End With
    
    objComissaoAux.colLinhas.Add objComissaoPorItem

    Comissao_Acumula = SUCESSO
    
    Exit Function
    
Erro_Comissao_Acumula:

    Comissao_Acumula = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147578)
            
    End Select
    
    Exit Function

End Function

Private Function Move_Comissao_Colecao(ByVal objComissao As Object, colComissoes As Collection) As Long
'Calcula os percentuais envolvendo a comissão e guarda a comissão na coleção
'objComissao RECEBE(Input) os valores que serão utilizados para calcular os percentuais e que serão guardados na coleção
'colComissao RETORNA(Output) a coleção com o obj adicionado

On Error GoTo Erro_Move_Comissao_Colecao

    '*** CÁLCULO OS PERCENTUAIS DE COMISSÃO - INÍCIO ***
    'Se já foi calculado um valor base
    If objComissao.dValorBase <> 0 Then
    
        'Calcula o percentual do valor total acumulado de comissões
        objComissao.dPercentual = objComissao.dValor / objComissao.dValorBase
    
    End If

    'Se foi encontrado um valor para essa comissão
    If objComissao.dValor <> 0 Then
        
        'Calcula o percentual do valor acumulado a ser pago na emissão
        objComissao.dPercentualEmissao = objComissao.dValorEmissao / objComissao.dValor
    
    End If
    
    '*** CÁLCULO OS PERCENTUAIS DE COMISSÃO - FIM ***
    
    'Se o Vendedor é diferente do vendedor direto => significa que essa é uma comissão indireta
    If objComissao.iCodVendedor <> iVendedorDireto Then objComissao.iIndireta = VENDEDOR_INDIRETO
    
    objComissao.iSeq = colComissoes.Count + 1
        
    'Adiciona o obj com os dados do vendedor anterior a coleção auxiliar
    colComissoes.Add objComissao
    
    Move_Comissao_Colecao = SUCESSO
    
    Exit Function

Erro_Move_Comissao_Colecao:

    Move_Comissao_Colecao = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147579)

    End Select

End Function

Private Sub Class_Terminate()
'Evento de encerramento da classe

Dim lErro As Long

On Error GoTo Erro_Class_Terminate

    'Libera memória alocada
    Set colcolComissoesRegras = Nothing
    Set colMnemonicosComissoes = Nothing
    Set colGridTransacao = Nothing

    If glExpComp <> 0 Then

        'libera o espaço alocado para guardar as formulas compiladas
        lErro = Finaliza_Formula_Contabil_Comissoes(glExpComp)
        If lErro <> SUCESSO Then gError 94948

    End If
    
    Exit Sub

Erro_Class_Terminate:

    Select Case gErr

        Case 94948

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147580)

    End Select

    Exit Sub

End Sub

'*** FUNÇÕES UTILIZADAS PELA TELA COMISSOESREGRAS - INÍCIO ***

'Incluida por Tulio em 30/04 durante o desenvolvimento da tela ComissoesRegras
Public Function Valida_Formula_Comissoes(ByVal sFormula As String, ByVal colMnemonicosComissoes As Collection, ByVal iTipoFormula As Integer) As Long
'Verifica se a formula do parametro sFormula esta correta
'sFormula RECEBE(Input) a fórmula que será validada
'colMnemonicosComissoes RECEBE(Inpu) os mnemônicos que servirão de base para validar a fórmula
'iTipoFormula RECEBE(Input) o tipo de dados que é esperado que a fórmula retorne

Dim objExeExp As New ClassExeExp
Dim iTam As Integer, iTipoCompilado As Integer
Dim lErro As Long

On Error GoTo Erro_Valida_Formula_Comissoes

    'instancia a colMnemonico de objExeExp como a colMnemonicos global da tela
    Set objExeExp.colMnemonico = colMnemonicosComissoes

    'Chama a funcao que compila a formula
    lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, iTipoCompilado, 0, iTam)
    If lErro <> SUCESSO Then gError 101623
    
    'testar se o tipo do mnemônico compilado é igual ao tipo esperado
    If iTipoCompilado <> iTipoFormula Then gError 101624
    
    Valida_Formula_Comissoes = SUCESSO
    
    Exit Function

Erro_Valida_Formula_Comissoes:

    Valida_Formula_Comissoes = gErr

    Select Case gErr
    
        Case 101623
        
        Case 101624
            
            iTam = Len(sFormula)
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, Obter_Tipo(iTipoCompilado), Obter_Tipo(iTipoFormula))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147581)

    End Select

    Exit Function

End Function

Function Obter_Tipo(ByVal iTipo As Integer) As String
'iTipo, parametro de INPUT, representa uma string...
'essa funcao retorna a string q ele representa..
'funcao utilizada para ser utilizado em mensagens de erro contribuindo
'para uma melhor elaboracao da mesma

    'seleciona o tipo, dependendo dele, retorna a string adequada..
    Select Case iTipo
        
        Case TIPO_NUMERICO
            Obter_Tipo = NUMERICO
            
        Case TIPO_DATA
            Obter_Tipo = Data
            
        Case TIPO_HORA
            Obter_Tipo = HORA

        Case TIPO_TEXTO
            Obter_Tipo = TEXTO

        Case TIPO_BOOLEANO
            Obter_Tipo = BOOLEANO

        Case Else
            Obter_Tipo = INVALIDO
        
    End Select
    
End Function

'*** FUNÇÕES UTILIZADAS PELA TELA COMISSOESREGRAS - FIM ***
