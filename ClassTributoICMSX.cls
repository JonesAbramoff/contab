VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoICMSX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ClassTributoDriver

'saidaCli: já foi feito
'entradacli: vou tratar como saida cli mas vou guardar a info se credita do imposto
    '???
'entradaforn: criar tratamento especifico (usar uforigem, padraotribentrada, )
'saidaforn: criar tratamento analogo ao de entradaforn mas nao vou ter info de credito

'cache vs atualizacao
'versoes nas tabelas: estados , ICMSExcecoes, ICMSAliquotasExternas
'manter comandos preparados na conexao browse
'ou ler sempre no bd
'no futuro posso extender interface p/tratar de maneira otimizada uma alteracao
'em complemento (frete, seguro,...) de maneira analoga ao tratamento dado a alteracao de item

'Variaveis a nivel de doc:
'ICMS_BASE 'base reduzida
'ICMS_VALOR
'ICMS_SUBST_BASE 'base reduzida
'ICMS_SUBST_VALOR
'ICMS_VALOR_CREDITO
'
'Variaveis a nivel de item:
'ICMS_TIPO
'ICMS_BASE
'ICMS_RED_BASE
'ICMS_ALIQUOTA
'ICMS_VALOR
'ICMS_BASE_SUBST
'''''ICMS_MARG_SUBST
'ICMS_ALIQUOTA_SUBST
'ICMS_VALOR_SUBST
'ICMS_VALOR_CREDITO
'
'red base: % de reducao 40% = multiplicar p/0.6

'a nivel de doc aparece a base reduzida.
'    no campo de obs base reduzida cfe dec lei ....
'??? Ver c/jogar a fundamentacao p/o item (e a propria nota)

'Todos os produtos devem estar associados a uma "CategoriaICMS"
'    itens pre - cadastrados
'        nao tributados (servicos e outros)
'        isentos
'        aliq padrao
'        aliq 25
'        aliq 7 via red base calculo
'
'Todos os clientes devem estar associados a uma "CategoriaICMS"
'    itens pre - cadastrados
'        consumidor final
'        revenda
'        revenda com subst
'
'(regra do Padua)
'Aliquota Interestadual nao pode ser maior que a interna do destinatario
'
'
'Venda para consumidor final nao contribuinte usar aliquota interna
'    acho que a categoria ICMS da filial cliente resolveria isto.
'
'O ICMS de subst é calculado pelo par (aliquota, margem de lucro) do estado destino
'
'ICMS solidario utiliza os mesmos campos de ICMS subst, a diferenca é que o solidario é resultado
'de um "regime especial" acordado entre uma Empresa e um Estado.
'
'Uma excecao a nivel de cliente é + forte que uma excecao a nivel de produto
'
'Checar tratamento de Ativo fixo, material de consumo, materia prima e estoque
'
'Duvidas:
'
'   talvez deva criar uma pesquisa de excecoes independente de docitem, p/saber como calcular ICMS sobre seguro, desp acessorias, etc
'    como vou tratar se um mesmo produto estiver vinculado a mais de uma categoria e para cada uma houver uma regra de excecao ?
'        acho que a categoria de prod ICMS deveria ser fixa e pre-definida
'            poderiamos assim pre-cadastrar as excecoes.
'        Se o usuario quiser dar um tratamento diferente p/determinado produto deve desvincular a categoria ICMS do produto e vincula-lo a uma nova categoria de ICMS especifica
'
'    Manter flag se calcula ou nao ICMS a nivel de produto ? Acho desnecessario, vide item acima.
'
'    Da mesma maneira devemos pre-definir uma categoria p/ICMS de Filiais-Clientes
'
'        pode ter cache p/agilizar calculos
'        calculo de icms saida a nivel de item:
'        (assim que tiver o produto, cliente e valor pode ser chamada, deve ser chamada a cada alteracao)
'        dados de entrada: (* p/obrigatorios )
'doc: tipo*, cliente*, se calcula icm, estado destino, se ipi na base
'Item:produto*, valor do produto*, valor icms(retido), aliquota, base de calculo, base subst, valor subst
'            dados de saida:
'            todos estes mnemonicos preenchidos, atraves de chamadas a DefinirMnemonico
'            apos o impacto a nivel de item, refletir a nivel de doc
'
'    acrescimos financeiros.etc
'        Art. 5º, § 1º, II
'        O assunto encargos financeiros está previsto no art. 196, § 9°, III, da Constituição Estadual, foi objeto da Lei n° 2085, de 12/02/93, e está regulamentado pela Resolução n° 2263, de 16/03/93.
'        os acrescimos financeiros entram no calculo
'
'Duvidas p / PADUA:
'    descontos:
        'incondicional "razoavel" nao integra, condicional sim (apostila pg 16)
        
'    a base vai incluir a soma da base do itens ou a soma da base * red base ?
'    frete, seguro, desp acess, ipi s/frete:
'        uso aliquota "fora da excecao" (interna, interestadual, etc ?) ou proporcional aos produtos ?
'        aumento base e valor icms subst ?
'        tipo trib varia item a item, como vou obter uma aliquota p/frete ?
'
'
'O valor do Ipi nao integra a base de calculo do ICMS qdo entre contribuintes do ICMS e destinada a comercializacao ou industrializacao.

'de acordo c/o informado pelo Padua (21/05) nao estou + testando se o frete é por conta do emitente ou d o destinatario p/inclui-lo ou nao na base do icms
'
'ELEMENTOS:
'==========
Private mvardFatorValor As Double 'para multiplicar o valor real obtendo o oficial
Private mvarcolEstados As Collection 'colecao de objEstado
            'colocar como chave a UF
Private mvarcolTiposTribICMS As New Collection 'colecao com os tipos de tributacao p/ICMS

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada

Private mvariSimplificado As Integer 'se a filialEmpresa está no Simples do ICMS
Private mvariInscrita As Integer 'se a filialEmpresa está ou nao inscrita no ICMS
Private mvardAliqInterna As Double 'Aliquota interna do estado da filialEmpresa
'Private mvardAliqExportacao As Double 'aliquota exportacao da filialEmpresa
Private mvardAliqImportacao As Double 'aliquota importacao da filialEmpresa
Private mvarcolAliqInterstSaida As Collection 'colecao de objICMSAliqExterna a partir da UF da FilialEmpresa
            'colocar como chave a UF
Private mvarcolAliqInterstEntrada As Collection 'colecao de objICMSAliqExterna com destino a UF da FilialEmpresa
            'colocar como chave a UF
Private mvarcolExcecoesInternas As Collection 'colecao de excecoes ICMS
            'já ordenada, primeiro a nivel de catcliente+catproduto,
            'depois a nivel de catcliente e por ultimo a nivel de catproduto
Private mvarsUFFilialEmpresa As String 'UF da filial empresa "corrente"
Private mvarsUFcolExcecoesAlt As String 'UF para a qual esta preenchida mvarcolExcecoesAlt
Private mvarcolExcecoesAlt As Collection 'colecao de excecoes ICMS p/mvarsUFcolExcecoesAlt
            'já ordenada, primeiro a nivel de catcliente+catproduto,
            'depois a nivel de catcliente e por ultimo a nivel de catproduto
Private mvariSimplesNacional As Integer

Private mvariZerarTributo As Integer
Private mvariZerarTributoST As Integer 'zerar subst tributaria
Private mvariNotaInterna As Integer

'Interface:
'==========

Private Function ClassTributoDriver_Ativar() As Long

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_Ativar

    'carregar dados de todos os estados (p/poder obter a aliquota interna de cada um deles)
    Set mvarcolEstados = New Collection
    lErro = CF("Estados_Le_Todos", mvarcolEstados)
    If lErro <> SUCESSO Then Error 27213
    
    lErro = CF("TiposTribICMS_Le_Todos", mvarcolTiposTribICMS)
    If lErro <> SUCESSO Then Error 27280
    
    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err

        Case 27213, 27280
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153816)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Desativar() As Long
    ClassTributoDriver_Desativar = SUCESSO
End Function

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao do doc como um todos e de seus itens
Dim lErro As Long, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
Dim objVarSerie As New ClassVariavelCalculo
Dim dtDataRef As Date, dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo, objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32022
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 27867
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 27868
    
    lErro = objDoc.ObterVar("SERIE", objVarSerie)
    If lErro <> SUCESSO Then Error 27867
            
    lErro = CF("NFiscal_ObtemFatorValor", mvariFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, mvardFatorValor)
    If lErro <> SUCESSO Then Error 27867
    
    mvariZerarTributo = 0
    mvariZerarTributoST = 0
    If mvariInscrita <> ICMS_INSCRITA Or mvariSimplificado <> 0 Then
        mvariZerarTributo = 1
        If mvariInscrita <> ICMS_INSCRITA Then mvariZerarTributoST = 1
    Else
        If mvariSimplesNacional <> 0 Then
            
            lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
            If lErro <> SUCESSO Then Error 27868
            
            dtDataEmissao = objVarDataEmissao.vValor
            
            lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
            If lErro <> SUCESSO Then Error 27868
            
            dtDataEntrada = objVarDataEntrada.vValor
            
            Select Case objTipoDocInfo.iTipoTrib
                Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
                    dtDataRef = dtDataEntrada
                    
                Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
                    dtDataRef = dtDataEmissao
            End Select
            
            If dtDataRef >= DATA_INICIO_SIMPLES_NACIONAL Then mvariZerarTributo = 1
            
        End If
        
    End If
    
    mvariNotaInterna = IIf(objTipoDocInfo.iEmitente = EMITENTE_EMPRESA, 1, 0)
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Then
    
        lErro = AtualizarImposto_Importacao(objDoc, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27870
    
    Else
    
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_CLI
                If objTipoDocInfo.iCodigo = DOCINFO_NFIEDV Then
                    lErro = AtualizarImposto_SaidaCli(objDoc, iRecalculaAutomatica)
                Else
                    lErro = AtualizarImposto_EntradaCli(objDoc, iRecalculaAutomatica)
                End If
                If lErro <> SUCESSO Then Error 27869
                
            Case TRIB_ENTRADA_FORN
                lErro = AtualizarImposto_EntradaForn(objDoc, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27870
            
            Case TRIB_SAIDA_CLI
                lErro = AtualizarImposto_SaidaCli(objDoc, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27871
            
            Case TRIB_SAIDA_FORN
                lErro = AtualizarImposto_SaidaForn(objDoc, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27872
            
        End Select
    
    End If
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 27867 To 27872, 32022
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153817)

    End Select

    Exit Function

End Function

'Private Function ClassTributoDriver_AtualizarImpostoItem(objDocItem As ClassTributoDocItem) As Long
'
'    ClassTributoDriver_AtualizarImpostoItem = ClassTributoDriver_AtualizarImposto(objDocItem.objDoc)
'
'End Function

'FUNCOES AUXILIARES:
'===================
Private Function AtualizarImposto_SaidaCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_SaidaCli

'    'se a filialEmpresa está inscrita no ICMS e não está no Simples do ICMS
'    If mvariInscrita = ICMS_INSCRITA And mvariSimplificado = 0 Then
    
        lErro = PrepararAtualizacao_SaidaCli(objDoc)
        If lErro <> SUCESSO And lErro <> 27720 Then Error 27291
        
        'se nao houve falta de dados
        If lErro <> 27720 Then
            
''            lErro = objTRBConfig.gobjTributacao.NaturezaOp_GerarPadraoCli(objDoc)
''            If lErro <> SUCESSO Then Error 27659
            
            'se estiver recalculando marcar valores a nivel de doc c/preench auto.
            If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
            
            Call objDoc.ObterColDocItens(colItens)
            
            'percorrer os itens atualizando-os
            For Each objDocItem In colItens
                lErro = AtualizarItem_SaidaCli(objDocItem, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27292
            Next
            
            'atualizar frete, seguro, etc
            lErro = AtualizarDocCompl_SaidaCli(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27479
            
            'atualizar o docto como um todo
            lErro = AtualizarDoc(objDoc)
            If lErro <> SUCESSO Then Error 27293
        
        End If
        
'    End If
    
    AtualizarImposto_SaidaCli = SUCESSO

    Exit Function

Erro_AtualizarImposto_SaidaCli:

    AtualizarImposto_SaidaCli = Err

    Select Case Err

        Case 27291, 27292, 27293, 27479, 27659
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153818)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_EntradaCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_EntradaCli

    'se a filialEmpresa está inscrita no ICMS e não está no Simples
'    If mvariZerarTributo = 0 Then

        lErro = PrepararAtualizacao_EntradaCli(objDoc)
        If lErro <> SUCESSO And lErro <> 27882 Then Error 27957

        'se nao houve falta de dados
        If lErro <> 27882 Then

''            lErro = objTRBConfig.gobjTributacao.NaturezaOp_GerarPadraoCli(objDoc)
''            If lErro <> SUCESSO Then Error 27958

            'se estiver recalculando marcar valores a nivel de doc c/preench auto.
            If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
    
            Call objDoc.ObterColDocItens(colItens)

            'percorrer os itens atualizando-os
            For Each objDocItem In colItens
                lErro = AtualizarItem_EntradaCli(objDocItem, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27959
            Next

            'atualizar frete, seguro, etc
            lErro = AtualizarDocCompl_EntCli(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27960

            'atualizar o docto como um todo
            lErro = AtualizarDoc(objDoc)
            If lErro <> SUCESSO Then Error 27961

        End If

'    End If
    
    AtualizarImposto_EntradaCli = SUCESSO

    Exit Function

Erro_AtualizarImposto_EntradaCli:

    AtualizarImposto_EntradaCli = Err

    Select Case Err

        Case 27957 To 27961
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153819)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_SaidaForn(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_SaidaForn

    'se a filialEmpresa está inscrita no ICMS e não está no Simples
'    If mvariZerarTributo = 0 Then

        lErro = PrepararAtualizacao_Forn(objDoc)
        If lErro <> SUCESSO And lErro <> 27973 Then Error 22723

        'se nao houve falta de dados
        If lErro <> 27973 Then

''            lErro = objTRBConfig.gobjTributacao.NaturezaOp_GerarPadraoForn(objDoc)
''            If lErro <> SUCESSO Then Error 22724

            'se estiver recalculando marcar valores a nivel de doc c/preench auto.
            If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
            
            Call objDoc.ObterColDocItens(colItens)

            'percorrer os itens atualizando-os
            For Each objDocItem In colItens
                lErro = AtualizarItem_SaidaForn(objDocItem, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 22725
            Next

            'atualizar frete, seguro, etc
            lErro = AtualizarDocCompl_SaidaForn(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22726

            'atualizar o docto como um todo
            lErro = AtualizarDoc(objDoc)
            If lErro <> SUCESSO Then Error 22727

        End If

'    End If
    
    AtualizarImposto_SaidaForn = SUCESSO

    Exit Function

Erro_AtualizarImposto_SaidaForn:

    AtualizarImposto_SaidaForn = Err

    Select Case Err

        Case 22723 To 22727
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153820)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_EntradaForn(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_EntradaForn

''    'se a filialEmpresa está inscrita no ICMS e não está no Simples do ICMS
''    If mvariInscrita = ICMS_INSCRITA  And mvariSimplificado = 0 Then
''
        lErro = PrepararAtualizacao_Forn(objDoc)
        If lErro <> SUCESSO And lErro <> 27973 Then Error 22718

        'se nao houve falta de dados
        If lErro <> 27973 Then

''            lErro = objTRBConfig.gobjTributacao.NaturezaOp_GerarPadraoForn(objDoc)
''            If lErro <> SUCESSO Then Error 22719

            'se estiver recalculando marcar valores a nivel de doc c/preench auto.
            If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
            
            Call objDoc.ObterColDocItens(colItens)

            'percorrer os itens atualizando-os
            For Each objDocItem In colItens
                lErro = AtualizarItem_EntradaForn(objDocItem, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 22720
            Next

            'atualizar frete, seguro, etc
            lErro = AtualizarDocCompl_EntForn(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22721

            'atualizar o docto como um todo
            lErro = AtualizarDoc(objDoc)
            If lErro <> SUCESSO Then Error 22722

        End If

''    End If
    
    AtualizarImposto_EntradaForn = SUCESSO

    Exit Function

Erro_AtualizarImposto_EntradaForn:

    AtualizarImposto_EntradaForn = Err

    Select Case Err

        Case 22718 To 22722
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153821)

    End Select

    Exit Function

End Function

Private Function PrepararAtualizacao_SaidaCli(objDoc As ClassTributoDoc) As Long
Dim lErro As Long, sUFDestino As String
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarUFDestino As New ClassVariavelCalculo
On Error GoTo Erro_PrepararAtualizacao_SaidaCli

''    sUFDestino = ""
''
''    'obter codigo do cliente
''    lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
''    If lErro <> SUCESSO Then Error 27717
''
''    'obter identificador da filial do cliente
''    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
''    If lErro <> SUCESSO Then Error 27718
''
''    'se o cliente e a filial já foram informados
''    If (objVarFilialCliente.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.iStatus <> VAR_PREENCH_VAZIO) Then
''
''        lErro = objDoc.ObterVar("CLIENTE_UF", objVarUFDestino)
''        If lErro <> SUCESSO Then Error 27719
''
''        sUFDestino = objVarUFDestino.vValor
''
''    End If
''
''    'faltam dados
''    If sUFDestino = "" Then Error 27720

    PrepararAtualizacao_SaidaCli = SUCESSO

    Exit Function

Erro_PrepararAtualizacao_SaidaCli:

    PrepararAtualizacao_SaidaCli = Err

    Select Case Err
    
        Case 27717 To 27719
        
        Case 27720 'faltam dados p/calcular tributacao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153822)

    End Select

    Exit Function

End Function

''Private Function AtualizarDoc_ObterAliquotaBasica(objDoc As ClassTributoDoc, dAliquota As Double) As Long
'''obtem a aliquota a ser usada p/ICMS sobre frete, seguro e despesas acessorias
'''??? estou fazendo a media ponderada por base de calculo nos itens em que incide ICMS
''Dim lErro As Long, dBase As Double, dAcum As Double, dAliquotaComum As Double
''Dim colItens As Collection
''Dim objDocItem As ClassTributoDocItem, objVarICMSBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo
''On Error GoTo Erro_AtualizarDoc_ObterAliquotaBasica
''
''    Call objDoc.ObterColDocItens(colItens)
''
''    dBase = 0
''    dAcum = 0
''    dAliquotaComum = -1
''
''    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
''    For Each objDocItem In colItens
''
''        'obter do item ICMS_BASE e ICMS_ALIQUOTA
''
''        lErro = objDocItem.ObterVar("ICMS_BASE", objVarICMSBase)
''        If lErro <> SUCESSO Then Error 27329
''
''        lErro = objDocItem.ObterVar("ICMS_ALIQUOTA", objVarAliquota)
''        If lErro <> SUCESSO Then Error 27330
''
''        If objVarICMSBase.vValor <> 0 Then
''
''            dBase = dBase + objVarICMSBase.vValor
''            dAcum = dAcum + (objVarICMSBase.vValor * objVarAliquota.vValor)
''
''            'na 1a passagem
''            If dAliquotaComum = -1 Then
''                dAliquotaComum = objVarAliquota.vValor
''            Else
''                'se nao existem aliquotas diferentes dentro da nota
''                If dAliquotaComum <> -2 Then
''                    'indicar que existem aliquotas diferentes dentro da nota, se for o caso
''                    If objVarAliquota.vValor <> dAliquotaComum Then dAliquotaComum = -2
''                End If
''            End If
''
''        End If
''
''    Next
''
''    If dAliquotaComum <> -1 Then
''        If dAliquotaComum = -2 And dBase <> 0 Then
''            dAliquota = dAcum / dBase
''        Else
''            dAliquota = dAliquotaComum
''        End If
''    Else
''        dAliquota = 0
''    End If
''
''    AtualizarDoc_ObterAliquotaBasica = SUCESSO
''
''    Exit Function
''
''Erro_AtualizarDoc_ObterAliquotaBasica:
''
''    AtualizarDoc_ObterAliquotaBasica = Err
''
''    Select Case Err
''
''        Case 27329, 27330
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153823)
''
''    End Select
''
''    Exit Function
''
''End Function

''Private Function AtualizarDoc_SaidaCli(objDoc As ClassTributoDoc) As Long
'''atualiza base de ICMS, valor de ICMS, base ICMS subst, valor ICMS Subst
'''    leva em conta dados dos itens + dados de frete, seg, desp acess.
''Dim colItens As Collection
''Dim objDocItem As ClassTributoDocItem, dAliquota As Double
''Dim dValor As Double, dValorSubst As Double, dBase As Double, dBaseSubst As Double, dBaseReduzida As Double, dBaseSubstReduzida As Double, dValorCredito As Double
''Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
''Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
''Dim objVarFreteValor As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
''Dim objVarSeguroValor As New ClassVariavelCalculo, objVarDespAcesValor As New ClassVariavelCalculo
''Dim lErro As Long, dPercRedBase As Double
''On Error GoTo Erro_AtualizarDoc_SaidaCli
''
''    dValor = 0
''    dValorSubst = 0
''    dBase = 0
''    dBaseReduzida = 0
''    dBaseSubst = 0
''    dBaseSubstReduzida = 0
''    dValorCredito = 0
''
''    Call objDoc.ObterColDocItens(colItens)
''
''    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
''    For Each objDocItem In colItens
''
''        'obter do item ICMS_BASE, ICMS_VALOR, ICMS_BASE_SUBST, ICMS_VALOR_SUBST
''
''        lErro = objDocItem.ObterVar("ICMS_BASE", objVarICMSBase)
''        If lErro <> SUCESSO Then Error 27226
''
''        lErro = objDocItem.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
''        If lErro <> SUCESSO Then Error 27442
''
''        dPercRedBase = objVarPercRedBase.vValor
''
''        lErro = objDocItem.ObterVar("ICMS_VALOR", objVarICMSValor)
''        If lErro <> SUCESSO Then Error 27227
''
''        lErro = objDocItem.ObterVar("ICMS_BASE_SUBST", objVarICMSBaseSubst)
''        If lErro <> SUCESSO Then Error 27228
''
''        lErro = objDocItem.ObterVar("ICMS_VALOR_SUBST", objVarICMSValorSubst)
''        If lErro <> SUCESSO Then Error 27229
''
''        dBase = dBase + objVarICMSBase.vValor
''        dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBase))
''        dBaseSubst = dBaseSubst + objVarICMSBaseSubst.vValor
''        dBaseSubstReduzida = dBaseSubstReduzida + (objVarICMSBaseSubst.vValor * (1 - dPercRedBase))
''        dValor = dValor + objVarICMSValor.vValor
''        dValorSubst = dValorSubst + objVarICMSValorSubst.vValor
''
''    Next
''
''    'tratar dados globais (frete, desp acessorias,...) inclusive vendo o impacto do IPI sobre o frete,... na base de calculo do ICMS
''    lErro = AtualizarDoc_ObterDadosCompl(objDoc, dValor, dValorSubst, dBase, dBaseReduzida, dBaseSubst, dBaseSubstReduzida, dValorCredito)
''    If lErro <> SUCESSO Then Error 27481
''
''    lErro = AtualizarDoc_GuardarValores(objDoc, dBaseReduzida, dValor, dBaseSubstReduzida, dValorSubst, dValorCredito)
''    If lErro <> SUCESSO Then Error 27878
''
''    AtualizarDoc_SaidaCli = SUCESSO
''
''    Exit Function
''
''Erro_AtualizarDoc_SaidaCli:
''
''    AtualizarDoc_SaidaCli = Err
''
''    Select Case Err
''
''        Case 27226 To 27234, 27331, 27442, 27481, 27878
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153824)
''
''    End Select
''
''    Exit Function
''
''End Function

Private Function ObterAliquotaBasicaSaidaCli(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFDestino As String) As Long
'preenche a aliquota
Dim lErro As Long
Dim objFilCliInscEst As New ClassVariavelCalculo
On Error GoTo Erro_ObterAliquotaBasicaSaidaCli

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFDestino = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            
        Else
          
            'se for exportacao
            If sUFDestino = "EX" Then
            
                'setar aliquota igual a zero ( exportacao nao tem ICMS )
                dAliquota = 0
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
                If lErro <> SUCESSO Then Error 27284
                
                'se filial cliente nao inscrita ICMS
                If objFilCliInscEst.vValor = "" Then
                    
                    'usar aliquota interna
                    dAliquota = mvardAliqInterna
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstSaida_Pesquisar(sUFDestino, dAliquota)
                    If lErro <> SUCESSO Then Error 27285
                
                End If
            End If
        End If
    Else
        dAliquota = 0
    End If

    ObterAliquotaBasicaSaidaCli = SUCESSO

    Exit Function

Erro_ObterAliquotaBasicaSaidaCli:

    ObterAliquotaBasicaSaidaCli = Err

    Select Case Err

        Case 27284, 27285
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153825)

    End Select

    Exit Function

End Function


Private Function AtualizarItem_ObterValoresAtuais(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer) As Long
'obtem os valores atuais da variaveis associadas a um item
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_ObterValoresAtuais

    lErro = objDocItem.ObterVar("ICMS_TIPO", objVarTipo)
    If lErro <> SUCESSO Then Error 27298
    If iRecalculaAutomatica = 1 Then objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("ICMS_BASE", objVarBaseCalculo)
    If lErro <> SUCESSO Then Error 27299
    If iRecalculaAutomatica = 1 Then objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
    If lErro <> SUCESSO Then Error 27300
    If iRecalculaAutomatica = 1 Then objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("ICMS_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27301
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("ICMS_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27302
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("ICMS_BASE_SUBST", objVarBaseSubst)
    If lErro <> SUCESSO Then Error 27303
    If iRecalculaAutomatica = 1 Then objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("ICMS_RED_BASE_SUBST", objVarPercRedBaseSubst)
    If lErro <> SUCESSO Then Error 27300
    If iRecalculaAutomatica = 1 Then objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("ICMS_MARG_SUBST", objVarMargSubst)
    If lErro <> SUCESSO Then Error 27304
    If iRecalculaAutomatica = 1 Then objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_SUBST", objVarAliqSubst)
    If lErro <> SUCESSO Then Error 27305
    If iRecalculaAutomatica = 1 Then objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("ICMS_VALOR_SUBST", objVarValorSubst)
    If lErro <> SUCESSO Then Error 27306
    If iRecalculaAutomatica = 1 Then objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("ICMS_VALOR_CREDITO", objVarValorCredito)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarItem_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarItem_ObterValoresAtuais:

    AtualizarItem_ObterValoresAtuais = Err

    Select Case Err

        Case 27298 To 27306, 27876
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153826)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_GuardarValores(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, _
        dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double, dPercRedBaseSubst As Double, dValorCredito As Double) As Integer
'atualiza as variaveis que tiveram seus valores recalculados pelo sistema
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_GuardarValores

    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTipo.vValor = iTipo
        objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTipo)
        If lErro <> SUCESSO Then Error 27307
        
    End If
    
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseCalculo.vValor = dBaseCalculo
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarBaseCalculo)
        If lErro <> SUCESSO Then Error 27308
        
    End If
    
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarPercRedBase.vValor = dPercRedBase
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27309
        
    End If
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliquota.vValor = dAliquota
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then Error 27310
        
    End If
                
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then Error 27311
        
    End If
                
    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseSubst.vValor = dBaseSubst
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarBaseSubst)
        If lErro <> SUCESSO Then Error 27312
        
    End If
                
    If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then

        objVarMargSubst.vValor = dMargSubst
        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarMargSubst)
        If lErro <> SUCESSO Then Error 27313

    End If
                
    If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliqSubst.vValor = dAliqSubst
        objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarAliqSubst)
        If lErro <> SUCESSO Then Error 27314
        
    End If
                
    If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorSubst.vValor = dValorSubst
        objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValorSubst)
        If lErro <> SUCESSO Then Error 27315
        
    End If
            
    If objVarPercRedBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarPercRedBaseSubst.vValor = dPercRedBaseSubst
        objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarPercRedBaseSubst)
        If lErro <> SUCESSO Then Error 27309
        
    End If
    
    If objVarValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorCredito.vValor = dValorCredito
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then Error 27874
        
    End If
            
    AtualizarItem_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores:

    AtualizarItem_GuardarValores = Err

    Select Case Err

        Case 27307 To 27315, 27874
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153827)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarExcecao(objExcecao As ClassICMSExcecao, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, objVarMargSubst As ClassVariavelCalculo, dMargSubst As Double, dPercRedBaseSubst As Double, ByVal bSemSubstTributaria, Optional iUsaPauta As Integer, Optional dValorPauta As Double) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarItem_TratarExcecao

    'se ICMS_TIPO nao tiver sido preenchido manualmente
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objExcecao.iTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27316
    
    'se ICMS_ALIQUOTA nao tiver sido preenchido manualmente
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        'se ICMS_TIPO permite aliquota <> 0
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            'ICMS_ALIQUOTA = aliquota da excecao
            dAliquota = objExcecao.dAliquota
        Else
            dAliquota = 0
        End If
    Else
        dAliquota = objVarAliquota.vValor
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = objExcecao.dPercRedBaseCalculo
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
    
    'se ICMS_TIPO permite subst
    If objTipoTribICMS.iPermiteMargLucro And (bSemSubstTributaria = False) Then
        'ICMS_MARG_SUBST = % da excecao
        dMargSubst = objExcecao.dPercMargemLucro
        dPercRedBaseSubst = objExcecao.dPercRedBaseCalculoSubst
        dValorPauta = objExcecao.dValorPauta
        iUsaPauta = objExcecao.iUsaPauta
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
        dValorPauta = 0
        iUsaPauta = 0
    End If
     
    AtualizarItem_TratarExcecao = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarExcecao:

    AtualizarItem_TratarExcecao = Err

    Select Case Err

        Case 27316
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153828)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFDestino As String) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoSaidaCli
        
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objTipoTrib.iICMSTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27317
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaSaidaCli(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFDestino)
            If lErro <> SUCESSO Then Error 27318
        Else
            dAliquota = 0
        End If
    Else
    
        dAliquota = objVarAliquota.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecaoSaidaCli = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoSaidaCli:

    AtualizarItem_TratarNaoExcecaoSaidaCli = Err

    Select Case Err

        Case 27317, 27318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153829)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_SaidaCli(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long, objExcecao As ClassICMSExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double
Dim objVarValorDescGlobal As New ClassVariavelCalculo, dValorPauta As Double, iUsaPauta As Integer, dQtdeSubst As Double, objVarQtdeSubst As New ClassVariavelCalculo
Dim objVarValorDespGlobal As New ClassVariavelCalculo, bExcecaoComSubstTrib As Boolean
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, dAliquotaInterestadual As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFDestino As String, dPercRedBaseSubst As Double
Dim objVarUFDestino As New ClassVariavelCalculo, objFilCliInscEst As New ClassVariavelCalculo, dBaseSubstReduzida As Double
Dim dBaseICMSProprioSTSimples As Double, objVarPercRedBaseSubst As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_SaidaCli
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dBaseICMSProprioSTSimples = 0
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dAliqSubst = 0
    dMargSubst = 0
    dValorSubst = 0
    dValorCredito = 0
    dPercRedBaseSubst = 0
    iUsaPauta = 0
    dValorPauta = 0
    bExcecaoComSubstTrib = False
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27319
    
    'se a filialEmpresa nao está inscrita no ICMS ou está no Simples
    If mvariZerarTributo <> 0 And mvariZerarTributoST <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27320
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27373
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'obter UF Destino
    lErro = objDocItem.objDoc.ObterVar("CLIENTE_UF", objVarUFDestino)
    If lErro <> SUCESSO Then Error 27721
    
    sUFDestino = objVarUFDestino.vValor
    If sUFDestino = "" Then sUFDestino = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
    If lErro <> SUCESSO Then Error 27722
    
    'se filial cliente inscrita ICMS e interestadual
    If objFilCliInscEst.vValor <> "" And sUFDestino <> mvarsUFFilialEmpresa And sUFDestino <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFDestino)
        If lErro <> SUCESSO Then Error 27723
        
        dAliquotaInterestadual = dAliquota
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, sUFDestino, mvarsUFFilialEmpresa)
        If lErro <> SUCESSO Then Error 27724
            
        If Not (objExcecao Is Nothing) Then
'            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
'            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dValorPauta <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta)
                If lErro <> SUCESSO Then Error 27322
                
'                If dMargSubst <> 0 Or dValorPauta <> 0 Then dAliquota = dAliquotaInterestadual
                dAliquota = dAliquotaInterestadual
                
                If TipoICMSComSubstTrib(objTipoTribICMS.iTipo) Then bExcecaoComSubstTrib = True
'            End If
        End If
    
    Else
        If sUFDestino <> "EX" Then
            'tratar como operacao interna
        
            'tentar achar excecao
            lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, "")
            If lErro <> SUCESSO Then Error 27321
            
            'se achou excecao
            If (Not (objExcecao Is Nothing)) Then
            
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, objFilCliInscEst.vValor = "", iUsaPauta, dValorPauta)
                If lErro <> SUCESSO Then Error 27322
                                
                If TipoICMSComSubstTrib(objTipoTribICMS.iTipo) And objFilCliInscEst.vValor <> "" Then bExcecaoComSubstTrib = True
            
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa)
                If lErro <> SUCESSO Then Error 27323
            
            End If
            
        Else  'exportacao
        
            lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFDestino)
            If lErro <> SUCESSO Then Error 27727
        
        End If
        
    End If

    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
        
            lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, True, True, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27324

            'se inclui o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseCalculo = dBaseCalculo + objVar.vValor
            End If
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.0000"))
    
    Else
        dValor = objVarValor.vValor
    End If

    If objTipoTrib.iICMSCredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        
        If bExcecaoComSubstTrib Then
            
            If iUsaPauta = 0 Then
            
                dBaseSubst = dBaseCalculo
                
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = dBaseSubst + objVar.vValor
                
                End If
                
                'incluir um valor proporcional de frete, seguro e outras despesas na base do item
                lErro = objDocItem.ObterVar("PRODUTO_DESP_GLOBAL", objVarValorDespGlobal)
                If lErro <> SUCESSO Then Error 27324

                dBaseSubst = dBaseSubst + Round(objVarValorDespGlobal.vValor * mvardFatorValor, 2)
                
                dBaseSubst = dBaseSubst * (1 + dMargSubst)
                
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
                
            Else
        
                'obter qtde do item de objdocitem
                lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = CDbl(Format(objVar.vValor * dValorPauta, "0.0000"))
                dBaseSubstReduzida = dBaseSubst
                
            End If
            
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst
    End If
    
    'se é para zerar a parte propria (mantendo apenas a subst tributaria)
    dBaseICMSProprioSTSimples = 0
    If mvariZerarTributo <> 0 Then
        If mvariSimplesNacional <> 0 Then
            'incluir um valor proporcional de frete, seguro e outras despesas na base do item
            lErro = objDocItem.ObterVar("PRODUTO_DESP_GLOBAL", objVarValorDespGlobal)
            If lErro <> SUCESSO Then Error 27324
            dBaseICMSProprioSTSimples = dBaseCalculo + Round(objVarValorDespGlobal.vValor * mvardFatorValor, 2)
        End If
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dValorCredito = 0
    End If
    
    If dBaseSubst <> 0 Then
        
        'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
        If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objExcecao Is Nothing Then
                dAliqSubst = 0
            Else
                dAliqSubst = objExcecao.dAliquota
            End If
        Else
            dAliqSubst = objVarAliqSubst.vValor
        End If
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If mvariSimplesNacional = 0 Then
                dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
            Else
                dValorSubst = dBaseSubstReduzida * dAliqSubst - CDbl(Format(dBaseICMSProprioSTSimples * 0.07, "0.0000"))
            End If
            If dValorSubst < 0 Then dValorSubst = 0
            dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dPercRedBaseSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27328
    
    AtualizarItem_SaidaCli = SUCESSO

    Exit Function

Erro_AtualizarItem_SaidaCli:

    AtualizarItem_SaidaCli = Err

    Select Case Err

        Case 27319 To 27328, 27373, 27687 To 27690, 27721 To 27724, 27727, 27765
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153830)

    End Select

    Exit Function

End Function

Private Function ExcecaoCli_Pesquisar(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, sUFDestino As String, sUFOrigem As String) As Long
'uma excecao a nivel de cliente tem + prioridade que excecoes a nivel de produto apenas
'e uma excecao a nivel de cliente+produto tem a maior prioridade de todas.
'A colecao já está ordenada de forma que pararei na 1a excecao que encontrar
'retorna Nothing se nao achar nenhuma excecao adequada
Dim lErro As Long, objICMSExcecao As ClassICMSExcecao, iAchou As Integer
Dim objVarUFDestino As New ClassVariavelCalculo, objVarFilCliColCateg As New ClassVariavelCalculo
Dim objVarProdColCateg As New ClassVariavelCalculo
Dim colExcecoes As New Collection

On Error GoTo Erro_ExcecaoCli_Pesquisar

    'obter a colecao de categorias da filialcliente
    lErro = objDocItem.objDoc.ObterVar("I_FILIAL_CLIENTE_COL_CATEG", objVarFilCliColCateg)
    If lErro <> SUCESSO Then Error 27220
    
    'obter a colecao de categorias do produto
    lErro = objDocItem.ObterVar("I_PRODUTO_COL_CATEG", objVarProdColCateg)
    If lErro <> SUCESSO Then Error 27221
    
    If sUFDestino = mvarsUFFilialEmpresa Then
        Set colExcecoes = mvarcolExcecoesInternas
    Else
        If mvarsUFcolExcecoesAlt <> sUFDestino Then
            'carregar colecao de excecoes ICMS para o estado destino
            Set mvarcolExcecoesAlt = New Collection
            mvarsUFcolExcecoesAlt = ""
            lErro = CF("ICMSExcecoes_Le_Estado", sUFDestino, mvarcolExcecoesAlt)
            If lErro <> SUCESSO Then Error 27725
            mvarsUFcolExcecoesAlt = sUFDestino
        End If
        Set colExcecoes = mvarcolExcecoesAlt
    End If
    
    iAchou = 0
    
    'percorrer a colecao de excecoes
    For Each objICMSExcecao In colExcecoes
    
        If sUFOrigem = "" Or Trim(objICMSExcecao.sEstadoOrigem) = "" Or objICMSExcecao.sEstadoOrigem = sUFOrigem Then
            
            If objICMSExcecao.sEstadoDestino = sUFDestino And objICMSExcecao.sCategoriaFornecedor = "" Then
            
                'se a excecao se aplica a uma categoria de clientes para uma categoria de produtos
                If objICMSExcecao.sCategoriaCliente <> "" And objICMSExcecao.sCategoriaProduto <> "" Then
                
                    'verifica se a filial cliente se enquadra
                    lErro = CategoriaCliente_Pesquisar(objICMSExcecao.sCategoriaCliente, objICMSExcecao.sCategoriaClienteItem, objVarFilCliColCateg.vValor)
                    If lErro <> SUCESSO And lErro <> 27215 Then Error 27224
                    If lErro = SUCESSO Then
                    
                        'verifica se o produto se enquadra
                        lErro = CategoriaProduto_Pesquisar(objICMSExcecao.sCategoriaProduto, objICMSExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27216 Then Error 27225
                                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
                        
                    End If
                    
                Else
                
                    'se a excecao se aplica a uma categoria de clientes
                    If objICMSExcecao.sCategoriaCliente <> "" Then
                        'verifica se a filial cliente se enquadra
                        lErro = CategoriaCliente_Pesquisar(objICMSExcecao.sCategoriaCliente, objICMSExcecao.sCategoriaClienteItem, objVarFilCliColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27215 Then Error 27222
                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
            
                    Else
                    'se a excecao se aplica a uma categoria de produtos
                    'If objICMSExcecao.sCategoriaProduto <> "" Then
                        'verifica se o produto se enquadra
                        lErro = CategoriaProduto_Pesquisar(objICMSExcecao.sCategoriaProduto, objICMSExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27216 Then Error 27223
                                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
                    End If
                End If
            End If
    
        End If
    
    Next
    
    If iAchou <> 1 Then Set objExcecao = Nothing
    
    ExcecaoCli_Pesquisar = SUCESSO

    Exit Function

Erro_ExcecaoCli_Pesquisar:

    ExcecaoCli_Pesquisar = Err

    Select Case Err
        
        Case 27217 To 27225, 27725

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153831)

    End Select

    Exit Function

End Function

Private Function AliqInterna_Pesquisar(sSigla As String, dAliquota As Double) As Long
'retorna a aliquota interna de uma UF
Dim lErro As Long, objEstado As ClassEstado
On Error GoTo Erro_AliqInterna_Pesquisar

    Set objEstado = mvarcolEstados.Item(sSigla)
    
    dAliquota = objEstado.dICMSAliquotaInterna
    
    AliqInterna_Pesquisar = SUCESSO

    Exit Function

Erro_AliqInterna_Pesquisar:

    AliqInterna_Pesquisar = Err

    Select Case Err

        Case 5 ' UF nao encontrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ICMS_ALIQ_INT_INEXISTENTE", Err, Error)
            dAliquota = 0
            AliqInterna_Pesquisar = 27208
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153832)

    End Select

    Exit Function

End Function

Private Function AliqInterstSaida_Pesquisar(sSiglaDestino As String, dAliquota As Double) As Long
'retorna a aliquota interestadual partindo do estado da filial empresa corrente
Dim lErro As Long, objICMSAliqExterna As ClassICMSAliqExterna
On Error GoTo Erro_AliqInterstSaida_Pesquisar
    
    Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sSiglaDestino)
    
    dAliquota = objICMSAliqExterna.dAliquota
    
    AliqInterstSaida_Pesquisar = SUCESSO

    Exit Function

Erro_AliqInterstSaida_Pesquisar:

    AliqInterstSaida_Pesquisar = Err

    Select Case Err

        Case 5 ' UF nao encontrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ICMS_ALIQ_INTERSTADUAL_INEXISTENTE", Err, Error)
            dAliquota = 0
            AliqInterstSaida_Pesquisar = 27209
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153833)

    End Select

    Exit Function

End Function

Private Function CategoriaCliente_Pesquisar(sCategoriaCliente As String, sCategoriaClienteItem As String, colCateg As Collection) As Long
'verifica se o par (sCategoriaCliente, sCategoriaClienteItem está na colecao colCateg
'retorna SUCESSO se achar, 27215 se nao achar ou outro erro
Dim lErro As Long, objFilialCliCategoria As ClassFilialCliCategoria, iAchou As Integer
On Error GoTo Erro_CategoriaCliente_Pesquisar

    iAchou = 0
    
    For Each objFilialCliCategoria In colCateg
    
        If objFilialCliCategoria.sCategoria = sCategoriaCliente And objFilialCliCategoria.sItem = sCategoriaClienteItem Then
        
            iAchou = 1
            Exit For
            
        End If
        
    Next

    If iAchou = 0 Then Error 27215
    
    CategoriaCliente_Pesquisar = SUCESSO

    Exit Function

Erro_CategoriaCliente_Pesquisar:

    CategoriaCliente_Pesquisar = Err

    Select Case Err

        Case 27215 'nao encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153834)

    End Select

    Exit Function

End Function

Private Function CategoriaProduto_Pesquisar(sCategoriaProduto As String, sCategoriaProdutoItem As String, colCateg As Collection) As Long
'verifica se o par (sCategoriaProduto, sCategoriaProdutoItem está na colecao colCateg
'retorna SUCESSO se achar, 27216 se nao achar ou outro erro
Dim lErro As Long, objProdutoCategoria As ClassProdutoCategoria, iAchou As Integer
On Error GoTo Erro_CategoriaProduto_Pesquisar

    iAchou = 0
    
    For Each objProdutoCategoria In colCateg
    
        If objProdutoCategoria.sCategoria = sCategoriaProduto And objProdutoCategoria.sItem = sCategoriaProdutoItem Then
        
            iAchou = 1
            Exit For
            
        End If
        
    Next

    If iAchou = 0 Then Error 27216
    
    CategoriaProduto_Pesquisar = SUCESSO

    Exit Function

Erro_CategoriaProduto_Pesquisar:

    CategoriaProduto_Pesquisar = Err

    Select Case Err

        Case 27216 'nao encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153835)

    End Select

    Exit Function

End Function

Private Function TipoICMS_Obter(objTipoTribICMS As ClassTipoTribICMS) As Long
'Preenche objTipoTribICMS de acordo com o objTipoTribICMS.iTipo que vier preenchido
Dim lErro As Long, iAchou As Integer
Dim objTipoTribICMSTemp As ClassTipoTribICMS
On Error GoTo Erro_TipoICMS_Obter

    iAchou = 0
    
    For Each objTipoTribICMSTemp In mvarcolTiposTribICMS
        If objTipoTribICMS.iTipo = objTipoTribICMSTemp.iTipo Then
            iAchou = 1
            Call objTipoTribICMS.Copiar(objTipoTribICMSTemp)
            Exit For
        End If
    Next

    If iAchou = 0 Then Error 27297
    
    TipoICMS_Obter = SUCESSO

    Exit Function

Erro_TipoICMS_Obter:

    TipoICMS_Obter = Err

    Select Case Err

        Case 27297
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_ICMS_INEXISTENTE", Err, objTipoTribICMS.iTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153836)

    End Select

    Exit Function

End Function

Private Function TipoICMSPermiteAliq(iTipoICMS As Integer, iPermiteAliq As Integer) As Long
'retorna em iPermiteAliq se o tipo de ICMS informado pode ter aliquota diferente de zero
Dim lErro As Long, iAchou As Integer
Dim objTipoTribICMS As ClassTipoTribICMS
On Error GoTo Erro_TipoICMSPermiteAliq

    iAchou = 0
    
    For Each objTipoTribICMS In mvarcolTiposTribICMS
        If objTipoTribICMS.iTipo = iTipoICMS Then
            iAchou = 1
            iPermiteAliq = objTipoTribICMS.iPermiteAliquota
            Exit For
        End If
    Next

    If iAchou = 0 Then Error 27281
    
    TipoICMSPermiteAliq = SUCESSO

    Exit Function

Erro_TipoICMSPermiteAliq:

    TipoICMSPermiteAliq = Err

    Select Case Err

        Case 27281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_ICMS_INEXISTENTE", Err, iTipoICMS)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153837)

    End Select

    Exit Function

End Function

''Function TiposTribICMS_Le_Todos(colTiposTribICMS As Collection) As Long
'''preenche colTiposTribICMS com objetos ClassTipoTribICMS a partir da tabela TiposTribICMS
''Dim lErro As Long, lComando As Long
''Dim objTipoTribICMS As ClassTipoTribICMS
''Dim iPermiteAliquota As Integer, iPermiteMargLucro As Integer, iPermiteReducaoBase As Integer, iTipo As Integer, sDescricao As String
''On Error GoTo Erro_TiposTribICMS_Le_Todos
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 27276
''
''    sDescricao = String(STRING_TIPO_ICMS_DESCRICAO, 0)
''
''    lErro = Comando_Executar(lComando, "SELECT Tipo, Descricao, PermiteAliquota, PermiteMargLucro, PermiteReducaoBase FROM TiposTribICMS", _
''        iTipo, sDescricao, iPermiteAliquota, iPermiteMargLucro, iPermiteReducaoBase)
''    If lErro <> AD_SQL_SUCESSO Then Error 27277
''
''    lErro = Comando_BuscarProximo(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27278
''
''    Do While lErro <> AD_SQL_SEM_DADOS
''
''        Set objTipoTribICMS = New ClassTipoTribICMS
''
''        objTipoTribICMS.iPermiteAliquota = iPermiteAliquota
''        objTipoTribICMS.iPermiteMargLucro = iPermiteMargLucro
''        objTipoTribICMS.iPermiteReducaoBase = iPermiteReducaoBase
''        objTipoTribICMS.iTipo = iTipo
''        objTipoTribICMS.sDescricao = sDescricao
''
''        colTiposTribICMS.Add objTipoTribICMS
''
''        lErro = Comando_BuscarProximo(lComando)
''        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27279
''
''    Loop
''
''    lErro = Comando_Fechar(lComando)
''
''    TiposTribICMS_Le_Todos = SUCESSO
''
''    Exit Function
''
''Erro_TiposTribICMS_Le_Todos:
''
''    TiposTribICMS_Le_Todos = Err
''
''    Select Case Err
''
''        Case 27276
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 27277, 27278, 27279
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOS_TRIB_ICMS", Err)
''
''        Case Else
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153838)
''
''    End Select
''
''    Exit Function
''
''End Function


Private Function AtualizarDoc_ObterDadosCompl(objDoc As ClassTributoDoc, dValor As Double, dValorSubst As Double, dBase As Double, dBaseReduzida As Double, dBaseSubst As Double, dBaseSubstReduzida As Double, dValorCredito As Double) As Long
'tratar dados globais (frete, desp acessorias,...) inclusive vendo o impacto do IPI sobre o frete,... na base de calculo do ICMS
Dim objTributoDocCompl As ClassTributoDocCompl, colDocCompl As New Collection, objVarItemTipo As New ClassVariavelCalculo
Dim objVarICMSBase As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
Dim objVarICMSValor As New ClassVariavelCalculo, objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim lErro As Long, dPercRedBase As Double, objVarICMSValorCredito As New ClassVariavelCalculo
On Error GoTo Erro_AtualizarDoc_ObterDadosCompl

    lErro = objDoc.ObterColDocCompl(colDocCompl)
    If lErro <> SUCESSO Then Error 27482

    For Each objTributoDocCompl In colDocCompl
    
        'obter do item ICMS_BASE, ICMS_VALOR, ICMS_BASE_SUBST, ICMS_VALOR_SUBST
        
        lErro = objTributoDocCompl.ObterVar("ICMS_BASE", objVarICMSBase)
        If lErro <> SUCESSO Then Error 27474
        
        lErro = objTributoDocCompl.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27475
        
        dPercRedBase = objVarPercRedBase.vValor
        
        lErro = objTributoDocCompl.ObterVar("ICMS_VALOR", objVarICMSValor)
        If lErro <> SUCESSO Then Error 27476
        
        lErro = objTributoDocCompl.ObterVar("ICMS_BASE_SUBST", objVarICMSBaseSubst)
        If lErro <> SUCESSO Then Error 27477
        
        lErro = objTributoDocCompl.ObterVar("ICMS_VALOR_SUBST", objVarICMSValorSubst)
        If lErro <> SUCESSO Then Error 27478
        
        lErro = objTributoDocCompl.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
        If lErro <> SUCESSO Then Error 27948
        
        lErro = objTributoDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32011
        
        If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
            dBase = dBase + objVarICMSBase.vValor
            dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBase))
            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.0000"))
            dBaseSubst = dBaseSubst + objVarICMSBaseSubst.vValor
            dBaseSubstReduzida = dBaseSubstReduzida + (objVarICMSBaseSubst.vValor * (1 - dPercRedBase))
            ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            dValor = dValor + objVarICMSValor.vValor
            dValorSubst = dValorSubst + objVarICMSValorSubst.vValor
            dValorCredito = dValorCredito + objVarICMSValorCredito.vValor
            
''        Else
''
''            dBase = dBase - objVarICMSBase.vValor
''            dBaseReduzida = dBaseReduzida - (objVarICMSBase.vValor * (1 - dPercRedBase))
''            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.0000"))
''            dBaseSubst = dBaseSubst - objVarICMSBaseSubst.vValor
''            dBaseSubstReduzida = dBaseSubstReduzida - (objVarICMSBaseSubst.vValor * (1 - dPercRedBase))
''            ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
''            dValor = dValor - objVarICMSValor.vValor
''            dValorSubst = dValorSubst - objVarICMSValorSubst.vValor
''            dValorCredito = dValorCredito - objVarICMSValorCredito.vValor
        
        End If
        
    Next
    
    AtualizarDoc_ObterDadosCompl = SUCESSO

    Exit Function

Erro_AtualizarDoc_ObterDadosCompl:

    AtualizarDoc_ObterDadosCompl = Err

    Select Case Err

        Case 27474 To 27478, 27482, 27948, 32011
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153839)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_ObterValoresAtuais(objTributoDocCompl As ClassTributoDocCompl, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer) As Long
'obtem os valores atuais da variaveis associadas a um Compl
Dim lErro As Long
On Error GoTo Erro_AtualizarCompl_ObterValoresAtuais

    lErro = objTributoDocCompl.ObterVar("ICMS_TIPO", objVarTipo)
    If lErro <> SUCESSO Then Error 27446
    If iRecalculaAutomatica = 1 Then objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objTributoDocCompl.ObterVar("ICMS_BASE", objVarBaseCalculo)
    If lErro <> SUCESSO Then Error 27447
    If iRecalculaAutomatica = 1 Then objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objTributoDocCompl.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
    If lErro <> SUCESSO Then Error 27448
    If iRecalculaAutomatica = 1 Then objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objTributoDocCompl.ObterVar("ICMS_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27449
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("ICMS_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27450
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("ICMS_BASE_SUBST", objVarBaseSubst)
    If lErro <> SUCESSO Then Error 27451
    If iRecalculaAutomatica = 1 Then objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
'    lErro = objTributoDocCompl.ObterVar("ICMS_MARG_SUBST", objVarMargSubst)
'    If lErro <> SUCESSO Then Error 27452
'
    lErro = objTributoDocCompl.ObterVar("ICMS_ALIQUOTA_SUBST", objVarAliqSubst)
    If lErro <> SUCESSO Then Error 27453
    If iRecalculaAutomatica = 1 Then objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("ICMS_VALOR_SUBST", objVarValorSubst)
    If lErro <> SUCESSO Then Error 27454
    If iRecalculaAutomatica = 1 Then objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("ICMS_VALOR_CREDITO", objVarValorCredito)
    If lErro <> SUCESSO Then Error 27875
    If iRecalculaAutomatica = 1 Then objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarCompl_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarCompl_ObterValoresAtuais:

    AtualizarCompl_ObterValoresAtuais = Err

    Select Case Err

        Case 27446 To 27454, 27875
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153840)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_GuardarValores(objTributoDocCompl As ClassTributoDocCompl, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, _
        dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double, dValorCredito As Double) As Integer
'atualiza as variaveis que tiveram seus valores recalculados pelo sistema
Dim lErro As Long
On Error GoTo Erro_AtualizarCompl_GuardarValores

    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTipo.vValor = iTipo
        objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarTipo)
        If lErro <> SUCESSO Then Error 27455
        
    End If
    
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseCalculo.vValor = dBaseCalculo
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarBaseCalculo)
        If lErro <> SUCESSO Then Error 27456
        
    End If
    
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarPercRedBase.vValor = dPercRedBase
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27457
        
    End If
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliquota.vValor = dAliquota
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then Error 27458
        
    End If
                
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then Error 27459
        
    End If
                
    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseSubst.vValor = dBaseSubst
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarBaseSubst)
        If lErro <> SUCESSO Then Error 27460
        
    End If
                
''    If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
''
''        objVarMargSubst.vValor = dMargSubst
''        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
''        lErro = objTributoDocCompl.GuardarVar(objVarMargSubst)
''        If lErro <> SUCESSO Then Error 27461
''
''    End If
                
    If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliqSubst.vValor = dAliqSubst
        objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarAliqSubst)
        If lErro <> SUCESSO Then Error 27462
        
    End If
                
    If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorSubst.vValor = dValorSubst
        objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarValorSubst)
        If lErro <> SUCESSO Then Error 27463
        
    End If
            
    If objVarValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorCredito.vValor = dValorCredito
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then Error 27873
        
    End If
            
    AtualizarCompl_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarCompl_GuardarValores:

    AtualizarCompl_GuardarValores = Err

    Select Case Err

        Case 27455 To 27463, 27873
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153841)

    End Select

    Exit Function

End Function

Private Function AtualizarDocCompl_SaidaCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarUFDestino As New ClassVariavelCalculo, objVarItemTipo As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double

Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection, dBaseSubstReduzida As Double
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFCli As String, dLixo As Double, dPercRedBaseSubst As Double

On Error GoTo Erro_AtualizarDocCompl_SaidaCli
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    lErro = objDoc.ObterVar("CLIENTE_UF", objVarUFDestino)
    If lErro <> SUCESSO Then Error 27726
            
    sUFCli = objVarUFDestino.vValor
    If sUFCli = "" Then sUFCli = mvarsUFFilialEmpresa
    
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        iTipo = ICMS_TIPO_NAO_TRIBUTADO
        
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dBaseSubst = 0
        dBaseSubstReduzida = 0
        dAliqSubst = 0
        dMargSubst = 0
        dValorSubst = 0
        dValorCredito = 0
        dPercRedBaseSubst = 0
        
        lErro = AtualizarCompl_ObterValoresAtuais(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27464
        
        'se a filialEmpresa nao está inscrita no ICMS ou está no Simples
        If mvariZerarTributo <> 0 Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27465
        
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 27466
        
        'se nao incide ICMS sobre a operacao
        If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
        'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
        If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
            iTipo = objVarTipo.vValor
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            objVarTipo.vValor = objTipoTrib.iICMSTipo
            iTipo = objTipoTrib.iICMSTipo
            
            If TipoICMSZerar(iTipo) Then
                Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
        Else
            iTipo = objVarTipo.vValor
        End If
        
        objTipoTribICMS.iTipo = iTipo
        lErro = TipoICMS_Obter(objTipoTribICMS)
        If lErro <> SUCESSO Then Error 27467
        
        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32252
        
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
            If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
                lErro = ObterAliquotaBasicaSaidaCli(objDocCompl.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFCli)
                If lErro <> SUCESSO Then Error 27468
                
            Else
            
                dAliquota = 0
                
            End If
            
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
        If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBase = 0
        Else
            dPercRedBase = objVarPercRedBase.vValor
        End If
            
'        'se ICMS_MARG_SUBST nao tiver sido preenchido manualmente...
'        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
'            dMargSubst = 0
'        Else
'            dMargSubst = objVarMargSubst.vValor
'        End If
        dMargSubst = 0
        
        'se ICMS_BASE nao preenchida manualmente
        If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
                
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27469
    
                dBaseCalculo = Round(objVar.vValor * IIf(mvardFatorValor = 0, 0, 1), 2)
    
                'se inclui o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27470
                    
                    dBaseCalculo = dBaseCalculo + objVar.vValor
                End If
            Else
                dBaseCalculo = 0
            End If
        Else
            dBaseCalculo = objVarBaseCalculo.vValor
        End If
        
        'se ICMS_VALOR nao preenchido manualmente
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
            dValor = CDbl(Format(dValor, "0.0000"))
            
        Else
            dValor = objVarValor.vValor
        End If
    
        If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If dMargSubst <> 0 Then
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
                Else
                    dBaseSubst = dBaseCalculo * (1 + dMargSubst)
                End If
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            Else
                dBaseSubst = 0
                dBaseSubstReduzida = dBaseSubst
            End If
        Else
            dBaseSubst = objVarBaseSubst.vValor
            dBaseSubstReduzida = dBaseSubst
        End If
                    
        If dBaseSubst <> 0 Then
        
            'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
            If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                
                'setar aliquota ICMS subst igual a aliq interna do estado destino
                lErro = AliqInterna_Pesquisar(sUFCli, dAliqSubst)
                If lErro <> SUCESSO Then Error 27472
                
            Else
                dAliqSubst = objVarAliqSubst.vValor
            End If
            
            'se ICMS_VALOR_SUBST nao preenchido manualmente
            If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
                dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
                dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
            Else
                dValorSubst = objVarValorSubst.vValor
            End If
            
        End If
    
        If objTipoTrib.iICMSCredita <> 0 Then
        
            dValorCredito = dValorCredito + dValor
            
        End If

Label_Atualizar:
        
        lErro = AtualizarCompl_GuardarValores(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, _
                                    iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dValorCredito)
        If lErro <> SUCESSO Then Error 27473
    
    Next
    
    AtualizarDocCompl_SaidaCli = SUCESSO

    Exit Function

Erro_AtualizarDocCompl_SaidaCli:

    AtualizarDocCompl_SaidaCli = Err

    Select Case Err

        Case 27464 To 27473, 27691 To 27694, 27766, 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153842)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_GuardarValores(objDoc As ClassTributoDoc, dBaseReduzida As Double, dValor As Double, dBaseSubstReduzida As Double, dValorSubst As Double, dValorCredito As Double) As Long
'atualiza ICMS_BASE, ICMS_VALOR, ICMS_SUBST_BASE, ICMS_SUBST_VALOR e ICMS_VALOR_CREDITO nao preenchidos manualmente
Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim objVarICMSValorCredito As New ClassVariavelCalculo
Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_GuardarValores

    lErro = objDoc.ObterVar("ICMS_BASE", objVarICMSBase)
    If lErro <> SUCESSO Then Error 27235
    If objVarICMSBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSBase.vValor = dBaseReduzida
        objVarICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSBase)
        If lErro <> SUCESSO Then Error 27236
    End If
    
    lErro = objDoc.ObterVar("ICMS_VALOR", objVarICMSValor)
    If lErro <> SUCESSO Then Error 27237
    If objVarICMSValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValor.vValor = dValor
        objVarICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValor)
        If lErro <> SUCESSO Then Error 27238
    End If
    
    lErro = objDoc.ObterVar("ICMS_SUBST_BASE", objVarICMSBaseSubst)
    If lErro <> SUCESSO Then Error 27239
    If objVarICMSBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSBaseSubst.vValor = dBaseSubstReduzida
        objVarICMSBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSBaseSubst)
        If lErro <> SUCESSO Then Error 27240
    End If
    
    lErro = objDoc.ObterVar("ICMS_SUBST_VALOR", objVarICMSValorSubst)
    If lErro <> SUCESSO Then Error 27241
    If objVarICMSValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValorSubst.vValor = dValorSubst
        objVarICMSValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValorSubst)
        If lErro <> SUCESSO Then Error 27242
    End If

    lErro = objDoc.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
    If lErro <> SUCESSO Then Error 27876
    If objVarICMSValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValorCredito.vValor = dValorCredito
        objVarICMSValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValorCredito)
        If lErro <> SUCESSO Then Error 27877
    End If

    AtualizarDoc_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarDoc_GuardarValores:

    AtualizarDoc_GuardarValores = Err

    Select Case Err

        Case 27235 To 27242, 27876, 27877
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153843)

    End Select

    Exit Function

End Function

Private Function PrepararAtualizacao_EntradaCli(objDoc As ClassTributoDoc) As Long
Dim lErro As Long, sUFOrigem As String
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarUFOrigem As New ClassVariavelCalculo
On Error GoTo Erro_PrepararAtualizacao_EntradaCli

''    sUFOrigem = ""
''
''    'obter codigo do cliente
''    lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
''    If lErro <> SUCESSO Then Error 27879
''
''    'obter identificador da filial do cliente
''    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
''    If lErro <> SUCESSO Then Error 27880
''
''    'se o cliente e a filial já foram informados
''    If (objVarFilialCliente.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.iStatus <> VAR_PREENCH_VAZIO) Then
''
''        lErro = objDoc.ObterVar("CLIENTE_UF", objVarUFOrigem)
''        If lErro <> SUCESSO Then Error 27881
''
''        sUFOrigem = objVarUFOrigem.vValor
''
''    End If
''
''    'faltam dados
''    If sUFOrigem = "" Then Error 27882

    PrepararAtualizacao_EntradaCli = SUCESSO

    Exit Function

Erro_PrepararAtualizacao_EntradaCli:

    PrepararAtualizacao_EntradaCli = Err

    Select Case Err
    
        Case 27879 To 27881
        
        Case 27882 'faltam dados p/calcular tributacao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153844)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_EntradaCli(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long, objExcecao As ClassICMSExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double, objVarPercRedBaseSubst As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo, dPercRedBaseSubst As Double

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFCli As String, dBaseSubstReduzida As Double
Dim objVarUFCli As New ClassVariavelCalculo, objFilCliInscEst As New ClassVariavelCalculo
Dim dValorPauta As Double, iUsaPauta As Integer, dQtdeSubst As Double, objVarQtdeSubst As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_EntradaCli
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dAliqSubst = 0
    dMargSubst = 0
    dValorSubst = 0
    dValorCredito = 0
    dPercRedBaseSubst = 0
    iUsaPauta = 0
    dValorPauta = 0

    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27912
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27916
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27920
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'obter UF Cliente
    lErro = objDocItem.objDoc.ObterVar("CLIENTE_UF", objVarUFCli)
    If lErro <> SUCESSO Then Error 27921
    
    sUFCli = objVarUFCli.vValor
    If sUFCli = "" Then sUFCli = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
    If lErro <> SUCESSO Then Error 27922
    
    'se filial cliente inscrita ICMS e interestadual
    If objFilCliInscEst.vValor <> "" And sUFCli <> mvarsUFFilialEmpresa And sUFCli <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFCli)
        If lErro <> SUCESSO Then Error 27923
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, sUFCli)
        If lErro <> SUCESSO Then Error 27924
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dValorPauta <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta)
                If lErro <> SUCESSO Then Error 27926
            End If
        End If
    
    Else
        If sUFCli <> "EX" Then
            'tratar como operacao interna
        
            'tentar achar excecao
            lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, "")
            If lErro <> SUCESSO Then Error 27925
            
            'se achou excecao
            If (Not (objExcecao Is Nothing)) Then
            
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, objFilCliInscEst.vValor = "", iUsaPauta, dValorPauta)
                If lErro <> SUCESSO Then Error 27926
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFCli)
                If lErro <> SUCESSO Then Error 27927
            
            End If
            
        Else  'importacao
        
            lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFCli)
            If lErro <> SUCESSO Then Error 27928
        
        End If
        
    End If

    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, True, True, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27929

            'se inclui o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27930
                
                dBaseCalculo = dBaseCalculo + objVar.vValor
            End If
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.0000"))
    
    Else
        dValor = objVarValor.vValor
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        If dMargSubst <> 0 Or dValorPauta <> 0 Then
            
            If iUsaPauta = 0 Then
            
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
                Else
                    dBaseSubst = dBaseCalculo * (1 + dMargSubst)
                End If
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            Else
            
                'obter qtde do item de objdocitem
                lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = CDbl(Format(objVar.vValor * dValorPauta, "0.0000"))
                dBaseSubstReduzida = dBaseSubst
            
            End If
        
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst
    End If
    
    If dBaseSubst <> 0 Then
        
        'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
        If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objExcecao Is Nothing Then
                dAliqSubst = 0
            Else
                dAliqSubst = objExcecao.dAliquota
            End If
        Else
            dAliqSubst = objVarAliqSubst.vValor
        End If
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
            dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
    End If

    If objTipoTrib.iICMSCredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dPercRedBaseSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27931
    
    AtualizarItem_EntradaCli = SUCESSO

    Exit Function

Erro_AtualizarItem_EntradaCli:

    AtualizarItem_EntradaCli = Err

    Select Case Err

        Case 27912 To 27931
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153845)

    End Select

    Exit Function

End Function

Private Function AliqInterstEntrada_Pesquisar(sSiglaOrigem As String, dAliquota As Double) As Long
'retorna a aliquota interestadual chegando ao estado da filial empresa corrente de sSiglaOrigem
Dim lErro As Long, objICMSAliqExterna As ClassICMSAliqExterna
On Error GoTo Erro_AliqInterstEntrada_Pesquisar
    
    Set objICMSAliqExterna = mvarcolAliqInterstEntrada.Item(sSiglaOrigem)
    
    dAliquota = objICMSAliqExterna.dAliquota
    
    AliqInterstEntrada_Pesquisar = SUCESSO

    Exit Function

Erro_AliqInterstEntrada_Pesquisar:

    AliqInterstEntrada_Pesquisar = Err

    Select Case Err

        Case 5 ' UF nao encontrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ICMS_ALIQ_INTERSTADUAL_INEXISTENTE", Err, Error)
            dAliquota = 0
            AliqInterstEntrada_Pesquisar = 27903
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153846)

    End Select

    Exit Function

End Function

Private Function ObterAliquotaBasicaEntCli(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFCli As String) As Long
'preenche a aliquota
Dim lErro As Long
Dim objFilCliInscEst As New ClassVariavelCalculo
On Error GoTo Erro_ObterAliquotaBasicaEntCli

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFCli = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            
        Else
          
            'se for importacao
            If sUFCli = "EX" Then
            
                dAliquota = mvardAliqImportacao
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
                If lErro <> SUCESSO Then Error 27904
                
                'se filial cliente nao inscrita ICMS
                If objFilCliInscEst.vValor = "" Then
                    
                    dAliquota = 0
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstEntrada_Pesquisar(sUFCli, dAliquota)
                    If lErro <> SUCESSO Then Error 27905
                
                End If
            End If
        End If
    Else
        dAliquota = 0
    End If

    ObterAliquotaBasicaEntCli = SUCESSO

    Exit Function

Erro_ObterAliquotaBasicaEntCli:

    ObterAliquotaBasicaEntCli = Err

    Select Case Err

        Case 27904, 27905
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153847)

    End Select

    Exit Function

End Function

Private Function ObterAliquotaBasicaEntForn(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFForn As String) As Long
'preenche a aliquota
Dim lErro As Long
Dim objInscEst As New ClassVariavelCalculo
On Error GoTo Erro_ObterAliquotaBasicaEntForn

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFForn = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            
        Else
          
            'se for importacao
            If sUFForn = "EX" Then
            
                dAliquota = mvardAliqImportacao
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_FORN_INSC_EST", objInscEst)
                If lErro <> SUCESSO Then Error 27906
                
                'se filial fornecedor nao inscrita ICMS
                If objInscEst.vValor = "" Then
                    
                    dAliquota = 0
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstEntrada_Pesquisar(sUFForn, dAliquota)
                    If lErro <> SUCESSO Then Error 27907
                
                End If
            End If
        End If
    Else
        dAliquota = 0
    End If

    ObterAliquotaBasicaEntForn = SUCESSO

    Exit Function

Erro_ObterAliquotaBasicaEntForn:

    ObterAliquotaBasicaEntForn = Err

    Select Case Err

        Case 27906, 27907
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153848)

    End Select

    Exit Function

End Function

Private Function ObterAliquotaBasicaSaidaForn(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFForn As String) As Long
'preenche a aliquota
Dim lErro As Long
Dim objInscEst As New ClassVariavelCalculo
On Error GoTo Erro_ObterAliquotaBasicaSaidaForn

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFForn = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            
        Else
          
            'se for exportacao
            If sUFForn = "EX" Then
            
                'setar aliquota igual a zero ( exportacao nao tem ICMS )
                dAliquota = 0
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_FORN_INSC_EST", objInscEst)
                If lErro <> SUCESSO Then Error 27908
                
                'se filial forn nao inscrita ICMS
                If objInscEst.vValor = "" Then
                    
                    'usar aliquota interna
                    dAliquota = mvardAliqInterna
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstSaida_Pesquisar(sUFForn, dAliquota)
                    If lErro <> SUCESSO Then Error 27909
                
                End If
            End If
        End If
    Else
        dAliquota = 0
    End If

    ObterAliquotaBasicaSaidaForn = SUCESSO

    Exit Function

Erro_ObterAliquotaBasicaSaidaForn:

    ObterAliquotaBasicaSaidaForn = Err

    Select Case Err

        Case 27908, 27909
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153849)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecaoEntCli(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFCli As String) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoEntCli
        
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objTipoTrib.iICMSTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27910
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaEntCli(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFCli)
            If lErro <> SUCESSO Then Error 27911
        Else
            dAliquota = 0
        End If
    Else
    
        dAliquota = objVarAliquota.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecaoEntCli = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoEntCli:

    AtualizarItem_TratarNaoExcecaoEntCli = Err

    Select Case Err

        Case 27910, 27911
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153850)

    End Select

    Exit Function

End Function

Private Function AtualizarDocCompl_EntCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarUFCli As New ClassVariavelCalculo, objVarItemTipo As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double, dPercRedBaseSubst As Double

Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection, dBaseSubstReduzida As Double
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFCli As String, dLixo As Double

On Error GoTo Erro_AtualizarDocCompl_EntCli
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    lErro = objDoc.ObterVar("CLIENTE_UF", objVarUFCli)
    If lErro <> SUCESSO Then Error 27932
            
    sUFCli = objVarUFCli.vValor
    If sUFCli = "" Then sUFCli = mvarsUFFilialEmpresa
            
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        iTipo = ICMS_TIPO_NAO_TRIBUTADO
        
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dBaseSubst = 0
        dBaseSubstReduzida = 0
        dAliqSubst = 0
        dMargSubst = 0
        dValorSubst = 0
        dValorCredito = 0
        dPercRedBaseSubst = 0
        
        lErro = AtualizarCompl_ObterValoresAtuais(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27933
        
        'se a filialEmpresa nao está inscrita no ICMS ou está no Simples
        If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27937
                
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 27941
        
        'se nao incide ICMS sobre a operacao
        If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
        'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
        If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
            iTipo = objVarTipo.vValor
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            objVarTipo.vValor = objTipoTrib.iICMSTipo
            iTipo = objTipoTrib.iICMSTipo
            
            If TipoICMSZerar(iTipo) Then
                Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
        Else
            iTipo = objVarTipo.vValor
        End If
        
        objTipoTribICMS.iTipo = iTipo
        lErro = TipoICMS_Obter(objTipoTribICMS)
        If lErro <> SUCESSO Then Error 27942
        
        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32252
        
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then

            If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
                lErro = ObterAliquotaBasicaEntCli(objDocCompl.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFCli)
                If lErro <> SUCESSO Then Error 27943
                
            Else
            
                dAliquota = 0
                
            End If
            
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
        If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        
            dPercRedBase = 0
            
        Else
            dPercRedBase = objVarPercRedBase.vValor
        End If
            
'        'se ICMS_MARG_SUBST nao tiver sido preenchido manualmente...
'        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
'            dMargSubst = 0
'        Else
'            dMargSubst = objVarMargSubst.vValor
'        End If
        dMargSubst = 0
        
        'se ICMS_BASE nao preenchida manualmente
        If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
                
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27944
    
                dBaseCalculo = Round(objVar.vValor * IIf(mvardFatorValor = 0, 0, 1), 2)
    
                'se inclui o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27945
                    
                    dBaseCalculo = dBaseCalculo + objVar.vValor
                End If
            Else
                dBaseCalculo = 0
            End If
        Else
            dBaseCalculo = objVarBaseCalculo.vValor
        End If
        
        'se ICMS_VALOR nao preenchido manualmente
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
            dValor = CDbl(Format(dValor, "0.0000"))
            
        Else
            dValor = objVarValor.vValor
        End If
    
        If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If dMargSubst <> 0 Then
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
                Else
                    dBaseSubst = dBaseCalculo * (1 + dMargSubst)
                End If
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            Else
                dBaseSubst = 0
                dBaseSubstReduzida = dBaseSubst
            End If
        Else
            dBaseSubst = objVarBaseSubst.vValor
            dBaseSubstReduzida = dBaseSubst
        End If
            
        If dBaseSubst <> 0 Then
        
            'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
            If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                
                'setar aliquota ICMS subst igual a aliq interna do estado destino
                lErro = AliqInterna_Pesquisar(mvarsUFFilialEmpresa, dAliqSubst)
                If lErro <> SUCESSO Then Error 27946
                
            Else
                dAliqSubst = objVarAliqSubst.vValor
            End If
            
            'se ICMS_VALOR_SUBST nao preenchido manualmente
            If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
                dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
                dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
            Else
                dValorSubst = objVarValorSubst.vValor
            End If
            
        End If
    
        If objTipoTrib.iICMSCredita <> 0 Then
        
            dValorCredito = dValorCredito + dValor
            
        End If
        
Label_Atualizar:
        
        lErro = AtualizarCompl_GuardarValores(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, _
                                    iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dValorCredito)
        If lErro <> SUCESSO Then Error 27947
    
    Next
    
    AtualizarDocCompl_EntCli = SUCESSO

    Exit Function

Erro_AtualizarDocCompl_EntCli:

    AtualizarDocCompl_EntCli = Err

    Select Case Err

        Case 27932 To 27947, 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153851)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc) As Long
'atualiza base de ICMS, valor de ICMS, base ICMS subst, valor ICMS Subst
'    leva em conta dados dos itens + dados de frete, seg, desp acess.
Dim colItens As Collection, dPercRedBaseSubst As Double
Dim objDocItem As ClassTributoDocItem, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dBase As Double, dBaseSubst As Double, dBaseReduzida As Double, dBaseSubstReduzida As Double, dValorCredito As Double
Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim objVarFreteValor As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
Dim objVarSeguroValor As New ClassVariavelCalculo, objVarDespAcesValor As New ClassVariavelCalculo
Dim lErro As Long, dPercRedBase As Double, objVarICMSValorCredito As New ClassVariavelCalculo
On Error GoTo Erro_AtualizarDoc

    dValor = 0
    dValorSubst = 0
    dBase = 0
    dBaseReduzida = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dValorCredito = 0
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
        
        'obter do item ICMS_BASE, ICMS_VALOR, ICMS_BASE_SUBST, ICMS_VALOR_SUBST
        
        lErro = objDocItem.ObterVar("ICMS_BASE", objVarICMSBase)
        If lErro <> SUCESSO Then Error 27949
        
        lErro = objDocItem.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27950
         
        dPercRedBase = objVarPercRedBase.vValor
        
        lErro = objDocItem.ObterVar("ICMS_VALOR", objVarICMSValor)
        If lErro <> SUCESSO Then Error 27951
        
        lErro = objDocItem.ObterVar("ICMS_BASE_SUBST", objVarICMSBaseSubst)
        If lErro <> SUCESSO Then Error 27952
        
        lErro = objDocItem.ObterVar("ICMS_VALOR_SUBST", objVarICMSValorSubst)
        If lErro <> SUCESSO Then Error 27953
        
        lErro = objDocItem.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
        If lErro <> SUCESSO Then Error 27954
        
        dBase = dBase + objVarICMSBase.vValor
        dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBase))
        ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
        dBaseSubst = dBaseSubst + objVarICMSBaseSubst.vValor
        dPercRedBaseSubst = 0 '???? ainda nao está sendo armazenado
        dBaseSubstReduzida = dBaseSubstReduzida + (objVarICMSBaseSubst.vValor * (1 - dPercRedBaseSubst))
        ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.00"))
        dValor = dValor + objVarICMSValor.vValor
        dValorSubst = dValorSubst + objVarICMSValorSubst.vValor
        dValorCredito = dValorCredito + objVarICMSValorCredito.vValor
        
    Next
    
    'tratar dados globais (frete, desp acessorias,...) inclusive vendo o impacto do IPI sobre o frete,... na base de calculo do ICMS
    lErro = AtualizarDoc_ObterDadosCompl(objDoc, dValor, dValorSubst, dBase, dBaseReduzida, dBaseSubst, dBaseSubstReduzida, dValorCredito)
    If lErro <> SUCESSO Then Error 27955
        
    'arredondar valores globais
    dBase = Arredonda_Moeda(dBase, 2)
    dBaseReduzida = Arredonda_Moeda(dBaseReduzida, 2)
    dBaseSubst = Arredonda_Moeda(dBaseSubst, 2)
    dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida, 2)
    dValor = Arredonda_Moeda(dValor, 2)
    dValorSubst = Arredonda_Moeda(dValorSubst, 2)
    dValorCredito = Arredonda_Moeda(dValorCredito, 2)
        
    lErro = AtualizarDoc_GuardarValores(objDoc, dBaseReduzida, dValor, dBaseSubstReduzida, dValorSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27956
        
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27949 To 27956
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153852)

    End Select

    Exit Function

End Function

Private Function PrepararAtualizacao_Forn(objDoc As ClassTributoDoc) As Long
Dim lErro As Long, sUFOrigem As String
Dim objVarForn As New ClassVariavelCalculo, objVarFilialForn As New ClassVariavelCalculo
Dim objVarUFOrigem As New ClassVariavelCalculo
On Error GoTo Erro_PrepararAtualizacao_Forn

''    sUFOrigem = ""
''
''    'obter codigo do forn
''    lErro = objDoc.ObterVar("FORN_CODIGO", objVarForn)
''    If lErro <> SUCESSO Then Error 27970
''
''    'obter identificador da filial do forn
''    lErro = objDoc.ObterVar("FILIAL_FORN_CODIGO", objVarFilialForn)
''    If lErro <> SUCESSO Then Error 27971
''
''    'se o forn e a filial já foram informados
''    If (objVarFilialForn.iStatus <> VAR_PREENCH_VAZIO And objVarForn.iStatus <> VAR_PREENCH_VAZIO) Then
''
''        lErro = objDoc.ObterVar("FORN_UF", objVarUFOrigem)
''        If lErro <> SUCESSO Then Error 27972
''
''        sUFOrigem = objVarUFOrigem.vValor
''
''    End If
''
''    'faltam dados
''    If sUFOrigem = "" Then Error 27973

    PrepararAtualizacao_Forn = SUCESSO

    Exit Function

Erro_PrepararAtualizacao_Forn:

    PrepararAtualizacao_Forn = Err

    Select Case Err
    
        Case 27970 To 27972
        
        Case 27973 'faltam dados p/calcular tributacao
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153853)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecaoEntForn(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFForn As String) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoEntForn
        
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objTipoTrib.iICMSTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 22700
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaEntForn(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn)
            If lErro <> SUCESSO Then Error 22701
        Else
            dAliquota = 0
        End If
    Else
    
        dAliquota = objVarAliquota.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecaoEntForn = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoEntForn:

    AtualizarItem_TratarNaoExcecaoEntForn = Err

    Select Case Err

        Case 22700, 22701
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153854)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_EntradaForn(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long, objExcecao As ClassICMSExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double
Dim objVarTipoTribDoc As New ClassVariavelCalculo, objVarTipoDocSigla As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dPercRedBaseSubst As Double
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, objVarPercRedBaseSubst As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_EntradaForn
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dBaseSubst = 0
    '???dBaseSubstReduzida = 0
    dAliqSubst = 0
    dMargSubst = 0
    dValorSubst = 0
    dValorCredito = 0
    dPercRedBaseSubst = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27974
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27978
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27985
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'obter UF Forn
    lErro = objDocItem.objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 27986
    
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_INSC_EST", objFilFornInscEst)
    If lErro <> SUCESSO Then Error 27987
    
    'se filial forn inscrita ICMS e interestadual
    If objFilFornInscEst.vValor <> "" And sUFForn <> mvarsUFFilialEmpresa And sUFForn <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
        If lErro <> SUCESSO Then Error 27988
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, sUFForn)
        If lErro <> SUCESSO Then Error 27989
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False)
                If lErro <> SUCESSO Then Error 27991
            End If
        End If
    
    Else
        If sUFForn <> "EX" Then
            'tratar como operacao interna
        
            'tentar achar excecao
            lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, "")
            If lErro <> SUCESSO Then Error 27990
            
            'se achou excecao
            If (Not (objExcecao Is Nothing)) Then
            
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False)
                If lErro <> SUCESSO Then Error 27991
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
                If lErro <> SUCESSO Then Error 27992
            
            End If
            
        Else  'importacao
        
            lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
            If lErro <> SUCESSO Then Error 27993
        
        End If
        
    End If

    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, True, True, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27994

            'se inclui o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27995
                
                dBaseCalculo = dBaseCalculo + objVar.vValor
            End If
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.00"))
    
    Else
        If dBaseCalculo <> 0 Then
            dValor = objVarValor.vValor
        Else
            dValor = 0
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        End If
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        If dMargSubst <> 0 Then
            'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI = 0 Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
            Else
                dBaseSubst = dBaseCalculo * (1 + dMargSubst)
            End If
            dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
        Else
            dBaseSubst = 0
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
    End If
    
    If dBaseSubst <> 0 Then
        
        'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
        If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objExcecao Is Nothing Then
                dAliqSubst = 0
            Else
                dAliqSubst = objExcecao.dAliquota
            End If
        Else
            dAliqSubst = objVarAliqSubst.vValor
        End If
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dValor
            dValorSubst = CDbl(Format(dValorSubst, "0.00"))
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
    
        dValorSubst = 0
        objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dMargSubst = 0
        dPercRedBaseSubst = 0
    
    End If

    If objTipoTrib.iICMSCredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27996
    
    AtualizarItem_EntradaForn = SUCESSO

    Exit Function

Erro_AtualizarItem_EntradaForn:

    AtualizarItem_EntradaForn = Err

    Select Case Err

        Case 27974 To 27996
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153855)

    End Select

    Exit Function

End Function

Private Function AtualizarDocCompl_EntForn(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarUFForn As New ClassVariavelCalculo, objVarItemTipo As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double, dPercRedBaseSubst As Double

Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dLixo As Double
On Error GoTo Erro_AtualizarDocCompl_EntForn
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    lErro = objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 22702
    
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
            
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        iTipo = ICMS_TIPO_NAO_TRIBUTADO
        
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dBaseSubst = 0
        '???dBaseSubstReduzida = 0
        dAliqSubst = 0
        dMargSubst = 0
        dValorSubst = 0
        dValorCredito = 0
        dPercRedBaseSubst = 0
        
        lErro = AtualizarCompl_ObterValoresAtuais(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 22703
        
        'se a filialEmpresa nao está inscrita no ICMS ou está no Simples
        If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 22707
                
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 22711
        
        'se nao incide ICMS sobre a operacao
        If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
        'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
        If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
            iTipo = objVarTipo.vValor
            
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            objVarTipo.vValor = objTipoTrib.iICMSTipo
            iTipo = objTipoTrib.iICMSTipo
            
            If TipoICMSZerar(iTipo) Then
                Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
        Else
            iTipo = objVarTipo.vValor
        End If
        
        objTipoTribICMS.iTipo = iTipo
        lErro = TipoICMS_Obter(objTipoTribICMS)
        If lErro <> SUCESSO Then Error 22712
        
        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32252
        
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
            If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
                lErro = ObterAliquotaBasicaEntForn(objDocCompl.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn)
                If lErro <> SUCESSO Then Error 22713
                
            Else
            
                dAliquota = 0
                
            End If
            
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
        If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBase = 0
        Else
            dPercRedBase = objVarPercRedBase.vValor
        End If
            
'        'se ICMS_MARG_SUBST nao tiver sido preenchido manualmente...
'        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
'            dMargSubst = 0
'        Else
'            dMargSubst = objVarMargSubst.vValor
'        End If
        dMargSubst = 0
        
        'se ICMS_BASE nao preenchida manualmente
        If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
                
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVar)
                If lErro <> SUCESSO Then Error 22714
    
                dBaseCalculo = Round(objVar.vValor * IIf(mvardFatorValor = 0, 0, 1), 2)
    
                'se inclui o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 22715
                    
                    dBaseCalculo = dBaseCalculo + objVar.vValor
                End If
            Else
                dBaseCalculo = 0
            End If
        Else
            dBaseCalculo = objVarBaseCalculo.vValor
        End If
        
        'se ICMS_VALOR nao preenchido manualmente
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
            dValor = CDbl(Format(dValor, "0.00"))
            
        Else
            dValor = objVarValor.vValor
        End If
    
        If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If dMargSubst <> 0 Then
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
                Else
                    dBaseSubst = dBaseCalculo * (1 + dMargSubst)
                End If
                dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
            Else
                dBaseSubst = 0
            End If
        Else
            dBaseSubst = objVarBaseSubst.vValor
        End If
            
        If dBaseSubst <> 0 Then
        
            'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
            If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                
                'setar aliquota ICMS subst igual a aliq interna do estado destino
                lErro = AliqInterna_Pesquisar(mvarsUFFilialEmpresa, dAliqSubst)
                If lErro <> SUCESSO Then Error 22716
                
            Else
                dAliqSubst = objVarAliqSubst.vValor
            End If
            
            'se ICMS_VALOR_SUBST nao preenchido manualmente
            If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
                dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dValor
                dValorSubst = CDbl(Format(dValorSubst, "0.00"))
            Else
                dValorSubst = objVarValorSubst.vValor
            End If
            
        End If
    
        If objTipoTrib.iICMSCredita <> 0 Then
        
            dValorCredito = dValorCredito + dValor
            
        End If
        
Label_Atualizar:
        
        lErro = AtualizarCompl_GuardarValores(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, _
                                    iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dValorCredito)
        If lErro <> SUCESSO Then Error 22717
    
    Next
    
    AtualizarDocCompl_EntForn = SUCESSO

    Exit Function

Erro_AtualizarDocCompl_EntForn:

    AtualizarDocCompl_EntForn = Err

    Select Case Err

        Case 22702 To 22717, 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153856)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_SaidaForn(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long, objExcecao As ClassICMSExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double
Dim objVarTipoTribDoc As New ClassVariavelCalculo, objVarTipoDocSigla As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo, dPercRedBaseSubst As Double

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, sUFDestino As String, objVarPercRedBaseSubst As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_SaidaForn
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dBaseSubst = 0
    '???dBaseSubstReduzida = 0
    dAliqSubst = 0
    dMargSubst = 0
    dValorSubst = 0
    dValorCredito = 0
    dPercRedBaseSubst = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 22728
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 22732
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 22739
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'obter UF Forn
    lErro = objDocItem.objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 22740
    
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
    
    sUFDestino = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_INSC_EST", objFilFornInscEst)
    If lErro <> SUCESSO Then Error 22741
     
    'se filial forn inscrita ICMS e interestadual
    If objFilFornInscEst.vValor <> "" And sUFForn <> mvarsUFFilialEmpresa And sUFForn <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
        If lErro <> SUCESSO Then Error 22742
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, sUFForn, mvarsUFFilialEmpresa)
        If lErro <> SUCESSO Then Error 22743
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False)
                If lErro <> SUCESSO Then Error 22745
            End If
        End If
    
    Else
        If sUFForn <> "EX" Then
            'tratar como operacao interna
        
            'tentar achar excecao
            lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, sUFForn, "")
            If lErro <> SUCESSO Then Error 22744
            
            'se achou excecao
            If (Not (objExcecao Is Nothing)) Then
            
                lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, False)
                If lErro <> SUCESSO Then Error 22745
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
                If lErro <> SUCESSO Then Error 22746
            
            End If
            
        Else  'exportacao
        
            lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn)
            If lErro <> SUCESSO Then Error 22747
        
        End If
        
    End If

    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = Item_ObtemBaseCalculo(mvariFilialEmpresa, objDocItem, mvardFatorValor, True, True, dBaseCalculo)
            If lErro <> SUCESSO Then Error 22748

            'se inclui o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 22749
                
                dBaseCalculo = dBaseCalculo + objVar.vValor
            End If
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.00"))
    
    Else
        dValor = objVarValor.vValor
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        If dMargSubst <> 0 Then
            'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI = 0 Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
            Else
                dBaseSubst = dBaseCalculo * (1 + dMargSubst)
            End If
            dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
        Else
            dBaseSubst = 0
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
    End If
    
    If dBaseSubst <> 0 Then
        
        'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
        If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objExcecao Is Nothing Then
                dAliqSubst = 0
            Else
                dAliqSubst = objExcecao.dAliquota
            End If
        Else
            dAliqSubst = objVarAliqSubst.vValor
        End If
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dValor
            dValorSubst = CDbl(Format(dValorSubst, "0.00"))
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
    End If

    If objTipoTrib.iICMSCredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 22750
    
    AtualizarItem_SaidaForn = SUCESSO

    Exit Function

Erro_AtualizarItem_SaidaForn:

    AtualizarItem_SaidaForn = Err

    Select Case Err

        Case 22728 To 22750
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153857)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFForn As String) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoSaidaForn
        
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objTipoTrib.iICMSTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 22751
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaSaidaForn(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn)
            If lErro <> SUCESSO Then Error 22752
        Else
            dAliquota = 0
        End If
    Else
    
        dAliquota = objVarAliquota.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecaoSaidaForn = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoSaidaForn:

    AtualizarItem_TratarNaoExcecaoSaidaForn = Err

    Select Case Err

        Case 22751, 22752
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153858)

    End Select

    Exit Function

End Function

Private Function AtualizarDocCompl_SaidaForn(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarUFForn As New ClassVariavelCalculo, objVarItemTipo As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double

Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dLixo As Double, dPercRedBaseSubst As Double

On Error GoTo Erro_AtualizarDocCompl_SaidaForn
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    lErro = objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 22753
            
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
    
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        iTipo = ICMS_TIPO_NAO_TRIBUTADO
        
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dBaseSubst = 0
        '???dBaseSubstReduzida = 0
        dAliqSubst = 0
        dMargSubst = 0
        dValorSubst = 0
        dValorCredito = 0
        dPercRedBaseSubst = 0
        
        lErro = AtualizarCompl_ObterValoresAtuais(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 22754
        
        'se a filialEmpresa nao está inscrita no ICMS ou está no Simples
        If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 22758
                
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 22762
        
        'se nao incide ICMS sobre a operacao
        If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
        'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
        If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
            iTipo = objVarTipo.vValor
            Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            objVarTipo.vValor = objTipoTrib.iICMSTipo
            iTipo = objTipoTrib.iICMSTipo
            
            If TipoICMSZerar(iTipo) Then
                Call AtualizarCompl_ZerarValores(objDocCompl, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
        Else
            iTipo = objVarTipo.vValor
        End If
        
        objTipoTribICMS.iTipo = iTipo
        lErro = TipoICMS_Obter(objTipoTribICMS)
        If lErro <> SUCESSO Then Error 22763
        
        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32252
        
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
            If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
                lErro = ObterAliquotaBasicaEntForn(objDocCompl.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn)
                If lErro <> SUCESSO Then Error 22764
                
            Else
            
                dAliquota = 0
            
            End If
            
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
        If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBase = 0
        Else
            dPercRedBase = objVarPercRedBase.vValor
        End If
            
'        'se ICMS_MARG_SUBST nao tiver sido preenchido manualmente...
'        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
'            dMargSubst = 0
'        Else
'            dMargSubst = objVarMargSubst.vValor
'        End If
        dMargSubst = 0
        
        'se ICMS_BASE nao preenchida manualmente
        If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
                
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVar)
                If lErro <> SUCESSO Then Error 22765
    
                dBaseCalculo = Round(objVar.vValor * IIf(mvardFatorValor = 0, 0, 1), 2)
    
                'se inclui o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 22766
                    
                    dBaseCalculo = dBaseCalculo + objVar.vValor
                End If
            Else
                dBaseCalculo = 0
            End If
        Else
            dBaseCalculo = objVarBaseCalculo.vValor
        End If
        
        'se ICMS_VALOR nao preenchido manualmente
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
            dValor = CDbl(Format(dValor, "0.00"))
            
        Else
            dValor = objVarValor.vValor
        End If
    
        If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If dMargSubst <> 0 Then
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocCompl.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = (dBaseCalculo + objVar.vValor) * (1 + dMargSubst)
                Else
                    dBaseSubst = dBaseCalculo * (1 + dMargSubst)
                End If
                dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
            Else
                dBaseSubst = 0
            End If
        Else
            dBaseSubst = objVarBaseSubst.vValor
        End If
            
        If dBaseSubst <> 0 Then
        
            'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
            If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                
                'setar aliquota ICMS subst igual a aliq interna do estado destino
                lErro = AliqInterna_Pesquisar(mvarsUFFilialEmpresa, dAliqSubst)
                If lErro <> SUCESSO Then Error 22767
                
            Else
                dAliqSubst = objVarAliqSubst.vValor
            End If
            
            'se ICMS_VALOR_SUBST nao preenchido manualmente
            If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
                dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dValor
                dValorSubst = CDbl(Format(dValorSubst, "0.00"))
            Else
                dValorSubst = objVarValorSubst.vValor
            End If
            
        End If
            
        If objTipoTrib.iICMSCredita <> 0 Then
        
            dValorCredito = dValorCredito + dValor
            
        End If

Label_Atualizar:
        
        lErro = AtualizarCompl_GuardarValores(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, _
                                    iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dValorCredito)
        If lErro <> SUCESSO Then Error 22768
    
    Next
    
    AtualizarDocCompl_SaidaForn = SUCESSO

    Exit Function

Erro_AtualizarDocCompl_SaidaForn:

    AtualizarDocCompl_SaidaForn = Err

    Select Case Err

        Case 22753 To 22768, 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153859)

    End Select

    Exit Function

End Function


Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objEndereco As New ClassEndereco
Dim objEstado As ClassEstado, objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32021
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        mvariFilialEmpresa = 0 'se der erro no meio forca a carga na vez seguinte
                
        mvariSimplificado = objFilialEmpresa.iICMSPorEstimativa
        mvariSimplesNacional = objFilialEmpresa.iSuperSimples
        
        If objFilialEmpresa.sInscricaoEstadual = "" Then
            mvariInscrita = ICMS_NAO_INSCRITA
        Else
            mvariInscrita = ICMS_INSCRITA
        End If
             
        'obter sigla da UF da propria filial empresa
        mvarsUFFilialEmpresa = objFilialEmpresa.objEndereco.sSiglaEstado
        
        'carregar aliquotas interestaduais a partir do Estado da filialEmpresa
        Set mvarcolAliqInterstSaida = New Collection
        lErro = CF("ICMSAliqExternas_Le_EstadoOrigem", mvarsUFFilialEmpresa, mvarcolAliqInterstSaida)
        If lErro <> SUCESSO Then Error 27212
        
        'carregar aliquotas interestaduais com destino ao Estado da filialEmpresa
        Set mvarcolAliqInterstEntrada = New Collection
        lErro = CF("ICMSAliqExternas_Le_EstadoDestino", mvarsUFFilialEmpresa, mvarcolAliqInterstEntrada)
        If lErro <> SUCESSO Then Error 27862
        
        'carregar aliquota interna, a de imp e a de exp do Estado da filialEmpresa
        Set objEstado = mvarcolEstados.Item(mvarsUFFilialEmpresa)
        
        mvardAliqInterna = objEstado.dICMSAliquotaInterna
        'mvardAliqExportacao = objEstado.dICMSAliquotaExportacao
        mvardAliqImportacao = objEstado.dICMSAliquotaImportacao
        
        'carregar colecao de excecoes ICMS para o estado da filialEmpresa
        Set mvarcolExcecoesInternas = New Collection
        lErro = CF("ICMSExcecoes_Le_Estado", mvarsUFFilialEmpresa, mvarcolExcecoesInternas)
        If lErro <> SUCESSO Then Error 27214
    
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 27212, 27214, 27862, 32021
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153860)

    End Select

    Exit Function

End Function

'Private Function ObterInfoProporcionalItens(objDoc As ClassTributoDoc, dAliquota As Double, dPercRedBase As Double) As Long
'
'Dim lErro As Long, dValor As Double, dBase As Double, dBaseReduzida As Double
'Dim colItens As Collection, dPercRedBaseItem As Double
'Dim objDocItem As ClassTributoDocItem, objVarPercRedBase As New ClassVariavelCalculo
'Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
'
'On Error GoTo Erro_ObterInfoProporcionalItens
'
'    dValor = 0
'    dBase = 0
'    dBaseReduzida = 0
'
'    Call objDoc.ObterColDocItens(colItens)
'
'    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
'    For Each objDocItem In colItens
'
'        'obter do item ICMS_BASE, ICMS_VALOR, ICMS_BASE_SUBST, ICMS_VALOR_SUBST
'
'        lErro = objDocItem.ObterVar("ICMS_BASE", objVarICMSBase)
'        If lErro <> SUCESSO Then gError 81612
'
'        lErro = objDocItem.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
'        If lErro <> SUCESSO Then gError 81613
'
'        dPercRedBaseItem = objVarPercRedBase.vValor
'
'        lErro = objDocItem.ObterVar("ICMS_VALOR", objVarICMSValor)
'        If lErro <> SUCESSO Then gError 81614
'
'        dBase = dBase + objVarICMSBase.vValor
'        dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBaseItem))
'        dValor = dValor + objVarICMSValor.vValor
'
'    Next
'
'    If dBaseReduzida <> 0 Then
'        dAliquota = dValor / dBaseReduzida
'    Else
'        dAliquota = 0
'    End If
'
'    If dBase <> 0 Then
'        dPercRedBase = (dBase - dBaseReduzida) / dBase
'    Else
'        dPercRedBase = 0
'    End If
'
'    ObterInfoProporcionalItens = SUCESSO
'
'    Exit Function
'
'Erro_ObterInfoProporcionalItens:
'
'    ObterInfoProporcionalItens = gErr
'
'    Select Case gErr
'
'        Case 81612, 81613, 81614
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153861)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function AtualizarDoc_Recalculo(objDoc As ClassTributoDoc) As Long
'marca os valores a nivel de doc como VAR_PREENCH_AUTOMATICO

Dim lErro As Long
Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim objVarICMSValorCredito As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc_Recalculo

    lErro = objDoc.ObterVar("ICMS_BASE", objVarICMSBase)
    If lErro <> SUCESSO Then gError 81633
    objVarICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSBase)
    If lErro <> SUCESSO Then gError 81641
    
    lErro = objDoc.ObterVar("ICMS_VALOR", objVarICMSValor)
    If lErro <> SUCESSO Then gError 81634
    objVarICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSValor)
    If lErro <> SUCESSO Then gError 81642
    
    lErro = objDoc.ObterVar("ICMS_SUBST_BASE", objVarICMSBaseSubst)
    If lErro <> SUCESSO Then gError 81635
    objVarICMSBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSBaseSubst)
    If lErro <> SUCESSO Then gError 81643
    
    lErro = objDoc.ObterVar("ICMS_SUBST_VALOR", objVarICMSValorSubst)
    If lErro <> SUCESSO Then gError 81636
    objVarICMSValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSValorSubst)
    If lErro <> SUCESSO Then gError 81644

    lErro = objDoc.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
    If lErro <> SUCESSO Then gError 81637
    objVarICMSValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSValorCredito)
    If lErro <> SUCESSO Then gError 81645

    AtualizarDoc_Recalculo = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_Recalculo:

    AtualizarDoc_Recalculo = gErr
     
    Select Case gErr
                  
        Case 81633 To 81637, 81641 To 81645
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153862)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarItem_ZerarValores(objDocItem As ClassTributoDocItem, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo) As Long
'zera os valores atuais da variaveis associadas a um item

Dim lErro As Long
On Error GoTo Erro_AtualizarItem_ZerarValores

    objVarBaseCalculo.vValor = 0
    objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarPercRedBase.vValor = 0
    objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarAliquota.vValor = 0
    objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValor.vValor = 0
    objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarBaseSubst.vValor = 0
    objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarPercRedBaseSubst.vValor = 0
    objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarAliqSubst.vValor = 0
    objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValorSubst.vValor = 0
    objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarMargSubst.vValor = 0
    objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValorCredito.vValor = 0
    objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarItem_ZerarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_ZerarValores:

    AtualizarItem_ZerarValores = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153863)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_ZerarValores(objTributoDocCompl As ClassTributoDocCompl, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo) As Long
'obtem os valores atuais da variaveis associadas a um Compl
Dim lErro As Long
On Error GoTo Erro_AtualizarCompl_ZerarValores

    objVarBaseCalculo.vValor = 0
    objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO

    objVarPercRedBase.vValor = 0
    objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO

    objVarAliquota.vValor = 0
    objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValor.vValor = 0
    objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarBaseSubst.vValor = 0
    objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarAliqSubst.vValor = 0
    objVarAliqSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValorSubst.vValor = 0
    objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
            
    objVarValorCredito.vValor = 0
    objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarCompl_ZerarValores = SUCESSO

    Exit Function

Erro_AtualizarCompl_ZerarValores:

    AtualizarCompl_ZerarValores = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153864)

    End Select

    Exit Function

End Function

Private Function ExcecaoForn_Pesquisar(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, sUFDestino As String, sUFOrigem As String) As Long
'uma excecao a nivel de fornecedor tem + prioridade que excecoes a nivel de produto apenas
'e uma excecao a nivel de fornecedor+produto tem a maior prioridade de todas.
'A colecao já está ordenada de forma que pararei na 1a excecao que encontrar
'retorna Nothing se nao achar nenhuma excecao adequada
Dim lErro As Long, objICMSExcecao As ClassICMSExcecao, iAchou As Integer
Dim objVarUFDestino As New ClassVariavelCalculo, objVarFilFornColCateg As New ClassVariavelCalculo
Dim objVarProdColCateg As New ClassVariavelCalculo
Dim colExcecoes As New Collection

On Error GoTo Erro_ExcecaoForn_Pesquisar

    'obter a colecao de categorias da filialfornecedor
    lErro = objDocItem.objDoc.ObterVar("I_FILIAL_FORN_COL_CATEG", objVarFilFornColCateg)
    If lErro <> SUCESSO Then Error 27220
    
    'obter a colecao de categorias do produto
    lErro = objDocItem.ObterVar("I_PRODUTO_COL_CATEG", objVarProdColCateg)
    If lErro <> SUCESSO Then Error 27221
    
    If sUFDestino = mvarsUFFilialEmpresa Then
        Set colExcecoes = mvarcolExcecoesInternas
    Else
        If mvarsUFcolExcecoesAlt <> sUFDestino Then
            'carregar colecao de excecoes ICMS para o estado destino
            Set mvarcolExcecoesAlt = New Collection
            mvarsUFcolExcecoesAlt = ""
            lErro = CF("ICMSExcecoes_Le_Estado", sUFDestino, mvarcolExcecoesAlt)
            If lErro <> SUCESSO Then Error 27725
            mvarsUFcolExcecoesAlt = sUFDestino
        End If
        Set colExcecoes = mvarcolExcecoesAlt
    End If
    
    iAchou = 0
    
    'percorrer a colecao de excecoes
    For Each objICMSExcecao In colExcecoes
    
        If sUFOrigem = "" Or Trim(objICMSExcecao.sEstadoOrigem) = "" Or objICMSExcecao.sEstadoOrigem = sUFOrigem Then
            
            If objICMSExcecao.sEstadoDestino = sUFDestino And objICMSExcecao.sCategoriaCliente = "" Then
            
                'se a excecao se aplica a uma categoria de fornecedores para uma categoria de produtos
                If objICMSExcecao.sCategoriaFornecedor <> "" And objICMSExcecao.sCategoriaProduto <> "" Then
                
                    'verifica se a filial fornecedor se enquadra
                    lErro = CategoriaFornecedor_Pesquisar(objICMSExcecao.sCategoriaFornecedor, objICMSExcecao.sCategoriaFornecedorItem, objVarFilFornColCateg.vValor)
                    If lErro <> SUCESSO And lErro <> 27215 Then Error 27224
                    If lErro = SUCESSO Then
                    
                        'verifica se o produto se enquadra
                        lErro = CategoriaProduto_Pesquisar(objICMSExcecao.sCategoriaProduto, objICMSExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27216 Then Error 27225
                                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
                    
                    End If
                                    
                Else
                
                    'se a excecao se aplica a uma categoria de fornecedores
                    If objICMSExcecao.sCategoriaFornecedor <> "" Then
                        'verifica se a filial fornecedor se enquadra
                        lErro = CategoriaFornecedor_Pesquisar(objICMSExcecao.sCategoriaFornecedor, objICMSExcecao.sCategoriaFornecedorItem, objVarFilFornColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27215 Then Error 27222
                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
            
                    Else
                    'se a excecao se aplica a uma categoria de produtos
                    'If objICMSExcecao.sCategoriaProduto <> "" Then
                        'verifica se o produto se enquadra
                        lErro = CategoriaProduto_Pesquisar(objICMSExcecao.sCategoriaProduto, objICMSExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                        If lErro <> SUCESSO And lErro <> 27216 Then Error 27223
                                        
                        If lErro = SUCESSO Then
                            Set objExcecao = objICMSExcecao
                            iAchou = 1
                            Exit For
                        End If
                    End If
                End If
            End If
    
        End If
        
    Next
    
    If iAchou <> 1 Then Set objExcecao = Nothing
    
    ExcecaoForn_Pesquisar = SUCESSO

    Exit Function

Erro_ExcecaoForn_Pesquisar:

    ExcecaoForn_Pesquisar = Err

    Select Case Err
        
        Case 27217 To 27225, 27725

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153865)

    End Select

    Exit Function

End Function

Private Function CategoriaFornecedor_Pesquisar(sCategoriaFornecedor As String, sCategoriaFornecedorItem As String, colCateg As Collection) As Long
'verifica se o par (sCategoriaFornecedor, sCategoriaFornecedorItem está na colecao colCateg
'retorna SUCESSO se achar, 27215 se nao achar ou outro erro
Dim lErro As Long, objFilialFornCategoria As ClassFilialFornCategoria, iAchou As Integer
On Error GoTo Erro_CategoriaFornecedor_Pesquisar

    iAchou = 0
    
    For Each objFilialFornCategoria In colCateg
    
        If objFilialFornCategoria.sCategoria = sCategoriaFornecedor And objFilialFornCategoria.sItem = sCategoriaFornecedorItem Then
        
            iAchou = 1
            Exit For
            
        End If
        
    Next

    If iAchou = 0 Then Error 27215
    
    CategoriaFornecedor_Pesquisar = SUCESSO

    Exit Function

Erro_CategoriaFornecedor_Pesquisar:

    CategoriaFornecedor_Pesquisar = Err

    Select Case Err

        Case 27215 'nao encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153866)

    End Select

    Exit Function

End Function

'Private Function Item_ObterBaseSubstProp(ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, dBaseSusbtProp As Double) As Long
''obter o adicional na base da margem para calculo da subst tributaria
''proporcional ao item devido aos valores de frete, seguro e outras despesas
'
'    'percorrer os complementos (frete, seguro,...) atualizando-os
'    For Each objDocCompl In colDocCompl
'
'        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
'        If lErro <> SUCESSO Then Error 32252
'
'        If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
'        End If
'
'    Next
'
'    Call objDoc.ObterColDocItens(colItens)
'
'    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
'    For Each objDocItem In colItens
'
'            lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
'            If lErro <> SUCESSO Then Error 27324
'
'            lErro = objDocItem.ObterVar("PRODUTO_DESC_GLOBAL", objVarValorDescGlobal)
'            If lErro <> SUCESSO Then Error 27324
'
'            dBaseCalculo = objVar.vValor - objVarValorDescGlobal.vValor
'
'    Next
'
'End Function
'
Private Function AtualizarImposto_Importacao(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'trata nf interna de importacao

Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long

On Error GoTo Erro_AtualizarImposto_Importacao

    'se estiver recalculando marcar valores a nivel de doc c/preench auto.
    If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
    
    Call objDoc.ObterColDocItens(colItens)

    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
        lErro = AtualizarItem_Importacao(objDocItem, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 22720
    Next

    'atualizar o docto como um todo
    lErro = AtualizarDoc_Importacao(objDoc)
    If lErro <> SUCESSO Then Error 22722

    AtualizarImposto_Importacao = SUCESSO

    Exit Function

Erro_AtualizarImposto_Importacao:

    AtualizarImposto_Importacao = Err

    Select Case Err

        Case 22718 To 22722
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153821)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_Importacao(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
'atualiza tipo ICMS, base de calculo, red da base, aliquota, valor ICMS, ICMS subst, aliq ICMS subst, marg subst
Dim lErro As Long, objExcecao As ClassICMSExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribICMS As New ClassTipoTribICMS
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo, objVarBaseSubst As New ClassVariavelCalculo
Dim objVarMargSubst As New ClassVariavelCalculo, objVarAliqSubst As New ClassVariavelCalculo, objVarValorSubst As New ClassVariavelCalculo
Dim objVarValorCredito As New ClassVariavelCalculo, dValorCredito As Double
Dim objVarTipoTribDoc As New ClassVariavelCalculo, objVarTipoDocSigla As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo
Dim objVarAliquotaDI As New ClassVariavelCalculo, objVarRateio As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dPercRedBaseSubst As Double
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, dBaseAux As Double, objVarPercRedBaseSubst As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_Importacao
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dBaseSubst = 0
    '???dBaseSubstReduzida = 0
    dAliqSubst = 0
    dMargSubst = 0
    dValorSubst = 0
    dValorCredito = 0
    dPercRedBaseSubst = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27974
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27978
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27985
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
                
    'tentar achar excecao
    lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, "")
    If lErro <> SUCESSO Then Error 27990
    
    'se achou excecao
    If (Not (objExcecao Is Nothing)) Then
    
        lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, dPercRedBaseSubst, True)
        If lErro <> SUCESSO Then Error 27991
            
    Else 'se NAO achou uma excecao
    
        lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, "EX")
        If lErro <> SUCESSO Then Error 27992
    
    End If
        
    'forçar a aliquota com o que está na DI
    'se ICMS_ALIQUOTA nao tiver sido preenchido manualmente
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        'se ICMS_TIPO permite aliquota <> 0
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            'pegar da DI
            lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_ADICAO_DI", objVarAliquotaDI)
            If lErro <> SUCESSO Then Error 27994
            
            dAliquota = objVarAliquotaDI.vValor
        End If
    End If
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVar)
            If lErro <> SUCESSO Then Error 27994

            dBaseCalculo = objVar.vValor

            'se inclui o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27995
                
                dBaseCalculo = dBaseCalculo + objVar.vValor
                                
            End If
            
            'NA IMPORTACAO O PROPRIO ICMS ENTRA NA BASE
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
                dBaseCalculo = dBaseCalculo / (1 - dAliquota)
            Else
                dBaseCalculo = dBaseCalculo + objVarValor.vValor
            End If
            dBaseCalculo = Arredonda_Moeda(dBaseCalculo, 2)
        
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = (dBaseCalculo * (1 - dPercRedBase)) * dAliquota
        dValor = Arredonda_Moeda(dValor, 4)
    
    Else
        
        dValor = objVarValor.vValor
    
    End If

    objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO

    If objTipoTrib.iICMSCredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27996
    
    AtualizarItem_Importacao = SUCESSO

    Exit Function

Erro_AtualizarItem_Importacao:

    AtualizarItem_Importacao = Err

    Select Case Err

        Case 27974 To 27996
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153855)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_Importacao(objDoc As ClassTributoDoc) As Long
'atualiza base de ICMS, valor de ICMS, base ICMS subst, valor ICMS Subst
'    leva em conta dados dos itens + dados de frete, seg, desp acess.
Dim colItens As Collection, dPercRedBaseSubst As Double
Dim objDocItem As ClassTributoDocItem, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dBase As Double, dBaseSubst As Double, dBaseReduzida As Double, dBaseSubstReduzida As Double, dValorCredito As Double
Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim objVarFreteValor As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
Dim objVarSeguroValor As New ClassVariavelCalculo, objVarDespAcesValor As New ClassVariavelCalculo
Dim lErro As Long, dPercRedBase As Double, objVarICMSValorCredito As New ClassVariavelCalculo
Dim objVarDespImpICMSBase As New ClassVariavelCalculo, objVarDespImpICMSValor As New ClassVariavelCalculo
Dim objVarDespImpICMSCredito As New ClassVariavelCalculo, dBaseImp As Double, dBaseImpReduzida As Double, dAliquotaImp As Double
Dim objVarAliquotaDI As New ClassVariavelCalculo, objVarDespImpValorRateado As New ClassVariavelCalculo
Dim objVarICMSBaseSimples As New ClassVariavelCalculo, dICMSBaseSimples As Double

On Error GoTo Erro_AtualizarDoc_Importacao

    dValor = 0
    dValorSubst = 0
    dBase = 0
    dBaseReduzida = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dValorCredito = 0
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
        
        'obter do item ICMS_BASE, ICMS_VALOR, ICMS_BASE_SUBST, ICMS_VALOR_SUBST
        
        lErro = objDocItem.ObterVar("ICMS_BASE", objVarICMSBase)
        If lErro <> SUCESSO Then Error 27949
        
        lErro = objDocItem.ObterVar("ICMS_RED_BASE", objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27950
         
        dPercRedBase = objVarPercRedBase.vValor
        
        lErro = objDocItem.ObterVar("ICMS_VALOR", objVarICMSValor)
        If lErro <> SUCESSO Then Error 27951
        
        lErro = objDocItem.ObterVar("ICMS_BASE_SUBST", objVarICMSBaseSubst)
        If lErro <> SUCESSO Then Error 27952
        
        lErro = objDocItem.ObterVar("ICMS_VALOR_SUBST", objVarICMSValorSubst)
        If lErro <> SUCESSO Then Error 27953
        
        lErro = objDocItem.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
        If lErro <> SUCESSO Then Error 27954
        
        dBase = dBase + objVarICMSBase.vValor
        dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBase))
        ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
        dBaseSubst = dBaseSubst + objVarICMSBaseSubst.vValor
        dPercRedBaseSubst = 0 '???? ainda nao está sendo armazenado
        dBaseSubstReduzida = dBaseSubstReduzida + (objVarICMSBaseSubst.vValor * (1 - dPercRedBaseSubst))
        ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.00"))
        dValor = dValor + objVarICMSValor.vValor
        dValorSubst = dValorSubst + objVarICMSValorSubst.vValor
        dValorCredito = dValorCredito + objVarICMSValorCredito.vValor
        
        'especifico de importacao
        lErro = objDocItem.ObterVar("DESP_IMP_VALOR_RATEADO", objVarDespImpValorRateado)
        If lErro <> SUCESSO Then Error 27949
        
        lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_ADICAO_DI", objVarAliquotaDI)
        If lErro <> SUCESSO Then Error 27994
        
        dAliquotaImp = objVarAliquotaDI.vValor
        
        dBaseImp = Arredonda_Moeda(objVarDespImpValorRateado.vValor / (1 - dAliquotaImp))
        
        objVarDespImpICMSBase.vValor = dBaseImp
        objVarDespImpICMSValor.vValor = Arredonda_Moeda(dBaseImp * dAliquotaImp)
        objVarDespImpICMSCredito.vValor = objVarDespImpICMSValor.vValor
        
        objVarDespImpICMSBase.sIdentificador = "DESP_IMP_ICMS_BASE"
        objVarDespImpICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarDespImpICMSBase)
        If lErro <> SUCESSO Then Error 27307
        
        objVarDespImpICMSValor.sIdentificador = "DESP_IMP_ICMS_VALOR"
        objVarDespImpICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarDespImpICMSValor)
        If lErro <> SUCESSO Then Error 27307
        
        objVarDespImpICMSCredito.sIdentificador = "DESP_IMP_ICMS_CREDITO"
        objVarDespImpICMSCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarDespImpICMSCredito)
        If lErro <> SUCESSO Then Error 27307
        
        dBase = dBase + objVarDespImpICMSBase.vValor
        dBaseReduzida = dBaseReduzida + objVarDespImpICMSBase.vValor
        dValor = dValor + objVarDespImpICMSValor.vValor
        dValorCredito = dValorCredito + objVarDespImpICMSCredito.vValor
    
    Next
    
    'conta de chegar
    lErro = objDoc.ObterVar("BASE_ICMS_IMPORTACAO_SIMPLES", objVarICMSBaseSimples)
    If lErro <> SUCESSO Then Error 27949
    dICMSBaseSimples = Arredonda_Moeda(objVarICMSBaseSimples.vValor + dValor, 2)
    If dBaseReduzida = dBase And dICMSBaseSimples <> dBase And Abs(dICMSBaseSimples - dBase) <= 0.02 Then
        dBase = dICMSBaseSimples
        dBaseReduzida = dBase
    End If
        
    'arredondar valores globais
    dBase = Arredonda_Moeda(dBase, 2)
    dBaseReduzida = Arredonda_Moeda(dBaseReduzida, 2)
    dBaseSubst = Arredonda_Moeda(dBaseSubst, 2)
    dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida, 2)
    dValor = Arredonda_Moeda(dValor, 2)
    dValorSubst = Arredonda_Moeda(dValorSubst, 2)
    dValorCredito = Arredonda_Moeda(dValorCredito, 2)
        
    lErro = AtualizarDoc_GuardarValores(objDoc, dBaseReduzida, dValor, dBaseSubstReduzida, dValorSubst, dValorCredito)
    If lErro <> SUCESSO Then Error 27956
        
    AtualizarDoc_Importacao = SUCESSO

    Exit Function

Erro_AtualizarDoc_Importacao:

    AtualizarDoc_Importacao = Err

    Select Case Err

        Case 27949 To 27956
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153852)

    End Select

    Exit Function

End Function

Private Function TipoICMSZerar(ByVal iTipoICMS As Integer) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMS
    
        Case ICMS_TIPO_COBR_SUBST_ANT, ICMS_TIPO_NAO_TRIBUTADO, ICMS_TIPO_ISENTO
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    TipoICMSZerar = bRet
    
End Function

Private Function TipoICMSComSubstTrib(ByVal iTipoICMS As Integer) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMS
    
        Case ICMS_TIPO_RED_BASE_E_SUBST, ICMS_TIPO_COM_SUBST_TRIB
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    TipoICMSComSubstTrib = bRet
    
End Function




