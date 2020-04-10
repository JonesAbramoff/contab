VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
Private mvarcolEstados As Collection 'colecao de objEstado
            'colocar como chave a UF
Private mvarcolTiposTribICMS As New Collection 'colecao com os tipos de tributacao p/ICMS

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada

Private mvariSimplificado As Integer 'se a filialEmpresa está no Simples do ICMS
Private mvariInscrita As Integer 'se a filialEmpresa está ou nao inscrita no ICMS
Private mvardAliqInterna As Double 'Aliquota interna do estado da filialEmpresa
Private mvardAliqInternaAtual As Double 'Aliquota interna do estado da filialEmpresa
Private mvardAliqInternaAnt As Double 'Aliquota interna do estado da filialEmpresa
Private mvardAliqInternaFCP As Double
Private mvardAliqInternaFCPAtual As Double
Private mvardAliqInternaFCPAnt As Double
'Private mvardAliqExportacao As Double 'aliquota exportacao da filialEmpresa
Private mvardAliqImportacao As Double 'aliquota importacao da filialEmpresa
Private mvardAliqImportacaoAtual As Double 'aliquota importacao da filialEmpresa
Private mvardAliqImportacaoAnt As Double 'aliquota importacao da filialEmpresa
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
Private mvardtDataRef As Date

Private mvardtDataIniAliqInternaAtual As Date
Private mvardtDataIniAliqImportacaoAtual As Date
Private mvardtDataIniAliqFCPAtual  As Date

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
Dim dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo, objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32022
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 27867
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 27868
    
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then Error 27868
    
    dtDataEmissao = objVarDataEmissao.vValor
    
    lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
    If lErro <> SUCESSO Then Error 27868
    
    dtDataEntrada = objVarDataEntrada.vValor
    
    Select Case objTipoDocInfo.iTipoTrib
        Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
            mvardtDataRef = dtDataEntrada
            
        Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
            mvardtDataRef = dtDataEmissao
    End Select
    
    If mvardtDataIniAliqInternaAtual <> DATA_NULA And mvardtDataIniAliqInternaAtual > mvardtDataRef Then
        mvardAliqInterna = mvardAliqInternaAnt
    Else
        mvardAliqInterna = mvardAliqInternaAtual
    End If
    
    If mvardtDataIniAliqFCPAtual <> DATA_NULA And mvardtDataIniAliqFCPAtual > mvardtDataRef Then
        mvardAliqInternaFCP = mvardAliqInternaFCPAnt
    Else
        mvardAliqInternaFCP = mvardAliqInternaFCPAtual
    End If
    
    If mvardtDataIniAliqImportacaoAtual <> DATA_NULA And mvardtDataIniAliqImportacaoAtual > mvardtDataRef Then
        mvardAliqImportacao = mvardAliqImportacaoAnt
    Else
        mvardAliqImportacao = mvardAliqImportacaoAtual
    End If

    mvariZerarTributo = 0
    mvariZerarTributoST = 0
    If mvariInscrita <> ICMS_INSCRITA Or mvariSimplificado <> 0 Then
        mvariZerarTributo = 1
        If mvariInscrita <> ICMS_INSCRITA Then mvariZerarTributoST = 1
    Else
        If mvariSimplesNacional <> 0 Then
            
            If mvardtDataRef >= DATA_INICIO_SIMPLES_NACIONAL Then mvariZerarTributo = 1
            
        End If
        
    End If
    
    mvariNotaInterna = IIf(objTipoDocInfo.iEmitente = EMITENTE_EMPRESA, 1, 0)
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC Then
    
        lErro = AtualizarImposto_Importacao(objDoc, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27870
    
    Else
    
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_CLI
                If objTipoDocInfo.iCodigo = DOCINFO_NFIEDV Then
                    lErro = AtualizarImposto_SaidaCli(objDoc, iRecalculaAutomatica, objTipoDocInfo.iCodigo)
                Else
                    lErro = AtualizarImposto_EntradaCli(objDoc, iRecalculaAutomatica)
                End If
                If lErro <> SUCESSO Then Error 27869
                
            Case TRIB_ENTRADA_FORN
                lErro = AtualizarImposto_EntradaForn(objDoc, iRecalculaAutomatica)
                If lErro <> SUCESSO Then Error 27870
            
            Case TRIB_SAIDA_CLI
                lErro = AtualizarImposto_SaidaCli(objDoc, iRecalculaAutomatica, objTipoDocInfo.iCodigo)
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

'FUNCOES AUXILIARES:
'===================
Private Function AtualizarImposto_SaidaCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, ByVal iTipoNFiscal As Integer) As Long

Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objTRBConfig As New ClassTRBConfig
Dim lErro As Long
Dim objVarNatOpDoc As New ClassVariavelCalculo, objVarNatOp As ClassVariavelCalculo, sNatOpDoc As String
Dim bAtualiza As Boolean

On Error GoTo Erro_AtualizarImposto_SaidaCli

    lErro = PrepararAtualizacao_SaidaCli(objDoc)
    If lErro <> SUCESSO And lErro <> 27720 Then Error 27291
    
    'se nao houve falta de dados
    If lErro <> 27720 Then
        
        'se estiver recalculando marcar valores a nivel de doc c/preench auto.
        If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
        
        Call objDoc.ObterColDocItens(colItens)
        
        'percorrer os itens atualizando-os
        For Each objDocItem In colItens
        
            bAtualiza = True
            If gobjCRFAT.iOtimizaTrib = MARCADO Then
                If objDocItem.iJaCalculado = MARCADO Then bAtualiza = False
            End If
            
            If bAtualiza Then
                lErro = AtualizarItem_SaidaCli(objDoc, objDocItem, iRecalculaAutomatica, iTipoNFiscal)
                If lErro <> SUCESSO Then Error 27292
            End If
        
        Next
        
        'ajustar cfop principal do doc caso seja manual e todos os itens tenham o mesmo cfop
        lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOpDoc)
        If lErro <> SUCESSO Then Error 27292
        If objVarNatOpDoc.iStatus = VAR_PREENCH_AUTOMATICO Then
        
            sNatOpDoc = ""
            For Each objDocItem In colItens
            
                lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                If lErro <> SUCESSO Then Error 27292
                
                If objVarNatOp.vValor <> "" Then
                    If sNatOpDoc = "" Then
                        sNatOpDoc = objVarNatOp.vValor
                    Else
                        If sNatOpDoc <> objVarNatOp.vValor Then
                            sNatOpDoc = ""
                            Exit For
                        End If
                    End If
                End If
                
            Next
        
            If sNatOpDoc <> "" Then
                
                objVarNatOpDoc.vValor = sNatOpDoc
                lErro = objDoc.GuardarVar(objVarNatOpDoc)
                If lErro <> SUCESSO Then Error 27292
                
            End If
            
        End If
        
        'atualizar o docto como um todo
        lErro = AtualizarDoc(objDoc)
        If lErro <> SUCESSO Then Error 27293
    
    End If
    
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

    lErro = PrepararAtualizacao_EntradaCli(objDoc)
    If lErro <> SUCESSO And lErro <> 27882 Then Error 27957

    'se nao houve falta de dados
    If lErro <> 27882 Then

        'se estiver recalculando marcar valores a nivel de doc c/preench auto.
        If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)

        Call objDoc.ObterColDocItens(colItens)

        'percorrer os itens atualizando-os
        For Each objDocItem In colItens
            lErro = AtualizarItem_EntradaCli(objDoc, objDocItem, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27959
        Next

        'atualizar o docto como um todo
        lErro = AtualizarDoc(objDoc)
        If lErro <> SUCESSO Then Error 27961

    End If
   
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

    lErro = PrepararAtualizacao_Forn(objDoc)
    If lErro <> SUCESSO And lErro <> 27973 Then Error 22723

    'se nao houve falta de dados
    If lErro <> 27973 Then

        'se estiver recalculando marcar valores a nivel de doc c/preench auto.
        If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
        
        Call objDoc.ObterColDocItens(colItens)

        'percorrer os itens atualizando-os
        For Each objDocItem In colItens
            lErro = AtualizarItem_SaidaForn(objDoc, objDocItem, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22725
        Next

        'atualizar o docto como um todo
        lErro = AtualizarDoc(objDoc)
        If lErro <> SUCESSO Then Error 22727

    End If
    
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

    lErro = PrepararAtualizacao_Forn(objDoc)
    If lErro <> SUCESSO And lErro <> 27973 Then Error 22718

    'se nao houve falta de dados
    If lErro <> 27973 Then

        'se estiver recalculando marcar valores a nivel de doc c/preench auto.
        If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
        
        Call objDoc.ObterColDocItens(colItens)

        'percorrer os itens atualizando-os
        For Each objDocItem In colItens
            lErro = AtualizarItem_EntradaForn(objDoc, objDocItem, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22720
        Next

        'atualizar o docto como um todo
        lErro = AtualizarDoc(objDoc)
        If lErro <> SUCESSO Then Error 22722

    End If
    
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

Private Function ObterAliquotaBasicaSaidaCli(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFDestino As String, ByVal objDocItem As ClassTributoDocItem, dAliquotaFCP As Double) As Long
'preenche a aliquota
Dim lErro As Long
Dim objFilCliInscEst As New ClassVariavelCalculo
Dim objFilCliIEIsento As New ClassVariavelCalculo

On Error GoTo Erro_ObterAliquotaBasicaSaidaCli

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFDestino = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            dAliquotaFCP = mvardAliqInternaFCP
        Else
          
            'se for exportacao
            If sUFDestino = "EX" Then
            
                'setar aliquota igual a zero ( exportacao nao tem ICMS )
                dAliquota = 0
                dAliquotaFCP = 0
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
                If lErro <> SUCESSO Then Error 27284
                
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_IEISENTO", objFilCliIEIsento)
                If lErro <> SUCESSO Then Error 27904
                
                'se filial cliente nao inscrita ICMS
                If objFilCliIEIsento.vValor = MARCADO Then
                    
                    'usar aliquota interna
                    dAliquota = mvardAliqInterna
                    dAliquotaFCP = mvardAliqInternaFCP
                    
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstSaida_Pesquisar(sUFDestino, dAliquota, objDocItem, dAliquotaFCP)
                    If lErro <> SUCESSO Then Error 27285
                
                End If
            End If
        End If
    Else
        dAliquota = 0
        dAliquotaFCP = 0
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

Private Function AtualizarItem_ObterValoresAtuais(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer, objVarModalidade As ClassVariavelCalculo, objVarModalidadeST As ClassVariavelCalculo, objVarVlrCredSimples As ClassVariavelCalculo, objVarAliCredSimples As ClassVariavelCalculo, objVarTipoICMSSimples As ClassVariavelCalculo, _
    objVarCobrAntValor As ClassVariavelCalculo, objVarCobrAntBase As ClassVariavelCalculo, objVarPBOP As ClassVariavelCalculo, objVarUFD As ClassVariavelCalculo, objVarBCSTRet As ClassVariavelCalculo, objVarICMSSTRet As ClassVariavelCalculo, objVarBCSTDest As ClassVariavelCalculo, objVarICMSSTDest As ClassVariavelCalculo, objVarVI As ClassVariavelCalculo, objVarMotivo As ClassVariavelCalculo, objVarOrigemProd As ClassVariavelCalculo, objVarFCI As ClassVariavelCalculo, objVar51ValorOp As ClassVariavelCalculo, objVarPercDifer As ClassVariavelCalculo, objVarValorDif As ClassVariavelCalculo _
    , objVarInterestBCUFDest As ClassVariavelCalculo, objVarInterestPercFCPUFDest As ClassVariavelCalculo, objVarInterestAliqUFDest As ClassVariavelCalculo, objVarInterestAliq As ClassVariavelCalculo, objVarInterestPercPartilha As ClassVariavelCalculo, objVarInterestVlrFCPUFDest As ClassVariavelCalculo, objVarInterestVlrUFDest As ClassVariavelCalculo, objVarInterestVlrUFRemet As ClassVariavelCalculo, _
    objVarICMSInterestBCFCPUFDest As ClassVariavelCalculo, objVarICMSvBCFCP As ClassVariavelCalculo, objVarICMSpFCP As ClassVariavelCalculo, objVarICMSvFCP As ClassVariavelCalculo, objVarICMSvBCFCPST As ClassVariavelCalculo, objVarICMSpFCPST As ClassVariavelCalculo, objVarICMSvFCPST As ClassVariavelCalculo, objVarICMSvBCFCPSTRet As ClassVariavelCalculo, objVarICMSpFCPSTRet As ClassVariavelCalculo, objVarICMSvFCPSTRet As ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As ClassVariavelCalculo, objVarICMSSTBaseDupla As ClassVariavelCalculo, objVarICMSInterestBaseDupla As ClassVariavelCalculo, objVarcBenef As ClassVariavelCalculo) As Long
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
            
    lErro = objDocItem.ObterVar("ICMS_BASE_MODALIDADE", objVarModalidade)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarModalidade.iStatus = VAR_PREENCH_AUTOMATICO
    If objVarModalidade.iStatus <> VAR_PREENCH_MANUAL Then
        objVarModalidade.vValor = ICMS_MODALIDADE_VALOR
    End If
    
    lErro = objDocItem.ObterVar("ICMS_ST_BASE_MODALIDADE", objVarModalidadeST)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarModalidadeST.iStatus = VAR_PREENCH_AUTOMATICO
    If objVarModalidadeST.iStatus <> VAR_PREENCH_MANUAL Then
        objVarModalidadeST.vValor = ICMS_ST_MODALIDADE_MARGEM
    End If
    
    lErro = objDocItem.ObterVar("ICMS_VALOR_CREDITO_SIMPLES", objVarVlrCredSimples)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_CREDITO_SIMPLES", objVarAliCredSimples)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_TIPO_SIMPLES", objVarTipoICMSSimples)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_PERC_BASE_OPER_PROP", objVarPBOP)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ST_UF_DEVIDO", objVarUFD)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ST_BASE_RETIDO_UF", objVarBCSTRet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarBCSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ST_VALOR_RETIDO_UF", objVarICMSSTRet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ST_BASE_UF_DESTINO", objVarBCSTDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarBCSTDest.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ST_VALOR_UF_DESTINO", objVarICMSSTDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSSTDest.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_VALOR_ISENTO", objVarVI)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_MOTIVO_ISENTO", objVarMotivo)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_BASE_ST_COBR_ANT", objVarCobrAntBase)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_VALOR_ST_COBR_ANT", objVarCobrAntValor)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
         
    lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
         
    lErro = objDocItem.ObterVar("FCI", objVarFCI)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarFCI.iStatus = VAR_PREENCH_AUTOMATICO
         
    lErro = objDocItem.ObterVar("ICMS_51_VALOR_OP", objVar51ValorOp)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_PERC_DIFER", objVarPercDifer)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_VALOR_DIF", objVarValorDif)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
               
    'NT 2015.003 (ICMS Interestadual)
    lErro = objDocItem.ObterVar("ICMS_INTEREST_BC_UF_DEST", objVarInterestBCUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestBCUFDest.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_PERC_FCP_UF_DEST", objVarInterestPercFCPUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestPercFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_ALIQ_UF_DEST", objVarInterestAliqUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestAliqUFDest.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_ALIQ", objVarInterestAliq)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestAliq.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_PERC_PARTILHA", objVarInterestPercPartilha)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestPercPartilha.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_FCP_UF_DEST", objVarInterestVlrFCPUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestVlrFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_UF_DEST", objVarInterestVlrUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestVlrUFDest.iStatus = VAR_PREENCH_AUTOMATICO
               
    lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_UF_REMET", objVarInterestVlrUFRemet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarInterestVlrUFRemet.iStatus = VAR_PREENCH_AUTOMATICO
    
    'NFe 4.00
    lErro = objDocItem.ObterVar("ICMS_INTEREST_BC_FCP_UF_DEST", objVarICMSInterestBCFCPUFDest)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSInterestBCFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_BC_FCP", objVarICMSvBCFCP)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ALIQ_FCP", objVarICMSpFCP)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_VLR_FCP", objVarICMSvFCP)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_BC_FCP_ST", objVarICMSvBCFCPST)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ALIQ_FCP_ST", objVarICMSpFCPST)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST", objVarICMSvFCPST)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_BC_FCP_ST_RET", objVarICMSvBCFCPSTRet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_ALIQ_FCP_ST_RET", objVarICMSpFCPSTRet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSpFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST_RET", objVarICMSvFCPSTRet)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
  
    lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_ST_COBR_ANT", objVarICMSSTCobrAntAliquota)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
       
    lErro = objDocItem.ObterVar("ICMS_ST_BASE_DUPLA", objVarICMSSTBaseDupla)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSSTBaseDupla.iStatus = VAR_PREENCH_AUTOMATICO
       
    lErro = objDocItem.ObterVar("ICMS_INTEREST_BASE_DUPLA", objVarICMSInterestBaseDupla)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarICMSInterestBaseDupla.iStatus = VAR_PREENCH_AUTOMATICO
       
    lErro = objDocItem.ObterVar("ICMS_CBENEF", objVarcBenef)
    If lErro <> SUCESSO Then Error 27876
    If iRecalculaAutomatica = 1 Then objVarcBenef.iStatus = VAR_PREENCH_AUTOMATICO
    
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

Private Function AtualizarItem_GuardarValores(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, objVarModalidade As ClassVariavelCalculo, objVarModalidadeST As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double, dPercRedBaseSubst As Double, dValorCredito As Double, iModalidade As Integer, iModalidadeST As Integer, objVarVlrCredSimples As ClassVariavelCalculo, objVarAliCredSimples As ClassVariavelCalculo, objVarTipoICMSSimples As ClassVariavelCalculo, _
        objVarCobrAntValor As ClassVariavelCalculo, objVarCobrAntBase As ClassVariavelCalculo, objVarPBOP As ClassVariavelCalculo, objVarUFD As ClassVariavelCalculo, objVarBCSTRet As ClassVariavelCalculo, objVarICMSSTRet As ClassVariavelCalculo, objVarBCSTDest As ClassVariavelCalculo, objVarICMSSTDest As ClassVariavelCalculo, objVarVI As ClassVariavelCalculo, objVarMotivo As ClassVariavelCalculo, objVarRegTrib As ClassVariavelCalculo, objVarOrigemProd As ClassVariavelCalculo, objVarFCI As ClassVariavelCalculo, objVar51ValorOp As ClassVariavelCalculo, objVarPercDifer As ClassVariavelCalculo, objVarValorDif As ClassVariavelCalculo, d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, iMotivo As Integer) As Integer

'atualiza as variaveis que tiveram seus valores recalculados pelo sistema
Dim lErro As Long
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objTipoTribICMS As New ClassTipoTribICMS
Dim objVarCSOSN As New ClassVariavelCalculo, objVarCST As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo

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
    
    If objVarModalidade.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarModalidade.vValor = iModalidade
        objVarModalidade.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarModalidade)
        If lErro <> SUCESSO Then Error 27874
        
    End If
    
    If objVarModalidadeST.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarModalidadeST.vValor = iModalidadeST
        objVarModalidadeST.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarModalidadeST)
        If lErro <> SUCESSO Then Error 27874
        
    End If
    
    If objVarVlrCredSimples.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarVlrCredSimples)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarAliCredSimples.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarAliCredSimples)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarTipoICMSSimples)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarPBOP.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarPBOP)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarUFD.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarUFD)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarBCSTRet.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarBCSTRet)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarICMSSTRet.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarICMSSTRet)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarBCSTDest.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarBCSTDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarICMSSTDest.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarICMSSTDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarVI.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarVI)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarMotivo.iStatus <> VAR_PREENCH_MANUAL Then
        objVarMotivo.vValor = iMotivo
        objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarMotivo)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarCobrAntBase.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarCobrAntBase)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarCobrAntValor.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarCobrAntValor)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarOrigemProd)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarFCI.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = objDocItem.GuardarVar(objVarFCI)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
        objVar51ValorOp.vValor = d51ValorOp
        objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVar51ValorOp)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
        objVarPercDifer.vValor = dPercDifer
        objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarPercDifer)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
        objVarValorDif.vValor = dValorDif
        objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValorDif)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    'Código para gravar o CST e o CSOSN
    objTipoTribICMS.iTipo = objVarTipo.vValor
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27874
                    
    objVarCST.sIdentificador = "ICMS_CST"
    objVarCST.vValor = CStr(objVarOrigemProd.vValor) & Format(objTipoTribICMS.iTipoTribCST, "00")
    objVarCST.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVarCST)
    If lErro <> SUCESSO Then Error 27874

    If objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES Then
            
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        
        'só para nao dar erro de vb
        If objTipoICMSSimples.iTipo <> 0 Then
        
            Set objVarItemTipoObj.vValor = objTipoICMSSimples
            lErro = objDocItem.objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
            If lErro <> SUCESSO Then Error 27874
    
            Set objTipoICMSSimples = objVarItemTipoObj.vValor
            
            objVarCSOSN.sIdentificador = "ICMS_CSOSN"
            objVarCSOSN.vValor = CStr(objVarOrigemProd.vValor) & Format(objTipoICMSSimples.iCSOSN, "000")
            objVarCSOSN.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDocItem.GuardarVar(objVarCSOSN)
            If lErro <> SUCESSO Then Error 27874
        
        Else
        
            objVarCSOSN.sIdentificador = "ICMS_CSOSN"
            objVarCSOSN.vValor = CStr(objVarOrigemProd.vValor) & Format(0, "000")
            objVarCSOSN.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDocItem.GuardarVar(objVarCSOSN)
            If lErro <> SUCESSO Then Error 27874
        
        End If
        
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

Private Function AtualizarItem_TratarExcecao(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, objVarMargSubst As ClassVariavelCalculo, dMargSubst As Double, objVarPercRedBaseSubst As ClassVariavelCalculo, dPercRedBaseSubst As Double, ByVal bSemSubstTributaria As Boolean, iUsaPauta As Integer, dValorPauta As Double, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo, iICMSSTBaseDupla As Integer, objVarICMSSTBaseDupla As ClassVariavelCalculo, ByRef iMotivo As Integer, ByRef objVarMotivo As ClassVariavelCalculo, ByRef scBenef As String, ByRef objVarcBenef As ClassVariavelCalculo) As Long

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
    
    If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
        'se ICMS_TIPO permite aliquota <> 0
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            'ICMS_ALIQUOTA = aliquota da excecao
            dICMSpFCP = objExcecao.dICMSPercFCP
        Else
            dICMSpFCP = 0
        End If
    Else
        dICMSpFCP = objVarAliquota.vValor
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = objExcecao.dPercRedBaseCalculo
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
    
    'se ICMS_TIPO permite subst
    If objTipoTribICMS.iPermiteMargLucro And (bSemSubstTributaria = False) Then
        If objVarPercRedBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBaseSubst = objExcecao.dPercRedBaseCalculoSubst
        Else
            dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        End If
        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dMargSubst = objExcecao.dPercMargemLucro
        Else
            dMargSubst = objVarMargSubst.vValor
        End If
        dValorPauta = objExcecao.dValorPauta
        iUsaPauta = objExcecao.iUsaPauta
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
        dValorPauta = 0
        iUsaPauta = 0
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        If objExcecao.iICMSSTBaseDupla = MARCADO And mvardtDataRef >= objExcecao.dtICMSSTBaseDuplaIni Then
            iICMSSTBaseDupla = MARCADO
        Else
            iICMSSTBaseDupla = DESMARCADO
        End If
    Else
        iICMSSTBaseDupla = objVarICMSSTBaseDupla.vValor
    End If
     
    If objVarMotivo.iStatus <> VAR_PREENCH_MANUAL Then
       iMotivo = objExcecao.iICMSMotivo
    Else
       iMotivo = objVarMotivo.vValor
    End If
     
    If objVarcBenef.iStatus <> VAR_PREENCH_MANUAL Then
       scBenef = objExcecao.scBenef
    Else
       scBenef = objVarcBenef.vValor
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

Private Function AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFDestino As String, ByVal bPodeUsarProt21 As Boolean, ByVal sUFCli As String, bProt21 As Boolean, dAliqProt21 As Double, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo) As Long

Dim lErro As Long, dAliquotaDestino As Double, objVarCliCodigo As New ClassVariavelCalculo
Dim objVarNaoCalcST As New ClassVariavelCalculo, lComando As Long, lComando2 As Long, iTipoCliente As Integer
Dim dAliquotaFCP As Double

On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoSaidaCli
                    
    'para a TA nao aceitar o prot 21 para clientes do tipo 80
    If UCase(gsNomeEmpresa) = UCase("IMI HYDRONIC ENGENHARIA LTDA") And bPodeUsarProt21 Then
    
        lComando2 = Comando_Abrir()
        If lComando2 = 0 Then Error 27317
    
        lErro = objDocItem.objDoc.ObterVar("CLIENTE_CODIGO", objVarCliCodigo)
        If lErro <> SUCESSO Then Error 27317
    
        If objVarCliCodigo.iStatus <> VAR_PREENCH_VAZIO Then
        
            lErro = Comando_Executar(lComando2, "SELECT Tipo FROM Clientes WHERE Codigo = ?", iTipoCliente, StrParaLong(objVarCliCodigo.vValor))
            If lErro <> AD_SQL_SUCESSO Then Error 27317
            
            lErro = Comando_BuscarProximo(lComando2)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27317
            
            If lErro = AD_SQL_SUCESSO Then
                
                If iTipoCliente = 80 Then bPodeUsarProt21 = False
            
            End If
            
        End If
        
        Call Comando_Fechar(lComando2)
    
    End If
                
    If bPodeUsarProt21 Then
    
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 27317
        
    End If
    
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        lErro = objDocItem.objDoc.ObterVar("ICMS_ST_NAO_CALCULAR", objVarNaoCalcST)
        If lErro <> SUCESSO Then Error 27317

        If objVarNaoCalcST.vValor <> MARCADO Then
            iTipo = objTipoTrib.iICMSTipo
        Else
            iTipo = ICMS_TIPO_NORMAL
        End If
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27317
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Or objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaSaidaCli(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFDestino, objDocItem, dAliquotaFCP)
            If lErro <> SUCESSO Then Error 27318
        Else
            dAliquota = 0
            dAliquotaFCP = 0
        End If
'    Else
'
'        dAliquota = objVarAliquota.vValor
        
    End If
    
    dICMSpFCP = dAliquotaFCP
    
    If objVarAliquota.iStatus = VAR_PREENCH_MANUAL Then dAliquota = objVarAliquota.vValor
    If objVarICMSpFCP.iStatus = VAR_PREENCH_MANUAL Then dICMSpFCP = objVarICMSpFCP.vValor
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    If bPodeUsarProt21 And iTipo = ICMS_TIPO_NORMAL And objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        lErro = Comando_Executar(lComando, "SELECT ICMSAliquotaInterna FROM Estados WHERE Sigla = ? AND ICMSProt21Ini <> ? AND ICMSProt21Ini <= ? AND (ICMSProt21Fim = ? OR ICMSProt21Fim >= ?)", _
            dAliquotaDestino, sUFCli, DATA_NULA, mvardtDataRef, DATA_NULA, mvardtDataRef)
        If lErro <> AD_SQL_SUCESSO Then Error 27318
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27318
        
        If lErro = AD_SQL_SUCESSO Then
            
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
                dAliquota = mvarcolAliqInterstSaida.Item(sUFCli).dAliquota
            
            End If
            
            iTipo = ICMS_TIPO_COM_SUBST_TRIB
        
            objTipoTribICMS.iTipo = iTipo
            lErro = TipoICMS_Obter(objTipoTribICMS)
            If lErro <> SUCESSO Then Error 27317
    
            dAliqProt21 = dAliquotaDestino
            bProt21 = True
        
        End If
   
    End If
        
    If bPodeUsarProt21 Then Call Comando_Fechar(lComando)
    
    AtualizarItem_TratarNaoExcecaoSaidaCli = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoSaidaCli:

    AtualizarItem_TratarNaoExcecaoSaidaCli = Err

    Select Case Err

        Case 27317, 27318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153829)

    End Select

    If bPodeUsarProt21 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function AtualizarItem_SaidaCli(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer, ByVal iTipoNFiscal As Integer) As Long
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
Dim objVarModalidade As New ClassVariavelCalculo, objVarModalidadeST As New ClassVariavelCalculo
Dim iModalidade As Integer, iModalidadeST As Integer, objVarProdOutrosValores As ClassVariavelCalculo
Dim dICMSProprioSTSimples As Double, dAliquotaPadrao As Double
Dim objVarNatOp As New ClassVariavelCalculo, objVarRegTrib As New ClassVariavelCalculo
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objVarEmissao As New ClassVariavelCalculo, objDASAliquotas As New ClassDASAliquotas
Dim objVarTipoICMSSimples As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo, objVarVlrCredSimples As New ClassVariavelCalculo, objVarAliCredSimples As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, objVarProdVlr As New ClassVariavelCalculo
Dim objVarPBOP As New ClassVariavelCalculo, objVarUFD As New ClassVariavelCalculo, objVarBCSTRet As New ClassVariavelCalculo
Dim objVarICMSSTRet As New ClassVariavelCalculo, objVarBCSTDest As New ClassVariavelCalculo, objVarICMSSTDest As New ClassVariavelCalculo
Dim objVarVI As New ClassVariavelCalculo, objVarMotivo As New ClassVariavelCalculo, objVarCobrAntValor As New ClassVariavelCalculo, objVarCobrAntBase As New ClassVariavelCalculo
Dim dVI As Double, iMotivo As Integer, objFilCliIEIsento As New ClassVariavelCalculo, dVCA As Double, dBCA As Double
Dim objVarOrigemProd As New ClassVariavelCalculo, iZerarAliquota As Integer, iZerarST As Integer
Dim objFilial As ClassFilialCliente, objVarFilCliObj As New ClassVariavelCalculo
Dim dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo, iSTCobrAnt As Integer
Dim bProt21 As Boolean, dAliqProt21 As Double, objVarItemNFOrig As New ClassVariavelCalculo, objItemNFOrig As ClassItemNF
Dim dMargSubstOrig As Double, dValorPautaOrig As Double, objVarFCI As New ClassVariavelCalculo
Dim objVar51ValorOp As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo, objVarValorDif As New ClassVariavelCalculo
Dim d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, bDiferimentoParcial As Boolean

'NT 2015.003 (ICMS Interestadual)
Dim objVarInterestBCUFDest As New ClassVariavelCalculo, objVarInterestPercFCPUFDest As New ClassVariavelCalculo, objVarInterestAliqUFDest As New ClassVariavelCalculo, objVarInterestAliq As New ClassVariavelCalculo
Dim objVarInterestPercPartilha As New ClassVariavelCalculo, objVarInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarInterestVlrUFDest As New ClassVariavelCalculo, objVarInterestVlrUFRemet As New ClassVariavelCalculo
Dim dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double
Dim dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double
Dim objVarIndConsumidorFinal As New ClassVariavelCalculo, bConsumidorFinal As Boolean, bEC872015 As Boolean, dDifal As Double
Dim dBC_EC872015 As Double, dPercRedBC_EC872015 As Double
Dim dAux1 As Double, dAux2 As Double, objVarTipoDocInfo As New ClassVariavelCalculo
Dim sCFOPItemNFOrig As String

'NFe 4.00
Dim objVarICMSInterestBCFCPUFDest As New ClassVariavelCalculo, objVarICMSvBCFCPST As New ClassVariavelCalculo, objVarICMSpFCPST As New ClassVariavelCalculo
Dim objVarICMSvFCPST As New ClassVariavelCalculo, objVarICMSvBCFCPSTRet As New ClassVariavelCalculo, objVarICMSpFCPSTRet  As New ClassVariavelCalculo
Dim objVarICMSvFCPSTRet As New ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As New ClassVariavelCalculo, objVarICMSvBCFCP As New ClassVariavelCalculo
Dim objVarICMSpFCP As New ClassVariavelCalculo, objVarICMSvFCP As New ClassVariavelCalculo
Dim dICMSInterestBCFCPUFDest As Double, dICMSvBCFCPST As Double, dICMSpFCPST As Double
Dim dICMSvFCPST As Double, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double
Dim dICMSvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dICMSvBCFCP As Double
Dim dICMSpFCP As Double, dICMSvFCP As Double, dAliquotaPadraoFCP As Double, dAliquotaInterestadualFCP As Double

'Base Dupla
Dim objVarICMSSTBaseDupla As New ClassVariavelCalculo, objVarICMSInterestBaseDupla As New ClassVariavelCalculo
Dim iICMSSTBaseDupla As Integer, iICMSInterestBaseDupla As Integer
Dim objVarcBenef As New ClassVariavelCalculo, scBenef As String
Dim objVarTipoAux As New ClassVariavelCalculo, iTipoAux As Integer
Dim objTipoTribICMSAux As New ClassTipoTribICMS

On Error GoTo Erro_AtualizarItem_SaidaCli
    
    bProt21 = False
    dAliqProt21 = 0
    bDiferimentoParcial = False
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    
    dICMSProprioSTSimples = 0
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
    iModalidade = 0
    iModalidadeST = 0
    bExcecaoComSubstTrib = False
    
    iZerarAliquota = 0
    iZerarST = 0
    
    dMargSubstOrig = 0
    dValorPautaOrig = 0
    
    'NT 2015.003 (ICMS Interestadual)
    bConsumidorFinal = False
    bEC872015 = False
    dInterestBCUFDest = 0
    dInterestPercFCPUFDest = 0
    dInterestAliqUFDest = 0
    dInterestAliq = 0
    dInterestPercPartilha = 0
    dInterestVlrFCPUFDest = 0
    dInterestVlrUFDest = 0
    dInterestVlrUFRemet = 0
    dBC_EC872015 = 0
    dPercRedBC_EC872015 = 0
    
    'NFe 4.00
    dICMSInterestBCFCPUFDest = 0
    dICMSvBCFCPST = 0
    dICMSpFCPST = 0
    dICMSvFCPST = 0
    dICMSvBCFCPSTRet = 0
    dICMSpFCPSTRet = 0
    dICMSvFCPSTRet = 0
    dICMSSTCobrAntAliquota = 0
    dICMSvBCFCP = 0
    dICMSpFCP = 0
    dICMSvFCP = 0
    scBenef = ""
       
    If iTipoNFiscal = DOCINFO_NFIEDV Then
    
        lErro = objDocItem.ObterVar("OBJ_ITEM_NF_ORIG", objVarItemNFOrig)
        If lErro <> SUCESSO Then Error 27721
        
        Set objItemNFOrig = objVarItemNFOrig.vValor
        
        If Not (objItemNFOrig Is Nothing) Then
        
            dMargSubstOrig = objItemNFOrig.objTributacaoItemNF.dICMSSubstPercMVA
            If objItemNFOrig.objTributacaoItemNF.dQtdTrib <> 0 Then
                dValorPautaOrig = objItemNFOrig.objTributacaoItemNF.dICMSSubstBase / objItemNFOrig.objTributacaoItemNF.dQtdTrib
            End If
        
        End If
    
    End If
                                   
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, objVarICMSInterestBaseDupla, objVarcBenef)
    If lErro <> SUCESSO Then Error 27319
    
    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27319
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarOrigemProd.vValor = objProduto.iOrigemMercadoria
            objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDocItem.GuardarVar(objVarOrigemProd)
            If lErro <> SUCESSO Then Error 27319
        
        End If
        
    End If
    
    If objVarFCI.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27319
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarFCI.vValor = objProduto.sFCI
            objVarFCI.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    lErro = objDocItem.ObterVar("REGIME_TRIBUTARIO", objVarRegTrib)
    If lErro <> SUCESSO Then Error 27320
    If iRecalculaAutomatica = 1 Then objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
    
    If objVarRegTrib.iStatus <> VAR_PREENCH_MANUAL Then
    
        'Se não é do Simples
        If mvariSimplesNacional = DESMARCADO Then
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_NORMAL
        Else
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES
        End If
        objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
        
        lErro = objDocItem.GuardarVar(objVarRegTrib)
        If lErro <> SUCESSO Then Error 27320
    
    End If
       
    'se a filialEmpresa nao está inscrita no ICMS
    If mvariZerarTributo <> 0 And mvariZerarTributoST <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                        objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    
    End If
       
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27320
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27373
    
    'NFe 4.00
    dICMSpFCP = objVarICMSpFCP.vValor
    dICMSpFCPST = objVarICMSpFCPST.vValor
    dICMSpFCPSTRet = objVarICMSpFCPSTRet.vValor
    'NFe 4.00
        
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                        objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And (TipoICMSZerar(objVarTipo.vValor, iTipoNFiscal)) Then
        iTipo = objVarTipo.vValor
        dVI = objVarVI.vValor
        dVCA = objVarCobrAntValor.vValor
        dBCA = objVarCobrAntBase.vValor
        dICMSSTCobrAntAliquota = objVarICMSSTCobrAntAliquota.vValor
        If objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_MANUAL Then
            dVCA = Arredonda_Moeda(dBCA * dICMSSTCobrAntAliquota)
        End If
        iMotivo = objVarMotivo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                        objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, iTipo = ICMS_TIPO_COBR_SUBST_ANT)
        'Manter campos próprios desses tipos de tributação
        If iTipo = ICMS_TIPO_COBR_SUBST_ANT Then
            objVarCobrAntBase.vValor = dBCA
            objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
            objVarCobrAntValor.vValor = dVCA
        
            'NFe 4.00
            If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvBCFCPSTRet = dBCA
                objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvBCFCPSTRet = objVarICMSvBCFCPST.vValor
            End If
            
            If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvFCPSTRet = Arredonda_Moeda(dICMSvBCFCPSTRet * dICMSpFCPSTRet)
                objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvFCPSTRet = objVarICMSvFCPSTRet.vValor
            End If
            'NFe 4.00
        
        Else
            objVarVI.vValor = dVI
            objVarMotivo.vValor = iMotivo
        End If
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    iModalidade = objVarModalidade.vValor
    iModalidadeST = objVarModalidadeST.vValor
        
    If iTipoNFiscal <> DOCINFO_NFISSAT And iTipoNFiscal <> DOCINFO_TNFISFV Then
    
        'obter UF Destino
        lErro = objDocItem.objDoc.ObterVar("CLIENTE_UF", objVarUFDestino)
        If lErro <> SUCESSO Then Error 27721
        
        sUFDestino = objVarUFDestino.vValor
        If sUFDestino = "" Then sUFDestino = mvarsUFFilialEmpresa
    
        lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
        If lErro <> SUCESSO Then Error 27722
        
        lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_IEISENTO", objFilCliIEIsento)
        If lErro <> SUCESSO Then Error 27722
    
    Else
    
        sUFDestino = mvarsUFFilialEmpresa
        
        If iTipoNFiscal = DOCINFO_TNFISFV Then
            objFilCliIEIsento.vValor = DESMARCADO
        Else
            objFilCliIEIsento.vValor = MARCADO
        End If
    
    End If
    
    iICMSInterestBaseDupla = objVarICMSInterestBaseDupla.vValor
    If objVarICMSInterestBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_INTER", sUFDestino, mvardtDataRef, iICMSInterestBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    iICMSSTBaseDupla = objVarICMSSTBaseDupla.vValor
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_ST", sUFDestino, mvardtDataRef, iICMSSTBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = objDoc.ObterVar("IND_CONSUMIDOR_FINAL", objVarIndConsumidorFinal)
    If lErro <> SUCESSO Then Error 27319
    
    bConsumidorFinal = objVarIndConsumidorFinal.vValor
    
    'testar se vai adotar a mudança da EC 87/2015
    If bConsumidorFinal And (gobjCRFAT.iEC872015Teste Or mvardtDataRef >= gobjCRFAT.dtEC872015) And sUFDestino <> mvarsUFFilialEmpresa And sUFDestino <> "EX" Then
    
        bEC872015 = True
        bProt21 = 0
        dAliqProt21 = 0
        
        'tentar achar excecao
        lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, sUFDestino, mvarsUFFilialEmpresa, (objFilCliIEIsento.vValor = MARCADO And gobjCRFAT.iICMSSTComIEIsento = DESMARCADO) Or iTipoNFiscal = DOCINFO_NFIEDV, EXCECAO_ICMS_TIPO_APLICACAO_CONSUMIDOR_FINAL)
        If lErro <> SUCESSO Then Error 27321
    
        'se achou excecao
        If (Not (objExcecao Is Nothing)) Then
        
            lErro = AtualizarItem_TratarExcecaoEC872015(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, objFilCliIEIsento.vValor = MARCADO And gobjCRFAT.iICMSSTComIEIsento = DESMARCADO, iUsaPauta, dValorPauta, dInterestAliqUFDest, dInterestPercFCPUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, dICMSpFCP, objVarICMSpFCP, iMotivo, objVarMotivo, scBenef, objVarcBenef)
            If lErro <> SUCESSO Then Error 27322
            
            dPercRedBC_EC872015 = dPercRedBase
                            
        Else 'se NAO achou uma excecao
        
            lErro = AtualizarItem_TratarNaoExcecaoEC872015(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa, sUFDestino, dInterestAliqUFDest, dInterestPercFCPUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 27323
        
        End If
        
        'CST 60 Cobrado anteriormente por ST tem que calcular a alíquota Interestadual
        'Caso necessários, para outros tipo com objTipoTribICMS.iPermiteAliquota = 0 que também precisem calcular o ICMS Interestadual, INCLUIR no IF abaixo
        'OBS-> CST 60 para consumidor final (EC 87/2015) sem os dados do ICMS Interestadual dá erro no envio da NFe
        If objTipoTribICMS.iTipo = 8 Then
        
            'Cria cópia do tipo de tributação de ICMS para não mudar os dados originais
            'objTipoTribICMSAux, objVarTipoAux, iTipoAux
            Call objTipoTribICMSAux.Copiar(objTipoTribICMS)
        
            'Obtem os dados padrões para calcular o ICMS Interestadual
            lErro = AtualizarItem_TratarNaoExcecaoEC872015(objDocItem, objTipoTrib, objTipoTribICMSAux, objVarTipoAux, iTipoAux, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa, sUFDestino, dInterestAliqUFDest, dInterestPercFCPUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 27323
        
        End If
        
        dAliquotaInterestadual = dAliquota
        dAliquotaInterestadualFCP = dICMSpFCP
        
    Else 'nao usa a EC 87/2015
    
        'se filial cliente inscrita ICMS e interestadual
        If objFilCliIEIsento.vValor = DESMARCADO And sUFDestino <> mvarsUFFilialEmpresa And sUFDestino <> "EX" Then
            
            'adotar a aliquota interestadual para operacao basica
            lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFDestino, False, sUFDestino, bProt21, dAliqProt21, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 27723
            
            dAliquotaInterestadual = dAliquota
            dAliquotaInterestadualFCP = dICMSpFCP
            
            'ver se na UF de destino tem excecao que possa incluir subst tribut
            lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, sUFDestino, mvarsUFFilialEmpresa, (objFilCliIEIsento.vValor = MARCADO And gobjCRFAT.iICMSSTComIEIsento = DESMARCADO) Or iTipoNFiscal = DOCINFO_NFIEDV, EXCECAO_ICMS_TIPO_APLICACAO_NAO_CONSUMIDOR_FINAL)
            If lErro <> SUCESSO Then Error 27724
            
            If Not (objExcecao Is Nothing) Then
    '            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
    '            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dValorPauta <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                    lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                    If lErro <> SUCESSO Then Error 27322
                    
    '                If dMargSubst <> 0 Or dValorPauta <> 0 Then dAliquota = dAliquotaInterestadual
                    dAliquota = dAliquotaInterestadual
                    dICMSpFCP = dAliquotaInterestadualFCP
                    
    '                If TipoICMSComSubstTrib(objTipoTribICMS.iTipo) Then bExcecaoComSubstTrib = True
                    
    '            End If
            End If
        
        Else
            
            If sUFDestino <> "EX" Then
                'tratar como operacao interna
            
                'tenta usar o protocolo 21
                If gobjCRFAT.iICMSProt21 <> 0 And objFilCliIEIsento.vValor = MARCADO And sUFDestino <> mvarsUFFilialEmpresa Then
                
                    lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa, True, sUFDestino, bProt21, dAliqProt21, dICMSpFCP, objVarICMSpFCP)
                    If lErro <> SUCESSO Then Error 27323
                
                End If
                
                If bProt21 = 0 Then
                
                    'tentar achar excecao
                    lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, mvarsUFFilialEmpresa, (objFilCliIEIsento.vValor = MARCADO And gobjCRFAT.iICMSSTComIEIsento = DESMARCADO) Or iTipoNFiscal = DOCINFO_NFIEDV)
                    If lErro <> SUCESSO Then Error 27321
                    
                    'se achou excecao
                    If (Not (objExcecao Is Nothing)) Then
                    
                        lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, objFilCliIEIsento.vValor = MARCADO And gobjCRFAT.iICMSSTComIEIsento = DESMARCADO, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                        If lErro <> SUCESSO Then Error 27322
                                        
                        'If TipoICMSComSubstTrib(objTipoTribICMS.iTipo) And (objFilCliIEIsento.vValor = DESMARCADO Or gobjCRFAT.iICMSSTComIEIsento = MARCADO) Then bExcecaoComSubstTrib = True
                    
                    Else 'se NAO achou uma excecao
                    
                        lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa, False, sUFDestino, bProt21, dAliqProt21, dICMSpFCP, objVarICMSpFCP)
                        If lErro <> SUCESSO Then Error 27323
                    
                    End If
                
                End If
                
            Else  'exportacao
            
                lErro = AtualizarItem_TratarNaoExcecaoSaidaCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFDestino, False, sUFDestino, bProt21, dAliqProt21, dICMSpFCP, objVarICMSpFCP)
                If lErro <> SUCESSO Then Error 27727
                
                'se ICMS_TIPO nao tiver sido preenchido manualmente...
                If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
                
                    'Coloca como não tributado
                    iTipo = ICMS_TIPO_NAO_TRIBUTADO
                    objVarTipo.vValor = iTipo
                    objTipoTribICMS.iTipo = iTipo
                    lErro = TipoICMS_Obter(objTipoTribICMS)
                    If lErro <> SUCESSO Then Error 27322
            
                End If
            
            End If
            
        End If
    
    End If
    
    If iTipoNFiscal = DOCINFO_NFIEDV Then

        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If iTipo <> ICMS_TIPO_NAO_TRIBUTADO And iTipo <> ICMS_TIPO_OUTRAS Then
        
                'Coloca como outras
                iTipo = ICMS_TIPO_OUTRAS
                objVarTipo.vValor = iTipo
                objTipoTribICMS.iTipo = iTipo
                lErro = TipoICMS_Obter(objTipoTribICMS)
                If lErro <> SUCESSO Then Error 27322
    
            End If
    
        End If
        
    End If
    
    If bProt21 Then
    
        'obter NatOp do item
        lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
        If lErro <> SUCESSO Then Error 27322
        
        If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
            
            Select Case objVarNatOp.vValor
                
                Case "6101"
                    objVarNatOp.vValor = "6107"
                Case "6102"
                    objVarNatOp.vValor = "6108"
                    
            End Select
            
            lErro = objDocItem.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 27322
        
        End If
    
    Else
    
        Select Case objTipoTribICMS.iTipo
        
            Case ICMS_TIPO_COM_SUBST_TRIB, ICMS_TIPO_RED_BASE_E_SUBST
        
                'obter NatOp do item
                lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                If lErro <> SUCESSO Then Error 27322
                
                If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
                    
                    Select Case objVarNatOp.vValor
                        Case "5101"
                            objVarNatOp.vValor = "5401"
                        Case "5102"
                            objVarNatOp.vValor = "5403"
                        Case "6101"
                            objVarNatOp.vValor = "6401"
                        Case "6102"
                            iSTCobrAnt = 0
                            lErro = ItemComSTCobradaAnt(objDoc, objDocItem, iSTCobrAnt)
                            If lErro <> SUCESSO Then Error 27765
                            
                            If iSTCobrAnt = 0 Then
                                objVarNatOp.vValor = "6403"
                            Else
                                objVarNatOp.vValor = "6404"
                            End If
                            
                    End Select
                    
                    lErro = objDocItem.GuardarVar(objVarNatOp)
                    If lErro <> SUCESSO Then Error 27322
                
                End If
    
            Case ICMS_TIPO_COBR_SUBST_ANT
        
                'obter NatOp do item
                lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                If lErro <> SUCESSO Then Error 27322
                
                If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
                    
                    Select Case objVarNatOp.vValor
                        Case "5101", "5102"
                            objVarNatOp.vValor = "5405"
                        Case "6101", "6102"
                            objVarNatOp.vValor = "6404"
                    End Select
                    
                    lErro = objDocItem.GuardarVar(objVarNatOp)
                    If lErro <> SUCESSO Then Error 27322
                
                End If
    
            Case ICMS_TIPO_NORMAL
                If mvariSimplesNacional <> 0 Then
                
                    'se ICMS_TIPO nao tiver sido preenchido manualmente...
                    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
                    
                        iTipo = ICMS_TIPO_NAO_TRIBUTADO
                        objVarTipo.vValor = iTipo
                        objTipoTribICMS.iTipo = iTipo
                        lErro = TipoICMS_Obter(objTipoTribICMS)
                        If lErro <> SUCESSO Then Error 27322
                
                    End If
                    
                End If
            
                If objFilCliIEIsento.vValor = MARCADO Then
            
                    'obter NatOp do item
                    lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                    If lErro <> SUCESSO Then Error 27322
                    
                    If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
                        
                        Select Case objVarNatOp.vValor
                            
                            Case "6101"
                                objVarNatOp.vValor = "6107"
                            Case "6102"
                                objVarNatOp.vValor = "6108"
                                
                        End Select
                        
                        lErro = objDocItem.GuardarVar(objVarNatOp)
                        If lErro <> SUCESSO Then Error 27322
                    
                    End If
        
                End If
                
                If gobjCRFAT.iTrataSTEntEmCST00 <> 0 Then
                
                    'obter NatOp do item
                    lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                    If lErro <> SUCESSO Then Error 27322
                    
                    If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then

                        Select Case objVarNatOp.vValor
                             Case "6102"
                                iSTCobrAnt = 0
                                lErro = ItemComSTCobradaAnt(objDoc, objDocItem, iSTCobrAnt)
                                If lErro <> SUCESSO Then Error 27765
                                
                                If iSTCobrAnt <> 0 Then
                                    objVarNatOp.vValor = "6404"
                                End If
                         End Select
                    
                        lErro = objDocItem.GuardarVar(objVarNatOp)
                        If lErro <> SUCESSO Then Error 27322
                        
                    End If
                End If
        
            Case ICMS_TIPO_RED_BASE
            
                If objFilCliIEIsento.vValor = MARCADO Then
            
                    'obter NatOp do item
                    lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
                    If lErro <> SUCESSO Then Error 27322
                    
                    If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
                        
                        Select Case objVarNatOp.vValor
                            
                            Case "6101"
                                objVarNatOp.vValor = "6107"
                            Case "6102"
                                objVarNatOp.vValor = "6108"
                                
                        End Select
                        
                        lErro = objDocItem.GuardarVar(objVarNatOp)
                        If lErro <> SUCESSO Then Error 27322
                    
                    End If
        
                End If
        
        End Select
    
    End If
    
    If objTipoTribICMS.iPermiteST = MARCADO And (objFilCliIEIsento.vValor = DESMARCADO Or gobjCRFAT.iICMSSTComIEIsento = MARCADO) And (dMargSubst <> 0 Or dValorPauta <> 0) Then bExcecaoComSubstTrib = True
    If objTipoTribICMS.iPermiteST = MARCADO And (objFilCliIEIsento.vValor = DESMARCADO Or gobjCRFAT.iICMSSTComIEIsento = MARCADO) And Not (objExcecao Is Nothing) And bExcecaoComSubstTrib = False And bProt21 = False And dAliquotaInterestadual <> 0 Then
        If (objExcecao.iTipo = ICMS_TIPO_COM_SUBST_TRIB Or objExcecao.iTipo = ICMS_TIPO_RED_BASE_E_SUBST) Then bExcecaoComSubstTrib = True
    End If
    
    If objTipoTribICMS.iPermiteReducaoBase <> 1 Then
        dPercRedBase = 0
        dPercRedBaseSubst = 0
    End If
    
    If objTipoTribICMS.iPermiteMargLucro <> 1 Then
        dMargSubst = 0
    End If
    
    'Calcula a permissão de crédito de ICMS
    'VAI ESTAR FALTANDO ZERAR OUTROS VALORES COMO A BASE E ALÍQUOTA NORMAL QUANDO NÃO PERMITE, ETC
    If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        objVarAliCredSimples.vValor = 0
        objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        objVarVlrCredSimples.vValor = 0
        objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        
        If objTipoTribICMS.iPermiteMotDesoneracao = DESMARCADO Then
            objVarMotivo.vValor = 0
            objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVI.vValor = 0
            objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermitePercBaseOperProp = DESMARCADO Then
            objVarPBOP.vValor = 0
            objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteUFICMSSTDevido = DESMARCADO Then
            objVarUFD.vValor = ""
            objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoTribICMS.iPermiteST = DESMARCADO Then iZerarST = MARCADO
        
    Else
               
        'Se não é manual ou não está preenchido calcula
        If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Or objVarTipoICMSSimples.vValor = 0 Then
        
            If (Not (objExcecao Is Nothing)) Then
                objVarTipoICMSSimples.vValor = objExcecao.iTipoSimples
                If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            Else
                objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            End If
            If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = 1
            objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
            Select Case objVarTipoICMSSimples.vValor
                Case 1, 2, 4, 5
                
                    If iTipoNFiscal <> DOCINFO_NFISSAT And iTipoNFiscal <> DOCINFO_TNFISFV Then
                
                        lErro = objDoc.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilCliObj)
                        If lErro <> SUCESSO Then Error 27916
                            
                        Set objFilial = objVarFilCliObj.vValor
                    
                    Else
                    
                        Set objFilial = New ClassFilialCliente
                        If iTipoNFiscal = DOCINFO_TNFISFV Then
                            objFilial.iIEIsento = DESMARCADO
                            objFilial.iIENaoContrib = DESMARCADO
                        Else
                            objFilial.iIEIsento = MARCADO
                            objFilial.iIENaoContrib = MARCADO
                        End If
                    
                    End If
                    
                    'Se cliente não tem direito a crédito ou o tipo está marcado como não tendo direito a crédito
                    If objFilial.iIEIsento = MARCADO Or objFilial.iIENaoContrib = MARCADO Or objFilial.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES Or objVarTipoICMSSimples.vValor = 2 Or objVarTipoICMSSimples.vValor = 5 Then
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                        End If
                    Else
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 4 'Com crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 1 'Com crédito - Sem ST
                        End If
                    End If
                Case 3, 6
                    If bExcecaoComSubstTrib Then
                        objVarTipoICMSSimples.vValor = 6 'Com crédito - Com ST
                    Else
                        objVarTipoICMSSimples.vValor = 3 'Com crédito - Sem ST
                    End If
            End Select
        End If
        
        If iTipoNFiscal = DOCINFO_NFIEDV Then
            If objVarTipoICMSSimples.vValor <> 8 Then objVarTipoICMSSimples.vValor = 10
        End If
          
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        Set objVarItemTipoObj.vValor = objTipoICMSSimples
        lErro = objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
        If lErro <> SUCESSO Then Error 27320

        Set objTipoICMSSimples = objVarItemTipoObj.vValor
       
        If objTipoICMSSimples.iPermiteSTRetAnt = DESMARCADO Or iTipoNFiscal = DOCINFO_NFIEDV Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        
        If objTipoICMSSimples.iPermiteCredito = DESMARCADO Or iTipoNFiscal = DOCINFO_NFIEDV Then
            objVarAliCredSimples.vValor = 0
            objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVlrCredSimples.vValor = 0
            objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            
            If objVarAliCredSimples.iStatus <> VAR_PREENCH_MANUAL Then
            
                objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
                lErro = objDoc.ObterVar("DATA_EMISSAO", objVarEmissao)
                If lErro <> SUCESSO Then Error 27320
            
                objDASAliquotas.iFilialEmpresa = mvariFilialEmpresa
                objDASAliquotas.iAno = Year(DateAdd("m", -1, IIf(objVarEmissao.vValor = DATA_NULA, gdtDataAtual, objVarEmissao.vValor)))
                objDASAliquotas.iMes = Month(DateAdd("m", -1, IIf(objVarEmissao.vValor = DATA_NULA, gdtDataAtual, objVarEmissao.vValor)))
                
                lErro = CF("DASAliquotas_Le", objDASAliquotas)
                If lErro <> SUCESSO And lErro <> 200911 Then Error 27320
                
                lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
                If lErro <> SUCESSO Then Error 27320
                
                Set objProduto = objVarProd.vValor
                
                If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then
                    objVarAliCredSimples.vValor = objDASAliquotas.dAliquotaICMSServ
                Else
                    objVarAliCredSimples.vValor = objDASAliquotas.dAliquotaICMS
                End If
                                
            End If

            If objVarVlrCredSimples.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                
                objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
                
                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdVlr)
                If lErro <> SUCESSO Then Error 27320
                
                objVarVlrCredSimples.vValor = Arredonda_Moeda(objVarProdVlr.vValor * objVarAliCredSimples.vValor)
            
            End If
            
        End If
        If objTipoICMSSimples.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoICMSSimples.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    
    End If
    'Calcula a permissão de crédito de ICMS
    
    If iTipoNFiscal = DOCINFO_NFIEDV Then
                       
        If Not (objItemNFOrig Is Nothing) Then
                        
            lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
            If lErro <> SUCESSO Then Error 27322
                        
            If objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO Then
            
                sCFOPItemNFOrig = objItemNFOrig.objTributacao.sNaturezaOp
                
                Select Case sCFOPItemNFOrig
                
                    Case "5101"
                        objVarNatOp.vValor = "1201"
                    Case "6101"
                        objVarNatOp.vValor = "2201"
                    Case "5102"
                        objVarNatOp.vValor = "1202"
                    Case "6102"
                        objVarNatOp.vValor = "2202"
                    Case "5401"
                        objVarNatOp.vValor = "1410"
                    Case "6401"
                        objVarNatOp.vValor = "2410"
                    Case Else
                        If left(sCFOPItemNFOrig, 1) = "5" Then
                            objVarNatOp.vValor = "1949"
                        Else
                            objVarNatOp.vValor = "2949"
                        End If
                        
                End Select
                
                lErro = objDocItem.GuardarVar(objVarNatOp)
                If lErro <> SUCESSO Then Error 27322
        
            End If
            
        End If
    
    End If
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo, iTipoNFiscal) Then
        
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27324
            
        Else
            
            dBaseCalculo = 0
            If gobjCRFAT.iVendaFuturaTributacao = 1 Then
                dAliquota = 0
                dICMSpFCP = 0
            End If
        End If
    Else
        
        dBaseCalculo = objVarBaseCalculo.vValor
    
    End If

    If bEC872015 Then
        
        dBC_EC872015 = 0
    
        If objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES Then
                
            If TipoICMSSimplesComBaseEC872015(objVarTipoICMSSimples.vValor) Then
            
                lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBC_EC872015)
                If lErro <> SUCESSO Then Error 27324
        
            End If
    
        Else
        
            If TipoICMSComBaseEC872015(objTipoTribICMS.iTipo) Then
                
                If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
            
                    lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBC_EC872015)
                    If lErro <> SUCESSO Then Error 27324
                    
                Else
                
                    dBC_EC872015 = dBaseCalculo
                    
                End If
        
            End If
        
        End If
    
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        'dValor = CDbl(Format(dValor, "0.0000"))
        dValor = Arredonda_Moeda(dValor)
    
    Else
        dValor = objVarValor.vValor
    End If

    If iTipoNFiscal = DOCINFO_NFIEDV Then
    
        If dMargSubst <> 0 And objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL And dMargSubstOrig <> 0 Then dMargSubst = dMargSubstOrig
        If dValorPauta <> 0 And iUsaPauta <> 0 And dValorPautaOrig <> 0 Then dValorPauta = dValorPautaOrig
        
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        
        If bExcecaoComSubstTrib Or bProt21 Then
            
            If iUsaPauta = 0 Then
            
                dBaseSubst = dBaseCalculo
                
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = dBaseSubst + objVar.vValor
                
                End If
                
                If iICMSSTBaseDupla = MARCADO Then
                
                    'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
                    If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                        If objExcecao Is Nothing Then
                            If bProt21 Then
                                dAliqSubst = dAliqProt21
                            Else
                                dAliqSubst = 0
                            End If
                        Else
                            dAliqSubst = objExcecao.dAliquota
                        End If
                    Else
                        dAliqSubst = objVarAliqSubst.vValor
                    End If
        
                    dBaseSubst = (dBaseSubst - dValor) / (1 - dAliqSubst)
                End If
                
                dBaseSubst = dBaseSubst * (1 + dMargSubst)
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
'                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
'                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
                dBaseSubst = Arredonda_Moeda(dBaseSubst)
                dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida)
                
            Else
        
                'obter qtde do item de objdocitem
                lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                'dBaseSubst = CDbl(Format(objVar.vValor * dValorPauta, "0.0000"))
                dBaseSubst = Arredonda_Moeda(objVar.vValor * dValorPauta)
                dBaseSubstReduzida = dBaseSubst
                
                iModalidadeST = ICMS_ST_MODALIDADE_PAUTA
                
            End If
            
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dMargSubst = objVarMargSubst.vValor
        'dBaseSubst = dBaseSubst * (1 + dMargSubst)
        dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
    End If
    
    If iZerarST Then
        dBaseSubst = 0
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dMargSubst = 0
        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dPercRedBaseSubst = 0
        objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dBaseSubstReduzida = 0
        
        objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        dICMSpFCPST = 0
    End If
    
    If bEC872015 Then
    
'        If objVarInterestBCUFDest.iStatus <> VAR_PREENCH_MANUAL Then
'            dInterestBCUFDest = dBC_EC872015
'        Else
'            dInterestBCUFDest = objVarInterestBCUFDest.vValor
'        End If
        
        If objVarInterestAliq.iStatus <> VAR_PREENCH_MANUAL Then
            dInterestAliq = dAliquota
        Else
            dInterestAliq = objVarInterestAliq.vValor
        End If
        
        If objVarInterestBCUFDest.iStatus <> VAR_PREENCH_MANUAL Then
            If iICMSInterestBaseDupla = DESMARCADO Then
                dInterestBCUFDest = dBC_EC872015
            Else
                If objFilCliIEIsento.vValor = MARCADO Then
                    dInterestBCUFDest = (dBC_EC872015) / (1 - dInterestAliqUFDest)
                Else
                    dInterestBCUFDest = (dBC_EC872015 - dValor) / (1 - dInterestAliqUFDest)
                End If
            End If
        Else
            dInterestBCUFDest = objVarInterestBCUFDest.vValor
        End If
    End If
    
    'se é para zerar a parte propria (mantendo apenas a subst tributaria)
    dBaseICMSProprioSTSimples = 0
    dICMSProprioSTSimples = 0
    If mvariZerarTributo <> 0 Or iZerarAliquota = MARCADO Or iTipo = ICMS_TIPO_COM_DIFERIMENTO Then
        If mvariSimplesNacional <> 0 Then
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseICMSProprioSTSimples)
            If lErro <> SUCESSO Then Error 27324
        
            dAliquotaPadrao = 0
            If sUFDestino = mvarsUFFilialEmpresa Then
                
                'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
                If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                    If objExcecao Is Nothing Then
                        dAliquotaPadrao = 0
                    Else
                        dAliquotaPadrao = objExcecao.dAliquota
                    End If
                Else
                    dAliquotaPadrao = objVarAliqSubst.vValor
                End If
            
                dICMSProprioSTSimples = dBaseICMSProprioSTSimples * Arredonda_Moeda((1 - dPercRedBaseSubst) * dAliquotaPadrao, 2)
            
            Else
                
                lErro = ObterAliquotaBasicaSaidaCli(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquotaPadrao, sUFDestino, objDocItem, dAliquotaPadraoFCP)
                If lErro <> SUCESSO Then Error 27318
            
                dICMSProprioSTSimples = dBaseICMSProprioSTSimples * Arredonda_Moeda((1 - dPercRedBase) * dAliquotaPadrao, 2)
            
            End If
            
        End If
        
        If iZerarAliquota = MARCADO Then
            objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        
        'http://portalnfe.fazenda.mg.gov.br/respostas_III.html
        'Na versão 3.10 do leiaute da NF-e foram criados campos para informação do diferimento parcial.
        'Esses campos também deverão ser preenchidos no caso do diferimento total?
        'Os campos modBC (Modalidade de determinação da base de cálculo do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação), pDif (Percentual do diferimento), vICMSDif (Valor do ICMS diferido) e vICMS (Valor do ICMS)
        'só devem ser preenchidos no caso do diferimento ser parcial.
        'Esses campos NÃO serão de preenchimento obrigatório para os casos do diferimento total.
        'Alertamos que os campos modBC (Modalidade de determinação da BC do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação) devem ser preenchidos com o valor como se não tivesse o diferimento.

        If objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO And mvariSimplesNacional = 0 And mvariZerarTributo = 0 And iZerarAliquota <> MARCADO Then
        
            If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
                dPercDifer = 1 '??? deveria vir da excecao
                objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dPercDifer = objVarPercDifer.vValor
            End If
            
            If dPercDifer <> 1 And dPercDifer <> 0 Then bDiferimentoParcial = True
        
            If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
                d51ValorOp = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
                objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                d51ValorOp = objVar51ValorOp.vValor
            End If
            
            If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
                dValorDif = Arredonda_Moeda(d51ValorOp * dPercDifer)
                objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dValorDif = objVarValorDif.vValor
            End If
        
        Else
        
            d51ValorOp = 0
            objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            dPercDifer = 0
            objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            dValorDif = 0
            objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Then
            
            If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then dBaseCalculo = 0
            If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then dPercRedBase = 0
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
        
            If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then dICMSpFCP = 0
        
        Else
        
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(d51ValorOp - dValorDif)
        
        End If
        
    End If
    
    If dBaseSubst <> 0 Then
        
        'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
        If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objExcecao Is Nothing Then
                If bProt21 Then
                    dAliqSubst = dAliqProt21
                Else
                    dAliqSubst = 0
                End If
            Else
                dAliqSubst = objExcecao.dAliquota
            End If
        Else
            dAliqSubst = objVarAliqSubst.vValor
        End If
        
        'NFe 4.00
        If objVarICMSpFCPST.iStatus <> VAR_PREENCH_MANUAL Then
            If Not (objExcecao Is Nothing) Then
                dICMSpFCPST = objExcecao.dICMSPercFCP
            End If
        Else
            dICMSpFCPST = objVarICMSpFCPST.vValor
        End If
        'NFe 4.00
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If mvariSimplesNacional = 0 Then
                dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
            Else
                If mvardtDataRef < #8/1/2009# Then
                    'dValorSubst = dBaseSubstReduzida * dAliqSubst - CDbl(Format(dBaseICMSProprioSTSimples * 0.07, "0.0000"))
                    dValorSubst = dBaseSubstReduzida * dAliqSubst - Arredonda_Moeda(dBaseICMSProprioSTSimples * 0.07)
                Else
                    dValorSubst = dBaseSubstReduzida * dAliqSubst - dICMSProprioSTSimples
                End If
            End If
            If dValorSubst < 0 Then dValorSubst = 0
            'dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
            dValorSubst = Arredonda_Moeda(dValorSubst)
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
    End If
    
    
    If iTipo = ICMS_TIPO_COBR_SUBST_ANT And gobjCRFAT.iTribCalcSTCobrAnt <> 0 And (objVarCobrAntBase.iStatus <> VAR_PREENCH_MANUAL Or objVarCobrAntValor.iStatus <> VAR_PREENCH_MANUAL) Then
    
        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27325
        
        Set objProduto = objVarProd.vValor
        
        'obter qtde do item de objdocitem
        lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVar)
        If lErro <> SUCESSO Then Error 27325
        
        lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
        If lErro <> SUCESSO Then Error 27325
        
        dtDataEmissao = objVarDataEmissao.vValor
        
        dBCA = 0
        dVCA = 0
        
        'tentar obter os valores de base e valor de st cobrados anteriormente por ST
        lErro = CF("ICMSSTCobrAnt_Obter", mvariFilialEmpresa, dtDataEmissao, objProduto.sCodigo, StrParaDbl(objVar.vValor), "", dBCA, dVCA)
        If lErro <> SUCESSO And lErro <> ERRO_SEM_MENSAGEM Then Error 27325
    
        If objVarCobrAntBase.iStatus <> VAR_PREENCH_MANUAL Then
        
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = dBCA
            
        End If
    
'        If objVarCobrAntValor.iStatus <> VAR_PREENCH_MANUAL Then
'
'            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
'            objVarCobrAntValor.vValor = dVCA
'
'        End If
        
        'NFe 4.00
        If objVarICMSSTCobrAntAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            objVarICMSSTCobrAntAliquota.vValor = dVCA / dBCA
            
            If objVarCobrAntValor.iStatus <> VAR_PREENCH_MANUAL Then
            
                objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
                objVarCobrAntValor.vValor = dVCA
                
            End If
        Else
    
            If objVarCobrAntValor.iStatus <> VAR_PREENCH_MANUAL Then
            
                objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
                objVarCobrAntValor.vValor = Arredonda_Moeda(dBCA * objVarICMSSTCobrAntAliquota.vValor)
                
            End If
    
        End If
        
    End If
    
    'NFe 4.00
    'Base de cálculo dos FCP iguais as bases correspondentes do ICMS
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCP = dBaseCalculo
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCP = objVarICMSvBCFCP.vValor
    End If
    
    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCPST = dBaseSubst
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCPST = objVarICMSvBCFCPST.vValor
    End If
           
    'Calculando o valor com a base e alíquota
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCP = Arredonda_Moeda(dICMSvBCFCP * dICMSpFCP)
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCP = objVarICMSvFCP.vValor
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCPST = Arredonda_Moeda(dICMSvBCFCPST * dICMSpFCPST)
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCPST = objVarICMSvFCPST.vValor
    End If
    'NFe 4.00

    'Se credita e não é com diferimento total
    If objTipoTrib.iICMSCredita <> 0 And (objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Or bDiferimentoParcial) And (mvariSimplesNacional = 0 Or mvardtDataRef < DATA_INICIO_SIMPLES_NACIONAL) Then
    
        dValorCredito = dValor
        
    End If

Label_Atualizar:
    
    If objVarVI.iStatus = VAR_PREENCH_AUTOMATICO And iMotivo <> 0 And objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        Select Case mvarsUFFilialEmpresa
        
            Case "RJ"
            
                Select Case iTipo
                
                    Case ICMS_TIPO_RED_BASE 'CST 20
                    
                    Case ICMS_TIPO_ISENTO 'CST 40
                    
                    Case ICMS_TIPO_RED_BASE_E_SUBST 'CST 70
                    
                End Select
        
        End Select
        
    End If
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dPercRedBaseSubst, dValorCredito, iModalidade, iModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarRegTrib, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, d51ValorOp, dPercDifer, dValorDif, iMotivo)
    If lErro <> SUCESSO Then Error 27328
    
    lErro = AtualizarItem_GuardarValores3(objDocItem, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, dICMSvBCFCP, dICMSpFCP, dICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, dICMSvBCFCPST, dICMSpFCPST, dICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, dICMSvBCFCPSTRet, dICMSpFCPSTRet, dICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, dICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, iICMSSTBaseDupla, objVarICMSInterestBaseDupla, iICMSInterestBaseDupla, objVarcBenef, scBenef)
    If lErro <> SUCESSO Then Error 27328
    
    If bEC872015 Then
    
        'Caso Diaglab
        If mvariSimplesNacional <> 0 And gobjCRFAT.iEC872015Simples = 0 And dInterestAliq = 0 And objTipoTribICMS.iTipoTribCST = 60 Then
        
            lErro = ObterAliquotaBasicaEC872015(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dInterestAliq, sUFDestino, objDocItem, dAux1, dAux2, dInterestPercFCPUFDest)
            If lErro <> SUCESSO Then Error 27318
        
        End If
        
        If objVarICMSInterestBCFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
            dICMSInterestBCFCPUFDest = dInterestBCUFDest
            objVarICMSInterestBCFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            dICMSInterestBCFCPUFDest = objVarICMSInterestBCFCPUFDest.vValor
        End If
    
        If dInterestBCUFDest <> 0 And dInterestAliq <> 0 And objFilCliIEIsento.vValor = MARCADO Then
        
            If objVarInterestVlrFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
                dInterestVlrFCPUFDest = Arredonda_Moeda(dICMSInterestBCFCPUFDest * dInterestPercFCPUFDest)
            Else
                dInterestVlrFCPUFDest = objVarInterestVlrFCPUFDest.vValor
            End If
                
            If objVarInterestPercPartilha.iStatus <> VAR_PREENCH_MANUAL Then
                
                lErro = objDoc.ObterVar("DATA_EMISSAO", objVarEmissao)
                If lErro <> SUCESSO Then Error 27320
            
                If objVarEmissao.iStatus = VAR_PREENCH_VAZIO Then
                
                    dInterestPercPartilha = 0
                    
                Else
                
                    If Year(objVarEmissao.vValor) < 2015 Then
                    
                        dInterestPercPartilha = 0
                        
                    Else
                    
                        Select Case Year(objVarEmissao.vValor)
                            
                            Case 2015, 2016
                                dInterestPercPartilha = 0.4
                            
                            Case 2017
                                dInterestPercPartilha = 0.6
                            
                            Case 2018
                                dInterestPercPartilha = 0.8
                        
                            Case Else
                                dInterestPercPartilha = 1
                                
                        End Select
                    
                    End If
                            
                End If
                
            Else
            
                dInterestPercPartilha = objVarInterestPercPartilha.vValor
            
            End If
                
            If (dInterestAliqUFDest > dInterestAliq) Then
                dDifal = Arredonda_Moeda(dInterestBCUFDest * (dInterestAliqUFDest - dInterestAliq), 2)
            Else
                dDifal = 0
            End If
                
            If dDifal <> 0 Then
            
                If objVarInterestVlrUFDest.iStatus <> VAR_PREENCH_MANUAL Then
                    dInterestVlrUFDest = Arredonda_Moeda((dDifal * dInterestPercPartilha), 2)
                Else
                    dInterestVlrUFDest = objVarInterestVlrUFDest.vValor
                End If
                
                If objVarInterestVlrUFRemet.iStatus <> VAR_PREENCH_MANUAL Then
                        
                    dInterestVlrUFRemet = Arredonda_Moeda(dDifal - dInterestVlrUFDest, 2)
                
                Else
                    dInterestVlrUFRemet = objVarInterestVlrUFRemet.vValor
                End If
    
            End If
            
            'se vale a liminar do Tofoli
            If mvariSimplesNacional <> 0 And gobjCRFAT.iEC872015Simples = 0 Then
            
                If objVarInterestBCUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestBCUFDest = 0
                If objVarInterestPercFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestPercFCPUFDest = 0
                If objVarInterestAliqUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestAliqUFDest = 0
                'If objVarInterestAliq.iStatus <> VAR_PREENCH_MANUAL Then dInterestAliq = 0
                'If objVarInterestPercPartilha.iStatus <> VAR_PREENCH_MANUAL Then dInterestPercPartilha = 0
                If objVarInterestVlrFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrFCPUFDest = 0
                If objVarInterestVlrUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrUFDest = 0
                If objVarInterestVlrUFRemet.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrUFRemet = 0
                If objVarICMSInterestBCFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dICMSInterestBCFCPUFDest = 0
            
            End If
            
        Else
        
            'zerar toda a parte da aba nova
            If objVarInterestBCUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestBCUFDest = 0
            If objVarInterestPercFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestPercFCPUFDest = 0
            If objVarInterestAliqUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestAliqUFDest = 0
            If objVarInterestAliq.iStatus <> VAR_PREENCH_MANUAL Then dInterestAliq = 0
            If objVarInterestPercPartilha.iStatus <> VAR_PREENCH_MANUAL Then dInterestPercPartilha = 0
            If objVarInterestVlrFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrFCPUFDest = 0
            If objVarInterestVlrUFDest.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrUFDest = 0
            If objVarInterestVlrUFRemet.iStatus <> VAR_PREENCH_MANUAL Then dInterestVlrUFRemet = 0
            If objVarICMSInterestBCFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then dICMSInterestBCFCPUFDest = 0
        
        End If
    
    End If
    
    lErro = AtualizarItem_GuardarValores2(objDocItem, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                dInterestBCUFDest, dInterestPercFCPUFDest, dInterestAliqUFDest, dInterestAliq, dInterestPercPartilha, dInterestVlrFCPUFDest, dInterestVlrUFDest, dInterestVlrUFRemet, objVarICMSInterestBCFCPUFDest, dICMSInterestBCFCPUFDest)
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

Private Function ExcecaoCli_Pesquisar(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, sUFDestino As String, sUFOrigem As String, ByVal bSemSubstTributaria As Boolean, Optional ByVal iTipoAplicacao As Integer = 0) As Long
'uma excecao a nivel de cliente tem + prioridade que excecoes a nivel de produto apenas
'e uma excecao a nivel de cliente+produto tem a maior prioridade de todas.
'A colecao já está ordenada de forma que pararei na 1a excecao que encontrar
'retorna Nothing se nao achar nenhuma excecao adequada
Dim lErro As Long, objICMSExcecao As ClassICMSExcecao, iAchou As Integer
Dim objVarUFDestino As New ClassVariavelCalculo, objVarFilCliColCateg As New ClassVariavelCalculo
Dim objVarProdColCateg As New ClassVariavelCalculo
Dim colExcecoes As New Collection
Dim objVarNaoCalcST As New ClassVariavelCalculo, bValida As Boolean

On Error GoTo Erro_ExcecaoCli_Pesquisar

    'Incluído 14/07/2011
    lErro = objDocItem.objDoc.ObterVar("ICMS_ST_NAO_CALCULAR", objVarNaoCalcST)
    If lErro <> SUCESSO Then Error 27220
    
    If objVarNaoCalcST.vValor <> MARCADO Then
    'Fim - Incluído 14/07/2011
    
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
        
            lErro = OrigemMercValida(objDocItem, objICMSExcecao.iGrupoOrigemMercadoria, bValida)
            If lErro <> SUCESSO Then Error 27725
            
            'Se filtrou pelo tipo de aplicação
            If iTipoAplicacao <> EXCECAO_ICMS_TIPO_APLICACAO_TODOS Then
                'Ou vale para todas ou vale para que foi filtrada
                If objICMSExcecao.iTipoAplicacao <> EXCECAO_ICMS_TIPO_APLICACAO_TODOS And objICMSExcecao.iTipoAplicacao <> iTipoAplicacao Then
                    bValida = False
                End If
            End If
        
            If bValida And objICMSExcecao.iTipoCliForn = 0 And (sUFOrigem = "" Or Trim(objICMSExcecao.sEstadoOrigem) = "" Or objICMSExcecao.sEstadoOrigem = sUFOrigem) Then
            
                If Not (bSemSubstTributaria And TipoICMSComSubstTrib(objICMSExcecao.iTipo)) Then
                
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
        
            End If
        
        Next
        
        If iAchou <> 1 Then Set objExcecao = Nothing
        
    'Incluído 14/07/2011
    End If
    'Fim - Incluído 14/07/2011
    
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

Private Function OrigemMercValida(ByVal objDocItem As ClassTributoDocItem, ByVal iGrupoOrigemMercadoria As Integer, bValida As Boolean) As Long

Dim lErro As Long, objVarOrigemProd As New ClassVariavelCalculo

On Error GoTo Erro_OrigemMercValida
    
    If iGrupoOrigemMercadoria = 0 Then
    
        bValida = True
    
    Else
    
        lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
        If lErro <> SUCESSO Then Error 27876
        
        Select Case objVarOrigemProd.vValor
        
            Case PRODUTO_ORIGEM_IMP_DIRETA, PRODUTO_ORIGEM_EST_MERC_INTERNO, PRODUTO_ORIGEM_NAC_COM_IMP, PRODUTO_ORIGEM_NAC_COM_IMP2
                bValida = IIf(iGrupoOrigemMercadoria = 1, True, False)
                
            Case Else
                bValida = IIf(iGrupoOrigemMercadoria = 2, True, False)
    
        End Select
        
    End If

    OrigemMercValida = SUCESSO

    Exit Function

Erro_OrigemMercValida:

    OrigemMercValida = Err

    Select Case Err

        Case 27876
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153833)

    End Select

    Exit Function

End Function

Private Function AliqInterstSaida_Pesquisar(sSiglaDestino As String, dAliquota As Double, ByVal objDocItem As ClassTributoDocItem, dAliquotaFCP As Double) As Long
'retorna a aliquota interestadual partindo do estado da filial empresa corrente
Dim lErro As Long, objICMSAliqExterna As ClassICMSAliqExterna
Dim objVarOrigemProd As New ClassVariavelCalculo

On Error GoTo Erro_AliqInterstSaida_Pesquisar
    
    If gobjCRFAT.dtResolucao13Senado <> DATA_NULA And mvardtDataRef >= gobjCRFAT.dtResolucao13Senado Then
    
        lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
        If lErro <> SUCESSO Then Error 27876
        
        Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sSiglaDestino)
        
        Select Case objVarOrigemProd.vValor
        
            Case PRODUTO_ORIGEM_IMP_DIRETA, PRODUTO_ORIGEM_EST_MERC_INTERNO, PRODUTO_ORIGEM_NAC_COM_IMP, PRODUTO_ORIGEM_NAC_COM_IMP2
                dAliquota = 0.04
                
            Case Else
'                Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sSiglaDestino)
                
                dAliquota = objICMSAliqExterna.dAliquota
        
        End Select
        
        dAliquotaFCP = objICMSAliqExterna.dAliquotaFCP
        
    Else
    
        Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sSiglaDestino)
        
        dAliquota = objICMSAliqExterna.dAliquota
        dAliquotaFCP = objICMSAliqExterna.dAliquotaFCP
    
    End If
    
    AliqInterstSaida_Pesquisar = SUCESSO

    Exit Function

Erro_AliqInterstSaida_Pesquisar:

    AliqInterstSaida_Pesquisar = Err

    Select Case Err

        Case 27876
        
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
    
        If Trim(UCase(objFilialCliCategoria.sCategoria)) = Trim(UCase(sCategoriaCliente)) And Trim(UCase(objFilialCliCategoria.sItem)) = Trim(UCase(sCategoriaClienteItem)) Then
        
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
    
        If Trim(UCase(objProdutoCategoria.sCategoria)) = Trim(UCase(sCategoriaProduto)) And Trim(UCase(objProdutoCategoria.sItem)) = Trim(UCase(sCategoriaProdutoItem)) Then
        
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

Private Function AtualizarDoc_GuardarValores(ByVal objDoc As ClassTributoDoc, ByVal dBaseReduzida As Double, ByVal dValor As Double, ByVal dBaseSubstReduzida As Double, ByVal dValorSubst As Double, ByVal dValorCredito As Double, ByVal dValorIsento As Double, ByVal dICMSInterestVlrFCPUFDest As Double, ByVal dICMSInterestVlrUFDest As Double, ByVal dICMSInterestVlrUFRemet As Double, ByVal dICMSVlrFCP As Double, ByVal dICMSVlrFCPST As Double, ByVal dICMSVlrFCPSTRet As Double, ByVal dIPIVlrDevolvido As Double) As Long
'atualiza ICMS_BASE, ICMS_VALOR, ICMS_SUBST_BASE, ICMS_SUBST_VALOR e ICMS_VALOR_CREDITO nao preenchidos manualmente
Dim objVarICMSBase As New ClassVariavelCalculo, objVarICMSValor As New ClassVariavelCalculo
Dim objVarICMSBaseSubst As New ClassVariavelCalculo, objVarICMSValorSubst As New ClassVariavelCalculo
Dim objVarICMSValorCredito As New ClassVariavelCalculo, objVarICMSValorIsento As New ClassVariavelCalculo

'NT 2015.003 (ICMS Interestadual)
Dim objVarICMSInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarICMSInterestVlrUFDest As New ClassVariavelCalculo, objVarICMSInterestVlrUFRemet As New ClassVariavelCalculo

'NFe 4.00
Dim objVarICMSVlrFCP As New ClassVariavelCalculo, objVarICMSVlrFCPST As New ClassVariavelCalculo, objVarICMSVlrFCPSTRet As New ClassVariavelCalculo, objVarIPIVlrDevolvido As New ClassVariavelCalculo

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_GuardarValores

    lErro = objDoc.ObterVar("ICMS_BASE", objVarICMSBase)
    If lErro <> SUCESSO Then Error 27235
    'If objVarICMSBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSBase.vValor = dBaseReduzida
        objVarICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSBase)
        If lErro <> SUCESSO Then Error 27236
    'End If
    
    lErro = objDoc.ObterVar("ICMS_VALOR", objVarICMSValor)
    If lErro <> SUCESSO Then Error 27237
    'If objVarICMSValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValor.vValor = dValor
        objVarICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValor)
        If lErro <> SUCESSO Then Error 27238
    'End If
    
    lErro = objDoc.ObterVar("ICMS_SUBST_BASE", objVarICMSBaseSubst)
    If lErro <> SUCESSO Then Error 27239
    'If objVarICMSBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSBaseSubst.vValor = dBaseSubstReduzida
        objVarICMSBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSBaseSubst)
        If lErro <> SUCESSO Then Error 27240
    'End If
    
    lErro = objDoc.ObterVar("ICMS_SUBST_VALOR", objVarICMSValorSubst)
    If lErro <> SUCESSO Then Error 27241
    'If objVarICMSValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValorSubst.vValor = dValorSubst
        objVarICMSValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValorSubst)
        If lErro <> SUCESSO Then Error 27242
    'End If

    lErro = objDoc.ObterVar("ICMS_VALOR_CREDITO", objVarICMSValorCredito)
    If lErro <> SUCESSO Then Error 27876
    'If objVarICMSValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSValorCredito.vValor = dValorCredito
        objVarICMSValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarICMSValorCredito)
        If lErro <> SUCESSO Then Error 27877
    'End If

    'nfe 3.10
    lErro = objDoc.ObterVar("ICMS_VALOR_ISENTO", objVarICMSValorIsento)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSValorIsento.vValor = dValorIsento
    objVarICMSValorIsento.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSValorIsento)
    If lErro <> SUCESSO Then Error 27236
    
    'NT 2015.003 (ICMS Interestadual)
    lErro = objDoc.ObterVar("ICMS_INTEREST_VLR_FCP_UF_DEST", objVarICMSInterestVlrFCPUFDest)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSInterestVlrFCPUFDest.vValor = dICMSInterestVlrFCPUFDest
    objVarICMSInterestVlrFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSInterestVlrFCPUFDest)
    If lErro <> SUCESSO Then Error 27236
    
    lErro = objDoc.ObterVar("ICMS_INTEREST_VLR_UF_DEST", objVarICMSInterestVlrUFDest)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSInterestVlrUFDest.vValor = dICMSInterestVlrUFDest
    objVarICMSInterestVlrUFDest.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSInterestVlrUFDest)
    If lErro <> SUCESSO Then Error 27236
    
    lErro = objDoc.ObterVar("ICMS_INTEREST_VLR_UF_REMET", objVarICMSInterestVlrUFRemet)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSInterestVlrUFRemet.vValor = dICMSInterestVlrUFRemet
    objVarICMSInterestVlrUFRemet.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSInterestVlrUFRemet)
    If lErro <> SUCESSO Then Error 27236
    
    'NFe 4.00
    lErro = objDoc.ObterVar("ICMS_VLR_FCP", objVarICMSVlrFCP)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSVlrFCP.vValor = dICMSVlrFCP
    objVarICMSVlrFCP.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSVlrFCP)
    If lErro <> SUCESSO Then Error 27236
    
    lErro = objDoc.ObterVar("ICMS_VLR_FCP_ST", objVarICMSVlrFCPST)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSVlrFCPST.vValor = dICMSVlrFCPST
    objVarICMSVlrFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSVlrFCPST)
    If lErro <> SUCESSO Then Error 27236
    
    lErro = objDoc.ObterVar("ICMS_VLR_FCP_RET", objVarICMSVlrFCPSTRet)
    If lErro <> SUCESSO Then Error 27235
    objVarICMSVlrFCPSTRet.vValor = dICMSVlrFCPSTRet
    objVarICMSVlrFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarICMSVlrFCPSTRet)
    If lErro <> SUCESSO Then Error 27236
    
    lErro = objDoc.ObterVar("IPI_VLR_DEVOLVIDO", objVarIPIVlrDevolvido)
    If lErro <> SUCESSO Then Error 27235
    objVarIPIVlrDevolvido.vValor = dIPIVlrDevolvido
    objVarIPIVlrDevolvido.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarIPIVlrDevolvido)
    If lErro <> SUCESSO Then Error 27236

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

Private Function AtualizarItem_EntradaCli(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
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
Dim objVarModalidade As New ClassVariavelCalculo, objVarModalidadeST As New ClassVariavelCalculo
Dim iModalidade As Integer, iModalidadeST As Integer, objVarProdOutrosValores As ClassVariavelCalculo
Dim objFilial As ClassFilialCliente, objVarFilFornObj As New ClassVariavelCalculo, objVarRegTrib As New ClassVariavelCalculo
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objVarEmissao As New ClassVariavelCalculo, objDASAliquotas As New ClassDASAliquotas
Dim objVarTipoICMSSimples As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo, objVarVlrCredSimples As New ClassVariavelCalculo, objVarAliCredSimples As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, objVarProdVlr As New ClassVariavelCalculo
Dim objVarPBOP As New ClassVariavelCalculo, objVarUFD As New ClassVariavelCalculo, objVarBCSTRet As New ClassVariavelCalculo
Dim objVarICMSSTRet As New ClassVariavelCalculo, objVarBCSTDest As New ClassVariavelCalculo, objVarICMSSTDest As New ClassVariavelCalculo
Dim objVarVI As New ClassVariavelCalculo, objVarMotivo As New ClassVariavelCalculo, objVarCobrAntValor As New ClassVariavelCalculo, objVarCobrAntBase As New ClassVariavelCalculo
Dim objFilCliIEIsento As New ClassVariavelCalculo, bExcecaoComSubstTrib As Boolean
Dim dVI As Double, iMotivo As Integer, dVCA As Double, dBCA As Double
Dim objVarOrigemProd As New ClassVariavelCalculo, iZerarAliquota As Integer, iZerarST As Integer, objVarFCI As New ClassVariavelCalculo
Dim objVar51ValorOp As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo, objVarValorDif As New ClassVariavelCalculo
Dim d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, bDiferimentoParcial As Boolean

'NT 2015.003 (ICMS Interestadual)
Dim objVarInterestBCUFDest As New ClassVariavelCalculo, objVarInterestPercFCPUFDest As New ClassVariavelCalculo, objVarInterestAliqUFDest As New ClassVariavelCalculo, objVarInterestAliq As New ClassVariavelCalculo
Dim objVarInterestPercPartilha As New ClassVariavelCalculo, objVarInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarInterestVlrUFDest As New ClassVariavelCalculo, objVarInterestVlrUFRemet As New ClassVariavelCalculo
Dim dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double
Dim dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double

'NFe 4.00
Dim objVarICMSInterestBCFCPUFDest As New ClassVariavelCalculo, objVarICMSvBCFCPST As New ClassVariavelCalculo, objVarICMSpFCPST As New ClassVariavelCalculo
Dim objVarICMSvFCPST As New ClassVariavelCalculo, objVarICMSvBCFCPSTRet As New ClassVariavelCalculo, objVarICMSpFCPSTRet  As New ClassVariavelCalculo
Dim objVarICMSvFCPSTRet As New ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As New ClassVariavelCalculo, objVarICMSvBCFCP As New ClassVariavelCalculo
Dim objVarICMSpFCP As New ClassVariavelCalculo, objVarICMSvFCP As New ClassVariavelCalculo
Dim dICMSInterestBCFCPUFDest As Double, dICMSvBCFCPST As Double, dICMSpFCPST As Double
Dim dICMSvFCPST As Double, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double
Dim dICMSvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dICMSvBCFCP As Double
Dim dICMSpFCP As Double, dICMSvFCP As Double, dAliquotaPadraoFCP As Double

'Base Dupla
Dim objVarICMSSTBaseDupla As New ClassVariavelCalculo, objVarICMSInterestBaseDupla As New ClassVariavelCalculo
Dim iICMSSTBaseDupla As Integer, iICMSInterestBaseDupla As Integer
Dim objVarcBenef As New ClassVariavelCalculo, scBenef As String

On Error GoTo Erro_AtualizarItem_EntradaCli
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    bDiferimentoParcial = False
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
    iModalidade = 0
    iModalidadeST = 0
    bExcecaoComSubstTrib = False
    
    iZerarAliquota = 0
    iZerarST = 0
    
    'NFe 4.00
    dICMSInterestBCFCPUFDest = 0
    dICMSvBCFCPST = 0
    dICMSpFCPST = 0
    dICMSvFCPST = 0
    dICMSvBCFCPSTRet = 0
    dICMSpFCPSTRet = 0
    dICMSvFCPSTRet = 0
    dICMSSTCobrAntAliquota = 0
    dICMSvBCFCP = 0
    dICMSpFCP = 0
    dICMSvFCP = 0
    scBenef = ""
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, objVarICMSInterestBaseDupla, objVarcBenef)
    If lErro <> SUCESSO Then Error 27912
    
    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27912
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarOrigemProd.vValor = objProduto.iOrigemMercadoria
            objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    objVarRegTrib.sIdentificador = "REGIME_TRIBUTARIO"
    'Se é uma nota interna
    If mvariNotaInterna = MARCADO Then
        'Se não é do Simples
        If mvariSimplesNacional = DESMARCADO Then
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_NORMAL
        Else
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES
        End If
    Else
        lErro = objDoc.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilFornObj)
        If lErro <> SUCESSO Then Error 27916
            
        Set objFilial = objVarFilFornObj.vValor
        
        objVarRegTrib.vValor = objFilial.iRegimeTributario
    End If
    objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 27916
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 And mvariZerarTributoST <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27916
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27920
    
    'NFe 4.00
    dICMSpFCP = objVarICMSpFCP.vValor
    dICMSpFCPST = objVarICMSpFCPST.vValor
    dICMSpFCPSTRet = objVarICMSpFCPSTRet.vValor
    'NFe 4.00
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        dVI = objVarVI.vValor
        dVCA = objVarCobrAntValor.vValor
        dBCA = objVarCobrAntBase.vValor
        dICMSSTCobrAntAliquota = objVarICMSSTCobrAntAliquota.vValor
        If objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_MANUAL Then
            dVCA = Arredonda_Moeda(dBCA * dICMSSTCobrAntAliquota)
        End If
        iMotivo = objVarMotivo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                 objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, iTipo = ICMS_TIPO_COBR_SUBST_ANT)
       'Manter campos próprios desses tipos de tributação
        If iTipo = ICMS_TIPO_COBR_SUBST_ANT Then
            objVarCobrAntBase.vValor = dBCA
            objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
            objVarCobrAntValor.vValor = dVCA
        
            'NFe 4.00
            If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvBCFCPSTRet = dBCA
                objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvBCFCPSTRet = objVarICMSvBCFCPST.vValor
            End If
            
            If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvFCPSTRet = Arredonda_Moeda(dICMSvBCFCPSTRet * dICMSpFCPSTRet)
                objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvFCPSTRet = objVarICMSvFCPSTRet.vValor
            End If
            'NFe 4.00
        Else
            objVarVI.vValor = dVI
            objVarMotivo.vValor = iMotivo
        End If
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    iModalidade = objVarModalidade.vValor
    iModalidadeST = objVarModalidadeST.vValor
        
    'obter UF Cliente
    lErro = objDocItem.objDoc.ObterVar("CLIENTE_UF", objVarUFCli)
    If lErro <> SUCESSO Then Error 27921
    
    sUFCli = objVarUFCli.vValor
    If sUFCli = "" Then sUFCli = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
    If lErro <> SUCESSO Then Error 27922
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_CLIENTE_IEISENTO", objFilCliIEIsento)
    If lErro <> SUCESSO Then Error 27722
    
    iICMSInterestBaseDupla = objVarICMSInterestBaseDupla.vValor
    If objVarICMSInterestBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_INTER", mvarsUFFilialEmpresa, mvardtDataRef, iICMSInterestBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    iICMSSTBaseDupla = objVarICMSSTBaseDupla.vValor
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_ST", mvarsUFFilialEmpresa, mvardtDataRef, iICMSSTBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    'se filial cliente inscrita ICMS e interestadual
    If objFilCliIEIsento.vValor = DESMARCADO And sUFCli <> mvarsUFFilialEmpresa And sUFCli <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFCli, dICMSpFCP, objVarICMSpFCP)
        If lErro <> SUCESSO Then Error 27923
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, sUFCli, objFilCliIEIsento.vValor = MARCADO)
        If lErro <> SUCESSO Then Error 27924
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dValorPauta <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                If lErro <> SUCESSO Then Error 27926
            End If
        End If
    
    Else
        If sUFCli <> "EX" Then
            'tratar como operacao interna
        
            'tentar achar excecao
            lErro = ExcecaoCli_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, mvarsUFFilialEmpresa, objFilCliIEIsento.vValor = MARCADO)
            If lErro <> SUCESSO Then Error 27925
            
            'se achou excecao
            If (Not (objExcecao Is Nothing)) Then
            
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, objFilCliIEIsento.vValor = MARCADO, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                If lErro <> SUCESSO Then Error 27926
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, mvarsUFFilialEmpresa, dICMSpFCP, objVarICMSpFCP)
                If lErro <> SUCESSO Then Error 27927
            
            End If
            
        Else  'importacao
        
            lErro = AtualizarItem_TratarNaoExcecaoEntCli(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFCli, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 27928
        
        End If
        
    End If
    
    If objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES Then
    
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        
            iTipo = ICMS_TIPO_NAO_TRIBUTADO
            objVarTipo.vValor = iTipo
            objTipoTribICMS.iTipo = iTipo
            lErro = TipoICMS_Obter(objTipoTribICMS)
            If lErro <> SUCESSO Then Error 27928
    
        End If
        
    End If
    
    If objTipoTribICMS.iPermiteST = MARCADO And (dMargSubst <> 0 Or dValorPauta <> 0) Then bExcecaoComSubstTrib = True

    If objTipoTribICMS.iPermiteReducaoBase <> 1 Then
        dPercRedBase = 0
        dPercRedBaseSubst = 0
    End If
    
    If objTipoTribICMS.iPermiteMargLucro <> 1 Then
        dMargSubst = 0
    End If
    
    'Calcula a permissão de crédito de ICMS
    'VAI ESTAR FALTANDO ZERAR OUTROS VALORES COMO A BASE E ALÍQUOTA NORMAL QUANDO NÃO PERMITE, ETC
    If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        objVarAliCredSimples.vValor = 0
        objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        objVarVlrCredSimples.vValor = 0
        objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        
        If objTipoTribICMS.iPermiteMotDesoneracao = DESMARCADO Then
            objVarMotivo.vValor = 0
            objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVI.vValor = 0
            objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermitePercBaseOperProp = DESMARCADO Then
            objVarPBOP.vValor = 0
            objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTUFDest = DESMARCADO Then
            objVarUFD.vValor = ""
            objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoTribICMS.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    Else
        
        'Se não é manual ou não está preenchido calcula
        If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Or objVarTipoICMSSimples.vValor = 0 Then
        
            If (Not (objExcecao Is Nothing)) Then
                objVarTipoICMSSimples.vValor = objExcecao.iTipoSimples
                If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            Else
                objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            End If
            If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = 1
            objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
            Select Case objVarTipoICMSSimples.vValor
                Case 1, 2, 4, 5
                    If mvariNotaInterna = MARCADO Then
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                        End If
                    Else
                        If mvariInscrita = ICMS_NAO_INSCRITA Or mvariSimplesNacional = MARCADO Or objVarTipoICMSSimples.vValor = 2 Or objVarTipoICMSSimples.vValor = 5 Then
                            If bExcecaoComSubstTrib Then
                                objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                            Else
                                objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                            End If
                        Else
                            If bExcecaoComSubstTrib Then
                                objVarTipoICMSSimples.vValor = 4 'Com crédito - Com ST
                            Else
                                objVarTipoICMSSimples.vValor = 1 'Com crédito - Sem ST
                            End If
                        End If
                    End If
                Case 3, 6
                    If bExcecaoComSubstTrib Then
                        objVarTipoICMSSimples.vValor = 6 'Com ST
                    Else
                        objVarTipoICMSSimples.vValor = 3 'Sem ST
                    End If
            End Select
        End If
          
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        Set objVarItemTipoObj.vValor = objTipoICMSSimples
        lErro = objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
        If lErro <> SUCESSO Then Error 27320

        Set objTipoICMSSimples = objVarItemTipoObj.vValor
       
        If objTipoICMSSimples.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoICMSSimples.iPermiteCredito = DESMARCADO Then
            objVarAliCredSimples.vValor = 0
            objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVlrCredSimples.vValor = 0
            objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoICMSSimples.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoICMSSimples.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    
    End If
    'Calcula a permissão de crédito de ICMS
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27929

        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        'dValor = CDbl(Format(dValor, "0.0000"))
        dValor = Arredonda_Moeda(dValor)
    
    Else
        dValor = objVarValor.vValor
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
    
        If bExcecaoComSubstTrib Then 'dMargSubst <> 0 Or dValorPauta <> 0 Then
            
            If iUsaPauta = 0 Then
            
                 dBaseSubst = dBaseCalculo
                                 
                'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
                If objTipoTrib.iICMSBaseComIPI = 0 Then
                
                    'obter valor do IPI do item de objdocitem
                    lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                    If lErro <> SUCESSO Then Error 27325
                    
                    dBaseSubst = dBaseSubst + objVar.vValor
                
                End If

                If iICMSSTBaseDupla = MARCADO Then
                
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
        
                    dBaseSubst = (dBaseSubst - dValor) / (1 - dAliqSubst)
                End If
                
                dBaseSubst = dBaseSubst * (1 + dMargSubst)
                dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
'                dBaseSubst = CDbl(Format(dBaseSubst, "0.0000"))
'                dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
                dBaseSubst = Arredonda_Moeda(dBaseSubst)
                dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida)
            Else
            
                'obter qtde do item de objdocitem
                lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                'dBaseSubst = CDbl(Format(objVar.vValor * dValorPauta, "0.0000"))
                dBaseSubst = Arredonda_Moeda(objVar.vValor * dValorPauta)
                dBaseSubstReduzida = dBaseSubst
                
                iModalidade = ICMS_ST_MODALIDADE_PAUTA
            
            End If
        
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dMargSubst = objVarMargSubst.vValor
        'dBaseSubst = dBaseSubst * (1 + dMargSubst)
        dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
    End If
    
    If iZerarST Then
        dBaseSubst = 0
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dMargSubst = 0
        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dPercRedBaseSubst = 0
        objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dBaseSubstReduzida = 0
        
        objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        dICMSpFCPST = 0
    End If
    
    'se é para zerar a parte propria (mantendo apenas a subst tributaria)
    If iZerarAliquota = MARCADO Or iTipo = ICMS_TIPO_COM_DIFERIMENTO Then
        If iZerarAliquota = MARCADO Then
            objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        
        'http://portalnfe.fazenda.mg.gov.br/respostas_III.html
        'Na versão 3.10 do leiaute da NF-e foram criados campos para informação do diferimento parcial.
        'Esses campos também deverão ser preenchidos no caso do diferimento total?
        'Os campos modBC (Modalidade de determinação da base de cálculo do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação), pDif (Percentual do diferimento), vICMSDif (Valor do ICMS diferido) e vICMS (Valor do ICMS)
        'só devem ser preenchidos no caso do diferimento ser parcial.
        'Esses campos NÃO serão de preenchimento obrigatório para os casos do diferimento total.
        'Alertamos que os campos modBC (Modalidade de determinação da BC do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação) devem ser preenchidos com o valor como se não tivesse o diferimento.

        If objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO And mvariSimplesNacional = 0 And mvariZerarTributo = 0 And iZerarAliquota <> MARCADO Then
        
            If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
                dPercDifer = 1 '??? deveria vir da excecao
                objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dPercDifer = objVarPercDifer.vValor
            End If
            
            If dPercDifer <> 1 And dPercDifer <> 0 Then bDiferimentoParcial = True
        
            If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
                d51ValorOp = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
                objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                d51ValorOp = objVar51ValorOp.vValor
            End If
            
            If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
                dValorDif = Arredonda_Moeda(d51ValorOp * dPercDifer)
                objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dValorDif = objVarValorDif.vValor
            End If
        
        Else
        
            d51ValorOp = 0
            objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            dPercDifer = 0
            objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            dValorDif = 0
            objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Then
            
            If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then dBaseCalculo = 0
            If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then dPercRedBase = 0
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
        
            If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then dICMSpFCP = 0
        
        Else
        
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(d51ValorOp - dValorDif)
        
        End If
        
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
        
        'NFe 4.00
        If objVarICMSpFCPST.iStatus <> VAR_PREENCH_MANUAL Then
            If Not (objExcecao Is Nothing) Then
                dICMSpFCPST = objExcecao.dICMSPercFCP
            End If
        Else
            dICMSpFCPST = objVarICMSpFCPST.vValor
        End If
        'NFe 4.00
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dValorSubst = dBaseSubstReduzida * dAliqSubst - dValor
            'dValorSubst = CDbl(Format(dValorSubst, "0.0000"))
            dValorSubst = Arredonda_Moeda(dValorSubst)
        Else
            dValorSubst = objVarValorSubst.vValor
        End If
        
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
    End If
    
    'NFe 4.00
    'Base de cálculo dos FCP iguais as bases correspondentes do ICMS
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCP = dBaseCalculo
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCP = objVarICMSvBCFCP.vValor
    End If
    
    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCPST = dBaseSubst
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCPST = objVarICMSvBCFCPST.vValor
    End If
           
    'Calculando o valor com a base e alíquota
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCP = Arredonda_Moeda(dICMSvBCFCP * dICMSpFCP)
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCP = objVarICMSvFCP.vValor
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCPST = Arredonda_Moeda(dICMSvBCFCPST * dICMSpFCPST)
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCPST = objVarICMSvFCPST.vValor
    End If
    'NFe 4.00

    'Se credita e não é com diferimento total
    If objTipoTrib.iICMSCredita <> 0 And (objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Or bDiferimentoParcial) And (mvariSimplesNacional = 0 Or mvardtDataRef < DATA_INICIO_SIMPLES_NACIONAL) Then
    
        dValorCredito = dValor
        
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubstReduzida, dPercRedBaseSubst, dValorCredito, iModalidade, iModalidadeST, _
                                objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarRegTrib, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, d51ValorOp, dPercDifer, dValorDif, iMotivo)
    If lErro <> SUCESSO Then Error 27931
    
    lErro = AtualizarItem_GuardarValores2(objDocItem, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                dInterestBCUFDest, dInterestPercFCPUFDest, dInterestAliqUFDest, dInterestAliq, dInterestPercPartilha, dInterestVlrFCPUFDest, dInterestVlrUFDest, dInterestVlrUFRemet, objVarICMSInterestBCFCPUFDest, dICMSInterestBCFCPUFDest)
    If lErro <> SUCESSO Then Error 27931
    
    lErro = AtualizarItem_GuardarValores3(objDocItem, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, dICMSvBCFCP, dICMSpFCP, dICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, dICMSvBCFCPST, dICMSpFCPST, dICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, dICMSvBCFCPSTRet, dICMSpFCPSTRet, dICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, dICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, iICMSSTBaseDupla, objVarICMSInterestBaseDupla, iICMSInterestBaseDupla, objVarcBenef, scBenef)
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

Private Function AliqInterstEntrada_Pesquisar(sSiglaOrigem As String, dAliquota As Double, dAliquotaFCP As Double) As Long
'retorna a aliquota interestadual chegando ao estado da filial empresa corrente de sSiglaOrigem
Dim lErro As Long, objICMSAliqExterna As ClassICMSAliqExterna
On Error GoTo Erro_AliqInterstEntrada_Pesquisar
    
    Set objICMSAliqExterna = mvarcolAliqInterstEntrada.Item(sSiglaOrigem)
    
    dAliquota = objICMSAliqExterna.dAliquota
    dAliquotaFCP = objICMSAliqExterna.dAliquotaFCP
    
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

Private Function ObterAliquotaBasicaEntCli(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFCli As String, ByVal objDocItem As ClassTributoDocItem, dAliquotaFCP As Double) As Long
'preenche a aliquota
Dim lErro As Long
Dim objFilCliInscEst As New ClassVariavelCalculo
Dim objFilCliIEIsento As New ClassVariavelCalculo

On Error GoTo Erro_ObterAliquotaBasicaEntCli

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFCli = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            dAliquotaFCP = mvardAliqInternaFCP
            
        Else
          
            'se for importacao
            If sUFCli = "EX" Then
            
                dAliquota = mvardAliqImportacao
                dAliquotaFCP = mvardAliqInternaFCP
                
            Else 'se for interestadual
                       
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_INSC_EST", objFilCliInscEst)
                If lErro <> SUCESSO Then Error 27904
    
                lErro = objDoc.ObterVar("FILIAL_CLIENTE_IEISENTO", objFilCliIEIsento)
                If lErro <> SUCESSO Then Error 27904
    
                'se filial cliente nao inscrita ICMS
                If objFilCliIEIsento.vValor = MARCADO Then
                    
                    dAliquota = 0
                    dAliquotaFCP = 0
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstEntrada_Pesquisar(sUFCli, dAliquota, dAliquotaFCP)
                    If lErro <> SUCESSO Then Error 27905
                
                End If
            End If
        End If
    Else
        dAliquota = 0
        dAliquotaFCP = 0
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

Private Function ObterAliquotaBasicaEntForn(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFForn As String, ByVal objDocItem As ClassTributoDocItem, dAliquotaFCP As Double) As Long
'preenche a aliquota
Dim lErro As Long
Dim objInscEst As New ClassVariavelCalculo
Dim objIEIsento As New ClassVariavelCalculo

On Error GoTo Erro_ObterAliquotaBasicaEntForn

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFForn = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            dAliquotaFCP = mvardAliqInternaFCP
            
        Else
          
            'se for importacao
            If sUFForn = "EX" Then
            
                dAliquota = mvardAliqImportacao
                dAliquotaFCP = mvardAliqInternaFCP
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_FORN_INSC_EST", objInscEst)
                If lErro <> SUCESSO Then Error 27906
                
                lErro = objDoc.ObterVar("FILIAL_FORN_IEISENTO", objIEIsento)
                If lErro <> SUCESSO Then Error 27906
                
                'se filial fornecedor nao inscrita ICMS
                If objIEIsento.vValor = MARCADO Then
                    
                    dAliquota = 0
                    dAliquotaFCP = 0
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstEntrada_Pesquisar(sUFForn, dAliquota, dAliquotaFCP)
                    If lErro <> SUCESSO Then Error 27907
                
                End If
            End If
        End If
    Else
        dAliquota = 0
        dAliquotaFCP = 0
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

Private Function ObterAliquotaBasicaSaidaForn(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquota As Double, sUFForn As String, ByVal objDocItem As ClassTributoDocItem, dAliquotaFCP As Double) As Long
'preenche a aliquota
Dim lErro As Long
Dim objInscEst As New ClassVariavelCalculo
Dim objIEIsento As New ClassVariavelCalculo

On Error GoTo Erro_ObterAliquotaBasicaSaidaForn

    'se ICMS_TIPO compativel com aliquota <> 0
    If objTipoTribICMS.iPermiteAliquota <> 0 Then
    
        'se for interna
        If sUFForn = mvarsUFFilialEmpresa Then
        
            'setar aliquota igual a interna
            dAliquota = mvardAliqInterna
            dAliquotaFCP = mvardAliqInternaFCP
            
        Else
          
            'se for exportacao
            If sUFForn = "EX" Then
            
                'setar aliquota igual a zero ( exportacao nao tem ICMS )
                dAliquota = 0
                dAliquotaFCP = 0
                
            Else 'se for interestadual
            
                lErro = objDoc.ObterVar("FILIAL_FORN_INSC_EST", objInscEst)
                If lErro <> SUCESSO Then Error 27908
                
                lErro = objDoc.ObterVar("FILIAL_FORN_IEISENTO", objIEIsento)
                If lErro <> SUCESSO Then Error 27908
                
                'se filial forn nao inscrita ICMS
                If objIEIsento.vValor = MARCADO Then
                    
                    'usar aliquota interna
                    dAliquota = mvardAliqInterna
                    dAliquotaFCP = mvardAliqInternaFCP
                
                Else
                    
                    'usar aliq interestadual
                    lErro = AliqInterstSaida_Pesquisar(sUFForn, dAliquota, objDocItem, dAliquotaFCP)
                    If lErro <> SUCESSO Then Error 27909
                
                End If
            End If
        End If
    Else
        dAliquota = 0
        dAliquotaFCP = 0
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

Private Function AtualizarItem_TratarNaoExcecaoEntCli(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFCli As String, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo) As Long
Dim lErro As Long
Dim dAliquotaFCP As Double
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
        
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Or objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaEntCli(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFCli, objDocItem, dAliquotaFCP)
            If lErro <> SUCESSO Then Error 27911
        Else
            dAliquota = 0
            dAliquotaFCP = 0
        End If
'    Else
'
'        dAliquota = objVarAliquota.vValor
        
    End If
    
    dICMSpFCP = dAliquotaFCP
    
    If objVarAliquota.iStatus = VAR_PREENCH_MANUAL Then dAliquota = objVarAliquota.vValor
    If objVarICMSpFCP.iStatus = VAR_PREENCH_MANUAL Then dICMSpFCP = objVarICMSpFCP.vValor
    
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
Dim lErro As Long, dPercRedBase As Double, objVarICMSValorCredito As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo
Dim objVarCST As New ClassVariavelCalculo, objVarVI As New ClassVariavelCalculo, dValorIsento As Double

'NT 2015.003 (ICMS Interestadual)
Dim dICMSInterestVlrFCPUFDest As Double, dICMSInterestVlrUFDest As Double, dICMSInterestVlrUFRemet As Double
Dim objVarICMSInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarICMSInterestVlrUFDest As New ClassVariavelCalculo, objVarICMSInterestVlrUFRemet As New ClassVariavelCalculo

'NFE 4.00
Dim dICMSVlrFCP As Double, dICMSVlrFCPST As Double, dICMSVlrFCPSTRet As Double, dIPIVlrDevolvido As Double
Dim objVarICMSVlrFCP As New ClassVariavelCalculo, objVarICMSVlrFCPST As New ClassVariavelCalculo, objVarICMSVlrFCPSTRet As New ClassVariavelCalculo, objVarIPIVlrDevolvido As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc

    dValor = 0
    dValorSubst = 0
    dBase = 0
    dBaseReduzida = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dValorCredito = 0
    dValorIsento = 0
    
    'NT 2015.003 (ICMS Interestadual)
    dICMSInterestVlrFCPUFDest = 0
    dICMSInterestVlrUFDest = 0
    dICMSInterestVlrUFRemet = 0
    
    'NFe 4.00
    dICMSVlrFCP = 0
    dICMSVlrFCPST = 0
    dICMSVlrFCPSTRet = 0
    dIPIVlrDevolvido = 0

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
        
        lErro = objDocItem.ObterVar("ICMS_CST", objVarCST)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_PERC_DIFER", objVarPercDifer)
        If lErro <> SUCESSO Then Error 27954
               
        'se nao for diferimento total
'        If Not (right(objVarCST.vValor, 2) = "51" And (objVarPercDifer.vValor = 1 Or objVarPercDifer.vValor = 0)) Then
        
            dBase = dBase + objVarICMSBase.vValor
            dBaseReduzida = dBaseReduzida + (objVarICMSBase.vValor * (1 - dPercRedBase))
            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
            dBaseSubst = dBaseSubst + objVarICMSBaseSubst.vValor
            dPercRedBaseSubst = objVarPercRedBase.vValor
            dPercRedBaseSubst = 0 '???? ainda nao está sendo armazenado
            dBaseSubstReduzida = dBaseSubstReduzida + (objVarICMSBaseSubst.vValor * (1 - dPercRedBaseSubst))
            ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.00"))
            dValor = dValor + objVarICMSValor.vValor
            dValorSubst = dValorSubst + objVarICMSValorSubst.vValor
            dValorCredito = dValorCredito + objVarICMSValorCredito.vValor
            
'        End If
       
        
        'nfe 3.10
        lErro = objDocItem.ObterVar("ICMS_VALOR_ISENTO", objVarVI)
        If lErro <> SUCESSO Then Error 27949
        
        dValorIsento = dValorIsento + objVarVI.vValor
        
        
        lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_FCP_UF_DEST", objVarICMSInterestVlrFCPUFDest)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_UF_DEST", objVarICMSInterestVlrUFDest)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_INTEREST_VLR_UF_REMET", objVarICMSInterestVlrUFRemet)
        If lErro <> SUCESSO Then Error 27954
        
        dICMSInterestVlrFCPUFDest = dICMSInterestVlrFCPUFDest + objVarICMSInterestVlrFCPUFDest.vValor
        dICMSInterestVlrUFDest = dICMSInterestVlrUFDest + objVarICMSInterestVlrUFDest.vValor
        dICMSInterestVlrUFRemet = dICMSInterestVlrUFRemet + objVarICMSInterestVlrUFRemet.vValor
        
        'NFe 4.00
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP", objVarICMSVlrFCP)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST", objVarICMSVlrFCPST)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST_RET", objVarICMSVlrFCPSTRet)
        If lErro <> SUCESSO Then Error 27954
           
        lErro = objDocItem.ObterVar("IPI_VLR_DEVOLVIDO", objVarIPIVlrDevolvido)
        If lErro <> SUCESSO Then Error 27954
           
        dICMSVlrFCP = dICMSVlrFCP + objVarICMSVlrFCP.vValor
        dICMSVlrFCPST = dICMSVlrFCPST + objVarICMSVlrFCPST.vValor
        dICMSVlrFCPSTRet = dICMSVlrFCPSTRet + objVarICMSVlrFCPSTRet.vValor
        dIPIVlrDevolvido = dIPIVlrDevolvido + objVarIPIVlrDevolvido.vValor
                

    Next
        
    'arredondar valores globais
    dBase = Arredonda_Moeda(dBase, 2)
    dBaseReduzida = Arredonda_Moeda(dBaseReduzida, 2)
    dBaseSubst = Arredonda_Moeda(dBaseSubst, 2)
    dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida, 2)
    dValor = Arredonda_Moeda(dValor, 2)
    dValorSubst = Arredonda_Moeda(dValorSubst, 2)
    dValorCredito = Arredonda_Moeda(dValorCredito, 2)
        
    'nfe 3.10
    dValorIsento = Arredonda_Moeda(dValorIsento, 2)
    
    'NT 2015.003 (ICMS Interestadual)
    dICMSInterestVlrFCPUFDest = Arredonda_Moeda(dICMSInterestVlrFCPUFDest, 2)
    dICMSInterestVlrUFDest = Arredonda_Moeda(dICMSInterestVlrUFDest, 2)
    dICMSInterestVlrUFRemet = Arredonda_Moeda(dICMSInterestVlrUFRemet, 2)
    
    'nfe 4.00
    dICMSVlrFCP = Arredonda_Moeda(dICMSVlrFCP, 2)
    dICMSVlrFCPST = Arredonda_Moeda(dICMSVlrFCPST, 2)
    dICMSVlrFCPSTRet = Arredonda_Moeda(dICMSVlrFCPSTRet, 2)
    dIPIVlrDevolvido = Arredonda_Moeda(dIPIVlrDevolvido, 2)
    
    lErro = AtualizarDoc_GuardarValores(objDoc, dBaseReduzida, dValor, dBaseSubstReduzida, dValorSubst, dValorCredito, dValorIsento, dICMSInterestVlrFCPUFDest, dICMSInterestVlrUFDest, dICMSInterestVlrUFRemet, dICMSVlrFCP, dICMSVlrFCPST, dICMSVlrFCPSTRet, dIPIVlrDevolvido)
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

Private Function AtualizarItem_TratarNaoExcecaoEntForn(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFForn As String, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo) As Long
Dim lErro As Long
Dim dAliquotaFCP As Double
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
        
    'If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Or objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaEntForn(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn, objDocItem, dAliquotaFCP)
            If lErro <> SUCESSO Then Error 22701
        Else
            dAliquota = 0
            dAliquotaFCP = 0
        End If
'    Else
'
'        dAliquota = objVarAliquota.vValor
        
    End If
    
    dICMSpFCP = dAliquotaFCP
    
    If objVarAliquota.iStatus = VAR_PREENCH_MANUAL Then dAliquota = objVarAliquota.vValor
    If objVarICMSpFCP.iStatus = VAR_PREENCH_MANUAL Then dICMSpFCP = objVarICMSpFCP.vValor
    
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

Private Function AtualizarItem_EntradaForn(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
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
Dim dBaseSubstReduzida As Double
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dPercRedBaseSubst As Double
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, objVarPercRedBaseSubst As New ClassVariavelCalculo
Dim objVarModalidade As New ClassVariavelCalculo, objVarModalidadeST As New ClassVariavelCalculo
Dim iModalidade As Integer, iModalidadeST As Integer, objVarProdOutrosValores As ClassVariavelCalculo
Dim objFilial As ClassFilialFornecedor, objVarFilFornObj As New ClassVariavelCalculo, objVarRegTrib As New ClassVariavelCalculo
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objVarEmissao As New ClassVariavelCalculo, objDASAliquotas As New ClassDASAliquotas
Dim objVarTipoICMSSimples As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo, objVarVlrCredSimples As New ClassVariavelCalculo, objVarAliCredSimples As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, objVarProdVlr As New ClassVariavelCalculo
Dim objVarPBOP As New ClassVariavelCalculo, objVarUFD As New ClassVariavelCalculo, objVarBCSTRet As New ClassVariavelCalculo
Dim objVarICMSSTRet As New ClassVariavelCalculo, objVarBCSTDest As New ClassVariavelCalculo, objVarICMSSTDest As New ClassVariavelCalculo
Dim objVarVI As New ClassVariavelCalculo, objVarMotivo As New ClassVariavelCalculo, objVarCobrAntValor As New ClassVariavelCalculo, objVarCobrAntBase As New ClassVariavelCalculo
Dim dAliquotaAux As Double, dPercRedBaseAux As Double
Dim objFilFornIEIsento As New ClassVariavelCalculo, bExcecaoComSubstTrib As Boolean
Dim dVI As Double, iMotivo As Integer, dVCA As Double, dBCA As Double
Dim objVarOrigemProd As New ClassVariavelCalculo, iZerarAliquota As Integer, iZerarST As Integer
Dim dBaseICMSProprioSTSimples As Double, dICMSProprioSTSimples As Double, dAliquotaPadrao As Double, objVarFCI As New ClassVariavelCalculo
Dim objVar51ValorOp As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo, objVarValorDif As New ClassVariavelCalculo
Dim d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, bDiferimentoParcial As Boolean

'NT 2015.003 (ICMS Interestadual)
Dim objVarInterestBCUFDest As New ClassVariavelCalculo, objVarInterestPercFCPUFDest As New ClassVariavelCalculo, objVarInterestAliqUFDest As New ClassVariavelCalculo, objVarInterestAliq As New ClassVariavelCalculo
Dim objVarInterestPercPartilha As New ClassVariavelCalculo, objVarInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarInterestVlrUFDest As New ClassVariavelCalculo, objVarInterestVlrUFRemet As New ClassVariavelCalculo
Dim dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double
Dim dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double

'NFe 4.00
Dim objVarICMSInterestBCFCPUFDest As New ClassVariavelCalculo, objVarICMSvBCFCPST As New ClassVariavelCalculo, objVarICMSpFCPST As New ClassVariavelCalculo
Dim objVarICMSvFCPST As New ClassVariavelCalculo, objVarICMSvBCFCPSTRet As New ClassVariavelCalculo, objVarICMSpFCPSTRet  As New ClassVariavelCalculo
Dim objVarICMSvFCPSTRet As New ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As New ClassVariavelCalculo, objVarICMSvBCFCP As New ClassVariavelCalculo
Dim objVarICMSpFCP As New ClassVariavelCalculo, objVarICMSvFCP As New ClassVariavelCalculo
Dim dICMSInterestBCFCPUFDest As Double, dICMSvBCFCPST As Double, dICMSpFCPST As Double
Dim dICMSvFCPST As Double, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double
Dim dICMSvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dICMSvBCFCP As Double
Dim dICMSpFCP As Double, dICMSvFCP As Double, dAliquotaPadraoFCP As Double

'Base Dupla
Dim objVarICMSSTBaseDupla As New ClassVariavelCalculo, objVarICMSInterestBaseDupla As New ClassVariavelCalculo
Dim iICMSSTBaseDupla As Integer, iICMSInterestBaseDupla As Integer
Dim objVarcBenef As New ClassVariavelCalculo, scBenef As String
Dim dValorPauta As Double, iUsaPauta As Integer
On Error GoTo Erro_AtualizarItem_EntradaForn
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    bDiferimentoParcial = False
    dICMSProprioSTSimples = 0
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
    iModalidade = 0
    iModalidadeST = 0
    bExcecaoComSubstTrib = False
    
    iZerarAliquota = 0
    iZerarST = 0
    
    'NFe 4.00
    dICMSInterestBCFCPUFDest = 0
    dICMSvBCFCPST = 0
    dICMSpFCPST = 0
    dICMSvFCPST = 0
    dICMSvBCFCPSTRet = 0
    dICMSpFCPSTRet = 0
    dICMSvFCPSTRet = 0
    dICMSSTCobrAntAliquota = 0
    dICMSvBCFCP = 0
    dICMSpFCP = 0
    dICMSvFCP = 0
    scBenef = ""
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                                 objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, objVarICMSInterestBaseDupla, objVarcBenef)
    If lErro <> SUCESSO Then Error 27974
    
    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27974
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarOrigemProd.vValor = objProduto.iOrigemMercadoria
            objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    objVarRegTrib.sIdentificador = "REGIME_TRIBUTARIO"
    'Se é uma nota interna
    If mvariNotaInterna = MARCADO Then
        'Se não é do Simples
        If mvariSimplesNacional = DESMARCADO Then
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_NORMAL
        Else
            objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES
        End If
    Else
        lErro = objDoc.ObterVar("FILIAL_FORN_OBJ", objVarFilFornObj)
        If lErro <> SUCESSO Then Error 27978
            
        Set objFilial = objVarFilFornObj.vValor
        
        objVarRegTrib.vValor = objFilial.iRegimeTributario
    End If
    objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 27916
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 And mvariZerarTributoST <> 0 Then

        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                 objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)

        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 27978
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27978
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27985
    
    'NFe 4.00
    dICMSpFCP = objVarICMSpFCP.vValor
    dICMSpFCPST = objVarICMSpFCPST.vValor
    dICMSpFCPSTRet = objVarICMSpFCPSTRet.vValor
    'NFe 4.00
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                 objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        dVI = objVarVI.vValor
        dVCA = objVarCobrAntValor.vValor
        dBCA = objVarCobrAntBase.vValor
        dICMSSTCobrAntAliquota = objVarICMSSTCobrAntAliquota.vValor
        If objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_MANUAL Then
            dVCA = Arredonda_Moeda(dBCA * dICMSSTCobrAntAliquota)
        End If
        iMotivo = objVarMotivo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                 objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, iTipo = ICMS_TIPO_COBR_SUBST_ANT)
        'Manter campos próprios desses tipos de tributação
        If iTipo = ICMS_TIPO_COBR_SUBST_ANT Then
            objVarCobrAntBase.vValor = dBCA
            objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
            objVarCobrAntValor.vValor = dVCA
        
            'NFe 4.00
            If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvBCFCPSTRet = dBCA
                objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvBCFCPSTRet = objVarICMSvBCFCPST.vValor
            End If
            
            If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvFCPSTRet = Arredonda_Moeda(dICMSvBCFCPSTRet * dICMSpFCPSTRet)
                objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvFCPSTRet = objVarICMSvFCPSTRet.vValor
            End If
            'NFe 4.00
        Else
            objVarVI.vValor = dVI
            objVarMotivo.vValor = iMotivo
        End If
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    iModalidade = objVarModalidade.vValor
    iModalidadeST = objVarModalidadeST.vValor
        
    'obter UF Forn
    lErro = objDocItem.objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 27986
    
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_INSC_EST", objFilFornInscEst)
    If lErro <> SUCESSO Then Error 27987

    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_IEISENTO", objFilFornIEIsento)
    If lErro <> SUCESSO Then Error 27987
    
    iICMSInterestBaseDupla = objVarICMSInterestBaseDupla.vValor
    If objVarICMSInterestBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_INTER", mvarsUFFilialEmpresa, mvardtDataRef, iICMSInterestBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    iICMSSTBaseDupla = objVarICMSSTBaseDupla.vValor
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_ST", mvarsUFFilialEmpresa, mvardtDataRef, iICMSSTBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    'se filial forn inscrita ICMS e interestadual
    If objFilFornIEIsento.vValor = DESMARCADO And sUFForn <> mvarsUFFilialEmpresa And sUFForn <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
        If lErro <> SUCESSO Then Error 27988
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, sUFForn)
        If lErro <> SUCESSO Then Error 27989
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                '??? atencao: usando dAliquotaAux e dPercRedBaseAux para manter a utilizacao dos valores interestaduais para a operacao basica...
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquotaAux, objVarPercRedBase, dPercRedBaseAux, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
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
            
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                If lErro <> SUCESSO Then Error 27991
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
                If lErro <> SUCESSO Then Error 27992
            
            End If
            
        Else  'importacao
        
            lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 27993
        
        End If
        
    End If
    
    If objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES Then
    
        'se ICMS_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL And Not ((Not (objExcecao Is Nothing)) And mvariNotaInterna = 0) Then
        
            iTipo = ICMS_TIPO_NAO_TRIBUTADO
            objVarTipo.vValor = iTipo
            objTipoTribICMS.iTipo = iTipo
            lErro = TipoICMS_Obter(objTipoTribICMS)
            If lErro <> SUCESSO Then Error 27993
    
        End If
        
    End If

    If objTipoTribICMS.iPermiteST = MARCADO And (dMargSubst <> 0) Then bExcecaoComSubstTrib = True

    If objTipoTribICMS.iPermiteReducaoBase <> 1 Then
        dPercRedBase = 0
        dPercRedBaseSubst = 0
    End If
    
    If objTipoTribICMS.iPermiteMargLucro <> 1 Then
        dMargSubst = 0
    End If
    
    'Calcula a permissão de crédito de ICMS
    'VAI ESTAR FALTANDO ZERAR OUTROS VALORES COMO A BASE E ALÍQUOTA NORMAL QUANDO NÃO PERMITE, ETC
    If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        objVarAliCredSimples.vValor = 0
        objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        objVarVlrCredSimples.vValor = 0
        objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        
        If objTipoTribICMS.iPermiteMotDesoneracao = DESMARCADO Then
            objVarMotivo.vValor = 0
            objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVI.vValor = 0
            objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermitePercBaseOperProp = DESMARCADO Then
            objVarPBOP.vValor = 0
            objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTUFDest = DESMARCADO Then
            objVarUFD.vValor = ""
            objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoTribICMS.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    Else
        
        'Se não é manual ou não está preenchido calcula
        If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Or objVarTipoICMSSimples.vValor = 0 Then
        
            If (Not (objExcecao Is Nothing)) Then
                objVarTipoICMSSimples.vValor = objExcecao.iTipoSimples
                If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            Else
                objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            End If
            If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = 1
            objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
            Select Case objVarTipoICMSSimples.vValor
                Case 1, 2, 4, 5
                    If mvariNotaInterna = MARCADO Then
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                        End If
                    Else
                        If mvariInscrita = ICMS_NAO_INSCRITA Or mvariSimplesNacional = MARCADO Or objVarTipoICMSSimples.vValor = 2 Or objVarTipoICMSSimples.vValor = 5 Then
                            If bExcecaoComSubstTrib Then
                                objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                            Else
                                objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                            End If
                        Else
                            If bExcecaoComSubstTrib Then
                                objVarTipoICMSSimples.vValor = 4 'Com crédito - Com ST
                            Else
                                objVarTipoICMSSimples.vValor = 1 'Com crédito - Sem ST
                            End If
                        End If
                    End If
                Case 3, 6
                    If bExcecaoComSubstTrib Then
                        objVarTipoICMSSimples.vValor = 6 'Com ST
                    Else
                        objVarTipoICMSSimples.vValor = 3 'Sem ST
                    End If
            End Select
        End If
          
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        Set objVarItemTipoObj.vValor = objTipoICMSSimples
        lErro = objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
        If lErro <> SUCESSO Then Error 27320

        Set objTipoICMSSimples = objVarItemTipoObj.vValor
       
        If objTipoICMSSimples.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoICMSSimples.iPermiteCredito = DESMARCADO Then
            objVarAliCredSimples.vValor = 0
            objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVlrCredSimples.vValor = 0
            objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoICMSSimples.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoICMSSimples.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    
    End If
    'Calcula a permissão de crédito de ICMS
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
       
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then

            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27994

        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = Arredonda_Moeda(dValor)
    
    Else
        If dBaseCalculo <> 0 Then
            dValor = objVarValor.vValor
        Else
            dValor = 0
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        End If
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        If bExcecaoComSubstTrib Then 'dMargSubst <> 0 Then
            
            'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI = 0 Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = (dBaseCalculo + objVar.vValor) '* (1 + dMargSubst)
            Else
                dBaseSubst = dBaseCalculo '* (1 + dMargSubst)
            End If
            
             If iICMSSTBaseDupla = MARCADO Then
            
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
    
                dBaseSubst = (dBaseSubst - dValor) / (1 - dAliqSubst)
            End If
            dBaseSubst = dBaseSubst * (1 + dMargSubst)
            
            dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
            dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
            'dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida)
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dMargSubst = objVarMargSubst.vValor
        'dBaseSubst = dBaseSubst * (1 + dMargSubst)
        dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
    End If
    
    If iZerarST Then
        dBaseSubst = 0
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dMargSubst = 0
        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dPercRedBaseSubst = 0
        objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dBaseSubstReduzida = 0
        
        objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        dICMSpFCPST = 0
    End If
    
    'se é para zerar a parte propria (mantendo apenas a subst tributaria)
    If iZerarAliquota = MARCADO Or iTipo = ICMS_TIPO_COM_DIFERIMENTO Then

        If objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES Then
        
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseICMSProprioSTSimples)
            If lErro <> SUCESSO Then Error 27324
        
            dAliquotaPadrao = 0
            If sUFForn = mvarsUFFilialEmpresa Then
                
                'se ICMS_ALIQUOTA_SUBST nao preenchido manualmente
                If objVarAliqSubst.iStatus <> VAR_PREENCH_MANUAL Then
                    If objExcecao Is Nothing Then
                        dAliquotaPadrao = 0
                    Else
                        dAliquotaPadrao = objExcecao.dAliquota
                    End If
                Else
                    dAliquotaPadrao = objVarAliqSubst.vValor
                End If
            
                dICMSProprioSTSimples = dBaseICMSProprioSTSimples * Arredonda_Moeda((1 - dPercRedBaseSubst) * dAliquotaPadrao, 2)
            
            Else
                
                lErro = ObterAliquotaBasicaEntForn(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquotaPadrao, mvarsUFFilialEmpresa, objDocItem, dAliquotaPadraoFCP)
                If lErro <> SUCESSO Then Error 27318
            
                dICMSProprioSTSimples = dBaseICMSProprioSTSimples * Arredonda_Moeda((1 - dPercRedBase) * dAliquotaPadrao, 2)
            
            End If
            
        End If
    
        If iZerarAliquota = MARCADO Then
            objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        
        'http://portalnfe.fazenda.mg.gov.br/respostas_III.html
        'Na versão 3.10 do leiaute da NF-e foram criados campos para informação do diferimento parcial.
        'Esses campos também deverão ser preenchidos no caso do diferimento total?
        'Os campos modBC (Modalidade de determinação da base de cálculo do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação), pDif (Percentual do diferimento), vICMSDif (Valor do ICMS diferido) e vICMS (Valor do ICMS)
        'só devem ser preenchidos no caso do diferimento ser parcial.
        'Esses campos NÃO serão de preenchimento obrigatório para os casos do diferimento total.
        'Alertamos que os campos modBC (Modalidade de determinação da BC do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação) devem ser preenchidos com o valor como se não tivesse o diferimento.

        If objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO And mvariSimplesNacional = 0 And mvariZerarTributo = 0 And iZerarAliquota <> MARCADO Then
        
            If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
                dPercDifer = 1 '??? deveria vir da excecao
                objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dPercDifer = objVarPercDifer.vValor
            End If
            
            If dPercDifer <> 1 And dPercDifer <> 0 Then bDiferimentoParcial = True
        
            If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
                d51ValorOp = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
                objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                d51ValorOp = objVar51ValorOp.vValor
            End If
            
            If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
                dValorDif = Arredonda_Moeda(d51ValorOp * dPercDifer)
                objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dValorDif = objVarValorDif.vValor
            End If
        
        Else
        
            d51ValorOp = 0
            objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            dPercDifer = 0
            objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            dValorDif = 0
            objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Then
            
            If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then dBaseCalculo = 0
            If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then dPercRedBase = 0
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
        
            If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then dICMSpFCP = 0
        
        Else
        
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(d51ValorOp - dValorDif)
        
        End If
        
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
        
        'NFe 4.00
        If objVarICMSpFCPST.iStatus <> VAR_PREENCH_MANUAL Then
            If Not (objExcecao Is Nothing) Then
                dICMSpFCPST = objExcecao.dICMSPercFCP
            End If
        Else
            dICMSpFCPST = objVarICMSpFCPST.vValor
        End If
        'NFe 4.00
        
        'se ICMS_VALOR_SUBST nao preenchido manualmente
        If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
            If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
                dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dValor
            Else
                dValorSubst = dBaseSubst * (1 - dPercRedBaseSubst) * dAliqSubst - dICMSProprioSTSimples
            End If
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
    
    'NFe 4.00
    'Base de cálculo dos FCP iguais as bases correspondentes do ICMS
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCP = dBaseCalculo
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCP = objVarICMSvBCFCP.vValor
    End If
    
    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCPST = dBaseSubst
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCPST = objVarICMSvBCFCPST.vValor
    End If
           
    'Calculando o valor com a base e alíquota
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCP = Arredonda_Moeda(dICMSvBCFCP * dICMSpFCP)
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCP = objVarICMSvFCP.vValor
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCPST = Arredonda_Moeda(dICMSvBCFCPST * dICMSpFCPST)
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCPST = objVarICMSvFCPST.vValor
    End If
    'NFe 4.00

    'Se credita e não é com diferimento total
    If dValorSubst = 0 And objTipoTrib.iICMSCredita <> 0 And (objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Or bDiferimentoParcial) And (mvariSimplesNacional = 0 Or mvardtDataRef < DATA_INICIO_SIMPLES_NACIONAL) Then
    
        dValorCredito = dValor
        
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito, iModalidade, iModalidadeST, _
                                objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarRegTrib, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, d51ValorOp, dPercDifer, dValorDif, iMotivo)
    If lErro <> SUCESSO Then Error 27996
    
    lErro = AtualizarItem_GuardarValores2(objDocItem, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                dInterestBCUFDest, dInterestPercFCPUFDest, dInterestAliqUFDest, dInterestAliq, dInterestPercPartilha, dInterestVlrFCPUFDest, dInterestVlrUFDest, dInterestVlrUFRemet, objVarICMSInterestBCFCPUFDest, dICMSInterestBCFCPUFDest)
    If lErro <> SUCESSO Then Error 27996
    
    lErro = AtualizarItem_GuardarValores3(objDocItem, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, dICMSvBCFCP, dICMSpFCP, dICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, dICMSvBCFCPST, dICMSpFCPST, dICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, dICMSvBCFCPSTRet, dICMSpFCPSTRet, dICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, dICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, iICMSSTBaseDupla, objVarICMSInterestBaseDupla, iICMSInterestBaseDupla, objVarcBenef, scBenef)
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

Private Function AtualizarItem_SaidaForn(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
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
Dim dBaseSubstReduzida As Double
Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, sUFDestino As String, objVarPercRedBaseSubst As New ClassVariavelCalculo
Dim objVarModalidade As New ClassVariavelCalculo, objVarModalidadeST As New ClassVariavelCalculo
Dim iModalidade As Integer, iModalidadeST As Integer, objVarProdOutrosValores As ClassVariavelCalculo, objVarRegTrib As New ClassVariavelCalculo
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objVarEmissao As New ClassVariavelCalculo, objDASAliquotas As New ClassDASAliquotas
Dim objVarTipoICMSSimples As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo, objVarVlrCredSimples As New ClassVariavelCalculo, objVarAliCredSimples As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, objVarProdVlr As New ClassVariavelCalculo
Dim objVarPBOP As New ClassVariavelCalculo, objVarUFD As New ClassVariavelCalculo, objVarBCSTRet As New ClassVariavelCalculo
Dim objVarICMSSTRet As New ClassVariavelCalculo, objVarBCSTDest As New ClassVariavelCalculo, objVarICMSSTDest As New ClassVariavelCalculo
Dim objVarVI As New ClassVariavelCalculo, objVarMotivo As New ClassVariavelCalculo, objVarCobrAntValor As New ClassVariavelCalculo, objVarCobrAntBase As New ClassVariavelCalculo
Dim objFilFornIEIsento As New ClassVariavelCalculo, bExcecaoComSubstTrib As Boolean
Dim dVI As Double, iMotivo As Integer, dVCA As Double, dBCA As Double
Dim objVarOrigemProd As New ClassVariavelCalculo, iZerarAliquota As Integer, iZerarST As Integer
Dim objFilial As ClassFilialFornecedor, objVarFilFornObj As New ClassVariavelCalculo, objVarFCI As New ClassVariavelCalculo
Dim objVar51ValorOp As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo, objVarValorDif As New ClassVariavelCalculo
Dim d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, bDiferimentoParcial As Boolean

'NT 2015.003 (ICMS Interestadual)
Dim objVarInterestBCUFDest As New ClassVariavelCalculo, objVarInterestPercFCPUFDest As New ClassVariavelCalculo, objVarInterestAliqUFDest As New ClassVariavelCalculo, objVarInterestAliq As New ClassVariavelCalculo
Dim objVarInterestPercPartilha As New ClassVariavelCalculo, objVarInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarInterestVlrUFDest As New ClassVariavelCalculo, objVarInterestVlrUFRemet As New ClassVariavelCalculo
Dim dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double
Dim dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double

'NFe 4.00
Dim objVarICMSInterestBCFCPUFDest As New ClassVariavelCalculo, objVarICMSvBCFCPST As New ClassVariavelCalculo, objVarICMSpFCPST As New ClassVariavelCalculo
Dim objVarICMSvFCPST As New ClassVariavelCalculo, objVarICMSvBCFCPSTRet As New ClassVariavelCalculo, objVarICMSpFCPSTRet  As New ClassVariavelCalculo
Dim objVarICMSvFCPSTRet As New ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As New ClassVariavelCalculo, objVarICMSvBCFCP As New ClassVariavelCalculo
Dim objVarICMSpFCP As New ClassVariavelCalculo, objVarICMSvFCP As New ClassVariavelCalculo
Dim dICMSInterestBCFCPUFDest As Double, dICMSvBCFCPST As Double, dICMSpFCPST As Double
Dim dICMSvFCPST As Double, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double
Dim dICMSvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dICMSvBCFCP As Double
Dim dICMSpFCP As Double, dICMSvFCP As Double, dAliquotaPadraoFCP As Double

'Base Dupla
Dim objVarICMSSTBaseDupla As New ClassVariavelCalculo, objVarICMSInterestBaseDupla As New ClassVariavelCalculo
Dim iICMSSTBaseDupla As Integer, iICMSInterestBaseDupla As Integer
Dim objVarcBenef As New ClassVariavelCalculo, scBenef As String
Dim dValorPauta As Double, iUsaPauta As Integer

On Error GoTo Erro_AtualizarItem_SaidaForn
    
    iTipo = ICMS_TIPO_NAO_TRIBUTADO
    bDiferimentoParcial = False
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
    iModalidade = 0
    iModalidadeST = 0
    bExcecaoComSubstTrib = False
    
    iZerarAliquota = 0
    iZerarST = 0
    
    'NFe 4.00
    dICMSInterestBCFCPUFDest = 0
    dICMSvBCFCPST = 0
    dICMSpFCPST = 0
    dICMSvFCPST = 0
    dICMSvBCFCPSTRet = 0
    dICMSpFCPSTRet = 0
    dICMSvFCPSTRet = 0
    dICMSSTCobrAntAliquota = 0
    dICMSvBCFCP = 0
    dICMSpFCP = 0
    dICMSvFCP = 0
    scBenef = ""
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, objVarICMSInterestBaseDupla, objVarcBenef)
    
    If lErro <> SUCESSO Then Error 22728
    
    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 22728
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarOrigemProd.vValor = objProduto.iOrigemMercadoria
            objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    If objVarFCI.iStatus <> VAR_PREENCH_MANUAL Then

        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then Error 27319
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            objVarFCI.vValor = objProduto.sFCI
            objVarFCI.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
    End If
    
    objVarRegTrib.sIdentificador = "REGIME_TRIBUTARIO"
    'Se não é do Simples
    If mvariSimplesNacional = DESMARCADO Then
        objVarRegTrib.vValor = REGIME_TRIBUTARIO_NORMAL
    Else
        objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES
    End If
    objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 27320
    
    If mvariZerarTributo <> 0 And mvariZerarTributoST <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
            
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 22732
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 22732
    
    'NFe 4.00
    dICMSpFCP = objVarICMSpFCP.vValor
    dICMSpFCPST = objVarICMSpFCPST.vValor
    dICMSpFCPSTRet = objVarICMSpFCPSTRet.vValor
    'NFe 4.00
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 22739
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)

        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        dVI = objVarVI.vValor
        dVCA = objVarCobrAntValor.vValor
        dBCA = objVarCobrAntBase.vValor
        dICMSSTCobrAntAliquota = objVarICMSSTCobrAntAliquota.vValor
        If objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_MANUAL Then
            dVCA = Arredonda_Moeda(dBCA * dICMSSTCobrAntAliquota)
        End If
        iMotivo = objVarMotivo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, iTipo = ICMS_TIPO_COBR_SUBST_ANT)
        'Manter campos próprios desses tipos de tributação
        If iTipo = ICMS_TIPO_COBR_SUBST_ANT Then
            objVarCobrAntBase.vValor = dBCA
            objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
            objVarCobrAntValor.vValor = dVCA
        
            'NFe 4.00
            If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvBCFCPSTRet = dBCA
                objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvBCFCPSTRet = objVarICMSvBCFCPST.vValor
            End If
            
            If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvFCPSTRet = Arredonda_Moeda(dICMSvBCFCPSTRet * dICMSpFCPSTRet)
                objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvFCPSTRet = objVarICMSvFCPSTRet.vValor
            End If
            'NFe 4.00
       Else
            objVarVI.vValor = dVI
            objVarMotivo.vValor = iMotivo
        End If
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    iModalidade = objVarModalidade.vValor
    iModalidadeST = objVarModalidadeST.vValor
        
    'obter UF Forn
    lErro = objDocItem.objDoc.ObterVar("FORN_UF", objVarUFForn)
    If lErro <> SUCESSO Then Error 22740
    
    sUFForn = objVarUFForn.vValor
    If sUFForn = "" Then sUFForn = mvarsUFFilialEmpresa
    
    sUFDestino = mvarsUFFilialEmpresa
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_INSC_EST", objFilFornInscEst)
    If lErro <> SUCESSO Then Error 22741
    
    lErro = objDocItem.objDoc.ObterVar("FILIAL_FORN_IEISENTO", objFilFornIEIsento)
    If lErro <> SUCESSO Then Error 22741
     
    iICMSInterestBaseDupla = objVarICMSInterestBaseDupla.vValor
    If objVarICMSInterestBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_INTER", sUFForn, mvardtDataRef, iICMSInterestBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    iICMSSTBaseDupla = objVarICMSSTBaseDupla.vValor
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
        lErro = CF("TribBaseDupla_Verifica", "ICMS_ST", sUFForn, mvardtDataRef, iICMSSTBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    'se filial forn inscrita ICMS e interestadual
    If objFilFornIEIsento.vValor = DESMARCADO And sUFForn <> mvarsUFFilialEmpresa And sUFForn <> "EX" Then
        
        'adotar a aliquota interestadual para operacao basica
        lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
        If lErro <> SUCESSO Then Error 22742
        
        'ver se na UF de destino tem excecao que possa incluir subst tribut
        lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, sUFForn, mvarsUFFilialEmpresa)
        If lErro <> SUCESSO Then Error 22743
            
        If Not (objExcecao Is Nothing) Then
            'se a excecao encontrada envolver subst tributaria ou reducao da base de calculo
            If objExcecao.dPercMargemLucro <> 0 Or objExcecao.dPercRedBaseCalculo <> 0 Then
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
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
            
                lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, False, iUsaPauta, dValorPauta, dICMSpFCP, objVarICMSpFCP, iICMSSTBaseDupla, objVarICMSSTBaseDupla, iMotivo, objVarMotivo, scBenef, objVarcBenef)
                If lErro <> SUCESSO Then Error 22745
                    
            Else 'se NAO achou uma excecao
            
                lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
                If lErro <> SUCESSO Then Error 22746
            
            End If
            
        Else  'exportacao
        
            lErro = AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, sUFForn, dICMSpFCP, objVarICMSpFCP)
            If lErro <> SUCESSO Then Error 22747
            
            'se ICMS_TIPO nao tiver sido preenchido manualmente...
            If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            
                'Coloca como não tributado
                iTipo = ICMS_TIPO_NAO_TRIBUTADO
                objVarTipo.vValor = iTipo
                objTipoTribICMS.iTipo = iTipo
                lErro = TipoICMS_Obter(objTipoTribICMS)
                If lErro <> SUCESSO Then Error 22747
        
            End If
        
        End If
        
    End If

    If objTipoTribICMS.iPermiteST = MARCADO And objFilFornIEIsento.vValor = DESMARCADO And (dMargSubst <> 0) Then bExcecaoComSubstTrib = True

    If objTipoTribICMS.iPermiteReducaoBase <> 1 Then
        dPercRedBase = 0
        dPercRedBaseSubst = 0
    End If
    
    If objTipoTribICMS.iPermiteMargLucro <> 1 Then
        dMargSubst = 0
    End If
    
    'Calcula a permissão de crédito de ICMS
    'VAI ESTAR FALTANDO ZERAR OUTROS VALORES COMO A BASE E ALÍQUOTA NORMAL QUANDO NÃO PERMITE, ETC
    If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        objVarAliCredSimples.vValor = 0
        objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        objVarVlrCredSimples.vValor = 0
        objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        
        If objTipoTribICMS.iPermiteMotDesoneracao = DESMARCADO Then
            objVarMotivo.vValor = 0
            objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVI.vValor = 0
            objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermitePercBaseOperProp = DESMARCADO Then
            objVarPBOP.vValor = 0
            objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTUFDest = DESMARCADO Then
            objVarUFD.vValor = ""
            objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoTribICMS.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    Else
    
        'Se não é manual ou não está preenchido calcula
        If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Or objVarTipoICMSSimples.vValor = 0 Then
        
            If (Not (objExcecao Is Nothing)) Then
                objVarTipoICMSSimples.vValor = objExcecao.iTipoSimples
                If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            Else
                objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            End If
            If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = 1
            objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
            Select Case objVarTipoICMSSimples.vValor
                Case 1, 2, 4, 5
                    lErro = objDoc.ObterVar("FILIAL_FORN_OBJ", objVarFilFornObj)
                    If lErro <> SUCESSO Then Error 27916
                        
                    Set objFilial = objVarFilFornObj.vValor
                    
                    If objFilial.iIEIsento = MARCADO Or objFilial.iIENaoContrib = MARCADO Or objFilial.iRegimeTributario = REGIME_TRIBUTARIO_SIMPLES Or objVarTipoICMSSimples.vValor = 2 Or objVarTipoICMSSimples.vValor = 5 Then
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 5 'Sem crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 2 'Sem crédito - Sem ST
                        End If
                    Else
                        If bExcecaoComSubstTrib Then
                            objVarTipoICMSSimples.vValor = 4 'Com crédito - Com ST
                        Else
                            objVarTipoICMSSimples.vValor = 1 'Com crédito - Sem ST
                        End If
                    End If
                Case 3, 6
                    If bExcecaoComSubstTrib Then
                        objVarTipoICMSSimples.vValor = 6 'Com crédito - Com ST
                    Else
                        objVarTipoICMSSimples.vValor = 3 'Com crédito - Sem ST
                    End If
            End Select
        End If
          
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        Set objVarItemTipoObj.vValor = objTipoICMSSimples
        lErro = objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
        If lErro <> SUCESSO Then Error 27320

        Set objTipoICMSSimples = objVarItemTipoObj.vValor
       
        If objTipoICMSSimples.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoICMSSimples.iPermiteCredito = DESMARCADO Then
            objVarAliCredSimples.vValor = 0
            objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVlrCredSimples.vValor = 0
            objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            
            If objVarAliCredSimples.iStatus <> VAR_PREENCH_MANUAL Then
            
                objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            
                lErro = objDoc.ObterVar("DATA_EMISSAO", objVarEmissao)
                If lErro <> SUCESSO Then Error 27320
            
                objDASAliquotas.iFilialEmpresa = mvariFilialEmpresa
                objDASAliquotas.iAno = Year(DateAdd("m", -1, IIf(objVarEmissao.vValor = DATA_NULA, gdtDataAtual, objVarEmissao.vValor)))
                objDASAliquotas.iMes = Month(DateAdd("m", -1, IIf(objVarEmissao.vValor = DATA_NULA, gdtDataAtual, objVarEmissao.vValor)))
                
                lErro = CF("DASAliquotas_Le", objDASAliquotas)
                If lErro <> SUCESSO And lErro <> 200911 Then Error 27320
                
                lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
                If lErro <> SUCESSO Then Error 27320
                
                Set objProduto = objVarProd.vValor
                
                If objProduto.iNatureza = NATUREZA_PROD_SERVICO Then
                    objVarAliCredSimples.vValor = objDASAliquotas.dAliquotaICMSServ
                Else
                    objVarAliCredSimples.vValor = objDASAliquotas.dAliquotaICMS
                End If
                                
            End If

            If objVarVlrCredSimples.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                
                objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
                
                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdVlr)
                If lErro <> SUCESSO Then Error 27320
                
                objVarVlrCredSimples.vValor = Arredonda_Moeda(objVarProdVlr.vValor * objVarAliCredSimples.vValor)
            
            End If
            
        End If
        If objTipoICMSSimples.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
        If objTipoICMSSimples.iPermiteST = DESMARCADO Then iZerarST = MARCADO
    
    End If
    'Calcula a permissão de crédito de ICMS
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27324
            
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If

    'se ICMS_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = Arredonda_Moeda(dValor)
    
    Else
        dValor = objVarValor.vValor
    End If

    If objVarBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
        If bExcecaoComSubstTrib Then 'dMargSubst <> 0 Then
            'se ainda nao incluiu o IPI na base de calculo em funcao do tipotrib
            If objTipoTrib.iICMSBaseComIPI = 0 Then
            
                'obter valor do IPI do item de objdocitem
                lErro = objDocItem.ObterVar("IPI_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27325
                
                dBaseSubst = (dBaseCalculo + objVar.vValor) '* (1 + dMargSubst)
                                
            Else
                dBaseSubst = dBaseCalculo '* (1 + dMargSubst)
            End If
            
            If iICMSSTBaseDupla = MARCADO Then
            
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
    
                dBaseSubst = (dBaseSubst - dValor) / (1 - dAliqSubst)
            End If
            
            dBaseSubst = dBaseSubst * (1 + dMargSubst)
            
            dBaseSubst = CDbl(Format(dBaseSubst, "0.00"))
            dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
            'dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.0000"))
            dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida)
        
        Else
            dBaseSubst = 0
            dBaseSubstReduzida = dBaseSubst
        End If
    Else
        dBaseSubst = objVarBaseSubst.vValor
        dMargSubst = objVarMargSubst.vValor
        'dBaseSubst = dBaseSubst * (1 + dMargSubst)
        dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        dBaseSubstReduzida = dBaseSubst * (1 - dPercRedBaseSubst)
    End If
    
    If iZerarST Then
        dBaseSubst = 0
        objVarBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dMargSubst = 0
        objVarMargSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dPercRedBaseSubst = 0
        objVarPercRedBaseSubst.iStatus = VAR_PREENCH_AUTOMATICO
        dBaseSubstReduzida = 0
        
        objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        dICMSpFCPST = 0
    End If
    
    'se é para zerar a parte propria
    If iZerarAliquota = MARCADO Or iTipo = ICMS_TIPO_COM_DIFERIMENTO Then
        If iZerarAliquota = MARCADO Then
            objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        
        'http://portalnfe.fazenda.mg.gov.br/respostas_III.html
        'Na versão 3.10 do leiaute da NF-e foram criados campos para informação do diferimento parcial.
        'Esses campos também deverão ser preenchidos no caso do diferimento total?
        'Os campos modBC (Modalidade de determinação da base de cálculo do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação), pDif (Percentual do diferimento), vICMSDif (Valor do ICMS diferido) e vICMS (Valor do ICMS)
        'só devem ser preenchidos no caso do diferimento ser parcial.
        'Esses campos NÃO serão de preenchimento obrigatório para os casos do diferimento total.
        'Alertamos que os campos modBC (Modalidade de determinação da BC do ICMS), pRedBC (Percentual da Redução de BC), vBC (Valor da BC do ICMS),
        'pICMS (Alíquota do imposto), vICMSOp (Valor do ICMS da Operação) devem ser preenchidos com o valor como se não tivesse o diferimento.

        If objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO And mvariSimplesNacional = 0 And mvariZerarTributo = 0 And iZerarAliquota <> MARCADO Then
        
            If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
                dPercDifer = 1 '??? deveria vir da excecao
                objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dPercDifer = objVarPercDifer.vValor
            End If
            
            If dPercDifer <> 1 And dPercDifer <> 0 Then bDiferimentoParcial = True
        
            If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
                d51ValorOp = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
                objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                d51ValorOp = objVar51ValorOp.vValor
            End If
            
            If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
                dValorDif = Arredonda_Moeda(d51ValorOp * dPercDifer)
                objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dValorDif = objVarValorDif.vValor
            End If
        
        Else
        
            d51ValorOp = 0
            objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
            dPercDifer = 0
            objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
            dValorDif = 0
            objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
        
        If objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Then
            
            If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then dBaseCalculo = 0
            If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then dPercRedBase = 0
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase) * dAliquota)
        
            If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then dICMSpFCP = 0
        
        Else
        
            If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(d51ValorOp - dValorDif)
        
        End If
        
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
        
        'NFe 4.00
        If objVarICMSpFCPST.iStatus <> VAR_PREENCH_MANUAL Then
            If Not (objExcecao Is Nothing) Then
                dICMSpFCPST = objExcecao.dICMSPercFCP
            End If
        Else
            dICMSpFCPST = objVarICMSpFCPST.vValor
        End If
        'NFe 4.00
        
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
    
    'NFe 4.00
    'Base de cálculo dos FCP iguais as bases correspondentes do ICMS
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCP = dBaseCalculo
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCP = objVarICMSvBCFCP.vValor
    End If
    
    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCPST = dBaseSubst
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCPST = objVarICMSvBCFCPST.vValor
    End If
           
    'Calculando o valor com a base e alíquota
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCP = Arredonda_Moeda(dICMSvBCFCP * dICMSpFCP)
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCP = objVarICMSvFCP.vValor
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCPST = Arredonda_Moeda(dICMSvBCFCPST * dICMSpFCPST)
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCPST = objVarICMSvFCPST.vValor
    End If
    'NFe 4.00

    'Se credita e não é com diferimento total
    If dValorSubst = 0 And objTipoTrib.iICMSCredita <> 0 And (objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Or bDiferimentoParcial) And (mvariSimplesNacional = 0 Or mvardtDataRef < DATA_INICIO_SIMPLES_NACIONAL) Then
    
        dValorCredito = dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito, iModalidade, iModalidadeST, _
                                objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarRegTrib, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, d51ValorOp, dPercDifer, dValorDif, iMotivo)
    If lErro <> SUCESSO Then Error 22750
    
    lErro = AtualizarItem_GuardarValores2(objDocItem, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                dInterestBCUFDest, dInterestPercFCPUFDest, dInterestAliqUFDest, dInterestAliq, dInterestPercPartilha, dInterestVlrFCPUFDest, dInterestVlrUFDest, dInterestVlrUFRemet, objVarICMSInterestBCFCPUFDest, dICMSInterestBCFCPUFDest)
    If lErro <> SUCESSO Then Error 22750
    
    lErro = AtualizarItem_GuardarValores3(objDocItem, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, dICMSvBCFCP, dICMSpFCP, dICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, dICMSvBCFCPST, dICMSpFCPST, dICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, dICMSvBCFCPSTRet, dICMSpFCPSTRet, dICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, dICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, iICMSSTBaseDupla, objVarICMSInterestBaseDupla, iICMSInterestBaseDupla, objVarcBenef, scBenef)
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

Private Function AtualizarItem_TratarNaoExcecaoSaidaForn(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, sUFForn As String, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo) As Long
Dim lErro As Long
Dim dAliquotaFCP As Double

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
        
'    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Or objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            lErro = ObterAliquotaBasicaSaidaForn(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquota, sUFForn, objDocItem, dAliquotaFCP)
            If lErro <> SUCESSO Then Error 22752
        Else
            dAliquota = 0
            dAliquotaFCP = 0
        End If
'    Else
'
'        dAliquota = objVarAliquota.vValor
        
    End If
    
    dICMSpFCP = dAliquotaFCP
    
    If objVarAliquota.iStatus = VAR_PREENCH_MANUAL Then dAliquota = objVarAliquota.vValor
    If objVarICMSpFCP.iStatus = VAR_PREENCH_MANUAL Then dICMSpFCP = objVarICMSpFCP.vValor
    
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
       
        mvardAliqInternaAtual = objEstado.dICMSAliquotaInterna
        mvardAliqInternaFCPAtual = objEstado.dICMSPercFCP
        mvardAliqImportacaoAtual = objEstado.dICMSAliquotaImportacao
        
        mvardAliqInternaAnt = objEstado.dICMSAliquotaInternaAnt
        mvardAliqInternaFCPAnt = objEstado.dICMSPercFCPAnt
        mvardAliqImportacaoAnt = objEstado.dICMSAliquotaImportacaoAnt
        
        mvardAliqInterna = mvardAliqInternaAtual
        mvardAliqInternaFCP = mvardAliqInternaFCPAtual
        mvardAliqImportacao = mvardAliqImportacaoAtual
        
        mvardtDataIniAliqInternaAtual = objEstado.dtDataIniAliqInternaAtual
        mvardtDataIniAliqImportacaoAtual = objEstado.dtDataIniAliqImportacaoAtual
        mvardtDataIniAliqFCPAtual = objEstado.dtDataIniAliqFCPAtual
        
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

Private Function AtualizarItem_ZerarValores(objDocItem As ClassTributoDocItem, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarBaseSubst As ClassVariavelCalculo, objVarPercRedBaseSubst As ClassVariavelCalculo, objVarMargSubst As ClassVariavelCalculo, objVarAliqSubst As ClassVariavelCalculo, objVarValorSubst As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, objVarModalidade As ClassVariavelCalculo, objVarModalidadeST As ClassVariavelCalculo, objVarVlrCredSimples As ClassVariavelCalculo, objVarAliCredSimples As ClassVariavelCalculo, objVarTipoICMSSimples As ClassVariavelCalculo, _
    objVarCobrAntValor As ClassVariavelCalculo, objVarCobrAntBase As ClassVariavelCalculo, objVarPBOP As ClassVariavelCalculo, objVarUFD As ClassVariavelCalculo, objVarBCSTRet As ClassVariavelCalculo, objVarICMSSTRet As ClassVariavelCalculo, objVarBCSTDest As ClassVariavelCalculo, objVarICMSSTDest As ClassVariavelCalculo, objVarVI As ClassVariavelCalculo, objVarMotivo As ClassVariavelCalculo, objVar51ValorOp As ClassVariavelCalculo, objVarPercDifer As ClassVariavelCalculo, objVarValorDif As ClassVariavelCalculo, _
    objVarICMSInterestBCFCPUFDest As ClassVariavelCalculo, objVarICMSvBCFCP As ClassVariavelCalculo, objVarICMSpFCP As ClassVariavelCalculo, objVarICMSvFCP As ClassVariavelCalculo, objVarICMSvBCFCPST As ClassVariavelCalculo, objVarICMSpFCPST As ClassVariavelCalculo, objVarICMSvFCPST As ClassVariavelCalculo, objVarICMSvBCFCPSTRet As ClassVariavelCalculo, objVarICMSpFCPSTRet As ClassVariavelCalculo, objVarICMSvFCPSTRet As ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As ClassVariavelCalculo, Optional ByVal bManterCobrAnt As Boolean = False) As Long

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
    
    objVarModalidade.vValor = ICMS_MODALIDADE_VALOR
    objVarModalidade.iStatus = VAR_PREENCH_AUTOMATICO

    objVarModalidadeST.vValor = ICMS_ST_MODALIDADE_MARGEM
    objVarModalidadeST.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarVlrCredSimples.vValor = 0
    objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
     
    objVarAliCredSimples.vValor = 0
    objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
    
    If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Then
        objVarTipoICMSSimples.vValor = 8
        objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
    End If
    
    objVarPBOP.vValor = 0
    objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarUFD.vValor = ""
    objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarBCSTRet.vValor = 0
    objVarBCSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarICMSSTRet.vValor = 0
    objVarICMSSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarBCSTDest.vValor = 0
    objVarBCSTDest.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarICMSSTDest.vValor = 0
    objVarICMSSTDest.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarVI.vValor = 0
    objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
   
    objVarMotivo.vValor = 0
    objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
    
    If Not bManterCobrAnt Then
        objVarCobrAntValor.vValor = 0
        objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
        
        objVarCobrAntBase.vValor = 0
        objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
    
        objVarICMSSTCobrAntAliquota.vValor = 0
        objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
    
    End If
         
    objVar51ValorOp.vValor = 0
    objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarPercDifer.vValor = 0
    objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
    
    objVarValorDif.vValor = 0
    objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
    
    'NFe 4.00
    objVarICMSInterestBCFCPUFDest.vValor = 0
    objVarICMSInterestBCFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSvBCFCP.vValor = 0
    objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSpFCP.vValor = 0
    objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSvFCP.vValor = 0
    objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSvBCFCPST.vValor = 0
    objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSpFCPST.vValor = 0
    objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO

    objVarICMSvFCPST.vValor = 0
    objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO

    If Not bManterCobrAnt Then
        objVarICMSvBCFCPSTRet.vValor = 0
        objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
        objVarICMSpFCPSTRet.vValor = 0
        objVarICMSpFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    
        objVarICMSvFCPSTRet.vValor = 0
        objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
    End If
    
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

Private Function ExcecaoForn_Pesquisar(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, sUFDestino As String, sUFOrigem As String) As Long
'uma excecao a nivel de fornecedor tem + prioridade que excecoes a nivel de produto apenas
'e uma excecao a nivel de fornecedor+produto tem a maior prioridade de todas.
'A colecao já está ordenada de forma que pararei na 1a excecao que encontrar
'retorna Nothing se nao achar nenhuma excecao adequada
Dim lErro As Long, objICMSExcecao As ClassICMSExcecao, iAchou As Integer
Dim objVarUFDestino As New ClassVariavelCalculo, objVarFilFornColCateg As New ClassVariavelCalculo
Dim objVarProdColCateg As New ClassVariavelCalculo, bValida As Boolean
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
    
        lErro = OrigemMercValida(objDocItem, objICMSExcecao.iGrupoOrigemMercadoria, bValida)
        If lErro <> SUCESSO Then Error 27725
        
        If bValida And objICMSExcecao.iTipoCliForn = 1 And (sUFOrigem = "" Or Trim(objICMSExcecao.sEstadoOrigem) = "" Or objICMSExcecao.sEstadoOrigem = sUFOrigem) Then
            
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
    
        If UCase(objFilialFornCategoria.sCategoria) = UCase(sCategoriaFornecedor) And UCase(objFilialFornCategoria.sItem) = UCase(sCategoriaFornecedorItem) Then
        
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
        lErro = AtualizarItem_Importacao(objDoc, objDocItem, iRecalculaAutomatica)
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

Private Function AtualizarItem_Importacao(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long
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
Dim objVarAliquotaDI As New ClassVariavelCalculo, objVarRateio As New ClassVariavelCalculo, objVarPercRedBaseDI As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, dAliquotaEfetiva As Double
Dim dValor As Double, dValorSubst As Double, dAliqSubst As Double, dMargSubst As Double, dBaseSubst As Double
Dim objTRBConfig As New ClassTRBConfig, sUFForn As String, dPercRedBaseSubst As Double
Dim objVarUFForn As New ClassVariavelCalculo, objFilFornInscEst As New ClassVariavelCalculo
Dim objVarProduto As New ClassVariavelCalculo, dBaseAux As Double, objVarPercRedBaseSubst As New ClassVariavelCalculo
Dim objVarModalidade As New ClassVariavelCalculo, objVarModalidadeST As New ClassVariavelCalculo
Dim iModalidade As Integer, iModalidadeST As Integer, objVarProdOutrosValores As ClassVariavelCalculo
Dim objVarDespImpValorRateado As New ClassVariavelCalculo
Dim dImpICMSBase As Double, dImpICMSValor As Double, dImpICMSCredito As Double

Dim objVarRegTrib As New ClassVariavelCalculo, dBaseInicial As Double
Dim objTipoICMSSimples As New ClassTipoTribICMSSimples, objVarEmissao As New ClassVariavelCalculo, objDASAliquotas As New ClassDASAliquotas
Dim objVarTipoICMSSimples As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo, objVarVlrCredSimples As New ClassVariavelCalculo, objVarAliCredSimples As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, objVarProdVlr As New ClassVariavelCalculo
Dim objVarPBOP As New ClassVariavelCalculo, objVarUFD As New ClassVariavelCalculo, objVarBCSTRet As New ClassVariavelCalculo
Dim objVarICMSSTRet As New ClassVariavelCalculo, objVarBCSTDest As New ClassVariavelCalculo, objVarICMSSTDest As New ClassVariavelCalculo
Dim objVarVI As New ClassVariavelCalculo, objVarMotivo As New ClassVariavelCalculo, objVarCobrAntValor As New ClassVariavelCalculo, objVarCobrAntBase As New ClassVariavelCalculo
Dim objVarOrigemProd As New ClassVariavelCalculo
Dim dVI As Double, iMotivo As Integer, dVCA As Double, dBCA As Double, iZerarAliquota As Integer, objVarFCI As New ClassVariavelCalculo
Dim objVar51ValorOp As New ClassVariavelCalculo, objVarPercDifer As New ClassVariavelCalculo, objVarValorDif As New ClassVariavelCalculo
Dim d51ValorOp As Double, dPercDifer As Double, dValorDif As Double, bDiferimentoParcial As Boolean, objVarIIValor As New ClassVariavelCalculo

'NT 2015.003 (ICMS Interestadual)
Dim objVarInterestBCUFDest As New ClassVariavelCalculo, objVarInterestPercFCPUFDest As New ClassVariavelCalculo, objVarInterestAliqUFDest As New ClassVariavelCalculo, objVarInterestAliq As New ClassVariavelCalculo
Dim objVarInterestPercPartilha As New ClassVariavelCalculo, objVarInterestVlrFCPUFDest As New ClassVariavelCalculo, objVarInterestVlrUFDest As New ClassVariavelCalculo, objVarInterestVlrUFRemet As New ClassVariavelCalculo
Dim dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double
Dim dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double

'NFe 4.00
Dim objVarICMSInterestBCFCPUFDest As New ClassVariavelCalculo, objVarICMSvBCFCPST As New ClassVariavelCalculo, objVarICMSpFCPST As New ClassVariavelCalculo
Dim objVarICMSvFCPST As New ClassVariavelCalculo, objVarICMSvBCFCPSTRet As New ClassVariavelCalculo, objVarICMSpFCPSTRet  As New ClassVariavelCalculo
Dim objVarICMSvFCPSTRet As New ClassVariavelCalculo, objVarICMSSTCobrAntAliquota As New ClassVariavelCalculo, objVarICMSvBCFCP As New ClassVariavelCalculo
Dim objVarICMSpFCP As New ClassVariavelCalculo, objVarICMSvFCP As New ClassVariavelCalculo
Dim dICMSInterestBCFCPUFDest As Double, dICMSvBCFCPST As Double, dICMSpFCPST As Double
Dim dICMSvFCPST As Double, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double
Dim dICMSvFCPSTRet As Double, dICMSSTCobrAntAliquota As Double, dICMSvBCFCP As Double
Dim dICMSpFCP As Double, dICMSvFCP As Double, dAliquotaPadraoFCP As Double

'Base Dupla
Dim objVarICMSSTBaseDupla As New ClassVariavelCalculo, objVarICMSInterestBaseDupla As New ClassVariavelCalculo
Dim iICMSSTBaseDupla As Integer, iICMSInterestBaseDupla As Integer
Dim objVarcBenef As New ClassVariavelCalculo, scBenef As String

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
    iModalidade = 0
    iModalidadeST = 0
    dImpICMSBase = 0
    dImpICMSValor = 0
    dImpICMSCredito = 0
    
    'NFe 4.00
    dICMSInterestBCFCPUFDest = 0
    dICMSvBCFCPST = 0
    dICMSpFCPST = 0
    dICMSvFCPST = 0
    dICMSvBCFCPSTRet = 0
    dICMSpFCPSTRet = 0
    dICMSvFCPSTRet = 0
    dICMSSTCobrAntAliquota = 0
    dICMSvBCFCP = 0
    dICMSpFCP = 0
    dICMSvFCP = 0
    scBenef = ""
    
    'especifico de importacao
    lErro = objDocItem.ObterVar("DESP_IMP_VALOR_RATEADO", objVarDespImpValorRateado)
    If lErro <> SUCESSO Then Error 27949
        
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, iRecalculaAutomatica, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, objVarICMSInterestBaseDupla, objVarcBenef)
    If lErro <> SUCESSO Then Error 27974
    
    objVarRegTrib.sIdentificador = "REGIME_TRIBUTARIO"
    'Se não é do Simples
    If mvariSimplesNacional = DESMARCADO Or gobjCRFAT.iNFImportacaoTribFlag02 <> DESMARCADO Then
        objVarRegTrib.vValor = REGIME_TRIBUTARIO_NORMAL
        mvariZerarTributo = 0
    Else
        objVarRegTrib.vValor = REGIME_TRIBUTARIO_SIMPLES
    End If
    objVarRegTrib.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDocItem.GuardarVar(objVarRegTrib)
    If lErro <> SUCESSO Then Error 27974
    
    'gobjCRFAT
    'NFImportacaoTribFlag01 - Permite a configuração do imposto mesmo no simples
    'NFImportacaoTribFlag02 - ***Descrever para que serve***
    'NFImportacaoTribFlag03 - ***Descrever para que serve***
    'NFImportacaoTribFlag04 - ***Descrever para que serve***
    'NFImportacaoTribFlag05 - ***Descrever para que serve***
    'NFImportacaoTribFlag06 - ***Descrever para que serve***
    'NFImportacaoTribFlag07 - ***Descrever para que serve***
    'NFImportacaoTribFlag08 - ***Descrever para que serve***
    'NFImportacaoTribFlag09 - ***Descrever para que serve***
    'NFImportacaoTribFlag10 - ***Descrever para que serve***
    
    If gobjCRFAT.iNFImportacaoTribFlag01 = DESMARCADO Then
        If mvariZerarTributo <> 0 Then
            
            Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)
            Call AtualizarItem_GuardarValoresImp(objDocItem, 0, 0, 0)
            
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27978
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27985
    
    'NFe 4.00
    dICMSpFCP = objVarICMSpFCP.vValor
    dICMSpFCPST = objVarICMSpFCPST.vValor
    dICMSpFCPSTRet = objVarICMSpFCPSTRet.vValor
    'NFe 4.00
    
    'se nao incide ICMS sobre a operacao
    If objTipoTrib.iICMSIncide = TRIBUTO_NAO_INCIDE Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota)

        Call AtualizarItem_GuardarValoresImp(objDocItem, 0, 0, 0)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula ICMS em funcao do ICMS_TIPO setado manualmente
    'como ICMS_TIPO_NAO_TRIBUTADO ou ICMS_TIPO_ISENTO
    If objVarTipo.iStatus = VAR_PREENCH_MANUAL And TipoICMSZerar(objVarTipo.vValor) Then
        iTipo = objVarTipo.vValor
        dVI = objVarVI.vValor
        dVCA = objVarCobrAntValor.vValor
        dBCA = objVarCobrAntBase.vValor
        dICMSSTCobrAntAliquota = objVarICMSSTCobrAntAliquota.vValor
        If objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_MANUAL Then
            dVCA = Arredonda_Moeda(dBCA * dICMSSTCobrAntAliquota)
        End If
        iMotivo = objVarMotivo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarPercRedBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVar51ValorOp, objVarPercDifer, objVarValorDif, _
                                                objVarICMSInterestBCFCPUFDest, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, iTipo = ICMS_TIPO_COBR_SUBST_ANT)
        Call AtualizarItem_GuardarValoresImp(objDocItem, 0, 0, 0)
        'Manter campos próprios desses tipos de tributação
        If iTipo = ICMS_TIPO_COBR_SUBST_ANT Then
            objVarCobrAntBase.vValor = dBCA
            objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
            objVarCobrAntValor.vValor = dVCA
        
            'NFe 4.00
            If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvBCFCPSTRet = dBCA
                objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvBCFCPSTRet = objVarICMSvBCFCPST.vValor
            End If
            
            If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
                dICMSvFCPSTRet = Arredonda_Moeda(dICMSvBCFCPSTRet * dICMSpFCPSTRet)
                objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
            Else
                dICMSvFCPSTRet = objVarICMSvFCPSTRet.vValor
            End If
            'NFe 4.00
        Else
            objVarVI.vValor = dVI
            objVarMotivo.vValor = iMotivo
        End If
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    iModalidade = objVarModalidade.vValor
    iModalidadeST = objVarModalidadeST.vValor
                
    'tentar achar excecao
    lErro = ExcecaoForn_Pesquisar(objDocItem, objExcecao, mvarsUFFilialEmpresa, "")
    If lErro <> SUCESSO Then Error 27990
    
'    'se achou excecao
'    If (Not (objExcecao Is Nothing)) Then
'
'        lErro = AtualizarItem_TratarExcecao(objDocItem, objExcecao, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarMargSubst, dMargSubst, objVarPercRedBaseSubst, dPercRedBaseSubst, true)
'        If lErro <> SUCESSO Then Error 27991
'
'    Else 'se NAO achou uma excecao
    
        lErro = AtualizarItem_TratarNaoExcecaoEntForn(objDocItem, objTipoTrib, objTipoTribICMS, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, "EX", dICMSpFCP, objVarICMSpFCP)
        If lErro <> SUCESSO Then Error 27992
    
'    End If
        
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
    Else
        dAliquota = objVarAliquota.vValor
    End If
    
    'Calcula a permissão de crédito de ICMS
    'VAI ESTAR FALTANDO ZERAR OUTROS VALORES COMO A BASE E ALÍQUOTA NORMAL QUANDO NÃO PERMITE, ETC
    If objVarRegTrib.vValor <> REGIME_TRIBUTARIO_SIMPLES Then
    
        objVarAliCredSimples.vValor = 0
        objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        objVarVlrCredSimples.vValor = 0
        objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        
        If objTipoTribICMS.iPermiteMotDesoneracao = DESMARCADO Then
            objVarMotivo.vValor = 0
            objVarMotivo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVI.vValor = 0
            objVarVI.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermitePercBaseOperProp = DESMARCADO Then
            objVarPBOP.vValor = 0
            objVarPBOP.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteSTUFDest = DESMARCADO Then
            objVarUFD.vValor = ""
            objVarUFD.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoTribICMS.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
    Else
        
        'Se não é manual ou não está preenchido calcula
        If objVarTipoICMSSimples.iStatus <> VAR_PREENCH_MANUAL Or objVarTipoICMSSimples.vValor = 0 Then
        
            objVarTipoICMSSimples.vValor = objTipoTrib.iICMSSimplesTipo
            If objVarTipoICMSSimples.vValor < 1 Then objVarTipoICMSSimples.vValor = 1
            objVarTipoICMSSimples.iStatus = VAR_PREENCH_AUTOMATICO
            Select Case objVarTipoICMSSimples.vValor
                Case 1, 2
                    objVarTipoICMSSimples.vValor = 2 'Sem crédito
            End Select
        End If
          
        objTipoICMSSimples.iTipo = objVarTipoICMSSimples.vValor
        Set objVarItemTipoObj.vValor = objTipoICMSSimples
        lErro = objDoc.ObterVar("ICMS_SIMPLES_TIPO_OBJ", objVarItemTipoObj)
        If lErro <> SUCESSO Then Error 27320

        Set objTipoICMSSimples = objVarItemTipoObj.vValor
       
        If objTipoICMSSimples.iPermiteSTRetAnt = DESMARCADO Then
            objVarCobrAntValor.vValor = 0
            objVarCobrAntValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarCobrAntBase.vValor = 0
            objVarCobrAntBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarICMSSTCobrAntAliquota.vValor = 0
            objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        
        If objTipoICMSSimples.iPermiteCredito = DESMARCADO Then
            objVarAliCredSimples.vValor = 0
            objVarAliCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
            objVarVlrCredSimples.vValor = 0
            objVarVlrCredSimples.iStatus = VAR_PREENCH_AUTOMATICO
        End If
        If objTipoICMSSimples.iPermiteAliquota = DESMARCADO Then iZerarAliquota = MARCADO
    
    End If
    'Calcula a permissão de crédito de ICMS
    
    'forçar perc red base com o que está na DI
    'se ICMS_PERC nao tiver sido preenchido manualmente
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        'pegar da DI
        lErro = objDocItem.ObterVar("ICMS_PERCREDBASE_ADICAO_DI", objVarPercRedBaseDI)
        If lErro <> SUCESSO Then Error 27994
        
        dPercRedBase = objVarPercRedBaseDI.vValor
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
    
    If dPercRedBase <> 0 Then
        
        If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
            dAliquotaEfetiva = Round(dAliquota * (1 - dPercRedBase), 4)
        Else
            dAliquotaEfetiva = dAliquota
        End If
        
        'se icms normal
        If objTipoTribICMS.iTipo = ICMS_TIPO_NORMAL And objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        
            iTipo = ICMS_TIPO_RED_BASE
            objVarTipo.vValor = iTipo
            
            objTipoTribICMS.iTipo = iTipo
            lErro = TipoICMS_Obter(objTipoTribICMS)
            If lErro <> SUCESSO Then Error 22700
        
        End If
    Else
        dAliquotaEfetiva = dAliquota
    End If
    
    'se ICMS_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If Not TipoICMSZerar(objTipoTribICMS.iTipo) Then
            
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ICMS_II", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27994
            
            dBaseInicial = dBaseCalculo + objVarDespImpValorRateado.vValor
            'para atender a International (GoldenFive) em 30/08/2015
            If gobjCRFAT.iNFImportacaoTribFlag11 = MARCADO And objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO Then
            
                dBaseCalculo = dBaseInicial
                
            Else
            
                If gobjCRFAT.iNFImportacaoTribFlag12 Then
                
                    lErro = objDocItem.ObterVar("II_VALOR", objVarIIValor)
                    If lErro <> SUCESSO Then Error 27994
                    
                    dBaseInicial = Arredonda_Moeda(dBaseInicial + objVarIIValor.vValor, 2)
                        
                End If

                'NA IMPORTACAO O PROPRIO ICMS ENTRA NA BASE
                If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
                    
                    If dAliquota <> 0 Then
                        If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                            dBaseCalculo = dBaseInicial / (1 - dAliquotaEfetiva)
                        Else
                            dBaseCalculo = ((dBaseInicial * (1 - dPercRedBase)) / (1 - dAliquota)) / (1 - dPercRedBase)
                        End If
                    Else
                        If TipoICMSZerarBase(objTipoTribICMS.iTipo) Then
                            dBaseCalculo = 0
                        Else
                            dBaseCalculo = dBaseInicial
                        End If
                    End If
                Else
                    If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                        dBaseCalculo = dBaseInicial + objVarValor.vValor
                    Else
                        If dAliquota <> 0 Then
                            dBaseCalculo = (objVarValor.vValor / dAliquota) / (1 - dPercRedBase)
                        Else
                            dBaseCalculo = dBaseInicial + objVarValor.vValor
                        End If
                    End If
                End If
                    
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
        
        If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
            dValor = Arredonda_Moeda(dBaseCalculo * (1 - dPercRedBase), 2) * dAliquota '(dBaseCalculo * dAliquotaEfetiva)
        Else
            dValor = (dBaseCalculo * (1 - dPercRedBase) * dAliquota)
        End If
        dValor = Arredonda_Moeda(dValor, 2)
    
    Else
        
        dValor = objVarValor.vValor
    
    End If
    
    'se é para zerar a parte propria
    If iZerarAliquota = MARCADO Then
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dValorCredito = 0
        dICMSpFCP = 0
    End If

    If objVarValorSubst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarValorSubst.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dValorSubst = objVarValorSubst.vValor
    End If

    If mvariSimplesNacional = MARCADO Then
    
        dValorCredito = 0
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        
    Else
    
        If objTipoTrib.iICMSCredita <> 0 And objTipoTribICMS.iTipo <> ICMS_TIPO_COM_DIFERIMENTO Then
        
            dValorCredito = dValorCredito + dValor
            
        End If
    
    End If
    
    If dBaseCalculo <> 0 Then
        
        If dAliquotaEfetiva <> 0 Then
        
            'para atender a International (GoldenFive) em 30/08/2015
            If gobjCRFAT.iNFImportacaoTribFlag11 = MARCADO And objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO Then
            
                dImpICMSBase = Arredonda_Moeda(objVarDespImpValorRateado.vValor)
                dImpICMSValor = Arredonda_Moeda(dImpICMSBase * dAliquotaEfetiva)
                
            Else
            
                If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                    dImpICMSBase = Arredonda_Moeda(objVarDespImpValorRateado.vValor / (1 - dAliquotaEfetiva))
                    dImpICMSValor = Arredonda_Moeda(dImpICMSBase * dAliquotaEfetiva)
                Else
                    dImpICMSBase = Arredonda_Moeda(objVarDespImpValorRateado.vValor * (1 - dPercRedBase) / (1 - dAliquota))
                    dImpICMSValor = Arredonda_Moeda(dImpICMSBase * dAliquota)
                End If
                
            End If
            If dValorCredito <> 0 Then dImpICMSCredito = dImpICMSValor
        
        End If
    End If

    If objTipoTribICMS.iTipo = ICMS_TIPO_COM_DIFERIMENTO And mvariSimplesNacional = 0 And mvariZerarTributo = 0 And iZerarAliquota <> MARCADO Then
    
        If objVarPercDifer.iStatus <> VAR_PREENCH_MANUAL Then
            dPercDifer = 1 '??? deveria vir da excecao
            objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            dPercDifer = objVarPercDifer.vValor
        End If
        
        If dPercDifer <> 1 And dPercDifer <> 0 Then bDiferimentoParcial = True
    
        If objVar51ValorOp.iStatus <> VAR_PREENCH_MANUAL Then
            d51ValorOp = dValor
            objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            d51ValorOp = objVar51ValorOp.vValor
        End If
        
        If objVarValorDif.iStatus <> VAR_PREENCH_MANUAL Then
            dValorDif = Arredonda_Moeda(d51ValorOp * dPercDifer)
            objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
        Else
            dValorDif = objVarValorDif.vValor
        End If
    
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then dValor = Arredonda_Moeda(d51ValorOp - dValorDif)
    
    Else
    
        d51ValorOp = 0
        objVar51ValorOp.iStatus = VAR_PREENCH_AUTOMATICO
        dPercDifer = 0
        objVarPercDifer.iStatus = VAR_PREENCH_AUTOMATICO
        dValorDif = 0
        objVarValorDif.iStatus = VAR_PREENCH_AUTOMATICO
    
    End If
    
    'NFe 4.00
    'Base de cálculo dos FCP iguais as bases correspondentes do ICMS
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCP = dBaseCalculo
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCP = objVarICMSvBCFCP.vValor
    End If
    
    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvBCFCPST = dBaseSubst
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvBCFCPST = objVarICMSvBCFCPST.vValor
    End If
           
    'Calculando o valor com a base e alíquota
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCP = Arredonda_Moeda(dICMSvBCFCP * dICMSpFCP)
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCP = objVarICMSvFCP.vValor
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
        dICMSvFCPST = Arredonda_Moeda(dICMSvBCFCPST * dICMSpFCPST)
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dICMSvFCPST = objVarICMSvFCPST.vValor
    End If
    'NFe 4.00
    

'        objVarDespImpICMSBase.sIdentificador = "DESP_IMP_ICMS_BASE"
'        objVarDespImpICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSBase)
'        If lErro <> SUCESSO Then Error 27307
'
'        objVarDespImpICMSValor.sIdentificador = "DESP_IMP_ICMS_VALOR"
'        objVarDespImpICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSValor)
'        If lErro <> SUCESSO Then Error 27307
'
'        objVarDespImpICMSCredito.sIdentificador = "DESP_IMP_ICMS_CREDITO"
'        objVarDespImpICMSCredito.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSCredito)
'        If lErro <> SUCESSO Then Error 27307

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarBaseSubst, objVarMargSubst, objVarAliqSubst, objVarValorSubst, objVarPercRedBaseSubst, objVarValorCredito, objVarModalidade, objVarModalidadeST, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorSubst, dAliqSubst, dMargSubst, dBaseSubst, dPercRedBaseSubst, dValorCredito, iModalidade, iModalidadeST, _
                                objVarVlrCredSimples, objVarAliCredSimples, objVarTipoICMSSimples, objVarCobrAntValor, objVarCobrAntBase, objVarPBOP, objVarUFD, objVarBCSTRet, objVarICMSSTRet, objVarBCSTDest, objVarICMSSTDest, objVarVI, objVarMotivo, objVarRegTrib, objVarOrigemProd, objVarFCI, objVar51ValorOp, objVarPercDifer, objVarValorDif, d51ValorOp, dPercDifer, dValorDif, iMotivo)
    If lErro <> SUCESSO Then Error 27996
    
    Call AtualizarItem_GuardarValoresImp(objDocItem, dImpICMSBase, dImpICMSValor, dImpICMSCredito)
    
    lErro = AtualizarItem_GuardarValores2(objDocItem, objVarInterestBCUFDest, objVarInterestPercFCPUFDest, objVarInterestAliqUFDest, objVarInterestAliq, objVarInterestPercPartilha, objVarInterestVlrFCPUFDest, objVarInterestVlrUFDest, objVarInterestVlrUFRemet, _
                                dInterestBCUFDest, dInterestPercFCPUFDest, dInterestAliqUFDest, dInterestAliq, dInterestPercPartilha, dInterestVlrFCPUFDest, dInterestVlrUFDest, dInterestVlrUFRemet, objVarICMSInterestBCFCPUFDest, dICMSInterestBCFCPUFDest)
    If lErro <> SUCESSO Then Error 27996

    lErro = AtualizarItem_GuardarValores3(objDocItem, objVarICMSvBCFCP, objVarICMSpFCP, objVarICMSvFCP, dICMSvBCFCP, dICMSpFCP, dICMSvFCP, objVarICMSvBCFCPST, objVarICMSpFCPST, objVarICMSvFCPST, dICMSvBCFCPST, dICMSpFCPST, dICMSvFCPST, objVarICMSvBCFCPSTRet, objVarICMSpFCPSTRet, objVarICMSvFCPSTRet, dICMSvBCFCPSTRet, dICMSpFCPSTRet, dICMSvFCPSTRet, objVarICMSSTCobrAntAliquota, dICMSSTCobrAntAliquota, objVarICMSSTBaseDupla, iICMSSTBaseDupla, objVarICMSInterestBaseDupla, iICMSInterestBaseDupla, objVarcBenef, scBenef)
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
Dim objVarCST As New ClassVariavelCalculo, objVarVI As New ClassVariavelCalculo, dValorIsento As Double

'NFE 4.00
Dim dICMSVlrFCP As Double, dICMSVlrFCPST As Double, dICMSVlrFCPSTRet As Double, dIPIVlrDevolvido As Double
Dim objVarICMSVlrFCP As New ClassVariavelCalculo, objVarICMSVlrFCPST As New ClassVariavelCalculo, objVarICMSVlrFCPSTRet As New ClassVariavelCalculo, objVarIPIVlrDevolvido As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc_Importacao

    dValor = 0
    dValorSubst = 0
    dBase = 0
    dBaseReduzida = 0
    dBaseSubst = 0
    dBaseSubstReduzida = 0
    dValorCredito = 0
    dValorIsento = 0
        
    'NFe 4.00
    dICMSVlrFCP = 0
    dICMSVlrFCPST = 0
    dICMSVlrFCPSTRet = 0
    dIPIVlrDevolvido = 0
    
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
        
        lErro = objDocItem.ObterVar("ICMS_CST", objVarCST)
        If lErro <> SUCESSO Then Error 27954
        
'        If right(objVarCST.vValor, 2) <> "51" Then
            dBase = dBase + Arredonda_Moeda(objVarICMSBase.vValor, 2)
'            If gobjCRFAT.iNFImportacaoTribFlag09 = DESMARCADO Then
                dBaseReduzida = dBaseReduzida + Arredonda_Moeda(objVarICMSBase.vValor * (1 - dPercRedBase), 2)
'            Else
'                dBaseReduzida = dBaseReduzida + objVarICMSBase.vValor
'            End If
            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
            dBaseSubst = dBaseSubst + Arredonda_Moeda(objVarICMSBaseSubst.vValor, 2)
            dPercRedBaseSubst = 0 '???? ainda nao está sendo armazenado
            dBaseSubstReduzida = dBaseSubstReduzida + Arredonda_Moeda(objVarICMSBaseSubst.vValor * (1 - dPercRedBaseSubst), 2)
            ''dBaseSubstReduzida = CDbl(Format(dBaseSubstReduzida, "0.00"))
            dValor = dValor + Arredonda_Moeda(objVarICMSValor.vValor, 2)
            dValorSubst = dValorSubst + Arredonda_Moeda(objVarICMSValorSubst.vValor, 2)
            dValorCredito = dValorCredito + Arredonda_Moeda(objVarICMSValorCredito.vValor, 2)
'        End If
        
        'nfe 3.10
        lErro = objDocItem.ObterVar("ICMS_VALOR_ISENTO", objVarVI)
        If lErro <> SUCESSO Then Error 27949
        
        dValorIsento = dValorIsento + Arredonda_Moeda(objVarVI.vValor, 2)

'        'especifico de importacao
'        'lErro = objDocItem.ObterVar("PRODUTO_OUTROS_VALORES", objVarDespImpValorRateado)
'        lErro = objDocItem.ObterVar("DESP_IMP_VALOR_RATEADO", objVarDespImpValorRateado)
'        If lErro <> SUCESSO Then Error 27949
'
'        lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_ADICAO_DI", objVarAliquotaDI)
'        If lErro <> SUCESSO Then Error 27994
'
'        If dBaseReduzida <> 0 Then
'            dAliquotaImp = objVarAliquotaDI.vValor
'            dBaseImp = Arredonda_Moeda(objVarDespImpValorRateado.vValor / (1 - dAliquotaImp))
'        Else
'            dAliquotaImp = 0
'            dBaseImp = 0
'        End If
'
'        objVarDespImpICMSBase.vValor = dBaseImp
'        objVarDespImpICMSValor.vValor = Arredonda_Moeda(dBaseImp * dAliquotaImp)
'        objVarDespImpICMSCredito.vValor = objVarDespImpICMSValor.vValor
'
'        objVarDespImpICMSBase.sIdentificador = "DESP_IMP_ICMS_BASE"
'        objVarDespImpICMSBase.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSBase)
'        If lErro <> SUCESSO Then Error 27307
'
'        objVarDespImpICMSValor.sIdentificador = "DESP_IMP_ICMS_VALOR"
'        objVarDespImpICMSValor.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSValor)
'        If lErro <> SUCESSO Then Error 27307
'
'        objVarDespImpICMSCredito.sIdentificador = "DESP_IMP_ICMS_CREDITO"
'        objVarDespImpICMSCredito.iStatus = VAR_PREENCH_AUTOMATICO
'        lErro = objDocItem.GuardarVar(objVarDespImpICMSCredito)
'        If lErro <> SUCESSO Then Error 27307
        
'        If dBaseReduzida <> 0 Then
'            dBase = dBase + objVarDespImpICMSBase.vValor
'            dBaseReduzida = dBaseReduzida + objVarDespImpICMSBase.vValor
'            dValor = dValor + objVarDespImpICMSValor.vValor
'            dValorCredito = dValorCredito + objVarDespImpICMSCredito.vValor
'        End If

        'NFe 4.00
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP", objVarICMSVlrFCP)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST", objVarICMSVlrFCPST)
        If lErro <> SUCESSO Then Error 27954
        
        lErro = objDocItem.ObterVar("ICMS_VLR_FCP_ST_RET", objVarICMSVlrFCPSTRet)
        If lErro <> SUCESSO Then Error 27954
           
        lErro = objDocItem.ObterVar("IPI_VLR_DEVOLVIDO", objVarIPIVlrDevolvido)
        If lErro <> SUCESSO Then Error 27954
           
        dICMSVlrFCP = dICMSVlrFCP + objVarICMSVlrFCP.vValor
        dICMSVlrFCPST = dICMSVlrFCPST + objVarICMSVlrFCPST.vValor
        dICMSVlrFCPSTRet = dICMSVlrFCPSTRet + objVarICMSVlrFCPSTRet.vValor
        dIPIVlrDevolvido = dIPIVlrDevolvido + objVarIPIVlrDevolvido.vValor
    Next
    
    'conta de chegar
''    lErro = objDoc.ObterVar("BASE_ICMS_IMPORTACAO_SIMPLES", objVarICMSBaseSimples)
''    If lErro <> SUCESSO Then Error 27949
''    dICMSBaseSimples = Arredonda_Moeda(objVarICMSBaseSimples.vValor + dValor, 2)
''    If dBaseReduzida = dBase And dICMSBaseSimples <> dBase And Abs(dICMSBaseSimples - dBase) <= 0.02 Then
''        dBase = dICMSBaseSimples
''        dBaseReduzida = dBase
''    End If
        
    'arredondar valores globais
    dBase = Arredonda_Moeda(dBase, 2)
    dBaseReduzida = Arredonda_Moeda(dBaseReduzida, 2)
    dBaseSubst = Arredonda_Moeda(dBaseSubst, 2)
    dBaseSubstReduzida = Arredonda_Moeda(dBaseSubstReduzida, 2)
    dValor = Arredonda_Moeda(dValor, 2)
    dValorSubst = Arredonda_Moeda(dValorSubst, 2)
    dValorCredito = Arredonda_Moeda(dValorCredito, 2)
        
    'nfe 3.10
    dValorIsento = Arredonda_Moeda(dValorIsento, 2)
    
    'nfe 4.00
    dICMSVlrFCP = Arredonda_Moeda(dICMSVlrFCP, 2)
    dICMSVlrFCPST = Arredonda_Moeda(dICMSVlrFCPST, 2)
    dICMSVlrFCPSTRet = Arredonda_Moeda(dICMSVlrFCPSTRet, 2)
    dIPIVlrDevolvido = Arredonda_Moeda(dIPIVlrDevolvido, 2)
    
    lErro = AtualizarDoc_GuardarValores(objDoc, dBaseReduzida, dValor, dBaseSubstReduzida, dValorSubst, dValorCredito, dValorIsento, 0, 0, 0, dICMSVlrFCP, dICMSVlrFCPST, dICMSVlrFCPSTRet, dIPIVlrDevolvido)
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

Private Function TipoICMSZerarBase(ByVal iTipoICMS As Integer) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMS
    
        Case ICMS_TIPO_COBR_SUBST_ANT, ICMS_TIPO_NAO_TRIBUTADO, ICMS_TIPO_ISENTO, ICMS_TIPO_SUSPENSO
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    TipoICMSZerarBase = bRet
    
End Function

Private Function TipoICMSZerar(ByVal iTipoICMS As Integer, Optional ByVal iTipoNFiscal As Integer = 0) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMS
    
        Case ICMS_TIPO_COBR_SUBST_ANT, ICMS_TIPO_NAO_TRIBUTADO, ICMS_TIPO_ISENTO
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    If (iTipoNFiscal = DOCINFO_NFIFVEFPV Or iTipoNFiscal = DOCINFO_NFIFVEF) And gobjCRFAT.iVendaFuturaTributacao = 1 Then bRet = True
   
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

Private Function AtualizarItem_GuardarValoresImp(objDocItem As ClassTributoDocItem, ByVal dImpICMSBase As Double, ByVal dImpICMSValor As Double, ByVal dImpICMSCredito As Double) As Long

Dim lErro As Long
Dim objVarDespImpICMSBase As New ClassVariavelCalculo, objVarDespImpICMSValor As New ClassVariavelCalculo
Dim objVarDespImpICMSCredito As New ClassVariavelCalculo

    objVarDespImpICMSBase.vValor = dImpICMSBase
    objVarDespImpICMSValor.vValor = dImpICMSValor
    objVarDespImpICMSCredito.vValor = dImpICMSCredito
    
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

End Function

Private Function ItemComSTCobradaAnt(ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, iSTCobrAnt As Integer) As Long
'retorna se o produto foi comprado com icms ST

Dim lErro As Long, objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto, sProduto As String
Dim lComando As Long, lNumIntDocItem As Long

On Error GoTo Erro_ItemComSTCobradaAnt

    iSTCobrAnt = 0
    
    If left(gsNomeEmpresa, 3) = "PBJ" Or gobjCRFAT.iTrataSTEntEmCST00 = 2 Then
    
        iSTCobrAnt = 1
    
    Else
    
        lComando = Comando_Abrir()
        If lComando = 0 Then gError 201490
    
        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
        If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
        
            Set objProduto = objVarProd.vValor
            
            sProduto = objProduto.sCodigo
                
            lErro = Comando_Executar(lComando, "SELECT NumIntDocItem FROM TributacaoDocItem TDI, NFiscal NF, TiposDocInfo As T WHERE TDI.TipoDoc=0 AND TDI.NumIntDoc = NF.NumIntDoc AND NF.TipoNFiscal = T.Codigo AND TDI.Produto = ? AND TDI.ICMSSubstValor <> 0 AND TDI.NaturezaOp < '4' AND T.Compras = 1 AND NF.Status <> 7 AND NF.FilialEmpresa = ? ", _
                lNumIntDocItem, sProduto, mvariFilialEmpresa)
            If lErro <> AD_SQL_SUCESSO Then gError 201491
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201492
                
            If lErro = AD_SQL_SUCESSO Then iSTCobrAnt = 1
                
        End If
                
        'Fecha os comandos
        Call Comando_Fechar(lComando)
    
    End If
    
    ItemComSTCobradaAnt = SUCESSO
    
    Exit Function
    
Erro_ItemComSTCobradaAnt:

    ItemComSTCobradaAnt = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201490
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201489)

    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecaoEC872015(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, ByVal sUFOrigem As String, sUFDestino As String, dInterestAliqUFDest As Double, dInterestPercFCPUFDest As Double, objVarInterestPercFCPUFDest As ClassVariavelCalculo, objVarInterestAliqUFDest As ClassVariavelCalculo, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo) As Long

Dim lErro As Long, objVarCliCodigo As New ClassVariavelCalculo, dAliquotaPropria As Double
Dim objVarNaoCalcST As New ClassVariavelCalculo, iTipoCliente As Integer, dAliquotaNoDestino As Double, dAliquotaFCPNoDestino As Double
Dim dAliquotaPropriaFCP As Double

On Error GoTo Erro_AtualizarItem_TratarNaoExcecaoEC872015
                    
    'se ICMS_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        lErro = objDocItem.objDoc.ObterVar("ICMS_ST_NAO_CALCULAR", objVarNaoCalcST)
        If lErro <> SUCESSO Then Error 27317

        If objVarNaoCalcST.vValor <> MARCADO Then
            iTipo = objTipoTrib.iICMSTipo
        Else
            iTipo = ICMS_TIPO_NORMAL
        End If
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribICMS.iTipo = iTipo
    lErro = TipoICMS_Obter(objTipoTribICMS)
    If lErro <> SUCESSO Then Error 27317
        
    lErro = ObterAliquotaBasicaEC872015(objDocItem.objDoc, objTipoTrib, objTipoTribICMS, dAliquotaPropria, sUFDestino, objDocItem, dAliquotaNoDestino, dAliquotaFCPNoDestino, dAliquotaPropriaFCP)
    If lErro <> SUCESSO Then Error 27318
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            dAliquota = dAliquotaPropria
        Else
            dAliquota = 0
        End If
    Else
        dAliquota = objVarAliquota.vValor
    End If
    
    If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            dICMSpFCP = dAliquotaPropriaFCP
        Else
            dICMSpFCP = 0
        End If
    Else
        dICMSpFCP = objVarICMSpFCP.vValor
    End If
    
    If objVarInterestPercFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            dInterestPercFCPUFDest = dAliquotaFCPNoDestino
        Else
            dInterestPercFCPUFDest = 0
        End If
    Else
    
        dInterestPercFCPUFDest = objVarInterestPercFCPUFDest.vValor
        
    End If
    
    If objVarInterestAliqUFDest.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            dInterestAliqUFDest = dAliquotaNoDestino
        Else
            dInterestAliqUFDest = 0
        End If
    Else
    
        dInterestAliqUFDest = objVarInterestAliqUFDest.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        dPercRedBase = 0
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecaoEC872015 = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecaoEC872015:

    AtualizarItem_TratarNaoExcecaoEC872015 = Err

    Select Case Err

        Case 27317, 27318
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153829)

    End Select

    Exit Function

End Function

Private Function ObterAliquotaBasicaEC872015(objDoc As ClassTributoDoc, objTributacaoTipo As ClassTipoDeTributacaoMovto, objTipoTribICMS As ClassTipoTribICMS, dAliquotaPropria As Double, sUFDestino As String, ByVal objDocItem As ClassTributoDocItem, dInterestAliqUFDest As Double, dInterestPercFCPUFDest As Double, dAliquotaPropriaFCP As Double) As Long
'preenche a aliquota

Dim lErro As Long, objICMSAliqExterna As ClassICMSAliqExterna
Dim objVarOrigemProd As New ClassVariavelCalculo
Dim objEstado As ClassEstado
Dim dICMSPercFCPAux As Double, dICMSAliquotaInternaAux As Double

On Error GoTo Erro_ObterAliquotaBasicaEC872015

    'usar aliq interestadual para o icms proprio
    
    lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
    If lErro <> SUCESSO Then Error 27876
    
    Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sUFDestino)
    
    Select Case objVarOrigemProd.vValor
    
        Case PRODUTO_ORIGEM_IMP_DIRETA, PRODUTO_ORIGEM_EST_MERC_INTERNO, PRODUTO_ORIGEM_NAC_COM_IMP, PRODUTO_ORIGEM_NAC_COM_IMP2
            dAliquotaPropria = 0.04
            
        Case Else
            'Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(sUFDestino)
            
            dAliquotaPropria = objICMSAliqExterna.dAliquota
    
    End Select
    dAliquotaPropriaFCP = objICMSAliqExterna.dAliquotaFCP
    
    'obter o icms padrao e fcp (=fundo de combate à pobreza) no destino
    Set objEstado = mvarcolEstados.Item(sUFDestino)
    
    If objEstado.dtDataIniAliqInternaAtual <> DATA_NULA And objEstado.dtDataIniAliqInternaAtual > mvardtDataRef Then
        dICMSAliquotaInternaAux = objEstado.dICMSAliquotaInternaAnt
    Else
        dICMSAliquotaInternaAux = objEstado.dICMSAliquotaInterna
    End If
    
    If objEstado.dtDataIniAliqFCPAtual <> DATA_NULA And objEstado.dtDataIniAliqFCPAtual > mvardtDataRef Then
        dICMSPercFCPAux = objEstado.dICMSPercFCPAnt
    Else
        dICMSPercFCPAux = objEstado.dICMSPercFCP
    End If
    
    dInterestPercFCPUFDest = dICMSPercFCPAux
    dInterestAliqUFDest = dICMSAliquotaInternaAux - dICMSPercFCPAux 'sem incluir o fcp
    
    ObterAliquotaBasicaEC872015 = SUCESSO

    Exit Function

Erro_ObterAliquotaBasicaEC872015:

    ObterAliquotaBasicaEC872015 = Err

    Select Case Err

        Case 27284, 27285
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153825)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarExcecaoEC872015(objDocItem As ClassTributoDocItem, objExcecao As ClassICMSExcecao, objTipoTribICMS As ClassTipoTribICMS, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, objVarMargSubst As ClassVariavelCalculo, dMargSubst As Double, objVarPercRedBaseSubst As ClassVariavelCalculo, dPercRedBaseSubst As Double, ByVal bSemSubstTributaria As Boolean, iUsaPauta As Integer, dValorPauta As Double, dInterestAliqUFDest As Double, dInterestPercFCPUFDest As Double, objVarInterestPercFCPUFDest As ClassVariavelCalculo, objVarInterestAliqUFDest As ClassVariavelCalculo, dICMSpFCP As Double, objVarICMSpFCP As ClassVariavelCalculo, ByRef iMotivo As Integer, ByRef objVarMotivo As ClassVariavelCalculo, ByRef scBenef As String, ByRef objVarcBenef As ClassVariavelCalculo) As Long

Dim lErro As Long
Dim objVarOrigemProd As New ClassVariavelCalculo, objICMSAliqExterna As ClassICMSAliqExterna

On Error GoTo Erro_AtualizarItem_TratarExcecaoEC872015

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
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Or objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            
            'a aliquota propria é a interestadual variando de acordo com a origem
            
            lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
            If lErro <> SUCESSO Then Error 27876
            
            Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(objExcecao.sEstadoDestino)
            
            Select Case objVarOrigemProd.vValor
            
                Case PRODUTO_ORIGEM_IMP_DIRETA, PRODUTO_ORIGEM_EST_MERC_INTERNO, PRODUTO_ORIGEM_NAC_COM_IMP, PRODUTO_ORIGEM_NAC_COM_IMP2
                    dAliquota = 0.04
                    
                Case Else
'                    Set objICMSAliqExterna = mvarcolAliqInterstSaida.Item(objExcecao.sEstadoDestino)
                    
                    dAliquota = objICMSAliqExterna.dAliquota
            
            End Select
            dICMSpFCP = objICMSAliqExterna.dAliquotaFCP
        
        Else
            dAliquota = 0
            dICMSpFCP = 0
        End If
'    Else
'
'        dAliquota = objVarAliquota.vValor
        
    End If
    
    'Se estiver manual ignora o que foi calculado
    If objVarAliquota.iStatus = VAR_PREENCH_MANUAL Then dAliquota = objVarAliquota.vValor
    If objVarICMSpFCP.iStatus = VAR_PREENCH_MANUAL Then dICMSpFCP = objVarICMSpFCP.vValor
    
    If objVarInterestPercFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            'o FCP vem da excecao
            dInterestPercFCPUFDest = objExcecao.dICMSPercFCP
        Else
            dInterestPercFCPUFDest = 0
        End If
    Else
    
        dInterestPercFCPUFDest = objVarInterestPercFCPUFDest.vValor
        
    End If
    
    If objVarInterestAliqUFDest.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribICMS.iPermiteAliquota <> 0 Then
            'a aliquota do destino vem da excecao
            dInterestAliqUFDest = objExcecao.dAliquota
        Else
            dInterestAliqUFDest = 0
        End If
    Else
    
        dInterestAliqUFDest = objVarInterestAliqUFDest.vValor
        
    End If
    
    'se ICMS_RED_BASE nao tiver sido preenchido manualmente
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        'dPercRedBase = 0 '??? acho que nao faz sentido ter reducao: objExcecao.dPercRedBaseCalculo
        dPercRedBase = objExcecao.dPercRedBaseCalculo
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
    
    'se ICMS_TIPO permite subst
    If objTipoTribICMS.iPermiteMargLucro And (bSemSubstTributaria = False) Then
        If objVarPercRedBaseSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBaseSubst = objExcecao.dPercRedBaseCalculoSubst
        Else
            dPercRedBaseSubst = objVarPercRedBaseSubst.vValor
        End If
        If objVarMargSubst.iStatus <> VAR_PREENCH_MANUAL Then
            dMargSubst = objExcecao.dPercMargemLucro
        Else
            dMargSubst = objVarMargSubst.vValor
        End If
        dValorPauta = objExcecao.dValorPauta
        iUsaPauta = objExcecao.iUsaPauta
    Else
        dMargSubst = 0
        dPercRedBaseSubst = 0
        dValorPauta = 0
        iUsaPauta = 0
    End If
     
    If objVarMotivo.iStatus <> VAR_PREENCH_MANUAL Then
       iMotivo = objExcecao.iICMSMotivo
    Else
       iMotivo = objVarMotivo.vValor
    End If
     
    If objVarcBenef.iStatus <> VAR_PREENCH_MANUAL Then
       scBenef = objExcecao.scBenef
    Else
       scBenef = objVarcBenef.vValor
    End If
     
    AtualizarItem_TratarExcecaoEC872015 = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarExcecaoEC872015:

    AtualizarItem_TratarExcecaoEC872015 = Err

    Select Case Err

        Case 27316
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153828)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_GuardarValores2(objDocItem As ClassTributoDocItem, objVarInterestBCUFDest As ClassVariavelCalculo, objVarInterestPercFCPUFDest As ClassVariavelCalculo, objVarInterestAliqUFDest As ClassVariavelCalculo, objVarInterestAliq As ClassVariavelCalculo, objVarInterestPercPartilha As ClassVariavelCalculo, _
        objVarInterestVlrFCPUFDest As ClassVariavelCalculo, objVarInterestVlrUFDest As ClassVariavelCalculo, objVarInterestVlrUFRemet As ClassVariavelCalculo, dInterestBCUFDest As Double, dInterestPercFCPUFDest As Double, dInterestAliqUFDest As Double, dInterestAliq As Double, dInterestPercPartilha As Double, dInterestVlrFCPUFDest As Double, dInterestVlrUFDest As Double, dInterestVlrUFRemet As Double, objVarICMSInterestBCFCPUFDest As ClassVariavelCalculo, dICMSInterestBCFCPUFDest As Double) As Integer
'atualiza as variaveis que tiveram seus valores recalculados pelo sistema

Dim lErro As Long

On Error GoTo Erro_AtualizarItem_GuardarValores2

    'NT 2015.003 (ICMS Interestadual)
    If objVarInterestBCUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestBCUFDest.vValor = dInterestBCUFDest
        objVarInterestBCUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestBCUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestPercFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestPercFCPUFDest.vValor = dInterestPercFCPUFDest
        objVarInterestPercFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestPercFCPUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestAliqUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestAliqUFDest.vValor = dInterestAliqUFDest
        objVarInterestAliqUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestAliqUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestAliq.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestAliq.vValor = dInterestAliq
        objVarInterestAliq.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestAliq)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestPercPartilha.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestPercPartilha.vValor = dInterestPercPartilha
        objVarInterestPercPartilha.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestPercPartilha)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestVlrFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestVlrFCPUFDest.vValor = dInterestVlrFCPUFDest
        objVarInterestVlrFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestVlrFCPUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestVlrUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestVlrUFDest.vValor = dInterestVlrUFDest
        objVarInterestVlrUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestVlrUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarInterestVlrUFRemet.iStatus <> VAR_PREENCH_MANUAL Then
        objVarInterestVlrUFRemet.vValor = dInterestVlrUFRemet
        objVarInterestVlrUFRemet.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarInterestVlrUFRemet)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    If objVarICMSInterestBCFCPUFDest.iStatus <> VAR_PREENCH_MANUAL Then
        objVarICMSInterestBCFCPUFDest.vValor = dICMSInterestBCFCPUFDest
        objVarICMSInterestBCFCPUFDest.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSInterestBCFCPUFDest)
        If lErro <> SUCESSO Then Error 27874
    End If
    
    AtualizarItem_GuardarValores2 = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores2:

    AtualizarItem_GuardarValores2 = Err

    Select Case Err

        Case 27307 To 27315, 27874
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153827)

    End Select

    Exit Function

End Function

Private Function TipoICMSSimplesComBaseEC872015(ByVal iTipoICMSSimples As Integer) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMSSimples
    
        Case 1, 2, 9
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    TipoICMSSimplesComBaseEC872015 = bRet
    
End Function

Private Function TipoICMSComBaseEC872015(ByVal iTipoICMS As Integer) As Boolean

Dim bRet As Boolean

    Select Case iTipoICMS
    
        Case ICMS_TIPO_NORMAL, ICMS_TIPO_RED_BASE, ICMS_TIPO_COBR_SUBST_ANT
            bRet = True
        
        Case Else
            bRet = False
        
    End Select

    TipoICMSComBaseEC872015 = bRet
    
End Function

Private Function AtualizarItem_GuardarValores3(objDocItem As ClassTributoDocItem, objVarICMSvBCFCP As ClassVariavelCalculo, objVarICMSpFCP As ClassVariavelCalculo, objVarICMSvFCP As ClassVariavelCalculo, dICMSvBCFCP As Double, dICMSpFCP As Double, dICMSvFCP As Double, objVarICMSvBCFCPST As ClassVariavelCalculo, objVarICMSpFCPST As ClassVariavelCalculo, objVarICMSvFCPST As ClassVariavelCalculo, dICMSvBCFCPST As Double, dICMSpFCPST As Double, dICMSvFCPST As Double, objVarICMSvBCFCPSTRet As ClassVariavelCalculo, objVarICMSpFCPSTRet As ClassVariavelCalculo, objVarICMSvFCPSTRet As ClassVariavelCalculo, dICMSvBCFCPSTRet As Double, dICMSpFCPSTRet As Double, dICMSvFCPSTRet As Double, objVarICMSSTCobrAntAliquota As ClassVariavelCalculo, dICMSSTCobrAntAliquota As Double, objVarICMSSTBaseDupla As ClassVariavelCalculo, iICMSSTBaseDupla As Integer, objVarICMSInterestBaseDupla As ClassVariavelCalculo, iICMSInterestBaseDupla As Integer, objVarcBenef As ClassVariavelCalculo, scBenef As String) As Integer

Dim lErro As Long

On Error GoTo Erro_AtualizarItem_GuardarValores3
    
    If objVarICMSvBCFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvBCFCP.vValor = dICMSvBCFCP
        objVarICMSvBCFCP.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvBCFCP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSpFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSpFCP.vValor = dICMSpFCP
        objVarICMSpFCP.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSpFCP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSvFCP.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvFCP.vValor = dICMSvFCP
        objVarICMSvFCP.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvFCP)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    If objVarICMSvBCFCPST.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvBCFCPST.vValor = dICMSvBCFCPST
        objVarICMSvBCFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvBCFCPST)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSpFCPST.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSpFCPST.vValor = dICMSpFCPST
        objVarICMSpFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSpFCPST)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSvFCPST.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvFCPST.vValor = dICMSvFCPST
        objVarICMSvFCPST.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvFCPST)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSvBCFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvBCFCPSTRet.vValor = dICMSvBCFCPSTRet
        objVarICMSvBCFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvBCFCPSTRet)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSpFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSpFCPSTRet.vValor = dICMSpFCPSTRet
        objVarICMSpFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSpFCPSTRet)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSvFCPSTRet.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSvFCPSTRet.vValor = dICMSvFCPSTRet
        objVarICMSvFCPSTRet.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSvFCPSTRet)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSSTCobrAntAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSSTCobrAntAliquota.vValor = dICMSSTCobrAntAliquota
        objVarICMSSTCobrAntAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSSTCobrAntAliquota)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSSTBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSSTBaseDupla.vValor = iICMSSTBaseDupla
        objVarICMSSTBaseDupla.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSSTBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarICMSInterestBaseDupla.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarICMSInterestBaseDupla.vValor = iICMSInterestBaseDupla
        objVarICMSInterestBaseDupla.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarICMSInterestBaseDupla)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
     
    If objVarcBenef.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarcBenef.vValor = scBenef
        objVarcBenef.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarcBenef)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
     
    AtualizarItem_GuardarValores3 = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores3:

    AtualizarItem_GuardarValores3 = Err

    Select Case Err

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153827)

    End Select

    Exit Function

End Function
