VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoIPIArtmill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'saidaCli: já foi feito
'entradacli: vou tratar como saida cli mas vou guardar a info se credita do imposto
'entradaforn: criar tratamento especifico (usar padraotribentrada, )
'saidaforn: criar tratamento analogo ao de entradaforn mas nao vou ter info de credito

'ver apostila IOB pg 16
'a base de calculo tem que poder ser informada p/tratar casos de interdependencia entre empresas
'inclui frete, despesas acessorias, ...
'nao podem ser abatidos descontos, abatimentos,...
'incluindo embalagem, mesmo que esta seja cobrada em separado.

'classificacao dos clientes p/IPI, que deve ser colocada a nivel de categoria:
' industrial, revendedor, equiparado a industrial,

'grupos p/tratamento em separado:
'produtos nao industrializados,
'produtos acabados,
'materias-primas p/producao,
'materias-primas p/consumo no processo de producao,
'materias-primas p/emprego no processo de producao,
'produtos intermediarios p/producao,
'produtos intermediarios p/consumo no processo de producao,
'produtos intermediarios p/emprego no processo de producao,
'material de embalagem p/emprego no processo de producao

'processo: ver apostila pg 26

''
''Duvidas:
''    opcao do calculo do imposto sobre 50% do valor tributavel
''
''
''fontes de consulta:
''
''FilialEmpresa: se a filial é ou nao contribuinte do IPI
''TiposDocInfo: se calcula IPI
''Produtos: aliquota, peso (p/ratear o frete)
''Clientes (ou filial cliente ?): se é ZFM (AM, AC, RO, RR) (pedir cod SUFRAMA no cadastro?)
''IPIExcecoes:
''
''variaveis:
''    como a Empresa (ou a Filial ?) trata o IPI, tipo do doc, natureza da operacao,
''    produto, tipo IPI, qtde, valor produto, aliq IPI, valor IPI, se o cliente é ZFM,
''    valor do credito
''
'''se a filial nao estiver inscrita no IPI
''    tipo
''    aliquota zero, valor zero
''
''  pode ter cache p/agilizar calculos

Implements ClassTributoDriver

Private mvardFatorValor As Double 'para multiplicar o valor real obtendo o oficial
Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvariContribuinte As Integer
Private mvariSimplesFederal As Integer

Private mvarcolExcecoes As New Collection 'colecao de excecoes IPI
            'já ordenada, primeiro a nivel de catcliente+catproduto,
            'depois a nivel de catcliente e por ultimo a nivel de catproduto
Private mvarcolTiposTribIPI As New Collection 'colecao com os tipos de tributacao p/IPI

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao do doc como um todos e de seus itens
Dim lErro As Long, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
Dim objVarSerie As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32024
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 22777
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22778
    
    lErro = objDoc.ObterVar("SERIE", objVarSerie)
    If lErro <> SUCESSO Then Error 22777
            
    lErro = CF("NFiscal_ObtemFatorValor", mvariFilialEmpresa, objVarTipoDocInfo.vValor, objVarSerie.vValor, mvardFatorValor)
    If lErro <> SUCESSO Then Error 22777
    
    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_ENTRADA_CLI
            lErro = AtualizarImposto_Cli(objDoc, 1, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22779

        Case TRIB_ENTRADA_FORN
            lErro = AtualizarImposto_Forn(objDoc, 1, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22780

        Case TRIB_SAIDA_CLI
            lErro = AtualizarImposto_Cli(objDoc, 0, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22781

        Case TRIB_SAIDA_FORN
            lErro = AtualizarImposto_Forn(objDoc, 0, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22782
        
    End Select
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 22777 To 22782, 32024
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154065)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
'permite a inicializacao do driver
'pode ser utilizado p/criar cache etc
Dim lErro As Long
On Error GoTo Erro_ClassTributoDriver_Ativar

    lErro = CF("TiposTribIPI_Le_Todos", mvarcolTiposTribIPI)
    If lErro <> SUCESSO Then Error 27517
    
    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err
        
        Case 27515
            '???
        
        Case 27516, 27517
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154066)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Desativar() As Long
    'informa que o driver nao vai mais ser necessario
    'pode ser utilizado p/liberar cache etc
    ClassTributoDriver_Desativar = SUCESSO
    
End Function

'FUNCOES AUXILIARES:
'===================
Private Function PrepararAtualizacao_Cli(objDoc As ClassTributoDoc) As Long
Dim lErro As Long, lCliente As Long, iFilial As Integer
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
On Error GoTo Erro_PrepararAtualizacao_Cli

    'obter codigo do cliente
    lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
    If lErro <> SUCESSO Then Error 27595
    
    'obter identificador da filial do cliente
    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
    If lErro <> SUCESSO Then Error 27596
            
    lCliente = objVarCliente.vValor
    iFilial = objVarFilialCliente.vValor
    
    'carregar colecao de excecoes IPI
    lErro = CF("IPIExcecoes_Le_FilialCliente", lCliente, iFilial, mvarcolExcecoes)
    If lErro <> SUCESSO Then Error 27530
    
    PrepararAtualizacao_Cli = SUCESSO

    Exit Function

Erro_PrepararAtualizacao_Cli:

    PrepararAtualizacao_Cli = Err

    Select Case Err
    
        Case 27530, 27595, 27596, 27658
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154067)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc) As Long
'atualiza base de IPI, valor de IPI
'    leva em conta dados dos itens + dados de frete, seg, desp acess.
Dim colItens As Collection
Dim objDocItem As ClassTributoDocItem, dAliquota As Double
Dim dValor As Double, dBase As Double, dBaseReduzida As Double
Dim objVarIPIBase As New ClassVariavelCalculo, objVarIPIValor As New ClassVariavelCalculo
Dim objVarFreteValor As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
Dim objVarSeguroValor As New ClassVariavelCalculo, objVarDespAcesValor As New ClassVariavelCalculo
Dim lErro As Long, dPercRedBase As Double, dValorCredito As Double, objVarIPIValorCredito As New ClassVariavelCalculo
On Error GoTo Erro_AtualizarDoc

    dValor = 0
    dBase = 0
    dBaseReduzida = 0
    dValorCredito = 0
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
        
        'obter do item IPI_BASE, IPI_VALOR
        
        lErro = objDocItem.ObterVar("IPI_BASE", objVarIPIBase)
        If lErro <> SUCESSO Then Error 27531
        
        lErro = objDocItem.ObterVar("IPI_RED_BASE", objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27532
        
        dPercRedBase = objVarPercRedBase.vValor
        
        lErro = objDocItem.ObterVar("IPI_VALOR", objVarIPIValor)
        If lErro <> SUCESSO Then Error 27533
        
        lErro = objDocItem.ObterVar("IPI_VALOR_CREDITO", objVarIPIValorCredito)
        If lErro <> SUCESSO Then Error 22787
        
        dBase = dBase + objVarIPIBase.vValor
        
        dBaseReduzida = dBaseReduzida + (objVarIPIBase.vValor * (1 - dPercRedBase))
        ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
        
        dValor = dValor + objVarIPIValor.vValor
        dValorCredito = dValorCredito + objVarIPIValorCredito.vValor
        
    Next
    
    'tratar dados globais (frete, desp acessorias,...)
    lErro = AtualizarDoc_ObterDadosCompl(objDoc, dValor, dBase, dBaseReduzida, dValorCredito)
    If lErro <> SUCESSO Then Error 27534
        
    'arredondar valores calculados
    dValor = Round(dValor, 2)
    dBase = Round(dBase, 2)
    dBaseReduzida = Round(dBaseReduzida, 2)
    dValorCredito = Round(dValorCredito, 2)

    'atualiza IPI_BASE, IPI_VALOR nao preenchidos manualmente
    lErro = objDoc.ObterVar("IPI_BASE", objVarIPIBase)
    If lErro <> SUCESSO Then Error 27535
    If objVarIPIBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIBase.vValor = dBaseReduzida
        objVarIPIBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIBase)
        If lErro <> SUCESSO Then Error 27536
    End If
    
    lErro = objDoc.ObterVar("IPI_VALOR", objVarIPIValor)
    If lErro <> SUCESSO Then Error 27536
    If objVarIPIValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIValor.vValor = dValor
        objVarIPIValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIValor)
        If lErro <> SUCESSO Then Error 27537
    End If
    
    lErro = objDoc.ObterVar("IPI_VALOR_CREDITO", objVarIPIValorCredito)
    If lErro <> SUCESSO Then Error 22785
    If objVarIPIValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIValorCredito.vValor = dValorCredito
        objVarIPIValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIValorCredito)
        If lErro <> SUCESSO Then Error 22786
    End If
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27531 To 27537, 22785, 22786, 22787
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154068)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_ObterValoresAtuais(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer) As Long
'obtem os valores atuais da variaveis associadas a um item
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_ObterValoresAtuais

    lErro = objDocItem.ObterVar("IPI_TIPO", objVarTipo)
    If lErro <> SUCESSO Then Error 27538
    If iRecalculaAutomatica = 1 Then objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("IPI_BASE", objVarBaseCalculo)
    If lErro <> SUCESSO Then Error 27539
    If iRecalculaAutomatica = 1 Then objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("IPI_RED_BASE", objVarPercRedBase)
    If lErro <> SUCESSO Then Error 27540
    If iRecalculaAutomatica = 1 Then objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objDocItem.ObterVar("IPI_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27541
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("IPI_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27542
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("IPI_VALOR_CREDITO", objVarValorCredito)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarItem_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarItem_ObterValoresAtuais:

    AtualizarItem_ObterValoresAtuais = Err

    Select Case Err

        Case 27538 To 27542, 22788
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154069)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_GuardarValores(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, dValor As Double, dValorCredito As Double) As Integer
'atualiza as variaveis que tiveram seus valores recalculados pelo sistema
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_GuardarValores

    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTipo.vValor = iTipo
        objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTipo)
        If lErro <> SUCESSO Then Error 27543
        
    End If
    
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseCalculo.vValor = dBaseCalculo
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarBaseCalculo)
        If lErro <> SUCESSO Then Error 27544
        
    End If
    
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarPercRedBase.vValor = dPercRedBase
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27545
        
    End If
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliquota.vValor = dAliquota
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then Error 27546
        
    End If
                
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then Error 27547
        
    End If
                
    If objVarValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorCredito.vValor = dValorCredito
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then Error 22783
        
    End If
                
    AtualizarItem_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores:

    AtualizarItem_GuardarValores = Err

    Select Case Err

        Case 27543 To 27547, 22783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154070)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarExcecao(objExcecao As ClassIPIExcecao, objTipoTribIPI As ClassTipoTribIPI, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarExcecao

    'se IPI_TIPO nao tiver sido preenchido manualmente
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objExcecao.iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribIPI.iTipo = iTipo
    lErro = TipoIPI_Obter(objTipoTribIPI)
    If lErro <> SUCESSO Then Error 27548
    
    'se IPI_ALIQUOTA nao tiver sido preenchido manualmente
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        'se IPI_TIPO permite aliquota <> 0
        If objTipoTribIPI.iPermiteAliquota <> 0 Then
            'IPI_ALIQUOTA = aliquota da excecao
            dAliquota = objExcecao.dAliquota
        Else
            dAliquota = 0
        End If
    Else
        dAliquota = objVarAliquota.vValor
    End If
    
    'se IPI_RED_BASE nao tiver sido preenchido manualmente
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        'IPI_RED_BASE = % da excecao
        dPercRedBase = objExcecao.dPercRedBaseCalculo
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
    
    AtualizarItem_TratarExcecao = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarExcecao:

    AtualizarItem_TratarExcecao = Err

    Select Case Err

        Case 27548
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154071)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecao(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribIPI As ClassTipoTribIPI, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double) As Long
Dim lErro As Long, objVarObjProduto As New ClassVariavelCalculo
Dim objProduto As New ClassProduto
On Error GoTo Erro_AtualizarItem_TratarNaoExcecao
        
    'se IPI_TIPO nao tiver sido preenchido manualmente...
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objTipoTrib.iIPITipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribIPI.iTipo = iTipo
    lErro = TipoIPI_Obter(objTipoTribIPI)
    If lErro <> SUCESSO Then Error 27549
    
    'se IPI_ALIQUOTA nao tiver sido preenchido manualmente
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        'se IPI_TIPO permite aliquota <> 0
        If objTipoTribIPI.iPermiteAliquota <> 0 Then
        
            lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
            If lErro <> SUCESSO Then Error 27597
        
            Set objProduto = objVarObjProduto.vValor
            dAliquota = objProduto.dIPIAliquota
        Else
            dAliquota = 0
        End If
    Else
        dAliquota = objVarAliquota.vValor
    End If
        
    'se IPI_RED_BASE nao tiver sido preenchido manualmente...
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
        If iTipo = IPI_TIPO_50PCTO Then
            dPercRedBase = 0.5
        Else
            dPercRedBase = 0
        End If
    Else
        dPercRedBase = objVarPercRedBase.vValor
    End If
        
    AtualizarItem_TratarNaoExcecao = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecao:

    AtualizarItem_TratarNaoExcecao = Err

    Select Case Err

        Case 27549, 27550, 27597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154072)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_Cli(objDocItem As ClassTributoDocItem, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'atualiza tipo IPI, base de calculo, red da base, aliquota, valor
Dim lErro As Long, objExcecao As ClassIPIExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, objVarValorCredito As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dValorCredito As Double
On Error GoTo Erro_AtualizarItem_Cli
    
    iTipo = IPI_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dValorCredito = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27551
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27552
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27553
    
    'se nao incide IPI sobre a operacao
    If objTipoTrib.iIPIIncide = TRIBUTO_NAO_INCIDE Then
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula IPI em funcao do IPI_TIPO setado manualmente
    'como IPI_TIPO_NAO_TRIBUTADO ou IPI_TIPO_ISENTO
    '??? conferir com Padua
    If (objVarTipo.iStatus = VAR_PREENCH_MANUAL And (objVarTipo.vValor = IPI_TIPO_NAO_TRIBUTADO Or objVarTipo.vValor = IPI_TIPO_ISENTO)) Then
        iTipo = objVarTipo.vValor
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    If iEntrada = 1 Then
        'tentar achar excecao
        lErro = Excecao_Pesquisar(objDocItem, objExcecao)
        If lErro <> SUCESSO Then Error 27554
    End If
    
    'se achou excecao
    If (Not (objExcecao Is Nothing)) Then
    
        lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase)
        If lErro <> SUCESSO Then Error 27555
            
    Else 'se NAO achou uma excecao
    
        lErro = AtualizarItem_TratarNaoExcecao(objDocItem, objTipoTrib, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase)
        If lErro <> SUCESSO Then Error 27556
    
    End If

    'se IPI_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO Then
            
            lErro = Item_ObtemBaseCalculo(objDocItem, mvardFatorValor, False, False, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27557
        
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If
    
    'se IPI_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.0000"))

    Else
        dValor = objVarValor.vValor
    End If

    If objTipoTrib.iIPICredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorCredito)
    If lErro <> SUCESSO Then Error 27558
    
    AtualizarItem_Cli = SUCESSO

    Exit Function

Erro_AtualizarItem_Cli:

    AtualizarItem_Cli = Err

    Select Case Err

        Case 27551 To 27558, 27699 To 27702, 27767, 59312
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154073)

    End Select

    Exit Function

End Function

Private Function Excecao_Pesquisar(objDocItem As ClassTributoDocItem, objExcecao As ClassIPIExcecao) As Long
'uma excecao a nivel de cliente+produto tem maior prioridade do que uma apenas p/cliente
'A colecao já está ordenada de forma que pararei na 1a excecao que encontrar
'retorna Nothing se nao achar nenhuma excecao adequada
Dim lErro As Long, objIPIExcecao As ClassIPIExcecao, iAchou As Integer
Dim objVarFilCliColCateg As New ClassVariavelCalculo
Dim objVarProdColCateg As New ClassVariavelCalculo
On Error GoTo Erro_Excecao_Pesquisar

    'obter a colecao de categorias da filialcliente
    lErro = objDocItem.objDoc.ObterVar("I_FILIAL_CLIENTE_COL_CATEG", objVarFilCliColCateg)
    If lErro <> SUCESSO Then Error 27559
    
    'obter a colecao de categorias do produto
    lErro = objDocItem.ObterVar("I_PRODUTO_COL_CATEG", objVarProdColCateg)
    If lErro <> SUCESSO Then Error 27560
    
    iAchou = 0
    
    'percorrer a colecao de excecoes
    For Each objIPIExcecao In mvarcolExcecoes
    
        'verifica se a filial cliente se enquadra
        lErro = CategoriaCliente_Pesquisar(objIPIExcecao.sCategoriaCliente, objIPIExcecao.sCategoriaClienteItem, objVarFilCliColCateg.vValor)
        If lErro <> SUCESSO And lErro <> 27563 Then Error 27561
            
        'se a excecao se aplica a uma (categoria de clientes para uma) categoria de produtos
        If objIPIExcecao.sCategoriaProduto <> "" Then
        
            'verifica se o produto se enquadra
            lErro = CategoriaProduto_Pesquisar(objIPIExcecao.sCategoriaProduto, objIPIExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
            If lErro <> SUCESSO And lErro <> 27564 Then Error 27562
                            
            If lErro = SUCESSO Then
                Set objExcecao = objIPIExcecao
                iAchou = 1
                Exit For
            End If
            
        Else 'se a excecao se aplica a uma categoria de clientes independente de produto
        
            If lErro = SUCESSO Then
                Set objExcecao = objIPIExcecao
                iAchou = 1
                Exit For
            End If
    
        End If
    
    Next
    
    If iAchou <> 1 Then Set objExcecao = Nothing
    
    Excecao_Pesquisar = SUCESSO

    Exit Function

Erro_Excecao_Pesquisar:

    Excecao_Pesquisar = Err

    Select Case Err
        
        Case 27559 To 27562

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154074)

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

    If iAchou = 0 Then Error 27563
    
    CategoriaCliente_Pesquisar = SUCESSO

    Exit Function

Erro_CategoriaCliente_Pesquisar:

    CategoriaCliente_Pesquisar = Err

    Select Case Err

        Case 27563 'nao encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154075)

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

    If iAchou = 0 Then Error 27564
    
    CategoriaProduto_Pesquisar = SUCESSO

    Exit Function

Erro_CategoriaProduto_Pesquisar:

    CategoriaProduto_Pesquisar = Err

    Select Case Err

        Case 27564 'nao encontrou
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154076)

    End Select

    Exit Function

End Function


Private Function TipoIPI_Obter(objTipoTribIPI As ClassTipoTribIPI) As Long
'Preenche objTipoTribIPI de acordo com o objTipoTribIPI.iTipo que vier preenchido
Dim lErro As Long, iAchou As Integer
Dim objTipoTribIPITemp As ClassTipoTribIPI
On Error GoTo Erro_TipoIPI_Obter

    iAchou = 0
    
    For Each objTipoTribIPITemp In mvarcolTiposTribIPI
        If objTipoTribIPI.iTipo = objTipoTribIPITemp.iTipo Then
            iAchou = 1
            Call objTipoTribIPI.Copiar(objTipoTribIPITemp)
            Exit For
        End If
    Next

    If iAchou = 0 Then Error 27565
    
    TipoIPI_Obter = SUCESSO

    Exit Function

Erro_TipoIPI_Obter:

    TipoIPI_Obter = Err

    Select Case Err

        Case 27565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_IPI_INEXISTENTE", Err, objTipoTribIPI.iTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154077)

    End Select

    Exit Function

End Function

Private Function TipoIPIPermiteAliq(iTipoIPI As Integer, iPermiteAliq As Integer) As Long
'retorna em iPermiteAliq se o tipo de IPI informado pode ter aliquota diferente de zero
Dim lErro As Long, iAchou As Integer
Dim objTipoTribIPI As ClassTipoTribIPI
On Error GoTo Erro_TipoIPIPermiteAliq

    iAchou = 0
    
    For Each objTipoTribIPI In mvarcolTiposTribIPI
        If objTipoTribIPI.iTipo = iTipoIPI Then
            iAchou = 1
            iPermiteAliq = objTipoTribIPI.iPermiteAliquota
            Exit For
        End If
    Next

    If iAchou = 0 Then Error 27566
    
    TipoIPIPermiteAliq = SUCESSO

    Exit Function

Erro_TipoIPIPermiteAliq:

    TipoIPIPermiteAliq = Err

    Select Case Err

        Case 27566
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_IPI_INEXISTENTE", Err, iTipoIPI)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154078)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_ObterDadosCompl(objDoc As ClassTributoDoc, dValor As Double, dBase As Double, dBaseReduzida As Double, dValorCredito As Double) As Long
'tratar dados globais (frete, desp acessorias,...)
Dim objTributoDocCompl As ClassTributoDocCompl, colDocCompl As New Collection, objVarItemTipo As New ClassVariavelCalculo
Dim objVarIPIBase As New ClassVariavelCalculo, objVarPercRedBase As New ClassVariavelCalculo
Dim objVarIPIValor As New ClassVariavelCalculo, objVarIPIValorCredito As New ClassVariavelCalculo
Dim lErro As Long, dPercRedBase As Double
On Error GoTo Erro_AtualizarDoc_ObterDadosCompl

    lErro = objDoc.ObterColDocCompl(colDocCompl)
    If lErro <> SUCESSO Then Error 27567

    For Each objTributoDocCompl In colDocCompl
    
        'obter do item IPI_BASE, IPI_VALOR
        
        lErro = objTributoDocCompl.ObterVar("IPI_BASE", objVarIPIBase)
        If lErro <> SUCESSO Then Error 27568
        
        lErro = objTributoDocCompl.ObterVar("IPI_RED_BASE", objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27569
        
        dPercRedBase = objVarPercRedBase.vValor
        
        lErro = objTributoDocCompl.ObterVar("IPI_VALOR", objVarIPIValor)
        If lErro <> SUCESSO Then Error 27570
        
        lErro = objTributoDocCompl.ObterVar("IPI_VALOR_CREDITO", objVarIPIValorCredito)
        If lErro <> SUCESSO Then Error 22803
        
        lErro = objTributoDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
        If lErro <> SUCESSO Then Error 32012
        
        If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
        
            dBase = dBase + objVarIPIBase.vValor
            dBaseReduzida = dBaseReduzida + (objVarIPIBase.vValor * (1 - dPercRedBase))
            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
            dValor = dValor + objVarIPIValor.vValor
            dValorCredito = dValorCredito + objVarIPIValorCredito.vValor
            
        Else
        
            dBase = dBase - objVarIPIBase.vValor
            dBaseReduzida = dBaseReduzida - (objVarIPIBase.vValor * (1 - dPercRedBase))
            ''dBaseReduzida = CDbl(Format(dBaseReduzida, "0.00"))
            dValor = dValor - objVarIPIValor.vValor
            dValorCredito = dValorCredito - objVarIPIValorCredito.vValor
        
        End If
        
    Next
    
    AtualizarDoc_ObterDadosCompl = SUCESSO

    Exit Function

Erro_AtualizarDoc_ObterDadosCompl:

    AtualizarDoc_ObterDadosCompl = Err

    Select Case Err

        Case 27567 To 27570, 22803, 32012
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154079)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_ObterValoresAtuais(objTributoDocCompl As ClassTributoDocCompl, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer) As Long
'obtem os valores atuais da variaveis associadas a um Compl
Dim lErro As Long
On Error GoTo Erro_AtualizarCompl_ObterValoresAtuais

    lErro = objTributoDocCompl.ObterVar("IPI_TIPO", objVarTipo)
    If lErro <> SUCESSO Then Error 27571
    If iRecalculaAutomatica = 1 Then objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objTributoDocCompl.ObterVar("IPI_BASE", objVarBaseCalculo)
    If lErro <> SUCESSO Then Error 27572
    If iRecalculaAutomatica = 1 Then objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objTributoDocCompl.ObterVar("IPI_RED_BASE", objVarPercRedBase)
    If lErro <> SUCESSO Then Error 27573
    If iRecalculaAutomatica = 1 Then objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO

    lErro = objTributoDocCompl.ObterVar("IPI_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27574
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("IPI_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27575
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objTributoDocCompl.ObterVar("IPI_VALOR_CREDITO", objVarValorCredito)
    If lErro <> SUCESSO Then Error 22804
    If iRecalculaAutomatica = 1 Then objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarCompl_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarCompl_ObterValoresAtuais:

    AtualizarCompl_ObterValoresAtuais = Err

    Select Case Err

        Case 27571 To 27575, 22804
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154080)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_GuardarValores(objTributoDocCompl As ClassTributoDocCompl, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, _
        dValor As Double, dValorCredito As Double) As Integer
'atualiza as variaveis que tiveram seus valores recalculados pelo sistema
Dim lErro As Long
On Error GoTo Erro_AtualizarCompl_GuardarValores

    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTipo.vValor = iTipo
        objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarTipo)
        If lErro <> SUCESSO Then Error 27576
        
    End If
    
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarBaseCalculo.vValor = dBaseCalculo
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarBaseCalculo)
        If lErro <> SUCESSO Then Error 27577
        
    End If
    
    If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarPercRedBase.vValor = dPercRedBase
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarPercRedBase)
        If lErro <> SUCESSO Then Error 27578
        
    End If
    
    If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarAliquota.vValor = dAliquota
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarAliquota)
        If lErro <> SUCESSO Then Error 27579
        
    End If
                
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then Error 27580
        
    End If
                
    If objVarValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorCredito.vValor = dValorCredito
        objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objTributoDocCompl.GuardarVar(objVarValorCredito)
        If lErro <> SUCESSO Then Error 22784
        
    End If
                
    AtualizarCompl_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarCompl_GuardarValores:

    AtualizarCompl_GuardarValores = Err

    Select Case Err

        Case 27576 To 27580, 22784
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154081)

    End Select

    Exit Function

End Function

Private Function ObterAliquotaCompl(objDoc As ClassTributoDoc, dAliquota As Double) As Long
'obtem a aliquota a ser utilizada p/frete, seguros, etc
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim dValorAcum As Double, dValorPonderadoAcum As Double, dAliquotaUnica As Double
Dim objVarValor As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim lErro As Long, dBaseCalculo As Double

On Error GoTo Erro_ObterAliquotaCompl

    Call objDoc.ObterColDocItens(colItens)
    
    dValorAcum = 0
    dValorPonderadoAcum = 0
    dAliquotaUnica = -1
    
    'percorrer os itens
    For Each objDocItem In colItens
        
        lErro = Item_ObtemBaseCalculo(objDocItem, mvardFatorValor, False, False, dBaseCalculo)
        If lErro <> SUCESSO Then Error 27601
            
        lErro = objDocItem.ObterVar("IPI_ALIQUOTA", objVarAliquota)
        If lErro <> SUCESSO Then Error 27602
        
        If dAliquotaUnica = -1 Then
            dAliquotaUnica = objVarAliquota.vValor
        Else
            If objVarAliquota.vValor <> dAliquotaUnica Then
                dAliquotaUnica = -2
            End If
        End If
        
        dValorPonderadoAcum = dValorPonderadoAcum + (objVarAliquota.vValor * dBaseCalculo)
        dValorAcum = dValorAcum + dBaseCalculo
        
    Next
    
    If dAliquotaUnica <> -1 And dAliquotaUnica <> -2 Then
        dAliquota = dAliquotaUnica
    Else
        If dValorAcum <> 0 Then
            dAliquota = dValorPonderadoAcum / dValorAcum
        Else
            dAliquota = 0
        End If
    End If
    
    ObterAliquotaCompl = SUCESSO

    Exit Function

Erro_ObterAliquotaCompl:

    ObterAliquotaCompl = Err

    Select Case Err

        Case 27601, 27602
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154082)

    End Select

    Exit Function

End Function

Private Function AtualizarDocCompl(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'atualiza tipo IPI, base de calculo, red da base, aliquota, valor
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo
Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, objVarTipoItem As New ClassVariavelCalculo, objVarValorCredito As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dValorCredito As Double
On Error GoTo Erro_AtualizarDocCompl
    
    Call objDoc.ObterColDocCompl(colDocCompl)
    
    'percorrer os complementos (frete, seguro,...) atualizando-os
    For Each objDocCompl In colDocCompl
    
        iTipo = IPI_TIPO_NAO_TRIBUTADO
        
        dBaseCalculo = 0
        dPercRedBase = 0
        dAliquota = 0
        dValor = 0
        dValorCredito = 0
        
        lErro = AtualizarCompl_ObterValoresAtuais(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27581
        
        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27582
        
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 27583
        
        'se nao incide IPI sobre a operacao
        If objTipoTrib.iIPIIncide = TRIBUTO_NAO_INCIDE Then
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        'se nao calcula IPI em funcao do IPI_TIPO setado manualmente
        'como IPI_TIPO_NAO_TRIBUTADO ou IPI_TIPO_ISENTO
        '??? rever c/Padua
        If (objVarTipo.iStatus = VAR_PREENCH_MANUAL And (objVarTipo.vValor = IPI_TIPO_NAO_TRIBUTADO Or objVarTipo.vValor = IPI_TIPO_ISENTO)) Then
            iTipo = objVarTipo.vValor
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        End If
            
        lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarTipoItem)
        If lErro <> SUCESSO Then Error 27599
        
        'se IPI_TIPO nao tiver sido preenchido manualmente...
        If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
            
            iTipo = objTipoTrib.iIPITipo
            objVarTipo.vValor = iTipo
            
            '??? rever c/Padua
            If (iTipo = IPI_TIPO_NAO_TRIBUTADO Or iTipo = IPI_TIPO_ISENTO) Then
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
            
            If objVarTipoItem.vValor = ITEM_TIPO_FRETE And objTipoTrib.iIPIFrete = 0 Then
                iTipo = IPI_TIPO_NAO_TRIBUTADO
                objVarTipo.vValor = iTipo
                'atualizar variaveis e sair
                GoTo Label_Atualizar
            End If
            
        Else
            iTipo = objVarTipo.vValor
        End If
        
        objTipoTribIPI.iTipo = iTipo
        lErro = TipoIPI_Obter(objTipoTribIPI)
        If lErro <> SUCESSO Then Error 27584
        
        'se IPI_RED_BASE nao tiver sido preenchido manualmente...
        If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
            dPercRedBase = 0
        Else
            dPercRedBase = objVarPercRedBase.vValor
        End If
            
        'se IPI_BASE nao preenchida manualmente
        If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        
            If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO And objVarTipoItem.vValor <> ITEM_TIPO_DESCONTO Then
                
                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVar)
                If lErro <> SUCESSO Then Error 27586
    
                objVar.vValor = Round(objVar.vValor * mvardFatorValor, 2)
                
                dBaseCalculo = objVar.vValor
            Else
                dBaseCalculo = 0
            End If
        Else
            dBaseCalculo = objVarBaseCalculo.vValor
        End If
        
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
        
            'obter aliquota proporcional aos produtos "normais"
            lErro = ObterAliquotaCompl(objDocCompl.objDoc, dAliquota)
            If lErro <> SUCESSO Then Error 27585
        
        Else
        
            dAliquota = objVarAliquota.vValor
            
        End If
        
        'se IPI_VALOR nao preenchido manualmente
        If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
            
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
            ''dValor = CDbl(Format(dValor, "0.00"))
            
        Else
            dValor = objVarValor.vValor
        End If
    
        If objTipoTrib.iIPICredita <> 0 Then
        
            dValorCredito = dValorCredito + dValor
        
        End If
        
Label_Atualizar:
        
        lErro = AtualizarCompl_GuardarValores(objDocCompl, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, _
                                    iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorCredito)
        If lErro <> SUCESSO Then Error 27587
    
    Next
    
    AtualizarDocCompl = SUCESSO

    Exit Function

Erro_AtualizarDocCompl:

    AtualizarDocCompl = Err

    Select Case Err

        Case 27581 To 27587, 27695 To 27698, 27768
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154083)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_Cli(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo
'recalcula a tributacao do doc como um todos e de seus itens
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_Cli

    If mvariContribuinte <> NAO_CONTRIBUINTE_IPI And mvariSimplesFederal = 0 Then
        
        lErro = PrepararAtualizacao_Cli(objDoc)
        If lErro <> SUCESSO Then Error 27518
        
        'se estiver recalculando marcar valores a nivel de doc c/preench auto.
        If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
    
        Call objDoc.ObterColDocItens(colItens)
        
        'percorrer os itens atualizando-os
        For Each objDocItem In colItens
            lErro = AtualizarItem_Cli(objDocItem, iEntrada, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27519
        Next
        
        'percorrer os complementos (frete, seguro,...) atualizando-os
        lErro = AtualizarDocCompl(objDoc, iEntrada, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27520
        
        'atualizar o docto como um todo
        lErro = AtualizarDoc(objDoc)
        If lErro <> SUCESSO Then Error 27521
        
    End If
    
    AtualizarImposto_Cli = SUCESSO

    Exit Function

Erro_AtualizarImposto_Cli:

    AtualizarImposto_Cli = Err

    Select Case Err

        Case 27518, 27519, 27520, 27521
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154084)

    End Select

    Exit Function

End Function

Private Function PrepararAtualizacao_Forn(objDoc As ClassTributoDoc) As Long
        
    PrepararAtualizacao_Forn = SUCESSO

End Function

Private Function AtualizarItem_Forn(objDocItem As ClassTributoDocItem, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'atualiza tipo IPI, base de calculo, red da base, aliquota, valor
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, objVarValorCredito As New ClassVariavelCalculo, objVarTipoTribDoc As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dValorCredito As Double, iTipoTrib As Integer
Dim objVarTipoDocSigla As New ClassVariavelCalculo, objVarProduto As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarItem_Forn
    
    iTipo = IPI_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dValorCredito = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 22789
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 22793
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 22799
    
    'se nao incide IPI sobre a operacao
    If objTipoTrib.iIPIIncide = TRIBUTO_NAO_INCIDE Then
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula IPI em funcao do IPI_TIPO setado manualmente
    'como IPI_TIPO_NAO_TRIBUTADO ou IPI_TIPO_ISENTO
    '??? conferir com Padua
    If (objVarTipo.iStatus = VAR_PREENCH_MANUAL And (objVarTipo.vValor = IPI_TIPO_NAO_TRIBUTADO Or objVarTipo.vValor = IPI_TIPO_ISENTO)) Then
        iTipo = objVarTipo.vValor
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
            
    lErro = AtualizarItem_TratarNaoExcecao(objDocItem, objTipoTrib, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase)
    If lErro <> SUCESSO Then Error 22800
    
    'se IPI_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO Then
            
            lErro = Item_ObtemBaseCalculo(objDocItem, mvardFatorValor, False, False, dBaseCalculo)
            If lErro <> SUCESSO Then Error 22801
            
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If
    
    'se IPI_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        dValor = CDbl(Format(dValor, "0.0000"))
    
    Else
        dValor = objVarValor.vValor
    End If

    If objTipoTrib.iIPICredita <> 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorCredito)
    If lErro <> SUCESSO Then Error 22802
    
    AtualizarItem_Forn = SUCESSO

    Exit Function

Erro_AtualizarItem_Forn:

    AtualizarItem_Forn = Err

    Select Case Err

        Case 22789 To 22802
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154085)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_Forn(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo
'recalcula a tributacao do doc como um todos e de seus itens
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long
On Error GoTo Erro_AtualizarImposto_Forn

    lErro = PrepararAtualizacao_Forn(objDoc)
    If lErro <> SUCESSO Then Error 22805
    
    'se estiver recalculando marcar valores a nivel de doc c/preench auto.
    If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
        lErro = AtualizarItem_Forn(objDocItem, iEntrada, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 22806
    Next
    
    'percorrer os complementos (frete, seguro,...) atualizando-os
    lErro = AtualizarDocCompl(objDoc, iEntrada, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 22807
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc)
    If lErro <> SUCESSO Then Error 22808
        
    AtualizarImposto_Forn = SUCESSO

    Exit Function

Erro_AtualizarImposto_Forn:

    AtualizarImposto_Forn = Err

    Select Case Err

        Case 22805 To 22808
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154086)

    End Select

    Exit Function

End Function


Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32023
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        mvariContribuinte = objFilialEmpresa.iContribuinteIPI
        mvariSimplesFederal = objFilialEmpresa.iSimplesFederal
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154087)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_Recalculo(objDoc As ClassTributoDoc) As Long
'marca os valores a nivel de doc como VAR_PREENCH_AUTOMATICO

Dim lErro As Long
Dim objVarIPIBase As New ClassVariavelCalculo, objVarIPIValor As New ClassVariavelCalculo
Dim objVarIPIValorCredito As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc_Recalculo

    lErro = objDoc.ObterVar("IPI_BASE", objVarIPIBase)
    If lErro <> SUCESSO Then gError 81630
    objVarIPIBase.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarIPIBase)
    If lErro <> SUCESSO Then gError 81638
    
    lErro = objDoc.ObterVar("IPI_VALOR", objVarIPIValor)
    If lErro <> SUCESSO Then gError 81631
    objVarIPIValor.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarIPIValor)
    If lErro <> SUCESSO Then gError 81639
    
    lErro = objDoc.ObterVar("IPI_VALOR_CREDITO", objVarIPIValorCredito)
    If lErro <> SUCESSO Then gError 81632
    objVarIPIValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarIPIValorCredito)
    If lErro <> SUCESSO Then gError 81640
    
    AtualizarDoc_Recalculo = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_Recalculo:

    AtualizarDoc_Recalculo = gErr
     
    Select Case gErr
                  
        Case 81630, 81631, 81632, 81638, 81639, 81640
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154088)
     
    End Select
     
    Exit Function

End Function



