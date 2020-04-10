VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoIPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvariContribuinte As Integer
'Private mvariSimplesFederal As Integer
Private mvariSimplesNacional As Integer
Private mvariNotaInterna As Integer

Private mvariZerarTributo As Integer
Private mvariZerarTribSimples As Integer

Private mvarcolExcecoes As New Collection 'colecao de excecoes IPI
            'já ordenada, primeiro a nivel de catcliente+catproduto,
            'depois a nivel de catcliente e por ultimo a nivel de catproduto
Private mvarcolTiposTribIPI As New Collection 'colecao com os tipos de tributacao p/IPI

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao do doc como um todos e de seus itens
Dim lErro As Long, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig
Dim dtDataRef As Date, dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo, objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32024
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 22777
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22778
    
    mvariZerarTributo = 0
    mvariZerarTribSimples = 0
    If mvariContribuinte = NAO_CONTRIBUINTE_IPI Or objTipoDocInfo.iModDocFisE = 35 Then
        mvariZerarTributo = 1
    End If
    
    If mvariSimplesNacional <> 0 Then
        
        lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
        If lErro <> SUCESSO Then Error 22778
        
        dtDataEmissao = objVarDataEmissao.vValor
        
        lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
        If lErro <> SUCESSO Then Error 22778
        
        dtDataEntrada = objVarDataEntrada.vValor
        
        Select Case objTipoDocInfo.iTipoTrib
            Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
                dtDataRef = dtDataEntrada
                
            Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
                dtDataRef = dtDataEmissao
        End Select
        
        If dtDataRef >= DATA_INICIO_SIMPLES_NACIONAL Then mvariZerarTribSimples = 1 '??? deveria comparar com a data de opcao pelo simples da empresa
        
    End If
        
    mvariNotaInterna = IIf(objTipoDocInfo.iEmitente = EMITENTE_EMPRESA, 1, 0)
    
    If (objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC) And (mvariSimplesNacional = 0 Or gobjCRFAT.iNFImportacaoTribFlag06 = MARCADO) Then mvariZerarTributo = 0
    
    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_ENTRADA_CLI
            lErro = AtualizarImposto_Cli(objDoc, 1, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22779

        Case TRIB_ENTRADA_FORN
            lErro = AtualizarImposto_Forn(objDoc, 1, iRecalculaAutomatica, objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC)
            If lErro <> SUCESSO Then Error 22780

        Case TRIB_SAIDA_CLI
            lErro = AtualizarImposto_Cli(objDoc, 0, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22781

        Case TRIB_SAIDA_FORN
            lErro = AtualizarImposto_Forn(objDoc, 0, iRecalculaAutomatica, False)
            If lErro <> SUCESSO Then Error 22782
        
    End Select
    
    lErro = AtualizarImposto_Dev(objDoc, 0, iRecalculaAutomatica, objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22781
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 22777 To 22782, 32024
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154041)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154042)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154043)

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
        
    'arredondar valores calculados
    dValor = Arredonda_Moeda(dValor, 2)
    dBase = Arredonda_Moeda(dBase, 2)
    dBaseReduzida = Arredonda_Moeda(dBaseReduzida, 2)
    dValorCredito = Arredonda_Moeda(dValorCredito, 2)

    'atualiza IPI_BASE, IPI_VALOR nao preenchidos manualmente
    lErro = objDoc.ObterVar("IPI_BASE", objVarIPIBase)
    If lErro <> SUCESSO Then Error 27535
    'If objVarIPIBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIBase.vValor = dBaseReduzida
        objVarIPIBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIBase)
        If lErro <> SUCESSO Then Error 27536
    'End If
    
    lErro = objDoc.ObterVar("IPI_VALOR", objVarIPIValor)
    If lErro <> SUCESSO Then Error 27536
    'If objVarIPIValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIValor.vValor = dValor
        objVarIPIValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIValor)
        If lErro <> SUCESSO Then Error 27537
    'End If
    
    lErro = objDoc.ObterVar("IPI_VALOR_CREDITO", objVarIPIValorCredito)
    If lErro <> SUCESSO Then Error 22785
    'If objVarIPIValorCredito.iStatus <> VAR_PREENCH_MANUAL Then
        objVarIPIValorCredito.vValor = dValorCredito
        objVarIPIValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIPIValorCredito)
        If lErro <> SUCESSO Then Error 22786
    'End If
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27531 To 27537, 22785, 22786, 22787
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154044)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_ObterValoresAtuais(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, iRecalculaAutomatica As Integer, objVarItemOutrosValores As ClassVariavelCalculo, objVarItemTipoCalculo As ClassVariavelCalculo, objVarItemValorUN As ClassVariavelCalculo, objVarItemQtdeUN As ClassVariavelCalculo, objVarCodEnq As ClassVariavelCalculo) As Long
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
            
    lErro = objDocItem.ObterVar("PRODUTO_OUTROS_VALORES", objVarItemOutrosValores)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarItemOutrosValores.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("IPI_TIPO_CALCULO", objVarItemTipoCalculo)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("IPI_UM_PADRAO_VALOR", objVarItemValorUN)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarItemValorUN.iStatus = VAR_PREENCH_AUTOMATICO
            
    lErro = objDocItem.ObterVar("IPI_UM_PADRAO_QTDE", objVarItemQtdeUN)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarItemQtdeUN.iStatus = VAR_PREENCH_AUTOMATICO
    
    If objVarItemTipoCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        objVarItemTipoCalculo.vValor = TRIB_TIPO_CALCULO_PERCENTUAL
    End If
            
    lErro = objDocItem.ObterVar("IPI_ENQUADRAMENTO_CODIGO", objVarCodEnq)
    If lErro <> SUCESSO Then Error 22788
    If iRecalculaAutomatica = 1 Then objVarCodEnq.iStatus = VAR_PREENCH_AUTOMATICO
            
    AtualizarItem_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarItem_ObterValoresAtuais:

    AtualizarItem_ObterValoresAtuais = Err

    Select Case Err

        Case 27538 To 27542, 22788
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154045)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_GuardarValores(objDocItem As ClassTributoDocItem, objVarTipo As ClassVariavelCalculo, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, objVarTipoCalc As ClassVariavelCalculo, objVarQtdeUN As ClassVariavelCalculo, objVarValorUN As ClassVariavelCalculo, objVarCodEnq As ClassVariavelCalculo, _
            iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double, dValor As Double, dValorCredito As Double, iTipoCalc As Integer, dQtdeUN As Double, dValorUN As Double, sCodEnq As String) As Integer
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
    
    If objVarTipoCalc.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTipoCalc.vValor = iTipoCalc
        objVarTipoCalc.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTipoCalc)
        If lErro <> SUCESSO Then Error 22783
        
    End If
    
    If objVarQtdeUN.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarQtdeUN.vValor = dQtdeUN
        objVarQtdeUN.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarQtdeUN)
        If lErro <> SUCESSO Then Error 22783
        
    End If
    
    If objVarValorUN.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorUN.vValor = dValorUN
        objVarValorUN.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarValorUN)
        If lErro <> SUCESSO Then Error 22783
        
    End If
                
    If objVarCodEnq.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarCodEnq.vValor = sCodEnq
        objVarCodEnq.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarCodEnq)
        If lErro <> SUCESSO Then Error 22783
        
    End If
    
    AtualizarItem_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores:

    AtualizarItem_GuardarValores = Err

    Select Case Err

        Case 27543 To 27547, 22783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154046)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarExcecao(objExcecao As ClassIPIExcecao, objTipoTribIPI As ClassTipoTribIPI, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, objVarItemTipoCalculo As ClassVariavelCalculo, iTipoCalc As Integer, objVarAliquotaRS As ClassVariavelCalculo, dAliquotaRS As Double) As Long
Dim lErro As Long
On Error GoTo Erro_AtualizarItem_TratarExcecao

    'se IPI_TIPO nao tiver sido preenchido manualmente
    If objVarTipo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipo = objExcecao.iTipo
        objVarTipo.vValor = iTipo
    Else
        iTipo = objVarTipo.vValor
    End If
    
    objTipoTribIPI.iTipo = iTipo
    lErro = TipoIPI_Obter(objTipoTribIPI)
    If lErro <> SUCESSO Then Error 27548
    
    If objVarItemTipoCalculo.iStatus <> VAR_PREENCH_MANUAL Then
        iTipoCalc = objExcecao.iTipoCalculo
    Else
        iTipoCalc = objVarItemTipoCalculo.vValor
    End If
       
    If objTipoTribIPI.iPermiteAliquota = MARCADO Then
    
        'se IPI_ALIQUOTA nao tiver sido preenchido manualmente
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            'IPI_ALIQUOTA = aliquota da excecao
            dAliquota = objExcecao.dAliquota
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        If objVarAliquotaRS.iStatus <> VAR_PREENCH_MANUAL Then
            'IPI_ALIQUOTA = aliquota da excecao
            dAliquotaRS = objExcecao.dAliquotaRS
        Else
            dAliquotaRS = objVarAliquotaRS.vValor
        End If
        
        If objTipoTribIPI.iPermiteReducaoBase = MARCADO Then
            
            'se IPI_RED_BASE nao tiver sido preenchido manualmente
            If objVarPercRedBase.iStatus <> VAR_PREENCH_MANUAL Then
                'IPI_RED_BASE = % da excecao
                dPercRedBase = objExcecao.dPercRedBaseCalculo
            Else
                dPercRedBase = objVarPercRedBase.vValor
            End If
        Else
            dPercRedBase = 0
        End If
        
    Else
        dAliquota = 0
        dAliquotaRS = 0
        dPercRedBase = 0
    End If
            
    'Mesmo que seja manual, se o tipotrib não permitir o tipo de calculo escolhido tem que fazer os ajustes
    Select Case objTipoTribIPI.iTipoCalculo
    
        Case TRIB_TIPO_CALCULO_PERMITE_NADA
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquota = 0
            dAliquotaRS = 0
            iTipoCalc = 0
        
        Case TRIB_TIPO_CALCULO_PERMITE_PERC
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquotaRS = 0
            iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL
        
        Case TRIB_TIPO_CALCULO_PERMITE_VALOR
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquota = 0
            iTipoCalc = TRIB_TIPO_CALCULO_VALOR
            
        Case TRIB_TIPO_CALCULO_PERMITE_AMBOS
        
            If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquota = 0
            Else
                objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquotaRS = 0
            End If
            
        Case TRIB_TIPO_CALCULO_PERMITE_AMBOS_MANUAL
        
            If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquota = 0
            Else
                objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquotaRS = 0
            End If
            
            'Zera se não tiver sido alterado manualmente
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarAliquotaRS.iStatus <> VAR_PREENCH_MANUAL Then dAliquotaRS = 0
            
    End Select
            
    AtualizarItem_TratarExcecao = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarExcecao:

    AtualizarItem_TratarExcecao = Err

    Select Case Err

        Case 27548
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154047)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_TratarNaoExcecao(objDocItem As ClassTributoDocItem, objTipoTrib As ClassTipoDeTributacaoMovto, objTipoTribIPI As ClassTipoTribIPI, objVarTipo As ClassVariavelCalculo, iTipo As Integer, objVarAliquota As ClassVariavelCalculo, dAliquota As Double, objVarPercRedBase As ClassVariavelCalculo, dPercRedBase As Double, objVarItemTipoCalculo As ClassVariavelCalculo, iTipoCalc As Integer, objVarAliquotaRS As ClassVariavelCalculo, dAliquotaRS As Double) As Long
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
    
    iTipoCalc = objVarItemTipoCalculo.vValor
    
    'se IPI_TIPO permite aliquota <> 0
    If objTipoTribIPI.iPermiteAliquota = MARCADO Then
    
        'se IPI_ALIQUOTA nao tiver sido preenchido manualmente
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            
            lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
            If lErro <> SUCESSO Then Error 27597
        
            Set objProduto = objVarObjProduto.vValor
            dAliquota = objProduto.dIPIAliquota
        Else
            dAliquota = objVarAliquota.vValor
        End If
        
        dAliquotaRS = objVarAliquotaRS.vValor
        
        If objTipoTribIPI.iPermiteReducaoBase = MARCADO Then
        
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
            
        Else
            dPercRedBase = 0
        End If
        
    Else
        dAliquota = 0
        dPercRedBase = 0
        dAliquotaRS = 0
    End If
    
    'Mesmo que seja manual, se o tipotrib não permitir o tipo de calculo escolhido tem que fazer os ajustes
    Select Case objTipoTribIPI.iTipoCalculo
    
        Case TRIB_TIPO_CALCULO_PERMITE_NADA
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquota = 0
            dAliquotaRS = 0
            iTipoCalc = 0
        
        Case TRIB_TIPO_CALCULO_PERMITE_PERC
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquotaRS = 0
            iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL
        
        Case TRIB_TIPO_CALCULO_PERMITE_VALOR
            objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
            dAliquota = 0
            iTipoCalc = TRIB_TIPO_CALCULO_VALOR
            
        Case TRIB_TIPO_CALCULO_PERMITE_AMBOS
        
            If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquota = 0
            Else
                objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquotaRS = 0
            End If
            
        Case TRIB_TIPO_CALCULO_PERMITE_AMBOS_MANUAL
        
            If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquota = 0
            Else
                objVarAliquotaRS.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquotaRS = 0
            End If
            
            'Zera se não tiver sido alterado manualmente
            If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
            If objVarAliquotaRS.iStatus <> VAR_PREENCH_MANUAL Then dAliquotaRS = 0
            
    End Select
        
    AtualizarItem_TratarNaoExcecao = SUCESSO

    Exit Function

Erro_AtualizarItem_TratarNaoExcecao:

    AtualizarItem_TratarNaoExcecao = Err

    Select Case Err

        Case 27549, 27550, 27597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154048)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_Cli(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'atualiza tipo IPI, base de calculo, red da base, aliquota, valor
Dim lErro As Long, objExcecao As ClassIPIExcecao

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, objVarValorCredito As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dValorCredito As Double
Dim objVarItemOutrosValores As New ClassVariavelCalculo, objVarItemTipoCalculo As New ClassVariavelCalculo
Dim objVarItemValorUN As New ClassVariavelCalculo, objVarItemQtdeUN As New ClassVariavelCalculo
Dim dQtdeUN As Double, dValorUN As Double, iTipoCalc As Integer, objVarItemQtdeTrib As New ClassVariavelCalculo

Dim objVarObjProduto As New ClassVariavelCalculo
Dim objProduto As New ClassProduto, bAliqZero As Boolean

Dim objVarCodEnq As New ClassVariavelCalculo, sCodEnq As String

On Error GoTo Erro_AtualizarItem_Cli
    
    iTipo = IPI_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dValorCredito = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, iRecalculaAutomatica, objVarItemOutrosValores, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN, objVarCodEnq)
    If lErro <> SUCESSO Then Error 27551
    
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 Then
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 27552
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 27553
    
    If objVarCodEnq.iStatus <> VAR_PREENCH_MANUAL Then
        sCodEnq = objTipoTrib.sIPICodEnq
    Else
        sCodEnq = objVarCodEnq.vValor
    End If
    
    'se nao incide IPI sobre a operacao
    If objTipoTrib.iIPIIncide = TRIBUTO_NAO_INCIDE Then
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula IPI em funcao do IPI_TIPO setado manualmente
    'como IPI_TIPO_NAO_TRIBUTADO ou IPI_TIPO_ISENTO
    '??? conferir com Padua
    If (objVarTipo.iStatus = VAR_PREENCH_MANUAL And (objVarTipo.vValor = IPI_TIPO_NAO_TRIBUTADO Or objVarTipo.vValor = IPI_TIPO_ISENTO)) Then
        iTipo = objVarTipo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
'    If iEntrada = 1 Then
        'tentar achar excecao
        lErro = Excecao_Pesquisar(objDocItem, objExcecao)
        If lErro <> SUCESSO Then Error 27554
'    End If
    
    'se achou excecao
    If (Not (objExcecao Is Nothing)) Then
    
        lErro = AtualizarItem_TratarExcecao(objExcecao, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarItemTipoCalculo, iTipoCalc, objVarItemValorUN, dValorUN)
        If lErro <> SUCESSO Then Error 27555
            
    Else 'se NAO achou uma excecao
    
        lErro = AtualizarItem_TratarNaoExcecao(objDocItem, objTipoTrib, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarItemTipoCalculo, iTipoCalc, objVarItemValorUN, dValorUN)
        If lErro <> SUCESSO Then Error 27556
    
    End If
        
    If objTipoTribIPI.iPermiteAliquota = DESMARCADO Then
        iTipo = objVarTipo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If

    If mvariNotaInterna <> 0 And mvariZerarTribSimples <> 0 Then
        
        iTipo = objVarTipo.vValor
        If iTipo = IPI_TIPO_NORMAL Then iTipo = IPI_TIPO_OUTROS
        
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    
    End If
    
    If gobjCRFAT.iTribIPINaoIncide <> 0 And mvariNotaInterna <> 0 And objVarTipo.iStatus <> VAR_PREENCH_MANUAL And objVarTipo.vValor = IPI_TIPO_NORMAL Then
    
        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
        If lErro <> SUCESSO Then Error 27597
    
        Set objProduto = objVarObjProduto.vValor
        
        If objProduto.iIPIIncide = 0 Then
            
            iTipo = objVarTipo.vValor
            
            Select Case gobjCRFAT.iTribIPINaoIncide
            
                Case 1
                    iTipo = IPI_TIPO_NAO_TRIBUTADO
                    objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
                
            End Select
            
            Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
            
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        
        End If
    
    End If
    
'    If objVarCodEnq.iStatus <> VAR_PREENCH_MANUAL Then
'        sCodEnq = objTipoTrib.sIPICodEnq
'    Else
'        sCodEnq = objVarCodEnq.vValor
'    End If
'
    'se IPI_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO Then
        
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "IPI", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 27557
            
        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If
    
    'iTipoCalc = objVarItemTipoCalculo.vValor
    
    bAliqZero = False
    If iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL Then
        dQtdeUN = 0
        dValorUN = 0
        objVarItemValorUN.iStatus = VAR_PREENCH_AUTOMATICO
        objVarItemQtdeUN.iStatus = VAR_PREENCH_AUTOMATICO
        If dAliquota = 0 Then bAliqZero = True
    Else
        dBaseCalculo = 0
        dAliquota = 0
        dPercRedBase = 0
        
        If objVarItemQtdeUN.iStatus = VAR_PREENCH_MANUAL Then
            dQtdeUN = objVarItemQtdeUN.vValor
        Else
            lErro = objDocItem.ObterVar("PRODUTO_QTDE_UM_TRIB", objVarItemQtdeTrib)
            If lErro <> SUCESSO Then Error 27557
            
            dQtdeUN = objVarItemQtdeTrib.vValor
        End If
            
        'dValorUN = objVarItemValorUN.vValor
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
        If dValorUN = 0 Then bAliqZero = True
    End If
    
    If bAliqZero And mvariNotaInterna <> 0 And objVarTipo.iStatus <> VAR_PREENCH_MANUAL And objVarTipo.vValor = IPI_TIPO_NORMAL Then
    
        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
        If lErro <> SUCESSO Then Error 27597
    
        Set objProduto = objVarObjProduto.vValor
        
        If objProduto.iIPIIncide <> 0 Then
        
            iTipo = IPI_TIPO_ALIQUOTA_ZERO
            objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO
            
            If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
            
                dBaseCalculo = 0
                objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
                
            End If
            
            Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
            
            'atualizar variaveis e sair
            GoTo Label_Atualizar
        
        End If
    
    End If
    
    'se IPI_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        If iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL Then
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        Else
            dValor = dQtdeUN * dValorUN
        End If
        
        'dValor = CDbl(Format(dValor, "0.0000"))
        dValor = Arredonda_Moeda(dValor)

    Else
        dValor = objVarValor.vValor
    End If

    If objTipoTrib.iIPICredita <> 0 And mvariZerarTributo = 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemQtdeUN, objVarItemValorUN, objVarCodEnq, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorCredito, iTipoCalc, dQtdeUN, dValorUN, sCodEnq)
    If lErro <> SUCESSO Then Error 27558
    
    AtualizarItem_Cli = SUCESSO

    Exit Function

Erro_AtualizarItem_Cli:

    AtualizarItem_Cli = Err

    Select Case Err

        Case 27551 To 27558, 27699 To 27702, 27767, 59312, 27597
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154049)

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
    
        'se a excecao se aplica a uma categoria de clientes para uma categoria de produtos
        If objIPIExcecao.sCategoriaCliente <> "" And objIPIExcecao.sCategoriaProduto <> "" Then
        
            'verifica se a filial cliente se enquadra
            lErro = CategoriaCliente_Pesquisar(objIPIExcecao.sCategoriaCliente, objIPIExcecao.sCategoriaClienteItem, objVarFilCliColCateg.vValor)
            If lErro <> SUCESSO And lErro <> 27563 Then Error 27561
            If lErro = SUCESSO Then
        
                'verifica se o produto se enquadra
                lErro = CategoriaProduto_Pesquisar(objIPIExcecao.sCategoriaProduto, objIPIExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                If lErro <> SUCESSO And lErro <> 27564 Then Error 27562
                                
                If lErro = SUCESSO Then
                    Set objExcecao = objIPIExcecao
                    iAchou = 1
                    Exit For
                End If
            
            End If
        
        Else
        
            'se a excecao se aplica a uma categoria de clientes
            If objIPIExcecao.sCategoriaCliente <> "" Then
                'verifica se a filial cliente se enquadra
            
                'verifica se a filial cliente se enquadra
                lErro = CategoriaCliente_Pesquisar(objIPIExcecao.sCategoriaCliente, objIPIExcecao.sCategoriaClienteItem, objVarFilCliColCateg.vValor)
                If lErro <> SUCESSO And lErro <> 27563 Then Error 27561
                If lErro = SUCESSO Then
                    Set objExcecao = objIPIExcecao
                    iAchou = 1
                    Exit For
                End If
            
            Else
            
                'verifica se o produto se enquadra
                lErro = CategoriaProduto_Pesquisar(objIPIExcecao.sCategoriaProduto, objIPIExcecao.sCategoriaProdutoItem, objVarProdColCateg.vValor)
                If lErro <> SUCESSO And lErro <> 27564 Then Error 27562
                                
                If lErro = SUCESSO Then
                    Set objExcecao = objIPIExcecao
                    iAchou = 1
                    Exit For
                End If
                
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154050)

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
    
        If UCase(objFilialCliCategoria.sCategoria) = UCase(sCategoriaCliente) And UCase(objFilialCliCategoria.sItem) = UCase(sCategoriaClienteItem) Then
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154051)

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
    
        If UCase(objProdutoCategoria.sCategoria) = UCase(sCategoriaProduto) And UCase(objProdutoCategoria.sItem) = UCase(sCategoriaProdutoItem) Then
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154052)

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
    
    If iAchou = 0 Then
        'Se não achou pega o Outras
        For Each objTipoTribIPITemp In mvarcolTiposTribIPI
            If 6 = objTipoTribIPITemp.iTipo Then
                iAchou = 1
                Call objTipoTribIPI.Copiar(objTipoTribIPITemp)
                Exit For
            End If
        Next
    End If

    If iAchou = 0 Then Error 27565
    
    TipoIPI_Obter = SUCESSO

    Exit Function

Erro_TipoIPI_Obter:

    TipoIPI_Obter = Err

    Select Case Err

        Case 27565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIB_IPI_INEXISTENTE", Err, objTipoTribIPI.iTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154053)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154054)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_Cli(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo
'recalcula a tributacao do doc como um todos e de seus itens
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long, bAtualiza As Boolean
On Error GoTo Erro_AtualizarImposto_Cli
       
    lErro = PrepararAtualizacao_Cli(objDoc)
    If lErro <> SUCESSO Then Error 27518
    
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
            lErro = AtualizarItem_Cli(objDoc, objDocItem, iEntrada, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27519
        End If
    Next
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc)
    If lErro <> SUCESSO Then Error 27521
    
    AtualizarImposto_Cli = SUCESSO

    Exit Function

Erro_AtualizarImposto_Cli:

    AtualizarImposto_Cli = Err

    Select Case Err

        Case 27518, 27519, 27520, 27521
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154060)

    End Select

    Exit Function

End Function

Private Function PrepararAtualizacao_Forn(objDoc As ClassTributoDoc) As Long
        
    PrepararAtualizacao_Forn = SUCESSO

End Function

Private Function AtualizarItem_Forn(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iEntrada As Integer, iRecalculaAutomatica As Integer, ByVal bNFImportacao As Boolean) As Long
'atualiza tipo IPI, base de calculo, red da base, aliquota, valor
Dim lErro As Long

Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTribIPI As New ClassTipoTribIPI
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVar As New ClassVariavelCalculo
Dim objVarTipo As New ClassVariavelCalculo, objVarBaseCalculo As New ClassVariavelCalculo
Dim objVarPercRedBase As New ClassVariavelCalculo, objVarAliquota As New ClassVariavelCalculo
Dim objVarValor As New ClassVariavelCalculo
Dim objVarValorDescGlobal As New ClassVariavelCalculo

Dim iTipo As Integer, dBaseCalculo As Double, dPercRedBase As Double, dAliquota As Double
Dim dValor As Double, objVarValorCredito As New ClassVariavelCalculo, objVarTipoTribDoc As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, dValorCredito As Double, iTipoTrib As Integer
Dim objVarTipoDocSigla As New ClassVariavelCalculo, objVarProduto As New ClassVariavelCalculo
Dim objVarItemOutrosValores As New ClassVariavelCalculo, objVarItemTipoCalculo As New ClassVariavelCalculo
Dim objVarItemValorUN As New ClassVariavelCalculo, objVarItemQtdeUN As New ClassVariavelCalculo
Dim dQtdeUN As Double, dValorUN As Double, iTipoCalc As Integer
Dim objVarAliquotaDI As New ClassVariavelCalculo, objVarValorUnitDI As New ClassVariavelCalculo
Dim objVarQtde As New ClassVariavelCalculo, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, objVarItemQtdeTrib As ClassVariavelCalculo
Dim objFilial As ClassFilialFornecedor, objVarFilFornObj As New ClassVariavelCalculo, objVarIIValor As New ClassVariavelCalculo

Dim objVarCodEnq As New ClassVariavelCalculo, sCodEnq As String

On Error GoTo Erro_AtualizarItem_Forn
    
    iTipo = IPI_TIPO_NAO_TRIBUTADO
    
    dBaseCalculo = 0
    dPercRedBase = 0
    dAliquota = 0
    dValor = 0
    dValorCredito = 0
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, iRecalculaAutomatica, objVarItemOutrosValores, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN, objVarCodEnq)
    If lErro <> SUCESSO Then Error 22789
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 27994
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 27994
        
    If mvariNotaInterna <> 0 And mvariZerarTributo <> 0 And objTipoDocInfo.sNomeTelaNFiscal <> "NFiscalDev" Then
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Then
    
        '??? deveria pegar ipi da nf original
    
    End If
    
    lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then Error 22793
    
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
    If lErro <> SUCESSO Then Error 22799
    
    If objVarCodEnq.iStatus <> VAR_PREENCH_MANUAL Then
        sCodEnq = objTipoTrib.sIPICodEnq
    Else
        sCodEnq = objVarCodEnq.vValor
    End If
    
    'se nao incide IPI sobre a operacao
    If objTipoTrib.iIPIIncide = TRIBUTO_NAO_INCIDE Then
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
        
    'se nao calcula IPI em funcao do IPI_TIPO setado manualmente
    'como IPI_TIPO_NAO_TRIBUTADO ou IPI_TIPO_ISENTO
    '??? conferir com Padua
    If (objVarTipo.iStatus = VAR_PREENCH_MANUAL And (objVarTipo.vValor = IPI_TIPO_NAO_TRIBUTADO Or objVarTipo.vValor = IPI_TIPO_ISENTO)) Then
        iTipo = objVarTipo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
            
    lErro = AtualizarItem_TratarNaoExcecao(objDocItem, objTipoTrib, objTipoTribIPI, objVarTipo, iTipo, objVarAliquota, dAliquota, objVarPercRedBase, dPercRedBase, objVarItemTipoCalculo, iTipoCalc, objVarItemValorUN, dValorUN)
    If lErro <> SUCESSO Then Error 22800
    
    If objTipoTribIPI.iPermiteAliquota = DESMARCADO Then
        iTipo = objVarTipo.vValor
        Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
        'atualizar variaveis e sair
        GoTo Label_Atualizar
    End If
    
    If mvariZerarTribSimples = 1 And bNFImportacao And gobjCRFAT.iNFImportacaoTribFlag06 = MARCADO Then mvariZerarTribSimples = 0
    
    If mvariNotaInterna <> 0 And mvariZerarTribSimples <> 0 Then
    
        lErro = objDoc.ObterVar("FILIAL_FORN_OBJ", objVarFilFornObj)
        If lErro <> SUCESSO Then Error 27994
            
        Set objFilial = objVarFilFornObj.vValor
    
        If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" And objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO And objFilial.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL Then
        
            iTipo = objVarTipo.vValor
            If iTipo = IPI_TIPO_NORMAL Then iTipo = IPI_TIPO_OUTROS
            
            '??? deveria obter aliquota e outras informacoes da nf de entrada
        
        Else
        
            If (objVarTipo.vValor <> IPI_TIPO_OUTROS Or objVarTipo.iStatus = VAR_PREENCH_AUTOMATICO) Then
            
                iTipo = objVarTipo.vValor
                If iTipo = IPI_TIPO_NORMAL Then iTipo = IPI_TIPO_OUTROS
                
                dAliquota = 0
                
                Call AtualizarItem_ZerarValores(objDocItem, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemValorUN, objVarItemQtdeUN)
                
                'atualizar variaveis e sair
                GoTo Label_Atualizar
        
            End If
    
        End If
    
    End If
    
    If bNFImportacao Then
        'forçar a aliquota com o que está na DI
        'se IPI_ALIQUOTA nao tiver sido preenchido manualmente
        If objVarAliquota.iStatus <> VAR_PREENCH_MANUAL Then
            'pegar da DI
            lErro = objDocItem.ObterVar("IPI_ALIQUOTA_ADICAO_DI", objVarAliquotaDI)
            If lErro <> SUCESSO Then Error 27994
            
            dAliquota = objVarAliquotaDI.vValor
        End If
        
        If objVarItemValorUN.iStatus <> VAR_PREENCH_MANUAL Then
            'pegar da DI
            lErro = objDocItem.ObterVar("IPI_VALORUNIT_ADICAO_DI", objVarValorUnitDI)
            If lErro <> SUCESSO Then Error 27994
            
            If objVarValorUnitDI.vValor <> 0 Then
                objVarItemValorUN.vValor = objVarValorUnitDI.vValor
                
                lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVarQtde)
                If lErro <> SUCESSO Then Error 27994
                objVarItemQtdeUN.vValor = objVarQtde.vValor
                
                objVarItemTipoCalculo.vValor = TRIB_TIPO_CALCULO_VALOR
            End If
            
        End If
        
    End If
    
'    If objVarCodEnq.iStatus <> VAR_PREENCH_MANUAL Then
'        sCodEnq = objTipoTrib.sIPICodEnq
'    Else
'        sCodEnq = objVarCodEnq.vValor
'    End If
    
    'se IPI_BASE nao preenchida manualmente
    If objVarBaseCalculo.iStatus <> VAR_PREENCH_MANUAL Then
    
        If objTipoTribIPI.iTipo <> IPI_TIPO_NAO_TRIBUTADO Then
            
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "IPI", objDoc, objDocItem, dBaseCalculo)
            If lErro <> SUCESSO Then Error 22801

            If gobjCRFAT.iNFImportacaoTribFlag12 <> 0 And bNFImportacao Then
            
                lErro = objDocItem.ObterVar("II_VALOR", objVarIIValor)
                If lErro <> SUCESSO Then Error 27994
                
                dBaseCalculo = Arredonda_Moeda(dBaseCalculo + objVarIIValor.vValor, 2)
                    
            End If

        Else
            dBaseCalculo = 0
        End If
    Else
        dBaseCalculo = objVarBaseCalculo.vValor
    End If
    
    'iTipoCalc = objVarItemTipoCalculo.vValor
    
    If iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL Then
        dQtdeUN = 0
        dValorUN = 0
        objVarItemValorUN.iStatus = VAR_PREENCH_AUTOMATICO
        objVarItemQtdeUN.iStatus = VAR_PREENCH_AUTOMATICO
    Else
        dBaseCalculo = 0
        dAliquota = 0
        dPercRedBase = 0
        
        If objVarItemQtdeUN.iStatus = VAR_PREENCH_MANUAL Then
            dQtdeUN = objVarItemQtdeUN.vValor
        Else
            lErro = objDocItem.ObterVar("PRODUTO_QTDE_UM_TRIB", objVarItemQtdeTrib)
            If lErro <> SUCESSO Then Error 22801
            
            dQtdeUN = objVarItemQtdeTrib.vValor
        End If
            
        'dValorUN = objVarItemValorUN.vValor
        objVarBaseCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
        objVarPercRedBase.iStatus = VAR_PREENCH_AUTOMATICO
    End If
    
    'se IPI_VALOR nao preenchido manualmente
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        
        If objVarItemTipoCalculo.vValor = TRIB_TIPO_CALCULO_PERCENTUAL Then
            dValor = dBaseCalculo * (1 - dPercRedBase) * dAliquota
        Else
            dValor = dQtdeUN * dValorUN
        End If
        
        'dValor = CDbl(Format(dValor, "0.0000"))
        dValor = Arredonda_Moeda(dValor)
    
    Else
        dValor = objVarValor.vValor
    End If

    If objTipoTrib.iIPICredita <> 0 And mvariZerarTributo = 0 Then
    
        dValorCredito = dValorCredito + dValor
        
    End If

Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTipo, objVarBaseCalculo, objVarPercRedBase, objVarAliquota, objVarValor, objVarValorCredito, objVarItemTipoCalculo, objVarItemQtdeUN, objVarItemValorUN, objVarCodEnq, _
                                iTipo, dBaseCalculo, dPercRedBase, dAliquota, dValor, dValorCredito, iTipoCalc, dQtdeUN, dValorUN, sCodEnq)
    If lErro <> SUCESSO Then Error 22802
    
    AtualizarItem_Forn = SUCESSO

    Exit Function

Erro_AtualizarItem_Forn:

    AtualizarItem_Forn = Err

    Select Case Err

        Case 22789 To 22802
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154061)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_Forn(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer, ByVal bNFImportacao As Boolean) As Long
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
        lErro = AtualizarItem_Forn(objDoc, objDocItem, iEntrada, iRecalculaAutomatica, bNFImportacao)
        If lErro <> SUCESSO Then Error 22806
    Next
    
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154062)

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
'        mvariSimplesFederal = objFilialEmpresa.iSimplesFederal
        mvariSimplesNacional = objFilialEmpresa.iSuperSimples
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154063)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154064)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarItem_ZerarValores(objDocItem As ClassTributoDocItem, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo, objVarItemTipoCalculo As ClassVariavelCalculo, objVarItemValorUN As ClassVariavelCalculo, objVarItemQtdeUN As ClassVariavelCalculo) As Integer
'zera valores calculados para o tributo

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
            
    objVarValorCredito.vValor = 0
    objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
                
    objVarItemValorUN.vValor = 0
    objVarItemValorUN.iStatus = VAR_PREENCH_AUTOMATICO
                
    objVarItemQtdeUN.vValor = 0
    objVarItemQtdeUN.iStatus = VAR_PREENCH_AUTOMATICO
                
    objVarItemTipoCalculo.vValor = TRIB_TIPO_CALCULO_PERCENTUAL
    objVarItemTipoCalculo.iStatus = VAR_PREENCH_AUTOMATICO
                
    AtualizarItem_ZerarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_ZerarValores:

    AtualizarItem_ZerarValores = Err

    Select Case Err

        Case 27543 To 27547, 22783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154046)

    End Select

    Exit Function

End Function

Private Function AtualizarCompl_ZerarValores(objDocCompl As ClassTributoDocCompl, objVarBaseCalculo As ClassVariavelCalculo, objVarPercRedBase As ClassVariavelCalculo, objVarAliquota As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarValorCredito As ClassVariavelCalculo) As Integer
'zera valores calculados para o tributo

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
            
    objVarValorCredito.vValor = 0
    objVarValorCredito.iStatus = VAR_PREENCH_AUTOMATICO
                
    AtualizarCompl_ZerarValores = SUCESSO

    Exit Function

Erro_AtualizarCompl_ZerarValores:

    AtualizarCompl_ZerarValores = Err

    Select Case Err

        Case 27543 To 27547, 22783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154046)

    End Select

    Exit Function

End Function

Private Function AtualizarImposto_Dev(objDoc As ClassTributoDoc, iEntrada As Integer, iRecalculaAutomatica As Integer, objTipoDocInfo As ClassTipoDocInfo) As Long

Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long, bAtualiza As Boolean, objItemNFOrig As ClassItemNF
Dim dPercDev As Double, dIPIDevol As Double, objVarItemNFOrig As New ClassVariavelCalculo, objVarQtde As New ClassVariavelCalculo
Dim objVarIPIVlrDevolvido As New ClassVariavelCalculo, objVarpDevol As New ClassVariavelCalculo
Dim bCalcIPIDevol As Double, objVarFinalidadeNFe As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarImposto_Dev

    Call objDoc.ObterColDocItens(colItens)
    
    lErro = objDoc.ObterVar("FINALIDADE_NFE", objVarFinalidadeNFe)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
  
    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
    
        dPercDev = 1
        dIPIDevol = 0
    
        bAtualiza = True
        If gobjCRFAT.iOtimizaTrib = MARCADO Then
            If objDocItem.iJaCalculado = MARCADO Then bAtualiza = False
        End If
        
        If bAtualiza Then
                                    
            lErro = objDocItem.ObterVar("IPI_VLR_DEVOLVIDO", objVarIPIVlrDevolvido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = objDocItem.ObterVar("PERCENTUAL_DEVOLVIDO", objVarpDevol)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
            If objVarFinalidadeNFe.vValor <> 4 Then
                'Se não é uma devolução apenas zera os dados
                dPercDev = 0
                dIPIDevol = 0
                bCalcIPIDevol = True
                objVarpDevol.iStatus = VAR_PREENCH_AUTOMATICO
                objVarIPIVlrDevolvido.iStatus = VAR_PREENCH_AUTOMATICO
            Else
            
                lErro = objDocItem.ObterVar("OBJ_ITEM_NF_ORIG", objVarItemNFOrig)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Set objItemNFOrig = objVarItemNFOrig.vValor
                
                If Not (objItemNFOrig Is Nothing) Then
                
                    lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVarQtde)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                    If objVarQtde.vValor > 0 Then
                        dPercDev = objVarQtde.vValor / objItemNFOrig.dQuantidade
                    End If
                    
                    dIPIDevol = objItemNFOrig.objTributacao.dIPIValor * dPercDev
                
                End If
                            
                bCalcIPIDevol = False
                Select Case gobjCRFAT.iCalcIPIDevolvido
                
                    Case 1 'Só para não contribuinte
                        If mvariContribuinte = NAO_CONTRIBUINTE_IPI Then bCalcIPIDevol = True
                    
                    Case 2 'Sempre
                        bCalcIPIDevol = True
                        
                    Case 0 'Nunca
                    
                End Select
            
            End If

        
            If objVarpDevol.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarpDevol.vValor = dPercDev
                objVarpDevol.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarpDevol)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            End If
            
            If bCalcIPIDevol Then
                If objVarIPIVlrDevolvido.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                    objVarIPIVlrDevolvido.vValor = dIPIDevol
                    objVarIPIVlrDevolvido.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarIPIVlrDevolvido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                End If
            End If
        
        
        End If
        
    Next
    
    AtualizarImposto_Dev = SUCESSO

    Exit Function

Erro_AtualizarImposto_Dev:

    AtualizarImposto_Dev = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154060)

    End Select

    Exit Function

End Function

