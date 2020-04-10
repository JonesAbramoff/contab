VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoDocAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'area de trabalho da calculadora de tributos que fica amarrada ao doc (pedido ou NF).
'Pode ser usada pelo Doc, diretamente pela calculadora ou pelos calculadores de imposto
'Ex. colecao de valores intermediarios,
'    colecao de mnemonicos alterados (p/cada tratador de impostos poder saber se a alteracao o afeta ou nao),
'    colecao de impostos com indicacao se calcula ou nao,
'    se já está calculado ou nao, indicacao de automatico ou nao.
'    colecao de valores "anteriores" p/poder saber se algo foi trocado p/decidir se recalcula o imposto
'
'ELEMENTOS:
'==========
'natop
'cliente
'filialCliente
'tipodoc
'tipotrib (default p/os itens) (???)
'
'estado destino
'pais destino
'colecao de categorias cliente c/seus valores
'se cliente é inscrito

Private mvarobjDoc As ClassTributoDoc
Private mvarobjVarAux As ClassVarAux

'Interface:
'==========

Function Ativar(objDoc As ClassTributoDoc) As Long

    'guarda objDoc
    Set mvarobjDoc = objDoc
    
    'cria objVarAux
    Set mvarobjVarAux = New ClassVarAux
    
End Function

Function Alterado() As Long
'marcar tendo ocorrido uma alteracao que impeça o aproveitamento do calculo anterior dos itens

Dim objDocItem As ClassTributoDocItem
Dim colDocItens As Collection

    Call mvarobjDoc.ObterColDocItens(colDocItens)
    
    For Each objDocItem In colDocItens
    
        Call objDocItem.Alterado
        
    Next

End Function

Function Desativar() As Long

    'apaga referencia a objDoc
    Set mvarobjDoc = Nothing
    
    'apaga referencia a objVarAux
    Set mvarobjVarAux = Nothing
    
End Function
    
Function ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'vai ser chamado pelo objDoc quando este nao conseguir "tratar"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long, colCateg As New Collection, objFilialCliente As New ClassFilialCliente, sUF As String
Dim objFilialForn As ClassFilialFornecedor, objFilialEmpresa As AdmFiliais
Dim objNatOpPadrao As New ClassNatOpPadrao, sCidadeIBGE As String
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo, bPadrao As Boolean
Dim lCliente As Long, iFilialCli As Integer, iAux As Integer, lComando As Long

On Error GoTo Erro_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    Select Case sIdentificador
        
        Case "FILIAL_FATURAMENTO"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            
            If giFilialEmpresa = 51 And left(UCase(gsNomeEmpresa), 5) = "FEMAR" Then
            
                bPadrao = True
            
                lErro = mvarobjDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
                If lErro <> SUCESSO Then Error 27380
                    
                If objVarCliente.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.vValor <> 0 Then
                
                    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
                    If lErro <> SUCESSO Then Error 27381
        
                    If objVarFilialCliente.iStatus <> VAR_PREENCH_VAZIO And objVarFilialCliente.vValor <> 0 Then
                    
                        lCliente = objVarCliente.vValor
                        iFilialCli = objVarFilialCliente.vValor
                        
                        lComando = Comando_Abrir()
                        If lComando = 0 Then Error 30000
                        
                        lErro = Comando_Executar(lComando, "SELECT Filial FROM FilialClienteCategorias WHERE Categoria = 'Humaita-AM' AND Item = 'Humaita-AM' AND Cliente = ? AND Filial = ?", iAux, lCliente, iFilialCli)
                        If lErro <> AD_SQL_SUCESSO Then Error 30001
                    
                        lErro = Comando_BuscarProximo(lComando)
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30002
                    
                        If lErro <> AD_SQL_SEM_DADOS Then
                            objVariavelCalculo.vValor = 52 'Humaitá
                            bPadrao = False
                        End If
                        
                        Call Comando_Fechar(lComando)
                    
                    End If
                    
                End If
        
                If bPadrao Then objVariavelCalculo.vValor = giFilialEmpresa
        
            Else
            
                objVariavelCalculo.vValor = giFilialEmpresa
                
            End If
        
        Case "FILIAL_FAT_OBJ"
            lErro = ObterObjFilialFat(objFilialEmpresa)
            If lErro <> SUCESSO Then Error 32013
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = objFilialEmpresa
        
        Case "I_FILIAL_FAT_CODIGO"
        
        
        Case "I_FILIAL_CLIENTE_COL_CATEG"
        
            lErro = ObterColCategFilCli(colCateg)
            If lErro <> SUCESSO Then Error 27356
            
            'objVariavelCalculo.sIdentificador = "I_FILIAL_CLIENTE_COL_CATEG"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = colCateg
            
        Case "I_FILIAL_FORN_COL_CATEG"
        
            lErro = ObterColCategFilForn(colCateg)
            If lErro <> SUCESSO Then Error 27356
            
            'objVariavelCalculo.sIdentificador = "I_FILIAL_FORN_COL_CATEG"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = colCateg
            
        Case "FILIAL_FORN_INSC_EST"
            lErro = ObterObjFilialForn(objFilialForn)
            If lErro <> SUCESSO Then Error 27841
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = objFilialForn.sInscricaoEstadual
            
        Case "FILIAL_FORN_IEISENTO"
            lErro = ObterObjFilialForn(objFilialForn)
            If lErro <> SUCESSO Then Error 27841
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            If Len(Trim(objFilialForn.sInscricaoEstadual)) = 0 Then
                objVariavelCalculo.vValor = MARCADO
            Else
                objVariavelCalculo.vValor = objFilialForn.iIENaoContrib 'era iiesento mas mudei em 29/02/2012...Jones
            End If
            
        Case "FORN_UF" '(EX p/exterior)
            lErro = ObterUFForn(sUF)
            If lErro <> SUCESSO Then Error 27842
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = sUF
        
        Case "FILIAL_FORN_OBJ"
            lErro = ObterObjFilialForn(objFilialForn)
            If lErro <> SUCESSO Then Error 27843
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = objFilialForn
        
        Case "FILIAL_CLIENTE_INSC_EST"
            lErro = ObterObjFilialCliente(objFilialCliente)
            If lErro <> SUCESSO Then Error 27394
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = objFilialCliente.sInscricaoEstadual
            
        Case "FILIAL_CLIENTE_IEISENTO"
            lErro = ObterObjFilialCliente(objFilialCliente)
            If lErro <> SUCESSO Then Error 27394
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'            If Len(Trim(objFilialCliente.sInscricaoEstadual)) = 0 Then
'                objVariavelCalculo.vValor = MARCADO
'            Else
                objVariavelCalculo.vValor = objFilialCliente.iIENaoContrib 'era iiesento mas mudei em 29/02/2012...Jones
'            End If
            
        Case "CLIENTE_UF" '(EX p/exterior)
            lErro = ObterUFCliente(sUF)
            If lErro <> SUCESSO Then Error 27395
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = sUF
        
        Case "FILIAL_CLIENTE_OBJ"
            lErro = ObterObjFilialCliente(objFilialCliente)
            If lErro <> SUCESSO Then Error 27396
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = objFilialCliente
        
        Case "NATOPPADRAO_OBJ"
            lErro = ObterObjNatOpPadrao(objNatOpPadrao)
            If lErro <> SUCESSO And lErro <> 56972 Then Error 56955
            
            If lErro = SUCESSO Then
            
                objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
                Set objVariavelCalculo.vValor = objNatOpPadrao
                
            Else
            
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
                Set objVariavelCalculo.vValor = Nothing
            
            End If
        
        Case "FILIAL_EMPRESA_FAT_CIDADE_IBGE" 'codigo do ibge da cidade da filialempresa de faturamento
            lErro = ObterCidadeIBGEFilialFat(sCidadeIBGE)
            If lErro <> SUCESSO Then Error 27396
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = sCidadeIBGE
            
        Case "FILIAL_CLIENTE_CIDADE_IBGE" 'codigo do ibge da cidade da filial do cliente
            lErro = ObterCidadeIBGECliente(sCidadeIBGE)
            If lErro <> SUCESSO Then Error 27396
            
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            objVariavelCalculo.vValor = sCidadeIBGE
            
        Case Else '???
            objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            objVariavelCalculo.vValor = 0 '???
            
    End Select

    ObterVar = SUCESSO

    Exit Function

Erro_ObterVar:

    ObterVar = Err

    Select Case Err

        Case 27356, 27394 To 27396, 27841, 27842, 27843, 32013, 56955
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153787)

    End Select

    Exit Function

End Function

Function GuardarVar(objVariavelCalculo As ClassVariavelCalculo)
'vai ser chamado pelo objDoc quando este nao conseguir "tratar"
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_GuardarVar
        
    lErro = mvarobjVarAux.GuardarVar(objVariavelCalculo)
    If lErro <> SUCESSO Then Error 27357
    
    GuardarVar = SUCESSO

    Exit Function

Erro_GuardarVar:

    GuardarVar = Err

    Select Case Err

        Case 27357
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153788)

    End Select

    Exit Function

End Function

Function ObterNaturezaOpPadrao(sNatOpPadrao As String) As Long
Dim lErro As Long
Dim objTRBConfig As New ClassTRBConfig
On Error GoTo Erro_ObterNaturezaOpPadrao

    lErro = objTRBConfig.gobjTributacao.ObterNaturezaOpPadrao(mvarobjDoc, sNatOpPadrao)
    If lErro <> SUCESSO Then Error 27404

    ObterNaturezaOpPadrao = SUCESSO

    Exit Function

Erro_ObterNaturezaOpPadrao:

    ObterNaturezaOpPadrao = Err

    Select Case Err

        Case 27404
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153789)

    End Select

    Exit Function

End Function
    
Function ObterTribPadrao(iTipoTrib As Integer) As Long
'retorna o tipo de tributacao padrao
Dim lErro As Long
Dim objTRBConfig As New ClassTRBConfig, objFilialEmpresa As AdmFiliais
Dim objTipoDocInfo As New ClassTipoDocInfo, objVarFilialFatObj As New ClassVariavelCalculo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, iTipoAtividadeEmp As Integer
Dim objVarNatOpPadraoObj As New ClassVariavelCalculo, objNatOpPadrao As ClassNatOpPadrao

On Error GoTo Erro_ObterTribPadrao

    lErro = mvarobjDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 27890
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 27891

    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_ENTRADA_CLI, TRIB_SAIDA_CLI
            lErro = ObterTribPadraoCli(iTipoTrib)
            If lErro <> SUCESSO Then Error 27892
            
        Case TRIB_ENTRADA_FORN, TRIB_SAIDA_FORN
            lErro = ObterTribPadraoForn(iTipoTrib)
            If lErro <> SUCESSO Then Error 27893
            
    End Select
    
    If iTipoTrib = 0 Then
    
        lErro = mvarobjDoc.ObterVar("NATOPPADRAO_OBJ", objVarNatOpPadraoObj)
        If lErro <> SUCESSO Then Error 56936
        
        Set objNatOpPadrao = objVarNatOpPadraoObj.vValor
        If Not (objNatOpPadrao Is Nothing) Then iTipoTrib = objNatOpPadrao.iTipoTribEmp
    
    End If
    
    ObterTribPadrao = SUCESSO

    Exit Function

Erro_ObterTribPadrao:

    ObterTribPadrao = Err

    Select Case Err

        Case 56936, 56937
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153790)

    End Select

    Exit Function

End Function

Function ObterTribPadraoCli(iTipoTrib As Integer) As Long
'obter o codigo do tipo de tributacao default para o doc
Dim lErro As Long
Dim objVarTipoTrib As New ClassVariavelCalculo
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarClienteAtual As New ClassVariavelCalculo, objVarFilialClienteAtual As New ClassVariavelCalculo
Dim objVarNatOp As New ClassVariavelCalculo, objVarTipoDoc As New ClassVariavelCalculo
Dim objVarNatOpAtual As New ClassVariavelCalculo, objVarTipoDocAtual As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig

On Error GoTo Erro_ObterTribPadraoCli

    lErro = mvarobjDoc.ObterVar("TIPO_DOC_SIGLA", objVarTipoDoc)
    If lErro <> SUCESSO Then Error 27341
        
    lErro = mvarobjDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
    If lErro <> SUCESSO Then Error 27342
        
    lErro = mvarobjDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
    If lErro <> SUCESSO Then Error 27343
        
    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
    If lErro <> SUCESSO Then Error 27344

    lErro = mvarobjVarAux.ObterVar("I_OTP_TIPO_DOC_INFO", objVarTipoDocAtual)
    If lErro <> SUCESSO Then Error 27345
        
    lErro = mvarobjVarAux.ObterVar("I_OTP_NAT_OP", objVarNatOpAtual)
    If lErro <> SUCESSO Then Error 27346

    lErro = mvarobjVarAux.ObterVar("I_OTP_CLIENTE_CODIGO", objVarClienteAtual)
    If lErro <> SUCESSO Then Error 27347
        
    lErro = mvarobjVarAux.ObterVar("I_OTP_FILIAL_CLIENTE_CODIGO", objVarFilialClienteAtual)
    If lErro <> SUCESSO Then Error 27348
    
    'se nao houve modificacao
    If objVarClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilialClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarTipoDocAtual.iStatus <> VAR_PREENCH_VAZIO And objVarNatOpAtual.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.vValor = objVarClienteAtual.vValor And objVarFilialCliente.vValor = objVarFilialClienteAtual.vValor And _
            objVarTipoDocAtual.vValor = objVarTipoDocAtual.vValor And objVarNatOp.vValor = objVarNatOpAtual.vValor Then

        'buscar o resultado obtido anteriormente
        lErro = mvarobjVarAux.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27349
            
        iTipoTrib = objVarTipoTrib.vValor
            
    Else

        'chamar funcao em Tributacao
        lErro = CF("ObterTribPadraoCli", iTipoTrib, objVarTipoDoc.vValor, objVarNatOp.vValor, objVarCliente.vValor, objVarFilialCliente.vValor)
        If lErro <> SUCESSO Then Error 27350
        
        'atualizar "TIPO_TRIB"
        objVarTipoTrib.vValor = iTipoTrib
        objVarTipoTrib.sIdentificador = "TIPO_TRIB"
        objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27351
    
        'atualizar "I_OTP_TIPO_DOC_INFO"
        objVarTipoDocAtual.vValor = objVarTipoDoc.vValor
        objVarTipoDocAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoDocAtual)
        If lErro <> SUCESSO Then Error 27352
    
        'atualizar "I_OTP_NAT_OP"
        objVarNatOpAtual.vValor = objVarNatOp.vValor
        objVarNatOpAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarNatOpAtual)
        If lErro <> SUCESSO Then Error 27353
        
        'atualizar "I_OTP_CLIENTE_CODIGO"
        objVarClienteAtual.vValor = objVarCliente.vValor
        objVarClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarClienteAtual)
        If lErro <> SUCESSO Then Error 27354
    
        'atualizar "I_OTP_CLIENTE_FILIAL"
        objVarFilialClienteAtual.vValor = objVarFilialCliente.vValor
        objVarFilialClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialClienteAtual)
        If lErro <> SUCESSO Then Error 27355
    
    End If
    
    ObterTribPadraoCli = SUCESSO

    Exit Function

Erro_ObterTribPadraoCli:

    ObterTribPadraoCli = Err

    Select Case Err

        Case 27341 To 27355
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153791)

    End Select

    Exit Function

End Function

Function ObterColCategFilCli(colCateg As Collection) As Long
'faz com que colCateg aponte p/uma colecao com as (categoria,item) associadas a filial cliente associada ao doc
Dim lErro As Long, objFilialCliente As New ClassFilialCliente
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarClienteAtual As New ClassVariavelCalculo, objVarFilialClienteAtual As New ClassVariavelCalculo
Dim objVarColCateg As New ClassVariavelCalculo
On Error GoTo Erro_ObterColCategFilCli

    lErro = mvarobjDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
    If lErro <> SUCESSO Then Error 27332
        
    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
    If lErro <> SUCESSO Then Error 27333
        
    lErro = mvarobjVarAux.ObterVar("I_CCC_CLIENTE_CODIGO", objVarClienteAtual)
    If lErro <> SUCESSO Then Error 27334
        
    lErro = mvarobjVarAux.ObterVar("I_CCC_FILIAL_CLIENTE_CODIGO", objVarFilialClienteAtual)
    If lErro <> SUCESSO Then Error 27335

    'se nao houve modificacao
    If objVarClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilialClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.vValor = objVarClienteAtual.vValor And objVarCliente.vValor = objVarClienteAtual.vValor Then
        
        lErro = mvarobjVarAux.ObterVar("I_FILIAL_CLIENTE_COL_CATEG", objVarColCateg)
        If lErro <> SUCESSO Then Error 27336
        
        Set colCateg = objVarColCateg.vValor
    
    Else
    
        'fazer a leitura
        objFilialCliente.lCodCliente = objVarCliente.vValor
        objFilialCliente.iCodFilial = objVarFilialCliente.vValor
        lErro = CF("ClienteCategorias_Le_FilialCliente", objFilialCliente, colCateg)
        If lErro <> SUCESSO Then Error 27337

        Set objVarColCateg.vValor = colCateg
        objVarColCateg.sIdentificador = "I_FILIAL_CLIENTE_COL_CATEG"
        objVarColCateg.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "I_FILIAL_CLIENTE_COL_CATEG"
        lErro = mvarobjVarAux.GuardarVar(objVarColCateg)
        If lErro <> SUCESSO Then Error 27338
        
        'atualizar "I_CCC_CLIENTE_CODIGO"
        objVarClienteAtual.vValor = objVarCliente.vValor
        objVarClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarClienteAtual)
        If lErro <> SUCESSO Then Error 27339
    
        'atualizar "I_CCC_CLIENTE_FILIAL"
        objVarFilialClienteAtual.vValor = objVarFilialCliente.vValor
        objVarFilialClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialClienteAtual)
        If lErro <> SUCESSO Then Error 27340
    
    End If
    
    ObterColCategFilCli = SUCESSO

    Exit Function

Erro_ObterColCategFilCli:

    ObterColCategFilCli = Err

    Select Case Err

        Case 27332 To 27340
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153792)

    End Select

    Exit Function

End Function

Function ObterObjFilialCliente(objFilialCliente As ClassFilialCliente) As Long
Dim lErro As Long
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarClienteAtual As New ClassVariavelCalculo, objVarFilialClienteAtual As New ClassVariavelCalculo
Dim objVarFilialClienteObj As New ClassVariavelCalculo
On Error GoTo Erro_ObterObjFilialCliente

    lErro = mvarobjDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
    If lErro <> SUCESSO Then Error 27380
        
    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
    If lErro <> SUCESSO Then Error 27381
        
    lErro = mvarobjVarAux.ObterVar("I_OFC_CLIENTE_CODIGO", objVarClienteAtual)
    If lErro <> SUCESSO Then Error 27382
        
    lErro = mvarobjVarAux.ObterVar("I_OFC_FILIAL_CLIENTE_CODIGO", objVarFilialClienteAtual)
    If lErro <> SUCESSO Then Error 27383

    'se nao houve modificacao
    If objVarClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilialClienteAtual.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.vValor = objVarClienteAtual.vValor And objVarFilialCliente.vValor = objVarFilialClienteAtual.vValor Then
        
        lErro = mvarobjVarAux.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilialClienteObj)
        If lErro <> SUCESSO Then Error 27384
        
        Set objFilialCliente = objVarFilialClienteObj.vValor
        
    Else
    
        Set objFilialCliente = New ClassFilialCliente
        
        'fazer a leitura
        objFilialCliente.lCodCliente = objVarCliente.vValor
        objFilialCliente.iCodFilial = objVarFilialCliente.vValor
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO Then Error 27385

        Set objVarFilialClienteObj.vValor = objFilialCliente
        objVarFilialClienteObj.sIdentificador = "FILIAL_CLIENTE_OBJ"
        objVarFilialClienteObj.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "FILIAL_CLIENTE_OBJ"
        lErro = mvarobjVarAux.GuardarVar(objVarFilialClienteObj)
        If lErro <> SUCESSO Then Error 27386
        
        'atualizar "I_OFC_CLIENTE_CODIGO"
        objVarClienteAtual.vValor = objVarCliente.vValor
        objVarClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarClienteAtual)
        If lErro <> SUCESSO Then Error 27387
    
        'atualizar "I_OFC_FILIAL_CLIENTE_CODIGO"
        objVarFilialClienteAtual.vValor = objVarFilialCliente.vValor
        objVarFilialClienteAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialClienteAtual)
        If lErro <> SUCESSO Then Error 27388
    
    End If
    
    ObterObjFilialCliente = SUCESSO

    Exit Function

Erro_ObterObjFilialCliente:

    ObterObjFilialCliente = Err

    Select Case Err

        Case 27380 To 27388
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153793)

    End Select

    Exit Function

End Function

Function ObterUFCliente(sUF As String) As Long
Dim lErro As Long
Dim objVarFilCliObj As New ClassVariavelCalculo, objVarClienteEndAtual As New ClassVariavelCalculo
Dim objVarEndereco As New ClassVariavelCalculo
Dim objFilialCliente As ClassFilialCliente, objEndereco As New ClassEndereco
On Error GoTo Erro_ObterUFCliente

    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilCliObj)
    If lErro <> SUCESSO Then Error 27393
                    
    lErro = mvarobjVarAux.ObterVar("I_OUC_CLIENTE_ENDERECO", objVarClienteEndAtual)
    If lErro <> SUCESSO Then Error 27389
        
    Set objFilialCliente = objVarFilCliObj.vValor
    
    'se nao houve modificacao
    If objVarClienteEndAtual.iStatus <> VAR_PREENCH_VAZIO And objVarClienteEndAtual.vValor = objFilialCliente.lEndereco Then
        
        lErro = mvarobjVarAux.ObterVar("CLIENTE_ENDERECO_OBJ", objVarEndereco)
        If lErro <> SUCESSO Then Error 27390
    
        Set objEndereco = objVarEndereco.vValor
        
    Else
    
        objEndereco.lCodigo = objFilialCliente.lEndereco
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO Then Error 27391
        
        Set objVarEndereco.vValor = objEndereco
        objVarEndereco.sIdentificador = "CLIENTE_ENDERECO_OBJ"
        objVarEndereco.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarEndereco)
        If lErro <> SUCESSO Then Error 27778
        
        'atualizar "I_OUC_CLIENTE_ENDERECO"
        objVarClienteEndAtual.vValor = objFilialCliente.lEndereco
        objVarClienteEndAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarClienteEndAtual)
        If lErro <> SUCESSO Then Error 27392
    
    End If
    
    If objEndereco.iCodigoPais <> PAIS_BRASIL And objEndereco.iCodigoPais <> 0 Then
        sUF = "EX"
    Else
        sUF = objEndereco.sSiglaEstado
    End If
    
    ObterUFCliente = SUCESSO

    Exit Function

Erro_ObterUFCliente:

    ObterUFCliente = Err

    Select Case Err

        Case 27389 To 27393, 27778
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153794)

    End Select

    Exit Function

End Function

Function ObterObjFilialForn(objFilialForn As ClassFilialFornecedor) As Long
Dim lErro As Long
Dim objVarForn As New ClassVariavelCalculo, objVarFilialForn As New ClassVariavelCalculo
Dim objVarFornAtual As New ClassVariavelCalculo, objVarFilialFornAtual As New ClassVariavelCalculo
Dim objVarFilialFornObj As New ClassVariavelCalculo
On Error GoTo Erro_ObterObjFilialForn

    lErro = mvarobjDoc.ObterVar("FORN_CODIGO", objVarForn)
    If lErro <> SUCESSO Then Error 27844
        
    lErro = mvarobjDoc.ObterVar("FILIAL_FORN_CODIGO", objVarFilialForn)
    If lErro <> SUCESSO Then Error 27845
        
    lErro = mvarobjVarAux.ObterVar("I_OFF_FORN_CODIGO", objVarFornAtual)
    If lErro <> SUCESSO Then Error 27846
        
    lErro = mvarobjVarAux.ObterVar("I_OFF_FILIAL_FORN_CODIGO", objVarFilialFornAtual)
    If lErro <> SUCESSO Then Error 27847

    'se nao houve modificacao
    If objVarFornAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilialFornAtual.iStatus <> VAR_PREENCH_VAZIO And objVarForn.vValor = objVarFornAtual.vValor And objVarFilialForn.vValor = objVarFilialFornAtual.vValor Then
        
        lErro = mvarobjVarAux.ObterVar("FILIAL_FORN_OBJ", objVarFilialFornObj)
        If lErro <> SUCESSO Then Error 27848
        
        Set objFilialForn = objVarFilialFornObj.vValor
        
    Else
    
        Set objFilialForn = New ClassFilialFornecedor
        
        'fazer a leitura
        objFilialForn.lCodFornecedor = objVarForn.vValor
        objFilialForn.iCodFilial = objVarFilialForn.vValor
        lErro = CF("FilialFornecedor_Le", objFilialForn)
        If lErro <> SUCESSO Then Error 27849

        Set objVarFilialFornObj.vValor = objFilialForn
        objVarFilialFornObj.sIdentificador = "FILIAL_FORN_OBJ"
        objVarFilialFornObj.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "FILIAL_FORN_OBJ"
        lErro = mvarobjVarAux.GuardarVar(objVarFilialFornObj)
        If lErro <> SUCESSO Then Error 27850
        
        'atualizar "I_OFF_FORN_CODIGO"
        objVarFornAtual.vValor = objVarForn.vValor
        objVarFornAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFornAtual)
        If lErro <> SUCESSO Then Error 27851
     
        'atualizar "I_OFF_FILIAL_FORN_CODIGO"
        objVarFilialFornAtual.vValor = objVarFilialForn.vValor
        objVarFilialFornAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialFornAtual)
        If lErro <> SUCESSO Then Error 27852
    
    End If
    
    ObterObjFilialForn = SUCESSO

    Exit Function

Erro_ObterObjFilialForn:

    ObterObjFilialForn = Err

    Select Case Err

        Case 27844 To 27852
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153795)

    End Select

    Exit Function

End Function

Function ObterUFForn(sUF As String) As Long
Dim lErro As Long
Dim objVarFilFornObj As New ClassVariavelCalculo, objVarFornEndAtual As New ClassVariavelCalculo
Dim objVarEndereco As New ClassVariavelCalculo
Dim objFilialForn As ClassFilialFornecedor, objEndereco As New ClassEndereco
On Error GoTo Erro_ObterUFForn

    lErro = mvarobjDoc.ObterVar("FILIAL_FORN_OBJ", objVarFilFornObj)
    If lErro <> SUCESSO Then Error 27853
                    
    lErro = mvarobjVarAux.ObterVar("I_OUF_ENDERECO", objVarFornEndAtual)
    If lErro <> SUCESSO Then Error 27854
        
    Set objFilialForn = objVarFilFornObj.vValor
    
    'se nao houve modificacao
    If objVarFornEndAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFornEndAtual.vValor = objFilialForn.lEndereco Then
        
        lErro = mvarobjVarAux.ObterVar("FORN_ENDERECO_OBJ", objVarEndereco)
        If lErro <> SUCESSO Then Error 27855
    
        Set objEndereco = objVarEndereco.vValor
        
    Else
    
        objEndereco.lCodigo = objFilialForn.lEndereco
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO Then Error 27856
        
        Set objVarEndereco.vValor = objEndereco
        objVarEndereco.sIdentificador = "FORN_ENDERECO_OBJ"
        objVarEndereco.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarEndereco)
        If lErro <> SUCESSO Then Error 27857
        
        'atualizar "I_OUF_ENDERECO"
        objVarFornEndAtual.vValor = objFilialForn.lEndereco
        objVarFornEndAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFornEndAtual)
        If lErro <> SUCESSO Then Error 27858
    
    End If
    
    If objEndereco.iCodigoPais <> PAIS_BRASIL And objEndereco.iCodigoPais <> 0 Then
        sUF = "EX"
    Else
        sUF = objEndereco.sSiglaEstado
    End If
    
    ObterUFForn = SUCESSO

    Exit Function

Erro_ObterUFForn:

    ObterUFForn = Err

    Select Case Err

        Case 27853 To 27858
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153796)

    End Select

    Exit Function

End Function

Private Function ObterTribPadraoForn(iTipoTrib As Integer) As Long
'obter o codigo do tipo de tributacao default para o doc

Dim lErro As Long
Dim objVarTipoTrib As New ClassVariavelCalculo
Dim objVarNatOp As New ClassVariavelCalculo, objVarTipoDoc As New ClassVariavelCalculo
Dim objVarNatOpAtual As New ClassVariavelCalculo, objVarTipoDocAtual As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig

On Error GoTo Erro_ObterTribPadraoForn

    lErro = mvarobjDoc.ObterVar("TIPO_DOC_SIGLA", objVarTipoDoc)
    If lErro <> SUCESSO Then Error 27894
        
    lErro = mvarobjDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
    If lErro <> SUCESSO Then Error 27895
        
    lErro = mvarobjVarAux.ObterVar("I_OTP_TIPO_DOC_INFO", objVarTipoDocAtual)
    If lErro <> SUCESSO Then Error 27896
        
    lErro = mvarobjVarAux.ObterVar("I_OTP_NAT_OP", objVarNatOpAtual)
    If lErro <> SUCESSO Then Error 27897

    'se nao houve modificacao
    If objVarTipoDocAtual.iStatus <> VAR_PREENCH_VAZIO And objVarNatOpAtual.iStatus <> VAR_PREENCH_VAZIO And _
            objVarTipoDocAtual.vValor = objVarTipoDocAtual.vValor And objVarNatOp.vValor = objVarNatOpAtual.vValor Then

        'buscar o resultado obtido anteriormente
        lErro = mvarobjVarAux.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27898
            
        iTipoTrib = objVarTipoTrib.vValor
            
    Else

        'chamar funcao em Tributacao passando "" como produto
        lErro = CF("ObterTribPadraoProd", iTipoTrib, objVarTipoDoc.vValor, objVarNatOp.vValor, "")
        If lErro <> SUCESSO Then Error 27899
        
        'atualizar "TIPO_TRIB"
        objVarTipoTrib.vValor = iTipoTrib
        objVarTipoTrib.sIdentificador = "TIPO_TRIB"
        objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27900
    
        'atualizar "I_OTP_TIPO_DOC_INFO"
        objVarTipoDocAtual.vValor = objVarTipoDoc.vValor
        objVarTipoDocAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoDocAtual)
        If lErro <> SUCESSO Then Error 27901
    
        'atualizar "I_OTP_NAT_OP"
        objVarNatOpAtual.vValor = objVarNatOp.vValor
        objVarNatOpAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarNatOpAtual)
        If lErro <> SUCESSO Then Error 27902
        
    End If
    
    ObterTribPadraoForn = SUCESSO

    Exit Function

Erro_ObterTribPadraoForn:

    ObterTribPadraoForn = Err

    Select Case Err

        Case 27894 To 27902
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153797)

    End Select

    Exit Function

End Function

Private Function ObterCidadeIBGEFilialFat(sCidadeIBGE As String) As Long

Dim lErro As Long
Dim objVarFilialFatObj As New ClassVariavelCalculo
Dim objFilialEmpresa As AdmFiliais, objCidades As New ClassCidades

On Error GoTo Erro_ObterCidadeIBGEFilialFat
    
    lErro = mvarobjDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFatObj)
    If lErro <> SUCESSO Then Error 32014
    
    Set objFilialEmpresa = objVarFilialFatObj.vValor
    
    objCidades.sDescricao = objFilialEmpresa.objEndereco.sCidade
    lErro = CF("Cidade_Le_Nome", objCidades)
    If lErro <> SUCESSO Then Error 32014
    
    sCidadeIBGE = objCidades.sCodIBGE
    
    ObterCidadeIBGEFilialFat = SUCESSO

    Exit Function

Erro_ObterCidadeIBGEFilialFat:

    ObterCidadeIBGEFilialFat = Err

    Select Case Err

        Case 32014 To 32019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153798)

    End Select

    Exit Function

End Function

Function ObterObjFilialFat(objFilial As AdmFiliais) As Long

Dim lErro As Long
Dim objVarFilial As New ClassVariavelCalculo
Dim objVarFilialAtual As New ClassVariavelCalculo
Dim objVarFilialObj As New ClassVariavelCalculo
On Error GoTo Erro_ObterObjFilialFat
    
    lErro = mvarobjDoc.ObterVar("FILIAL_FATURAMENTO", objVarFilial)
    If lErro <> SUCESSO Then Error 32014
        
    lErro = mvarobjVarAux.ObterVar("I_FILIAL_FAT_CODIGO", objVarFilialAtual)
    If lErro <> SUCESSO Then Error 32015
        
    'se nao houve modificacao
    If objVarFilialAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilial.vValor = objVarFilialAtual.vValor Then
        
        lErro = mvarobjVarAux.ObterVar("FILIAL_FAT_OBJ", objVarFilialObj)
        If lErro <> SUCESSO Then Error 32016
        
        Set objFilial = objVarFilialObj.vValor
        
    Else
    
        If Not (gobjCRFAT.objFilialEmpresa Is Nothing) Then
        
            If objVarFilial.vValor = gobjCRFAT.objFilialEmpresa.iCodFilial Then
            
                Set objFilial = gobjCRFAT.objFilialEmpresa
                
            Else
                
                Set objFilial = New AdmFiliais
                
                'fazer a leitura
                objFilial.iCodFilial = objVarFilial.vValor
                lErro = CF("FilialEmpresa_Le", objFilial)
                If lErro <> SUCESSO Then Error 32017
    
            End If
        
        Else
        
            Set objFilial = New AdmFiliais
            
            'fazer a leitura
            objFilial.iCodFilial = objVarFilial.vValor
            lErro = CF("FilialEmpresa_Le", objFilial)
            If lErro <> SUCESSO Then Error 32017
        
        End If
        
        Set objVarFilialObj.vValor = objFilial
        objVarFilialObj.sIdentificador = "FILIAL_FAT_OBJ"
        objVarFilialObj.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "FILIAL_FAT_OBJ"
        lErro = mvarobjVarAux.GuardarVar(objVarFilialObj)
        If lErro <> SUCESSO Then Error 32018
         
        'atualizar "I_FILIAL_FAT_CODIGO"
        objVarFilialAtual.vValor = objVarFilial.vValor
        objVarFilialAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialAtual)
        If lErro <> SUCESSO Then Error 32019
     
    End If
    
    ObterObjFilialFat = SUCESSO

    Exit Function

Erro_ObterObjFilialFat:

    ObterObjFilialFat = Err

    Select Case Err

        Case 32014 To 32019
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153798)

    End Select

    Exit Function

End Function

Function ObterObjNatOpPadrao(objNatOpPadrao As ClassNatOpPadrao) As Long

Dim lErro As Long, objVarNatOpPadraoObj As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, objFilialEmpresa As AdmFiliais
Dim objTipoDocInfo As New ClassTipoDocInfo, objVarFilialFatObj As New ClassVariavelCalculo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, iTipoAtividadeEmp As Integer
Dim objVarTipoAtivTribAtual As New ClassVariavelCalculo, objVarTipoOpTribAtual As New ClassVariavelCalculo

On Error GoTo Erro_ObterObjNatOpPadrao
    
    lErro = mvarobjDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFatObj)
    If lErro <> SUCESSO Then Error 56945
    
    Set objFilialEmpresa = objVarFilialFatObj.vValor
    
    lErro = mvarobjDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 56946
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 56947

    lErro = mvarobjVarAux.ObterVar("I_NOPP_TIPOOPERTRIB", objVarTipoOpTribAtual)
    If lErro <> SUCESSO Then Error 56948
        
    lErro = mvarobjVarAux.ObterVar("I_NOPP_TIPOATIVTRIB", objVarTipoAtivTribAtual)
    If lErro <> SUCESSO Then Error 56949
        
    'se nao houve modificacao
    If objVarTipoOpTribAtual.iStatus <> VAR_PREENCH_VAZIO And objVarTipoOpTribAtual.vValor = objTipoDocInfo.iTipoOperacaoTrib And _
        objVarTipoAtivTribAtual.iStatus <> VAR_PREENCH_VAZIO And objVarTipoAtivTribAtual.vValor = objFilialEmpresa.iTipoTribAtividade Then
        
        lErro = mvarobjVarAux.ObterVar("NATOPPADRAO_OBJ", objVarNatOpPadraoObj)
        If lErro <> SUCESSO Then Error 56950
        
        Set objNatOpPadrao = objVarNatOpPadraoObj.vValor
        
    Else
    
        Set objNatOpPadrao = New ClassNatOpPadrao
        
        'fazer a leitura
        objNatOpPadrao.iTipoOperacao = objTipoDocInfo.iTipoOperacaoTrib
        objNatOpPadrao.iTipoAtividadeEmp = objFilialEmpresa.iTipoTribAtividade
        lErro = CF("NatOpPadrao_Le_Padrao", objNatOpPadrao)
        If lErro <> SUCESSO And lErro <> 56944 Then Error 56951
        If lErro <> SUCESSO Then Error 56972
    
        Set objVarNatOpPadraoObj.vValor = objNatOpPadrao
        objVarNatOpPadraoObj.sIdentificador = "NATOPPADRAO_OBJ"
        objVarNatOpPadraoObj.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "NATOPPADRAO_OBJ"
        lErro = mvarobjVarAux.GuardarVar(objVarNatOpPadraoObj)
        If lErro <> SUCESSO Then Error 56952
         
        'atualizar "I_NOPP_TIPOOPERTRIB"
        objVarTipoOpTribAtual.vValor = objNatOpPadrao.iTipoOperacao
        objVarTipoOpTribAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoOpTribAtual)
        If lErro <> SUCESSO Then Error 56953
    
        'atualizar "I_NOPP_TIPOATIVTRIB"
        objVarTipoAtivTribAtual.vValor = objNatOpPadrao.iTipoAtividadeEmp
        objVarTipoAtivTribAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarTipoAtivTribAtual)
        If lErro <> SUCESSO Then Error 56954
    
    End If
    
    ObterObjNatOpPadrao = SUCESSO
     
    Exit Function
    
Erro_ObterObjNatOpPadrao:

    ObterObjNatOpPadrao = Err
     
    Select Case Err
          
        Case 56945 To 56954, 56972
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153799)
     
    End Select
     
    Exit Function
    
End Function

'''??? transferir p/TRBSelect junto com a cte ERRO_LEITURA_NATOPPADRAO acima
''Function ObterTribPadraoOp(objTipoDocInfo As ClassTipoDocInfo, iTipoTrib As Integer, iTipoTribAtividade As Integer) As Long
''
''Dim lErro As Long, lComando As Long
''
''On Error GoTo Erro_ObterTribPadraoOp
''
''    lComando = Comando_Abrir()
''    If lComando = 0 Then Error 56938
''
''    lErro = Comando_Executar(lComando, "SELECT TipoTribEmp FROM NatOpPadrao WHERE TipoOperacao = ? AND TipoAtividadeEmp = ? AND Padrao = ?", _
''        iTipoTrib, objTipoDocInfo.iTipoOperacaoTrib, iTipoTribAtividade, 1)
''    If lErro <> AD_SQL_SUCESSO Then Error 56939
''
''    lErro = Comando_BuscarProximo(lComando)
''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 56940
''
''    'se nao achei nenhum padrao
''    If lErro <> AD_SQL_SUCESSO Then iTipoTrib = 0
''
''    lErro = Comando_Fechar(lComando)
''
''    ObterTribPadraoOp = SUCESSO
''
''    Exit Function
''
''Erro_ObterTribPadraoOp:
''
''    ObterTribPadraoOp = Err
''
''    Select Case Err
''
''        Case 56938
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''
''        Case 56939, 56940
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATOPPADRAO", Err)
''
''        Case Else
''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153800)
''
''    End Select
''
''    Call Comando_Fechar(lComando)
''
''    Exit Function
''
''End Function

Function ObterColCategFilForn(colCateg As Collection) As Long
'faz com que colCateg aponte p/uma colecao com as (categoria,item) associadas a filial fornecedor associada ao doc
Dim lErro As Long, objFilialFornecedor As New ClassFilialFornecedor
Dim objVarFornecedor As New ClassVariavelCalculo, objVarFilialFornecedor As New ClassVariavelCalculo
Dim objVarFornecedorAtual As New ClassVariavelCalculo, objVarFilialFornecedorAtual As New ClassVariavelCalculo
Dim objVarColCateg As New ClassVariavelCalculo
On Error GoTo Erro_ObterColCategFilForn

    lErro = mvarobjDoc.ObterVar("FORN_CODIGO", objVarFornecedor)
    If lErro <> SUCESSO Then Error 27332
        
    lErro = mvarobjDoc.ObterVar("FILIAL_FORN_CODIGO", objVarFilialFornecedor)
    If lErro <> SUCESSO Then Error 27333
        
    lErro = mvarobjVarAux.ObterVar("I_CCF_FORN_CODIGO", objVarFornecedorAtual)
    If lErro <> SUCESSO Then Error 27334
        
    lErro = mvarobjVarAux.ObterVar("I_CCF_FILIAL_FORN_CODIGO", objVarFilialFornecedorAtual)
    If lErro <> SUCESSO Then Error 27335

    'se nao houve modificacao
    If objVarFornecedorAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFilialFornecedorAtual.iStatus <> VAR_PREENCH_VAZIO And objVarFornecedor.vValor = objVarFornecedorAtual.vValor And objVarFornecedor.vValor = objVarFornecedorAtual.vValor Then
        
        lErro = mvarobjVarAux.ObterVar("I_FILIAL_FORN_COL_CATEG", objVarColCateg)
        If lErro <> SUCESSO Then Error 27336
        
        Set colCateg = objVarColCateg.vValor
    
    Else
    
        'fazer a leitura
        objFilialFornecedor.lCodFornecedor = objVarFornecedor.vValor
        objFilialFornecedor.iCodFilial = objVarFilialFornecedor.vValor
        lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilialFornecedor, colCateg)
        If lErro <> SUCESSO And lErro <> 127085 Then Error 27337

        Set objVarColCateg.vValor = colCateg
        objVarColCateg.sIdentificador = "I_FILIAL_FORN_COL_CATEG"
        objVarColCateg.iStatus = VAR_PREENCH_AUTOMATICO
        
        'atualizar "I_FILIAL_FORN_COL_CATEG"
        lErro = mvarobjVarAux.GuardarVar(objVarColCateg)
        If lErro <> SUCESSO Then Error 27338
        
        'atualizar "I_CCF_FORN_CODIGO"
        objVarFornecedorAtual.vValor = objVarFornecedor.vValor
        objVarFornecedorAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFornecedorAtual)
        If lErro <> SUCESSO Then Error 27339
    
        'atualizar "I_CCF_FILIAL_FORN_CODIGO"
        objVarFilialFornecedorAtual.vValor = objVarFilialFornecedor.vValor
        objVarFilialFornecedorAtual.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = mvarobjVarAux.GuardarVar(objVarFilialFornecedorAtual)
        If lErro <> SUCESSO Then Error 27340
    
    End If
    
    ObterColCategFilForn = SUCESSO

    Exit Function

Erro_ObterColCategFilForn:

    ObterColCategFilForn = Err

    Select Case Err

        Case 27332 To 27340
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153801)

    End Select

    Exit Function

End Function

Private Function ObterCidadeIBGECliente(sCidadeIBGECodigo As String) As Long

Dim lErro As Long
Dim objVarFilCliObj As New ClassVariavelCalculo
Dim objFilialCliente As ClassFilialCliente, objEndereco As New ClassEndereco
Dim objCidade As New ClassCidades

On Error GoTo Erro_ObterCidadeIBGECliente

    lErro = mvarobjDoc.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilCliObj)
    If lErro <> SUCESSO Then Error 27393
    
    Set objFilialCliente = objVarFilCliObj.vValor
                    
    objEndereco.lCodigo = objFilialCliente.lEndereco
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO Then Error 27391
        
    objCidade.sDescricao = objEndereco.sCidade
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then Error 27391
    
    If lErro = SUCESSO Then
        sCidadeIBGECodigo = objCidade.sCodIBGE
    Else
        sCidadeIBGECodigo = ""
    End If
    
    ObterCidadeIBGECliente = SUCESSO

    Exit Function

Erro_ObterCidadeIBGECliente:

    ObterCidadeIBGECliente = Err

    Select Case Err

        Case 27389 To 27393, 27778
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153794)

    End Select

    Exit Function

End Function

