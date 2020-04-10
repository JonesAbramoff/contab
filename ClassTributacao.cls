VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'existirá um objeto global desta classe

'Responderá a questoes sobre a Empresa, a filial,

'Performance:
'TributacaoTipo_Obter
'    poderia manter comando preparado ou cache+versao da tabela
'
'poderia criar indirecao qto a quem implementa os drivers tratadores de imposto
'
'carregará (logo de cara ou a medida que for necessario)
'dados da FilialEmpresa,
'Movimentos
'colecao de objetos tratadores de impostos, já na ordem em que devem ser chamados
'
'Visivel p/Docs (Pedido, NF,...):
'
'ELEMENTOS:
'==========
Public colDriversTributos As New Collection 'Colecao de objetos tratadores de tributos na ordem em que devem ser chamados
Private mvarobjTributacaoTipoLido As New ClassTipoDeTributacaoMovto
Private mvarobjTipoDocInfoLido As New ClassTipoDocInfo

'Interface p / doc:
'=================
Function Tributacao_Doc_Iniciar(objDoc As ClassTributoDoc)
'pode fazer inicializacoes ligadas a melhoraria de performance
    Tributacao_Doc_Iniciar = SUCESSO
    
End Function

Function Tributacao_Doc_Terminar(objDoc As ClassTributoDoc)
'pode fazer finalizacoes ligadas a melhoraria de performance
    Tributacao_Doc_Terminar = SUCESSO
    
End Function

Function AtualizaImpostos(objDoc As ClassTributoDoc, Optional iRecalculaAutomatica As Integer = 0) As Long
'faz a atualizacao de todas as variaveis associadas ao calculo de impostos
'recalcula todos os itens e o doc como um todo
Dim lErro As Long, iFaltaInfo As Integer
Dim objTributoDriver As ClassTributoDriver
'Dim colItens As Collection, objDocItem As ClassTributoDocItem

On Error GoTo Erro_AtualizaImpostos
    
    lErro = PreparaAtualizacaoImpostos(objDoc, iFaltaInfo, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27243

    'se nao faltam informacoes p/calcular tributacao
    If iFaltaInfo = 0 Then
        
        'vai passar pelos impostos em sequencia
        For Each objTributoDriver In colDriversTributos
        
            lErro = objTributoDriver.AtualizarImposto(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 27244
            
        Next
        
        lErro = FinalizaAtualizacaoImpostos(objDoc, iRecalculaAutomatica)
        If lErro <> SUCESSO Then Error 27243
    
    End If
    
    AtualizaImpostos = SUCESSO

    Exit Function

Erro_AtualizaImpostos:

    AtualizaImpostos = Err

    Select Case Err

        Case 27243, 27244
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153774)

    End Select

    Exit Function

End Function

Private Function ObterNatOpETribCli2(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'obtem natop e tributacao padrao para os itens e os complementos(frete,...)
Dim lErro As Long
Dim objVarNatOpDoc As New ClassVariavelCalculo, objVarNatOp As New ClassVariavelCalculo
Dim objVarTipoTribDoc As New ClassVariavelCalculo, objVarTipoTrib As New ClassVariavelCalculo
Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection
Dim objTributacaoTipo As ClassTipoDeTributacaoMovto, objVarNatBCCred As New ClassVariavelCalculo
Dim objNaturezaOp As ClassNaturezaOp

Dim objVarObjProduto As New ClassVariavelCalculo
Dim objProduto As ClassProduto
Dim objNatOpPadrao As ClassNatOpPadrao
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarTipoDoc As New ClassVariavelCalculo, objVarTipoDocInfo As New ClassVariavelCalculo
Dim iTipoTrib As Integer, bAtualiza As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo, bVenda As Boolean, sNatOp As String, iTipoOperacaoTrib As Integer

On Error GoTo Erro_ObterNatOpETribCli2

    If gobjCRFAT.iCalculaCFOPItem <> 0 Then
    
        lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
        If lErro <> SUCESSO Then Error 22810
        
        objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
        lErro = TipoDocInfo_Obter(objTipoDocInfo)
        If lErro <> SUCESSO Then Error 22810
    
        bVenda = IIf(objTipoDocInfo.iTipoOperacaoTrib = 2 Or objTipoDocInfo.iTipoOperacaoTrib = 19, True, False)
    
    Else
    
        bVenda = False
    
    End If
    
    lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOpDoc)
    If lErro <> SUCESSO Then Error 22811
        
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTribDoc)
    If lErro <> SUCESSO Then Error 22817
        
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
    
        bAtualiza = True
        If gobjCRFAT.iOtimizaTrib = MARCADO Then
            If objDocItem.iJaCalculado = MARCADO Then bAtualiza = False
        End If
        If bAtualiza Then
        
            Set objProduto = Nothing
        
            'obter NatOp do item
            lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
            If lErro <> SUCESSO Then Error 22810
            
            If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
                If gobjCRFAT.iCalculaCFOPItem = 0 Or objVarNatOpDoc.iStatus = VAR_PREENCH_MANUAL Or bVenda = False Then
                
                    objVarNatOp.vValor = objVarNatOpDoc.vValor
                
                Else
                
                    Set objNatOpPadrao = New ClassNatOpPadrao
                
                    lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
                    If lErro <> SUCESSO Then Error 22810
                
                    Set objProduto = objVarObjProduto.vValor
            
                    iTipoOperacaoTrib = 2
            
                    If objProduto.iNatureza = 8 Then
                        objNatOpPadrao.iTipoAtividadeEmp = 3 'prestacao de servicos
                        iTipoOperacaoTrib = 19
                    ElseIf objProduto.iCompras = 1 Then 'se for produto comprado
                        objNatOpPadrao.iTipoAtividadeEmp = 1 'comercio
                    Else ' se for produto produzido
                        objNatOpPadrao.iTipoAtividadeEmp = 2 'industrial
                    End If
                            
                    'fazer a leitura
                    objNatOpPadrao.iTipoOperacao = iTipoOperacaoTrib
                    
                    lErro = CF("NatOpPadrao_Le_Padrao", objNatOpPadrao)
                    If lErro <> SUCESSO And lErro <> 56944 Then Error 56951
                    If lErro = SUCESSO Then
            
                        sNatOp = objNatOpPadrao.sCFOPEmp
                        If Len(sNatOp) = 3 Then sNatOp = NaturezaOp_Conv34(sNatOp, gdtDataAtual)
            
                        If left(sNatOp, 1) <> left(objVarNatOpDoc.vValor, 1) Then
                            sNatOp = left(objVarNatOpDoc.vValor, 1) + Mid(sNatOp, 2)
                        End If
                        
                        objVarNatOp.vValor = sNatOp
                        
                    Else
                    
                        objVarNatOp.vValor = objVarNatOpDoc.vValor
                        
                    End If
                
                End If
                
                objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
                
                lErro = objDocItem.GuardarVar(objVarNatOp)
                If lErro <> SUCESSO Then Error 22812
            
            End If
            
            lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then Error 22813
            
            If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
                If gobjCRFAT.iCalculaCFOPItem = 0 Or bVenda = False Then
                
                    If objVarNatOpDoc.vValor = objVarNatOp.vValor Then
                
                        objVarTipoTrib.vValor = objVarTipoTribDoc.vValor
                        objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
                        
                        lErro = objDocItem.GuardarVar(objVarTipoTrib)
                        If lErro <> SUCESSO Then Error 22816
                    
                    End If
                    
                Else
    
                    If objVarTipoTribDoc.iStatus = VAR_PREENCH_MANUAL And objVarNatOp.vValor = objVarNatOpDoc.vValor Then
                    
                        objVarTipoTrib.vValor = objVarTipoTribDoc.vValor
                        objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
                        
                        lErro = objDocItem.GuardarVar(objVarTipoTrib)
                        If lErro <> SUCESSO Then Error 22816
                    
                    Else
                    
                        lErro = objDoc.ObterVar("TIPO_DOC_SIGLA", objVarTipoDoc)
                        If lErro <> SUCESSO Then Error 27341
                            
                        lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
                        If lErro <> SUCESSO Then Error 27343
                            
                        lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
                        If lErro <> SUCESSO Then Error 27344
            
                        'chamar funcao em Tributacao
                        lErro = CF("ObterTribPadraoCli", iTipoTrib, objVarTipoDoc.vValor, objVarNatOp.vValor, objVarCliente.vValor, objVarFilialCliente.vValor)
                        If lErro <> SUCESSO Then Error 27350
                    
                        If iTipoTrib = 0 Then
                        
                            Set objNatOpPadrao = New ClassNatOpPadrao
                        
                            If objProduto Is Nothing Then
                        
                                lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
                                If lErro <> SUCESSO Then Error 22810
                            
                                Set objProduto = objVarObjProduto.vValor
            
                            End If
            
                            iTipoOperacaoTrib = 2
            
                            If objProduto.iNatureza = 8 Then
                                objNatOpPadrao.iTipoAtividadeEmp = 3 'prestacao de servicos
                                iTipoOperacaoTrib = 19
                            ElseIf objProduto.iCompras = 1 Then 'se for produto comprado
                                objNatOpPadrao.iTipoAtividadeEmp = 1 'comercio
                            Else ' se for produto produzido
                                objNatOpPadrao.iTipoAtividadeEmp = 2 'industrial
                            End If
                                    
                            'fazer a leitura
                            objNatOpPadrao.iTipoOperacao = iTipoOperacaoTrib
                            
                            lErro = CF("NatOpPadrao_Le_Padrao", objNatOpPadrao)
                            If lErro <> SUCESSO And lErro <> 56944 Then Error 56951
                            If lErro = SUCESSO Then
                    
                                iTipoTrib = objNatOpPadrao.iTipoTribEmp
                                
                            Else
                            
                                objVarNatOp.vValor = objVarTipoTribDoc.vValor
                                
                            End If
                        
                        End If
                        
                        'atualizar "TIPO_TRIB"
                        objVarTipoTrib.vValor = iTipoTrib
                    
                        objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
                        
                        lErro = objDocItem.GuardarVar(objVarTipoTrib)
                        If lErro <> SUCESSO Then Error 22816
                    
                    End If
                    
               End If
                    
            End If
            
            lErro = objDocItem.ObterVar("NAT_BC_CRED", objVarNatBCCred)
            If lErro <> SUCESSO Then Error 22813
                    
            If objVarNatBCCred.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
                objVarNatBCCred.iStatus = VAR_PREENCH_AUTOMATICO
            
                Set objNaturezaOp = New ClassNaturezaOp
            
                objNaturezaOp.sCodigo = objVarNatOp.vValor
                lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
                If lErro <> SUCESSO And lErro <> 17958 Then Error 22816
            
                If Len(Trim(objNaturezaOp.sNatBCCredPadrao)) > 0 Then
                    objVarNatBCCred.vValor = objNaturezaOp.sNatBCCredPadrao
                Else
                    Set objTributacaoTipo = New ClassTipoDeTributacaoMovto
                    objTributacaoTipo.iTipo = objVarTipoTrib.vValor
                    Call TributacaoTipo_Obter(objTributacaoTipo)
                
                    objVarNatBCCred.vValor = objTributacaoTipo.sNatBCCred
                End If
                
                lErro = objDocItem.GuardarVar(objVarNatBCCred)
                If lErro <> SUCESSO Then Error 22812
            
            End If
        End If

    Next
    
'    Call objDoc.ObterColDocCompl(colDocCompl)
'
'    For Each objDocCompl In colDocCompl
'
'        'obter NatOp do item
'        lErro = objDocCompl.ObterVar("NAT_OPERACAO", objVarNatOp)
'        If lErro <> SUCESSO Then Error 22818
'
'        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
'
'            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
'            objVarNatOp.vValor = objVarNatOpDoc.vValor
'
'            lErro = objDocCompl.GuardarVar(objVarNatOp)
'            If lErro <> SUCESSO Then Error 22819
'
'        End If
'
'        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
'        If lErro <> SUCESSO Then Error 22820
'
'        If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
'
'            If objVarNatOpDoc.vValor = objVarNatOp.vValor Then
'
'                objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
'                objVarTipoTrib.vValor = objVarTipoTribDoc.vValor
'
'                lErro = objDocCompl.GuardarVar(objVarTipoTrib)
'                If lErro <> SUCESSO Then Error 22821
'
'            End If
'
'        End If
'
'    Next
    
    ObterNatOpETribCli2 = SUCESSO

    Exit Function

Erro_ObterNatOpETribCli2:

    ObterNatOpETribCli2 = Err

    Select Case Err

        Case 22810 To 22821
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153775)

    End Select

    Exit Function

End Function

Private Function ObterNatOpETrib(objDoc As ClassTributoDoc, iFaltaInfo As Integer, iRecalculaAutomatica As Integer) As Long
Dim lErro As Long, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, iIndConsumidorFinal As Integer
Dim objVarIndConsumidorFinal As New ClassVariavelCalculo, objVarIndPresenca As New ClassVariavelCalculo

'NT 2015.003 (ICMS Interestadual)
Dim objFilCliIEIsento As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo

On Error GoTo Erro_ObterNatOpETrib

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 22835
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22836
    
    lErro = objDoc.ObterVar("IND_CONSUMIDOR_FINAL", objVarIndConsumidorFinal)
    If lErro <> SUCESSO Then Error 22835
    
    If objVarIndConsumidorFinal.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
    
        iIndConsumidorFinal = objTipoDocInfo.iIndConsumidorFinal
        
        If iIndConsumidorFinal = 0 Then
        
            'obter identificador da filial do cliente
            lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
            If lErro <> SUCESSO Then Error 22835
            
            'se o cliente e a filial já foram informados
            If (objVarFilialCliente.iStatus <> VAR_PREENCH_VAZIO And objVarFilialCliente.vValor <> 0) Then
            
                If objTipoDocInfo.iCodigo <> DOCINFO_NFISSAT And objTipoDocInfo.iCodigo <> DOCINFO_TNFISFV Then
                
                    lErro = objDoc.ObterVar("FILIAL_CLIENTE_IEISENTO", objFilCliIEIsento)
                    If lErro <> SUCESSO Then Error 22835
                
                Else
                
                    If objTipoDocInfo.iCodigo = DOCINFO_TNFISFV Then
                        objFilCliIEIsento.vValor = DESMARCADO
                    Else
                        objFilCliIEIsento.vValor = MARCADO
                    End If
                
                End If
            
                If objFilCliIEIsento.vValor = MARCADO Then iIndConsumidorFinal = 1
            
            End If
            
        End If
    
        objVarIndConsumidorFinal.vValor = iIndConsumidorFinal
        objVarIndConsumidorFinal.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIndConsumidorFinal)
        If lErro <> SUCESSO Then Error 22835
    
    End If
    
    lErro = objDoc.ObterVar("IND_PRESENCA", objVarIndPresenca)
    If lErro <> SUCESSO Then Error 22835
    
    If objVarIndPresenca.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
    
        objVarIndPresenca.vValor = objTipoDocInfo.iIndPresenca
        objVarIndPresenca.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarIndPresenca)
        If lErro <> SUCESSO Then Error 22835
    
    End If
    
    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_SAIDA_CLI, TRIB_ENTRADA_CLI
            lErro = NaturezaOp_GerarPadraoCli(objDoc, iFaltaInfo, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22837
            
            lErro = ObterNatOpETribCli2(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22838
            
        Case TRIB_SAIDA_FORN, TRIB_ENTRADA_FORN
            lErro = NaturezaOp_GerarPadraoForn(objDoc, iFaltaInfo, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22839
        
            lErro = ObterNatOpETribForn2(objDoc, iRecalculaAutomatica)
            If lErro <> SUCESSO Then Error 22840
        
    End Select
    
    ObterNatOpETrib = SUCESSO

    Exit Function

Erro_ObterNatOpETrib:

    ObterNatOpETrib = Err

    Select Case Err

        Case 22835 To 22840

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153776)

    End Select

    Exit Function

End Function

Private Function ValorDescGlobal_Ratear(objDoc As ClassTributoDoc) As Long
'faz o rateio do desconto global entre os itens do doc (nf ou pv)

'Dim lErro As Long, dValorDescGlobal As Double, dValorDescItem As Double, dValorDescAcum As Double
'Dim objVarValorDescGlobal As New ClassVariavelCalculo, objVarProdutoValor As New ClassVariavelCalculo
'Dim colItens As Collection, dValorBaseRateio As Double, objVarDescontoValor As New ClassVariavelCalculo
'Dim objDocItem As ClassTributoDocItem, dDif As Double, objVarItemTipo As New ClassVariavelCalculo
'Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection

On Error GoTo Erro_ValorDescGlobal_Ratear

'    Call objDoc.ObterColDocItens(colItens)
'
'    If colItens.Count <> 0 Then
'
'        Call objDoc.ObterColDocCompl(colDocCompl)
'
'        For Each objDocCompl In colDocCompl
'
'            lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
'            If lErro <> SUCESSO Then gError 124116
'
'            If objVarItemTipo.vValor = ITEM_TIPO_DESCONTO Then
'
'                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVarValorDescGlobal)
'                If lErro <> SUCESSO Then gError 124110
'
'                Exit For
'
'            End If
'
'        Next
'
'        dValorDescGlobal = objVarValorDescGlobal.vValor
'
'        If dValorDescGlobal <> 0 Then
'
'            'percorrer os itens obtendo valores
'            For Each objDocItem In colItens
'
'                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
'                If lErro <> SUCESSO Then gError 124111
'
'                dValorBaseRateio = dValorBaseRateio + objVarProdutoValor.vValor
'
'            Next
'
'            objVarDescontoValor.sIdentificador = "PRODUTO_DESC_GLOBAL"
'
'            'percorrer os itens guardando valores "aproximados"
'            For Each objDocItem In colItens
'
'                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
'                If lErro <> SUCESSO Then gError 124112
'
'                dValorDescItem = Int(100 * objVarProdutoValor.vValor * dValorDescGlobal / dValorBaseRateio) / 100
'
'                objVarDescontoValor.vValor = dValorDescItem
'                dValorDescAcum = dValorDescAcum + dValorDescItem
'
'                lErro = objDocItem.GuardarVar(objVarDescontoValor)
'                If lErro <> SUCESSO Then gError 124113
'
'            Next
'
'            dDif = dValorDescGlobal - dValorDescAcum
'
'            Do While dDif >= DELTA_VALORMONETARIO
'
'                'percorrer os itens aumentando em R$ 0,01 o valor do desconto
'                For Each objDocItem In colItens
'
'                    lErro = objDocItem.ObterVar("PRODUTO_DESC_GLOBAL", objVarDescontoValor)
'                    If lErro <> SUCESSO Then gError 124114
'
'                    objVarDescontoValor.vValor = objVarDescontoValor.vValor + 0.01
'                    lErro = objDocItem.GuardarVar(objVarDescontoValor)
'                    If lErro <> SUCESSO Then gError 124115
'
'                    dDif = dDif - 0.01
'                    If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
'
'                Next
'
'            Loop
'
'        Else
'
'            objVarDescontoValor.sIdentificador = "PRODUTO_DESC_GLOBAL"
'
'            'percorrer os itens guardando valores "aproximados"
'            For Each objDocItem In colItens
'
'                objVarDescontoValor.vValor = 0
'
'                lErro = objDocItem.GuardarVar(objVarDescontoValor)
'                If lErro <> SUCESSO Then gError 124113
'
'            Next
'
'        End If
'
'    End If
    
    ValorDescGlobal_Ratear = SUCESSO
     
    Exit Function
    
Erro_ValorDescGlobal_Ratear:

    ValorDescGlobal_Ratear = gErr
     
    Select Case gErr
          
        Case 124110 To 124115
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153777)
     
    End Select
     
    Exit Function

End Function

Private Function PreparaAtualizacaoImpostos(objDoc As ClassTributoDoc, iFaltaInfo As Integer, iRecalculaAutomatica As Integer) As Long
Dim lErro As Long, objVarFilialFat As New ClassVariavelCalculo
On Error GoTo Erro_PreparaAtualizacaoImpostos

    'para forcar a carga do bd
    mvarobjTributacaoTipoLido.iTipo = 0
    mvarobjTipoDocInfoLido.iCodigo = 0
    
    lErro = ObterNatOpETrib(objDoc, iFaltaInfo, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 22809
    
    lErro = ValorDescGlobal_Ratear(objDoc)
    If lErro <> SUCESSO Then Error 22809
    
    lErro = ValorDespGlobal_Ratear(objDoc)
    If lErro <> SUCESSO Then Error 22809
    
    PreparaAtualizacaoImpostos = SUCESSO

    Exit Function

Erro_PreparaAtualizacaoImpostos:

    PreparaAtualizacaoImpostos = Err

    Select Case Err

        Case 22809, 32012
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153778)

    End Select

    Exit Function

End Function

Function ObterNaturezaOpPadrao(objDoc As ClassTributoDoc, sNatOp As String) As Long
'retorna em sNatOp a natureza de operacao padrao de acordo com o tipo do docto

'??? pendencias
'512 deveria ir p/619 se o cliente de outro estado nao for contribuinte
'511 deveria ir p/618 se o cliente de outro estado nao for contribuinte

Dim lErro As Long, sUFFilial As String, sUFCliForn As String
Dim objTipoDocInfo As New ClassTipoDocInfo, objVarUF As New ClassVariavelCalculo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, objFilialEmpresa As AdmFiliais
Dim objVarFilialFatObj As New ClassVariavelCalculo
On Error GoTo Erro_ObterNaturezaOpPadrao

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 27883
    
    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFatObj)
    If lErro <> SUCESSO Then Error 32020
    
    Set objFilialEmpresa = objVarFilialFatObj.vValor
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    lErro = TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 27884
    
    lErro = ObterNaturezaOpPadrao2(objDoc, sNatOp)
    If lErro <> SUCESSO Then Error 56956
    
    sUFFilial = objFilialEmpresa.objEndereco.sSiglaEstado
       
    If objTipoDocInfo.iCodigo <> DOCINFO_NFISSAT And objTipoDocInfo.iCodigo <> DOCINFO_TNFISFV Then
    
        'obtem a UF do cliente/fornecedor
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_CLI, TRIB_SAIDA_CLI
                lErro = objDoc.ObterVar("CLIENTE_UF", objVarUF)
                If lErro <> SUCESSO Then Error 27885
                
            Case TRIB_ENTRADA_FORN, TRIB_SAIDA_FORN
                lErro = objDoc.ObterVar("FORN_UF", objVarUF)
                If lErro <> SUCESSO Then Error 27886
                
        End Select
        
        sUFCliForn = objVarUF.vValor
    
    End If
    
    If sUFCliForn = "" Then sUFCliForn = sUFFilial
    
    If left(sNatOp, 1) = "1" Or left(sNatOp, 1) = "5" Then
    
        'aplico a modificacao em termos de localizacao geografica(interna, initerestadual,internacional)
        If sUFFilial = sUFCliForn Then 'interna
            'nao faz nada
        Else
            If sUFCliForn = "EX" Then 'internacional
                sNatOp = CStr(CInt(sNatOp) + IIf(Len(sNatOp) = 3, 200, 2000))
            Else 'interestadual
                sNatOp = CStr(CInt(sNatOp) + IIf(Len(sNatOp) = 3, 100, 1000))
            End If
        End If
    
    End If
    
    ObterNaturezaOpPadrao = SUCESSO

    Exit Function

Erro_ObterNaturezaOpPadrao:

    ObterNaturezaOpPadrao = Err

    Select Case Err

        Case 27406, 27883 To 27886, 56956
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153779)

    End Select

    Exit Function

End Function

Function ObterNaturezaOpPadrao2(objDoc As ClassTributoDoc, sNatOp As String) As Long
'retorna em sNatOp a natureza de operacao padrao de acordo com o tipo do docto
'anda nao estao disponiveis dados de cliente ou fornecedor

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo
Dim objVarNatOpPadraoObj As New ClassVariavelCalculo, objNatOpPadrao As ClassNatOpPadrao
Dim objVarDataRef As New ClassVariavelCalculo, dtDataRef As Date

On Error GoTo Erro_ObterNaturezaOpPadrao2

    lErro = objDoc.ObterVar("NATOPPADRAO_OBJ", objVarNatOpPadraoObj)
    If lErro <> SUCESSO Then Error 56932
    
    Set objNatOpPadrao = objVarNatOpPadraoObj.vValor
    
    If Not (objNatOpPadrao Is Nothing) Then sNatOp = objNatOpPadrao.sCFOPEmp
    
    If sNatOp = "" Then
    
        lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
        If lErro <> SUCESSO Then Error 56932
        
        objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
        lErro = TipoDocInfo_Obter(objTipoDocInfo)
        If lErro <> SUCESSO Then Error 56933
            
        sNatOp = objTipoDocInfo.sNaturezaOperacaoPadrao
    
    End If
    
    If Len(sNatOp) = 3 Then
        
        'se for uma entrada
        If sNatOp < NATUREZA_SAIDA_COD_INICIAL Then
            
            lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataRef)
            If lErro <> SUCESSO Then Error 56932
        
        Else
        
            lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataRef)
            If lErro <> SUCESSO Then Error 56932
        
        End If
        
        dtDataRef = objVarDataRef.vValor
        
        sNatOp = NaturezaOp_Conv34(sNatOp, dtDataRef)
        
    End If
    
    ObterNaturezaOpPadrao2 = SUCESSO

    Exit Function

Erro_ObterNaturezaOpPadrao2:

    ObterNaturezaOpPadrao2 = Err

    Select Case Err

        Case 56932, 56933
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153780)

    End Select

    Exit Function

End Function

'Visivel p/drivers de impostos:
'====================================

Function TributacaoTipo_Obter(objTributacaoTipo As ClassTipoDeTributacaoMovto) As Long
'preenche objTributacaoTipo a partir de objTributacaoTipo.iTipo
Dim lErro As Long

On Error GoTo Erro_TributacaoTipo_Obter

    'se nao é igual ao ultimo tipo "lido"
    If mvarobjTributacaoTipoLido.iTipo <> objTributacaoTipo.iTipo Then
        
        'ler do bd
        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
        If lErro <> SUCESSO Then Error 27255
        
        Call mvarobjTributacaoTipoLido.Copiar(objTributacaoTipo)
        
    Else
    
        Call objTributacaoTipo.Copiar(mvarobjTributacaoTipoLido)
    
    End If
    
    TributacaoTipo_Obter = SUCESSO

    Exit Function

Erro_TributacaoTipo_Obter:

    TributacaoTipo_Obter = Err

    Select Case Err

        Case 27255
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153781)

    End Select

    Exit Function

End Function


Function NaturezaOp_GerarPadraoCli(objDoc As ClassTributoDoc, iFaltaInfo As Integer, iRecalculaAutomatica As Integer) As Long
Dim lErro As Long, iTipoTrib As Integer
Dim objVarCliente As New ClassVariavelCalculo, objVarFilialCliente As New ClassVariavelCalculo
Dim objVarNatOp As New ClassVariavelCalculo, objVarTipoTrib As New ClassVariavelCalculo, sNatOp As String
On Error GoTo Erro_NaturezaOp_GerarPadraoCli

    'obter codigo do cliente
    lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCliente)
    If lErro <> SUCESSO Then Error 27656
    
    'obter identificador da filial do cliente
    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilialCliente)
    If lErro <> SUCESSO Then Error 27657
    
    'obter natureza de operacao
    lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
    If lErro <> SUCESSO Then Error 27650
        
    'se o cliente e a filial já foram informados
    If (objVarFilialCliente.iStatus <> VAR_PREENCH_VAZIO And objVarCliente.iStatus <> VAR_PREENCH_VAZIO And objVarFilialCliente.vValor <> 0 And objVarCliente.vValor <> 0) Then
        
        iFaltaInfo = 0
        
        'se a natureza de operacao nao estiver preenchida
        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
            'obter natureza de operacao padrao
            lErro = objDoc.ObterNatOpPadrao(sNatOp)
            If lErro <> SUCESSO Then Error 27651
        
            objVarNatOp.vValor = sNatOp
            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
            
            lErro = objDoc.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 27652
            
        End If
        
        If objVarNatOp.vValor <> "" Then
        
            'obter tipo de tributacao padrao
            lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then Error 27653
            
            'se o tipo de tributacao nao estiver preenchido
            If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
                'obter tipo de tributacao padrao
                lErro = objDoc.ObterTribPadrao(iTipoTrib)
                If lErro <> SUCESSO Then Error 27654
                
                objVarTipoTrib.vValor = iTipoTrib
                objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
                
                lErro = objDoc.GuardarVar(objVarTipoTrib)
                If lErro <> SUCESSO Then Error 27655
            
            End If
        
        End If
        
    Else
    
        iFaltaInfo = 1
        
        'se a natureza de operacao nao estiver preenchida
        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
        
            lErro = ObterNaturezaOpPadrao2(objDoc, sNatOp)
            If lErro <> SUCESSO Then Error 56934
            
            objVarNatOp.vValor = sNatOp
            lErro = objDoc.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 56935
    
        End If
        
    End If
        
    NaturezaOp_GerarPadraoCli = SUCESSO

    Exit Function

Erro_NaturezaOp_GerarPadraoCli:

    NaturezaOp_GerarPadraoCli = Err

    Select Case Err
    
        Case 27656, 27657, 27650 To 27655, 56934, 56935
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153782)

    End Select

    Exit Function

End Function

Function TipoDocInfo_Obter(objTipoDocInfo As ClassTipoDocInfo) As Long
'preenche objTipoDocInfo a partir de objTipoDocInfo.iCodigo
Dim lErro As Long

On Error GoTo Erro_TipoDocInfo_Obter

    'se nao é igual ao ultimo tipo "lido"
    If mvarobjTipoDocInfoLido.iCodigo <> objTipoDocInfo.iCodigo Then
        
        'ler do bd
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then Error 27861
        
        Call mvarobjTipoDocInfoLido.Copiar(objTipoDocInfo)
        
    Else
    
        Call objTipoDocInfo.Copiar(mvarobjTipoDocInfoLido)
    
    End If
    
    TipoDocInfo_Obter = SUCESSO

    Exit Function

Erro_TipoDocInfo_Obter:

    TipoDocInfo_Obter = Err

    Select Case Err

        Case 27861
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153783)

    End Select

    Exit Function

End Function

Function NaturezaOp_GerarPadraoForn(objDoc As ClassTributoDoc, iFaltaInfo As Integer, iRecalculaAutomatica As Integer) As Long

Dim lErro As Long, iTipoTrib As Integer
Dim objVarForn As New ClassVariavelCalculo, objVarFilialForn As New ClassVariavelCalculo
Dim objVarNatOp As New ClassVariavelCalculo, objVarTipoTrib As New ClassVariavelCalculo, sNatOp As String

On Error GoTo Erro_NaturezaOp_GerarPadraoForn

    'obter codigo do fornecedor
    lErro = objDoc.ObterVar("FORN_CODIGO", objVarForn)
    If lErro <> SUCESSO Then Error 27962
    
    'obter identificador da filial do fornecedor
    lErro = objDoc.ObterVar("FILIAL_FORN_CODIGO", objVarFilialForn)
    If lErro <> SUCESSO Then Error 27963
    
    'obter natureza de operacao
    lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
    If lErro <> SUCESSO Then Error 27964
        
    'se o Forn e a filial já foram informados
    If (objVarFilialForn.iStatus <> VAR_PREENCH_VAZIO And objVarForn.iStatus <> VAR_PREENCH_VAZIO And objVarFilialForn.vValor <> 0 And objVarForn.vValor <> 0) Then
        
        iFaltaInfo = 0
        
        'se a natureza de operacao nao estiver preenchida
        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
            'obter natureza de operacao padrao
            lErro = objDoc.ObterNatOpPadrao(sNatOp)
            If lErro <> SUCESSO Then Error 27965
        
            objVarNatOp.vValor = sNatOp
            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
            
            lErro = objDoc.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 27966
            
        End If
        
        If objVarNatOp.vValor <> "" Then
        
            'obter tipo de tributacao padrao
            lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then Error 27967
            
            'se o tipo de tributacao nao estiver preenchido
            If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
            
                'obter tipo de tributacao padrao
                lErro = objDoc.ObterTribPadrao(iTipoTrib)
                If lErro <> SUCESSO Then Error 27968
                
                objVarTipoTrib.vValor = iTipoTrib
                objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
                
                lErro = objDoc.GuardarVar(objVarTipoTrib)
                If lErro <> SUCESSO Then Error 27969
            
            End If
        
        End If
        
    Else
        iFaltaInfo = 1
    
        'se a natureza de operacao nao estiver preenchida
        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
        
            lErro = ObterNaturezaOpPadrao2(objDoc, sNatOp)
            If lErro <> SUCESSO Then Error 56966
            
            objVarNatOp.vValor = sNatOp
            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDoc.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 56967
    
        End If
            
    End If
        
    NaturezaOp_GerarPadraoForn = SUCESSO

    Exit Function

Erro_NaturezaOp_GerarPadraoForn:

    NaturezaOp_GerarPadraoForn = Err

    Select Case Err
    
        Case 27962 To 27969, 56966, 56967
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153784)

    End Select

    Exit Function

End Function

Private Function ObterNatOpETribForn2(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'obtem natop e tributacao padrao para os itens e os complementos(frete,...)

Dim lErro As Long, iTipoTrib As Integer
Dim objVarNatOpDoc As New ClassVariavelCalculo, objVarNatOp As New ClassVariavelCalculo
Dim objVarTipoTribDoc As New ClassVariavelCalculo, objVarTipoTrib As New ClassVariavelCalculo
Dim colItens As Collection, objDocItem As ClassTributoDocItem, objVarTipoDocSigla As New ClassVariavelCalculo
Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection, objVarProduto As New ClassVariavelCalculo
Dim objTributacaoTipo As ClassTipoDeTributacaoMovto, objVarNatBCCred As New ClassVariavelCalculo
Dim objNaturezaOp As ClassNaturezaOp

On Error GoTo Erro_ObterNatOpETribForn2

    lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOpDoc)
    If lErro <> SUCESSO Then Error 22822
        
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTribDoc)
    If lErro <> SUCESSO Then Error 22823
        
    lErro = objDoc.ObterVar("TIPO_DOC_SIGLA", objVarTipoDocSigla)
    If lErro <> SUCESSO Then Error 22824
                
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
    
        'obter NatOp do item
        lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
        If lErro <> SUCESSO Then Error 22825
        
        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
        
            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
            objVarNatOp.vValor = objVarNatOpDoc.vValor
            
            lErro = objDocItem.GuardarVar(objVarNatOp)
            If lErro <> SUCESSO Then Error 22826
        
        End If
        
        lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 22827
        
        If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
        
            If objVarTipoTribDoc.iStatus = VAR_PREENCH_MANUAL And objVarNatOpDoc.vValor = objVarNatOp.vValor Then
            
                iTipoTrib = objVarTipoTribDoc.vValor
                        
            Else
                
                'obter trib + especifica p/produto
                
                lErro = objDocItem.ObterVar("PRODUTO_CODIGO", objVarProduto)
                If lErro <> SUCESSO Then Error 22828
                
                lErro = CF("ObterTribPadraoProd", iTipoTrib, objVarTipoDocSigla.vValor, objVarNatOp.vValor, objVarProduto.vValor)
                If lErro <> SUCESSO Then Error 22829
                
                If iTipoTrib = 0 Then
                    iTipoTrib = objVarTipoTribDoc.vValor
                Else
                    If objVarTipoTribDoc.iStatus <> VAR_PREENCH_MANUAL And objVarNatOpDoc.vValor = objVarNatOp.vValor Then
                    
                        objVarTipoTribDoc.vValor = iTipoTrib
                        lErro = objDoc.GuardarVar(objVarTipoTribDoc)
                        If lErro <> SUCESSO Then Error 22830
                        
                    End If
                End If
                
            End If
                
            objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
            objVarTipoTrib.vValor = iTipoTrib
            
            lErro = objDocItem.GuardarVar(objVarTipoTrib)
            If lErro <> SUCESSO Then Error 22830
                
        End If
        
        lErro = objDocItem.ObterVar("NAT_BC_CRED", objVarNatBCCred)
        If lErro <> SUCESSO Then Error 22828
                
        If objVarNatBCCred.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
        
            objVarNatBCCred.iStatus = VAR_PREENCH_AUTOMATICO
        
            Set objNaturezaOp = New ClassNaturezaOp
        
            objNaturezaOp.sCodigo = objVarNatOp.vValor
            lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
            If lErro <> SUCESSO And lErro <> 17958 Then Error 22816
        
            If Len(Trim(objNaturezaOp.sNatBCCredPadrao)) > 0 Then
                objVarNatBCCred.vValor = objNaturezaOp.sNatBCCredPadrao
            Else
                Set objTributacaoTipo = New ClassTipoDeTributacaoMovto
                objTributacaoTipo.iTipo = objVarTipoTrib.vValor
                Call TributacaoTipo_Obter(objTributacaoTipo)
            
                objVarNatBCCred.vValor = objTributacaoTipo.sNatBCCred
            End If
            
            lErro = objDocItem.GuardarVar(objVarNatBCCred)
            If lErro <> SUCESSO Then Error 22830
        
        End If

    Next
    
'    Call objDoc.ObterColDocCompl(colDocCompl)
'
'    For Each objDocCompl In colDocCompl
'
'        'obter NatOp do item
'        lErro = objDocCompl.ObterVar("NAT_OPERACAO", objVarNatOp)
'        If lErro <> SUCESSO Then Error 22831
'
'        If objVarNatOp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
'
'            objVarNatOp.iStatus = VAR_PREENCH_AUTOMATICO
'            objVarNatOp.vValor = objVarNatOpDoc.vValor
'
'            lErro = objDocCompl.GuardarVar(objVarNatOp)
'            If lErro <> SUCESSO Then Error 22832
'
'        End If
'
'        lErro = objDocCompl.ObterVar("TIPO_TRIB", objVarTipoTrib)
'        If lErro <> SUCESSO Then Error 22833
'
'        If objVarTipoTrib.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = 1 Then
'
'            If objVarNatOpDoc.vValor = objVarNatOp.vValor Then
'
'                objVarTipoTrib.iStatus = VAR_PREENCH_AUTOMATICO
'                objVarTipoTrib.vValor = objVarTipoTribDoc.vValor
'
'                lErro = objDocCompl.GuardarVar(objVarTipoTrib)
'                If lErro <> SUCESSO Then Error 22834
'
'            End If
'
'        End If
'
'    Next
    
    ObterNatOpETribForn2 = SUCESSO

    Exit Function

Erro_ObterNatOpETribForn2:

    ObterNatOpETribForn2 = Err

    Select Case Err

        Case 22822 To 22834
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153785)

    End Select

    Exit Function

End Function

Private Sub Class_Initialize()

Dim objTributoDriver As ClassTributoDriver
Dim lErro As Long

On Error GoTo Erro_Class_Initialize

    lErro = CF("Tributacao_Iniciar", colDriversTributos)
    If lErro <> SUCESSO Then Error 27247
    
'    Set objTributoDriver = New ClassTributoISS
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 27247
'    colDriversTributos.Add objTributoDriver
'
'    Set objTributoDriver = New ClassTributoIR
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 27248
'    colDriversTributos.Add objTributoDriver
'
'    Set objTributoDriver = New ClassTributoIPI
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 27249
'    colDriversTributos.Add objTributoDriver
'
'    Set objTributoDriver = New ClassTributoICMS
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 27250
'    colDriversTributos.Add objTributoDriver
'
'    Set objTributoDriver = New ClassTributoINSS
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 32251
'    colDriversTributos.Add objTributoDriver
'
'    'valendo a partir de 01/12/2002 pela mp 66
'    Set objTributoDriver = New ClassTributoPIS
'    lErro = objTributoDriver.Ativar()
'    If lErro <> SUCESSO Then Error 32251
'    colDriversTributos.Add objTributoDriver
        
    Exit Sub
     
Erro_Class_Initialize:

    Select Case Err
          
        Case 27247 To 27250, 32251
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153786)
     
    End Select
     
    Exit Sub
            
End Sub

Private Function ValorDespGlobal_Ratear(objDoc As ClassTributoDoc) As Long
'faz o rateio do (frete+seguro+outras despesas) entre os itens do doc (nf ou pv)

'Dim lErro As Long, dValorDespGlobal As Double, dValorDespItem As Double, dValorDespAcum As Double
'Dim objVarValorDespGlobal As New ClassVariavelCalculo, objVarProdutoValor As New ClassVariavelCalculo
'Dim colItens As Collection, dValorBaseRateio As Double, objVarDespValor As New ClassVariavelCalculo
'Dim objDocItem As ClassTributoDocItem, dDif As Double, objVarItemTipo As New ClassVariavelCalculo
'Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection
Dim lErro As Long, colItens As Collection, objVarValorGlobal As New ClassVariavelCalculo
Dim dValorItens As Double, dValorGlobal As Double, dValorBaseRateio As Double
Dim objVarValorItem As New ClassVariavelCalculo, objVarProdutoValor As New ClassVariavelCalculo
Dim dValorItem As Double, dDif As Double, dValorAcum As Double, objDocItem As ClassTributoDocItem
Dim iIndice As Integer, sID As String, sIDItem As String, iStatus As Integer
Dim objVarTipoDocInfo As New ClassVariavelCalculo, objTipoDocInfo As New ClassTipoDocInfo, dProdQtde As Double, dValorItemOrig As Double
Dim objVarItemNFOrig As New ClassVariavelCalculo, objItemNFOrig As ClassItemNF, objVarProdutoQtde As New ClassVariavelCalculo

On Error GoTo Erro_ValorDespGlobal_Ratear

    Call objDoc.ObterColDocItens(colItens)
        
    If colItens.Count <> 0 Then
    
        lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
        If lErro <> SUCESSO Then gError 27867
    
        objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor

        If objTipoDocInfo.iCodigo = DOCINFO_NFIEDV Then
    
             For iIndice = 2 To 4
                 
                 Select Case iIndice
                     Case 1
                         sID = "DESCONTO_VALOR"
                         sIDItem = "VALOR_DESCONTO_ITEM"
                     Case 2
                         sID = "FRETE_VALOR"
                         sIDItem = "VALOR_FRETE_ITEM"
                     Case 3
                         sID = "SEGURO_VALOR"
                         sIDItem = "VALOR_SEGURO_ITEM"
                     Case 4
                         sID = "DESP_ACESS_VALOR"
                         sIDItem = "VALOR_DESPESAS_ITEM"
                 End Select
             
                 lErro = objDoc.ObterVar(sID, objVarValorGlobal)
                 If lErro <> SUCESSO Then gError 124116
                 
                 dValorItens = 0
                 
                 For Each objDocItem In colItens
                 
                    lErro = objDocItem.ObterVar(sIDItem, objVarValorItem)
                    If lErro <> SUCESSO Then gError 124116
                     
                    If objVarValorItem.iStatus <> VAR_PREENCH_MANUAL Then
                         
                        dValorItem = 0
                        
                        lErro = objDocItem.ObterVar("OBJ_ITEM_NF_ORIG", objVarItemNFOrig)
                        If lErro <> SUCESSO Then Error 27721
                        
                        Set objItemNFOrig = objVarItemNFOrig.vValor
                        
                        If Not (objItemNFOrig Is Nothing) Then
        
                            lErro = objDocItem.ObterVar("PRODUTO_QTDE", objVarProdutoQtde)
                            If lErro <> SUCESSO Then gError 124111
                            
                            dProdQtde = objVarProdutoQtde.vValor
                 
                            Select Case iIndice
                            
                                Case 2 'frete
                                    dValorItemOrig = objItemNFOrig.objTributacaoItemNF.dValorFreteItem
                                    
                                Case 3 'seguro
                                    dValorItemOrig = objItemNFOrig.objTributacaoItemNF.dValorSeguroItem
                                
                                Case 4 'outras despesas (vai incorporar ICMS ST pq na entrada nao pode destacar ST)
                                    dValorItemOrig = objItemNFOrig.objTributacaoItemNF.dValorOutrasDespesasItem + objItemNFOrig.objTributacaoItemNF.dICMSSubstValor
                            
                            End Select
        
                            'se for uma devolucao parcial precisa ratear
                            If dProdQtde <> 0 And objItemNFOrig.objTributacaoItemNF.dQtdTrib <> 0 And Abs(objItemNFOrig.objTributacaoItemNF.dQtdTrib - dProdQtde) > QTDE_ESTOQUE_DELTA Then
                            
                                dValorItem = Arredonda_Moeda(dValorItemOrig * dProdQtde / objItemNFOrig.objTributacaoItemNF.dQtdTrib)
                            
                            Else
                            
                                dValorItem = dValorItemOrig
                            
                            End If
                            
                        End If
                        
                        objVarValorItem.vValor = dValorItem
        
                        lErro = objDocItem.GuardarVar(objVarValorItem)
                        If lErro <> SUCESSO Then gError 124113
                    
                    End If
                     
                    dValorItens = dValorItens + objVarValorItem.vValor
                 
                 Next
                 
                 'If objVarValorGlobal.iStatus <> VAR_PREENCH_MANUAL Then
                 'O teste acima não funciona pois não existe os iManual para esses campos no BD
                 dValorGlobal = Arredonda_Moeda(objVarValorGlobal.vValor)
                If Abs(dValorGlobal - dValorItens) > DELTA_VALORMONETARIO Then
                 
                    objVarValorGlobal.iStatus = VAR_PREENCH_MANUAL
                    objVarValorGlobal.vValor = dValorItens 'Coloca a soma dos itens no valor global
                    lErro = objDoc.GuardarVar(objVarValorGlobal)
                    If lErro <> SUCESSO Then gError 124115
                         
                End If
                         
            Next
            
        Else
    
             For iIndice = 2 To 4
                 
                 Select Case iIndice
                     Case 1
                         sID = "DESCONTO_VALOR"
                         sIDItem = "VALOR_DESCONTO_ITEM"
                     Case 2
                         sID = "FRETE_VALOR"
                         sIDItem = "VALOR_FRETE_ITEM"
                     Case 3
                         sID = "SEGURO_VALOR"
                         sIDItem = "VALOR_SEGURO_ITEM"
                     Case 4
                         sID = "DESP_ACESS_VALOR"
                         sIDItem = "VALOR_DESPESAS_ITEM"
                 End Select
             
                 lErro = objDoc.ObterVar(sID, objVarValorGlobal)
                 If lErro <> SUCESSO Then gError 124116
                 
                 dValorGlobal = Arredonda_Moeda(objVarValorGlobal.vValor)
                 dValorItens = 0
                 dValorBaseRateio = 0
                 
                 objVarValorItem.sIdentificador = sIDItem
                 iStatus = VAR_PREENCH_AUTOMATICO
                 
                 For Each objDocItem In colItens
                 
                     lErro = objDocItem.ObterVar(sIDItem, objVarValorItem)
                     If lErro <> SUCESSO Then gError 124116
                     dValorItens = dValorItens + objVarValorItem.vValor
                     
                     If objVarValorItem.iStatus = VAR_PREENCH_MANUAL Then
                         iStatus = VAR_PREENCH_MANUAL
                     End If
                     
                     lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
                     If lErro <> SUCESSO Then gError 124111
        
                     dValorBaseRateio = dValorBaseRateio + objVarProdutoValor.vValor
                 
                 Next
                 
                 If Abs(dValorGlobal - dValorItens) > DELTA_VALORMONETARIO Then
                             
                     'Se mexeram nos itens acerta o global, senão acerta os itens
                     If iStatus <> VAR_PREENCH_AUTOMATICO Then
         
                         'percorrer os itens para marcar como automático para que
                         'quando alterar o valor de novo retestar onde ocorreu a alteração
                         For Each objDocItem In colItens
         
                             lErro = objDocItem.ObterVar(sIDItem, objVarValorItem)
                             If lErro <> SUCESSO Then gError 124114
         
                             objVarValorItem.iStatus = VAR_PREENCH_AUTOMATICO
                             lErro = objDocItem.GuardarVar(objVarValorItem)
                             If lErro <> SUCESSO Then gError 124115
         
                         Next
                         
                         lErro = objDoc.ObterVar(sID, objVarValorGlobal)
                         If lErro <> SUCESSO Then gError 124114
         
                         objVarValorGlobal.vValor = dValorItens 'Coloca a soma dos itens no valor global
                         lErro = objDoc.GuardarVar(objVarValorGlobal)
                         If lErro <> SUCESSO Then gError 124115
                     
                     Else
             
                         If dValorGlobal < DELTA_VALORMONETARIO Then
                         
                             For Each objDocItem In colItens
                         
                                 objVarValorItem.vValor = 0
                                 lErro = objDocItem.GuardarVar(objVarValorItem)
                                 If lErro <> SUCESSO Then gError 124115
                                 
                             Next
                                 
                         Else
                         
                             'percorrer os itens guardando valores "aproximados"
                             dValorAcum = 0
                             For Each objDocItem In colItens
                 
                                 lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
                                 If lErro <> SUCESSO Then gError 124112
                 
                                 If dValorBaseRateio > 0 Then
                                     dValorItem = Int(100 * objVarProdutoValor.vValor * dValorGlobal / dValorBaseRateio) / 100
                                 Else
                                     dValorItem = Int(100 * dValorGlobal / colItens.Count) / 100
                                 End If
                 
                                 objVarValorItem.vValor = dValorItem
                                 dValorAcum = dValorAcum + dValorItem
                 
                                 lErro = objDocItem.GuardarVar(objVarValorItem)
                                 If lErro <> SUCESSO Then gError 124113
                 
                             Next
                 
                             dDif = dValorGlobal - dValorAcum
                             
                             Do While dDif >= DELTA_VALORMONETARIO
                 
                                 'percorrer os itens aumentando em R$ 0,01 o valor do desconto
                                 For Each objDocItem In colItens
                 
                                     lErro = objDocItem.ObterVar(sIDItem, objVarValorItem)
                                     If lErro <> SUCESSO Then gError 124114
                 
                                     objVarValorItem.vValor = objVarValorItem.vValor + 0.01
                                     lErro = objDocItem.GuardarVar(objVarValorItem)
                                     If lErro <> SUCESSO Then gError 124115
                 
                                     dDif = dDif - 0.01
                                     If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
                 
                                 Next
                             Loop
                         End If
                     End If
                     
                 End If
                 
             Next
             
         End If
        
    End If
    
'        Call objDoc.ObterColDocCompl(colDocCompl)
'
'        For Each objDocCompl In colDocCompl
'
'            lErro = objDocCompl.ObterVar("ITEM_TIPO", objVarItemTipo)
'            If lErro <> SUCESSO Then gError 124116
'
'            If objVarItemTipo.vValor <> ITEM_TIPO_DESCONTO Then
'
'                lErro = objDocCompl.ObterVar("PRODUTO_VALOR", objVarValorDespGlobal)
'                If lErro <> SUCESSO Then gError 124110
'
'                dValorDespGlobal = Arredonda_Moeda(dValorDespGlobal + objVarValorDespGlobal.vValor)
'
'            End If
'
'        Next
'
'        If dValorDespGlobal <> 0 Then
'
'            'percorrer os itens obtendo valores
'            For Each objDocItem In colItens
'
'                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
'                If lErro <> SUCESSO Then gError 124111
'
'                dValorBaseRateio = dValorBaseRateio + objVarProdutoValor.vValor
'
'            Next
'
'            objVarDespValor.sIdentificador = "PRODUTO_DESP_GLOBAL"
'
'            'percorrer os itens guardando valores "aproximados"
'            For Each objDocItem In colItens
'
'                lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarProdutoValor)
'                If lErro <> SUCESSO Then gError 124112
'
'                dValorDespItem = Int(100 * objVarProdutoValor.vValor * dValorDespGlobal / dValorBaseRateio) / 100
'
'                objVarDespValor.vValor = dValorDespItem
'                dValorDespAcum = dValorDespAcum + dValorDespItem
'
'                lErro = objDocItem.GuardarVar(objVarDespValor)
'                If lErro <> SUCESSO Then gError 124113
'
'            Next
'
'            dDif = dValorDespGlobal - dValorDespAcum
'
'            Do While dDif >= DELTA_VALORMONETARIO
'
'                'percorrer os itens aumentando em R$ 0,01 o valor do desconto
'                For Each objDocItem In colItens
'
'                    lErro = objDocItem.ObterVar("PRODUTO_DESP_GLOBAL", objVarDespValor)
'                    If lErro <> SUCESSO Then gError 124114
'
'                    objVarDespValor.vValor = objVarDespValor.vValor + 0.01
'                    lErro = objDocItem.GuardarVar(objVarDespValor)
'                    If lErro <> SUCESSO Then gError 124115
'
'                    dDif = dDif - 0.01
'                    If Abs(dDif) < DELTA_VALORMONETARIO Then Exit For
'
'                Next
'
'            Loop
'
'        Else
'
'            objVarDespValor.sIdentificador = "PRODUTO_DESP_GLOBAL"
'
'            'percorrer os itens guardando valores "aproximados"
'            For Each objDocItem In colItens
'
'                objVarDespValor.vValor = 0
'
'                lErro = objDocItem.GuardarVar(objVarDespValor)
'                If lErro <> SUCESSO Then gError 124113
'
'            Next
'
'        End If
'
'    End If
    
    ValorDespGlobal_Ratear = SUCESSO
     
    Exit Function
    
Erro_ValorDespGlobal_Ratear:

    ValorDespGlobal_Ratear = gErr
     
    Select Case gErr
          
        Case 124110 To 124115
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153777)
     
    End Select
     
    Exit Function

End Function

Private Function Doc_ObterFinalidadeNFePadrao(ByVal objDoc As ClassTributoDoc, iFinalidadeNFe As Integer) As Long

Dim lErro As Long, lComando As Long, iFinalidadeEspec As Integer, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, objVarNatOp As New ClassVariavelCalculo

On Error GoTo Erro_Doc_ObterFinalidadeNFePadrao

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iFinalidadeNFe = objTipoDocInfo.iFinalidadeNFe
    
    If iFinalidadeNFe = 1 Then
    
        'obter natureza de operacao
        lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        If Len(Trim(objVarNatOp.vValor)) <> 0 Then
        
            lComando = Comando_Abrir()
            If lComando = 0 Then gError 201575
            
            lErro = Comando_Executar(lComando, "SELECT FinalidadeEspec FROM NaturezaOp WHERE Codigo = ?", iFinalidadeEspec, Trim(objVarNatOp.vValor))
            If lErro <> AD_SQL_SUCESSO Then gError 201576
        
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201577
                
            If lErro = AD_SQL_SUCESSO Then
                
                Select Case iFinalidadeEspec
                
                    Case 1 'Devolução
                        iFinalidadeNFe = 4
                        
                    Case 2 'Combustível
                    
                    Case 3 'Serviço de Transporte
                        
                    Case Else
                        
                End Select
                
            End If
        
            Call Comando_Fechar(lComando)
            
        End If
    
    End If
    
    Doc_ObterFinalidadeNFePadrao = SUCESSO
    
    Exit Function
    
Erro_Doc_ObterFinalidadeNFePadrao:

    Doc_ObterFinalidadeNFePadrao = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201575
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 201576, 201577
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201574)

    End Select
    
    If lComando <> 0 Then Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function FinalizaAtualizacaoImpostos(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long

Dim lErro As Long, objVarNatOp As New ClassVariavelCalculo, objVarDestino As New ClassVariavelCalculo
Dim colItens As Collection, objDocItem As ClassTributoDocItem, objVarTipoTribDoc As New ClassVariavelCalculo, iFinalidadeNFe As Integer
Dim objVarFinalidadeNFe As New ClassVariavelCalculo, iTipoTribDoc As Integer, iTipoTribItens As Integer, objVarTipoTrib As New ClassVariavelCalculo

On Error GoTo Erro_FinalizaAtualizacaoImpostos

    iTipoTribItens = -1
    iTipoTribDoc = -1
    
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTribDoc)
    If lErro <> SUCESSO Then Error 22809
        
    If objVarTipoTribDoc.iStatus = VAR_PREENCH_AUTOMATICO Then
    
        iTipoTribDoc = objVarTipoTribDoc.vValor
    
    End If
    
    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens marcando-os como já calculados
    For Each objDocItem In colItens

        If iTipoTribDoc <> -1 Then

            lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
            If lErro <> SUCESSO Then Error 22809
            
            If iTipoTribItens = -1 Then
                
                iTipoTribItens = objVarTipoTrib.vValor
                
            Else
            
                If iTipoTribItens <> objVarTipoTrib.vValor Then iTipoTribDoc = -1
            
            End If
        
        End If
        
        Call objDocItem.objTributacaoDocItem.Calculado

    Next
    
    'se o tipo de tributacao do doc é automatico, todos os itens tem o mesmo tipo e o do doc está diferente dos itens
    If iTipoTribDoc <> -1 And iTipoTribItens <> -1 And iTipoTribDoc <> iTipoTribItens Then
    
        objVarTipoTribDoc.vValor = iTipoTribItens
        lErro = objDoc.GuardarVar(objVarTipoTribDoc)
        If lErro <> SUCESSO Then Error 22809
    
    End If
    
    lErro = objDoc.ObterVar("DESTINO", objVarDestino)
    If lErro <> SUCESSO Then Error 22809
    
    If objVarDestino.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
    
        'obter natureza de operacao
        lErro = objDoc.ObterVar("NAT_OPERACAO", objVarNatOp)
        If lErro <> SUCESSO Then Error 22809
        
        Select Case left(objVarNatOp.vValor, 1)
        
            Case "2", "6"
                objVarDestino.vValor = 2
            
            Case "3", "7"
                objVarDestino.vValor = 3
            
            Case Else
                objVarDestino.vValor = 1
        
        End Select
        
        objVarDestino.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarDestino)
        If lErro <> SUCESSO Then Error 22809
        
    End If
        
    lErro = objDoc.ObterVar("FINALIDADE_NFE", objVarFinalidadeNFe)
    If lErro <> SUCESSO Then Error 22835
    
    If objVarFinalidadeNFe.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
    
        lErro = Doc_ObterFinalidadeNFePadrao(objDoc, iFinalidadeNFe)
        If lErro <> SUCESSO Then Error 22809
        
        objVarFinalidadeNFe.vValor = iFinalidadeNFe
        objVarFinalidadeNFe.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarFinalidadeNFe)
        If lErro <> SUCESSO Then Error 22835
    
    End If
    
    FinalizaAtualizacaoImpostos = SUCESSO

    Exit Function

Erro_FinalizaAtualizacaoImpostos:

    FinalizaAtualizacaoImpostos = Err

    Select Case Err

        Case 22809
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153778)

    End Select

    Exit Function

End Function

