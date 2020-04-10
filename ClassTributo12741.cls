VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributo12741"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'para tratamento da lei 12741/12 que exige o destaque do valor aproximado dos tributos em venda ao consumidor

Implements ClassTributoDriver

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvarsUFFilialEmpresa As String 'UF da filial empresa "corrente"
Private mvariNotaInterna As Integer
Private mvardtDataEmissao As Date

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao

Dim lErro As Long, objVarTipoDocInfo As New ClassVariavelCalculo, objTipoDocInfo As New ClassTipoDocInfo, objTRBConfig As New ClassTRBConfig, objVarDataEmissao As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    mvardtDataEmissao = objVarDataEmissao.vValor
    
    mvariNotaInterna = IIf(objTipoDocInfo.iEmitente = EMITENTE_EMPRESA, 1, 0)
        
    If (objTipoDocInfo.iTipoTrib = TRIB_ENTRADA_CLI Or objTipoDocInfo.iTipoTrib = TRIB_SAIDA_FORN Or (objTipoDocInfo.iTipoTrib = TRIB_ENTRADA_FORN And mvariNotaInterna <> 0)) Then
    
        lErro = AtualizarImposto_Zerar(objDoc, iRecalculaAutomatica)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Else
    
        Select Case objTipoDocInfo.iTipoTrib
        
            Case TRIB_ENTRADA_FORN
                lErro = AtualizarImposto_Zerar(objDoc, iRecalculaAutomatica) '??? AtualizarImposto_EntradaForn
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Case TRIB_SAIDA_CLI
                If objTipoDocInfo.iFaturamento <> 1 And objTipoDocInfo.iCodigo <> DOCINFO_PVN And objTipoDocInfo.iCodigo <> DOCINFO_PSRVN Then
                
                    lErro = AtualizarImposto_Zerar(objDoc, iRecalculaAutomatica)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Else
                    
                    lErro = AtualizarImposto_SaidaCli(objDoc, iRecalculaAutomatica, objTipoDocInfo)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                End If
            
        End Select
        
    End If
        
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153676)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
'permite a inicializacao do driver
'pode ser utilizado p/criar cache etc
Dim lErro As Long
On Error GoTo Erro_ClassTributoDriver_Ativar

    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153677)

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

Private Function AtualizarImposto_SaidaCli(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long

Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long, objVarTotTribTipo As New ClassVariavelCalculo, iTotTribTipo As Integer, objVarTotTribItemTipo As New ClassVariavelCalculo
Dim bAtualiza As Boolean

On Error GoTo Erro_AtualizarImposto_SaidaCli

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
            lErro = AtualizarItem_SaidaCli(objDoc, objDocItem, iRecalculaAutomatica, objTipoDocInfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
    Next

    'se todos os itens tiverem o mesmo tipo de calculo entao coloca-lo para o doc como um todo
    lErro = objDoc.ObterVar("TOT_TRIB_TIPO", objVarTotTribTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If objVarTotTribTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        iTotTribTipo = -1
        For Each objDocItem In colItens
        
            lErro = objDocItem.ObterVar("TOT_TRIB_TIPO", objVarTotTribItemTipo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If iTotTribTipo = -1 Then
                iTotTribTipo = objVarTotTribItemTipo.vValor
            Else
                If iTotTribTipo <> objVarTotTribItemTipo.vValor Then
                    iTotTribTipo = -1
                    Exit For
                End If
            End If
        
        Next
    
        If iTotTribTipo <> -1 Then
        
            objVarTotTribTipo.iStatus = VAR_PREENCH_AUTOMATICO
            objVarTotTribTipo.vValor = iTotTribTipo
            lErro = objDoc.GuardarVar(objVarTotTribTipo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
    End If
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    AtualizarImposto_SaidaCli = SUCESSO
    
    Exit Function
    
Erro_AtualizarImposto_SaidaCli:

    AtualizarImposto_SaidaCli = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201509)

    End Select
    
    Exit Function

End Function

Private Function AtualizarItem_ObterValoresAtuais(objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer, objVarTotTrib As ClassVariavelCalculo, objVarTotTribTipo As ClassVariavelCalculo, objVarOrigemProd As ClassVariavelCalculo, objVarTotTribFed As ClassVariavelCalculo, objVarTotTribEst As ClassVariavelCalculo, objVarTotTribMunic As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarItem_ObterValoresAtuais

    lErro = objDocItem.ObterVar("TOT_TRIB", objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTrib.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("TOT_TRIB_TIPO", objVarTotTribTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("ORIGEM_MERCADORIA", objVarOrigemProd)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
         
    lErro = objDocItem.ObterVar("TOT_TRIB_FED", objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribFed.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("TOT_TRIB_EST", objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribEst.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDocItem.ObterVar("TOT_TRIB_MUNIC", objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribMunic.iStatus = VAR_PREENCH_AUTOMATICO
        
    AtualizarItem_ObterValoresAtuais = SUCESSO

    Exit Function

Erro_AtualizarItem_ObterValoresAtuais:

    AtualizarItem_ObterValoresAtuais = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201511)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_GuardarValores(objDocItem As ClassTributoDocItem, objVarTotTrib As ClassVariavelCalculo, ByVal dTotTrib As Double, objVarTotTribTipo As ClassVariavelCalculo, ByVal iTotTrib As Integer, objVarTotTribFed As ClassVariavelCalculo, ByVal dTotTribAliqFed As Double, objVarTotTribEst As ClassVariavelCalculo, ByVal dTotTribAliqEst As Double, objVarTotTribMunic As ClassVariavelCalculo, ByVal dTotTribAliqMunic As Double) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarItem_GuardarValores

    If objVarTotTrib.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTotTrib.vValor = dTotTrib
        objVarTotTrib.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTotTrib)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarTotTribTipo.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTotTribTipo.vValor = iTotTrib
        objVarTotTribTipo.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTotTribTipo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarTotTribFed.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTotTribFed.vValor = dTotTribAliqFed
        objVarTotTribFed.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTotTribFed)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarTotTribEst.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTotTribEst.vValor = dTotTribAliqEst
        objVarTotTribEst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTotTribEst)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If objVarTotTribMunic.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarTotTribMunic.vValor = dTotTribAliqMunic
        objVarTotTribMunic.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDocItem.GuardarVar(objVarTotTribMunic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    AtualizarItem_GuardarValores = SUCESSO

    Exit Function

Erro_AtualizarItem_GuardarValores:

    AtualizarItem_GuardarValores = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201512)

    End Select

    Exit Function

End Function

Private Function AtualizarItem_SaidaCli(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer, objTipoDocInfo As ClassTipoDocInfo) As Long

Dim lErro As Long, objVarTotTrib As New ClassVariavelCalculo, objVarTotTribTipo As New ClassVariavelCalculo
Dim dTotTrib As Double, iTotTribTipo As Integer, objVarOrigemProd As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto
Dim objFilial As ClassFilialCliente, objVarFilCliObj As New ClassVariavelCalculo
Dim iDocTotTribTipo As Integer, objVarDocTotTribTipo As New ClassVariavelCalculo
Dim dBaseItem As Double, dPerc As Double
Dim objVarTotTribFed As New ClassVariavelCalculo, objVarTotTribEst As New ClassVariavelCalculo, objVarTotTribMunic As New ClassVariavelCalculo
Dim dTotTribAliqFed As Double, dTotTribAliqEst As Double, dTotTribAliqMunic As Double
Dim dTotTribAliqFedAux As Double, dTotTribAliqEstAux As Double, dTotTribAliqMunicAux As Double

On Error GoTo Erro_AtualizarItem_SaidaCli

    dTotTrib = 0
    dTotTribAliqFed = 0
    dTotTribAliqEst = 0
    dTotTribAliqMunic = 0
    iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA

    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, iRecalculaAutomatica, objVarTotTrib, objVarTotTribTipo, objVarOrigemProd, objVarTotTribFed, objVarTotTribEst, objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If objVarTotTribFed.iStatus = VAR_PREENCH_MANUAL Then dTotTribAliqFed = objVarTotTribFed.vValor
    If objVarTotTribEst.iStatus = VAR_PREENCH_MANUAL Then dTotTribAliqEst = objVarTotTribEst.vValor
    If objVarTotTribMunic.iStatus = VAR_PREENCH_MANUAL Then dTotTribAliqMunic = objVarTotTribMunic.vValor
    If objVarTotTrib.iStatus = VAR_PREENCH_MANUAL Then dTotTrib = objVarTotTrib.vValor
    
    If objVarTotTribTipo.iStatus = VAR_PREENCH_MANUAL Then
        
        iTotTribTipo = objVarTotTribTipo.vValor
    
    Else
    
        lErro = objDoc.ObterVar("TOT_TRIB_TIPO", objVarDocTotTribTipo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        If objVarDocTotTribTipo.iStatus = VAR_PREENCH_MANUAL Then
            iTotTribTipo = objVarDocTotTribTipo.vValor
        Else
         
            If mvardtDataEmissao < DATA_INICIO_LEI_12741 Or gobjCRFAT.iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA Then
         
                iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA
         
            Else
         
                If objTipoDocInfo.iCodigo <> DOCINFO_NFISSAT And objTipoDocInfo.iCodigo <> DOCINFO_TNFISFV Then
                
                    lErro = objDoc.ObterVar("FILIAL_CLIENTE_OBJ", objVarFilCliObj)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                    Set objFilial = objVarFilCliObj.vValor
                    If Len(Trim(objFilial.sCgc)) <= STRING_CPF Or gobjCRFAT.iTotTribCNPJ = 1 Then
                        iTotTribTipo = gobjCRFAT.iTotTribTipo
                    Else
                        iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA
                    End If
            
                Else
                
                    iTotTribTipo = gobjCRFAT.iTotTribTipo
                
                End If
            
            End If
            
        End If
        
    End If
    
    If iTotTribTipo <> LEI_12741_TIPO_NAO_CALCULA Then
    
        lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "PIS", objDoc, objDocItem, dBaseItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
        Select Case iTotTribTipo
        
            Case LEI_12741_TIPO_SIMPLES
                lErro = Simples_ObterPerc(mvariFilialEmpresa, mvardtDataEmissao, dPerc)
                If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
        
            Case Else 'LEI_12741_TIPO_AUTOMATICO
        
                lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
                    
                    Set objProduto = objVarProd.vValor
                                    
                    If objVarOrigemProd.iStatus <> VAR_PREENCH_MANUAL Then
                
                        objVarOrigemProd.vValor = objProduto.iOrigemMercadoria
                        objVarOrigemProd.iStatus = VAR_PREENCH_AUTOMATICO
                        
                    End If
            
                    lErro = CF("IBPTax_ObterPerc", mvariFilialEmpresa, mvardtDataEmissao, objProduto, objVarOrigemProd.vValor, dPerc, mvarsUFFilialEmpresa, dTotTribAliqFedAux, dTotTribAliqEstAux, dTotTribAliqMunicAux)
                    If lErro <> SUCESSO And lErro <> ERRO_ITEM_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
                    
                    If lErro = SUCESSO Then
                    
                        If objVarTotTribFed.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqFed = dTotTribAliqFedAux
                        If objVarTotTribEst.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqEst = dTotTribAliqEstAux
                        If objVarTotTribMunic.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqMunic = dTotTribAliqMunicAux
                
                    End If
                
                End If
    
        End Select
    
        If lErro = ERRO_ITEM_NAO_CADASTRADO Then
        
            If objVarTotTribTipo.iStatus <> VAR_PREENCH_MANUAL Then iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA
            dPerc = 0
            If objVarTotTribFed.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqFed = 0
            If objVarTotTribEst.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqEst = 0
            If objVarTotTribMunic.iStatus <> VAR_PREENCH_MANUAL Then dTotTribAliqMunic = 0
            
        End If
        
        If objVarTotTrib.iStatus <> VAR_PREENCH_MANUAL Then
        
            If (dTotTribAliqFed <> 0 Or dTotTribAliqEst <> 0 Or dTotTribAliqMunic <> 0) Then
                dTotTrib = Arredonda_Moeda(Arredonda_Moeda(dBaseItem * dTotTribAliqFed, 2) + Arredonda_Moeda(dBaseItem * dTotTribAliqEst, 2) + Arredonda_Moeda(dBaseItem * dTotTribAliqMunic, 2), 2)
            Else
                dTotTrib = Arredonda_Moeda(dBaseItem * dPerc, 2)
            End If
    
        End If
    
    End If
    
Label_Atualizar:
    
    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTotTrib, dTotTrib, objVarTotTribTipo, iTotTribTipo, objVarTotTribFed, dTotTribAliqFed, objVarTotTribEst, dTotTribAliqEst, objVarTotTribMunic, dTotTribAliqMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    AtualizarItem_SaidaCli = SUCESSO
    
    Exit Function
    
Erro_AtualizarItem_SaidaCli:

    AtualizarItem_SaidaCli = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201510)

    End Select
    
    Exit Function

End Function

Private Function AtualizarDoc_Recalculo(objDoc As ClassTributoDoc) As Long
'marca os valores a nivel de doc como VAR_PREENCH_AUTOMATICO

Dim lErro As Long
Dim objVarTotTrib As New ClassVariavelCalculo, objVarTotTribTipo As New ClassVariavelCalculo
Dim objVarTotTribFed As New ClassVariavelCalculo, objVarTotTribEst As New ClassVariavelCalculo, objVarTotTribMunic As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc_Recalculo

    lErro = objDoc.ObterVar("TOT_TRIB", objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    objVarTotTrib.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_TIPO", objVarTotTribTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    objVarTotTribTipo.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarTotTribTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_FED", objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    objVarTotTribFed.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_EST", objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    objVarTotTribEst.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_MUNIC", objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    objVarTotTribMunic.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    AtualizarDoc_Recalculo = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_Recalculo:

    AtualizarDoc_Recalculo = gErr
     
    Select Case gErr
                  
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153862)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc) As Long
'atualiza TotTrib e TotTribTipo tanto a nivel de doc como de item
'    leva em conta dados dos itens

Dim lErro As Long, colItens As Collection, dTotTrib As Double, objVarTotTrib As New ClassVariavelCalculo, objVarTotTribItem As New ClassVariavelCalculo
Dim objDocItem As ClassTributoDocItem, dBaseItem As Double, dTotTribAux As Double
Dim dTotTribFed As Double, objVarTotTribFed As New ClassVariavelCalculo
Dim dTotTribEst As Double, objVarTotTribEst As New ClassVariavelCalculo
Dim dTotTribMunic As Double, objVarTotTribMunic As New ClassVariavelCalculo
Dim objVarTotTribItemFed As New ClassVariavelCalculo, objVarTotTribItemEst As New ClassVariavelCalculo, objVarTotTribItemMunic As New ClassVariavelCalculo
Dim dTotTribFedAux As Double, dTotTribEstAux As Double, dTotTribMunicAux As Double

On Error GoTo Erro_AtualizarDoc

    lErro = objDoc.ObterVar("TOT_TRIB", objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    lErro = objDoc.ObterVar("TOT_TRIB_FED", objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_EST", objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objDoc.ObterVar("TOT_TRIB_MUNIC", objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call objDoc.ObterColDocItens(colItens)
                
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
    
        lErro = objDocItem.ObterVar("TOT_TRIB", objVarTotTribItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dTotTribAux = dTotTribAux + objVarTotTribItem.vValor
                
        lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "PIS", objDoc, objDocItem, dBaseItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        lErro = objDocItem.ObterVar("TOT_TRIB_FED", objVarTotTribItemFed)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dTotTribFedAux = dTotTribFedAux + Arredonda_Moeda(objVarTotTribItemFed.vValor * dBaseItem)
                
        lErro = objDocItem.ObterVar("TOT_TRIB_EST", objVarTotTribItemEst)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dTotTribEstAux = dTotTribEstAux + Arredonda_Moeda(objVarTotTribItemEst.vValor * dBaseItem)
                
        lErro = objDocItem.ObterVar("TOT_TRIB_MUNIC", objVarTotTribItemMunic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        dTotTribMunicAux = dTotTribMunicAux + Arredonda_Moeda(objVarTotTribItemMunic.vValor * dBaseItem)
                
    Next
        
    If objVarTotTribFed.iStatus <> VAR_PREENCH_MANUAL Then
        dTotTribFed = Arredonda_Moeda(dTotTribFedAux, 2)
    Else
        dTotTribFed = objVarTotTribFed.vValor
    End If
    
    If objVarTotTribEst.iStatus <> VAR_PREENCH_MANUAL Then
        dTotTribEst = Arredonda_Moeda(dTotTribEstAux, 2)
    Else
        dTotTribEst = objVarTotTribEst.vValor
    End If
    
    If objVarTotTribMunic.iStatus <> VAR_PREENCH_MANUAL Then
        dTotTribMunic = Arredonda_Moeda(dTotTribMunicAux, 2)
    Else
        dTotTribMunic = objVarTotTribMunic.vValor
    End If
    
    If objVarTotTrib.iStatus <> VAR_PREENCH_MANUAL Then
        If (dTotTribFed <> 0 Or dTotTribEst <> 0 Or dTotTribMunic <> 0) And dTotTribAux = 0 Then
            dTotTrib = dTotTribFed + dTotTribEst + dTotTribMunic
        Else
            dTotTrib = dTotTribAux
        End If
    Else
        dTotTrib = objVarTotTrib.vValor
    End If
    
    dTotTrib = Arredonda_Moeda(dTotTrib, 2)

    lErro = AtualizarDoc_GuardarValores(objDoc, dTotTrib, dTotTribFed, dTotTribEst, dTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153678)

    End Select

    Exit Function

End Function

Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32025
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        mvarsUFFilialEmpresa = objFilialEmpresa.objEndereco.sSiglaEstado
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153679)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_GuardarValores(objDoc As ClassTributoDoc, ByVal dTotTrib As Double, ByVal dTotTribFed As Double, ByVal dTotTribEst As Double, ByVal dTotTribMunic As Double) As Long

Dim lErro As Long, objVarTotTrib As New ClassVariavelCalculo
Dim objVarTotTribFed As New ClassVariavelCalculo, objVarTotTribEst As New ClassVariavelCalculo, objVarTotTribMunic As New ClassVariavelCalculo

On Error GoTo Erro_AtualizarDoc_GuardarValores

    lErro = objDoc.ObterVar("TOT_TRIB", objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objVarTotTrib.iStatus <> VAR_PREENCH_MANUAL Then
        objVarTotTrib.vValor = dTotTrib
        objVarTotTrib.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarTotTrib)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = objDoc.ObterVar("TOT_TRIB_FED", objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objVarTotTribFed.iStatus <> VAR_PREENCH_MANUAL Then
        objVarTotTribFed.vValor = dTotTribFed
        objVarTotTribFed.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarTotTribFed)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = objDoc.ObterVar("TOT_TRIB_EST", objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objVarTotTribEst.iStatus <> VAR_PREENCH_MANUAL Then
        objVarTotTribEst.vValor = dTotTribEst
        objVarTotTribEst.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarTotTribEst)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    lErro = objDoc.ObterVar("TOT_TRIB_MUNIC", objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objVarTotTribMunic.iStatus <> VAR_PREENCH_MANUAL Then
        objVarTotTribMunic.vValor = dTotTribMunic
        objVarTotTribMunic.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarTotTribMunic)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    AtualizarDoc_GuardarValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_GuardarValores:

    AtualizarDoc_GuardarValores = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153680)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarDoc_ObterValores(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, objVarTotTrib As ClassVariavelCalculo, objVarTotTribTipo As ClassVariavelCalculo, objVarTotTribFed As ClassVariavelCalculo, objVarTotTribEst As ClassVariavelCalculo, objVarTotTribMunic As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_ObterValores

    lErro = objDoc.ObterVar("TOT_TRIB", objVarTotTrib)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTrib.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("TOT_TRIB_TIPO", objVarTotTribTipo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribTipo.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("TOT_TRIB_FED", objVarTotTribFed)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribFed.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("TOT_TRIB_EST", objVarTotTribEst)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribEst.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("TOT_TRIB_MUNIC", objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If iRecalculaAutomatica = 1 Then objVarTotTribMunic.iStatus = VAR_PREENCH_AUTOMATICO
    
    AtualizarDoc_ObterValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_ObterValores:

    AtualizarDoc_ObterValores = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153681)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarImposto_Zerar(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long

Dim colItens As Collection, objDocItem As ClassTributoDocItem
Dim lErro As Long

On Error GoTo Erro_AtualizarImposto_Zerar

    'se estiver recalculando marcar valores a nivel de doc c/preench auto.
    If iRecalculaAutomatica = 1 Then Call AtualizarDoc_Recalculo(objDoc)
            
    Call objDoc.ObterColDocItens(colItens)
                
    'percorrer os itens atualizando-os
    For Each objDocItem In colItens
                
        lErro = AtualizarItem_Zerar(objDoc, objDocItem, iRecalculaAutomatica)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

    '??? se todos os itens tiverem o mesmo tipo de calculo entao coloca-lo para o doc como um todo
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    AtualizarImposto_Zerar = SUCESSO
    
    Exit Function
    
Erro_AtualizarImposto_Zerar:

    AtualizarImposto_Zerar = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201509)

    End Select
    
    Exit Function

End Function

Private Function AtualizarItem_Zerar(objDoc As ClassTributoDoc, objDocItem As ClassTributoDocItem, iRecalculaAutomatica As Integer) As Long

Dim lErro As Long, objVarTotTrib As New ClassVariavelCalculo, objVarTotTribTipo As New ClassVariavelCalculo
Dim dTotTrib As Double, iTotTribTipo As Integer, objVarOrigemProd As New ClassVariavelCalculo
Dim objVarTotTribFed As New ClassVariavelCalculo, objVarTotTribEst As New ClassVariavelCalculo, objVarTotTribMunic As New ClassVariavelCalculo
Dim dTotTribAliqFed As Double, dTotTribAliqEst As Double, dTotTribAliqMunic As Double

On Error GoTo Erro_AtualizarItem_Zerar

    dTotTrib = 0
    iTotTribTipo = LEI_12741_TIPO_NAO_CALCULA
    
    lErro = AtualizarItem_ObterValoresAtuais(objDocItem, iRecalculaAutomatica, objVarTotTrib, objVarTotTribTipo, objVarOrigemProd, objVarTotTribFed, objVarTotTribEst, objVarTotTribMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = AtualizarItem_GuardarValores(objDocItem, objVarTotTrib, dTotTrib, objVarTotTribTipo, iTotTribTipo, objVarTotTribFed, dTotTribAliqFed, objVarTotTribEst, dTotTribAliqEst, objVarTotTribMunic, dTotTribAliqMunic)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    AtualizarItem_Zerar = SUCESSO
    
    Exit Function
    
Erro_AtualizarItem_Zerar:

    AtualizarItem_Zerar = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201510)

    End Select
    
    Exit Function

End Function

'Private Function IBPTax_ObterPerc(ByVal iFilialEmpresa As Integer, ByVal dtDataEmissao As Date, ByVal sProduto As String, ByVal iOrigemMercadoria As Integer, dPerc As Double) As Long
'
'Dim lErro As Long, alComando(0 To 2) As Long, lCodVersao As Long, sCampo As String
'Dim iIndice As Integer
'
'On Error GoTo Erro_IBPTax_ObterPerc
'
'    'Abertura de Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_Abrir()
'        If alComando(iIndice) = 0 Then gError 201521
'    Next
'
'    lErro = Comando_Executar(alComando(0), "SELECT MAX(CodVersao) FROM IBPTax WHERE Validade < ?", lCodVersao, dtDataEmissao)
'    If lErro <> AD_SQL_SUCESSO Then gError 201522
'
'    lErro = Comando_BuscarProximo(alComando(0))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201523
'    If lErro <> AD_SQL_SUCESSO Then gError ERRO_ITEM_NAO_CADASTRADO
'
'    Select Case iOrigemMercadoria
'
'        Case 0, 3, 4, 5
'            sCampo = "AliqNac"
'
'        Case Else
'            sCampo = "AliqImp"
'
'    End Select
'
'    lErro = Comando_Executar(alComando(1), "SELECT " & sCampo & " FROM IBPTaxAliq I, Produtos P WHERE P.IPICodigo LIKE I.NCM AND I.ExTIPI = P.EXTIPI AND I.Tabela = 0 AND P.Codigo = ? ORDER BY I.NCM DESC", dPerc, sProduto)
'    If lErro <> AD_SQL_SUCESSO Then gError 201524
'
'    lErro = Comando_BuscarProximo(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
'    If lErro <> AD_SQL_SUCESSO Then
'
'        lErro = Comando_Executar(alComando(2), "SELECT " & sCampo & " FROM IBPTaxAliq I, Produtos P WHERE P.ISSQN = I.NCM AND I.Tabela = 2 AND P.Codigo = ?", dPerc, sProduto)
'        If lErro <> AD_SQL_SUCESSO Then gError 201524
'
'        lErro = Comando_BuscarProximo(alComando(2))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
'        If lErro <> AD_SQL_SUCESSO Then gError ERRO_ITEM_NAO_CADASTRADO
'
'    End If
'
'    'Fecha Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    IBPTax_ObterPerc = SUCESSO
'
'    Exit Function
'
'Erro_IBPTax_ObterPerc:
'
'    IBPTax_ObterPerc = gErr
'
'    Select Case gErr
'
'        Case 201521
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case ERRO_ITEM_NAO_CADASTRADO
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201520)
'
'    End Select
'
'    'Fecha Comando
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    Exit Function
'
'End Function

Private Function Simples_ObterPerc(ByVal iFilialEmpresa As Integer, ByVal dtDataEmissao As Date, dPerc As Double) As Long

Dim lErro As Long, alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_Simples_ObterPerc

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201521
    Next
    
    lErro = Comando_Executar(alComando(1), "SELECT AliquotaTotal FROM DASAliquotas WHERE ((Ano= ? AND Mes <= ?) OR Ano < ?) AND FilialEmpresa = ? ORDER BY Ano DESC, Mes DESC", dPerc, Year(dtDataEmissao), Month(dtDataEmissao), Year(dtDataEmissao), iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 201524
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201525
    If lErro <> AD_SQL_SUCESSO Then gError ERRO_ITEM_NAO_CADASTRADO
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Simples_ObterPerc = SUCESSO
    
    Exit Function
    
Erro_Simples_ObterPerc:

    Simples_ObterPerc = gErr

    Select Case gErr

        Case 201521
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case ERRO_ITEM_NAO_CADASTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201520)

    End Select
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

