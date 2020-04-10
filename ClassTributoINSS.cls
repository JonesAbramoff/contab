VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoINSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ClassTributoDriver

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32252
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 32252
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 32252
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154018)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154019)

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
Private Function AtualizarDoc(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza INSSBase, INSSDeducoes, INSSValor, INSSValorRetencao
'    leva em conta dados dos itens
Dim colItens As Collection
Dim objDocItem As ClassTributoDocItem
Dim dValor As Double, dValorBase As Double, iRetido As Integer, dValorDeducoes As Double
Dim objTipoTrib As New ClassTipoDeTributacaoMovto
Dim objVarBase As New ClassVariavelCalculo, objVarValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim lErro As Long, objVar As New ClassVariavelCalculo, dAliquota As Double
Dim objTRBConfig As New ClassTRBConfig, dValorMinimoRetencao As Double
Dim objVarObjProduto As New ClassVariavelCalculo, objVarRetencao As New ClassVariavelCalculo
Dim objProduto As New ClassProduto, objVarDeducoes As New ClassVariavelCalculo
Dim objTributoDocCompl As ClassTributoDocCompl, colDocCompl As New Collection, objVarItemTipo As New ClassVariavelCalculo
Dim dBaseItem As Double

On Error GoTo Erro_AtualizarDoc

    lErro = AtualizarDoc_ObterValores(objDoc, iRecalculaAutomatica, objVarBase, objVarDeducoes, objVarValor, objVarRetencao)
    If lErro <> SUCESSO Then gError 81592
    
    iRetido = 0
    dValorBase = 0
    dValorDeducoes = 0
    dValor = 0
    dValorMinimoRetencao = 0
    
    'obter do doc o "TIPO_TRIB"
    lErro = objDoc.ObterVar("TIPO_TRIB", objVarTipoTrib)
    If lErro <> SUCESSO Then gError 81601
        
    'obter mais info sobre o tipo trib
    objTipoTrib.iTipo = objVarTipoTrib.vValor
    If objTipoTrib.iTipo <> 0 Then
    
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then gError 81602
        
        If objTipoTrib.iINSSIncide <> TRIBUTO_NAO_INCIDE Then
        
            dAliquota = -1
            
            Call objDoc.ObterColDocItens(colItens)
            
            'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
            For Each objDocItem In colItens
            
                'obter do item o "TIPO_TRIB"
                lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
                If lErro <> SUCESSO Then gError 81593
                
                'obter mais info sobre o tipo trib
                objTipoTrib.iTipo = objVarTipoTrib.vValor
                If objTipoTrib.iTipo <> 0 Then
                
                    lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
                    If lErro <> SUCESSO Then gError 81594
            
                    If objTipoTrib.iINSSIncide <> TRIBUTO_NAO_INCIDE Then

                        lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "INSS", objDoc, objDocItem, dBaseItem)
                        If lErro <> SUCESSO Then gError 81595
                        
                        objVarItemValor.vValor = dBaseItem
                        
                        lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarObjProduto)
                        If lErro <> SUCESSO Then gError 81596
                    
                        Set objProduto = objVarObjProduto.vValor
                         
                        dAliquota = objTipoTrib.dINSSAliquota
                        dValorBase = dValorBase + Arredonda_Moeda(objVarItemValor.vValor, 2)
                        dValorDeducoes = dValorDeducoes + Arredonda_Moeda(objVarItemValor.vValor * (1 - objProduto.dINSSPercBase), 2)
                        dValor = dValor + Arredonda_Moeda((dValorBase - dValorDeducoes) * dAliquota, 2)
                        dValorMinimoRetencao = objTipoTrib.dINSSRetencaoMinima
            
                    End If
            
                End If
            
            Next
                
        Else
        
            objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
            objVarDeducoes.iStatus = VAR_PREENCH_AUTOMATICO
            objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
            objVarRetencao.iStatus = VAR_PREENCH_AUTOMATICO
        
        End If
    
    End If
    
    If objVarBase.iStatus = VAR_PREENCH_MANUAL Then dValorBase = objVarBase.vValor
    If objVarDeducoes.iStatus = VAR_PREENCH_MANUAL Then dValorDeducoes = objVarDeducoes.vValor
    If objVarValor.iStatus = VAR_PREENCH_MANUAL Then dValor = objVarValor.vValor
    If objVarRetencao.iStatus = VAR_PREENCH_MANUAL Then
        iRetido = objVarRetencao.vValor
    Else
        
        If dValor = 0 Or dValor < dValorMinimoRetencao Then
            iRetido = 0
        Else
            iRetido = 1
        End If
        
    End If
    
    lErro = AtualizarDoc_GuardarValores(objDoc, dValorBase, dValorDeducoes, dValor, iRetido, objVarBase, objVarDeducoes, objVarValor, objVarRetencao)
    If lErro <> SUCESSO Then gError 81597
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = gErr

    Select Case gErr

        Case 81592 To 81597, 81601, 81602, 81609, 81610, 81611
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154020)

    End Select

    Exit Function

End Function

Private Function AtualizarDoc_GuardarValores(objDoc As ClassTributoDoc, dValorBase As Double, dValorDeducoes As Double, dValor As Double, iRetido As Integer, objVarBase As ClassVariavelCalculo, objVarDeducoes As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarRetencao As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_GuardarValores

    If objVarBase.iStatus <> VAR_PREENCH_MANUAL Then
        objVarBase.vValor = dValorBase
        objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarBase)
        If lErro <> SUCESSO Then gError 81585
    End If
    
    If objVarDeducoes.iStatus <> VAR_PREENCH_MANUAL Then
        objVarDeducoes.vValor = dValorDeducoes
        objVarDeducoes.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarDeducoes)
        If lErro <> SUCESSO Then gError 81587
    End If
    
    If objVarValor.iStatus <> VAR_PREENCH_MANUAL Then
        objVarValor.vValor = dValor
        objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarValor)
        If lErro <> SUCESSO Then gError 81589
    End If
    
    If objVarRetencao.iStatus <> VAR_PREENCH_MANUAL Then
        objVarRetencao.vValor = iRetido
        objVarRetencao.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarRetencao)
        If lErro <> SUCESSO Then gError 81591
    End If
    
    AtualizarDoc_GuardarValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_GuardarValores:

    AtualizarDoc_GuardarValores = gErr
     
    Select Case gErr
          
        Case 81585, 81587, 81589, 81591
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154021)
     
    End Select
     
    Exit Function

End Function

Private Function AtualizarDoc_ObterValores(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer, objVarBase As ClassVariavelCalculo, objVarDeducoes As ClassVariavelCalculo, objVarValor As ClassVariavelCalculo, objVarRetencao As ClassVariavelCalculo) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarDoc_ObterValores

    lErro = objDoc.ObterVar("INSS_BASE", objVarBase)
    If lErro <> SUCESSO Then gError 81580
    If iRecalculaAutomatica = 1 Then objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("INSS_VALOR_DEDUCOES", objVarDeducoes)
    If lErro <> SUCESSO Then gError 81581
    If iRecalculaAutomatica = 1 Then objVarDeducoes.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("INSS_VALOR", objVarValor)
    If lErro <> SUCESSO Then gError 81582
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("INSS_RETIDO", objVarRetencao)
    If lErro <> SUCESSO Then gError 81583
    If iRecalculaAutomatica = 1 Then objVarRetencao.iStatus = VAR_PREENCH_AUTOMATICO
        
    AtualizarDoc_ObterValores = SUCESSO
     
    Exit Function
    
Erro_AtualizarDoc_ObterValores:

    AtualizarDoc_ObterValores = gErr
     
    Select Case gErr
          
        Case 81580 To 81583
                  
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154022)
     
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
        
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179087)

    End Select

    Exit Function

End Function
