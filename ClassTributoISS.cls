VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'saidaCli: já foi feito
'entradacli: vou tratar como saida cli mas vou guardar a info se credita do imposto
'entradaforn: criar tratamento especifico (usar uforigem, padraotribentrada, )
'saidaforn: criar tratamento analogo ao de entradaforn mas nao vou ter info de credito

Implements ClassTributoDriver

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvariContribuinte As Integer 'se tem ou nao inscr municipal
Private mvardPercPadrao As Double
Private mvariISSIncluso As Integer
Private mvariSimplesNacional As Integer
Private sCidadeIBGECodigoEmp As String

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then Error 32028
    
    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then Error 27731
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = Err

    Select Case Err

        Case 27731, 32028
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154161)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
'permite a inicializacao do driver
'pode ser utilizado p/criar cache etc
Dim lErro As Long, objFilialEmpresa As New AdmFiliais
On Error GoTo Erro_ClassTributoDriver_Ativar

    ClassTributoDriver_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDriver_Ativar:

    ClassTributoDriver_Ativar = Err

    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154162)

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

Private Function AtualizarItem_SemISS(objDocItem As ClassTributoDocItem) As Long

Dim objVar As New ClassVariavelCalculo, lErro As Long

    objVar.sIdentificador = "ISSQN"
    objVar.vValor = ""
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_TIPO"
    objVar.vValor = 1
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_CST"
    objVar.vValor = "N"
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_BASE"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_ALIQUOTA"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_CIDADE_IBGE"
    objVar.vValor = ""
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
                       
    'nfe 3.10
    objVar.sIdentificador = "ISS_COD_SERV"
    objVar.vValor = ""
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR_DEDUCAO"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR_OUTRASRET"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR_DESC_INCOND"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR_DESC_COND"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_VALOR_RET"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745

    objVar.sIdentificador = "ISS_MUNIC_INCID_IMP"
    objVar.vValor = ""
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    
    objVar.sIdentificador = "ISS_IND_EXIGIBILIDADE"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745

    objVar.sIdentificador = "ISS_COD_PAIS"
    objVar.vValor = 0
    objVar.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDocItem.GuardarVar(objVar)
    If lErro <> SUCESSO Then Error 27745
    'fim nfe 3.10
                   
End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'atualiza ISSBase, ISSAliquota, ISSValor
'    leva em conta dados dos itens
Dim colItens As Collection
Dim objDocItem As ClassTributoDocItem
Dim dValor As Double, dBase As Double, dAliquota As Double, dtDataEmissao As Date, objVarDataEmissao As New ClassVariavelCalculo
Dim objTipoTrib As New ClassTipoDeTributacaoMovto, objVarISSIncluso As New ClassVariavelCalculo
Dim objVarBase As New ClassVariavelCalculo, objVarValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim lErro As Long, objVarAliquota As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, objTipoDocInfo As New ClassTipoDocInfo, dISSRetido As Double
Dim objVarTipoDocInfo As New ClassVariavelCalculo, objVarISSRetido As New ClassVariavelCalculo
Dim objVarItemBase As New ClassVariavelCalculo
Dim objVarItemAliq As New ClassVariavelCalculo, objVarItemCST As New ClassVariavelCalculo
Dim objVarItemVlr As New ClassVariavelCalculo, objVarItemTipo As New ClassVariavelCalculo
Dim dBaseItem As Double, dPercExcecao As Double, objVarItemTipoCST As New ClassVariavelCalculo
Dim objTipoISS As ClassTipoTribISS, objVarItemTipoObj As New ClassVariavelCalculo, bAtualiza As Boolean
Dim objVarItemISSQN As New ClassVariavelCalculo

Dim objVarProd As New ClassVariavelCalculo, objProduto As ClassProduto

'nfe 3.10
Dim dISSValorDeducao As Double, objVarISSValorDeducao As New ClassVariavelCalculo
Dim dISSValorOutrasRet As Double, objVarISSValorOutrasRet As New ClassVariavelCalculo
Dim dISSValorDescIncond As Double, objVarISSValorDescIncond As New ClassVariavelCalculo
Dim dISSValorDescCond As Double, objVarISSValorDescCond As New ClassVariavelCalculo
Dim dtDataPrestServico As Date, objVarDataPrestServico As New ClassVariavelCalculo

Dim objVarISSValorDeducaoItem As New ClassVariavelCalculo, objVarItemDesc As New ClassVariavelCalculo
Dim objVarISSValorOutrasRetItem As New ClassVariavelCalculo
Dim objVarISSValorDescIncondItem As New ClassVariavelCalculo
Dim objVarISSValorDescCondItem As New ClassVariavelCalculo, objVarISSIndExigibilidade As New ClassVariavelCalculo
Dim objVarISSValorRetItem As New ClassVariavelCalculo, dISSValorRetItem As Double, objVarISSCodServ As New ClassVariavelCalculo
Dim objVarFilCliCidadeIBGE As New ClassVariavelCalculo, objVarItemISSCodPais As New ClassVariavelCalculo
Dim objVarItemCidadeIBGE As New ClassVariavelCalculo, objVarMunicIncideImp As New ClassVariavelCalculo, sCodTribMun As String
Dim objVarMunicFaturamento As New ClassVariavelCalculo
Dim bSimplesNacionalInterna As Boolean, bCalculaISSMesmoNoSimples As Boolean, sConteudo As String
'fim nfe 3.10

On Error GoTo Erro_AtualizarDoc

    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then Error 22777
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then Error 22778
    
    lErro = objDoc.ObterVar("ISS_INCLUSO", objVarISSIncluso)
    If lErro <> SUCESSO Then Error 27747
    If iRecalculaAutomatica = 1 Then objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("ISS_ALIQUOTA", objVarAliquota)
    If lErro <> SUCESSO Then Error 27737
    If iRecalculaAutomatica = 1 Then objVarAliquota.iStatus = VAR_PREENCH_AUTOMATICO
    
    lErro = objDoc.ObterVar("ISS_BASE", objVarBase)
    If lErro <> SUCESSO Then Error 27739
    If iRecalculaAutomatica = 1 Then objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("ISS_VALOR", objVarValor)
    If lErro <> SUCESSO Then Error 27746
    If iRecalculaAutomatica = 1 Then objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("ISS_RETIDO", objVarISSRetido)
    If lErro <> SUCESSO Then Error 27746
    If iRecalculaAutomatica = 1 Then objVarISSRetido.iStatus = VAR_PREENCH_AUTOMATICO
           
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then Error 27746
    If iRecalculaAutomatica = 1 Then objVarDataEmissao.iStatus = VAR_PREENCH_AUTOMATICO
    
    dtDataEmissao = objVarDataEmissao.vValor
           
    'nfe 3.10
    objVarMunicFaturamento.sIdentificador = "ISS_MUNIC_FATURAMENTO"
    If objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
        objVarMunicFaturamento.vValor = sCidadeIBGECodigoEmp
    Else
        objVarMunicFaturamento.vValor = ""
    End If
    objVarMunicFaturamento.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarMunicFaturamento)
    If lErro <> SUCESSO Then Error 27324
                   
    lErro = objDoc.ObterVar("DATA_PREST_SERVICO", objVarDataPrestServico)
    If lErro <> SUCESSO Then Error 27746
    If iRecalculaAutomatica = 1 Then objVarDataPrestServico.iStatus = VAR_PREENCH_AUTOMATICO
    
    If objVarDataPrestServico.iStatus <> VAR_PREENCH_MANUAL Then
        
        objVarDataPrestServico.vValor = dtDataEmissao
        objVarDataPrestServico.iStatus = VAR_PREENCH_AUTOMATICO
        lErro = objDoc.GuardarVar(objVarDataPrestServico)
        If lErro <> SUCESSO Then Error 27746
    
    End If
    'fim nfe 3.10
    
    bSimplesNacionalInterna = False
    bCalculaISSMesmoNoSimples = False
    
    'se está no Simples Nacional
    If objTipoDocInfo.iEmitente = EMITENTE_EMPRESA And dtDataEmissao >= DATA_INICIO_SIMPLES_NACIONAL And mvariSimplesNacional <> 0 Then
    
        bSimplesNacionalInterna = True
   
        lErro = CF("Config_Le", "FATConfig", "ISS_CALCULA_MESMO_NO_SIMPLES", giFilialEmpresa, sConteudo)
        If lErro <> SUCESSO And lErro <> 208279 Then gError ERRO_SEM_MENSAGEM
        
        If lErro = SUCESSO And StrParaInt(sConteudo) = MARCADO Then
            bCalculaISSMesmoNoSimples = True
        End If
   
    End If
    
    'para as notas da propria empresa
    If objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
     
         If objVarISSIncluso.iStatus <> VAR_PREENCH_MANUAL Then
             
             If mvariISSIncluso <> objVarISSIncluso.vValor Then
         
                 objVarISSIncluso.vValor = mvariISSIncluso
             
                 objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
                 lErro = objDoc.GuardarVar(objVarISSIncluso)
                 If lErro <> SUCESSO Then Error 27745
             
             End If
             
         End If
     
    Else 'para notas de fornecedores
     
         '??? enquanto nao coloco tratamento especifico p/ISS p/contas a pagar
         If objVarISSIncluso.vValor <> 1 Then
             
             objVarISSIncluso.vValor = 1 'incluso
             objVarISSIncluso.iStatus = VAR_PREENCH_AUTOMATICO
             lErro = objDoc.GuardarVar(objVarISSIncluso)
             If lErro <> SUCESSO Then Error 27745
         
         End If
     
    End If
    
    dBase = 0
    dValor = 0
    
    'nfe 3.10
    dISSValorDeducao = 0
    dISSValorOutrasRet = 0
    dISSValorDescIncond = 0
    dISSValorDescCond = 0
    'fim nfe 3.10

    Call objDoc.ObterColDocItens(colItens)
    
    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
    
        'obter do item "TIPO_TRIB" e "PRODUTO_VALOR"
        lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then Error 27740
        
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then Error 27743
        
        bAtualiza = True
        If gobjCRFAT.iOtimizaTrib = MARCADO Then
            If objDocItem.iJaCalculado = MARCADO Then bAtualiza = False
        End If
        
        If bAtualiza Then
    
            If objTipoTrib.iISSIncide = TRIBUTO_NAO_INCIDE Then
            
                lErro = AtualizarItem_SemISS(objDocItem)
                If lErro <> SUCESSO Then Error 27324
            
            End If
            
            lErro = objDocItem.ObterVar("ISS_TIPO", objVarItemTipo)
            If lErro <> SUCESSO Then Error 27324
            If objVarItemTipo.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemTipo.vValor = objTipoTrib.iISSTipo
                If objTipoTrib.iISSTipo = 0 Then objVarItemTipo.vValor = 1
                objVarItemTipo.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemTipo)
                If lErro <> SUCESSO Then Error 27324
            End If
            
            Set objTipoISS = New ClassTipoTribISS
            objTipoISS.iTipo = objVarItemTipo.vValor
            Set objVarItemTipoObj.vValor = objTipoISS
            lErro = objDocItem.objDoc.ObterVar("ISS_TIPO_OBJ", objVarItemTipoObj)
            If lErro <> SUCESSO Then Error 27874
            Set objTipoISS = objVarItemTipoObj.vValor
            
            If objTipoTrib.iISSIncide <> TRIBUTO_NAO_INCIDE Then
            
                lErro = objDocItem.ObterVar("PRODUTO_OBJ", objVarProd)
                If lErro <> SUCESSO Then Error 27324
                            
                If objVarProd.iStatus <> VAR_PREENCH_VAZIO Then
                
                    Set objProduto = objVarProd.vValor
                    
                Else
                
                    Set objProduto = Nothing
                
                End If
                
                'nfe 3.10
                lErro = objDocItem.ObterVar("ISS_IND_EXIGIBILIDADE", objVarISSIndExigibilidade)
                If lErro <> SUCESSO Then Error 27324
                If objVarISSIndExigibilidade.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                
                    objVarISSIndExigibilidade.vValor = objTipoTrib.iISSIndExigibilidade
                    objVarISSIndExigibilidade.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarISSIndExigibilidade)
                    If lErro <> SUCESSO Then Error 27324
                
                End If
                'fim nfe 3.10
                
                lErro = objDocItem.ObterVar("ISSQN", objVarItemISSQN)
                If lErro <> SUCESSO Then Error 27324
                If objVarItemISSQN.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                    objVarItemISSQN.vValor = IIf(objProduto Is Nothing, "", objProduto.sISSQN)
                    objVarItemISSQN.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarItemISSQN)
                    If lErro <> SUCESSO Then Error 27324
                End If
                
                lErro = objDocItem.ObterVar("ISS_COD_PAIS", objVarItemISSCodPais)
                If lErro <> SUCESSO Then Error 27324
                If objVarItemISSCodPais.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
               
                   objVarItemISSCodPais.vValor = 0
                   objVarItemISSCodPais.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarItemISSCodPais)
                   If lErro <> SUCESSO Then Error 27324
               
                End If
                
               lErro = objDocItem.ObterVar("ISS_CIDADE_IBGE", objVarItemCidadeIBGE)
               If lErro <> SUCESSO Then Error 27324
               If objVarItemCidadeIBGE.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                    If objVarItemISSCodPais.vValor = 0 And objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then

                        If objProduto Is Nothing Then
                        
                            objVarItemCidadeIBGE.vValor = sCidadeIBGECodigoEmp
                                               
                        Else
                        
                            '*****************************
                            'Preenche o código do IBGE da cidade onde o serviço foi realizado
                            'com a cidade do cliente ou a cidade da empresa
                            'de acordo com a configuração do produto
                            lErro = CF("ProdutoCNAE_Le", objProduto)
                            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 27324
                            
                            If lErro = SUCESSO Then
                            
                                If objProduto.objProdutoCNAE.iLocServCliente = MARCADO Then
                                
                                    'obter codigo do ibge da cidade do cliente
                                    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CIDADE_IBGE", objVarFilCliCidadeIBGE)
                                    If lErro <> SUCESSO Then Error 27324
                                    
                                    objVarItemCidadeIBGE.vValor = objVarFilCliCidadeIBGE.vValor
                                    
                                Else
                                    objVarItemCidadeIBGE.vValor = sCidadeIBGECodigoEmp
                                End If
                                
                            Else
                                objVarItemCidadeIBGE.vValor = sCidadeIBGECodigoEmp
                            End If
                            '*****************************
                        
                        End If
              
                    Else
                
                        objVarItemCidadeIBGE.vValor = ""
                        
                    End If
                    
                    objVarItemCidadeIBGE.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarItemCidadeIBGE)
                    If lErro <> SUCESSO Then Error 27324
              
               End If
                
               'nfe 3.10
               
               lErro = objDocItem.ObterVar("ISS_MUNIC_INCID_IMP", objVarMunicIncideImp)
               If lErro <> SUCESSO Then Error 27324
               If objVarMunicIncideImp.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
               
                    objVarMunicIncideImp.vValor = objVarItemCidadeIBGE.vValor
                    If objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
                        If Not (objProduto Is Nothing) Then
                            '*****************************
                            'Preenche o código do IBGE da cidade da incidência do imposto
                            'com a cidade do cliente ou a cidade da empresa
                            'de acordo com a configuração do produto
                            lErro = CF("ProdutoCNAE_Le", objProduto)
                            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 27324
                            
                            If lErro = SUCESSO Then
                            
                                If objProduto.objProdutoCNAE.iLocIncidImpCliente = MARCADO Then
                                
                                    'obter codigo do ibge da cidade do cliente
                                    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CIDADE_IBGE", objVarFilCliCidadeIBGE)
                                    If lErro <> SUCESSO Then Error 27324
                                    
                                    objVarMunicIncideImp.vValor = objVarFilCliCidadeIBGE.vValor
                                    
                                Else
                                    objVarMunicIncideImp.vValor = sCidadeIBGECodigoEmp
                                End If
                                
                            End If
                            '*****************************
                        End If
                    End If

                    'objVarMunicIncideImp.vValor = objVarItemCidadeIBGE.vValor
                    objVarMunicIncideImp.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarMunicIncideImp)
                    If lErro <> SUCESSO Then Error 27324
              
               End If
               
               lErro = objDocItem.ObterVar("ISS_COD_SERV", objVarISSCodServ)
               If lErro <> SUCESSO Then Error 27324
               If objVarISSCodServ.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                    sCodTribMun = ""
                    If Not (objProduto Is Nothing) And objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
                    
                        lErro = CF("Servico_ObterCodTribMun", objProduto.sCodigo, sCidadeIBGECodigoEmp, sCodTribMun)
                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 27324
                    
                    End If
                    
                    objVarISSCodServ.vValor = sCodTribMun
                    objVarISSCodServ.iStatus = VAR_PREENCH_AUTOMATICO
                    lErro = objDocItem.GuardarVar(objVarISSCodServ)
                    If lErro <> SUCESSO Then Error 27324
              
               End If
               'fim nfe 3.10
               
                lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "ISS", objDoc, objDocItem, dBaseItem)
                If lErro <> SUCESSO Then Error 27744
                
                objVarItemValor.vValor = dBaseItem
                
                lErro = objDocItem.ObterVar("ISS_BASE", objVarItemBase)
                If lErro <> SUCESSO Then Error 27324
                If objVarItemBase.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                    
                    If bSimplesNacionalInterna And Not bCalculaISSMesmoNoSimples Then
                    
                        objVarItemBase.vValor = 0
                    
                    Else
                    
                        objVarItemBase.vValor = objVarItemValor.vValor
                        
                    End If
                    
                End If
                
                lErro = objDocItem.ObterVar("ISS_ALIQUOTA", objVarItemAliq)
                If lErro <> SUCESSO Then Error 27324
                If objVarItemAliq.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                   
                   If objVarItemBase.vValor <> 0 Or (bSimplesNacionalInterna And objTipoISS.iComRetencao <> DESMARCADO) Then
                   
                        If bSimplesNacionalInterna And Not bCalculaISSMesmoNoSimples Then
                        
                            If objTipoISS.iComRetencao <> DESMARCADO Then
                            
                                dPercExcecao = 0
                                lErro = Excecao_Pesquisar(objDocItem, mvariFilialEmpresa, dPercExcecao)
                                'If lErro <> SUCESSO Then Error 27324
                                'Se não achou alíquota por produto pega o padrão
                                If lErro <> SUCESSO Or dPercExcecao = 0 Then
                                
                                    'obter aliquota do iss do simples (deveria pegar de DASAliquotas mas vou simplificar)
                                    objVarItemAliq.vValor = Arredonda_Moeda(mvardPercPadrao, 4) 'ex.: 2,79% = 0,0279
                            
                                Else
                                    
                                    objVarItemAliq.vValor = Arredonda_Moeda(dPercExcecao, 4)
                                
                                End If
                            
                            Else
                            
                                objVarItemAliq.vValor = 0
                                
                            End If
                        
                        Else
                   
                            dPercExcecao = 0
                            lErro = Excecao_Pesquisar(objDocItem, mvariFilialEmpresa, dPercExcecao)
                            'If lErro <> SUCESSO Then Error 27324
                            'Se não achou alíquota por produto pega o padrão
                            If lErro <> SUCESSO Or dPercExcecao = 0 Then
                                objVarItemAliq.vValor = mvardPercPadrao
                            Else
                                objVarItemAliq.vValor = dPercExcecao
                            End If
                            
                        End If
                   
                   Else
                        objVarItemAliq.vValor = 0
                   End If
                   
                   objVarItemAliq.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarItemAliq)
                   If lErro <> SUCESSO Then Error 27324
                   
               End If
               
               If objVarItemBase.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                
                   'se o iss nao está incluso no valor dos produtos e deve ser incorporado à base de calculo
                   If objVarISSIncluso.vValor = 0 And gobjCRFAT.iISSNaBase <> 0 Then
                       If 1 <> objVarItemAliq.vValor Then
                           objVarItemBase.vValor = Arredonda_Moeda(objVarItemBase.vValor / (1 - objVarItemAliq.vValor), 2)
                       End If
                   End If
                   objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarItemBase)
                   If lErro <> SUCESSO Then Error 27324
                   
               End If
                              
               lErro = objDocItem.ObterVar("ISS_VALOR", objVarItemVlr)
               If lErro <> SUCESSO Then Error 27324
               If objVarItemVlr.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                   objVarItemVlr.vValor = objVarItemBase.vValor * objVarItemAliq.vValor
                   
                   objVarItemVlr.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarItemVlr)
                   If lErro <> SUCESSO Then Error 27324
              
               End If
                
               'nfe 3.10
               lErro = objDocItem.ObterVar("ISS_VALOR_DEDUCAO", objVarISSValorDeducaoItem)
               If lErro <> SUCESSO Then Error 27324
               If objVarISSValorDeducaoItem.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                   objVarISSValorDeducaoItem.vValor = 0
                   
                   objVarISSValorDeducaoItem.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarISSValorDeducaoItem)
                   If lErro <> SUCESSO Then Error 27324
              
               End If
               dISSValorDeducao = dISSValorDeducao + objVarISSValorDeducaoItem.vValor
               
               lErro = objDocItem.ObterVar("ISS_VALOR_OUTRASRET", objVarISSValorOutrasRetItem)
               If lErro <> SUCESSO Then Error 27324
               If objVarISSValorOutrasRetItem.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                   objVarISSValorOutrasRetItem.vValor = 0
                   objVarISSValorOutrasRetItem.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarISSValorOutrasRetItem)
                   If lErro <> SUCESSO Then Error 27324
              
               End If
               dISSValorOutrasRet = dISSValorOutrasRet + objVarISSValorOutrasRetItem.vValor
               
               lErro = objDocItem.ObterVar("ISS_VALOR_DESC_INCOND", objVarISSValorDescIncondItem)
               If lErro <> SUCESSO Then Error 27324
               If objVarISSValorDescIncondItem.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                    'Obtem o desconto
                    lErro = objDocItem.ObterVar("VALOR_DESCONTO_GRID", objVarItemDesc)
                    If lErro <> SUCESSO Then Error 27324

                   objVarISSValorDescIncondItem.vValor = objVarItemDesc.vValor
                   objVarISSValorDescIncondItem.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarISSValorDescIncondItem)
                   If lErro <> SUCESSO Then Error 27324
              
               End If
               dISSValorDescIncond = dISSValorDescIncond + objVarISSValorDescIncondItem.vValor
               
               lErro = objDocItem.ObterVar("ISS_VALOR_DESC_COND", objVarISSValorDescCondItem)
               If lErro <> SUCESSO Then Error 27324
               If objVarISSValorDescCondItem.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then

                   objVarISSValorDescCondItem.vValor = 0
                   objVarISSValorDescCondItem.iStatus = VAR_PREENCH_AUTOMATICO
                   lErro = objDocItem.GuardarVar(objVarISSValorDescCondItem)
                   If lErro <> SUCESSO Then Error 27324
              
               End If
               dISSValorDescCond = dISSValorDescCond + objVarISSValorDescCondItem.vValor
               'fim nfe 3.10
                
                dBase = dBase + objVarItemBase.vValor
                dValor = dValor + objVarItemVlr.vValor
            
            End If
        
            lErro = objDocItem.ObterVar("ISS_VALOR_RET", objVarISSValorRetItem)
            If lErro <> SUCESSO Then Error 27324
            If objVarISSValorRetItem.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
            
                If objTipoISS.iComRetencao <> DESMARCADO Then
                
                    If bSimplesNacionalInterna Then
                    
                        dISSValorRetItem = Arredonda_Moeda(dBaseItem * objVarItemAliq.vValor)
                    
                    Else
                    
                        If objVarISSIncluso.vValor <> 1 Then
                            dISSValorRetItem = Arredonda_Moeda((objVarItemBase.vValor + objVarItemVlr.vValor) * objVarItemAliq.vValor)
                        Else
                            
                            dISSValorRetItem = Arredonda_Moeda(objVarItemVlr.vValor)
                        
                        End If
                    
                    End If
                    
                Else
                
                    dISSValorRetItem = 0
                
                End If
        
                objVarISSValorRetItem.vValor = dISSValorRetItem
                objVarISSValorRetItem.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarISSValorRetItem)
                If lErro <> SUCESSO Then Error 27324
            
            Else
                
                dISSValorRetItem = objVarISSValorRetItem.vValor
            
            End If
            
            dISSRetido = dISSRetido + dISSValorRetItem
        
            lErro = objDocItem.ObterVar("ISS_CST", objVarItemCST)
            If lErro <> SUCESSO Then Error 27745
            objVarItemCST.vValor = objTipoISS.sTipoTribCST
            objVarItemCST.iStatus = VAR_PREENCH_AUTOMATICO
            lErro = objDocItem.GuardarVar(objVarItemCST)
            If lErro <> SUCESSO Then Error 27745
            
        Else
            'Só lê o necessário para acumular os valores usados depois
            lErro = objDocItem.ObterVar("ISS_TIPO", objVarItemTipo)
            If lErro <> SUCESSO Then Error 27324
            
            Set objTipoISS = New ClassTipoTribISS
            objTipoISS.iTipo = objVarItemTipo.vValor
            Set objVarItemTipoObj.vValor = objTipoISS
            lErro = objDocItem.objDoc.ObterVar("ISS_TIPO_OBJ", objVarItemTipoObj)
            If lErro <> SUCESSO Then Error 27874
            Set objTipoISS = objVarItemTipoObj.vValor
            
            lErro = objDocItem.ObterVar("ISS_VALOR", objVarItemVlr)
            If lErro <> SUCESSO Then Error 27324
        
            lErro = objDocItem.ObterVar("ISS_BASE", objVarItemBase)
            If lErro <> SUCESSO Then Error 27324
        
            lErro = objDocItem.ObterVar("ISS_ALIQUOTA", objVarItemAliq)
            If lErro <> SUCESSO Then Error 27324
        
            'nfe 3.10
            lErro = objDocItem.ObterVar("ISS_VALOR_DEDUCAO", objVarISSValorDeducaoItem)
            If lErro <> SUCESSO Then Error 27324
            dISSValorDeducao = dISSValorDeducao + objVarISSValorDeducaoItem.vValor
             
            lErro = objDocItem.ObterVar("ISS_VALOR_OUTRASRET", objVarISSValorOutrasRetItem)
            If lErro <> SUCESSO Then Error 27324
            dISSValorOutrasRet = dISSValorOutrasRet + objVarISSValorOutrasRetItem.vValor
             
            lErro = objDocItem.ObterVar("ISS_VALOR_DESC_INCOND", objVarISSValorDescIncondItem)
            If lErro <> SUCESSO Then Error 27324
            dISSValorDescIncond = dISSValorDescIncond + objVarISSValorDescIncondItem.vValor
             
            lErro = objDocItem.ObterVar("ISS_VALOR_DESC_COND", objVarISSValorDescCondItem)
            If lErro <> SUCESSO Then Error 27324
            dISSValorDescCond = dISSValorDescCond + objVarISSValorDescCondItem.vValor
            
            lErro = objDocItem.ObterVar("ISS_VALOR_RET", objVarISSValorRetItem)
            If lErro <> SUCESSO Then Error 27324
            dISSRetido = dISSRetido + objVarISSValorRetItem.vValor
            'fim nfe 3.10
            
            dBase = dBase + objVarItemBase.vValor
            dValor = dValor + objVarItemVlr.vValor
        
        End If
    
    Next

    dValor = CDbl(Format(dValor, "0.00"))
         
    objVarBase.vValor = dBase
    objVarBase.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarBase)
    If lErro <> SUCESSO Then Error 27745
         
    objVarValor.vValor = dValor
    objVarValor.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarValor)
    If lErro <> SUCESSO Then Error 27748
     
    If objVarISSRetido.iStatus <> VAR_PREENCH_MANUAL Then
     
         objVarISSRetido.vValor = dISSRetido
         objVarISSRetido.iStatus = VAR_PREENCH_AUTOMATICO
         lErro = objDoc.GuardarVar(objVarISSRetido)
         If lErro <> SUCESSO Then Error 27748
         
     End If

'    End If
    
    'nfe 3.10
    objVarISSValorDeducao.sIdentificador = "ISS_VALOR_DEDUCAO"
    objVarISSValorDeducao.vValor = dISSValorDeducao
    objVarISSValorDeducao.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarISSValorDeducao)
    If lErro <> SUCESSO Then Error 27746
    
    objVarISSValorOutrasRet.sIdentificador = "ISS_VALOR_OUTRAS_RET"
    objVarISSValorOutrasRet.vValor = dISSValorOutrasRet
    objVarISSValorOutrasRet.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarISSValorOutrasRet)
    If lErro <> SUCESSO Then Error 27746
    
    objVarISSValorDescIncond.sIdentificador = "ISS_VALOR_DESC_INCOND"
    objVarISSValorDescIncond.vValor = dISSValorDescIncond
    objVarISSValorDescIncond.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarISSValorDescIncond)
    If lErro <> SUCESSO Then Error 27746
    
    objVarISSValorDescCond.sIdentificador = "ISS_VALOR_DESC_COND"
    objVarISSValorDescCond.vValor = dISSValorDescCond
    objVarISSValorDescCond.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarISSValorDescCond)
    If lErro <> SUCESSO Then Error 27746
    'fim nfe 3.10
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = Err

    Select Case Err

        Case 27737 To 27748, 27773 To 27776, 22777, 22778
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154163)

    End Select

    Exit Function

End Function

Private Function TestaFilialEmpresa(objDoc As ClassTributoDoc) As Long

Dim lErro As Long, objFilialEmpresa As AdmFiliais
Dim objVarFilialFat As New ClassVariavelCalculo, objCidade As New ClassCidades
    
On Error GoTo Erro_TestaFilialEmpresa

    lErro = objDoc.ObterVar("FILIAL_FAT_OBJ", objVarFilialFat)
    If lErro <> SUCESSO Then Error 32027
    
    Set objFilialEmpresa = objVarFilialFat.vValor
    
    'se trocou a filial de faturamento
    If objFilialEmpresa.iCodFilial <> mvariFilialEmpresa Then
    
        If objFilialEmpresa.sInscricaoMunicipal <> "" Then
            mvariContribuinte = CONTRIBUINTE_ISS
            mvardPercPadrao = objFilialEmpresa.dISSPercPadrao
        Else
            mvariContribuinte = NAO_CONTRIBUINTE_ISS
        End If
        
        If mvardPercPadrao = 0 Then mvardPercPadrao = 0.05
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        mvariISSIncluso = objFilialEmpresa.iISSIncluso
        mvariSimplesNacional = objFilialEmpresa.iSuperSimples
        
        objCidade.sDescricao = objFilialEmpresa.objEndereco.sCidade
        
        lErro = CF("Cidade_Le_Nome", objCidade)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then Error 32027
        
        sCidadeIBGECodigoEmp = objCidade.sCodIBGE
    
    End If
    
    TestaFilialEmpresa = SUCESSO
    
    Exit Function
    
Erro_TestaFilialEmpresa:

    TestaFilialEmpresa = Err
    
    Select Case Err

        Case 32027
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154164)

    End Select

    Exit Function

End Function

Private Function Excecao_Pesquisar(objDocItem As ClassTributoDocItem, ByVal iFilialEmpresa As Integer, dPercExcecao As Double) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objVarProduto As New ClassVariavelCalculo
Dim objCidade As New ClassCidades
Dim objCodTrib As ClassCodTribMun
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Excecao_Pesquisar

    dPercExcecao = 0

    'obter a colecao de categorias da filialcliente
    lErro = objDocItem.ObterVar("PRODUTO_CODIGO", objVarProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objProduto.sCodigo = objVarProduto.vValor
    
    lErro = CF("ProdutoCNAE_Le", objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro = SUCESSO Then
        
        If objProduto.objProdutoCNAE.colCidades.Count > 0 Then
       
            objFilialEmpresa.iCodFilial = giFilialEmpresa
            
            'Le o nome da filial
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
                    
            objCidade.sDescricao = objFilialEmpresa.objEndereco.sCidade
            
            lErro = CF("Cidade_Le_Nome", objCidade)
            If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM

            For Each objCodTrib In objProduto.objProdutoCNAE.colCidades
                If objCodTrib.lCidade = objCidade.iCodigo Then
                    dPercExcecao = objCodTrib.dAliquota
                    Exit For
                End If
            Next

        End If
        
    End If
    
    Excecao_Pesquisar = SUCESSO

    Exit Function

Erro_Excecao_Pesquisar:

    Excecao_Pesquisar = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153831)

    End Select

    Exit Function

End Function
