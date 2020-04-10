VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributoCOFINS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ClassTributoDriver

Private mvariFilialEmpresa As Integer 'guarda a ultima filialempresa utilizada
Private mvariNaoCumulativo As Integer
'Private mvariSimplesFederal As Integer
Private mvariSimplesNacional As Integer

Private Function ClassTributoDriver_AtualizarImposto(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long
'recalcula a tributacao de todos os itens de um doc e dele como um todo

Dim lErro As Long

On Error GoTo Erro_ClassTributoDriver_AtualizarImposto

    'atualizar o docto como um todo
    lErro = AtualizarDoc(objDoc, iRecalculaAutomatica)
    If lErro <> SUCESSO Then gError 106538
    
    ClassTributoDriver_AtualizarImposto = SUCESSO

    Exit Function

Erro_ClassTributoDriver_AtualizarImposto:

    ClassTributoDriver_AtualizarImposto = gErr

    Select Case gErr

        Case 106538
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153805)

    End Select

    Exit Function

End Function

Private Function ClassTributoDriver_Ativar() As Long
    'permite a inicializacao do driver
    'pode ser utilizado p/criar cache etc

    ClassTributoDriver_Ativar = SUCESSO
    
End Function

Private Function ClassTributoDriver_Desativar() As Long
    'informa que o driver nao vai mais ser necessario
    'pode ser utilizado p/liberar cache etc

    ClassTributoDriver_Desativar = SUCESSO
    
End Function

Private Function PV_Venda(ByVal sNatOp As String) As Boolean

    Select Case sNatOp
    
        Case "5101" To "5125", "5301" To "5357", "5401" To "5405", "6101" To "6125", "6401" To "6404"
            PV_Venda = True
            
        Case Else
            PV_Venda = False
            
    End Select
    
End Function

Private Function AtualizarDoc(objDoc As ClassTributoDoc, iRecalculaAutomatica As Integer) As Long

Dim lErro As Long, colItens As Collection, dValorCredito As Double, dValorCreditoDoc As Double, objVarValorCreditoItem As New ClassVariavelCalculo
Dim objDocItem As ClassTributoDocItem, dtDataEmissao As Date, objVarValorCompl As New ClassVariavelCalculo
Dim objVarDataEmissao As New ClassVariavelCalculo, objVarItemValor As New ClassVariavelCalculo
Dim objVarTipoTrib As New ClassVariavelCalculo, objTipoTrib As New ClassTipoDeTributacaoMovto
Dim objDocCompl As ClassTributoDocCompl, colDocCompl As Collection, objVarItemTipo As New ClassVariavelCalculo
Dim objTRBConfig As New ClassTRBConfig, objTipoDocInfo As New ClassTipoDocInfo
Dim objVarTipoDocInfo As New ClassVariavelCalculo, dtDataRef As Date
Dim objVarDataEntrada As New ClassVariavelCalculo, dtDataEntrada As Date
Dim dAliquota As Double, dValorCreditoDocManual As Double, dFatorRateio As Double, dSaldoCreditoDocManual As Double
Dim objVarValorCreditoDoc As New ClassVariavelCalculo, objVarValorCreditoCompl As New ClassVariavelCalculo
Dim objVarItemTipoPIS As New ClassVariavelCalculo, objVarItemTipoObj As New ClassVariavelCalculo
Dim objVarItemVlr As New ClassVariavelCalculo, objVarItemAliquota As New ClassVariavelCalculo
Dim objVarItemBase As New ClassVariavelCalculo, objVarItemTipoCalc As New ClassVariavelCalculo
Dim objTipoPIS As ClassTipoTribPISCOFINS, objVarItemUNVlr As New ClassVariavelCalculo
Dim objVarItemUNQtd As New ClassVariavelCalculo, objVarItemQtdeTrib As New ClassVariavelCalculo
Dim iTipoCalc As Integer, dBase As Double, dQtd As Double, dUNValor As Double
Dim objVarItemVlrST As New ClassVariavelCalculo, objVarItemAliquotaST As New ClassVariavelCalculo
Dim objVarItemBaseST As New ClassVariavelCalculo, objVarItemTipoCalcST As New ClassVariavelCalculo
Dim objVarItemUNVlrST As New ClassVariavelCalculo, dAliquotaST As Double
Dim objVarItemUNQtdST As New ClassVariavelCalculo, objVarItemQtdeTribST As New ClassVariavelCalculo
Dim iTipoCalcST As Integer, dBaseST As Double, dQtdST As Double, dUNValorST As Double, iTipoCOFINS As Integer
Dim objVarAliquotaICMSDI As New ClassVariavelCalculo, objVarAliquotaPISDI As New ClassVariavelCalculo, objVarAliquotaCOFINSDI As New ClassVariavelCalculo, objVarIPIValor As New ClassVariavelCalculo
Dim objVarIIValor As New ClassVariavelCalculo, iTipoCalculo As Integer
Dim objVarNatOp As New ClassVariavelCalculo, bAtualiza As Boolean
Dim bExportacao As Boolean, objVarUFDestino As New ClassVariavelCalculo
Dim mvariNotaInterna As Integer

'nfe 3.10
Dim objVarCOFINSValor As New ClassVariavelCalculo
Dim dCOFINSValor As Double
'fim nfe 3.10

On Error GoTo Erro_AtualizarDoc

    'obtem dados da filialempresa
    lErro = TestaFilialEmpresa(objDoc)
    If lErro <> SUCESSO Then gError 106539
    
    lErro = objDoc.ObterVar("COFINS_VALOR_CREDITO", objVarValorCreditoDoc)
    If lErro <> SUCESSO Then gError 106539
    If iRecalculaAutomatica = 1 Then objVarValorCreditoDoc.iStatus = VAR_PREENCH_AUTOMATICO
        
    lErro = objDoc.ObterVar("DATA_EMISSAO", objVarDataEmissao)
    If lErro <> SUCESSO Then gError 106539
    
    dtDataEmissao = objVarDataEmissao.vValor
    
    lErro = objDoc.ObterVar("DATA_ENTRADA", objVarDataEntrada)
    If lErro <> SUCESSO Then gError 106552
    
    dtDataEntrada = objVarDataEntrada.vValor
    
    lErro = objDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
    If lErro <> SUCESSO Then gError 106550
    
    objTipoDocInfo.iCodigo = objVarTipoDocInfo.vValor
    
    lErro = objTRBConfig.gobjTributacao.TipoDocInfo_Obter(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 106551
    
    mvariNotaInterna = IIf(objTipoDocInfo.iEmitente = EMITENTE_EMPRESA, 1, 0)
    
    bExportacao = False
    Select Case objTipoDocInfo.iTipoTrib
    
        Case TRIB_ENTRADA_CLI, TRIB_ENTRADA_FORN
            dtDataRef = dtDataEntrada
            
        Case TRIB_SAIDA_CLI, TRIB_SAIDA_FORN
            dtDataRef = dtDataEmissao
            
            If objTipoDocInfo.iCodigo <> DOCINFO_NFISSAT And objTipoDocInfo.iCodigo <> DOCINFO_TNFISFV Then
            
                If objTipoDocInfo.iTipoTrib = TRIB_SAIDA_CLI Then
                    'obter UF Destino
                    lErro = objDoc.ObterVar("CLIENTE_UF", objVarUFDestino)
                    If lErro <> SUCESSO Then gError 106552
                    If objVarUFDestino.vValor = "EX" Then bExportacao = True
                End If
    
            End If
    
    End Select
    
    Call objDoc.ObterColDocItens(colItens)
    
    objVarValorCreditoItem.sIdentificador = "COFINS_VALOR_CREDITO"
    objVarValorCreditoItem.iStatus = VAR_PREENCH_AUTOMATICO

    'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
    For Each objDocItem In colItens
    
        bAtualiza = True
        If gobjCRFAT.iOtimizaTrib = MARCADO Then
            If objDocItem.iJaCalculado = MARCADO Then bAtualiza = False
        End If
    
        lErro = objDocItem.ObterVar("NAT_OPERACAO", objVarNatOp)
        If lErro <> SUCESSO Then gError 106540
        
        lErro = objDocItem.ObterVar("TIPO_TRIB", objVarTipoTrib)
        If lErro <> SUCESSO Then gError 106540
        
        'obter mais info sobre o tipo trib
        objTipoTrib.iTipo = objVarTipoTrib.vValor
        lErro = objTRBConfig.gobjTributacao.TributacaoTipo_Obter(objTipoTrib)
        If lErro <> SUCESSO Then gError 106541
        
        dValorCredito = 0
        
        If bAtualiza Then
            
            dAliquotaST = 0
            dBaseST = 0
            dQtdST = 0
            dUNValorST = 0
            iTipoCalcST = 0
            
            dAliquota = 0
            dBase = 0
            dQtd = 0
            dUNValor = 0
            iTipoCalc = 0
    
            lErro = CF("Tributacao_Imposto_ObtemBaseItem", mvariFilialEmpresa, "COFINS", objDoc, objDocItem, dBase)
            If lErro <> SUCESSO Then gError 130765
            
            '30/06/2009
            'Alterado para respeitar caso tenha sido alterado na tela, aproveita e também já guarda no obj caso seja automático
            lErro = objDocItem.ObterVar("COFINS_TIPO", objVarItemTipoPIS)
            If lErro <> SUCESSO Then gError 130765
    
            lErro = objDocItem.ObterVar("COFINS_TIPO_CALCULO", objVarItemTipoCalc)
            If lErro <> SUCESSO Then gError 130765
            
            lErro = objDocItem.ObterVar("COFINS_ALIQUOTA", objVarItemAliquota)
            If lErro <> SUCESSO Then gError 130765
            
            lErro = objDocItem.ObterVar("COFINS_ALIQUOTA_VALOR", objVarItemUNVlr)
            If lErro <> SUCESSO Then gError 130765
            
            Select Case objTipoDocInfo.iTipoTrib
            
                Case TRIB_SAIDA_CLI, TRIB_ENTRADA_CLI
                    lErro = Excecao_Pesquisar2(objDoc, objDocItem, objTipoDocInfo, iTipoCOFINS, iTipoCalc, dAliquota, dUNValor)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 130765
                    
                Case TRIB_SAIDA_FORN, TRIB_ENTRADA_FORN
                    lErro = Excecao_Pesquisar3(objDoc, objDocItem, objTipoDocInfo, iTipoCOFINS, iTipoCalc, dAliquota, dUNValor)
                    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 130765
                    
            End Select
            
            'se nao tem excecao, usa aliquota de 0,76%
            If lErro <> SUCESSO Then
                
                dUNValor = 0
                dAliquota = 0.076
                If mvariNaoCumulativo = 0 Then dAliquota = 0.03
                iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL
                
                If objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC Then
                
                    'pegar da DI
                    lErro = objDocItem.ObterVar("COFINS_ALIQUOTA_ADICAO_DI", objVarAliquotaCOFINSDI)
                    If lErro <> SUCESSO Then gError 106551
                    
                    If Abs(objVarAliquotaCOFINSDI.vValor - dAliquota) > 0.001 Then
                        dAliquota = objVarAliquotaCOFINSDI.vValor
                    End If
                    
                    'pegar da DI
                    lErro = objDocItem.ObterVar("PIS_ALIQUOTA_ADICAO_DI", objVarAliquotaPISDI)
                    If lErro <> SUCESSO Then gError 106551
                
                    'pegar da DI
                    lErro = objDocItem.ObterVar("ICMS_ALIQUOTA_ADICAO_DI", objVarAliquotaICMSDI)
                    If lErro <> SUCESSO Then gError 106551
                    
                    lErro = objDocItem.ObterVar("IPI_VALOR", objVarIPIValor)
                    If lErro <> SUCESSO Then gError 106551
                    
                    'Preço Item
                    lErro = objDocItem.ObterVar("PRODUTO_VALOR", objVarItemValor)
                    If lErro <> SUCESSO Then gError 106551
                    
                    lErro = objDocItem.ObterVar("II_VALOR", objVarIIValor)
                    If lErro <> SUCESSO Then gError 106551
                    
                    If mvariSimplesNacional = 0 Or gobjCRFAT.iNFImportacaoTribFlag07 = MARCADO Then
                        If dtDataRef = DATA_NULA Or dtDataRef >= DATA_PIS_NOVO_CALC Then
                            If gobjCRFAT.iNFImportacaoTribFlag12 = 0 Then
                                dBase = Arredonda_Moeda(objVarItemValor.vValor - objVarIIValor.vValor, 2)
                            Else
                                dBase = Arredonda_Moeda(objVarItemValor.vValor, 2)
                            End If
                        Else
                            dBase = Arredonda_Moeda((objVarItemValor.vValor - objVarIIValor.vValor + (objVarAliquotaICMSDI.vValor * (objVarIIValor.vValor + objVarIPIValor.vValor))) / ((1 - objVarAliquotaPISDI.vValor - objVarAliquotaCOFINSDI.vValor) * (1 - objVarAliquotaICMSDI.vValor)), 2)
                        End If
                    Else
                        dBase = 0
                    End If
                End If
                
                If objVarItemTipoPIS.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                    objVarItemTipoPIS.vValor = objTipoTrib.iCOFINSTipo
                    If objTipoTrib.iCOFINSTipo = 0 Then
                        If mvariSimplesNacional <> 0 Then
                            objVarItemTipoPIS.vValor = 99 'outros
                        Else
                            If objTipoTrib.iPISCredita <> TIPOTRIB_SEMCREDDEB And (mvariNaoCumulativo Or (objTipoDocInfo.iFaturamento <> 0 Or (objTipoDocInfo.iCodigo = DOCINFO_PVN And PV_Venda(objVarNatOp.vValor)))) Then
                                 objVarItemTipoPIS.vValor = 1 'Operação Tributável (Base de Cálculo = Valor da Operação Alíquota Nornal (Cumulativo/Não Cumulativo))
                             Else
                                 objVarItemTipoPIS.vValor = 8 'Operação sem Incidência de Contribuição
                             End If
                        End If
                    End If
                    If mvariNaoCumulativo = 0 And mvariSimplesNacional = 0 And objVarItemTipoPIS.vValor = 1 And (objTipoDocInfo.iTipo = DOCINFO_NF_INT_ENTRADA Or objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA) Then
                        If objTipoDocInfo.iFaturamento = 2 Then
                            objVarItemTipoPIS.vValor = 98 'outros
                        Else
                            objVarItemTipoPIS.vValor = 70 'operacao de aquisicao sem credito
                        End If
                    End If
                    
                    If mvariSimplesNacional <> 0 And objTipoDocInfo.iEmitente = EMITENTE_EMPRESA Then
                    
                        objVarItemTipoPIS.vValor = 99 'outros
                        
                    End If
                    
                    'If bExportacao Then objVarItemTipoPIS.vValor = 8 'outros
                    
                    objVarItemTipoPIS.iStatus = VAR_PREENCH_AUTOMATICO
                    
                End If
                
            Else
                If objVarItemTipoPIS.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                    objVarItemTipoPIS.vValor = iTipoCOFINS
                End If
            End If
            
            lErro = objDocItem.GuardarVar(objVarItemTipoPIS)
            If lErro <> SUCESSO Then gError 130765
            
            Set objTipoPIS = New ClassTipoTribPISCOFINS
            
            objTipoPIS.iTipo = objVarItemTipoPIS.vValor
            Set objVarItemTipoObj.vValor = objTipoPIS
            lErro = objDoc.ObterVar("COFINS_TIPO_OBJ", objVarItemTipoObj)
            If lErro <> SUCESSO Then gError 130765
     
            Set objTipoPIS = objVarItemTipoObj.vValor
          
            lErro = objDocItem.ObterVar("PRODUTO_QTDE_UM_TRIB", objVarItemQtdeTrib)
            If lErro <> SUCESSO Then gError 130765
            
            dQtd = objVarItemQtdeTrib.vValor
            dQtdST = objVarItemQtdeTrib.vValor
                    
            lErro = objDocItem.ObterVar("COFINS_BASE", objVarItemBase)
            If lErro <> SUCESSO Then gError 130765
           
            lErro = objDocItem.ObterVar("COFINS_QTDE", objVarItemUNQtd)
            If lErro <> SUCESSO Then gError 130765
    
            lErro = objDocItem.ObterVar("COFINS_ST_TIPO_CALCULO", objVarItemTipoCalcST)
            If lErro <> SUCESSO Then gError 130765
            
            lErro = objDocItem.ObterVar("COFINS_ST_ALIQUOTA", objVarItemAliquotaST)
            If lErro <> SUCESSO Then gError 130765
            
            lErro = objDocItem.ObterVar("COFINS_ST_BASE", objVarItemBaseST)
            If lErro <> SUCESSO Then gError 130765
           
            lErro = objDocItem.ObterVar("COFINS_ST_ALIQUOTA_VALOR", objVarItemUNVlrST)
            If lErro <> SUCESSO Then gError 130765
           
            lErro = objDocItem.ObterVar("COFINS_ST_QTDE", objVarItemUNQtdST)
            If lErro <> SUCESSO Then gError 130765
           
            iTipoCalculo = objTipoPIS.iTipoCalculo
            If (gobjCRFAT.iPISSimplesCalc = 0 Or (gobjCRFAT.iPISSimplesCalc = 1 And mvariNotaInterna = MARCADO)) And mvariSimplesNacional <> 0 And (gobjCRFAT.iNFImportacaoTribFlag07 = DESMARCADO Or ((objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMP And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPSE And objTipoDocInfo.iCodigo <> DOCINFO_NFIEIMPPC))) Then iTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_NADA
            If mvariNaoCumulativo = 0 And (mvariSimplesNacional = 0 Or (gobjCRFAT.iNFImportacaoTribFlag07 = MARCADO And (objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC))) And objVarItemTipoPIS.vValor = 70 And ((objTipoDocInfo.iCodigo = DOCINFO_NFIEIMP Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPSE Or objTipoDocInfo.iCodigo = DOCINFO_NFIEIMPPC) Or objTipoDocInfo.iTipo = DOCINFO_NF_EXTERNA) Then iTipoCalculo = TRIB_TIPO_CALCULO_PERMITE_AMBOS
            
            Select Case iTipoCalculo
            
                Case TRIB_TIPO_CALCULO_PERMITE_NADA
                    objVarItemTipoCalc.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemVlr.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemUNVlr.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemUNQtd.iStatus = VAR_PREENCH_AUTOMATICO
                    dAliquota = 0
                    dBase = 0
                    dQtd = 0
                    dUNValor = 0
                    iTipoCalc = 0
                
                Case TRIB_TIPO_CALCULO_PERMITE_PERC
                    objVarItemTipoCalc.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemUNVlr.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemUNQtd.iStatus = VAR_PREENCH_AUTOMATICO
                    dQtd = 0
                    dUNValor = 0
                    iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL
                
                Case TRIB_TIPO_CALCULO_PERMITE_VALOR
                    objVarItemTipoCalc.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                    objVarItemAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                    dAliquota = 0
                    dBase = 0
                    iTipoCalc = TRIB_TIPO_CALCULO_VALOR
                    
                Case TRIB_TIPO_CALCULO_PERMITE_AMBOS
                
                    If objVarItemTipoCalc.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                        'iTipoCalc = TRIB_TIPO_CALCULO_PERMITE_PERC
                    Else
                        iTipoCalc = objVarItemTipoCalc.vValor
                    End If
                
                    If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                        objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                        objVarItemAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                        dAliquota = 0
                        dBase = 0
                    Else
                        objVarItemUNVlr.iStatus = VAR_PREENCH_AUTOMATICO
                        objVarItemUNQtd.iStatus = VAR_PREENCH_AUTOMATICO
                        dQtd = 0
                        dUNValor = 0
                    End If
                    
                Case TRIB_TIPO_CALCULO_PERMITE_AMBOS_MANUAL
                
                    If objVarItemTipoCalc.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                        'iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL
                    Else
                        iTipoCalc = objVarItemTipoCalc.vValor
                    End If
                
                    If iTipoCalc = TRIB_TIPO_CALCULO_VALOR Then
                        objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                        objVarItemAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                    Else
                        objVarItemUNVlr.iStatus = VAR_PREENCH_AUTOMATICO
                        objVarItemUNQtd.iStatus = VAR_PREENCH_AUTOMATICO
                        dQtd = 0
                        dUNValor = 0
                    End If
                    
                    'Zera se não tiver sido alterado manualmente
                    If objVarItemBase.iStatus <> VAR_PREENCH_MANUAL Then dBase = 0
                    If objVarItemAliquota.iStatus <> VAR_PREENCH_MANUAL Then dAliquota = 0
                    If objVarItemUNVlr.iStatus <> VAR_PREENCH_MANUAL Then dUNValor = 0
                    If objVarItemUNQtd.iStatus <> VAR_PREENCH_MANUAL Then dQtd = 0
                
            End Select
            
            If objVarItemTipoCalc.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemTipoCalc.vValor = iTipoCalc
                objVarItemTipoCalc.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemTipoCalc)
                If lErro <> SUCESSO Then gError 130765
            End If
            
            If objVarItemAliquota.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemAliquota.vValor = dAliquota
                objVarItemAliquota.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemAliquota)
                If lErro <> SUCESSO Then gError 130765
            Else
                dAliquota = objVarItemAliquota.vValor
            End If
            
            If objVarItemBase.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemBase.vValor = dBase
                objVarItemBase.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemBase)
                If lErro <> SUCESSO Then gError 130765
            Else
                dBase = objVarItemBase.vValor
            End If
            
            If objVarItemUNQtd.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemUNQtd.vValor = dQtd
                objVarItemUNQtd.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemUNQtd)
                If lErro <> SUCESSO Then gError 130765
            Else
                dQtd = objVarItemUNQtd.vValor
            End If
            
            If objVarItemUNVlr.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemUNVlr.vValor = dUNValor
                objVarItemUNVlr.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemUNVlr)
                If lErro <> SUCESSO Then gError 130765
            Else
                dUNValor = objVarItemUNVlr.vValor
            End If
            
            lErro = objDocItem.ObterVar("COFINS_VALOR", objVarItemVlr)
            If lErro <> SUCESSO Then gError 130765
            If objVarItemVlr.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                If iTipoCalc = TRIB_TIPO_CALCULO_PERCENTUAL Then
                    objVarItemVlr.vValor = CDbl(Format(dBase * dAliquota, "0.00"))
                Else
                    objVarItemVlr.vValor = CDbl(Format(dQtd * dUNValor, "0.00"))
                End If
                objVarItemVlr.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemVlr)
                If lErro <> SUCESSO Then gError 130765
            End If
            
            'nfe 3.10
            dCOFINSValor = dCOFINSValor + objVarItemVlr.vValor
            'fim nfe 3.10
            
            If objVarItemTipoCalcST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                iTipoCalcST = TRIB_TIPO_CALCULO_PERMITE_PERC
            Else
                iTipoCalcST = objVarItemTipoCalcST.vValor
            End If
        
            If iTipoCalcST = TRIB_TIPO_CALCULO_VALOR Then
                objVarItemBaseST.iStatus = VAR_PREENCH_AUTOMATICO
                objVarItemBaseST.iStatus = VAR_PREENCH_AUTOMATICO
                dAliquotaST = 0
                dBaseST = 0
            Else
                objVarItemUNVlrST.iStatus = VAR_PREENCH_AUTOMATICO
                objVarItemUNQtdST.iStatus = VAR_PREENCH_AUTOMATICO
                dQtdST = 0
                dUNValorST = 0
            End If
            
            If objVarItemTipoCalcST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemTipoCalcST.vValor = iTipoCalcST
                objVarItemTipoCalcST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemTipoCalcST)
                If lErro <> SUCESSO Then gError 130765
            End If
            
            If objVarItemAliquotaST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemAliquotaST.vValor = dAliquotaST
                objVarItemAliquotaST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemAliquotaST)
                If lErro <> SUCESSO Then gError 130765
            Else
                dAliquotaST = objVarItemAliquotaST.vValor
            End If
            
            If objVarItemBaseST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemBaseST.vValor = dBaseST
                objVarItemBaseST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemBaseST)
                If lErro <> SUCESSO Then gError 130765
            Else
                dBaseST = objVarItemBaseST.vValor
            End If
            
            If objVarItemUNQtdST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemUNQtdST.vValor = dQtdST
                objVarItemUNQtdST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemUNQtdST)
                If lErro <> SUCESSO Then gError 130765
            Else
                dQtdST = objVarItemUNQtdST.vValor
            End If
            
            If objVarItemUNVlrST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                objVarItemUNVlrST.vValor = dUNValorST
                objVarItemUNVlrST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemUNVlrST)
                If lErro <> SUCESSO Then gError 130765
            Else
                dUNValorST = objVarItemUNVlrST.vValor
            End If
            
            lErro = objDocItem.ObterVar("COFINS_ST_VALOR", objVarItemVlrST)
            If lErro <> SUCESSO Then gError 130765
            If objVarItemVlrST.iStatus <> VAR_PREENCH_MANUAL Or iRecalculaAutomatica = MARCADO Then
                If iTipoCalcST = TRIB_TIPO_CALCULO_PERCENTUAL Then
                    objVarItemVlrST.vValor = CDbl(Format(dBaseST * dAliquotaST, "0.00"))
                Else
                    objVarItemVlrST.vValor = CDbl(Format(dQtdST * dUNValorST, "0.00"))
                End If
                objVarItemVlrST.iStatus = VAR_PREENCH_AUTOMATICO
                lErro = objDocItem.GuardarVar(objVarItemVlrST)
                If lErro <> SUCESSO Then gError 130765
            End If
            
        Else
            
            'Le somente o que for necessário para calcular o crédito total no doc
            'objTipoTrib já está sendo obtido antes do IF
            'o dValorCredito também é zerado antes do IF
            lErro = objDocItem.ObterVar("COFINS_VALOR", objVarItemVlr)
            If lErro <> SUCESSO Then gError 130765
        
            'nfe 3.10
            dCOFINSValor = dCOFINSValor + objVarItemVlr.vValor
            'fim nfe 3.10
        
        End If
     
        If dtDataRef >= #2/1/2004# And mvariNaoCumulativo Then
    
            If objTipoTrib.iCOFINSCredita <> TIPOTRIB_SEMCREDDEB And mvariSimplesNacional = 0 Then
                
                'usa aliquota de 7,6%
                dValorCredito = objVarItemVlr.vValor  'CDbl(Format(objVarItemValor.vValor * dAliquota, "0.00"))
                
            End If
    
        End If
            
        If bAtualiza Then
            'guarda o valor do credito do COFINS do item
            objVarValorCreditoItem.vValor = dValorCredito
            lErro = objDocItem.GuardarVar(objVarValorCreditoItem)
            If lErro <> SUCESSO Then gError 106543
        End If
        
        dValorCreditoDoc = dValorCreditoDoc + dValorCredito
            
    Next
    
    If objVarValorCreditoDoc.iStatus <> VAR_PREENCH_MANUAL Then
    
        objVarValorCreditoDoc.iStatus = VAR_PREENCH_AUTOMATICO
        
        'guarda o valor do credito do COFINS do doc
        objVarValorCreditoDoc.vValor = dValorCreditoDoc
        lErro = objDoc.GuardarVar(objVarValorCreditoDoc)
        If lErro <> SUCESSO Then gError 106549
        
    Else
    
        dValorCreditoDocManual = objVarValorCreditoDoc.vValor
        If Abs(dValorCreditoDocManual - dValorCreditoDoc) > DELTA_VALORMONETARIO Then
        
            'ratear o valor manual do doc pelos itens e complementos
            If dValorCreditoDocManual = 0 Or dValorCreditoDoc = 0 Then
            
                'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
                For Each objDocItem In colItens
                
                    'guarda o valor do credito do COFINS do item
                    objVarValorCreditoItem.vValor = 0
                    lErro = objDocItem.GuardarVar(objVarValorCreditoItem)
                    If lErro <> SUCESSO Then gError 106543
                
                Next
            
            Else
                dFatorRateio = dValorCreditoDocManual / dValorCreditoDoc
                dSaldoCreditoDocManual = dValorCreditoDocManual
                
                'percorrer os itens vendo a contribuicao de cada um p/o Doc c/um todo
                For Each objDocItem In colItens
                
                    lErro = objDocItem.ObterVar("COFINS_VALOR_CREDITO", objVarValorCreditoItem)
                    If lErro <> SUCESSO Then gError 106542
                     
                    dValorCredito = Arredonda_Moeda(objVarValorCreditoItem.vValor * dFatorRateio)
                    dSaldoCreditoDocManual = dSaldoCreditoDocManual - dValorCredito
                    
                    'guarda o valor do credito do COFINS do item
                    objVarValorCreditoItem.vValor = dValorCredito
                    
                    lErro = objDocItem.GuardarVar(objVarValorCreditoItem)
                    If lErro <> SUCESSO Then gError 106543
                
                Next
                
                If Abs(dSaldoCreditoDocManual) > DELTA_VALORMONETARIO Then
                
                    'jogar esta diferenca para o 1o item com valor <> 0
                    For Each objDocItem In colItens
                    
                        lErro = objDocItem.ObterVar("COFINS_VALOR_CREDITO", objVarValorCreditoItem)
                        If lErro <> SUCESSO Then gError 106542
                         
                        If objVarValorCreditoItem.vValor > 0 Then
                        
                            dValorCredito = objVarValorCreditoItem.vValor + dSaldoCreditoDocManual
                            If dValorCredito < 0 Then dValorCredito = 0
                            
                            'guarda o valor do credito do COFINS do item
                            objVarValorCreditoItem.vValor = dValorCredito
                            
                            lErro = objDocItem.GuardarVar(objVarValorCreditoItem)
                            If lErro <> SUCESSO Then gError 106543
                    
                            Exit For
                            
                        End If
                        
                    Next
                    
                End If
                
            End If
        
        End If
        
    End If
    
    'nfe 3.10
    objVarCOFINSValor.sIdentificador = "COFINS_VALOR"
    objVarCOFINSValor.vValor = dCOFINSValor
    objVarCOFINSValor.iStatus = VAR_PREENCH_AUTOMATICO
    lErro = objDoc.GuardarVar(objVarCOFINSValor)
    If lErro <> SUCESSO Then Error 27746
    
    AtualizarDoc = SUCESSO

    Exit Function

Erro_AtualizarDoc:

    AtualizarDoc = gErr

    Select Case gErr

        Case 106539 To 106552, 130771
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153806)

    End Select

    Exit Function

End Function

Private Function Excecao_Pesquisar(ByVal objDocItem As ClassTributoDocItem, dAliquota As Double) As Long

Dim lErro As Long
Dim objVarProdColCateg As New ClassVariavelCalculo

On Error GoTo Erro_Excecao_Pesquisar

    'obter a colecao de categorias do produto
    lErro = objDocItem.ObterVar("I_PRODUTO_COL_CATEG", objVarProdColCateg)
    If lErro <> SUCESSO Then gError 130768
    
    lErro = CF("COFINS_Excecao_Pesquisar", objVarProdColCateg.vValor, dAliquota)
    If lErro <> SUCESSO And lErro <> 130767 Then gError 130769
    If lErro <> SUCESSO Then gError 130770
    
    Excecao_Pesquisar = SUCESSO
     
    Exit Function
    
Erro_Excecao_Pesquisar:

    Excecao_Pesquisar = gErr
     
    Select Case gErr
          
        Case 130768, 130769
        
        'nao encontrou excecao
        Case 130770
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130760)
     
    End Select
     
    Exit Function

End Function

Private Function Excecao_Pesquisar2(ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, ByVal objTipoDocInfo As ClassTipoDocInfo, iTipoCOFINS As Integer, iTipoCalc As Integer, dAliquota As Double, dUNValor As Double) As Long

Dim lErro As Long
Dim objVarCli As New ClassVariavelCalculo
Dim objVarFilCli As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo
Dim objExc As New ClassPISCOFINSExcecao
Dim objTRBConfig As New ClassTRBConfig

On Error GoTo Erro_Excecao_Pesquisar2

    lErro = objDoc.ObterVar("CLIENTE_CODIGO", objVarCli)
    If lErro <> SUCESSO Then gError 205501

    lErro = objDoc.ObterVar("FILIAL_CLIENTE_CODIGO", objVarFilCli)
    If lErro <> SUCESSO Then gError 205502

    lErro = objDocItem.ObterVar("PRODUTO_CODIGO", objVarProd)
    If lErro <> SUCESSO Then gError 205503

    lErro = CF("PisCofinsExcecao_Le_Cat", objVarCli.vValor, objVarFilCli.vValor, objVarProd.vValor, EXCECAO_PIS_COFINS_TIPO_COFINS, objExc)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205504

    If lErro <> SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
    
    iTipoCOFINS = objExc.iTipoCOFINS
    iTipoCalc = objExc.iCOFINSTipoCalculo
    dAliquota = objExc.dAliquotaCofinsPerc
    dUNValor = objExc.dAliquotaCofinsRS
    
    'Se é entrada
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Then
        'E tem um tipo específico para entrada
        If objExc.iTipoCOFINSE <> 0 Then
            iTipoCOFINS = objExc.iTipoCOFINSE
        End If
    End If
    
    Excecao_Pesquisar2 = SUCESSO
     
    Exit Function
    
Erro_Excecao_Pesquisar2:

    Excecao_Pesquisar2 = gErr
     
    Select Case gErr
          
        Case 205499 To 205504
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205505)
     
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
    
'        mvariSimplesFederal = objFilialEmpresa.iSimplesFederal
        mvariSimplesNacional = objFilialEmpresa.iSuperSimples
        mvariFilialEmpresa = objFilialEmpresa.iCodFilial
        
        mvariNaoCumulativo = objFilialEmpresa.iCOFINSNaoCumulativo
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

Private Function Excecao_Pesquisar3(ByVal objDoc As ClassTributoDoc, ByVal objDocItem As ClassTributoDocItem, ByVal objTipoDocInfo As ClassTipoDocInfo, iTipoCOFINS As Integer, iTipoCalc As Integer, dAliquota As Double, dUNValor As Double) As Long

Dim lErro As Long
Dim objVarForn As New ClassVariavelCalculo
Dim objVarFilForn As New ClassVariavelCalculo
Dim objVarProd As New ClassVariavelCalculo
Dim objExc As New ClassPISCOFINSExcecao
Dim objTRBConfig As New ClassTRBConfig

On Error GoTo Erro_Excecao_Pesquisar3

    lErro = objDoc.ObterVar("FORN_CODIGO", objVarForn)
    If lErro <> SUCESSO Then gError 205501

    lErro = objDoc.ObterVar("FILIAL_FORN_CODIGO", objVarFilForn)
    If lErro <> SUCESSO Then gError 205502

    lErro = objDocItem.ObterVar("PRODUTO_CODIGO", objVarProd)
    If lErro <> SUCESSO Then gError 205503

    lErro = CF("PisCofinsExcecao_Le_CatForn", objVarForn.vValor, objVarFilForn.vValor, objVarProd.vValor, EXCECAO_PIS_COFINS_TIPO_COFINS, objExc)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 205504

    If lErro <> SUCESSO Then gError ERRO_LEITURA_SEM_DADOS
    
    iTipoCOFINS = objExc.iTipoCOFINS
    iTipoCalc = objExc.iCOFINSTipoCalculo
    dAliquota = objExc.dAliquotaCofinsPerc
    dUNValor = objExc.dAliquotaCofinsRS
    
    'Se é entrada
    If objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFEXT Or objTipoDocInfo.iTipo = TIPODOCINFO_TIPO_NFIE Then
        'E tem um tipo específico para entrada
        If objExc.iTipoCOFINSE <> 0 Then
            iTipoCOFINS = objExc.iTipoCOFINSE
        End If
    End If
    
    Excecao_Pesquisar3 = SUCESSO
     
    Exit Function
    
Erro_Excecao_Pesquisar3:

    Excecao_Pesquisar3 = gErr
     
    Select Case gErr
          
        Case 205499 To 205504
        
        Case ERRO_LEITURA_SEM_DADOS
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205505)
     
    End Select
     
    Exit Function

End Function

