VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function RelCustoImportacao_Prepara(ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal sDI As String, lNumIntRel As Long) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim iIndice As Integer, alComando(1 To 4) As Long

On Error GoTo Erro_RelCustoImportacao_Prepara

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 184871
 
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 184872
    Next
    
    'Obter NumIntRel
    lErro = CF("Config_ObterNumInt", "ESTConfig", "NUM_PROX_REL_CUSTOIMP", lNumIntRel)
    If lErro <> SUCESSO Then gError 184873
    
    If sDI = "" Then
    
        lErro = RelCustoImportacao_Prepara2(dtDataIni, dtDataFim, sProd_I, sProd_F, lFornIni, lFornFim, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184875
    
    Else
    
        lErro = RelCustoImportacao_Prepara1(sProd_I, sProd_F, lFornIni, lFornFim, sDI, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184875
    
    End If
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 184874
 
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelCustoImportacao_Prepara = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara:

    RelCustoImportacao_Prepara = gErr

    Select Case gErr

        Case 184871
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 184872
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 184873, 184875
        
        Case 184874
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184870)

    End Select
    
    Call Transacao_Rollback
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function RelCustoImportacao_Prepara1(ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal sDI As String, ByVal lNumIntRel As Long, alComando() As Long) As Long

Dim lErro As Long, objProdutoFilial As ClassProdutoFilial, iFilialForn As Integer
Dim objDIInfo As ClassDIInfo, objDespesa As ClassImportCompl, objAdicao As ClassAdicaoDI
Dim objItem As ClassItemAdicaoDI, lForn As Long, dValorItemaAux As Double, iDias As Integer
Dim dValorAduaneiroDI As Double, dValorDespesasLegais As Double, dValorTotalCreditos As Double
Dim dValorAcertosFiscais As Double, dFatorItemAdicao As Double, dFatorItemDI As Double
Dim dFatorCustoFinanceiro As Double, dFatorSeguroProduto As Double, dValorTotalBruto As Double
Dim dOverhead As Double, dComissFornec As Double, dComissFornecNormal As Double, dComissFornecMobimax As Double

On Error GoTo Erro_RelCustoImportacao_Prepara1

    Set objDIInfo = New ClassDIInfo
        
    objDIInfo.sNumero = sDI
    
    lErro = CF("DIInfo_Le", objDIInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184876
    If lErro <> SUCESSO Then gError 184877
    
    dValorAduaneiroDI = Arredonda_Moeda(objDIInfo.dValorMercadoriaEmReal + objDIInfo.dValorFreteInternacEmReal + objDIInfo.dValorSeguroInternacEmReal)
    
    For Each objDespesa In objDIInfo.colDespesasDI
    
        Select Case objDespesa.iTipo
        
            Case IMPORTCOMPL_TIPO_ACERTOS_FISCAIS
                dValorAcertosFiscais = objDespesa.dValor
                
            Case IMPORTCOMPL_TIPO_CUSTO_FINANCEIRO
                iDias = objDespesa.iDias
                If iDias = 0 Then iDias = 30
                dFatorCustoFinanceiro = objDespesa.dPerc * iDias / 30
            
            Case IMPORTCOMPL_TIPO_SEGURO_PRODUTO
                iDias = objDespesa.iDias
                If iDias = 0 Then iDias = 30
                dFatorSeguroProduto = objDespesa.dPerc * iDias / 30
                
            Case IMPORTCOMPL_TIPO_OVERHEAD
                dOverhead = objDespesa.dPerc
            
            Case IMPORTCOMPL_TIPO_COMISSAO_FORNEC
                dComissFornecNormal = objDespesa.dPerc
                
            Case IMPORTCOMPL_TIPO_COMISSAO_FORNEC_MOBIMAX
                dComissFornecMobimax = objDespesa.dPerc
        
            Case IMPORTCOMPL_TIPO_TAXA_DE_LI
                'rateio depende de NCM
                '???
                dValorDespesasLegais = dValorDespesasLegais + objDespesa.dValor
                
            Case Else
                dValorDespesasLegais = dValorDespesasLegais + objDespesa.dValor
        
        End Select
    
    Next
    
    For Each objAdicao In objDIInfo.colAdicoesDI
    
        For Each objItem In objAdicao.colItensAdicaoDI
    
            If objItem.sProduto >= sProd_I And (sProd_F = "" Or (objItem.sProduto <= sProd_F)) Then
            
                lForn = 0
                lErro = Comando_Executar(alComando(4), "SELECT Fornecedor FROM FornecedorProdutoFF WHERE FilialEmpresa = ? AND Produto = ?", lForn, objDIInfo.iFilialEmpresa, objItem.sProduto)
                If lErro <> AD_SQL_SUCESSO Then gError 184879
                
                lErro = Comando_BuscarProximo(alComando(4))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184880
                
                If lErro <> AD_SQL_SUCESSO Then lForn = objDIInfo.lFornTrading
                            
                If lForn = 0 Or (lForn >= lFornIni And (lFornFim = 0 Or (lForn <= lFornFim))) Then
            
                    If objAdicao.dValorAduaneiro = 0 Then
                        dFatorItemAdicao = 0
                    Else
                        dFatorItemAdicao = objItem.dValorTotalCIFEmReal / objAdicao.dValorAduaneiro
                    End If
                    
                    If dValorAduaneiroDI = 0 Then
                        dFatorItemDI = 0
                    Else
                        dFatorItemDI = objItem.dValorTotalCIFEmReal / dValorAduaneiroDI
                    End If
                    
                    If lForn <> 0 Then
                        'obter comissao do fornecedor
                        lErro = Comando_Executar(alComando(2), "SELECT Filial FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Categoria = ? AND Item = ?", iFilialForn, lForn, "comissao", "especial")
                        If lErro <> AD_SQL_SUCESSO Then gError 184879
                        
                        lErro = Comando_BuscarProximo(alComando(2))
                        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184880
                        
                        'se o fornecedor tem comissao especial
                        If lErro = AD_SQL_SUCESSO Then
                            dComissFornec = dComissFornecMobimax
                            If dComissFornec = 0 Then dComissFornec = 0.1
                        Else
                            dComissFornec = dComissFornecNormal
                            If dComissFornec = 0 Then dComissFornec = 0.01
                        End If
                        
                    Else
                    
                        dComissFornec = 0.01
                        
                    End If
                    
                    dValorTotalCreditos = Arredonda_Moeda(dFatorItemAdicao * (objAdicao.dIPIValor + objAdicao.dPISValor + objAdicao.dCOFINSValor + objAdicao.dICMSValor))
                    'CIF + impostos + despesas legais
                    dValorItemaAux = Arredonda_Moeda(objItem.dValorTotalCIFEmReal + (dFatorItemAdicao * objAdicao.dIIValor) + dValorTotalCreditos + (dFatorItemDI * dValorDespesasLegais))
                    dValorTotalBruto = Arredonda_Moeda(dValorItemaAux + (dFatorItemDI * dValorAcertosFiscais) + (dValorItemaAux * (dFatorCustoFinanceiro + dFatorSeguroProduto + dOverhead)) + (objItem.dValorTotalCIFEmReal * dComissFornec))
                    
                    lErro = Comando_Executar(alComando(1), "INSERT INTO RelCustoImportacaoAux (NumIntRel, DINumIntDoc, " & _
                        "AdicaoDINumIntDoc, ItensAdicaoDINumIntDoc, ValorIIEmReal, ValorIPIEmReal, ValorPISEmReal, ValorCOFINSEmReal, " & _
                        "ValorICMSEmReal, ValorDespesasEmReal, ValorAcertosFiscais, ValorCustoFinanceiro, " & _
                        "ValorSeguro, ValorOverhead , ValorComissao, ValorTotalBruto, ValorTotalCreditos, ValorTotalLiquido, Fornecedor) " & _
                        " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        lNumIntRel, objDIInfo.lNumIntDoc, objAdicao.lNumIntDoc, objItem.lNumIntDoc, Arredonda_Moeda(dFatorItemAdicao * objAdicao.dIIValor), _
                        Arredonda_Moeda(dFatorItemAdicao * objAdicao.dIPIValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dPISValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dCOFINSValor), Arredonda_Moeda(dFatorItemAdicao * objAdicao.dICMSValor), _
                        Arredonda_Moeda(dFatorItemDI * dValorDespesasLegais), Arredonda_Moeda(dFatorItemDI * dValorAcertosFiscais), Arredonda_Moeda(dValorItemaAux * dFatorCustoFinanceiro), Arredonda_Moeda(dValorItemaAux * dFatorSeguroProduto), _
                        Arredonda_Moeda(dValorItemaAux * dOverhead), Arredonda_Moeda(objItem.dValorTotalCIFEmReal * dComissFornec), Arredonda_Moeda(dValorTotalBruto), Arredonda_Moeda(dValorTotalCreditos), Arredonda_Moeda(dValorTotalBruto - dValorTotalCreditos), lForn)
                    If lErro <> AD_SQL_SUCESSO Then gError 184881
            
                End If
                
            End If
    
        Next
        
    Next
    
    RelCustoImportacao_Prepara1 = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara1:

    RelCustoImportacao_Prepara1 = gErr

    Select Case gErr

        Case 184876, 184878
        
        Case 184877
            Call Rotina_Erro(vbOKOnly, "ERRO_DIINFO_NAO_CADASTRADO", gErr, objDIInfo.sNumero)
        
        Case 184879, 184880
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILFORNCATEG", gErr)
        
        Case 184881
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCUSTOIMPAUX", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184882)

    End Select
    
    Exit Function

End Function

Private Function RelCustoImportacao_Prepara2(ByVal dtDataIni As Date, ByVal dtDataFim As Date, ByVal sProd_I As String, ByVal sProd_F As String, ByVal lFornIni As Long, ByVal lFornFim As Long, ByVal lNumIntRel As Long, alComando() As Long) As Long

Dim lErro As Long, sDI As String

On Error GoTo Erro_RelCustoImportacao_Prepara2

    If dtDataFim = DATA_NULA Then dtDataFim = gdtDataAtual
    
    sDI = String(STRING_DI_NUMERO, 0)
    lErro = Comando_Executar(alComando(3), "SELECT Numero FROM DIInfo WHERE Data >= ? AND Data <= ? AND FilialEmpresa = ? ORDER BY Data", sDI, dtDataIni, dtDataFim, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 184884
    
    lErro = Comando_BuscarProximo(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184885
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = RelCustoImportacao_Prepara1(sProd_I, sProd_F, lFornIni, lFornFim, sDI, lNumIntRel, alComando())
        If lErro <> SUCESSO Then gError 184887
    
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184888
    
    Loop
    
    RelCustoImportacao_Prepara2 = SUCESSO
    
    Exit Function
    
Erro_RelCustoImportacao_Prepara2:

    RelCustoImportacao_Prepara2 = gErr

    Select Case gErr

        Case 184887
        
        Case 184884, 184885, 184888
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DIINFO_RELCUSTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184883)

    End Select
    
    Exit Function

End Function



