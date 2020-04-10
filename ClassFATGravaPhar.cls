VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGravaPhar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Incluído gravação dos campos comprimento, pesokg e vendapeso
Function NFiscalItens_Grava_Customizado(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim objItemNFPhar As ClassItemNF
Dim objItemNF As ClassItemNF
Dim lNumIntDoc As Long
Dim sProdutoKit As String
Dim iSeqPed As Integer
Dim objItemNFPhar1 As New ClassItemNF, iSeqNoPack As Integer

On Error GoTo Erro_NFiscalItens_Grava_Customizado

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197688
    Next

    iIndice = 0
    iSeqNoPack = 1
    
    For Each objItemNF In objNFiscal.ColItensNF
    
        iIndice = iIndice + 1
    
        If Not (objItemNF.objInfoUsu Is Nothing) Then
        
            'Pesquisa os ítens vinculados à Nota Fiscal
            lErro = Comando_ExecutarPos(alComando(0), "SELECT NumIntDoc FROM ItensNFiscal WHERE NumIntDoc = ? ", 0, lNumIntDoc, objItemNF.lNumIntDoc)
            If lErro <> AD_SQL_SUCESSO Then gError 197689
    
            lErro = Comando_BuscarPrimeiro(alComando(0))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197690
    
            If lErro = AD_SQL_SUCESSO Then
            
                Set objItemNFPhar = objItemNF.objInfoUsu
                
                If iIndice > 1 Then
                    Set objItemNFPhar1 = objNFiscal.ColItensNF(iIndice - 1).objInfoUsu
                End If
                
                If Len(Trim(objItemNFPhar.sProduto)) = 0 Or objItemNFPhar.sProduto <> objItemNFPhar1.sProduto Then
                    iSeqPed = iSeqPed + 1
                    iSeqNoPack = 1
                Else
                    iSeqNoPack = iSeqNoPack + 1
                End If
                
                lErro = Comando_ExecutarPos(alComando(1), "UPDATE ItensNFiscal SET ProdutoKit = ?, QuantKit = ?, UMKit = ?, SeqPed = ?, SeqNoPack = ?", alComando(0), objItemNFPhar.sProduto, objItemNFPhar.dQuantidade, objItemNFPhar.sUnidadeMed, iSeqPed, iSeqNoPack)
                If lErro <> AD_SQL_SUCESSO Then gError 197691
    
            End If

        End If

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscalItens_Grava_Customizado = SUCESSO

    Exit Function

Erro_NFiscalItens_Grava_Customizado:

    NFiscalItens_Grava_Customizado = gErr

    Select Case gErr

        Case 197688
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197689, 197690
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ITENSNFISCAL", gErr)

        Case 197691
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_ITENSNFISCAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197692)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MovEstoque_ItemNF_Customizado(objItemNF As ClassItemNF, objNFiscal As ClassNFiscal) As Long

Dim objItemNF1 As ClassItemNF
Dim objItemNF2 As ClassItemNF
Dim objItemNF3 As ClassItemNF
Dim objProduto As New ClassProduto
Dim lErro As Long
Dim dValorTotal As Double
Dim iCont As Integer
Dim iAchou As Integer

On Error GoTo Erro_MovEstoque_ItemNF_Customizado

    If Not (objItemNF.objInfoUsu Is Nothing) Then
        
        Set objItemNF1 = objItemNF.objInfoUsu
        
        If Len(Trim(objItemNF1.sProduto)) > 0 Then
        
            Set objItemNF1.ColAlocacoes = objItemNF.ColAlocacoes
            
            'Lê o produto
            objProduto.sCodigo = objItemNF1.sProduto
            
            
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 197706
                    
            objItemNF1.sDescricaoItem = objProduto.sDescricao
            objItemNF1.iControleEstoque = objProduto.iControleEstoque
            objItemNF1.iClasseUM = objProduto.iClasseUM
            objItemNF1.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF1.sUMVenda = objProduto.sSiglaUMVenda
            objItemNF1.lNumIntDoc = objItemNF.lNumIntDoc
            objItemNF1.iApropriacaoProd = objProduto.iApropriacaoCusto
            objItemNF1.lNumIntItemPedVenda = objItemNF.lNumIntItemPedVenda
            
            For Each objItemNF2 In objNFiscal.ColItensNF
                
                If Not (objItemNF2.objInfoUsu Is Nothing) Then
                
                    Set objItemNF3 = objItemNF2.objInfoUsu
                    
                    If objItemNF3.sProduto = objItemNF1.sProduto Then
                        
                        iCont = iCont + 1
                    
                        'se nao for o primeiro item de nota fiscal com o produto do pedido em questao ==> a alocacao nao ficara associada a ele
                        If objItemNF2 Is objItemNF And iCont = 1 Then iAchou = 1
                        
                        dValorTotal = dValorTotal + (objItemNF2.dValorTotal)
                    End If
                    
                End If
            Next
            
            objItemNF1.dPrecoUnitario = dValorTotal / objItemNF1.dQuantidade
            
            If iAchou = 1 Then
            
                Set objItemNF = objItemNF1
                
            Else
            
                Set objItemNF = Nothing
                
            End If
        
        End If
        
    End If
    
    MovEstoque_ItemNF_Customizado = SUCESSO
    
    Exit Function
    
Erro_MovEstoque_ItemNF_Customizado:

    MovEstoque_ItemNF_Customizado = gErr
    
    Select Case gErr

        Case 197706

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197707)

    End Select
    
End Function

Public Function NFe_Valida_Envio_Cust(ByVal objNFeFedLote As ClassNFeFedLote) As Long

Dim lErro As Long

On Error GoTo Erro_NFe_Valida_Envio_Cust

    lErro = CF("NF_Valida_PharSep", objNFeFedLote.lNumIntNF)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    NFe_Valida_Envio_Cust = SUCESSO
    
    Exit Function
    
Erro_NFe_Valida_Envio_Cust:

    NFe_Valida_Envio_Cust = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211193)

    End Select
    
End Function
