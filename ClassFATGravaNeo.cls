VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const DELTA_FILIALREAL_OFICIAL = 50

Function PedidoVenda_Valida_MargContr(ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
'Verifica se a margem de contribuicao do PV está dentro do aceitavel

Dim lErro As Long
Dim alComando(1 To 2) As Long, dQtdeUMVenda As Double
Dim iIndice As Integer, objProduto As New ClassProduto
Dim objItem As ClassItemPedido, dFator As Double
Dim dCoef As Double, dMargemMinima As Double
Dim dPreco As Double, dPrecoMinimo As Double, dPrecoItem_UMVenda As Double

On Error GoTo Erro_PedidoVenda_Valida_MargContr

    If objPedidoVenda.iFilialEmpresa > DELTA_FILIALREAL_OFICIAL Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 200257
        Next
    
        dMargemMinima = 0.15
        
        'Le o preço da tabela de preço para obter seu coeficiente
        lErro = Comando_Executar(alComando(1), "SELECT VlrCompCoef FROM TabelasDePreco WHERE AjustaTabVlrCompCoefMaior = ? AND VlrCompCoef <> 0 AND Codigo = ?", dCoef, MARCADO, objPedidoVenda.iTabelaPreco)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200258
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200259
        
        If lErro = AD_SQL_SUCESSO Then
        
            For Each objItem In objPedidoVenda.colItensPedido
        
                'Pesquisa no BD Ítens de Tabela de Preço com vigência atual
                lErro = Comando_Executar(alComando(2), "SELECT Preco FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", dPreco, objPedidoVenda.iFilialEmpresa, objPedidoVenda.iTabelaPreco, objItem.sProduto, objPedidoVenda.dtDataEmissao)
                If lErro <> AD_SQL_SUCESSO Then gError 200260
            
                'Verifica se Ítem existe
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200261
                
                If lErro = AD_SQL_SUCESSO Then
                            
                    objProduto.sCodigo = objItem.sProduto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 201037
                    If lErro <> SUCESSO Then gError 201038
                    
                    lErro = CF("UM_Conversao", objProduto.iClasseUM, objItem.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
                    If lErro <> SUCESSO Then gError 201031
                    
                    dQtdeUMVenda = objItem.dQuantidade * dFator
                    If dQtdeUMVenda = 0 Then dQtdeUMVenda = 1
                    
                    dPrecoItem_UMVenda = Arredonda_Moeda(objItem.dPrecoTotal / dQtdeUMVenda)
                    
                    dPrecoMinimo = Arredonda_Moeda((dPreco / dCoef) * (1 + dMargemMinima))
                    
                    'Se está sendo vendido com menos de 15% de margem  => Erro
                    If (dPrecoMinimo - dPrecoItem_UMVenda) > DELTA_VALORMONETARIO Then gError 200262
    
                End If
        
            Next
        
        End If
     
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next

    End If
    
    PedidoVenda_Valida_MargContr = SUCESSO

    Exit Function

Erro_PedidoVenda_Valida_MargContr:

    PedidoVenda_Valida_MargContr = gErr

    Select Case gErr

        Case 201037
        
        Case 200257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200258, 200259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO", gErr)
            
        Case 200260, 200261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", gErr)

        Case 200262
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FORA_VALOR_MINIMO", gErr, objItem.sProduto, objPedidoVenda.iFilialEmpresa, objPedidoVenda.iTabelaPreco, objPedidoVenda.dtDataEmissao, Format(dPrecoMinimo, "STANDARD"))

        Case 201038
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ENCONTRADO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200263)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFiscal_Valida_MargContr(ByVal objNF As ClassNFiscal) As Long
'Verifica se a margem de contribuicao da nfiscal está dentro do aceitavel

Dim lErro As Long
Dim alComando(1 To 2) As Long, dQtdeUMVenda As Double
Dim iIndice As Integer, objProduto As New ClassProduto
Dim objItem As ClassItemNF, dFator As Double, objTipoDocInfo As New ClassTipoDocInfo
Dim dCoef As Double, dMargemMinima As Double
Dim dPreco As Double, dPrecoMinimo As Double, dPrecoItem_UMVenda As Double

On Error GoTo Erro_NFiscal_Valida_MargContr

    If objNF.iFilialEmpresa > DELTA_FILIALREAL_OFICIAL Then
    
        objTipoDocInfo.iCodigo = objNF.iTipoNFiscal
        
        'Lê o Tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 35913
        
        If lErro = 31415 Then gError 35914 'Não encontrou
    
        If objTipoDocInfo.iComplementar = DOCINFO_NORMAL And objTipoDocInfo.iFaturamento = 1 Then
            
            For iIndice = LBound(alComando) To UBound(alComando)
                alComando(iIndice) = Comando_Abrir()
                If alComando(iIndice) = 0 Then gError 200257
            Next
        
            dMargemMinima = 0.15
            
            'Le o preço da tabela de preço para obter seu coeficiente
            lErro = Comando_Executar(alComando(1), "SELECT VlrCompCoef FROM TabelasDePreco WHERE AjustaTabVlrCompCoefMaior = ? AND VlrCompCoef <> 0 AND Codigo = ?", dCoef, MARCADO, objNF.iTabelaPreco)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200258
        
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200259
            
            If lErro = AD_SQL_SUCESSO Then
            
                For Each objItem In objNF.colItensNF
            
                    'Pesquisa no BD Ítens de Tabela de Preço com vigência atual
                    lErro = Comando_Executar(alComando(2), "SELECT Preco FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", dPreco, objNF.iFilialEmpresa, objNF.iTabelaPreco, objItem.sProduto, objNF.dtDataEmissao)
                    If lErro <> AD_SQL_SUCESSO Then gError 200260
                
                    'Verifica se Ítem existe
                    lErro = Comando_BuscarPrimeiro(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200261
                    
                    If lErro = AD_SQL_SUCESSO Then
                                
                        objProduto.sCodigo = objItem.sProduto
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 201037
                        If lErro <> SUCESSO Then gError 201038
                        
                        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItem.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
                        If lErro <> SUCESSO Then gError 201031
                        
                        dQtdeUMVenda = objItem.dQuantidade * dFator
                        If dQtdeUMVenda = 0 Then dQtdeUMVenda = 1
                        
                        dPrecoItem_UMVenda = Arredonda_Moeda(objItem.dValorTotal / dQtdeUMVenda)
                        
                        dPrecoMinimo = Arredonda_Moeda((dPreco / dCoef) * (1 + dMargemMinima))
                        
                        'Se está sendo vendido com menos de 15% de margem  => Erro
                        If (dPrecoMinimo - dPrecoItem_UMVenda) > DELTA_VALORMONETARIO Then gError 200262
        
                    End If
            
                Next
            
            End If
         
            For iIndice = LBound(alComando) To UBound(alComando)
                Call Comando_Fechar(alComando(iIndice))
            Next
    
        End If
    
    End If
    
    NFiscal_Valida_MargContr = SUCESSO

    Exit Function

Erro_NFiscal_Valida_MargContr:

    NFiscal_Valida_MargContr = gErr

    Select Case gErr

        Case 201037, 39513
        
        Case 200257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200258, 200259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO", gErr)
            
        Case 200260, 200261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", gErr)

        Case 200262
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FORA_VALOR_MINIMO", gErr, objItem.sProduto, objNF.iFilialEmpresa, objNF.iTabelaPreco, objNF.dtDataEmissao, Format(dPrecoMinimo, "STANDARD"))

        Case 201038
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ENCONTRADO", gErr, objProduto.sCodigo)
        
        Case 35914
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200263)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

