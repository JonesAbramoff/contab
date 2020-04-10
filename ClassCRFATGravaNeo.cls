VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function NFiscal_Valida_MargContr(ByVal objNF As ClassNFiscal) As Long
'Verifica se a margem de contribuicao da nfiscal está dentro do aceitavel

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, objProduto As New ClassProduto
Dim objItemNF As ClassItemNF, dFator As Double
Dim dCoef As Double, dMargemMinima As Double
Dim dPreco As Double, dPrecoMinimo As Double

On Error GoTo Erro_NFiscal_Valida_MargContr

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200257
    Next

    dMargemMinima = 1.15
    
    'Le o preço da tabela de preço 1, que servirá como base
    lErro = Comando_Executar(alComando(1), "SELECT VlrCompCoef FROM TabelasDePreco WHERE AjustaTabVlrCompCoefMaior = ? AND VlrCompCoef <> 0 AND Codigo = 1", dCoef, MARCADO)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200258

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200259
    
    If lErro = AD_SQL_SUCESSO Then
    
        For Each objItemNF In objNF.ColItensNF
    
             'Pesquisa no BD Ítens de Tabela de Preço com vigência atual
            lErro = Comando_Executar(alComando(2), "SELECT Preco FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia <= ? ORDER BY DataVigencia DESC", dPreco, objNF.iFilialEmpresa, objNF.iTabelaPreco, objItemNF.sProduto, objNF.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 200260
        
            'Verifica se Ítem existe
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200261
            
            If lErro = AD_SQL_SUCESSO Then
                        
                'Se está sendo vendido com menos de 15% de margem  => Erro
                If (Arredonda_Moeda((dPreco / dCoef) * dMargemMinima) - objItemNF.dPrecoUnitario) > DELTA_VALORMONETARIO Then gError 200262

            End If
    
        Next
    
    End If
 
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Valida_MargContr = SUCESSO

    Exit Function

Erro_NFiscal_Valida_MargContr:

    NFiscal_Valida_MargContr = gErr

    Select Case gErr

        Case 200257
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200258, 200259
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO", gErr)
            
        Case 200260, 200261
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", gErr)

        Case 200262
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FORA_VALOR_MINIMO", gErr, objItemNF.sProduto, objNF.iFilialEmpresa, objNF.iTabelaPreco, objNF.dtDataEmissao, Format((dPreco / dCoef) * 1.15, "STANDARD"))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200263)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

