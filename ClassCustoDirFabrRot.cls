VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoDirFabrRot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rotina_Rateio_CustoDirFabr(ByVal sNomeArqParam As String, ByVal objCustoDirFabr As ClassCustoDirFabr) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_Rateio_CustoDirFabr

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106632

    lErro = CustoDiretoFabricacao_Calcula(objCustoDirFabr)
    If lErro <> SUCESSO Then gError 106633
        
    Rotina_Rateio_CustoDirFabr = SUCESSO
     
    Exit Function
    
Erro_Rotina_Rateio_CustoDirFabr:

    Rotina_Rateio_CustoDirFabr = gErr
     
    Select Case gErr
          
        Case 106632, 106633
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149252)
     
    End Select
     
    Exit Function

End Function

Function CustoDiretoFabricacao_Calcula(ByVal objCustoDirFabr As ClassCustoDirFabr) As Long
'Calcula o custo para a producao de iFilialEmpresa no iAno.
'ATENCAO:
'Caso o sProduto esteja preenchido o calculo será apenas feito para o mesmo,
'senao será calculado para todos os produtos da sCodigoPrevVenda

Dim lErro As Long, iIndice As Integer, tCustDirFabr As typeCustoDirFabr
Dim lTransacao As Long, alComando(1 To 34) As Long

On Error GoTo Erro_CustoDiretoFabricacao_Calcula

    'Abertura de transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106590
 
    'Abertura dos comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If lErro <> SUCESSO Then gError 106591
    Next
    
    objCustoDirFabr.iRotinaOrigem = FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS
    
    'se estiver definido um produto
    If Len(Trim(objCustoDirFabr.sProduto)) <> 0 Then
    
        'pesquisar registro em CustoDirFabr por filialempresa e ano, preenchendo quantidades e custos com os dados do bd
        With tCustDirFabr
            lErro = Comando_ExecutarPos(alComando(1), "SELECT QuantFator1, QuantFator2, QuantFator3, QuantFator4, QuantFator5, QuantFator6, CustoTotal, CustoFator1, CustoFator2, CustoFator3, CustoFator4, CustoFator5, CustoFator6 FROM CustoDirFabr WHERE FilialEmpresa = ? AND Ano = ?", 0, _
                .dQuantFator1, .dQuantFator2, .dQuantFator3, .dQuantFator4, .dQuantFator5, .dQuantFator6, .dCustoTotal, .dCustoFator1, .dCustoFator2, .dCustoFator3, .dCustoFator4, .dCustoFator5, .dCustoFator6, objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 106592
        
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106593
        
        'se nao encontrar entao erro
        If lErro <> AD_SQL_SUCESSO Then gError 106594
        
        'fazer lock
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 106595
        
        'excluir registros em CustoDirFabrPlanMP por filialempresa, ano e produto
        lErro = Comando_Executar(alComando(2), "DELETE FROM CustoDirFabrPlanMP WHERE Filialempresa = ? AND Ano = ? AND Produto = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106618
        
        'excluir registros em CustoDirFabrPlanEmb por filialempresa, ano e produto
        lErro = Comando_Executar(alComando(33), "DELETE FROM CustoDirFabrPlanEmb WHERE Filialempresa = ? AND Ano = ? AND Produto = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106618
        
        'excluir registros em CustoDirFabrPlan por filialempresa, ano e produto
        lErro = Comando_Executar(alComando(34), "DELETE FROM CustoDirFabrPlan WHERE Filialempresa = ? AND Ano = ? AND Produto = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106596
        
        'excluir registro em CustoDirFabrProd por filialempresa, ano e produto
        lErro = Comando_Executar(alComando(3), "DELETE FROM CustoDirFabrProd WHERE Filialempresa = ? AND Ano = ? AND Produto = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106597
        
        'executar o recalculo do produto
        lErro = CustoDiretoFabricacao_Calcula1(objCustoDirFabr, alComando)
        If lErro <> SUCESSO Then gError 106598
    
    Else
        
        'procurar registro em CustoDirFabr por filialempresa e ano (executarpos)
        lErro = Comando_ExecutarPos(alComando(6), "SELECT Data FROM CustoDirFabr WHERE Filialempresa = ? AND Ano = ?", 0, tCustDirFabr.dtData, objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
        If lErro <> AD_SQL_SUCESSO Then gError 106599
        
        lErro = Comando_BuscarProximo(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106600
        
        'se houver registro em CustoDirFabr
        If lErro = AD_SQL_SUCESSO Then
        
            'fazer lock nele
            lErro = Comando_LockExclusive(alComando(6))
            If lErro <> AD_SQL_SUCESSO Then gError 106601
            
            'excluir registros em CustoDirFabrPlanMP por filialempresa e ano
            lErro = Comando_Executar(alComando(2), "DELETE FROM CustoDirFabrPlanMP WHERE Filialempresa = ? AND Ano = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
            If lErro <> AD_SQL_SUCESSO Then gError 106619
        
            'excluir registros em CustoDirFabrPlanEmb por filialempresa e ano
            lErro = Comando_Executar(alComando(33), "DELETE FROM CustoDirFabrPlanEmb WHERE Filialempresa = ? AND Ano = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
            If lErro <> AD_SQL_SUCESSO Then gError 106619
        
            'excluir registros de CustoDirFabrPlan por filialempresa e ano
            lErro = Comando_Executar(alComando(4), "DELETE FROM CustoDirFabrPlan WHERE Filialempresa = ? AND Ano = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
            If lErro <> AD_SQL_SUCESSO Then gError 106602
            
            'excluir registros de CustoDirFabrProd por filialempresa e ano
            lErro = Comando_Executar(alComando(5), "DELETE FROM CustoDirFabrProd WHERE Filialempresa = ? AND Ano = ?", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
            If lErro <> AD_SQL_SUCESSO Then gError 106603
        
        'se nao houver entao
        Else
        
            'inserir
            lErro = Comando_Executar(alComando(7), "INSERT INTO CustoDirFabr (Filialempresa, Ano, Data) VALUES (?,?,?)", objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno, DATA_NULA)
            If lErro <> AD_SQL_SUCESSO Then gError 106604
            
            'procurar registro em CustoDirFabr por filialempresa e ano (executarpos)
            lErro = Comando_ExecutarPos(alComando(6), "SELECT Data FROM CustoDirFabr WHERE Filialempresa = ? AND Ano = ?", 0, tCustDirFabr.dtData, objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.iAno)
            If lErro <> AD_SQL_SUCESSO Then gError 106626
            
            lErro = Comando_BuscarProximo(alComando(6))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106627
        
            'fazer lock nele
            lErro = Comando_LockExclusive(alComando(6))
            If lErro <> AD_SQL_SUCESSO Then gError 106628
            
        End If
        
        'calcular todos os produtos
        lErro = CustoDiretoFabricacao_Calcula1(objCustoDirFabr, alComando)
        If lErro <> SUCESSO Then gError 106605
        
        'atualizar registro em CustoDirFabr
        With objCustoDirFabr
            lErro = Comando_ExecutarPos(alComando(8), "UPDATE CustoDirFabr SET CodigoPrevVenda = ?, Data = ?, CustoTotal = ?, QuantFator1 = ?, CustoFator1 = ?, QuantFator2 = ?, CustoFator2 = ?, QuantFator3 = ?, CustoFator3 = ?, QuantFator4 = ?, CustoFator4 = ?, QuantFator5 = ?, CustoFator5 = ?, QuantFator6 = ?, CustoFator6 = ?, MesIni = ?, MesFim = ?", _
                alComando(6), .sCodigoPrevVenda, .dtData, .dCustoTotal, .dQuantFator1, .dCustoFator1, .dQuantFator2, .dCustoFator2, .dQuantFator3, .dCustoFator3, .dQuantFator4, .dCustoFator4, .dQuantFator5, .dCustoFator5, .dQuantFator6, .dCustoFator6, .iMesIni, .iMesFim)
            If lErro <> AD_SQL_SUCESSO Then gError 106606
        End With
        
    End If
        
    'Fechamento de transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 106607
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_RATEIO_CUSTOS_DIR_EXECUTADO_SUCESSO")
    
    CustoDiretoFabricacao_Calcula = SUCESSO
     
    Exit Function
    
Erro_CustoDiretoFabricacao_Calcula:

    CustoDiretoFabricacao_Calcula = gErr
     
    Select Case gErr
          
        Case 106590
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 106591
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106592 To 106607, 106618, 106619, 106626 To 106628
            '??? colocar erros adequados
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149253)
     
    End Select
     
    'Erro na transacao
    Call Transacao_Rollback
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function CustoDiretoFabricacao_Calcula1(ByVal objCustoDirFabr As ClassCustoDirFabr, alComando() As Long) As Long
'Se objCustoDirFabr.sProduto vier prenchido recalcula-o, senao calcula todos os produtos

Dim lErro As Long, tPrev As typePrevVendaMensal2, sProdutoLido As String, sProduto As String
Dim dAcumQuantMedia As Double, objCustoDirFabrProd As New ClassCustoDirFabrProd

On Error GoTo Erro_CustoDiretoFabricacao_Calcula1

    sProdutoLido = String(STRING_PRODUTO, 0)
    
    'ler prevvendamensal por filialempresa, ano e codigo da previsao de venda, ordenando por produto
    'descartando produtos que tenha registro em custoembmp para a filialempresa pois neste caso devem ser tratados como se fossem materias cmoprados
    With tPrev
        
        If objCustoDirFabr.sProduto = "" Then
            If objCustoDirFabr.iFilialEmpresa <> EMPRESA_TODA Then
                lErro = Comando_Executar(alComando(9), "SELECT Produto, Quantidade1, Valor1, DataAtualizacao1, Quantidade2, Valor2, DataAtualizacao2, Quantidade3, Valor3, DataAtualizacao3, Quantidade4, Valor4, DataAtualizacao4" & _
                    ", Quantidade5, Valor5, DataAtualizacao5, Quantidade6, Valor6, DataAtualizacao6, Quantidade7, Valor7, DataAtualizacao7, Quantidade8, Valor8, DataAtualizacao8, Quantidade9, Valor9, DataAtualizacao9" & _
                    ", Quantidade10, Valor10, DataAtualizacao10, Quantidade11, Valor11, DataAtualizacao11, Quantidade12, Valor12, DataAtualizacao12" & _
                    " FROM PrevVendaMensal WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? AND NOT EXISTS (SELECT Custo FROM CustoEmbMP WHERE CustoEmbMP.FilialEmpresa = PrevVendaMensal.FilialEmpresa AND CustoEmbMP.Produto = PrevVendaMensal.Produto) ORDER BY Produto", _
                    sProdutoLido, .adQuantidade(1), .adValor(1), .adtDataAtualizacao(1), .adQuantidade(2), .adValor(2), .adtDataAtualizacao(2), .adQuantidade(3), .adValor(3), .adtDataAtualizacao(3), .adQuantidade(4), .adValor(4), .adtDataAtualizacao(4), _
                    .adQuantidade(5), .adValor(5), .adtDataAtualizacao(5), .adQuantidade(6), .adValor(6), .adtDataAtualizacao(6), .adQuantidade(7), .adValor(7), .adtDataAtualizacao(7), .adQuantidade(8), .adValor(8), .adtDataAtualizacao(8), .adQuantidade(9), .adValor(9), .adtDataAtualizacao(9), _
                    .adQuantidade(10), .adValor(10), .adtDataAtualizacao(10), .adQuantidade(11), .adValor(11), .adtDataAtualizacao(11), .adQuantidade(12), .adValor(12), .adtDataAtualizacao(12), _
                    objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.sCodigoPrevVenda, objCustoDirFabr.iAno)
            Else
                lErro = Comando_Executar(alComando(9), "SELECT Produto, Quantidade1, Valor1, DataAtualizacao1, Quantidade2, Valor2, DataAtualizacao2, Quantidade3, Valor3, DataAtualizacao3, Quantidade4, Valor4, DataAtualizacao4" & _
                    ", Quantidade5, Valor5, DataAtualizacao5, Quantidade6, Valor6, DataAtualizacao6, Quantidade7, Valor7, DataAtualizacao7, Quantidade8, Valor8, DataAtualizacao8, Quantidade9, Valor9, DataAtualizacao9" & _
                    ", Quantidade10, Valor10, DataAtualizacao10, Quantidade11, Valor11, DataAtualizacao11, Quantidade12, Valor12, DataAtualizacao12" & _
                    " FROM PrevVendaMensal WHERE Codigo = ? AND Ano = ?  AND NOT EXISTS (SELECT Custo FROM CustoEmbMP WHERE CustoEmbMP.FilialEmpresa = PrevVendaMensal.FilialEmpresa AND CustoEmbMP.Produto = PrevVendaMensal.Produto) ORDER BY Produto", _
                    sProdutoLido, .adQuantidade(1), .adValor(1), .adtDataAtualizacao(1), .adQuantidade(2), .adValor(2), .adtDataAtualizacao(2), .adQuantidade(3), .adValor(3), .adtDataAtualizacao(3), .adQuantidade(4), .adValor(4), .adtDataAtualizacao(4), _
                    .adQuantidade(5), .adValor(5), .adtDataAtualizacao(5), .adQuantidade(6), .adValor(6), .adtDataAtualizacao(6), .adQuantidade(7), .adValor(7), .adtDataAtualizacao(7), .adQuantidade(8), .adValor(8), .adtDataAtualizacao(8), .adQuantidade(9), .adValor(9), .adtDataAtualizacao(9), _
                    .adQuantidade(10), .adValor(10), .adtDataAtualizacao(10), .adQuantidade(11), .adValor(11), .adtDataAtualizacao(11), .adQuantidade(12), .adValor(12), .adtDataAtualizacao(12), _
                    objCustoDirFabr.sCodigoPrevVenda, objCustoDirFabr.iAno)
            End If
        Else
            'se objCustoDirFabr tiver produto definido filtrar por ele
            If objCustoDirFabr.iFilialEmpresa <> EMPRESA_TODA Then
                lErro = Comando_Executar(alComando(9), "SELECT Produto, Quantidade1, Valor1, DataAtualizacao1, Quantidade2, Valor2, DataAtualizacao2, Quantidade3, Valor3, DataAtualizacao3, Quantidade4, Valor4, DataAtualizacao4" & _
                    ", Quantidade5, Valor5, DataAtualizacao5, Quantidade6, Valor6, DataAtualizacao6, Quantidade7, Valor7, DataAtualizacao7, Quantidade8, Valor8, DataAtualizacao8, Quantidade9, Valor9, DataAtualizacao9" & _
                    ", Quantidade10, Valor10, DataAtualizacao10, Quantidade11, Valor11, DataAtualizacao11, Quantidade12, Valor12, DataAtualizacao12" & _
                    " FROM PrevVendaMensal WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? AND Produto = ?  AND NOT EXISTS (SELECT Custo FROM CustoEmbMP WHERE CustoEmbMP.FilialEmpresa = PrevVendaMensal.FilialEmpresa AND CustoEmbMP.Produto = PrevVendaMensal.Produto) ORDER BY Produto", _
                    sProdutoLido, .adQuantidade(1), .adValor(1), .adtDataAtualizacao(1), .adQuantidade(2), .adValor(2), .adtDataAtualizacao(2), .adQuantidade(3), .adValor(3), .adtDataAtualizacao(3), .adQuantidade(4), .adValor(4), .adtDataAtualizacao(4), _
                    .adQuantidade(5), .adValor(5), .adtDataAtualizacao(5), .adQuantidade(6), .adValor(6), .adtDataAtualizacao(6), .adQuantidade(7), .adValor(7), .adtDataAtualizacao(7), .adQuantidade(8), .adValor(8), .adtDataAtualizacao(8), .adQuantidade(9), .adValor(9), .adtDataAtualizacao(9), _
                    .adQuantidade(10), .adValor(10), .adtDataAtualizacao(10), .adQuantidade(11), .adValor(11), .adtDataAtualizacao(11), .adQuantidade(12), .adValor(12), .adtDataAtualizacao(12), _
                    objCustoDirFabr.iFilialEmpresa, objCustoDirFabr.sCodigoPrevVenda, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
            Else
                lErro = Comando_Executar(alComando(9), "SELECT Produto, Quantidade1, Valor1, DataAtualizacao1, Quantidade2, Valor2, DataAtualizacao2, Quantidade3, Valor3, DataAtualizacao3, Quantidade4, Valor4, DataAtualizacao4" & _
                    ", Quantidade5, Valor5, DataAtualizacao5, Quantidade6, Valor6, DataAtualizacao6, Quantidade7, Valor7, DataAtualizacao7, Quantidade8, Valor8, DataAtualizacao8, Quantidade9, Valor9, DataAtualizacao9" & _
                    ", Quantidade10, Valor10, DataAtualizacao10, Quantidade11, Valor11, DataAtualizacao11, Quantidade12, Valor12, DataAtualizacao12" & _
                    " FROM PrevVendaMensal WHERE Codigo = ? AND Ano = ? AND Produto = ?  AND NOT EXISTS (SELECT Custo FROM CustoEmbMP WHERE CustoEmbMP.FilialEmpresa = PrevVendaMensal.FilialEmpresa AND CustoEmbMP.Produto = PrevVendaMensal.Produto) ORDER BY Produto", _
                    sProdutoLido, .adQuantidade(1), .adValor(1), .adtDataAtualizacao(1), .adQuantidade(2), .adValor(2), .adtDataAtualizacao(2), .adQuantidade(3), .adValor(3), .adtDataAtualizacao(3), .adQuantidade(4), .adValor(4), .adtDataAtualizacao(4), _
                    .adQuantidade(5), .adValor(5), .adtDataAtualizacao(5), .adQuantidade(6), .adValor(6), .adtDataAtualizacao(6), .adQuantidade(7), .adValor(7), .adtDataAtualizacao(7), .adQuantidade(8), .adValor(8), .adtDataAtualizacao(8), .adQuantidade(9), .adValor(9), .adtDataAtualizacao(9), _
                    .adQuantidade(10), .adValor(10), .adtDataAtualizacao(10), .adQuantidade(11), .adValor(11), .adtDataAtualizacao(11), .adQuantidade(12), .adValor(12), .adtDataAtualizacao(12), _
                    objCustoDirFabr.sCodigoPrevVenda, objCustoDirFabr.iAno, objCustoDirFabr.sProduto)
            End If
        End If
    
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106608
        
    'ler 1o registro
    lErro = Comando_BuscarProximo(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106609
    
    'se houver algum registro entao
    If lErro = AD_SQL_SUCESSO Then
    
        'setar produto sendo processado
        sProduto = sProdutoLido
        
        'loop enquanto houver registros
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de produto entao
            If sProduto <> sProdutoLido Then
            
                lErro = CF("FPreco_ConvQtdeOrcCusto", sProduto, dAcumQuantMedia)
                If lErro <> SUCESSO Then gError 124121
                
                'processá -lo
                lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, dAcumQuantMedia)
                If lErro <> SUCESSO Then gError 106610
                            
                'setar produto sendo processado
                sProduto = sProdutoLido
                
                'zerar acumuldores de qtde e de meses do produto
                dAcumQuantMedia = 0
            
            End If
            
            'acumular qtde media do produto para os meses validos
            Call CustoDiretoFabricacao_Calcula2(dAcumQuantMedia, tPrev, objCustoDirFabr.iMesIni, objCustoDirFabr.iMesFim)
            
            lErro = Comando_BuscarProximo(alComando(9))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106611
        
        Loop
        
        lErro = CF("FPreco_ConvQtdeOrcCusto", sProduto, dAcumQuantMedia)
        If lErro <> SUCESSO Then gError 124122
                
        'processar o ultimo produto
        lErro = CF("CustoDiretoFabricacao_Prod_Passo1", objCustoDirFabr, objCustoDirFabrProd, sProduto, dAcumQuantMedia)
        If lErro <> SUCESSO Then gError 106612

    End If
    
    CustoDiretoFabricacao_Calcula1 = SUCESSO
     
    Exit Function
    
Erro_CustoDiretoFabricacao_Calcula1:

    CustoDiretoFabricacao_Calcula1 = gErr
     
    Select Case gErr
          
        Case 106608 To 106612, 124121, 124122
            '??? colocar erros adequados
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149254)
     
    End Select
     
    Exit Function

End Function
        
'Private Function CustoDiretoFabricacao_Prod_Passo1(ByVal objCustoDirFabr As ClassCustoDirFabr, objCustoDirFabrProd As ClassCustoDirFabrProd, ByVal sProduto As String, ByVal dQtdeMedia As Double) As Long
''Processar sProduto e seus componentes
'
'Dim lErro As Long, objCustoDirFabrPlan As New ClassCustoDirFabrPlan, objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP
'
'On Error GoTo Erro_CustoDiretoFabricacao_Prod_Passo1
'
'    If dQtdeMedia <> 0 Then
'
'        Set objCustoDirFabrProd = New ClassCustoDirFabrProd
'        objCustoDirFabrProd.iFilialEmpresa = objCustoDirFabr.iFilialEmpresa
'        objCustoDirFabrProd.iAno = objCustoDirFabr.iAno
'        objCustoDirFabrProd.dtData = gdtDataHoje
'        objCustoDirFabrProd.dQuantPrevista = dQtdeMedia
'        objCustoDirFabrProd.sProduto = sProduto
'
'        'inicializar objCustoDirFabrPlan com dados do produto colocando-o como filho de si mesmo com participacao de 100%
'        objCustoDirFabrPlan.iFilialEmpresa = objCustoDirFabrProd.iFilialEmpresa
'        objCustoDirFabrPlan.iAno = objCustoDirFabrProd.iAno
'        objCustoDirFabrPlan.dQuantPai = dQtdeMedia
'        objCustoDirFabrPlan.dParticipacao = 1 '100%
'        objCustoDirFabrPlan.sProduto = objCustoDirFabrProd.sProduto
'        objCustoDirFabrPlan.sProdutoPai = objCustoDirFabrProd.sProduto
'        objCustoDirFabrPlan.sProdutoFilho = objCustoDirFabrProd.sProduto
'
'        'incluir uma linha referente ao proprio produto e mandar obter dados de seus componentes fabricados recursivamente
'        lErro = CustoDiretoFabricacao_Prod_Passo1A(objCustoDirFabr, objCustoDirFabrProd, objCustoDirFabrPlan, alComando)
'        If lErro <> SUCESSO Then gError 106613
'
'        'gravar registro em CustoDirFabrProd
'        With objCustoDirFabrProd
'            lErro = Comando_Executar(alComando(10), "INSERT INTO CustoDirFabrProd (FilialEmpresa, Ano, Produto, Data, QuantPrevista, QuantFator1, QuantFator2, QuantFator3, QuantFator4) VALUES (?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .dtData, .dQuantPrevista, .dQuantFator1, .dQuantFator2, .dQuantFator3, .dQuantFator4)
'        End With
'        If lErro <> SUCESSO Then gError 106614
'
'        For Each objCustoDirFabrPlan In objCustoDirFabrProd.colPlan
'
'            'gravar registros em CustoDirFabrPlan
'            With objCustoDirFabrPlan
'                lErro = Comando_Executar(alComando(11), "INSERT INTO CustoDirFabrPlan (FilialEmpresa, Ano, Produto, Seq,  ProdutoPai, ProdutoFilho, QuantPai, Participacao, PesoFator1, PesoFator2, PesoFator3, PesoFator4) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .iSeq, .sProdutoPai, .sProdutoFilho, .dQuantPai, .dParticipacao, .dPesoFator1, .dPesoFator2, .dPesoFator3, .dPesoFator4)
'            End With
'            If lErro <> SUCESSO Then gError 106615
'
'            For Each objCustoDirFabrPlanMP In objCustoDirFabrPlan.colMP
'                With objCustoDirFabrPlanMP
'                    lErro = Comando_Executar(alComando(12), "INSERT INTO CustoDirFabrPlanMP (FilialEmpresa, Ano, Produto, Seq, ProdutoMP, Qtde) VALUES (?,?,?,?,?,?)", .iFilialEmpresa, .iAno, .sProduto, .iSeq, .sProdutoMP, .dQtde)
'                End With
'                If lErro <> SUCESSO Then gError 106617
'            Next
'
'        Next
'
'        'se nao estiver reprocessando apenas 1 produto entao atualizar acumuladores de quantidades multiplicadas pelos fatores em objCustoDirFabr
'        If objCustoDirFabr.sProduto = "" Then
'            objCustoDirFabr.dQuantFator1 = objCustoDirFabr.dQuantFator1 + objCustoDirFabrProd.dQuantFator1
'            objCustoDirFabr.dQuantFator2 = objCustoDirFabr.dQuantFator2 + objCustoDirFabrProd.dQuantFator2
'            objCustoDirFabr.dQuantFator3 = objCustoDirFabr.dQuantFator3 + objCustoDirFabrProd.dQuantFator3
'            objCustoDirFabr.dQuantFator4 = objCustoDirFabr.dQuantFator4 + objCustoDirFabrProd.dQuantFator4
'        End If
'
'    End If
'
'    CustoDiretoFabricacao_Prod_Passo1 = SUCESSO
'
'    Exit Function
'
'Erro_CustoDiretoFabricacao_Prod_Passo1:
'
'    CustoDiretoFabricacao_Prod_Passo1 = gErr
'
'    Select Case gErr
'
'        Case 106613 To 106615, 106617
'            '??? colocar erros adequados
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149255)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function CustoDiretoFabricacao_Prod_Passo1A(ByVal objCustoDirFabr As ClassCustoDirFabr, ByVal objCustoDirFabrProd As ClassCustoDirFabrProd, ByVal objCustoDirFabrPlan As ClassCustoDirFabrPlan, alComando() As Long) As Long
''procesa o produto filho de objCustoDirFabrPlan, incorpora-o em objCustoDirFabrProd e se necessario processa os seus filhos
'
'Dim lErro As Long, dPesoFator1 As Double, dPesoFator2 As Double, dPesoFator3 As Double, dPesoFator4 As Double, sVersao As String
'Dim colCustoDirFabrPlan As New Collection, dQuantFator1 As Double, dQuantFator2 As Double, dQuantFator3 As Double, dQuantFator4 As Double
'Dim iNivel As Integer, sInsumo As String, dQuantidade As Double, dPercentualPerda As Double, dQuantidadeBase As Double, dQuantidadeProprio As Double, iCompras As Integer
'Dim objCustoDirFabrPlanAux As ClassCustoDirFabrPlan, bAchou As Boolean, dQuantPai As Double, objCustoDirFabrPlanMPAux As ClassCustoDirFabrPlanMP
'
'On Error GoTo Erro_CustoDiretoFabricacao_Prod_Passo1A
'
'    'pesquisa na versao padrao na tabela kit os pesos dos fatores
'    sVersao = String(STRING_KIT_VERSAO, 0)
'    lErro = Comando_Executar(alComando(13), "SELECT Versao, PesoFator1, PesoFator2, PesoFator3, PesoFator4 FROM Kit WHERE ProdutoRaiz = ? AND Situacao = ?", _
'        sVersao, dPesoFator1, dPesoFator2, dPesoFator3, dPesoFator4, objCustoDirFabrPlan.sProdutoFilho, KIT_SITUACAO_PADRAO)
'    If lErro <> AD_SQL_SUCESSO Then gError 106620
'
'    lErro = Comando_BuscarProximo(alComando(13))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106621
'
'    'completar objCustoDirFabrPlan
'    objCustoDirFabrPlan.dPesoFator1 = dPesoFator1
'    objCustoDirFabrPlan.dPesoFator2 = dPesoFator2
'    objCustoDirFabrPlan.dPesoFator3 = dPesoFator3
'    objCustoDirFabrPlan.dPesoFator4 = dPesoFator4
'    objCustoDirFabrPlan.iSeq = objCustoDirFabrProd.colPlan.Count + 1
'
'    'incluir objCustoDirFabrPlan na colecao de objCustoDirFabrProd
'    objCustoDirFabrProd.colPlan.Add objCustoDirFabrPlan
'
'    'acumular qtdes multiplicadas pelos fatores em objCustoDirFabrProd
'    dQuantPai = objCustoDirFabrPlan.dQuantPai * objCustoDirFabrPlan.dParticipacao
'    dQuantFator1 = ArredondaMod(objCustoDirFabrPlan.dPesoFator1 * dQuantPai, 0)
'    dQuantFator2 = ArredondaMod(objCustoDirFabrPlan.dPesoFator2 * dQuantPai, 0)
'    dQuantFator3 = ArredondaMod(objCustoDirFabrPlan.dPesoFator3 * dQuantPai, 0)
'    dQuantFator4 = ArredondaMod(objCustoDirFabrPlan.dPesoFator4 * dQuantPai, 0)
'
'    objCustoDirFabrProd.dQuantFator1 = objCustoDirFabrProd.dQuantFator1 + dQuantFator1
'    objCustoDirFabrProd.dQuantFator2 = objCustoDirFabrProd.dQuantFator2 + dQuantFator2
'    objCustoDirFabrProd.dQuantFator3 = objCustoDirFabrProd.dQuantFator3 + dQuantFator3
'    objCustoDirFabrProd.dQuantFator4 = objCustoDirFabrProd.dQuantFator4 + dQuantFator4
'
'    'pesquisa na versao padrao do produto filho de objCustoDirFabrPlan os componentes de nivel <= 1 ordenados por nivel, seq
'    sInsumo = String(STRING_PRODUTO, 0)
'    lErro = Comando_Executar(alComando(14), "SELECT Compras, Produto, ProdutoKit.Nivel, Quantidade, PercentualPerda FROM ProdutoKit, Produtos WHERE Versao = ? AND ProdutoKit.ProdutoRaiz = ? AND ProdutoKit.Nivel <= 1 AND ProdutoKit.Produto = Produtos.Codigo ORDER BY ProdutoKit.Nivel, Seq", _
'        iCompras, sInsumo, iNivel, dQuantidade, dPercentualPerda, sVersao, objCustoDirFabrPlan.sProdutoFilho)
'    If lErro <> AD_SQL_SUCESSO Then gError 106622
'
'    lErro = Comando_BuscarProximo(alComando(14))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106623
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        'se o nivel for 0 entao
'        If iNivel = 0 Then
'            dQuantidadeBase = dQuantidade
'        Else
'            'se o produto é produzido
'            If iCompras = PRODUTO_NAO_COMPRAVEL Then
'                'para cada insumo encontrado que nao seja o proprio produto
'                If sInsumo <> objCustoDirFabrPlan.sProdutoFilho Then
'
'                    'inclui-lo em colCustoDirFabrPlan se ainda nao estiver lá. se estiver acumular a qtde
'                    bAchou = False
'                    For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
'                        If objCustoDirFabrPlanAux.sProdutoFilho = sInsumo Then
'                            objCustoDirFabrPlanAux.dParticipacao = objCustoDirFabrPlanAux.dParticipacao + (dQuantidade / dQuantidadeBase)
'                            bAchou = True
'                            Exit For
'                        End If
'                    Next
'
'                    If bAchou = False Then
'                        Set objCustoDirFabrPlanAux = New ClassCustoDirFabrPlan
'
'                        objCustoDirFabrPlanAux.iFilialEmpresa = objCustoDirFabrProd.iFilialEmpresa
'                        objCustoDirFabrPlanAux.iAno = objCustoDirFabrProd.iAno
'                        objCustoDirFabrPlanAux.dQuantPai = ArredondaMod(dQuantPai, 0)
'                        objCustoDirFabrPlanAux.dParticipacao = dQuantidade / dQuantidadeBase
'                        objCustoDirFabrPlanAux.sProduto = objCustoDirFabrProd.sProduto
'                        objCustoDirFabrPlanAux.sProdutoPai = objCustoDirFabrPlan.sProdutoFilho
'                        objCustoDirFabrPlanAux.sProdutoFilho = sInsumo
'
'                        colCustoDirFabrPlan.Add objCustoDirFabrPlanAux
'                    End If
'
'                Else
'
'                    'acumular a sua participacao
'                    dQuantidadeProprio = dQuantidadeProprio + dQuantidade
'
'                End If
'
'            Else
'
'                'se ainda nao estiver na colecao objCustoDirFabrPlan.colMP entao incluir, se estiver acumular a qtde
'                bAchou = False
'                For Each objCustoDirFabrPlanMPAux In objCustoDirFabrPlan.colMP
'                    If objCustoDirFabrPlanMPAux.sProdutoMP = sInsumo Then
'                        objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde + ArredondaMod(dQuantPai * dQuantidade / dQuantidadeBase, 0)
'                        bAchou = True
'                        Exit For
'                    End If
'                Next
'
'                If bAchou = False Then
'
'                    Set objCustoDirFabrPlanMPAux = New ClassCustoDirFabrPlanMP
'
'                    objCustoDirFabrPlanMPAux.iFilialEmpresa = objCustoDirFabrPlan.iFilialEmpresa
'                    objCustoDirFabrPlanMPAux.iAno = objCustoDirFabrPlan.iAno
'                    objCustoDirFabrPlanMPAux.sProduto = objCustoDirFabrPlan.sProduto
'                    objCustoDirFabrPlanMPAux.iSeq = objCustoDirFabrPlan.iSeq
'                    objCustoDirFabrPlanMPAux.sProdutoMP = sInsumo
'                    objCustoDirFabrPlanMPAux.dQtde = ArredondaMod(dQuantPai * dQuantidade / dQuantidadeBase, 0)
'
'                    objCustoDirFabrPlan.colMP.Add objCustoDirFabrPlanMPAux
'
'                End If
'
'            End If
'
'        End If
'
'        lErro = Comando_BuscarProximo(alComando(14))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106624
'
'    Loop
'
'    'se o produto pai for insumo dele mesmo
'    If dQuantidadeProprio <> 0 Then
'
'        'percorrer colCustoDirFabrPlan dividindo pela participacao do proprio produto
'        For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
'
'            objCustoDirFabrPlanAux.dParticipacao = objCustoDirFabrPlanAux.dParticipacao * (dQuantidadeBase / dQuantidadeProprio)
'
'        Next
'
'        'percorrer objCustoDirFabrPlan.colMP dividindo pela participacao do proprio produto
'        For Each objCustoDirFabrPlanMPAux In objCustoDirFabrPlan.colMP
'
'            objCustoDirFabrPlanMPAux.dQtde = objCustoDirFabrPlanMPAux.dQtde * (dQuantidadeBase / dQuantidadeProprio)
'
'        Next
'
'    End If
'
'    'forçar o processamento dos filhos
'    For Each objCustoDirFabrPlanAux In colCustoDirFabrPlan
'
'        lErro = CustoDiretoFabricacao_Prod_Passo1A(objCustoDirFabr, objCustoDirFabrProd, objCustoDirFabrPlanAux, alComando)
'        If lErro <> SUCESSO Then gError 106625
'
'    Next
'
'    CustoDiretoFabricacao_Prod_Passo1A = SUCESSO
'
'    Exit Function
'
'Erro_CustoDiretoFabricacao_Prod_Passo1A:
'
'    CustoDiretoFabricacao_Prod_Passo1A = gErr
'
'    Select Case gErr
'
'        Case 106620 To 106625
'            '??? colocar erros adequados
'            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149256)
'
'    End Select
'
'    Exit Function
'
'End Function
'


