VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoFixoRot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private gobjCustoDirFabr As ClassCustoDirFabr
Private gcolCustoFixoProd As Collection
Private gcolItensCategoria As Collection

Private gdTaxaDescPadrao As Double 'taxa p/trazer custo de insumos a valor presente
Private gdTaxaValFut As Double 'taxa p/levar custo de insumos a valor futuro
Private gdDiasValFut As Double 'numero de dias p/levar custo de insumos a valor futuro

Private gcolProd As New Collection

'??? verificar se existem produtos que nao estao classificados p/rateio de custo fixo
'??? verificar se existem mps, embalagens, rotulos,... que nao tenham custo definido
'??? verificar se existem produtos sem os fatores informados
'??? verificar produtos que ficaram em custodirfabrprod mas nao tinham kit definido, como o 2214.

Function Rotina_CustoFixo_Calcula(ByVal sNomeArqParam As String, ByVal objCustoFixo As ClassCustoFixo) As Long

Dim lErro As Long

On Error GoTo Erro_Rotina_CustoFixo_Calcula

    lErro = Sistema_Abrir_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 106642

    lErro = CustoFixo_Calcula(objCustoFixo)
    If lErro <> SUCESSO Then gError 106643
        
    Rotina_CustoFixo_Calcula = SUCESSO
     
    Exit Function
    
Erro_Rotina_CustoFixo_Calcula:

    Rotina_CustoFixo_Calcula = gErr
     
    Select Case gErr
          
        Case 106642, 106643
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149257)
     
    End Select
     
    Exit Function

End Function

Private Function CustoFixo_Calcula(ByVal objCustoFixo As ClassCustoFixo) As Long
    
Dim lErro As Long, iIndice As Integer
Dim lTransacao As Long, alComando(1 To 16) As Long

On Error GoTo Erro_CustoFixo_Calcula

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 106761
 
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir
        If lErro <> SUCESSO Then gError 106762
    Next
    
    lErro = CustoFixo_Calcula1(objCustoFixo, alComando)
    If lErro <> SUCESSO Then gError 106762

    lErro = CustoFixo_Calcula2(objCustoFixo, alComando)
    If lErro <> SUCESSO Then gError 106771
                                
    lErro = CustoFixo_Calcula3(objCustoFixo, alComando)
    If lErro <> SUCESSO Then gError 106772
                                
    lErro = CustoFixo_Calcula6(objCustoFixo)
    If lErro <> SUCESSO Then gError 106773
    
    lErro = CustoFixo_Calcula7(objCustoFixo, alComando)
    If lErro <> SUCESSO Then gError 106774
                                
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 106763
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CustoFixo_Calcula = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula:

    CustoFixo_Calcula = gErr
     
    Select Case gErr
          
        Case 106761
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 106762
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 106763
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 106771 To 106774
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149258)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function CustoFixo_Calcula1(ByVal objCustoFixo As ClassCustoFixo, alComando() As Long) As Long
'limpar dados de execucao anterior para a mesma filialempresa e data
'bloquear p/evitar problemas de concorrencia

Dim lErro As Long, dtDataAtualizacao As Date

On Error GoTo Erro_CustoFixo_Calcula1

de_novo:

    'pesquisa se o custo já foi rodado anteriormente p/esta filialempresa-data de referencia
    lErro = Comando_ExecutarPos(alComando(1), "SELECT DataAtualizacao FROM CustoFixo WHERE FilialEmpresa = ? AND DataReferencia = ?", 0, dtDataAtualizacao, objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia)
    If lErro <> AD_SQL_SUCESSO Then gError 106764
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106765
    
    'se houver
    If lErro = AD_SQL_SUCESSO Then
    
        'fazer lock
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 106766
        
        'exluir os registros de custofixoprod com Automatico=1
        lErro = Comando_Executar(alComando(2), "DELETE FROM CustoFixoProd WHERE FilialEmpresa = ? AND DataReferencia = ? AND Automatico = ?", objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia, CUSTOFIXOPROD_AUTOMATICO)
        If lErro <> AD_SQL_SUCESSO Then gError 106767
        
        'percorrer em custofixoprod alterando custocalculado para zero.
        lErro = Comando_Executar(alComando(3), "UPDATE CustoFixoProd Set CustoCalculado = 0 WHERE FilialEmpresa = ? AND DataReferencia = ?", objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia)
        If lErro <> AD_SQL_SUCESSO Then gError 106768
    
        'alterar atributos da tabela custofixo
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE CustoFixo SET CustoFixo = ?, CustoFixoPapel = ?, DataAtualizacao = ?", alComando(1), objCustoFixo.dCusto1, objCustoFixo.dCusto2, objCustoFixo.dtDataAtualizacao)
        If lErro <> AD_SQL_SUCESSO Then gError 106769
        
    Else
    
        'inserir registro na tabela custofixo
        lErro = Comando_Executar(alComando(4), "INSERT INTO CustoFixo (FilialEmpresa, DataReferencia, CustoFixo, CustoFixoPapel, DataAtualizacao) VALUES (?,?,?,?,?)", objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia, objCustoFixo.dCusto1, objCustoFixo.dCusto2, objCustoFixo.dtDataAtualizacao)
        If lErro <> AD_SQL_SUCESSO Then gError 106770
    
        GoTo de_novo
        
    End If
    
    CustoFixo_Calcula1 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula1:

    CustoFixo_Calcula1 = gErr
     
    Select Case gErr
          
        '??? completar
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149259)
     
    End Select
     
    Exit Function

End Function

Private Function CustoFixo_Calcula2(ByVal objCustoFixo As ClassCustoFixo, alComando() As Long) As Long
'bloquear registro da tabela CustoDirFabr, se houver (p/evitar alguem rodando custos diretos simultaneamente a custos fixos)
'guardar info sobre custos diretos globais

Dim lErro As Long, tcustodir As typeCustoDirFabr

On Error GoTo Erro_CustoFixo_Calcula2

    With tcustodir
        .sCodigoPrevVenda = String(STRING_PREVVENDA_CODIGO, 0)
        lErro = Comando_ExecutarLockado(alComando(6), _
            "SELECT CodigoPrevVenda, CustoTotal, QuantFator1, CustoFator1, QuantFator2, CustoFator2, QuantFator3, CustoFator3, QuantFator4, CustoFator4, QuantFator5, CustoFator5, QuantFator6, CustoFator6, MesIni, MesFim FROM CustoDirFabr WHERE FilialEmpresa = ? AND Ano = ?", _
                .sCodigoPrevVenda, .dCustoTotal, .dQuantFator1, .dCustoFator1, .dQuantFator2, .dCustoFator2, .dQuantFator3, .dCustoFator3, .dQuantFator4, .dCustoFator4, .dQuantFator5, .dCustoFator5, .dQuantFator6, .dCustoFator6, .iMesIni, .iMesFim, objCustoFixo.iFilialEmpresa, Year(objCustoFixo.dtDataReferencia))
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106775
    
    lErro = Comando_BuscarProximo(alComando(6))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106776
    
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockShared(alComando(6))
        If lErro <> AD_SQL_SUCESSO Then gError 106777
    
        Set gobjCustoDirFabr = New ClassCustoDirFabr
        
        With tcustodir
            gobjCustoDirFabr.sCodigoPrevVenda = .sCodigoPrevVenda
            gobjCustoDirFabr.dCustoFator1 = .dCustoFator1
            gobjCustoDirFabr.dCustoFator2 = .dCustoFator2
            gobjCustoDirFabr.dCustoFator3 = .dCustoFator3
            gobjCustoDirFabr.dCustoFator4 = .dCustoFator4
            gobjCustoDirFabr.dCustoFator5 = .dCustoFator5
            gobjCustoDirFabr.dCustoFator6 = .dCustoFator6
            gobjCustoDirFabr.dQuantFator1 = .dQuantFator1
            gobjCustoDirFabr.dQuantFator2 = .dQuantFator2
            gobjCustoDirFabr.dQuantFator3 = .dQuantFator3
            gobjCustoDirFabr.dQuantFator4 = .dQuantFator4
            gobjCustoDirFabr.dQuantFator5 = .dQuantFator5
            gobjCustoDirFabr.dQuantFator6 = .dQuantFator6
            gobjCustoDirFabr.dCustoTotal = .dCustoTotal
            gobjCustoDirFabr.iAno = Year(objCustoFixo.dtDataReferencia)
            gobjCustoDirFabr.iFilialEmpresa = objCustoFixo.iFilialEmpresa
            gobjCustoDirFabr.iMesIni = .iMesIni
            gobjCustoDirFabr.iMesFim = .iMesFim
        End With
        
        With gobjCustoDirFabr
            .dCustoOutros = .dCustoTotal - (.dCustoFator1 + .dCustoFator2 + .dCustoFator3 + .dCustoFator4 + .dCustoFator5 + .dCustoFator6)
        End With
                
    End If

    lErro = CF("CalcMP_ObterValores", objCustoFixo.iFilialEmpresa, gdTaxaDescPadrao, gdTaxaValFut, gdDiasValFut)
    If lErro <> SUCESSO Then gError 106884
    
'    'obter valores globais para calculo de valor presente e futuro de insumo
'    sMnemonico = String(STRING_MNEMONICOFPRECO_MNEMONICO, 0)
'    sExpressao = String(STRING_MNEMONICOFPRECO_EXPRESSAO, 0)
'    lErro = Comando_Executar(alComando(7), "SELECT Mnemonico, Expressao FROM MnemonicoFPTipo WHERE TipoPlanilha = ? AND FilialEmpresa = ? AND Escopo = ? AND Mnemonico IN ('TaxaDescPadrao', 'TaxaValFut', 'iDiasValFut' )", _
'        sMnemonico, sExpressao, PLANILHA_TIPO_TODOS, objCustoFixo.iFilialEmpresa, MNEMONICOFPRECO_ESCOPO_GERAL)
'    If lErro <> AD_SQL_SUCESSO Then gError 106778
'
'    lErro = Comando_BuscarProximo(alComando(7))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106779
'
'    Do While lErro = AD_SQL_SUCESSO
'
'        Select Case sMnemonico
'
'            Case "TaxaDescPadrao"
'                gdTaxaDescPadrao = StrParaDbl(sExpressao)
'
'            Case "TaxaValFut"
'                gdTaxaValFut = StrParaDbl(sExpressao)
'
'            Case "DiasValFut"
'                gdDiasValFut = StrParaDbl(sExpressao)
'
'        End Select
'
'        lErro = Comando_BuscarProximo(alComando(7))
'        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106780
'
'    Loop
    
    CustoFixo_Calcula2 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula2:

    CustoFixo_Calcula2 = gErr
     
    Select Case gErr
          
        '??? completar
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149260)
     
    End Select
     
    Exit Function

End Function

Private Function CustoFixo_Calcula3(ByVal objCustoFixo As ClassCustoFixo, alComando() As Long) As Long

Dim lErro As Long, tPrev As typePrevVendaMensal2, sProdutoLido As String, sItemLido As String
Dim sProduto As String, sItem As String, dAcumQuantMedia As Double
Dim iIndiceCateg As Integer, dAcumCustoItemCateg As Double, dCustoFabricacao As Double
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem

On Error GoTo Erro_CustoFixo_Calcula3

    Set gcolItensCategoria = New Collection
    objCategoriaProduto.sCategoria = gobjFAT.sCategCustoFixo
    lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, gcolItensCategoria)
    If lErro <> SUCESSO And lErro <> 22541 Then gError 106781
    
    If gcolItensCategoria.Count = 0 Then
        With objCategoriaProdutoItem
            .dvalor1 = 1
            .dvalor2 = 1
        End With
        gcolItensCategoria.Add objCategoriaProdutoItem
    End If
    
    sProdutoLido = String(STRING_PRODUTO, 0)
    sItemLido = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)
    
    With tPrev
        lErro = Comando_Executar(alComando(8), "SELECT Item, PrevVendaMensal.Produto, Quantidade1, Valor1, DataAtualizacao1, Quantidade2, Valor2, DataAtualizacao2, Quantidade3, Valor3, DataAtualizacao3, Quantidade4, Valor4, DataAtualizacao4" & _
            ", Quantidade5, Valor5, DataAtualizacao5, Quantidade6, Valor6, DataAtualizacao6, Quantidade7, Valor7, DataAtualizacao7, Quantidade8, Valor8, DataAtualizacao8, Quantidade9, Valor9, DataAtualizacao9" & _
            ", Quantidade10, Valor10, DataAtualizacao10, Quantidade11, Valor11, DataAtualizacao11, Quantidade12, Valor12, DataAtualizacao12" & _
            " FROM PrevVendaMensal LEFT OUTER JOIN ProdutoCategoria ON PrevVendaMensal.Produto = ProdutoCategoria.Produto AND ProdutoCategoria.Categoria = ? WHERE FilialEmpresa = ? AND Codigo = ? AND Ano = ? ORDER BY Item, PrevVendaMensal.Produto", _
            sItemLido, sProdutoLido, .adQuantidade(1), .adValor(1), .adtDataAtualizacao(1), .adQuantidade(2), .adValor(2), .adtDataAtualizacao(2), .adQuantidade(3), .adValor(3), .adtDataAtualizacao(3), .adQuantidade(4), .adValor(4), .adtDataAtualizacao(4), _
            .adQuantidade(5), .adValor(5), .adtDataAtualizacao(5), .adQuantidade(6), .adValor(6), .adtDataAtualizacao(6), .adQuantidade(7), .adValor(7), .adtDataAtualizacao(7), .adQuantidade(8), .adValor(8), .adtDataAtualizacao(8), .adQuantidade(9), .adValor(9), .adtDataAtualizacao(9), _
            .adQuantidade(10), .adValor(10), .adtDataAtualizacao(10), .adQuantidade(11), .adValor(11), .adtDataAtualizacao(11), .adQuantidade(12), .adValor(12), .adtDataAtualizacao(12), _
            gobjFAT.sCategCustoFixo, gobjCustoDirFabr.iFilialEmpresa, gobjCustoDirFabr.sCodigoPrevVenda, gobjCustoDirFabr.iAno)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 106782
    
    lErro = Comando_BuscarProximo(alComando(8))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106783
    
    'se houver algum registro entao
    If lErro = AD_SQL_SUCESSO Then
    
        'setar produto sendo processado
        sProduto = sProdutoLido
        sItem = sItemLido
        lErro = ObtemIndiceCateg(sItem, iIndiceCateg)
        If lErro <> SUCESSO Then gError 106784
        
        'loop enquanto houver registros
        Do While lErro = AD_SQL_SUCESSO
        
            'se trocou de produto entao
            If sProduto <> sProdutoLido Then
            
                lErro = CF("FPreco_ConvQtdeOrcCusto", sProduto, dAcumQuantMedia)
                If lErro <> SUCESSO Then gError 124126
                                        
                'processá -lo
                lErro = CustoFixo_Calcula4(objCustoFixo, sProduto, dAcumQuantMedia, iIndiceCateg, dAcumCustoItemCateg, alComando)
                If lErro <> SUCESSO Then gError 106785
                            
                'setar produto sendo processado
                sProduto = sProdutoLido
                
                'zerar acumulador de qtde do produto
                dAcumQuantMedia = 0
            
                'se trocou de item na categoria
                If sItem <> sItemLido Then
                
                    'atribuir a valor3 na col dAcumCustoItemCateg
                    gcolItensCategoria.Item(iIndiceCateg).dvalor3 = dAcumCustoItemCateg
                    
                    dAcumCustoItemCateg = 0
                    
                    sItem = sItemLido
                    lErro = ObtemIndiceCateg(sItem, iIndiceCateg)
                    If lErro <> SUCESSO Then gError 106786
                    
                End If
                
            End If
            
            'acumular qtde media do produto para os meses validos
            Call CustoDiretoFabricacao_Calcula2(dAcumQuantMedia, tPrev, gobjCustoDirFabr.iMesIni, gobjCustoDirFabr.iMesFim)
            
            lErro = Comando_BuscarProximo(alComando(8))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106787
        
        Loop
    
        lErro = CF("FPreco_ConvQtdeOrcCusto", sProduto, dAcumQuantMedia)
        If lErro <> SUCESSO Then gError 124127
                                        
        'processar o ultimo produto
        lErro = CustoFixo_Calcula4(objCustoFixo, sProduto, dAcumQuantMedia, iIndiceCateg, dAcumCustoItemCateg, alComando)
        If lErro <> SUCESSO Then gError 106788
    
        gcolItensCategoria.Item(iIndiceCateg).dvalor3 = dAcumCustoItemCateg
        
    End If
    
    CustoFixo_Calcula3 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula3:

    CustoFixo_Calcula3 = gErr
     
    Select Case gErr
          
        Case 124126, 124127
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149261)
     
    End Select
     
    Exit Function
    
End Function
        
Private Function CustoFixo_Calcula4(ByVal objCustoFixo As ClassCustoFixo, ByVal sProduto As String, ByVal dAcumQuantMedia As Double, ByVal iIndiceCateg As Integer, dAcumCustoItemCateg As Double, alComando() As Long) As Long
'calcula o custo de fabricacao de um produto

Dim lErro As Long, tcustoemb As typeCustoEmbMP, tcustodirprod As typeCustoDirFabrProd, dFator As Double
Dim objCustoDirFabrProd As New ClassCustoDirFabrProd, dCustoProd As Double
Dim dCustoMP As Double, dCustoEmb As Double
Dim objCustoFixoRotProd As ClassCustoFixoRotProd, colMP As Collection, colEmb As Collection
Dim objCustoDirFabrPlanMP As ClassCustoDirFabrPlanMP, objCustoDirFabrPlanEmb As ClassCustoDirFabrPlanEmb

On Error GoTo Erro_CustoFixo_Calcula4

    If dAcumQuantMedia <> 0 Then
    
        'procura custo dele em custoembmp (vai ser util para produtos revendidos, ou seja, nao produzidos)
        With tcustoemb
            lErro = Comando_Executar(alComando(9), "SELECT Custo, AliquotaICMS, CondicaoPagto, FretePorKG FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", _
                .dCusto, .dAliquotaICMS, .iCondicaoPagto, .dFretePorKg, objCustoFixo.iFilialEmpresa, sProduto)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 106789
        
        lErro = Comando_BuscarProximo(alComando(9))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106790
        
        'se nao achou
        If lErro <> AD_SQL_SUCESSO Then
        
            lErro = CF("CustoDirFabrProd_ObtemInfo", objCustoDirFabrProd, gobjCustoDirFabr, alComando(10))
            If lErro <> SUCESSO Then gError 106827
                
            'se o produto fez parte do rateio de custos diretos de fabricacao
            If objCustoDirFabrProd.dQuantPrevista <> 0 Then dCustoProd = objCustoDirFabrProd.dCusto / objCustoDirFabrProd.dQuantPrevista * dAcumQuantMedia
            
            Set colMP = New Collection
            Set colEmb = New Collection
                        
            'obtem as colecoes de materia-prima e de embalagens e seus acessorios (embalagens de expedicao, rotulos, palletizacao,...) necessarios para produzir sProduto na Filialempresa
            lErro = CF("CustoDireto_ObtemEmbMP", objCustoFixo.iFilialEmpresa, gobjCustoDirFabr.iAno, sProduto, dAcumQuantMedia, colEmb, colMP, FORMACAO_PRECO_ROTINA_CUSTOFIXO)
            If lErro <> SUCESSO Then gError 106826
            
            'para cada materia prima
            For Each objCustoDirFabrPlanMP In colMP
            
                'obter o seu custo em custembmp.
                lErro = CF("CustoDiretoProd_ObtemValor", objCustoFixo.iFilialEmpresa, objCustoDirFabrPlanMP.sProdutoMP, gdTaxaDescPadrao, gdTaxaValFut, gdDiasValFut, dCustoMP, alComando(12))
                If lErro <> SUCESSO Then gError 106797
                
                'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
                dCustoProd = dCustoProd + (dCustoMP * objCustoDirFabrPlanMP.dQtde)
            
            Next
            
            'para cada embalagem
            For Each objCustoDirFabrPlanEmb In colEmb
            
                If objCustoDirFabrPlanEmb.sProdutoEmb <> "" Then
                
                    'obter o seu custo em custembmp.
                    lErro = CF("CustoDiretoProd_ObtemValor", objCustoFixo.iFilialEmpresa, objCustoDirFabrPlanEmb.sProdutoEmb, gdTaxaDescPadrao, gdTaxaValFut, gdDiasValFut, dCustoEmb, alComando(16))
                    If lErro <> SUCESSO Then gError 106842
                    
                    'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
                    dCustoProd = dCustoProd + (dCustoEmb * objCustoDirFabrPlanEmb.dQtde)
            
                End If
                
            Next
            
        Else
            
            'obter o seu custo em custembmp.
            lErro = CF("CustoDiretoProd_ObtemValor", objCustoFixo.iFilialEmpresa, sProduto, gdTaxaDescPadrao, gdTaxaValFut, gdDiasValFut, dCustoMP, alComando(12))
            If lErro <> SUCESSO Then gError 106800
                    
            'com a qtde e custo unitario, acumulo o custo do insumo no custo do produto.
            dCustoProd = dCustoMP * dAcumQuantMedia
            
        End If
        
        dAcumCustoItemCateg = dAcumCustoItemCateg + dCustoProd
        
        Set objCustoFixoRotProd = New ClassCustoFixoRotProd
        
        objCustoFixoRotProd.sProduto = sProduto
        objCustoFixoRotProd.dCustoFabricacao = dCustoProd
        objCustoFixoRotProd.iIndice = iIndiceCateg
        objCustoFixoRotProd.dQuantidade = dAcumQuantMedia
        
        gcolProd.Add objCustoFixoRotProd
    
    End If
    
    CustoFixo_Calcula4 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula4:

    CustoFixo_Calcula4 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149262)
     
    End Select
     
    Exit Function
    
End Function

Sub CustoDirFabrProd_Calcula(ByVal objCustoDirFabr As ClassCustoDirFabr, ByVal objCustoDirFabrProd As ClassCustoDirFabrProd, dCustoDireto As Double)

Dim lErro As Long, dCusto1 As Double, dCusto2 As Double, dCusto3 As Double, dCusto4 As Double, dCusto5 As Double, dCusto6 As Double, dCustoOutros As Double

On Error GoTo Erro_CustoDirFabrProd_Calcula

    If objCustoDirFabr.dQuantFator1 <> 0 Then
    
        dCusto1 = ((objCustoDirFabrProd.dQuantFator1 * objCustoDirFabr.dCustoFator1) / objCustoDirFabr.dQuantFator1)
    
    Else
    
        dCusto1 = 0
    
    End If
    
    If objCustoDirFabr.dQuantFator2 <> 0 Then
    
        dCusto2 = ((objCustoDirFabrProd.dQuantFator2 * objCustoDirFabr.dCustoFator2) / objCustoDirFabr.dQuantFator2)
    
    Else
    
        dCusto2 = 0
    
    End If
    
    If objCustoDirFabr.dQuantFator3 <> 0 Then
    
        dCusto3 = ((objCustoDirFabrProd.dQuantFator3 * objCustoDirFabr.dCustoFator3) / objCustoDirFabr.dQuantFator3)
    
    Else
    
        dCusto3 = 0
    
    End If
    
    If objCustoDirFabr.dQuantFator4 <> 0 Then
    
        dCusto4 = ((objCustoDirFabrProd.dQuantFator4 * objCustoDirFabr.dCustoFator4) / objCustoDirFabr.dQuantFator4)
    
    Else
    
        dCusto4 = 0
    
    End If
    
    If objCustoDirFabr.dQuantFator5 <> 0 Then
    
        dCusto5 = ((objCustoDirFabrProd.dQuantFator5 * objCustoDirFabr.dCustoFator5) / objCustoDirFabr.dQuantFator5)
    
    Else
    
        dCusto5 = 0
    
    End If
    
    If objCustoDirFabr.dQuantFator6 <> 0 Then
    
        dCusto6 = ((objCustoDirFabrProd.dQuantFator6 * objCustoDirFabr.dCustoFator6) / objCustoDirFabr.dQuantFator6)
    
    Else
    
        dCusto6 = 0
    
    End If
    
    dCustoOutros = (dCusto1 + dCusto2 + dCusto3 + dCusto4 + dCusto5 + dCusto6) / (objCustoDirFabr.dCustoTotal - objCustoDirFabr.dCustoOutros) * objCustoDirFabr.dCustoOutros
    
    dCustoDireto = dCusto1 + dCusto2 + dCusto3 + dCusto4 + dCusto5 + dCusto6 + dCustoOutros
    
    Exit Sub
     
Erro_CustoDirFabrProd_Calcula:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149263)
     
    End Select
     
    Exit Sub

End Sub

'Private Function CustoFixo_Calcula5(ByVal objCustoFixo As ClassCustoFixo, ByVal sProduto As String, dCusto As Double, alComando() As Long) As Long
''retorna o custo unitario de sProduto
'
'Dim lErro As Long, tcustoemb As typeCustoEmbMP, dCustoAux As Double
'Dim objCondicaoPagto As New ClassCondicaoPagto
'Dim dDias As Double
'
'On Error GoTo Erro_CustoFixo_Calcula5
'
'    With tcustoemb
'        lErro = Comando_Executar(alComando(12), "SELECT Custo, AliquotaICMS, CondicaoPagto, FretePorKG FROM CustoEmbMP WHERE FilialEmpresa = ? AND Produto = ?", _
'            .dCusto, .dAliquotaICMS, .iCondicaoPagto, .dFretePorKg, objCustoFixo.iFilialEmpresa, sProduto)
'    End With
'    If lErro <> AD_SQL_SUCESSO Then gError 106801
'
'    lErro = Comando_BuscarProximo(alComando(12))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106802
'
'    'se nao achou
'    If lErro <> AD_SQL_SUCESSO Then gError 106803
'
'    dCustoAux = tcustoemb.dCusto
'
'    'calcular a valor presente pela taxa gdTaxaDescPadrao com o prazo de pagto medio da condicao de pagto.
'
'    If tcustoemb.iCondicaoPagto <> 0 Then
'
'        objCondicaoPagto.iCodigo = tcustoemb.iCondicaoPagto
'        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
'        If lErro <> SUCESSO And lErro <> 19205 Then gError 106804
'
'        'Se a condição de pagamento não está cadastrada, erro
'        If lErro = 19205 Then gError 106805
'
'        'Calcula a média ponderada de dias para a taxa financeira
'        Call CF("CalculaDias_CondPagamento", objCondicaoPagto, dDias)
'
'        dCustoAux = dCustoAux / ((1 + gdTaxaDescPadrao / 100) ^ (dDias / 30))
'
'    End If
'
'    'abater o icms
'    dCustoAux = dCustoAux * (1 - tcustoemb.dAliquotaICMS)
'
'    'agregar o frete
'    dCustoAux = dCustoAux + tcustoemb.dFretePorKg
'
'    'aplicar fator referente à perda
'    'já deve estar no calculo do custodirfabr
'
'    'levar a valor futuro pela taxa gdTaxaValFut com numero de dias gdDiasValFut
'    dCustoAux = dCustoAux * ((1 + gdTaxaValFut / 100) ^ (gdDiasValFut / 30))
'
'    dCusto = dCustoAux
'
'    CustoFixo_Calcula5 = SUCESSO
'
'    Exit Function
'
'Erro_CustoFixo_Calcula5:
'
'    CustoFixo_Calcula5 = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149264)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function ObtemIndiceCateg(ByVal sItem As String, iIndiceCateg As Integer) As Long
'obtem o indice do item na colecao global de itens da categoria 'custofixo'

Dim lErro As Long, iIndice As Integer
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_ObtemIndiceCateg
    
    iIndiceCateg = 0
    
    For Each objCategoriaProdutoItem In gcolItensCategoria
    
        iIndice = iIndice + 1
        If UCase(sItem) = UCase(objCategoriaProdutoItem.sItem) Then
            iIndiceCateg = iIndice
            Exit For
        End If
    Next
    
    If iIndiceCateg = 0 Then gError 106806
    
    ObtemIndiceCateg = SUCESSO
     
    Exit Function
    
Erro_ObtemIndiceCateg:

    ObtemIndiceCateg = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149265)
     
    End Select
     
    Exit Function

End Function

Private Function CustoFixo_Calcula6(ByVal objCustoFixo As ClassCustoFixo) As Long

Dim lErro As Long, dAcumBase1 As Double, dAcumBase2 As Double
Dim objCategoriaProdutoItem As ClassCategoriaProdutoItem

On Error GoTo Erro_CustoFixo_Calcula6

    'percorrer a colecao de itens da categoria cf obtendo o valor total multiplicado pelo peso do item.
    For Each objCategoriaProdutoItem In gcolItensCategoria
    
        If objCategoriaProdutoItem.dvalor2 = 1 Then
            dAcumBase1 = dAcumBase1 + (objCategoriaProdutoItem.dvalor1 * objCategoriaProdutoItem.dvalor3)
        Else
            dAcumBase2 = dAcumBase2 + (objCategoriaProdutoItem.dvalor1 * objCategoriaProdutoItem.dvalor3)
        End If
        
    Next
    
    'colocar no valor 4 do item o fator p/absorver o custo fixo
    For Each objCategoriaProdutoItem In gcolItensCategoria
    
        If objCategoriaProdutoItem.dvalor2 = 1 Then
                
            If dAcumBase1 <> 0 And objCategoriaProdutoItem.dvalor3 <> 0 Then
                'obtem indice
                objCategoriaProdutoItem.dvalorAux1 = (objCategoriaProdutoItem.dvalor1 * objCategoriaProdutoItem.dvalor3) / dAcumBase1
                'obtem fator
                objCategoriaProdutoItem.dvalorAux2 = (objCategoriaProdutoItem.dvalorAux1 * objCustoFixo.dCusto1) / objCategoriaProdutoItem.dvalor3
            End If
        
        Else
        
            If dAcumBase2 <> 0 And objCategoriaProdutoItem.dvalor3 <> 0 Then
                'obtem indice
                objCategoriaProdutoItem.dvalorAux1 = (objCategoriaProdutoItem.dvalor1 * objCategoriaProdutoItem.dvalor3) / dAcumBase2
                'obtem fator
                objCategoriaProdutoItem.dvalorAux2 = (objCategoriaProdutoItem.dvalorAux1 * objCustoFixo.dCusto2) / objCategoriaProdutoItem.dvalor3
            End If
        
        End If
        
    Next
    
    CustoFixo_Calcula6 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula6:

    CustoFixo_Calcula6 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149266)
     
    End Select
     
    Exit Function

End Function
    
Private Function CustoFixo_Calcula7(ByVal objCustoFixo As ClassCustoFixo, alComando() As Long) As Long
'percorre a colecao com os produtos com custo fixo já calculado e atualiza a tabela CustoFixoProd

Dim lErro As Long, iIndiceCategLido As Integer, iIndiceCateg As Integer
Dim objCustoFixoRotProd As ClassCustoFixoRotProd, dFator As Double
Dim dCustoFixoUnitario As Double, iAutomatico As Integer

On Error GoTo Erro_CustoFixo_Calcula7

    'percorrer a colecao de produtos
    For Each objCustoFixoRotProd In gcolProd
    
        iIndiceCategLido = objCustoFixoRotProd.iIndice
        
        If iIndiceCateg <> iIndiceCategLido Then
        
            'obter fator correspondente
            dFator = gcolItensCategoria.Item(iIndiceCategLido).dvalorAux2
            
            iIndiceCateg = iIndiceCategLido
            
        End If
        
        'obter o custo fixo unitario
        dCustoFixoUnitario = objCustoFixoRotProd.dCustoFabricacao * dFator / objCustoFixoRotProd.dQuantidade
        
        'procurar o registro em custofixoprod
        lErro = Comando_ExecutarPos(alComando(13), "SELECT Automatico FROM CustoFixoProd WHERE FilialEmpresa = ? AND DataReferencia = ? AND Produto = ?", 0, _
            iAutomatico, objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia, objCustoFixoRotProd.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 106807
        
        lErro = Comando_BuscarProximo(alComando(13))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106808
            
        'se encontrar
        If lErro = AD_SQL_SUCESSO Then
                
            'atualizar o valor calculado
            lErro = Comando_ExecutarPos(alComando(14), "UPDATE CustoFixoProd SET CustoCalculado = ?, CustoFabricacao = ?, QtdePrevista = ?", alComando(13), dCustoFixoUnitario, objCustoFixoRotProd.dCustoFabricacao, objCustoFixoRotProd.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 106809
            
        Else
        
            'incluir registro
            lErro = Comando_Executar(alComando(15), "INSERT INTO CustoFixoProd (FilialEmpresa, DataReferencia, Produto, Custo, CustoCalculado, Automatico, CustoFabricacao, QtdePrevista) VALUES (?,?,?,?,?,?,?,?)", _
                objCustoFixo.iFilialEmpresa, objCustoFixo.dtDataReferencia, objCustoFixoRotProd.sProduto, 0, dCustoFixoUnitario, 1, objCustoFixoRotProd.dCustoFabricacao, objCustoFixoRotProd.dQuantidade)
            If lErro <> AD_SQL_SUCESSO Then gError 106810
            
        End If
        
    Next
    
    CustoFixo_Calcula7 = SUCESSO
     
    Exit Function
    
Erro_CustoFixo_Calcula7:

    CustoFixo_Calcula7 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149267)
     
    End Select
     
    Exit Function

End Function

