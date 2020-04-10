VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMargContrCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function RelAnaliseVendas_Prepara(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas) As Long

Dim lErro As Long, sSelect As String, iIndice As Integer
Dim lTransacao As Long, alComando(0 To 1) As Long, lNumIntRel As Long
Dim iVendedor As Integer, lCliente As Long, iFilialCli As Integer, sProduto As String, iClasseUM As Integer, sUnidadeMed As String, iFilialEmpresa As Integer, iFaturamento As Integer, iTipoNFiscal As Integer, iTabelaPreco As Integer, dQuantidade As Double, dValorTotal As Double, sUMBase As String

On Error GoTo Erro_RelAnaliseVendas_Prepara

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130348
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130349

    'obtem o nº automatico
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_ANALISE_VENDAS", lNumIntRel)
    If lErro <> SUCESSO Then gError 130350
    
    objRelAnaliseVendas.lNumIntRel = lNumIntRel
    
    'monta o SELECT
    lErro = RelAnaliseVendas_Prepara1(objRelAnaliseVendas, sSelect)
    If lErro <> SUCESSO Then gError 130351
    
    'executa a preparacao da parte fixa do SELECT
    sProduto = String(STRING_PRODUTO, 0)
    sUnidadeMed = String(STRING_UM_SIGLA, 0)
    sUMBase = String(STRING_UM_SIGLA, 0)
    lErro = RelAnaliseVendas_Prepara2(alComando(0), sSelect, iVendedor, lCliente, iFilialCli, sProduto, iClasseUM, sUnidadeMed, sUMBase, iFilialEmpresa, iFaturamento, iTipoNFiscal, iTabelaPreco, dQuantidade, dValorTotal)
    If lErro <> SUCESSO Then gError 130352
    
    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das parcelas
    lErro = RelAnaliseVendas_Prepara3(objRelAnaliseVendas, alComando(0))
    If lErro <> SUCESSO Then gError 130353
    
    'processa os registros retornados pelo SELECT executado acima
    lErro = RelAnaliseVendas_Prepara4(objRelAnaliseVendas, alComando, iVendedor, lCliente, iFilialCli, sProduto, iClasseUM, sUnidadeMed, sUMBase, iFilialEmpresa, iFaturamento, iTipoNFiscal, iTabelaPreco, dQuantidade, dValorTotal)
    If lErro <> SUCESSO Then gError 130354
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130355
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    RelAnaliseVendas_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara:

    RelAnaliseVendas_Prepara = gErr
     
    Select Case gErr
          
        Case 130348
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 130349
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 130351 To 130354
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151538)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara1(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, sSelect As String) As Long
'monta o SELECT em funcao dos parametros passados

Dim lErro As Long
Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String
Dim sGroupBy As String

On Error GoTo Erro_RelAnaliseVendas_Prepara1
    
    sFields = "NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, TiposDocInfo.Faturamento, TiposDocInfo.Codigo, NFiscal.TabelaPreco, " & _
            "SUM(ItensNFiscal.Quantidade) As Quantidade, SUM(ItensNFiscal.Quantidade * ItensNFiscal.PrecoUnitario - ItensNFiscal.ValorDesconto) AS ValorTotal"
    sFrom = " FROM NFiscal, TiposDocInfo, ItensNFiscal, Produtos, FiliaisClientes, UnidadesdeMedida, ClasseUM"
    sWhere = " WHERE NFiscal.DataEmissao >= ? AND NFiscal.DataEmissao <= ? AND NFiscal.NumIntDoc = ItensNFiscal.NumIntNF AND TiposDocInfo.Codigo = NFiscal.TipoNFiscal AND (TiposDocInfo.Faturamento IN (1, 2)) AND ItensNFiscal.Produto = Produtos.Codigo AND NFiscal.Status <> 7  AND (TiposDocInfo.Complementar = 0 OR TiposDocInfo.Codigo IN (?,?)) AND FiliaisClientes.CodCliente = NFiscal.Cliente AND FiliaisClientes.CodFilial = NFiscal.FilialCli AND UnidadesdeMedida.Classe = ClasseUM.Classe AND UnidadesdeMedida.Sigla = ItensNFiscal.UnidadeMed AND Produtos.ClasseUM = ClasseUM.Classe"
    sOrderBy = " ORDER BY NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase"
    sGroupBy = " GROUP BY NFiscal.FilialEmpresa, FiliaisClientes.Vendedor, NFiscal.Cliente, NFiscal.FilialCli, ItensNFiscal.Produto, Produtos.ClasseUM, ItensNFiscal.UnidadeMed, UnidadesdeMedida.SiglaUMBase, TiposDocInfo.Faturamento, TiposDocInfo.Codigo, NFiscal.TabelaPreco"
    
    If objRelAnaliseVendas.iFilialEmpresa <> EMPRESA_TODA Then
        sWhere = sWhere & " AND NFiscal.FilialEmpresa = ?"
    End If
    
    If objRelAnaliseVendas.iVendedorDe <> 0 Then
        sWhere = sWhere & " AND FiliaisClientes.Vendedor >= ?"
    End If
    
    If objRelAnaliseVendas.iVendedorAte <> 0 Then
        sWhere = sWhere & " AND FiliaisClientes.Vendedor <= ?"
    End If
        
    If objRelAnaliseVendas.lClienteDe <> 0 Then
        sWhere = sWhere & " AND NFiscal.Cliente >= ?"
    End If
    
    If objRelAnaliseVendas.lClienteAte <> 0 Then
        sWhere = sWhere & " AND NFiscal.Cliente <= ?"
    End If
        
    If Len(Trim(objRelAnaliseVendas.sProdutoDe)) <> 0 Then
        sWhere = sWhere & " AND ItensNFiscal.Produto >= ?"
    End If
    
    If Len(Trim(objRelAnaliseVendas.sProdutoAte)) <> 0 Then
        sWhere = sWhere & " AND ItensNFiscal.Produto <= ?"
    End If
        
    sSelect = "SELECT " & sFields & sFrom & sWhere & sGroupBy & sOrderBy
    
    RelAnaliseVendas_Prepara1 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara1:

    RelAnaliseVendas_Prepara1 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151539)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara2(ByVal lComando As Long, ByVal sSelect As String, viVendedor As Variant, vlCliente As Variant, viFilialCli As Variant, vsProduto As Variant, viClasseUM As Variant, vsUnidadeMed As Variant, vsUMBase As Variant, viFilialEmpresa As Variant, viFaturamento As Variant, viTipoNFiscal As Variant, viTabelaPreco As Variant, vdQuantidade As Variant, vdValorTotal As Variant) As Long
'executa a preparacao da parte fixa do SELECT

Dim lErro As Long

On Error GoTo Erro_RelAnaliseVendas_Prepara2
    
    lErro = Comando_PrepararInt(lComando, sSelect)
    If lErro <> AD_SQL_SUCESSO Then gError 130356

    lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 130357

    lErro = Comando_BindVarInt(lComando, viVendedor)
    If lErro <> AD_SQL_SUCESSO Then gError 130358

    lErro = Comando_BindVarInt(lComando, vlCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 130359

    lErro = Comando_BindVarInt(lComando, viFilialCli)
    If lErro <> AD_SQL_SUCESSO Then gError 130360

    lErro = Comando_BindVarInt(lComando, vsProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 130361

    lErro = Comando_BindVarInt(lComando, viClasseUM)
    If lErro <> AD_SQL_SUCESSO Then gError 130362

    lErro = Comando_BindVarInt(lComando, vsUnidadeMed)
    If lErro <> AD_SQL_SUCESSO Then gError 130363

    lErro = Comando_BindVarInt(lComando, vsUMBase)
    If lErro <> AD_SQL_SUCESSO Then gError 130363

    lErro = Comando_BindVarInt(lComando, viFaturamento)
    If lErro <> AD_SQL_SUCESSO Then gError 130364

    lErro = Comando_BindVarInt(lComando, viTipoNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 130365

    lErro = Comando_BindVarInt(lComando, viTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 130366

    lErro = Comando_BindVarInt(lComando, vdQuantidade)
    If lErro <> AD_SQL_SUCESSO Then gError 130367

    lErro = Comando_BindVarInt(lComando, vdValorTotal)
    If lErro <> AD_SQL_SUCESSO Then gError 130368

    RelAnaliseVendas_Prepara2 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara2:

    RelAnaliseVendas_Prepara2 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151540)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara3(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, ByVal lComando As Long) As Long
'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
'e executa o SELECT p/obtencao das parcelas

Dim lErro As Long
Dim viFilialEmpresa As Variant, viVendedorDe As Variant, viVendedorAte As Variant, vlClienteDe As Variant, vlClienteAte As Variant, vsProdutoDe As Variant, vsProdutoAte As Variant
Dim vdtDataEmissaoDe As Variant, vdtDataEmissaoAte As Variant, viTipoComplPreco1 As Variant, viTipoComplPreco2 As Variant

On Error GoTo Erro_RelAnaliseVendas_Prepara3
    
    vdtDataEmissaoDe = objRelAnaliseVendas.dtDataEmissaoDe
    lErro = Comando_BindVarInt(lComando, vdtDataEmissaoDe)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    vdtDataEmissaoAte = objRelAnaliseVendas.dtDataEmissaoAte
    lErro = Comando_BindVarInt(lComando, vdtDataEmissaoAte)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    viTipoComplPreco1 = CInt(DOCINFO_NFISCP)
    lErro = Comando_BindVarInt(lComando, viTipoComplPreco1)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    viTipoComplPreco2 = CInt(DOCINFO_NFISFCP)
    lErro = Comando_BindVarInt(lComando, viTipoComplPreco2)
    If lErro <> AD_SQL_SUCESSO Then gError 130369
    
    If objRelAnaliseVendas.iFilialEmpresa <> EMPRESA_TODA Then
        viFilialEmpresa = objRelAnaliseVendas.iFilialEmpresa
        lErro = Comando_BindVarInt(lComando, viFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 130369
    End If
    
    If objRelAnaliseVendas.iVendedorDe <> 0 Then
        viVendedorDe = objRelAnaliseVendas.iVendedorDe
        lErro = Comando_BindVarInt(lComando, viVendedorDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130370
    End If
    
    If objRelAnaliseVendas.iVendedorAte <> 0 Then
        viVendedorAte = objRelAnaliseVendas.iVendedorAte
        lErro = Comando_BindVarInt(lComando, viVendedorAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130371
    End If
        
    If objRelAnaliseVendas.lClienteDe <> 0 Then
        vlClienteDe = objRelAnaliseVendas.lClienteDe
        lErro = Comando_BindVarInt(lComando, vlClienteDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130372
    End If
    
    If objRelAnaliseVendas.lClienteAte <> 0 Then
        vlClienteAte = objRelAnaliseVendas.lClienteAte
        lErro = Comando_BindVarInt(lComando, vlClienteAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130373
    End If
        
    If Len(Trim(objRelAnaliseVendas.sProdutoDe)) <> 0 Then
        vsProdutoDe = objRelAnaliseVendas.sProdutoDe
        lErro = Comando_BindVarInt(lComando, vsProdutoDe)
        If lErro <> AD_SQL_SUCESSO Then gError 130374
    End If
    
    If Len(Trim(objRelAnaliseVendas.sProdutoAte)) <> 0 Then
        vsProdutoAte = objRelAnaliseVendas.sProdutoAte
        lErro = Comando_BindVarInt(lComando, vsProdutoAte)
        If lErro <> AD_SQL_SUCESSO Then gError 130375
    End If
    
    lErro = Comando_ExecutarInt(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 130376

    RelAnaliseVendas_Prepara3 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara3:

    RelAnaliseVendas_Prepara3 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151541)
     
    End Select
     
    Exit Function

End Function

Private Function RelAnaliseVendas_Prepara4(ByVal objRelAnaliseVendas As ClassRelAnaliseVendas, alComando() As Long, iVendedor As Integer, lCliente As Long, iFilialCli As Integer, sProduto As String, iClasseUM As Integer, sUnidadeMed As String, sUMBase As String, iFilialEmpresa As Integer, iFaturamento As Integer, iTipoNFiscal As Integer, iTabelaPreco As Integer, dQuantidade As Double, dValorTotal As Double) As Long

Dim lErro As Long, dFator As Double, sUMAnaliseAnt As String
Dim iVendedorAnt As Integer, lClienteAnt As Long, iFilialCliAnt As Integer, sProdutoAnt As String, sUMBaseAnt As String, iFilialEmpresaAnt As Integer, iTabelaPrecoAnt As Integer
Dim dQuantidadeAcum As Double, dValorTotalAcum As Double, iFilialEmpresaCalc As Integer
Dim dQuantidadeDevAcum As Double, dValorTotalDevAcum As Double, bFiltraLucroDe As Boolean, bFiltraLucroAte As Boolean, dLucroDe As Double, dLucroAte As Double
Dim objMargContr As New ClassMargContr, dPrecoPraticadoInformado As Double, bInsereRegistro As Boolean

On Error GoTo Erro_RelAnaliseVendas_Prepara4

    If Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Or Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Then
    
        If Len(Trim(objRelAnaliseVendas.sLucroDe)) <> 0 Then
            bFiltraLucroDe = True
            dLucroDe = StrParaDbl(objRelAnaliseVendas.sLucroDe)
        End If
        
        If Len(Trim(objRelAnaliseVendas.sLucroAte)) <> 0 Then
            bFiltraLucroAte = True
            dLucroAte = StrParaDbl(objRelAnaliseVendas.sLucroAte)
        End If
        
    End If
    
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130377

    If lErro = AD_SQL_SEM_DADOS Then gError 130378

    With objMargContr
        .iColunaCli = 3
        .iColunaPraticado = 4
        .iLinhaPercMarg = 17
        .iLinhaPercLucro = 20
        .iLinhaValorMarg = 21
    End With

    iVendedorAnt = iVendedor
    lClienteAnt = lCliente
    iFilialCliAnt = iFilialCli
    sProdutoAnt = sProduto
    sUMBaseAnt = sUMBase
    iFilialEmpresaAnt = iFilialEmpresa
    iTabelaPrecoAnt = iTabelaPreco
            
    Do While lErro = AD_SQL_SUCESSO

        'se trocou algo que compõe a chave do registro em RelAnaliseVendas
        If iVendedorAnt <> iVendedor Or lClienteAnt <> lCliente Or iFilialCliAnt <> iFilialCli Or sProdutoAnt <> sProduto Or iFilialEmpresaAnt <> iFilialEmpresa Then
        
            If dQuantidadeAcum <> 0 Then
                
                dPrecoPraticadoInformado = Arredonda_Moeda(dValorTotalAcum / dQuantidadeAcum, 2)
            
                With objMargContr
                
                    .dQuantidade = dQuantidadeAcum
                    .iCodFilial = iFilialCliAnt
                    .iFilialFaturamento = iFilialEmpresaAnt
                    .lCodCliente = lClienteAnt
                    .sProduto = sProdutoAnt
                    .iTabelaPreco = iTabelaPrecoAnt
                    .sUnidadeMed = sUMBaseAnt
                    .dPrecoPraticadoInformado = dPrecoPraticadoInformado
                    
                End With
    
                lErro = CF("MargContr_Calcula", objMargContr)
                If lErro <> SUCESSO Then gError 130382
                
                bInsereRegistro = True
                
                If bFiltraLucroDe Or bFiltraLucroDe Then
                
                    If bFiltraLucroDe Then
                    
                        Select Case objRelAnaliseVendas.iTipoFiltroProduto
                        
                            Case 1 '% lucro
                                If (objMargContr.dUltPercLucroPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                                
                            Case 2 '% marg
                                If (objMargContr.dUltPercMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                            Case 3 'R$ marg
                                If (objMargContr.dUltValorMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                        End Select
                        
                    End If
                    
                    If bInsereRegistro And bFiltraLucroAte Then
                    
                        Select Case objRelAnaliseVendas.iTipoFiltroProduto
                        
                            Case 1 '% lucro
                                If (objMargContr.dUltPercLucroPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                                
                            Case 2 '% marg
                                If (objMargContr.dUltPercMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                            Case 3 'R$ marg
                                If (objMargContr.dUltValorMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                            
                        End Select
                        
                    End If
                    
                End If
                
                If bInsereRegistro Then
                    lErro = Comando_Executar(alComando(1), "INSERT INTO RelAnaliseVendas (NumIntRel, Vendedor, Cliente, Filial, Produto, FilialEmpresa, UnidadeMed, Quantidade, PrecoVenda, PrecoCliente, PercMargem, PercLucro, ValorMargem, QtdeDevolvida, ValorDevolvido, TabelaDePreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                        objRelAnaliseVendas.lNumIntRel, iVendedorAnt, lClienteAnt, iFilialCliAnt, sProdutoAnt, iFilialEmpresaAnt, sUMBaseAnt, dQuantidadeAcum, dPrecoPraticadoInformado, objMargContr.dUltPrecoCli, objMargContr.dUltPercMargPrat, objMargContr.dUltPercLucroPrat, objMargContr.dUltValorMargPrat, dQuantidadeDevAcum, dValorTotalDevAcum, iTabelaPrecoAnt)
                    If lErro <> AD_SQL_SUCESSO Then gError 130390
                End If
            
            End If
            
            dQuantidadeAcum = 0
            dValorTotalAcum = 0
            dQuantidadeDevAcum = 0
            dValorTotalDevAcum = 0
            
            iVendedorAnt = iVendedor
            lClienteAnt = lCliente
            iFilialCliAnt = iFilialCli
            sProdutoAnt = sProduto
            iFilialEmpresaAnt = iFilialEmpresa
            sUMBaseAnt = sUMBase
            
        End If
            
        'Faturada
        If iFaturamento = 1 Then
            dValorTotalAcum = dValorTotalAcum + dValorTotal
            
'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 130383

                If iTipoNFiscal <> DOCINFO_NFISCP And iTipoNFiscal <> DOCINFO_NFISFCP Then
                    dQuantidadeAcum = dQuantidadeAcum + (dQuantidade * dFator)
                End If
'            End If

        Else
            'Devolvida
            dValorTotalDevAcum = dValorTotalDevAcum + dValorTotal

'            If sUnidadeMed <> "PC" And sUnidadeMed <> "DIV" And sUnidadeMed <> "TB" Then
                lErro = CF("UM_Conversao", iClasseUM, sUnidadeMed, sUMBase, dFator)
                If lErro <> SUCESSO Then gError 130384
                
                dQuantidadeDevAcum = dQuantidadeDevAcum + (dQuantidade * dFator)
'            End If
        
        End If
        
        iTabelaPrecoAnt = iTabelaPreco
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130385

    Loop
    
    If dQuantidadeAcum <> 0 Then
        dPrecoPraticadoInformado = Arredonda_Moeda(dValorTotalAcum / dQuantidadeAcum, 2)
            
        With objMargContr
        
            .dQuantidade = dQuantidadeAcum
            .iCodFilial = iFilialCliAnt
            .iFilialFaturamento = iFilialEmpresaAnt
            .lCodCliente = lClienteAnt
            .sProduto = sProdutoAnt
            .iTabelaPreco = iTabelaPrecoAnt
            .sUnidadeMed = sUMBaseAnt
            .dPrecoPraticadoInformado = dPrecoPraticadoInformado
            
        End With
    
        lErro = CF("MargContr_Calcula", objMargContr)
        If lErro <> SUCESSO Then gError 130391
        
        bInsereRegistro = True
        
        If bFiltraLucroDe Or bFiltraLucroDe Then
        
            If bFiltraLucroDe Then
            
                Select Case objRelAnaliseVendas.iTipoFiltroProduto
                
                    Case 1 '% lucro
                        If (objMargContr.dUltPercLucroPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                        
                    Case 2 '% marg
                        If (objMargContr.dUltPercMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                    Case 3 'R$ marg
                        If (objMargContr.dUltValorMargPrat - dLucroDe) < DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                End Select
                
            End If
            
            If bInsereRegistro And bFiltraLucroAte Then
            
                Select Case objRelAnaliseVendas.iTipoFiltroProduto
                
                    Case 1 '% lucro
                        If (objMargContr.dUltPercLucroPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                        
                    Case 2 '% marg
                        If (objMargContr.dUltPercMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                    Case 3 'R$ marg
                        If (objMargContr.dUltValorMargPrat - dLucroAte) > DELTA_VALORMONETARIO2 Then bInsereRegistro = False
                    
                End Select
                
            End If
            
        End If
        
        If bInsereRegistro Then
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelAnaliseVendas (NumIntRel, Vendedor, Cliente, Filial, Produto, FilialEmpresa, UnidadeMed, Quantidade, PrecoVenda, PrecoCliente, PercMargem, PercLucro, ValorMargem, QtdeDevolvida, ValorDevolvido, TabelaDePreco) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objRelAnaliseVendas.lNumIntRel, iVendedorAnt, lClienteAnt, iFilialCliAnt, sProdutoAnt, iFilialEmpresaAnt, sUMBaseAnt, dQuantidadeAcum, dPrecoPraticadoInformado, objMargContr.dUltPrecoCli, objMargContr.dUltPercMargPrat, objMargContr.dUltPercLucroPrat, objMargContr.dUltValorMargPrat, dQuantidadeDevAcum, dValorTotalDevAcum, iTabelaPrecoAnt)
            If lErro <> AD_SQL_SUCESSO Then gError 130392
           
        End If
        
    End If
    
    RelAnaliseVendas_Prepara4 = SUCESSO
     
    Exit Function
    
Erro_RelAnaliseVendas_Prepara4:

    RelAnaliseVendas_Prepara4 = gErr
     
    Select Case gErr
          
        Case 130378
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_VENDAS_NO_PERIODO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151542)
     
    End Select
     
    Exit Function

End Function

Private Function MargContr_CalculaDVV(ByVal objMargContr As ClassMargContr, ByVal iColuna As Integer) As Long

Dim lErro As Long, iLinha As Integer, objDVVLinCol As ClassDVVLinCol
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double
Dim objDVVLinColEsq As ClassDVVLinCol

On Error GoTo Erro_MargContr_CalculaDVV

    If objMargContr.colDVVLin.Count <> 0 Then
    
        For iLinha = 1 To objMargContr.colDVVLin.Count
                
            Set objDVVLinCol = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 1))
            
            If Len(Trim(objDVVLinCol.sFormula)) = 0 Then
            
                If iColuna > 1 Then
                
                    Set objDVVLinColEsq = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 2))
                    sFormula = objDVVLinColEsq.sFormula
                    If Len(Trim(sFormula)) = 0 Then
                    
                        If iColuna > 2 Then
                        
                            Set objDVVLinColEsq = objMargContr.colDVVLinCol(objMargContr.IndDVV(iLinha, iColuna - 3))
                            sFormula = objDVVLinColEsq.sFormula
                            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
                            
                        Else
                        
                            sFormula = "0"
                            
                        End If
                        
                    End If
                Else
                
                    sFormula = "0"
                    
                End If
            
            Else
            
                sFormula = objDVVLinCol.sFormula
                
            End If
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_DVV
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .sNomeRedCliente = objMargContr.sUltNomeRedCli
            .sUM = objMargContr.sUnidadeMed
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 130046
    
        Set objPlanilhas = colPlanilhas.Item(colPlanilhas.Count)
        
        Select Case (iColuna - 1)
        
            Case 1
                objMargContr.sDVVTotal1 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
            Case 2
                objMargContr.sDVVTotal2 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
            Case 3
                objMargContr.sDVVTotal3 = Format(objPlanilhas.dValor * 100, "###,##0.00###")
                
        End Select
    
    End If
    
    MargContr_CalculaDVV = SUCESSO
     
    Exit Function
    
Erro_MargContr_CalculaDVV:

    MargContr_CalculaDVV = gErr
     
    Select Case gErr
          
        Case 130046
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151543)
     
    End Select
     
    Exit Function

End Function

Private Function Analise_RecalcularColuna(ByVal objMargContr As ClassMargContr, ByVal iColuna As Integer, ByVal iColunaCli As Integer, ByVal iColunaPraticado As Integer) As Long

Dim lErro As Long, iLinha As Integer, objLinCol As ClassPlanMargContrLinCol, objLin As ClassPlanMargContrLin
Dim colPlanilhas As New Collection, objPlanilhas As ClassPlanilhas, sFormula As String
Dim objContexto As New ClassContextoPlan, dValor As Double, colPlanilhas2 As New Collection
Dim objMnemonicoValor As ClassMnemonicoValor, objContexto2 As New ClassContextoPlan, iCol As Integer
Dim dValorL1 As Double

On Error GoTo Erro_Analise_RecalcularColuna

    'calcular o preco para o cliente se a coluna for iColunaCli
    If iColuna = iColunaCli Then
    
        With objContexto2
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .dTaxaDescPadrao = objMargContr.dTaxaDescPadrao
            .dTaxaValFut = objMargContr.dTaxaValFut
            .dDiasValFut = objMargContr.dDiasValFut
        End With
        
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objMargContr, objContexto2)
        
        For iLinha = 1 To objMargContr.colPlanMargContrLin.Count
        
            Set objLin = objMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = objMargContr.colPlanMargContrLinCol(objMargContr.IndAnalise(iLinha, iColuna - 1))
            
            sFormula = objLinCol.sFormula
            If iLinha = 1 And Len(Trim(sFormula)) <> 0 Then Exit For
            If Len(Trim(objLin.sFormulaL1)) = 0 And Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas2.Add objPlanilhas
            
            'se possui formula p/calculo da linha 1 e o valor está preenchido...
            If Len(Trim(objLin.sFormulaL1)) <> 0 And ((Len(Trim(objLinCol.sFormula)) <> 0)) Then
            
                Set objPlanilhas = New ClassPlanilhas
                
                With objPlanilhas
                    .iTipoPlanilha = PLANILHA_TIPO_TODOS
                    .iFilialEmpresa = objMargContr.iFilialFaturamento
                    .iEscopo = MNEMONICOFPRECO_ESCOPO_GERAL
                    .iLinha = iLinha + 1
                    .sExpressao = objLin.sFormulaL1
                End With
                
                colPlanilhas2.Add objPlanilhas
        
                'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
                lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas2, dValor, objContexto2)
                If lErro <> SUCESSO Then gError 106721

                dValorL1 = dValor
                                
                Exit For
                
            End If
            
        Next
        
    End If
    
    If iColuna = iColunaCli Then
     
        objMargContr.dUltPrecoCli = dValorL1
        
    Else
    
        'guarda os resultados do grid dvv para todas as colunas
        Call DVVTotal_GuardaValores(objMargContr, objContexto)
        
        For iLinha = 1 To objMargContr.colPlanMargContrLin.Count
        
            Set objLin = objMargContr.colPlanMargContrLin(iLinha)
            Set objLinCol = objMargContr.colPlanMargContrLinCol(objMargContr.IndAnalise(iLinha, iColuna - 1))
                
            If iLinha = 1 And Len(Trim(objLinCol.sFormula)) = 0 And Len(Trim(objLin.sFormulaGeral)) = 0 And objLin.iEditavel <> 0 Then
            
                'pega conteudo da tela, da propria celula
                sFormula = CStr(dValorL1)
                If Len(Trim(sFormula)) <> 0 And objLin.iFormato = GRID_FORMATO_PERCENTAGEM Then sFormula = CStr(StrParaDbl(sFormula) / 100)
                Call TrocaPontoVirgula(sFormula)
                
            Else
            
                sFormula = objLinCol.sFormula
                If Len(Trim(sFormula)) = 0 Then sFormula = objLin.sFormulaGeral
            
            End If
            
            If Len(Trim(sFormula)) = 0 Then sFormula = "0"
            
            Set objPlanilhas = New ClassPlanilhas
            
            With objPlanilhas
                .iTipoPlanilha = PLANILHA_TIPO_TODOS
                .iFilialEmpresa = objMargContr.iFilialFaturamento
                .iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO
                .iLinha = iLinha
                .sExpressao = sFormula
            End With
            
            colPlanilhas.Add objPlanilhas
        
        Next
        
        With objContexto
            .iFilialFaturamento = objMargContr.iFilialFaturamento
            .sProduto = objMargContr.sProduto
            .dQuantidade = objMargContr.dQuantidade
            .iFilialCli = objMargContr.iCodFilial
            .lCliente = objMargContr.lCodCliente
            .iVendedor = objMargContr.iCodVendedor
            .iTabelaPreco = objMargContr.iTabelaPreco
            .iAno = Year(gdtDataAtual)
            .dtDataCustoFixo = objMargContr.dtDataCF
            .iRotinaOrigem = FORMACAO_PRECO_ANALISE_MARGCONTR_REL
            Set .colComissoes = objMargContr.colComissoes
            .dPrecoPraticado = objMargContr.dPrecoComissoes
            .dTaxaDescPadrao = objMargContr.dTaxaDescPadrao
            .dTaxaValFut = objMargContr.dTaxaValFut
            .dDiasValFut = objMargContr.dDiasValFut
        End With
        
        'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor
        lErro = CF("Avalia_Expressao_FPreco3", colPlanilhas, dValor, objContexto)
        If lErro <> SUCESSO Then gError 106721
    
        For Each objPlanilhas In colPlanilhas
        
            If objPlanilhas.iLinha = 1 Then
                objMargContr.dUltPrecoPraticado = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
            
            If objPlanilhas.iLinha = objMargContr.iLinhaPercMarg Then
                objMargContr.dUltPercMargPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = objMargContr.iLinhaPercLucro Then
                objMargContr.dUltPercLucroPrat = StrParaDbl(Format(objPlanilhas.dValor * 100, "###,##0.00###"))
            End If
    
            If objPlanilhas.iLinha = objMargContr.iLinhaValorMarg Then
                objMargContr.dUltValorMargPrat = StrParaDbl(Format(objPlanilhas.dValor, "###,##0.00###"))
            End If
    
        Next
    
    End If
    
    Analise_RecalcularColuna = SUCESSO
     
    Exit Function
    
Erro_Analise_RecalcularColuna:

    Analise_RecalcularColuna = gErr
     
    Select Case gErr
          
        Case 106721
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151544)
     
    End Select
     
    Exit Function

End Function

Private Sub DVVTotal_GuardaValores(ByVal objMargContr As ClassMargContr, objContexto As ClassContextoPlan)

Dim iCol As Integer, objMnemonicoValor As ClassMnemonicoValor

    'guarda os resultados do grid dvv para todas as colunas
    For iCol = 1 To 3
    
        Set objMnemonicoValor = New ClassMnemonicoValor
        Set objMnemonicoValor.colValor = New Collection
        
        objMnemonicoValor.sMnemonico = "DVVTotal"
        objMnemonicoValor.vParam(1) = CDbl(iCol)
    
        Select Case iCol
        
            Case 1
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal1) / 100)
            Case 2
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal2) / 100)
            Case 3
                objMnemonicoValor.colValor.Add StrParaDbl(StrParaDbl(objMargContr.sDVVTotal3) / 100)
        
        End Select
        
        objContexto.colMnemonicoValor.Add objMnemonicoValor
        
    Next

End Sub

Private Sub TrocaPontoVirgula(sNumero As String)
Dim iTam As Integer, sResult As String, i As Integer, sCaracter As String

    For i = 1 To Len(sNumero)
    
        sCaracter = Mid(sNumero, i, 1)
        Select Case sCaracter
        
            Case ","
                sCaracter = "."
                
            Case "."
                sCaracter = ""
            
        End Select
        
        sResult = sResult & sCaracter
        
    Next
        
    sNumero = sResult

End Sub

Private Function MargContr_TestaDVV(ByVal objMargContr As ClassMargContr) As Long

Dim lErro As Long, dPercDVV As Double, lComando As Long

On Error GoTo Erro_MargContr_TestaDVV

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 106666
    
    lErro = Comando_Executar(lComando, "SELECT PercDVV FROM DVVClienteProd WHERE FilialEmpresa = ? AND CodCliente = ? AND CodFilial = ? AND Produto = ?", dPercDVV, objMargContr.iFilialFaturamento, objMargContr.lCodCliente, objMargContr.iCodFilial, objMargContr.sProduto)
    If lErro <> AD_SQL_SUCESSO Then gError 106667

    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 106668
    If lErro <> AD_SQL_SUCESSO Then gError 130055
    
    Call Comando_Fechar(lComando)
    
    MargContr_TestaDVV = SUCESSO
     
    Exit Function
    
Erro_MargContr_TestaDVV:

    MargContr_TestaDVV = gErr
     
    Select Case gErr
          
        Case 106666
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 106667, 106668
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DVVCLIENTEPROD", gErr)

        Case 130055
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151545)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Public Function MargContr_Calcula(ByVal objMargContr As ClassMargContr) As Long

Dim lErro As Long, objTelaComissoes As Object, objCodigoNome As New AdmlCodigoNome
Dim dtDataCFTemp As Date, dTaxaDescPadraoTemp As Double, dTaxaValFutTemp As Double, dDiasValFutTemp As Double
Dim colComissoes As Collection, dPrecoComissoes As Double

On Error GoTo Erro_MargContr_Calcula

    If objMargContr.objTelaComissoes Is Nothing Then
    
        lErro = CF("MargContr_Le_DVV", objMargContr)
        If lErro <> SUCESSO Then gError 130379
        
        lErro = CF("MargContr_Le_Analise", objMargContr)
        If lErro <> SUCESSO Then gError 130380
        
        lErro = Chama_Tela_Nova_Instancia1("ComissoesCalcula", objTelaComissoes)
        If lErro <> SUCESSO Then gError 130381
        
        Set objMargContr.objTelaComissoes = objTelaComissoes
    
    Else
    
        Set objTelaComissoes = objMargContr.objTelaComissoes
        
    End If
    
    'se trocou de filialempresa
    If objMargContr.iUltFilialEmpresa <> objMargContr.iFilialFaturamento Then
        
        lErro = CF("FilialEmpresa_ObtemDataCustoFixo", objMargContr.iFilialFaturamento, dtDataCFTemp)
        If lErro <> SUCESSO Then gError 130386
    
        lErro = CF("CalcMP_ObterValores", objMargContr.iFilialFaturamento, dTaxaDescPadraoTemp, dTaxaValFutTemp, dDiasValFutTemp)
        If lErro <> SUCESSO Then gError 130387
        
        With objMargContr
            .dtDataCF = dtDataCFTemp
            .dTaxaDescPadrao = dTaxaDescPadraoTemp
            .dTaxaValFut = dTaxaValFutTemp
            .dDiasValFut = dDiasValFutTemp
        End With
        objMargContr.iUltFilialEmpresa = objMargContr.iFilialFaturamento
        
    End If

    'se trocou de cliente
    If objMargContr.lUltCliente <> objMargContr.lCodCliente Then
    
        objCodigoNome.lCodigo = objMargContr.lCodCliente
        lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
        If lErro <> SUCESSO And lErro <> 12553 Then gError 130388
        If lErro <> SUCESSO Then gError 130389
        
        objMargContr.sUltNomeRedCli = objCodigoNome.sNome
        
        objMargContr.lUltCliente = objMargContr.lCodCliente
        
    End If

    lErro = MargContr_TestaDVV(objMargContr)
    If lErro <> SUCESSO And lErro <> 130055 Then
    
        objMargContr.dUltPrecoCli = 0
        objMargContr.dUltPrecoPraticado = 0
        objMargContr.dUltPercMargPrat = 0
        objMargContr.dUltPercLucroPrat = 0
        objMargContr.dUltValorMargPrat = 0
    
    Else
    
        lErro = CF("MargContr_CalculaComissoes", colComissoes, objMargContr, dPrecoComissoes, objTelaComissoes, objMargContr.sUltNomeRedCli, objMargContr.sUnidadeMed)
        If lErro <> SUCESSO Then gError 130045
        
        Set objMargContr.colComissoes = colComissoes
        objMargContr.dPrecoComissoes = dPrecoComissoes
    
        'calcula a DVV para o cliente
        lErro = MargContr_CalculaDVV(objMargContr, objMargContr.iColunaCli)
        If lErro <> SUCESSO Then gError 130047
    
        'calcula a DVV para o preco praticado
        lErro = MargContr_CalculaDVV(objMargContr, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130052
    
        lErro = Analise_RecalcularColuna(objMargContr, objMargContr.iColunaCli, objMargContr.iColunaCli, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130053
        
        lErro = Analise_RecalcularColuna(objMargContr, objMargContr.iColunaPraticado, objMargContr.iColunaCli, objMargContr.iColunaPraticado)
        If lErro <> SUCESSO Then gError 130054
    
    End If

    MargContr_Calcula = SUCESSO
     
    Exit Function
    
Erro_MargContr_Calcula:

    MargContr_Calcula = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 151546)
     
    End Select
     
    Exit Function

End Function

