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
Option Explicit

Function RelComisCro_Prepara(lNumIntRel As Long, ByVal iFilialEmpresa As Integer, ByVal dtDataInicio As Date, ByVal dtDataFim As Date, ByVal iVendedor As Integer, ByVal iIncluirPrevisao As Integer) As Long

Dim lErro As Long, objCondicaoPagto As New ClassCondicaoPagto, dDias As Double, dCustoFinanceiro As Double
Dim lTransacao As Long, alComando(0 To 1) As Long, iIndice As Integer, sSelect As String
Dim lNumPedidoVenda As Long, lNumTitulo As Long, iNumParcela As Integer, dValorParcela As Double, lCodCliente As Long
Dim sClienteNome As String, iCondicaoPagto As Integer, dICMSValor As Double, dIPIValor As Double, dtDataVencimento As Date, lNumIntComissao As Long
Dim iCodVendedor As Integer, iStatusComissao As Integer, dtDataGeracaoComis As Date, dtDataBaixaComis As Date, dPercentualComis As Double
Dim dValorBaseComis As Double, dValorComissao As Double, sNomeReduzidoVend As String, dtEmissaoTitulo As Date

On Error GoTo Erro_RelComisCro_Prepara

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130497

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 130498
    Next

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_NUMINTREL", lNumIntRel)
    If lErro <> SUCESSO Then gError 130499
    
    sClienteNome = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    sNomeReduzidoVend = String(STRING_VENDEDOR_NOME, 0)
    
    sSelect = "SELECT EmissaoTitulo, NumPedidoVenda, NumTitulo, NumParcela, ValorParcela, CodCliente, " & _
        "ClienteNome, CondicaoPagto, ICMSValor, IPIValor, DataVencimento, NumIntComissao," & _
        "CodVendedor, StatusComissao, DataGeracaoComis, DataBaixaComis, PercentualComis," & _
        "ValorBaseComis , ValorComissao, NomeReduzidoVend FROM RelComissaoCroAux WHERE FilialEmpresa = ? "
        
    If iIncluirPrevisao Then
        sSelect = sSelect & " AND ((DataGeracaoComis = ? AND DataVencimento BETWEEN ? AND ?) OR (DataGeracaoComis between ? AND ?))"
        If iVendedor = 0 Then
            lErro = Comando_Executar(alComando(0), sSelect, dtEmissaoTitulo, lNumPedidoVenda, lNumTitulo, iNumParcela, dValorParcela, lCodCliente, _
                sClienteNome, iCondicaoPagto, dICMSValor, dIPIValor, dtDataVencimento, lNumIntComissao, _
                iCodVendedor, iStatusComissao, dtDataGeracaoComis, dtDataBaixaComis, dPercentualComis, _
                dValorBaseComis, dValorComissao, sNomeReduzidoVend, iFilialEmpresa, DATA_NULA, dtDataInicio, dtDataFim, dtDataInicio, dtDataFim)
        Else
            lErro = Comando_Executar(alComando(0), sSelect & " AND CodVendedor = ?", dtEmissaoTitulo, lNumPedidoVenda, lNumTitulo, iNumParcela, dValorParcela, lCodCliente, _
                sClienteNome, iCondicaoPagto, dICMSValor, dIPIValor, dtDataVencimento, lNumIntComissao, _
                iCodVendedor, iStatusComissao, dtDataGeracaoComis, dtDataBaixaComis, dPercentualComis, _
                dValorBaseComis, dValorComissao, sNomeReduzidoVend, iFilialEmpresa, DATA_NULA, dtDataInicio, dtDataFim, dtDataInicio, dtDataFim, iVendedor)
        End If
    Else
        sSelect = sSelect & " AND DataGeracaoComis between ? AND ?"
        If iVendedor = 0 Then
            lErro = Comando_Executar(alComando(0), sSelect, dtEmissaoTitulo, lNumPedidoVenda, lNumTitulo, iNumParcela, dValorParcela, lCodCliente, _
                sClienteNome, iCondicaoPagto, dICMSValor, dIPIValor, dtDataVencimento, lNumIntComissao, _
                iCodVendedor, iStatusComissao, dtDataGeracaoComis, dtDataBaixaComis, dPercentualComis, _
                dValorBaseComis, dValorComissao, sNomeReduzidoVend, iFilialEmpresa, dtDataInicio, dtDataFim)
        Else
            lErro = Comando_Executar(alComando(0), sSelect & " AND CodVendedor = ?", dtEmissaoTitulo, lNumPedidoVenda, lNumTitulo, iNumParcela, dValorParcela, lCodCliente, _
                sClienteNome, iCondicaoPagto, dICMSValor, dIPIValor, dtDataVencimento, lNumIntComissao, _
                iCodVendedor, iStatusComissao, dtDataGeracaoComis, dtDataBaixaComis, dPercentualComis, _
                dValorBaseComis, dValorComissao, sNomeReduzidoVend, iFilialEmpresa, dtDataInicio, dtDataFim, iVendedor)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 130500
        
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130501
    
    Do While lErro = AD_SQL_SUCESSO
    
        '??? poderia otimizar se guardasse em colecao em memoria
        Set objCondicaoPagto = New ClassCondicaoPagto
        objCondicaoPagto.iCodigo = iCondicaoPagto
        
        'le os dados da condicao a partir do codigo...
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 130502
        
        lErro = CF("CalculaDias_CondPagamento", objCondicaoPagto, dDias)
        If lErro <> SUCESSO Then gError 130502
        
        dCustoFinanceiro = Arredonda_Moeda(dValorParcela - (dIPIValor + dICMSValor + dValorBaseComis), 2)
        
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelComisCro (EmissaoTitulo, NumPedidoVenda, NumTitulo, NumParcela, ValorParcela, CodCliente, " & _
            "ClienteNome, CondicaoPagto, ICMSValor, IPIValor, DataVencimento, NumIntComissao," & _
            "CodVendedor, StatusComissao, DataGeracaoComis, DataBaixaComis, PercentualComis," & _
            "ValorBaseComis , ValorComissao, NomeReduzidoVend, FilialEmpresa, NumIntRel, Prazo, CustoFinanceiro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            dtEmissaoTitulo, lNumPedidoVenda, lNumTitulo, iNumParcela, dValorParcela, lCodCliente, _
            sClienteNome, iCondicaoPagto, dICMSValor, dIPIValor, dtDataVencimento, lNumIntComissao, _
            iCodVendedor, iStatusComissao, dtDataGeracaoComis, dtDataBaixaComis, dPercentualComis, _
            dValorBaseComis, dValorComissao, sNomeReduzidoVend, iFilialEmpresa, lNumIntRel, dDias, dCustoFinanceiro)
        If lErro <> AD_SQL_SUCESSO Then gError 130503

        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130504
    
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130505
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
    
    RelComisCro_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelComisCro_Prepara:

    RelComisCro_Prepara = gErr
     
    Select Case gErr
          
        Case 130499, 130502
        
        Case 130500, 130501, 130504
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELCOMISSAOCROAUX", gErr)
        
        Case 130503
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELCOMISCRO", gErr)
        
        Case 130498
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
 
        Case 130497
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 130505
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 130506)
     
    End Select
     
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
         Call Comando_Fechar(alComando(iIndice))
    Next
   
    Exit Function

End Function


