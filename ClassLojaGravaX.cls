VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLojaGrava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Venda_Gravar_CC(ByVal objVenda As ClassVenda) As Long
'grava os dados da venda no caixa central
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim lTransacao As Long
Dim colRegistro As New Collection
Dim objLog As New ClassLog
Dim objMovCx As ClassMovimentoCaixa
Dim objMovEstoque As New ClassMovEstoque
Dim lCodigo As Long
Dim alComando(1 To 3) As Long
Dim alComando1(1 To NUM_MAX_LCOMANDO_MOVESTOQUE) As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objCheque As ClassChequePre
Dim objTroca As ClassTroca
Dim colConfig As Object
Dim iAceitaEstoqueNegativo As Integer

On Error GoTo Erro_Venda_Gravar_CC

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126167
    Next

    For iIndice = LBound(alComando1) To UBound(alComando1)
        alComando1(iIndice) = Comando_Abrir()
        If alComando1(iIndice) = 0 Then gError 126168
    Next

    Do While iIndice1 < 2
        
        iIndice1 = iIndice1 + 1

        'gravar os dados do cupom
        lErro = CF("CupomFiscal_Gravar_Venda", objVenda.objCupomFiscal, objLog, objMovEstoque)
        If lErro <> SUCESSO Then gError 110015

        'gerar os NumMovto dos movimentos de caixa que serao utilizados nos cheques e trocas
        lErro = CF("MovimentosCaixa_Gera_NumMovto", objVenda.colMovimentosCaixa)
        If lErro <> SUCESSO Then gError 110016

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then Call CF("Mover_Separator_Log", objLog)

        'gravar os dados de carne caso existam
        lErro = CF("Carne_Gravar_Venda", objVenda.objCarne, objVenda.objCupomFiscal.lNumIntDoc, objLog)
        If lErro <> SUCESSO Then gError 110014
    
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then Call CF("Mover_Separator_Log", objLog)
    
        'gravar os dados de cheque caso existam
        lErro = CF("Cheque_Gravar_Venda", objVenda.colCheques, objVenda.colMovimentosCaixa, objLog)
        If lErro <> SUCESSO Then gError 110017
    
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then Call CF("Mover_Separator_Log", objLog)
    
        'gravar os dados de troca caso existam
        lErro = CF("Troca_Gravar_Venda", objVenda.colTroca, objVenda.colMovimentosCaixa, objLog, objVenda.objCupomFiscal, objMovEstoque)
        If lErro <> SUCESSO Then gError 110018
        
        For Each objMovCx In objVenda.colMovimentosCaixa
            objMovCx.lCupomFiscal = objVenda.objCupomFiscal.lNumIntDoc
        Next
        
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        
            Call CF("Mover_Separator_Log", objLog)
            
            'preenche o objLog com os dados da venda
            Call CF("Mover_Dados_Movcx_Log", objVenda.colMovimentosCaixa, objLog)
        
            objLog.iOperacao = INCLUSAO_CUPOM_FISCAL
        
            'grava o log
            lErro = CF("Log_Grava", objLog, objLog.colRegistro)
            If lErro <> SUCESSO Then gError 112093
            
        End If
        
        'gravar os movimentos de caixa
        lErro = CF("MovimentosCaixa_Gravar_Venda", objVenda.colMovimentosCaixa, objVenda.colCheques, objVenda.colTroca, objLog, objVenda.objCupomFiscal)
        If lErro <> SUCESSO Then gError 112094
    
        If objMovEstoque.colItens.Count > 0 Then
            
        
            Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
    
            colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, objMovEstoque.iFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
        
            'Lê as configurações em ESTConfig
            lErro = CF("ESTConfig_Le_Configs", colConfig)
            If lErro <> SUCESSO Then gError 126846
            
            
            gobjMAT.iAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
            
            
            'Obtém o Código do Movimento
            lErro = CF("MovEstoque_Automatico_EmTransacao1", alComando(1), alComando(2), alComando(3), objMovEstoque.iFilialEmpresa, lCodigo)
            If lErro <> SUCESSO Then gError 118195
            
            objMovEstoque.lCodigo = lCodigo
            
            'Grava os Movimentos de Estoque relacionados ao Cupom Fiscal
            lErro = CF("MovEstoque_Grava1", alComando1(), objMovEstoque)
            If lErro <> SUCESSO Then gError 118196
            
        
        End If

        If objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO Then Exit Do

        If objVenda.objCupomFiscal.iFilialEmpresa > FilialEmpresa_ConvFRFO(objVenda.objCupomFiscal.iFilialEmpresa) Then
        
            Set objMovEstoque = New ClassMovEstoque
        
            objVenda.objCupomFiscal.iFilialEmpresa = FilialEmpresa_ConvFRFO(objVenda.objCupomFiscal.iFilialEmpresa)
        
            For Each objMovCx In objVenda.colMovimentosCaixa
                objMovCx.iFilialEmpresa = objVenda.objCupomFiscal.iFilialEmpresa
            Next
            
            objVenda.objCarne.iFilialEmpresa = objVenda.objCupomFiscal.iFilialEmpresa
            
            For Each objCheque In objVenda.colCheques
                objCheque.iFilialEmpresa = objVenda.objCupomFiscal.iFilialEmpresa
                objCheque.iFilialEmpresaLoja = objVenda.objCupomFiscal.iFilialEmpresa
            Next
            
            For Each objTroca In objVenda.colTroca
                objTroca.iFilialEmpresa = objVenda.objCupomFiscal.iFilialEmpresa
            Next

        End If
        

    Loop

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Venda_Gravar_CC = SUCESSO

    Exit Function

Erro_Venda_Gravar_CC:

    Venda_Gravar_CC = gErr

    Select Case gErr

        Case 110014, 110015, 110016, 110017, 110018, 112093, 112094

        Case 126167, 126168
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 151081)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    For iIndice = LBound(alComando1) To UBound(alComando1)
        Call Comando_Fechar(alComando1(iIndice))
    Next

    Exit Function

End Function

