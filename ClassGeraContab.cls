VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeraContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Gera_Lancamentos(ByVal sModulo As String, ByVal objTransacao As Object, ByVal sModelo As String, ByVal colLancamentos As ClassColLancamentos, ByVal iSubTipoTransacao As Integer) As Long
'retorna os lançamentos normais (sem incluir os lançamentos de custo) do Modulo/Transacao passados como parametro . Os lançamentos são retornados em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objPadraoContab As New ClassPadraoContab
Dim colPadraoContabItem As New ClassColPadraoContabItem
Dim colMnemonico As New Collection
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objTransacaoCTB As New ClassTransacaoCTB

On Error GoTo Erro_Gera_Lancamentos

    objPadraoContab.sModulo = sModulo
    objPadraoContab.sModelo = sModelo

    objTransacaoCTB.sSigla = sModulo
    objTransacaoCTB.sTransacao = objTransacao.Name
    objTransacaoCTB.iSubTipo = iSubTipoTransacao
    
    lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
    If lErro <> SUCESSO Then Error 36473

    objPadraoContab.iTransacao = objTransacaoCTB.iCodigo
    objPadraoContab.iSubTipo = objTransacaoCTB.iSubTipo

    'le o lancamento padrão para o modulo/transação em questão
    lErro = CF("PadraoContab_Le1", objPadraoContab)
    If lErro <> SUCESSO And lErro <> 36160 Then Error 36054

    'se não encontrou o modelo em questão
    If lErro = 36160 Then Error 36231

    'le os lancamentos padrão
    lErro = CF("PadraoContabItem_Le", objPadraoContab, colPadraoContabItem)
    If lErro <> SUCESSO Then Error 20462
    
    'le os mnemonicos associados ao modulo/transacao
    lErro = CF("MnemonicoCTB_Le", sModulo, objTransacaoCTB.iCodigo, colMnemonico)
    If lErro <> SUCESSO Then Error 36020
    
    Set objExeExp.colMnemonico = colMnemonico
    
    'descobre os grids da tela em questão e coloca-os em colGridTransacao
    lErro = Transacao_Grids(objTransacao, colGridTransacao)
    If lErro <> SUCESSO Then Error 36060
    
    Set objExeExp.colGridTransacao = colGridTransacao
    
    Set objExeExp.colLancamentos = colLancamentos
    
    Set objExeExp.objTransacao = objTransacao
    
    Set objExeExp.colMnemonicoValor = colMnemonicoValor
    
    'processa lancamentos padrão (transforma os lancamentos padrão em lançamentos contábeis)
    lErro = Processa_Lancamentos_Padrão(colPadraoContabItem, objExeExp)
    If lErro <> SUCESSO Then Error 20467
    
    Call Comando_Fechar(lComando)
    
    Gera_Lancamentos = SUCESSO

    Exit Function

Erro_Gera_Lancamentos:

    Gera_Lancamentos = Err
    
    Select Case Err
    
        Case 20462, 20467, 36020, 36054, 36060, 36473
        
        Case 36231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOCONTAB_INEXISTENTE", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150541)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function Gera_Lancamentos_Custo(ByVal sModulo As String, ByVal objTransacao As Object, ByVal sModelo As String, ByVal colLancamentos As ClassColLancamentos, ByVal iSubTipoTransacao As Integer) As Long
'retorna somente os lançamentos de custo do Modulo/Transacao passados como parametro. Os lançamentos são retornados em colLancamentos

Dim lComando As Long
Dim lErro As Long
Dim objPadraoContab As New ClassPadraoContab
Dim colPadraoContabItem As New ClassColPadraoContabItem
Dim colMnemonico As New Collection
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objTransacaoCTB As New ClassTransacaoCTB
Dim objMnemonico As ClassMnemonico
Dim objPadraoContabItem As ClassPadraoContabItem
Dim iEscaninhoCustoConsig As Integer
Dim iEscaninhoCustoConsig_Cadastrado As Integer
Dim iEscaninhoCustoBenef As Integer
Dim iEscaninhoCustoBenef_Cadastrado As Integer

On Error GoTo Erro_Gera_Lancamentos_Custo

    objPadraoContab.sModulo = sModulo
    objPadraoContab.sModelo = sModelo

    objTransacaoCTB.sSigla = sModulo
    objTransacaoCTB.sTransacao = objTransacao.Name
    objTransacaoCTB.iSubTipo = iSubTipoTransacao
    
    lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
    If lErro <> SUCESSO Then Error 36474

    objPadraoContab.iTransacao = objTransacaoCTB.iCodigo
    objPadraoContab.iSubTipo = objTransacaoCTB.iSubTipo

    'le o lancamento padrão para o modulo/transação em questão
    lErro = CF("PadraoContab_Le1", objPadraoContab)
    If lErro <> SUCESSO And lErro <> 36160 Then Error 36054

    'se não encontrou o modelo em questão
    If lErro = 36160 Then Error 36231

    'le os lancamentos padrão
    lErro = CF("PadraoContabItem_Le_Custo", objPadraoContab, colPadraoContabItem)
    If lErro <> SUCESSO Then Error 20462
    
    If colPadraoContabItem.Count = 0 Then Error 36774
    
    'le os mnemonicos associados ao modulo/transacao
    lErro = CF("MnemonicoCTB_Le", sModulo, objTransacaoCTB.iCodigo, colMnemonico)
    If lErro <> SUCESSO Then Error 36020
        
    Set objExeExp.colMnemonico = colMnemonico
    
    'se a lista de mnemonicos da transação incluir o mnemonico ESCANINHO_CUSTO ==> colocar este na coluna Escaninho_Custo. Esta coluna é semelhante as demais e será avaliada como as demais. Só não estará disponivel para ser setada.
    For Each objMnemonico In colMnemonico
    
        If objMnemonico.sMnemonico = ESCANINHO_CUSTO Then
            
            For Each objPadraoContabItem In colPadraoContabItem
                If objPadraoContabItem.sDebito = QUANT_ALOCADA_CONSIG Or objPadraoContabItem.sCredito = QUANT_ALOCADA_CONSIG Then
                    objPadraoContabItem.sEscaninho_Custo = ESCANINHO_CUSTO_CONSIG
                    iEscaninhoCustoConsig = 1
                ElseIf objPadraoContabItem.sDebito = QUANT_BENEFICIADA Or objPadraoContabItem.sCredito = QUANT_BENEFICIADA Then
                    objPadraoContabItem.sEscaninho_Custo = ESCANINHO_CUSTO_BENEF
                    iEscaninhoCustoBenef = 1
                Else
                    objPadraoContabItem.sEscaninho_Custo = ESCANINHO_CUSTO
                End If
            Next
            
            Exit For
            
        End If
        
    Next
    
    'se utilizou o mnemonico escaninhocustoconsig ===> verifica se ele está cadastrado na tabela de Mnemonicos
    If iEscaninhoCustoConsig = 1 Then
    
        For Each objMnemonico In colMnemonico
    
            If objMnemonico.sMnemonico = ESCANINHO_CUSTO_CONSIG Then
                iEscaninhoCustoConsig_Cadastrado = 1
                Exit For
            End If
            
        Next
        
        'se o escaninhocustoconsig não estiver cadastrado ==> erro
        If iEscaninhoCustoConsig_Cadastrado <> 1 Then Error 20796
    
    'se utilizou o mnemonico escaninhocustobenef ===> verifica se ele está cadastrado na tabela de Mnemonicos
    ElseIf iEscaninhoCustoBenef = 1 Then
    
        For Each objMnemonico In colMnemonico
    
            If objMnemonico.sMnemonico = ESCANINHO_CUSTO_BENEF Then
                iEscaninhoCustoBenef_Cadastrado = 1
                Exit For
            End If
            
        Next
        
        'se o escaninhocustobenef não estiver cadastrado ==> erro
        If iEscaninhoCustoBenef_Cadastrado <> 1 Then Error 20797
        
    End If
    'descobre os grids da tela em questão e coloca-os em colGridTransacao
    lErro = Transacao_Grids(objTransacao, colGridTransacao)
    If lErro <> SUCESSO Then Error 36060
    
    Set objExeExp.colGridTransacao = colGridTransacao
    
    Set objExeExp.colLancamentos = colLancamentos
    
    Set objExeExp.objTransacao = objTransacao
    
    Set objExeExp.colMnemonicoValor = colMnemonicoValor
    
    'processa lancamentos padrão (transforma os lancamentos padrão em lançamentos contábeis)
    lErro = Processa_Lancamentos_Padrão(colPadraoContabItem, objExeExp)
    If lErro <> SUCESSO Then Error 20467
    
    Call Comando_Fechar(lComando)
    
    Gera_Lancamentos_Custo = SUCESSO

    Exit Function

Erro_Gera_Lancamentos_Custo:

    Gera_Lancamentos_Custo = Err
    
    Select Case Err
    
        Case 20462, 20467, 36020, 36054, 36060, 36474, 36774
        
        Case 20796
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_ESCANINHOCUSTOCONSIG_INEX", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao)
        
        Case 20797
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_ESCANINHOCUSTOBENEF_INEX", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao)
        
        Case 36231
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PADRAOCONTAB_INEXISTENTE", Err, objPadraoContab.sModulo, objPadraoContab.iTransacao, objPadraoContab.sModelo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150542)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function Processa_Lancamentos_Padrão(ByVal colPadraoContabItem As ClassColPadraoContabItem, ByVal objExeExp As ClassExeExp) As Long
'Transforma os lançamentos padrão (colPadraoContabItem) em lançamentos contábeis e retorna-os em colLancamentos dentro de objExeExp. objTransacao dentro de objExeExp indica o objeto onde residem as funcoes que retornam os valores dos mnemonicos

Dim lErro As Long
Dim objPadraoContabItem As ClassPadraoContabItem
Dim objLancamento As ClassLancamentos
Dim iTipo As Integer
Dim lExpComp As Long
Dim iUltimoLancamentoAntigo As Integer
Dim iUltimoLancamentoAtual As Integer
Dim iIndice As Integer

On Error GoTo Erro_Processa_Lancamentos_Padrão

    For Each objPadraoContabItem In colPadraoContabItem

        'inicializa o local que armazenará as formulas compiladas
        lErro = Inicializa_Formula_Contabil(lExpComp)
        If lErro <> SUCESSO Then gError 36064

        'se foi especificada uma formula para a conta ==> compila-a
        If Len(objPadraoContabItem.sConta) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sConta, CAMPO_CONTA, lExpComp)
            If lErro <> SUCESSO Then gError 36000
            
        End If
        
        'se foi especificada uma formula para o centro de custo/lucro ==> compila-a
        If Len(objPadraoContabItem.sCcl) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sCcl, CAMPO_CCL, lExpComp)
            If lErro <> SUCESSO Then gError 36038
            
        End If
        
        'se foi especificada uma formula para o crédito ==> compila-a
        If Len(objPadraoContabItem.sCredito) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sCredito, CAMPO_CREDITO, lExpComp)
            If lErro <> SUCESSO Then gError 36039
            
        End If
        
        'se foi especificada uma formula para o debito ==> compila-a
        If Len(objPadraoContabItem.sDebito) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sDebito, CAMPO_DEBITO, lExpComp)
            If lErro <> SUCESSO Then gError 36040
            
        End If
        
        'se foi especificada uma formula para o histórico ==> compila-a
        If Len(objPadraoContabItem.sHistorico) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sHistorico, CAMPO_HISTORICO, lExpComp)
            If lErro <> SUCESSO Then gError 36041
            
        End If
        
        'se foi especificada uma formula para o produto (utilizado em lançamento de custo) ==> compila-a
        If Len(objPadraoContabItem.sProduto) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sProduto, CAMPO_PRODUTO, lExpComp)
            If lErro <> SUCESSO Then gError 36409
            
        End If
        
        
        lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.iAglutina, CAMPO_AGLUTINA, lExpComp)
        If lErro <> SUCESSO Then gError 55860
        
        lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.iGerencial, CAMPO_GERENCIAL, lExpComp)
        If lErro <> SUCESSO Then gError 188067
        
        'se foi especificada uma formula para o escaninho de custo ==> compila-a
        If Len(objPadraoContabItem.sEscaninho_Custo) > 0 Then

            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sEscaninho_Custo, CAMPO_ESCANINHO_CUSTO, lExpComp)
            If lErro <> SUCESSO Then gError 20792
            
        End If
        
        
        iUltimoLancamentoAntigo = objExeExp.colLancamentos.Count
        
        'executa as formulas de um lançamento padrão
        lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_Campo, AddressOf MaqExp_Devolve_Valor_Mnemonico, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
        If lErro <> SUCESSO Then gError 36065
        
        iUltimoLancamentoAtual = objExeExp.colLancamentos.Count
        
        If iUltimoLancamentoAtual > iUltimoLancamentoAntigo Then
        
            iIndice = iUltimoLancamentoAntigo + 1
        
            Do While iIndice <= iUltimoLancamentoAtual
        
                'seta o flag de aglutinação
                objExeExp.colLancamentos.Item(iIndice).iAglutina = objPadraoContabItem.iAglutina
                
                Set objLancamento = objExeExp.colLancamentos.Item(iIndice)
                
                'se o lançamento não estiver preenchido ==> remova-o
                If objLancamento.sConta = "" And objLancamento.sCcl = "" And objLancamento.dCredito = 0 And objLancamento.dDebito = 0 And objLancamento.sHistorico = "" And objLancamento.sProduto = "" Then
                    objExeExp.colLancamentos.Remove (iIndice)
                    iUltimoLancamentoAtual = iUltimoLancamentoAtual - 1
                    iIndice = iIndice - 1
                End If
            
                iIndice = iIndice + 1
                
            Loop
            
        End If
        
        'libera o espaço alocado para guardar as formulas compiladas
        lErro = Finaliza_Formula_Contabil(lExpComp)
        If lErro <> SUCESSO Then gError 36066
        
    Next

    Processa_Lancamentos_Padrão = SUCESSO
    
    Exit Function
    
Erro_Processa_Lancamentos_Padrão:

    Processa_Lancamentos_Padrão = gErr
    
    Select Case gErr
    
        Case 20792, 36000, 36038, 36039, 36040, 36041, 36064, 36065, 36066, 36409, 55860, 188067
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150543)
        
    End Select
    
    Exit Function

End Function

Function Valida_Lancamentos_Padrao(ByVal sModulo As String, ByVal objTransacao As Object) As Long
'valida os lançamentos do Modulo/Transacao passados como parametro.

Dim lErro As Long
Dim objPadraoContab As New ClassPadraoContab
Dim colPadraoContabItem As New ClassColPadraoContabItem
Dim colMnemonico As New Collection
Dim objExeExp As New ClassExeExp
Dim objTransacaoCTB As New ClassTransacaoCTB

On Error GoTo Erro_Valida_Lancamentos_Padrao

    objPadraoContab.sModulo = sModulo

    objTransacaoCTB.sSigla = sModulo
    objTransacaoCTB.sTransacao = objTransacao.Name
    
    lErro = CF("TransacaoCTB_Le", objTransacaoCTB)
    If lErro <> SUCESSO Then Error 36474

    objPadraoContab.iTransacao = objTransacaoCTB.iCodigo

    'le o lancamentos padrão para o modulo/transação em questão
    lErro = CF("PadraoContab_Le", objPadraoContab)
    If lErro <> SUCESSO Then Error 36042
    
    'le os lancamentos padrão
    lErro = CF("PadraoContabItem_Le", objPadraoContab, colPadraoContabItem)
    If lErro <> SUCESSO Then Error 36046
    
    'le os mnemonicos associados ao modulo/transacao
    lErro = CF("MnemonicoCTB_Le", sModulo, objTransacaoCTB.iCodigo, colMnemonico)
    If lErro <> SUCESSO Then Error 36047
    
    Set objExeExp.colMnemonico = colMnemonico
    
    'compila lancamentos padrão
    lErro = Compila_Lancamentos_Padrão(colPadraoContabItem, objExeExp)
    If lErro <> SUCESSO Then Error 36048
    
    Valida_Lancamentos_Padrao = SUCESSO

    Exit Function

Erro_Valida_Lancamentos_Padrao:

    Valida_Lancamentos_Padrao = Err
    
    Select Case Err
    
        Case 36042, 36046, 36047, 36048
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150544)
        
    End Select
    
    Exit Function
    
End Function

Private Function Compila_Lancamentos_Padrão(ByVal colPadraoContabItem As ClassColPadraoContabItem, ByVal objExeExp As ClassExeExp) As Long
'Compila os lançamentos padrão (colPadraoContabItem)

Dim lErro As Long
Dim objPadraoContabItem As ClassPadraoContabItem
Dim iTipo As Integer
Dim iInicio As Integer
Dim iTam As Integer
Dim sFormula As String
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Compila_Lancamentos_Padrão

    For Each objPadraoContabItem In colPadraoContabItem

        If Len(objPadraoContabItem.sConta) > 0 Then
            
            lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sConta, iTipo, iInicio, iTam)
            If lErro <> SUCESSO Then Error 36049
            
            If iTipo <> TIPO_TEXTO Then Error 36131
            
        End If
        
        If Len(objPadraoContabItem.sCcl) > 0 Then

            sFormula = objPadraoContabItem.sCcl
            
            lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sCcl, iTipo, iInicio, iTam)
            If lErro <> SUCESSO Then Error 36050
            
            If iTipo <> TIPO_TEXTO Then Error 36132

        End If
        
        If Len(objPadraoContabItem.sCredito) > 0 Then

            lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sCredito, iTipo, iInicio, iTam)
            If lErro <> SUCESSO Then Error 36051
            
            If iTipo <> TIPO_NUMERICO Then Error 36133

        End If
        
        If Len(objPadraoContabItem.sDebito) > 0 Then


            lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sDebito, iTipo, iInicio, iTam)
            If lErro <> SUCESSO Then Error 36052
            
            If iTipo <> TIPO_NUMERICO Then Error 36134
            
        End If
        
        If Len(objPadraoContabItem.sHistorico) > 0 Then

            lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objPadraoContabItem.sHistorico, iTipo, iInicio, iTam)
            If lErro <> SUCESSO Then Error 36053
            
            If iTipo <> TIPO_TEXTO Then Error 36135
            
        End If
        
    Next

    Compila_Lancamentos_Padrão = SUCESSO
    
    Exit Function
    
Erro_Compila_Lancamentos_Padrão:

    Compila_Lancamentos_Padrão = Err
    
    Select Case Err
    
        Case 36049, 36050, 36051, 36052, 36053
        
        Case 36131
            Call Retorna_Tipo(TIPO_TEXTO, sTipo)
            Call Retorna_Tipo(iTipo, sTipo1)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, objPadraoContabItem.sConta, sTipo1, sTipo)
            
        Case 36132
            Call Retorna_Tipo(TIPO_TEXTO, sTipo)
            Call Retorna_Tipo(iTipo, sTipo1)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, objPadraoContabItem.sCcl, sTipo1, sTipo)
            
        Case 36133
            Call Retorna_Tipo(TIPO_NUMERICO, sTipo)
            Call Retorna_Tipo(iTipo, sTipo1)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, objPadraoContabItem.sCredito, sTipo1, sTipo)
        
        Case 36134
            Call Retorna_Tipo(TIPO_NUMERICO, sTipo)
            Call Retorna_Tipo(iTipo, sTipo1)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, objPadraoContabItem.sDebito, sTipo1, sTipo)
        
        Case 36135
            Call Retorna_Tipo(TIPO_TEXTO, sTipo)
            Call Retorna_Tipo(iTipo, sTipo1)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, objPadraoContabItem.sHistorico, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150545)
        
    End Select
    
    Exit Function

End Function

Function Valida_Formula(ByVal sModulo As String, ByVal iTransacao As Integer, ByVal sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim colMnemonico As New Collection
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_Formula

    If Len(sFormula) > 0 Then

        'le os mnemonicos associados ao modulo/transacao
        lErro = CF("MnemonicoCTB_Le", sModulo, iTransacao, colMnemonico)
        If lErro <> SUCESSO Then Error 36047
        
        Set objExeExp.colMnemonico = colMnemonico
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then Error 36049
        
        If iTipo <> iTipoCompilado Then Error 36136
    
    End If
    
    Valida_Formula = SUCESSO

    Exit Function

Erro_Valida_Formula:

    Valida_Formula = Err
    
    Select Case Err
    
        Case 36047, 36049
        
        Case 36136
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", Err, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150546)
        
    End Select
    
    Exit Function
    
End Function

Public Function Transacao_Grids(ByVal objTransacao As Object, ByVal colGridTransacao As ClassColGridTransacao) As Long
'descobre os grids da tela em questão e coloca-os em colGridTransacao
'ATENCAO: se objTransacao nao for uma tela o erro 438 será tratado como sucesso

Dim objTela As Object
Dim objControle As Control
Dim lErro As Long
Dim iLinhasExistentes As Integer

On Error GoTo Erro_Transacao_Grids

    Set objTela = objTransacao

    'pesquisa todos os elementos da tela
    For Each objControle In objTela.Controls

        'se o elemento em questão é um grid ==> armanzena-o
        If TypeName(objControle) = "MSFlexGrid" Then
            If IsNumeric(objControle.Tag) Then
                iLinhasExistentes = CInt(objControle.Tag)
            Else
                iLinhasExistentes = 0
            End If
            colGridTransacao.Add objControle.Name, iLinhasExistentes, objControle.Name
        End If
        
    Next
        
    Transacao_Grids = SUCESSO
    
    Exit Function
    
Erro_Transacao_Grids:

    Transacao_Grids = Err
    
    Select Case Err
    
        Case 438 'para tratar caso em que objTransacao nao seja uma tela
            Transacao_Grids = SUCESSO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 150547)
        
    End Select
    
    Exit Function
    
End Function

Function Retorna_Tipo(ByVal iTipo As Integer, sTipo As String) As Long

    Select Case iTipo
    
        Case TIPO_NUMERICO
            sTipo = NUMERICO
            
        Case TIPO_DATA
            sTipo = Data
            
        Case TIPO_HORA
            sTipo = HORA

        Case TIPO_TEXTO
            sTipo = TEXTO

        Case TIPO_BOOLEANO
            sTipo = BOOLEANO

        Case Else
            sTipo = INVALIDO
        
    End Select
    
    Retorna_Tipo = SUCESSO
    
End Function

Public Function Avalia_Expressao_FPreco1(ByVal colFormacaoPreco As Collection, dValor As Double, ByVal sProduto As String, Optional ByVal objContexto As Object) As Long
'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colFormacaoPreco.Item(?).dValor

Dim lErro As Long
Dim objFormacaoPreco As ClassFormacaoPreco
Dim iTipo As Integer
Dim lExpComp As Long
Dim iIndice As Integer
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objMnemonicoValor As ClassMnemonicoValor
Dim objMnemonicoFPreco As New ClassMnemonicoFPreco
Dim colMnemonico As New Collection

On Error GoTo Erro_Avalia_Expressao_FPreco1

    If colFormacaoPreco.Count > 0 Then
    
        Set objFormacaoPreco = colFormacaoPreco.Item(1)

        objMnemonicoFPreco.iFilialEmpresa = objFormacaoPreco.iFilialEmpresa
        objMnemonicoFPreco.iEscopo = objFormacaoPreco.iEscopo
        objMnemonicoFPreco.sItemCategoria = objFormacaoPreco.sItemCategoria
        objMnemonicoFPreco.iTabelaPreco = objFormacaoPreco.iTabelaPreco
        objMnemonicoFPreco.sProduto = objFormacaoPreco.sProduto

        'leitura dos mnemonicos no BD
        lErro = CF("MnemonicoFPreco_Le_Todos1", objMnemonicoFPreco, colMnemonico)
        If lErro <> SUCESSO Then gError 92284
    
        Set objExeExp.colMnemonico = colMnemonico
        
        Set objExeExp.colGridTransacao = colGridTransacao
        
        Set objExeExp.objTransacao = CreateObject("RotinasEST.ClassESTSelect")
        
        Set objExeExp.objContexto = objContexto
        
        objExeExp.sProduto = sProduto 'produto para o qual está sendo executada a planilha
        
        Set objExeExp.colMnemonicoValor = colMnemonicoValor
        
        For Each objFormacaoPreco In colFormacaoPreco
    
            'inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92280
    
            objExeExp.iLinhaAtual = objFormacaoPreco.iLinha
    
            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco, objFormacaoPreco.sExpressao, 1, lExpComp)
            If lErro <> SUCESSO Then gError 92281
                
            'executa as formulas de um lançamento padrão
            lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_Campo_FPreco, AddressOf MaqExp_Devolve_Valor_Mnemon_FPreco, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
            If lErro <> SUCESSO Then gError 92282
            
            Set objMnemonicoValor = New ClassMnemonicoValor
            Set objMnemonicoValor.colValor = New Collection
            
            objMnemonicoValor.sMnemonico = "L" & CStr(objFormacaoPreco.iLinha)
            objMnemonicoValor.colValor.Add objExeExp.vValor
            
            objExeExp.colMnemonicoValor.Add objMnemonicoValor
            
            objFormacaoPreco.dValor = CDbl(objExeExp.vValor)
            
            'libera o espaço alocado para guardar as formulas compiladas
            lErro = Finaliza_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92283
            
        Next
    
        dValor = CDbl(objExeExp.vValor)

    Else
    
        dValor = 0
        
    End If

    Avalia_Expressao_FPreco1 = SUCESSO
    
    Exit Function
    
Erro_Avalia_Expressao_FPreco1:

    Avalia_Expressao_FPreco1 = gErr
    
    Select Case gErr
    
        Case 92281, 92282
            Call Rotina_Erro(vbOKOnly, "erro na expressao " & objFormacaoPreco.sExpressao, gErr)
            
        Case 92280, 92283, 92284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150548)
        
    End Select
    
    Exit Function

End Function

Function Valida_FormulaFPreco(ByVal sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPreco

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinha
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92181
        
        If iTipo <> iTipoCompilado Then gError 92182
    
    End If
    
    Valida_FormulaFPreco = SUCESSO

    Exit Function

Erro_Valida_FormulaFPreco:

    Valida_FormulaFPreco = gErr
    
    Select Case gErr
    
        Case 92180, 92181
        
        Case 92182
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150549)
        
    End Select
    
    Exit Function
    
End Function

Function Valida_FormulaFPreco1(sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinhaAtual As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPreco1

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinhaAtual
        objExeExp.iLinha = iLinha
        objExeExp.sExpressao = sFormula
        objExeExp.sExpressao1 = sFormula
        objExeExp.iInicio_Expressao = -1
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco1, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92288
        
        If iTipo <> iTipoCompilado Then gError 92289
    
        sFormula = objExeExp.sExpressao1
    
    End If
    
    Valida_FormulaFPreco1 = SUCESSO

    Exit Function

Erro_Valida_FormulaFPreco1:

    Valida_FormulaFPreco1 = gErr
    
    Select Case gErr
    
        Case 92287, 92288
        
        Case 92289
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150550)
        
    End Select
    
    Exit Function
    
End Function

Function Valida_FormulaFPreco2(sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinhaAtual As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPreco2

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinhaAtual
        objExeExp.iLinha = iLinha
        objExeExp.sExpressao = sFormula
        objExeExp.sExpressao1 = sFormula
        objExeExp.iInicio_Expressao = -1
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco2, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92297
        
        If iTipo <> iTipoCompilado Then gError 92298
    
        sFormula = objExeExp.sExpressao1
    
    End If
    
    Valida_FormulaFPreco2 = SUCESSO

    Exit Function

Erro_Valida_FormulaFPreco2:

    Valida_FormulaFPreco2 = gErr
    
    Select Case gErr
    
        Case 92296, 92297
        
        Case 92298
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150551)
        
    End Select
    
    Exit Function
    
End Function

Public Function Avalia_Expressao_FPreco2(ByVal colPlanilhas As Collection, dValor As Double, ByVal sProduto As String) As Long
'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor

Dim lErro As Long
Dim objPlanilhas As ClassPlanilhas
Dim iTipo As Integer
Dim lExpComp As Long
Dim iIndice As Integer
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objMnemonicoValor As ClassMnemonicoValor
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim colMnemonico As New Collection

On Error GoTo Erro_Avalia_Expressao_FPreco2

    If colPlanilhas.Count > 0 Then
    
        Set objPlanilhas = colPlanilhas.Item(1)

        objMnemonicoFPTipo.iFilialEmpresa = objPlanilhas.iFilialEmpresa
        objMnemonicoFPTipo.iEscopo = objPlanilhas.iEscopo
        objMnemonicoFPTipo.sItemCategoria = objPlanilhas.sItemCategoria
        objMnemonicoFPTipo.iTabelaPreco = objPlanilhas.iTabelaPreco
        objMnemonicoFPTipo.sProduto = objPlanilhas.sProduto
        objMnemonicoFPTipo.iTipoPlanilha = objPlanilhas.iTipoPlanilha

        'leitura dos mnemonicos no BD
        lErro = CF("MnemonicoFPlanilha_Le_Tipo", objMnemonicoFPTipo, colMnemonico)
        If lErro <> SUCESSO Then gError 92284
    
        Set objExeExp.colMnemonico = colMnemonico
        
        Set objExeExp.colGridTransacao = colGridTransacao
        
        Set objExeExp.objTransacao = CreateObject("RotinasEST.ClassESTSelect")
        
        objExeExp.sProduto = sProduto 'produto para o qual está sendo executada a planilha
        
        Set objExeExp.colMnemonicoValor = colMnemonicoValor
        
        For Each objPlanilhas In colPlanilhas
    
            'inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92280
    
            objExeExp.iLinhaAtual = objPlanilhas.iLinha
    
            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPlanilha, objPlanilhas.sExpressao, 1, lExpComp)
            If lErro <> SUCESSO Then gError 92281
                
            'executa as formulas de um lançamento padrão
            lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_Campo_FPreco, AddressOf MaqExp_Devolve_Valor_Mnemon_FPlanilha, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
            If lErro <> SUCESSO Then gError 92282
            
            Set objMnemonicoValor = New ClassMnemonicoValor
            Set objMnemonicoValor.colValor = New Collection
            
            objMnemonicoValor.sMnemonico = "L" & CStr(objPlanilhas.iLinha)
            objMnemonicoValor.colValor.Add objExeExp.vValor
            
            objExeExp.colMnemonicoValor.Add objMnemonicoValor
            
            objPlanilhas.dValor = CDbl(objExeExp.vValor)
            
            'libera o espaço alocado para guardar as formulas compiladas
            lErro = Finaliza_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92283
            
        Next
    
        dValor = CDbl(objExeExp.vValor)

    Else
    
        dValor = 0
        
    End If

    Avalia_Expressao_FPreco2 = SUCESSO
    
    Exit Function
    
Erro_Avalia_Expressao_FPreco2:

    Avalia_Expressao_FPreco2 = gErr
    
    Select Case gErr
    
        Case 92281, 92282
            Call Rotina_Erro(vbOKOnly, "erro na expressao " & objPlanilhas.sExpressao, gErr)
            
        Case 92280, 92283, 92284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150552)
        
    End Select
    
    Exit Function

End Function

Function Valida_FormulaFPlanilha(ByVal sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPlanilha

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinha
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPlanilha, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92181
        
        If iTipo <> iTipoCompilado Then gError 92182
    
    End If
    
    Valida_FormulaFPlanilha = SUCESSO

    Exit Function

Erro_Valida_FormulaFPlanilha:

    Valida_FormulaFPlanilha = gErr
    
    Select Case gErr
    
        Case 92180, 92181
        
        Case 92182
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150553)
        
    End Select
    
    Exit Function
    
End Function

Function Valida_FormulaFPlanilha1(sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinhaAtual As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPlanilha1

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinhaAtual
        objExeExp.iLinha = iLinha
        objExeExp.sExpressao = sFormula
        objExeExp.sExpressao1 = sFormula
        objExeExp.iInicio_Expressao = -1
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco1, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92288
        
        If iTipo <> iTipoCompilado Then gError 92289
    
        sFormula = objExeExp.sExpressao1
    
    End If
    
    Valida_FormulaFPlanilha1 = SUCESSO

    Exit Function

Erro_Valida_FormulaFPlanilha1:

    Valida_FormulaFPlanilha1 = gErr
    
    Select Case gErr
    
        Case 92287, 92288
        
        Case 92289
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150554)
        
    End Select
    
    Exit Function
    
End Function

Function Valida_FormulaFPlanilha2(sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal iLinhaAtual As Integer, ByVal iLinha As Integer, colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_FormulaFPlanilha2

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        objExeExp.iLinhaAtual = iLinhaAtual
        objExeExp.iLinha = iLinha
        objExeExp.sExpressao = sFormula
        objExeExp.sExpressao1 = sFormula
        objExeExp.iInicio_Expressao = -1
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPreco2, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 92297
        
        If iTipo <> iTipoCompilado Then gError 92298
    
        sFormula = objExeExp.sExpressao1
    
    End If
    
    Valida_FormulaFPlanilha2 = SUCESSO

    Exit Function

Erro_Valida_FormulaFPlanilha2:

    Valida_FormulaFPlanilha2 = gErr
    
    Select Case gErr
    
        Case 92296, 92297
        
        Case 92298
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150555)
        
    End Select
    
    Exit Function
    
End Function

Public Function Avalia_Expressao_FPreco3(ByVal colPlanilhas As Collection, dValor As Double, ByVal objContexto As Object) As Long
'Executa as formulas da planilha de preço. Retorna o valor da planilha em dValor (que é o valor da última linha da planilha) e o valor de cada linha em colPlanilhas.Item(?).dValor

Dim lErro As Long
Dim objPlanilhas As ClassPlanilhas
Dim iTipo As Integer
Dim lExpComp As Long
Dim iIndice As Integer
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objMnemonicoValor As ClassMnemonicoValor
Dim objMnemonicoFPTipo As New ClassMnemonicoFPTipo
Dim colMnemonico As New Collection

On Error GoTo Erro_Avalia_Expressao_FPreco3

    If colPlanilhas.Count > 0 Then
    
        Set objPlanilhas = colPlanilhas.Item(1)

        objMnemonicoFPTipo.iFilialEmpresa = objPlanilhas.iFilialEmpresa
        objMnemonicoFPTipo.iEscopo = objPlanilhas.iEscopo
        objMnemonicoFPTipo.sItemCategoria = objPlanilhas.sItemCategoria
        objMnemonicoFPTipo.iTabelaPreco = objPlanilhas.iTabelaPreco
        objMnemonicoFPTipo.sProduto = objContexto.sProduto
        objMnemonicoFPTipo.iTipoPlanilha = objPlanilhas.iTipoPlanilha

        'leitura dos mnemonicos no BD
        lErro = CF("MnemonicoFPlanilha_Le_Tipo", objMnemonicoFPTipo, colMnemonico)
        If lErro <> SUCESSO Then gError 92284
    
        Set objExeExp.colMnemonico = colMnemonico
        
        Set objExeExp.colGridTransacao = colGridTransacao
        
        Set objExeExp.objTransacao = CreateObject("RotinasEST.ClassESTSelect")
        
        Set objExeExp.objContexto = objContexto
        
        objExeExp.sProduto = objContexto.sProduto 'produto para o qual está sendo executada a planilha
        
        Set objExeExp.colMnemonicoValor = colMnemonicoValor
        
        If Not (objContexto.colMnemonicoValor) Is Nothing Then
        
            For Each objMnemonicoValor In objContexto.colMnemonicoValor
        
                colMnemonicoValor.Add objMnemonicoValor
            
            Next
            
        End If
        
        For Each objPlanilhas In colPlanilhas
    
            'inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92280
    
            objExeExp.iLinhaAtual = objPlanilhas.iLinha
    
            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico_FPlanilha, objPlanilhas.sExpressao, 1, lExpComp)
            If lErro <> SUCESSO Then gError 92281
                
            'executa as formulas de um lançamento padrão
            lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_Campo_FPreco, AddressOf MaqExp_Devolve_Valor_Mnemon_FPlanilha, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
            If lErro <> SUCESSO Then gError 92282
            
            Set objMnemonicoValor = New ClassMnemonicoValor
            Set objMnemonicoValor.colValor = New Collection
            
            objMnemonicoValor.sMnemonico = "L" & CStr(objPlanilhas.iLinha)
            objMnemonicoValor.colValor.Add objExeExp.vValor
            
            objExeExp.colMnemonicoValor.Add objMnemonicoValor
            
            objPlanilhas.dValor = CDbl(objExeExp.vValor)
            
            'libera o espaço alocado para guardar as formulas compiladas
            lErro = Finaliza_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 92283
            
        Next
    
        dValor = CDbl(objExeExp.vValor)

    Else
    
        dValor = 0
        
    End If

    Avalia_Expressao_FPreco3 = SUCESSO
    
    Exit Function
    
Erro_Avalia_Expressao_FPreco3:

    Avalia_Expressao_FPreco3 = gErr
    
    Select Case gErr
    
        Case 92281, 92282
            Call Rotina_Erro(vbOKOnly, "erro na expressao " & objPlanilhas.sExpressao, gErr)
            
        Case 92280, 92283, 92284
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150556)
        
    End Select
    
    Exit Function

End Function

Function Valida_Formula_WFW(ByVal sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_Formula_WFW

    If Len(sFormula) > 0 Then

        Set objExeExp.colMnemonico = colMnemonico
        
        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 178131
        
        If iTipo <> iTipoCompilado Then gError 178132
    
    End If
    
    Valida_Formula_WFW = SUCESSO

    Exit Function

Erro_Valida_Formula_WFW:

    Valida_Formula_WFW = gErr
    
    Select Case gErr
    
        Case 178131
        
        Case 178132
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178133)
        
    End Select
    
    Exit Function
    
End Function

Private Function Executa_ExpressaoWFW1(ByVal colRegraWFW As Collection, ByVal objExeExp As ClassExeExp) As Long
'Executa as expressoes passadas como parametro

Dim lErro As Long
Dim lExpComp As Long
Dim iIndice As Integer
Dim objRegraWFW As ClassRegraWFW
Dim objRegraWFW1 As ClassRegraWFW

On Error GoTo Erro_Executa_ExpressaoWFW1

    For Each objRegraWFW In colRegraWFW
    
        If Len(objRegraWFW.sRegra) > 0 Then
    
            'se for a transacao "Geral" coloca objRegraWFW no contexto para poder ler a data/hora da ultima execucao verdadeira de cada regra
            If objRegraWFW.sModulo = "ADM" And objRegraWFW.iTransacao = 9999 Then
                Set objExeExp.objContexto = objRegraWFW
            End If
    
            'inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 178149
    
            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sRegra, CAMPO_REGRA, lExpComp)
            If lErro <> SUCESSO Then gError 178150

            'executa as formulas de um lançamento padrão
            lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_CampoWFW, AddressOf MaqExp_Devolve_Valor_Mnemonico, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
            If lErro <> SUCESSO Then gError 178151

            'libera o espaço alocado para guardar as formulas compiladas
            lErro = Finaliza_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError 178152

            Set objRegraWFW1 = objExeExp.colRegras.Item(objExeExp.colRegras.Count)
            
            If objRegraWFW1.dRegraRet <> 0 Then
            
                'inicializa o local que armazenará as formulas compiladas
                lErro = Inicializa_Formula_Contabil(lExpComp)
                If lErro <> SUCESSO Then gError 178153
            
                If Len(objRegraWFW.sEmailPara) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sEmailPara, CAMPO_EMAILPARA, lExpComp)
                    If lErro <> SUCESSO Then gError 178260
    
                End If
            
                If Len(objRegraWFW.sEmailAssunto) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sEmailAssunto, CAMPO_EMAILASSUNTO, lExpComp)
                    If lErro <> SUCESSO Then gError 178154
    
                End If
                
                If Len(objRegraWFW.sEmailMsg) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sEmailMsg, CAMPO_EMAILMSG, lExpComp)
                    If lErro <> SUCESSO Then gError 178155
    
                End If
                
                If Len(objRegraWFW.sAvisoMsg) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sAvisoMsg, CAMPO_AVISOMSG, lExpComp)
                    If lErro <> SUCESSO Then gError 178156
    
                End If
                
                If Len(objRegraWFW.sLogDoc) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sLogDoc, CAMPO_LOGDOC, lExpComp)
                    If lErro <> SUCESSO Then gError 178157
    
                End If
    
                If Len(objRegraWFW.sLogMsg) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sLogMsg, CAMPO_LOGMSG, lExpComp)
                    If lErro <> SUCESSO Then gError 178158
    
                End If

                If Len(objRegraWFW.sRelSel) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sRelSel, CAMPO_RELSEL, lExpComp)
                    If lErro <> SUCESSO Then gError 178158
    
                End If

                If Len(objRegraWFW.sRelAnexo) > 0 Then
            
                    lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegraWFW.sRelAnexo, CAMPO_RELANEXO, lExpComp)
                    If lErro <> SUCESSO Then gError 178158
    
                End If

                'executa as formulas de um lançamento padrão
                lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_CampoWFW, AddressOf MaqExp_Devolve_Valor_Mnemonico, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
                If lErro <> SUCESSO Then gError 178160

                'libera o espaço alocado para guardar as formulas compiladas
                lErro = Finaliza_Formula_Contabil(lExpComp)
                If lErro <> SUCESSO Then gError 178159
        
            End If

        End If
        
    Next

    Executa_ExpressaoWFW1 = SUCESSO
    
    Exit Function
    
Erro_Executa_ExpressaoWFW1:

    Executa_ExpressaoWFW1 = gErr
    
    Select Case gErr
    
        Case 178149 To 178160, 178260
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150543)
        
    End Select
    
    Exit Function

End Function

Function Executa_ExpressaoWFW(ByVal sModulo As String, ByVal objTransacao As Object, ByVal colRegraWFW As Collection, ByVal colRegraWFW1 As Collection, ByVal objContexto As Object) As Long
'retorna as regras de workflow executadas em  colRegraWFW
 
Dim lErro As Long
Dim colMnemonico As New Collection
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor
Dim objRegraWFW As ClassRegraWFW
 
On Error GoTo Erro_Executa_ExpressaoWFW
 
    If colRegraWFW1.Count > 0 Then
         
        Set objRegraWFW = colRegraWFW1(1)
 
        'le os mnemonicos associados ao modulo/transacao
        lErro = CF("MnemonicoWFW_Le", sModulo, objRegraWFW.iTransacao, colMnemonico)
        If lErro <> SUCESSO Then gError 178162
     
        Set objExeExp.colMnemonico = colMnemonico
     
        'descobre os grids da tela em questão e coloca-os em colGridTransacao
        lErro = Transacao_Grids(objTransacao, colGridTransacao)
        If lErro <> SUCESSO Then gError 178163
     
        Set objExeExp.colGridTransacao = colGridTransacao
         
        Set objExeExp.colRegras = colRegraWFW
         
        Set objExeExp.objTransacao = objTransacao
         
        Set objExeExp.colMnemonicoValor = colMnemonicoValor
         
        Set objExeExp.objContexto = objContexto
         
        'processa lancamentos padrão (transforma os lancamentos padrão em lançamentos contábeis)
        lErro = Executa_ExpressaoWFW1(colRegraWFW1, objExeExp)
        If lErro <> SUCESSO Then gError 178164
     
    End If
     
    Executa_ExpressaoWFW = SUCESSO
 
    Exit Function
 
Erro_Executa_ExpressaoWFW:
 
    Executa_ExpressaoWFW = gErr
     
    Select Case gErr
     
        Case 178161 To 178164
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150541)
         
    End Select
     
    Exit Function
     
End Function

'Function Executa_ExpressaoBrowser(ByVal colRegraWFW As Collection, ByVal colRegraWFW1 As Collection, ByVal objContexto As Object) As Long
''retorna as regras de workflow executadas em  colRegraWFW
'
'Dim lErro As Long
'Dim colMnemonico As New Collection
'Dim objExeExp As New ClassExeExp
'Dim colGridTransacao As New ClassColGridTransacao
'Dim colMnemonicoValor As New ClassColMnemonicoValor
'Dim objRegraWFW As ClassRegraWFW
'
'On Error GoTo Erro_Executa_ExpressaoWFW
'
'
'        Set objExeExp.colRegras = colRegraWFW
'
'        Set objExeExp.objTransacao = objTransacao
'
'        Set objExeExp.colMnemonicoValor = colMnemonicoValor
'
'        Set objExeExp.objContexto = objContexto
'
'        'processa lancamentos padrão (transforma os lancamentos padrão em lançamentos contábeis)
'        lErro = Executa_ExpressaoWFW1(colRegraWFW1, objExeExp)
'        If lErro <> SUCESSO Then gError 178164
'
'    End If
'
'    Executa_ExpressaoWFW = SUCESSO
'
'    Exit Function
'
'Erro_Executa_ExpressaoWFW:
'
'    Executa_ExpressaoWFW = gErr
'
'    Select Case gErr
'
'        Case 178161 To 178164
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150541)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'
Function Valida_Formula_Browser(ByVal sFormula As String, ByVal iTipo As Integer, iInicio As Integer, iTam As Integer, ByVal colMnemonico As Collection) As Long
'compila a formula sFormula pertencente ao modulo sModulo e transacao iTransacao. iTipo indica o tipo que deve ser retornado pela formula. iInicio indica a posicao inicial da formula em que houve erro. iTam indica o tamanho do erro.

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim iTipoCompilado As Integer
Dim sTipo1 As String
Dim sTipo As String

On Error GoTo Erro_Valida_Formula_Browser

    If Len(sFormula) > 0 Then

        'compila a formula passada como parametro
        lErro = Compila_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sFormula, iTipoCompilado, iInicio, iTam)
        If lErro <> SUCESSO Then gError 178261
        
        If iTipo <> iTipoCompilado Then gError 178262
    
    End If
    
    Valida_Formula_Browser = SUCESSO

    Exit Function

Erro_Valida_Formula_Browser:

    Valida_Formula_Browser = gErr
    
    Select Case gErr
    
        Case 178261
        
        Case 178262
            Call Retorna_Tipo(iTipo, sTipo)
            Call Retorna_Tipo(iTipoCompilado, sTipo1)
            iInicio = 0
            iTam = Len(sFormula)
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORMULA_INVALIDA", gErr, sFormula, sTipo1, sTipo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 178263)
        
    End Select
    
    Exit Function
    
End Function

'#############################################################################
'Inserido por Wagner 01/03/2007
Function Executa_Regras_CobrEmail(ByVal objTransacao As Object, ByVal colPadraoContabItem As ClassColPadraoContabItem, ByVal colLancamentos As ClassColLancamentos, colMnemonicos As Collection) As Long

Dim lErro As Long
Dim objExeExp As New ClassExeExp
Dim colGridTransacao As New ClassColGridTransacao
Dim colMnemonicoValor As New ClassColMnemonicoValor

On Error GoTo Erro_Executa_Regras_CobrEmail
    
    Set objExeExp.colMnemonico = colMnemonicos
    
    Set objExeExp.colGridTransacao = colGridTransacao
    
    Set objExeExp.colLancamentos = colLancamentos
    
    Set objExeExp.objTransacao = objTransacao
    
    Set objExeExp.colMnemonicoValor = colMnemonicoValor
    
    'processa lancamentos padrão (transforma os lancamentos padrão em lançamentos contábeis)
    lErro = Processa_Lancamentos_Padrão(colPadraoContabItem, objExeExp)
    If lErro <> SUCESSO Then gError 187160
    
    Executa_Regras_CobrEmail = SUCESSO

    Exit Function

Erro_Executa_Regras_CobrEmail:

    Executa_Regras_CobrEmail = gErr
    
    Select Case gErr
    
        Case 187160
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187161)
        
    End Select
    
    Exit Function
    
End Function
'#############################################################################

Function Executa_ExpressaoRegrasMsg(ByVal iTipoDoc As Integer, ByVal objTransacao As Object, colMensagens As Collection, Optional ByVal bSoEspecifico As Boolean = False) As Long
 
Dim lErro As Long
Dim colMnemonicos As New Collection
Dim objExeExp As ClassExeExp
Dim colGridTransacao As ClassColGridTransacao
Dim colMnemonicoValor As ClassColMnemonicoValor
Dim colRegras As New Collection
Dim colRegrasRet As Collection
Dim iTipoMsg As Integer, colcolRegrasTipo As New Collection, colRegrasTipo As Collection
Dim objRegra As ClassRegrasMsg

On Error GoTo Erro_Executa_ExpressaoRegrasMsg
         
    'leitura dos mnemonicos no BD
    lErro = CF("MnemonicoRegrasMsg_Le", colMnemonicos, iTipoDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    'leitura das regras no BD
    lErro = CF("RegrasMsg_Le", colRegras, iTipoDoc, bSoEspecifico)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iTipoMsg = -1
    For Each objRegra In colRegras
        If iTipoMsg <> objRegra.iTipoMsg Then
            iTipoMsg = objRegra.iTipoMsg
            Set colRegrasTipo = New Collection
            colcolRegrasTipo.Add colRegrasTipo
        End If
        colRegrasTipo.Add objRegra
    Next
    
    For Each colRegrasTipo In colcolRegrasTipo
    
        Set objExeExp = New ClassExeExp
        Set colGridTransacao = New ClassColGridTransacao
        Set colMnemonicoValor = New ClassColMnemonicoValor
        Set colRegrasRet = New Collection

        Set objExeExp.colMnemonico = colMnemonicos
        Set objExeExp.colGridTransacao = colGridTransacao
        Set objExeExp.colRegras = colRegrasRet
        Set objExeExp.objTransacao = objTransacao
        Set objExeExp.colMnemonicoValor = colMnemonicoValor
        
        lErro = Executa_ExpressaoRegrasMsg1(colRegrasTipo, colMensagens, objExeExp)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    Next
     
    Executa_ExpressaoRegrasMsg = SUCESSO
 
    Exit Function
 
Erro_Executa_ExpressaoRegrasMsg:
 
    Executa_ExpressaoRegrasMsg = gErr
     
    Select Case gErr
     
        Case ERRO_SEM_MENSAGEM
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150541)
         
    End Select
     
    Exit Function
     
End Function

Private Function Executa_ExpressaoRegrasMsg1(ByVal colRegras As Collection, ByVal colMensagens As Collection, ByVal objExeExp As ClassExeExp) As Long
'Executa as expressoes passadas como parametro

Dim lErro As Long
Dim lExpComp As Long
Dim iIndice As Integer
Dim objRegra As ClassRegrasMsg
Dim objRegraRet As ClassRegrasMsg
Dim sRegra As String, bRegraValida As Boolean
Dim objMsg As ClassMensagensRegra

On Error GoTo Erro_Executa_ExpressaoRegrasMsg1

    Call objExeExp.objTransacao.Inicializa_Calculo_RegrasMsg2(colRegras)

    For Each objRegra In colRegras
    
        bRegraValida = True
        For iIndice = 1 To 5
        
            Select Case iIndice
                Case 1
                    sRegra = objRegra.sRegra1
                Case 2
                    sRegra = objRegra.sRegra2
                Case 3
                    sRegra = objRegra.sRegra3
                Case 4
                    sRegra = objRegra.sRegra4
                Case 5
                    sRegra = objRegra.sRegra5
            End Select
    
            If Len(sRegra) > 0 Then
        
                'inicializa o local que armazenará as formulas compiladas
                lErro = Inicializa_Formula_Contabil(lExpComp)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, sRegra, CAMPO_REGRA, lExpComp)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'executa as formulas de um lançamento padrão
                lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_CampoMsg, AddressOf MaqExp_Devolve_Valor_Mnemonico, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'libera o espaço alocado para guardar as formulas compiladas
                lErro = Finaliza_Formula_Contabil(lExpComp)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                Set objRegraRet = objExeExp.colRegras.Item(objExeExp.colRegras.Count)
                
                If objRegraRet.dRegraRet = 0 Then bRegraValida = False 'Regra não é verdadeira
                
            End If
                       
            If Not bRegraValida Then Exit For
            
        Next
        
        Set objMsg = New ClassMensagensRegra
        colMensagens.Add objMsg
        
        objMsg.iTipoDoc = objRegra.iTipoDoc
        objMsg.iTipoMsg = objRegra.iTipoMsg
        objMsg.iSeq = objRegra.iSeq
            
        Call objExeExp.objTransacao.Retorno_Calculo_RegrasMsg(bRegraValida)
            
        If bRegraValida Then
        
            'inicializa o local que armazenará as formulas compiladas
            lErro = Inicializa_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                   
            lErro = CompilaExe_Formula_Contabil(objExeExp, AddressOf MaqExp_Testa_Mnemonico, objRegra.sMensagem, CAMPO_LOGMSG, lExpComp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'executa as formulas de um lançamento padrão
            lErro = Executa_Formula_Contabil(objExeExp, AddressOf MaqExp_Grid_Linhas, lExpComp, AddressOf MaqExp_Armazena_CampoMsg, AddressOf MaqExp_Devolve_Valor_Mnemonico, AddressOf MaqExp_Devolve_Valor_Total_Mnemonico)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'libera o espaço alocado para guardar as formulas compiladas
            lErro = Finaliza_Formula_Contabil(lExpComp)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Set objRegraRet = objExeExp.colRegras.Item(objExeExp.colRegras.Count)
            
            objMsg.sMensagem = objRegraRet.sMensagemRet
            
        Else
        
            objMsg.sMensagem = ""
    
        End If
                
    Next
    
    Call objExeExp.objTransacao.Finaliza_Calculo_RegrasMsg

    Executa_ExpressaoRegrasMsg1 = SUCESSO
    
    Exit Function
    
Erro_Executa_ExpressaoRegrasMsg1:

    Executa_ExpressaoRegrasMsg1 = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            'MsgBox (sRegra)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150543)
        
    End Select
    
    Call objExeExp.objTransacao.Finaliza_Calculo_RegrasMsg
    
    Exit Function

End Function
