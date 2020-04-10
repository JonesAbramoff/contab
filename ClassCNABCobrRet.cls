VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCNABCobrRet"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Function Processar_ArquivoRetorno_Cobranca(ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long

Dim lErro As Long
Dim objCNABBanco As Object

On Error GoTo Erro_Processar_ArquivoRetorno_Cobranca

   
    Select Case objCobrancaEletronica.objCobrador.iCodBanco
    
'        Case 1
'            Set objCNABBanco = New ClassCNABBcoBrasil
'            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_BcoBrasil(objCobrancaEletronica)
                
        Case 230
            Set objCNABBanco = New ClassCNABBandeirantes
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Bandeirantes(objCobrancaEletronica)
        
        Case 237
            Set objCNABBanco = New ClassCNABBradesco
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Bradesco(objCobrancaEletronica)
        
'        Case 275
'            Set objCNABBanco = New ClassCNABReal
'            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_BcoReal(objCobrancaEletronica)
                
        Case 320
            Set objCNABBanco = New ClassCNABBicBanco
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_BicBanco(objCobrancaEletronica)
                
'        Case 341
'            Set objCNABBanco = New ClassCNABItau
'            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Itau(objCobrancaEletronica)
        
        Case 389
            Set objCNABBanco = New ClassCNABMercantil
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Mercantil(objCobrancaEletronica)
        
'        Case 409
'            Set objCNABBanco = New ClassCNABUnibanco
'            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Unibanco(objCobrancaEletronica)
    
        Case 422
            Set objCNABBanco = New ClassCNABSafra
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_BancoSafra(objCobrancaEletronica)
        
'        Case 453
'            Set objCNABBanco = New ClassCNABRural
'            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Rural(objCobrancaEletronica)
'
        Case 604
            Set objCNABBanco = New ClassCNABBcoIndustrial
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_BancoIndustrial(objCobrancaEletronica)
        
        Case 719
             Set objCNABBanco = New ClassCNABPrimus
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca_Primus(objCobrancaEletronica)
       
        'cnab 240
        Case 1, 104, 479, 275, 356, 409, 399, 341, 745, 8, 33, 353, 246, 41, 453, 749, 748, 637, 85
             Set objCNABBanco = New ClassCNAB240
            lErro = objCNABBanco.Processar_ArquivoRetorno_Cobranca(objCobrancaEletronica)
        
        Case Else
            Error 59355
                
    End Select
    If lErro <> SUCESSO Then Error 51753
    
    Processar_ArquivoRetorno_Cobranca = SUCESSO
    
    Exit Function

Erro_Processar_ArquivoRetorno_Cobranca:

    Processar_ArquivoRetorno_Cobranca = Err
    
    Select Case Err
    
        Case 51753
        
        Case 59355
            '??? falta colocar msg de erro
            MsgBox ("Banco não cadastrado")
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147241)
    
    End Select
    
    Exit Function
    
End Function

Public Function CobrancaEletronica_Abre_TelaRetornoArq(ByVal sNomeArqParam As String, ByVal objCobrancaEletronica As ClassCobrancaEletronica) As Long

Dim lErro As Long

On Error GoTo Erro_CobrancaEletronica_Abre_TelaRetornoArq

    ProcessaArqRetCobr2.sNomeArqParam = sNomeArqParam

    Set ProcessaArqRetCobr2.gobjCobrancaEletronica = objCobrancaEletronica
    
    ProcessaArqRetCobr2.Show vbModal
    
    Set ProcessaArqRetCobr2 = Nothing

    CobrancaEletronica_Abre_TelaRetornoArq = SUCESSO
            
    Exit Function

Erro_CobrancaEletronica_Abre_TelaRetornoArq:

    CobrancaEletronica_Abre_TelaRetornoArq = Err
    
    Select Case Err
    
        Case 62293, 62294
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147242)
            
    End Select

''    Call Rotina_ErrosBatch2("Processamento do Retorno da Cobrança")
    
    Exit Function

End Function

Public Function CobrancaEletronica_ObtemNumTitulos(ByVal objCobrancaEletronica As ClassCobrancaEletronica, lQuantRegistros As Long) As Long

Dim lErro As Long
Dim objCNABBanco As Object
Dim sRegistro As String
Dim sCodBanco As String
Dim iCodBanco As Integer

On Error GoTo Erro_CobrancaEletronica_ObtemNumTitulos

    Select Case objCobrancaEletronica.objCobrador.iCodBanco
    
        Case 1
            Set objCNABBanco = New ClassCNABBcoBrasil
            lErro = objCNABBanco.CobrancaEletronica_ObtemNumTitulos(objCobrancaEletronica, lQuantRegistros)
            If lErro <> SUCESSO Then gError 81541
    
        Case 275
            Set objCNABBanco = New ClassCNABReal
            lErro = objCNABBanco.CobrancaEletronica_ObtemNumTitulos(objCobrancaEletronica, lQuantRegistros)
            If lErro <> SUCESSO Then gError 81541
    
        'cnab 240
        Case 104, 479, 356, 409, 399, 341, 745, 8, 33, 353, 246, 422, 41, 453, 749, 748, 85
            Set objCNABBanco = New ClassCNAB240
            lErro = objCNABBanco.CobrancaEletronica_ObtemNumTitulos(objCobrancaEletronica, lQuantRegistros)
            If lErro <> SUCESSO Then gError 81541
    
        Case Else
        
            'Abre o arquivo de retorno
            Open objCobrancaEletronica.sNomeArquivoRetorno For Input As #1
                
            iCodBanco = objCobrancaEletronica.objCobrador.iCodBanco
            
            Do While Not EOF(1)
            
                'Busca os registro do arquivo
                Line Input #1, sRegistro
                
                lQuantRegistros = lQuantRegistros + 1
                
                If lQuantRegistros = 1 Then
                    If Mid(sRegistro, 1, 9) <> "02RETORNO" Then gError 32250
                    
                    sCodBanco = Mid(sRegistro, 77, 3)
'                    If StrParaInt(sCodBanco) <> iCodBanco Then gError 32249
                    
                End If
                
            Loop
                
            'excluindo o header e o trailer
            lQuantRegistros = lQuantRegistros - 2
            
            Close #1
    
    End Select
    
    CobrancaEletronica_ObtemNumTitulos = SUCESSO
     
    Exit Function
    
Erro_CobrancaEletronica_ObtemNumTitulos:

    CobrancaEletronica_ObtemNumTitulos = gErr
     
    Select Case gErr
          
        Case 81541
        
        Case 32250
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQRETCOBR_ARQ_INVALIDO", gErr)
    
        Case 32249
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQRETCOBR_BCO_NAO_CONFERE", gErr, StrParaInt(sCodBanco), iCodBanco)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147243)
     
    End Select
     
    Close #1
    
    Exit Function

End Function

Function RetCobr_Preenche_InfoParcRec(ByVal objParcelaRec As ClassParcelaReceber, objInfoParcRec As ClassInfoParcRec) As Long
'Preenche infoparcrec obtendo informacoes da parcela e do titulo que está sofrendo baixa

Dim lErro As Long, objTituloRec As New ClassTituloReceber

On Error GoTo Erro_RetCobr_Preenche_InfoParcRec

    Set objInfoParcRec = New ClassInfoParcRec
    
    'Lê a parcela a receber
    lErro = CF("ParcelaReceber_Le", objParcelaRec)
    If lErro <> SUCESSO And lErro <> 19147 Then gError 51683
    If lErro <> SUCESSO Then gError 51697
        
    objTituloRec.lNumIntDoc = objParcelaRec.lNumIntTitulo
    lErro = CF("TituloReceber_Le", objTituloRec)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 81866
    If lErro <> SUCESSO Then gError 81867
            
    With objInfoParcRec
        .lNumIntParc = objParcelaRec.lNumIntDoc
        .iNumParcela = objParcelaRec.iNumParcela
        .iCobrador = objParcelaRec.iCobrador
        .iCarteiraCobrador = objParcelaRec.iCarteiraCobranca
        .lNumTitulo = objTituloRec.lNumTitulo
        .lCliente = objTituloRec.lCliente
        .iFilialCliente = objTituloRec.iFilial
        .iFilialEmpresa = objTituloRec.iFilialEmpresa
    End With

    RetCobr_Preenche_InfoParcRec = SUCESSO
     
    Exit Function
    
Erro_RetCobr_Preenche_InfoParcRec:

    RetCobr_Preenche_InfoParcRec = gErr
     
    Select Case gErr
          
        Case 51683, 81866
        
        Case 81867
            'Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", gErr)'Comentado por Wagner
            
        Case 51697
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", gErr, objParcelaRec.lNumIntDoc) 'Comentado por Wagner
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147244)
     
    End Select
     
    Exit Function

End Function
    
Function BorderoRetCobr_ConfirmaGrava(ByVal objBorderoRetCobr As ClassBorderoRetCobr) As Boolean

Dim lErro As Long, objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim bDesistiu As Boolean, dValorMov As Double

On Error GoTo Erro_BorderoRetCobr_ConfirmaGrava

    Call RetCobr_ConfirmaAtualizacao(objBorderoRetCobr, bDesistiu)
    If lErro <> SUCESSO Then Error 32260

    If objBorderoRetCobr.colRetCobrErros.Count <> 0 And Not ProcessaArqRetCobr2.bTeste Then
        '???
        If ProcessaArqRetCobr2.iGravaCriticas = MARCADO Then Call RetCobr_Grava_Erros(objBorderoRetCobr)
        'If MsgBox("Grava críticas do arquivo de retorno ?", vbYesNo, "Bordero de Retorno de Cobrança") = vbYes Then Call RetCobr_Grava_Erros(objBorderoRetCobr)
    End If

    If Not bDesistiu And Not ProcessaArqRetCobr2.bTeste Then

        'atualiza o valor a ser debitado ou creditado por carteira
        For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart

            dValorMov = Arredonda_Moeda(objBorderoRetCobrCart.dValorRecebido - objBorderoRetCobrCart.dValorCustas - objBorderoRetCobrCart.dValorTarifas + objBorderoRetCobrCart.dValorTarifasNaoCred, 2)
            objBorderoRetCobrCart.objMovCCI.dValor = Abs(dValorMov)

            If dValorMov >= 0 Then

                objBorderoRetCobrCart.objMovCCI.iTipo = MOVCCI_CRED_RET_COBRANCA

            Else

                objBorderoRetCobrCart.objMovCCI.iTipo = MOVCCI_DEB_RET_COBRANCA

            End If

        Next

        'Grava os dados recolhidos durante o processamento do arquivo de retorno
        lErro = CF("BorderoRetCobr_Grava", objBorderoRetCobr)
        If lErro <> SUCESSO Then Error 51829

    End If

    BorderoRetCobr_ConfirmaGrava = SUCESSO
     
    Exit Function
    
Erro_BorderoRetCobr_ConfirmaGrava:

    BorderoRetCobr_ConfirmaGrava = gErr
     
    Select Case gErr
          
        Case 32260, 51829
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147245)
     
    End Select
     
    Exit Function

End Function

Private Function RetCobr_ConfirmaAtualizacao(objBorderoRetCobr As ClassBorderoRetCobr, bDesistiu As Boolean) As Long
'Dá chance ao usuario para desistir de atualizar bordero de retorno de cobranca conferindo os valores do creditos, custas e tarifas que serao atualizados

Dim lErro As Long, objBorderoRetCobrCart As ClassBorderoRetCobrCart
Dim dValorRecebido As Double, dValorTarifas As Double, dValorCustas As Double, dValorBaixado As Double

On Error GoTo Erro_RetCobr_ConfirmaAtualizacao

    For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
    
        dValorBaixado = Arredonda_Moeda(dValorBaixado + objBorderoRetCobrCart.dValorBaixado, 2)
        dValorRecebido = Arredonda_Moeda(dValorRecebido + objBorderoRetCobrCart.dValorRecebido, 2)
        dValorTarifas = Arredonda_Moeda(dValorTarifas + objBorderoRetCobrCart.dValorTarifas, 2)
        dValorCustas = Arredonda_Moeda(dValorCustas + objBorderoRetCobrCart.dValorCustas, 2)
        
    Next

    bDesistiu = False
    
    If ProcessaArqRetCobr2.bTeste Then
        ProcessaArqRetCobr2.dValorRecebido = dValorRecebido
        ProcessaArqRetCobr2.dValorTarifas = dValorTarifas
        ProcessaArqRetCobr2.dValorCustas = dValorCustas
        ProcessaArqRetCobr2.dValorBaixado = dValorBaixado
    End If
    
'    If Not ProcessaArqRetCobr2.bTeste Then
'        '??? abrir form pedindo confirmacao
'        If MsgBox("Atenção: confirma os valores baixado " & Format(dValorBaixado, "standard") & " , recebido " & Format(dValorRecebido, "standard") & " , de tarifas " & Format(dValorTarifas, "standard") & " , de custas " & Format(dValorCustas, "standard") & " e prossegue com a atualização ?", vbYesNo, "Bordero de Retorno de Cobrança") <> vbYes Then
'
'            bDesistiu = True
'
'        End If
'    End If
    RetCobr_ConfirmaAtualizacao = SUCESSO
    
    Exit Function
    
Erro_RetCobr_ConfirmaAtualizacao:

    RetCobr_ConfirmaAtualizacao = Err
    
    Select Case Err
    
        Case 32259
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147246)
            
    End Select

    Exit Function

End Function

Function RetCobr_Grava_Erros_Trans(objBorderoRetCobr As ClassBorderoRetCobr) As Long

Dim lErro As Long, objDetRetCobr As ClassDetRetCobr, lComando As Long, lNumIntDoc As Long
Dim objOcorrRetParcRec As ClassOcorrRetParcRec, bAtualizou As Boolean, objBorderoRetCobrCart As ClassBorderoRetCobrCart

On Error GoTo Erro_RetCobr_Grava_Erros_Trans

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 32277
    
    lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_RETCOBRERROS", lNumIntDoc, objBorderoRetCobr.colRetCobrErros.Count)
    If lErro <> SUCESSO Then Error 32280
    
    lErro = Comando_Executar(lComando, "DELETE FROM RetCobrErros WHERE Banco = ? AND NomeArq = ? AND DataArq = ?", objBorderoRetCobr.iBanco, objBorderoRetCobr.sNomeArq, objBorderoRetCobr.dtDataRecepcao)
    If lErro <> AD_SQL_SUCESSO Then Error 32278
    
    For Each objDetRetCobr In objBorderoRetCobr.colRetCobrErros
    
        objDetRetCobr.lNumIntDoc = lNumIntDoc
        
        With objDetRetCobr
            lErro = Comando_Executar(lComando, "INSERT INTO RetCobrErros (Banco, NomeArq, DataArq, SeqReg, CodCobrador, NumCarteiraBanco, Ocorrencia, DataVencimento, ValorTitulo, ValorRecebido, ValorJuros, ValorDesconto, ValorTarifa, SeuNumero, NossoNumero, UsoDaEmpresa, DataCredito, TipoCritica, CodOcorrencia1, CodOcorrencia2, CodOcorrencia3, CodOcorrencia4, CodOcorrencia5, NumIntDoc, NumIntParc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objBorderoRetCobr.iBanco, objBorderoRetCobr.sNomeArq, objBorderoRetCobr.dtDataRecepcao, .lNumSequencial, objBorderoRetCobr.iCobrador, .iCodCarteiraNoBanco, .iCodOcorrencia, .dtDataVencimento, .dValorTitulo, .dValorRecebido, .dValorJuros, .dValorDesconto, .dValorTarifa, right(Trim(.sSeuNumero), 10), Trim(.sNossoNumero), Trim(.sUsoDaEmpresa), IIf(.dtDataCredito < DATA_NULA, DATA_NULA, .dtDataCredito), .iTipoCritica, .iCodOcorrencia1, .iCodOcorrencia2, .iCodOcorrencia3, .iCodOcorrencia4, .iCodOcorrencia5, lNumIntDoc, .lNumIntParc)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 32279
            
        'atualiza o lNumIntRetCobrErros na ocorrencia
        bAtualizou = False
        For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
        
            For Each objOcorrRetParcRec In objBorderoRetCobrCart.colOcorrRetParcRec
            
                If objOcorrRetParcRec.lNumSequencial = objDetRetCobr.lNumSequencial Then
                    objOcorrRetParcRec.lNumIntRetCobrErros = lNumIntDoc
                    bAtualizou = True
                    Exit For
                End If
                
                If bAtualizou = True Then Exit For
                
            Next
        
        Next
        
        lNumIntDoc = lNumIntDoc + 1
    
    Next
    
    Call Comando_Fechar(lComando)
    
    RetCobr_Grava_Erros_Trans = SUCESSO

    Exit Function

Erro_RetCobr_Grava_Erros_Trans:

    RetCobr_Grava_Erros_Trans = Err

    Select Case Err

        Case 32280
        
        Case 32277
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 32278
            Call Rotina_Erro(vbOKOnly, "ERRO_RETCOBR_EXCLUSAO_ERROS", Err)

        Case 32279
            Call Rotina_Erro(vbOKOnly, "ERRO_RETCOBR_INCLUSAO_ERRO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147247)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function RetCobr_Grava_Erros(objBorderoRetCobr As ClassBorderoRetCobr) As Long

Dim lErro As Long, lTransacao As Long

On Error GoTo Erro_RetCobr_Grava_Erros

    'Entra em transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 32274

    lErro = RetCobr_Grava_Erros_Trans(objBorderoRetCobr)
    If lErro <> SUCESSO Then Error 32275
    
    'Confirma transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 32276

    RetCobr_Grava_Erros = SUCESSO

    Exit Function

Erro_RetCobr_Grava_Erros:

    RetCobr_Grava_Erros = Err

    Select Case Err

        Case 32275
        
        Case 32274
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 32276
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147248)
            
    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Processa_ParcelaRec(ByVal objParcelaRec As ClassParcelaReceber, ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objBorderoRetCobrCart As ClassBorderoRetCobrCart, ByVal objDetRetCobr As ClassDetRetCobr, ByVal objInfoParcRec As ClassInfoParcRec) As Long
'Faz o processamento da parcela a receber

Dim lErro As Long, dValorBaixado As Double
Dim objOcorrRetParcRec As New ClassOcorrRetParcRec

On Error GoTo Erro_Processa_ParcelaRec

    If objDetRetCobr.iTipoCritica <> RETCOBR_CRITICA_TARIFAS Then
    
        'Acumula-se os valores creditados para o criação do movimento de conta corrente
        If objParcelaRec.iCarteiraCobranca <> CARTEIRA_DESCONTADA And objBorderoRetCobrCart.dtDataCredito >= objBorderoRetCobr.dtDataCredito Then
            
            If objDetRetCobr.dValorTarifa <> 0 And objBorderoRetCobr.iBanco = 341 Then
            
                objBorderoRetCobrCart.dValorTarifas = Arredonda_Moeda(objBorderoRetCobrCart.dValorTarifas + objDetRetCobr.dValorTarifa, 2)
                objDetRetCobr.dValorRecebido = Arredonda_Moeda(objDetRetCobr.dValorRecebido + objDetRetCobr.dValorTarifa, 2)
                
            End If
            
            objBorderoRetCobrCart.dValorRecebido = Arredonda_Moeda(objBorderoRetCobrCart.dValorRecebido + objDetRetCobr.dValorRecebido, 2)
            objBorderoRetCobrCart.dValorJuros = Arredonda_Moeda(objBorderoRetCobrCart.dValorJuros + objDetRetCobr.dValorJuros, 2)
            objBorderoRetCobrCart.dValorDescontos = Arredonda_Moeda(objBorderoRetCobrCart.dValorDescontos + objDetRetCobr.dValorDesconto, 2)
            
            dValorBaixado = Arredonda_Moeda(objDetRetCobr.dValorRecebido + objDetRetCobr.dValorDesconto - objDetRetCobr.dValorJuros, 2)
               
        End If
        
        If objParcelaRec.iCarteiraCobranca = CARTEIRA_DESCONTADA Then
            dValorBaixado = objDetRetCobr.dValorTitulo
        End If
        
        objBorderoRetCobrCart.dValorBaixado = Arredonda_Moeda(objBorderoRetCobrCart.dValorBaixado + dValorBaixado, 2)
        
        'se a baixa nao foi total incluir o reg p/sair no relatorio de criticas
        If Abs(dValorBaixado - objParcelaRec.dSaldo) > 0.01 Then
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_BAIXA_PARCIAL
            '...já foi incluido antes da chamada a esta funcao Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
            
        End If
        
        'Gera uma baixaparcrec p\ a parcela a receber em questão
        objBorderoRetCobrCart.colBaixaParcRec.Add 0, 0, objParcelaRec.lNumIntDoc, 0, STATUS_LANCADO, 0, objDetRetCobr.dValorJuros, objDetRetCobr.dValorDesconto, dValorBaixado, objDetRetCobr.dValorRecebido, objBorderoRetCobr.iCobrador
        
        'guarda a informacao que será util p/gerar a contabilizacao
        objBorderoRetCobrCart.colInfoParcRec.Add objInfoParcRec
        
    End If
    
    'Gera um ocorrenciaRetParcrec de acordo com os dados retornados
    objOcorrRetParcRec.dtDataComplementar = DATA_NULA
    objOcorrRetParcRec.lNumIntParc = objParcelaRec.lNumIntDoc
    objOcorrRetParcRec.lNumSeqOcorr = 0
    objOcorrRetParcRec.sNumTitCobrador = objDetRetCobr.sNossoNumero
    objOcorrRetParcRec.iCodMovimento = objDetRetCobr.iCodOcorrencia
    objOcorrRetParcRec.dtDataComplementar = DATA_NULA
    objOcorrRetParcRec.iCodOcorrencia1 = objDetRetCobr.iCodOcorrencia1
    objOcorrRetParcRec.iCodOcorrencia2 = objDetRetCobr.iCodOcorrencia2
    objOcorrRetParcRec.iCodOcorrencia3 = objDetRetCobr.iCodOcorrencia3
    objOcorrRetParcRec.iCodOcorrencia4 = objDetRetCobr.iCodOcorrencia4
    objOcorrRetParcRec.iCodOcorrencia5 = objDetRetCobr.iCodOcorrencia5
    objOcorrRetParcRec.lNumSequencial = objDetRetCobr.lNumSequencial
        
    'Adiciona a ocorrência na coleção
    objBorderoRetCobrCart.colOcorrRetParcRec.Add objOcorrRetParcRec
        
    Processa_ParcelaRec = SUCESSO
    
    Exit Function

Erro_Processa_ParcelaRec:

    Processa_ParcelaRec = Err
    
    Select Case Err
    
        Case 51691, 51692
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147249)
            
    End Select

    Exit Function

End Function

Private Sub RetCobrCart_Obtem(objDetRetCobr As ClassDetRetCobr, objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart)
'obtem dentro do bordero de retorno objBorderoRetCobrCart correspondente à carteira passada
'Se ainda nao houver em objBorderoRetCobr retorna nothing.

Dim dtDataCredito As Date, iCodCarteiraNoBanco As Integer, dtDataBaixa As Date
Dim iStatusRecurso As Integer

    iCodCarteiraNoBanco = objDetRetCobr.iCodCarteiraNoBanco
    dtDataCredito = IIf(objDetRetCobr.dtDataCredito <> DATA_NULA, objDetRetCobr.dtDataCredito, objBorderoRetCobr.dtDataCredito)
    If dtDataCredito = DATA_NULA Then dtDataCredito = objDetRetCobr.dtDataBaixa
    dtDataBaixa = objDetRetCobr.dtDataBaixa
    
    Set objBorderoRetCobrCart = Nothing
    
    If objBorderoRetCobr.iQuebraMovCCPorStatusRecurso = MARCADO Then
        Call Ocorrencia_Obtem_StatusRecurso(objBorderoRetCobr.iBanco, objDetRetCobr.sCodLiquidacao, iStatusRecurso)
    Else
        iStatusRecurso = 0
    End If
    For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
    
        If objBorderoRetCobrCart.iStatusRecurso = iStatusRecurso And objBorderoRetCobrCart.iCodCarteiraNoBanco = iCodCarteiraNoBanco And objBorderoRetCobrCart.dtDataCredito = dtDataCredito And objBorderoRetCobrCart.objMovCCI.dtDataBaixa = dtDataBaixa Then Exit For
        
    Next
    
    '??? titulo no itau vai numa carteira e volta em outra
    If objBorderoRetCobr.iBanco = 341 And objBorderoRetCobrCart Is Nothing And iCodCarteiraNoBanco = 358 Then
    
        iCodCarteiraNoBanco = 354
        
        For Each objBorderoRetCobrCart In objBorderoRetCobr.colBorderoRetCobrCart
        
            If objBorderoRetCobrCart.iStatusRecurso = iStatusRecurso And objBorderoRetCobrCart.iCodCarteiraNoBanco = iCodCarteiraNoBanco And objBorderoRetCobrCart.dtDataCredito = dtDataCredito And objBorderoRetCobrCart.objMovCCI.dtDataBaixa = dtDataBaixa Then Exit For
            
        Next
    
    End If

End Sub

Private Sub Ocorrencia_Obtem_StatusRecurso(ByVal iBanco As Integer, ByVal sCodLiquidacao As String, iStatusRecurso As Integer)
    iStatusRecurso = 0 'Disponível
    Select Case iBanco
        Case 341  'Itaú
            'AA---CAIXA ELETRÔNICO BANCO ITAÚ-------------------------------------------------------------DISPONÍVEL
            'AC---PAGAMENTO EM CARTÓRIO AUTOMATIZADO------------------------------------------------------A COMPENSAR
            'AO---ACERTO ONLINE---------------------------------------------------------------------------DISPONÍVEL
            'BC---BANCOS CORRESPONDENTES------------------------------------------------------------------DISPONÍVEL
            'BF---ITAÚ BANKFONE---------------------------------------------------------------------------DISPONÍVEL
            'BL---ITAÚ BANKLINE---------------------------------------------------------------------------DISPONÍVEL
            'B0---OUTROS BANCOS – RECEBIMENTO OFF-LINE----------------------------------------------------A COMPENSAR
            'B1---OUTROS BANCOS – PELO CÓDIGO DE BARRAS---------------------------------------------------A COMPENSAR
            'B2---OUTROS BANCOS – PELA LINHA DIGITÁVEL----------------------------------------------------A COMPENSAR
            'B3---OUTROS BANCOS – PELO AUTO ATENDIMENTO---------------------------------------------------A COMPENSAR
            'B4---OUTROS BANCOS – RECEBIMENTO EM CASA LOTÉRICA--------------------------------------------A COMPENSAR
            'B5---OUTROS BANCOS – CORRESPONDENTE----------------------------------------------------------A COMPENSAR
            'B6---OUTROS BANCOS – TELEFONE----------------------------------------------------------------A COMPENSAR
            'B7---OUTROS BANCOS – ARQUIVO ELETRÔNICO (Pagamento Efetuado por meio de troca de arquivos)---A COMPENSAR
            'CC---AGÊNCIA ITAÚ – COM CHEQUE DE OUTRO BANCO ou (CHEQUE ITAÚ)*------------------------------A COMPENSAR
            'CI---CORRESPONDENTE ITAÚ---------------------------------------------------------------------DISPONÍVEL
            'CK---SISPAG – SISTEMA DE CONTAS A PAGAR ITAÚ-------------------------------------------------DISPONÍVEL
            'CP---AGÊNCIA ITAÚ – POR DÉBITO EM CONTA CORRENTE, CHEQUE ITAÚ* OU DINHEIRO-------------------DISPONÍVEL
            'DG---AGÊNCIA ITAÚ – CAPTURADO EM OFF-LINE----------------------------------------------------DISPONÍVEL
            'LC---PAGAMENTO EM CARTÓRIO DE PROTESTO COM CHEQUE--------------------------------------------A COMPENSAR
            'EA---TERMINAL DE CAIXA-----------------------------------------------------------------------DISPONÍVEL
            'Q0---AGENDAMENTO – PAGAMENTO AGENDADO VIA BANKLINE OU OUTRO CANAL ELETRÔNICO E ...---------- DISPONÍVEL
            'RA---DIGITAÇÃO – REALIMENTAÇÃO AUTOMÁTICA----------------------------------------------------DISPONÍVEL
            'ST---PAGAMENTO VIA SELTEC**------------------------------------------------------------------DISPONÍVEL
            Select Case sCodLiquidacao
                Case "AC", "B0", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "CC", "LC"
                    iStatusRecurso = 1 'a compensar
            End Select
    End Select
End Sub

Function Prepara_CartCobrRet(ByVal objCobrancaEletronica As ClassCobrancaEletronica, ByVal objBorderoRetCobr As ClassBorderoRetCobr, objBorderoRetCobrCart As ClassBorderoRetCobrCart, ByVal objDetRetCobr As ClassDetRetCobr) As Long
'codigo comum ao tratamento de registros detalhe de baixas, custas e tarifas
'preenche objBorderoRetCobrCart e objDetRetCobr

Dim lErro As Long

On Error GoTo Erro_Prepara_CartCobrRet

    Call RetCobrCart_Obtem(objDetRetCobr, objBorderoRetCobr, objBorderoRetCobrCart)

    If objBorderoRetCobrCart Is Nothing Then
    
        Set objBorderoRetCobrCart = New ClassBorderoRetCobrCart
        
        lErro = CF("RetCobrCart_Inicializa", objCobrancaEletronica, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr)
        If lErro <> SUCESSO Then Error 32255
        
        Call objBorderoRetCobr.colBorderoRetCobrCart.Add(objBorderoRetCobrCart)
                
    End If
    
    Prepara_CartCobrRet = SUCESSO
    
    Exit Function
    
Erro_Prepara_CartCobrRet:

    Prepara_CartCobrRet = Err
    
    Select Case Err
    
        Case 32255
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147250)
            
    End Select

    Exit Function

End Function

Function RetCobrCart_Inicializa(ByVal objCobrancaEletronica As ClassCobrancaEletronica, ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objBorderoRetCobrCart As ClassBorderoRetCobrCart, ByVal objDetRetCobr As ClassDetRetCobr) As Long
'inicializa carteira de cobranca

Dim lErro As Long, objCarteiraCobrador As ClassCarteiraCobrador
Dim objContaCorrente As New ClassContasCorrentesInternas, dtDataMov As Date
Dim bAchou As Boolean, iStatusRecurso As Integer

On Error GoTo Erro_RetCobrCart_Inicializa

    objBorderoRetCobrCart.iCobrador = objBorderoRetCobr.iCobrador
    objBorderoRetCobrCart.iCodCarteiraNoBanco = objDetRetCobr.iCodCarteiraNoBanco
   
    Call objBorderoRetCobr.CarteiraCobranca_Obter(objBorderoRetCobrCart.iCodCarteiraNoBanco, objCarteiraCobrador)
    If objCarteiraCobrador Is Nothing Then

        '??? titulo no itau vai numa carteira e volta em outra
        If objCobrancaEletronica.objCobrador.iCodBanco = 341 And objDetRetCobr.iCodCarteiraNoBanco = 358 Then

            objBorderoRetCobrCart.iCodCarteiraNoBanco = 354
            Call objBorderoRetCobr.CarteiraCobranca_Obter(objBorderoRetCobrCart.iCodCarteiraNoBanco, objCarteiraCobrador)

        End If
        
        '######################################
        'Inserido por Wagner
        If objCobrancaEletronica.objCobrador.iCodBanco = 356 Then
        
            If objDetRetCobr.iCodCarteiraNoBanco = 1 Then
                For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
                        Exit For
                    End If
                Next
            End If
            
            If objDetRetCobr.iCodCarteiraNoBanco = 3 Then
                 For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
                        Exit For
                    End If
                Next
            End If
        
        End If
        '######################################

        If objCobrancaEletronica.objCobrador.iCodBanco = 33 Then
        
            bAchou = False
            
            If objDetRetCobr.iCodCarteiraNoBanco = 1 Or objDetRetCobr.iCodCarteiraNoBanco = 2 Then
                For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_SIMPLES Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
            
            If objDetRetCobr.iCodCarteiraNoBanco = 3 Or objDetRetCobr.iCodCarteiraNoBanco = 6 Then
                 For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CAUCIONADA Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
        
            If objDetRetCobr.iCodCarteiraNoBanco = 4 Then
                For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    If objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_DESCONTADA Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
            
            '??? nao deveria cair aqui
            If bAchou = False Then
            
                For Each objCarteiraCobrador In objBorderoRetCobr.colCartCobr
                    
                    If objDetRetCobr.iCodCarteiraNoBanco = 0 Then
                        If objBorderoRetCobr.colBorderoRetCobrCart.Count <> 0 Then
                            objDetRetCobr.iCodCarteiraNoBanco = objBorderoRetCobr.colBorderoRetCobrCart.Item(1).iCodCarteiraNoBanco
                        End If
                    End If
                    
                    bAchou = True
                    Exit For
                    
                Next
                
            End If
        
        End If
        
        '??? tarifa na cef pode vir zerada
        If objDetRetCobr.iCodCarteiraNoBanco = 0 Then

            If objBorderoRetCobr.colBorderoRetCobrCart.Count <> 0 Then

                objBorderoRetCobrCart.iCodCarteiraNoBanco = objBorderoRetCobr.colBorderoRetCobrCart.Item(1).iCodCarteiraNoBanco
                Call objBorderoRetCobr.CarteiraCobranca_Obter(objBorderoRetCobrCart.iCodCarteiraNoBanco, objCarteiraCobrador)
                
            End If

        End If
        
        If objCarteiraCobrador Is Nothing Then Error 32266

    End If
    
    objBorderoRetCobrCart.iCarteiraCobr = objCarteiraCobrador.iCodCarteiraCobranca
    
    'Pega a conta corrente do cobrador
    objContaCorrente.iCodigo = objCobrancaEletronica.objCobrador.iCodCCI
    
    'Lê a conta corrente do cobrador
    lErro = CF("ContaCorrenteInt_Le", objContaCorrente.iCodigo, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 51711
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 51712
    
    'Preenche parcialmente o objMovCCI
    With objBorderoRetCobrCart.objMovCCI
    
        .iFilialEmpresa = objContaCorrente.iFilialEmpresa
        .iCodConta = objContaCorrente.iCodigo
        .iExcluido = NAO_EXCLUIDO
        .iTipoMeioPagto = DINHEIRO
        .dtDataBaixa = objDetRetCobr.dtDataBaixa
        If objDetRetCobr.dtDataCredito <> DATA_NULA Then
            dtDataMov = objDetRetCobr.dtDataCredito
        ElseIf objBorderoRetCobr.dtDataCredito <> DATA_NULA Then
            dtDataMov = objBorderoRetCobr.dtDataCredito
        Else
            dtDataMov = objDetRetCobr.dtDataBaixa
        End If
        .dtDataMovimento = dtDataMov
        .dtDataContabil = dtDataMov
        .iConciliado = NAO_CONCILIADO
           
    End With
    
    objBorderoRetCobrCart.dtDataCredito = objBorderoRetCobrCart.objMovCCI.dtDataMovimento
    
    If objBorderoRetCobr.iQuebraMovCCPorStatusRecurso = MARCADO Then
        Call Ocorrencia_Obtem_StatusRecurso(objBorderoRetCobr.iBanco, objDetRetCobr.sCodLiquidacao, iStatusRecurso)
        objBorderoRetCobrCart.iStatusRecurso = iStatusRecurso
    Else
        objBorderoRetCobrCart.iStatusRecurso = 0
    End If
    
    RetCobrCart_Inicializa = SUCESSO
    
    Exit Function
    
Erro_RetCobrCart_Inicializa:

    RetCobrCart_Inicializa = Err
    
    Select Case Err

        Case 32266
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBR_INVALIDA", Err, objBorderoRetCobrCart.iCodCarteiraNoBanco, objBorderoRetCobr.iCobrador)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147251)

    End Select

    Exit Function

End Function

Function RetCobr_Obtem_Parcela(ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objDetRetCobr As ClassDetRetCobr, objParcelaRec As ClassParcelaReceber, ByVal iTipoSeuNumero As Integer, lComando As Long) As Long

Dim lErro As Long, sNumTitulo As String, sParcela As String, sFilialEmpTit As String

On Error GoTo Erro_RetCobr_Obtem_Parcela

    Call objBorderoRetCobr.Abrir_Comando

    'Tenta pegar pelo nosso número
    lErro = CF("RetCobr_Obtem_Parcela4", objBorderoRetCobr, objDetRetCobr, objParcelaRec)
    If lErro <> SUCESSO And lErro <> 130041 And lErro <> 130042 Then gError ERRO_SEM_MENSAGEM
    
    '130041 - Não achou, 130042 achou mas com erro
    If lErro = 130041 Then
        If iTipoSeuNumero = RETCOBR_TIPO_SEU_NUMERO4 And IsNumeric(objDetRetCobr.sSeuNumero) Then
        
            'Le pelo NumIntDoc da Ocorrência
            lErro = CF("RetCobr_Obtem_Parcela3", objDetRetCobr, objParcelaRec)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        Else
            
            'Le pelo número do título e parcela
            lErro = CF("RetCobr_Obtem_Parcela1", objDetRetCobr, sNumTitulo, sParcela, sFilialEmpTit, iTipoSeuNumero)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            lErro = CF("RetCobr_Obtem_Parcela2", objBorderoRetCobr, objDetRetCobr, objParcelaRec, sNumTitulo, sParcela, sFilialEmpTit)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
    End If
    
    objDetRetCobr.lNumIntParc = objParcelaRec.lNumIntDoc
    
    Call objBorderoRetCobr.Fechar_Comando
    
    RetCobr_Obtem_Parcela = SUCESSO
     
    Exit Function
    
Erro_RetCobr_Obtem_Parcela:

    RetCobr_Obtem_Parcela = gErr
     
    Select Case gErr
          
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147252)
     
    End Select
     
    Call objBorderoRetCobr.Fechar_Comando
     
    Exit Function

End Function

Function RetCobr_Obtem_Parcela1(ByVal objDetRetCobr As ClassDetRetCobr, sNumTitulo As String, sParcela As String, sFilialEmpTit As String, ByVal iTipoSeuNumero As Integer) As Long

Dim lErro As Long, iPosSep As Integer, sSeuNumero As String, sAux As String, iAux As Integer
Dim iTamMax As Integer

On Error GoTo Erro_RetCobr_Obtem_Parcela1

    sSeuNumero = objDetRetCobr.sSeuNumero
    
    Select Case iTipoSeuNumero
    
        Case RETCOBR_TIPO_SEU_NUMERO1 'geral
        
            If Len(Trim(sSeuNumero)) <> 10 Then
            
                iPosSep = InStr(sSeuNumero, SEPARADOR)
                If iPosSep <> 0 And IsNumeric(left(sSeuNumero, iPosSep - 1)) Then
                    sParcela = Trim(Mid(sSeuNumero, iPosSep + 1))
                    sNumTitulo = left(sSeuNumero, iPosSep - 1)
                End If
                
            Else
            
                sFilialEmpTit = left(sSeuNumero, 2)
                sParcela = Mid(sSeuNumero, 3, 2)
                sNumTitulo = Mid(sSeuNumero, 5, 6)
                
            End If
    
        Case RETCOBR_TIPO_SEU_NUMERO2 'bb
        
            If UCase(left(gsNomeEmpresa, 9)) = UCase("AsistBras") Then
                iTamMax = 7
            Else
                iTamMax = 6
            End If
        
            If Len(Trim(sSeuNumero)) = 0 Then sSeuNumero = right(Trim(objDetRetCobr.sNossoNumero), iTamMax)
            
            iPosSep = InStr(sSeuNumero, SEPARADOR)
            If iPosSep <> 0 Then
                sNumTitulo = Mid(sSeuNumero, 1, iPosSep - 1)
                sParcela = Mid(sSeuNumero, iPosSep + 1, 2)
            Else
                If IsNumeric(Trim(sSeuNumero)) Then
                    sNumTitulo = Mid(sSeuNumero, 1, iTamMax)
                    If StrParaLong(sNumTitulo) = 0 Then sNumTitulo = right(sSeuNumero, iTamMax)
                Else
                    sNumTitulo = Trim(sSeuNumero)
                End If
                sParcela = "1"
            End If
    
        Case Else
        
            sParcela = "1"
            sNumTitulo = "0"
        
            iPosSep = InStr(sSeuNumero, "/")
            If iPosSep <> 0 Then
                If IsNumeric(left(sSeuNumero, iPosSep - 1)) And IsNumeric(Trim(Mid(sSeuNumero, iPosSep + 1))) Then
                    sParcela = Trim(Mid(sSeuNumero, iPosSep + 1))
                    sNumTitulo = left(sSeuNumero, iPosSep - 1)
                End If
            End If
            
    End Select
    
    RetCobr_Obtem_Parcela1 = SUCESSO
     
    Exit Function
    
Erro_RetCobr_Obtem_Parcela1:

    RetCobr_Obtem_Parcela1 = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147253)
     
    End Select
     
    Exit Function

End Function

Function RetCobr_Obtem_Parcela2(ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objDetRetCobr As ClassDetRetCobr, ByVal objParcelaRec As ClassParcelaReceber, ByVal sNumTitulo As String, ByVal sParcela As String, ByVal sFilialEmpTit As String) As Long
'preenche objParcelaRec lendo do bd identificando a parcela por dados em objDetRetCobr

Dim lErro As Long, lNumTitulo As Long, iNumParcela As Integer
Dim lNumIntDoc As Long, iFilialEmpTit As Integer, sFiltro As String
Dim lComando As Long, objCobrador1 As New ClassCobrador, objCobrador2 As New ClassCobrador

On Error GoTo Erro_RetCobr_Obtem_Parcela2

    If IsNumeric(sNumTitulo) Then
        lNumTitulo = StrParaLong(right(Trim(sNumTitulo), 9))
    Else
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
        gError 32269
    End If
        
    Call CF("FilialEmpresa_Le_Filtro_Filial", sFiltro)
    
    If IsNumeric(sFilialEmpTit) Then
        'Busca na filial empresa do titulo com o mesmo cobrador ou cobrador = próprio com o número e vencimento
        iFilialEmpTit = StrParaInt(sFilialEmpTit)
        lErro = Comando_Executar(objBorderoRetCobr.alComando(4), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.FilialEmpresa = ? AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, iFilialEmpTit, lNumTitulo, objDetRetCobr.dtDataVencimento, STATUS_EXCLUIDO, STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
        lComando = objBorderoRetCobr.alComando(4)
    Else
        If objDetRetCobr.dtDataVencimento <> DATA_NULA Then
            'Busca com o mesmo cobrador ou cobrador = próprio com o número e vencimento aproximado
            lErro = Comando_Executar(objBorderoRetCobr.alComando(5), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento BETWEEN ? AND ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, lNumTitulo, DateAdd("d", -15, objDetRetCobr.dtDataVencimento), DateAdd("d", 15, objDetRetCobr.dtDataVencimento), STATUS_EXCLUIDO, STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
            lComando = objBorderoRetCobr.alComando(5)
        Else
            'Busca com o mesmo cobrador ou cobrador = próprio e com o número do título
            lErro = Comando_Executar(objBorderoRetCobr.alComando(6), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, lNumTitulo, STATUS_EXCLUIDO, STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
            lComando = objBorderoRetCobr.alComando(6)
        End If
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 32267
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32268
    
    If lErro = AD_SQL_SEM_DADOS Then
    
        'TENTA ACHAR COM BAIXADA
        If IsNumeric(sFilialEmpTit) Then
            'Busca na filial empresa do titulo com o mesmo cobrador ou cobrador = próprio com o número, parcelas e vencimento
            iFilialEmpTit = StrParaInt(sFilialEmpTit)
            lErro = Comando_Executar(objBorderoRetCobr.alComando(7), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRecTodas AS ParcelasRec,TitulosRecTodos AS TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.FilialEmpresa = ? AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND ParcelasRec.Status = ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, iFilialEmpTit, lNumTitulo, objDetRetCobr.dtDataVencimento, STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
            lComando = objBorderoRetCobr.alComando(7)
        Else
            If objDetRetCobr.dtDataVencimento <> DATA_NULA Then
                'Busca com o mesmo cobrador ou cobrador = próprio com o número e vencimento aproximado
                lErro = Comando_Executar(objBorderoRetCobr.alComando(8), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRecTodas AS ParcelasRec,TitulosRecTodos AS TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento BETWEEN ? AND ? AND ParcelasRec.Status = ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, lNumTitulo, DateAdd("d", -15, objDetRetCobr.dtDataVencimento), DateAdd("d", 15, objDetRetCobr.dtDataVencimento), STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
                lComando = objBorderoRetCobr.alComando(8)
            Else
                'Busca com o mesmo cobrador ou cobrador = próprio e com o número do título
                lErro = Comando_Executar(objBorderoRetCobr.alComando(9), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRecTodas AS ParcelasRec,TitulosRecTodos AS TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND ParcelasRec.Status = ? AND (ParcelasRec.Cobrador = ? OR ParcelasRec.Cobrador = ? )" & sFiltro, lNumIntDoc, lNumTitulo, STATUS_BAIXADO, objBorderoRetCobr.iCobrador, COBRADOR_PROPRIA_EMPRESA)
                lComando = objBorderoRetCobr.alComando(9)
            End If
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 32267
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32268
        
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Tenta achar em outro cobrador -> Existem casos em que vai em um conrador e volta em outro
            If IsNumeric(sFilialEmpTit) Then
                'Busca na filial empresa do titulo com o número e vencimento
                iFilialEmpTit = StrParaInt(sFilialEmpTit)
                lErro = Comando_Executar(objBorderoRetCobr.alComando(11), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.FilialEmpresa = ? AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ?" & sFiltro, lNumIntDoc, iFilialEmpTit, lNumTitulo, objDetRetCobr.dtDataVencimento, STATUS_EXCLUIDO, STATUS_BAIXADO)
                lComando = objBorderoRetCobr.alComando(11)
            Else
                If objDetRetCobr.dtDataVencimento <> DATA_NULA Then
                    'Busca com o número e vencimento aproximado
                    lErro = Comando_Executar(objBorderoRetCobr.alComando(12), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento BETWEEN ? AND ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? " & sFiltro, lNumIntDoc, lNumTitulo, DateAdd("d", -15, objDetRetCobr.dtDataVencimento), DateAdd("d", 15, objDetRetCobr.dtDataVencimento), STATUS_EXCLUIDO, STATUS_BAIXADO)
                    lComando = objBorderoRetCobr.alComando(12)
                Else
                    'Busca com o número do título
                    lErro = Comando_Executar(objBorderoRetCobr.alComando(13), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? " & sFiltro, lNumIntDoc, lNumTitulo, STATUS_EXCLUIDO, STATUS_BAIXADO)
                    lComando = objBorderoRetCobr.alComando(13)
                End If
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 32267
            
            lErro = Comando_BuscarPrimeiro(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32268
        
            If lErro = AD_SQL_SEM_DADOS Then
                objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
                gError 32269
            End If
        Else
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_BAIXADA
            gError 32269
        End If
    
    End If
    
    'Se chegou até aqui é porque de um jeito ou de outro achou pelo menos uma parcela aberta
    objParcelaRec.lNumIntDoc = lNumIntDoc

    'Lê a parcela a receber
    lErro = CF("ParcelaReceber_Le", objParcelaRec)
    If lErro <> SUCESSO And lErro <> 19147 Then gError 51683
    If lErro <> SUCESSO Then gError 51697
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32270
    
    'se há mais de uma parcela
    If lErro <> AD_SQL_SEM_DADOS Then
    
        If IsNumeric(sParcela) Then
            iNumParcela = CInt(sParcela)
        Else
            If sParcela >= "A" And sParcela <= "Z" Then
                iNumParcela = Asc(sParcela) - Asc("A") + 1
            Else
                If sParcela >= "a" And sParcela <= "z" Then
                    iNumParcela = Asc(sParcela) - Asc("a") + 1
                Else
                    iNumParcela = 0
                End If
            End If
        End If
    
        'se sei o # da parcela
        If iNumParcela <> 0 Then
            
            lErro = Comando_Executar(objBorderoRetCobr.alComando(10), "SELECT ParcelasRec.NumIntDoc FROM ParcelasRec, TitulosRec WHERE ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.NumTitulo = ? AND DataVencimento = ? AND NumParcela = ? AND ParcelasRec.Status <> ? AND ParcelasRec.Status <> ? " & sFiltro, lNumIntDoc, lNumTitulo, objDetRetCobr.dtDataVencimento, iNumParcela, STATUS_EXCLUIDO, STATUS_BAIXADO)
            If lErro <> AD_SQL_SUCESSO Then gError 32267
    
            lErro = Comando_BuscarPrimeiro(objBorderoRetCobr.alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32268
            
            If lErro = AD_SQL_SEM_DADOS Then
                   
                objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
                gError 32269
                
            End If
    
            objParcelaRec.lNumIntDoc = lNumIntDoc
        
            'Lê a parcela a receber
            lErro = CF("ParcelaReceber_Le", objParcelaRec)
            If lErro <> SUCESSO And lErro <> 19147 Then gError 51683
            If lErro <> SUCESSO Then gError 51697

            lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(10))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 32270
            
            'se há mais de uma parcela
            If lErro <> AD_SQL_SEM_DADOS Then
            
                '??? poderia abrir janela p/selecionar parcela passando as informacoes do registro
                objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
                gError 32271
                
            End If
            
        Else
        
            '??? poderia abrir janela p/selecionar parcela passando as informacoes do registro
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
            gError 32271
        
        End If
        
    End If
    
    'Testes adicionais
    If objDetRetCobr.iTipoOcorrencia = RETCOBR_DET_BAIXA Then
    
        'Se o valor recebido é maior que o saldo + 50% e a diferença é maior que 5 real
        If objParcelaRec.dSaldo <> 1 And (objDetRetCobr.dValorRecebido > (objParcelaRec.dSaldo * 1.5)) And Abs(objDetRetCobr.dValorRecebido - objParcelaRec.dSaldo) > 5 Then
    
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_VALOR_DIF
            gError 32272
    
        End If
            
    End If
    
    'O cobrador, se preenchido, tem que ser ao menos da mesma filial e banco
    If objParcelaRec.iCobrador <> COBRADOR_PROPRIA_EMPRESA And objBorderoRetCobr.iCobrador > COBRADOR_PROPRIA_EMPRESA Then
        
        objCobrador1.iCodigo = objParcelaRec.iCobrador
        
        lErro = CF("Cobrador_Le", objCobrador1)
        If lErro <> SUCESSO And lErro <> 19294 Then gError ERRO_SEM_MENSAGEM
    
        objCobrador2.iCodigo = objBorderoRetCobr.iCobrador
    
        lErro = CF("Cobrador_Le", objCobrador2)
        If lErro <> SUCESSO And lErro <> 19294 Then gError ERRO_SEM_MENSAGEM
        
        If objCobrador1.iFilialEmpresa <> objCobrador2.iFilialEmpresa Or objCobrador1.iCodBanco <> objCobrador2.iCodBanco Then
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
            gError 32269
        End If
    
    End If
    
    If objDetRetCobr.dtDataVencimento <> DATA_NULA Then

        'Se a diferença do vencimento ultrapassar 15 dias
        If Abs(DateDiff("d", objDetRetCobr.dtDataVencimento, objParcelaRec.dtDataVencimentoReal)) > 15 Then
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_VENC_DIF
            gError 32273
        
        End If

    End If
    
    RetCobr_Obtem_Parcela2 = SUCESSO

    Exit Function

Erro_RetCobr_Obtem_Parcela2:

    RetCobr_Obtem_Parcela2 = gErr

    Select Case gErr

        Case 32269
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_NENHUMA", gErr, tRegDetalhe.sSeuNumero)
            RetCobr_Obtem_Parcela2 = SUCESSO
        
        Case 32271
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_VARIAS", gErr, tRegDetalhe.sSeuNumero)
            RetCobr_Obtem_Parcela2 = SUCESSO
            
        Case 32272 'Valor não bate
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_VARIAS", gErr, tRegDetalhe.sSeuNumero)
            RetCobr_Obtem_Parcela2 = SUCESSO
            
        Case 32273 'Vencimento não bate
            'Call Rotina_Erro(vbOKOnly, "ERRO_PARCBORDRETCOBR_VARIAS", gErr, tRegDetalhe.sSeuNumero)
            RetCobr_Obtem_Parcela2 = SUCESSO
            
        Case 32267, 32268, 32270
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCREC_RETCOBR", gErr)
        
        Case 51683, 81769
        
        Case 51697
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", gErr, lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147254)
            
    End Select

    Exit Function

End Function

Function Processa_Registro_Detalhe_Geral(ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objCobrancaEletronica As ClassCobrancaEletronica, ByVal iTipoOcorrencia As Integer, ByVal iTipoSeuNumero As Integer, ByVal objDetRetCobr As ClassDetRetCobr, ByVal objBorderoRetCobrCart As ClassBorderoRetCobrCart, lComando As Long) As Long

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim objParcelaRec As New ClassParcelaReceber
Dim objInfoParcRec As ClassInfoParcRec
Dim objOcorrRetParcRec As New ClassOcorrRetParcRec
Dim objDetRetCobrAux As ClassDetRetCobr, bAchou As Boolean

On Error GoTo Erro_Processa_Registro_Detalhe_Geral

    objDetRetCobr.iTipoOcorrencia = iTipoOcorrencia
    
    Select Case iTipoOcorrencia
        
        'Quando ocorre baixa p/liquidacao
        Case RETCOBR_DET_BAIXA
               
            lErro = CF("RetCobr_Obtem_Parcela", objBorderoRetCobr, objDetRetCobr, objParcelaRec, iTipoSeuNumero, lComando)
            If lErro <> SUCESSO Then Error 32264
            'se houve erro
            If objDetRetCobr.iTipoCritica <> RETCOBR_CRITICA_SEM_ERRO Then
        
                'incluir ocorrencia de erro do retorno
                Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
        
            Else
        
                lErro = CF("RetCobr_Preenche_InfoParcRec", objParcelaRec, objInfoParcRec)
                '##############################
                'Alterado por Wagner
                'If lErro <> SUCESSO Then Error 51695
                
                'Se deu erro ao ler a Parcela ou o Titulo não processa e inclui erro
                If lErro <> SUCESSO Then
                
                    objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
                
                    'incluir ocorrencia de erro do retorno
                    Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                
                Else
                
                    'Busca outro pagamento integral para mesma parcela
                    bAchou = False
                    For Each objDetRetCobrAux In objBorderoRetCobr.colRetCobrErros
                        If objDetRetCobrAux.iTipoCritica = RETCOBR_CRITICA_LIQUIDACAO And objDetRetCobrAux.lNumIntParc = objDetRetCobr.lNumIntParc Then
                            bAchou = True
                        End If
                    Next
                    'Se está duplicado apenas adiciona o erro
                    If bAchou Then
                    
                        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PAGA_DUPLIC
                        
                        'incluir ocorrencia de erro do retorno
                        Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                   
                    Else
                    
                        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_LIQUIDACAO
                        'incluir ocorrencia de erro do retorno
                        Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                        
                        'Processa a parcela a receber
                        lErro = CF("Processa_ParcelaRec", objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, objInfoParcRec)
                        If lErro <> SUCESSO Then Error 51695
                                            
                    End If
                    
                End If
                '##############################
                              
            
            End If
            
        'cobranca de tarifas
        Case RETCOBR_DET_TARIFAS
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_TARIFAS
            Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
            
            lErro = CF("RetCobr_Obtem_Parcela", objBorderoRetCobr, objDetRetCobr, objParcelaRec, iTipoSeuNumero, lComando)
            '??? melhorar depois If lErro <> SUCESSO Then Error 32264
            
            'pode ter trocado na rotina acima p/sem_parc, baixada, etc
            If objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_TARIFAS Then
                lErro = CF("Processa_ParcelaRec", objParcelaRec, objBorderoRetCobr, objBorderoRetCobrCart, objDetRetCobr, objInfoParcRec)
                If lErro <> SUCESSO Then Error 51695
            End If
            
            objBorderoRetCobrCart.dValorTarifas = Arredonda_Moeda(objBorderoRetCobrCart.dValorTarifas + objDetRetCobr.dValorTarifa, 2)
            
            'Se o valor não costa como creditado na conta
            If objBorderoRetCobr.iVerificaCreditosTarifas = MARCADO Then
                If objDetRetCobr.dValorCreditado = 0 Then objBorderoRetCobrCart.dValorTarifasNaoCred = Arredonda_Moeda(objBorderoRetCobrCart.dValorTarifasNaoCred + objDetRetCobr.dValorTarifa, 2)
            End If
            
        'cobranca de tarifas
        Case RETCOBR_DET_CUSTAS
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_CUSTAS
            Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
            
            lErro = CF("RetCobr_Obtem_Parcela", objBorderoRetCobr, objDetRetCobr, objParcelaRec, iTipoSeuNumero, lComando)
            '??? melhorar depois If lErro <> SUCESSO Then Error 32264
            
            objBorderoRetCobrCart.dValorCustas = Arredonda_Moeda(objBorderoRetCobrCart.dValorCustas + objDetRetCobr.dValorTarifa, 2)
        
        Case RETCOBR_DET_REJEITADO
            
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_ENTRADA_REJEITADA
            Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
                        
        Case RETCOBR_DET_BAIXA_POR_PROTESTO
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_BAIXA_POR_PROTESTO
            Call objBorderoRetCobr.colRetCobrErros.Add(objDetRetCobr)
    
    End Select
    
    Set objCobrancaEletronica.objTelaAtualizacao.colRetCobrErros = objBorderoRetCobr.colRetCobrErros

    Processa_Registro_Detalhe_Geral = SUCESSO
     
    Exit Function
    
Erro_Processa_Registro_Detalhe_Geral:

    Processa_Registro_Detalhe_Geral = gErr
     
    Select Case gErr
          
        Case 32264, 51695 'Inserido por Wagner
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147255)
     
    End Select
     
    Exit Function

End Function

Function RetCobr_Obtem_Parcela3(ByVal objDetRetCobr As ClassDetRetCobr, ByVal objParcelaRec As ClassParcelaReceber) As Long

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec

On Error GoTo Erro_RetCobr_Obtem_Parcela3

    objOcorrRemParcRec.lNumIntDoc = StrParaLong(right(Trim(objDetRetCobr.sSeuNumero), 9))
    
    'Lê os dados da remessa em ocorremparcrec
    lErro = CF("OcorrRemParcRec_Le_NumIntOcorr", objOcorrRemParcRec)
    If lErro <> SUCESSO And lErro <> 51719 Then gError 51879
'    If lErro <> SUCESSO Then gError 51880
    If lErro <> SUCESSO Then
        objParcelaRec.lNumIntDoc = 0
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
        gError 51882
    End If

    objParcelaRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc
    
    'Lê a parcela a receber
    lErro = CF("ParcelaReceber_Le", objParcelaRec)
    If lErro <> SUCESSO And lErro <> 19147 Then gError 51881
    If lErro <> SUCESSO Then
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_SEM_PARC
        gError 51882
    End If
    
    RetCobr_Obtem_Parcela3 = SUCESSO
     
    Exit Function
    
Erro_RetCobr_Obtem_Parcela3:

    RetCobr_Obtem_Parcela3 = gErr
     
    Select Case gErr
          
        Case 51881, 51879
        
        Case 51880
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRREMPARCREC_NAO_CADASTRADA", gErr, objOcorrRemParcRec.lNumIntDoc)
            
        Case 51882
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", Err, objOcorrRemParcRec.lNumIntParc)
            RetCobr_Obtem_Parcela3 = SUCESSO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147256)
     
    End Select
     
    Exit Function

End Function

Function RetCobr_Obtem_Parcela4(ByVal objBorderoRetCobr As ClassBorderoRetCobr, ByVal objDetRetCobr As ClassDetRetCobr, objParcelaRec As ClassParcelaReceber) As Long
'preenche objParcelaRec a partir do "NossoNumero" ( cobrador/carteira/nossonumero ou cobrador/nossonumero )

Dim lErro As Long, lNumIntParc As Long, objCarteiraCobrador As ClassCarteiraCobrador, iLenNossoNum As Integer, sSQL As String
Dim dtDataVencimento As Date, lNumIntParcAnt As Long

On Error GoTo Erro_RetCobr_Obtem_Parcela4

    Call objBorderoRetCobr.CarteiraCobranca_Obter(objDetRetCobr.iCodCarteiraNoBanco, objCarteiraCobrador)
    
    If objCarteiraCobrador Is Nothing Then
        If objBorderoRetCobr.colCartCobr.Count <> 0 Then Set objCarteiraCobrador = objBorderoRetCobr.colCartCobr.Item(1)
    End If
    
    If Not (objCarteiraCobrador Is Nothing) And Len(Trim(objDetRetCobr.sNossoNumero)) > 0 Then
    
        lErro = CF("Banco_ObtemTamNossoNumero", objBorderoRetCobr.iBanco, iLenNossoNum)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If objBorderoRetCobr.iBanco = 104 Then
        
            If Len(objDetRetCobr.sNossoNumero) = 15 Then iLenNossoNum = 15
            If Len(objDetRetCobr.sNossoNumero) = 10 Then iLenNossoNum = 10
            
        End If
        
        'Busca pelo Nosso número, cobrador e carteira
        sSQL = "SELECT NumIntDoc FROM ParcelasRec WHERE {fn RIGHT(NumTitCobrador," & CStr(iLenNossoNum) & ")} = ? AND Cobrador = ? AND CarteiraCobranca = ?"
        lErro = Comando_Executar(objBorderoRetCobr.alComando(1), sSQL, lNumIntParc, Format(right(objDetRetCobr.sNossoNumero, iLenNossoNum), FormataCpoNum(0, iLenNossoNum)), objBorderoRetCobr.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
        If lErro <> AD_SQL_SUCESSO Then gError 130050
        
        lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130051
        
        'Se achou verifica se só tem uma
        If lErro = AD_SQL_SUCESSO Then
        
            lNumIntParcAnt = lNumIntParc
        
            lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130053
            
            If lErro = AD_SQL_SUCESSO Then
            
                objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
                gError 130042 'Várias parcelas encontradas
            
            End If
            
            lNumIntParc = lNumIntParcAnt
        
        Else
        'Se não achou
            
            'Repete a busca sem a carteira
            sSQL = "SELECT NumIntDoc FROM ParcelasRec WHERE {fn RIGHT(NumTitCobrador," & CStr(iLenNossoNum) & ")} = ? AND Cobrador = ?"
            lErro = Comando_Executar(objBorderoRetCobr.alComando(2), sSQL, lNumIntParc, Format(right(objDetRetCobr.sNossoNumero, iLenNossoNum), FormataCpoNum(0, iLenNossoNum)), objBorderoRetCobr.iCobrador)
            If lErro <> AD_SQL_SUCESSO Then gError 130052
            
            lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130053
        
            'Se achou verifica se só tem uma
            If lErro = AD_SQL_SUCESSO Then
            
                lNumIntParcAnt = lNumIntParc
            
                lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130053
                
                If lErro = AD_SQL_SUCESSO Then
                
                    objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_VARIAS_PARC
                    gError 130042 'Várias parcelas encontradas
                
                End If
                
                lNumIntParc = lNumIntParcAnt
            
            Else
            'Se não achou
            
                'Procura para ver se já está baixado
                If objDetRetCobr.dtDataVencimento = DATA_NULA Then
                    dtDataVencimento = Date
                Else
                    dtDataVencimento = objDetRetCobr.dtDataVencimento
                End If
            
                sSQL = "SELECT NumIntDoc FROM ParcelasRecBaixadas WHERE {fn RIGHT(NumTitCobrador," & CStr(iLenNossoNum) & ")} = ? AND Cobrador = ? AND DataVencimento BETWEEN ? AND ? "
                lErro = Comando_Executar(objBorderoRetCobr.alComando(3), sSQL, lNumIntParc, Format(right(objDetRetCobr.sNossoNumero, iLenNossoNum), FormataCpoNum(0, iLenNossoNum)), objBorderoRetCobr.iCobrador, DateAdd("d", -90, dtDataVencimento), DateAdd("d", 90, dtDataVencimento))
                If lErro <> AD_SQL_SUCESSO Then gError 130052
                
                lErro = Comando_BuscarProximo(objBorderoRetCobr.alComando(3))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 130053
                                
                If lErro = AD_SQL_SUCESSO Then
                
                    objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_BAIXADA
                    gError 130042 'Parcela encontrada, mas baixada
                    
                Else
                
                    gError 130041 'Parcela não encontrada
                
                End If
 
            End If
            
        End If
        
        objParcelaRec.lNumIntDoc = lNumIntParc
        
        'Lê a parcela a receber
        lErro = CF("ParcelaReceber_Le", objParcelaRec)
        If lErro <> SUCESSO And lErro <> 19147 Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 51882
    
    Else
    
        'Esse ero indica que é para continuar procurando
        gError 130041 '??? poderia ter um erro mais especifico mas as vezes vai numa carteira e volta em outra por alteracao do banco
        
    End If
    
    'Testes adicionais
    'Se está baixada
    If objParcelaRec.iStatus = STATUS_BAIXADO Then
    
        objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_BAIXADA
        gError 130042 'Parcela encontrada, mas baixada
    
    End If
       
    'Se é uma baixa testa o valor baixado
    If objDetRetCobr.iTipoOcorrencia = RETCOBR_DET_BAIXA Then
      
        'Se o valor recebido é maior que o saldo + 50% e a diferença é maior que 5 real
        If objParcelaRec.dSaldo <> 1 And objDetRetCobr.dValorRecebido > objParcelaRec.dSaldo * 1.5 And Abs(objDetRetCobr.dValorRecebido - objParcelaRec.dSaldo) > 5 Then
    
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_VALOR_DIF
            gError 130042
    
        End If
            
    End If
    
    'Se tem data de vencimento -> Testa contra o da parcela
    If objDetRetCobr.dtDataVencimento <> DATA_NULA Then

        'Se a diferença do vencimento ultrapassar 90 dias
        If Abs(DateDiff("d", objDetRetCobr.dtDataVencimento, objParcelaRec.dtDataVencimentoReal)) > 90 Then
        
            objDetRetCobr.iTipoCritica = RETCOBR_CRITICA_PARC_VENC_DIF
            gError 130042
        
        End If

    End If
    
    RetCobr_Obtem_Parcela4 = SUCESSO
     
    Exit Function
    
Erro_RetCobr_Obtem_Parcela4:

    RetCobr_Obtem_Parcela4 = gErr
     
    Select Case gErr
          
        Case 130041 'parcela nao encontrada
        
        Case 130042 'parcela encontrada, mas com erro
        
        Case 130050 To 130053
            Case Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAREC", gErr, lNumIntParc)
        
        Case 51882
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADA", gErr, lNumIntParc)
        
        Case ERRO_SEM_MENSAGEM
               
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 147257)
     
    End Select
     
    Exit Function

End Function
