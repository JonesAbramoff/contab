VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRGravaMiguez"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function TituloPagarRateio_Insere(ByVal objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim lComando As Long
Dim lCount As Long
Dim objTituloPagRateio As ClassTituloPagRateio
Dim ColRateio As Collection

On Error GoTo Erro_TituloPagarRateio_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129030

    If Not (objTituloPagar.objInfoUsu Is Nothing) Then

        Set ColRateio = objTituloPagar.objInfoUsu
        
        If objTituloPagar.lNumIntDoc <> 0 Then
         
            For Each objTituloPagRateio In ColRateio
                            
                lErro = Comando_Executar(lComando, "INSERT INTO TitulosPagRateio " & _
                         "( NumIntDocPag, Seq, NumIntDocRec, Cliente, Valor, Historico, Dossie, Referencia, Cobrar, Data_Ger_ND, ND) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
                         objTituloPagar.lNumIntDoc, objTituloPagRateio.iSeq, objTituloPagRateio.lNumIntDocRec, objTituloPagRateio.lCliente, objTituloPagRateio.dValor, objTituloPagRateio.lHistorico, objTituloPagRateio.sDossie, objTituloPagRateio.sReferencia, objTituloPagRateio.iCobrar, objTituloPagRateio.dtDataGerND, objTituloPagRateio.lND)
                If lErro <> AD_SQL_SUCESSO Then gError 129031
                        
            Next
             
        End If
         
    End If
        
    Call Comando_Fechar(lComando)
        
    TituloPagarRateio_Insere = SUCESSO
    
    Exit Function
    
Erro_TituloPagarRateio_Insere:

    TituloPagarRateio_Insere = gErr

    Select Case gErr
    
        Case 129030
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 129031
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function TituloPagarRateio_Exclui(ByVal objTituloPag As ClassTituloPagar) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDocRec As Integer

On Error GoTo Erro_TituloPagarRateio_Exclui
       
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 129038
    Next
          
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocRec FROM TitulosPagRateio WHERE NumIntDocPag = ? ", 0, lNumIntDocRec, objTituloPag.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 129039
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129094
    
    While lErro <> AD_SQL_SEM_DADOS
    
        If lNumIntDocRec <> 0 Then gError 129138

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 129055

        'Exclui Rateio
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TitulosPagRateio", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 129040

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129041

    Wend
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TituloPagarRateio_Exclui = SUCESSO
    
    Exit Function
    
Erro_TituloPagarRateio_Exclui:

    TituloPagarRateio_Exclui = gErr

    Select Case gErr
    
        Case 129038
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
              
       Case 129039, 129040, 129041, 129055, 129094
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)
        
       Case 129138
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSRECND_NUMINTDOC", gErr)
       
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NFFatPag_Grava(ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelas As ColParcelaPagar, ByVal objContabil As ClassContabil) As Long
'Grava Nota Fiscal Fatura a Pagar no BD (nível + alto)

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long, lNumIntNF As Long
Dim iFilialEmpresa As Integer
Dim objParcelaPagar As ClassParcelaPagar, dValorLivroAux As Double

On Error GoTo Erro_NFFatPag_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 18461
    
    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18462
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 18463
       
    'Verifica se o Título está cadastrado como um Título Baixado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa FROM TitulosPagBaixados WHERE Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ?", iFilialEmpresa, objTituloPagar.lFornecedor, objTituloPagar.iFilial, objTituloPagar.lNumTitulo, objTituloPagar.dtDataEmissao, objTituloPagar.sSiglaDocumento, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 18476
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18477
    
    'Se encontrou, erro
    If lErro = AD_SQL_SUCESSO Then gError 18478
    
    'Verifica se o Título a pagar possui um Adiantamento vinculado ao Fornecedor
    lErro = CF("TituloPagar_Verifica_Adiantamento", objTituloPagar.lFornecedor, objTituloPagar.iFilial)
    If lErro <> SUCESSO Then gError 75478
        
    'Grava no BD
    lErro = CF("NFFatPag_Grava_BD", objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 18470
    
    'Inserido por Wagner
    '########################
    lErro = TituloPagarRateio_Exclui(objTituloPagar)
    If lErro <> SUCESSO Then gError 129048
    
    lErro = TituloPagarRateio_Insere(objTituloPagar)
    If lErro <> SUCESSO Then gError 129049
    '########################
    
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'se o documento foi originado fora de CPR nao pode ter a contabilidade alterada em CPR
        lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloPagar.sSiglaDocumento, objTituloPagar.lNumIntDoc, lNumIntNF)
        If lErro <> SUCESSO And lErro <> 41542 Then gError 59224
        
        'se nao é um docto gerado fora do cpr
        If lErro <> SUCESSO Then
            
            dValorLivroAux = 0
                        
            For Each objParcelaPagar In colParcelas
            
                dValorLivroAux = dValorLivroAux + Round(objParcelaPagar.dValor, 2)
                
            Next
            
            'Grava os dados contábeis (contabilidade)
            lErro = objContabil.Contabil_Gravar_Registro(objTituloPagar.lNumIntDoc, objTituloPagar.lFornecedor, objTituloPagar.iFilial, DATA_NULA, , 0, , dValorLivroAux)
            If lErro <> SUCESSO Then gError 20473
    
        End If
        
    End If
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 18471
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    NFFatPag_Grava = SUCESSO
    
    Exit Function
    
Erro_NFFatPag_Grava:

    NFFatPag_Grava = gErr
    
    Select Case gErr
    
        Case 18461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 18462, 18463
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 18470, 20473, 59224, 75478
        
        Case 18471
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 18476, 18477
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFFATBAIXADA", gErr, objTituloPagar.lNumTitulo)
        
        Case 18478
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_BAIXADA_MODIFICACAO", gErr, objTituloPagar.lNumTitulo)
    
        '########
        Case 129048, 129049  'Tratado na função de origem
        '########
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Function NFFatPag_Exclui(ByVal objTituloPagar As ClassTituloPagar, ByVal objContabil As ClassContabil) As Long
'Executa a exclusão da Nota Fiscal Fatura passada por parâmetro

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lNumIntDoc As Long, lNumIntNF As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_NFFatPag_Exclui
    
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 18560

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 18561

    lComando1 = Comando_Abrir()
    If lComando = 0 Then gError 18576
    
    'Inserido por Wagner
    '########################
    lErro = TituloPagarRateio_Exclui(objTituloPagar)
    If lErro <> SUCESSO Then gError 129054
    '########################

    'Seleciona o Título na Tabela de Titulos a Pagar
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM TitulosPag WHERE NumIntDoc = ? AND FilialEmpresa = ? AND Status <> ? AND SiglaDocumento = ? ", 0, lNumIntDoc, objTituloPagar.lNumIntDoc, giFilialEmpresa, STATUS_EXCLUIDO, objTituloPagar.sSiglaDocumento)
    If lErro <> AD_SQL_SUCESSO Then gError 18562

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 18563
    
    If lErro = AD_SQL_SEM_DADOS Then gError 18564

    'Faz um "lock" no Título
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 18565

    'se o documento foi originado fora de CPR nao pode ser excluido em CPR
    lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloPagar.sSiglaDocumento, objTituloPagar.lNumIntDoc, lNumIntNF)
    If lErro <> SUCESSO And lErro <> 41542 Then gError 41543
    If lErro = SUCESSO Then
    
        objNFiscal.lNumIntDoc = lNumIntNF

        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 59450
        If lErro <> SUCESSO Then gError 59451
        If objNFiscal.iStatus <> STATUS_CANCELADO Then gError 18467

    End If
    
    'Exclui as parcelas referentes ao Título passado
    lErro = CF("ParcelasPagar_Exclui", objTituloPagar)
    If lErro <> SUCESSO Then gError 18566

    'Faz a exclusão lógica do Título
    lErro = Comando_ExecutarPos(lComando1, "UPDATE TitulosPag SET Status = ? ", lComando, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 18587

    'Exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
    lErro = objContabil.Contabil_Exclui(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO Then gError 20476

    'Finaliza transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 18574

    'Fecha comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    NFFatPag_Exclui = SUCESSO
    
    Exit Function
    
Erro_NFFatPag_Exclui:

    NFFatPag_Exclui = gErr
    
    Select Case gErr
    
        Case 18467
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAGAR_VINCULADA_NFISCAL", gErr, objTituloPagar.lNumTitulo)
            
        Case 18560
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 18561, 18576
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 18562, 18563
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
            
        Case 18564
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA", gErr, objTituloPagar.lNumIntDoc)
        
        Case 18565
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
        
        Case 18466, 18566, 18905, 20476, 41543, 59450, 59451
            
        Case 18574
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 18587
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
            
        '#########
        Case 129054 ' Tratado na função de origem
        '#########
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function TituloReceber_Grava(ByVal objTituloReceber As ClassTituloReceber, ByVal colComissaoEmissao As colComissao, ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal colcolDesconto As colcolDesconto, ByVal objContabil As ClassContabil) As Long
'Grava um Título à Receber

Dim lErro As Long
Dim lTransacao As Long, dValorLivroAux As Double
Dim lNumIntDoc As Long, lNumIntNF As Long

On Error GoTo Erro_TituloReceber_Grava

    'Inserido por Wagner
    '####################
    If objTituloReceber.sSiglaDocumento = "ND" Then
        lErro = TituloReceberRateio_Valida(objTituloReceber)
        If lErro <> SUCESSO And lErro <> 129135 And lErro <> 129137 Then gError 129116
    End If
    '####################

    'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 26225
        
    'Grava no BD
    lErro = CF("TituloReceber_Grava_BD", objTituloReceber, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto)
    If lErro <> SUCESSO Then gError 26239
    
    'Inserido por Wagner
    '####################
    If objTituloReceber.sSiglaDocumento = "ND" Then
        lErro = TituloReceberRateio_Altera(objTituloReceber)
        If lErro <> SUCESSO Then gError 129117
    End If
    '####################

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'se o documento foi originado fora de CPR nao pode ter a contabilidade alterada em CPR
        lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloReceber.sSiglaDocumento, objTituloReceber.lNumIntDoc, lNumIntNF)
        If lErro <> SUCESSO And lErro <> 41542 Then gError 59225
        
        'se nao é um docto gerado fora do cpr
        If lErro <> SUCESSO Then
            
            dValorLivroAux = Round(objTituloReceber.dValor - objTituloReceber.dValorIRRF - IIf(objTituloReceber.iINSSRetido <> 0, objTituloReceber.dValorINSS, 0), 2)
            
            'Grava os dados contábeis (contabilidade)
            lErro = objContabil.Contabil_Gravar_Registro(objTituloReceber.lNumIntDoc, objTituloReceber.lCliente, objTituloReceber.iFilial, DATA_NULA, , 0, , -dValorLivroAux)
            If lErro <> SUCESSO Then gError 20523

        End If
        
    End If
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 26240
    
    TituloReceber_Grava = SUCESSO
    
    Exit Function

Erro_TituloReceber_Grava:

    TituloReceber_Grava = gErr
    
    Select Case gErr
    
        Case 20523, 26239, 59225
    
        Case 26225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 26240
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 129116, 129117
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$)
            
    End Select

    'Fecha a transação
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function TituloReceber_Exclui(ByVal objTituloReceber As ClassTituloReceber, ByVal objContabil As ClassContabil) As Long
'Exclui o Título à Receber do Banco de Dados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TituloReceber_Exclui

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 26443

    lErro = CF("TituloReceber_Exclui_EmTransacao", objTituloReceber, objContabil)
    If lErro <> SUCESSO Then gError 62137
    
    'Inserido por Wagner
    '####################
    lErro = TituloReceberRateio_Exclui(objTituloReceber)
    If lErro <> SUCESSO Then gError 129137
    '####################
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 26453
    
    TituloReceber_Exclui = SUCESSO
    
    Exit Function

Erro_TituloReceber_Exclui:

    TituloReceber_Exclui = gErr
    
    Select Case gErr
        
        Case 26443
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 62137, 129137
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
            
    End Select

    'Encerra a transação e os comandos
    Call Transacao_Rollback
    
    Exit Function

End Function

Private Function TituloReceberRateio_Valida(ByVal objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim lComando As Long
Dim tTituloPagRateio As typeTituloPagRateio
Dim objTituloPagRateio As ClassTituloPagRateio
Dim dValor As Double
Dim iCont As Integer

On Error GoTo Erro_TituloReceberRateio_Valida

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 129129
    
    'Seleciona Rateios vinculados ao Título
    lErro = Comando_Executar(lComando, "SELECT NumIntDocRec, Valor, Cliente, Cobrar, ND FROM TitulosPagRateio WHERE ND = ? ", tTituloPagRateio.lNumIntDocRec, tTituloPagRateio.dValor, tTituloPagRateio.lCliente, tTituloPagRateio.iCobrar, tTituloPagRateio.lND, objTituloReceber.lNumTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 129130

    'Tenta ler o primeiro item do Rateio
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129131
        
    dValor = 0
    iCont = 0
        
    Do While lErro <> AD_SQL_SEM_DADOS
    
        If tTituloPagRateio.lCliente <> objTituloReceber.lCliente Then gError 129133
    
        If tTituloPagRateio.lNumIntDocRec <> 0 Then gError 129134
    
        dValor = dValor + tTituloPagRateio.dValor
        
        iCont = iCont + 1
    
        'Vai para o próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129132
    
    Loop
       
    If iCont = 0 Then gError 129137
    
    If dValor <> objTituloReceber.dValor Then gError 129135
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    TituloReceberRateio_Valida = SUCESSO
    
    Exit Function
    
Erro_TituloReceberRateio_Valida:

    TituloReceberRateio_Valida = gErr
    
    Select Case gErr
    
        Case 129129
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 129130 To 129132
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)
        
        Case 129133
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSRECND_CLIENTE", gErr)
        
        Case 129134
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOSRECND_NUMINTDOC", gErr)
                
        Case 129135
            Call Rotina_Aviso(vbOKOnly, "AVISO_TITULOSRECND_VALOR", gErr)
        
        Case 129137
            Call Rotina_Aviso(vbOKOnly, "AVISO_TITULOSRECND_ND", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function TituloReceberRateio_Exclui(ByVal objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_TituloReceberRateio_Exclui
       
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 129117
    Next
          
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Seq FROM TitulosPagRateio WHERE ND = ? ", 0, iSeq, objTituloReceber.lNumTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 129118
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129119
    
    While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 129120

        'Exclui Rateio
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPagRateio SET NumIntDocRec = ? ", alComando(1), 0)
        If lErro <> AD_SQL_SUCESSO Then gError 129121

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129122

    Wend
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TituloReceberRateio_Exclui = SUCESSO
    
    Exit Function
    
Erro_TituloReceberRateio_Exclui:

    TituloReceberRateio_Exclui = gErr

    Select Case gErr
    
        Case 129117
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
              
       Case 129118 To 129122
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)
        
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function TituloReceberRateio_Altera(ByVal objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iSeq As Integer

On Error GoTo Erro_TituloReceberRateio_Altera
       
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 129123
    Next
          
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Seq FROM TitulosPagRateio WHERE ND = ? ", 0, iSeq, objTituloReceber.lNumTitulo)
    If lErro <> AD_SQL_SUCESSO Then gError 129124
        
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129125
    
    While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 129126

        'Exclui Rateio
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPagRateio SET NumIntDocRec = ? ", alComando(1), objTituloReceber.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 129127

        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 129128

    Wend
        
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TituloReceberRateio_Altera = SUCESSO
    
    Exit Function
    
Erro_TituloReceberRateio_Altera:

    TituloReceberRateio_Altera = gErr

    Select Case gErr
    
        Case 129123
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
              
       Case 129124 To 129128
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGRATEIO", gErr)
        
       Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)
            
    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function TituloPagarRateio_Grava(ByVal objTituloPagar As ClassTituloPagar) As Long
'Grava Nota Fiscal Fatura a Pagar no BD (nível + alto)

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TituloPagarRateio_Grava

    'Abre a transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 132287
    
    lErro = TituloPagarRateio_Exclui(objTituloPagar)
    If lErro <> SUCESSO Then gError 132288
    
    lErro = TituloPagarRateio_Insere(objTituloPagar)
    If lErro <> SUCESSO Then gError 132289
    
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 132290
    
    TituloPagarRateio_Grava = SUCESSO
    
    Exit Function
    
Erro_TituloPagarRateio_Grava:

    TituloPagarRateio_Grava = gErr
    
    Select Case gErr
    
        Case 132287
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 132288, 132289 'Tratado na função de origem
    
        Case 132290
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
    
    Call Transacao_Rollback
    
    Exit Function

End Function

Public Function ParcelasReceber_Altera_CamposAlteraveis(ByVal colParcelaReceber As colParcelaReceber, ByVal iNumParcelas As Integer, ByVal bOrigemNF As Boolean) As Long
'Altera os campos Alteráveis das primeiras iNumParcelas parcelas da colecao
'As parcelas maiores que iNumParcelas, se houverem, serao novas e portanto serao incluidas e nao alteradas
'bOrigemNF indica se a alteracao está sendo feita da tela do FAT ou do CR. Se a origem for Nota fiscal nao poderao ser alterados os valores das parcelas

Dim lErro As Long, objParcRec As ClassParcelaReceber
Dim lComando As Long, dSaldo As Double, dValor As Double
Dim lComando1 As Long, objCobrador As New ClassCobrador, dSaldoChequePre As Double
Dim iIndice As Integer, iCobrador As Integer, iCarteiraCobranca As Integer
Dim lNumIntDoc As Long, dtDataVencimento As Date, dtDataVencimentoReal As Date
Dim objCarteiraCobrador As New ClassCarteiraCobrador, iNumParcela As Integer, bAlterouValor As Boolean
Dim dtDataPrevisao As Date
Dim bAlterouDataPrevisao As Boolean
Dim objParcRecMgz As ClassParcRecMgz

On Error GoTo Erro_ParcelasReceber_Altera_CamposAlteraveis

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26294
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26295
        
    '##################################################################
    'Alterado por Wagner
    'Seleciona as Parcelas a alterar
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumParcela, DataVencimento, DataVencimentoReal, Cobrador, CarteiraCobranca, Saldo, Valor, DataPrevisao FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? AND NumParcela <= ? ORDER BY NumParcela", 0, lNumIntDoc, iNumParcela, dtDataVencimento, dtDataVencimentoReal, iCobrador, iCarteiraCobranca, dSaldo, dValor, dtDataPrevisao, colParcelaReceber(1).lNumIntTitulo, STATUS_EXCLUIDO, iNumParcelas)
    If lErro <> AD_SQL_SUCESSO Then gError 26296
    '##################################################################
    
    'Loop de alteração das Parcelas
    For iIndice = 1 To iNumParcelas
        
        Set objParcRec = colParcelaReceber(iIndice)
        
        'Tenta ler a Parcela de número iIndice
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 26297
        If lErro = AD_SQL_SEM_DADOS Then gError 26298
        
        bAlterouValor = Abs(objParcRec.dValor - dValor) > DELTA_VALORMONETARIO
        
        'se o vencimento foi alterado
        '##################################
        'Alterado por Wagner
        bAlterouDataPrevisao = False
        If Not (objParcRec.objInfoUsu Is Nothing) Then
            Set objParcRecMgz = objParcRec.objInfoUsu
            
            If objParcRecMgz.dtDataPrevisao <> dtDataPrevisao Then bAlterouDataPrevisao = True

        End If
        
        
        If bAlterouDataPrevisao Or dtDataVencimento <> objParcRec.dtDataVencimento Or dtDataVencimentoReal <> objParcRec.dtDataVencimentoReal Or bAlterouValor Then
        '##################################

            'verifica se a parcela está em cobranca eletronica
            If iCobrador <> COBRADOR_PROPRIA_EMPRESA Then
            
                objCobrador.iCodigo = iCobrador
                lErro = CF("Cobrador_Le", objCobrador)
                If lErro <> SUCESSO And lErro <> 19294 Then gError 59200
                If lErro <> SUCESSO Then gError 59201
            
                If objCobrador.iCobrancaEletronica = 1 Then gError 59202
                
            End If

            'se o valor da parcela foi alterado
            If bAlterouValor Then
            
                If bOrigemNF Then gError 81881
                
                objParcRec.dSaldo = Round(objParcRec.dValor - (dValor - dSaldo), 2)
                
                'se o saldo que está no bd é diferente do valor da parcela é sinal de que há alguma baixa nao cancelada
                If Abs(dValor - dSaldo) > DELTA_VALORMONETARIO Then
                
                    'Se eu permitir alterar parcela com baixa posso ter que
                    'tratar o caso do valor ser reduzido de forma que a parcela
                    'fique baixada e indo adiante posso ter que baixar o titulo.
                    'para evitar ter que criar este codigo...
                    gError 81879
                                        
                End If
                
                'se o valor da parcela diminuiu
                If (objParcRec.dValor - dValor) < -DELTA_VALORMONETARIO Then
                
                    lErro = CF("ParcelaRec_ObtemSaldoChequePre", objParcRec.lNumIntDoc, dSaldoChequePre)
                    If lErro <> SUCESSO Then gError 81888
                    
                    'verifica se o saldo de cheques pre-datados vinculado à parcela é menor que o saldo da parcela
                    If (dSaldoChequePre - objParcRec.dSaldo) > DELTA_VALORMONETARIO Then gError 81889
                
                End If
                    
                'atualiza saldo da carteira de cobranca
                objCarteiraCobrador.iCobrador = iCobrador
                objCarteiraCobrador.iCodCarteiraCobranca = iCarteiraCobranca
                lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, 0, objParcRec.dSaldo - dSaldo)
                If lErro <> SUCESSO Then gError 81880
        
            End If
            
            If bAlterouValor Then
                        
                '####################################################
                'Alterado por Wagner
                'Altera os campos alteráveis
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET DataVencimento = ?, DataVencimentoReal =?, Valor = ?, Saldo =?, DataPrevisao = ?", lComando, objParcRec.dtDataVencimento, objParcRec.dtDataVencimentoReal, objParcRec.dValor, objParcRec.dSaldo, objParcRecMgz.dtDataPrevisao)
                If lErro <> AD_SQL_SUCESSO Then gError 26299
                '####################################################
                
            Else
            
                '####################################################
                'Alterado por Wagner
                'Altera os campos alteráveis
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET DataVencimento = ?, DataVencimentoReal =?, DataPrevisao = ?", lComando, objParcRec.dtDataVencimento, objParcRec.dtDataVencimentoReal, objParcRecMgz.dtDataPrevisao)
                If lErro <> AD_SQL_SUCESSO Then gError 26299
                '####################################################
            
            End If
        
        End If
        
    Next

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasReceber_Altera_CamposAlteraveis = SUCESSO
    
    Exit Function
    
Erro_ParcelasReceber_Altera_CamposAlteraveis:
    
    ParcelasReceber_Altera_CamposAlteraveis = gErr
    
    Select Case gErr
    
        Case 26294, 26295
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 26296, 26297
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr, colParcelaReceber(1).lNumIntTitulo)
       
        Case 26298
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_RECEBER_NAO_CADASTRADA", gErr, colParcelaReceber(1).lNumIntTitulo, iIndice)
        
        Case 26299
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", gErr, colParcelaReceber(1).lNumIntTitulo, iIndice)
        
        Case 59200, 81880, 81888
        
        Case 59201
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", gErr, iCobrador)
            
        Case 59202
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERE_VCTO_INSTR_COB_ELETR", gErr)
        
        Case 81879
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COM_BAIXA2", gErr, iNumParcela)
        
        Case 81881
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_ORIGEMNF_ALTVAL", gErr)
        
        Case 81889
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_SALDO_CHEQUEPRE", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    'fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function ParcelasReceber_Insere(ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, Optional ByVal iNumPrimeiraParcela As Integer = 1) As Long
'Insere as Parcelas de um Titulo passadas por parâmetro
'CHAMADA EM TRANSAÇÃO
'iNumPrimeiraParcela é util na alteracao de titulo onde apenas as parcelas novas sao inseridas

Dim lErro As Long
Dim iIndice As Integer
Dim lComando As Long, objCarteiraCobrador As New ClassCarteiraCobrador
Dim lNumIntDoc As Long
Dim lNumIntCom As Long
'Dim dValorCarteira As Double, iNumParcsCarteira As Integer
'Dim dValorCarteiraChequePre As Double, iNumParcsCarteiraChequePre As Integer
'Dim objChequePre As ClassChequePre
Dim colDesconto As colDesconto
Dim objComissao As ClassComissao, colComissao As New colComissao
Dim objParcRecMgz As ClassParcRecMgz
Dim dtDataPrevisao As Date

On Error GoTo Erro_ParcelasReceber_Insere

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26356
    
    'Gera número interno automático da 1a parcela e reserva numeros para outras parcelas, se houverem
    lErro = CF("ParcelaReceber_Automatico", lNumIntDoc, colParcelaReceber.Count - iNumPrimeiraParcela + 1)
    If lErro <> SUCESSO Then Error 26359

    'Para cada Parcela
    For iIndice = iNumPrimeiraParcela To colParcelaReceber.Count
    
        colParcelaReceber(iIndice).lNumIntDoc = lNumIntDoc
    
        'Armazena iProxSeqBaixa e Saldo
        colParcelaReceber(iIndice).iProxSeqBaixa = 1
        colParcelaReceber(iIndice).iProxSeqOcorr = 1
        colParcelaReceber(iIndice).dSaldo = colParcelaReceber(iIndice).dValor
        
        'Inicio da Alteracao Daniel em 10/10/2001
        If colParcelaReceber(iIndice).iCobrador = 0 Then
            objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA
        Else
            objCarteiraCobrador.iCobrador = colParcelaReceber(iIndice).iCobrador
        End If
        
        If colParcelaReceber(iIndice).iCarteiraCobranca = 0 Then
            objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CARTEIRA
        Else
            objCarteiraCobrador.iCodCarteiraCobranca = colParcelaReceber(iIndice).iCarteiraCobranca
        End If
        'Fim da Alteracao Daniel em 10/10/2001
        
        For Each objComissao In colcolComissao(iIndice).colComissao
            
            objComissao.iTipoTitulo = TIPO_PARCELA
            objComissao.lNumIntDoc = lNumIntDoc
            objComissao.iStatus = STATUS_PENDENTE
            objComissao.dtDataBaixa = DATA_NULA
            objComissao.dtDataGeracao = DATA_NULA
        
            colComissao.AddObj objComissao
            
        Next

'        dValorCarteira = dValorCarteira + colParcelaReceber(iIndice).dSaldo
'        iNumParcsCarteira = iNumParcsCarteira + 1
        
        '#######################################
        'Alterado por Wagner
        'Insere a Parcela no BD
        dtDataPrevisao = DATA_NULA
        
        If Not (colParcelaReceber(iIndice).objInfoUsu Is Nothing) Then
        
            Set objParcRecMgz = colParcelaReceber(iIndice).objInfoUsu
            
            dtDataPrevisao = objParcRecMgz.dtDataPrevisao
        
        End If
        
        With colParcelaReceber(iIndice)
            lErro = Comando_Executar(lComando, "INSERT INTO ParcelasRec (NumIntDoc,NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Saldo,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo, Desconto1Ate, Desconto1Valor,Desconto2Codigo, Desconto2Ate, Desconto2Valor,Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, DataPrevisao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lNumIntDoc, .lNumIntTitulo, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca, .sNumTitCobrador, .iProxSeqBaixa, .iProxSeqOcorr, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .iAceite, .iDescontada, dtDataPrevisao)
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 26360
        '#######################################

        lNumIntDoc = lNumIntDoc + 1
        
        lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, 1, colParcelaReceber(iIndice).dSaldo)
        If lErro <> SUCESSO Then Error 59274
    
    Next

'    'atualiza o saldo da carteira "em carteira"
'    If iNumParcsCarteira <> 0 Then
'
'        objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CARTEIRA
'
'        lErro = CF("CarteiraCobrador_AtualizaSaldo",objCarteiraCobrador, iNumParcsCarteira, dValorCarteira)
'        If lErro <> SUCESSO Then Error 59274
'
'    End If
        
'    'atualiza o saldo da carteira "em cheque pre datado"
'    If iNumParcsCarteiraChequePre <> 0 Then
'
'        objCarteiraCobrador.iCodCarteiraCobranca = CARTEIRA_CHEQUEPRE
'
'        lErro = CF("CarteiraCobrador_AtualizaSaldo",objCarteiraCobrador, iNumParcsCarteiraChequePre, dValorCarteiraChequePre)
'        If lErro <> SUCESSO Then Error 59274
'
'    End If
        
    'Grava as Comissões das Parcelas
    lErro = CF("Comissoes_Grava", colComissao)
    If lErro <> SUCESSO Then Error 26361
    
    'Fecha comando
    Call Comando_Fechar(lComando)
    
    ParcelasReceber_Insere = SUCESSO
    
    Exit Function

Erro_ParcelasReceber_Insere:

    ParcelasReceber_Insere = Err
    
    Select Case Err
    
        Case 26356
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26357, 26358, 26359, 59274, 59275 'Tratado na rotina chamada
        
        Case 26360
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", Err, colParcelaReceber(iIndice).lNumIntTitulo, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function
