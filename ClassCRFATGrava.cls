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

Const MOTIVO_PERDA = 4

Function Comissoes_Altera(ByVal colComissao As colComissao, ByVal iTipoTitulo As Integer, ByVal lNumIntDoc As Long) As Long
'Altera no BD as comissoes passadas em colComissao

Dim lErro As Long

On Error GoTo Erro_Comissoes_Altera

    'Exclui as Comissoes já cadastradas, inclusive as baixadas
    lErro = CF("Comissoes_Exclui", iTipoTitulo, lNumIntDoc, 1)
    If lErro <> SUCESSO Then Error 26276
        
    'Verifica se existe comissao a ser alterada
    If colComissao.Count > 0 Then

        'Insere as novas Comissões
        lErro = CF("Comissoes_Grava", colComissao)
        If lErro <> SUCESSO Then Error 26277

    End If

    Comissoes_Altera = SUCESSO

    Exit Function

Erro_Comissoes_Altera:

    Comissoes_Altera = Err

    Select Case Err

        Case 26276, 26277

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148831)

    End Select

    Exit Function

End Function

Function Comissoes_Grava(ByVal colComissao As colComissao) As Long
'Grava as Comissões passadas em colComissao

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long, iUltVendedor As Integer
Dim objComissao As ClassComissao
Dim iCodigo As Integer
Dim lNumIntCom As Long

On Error GoTo Erro_Comissoes_Grava

    If colComissao.Count <> 0 Then
    
        lComando = 0
        lComando1 = 0
    
        'Abre os comandos
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 26284
    
        lComando1 = Comando_Abrir()
        If lComando1 = 0 Then Error 26285
    
        'Gera o próximo número interno para a nova Comissão
        lErro = CF("Comissao_Automatico", lNumIntCom, colComissao.Count)
        If lErro <> SUCESSO Then Error 26291
    
        'Loop de gravação
        For Each objComissao In colComissao
    
            'para evitar ficar lockando varias vezes o mesmo vendedor
            If (iUltVendedor <> objComissao.iCodVendedor) Then
            
                lErro = CF("Vendedor_Lock", lComando, objComissao.iCodVendedor)
                If lErro <> SUCESSO Then Error 59073
                
                iUltVendedor = objComissao.iCodVendedor
            
            End If
            
            'Insere a comissao no BD
            lErro = Comando_Executar(lComando1, "INSERT INTO Comissoes (NumIntCom, FilialEmpresa, Status, TipoTitulo, NumIntDoc, CodVendedor, DataGeracao, DataBaixa, Percentual, ValorBase, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?)", lNumIntCom, giFilialEmpresa, objComissao.iStatus, objComissao.iTipoTitulo, objComissao.lNumIntDoc, objComissao.iCodVendedor, objComissao.dtDataGeracao, objComissao.dtDataBaixa, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor)
            If lErro <> AD_SQL_SUCESSO Then Error 26292
    
            lNumIntCom = lNumIntCom + 1
            
        Next
    
        'Fecha os comandos
        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando1)

    End If
    
    Comissoes_Grava = SUCESSO

    Exit Function

Erro_Comissoes_Grava:

    Comissoes_Grava = Err

    Select Case Err

        Case 26284, 26285
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err, Error$)

        Case 26291, 59073

        Case 26292
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOES", Err, objComissao.iTipoTitulo, objComissao.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148832)

    End Select

    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function Comissao_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Gera o próximo Número Interno de uma Comissao
'AVISO: ESSA ROTINA DEVE SER CHAMADA DE DENTRO DE UMA TRANSACAO

    Comissao_Automatico = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_COMISSAO", lNumIntDoc, iQtde)
    
End Function

Function Estado_Grava(ByVal objEstado As ClassEstado) As Long
'Grava as alíquotas do Estado no Banco de Dados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sSigla As String

On Error GoTo Erro_Estado_Grava

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28502

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 28503

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28504

    sSigla = String(STRING_ESTADO_SIGLA, 0)

    'Pesquisa no BD o Estado em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Sigla FROM Estados WHERE Sigla = ?", 0, sSigla, objEstado.sSigla)
    If lErro <> AD_SQL_SUCESSO Then Error 28505

    'Lê os dados do estado, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28506

    'Não encontrou o Estado ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 28507

    'Encontrou o Estado ==> atualiza
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Estados SET ICMSAliquotaInterna = ?, ICMSAliquotaExportacao = ?, ICMSAliquotaImportacao = ?", lComando, objEstado.dICMSAliquotaInterna, objEstado.dICMSAliquotaExportacao, objEstado.dICMSAliquotaImportacao)
    If lErro <> AD_SQL_SUCESSO Then Error 28508

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28509

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Estado_Grava = SUCESSO

    Exit Function

Erro_Estado_Grava:

    Estado_Grava = Err

    Select Case Err

        Case 28502, 28503
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28504
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28505, 28506
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTADOS1", Err, objEstado.sSigla)

        Case 28507
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, objEstado.sSigla)

        Case 28508
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_ESTADOS", Err, objEstado.sSigla)

        Case 28509
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148833)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

''''Private Function CarteirasCobrador_Grava(objCarteiraCobrador As ClassCarteiraCobrador) As Long
'''''Grava Carteira de cobranca no Banco de Dados
'''''NAO ALTERA DADOS ESTATISTICOS
'''''Chamada dentro de TRANSAÇÃO
''''
''''Dim lComando1 As Long, lComando2 As Long, lComando3 As Long
''''Dim lErro As Long, sContaContabil As String, sContaDuplDescontadas As String
''''Dim iStatus As Integer
''''Dim iCodigo As Integer
''''Dim objPlanoConta As New ClassPlanoConta
''''
''''On Error GoTo Erro_CarteirasCobrador_Grava
''''
''''    lComando1 = 0
''''    lComando2 = 0
''''    lComando3 = 0
''''
''''    lComando1 = Comando_Abrir()
''''    If lComando1 = 0 Then Error 23482
''''
''''    lComando2 = Comando_Abrir()
''''    If lComando2 = 0 Then Error 23483
''''
''''    lComando3 = Comando_Abrir()
''''    If lComando3 = 0 Then Error 41646
''''
''''    sContaContabil = String(STRING_CONTA, 0)
''''    sContaDuplDescontadas = String(STRING_CONTA, 0)
''''
''''    'Pesquisa no Banco de Dados a Carteira em questão
''''    lErro = Comando_ExecutarPos(lComando1, "SELECT ContaContabil,ContaDuplDescontadas FROM CarteirasCobrador WHERE Cobrador = ? And CodCarteiraCobranca = ?", 0, sContaContabil, sContaDuplDescontadas, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
''''    If lErro <> AD_SQL_SUCESSO Then Error 23485
''''
''''    'Lê os dados da Carteira, se estiver no Banco de Dados
''''    lErro = Comando_BuscarPrimeiro(lComando1)
''''    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23486
''''
''''    'Encontrou a carteira, então isto é uma alteração
''''    If lErro = AD_SQL_SUCESSO Then
''''
''''        'obs.: nao atualizar dados estatisticos pois a tela pode estar defasada em relacao ao bd
''''
''''        'se trocou alguma das ctas p/valor <> "" verificar se é uma cta valida
''''        If Len(Trim(objCarteiraCobrador.sContaContabil)) > 0 And objCarteiraCobrador.sContaContabil <> sContaContabil Then
''''
''''            objPlanoConta.sConta = objCarteiraCobrador.sContaContabil
''''            lErro = CF("PlanoConta_Le_Conta1",objCarteiraCobrador.sContaContabil, objPlanoConta)
''''            If lErro <> SUCESSO And lErro <> 6030 Then Error 42906
''''            If lErro = 6030 Then Error 42907
''''
''''        End If
''''
''''        If Len(Trim(objCarteiraCobrador.sContaDuplDescontadas)) > 0 And objCarteiraCobrador.sContaDuplDescontadas <> sContaDuplDescontadas Then
''''
''''            objPlanoConta.sConta = objCarteiraCobrador.sContaDuplDescontadas
''''            lErro = CF("PlanoConta_Le_Conta1",objCarteiraCobrador.sContaDuplDescontadas, objPlanoConta)
''''            If lErro <> SUCESSO And lErro <> 6030 Then Error 42908
''''            If lErro = 6030 Then Error 42909
''''
''''        End If
''''
''''        'Atualiza os dados da Carteira
''''        lErro = Comando_ExecutarPos(lComando2, "UPDATE CarteirasCobrador SET Desativada=?, ContaContabil=?, ContaDuplDescontadas=?, DiasDeRetencao=?, TaxaCobranca=?, TaxaDesconto=? ,FaixaNossoNumeroInicial=?, FaixaNossoNumeroFinal=?, FaixaNossoNumeroProx=?, NomeNoBanco=?, CodCarteiraNoBanco=?, NumCarteiraNoBanco=?, ImprimeBoleta = ?,ComRegistro=?,GeraNossoNumero=?", lComando1, objCarteiraCobrador.iDesativada, objCarteiraCobrador.sContaContabil, _
''''        objCarteiraCobrador.sContaDuplDescontadas, objCarteiraCobrador.iDiasDeRetencao, objCarteiraCobrador.dTaxaCobranca, objCarteiraCobrador.dTaxaDesconto, objCarteiraCobrador.sFaixaNossoNumeroInicial, objCarteiraCobrador.sFaixaNossoNumeroFinal, objCarteiraCobrador.sFaixaNossoNumeroProx, objCarteiraCobrador.sNomeNoBanco, objCarteiraCobrador.sCodCarteiraNoBanco, objCarteiraCobrador.iNumCarteiraNoBanco, objCarteiraCobrador.iImprimeBoleta, objCarteiraCobrador.iComRegistro, objCarteiraCobrador.iGeraNossoNumero)
''''        If lErro <> AD_SQL_SUCESSO Then Error 23487
''''
''''    'Não encontrou a carteira, então isto é uma inserção
''''    Else
''''
''''        'verificar se ctas <> "" são validas
''''        If Len(Trim(objCarteiraCobrador.sContaContabil)) > 0 Then
''''            objPlanoConta.sConta = objCarteiraCobrador.sContaContabil
''''            lErro = CF("PlanoConta_Le_Conta1",objCarteiraCobrador.sContaContabil, objPlanoConta)
''''            If lErro <> SUCESSO And lErro <> 6030 Then Error 42902
''''            If lErro = 6030 Then Error 42903
''''
''''        End If
''''
''''        If Len(Trim(objCarteiraCobrador.sContaDuplDescontadas)) > 0 Then
''''            objPlanoConta.sConta = objCarteiraCobrador.sContaDuplDescontadas
''''            lErro = CF("PlanoConta_Le_Conta1",objCarteiraCobrador.sContaDuplDescontadas, objPlanoConta)
''''            If lErro <> SUCESSO And lErro <> 6030 Then Error 42904
''''            If lErro = 6030 Then Error 42905
''''
''''        End If
''''
''''        lErro = CF("CarteiraDeCobranca_Lock",objCarteiraCobrador.iCodCarteiraCobranca, lComando3)
''''        If lErro <> SUCESSO And lErro <> 41642 Then Error 41644
''''        If lErro <> SUCESSO Then Error 41645
''''
''''        'Insere a carteira
''''        lErro = Comando_Executar(lComando2, "INSERT INTO CarteirasCobrador(Cobrador,CodCarteiraCobranca,Desativada,ContaContabil,ContaDuplDescontadas,DiasDeRetencao,TaxaCobranca,TaxaDesconto,FaixaNossoNumeroInicial,FaixaNossoNumeroFinal,FaixaNossoNumeroProx,NomeNoBanco,CodCarteiraNoBanco,NumCarteiraNoBanco,QuantidadeAtual,QuantidadeAtualBanco,SaldoAtual,SaldoAtualBanco, ImprimeBoleta,ComRegistro,GeraNossoNumero) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objCarteiraCobrador.iCobrador, _
''''        objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iDesativada, objCarteiraCobrador.sContaContabil, objCarteiraCobrador.sContaDuplDescontadas, objCarteiraCobrador.iDiasDeRetencao, objCarteiraCobrador.dTaxaCobranca, objCarteiraCobrador.dTaxaDesconto, objCarteiraCobrador.sFaixaNossoNumeroInicial, objCarteiraCobrador.sFaixaNossoNumeroFinal, objCarteiraCobrador.sFaixaNossoNumeroProx, objCarteiraCobrador.sNomeNoBanco, objCarteiraCobrador.sCodCarteiraNoBanco, _
''''        objCarteiraCobrador.iNumCarteiraNoBanco, objCarteiraCobrador.lQuantidadeAtual, objCarteiraCobrador.lQuantidadeAtualBanco, objCarteiraCobrador.dSaldoAtual, objCarteiraCobrador.dSaldoAtualBanco, objCarteiraCobrador.iImprimeBoleta, objCarteiraCobrador.iComRegistro, objCarteiraCobrador.iGeraNossoNumero)
''''
''''        If lErro <> AD_SQL_SUCESSO Then Error 23488
''''
''''    End If
''''
''''    Call Comando_Fechar(lComando3)
''''    Call Comando_Fechar(lComando2)
''''    Call Comando_Fechar(lComando1)
''''
''''    CarteirasCobrador_Grava = SUCESSO
''''
''''    Exit Function
''''
''''Erro_CarteirasCobrador_Grava:
''''
''''    CarteirasCobrador_Grava = Err
''''
''''    Select Case Err
''''
''''        Case 23482, 23483, 41646
''''            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
''''
''''        Case 23485, 23486
''''            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
''''
''''        Case 23487
''''            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CARTEIRASCOBRADOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
''''
''''        Case 23488
''''            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CARTEIRASCOBRADOR", Err, objCarteiraCobrador.iCodCarteiraCobranca)
''''
''''        Case 42902, 42904, 42906, 42908, 41644
''''
''''        Case 42903, 42905, 42907, 42909
''''            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", Err, objPlanoConta.sConta)
''''
''''        Case 41645
''''            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err)
''''
''''        Case Else
''''            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148834)
''''
''''    End Select
''''
''''    Call Transacao_Rollback
''''    Call Comando_Fechar(lComando1)
''''    Call Comando_Fechar(lComando2)
''''    Call Comando_Fechar(lComando3)
''''
''''    Exit Function
''''
''''End Function

Private Function CarteirasCobrador_Exclui(ByVal objCobrador As ClassCobrador) As Long
'Exclui as Carteiras do Cobrador
'Chamada por Cobrador_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim iCarteira As Integer
Dim lQuantidadeAtual As Long, lQuantidadeAtualBanco As Long
Dim lComando(2) As Long

On Error GoTo Erro_CarteirasCobrador_Exclui

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 23584
    Next

    'Seleciona as Carteira(s) a ser(em) excluida(s)
    lErro = Comando_ExecutarPos(lComando(0), "SELECT CodCarteiraCobranca, QuantidadeAtual,QuantidadeAtualBanco FROM CarteirasCobrador WHERE Cobrador =?", 0, iCodigo, lQuantidadeAtual, lQuantidadeAtualBanco, objCobrador.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23585

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23586

    While lErro = SUCESSO

        If lQuantidadeAtual <> 0 Or lQuantidadeAtualBanco <> 0 Then Error 42893
        'double-check: verifica se existe alguma parcela associada a carteira
        lErro = Comando_Executar(lComando(1), "SELECT ParcelasRec.CarteiraCobranca FROM ParcelasRec WHERE ParcelasRec.Cobrador = ? AND ParcelasRec.CarteiraCobranca = ? UNION SELECT ParcelasRecBaixadas.CarteiraCobranca FROM ParcelasRecBaixadas WHERE ParcelasRecBaixadas.Cobrador = ? AND ParcelasRecBaixadas.CarteiraCobranca = ? ", iCarteira, objCobrador.iCodigo, iCodigo, objCobrador.iCodigo, iCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 42894
        
        lErro = Comando_BuscarPrimeiro(lComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42895
        If lErro = AD_SQL_SUCESSO Then Error 42896
        
        'Exclui Carteira
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM CarteirasCobrador", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 23588

        lErro = Comando_BuscarProximo(lComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19413

    Wend

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    CarteirasCobrador_Exclui = SUCESSO

    Exit Function

Erro_CarteirasCobrador_Exclui:

    CarteirasCobrador_Exclui = Err

    Select Case Err

        Case 23584
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23585, 23586, 19413
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRADOR1", Err)

        Case 23587 'Tratado na rotina que chama

        Case 23588
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CARTEIRASCOBRADOR", Err, iCodigo, objCobrador.iCodigo)

        Case 42893, 42896
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_VINCULADA_PARCELAS", Err)

        Case 42894, 42895
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148835)

    End Select

    Call Transacao_Rollback

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Private Function CarteirasCobranca_VerificaExclusao(ByVal iCodigo As Integer) As Long
'verifica se a carteira nao está sendo utilizada por algum cobrador

Dim lErro As Long
Dim lComando As Long
Dim iCodigo1 As Integer

On Error GoTo Erro_CarteirasCobranca_VerificaExclusao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23444

    'pesquisa Carteira no CarteirasCobrador
    lErro = Comando_Executar(lComando, "SELECT CodCarteiraCobranca FROM CarteirasCobrador WHERE CodCarteiraCobranca = ? ", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23445

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23446

    If lErro = AD_SQL_SUCESSO Then Error 23447

    lErro = Comando_Fechar(lComando)

    CarteirasCobranca_VerificaExclusao = SUCESSO

    Exit Function

Erro_CarteirasCobranca_VerificaExclusao:

    CarteirasCobranca_VerificaExclusao = Err

    Select Case Err

        Case 23444
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23445, 23446
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTEIRASCOBRADOR", Err)

        Case 23447
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COM_COBRADOR", Err, iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148836)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CarteiraDeCobranca_Grava(ByVal objCarteiraCobranca As ClassCarteiraCobranca) As Long
'Grava a Carteira no Banco de Dados

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long, iValidaPara As Integer
Dim iStatus As Integer
Dim iCodigo As Integer

On Error GoTo Erro_CarteiraDeCobranca_Grava

    lComando1 = 0
    lComando2 = 0
    lTransacao = 0

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 23394

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 23395

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 23396

    'Pesquisa no Banco de Dados a Carteira em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo, ValidaPara FROM CarteirasCobranca WHERE Codigo=?", 0, iCodigo, iValidaPara, objCarteiraCobranca.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23397

    'Lê os dados da Carteira, se estiver no Banco de Dados
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23398

    'Encontrou a carteira, então isto é uma alteração
    If lErro = AD_SQL_SUCESSO Then

        If iValidaPara = CARTCOBR_PARA_EMPRESA Then Error 59252
        
        'Atualiza os dados da Carteira
        lErro = Comando_ExecutarPos(lComando2, "UPDATE CarteirasCobranca SET Descricao=?,ValidaPara=?", lComando1, objCarteiraCobranca.sDescricao, objCarteiraCobranca.iValidaPara)
        If lErro <> AD_SQL_SUCESSO Then Error 23399

    'Não encontrou a carteira, então isto é uma inserção
    Else

        'Insere a carteira
        lErro = Comando_Executar(lComando2, "INSERT INTO CarteirasCobranca (Codigo,Descricao,ValidaPara) VALUES (?,?,?)", objCarteiraCobranca.iCodigo, objCarteiraCobranca.sDescricao, objCarteiraCobranca.iValidaPara)
        If lErro <> AD_SQL_SUCESSO Then Error 23400

    End If

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 23401

    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando1)

    CarteiraDeCobranca_Grava = SUCESSO

    Exit Function

Erro_CarteiraDeCobranca_Grava:

    CarteiraDeCobranca_Grava = Err

    Select Case Err

        Case 23394, 23395
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23396
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 23397, 23398
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA1", Err, objCarteiraCobranca.iCodigo)

        Case 23399
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CARTEIRASCOBRANCA", Err, objCarteiraCobranca.iCodigo)

        Case 23400
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CARTEIRASCOBRANCA", Err, objCarteiraCobranca.iCodigo)

        Case 23401
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 59252
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_CARTEIRA_EMPRESA", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148837)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function CarteiraDeCobranca_Exclui(ByVal iCodigo As Integer) As Long
'Exclui a Carteira do Banco de Dados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sDescricao As String
Dim iValidaPara As Integer

On Error GoTo Erro_CarteiraDeCobranca_Exclui

    lComando = 0
    lComando1 = 0
    lTransacao = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 23435

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 23436

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 23437

    sDescricao = String(STRING_DESCRICAO_CARTCOBR, 0)

    'Seleciona a Carteira que possue o Código em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Descricao,ValidaPara FROM CarteirasCobranca WHERE Codigo=?", 0, sDescricao, iValidaPara, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 23438

    'Lê o primeira Carteira da seleção feita
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 23439

    'Se não existir Carteira com o código em questão
    If lErro = AD_SQL_SEM_DADOS Then Error 23440

    'Loca o registro que guarda o Código do Carteira
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 23441

    lErro = CarteirasCobranca_VerificaExclusao(iCodigo)
    If lErro <> SUCESSO Then Error 23448

    'Exclui a Carteira do Banco de Dados
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM CarteirasCobranca", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 23442

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 23443

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    CarteiraDeCobranca_Exclui = SUCESSO

    Exit Function

Erro_CarteiraDeCobranca_Exclui:

    CarteiraDeCobranca_Exclui = Err

    Select Case Err

        Case 23435, 23436
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 23437
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 23438, 23439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELA_CARTEIRASCOBRANCA1", Err, iCodigo)

        Case 23440
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADO", Err, iCodigo)

        Case 23441
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CARTEIRASCOBRANCA", Err, iCodigo)

        Case 23442
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CARTEIRASCOBRANCA", Err, iCodigo)

        Case 23443
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 23448

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148838)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function PadraoCobranca_ObterNaColecao(objPadraoCobranca As ClassPadraoCobranca, ByVal colPadroesCobranca As Collection) As Long

Dim objPadraoCobrancaTemp As New ClassPadraoCobranca
Dim lErro As Long

On Error GoTo Erro_PadraoCobranca_ObterNaColecao

    For Each objPadraoCobrancaTemp In colPadroesCobranca
        If objPadraoCobrancaTemp.iCodigo = objPadraoCobranca.iCodigo Then
            With objPadraoCobranca
                .dJuros = objPadraoCobrancaTemp.dJuros
                .iDiasDeProtesto1 = objPadraoCobrancaTemp.iDiasDeProtesto1
                .iDiasDeProtesto2 = objPadraoCobrancaTemp.iDiasDeProtesto2
                .iInativo = objPadraoCobrancaTemp.iInativo
                .iInstrucao1 = objPadraoCobrancaTemp.iInstrucao1
                .iInstrucao2 = objPadraoCobrancaTemp.iInstrucao2
                .sDescricao = objPadraoCobrancaTemp.sDescricao
            End With
            PadraoCobranca_ObterNaColecao = SUCESSO
            Exit Function
        End If
    Next

    Error 7555

Erro_PadraoCobranca_ObterNaColecao:

    PadraoCobranca_ObterNaColecao = Err
    
    Select Case Err

        Case 7555
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_INVALIDO", Err, objPadraoCobranca.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148839)

    End Select

    Exit Function

End Function

Function CondicaoPagto_Grava(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
'Atualiza ou insere um novo registro na tabela CondicoesPagto

Dim lErro As Long
Dim iCondicao As Integer
Dim iEmPagamento As Integer
Dim iEmRecebimento As Integer
Dim alComando(0 To 7) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim lPedido As Long
Dim iFilial As Integer, objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_CondicaoPagto_Grava

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 16424

    Next
    
    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16427

    'Pesquisa descrição reduzida de outras Condições de Pagamento no BD
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM CondicoesPagto WHERE Codigo <> ? AND DescReduzida = ?", iCondicao, objCondicaoPagto.iCodigo, objCondicaoPagto.sDescReduzida)
    If lErro <> AD_SQL_SUCESSO Then Error 16428

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16429
    
    'Se já existe CondPagto com a desc reduzida, erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 16430

    'Pesquisa CondicaoPagto no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT EmPagamento, EmRecebimento FROM CondicoesPagto WHERE Codigo = ? ", 0, iEmPagamento, iEmRecebimento, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16431

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16432

    If lErro = AD_SQL_SEM_DADOS Then
    'CondicaoPagto não pertence ao BD

        'Insere novo registro na tabela CondicoesPagto
        lErro = Comando_Executar(alComando(1), "INSERT INTO CondicoesPagto (Codigo, DescReduzida, Descricao, EmPagamento, EmRecebimento, NumeroParcelas, DiasParaPrimeiraParcela, IntervaloParcelas, Mensal, DiaDoMes, AcrescimoFinanceiro, Modificador, DataFixa, CargoMinimo, FormaPagamento,TetoParcela) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objCondicaoPagto.iCodigo, objCondicaoPagto.sDescReduzida, objCondicaoPagto.sDescricao, objCondicaoPagto.iEmPagamento, objCondicaoPagto.iEmRecebimento, objCondicaoPagto.iNumeroParcelas, objCondicaoPagto.iDiasParaPrimeiraParcela, objCondicaoPagto.iIntervaloParcelas, objCondicaoPagto.iMensal, objCondicaoPagto.iDiaDoMes, objCondicaoPagto.dAcrescimoFinanceiro, objCondicaoPagto.iModificador, objCondicaoPagto.iDataFixa, objCondicaoPagto.iCargoMinimo, objCondicaoPagto.iFormaPagamento, objCondicaoPagto.dTetoParcela)
        If lErro <> AD_SQL_SUCESSO Then Error 16433

    Else  'CondicaoPagto existe no BD
        
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 124023
        
        lErro = CondicaoPagto_ExcluiParcelas(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 124024
        
        'Se quer alterar EmPagamento de 1 para 0
        If iEmPagamento = 1 And objCondicaoPagto.iEmPagamento = 0 Then
        
            'Seleciona Fornecedores que tenham esta Condicao Pagto
            lErro = Comando_Executar(alComando(2), "SELECT CondicaoPagto FROM Fornecedores WHERE CondicaoPagto = ?", iCodigo, objCondicaoPagto.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 6937
        
            'Verifica resultado da pesquisa
            lErro = Comando_BuscarPrimeiro(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6938
            
            If lErro = AD_SQL_SUCESSO Then Error 6939
        
            'Seleciona TiposFornecedor que tenham esta CondicaoPagto
            lErro = Comando_Executar(alComando(3), "SELECT CondicaoPagto FROM TiposDeFornecedor WHERE CondicaoPagto = ?", iCodigo, objCondicaoPagto.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 6940
        
            'Verifica resultado da pesquisa
            lErro = Comando_BuscarPrimeiro(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6941
            
            If lErro = AD_SQL_SUCESSO Then Error 6942
        
        End If

        'Se quer alterar EmRecebimento de 1 para 0
        If iEmRecebimento = 1 And objCondicaoPagto.iEmRecebimento = 0 Then
        
            'Seleciona Clientes que tenham esta Condicao Pagto
            lErro = Comando_Executar(alComando(4), "SELECT CondicaoPagto FROM Clientes WHERE CondicaoPagto = ?", iCodigo, objCondicaoPagto.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 6943
        
            'Verifica resultado da pesquisa
            lErro = Comando_BuscarPrimeiro(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6944
            
            If lErro = AD_SQL_SUCESSO Then Error 6945
        
            'Seleciona TiposCliente que tenham esta Condicao Pagto
            lErro = Comando_Executar(alComando(5), "SELECT CondicaoPagto FROM TiposDeCliente WHERE CondicaoPagto = ?", iCodigo, objCondicaoPagto.iCodigo)
            If lErro <> AD_SQL_SUCESSO Then Error 6946
        
            'Verifica resultado da pesquisa
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6947
           
            If lErro = AD_SQL_SUCESSO Then Error 6948
            
            'Verifica se CondicaoPagto está vinculada a Pedidos de Venda ou a Pedidos de Venda Baixados
            lErro = CF("PedidosDeVenda_Le_CondicaoPagto", objCondicaoPagto, lPedido, iFilial)
            If lErro <> SUCESSO And lErro <> 33989 Then Error 33992
            
            'Se estiver vinculada, erro
            If lErro = SUCESSO Then Error 33993
 
        End If
       
        'Altera registro na tabela CondicoesPagto
        lErro = Comando_ExecutarPos(alComando(6), "UPDATE CondicoesPagto SET DescReduzida = ?, Descricao = ?, EmPagamento = ?, EmRecebimento = ?, NumeroParcelas = ?, DiasParaPrimeiraParcela = ?, IntervaloParcelas = ?, Mensal = ?, DiaDoMes = ?, AcrescimoFinanceiro = ?, Modificador = ?, DataFixa = ?, CargoMinimo = ?, FormaPagamento = ?, TetoParcela = ?", alComando(0), objCondicaoPagto.sDescReduzida, objCondicaoPagto.sDescricao, objCondicaoPagto.iEmPagamento, objCondicaoPagto.iEmRecebimento, objCondicaoPagto.iNumeroParcelas, objCondicaoPagto.iDiasParaPrimeiraParcela, objCondicaoPagto.iIntervaloParcelas, objCondicaoPagto.iMensal, objCondicaoPagto.iDiaDoMes, objCondicaoPagto.dAcrescimoFinanceiro, objCondicaoPagto.iModificador, objCondicaoPagto.iDataFixa, objCondicaoPagto.iCargoMinimo, objCondicaoPagto.iFormaPagamento, objCondicaoPagto.dTetoParcela)
        If lErro <> AD_SQL_SUCESSO Then Error 16434

    End If

    For Each objParc In objCondicaoPagto.colParcelas
    
        With objParc
            lErro = Comando_Executar(alComando(7), "INSERT INTO CondicoesPagtoParc (Codigo, Seq, TipoDataBase, TipoIntervalo, Dias, Modificador, PercReceb) VALUES (?,?,?,?,?,?,?)", .iCodigo, .iSeq, .iTipoDataBase, .iTipoIntervalo, .iDias, .iModificador, .dPercReceb)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 124030
    
    Next
    
    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16435

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CondicaoPagto_Grava = SUCESSO

    Exit Function

Erro_CondicaoPagto_Grava:

    CondicaoPagto_Grava = Err

        Select Case Err

        Case 6937, 6938
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES", Err)

        Case 6939
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_VINCULADO_COND_PAGTO", Err, objCondicaoPagto.iCodigo)

        Case 6940, 6941
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSFORNECEDOR", Err)

        Case 6942
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_VINCULADO_COND_PAGTO", Err, objCondicaoPagto.iCodigo)
        
        Case 6943, 6944
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 6945
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_VINCULADO_COND_PAGTO", Err, objCondicaoPagto.iCodigo)

        Case 6946, 6947
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTE", Err)

        Case 6948
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_CLIENTE_VINCULADO_COND_PAGTO", Err, objCondicaoPagto.iCodigo)
        
        Case 16424
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16427
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16428, 16429, 16431, 16432
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO", Err, objCondicaoPagto.iCodigo)

        Case 16430
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_REDUZIDA_CONDICAOPAGTO_REPETIDA", Err)

        Case 16433
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONDICAOPAGTO", Err, objCondicaoPagto.iCodigo)

        Case 16434
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONDICAOPAGTO", Err, objCondicaoPagto.iCodigo)

        Case 16435
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
             
        Case 124023
             Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONDICAOPAGTO", Err)
        
        Case 33992, 124024

        Case 33993
             Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_REL_PED_VENDA", Err, objCondicaoPagto.iCodigo, lPedido, iFilial)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148840)

    End Select

    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CondicaoPagto_Exclui(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
'Exclui CondicaoPagto do BD

Dim lErro As Long
Dim iCodigo As Integer
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lTotal As Long
Dim lPedido As Long
Dim iFilial As Integer

On Error GoTo Erro_CondicaoPagto_Exclui

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 16441

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 16442

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 16443
 
    'Pesquisa CondicaoPagto no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM CondicoesPagto WHERE Codigo = ? ", 0, iCodigo, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16444

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16445

    If lErro = AD_SQL_SEM_DADOS Then gError 16446

    'Lock da CondicaoPagto
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 16447

    'Verifica se CondicaoPagto está relacionada com Clientes
    lErro = Clientes_Le_CondicaoPagto(objCondicaoPagto.iCodigo, lTotal)
    If lErro <> SUCESSO And lErro <> 16461 Then gError 16448
    
    If lErro = SUCESSO Then gError 16449
    
    'Verifica se CondicaoPagto está relacionada com Fornecedores
    lErro = Fornecedores_Le_CondicaoPagto(objCondicaoPagto.iCodigo, lTotal)
    If lErro <> SUCESSO And lErro <> 16465 Then gError 16450
    
    If lErro = SUCESSO Then gError 16451

    'Verifica se CondicaoPagto está relacionada com Tipos De Cliente
    lErro = TiposCliente_Le_CondicaoPagto(objCondicaoPagto.iCodigo, lTotal)
    If lErro <> SUCESSO And lErro <> 16469 Then gError 16452
    
    If lErro = SUCESSO Then gError 16453
 
    'Verifica se CondicaoPagto está relacionada com Tipos De Fornecedor
    lErro = TiposFornecedor_Le_CondicaoPagto(objCondicaoPagto.iCodigo, lTotal)
    If lErro <> SUCESSO And lErro <> 16473 Then gError 16454
    
    If lErro = SUCESSO Then gError 16455
    
    'Verifica se CondicaoPagto está vinculada a Pedidos de Venda ou a Pedidos de Venda Baixados
    lErro = CF("PedidosDeVenda_Le_CondicaoPagto", objCondicaoPagto, lPedido, iFilial)
    If lErro <> SUCESSO And lErro <> 33989 Then gError 33990
    
    'Se estiver, erro
    If lErro = SUCESSO Then gError 33991
     
    'Verifica se a Condição de Pagamento está vinculada com Títulos a Pagar e Títulos a Pagar Baixados
    lErro = CF("TitulosPag_Le_CondicaoPagto", objCondicaoPagto.iCodigo)
    If lErro <> SUCESSO Then gError 58118
    
    'Verifica se a Condição de Pagamento está vinculada com Títulos a Receber e Títulos a Receber Baixados
    lErro = CF("TitulosRec_Le_CondicaoPagto", objCondicaoPagto.iCodigo)
    If lErro <> SUCESSO Then gError 58158
    
    'Verifica se a CondPagto é usada no COMPRAS
    lErro = CondPagto_Critica_Compras(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 32253
    
    'Verifica se a CondPagto é usada (Dependendo do Cliente)
    lErro = CF("CondPagto_Critica_Usu", objCondicaoPagto)
    If lErro <> SUCESSO Then gError 97905
    
    lErro = CondicaoPagto_ExcluiParcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 124022
    
    'Exclui CondicaoPagto da tabela CondicoesPagto
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM CondicoesPagto", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 16456

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 16457

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    CondicaoPagto_Exclui = SUCESSO

    Exit Function

Erro_CondicaoPagto_Exclui:

    CondicaoPagto_Exclui = gErr

    Select Case gErr
                       
        Case 16441, 16442
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 16443
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 16444, 16445
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO", gErr, objCondicaoPagto.iCodigo)

        Case 16446
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 16447
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONDICAOPAGTO", gErr)

        Case 16448, 16450, 16452, 16454, 33990, 58118, 58158, 32253, 97905, 124022
            'Erro tratado na rotina chamada
        
        Case 16449
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_RELACIONADA_COM_CLIENTE", gErr, lTotal)
            
        Case 16451
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_RELACIONADA_COM_FORNECEDOR", gErr, lTotal)
     
        Case 16453
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDECLIENTE", gErr, lTotal)
        
        Case 16455
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_RELACIONADA_COM_TIPOSDEFORNECEDOR", gErr, lTotal)
        
        Case 16456
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONDICAOPAGTO", gErr, objCondicaoPagto.iCodigo)

        Case 16457
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 33991
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAOPAGTO_REL_PED_VENDA", gErr, objCondicaoPagto.iCodigo, lPedido, iFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148841)

    End Select

    Call Transacao_Rollback
    
    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function PadraoCobranca_Grava(ByVal objPadraoCobranca As ClassPadraoCobranca) As Long
'Atualiza ou insere um novo registro na tabela PadroesCobranca

Dim lErro As Long, iInativo As Integer
Dim iCodigo As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim sDescricao As String
Dim iInstrucao1 As Integer
Dim iInstrucao2 As Integer

On Error GoTo Erro_PadraoCobranca_Grava

    lComando = 0
    lComando1 = 0
    lComando2 = 0
    lComando3 = 0
    lComando4 = 0
    lTransacao = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16573

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16574

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 16575
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 40609
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then Error 40610
    
    'inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16576

    'verifica se a descricao já é usada por um padrao com outro codigo
    lErro = Comando_Executar(lComando, "SELECT Inativo FROM PadroesCobranca WHERE Codigo <> ? AND Descricao = ?", iInativo, objPadraoCobranca.iCodigo, objPadraoCobranca.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 16577

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16578
    If lErro <> AD_SQL_SEM_DADOS Then Error 16579
    
    'pesquisa Padrao Cobranca no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo, Instrucao1, Instrucao2 FROM PadroesCobranca WHERE Codigo = ? ", 0, iCodigo, iInstrucao1, iInstrucao2, objPadraoCobranca.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16580

    'verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16581

    If lErro = AD_SQL_SEM_DADOS Then
    'PadraoCobranca nao pertence ao BD

        'para as instrucoes <> 0 chamar InstrCobranca_Lock
        If objPadraoCobranca.iInstrucao1 <> 0 Then
            
            'Verifica a Integridade da Instrucao com o BD
            lErro = CF("InstrCobranca_Lock", objPadraoCobranca.iInstrucao1, lComando3)
            If lErro <> SUCESSO And lErro <> 40603 Then Error 40605
        
        End If
        
        If objPadraoCobranca.iInstrucao2 <> 0 Then
        
            'Verifica a Integridade da Instrucao com o BD
            lErro = CF("InstrCobranca_Lock", objPadraoCobranca.iInstrucao1, lComando4)
            If lErro <> SUCESSO And lErro <> 40603 Then Error 40606
        
        End If
                
        'insere novo registro na tabela PadroesCobranca
        lErro = Comando_Executar(lComando2, "INSERT INTO PadroesCobranca (Codigo, Descricao, Inativo, Juros, Instrucao1, DiasDeProtesto1, Instrucao2, DiasDeProtesto2) VALUES (?,?,?,?,?,?,?,?)", objPadraoCobranca.iCodigo, objPadraoCobranca.sDescricao, objPadraoCobranca.iInativo, objPadraoCobranca.dJuros, objPadraoCobranca.iInstrucao1, objPadraoCobranca.iDiasDeProtesto1, objPadraoCobranca.iInstrucao2, objPadraoCobranca.iDiasDeProtesto2)
        If lErro <> AD_SQL_SUCESSO Then Error 16582

    Else
    'PadraoCobranca existe no BD

        'para as instrucoes <> 0 que foram modificadas, chamar InstrCobranca_Lock
        If objPadraoCobranca.iInstrucao1 <> 0 Then
            
            'Verifica a Integridade da Instrucao com o BD
            lErro = CF("InstrCobranca_Lock", objPadraoCobranca.iInstrucao1, lComando3)
            If lErro <> SUCESSO And lErro <> 40603 Then Error 40607
        
        End If
        
        If objPadraoCobranca.iInstrucao2 <> 0 Then
        
            'Verifica a Integridade da Instrucao com o BD
            lErro = CF("InstrCobranca_Lock", objPadraoCobranca.iInstrucao1, lComando4)
            If lErro <> SUCESSO And lErro <> 40603 Then Error 40608
        End If
        
        'altera registro na tabela PadroesCobranca
        lErro = Comando_ExecutarPos(lComando2, "UPDATE PadroesCobranca SET Codigo = ?, Descricao = ?, Inativo = ?, Juros = ?, Instrucao1 = ?, DiasDeProtesto1 = ?, Instrucao2 = ?, DiasDeProtesto2 = ? ", lComando1, objPadraoCobranca.iCodigo, objPadraoCobranca.sDescricao, objPadraoCobranca.iInativo, objPadraoCobranca.dJuros, objPadraoCobranca.iInstrucao1, objPadraoCobranca.iDiasDeProtesto1, objPadraoCobranca.iInstrucao2, objPadraoCobranca.iDiasDeProtesto2)
        If lErro <> AD_SQL_SUCESSO Then Error 16583

    End If

    'confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16584

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)
    lErro = Comando_Fechar(lComando2)
    lErro = Comando_Fechar(lComando3)
    lErro = Comando_Fechar(lComando4)

    PadraoCobranca_Grava = SUCESSO

    Exit Function

Erro_PadraoCobranca_Grava:

    PadraoCobranca_Grava = Err

        Select Case Err

        Case 16573, 16574, 16575, 40609, 40610
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16576
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16577, 16578, 16580, 16581
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 16579
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_PADRAO_COBRANCA_REPETIDA", Err)

        Case 16582
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 16583
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 16584
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
             
        Case 40605, 40606, 40607, 40608
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148842)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function PadraoCobranca_Exclui(ByVal objPadraoCobranca As ClassPadraoCobranca) As Long
'exclui Padrao Cobranca do BD

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim iCodigo As Integer
Dim lTotal As Long

On Error GoTo Erro_PadraoCobranca_Exclui

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16590

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 16591

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16592
 
    'pesquisa PadraoCobranca no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM PadroesCobranca WHERE Codigo = ? ", 0, iCodigo, objPadraoCobranca.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16593

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16594

    If lErro = AD_SQL_SEM_DADOS Then Error 16595

    'lock do Padrao Cobranca
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 16596
    
    'verifica se Padrao Cobranca esta relacionada com Tipos De Cliente
    lErro = CF("TiposCliente_Le_PadraoCobranca", objPadraoCobranca.iCodigo)
    If lErro = SUCESSO Then Error 16597

    If lErro <> 16604 Then Error 16598
    
    'verifica se Padrão Cobranca esta relacionada com tabela de clientes
    lErro = CF("Clientes_Le_PadraoCobranca", objPadraoCobranca.iCodigo)
    If lErro = SUCESSO Then Error 40611
    
    If lErro <> 40616 Then Error 40612
    
    'exclui Padrao Cobranca da tabela PadroesCobranca
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM PadroesCobranca", lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 16599

    'confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16600

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    PadraoCobranca_Exclui = SUCESSO

    Exit Function

Erro_PadraoCobranca_Exclui:

    PadraoCobranca_Exclui = Err

    Select Case Err
                       
        Case 16590, 16591
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16592
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16593, 16594
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAO_COBRANCA", Err)

        Case 16595
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO", Err, objPadraoCobranca.sDescricao)

        Case 16596
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 16597
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_RELACIONADO_COM_TIPOS_DE_CLIENTE", Err)
            
        Case 16598, 40612
            'Erro tratado na rotina chamada
        
        Case 16599
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PADRAO_COBRANCA", Err, objPadraoCobranca.iCodigo)

        Case 16600
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 40611
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_RELACIONADO_COM_CLIENTE", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148843)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Transportadora_Automatico(iCodigo As Integer) As Long
'obtem o proximo numero da Transportadora

Dim lErro As Long, lTemp As Long

On Error GoTo Erro_Transportadora_Automatico

    lErro = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_TRANSPORTADORA", "Transportadoras", "Codigo", lTemp)
    If lErro <> SUCESSO Then Error 22115

    iCodigo = lTemp

    Transportadora_Automatico = SUCESSO

    Exit Function

Erro_Transportadora_Automatico:

    Transportadora_Automatico = Err

   Select Case Err

        Case 22115

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148844)

    End Select

    Exit Function

End Function

Function Transportadora_Grava(ByVal objTransportadora As ClassTransportadora, ByVal objEndereco As ClassEndereco) As Long
'Grava A Transportadora no BD

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Transportadora_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 22064

    lErro = Transportadora_Grava_EmTrans(objTransportadora, objEndereco)
    If lErro <> SUCESSO Then gError 97728

    'Confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 22068

    Transportadora_Grava = SUCESSO

    Exit Function

Erro_Transportadora_Grava:

    Transportadora_Grava = gErr

    Select Case gErr

        Case 22064
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 22068
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 97728
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148845)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

'mario
Function CRFATConfig_ObterNumInt(ByVal lComandoSel As Long, ByVal lComandoMod As Long, ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_CRFATConfig_ObterNumInt

    sConteudo = String(STRING_CRFATCONFIG_CONTEUDO, 0)

    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Conteudo FROM CRFATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then Error 7365

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7366
    If lErro <> AD_SQL_SUCESSO Then Error 7365

    lNumInt = CLng(sConteudo)
    sConteudo = CStr(lNumInt + iQtde)

    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE CRFATConfig SET Conteudo = ?", lComandoSel, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 7368

    CRFATConfig_ObterNumInt = SUCESSO

    Exit Function

Erro_CRFATConfig_ObterNumInt:

    CRFATConfig_ObterNumInt = Err

    Select Case Err

        Case 7365, 7366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRFATCONFIG", Err, sCodigo)

        Case 7368
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CRFATCONFIG", Err, sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148846)

    End Select

    Exit Function

End Function

Function Cliente_Automatico(lCodigo As Long) As Long
'retorna o número do proximo Cliente disponivel

    Cliente_Automatico = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_CLIENTE", "Clientes", "Codigo", lCodigo)

End Function

Function Cliente_Grava(ByVal objCliente As ClassCliente, ByVal colEndereco As Collection) As Long
'Grava o Cliente passado como parâmetro

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Cliente_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12370

    lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
    If lErro <> SUCESSO Then Error 59487
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12388

    Cliente_Grava = SUCESSO

    Exit Function

Erro_Cliente_Grava:

    Cliente_Grava = Err

    Select Case Err

        Case 12370
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 12388
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)

        Case 59487
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148847)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function FilialCliente_Insere_Categorias(ByVal objFilialCliente As ClassFilialCliente, ByVal colCategoriaItem As Collection) As Long
'Percorre as Categorias na coleção incluindo-as no BD
'Chamada DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim lComando As Long, lComando1 As Long
Dim sCategoria As String, sItem As String
Dim objFilialCliCategoria As New ClassFilialCliCategoria

On Error GoTo Erro_FilialCliente_Insere_Categorias

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33035

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 33066
    
    'Para cada registro da coleção
    For Each objFilialCliCategoria In colCategoriaItem

        lErro = CF("CategoriaClienteItem_Lock", objFilialCliCategoria.sCategoria, objFilialCliCategoria.sItem, lComando)
        If lErro <> SUCESSO Then Error 33067
        
        'Insere a filial cliente categoria
        lErro = Comando_Executar(lComando1, "INSERT INTO FilialClienteCategorias (Cliente, Filial, Categoria, Item) VALUES(?,?,?,?)", objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, objFilialCliCategoria.sCategoria, objFilialCliCategoria.sItem)
        If lErro <> AD_SQL_SUCESSO Then Error 33070

    Next

    'Libera comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    FilialCliente_Insere_Categorias = SUCESSO

    Exit Function

Erro_FilialCliente_Insere_Categorias:

    FilialCliente_Insere_Categorias = Err

    Select Case Err

        Case 33035, 33066
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33067

        Case 33070
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALCLIENTECATEGORIAS", Err, objFilialCliente.lCodCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148848)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function FilialCliente_Atualiza_Categorias(ByVal objFilialCliente As ClassFilialCliente, ByVal colCategoriaItem As Collection) As Long
'atualiza as categorias da filial cliente no bd
'Chamada DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim lComando As Long, lComando1 As Long, lComando2 As Long
Dim lCliente As Long, sCategoria As String, sItem As String, sCategoria1 As String, sItem1 As String, iIndice As Integer
Dim objFilialCliCategoria As ClassFilialCliCategoria

On Error GoTo Erro_FilialCliente_Atualiza_Categorias

    'Inicializa commandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 33037

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 33038

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 33071

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Percorre as categorias associadas a filial que já estão no bd
    lErro = Comando_ExecutarPos(lComando, "SELECT Categoria, Item FROM FilialClienteCategorias WHERE Cliente = ? AND Filial = ?", 0, sCategoria, sItem, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 33039

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33040

    Do While lErro <> AD_SQL_SEM_DADOS

        iIndice = FilialCliente_Grava_CategoriaNaColecao(colCategoriaItem, sCategoria)

        If iIndice <> 0 Then
            
            Set objFilialCliCategoria = colCategoriaItem.Item(iIndice)
            
            'Se houver sido alterado o Item
            If objFilialCliCategoria.sItem <> sItem Then

                lErro = CF("CategoriaClienteItem_Lock", sCategoria, objFilialCliCategoria.sItem, lComando1)
                If lErro <> SUCESSO Then Error 33072
                
                'Atualiza a categoria da filial do cliente
                lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialClienteCategorias SET Item = ?", lComando, objFilialCliCategoria.sItem)
                If lErro <> AD_SQL_SUCESSO Then Error 33041

            End If

            'Retira da coleção
            colCategoriaItem.Remove (iIndice)

        Else
            
            'Excluir registro em FilialCliCategoria
            lErro = Comando_ExecutarPos(lComando2, "DELETE FROM FilialClienteCategorias", lComando)
            If lErro <> AD_SQL_SUCESSO Then Error 33075
            
        End If

        'Tenta ler a próxima categoria do cliente da filial
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 33040

    Loop

    'Chama FilialCliente_Insere_Categorias
    lErro = CF("FilialCliente_Insere_Categorias", objFilialCliente, colCategoriaItem)
    If lErro <> SUCESSO Then Error 1

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    FilialCliente_Atualiza_Categorias = SUCESSO

    Exit Function

Erro_FilialCliente_Atualiza_Categorias:

    FilialCliente_Atualiza_Categorias = Err

    Select Case Err

        Case 33037, 33038, 33071
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33039, 33040
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS2", Err, objFilialCliCategoria.lCliente)

        Case 33041
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_FILIALCLIENTECATEGORIAS", Err, objFilialCliCategoria.lCliente)

        Case 33042
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALCLIENTECATEGORIAS", Err, objFilialCliCategoria.lCliente)

        Case 33072

        Case 33075
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS1", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 33107
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_CATEGORIA_NAO_CADASTRADO", Err, sItem, sCategoria)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148849)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Private Function FilialCliente_Grava_CategoriaNaColecao(ByVal colCategoriaItem As Collection, ByVal sCategoria As String) As Long
'Retorna o Indice da Categoria na coleção ou Zero se não achá-la (na coleção)

Dim iIndice As Integer
Dim objFilailCliCategoria As New ClassFilialCliCategoria

    'Pesquisa a Categoria na coleção
    For iIndice = 1 To colCategoriaItem.Count

        Set objFilailCliCategoria = colCategoriaItem.Item(iIndice)

        'Se achou a categoria na coleção
        If objFilailCliCategoria.sCategoria = sCategoria Then

            FilialCliente_Grava_CategoriaNaColecao = iIndice

            Exit Function

        End If

    Next

    FilialCliente_Grava_CategoriaNaColecao = 0

End Function

Function Cliente_Exclui(ByVal objCliente As ClassCliente) As Long
'Exclui o cliente e suas filiais

Dim lErro As Long
Dim lComando(2) As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim lPedido As Long
Dim lNumIntCheque As Long
Dim iIndice As Integer
Dim iFilialEmpresa As Integer
Dim objLog As New ClassLog
Dim colEndereco As New Collection
Dim lCodigoLoja As Long
Dim lCodigoCliente As Long
Dim iTransferido As Integer
Dim lCodRelacionamento As Long

On Error GoTo Erro_Cliente_Exclui

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 12468
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 12472
    
    'guarda o código para as msgs de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lCodigoCliente = objCliente.lCodigoLoja
    Else
        lCodigoCliente = objCliente.lCodigo
    End If
    
    
    'Seleciona o Cliente a ser excluído
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lErro = Comando_ExecutarPos(lComando(0), "SELECT Transferido FROM Clientes WHERE CodigoLoja =?", 0, iTransferido, objCliente.lCodigoLoja)
    Else
        lErro = Comando_ExecutarPos(lComando(0), "SELECT Transferido FROM Clientes WHERE Codigo =?", 0, iTransferido, objCliente.lCodigo)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 12473
    
    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12474

    If lErro = AD_SQL_SEM_DADOS Then gError 12475

    'se o registro foi transferido ==> nao pode excluir
    If iTransferido = TRANSFERIDO Then gError 110113

    'Lock do Cliente
    lErro = Comando_LockExclusive(lComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 12270

    'Lock das Filiais
    lErro = CF("FiliaisClientes_Lock", lCodigoCliente)
    If lErro <> SUCESSO Then gError 12521

    If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then

        'Verifica se Cliente está relacionado à PedidoDeVenda ou a PedidoVendaBaixado.
        lErro = CF("PedidoVenda_Le_Cliente", objCliente, lPedido, iFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 33935 Then gError 33936
        
        'Cliente está relacionado com Pedido de Venda --> erro
        If lErro = SUCESSO Then gError 33937
    
        'Verifica se existem Notas Fiscais relacionadas ao Cliente
        lErro = CF("NFs_Cliente_Nao_Relacionado", lCodigoCliente)
        If lErro <> SUCESSO Then gError 6877
    
        'Verifica se Cliente está relacionado com ChequePre.
        lErro = CF("ChequePre_Le_Cliente", objCliente, lNumIntCheque)
        If lErro <> SUCESSO And lErro <> 33941 Then gError 33942
        
        'Cliente está relacionado com Cheque Pré --> erro
        If lErro = SUCESSO Then gError 33943
    
        'Preenche a coleção de códigos de Débitos Receber relacionados ao Cliente
        lErro = CF("DebitosRecCli_Le_Cliente", lCodigoCliente)
        If lErro <> SUCESSO And lErro <> 12452 Then gError 12447
    
        'Verifica se Cliente está relacionado com Débitos
        If lErro = SUCESSO Then gError 12448
    
        'Preenche a coleção de códigos de Recebimentos Antecipados relacionados ao Cliente
        lErro = CF("RecebAntecipados_Le_Cliente", lCodigoCliente)
        If lErro <> SUCESSO And lErro <> 12459 Then gError 12454
    
        'Verifica se Cliente está relacionado com Recebimentos Antecipados
        If lErro = SUCESSO Then gError 12455
        
        'Verifica se o Cliente está relacionado com Títulos a Receber e Títulos a Receber Baixados
        lErro = CF("TitulosReceber_Cliente_Nao_Relacionado", lCodigoCliente)
        If lErro <> SUCESSO Then gError 58101
        
        'Verifica se a CondPagto é usada (Dependendo do Cliente)
        lErro = CF("Cliente_Critica_Usu", objCliente)
        If lErro <> SUCESSO Then gError 97906
        
        'Inserido por Luiz Nogueira em 27/10/03
        'Verifica se existe algum relacionamento de cliente vinculado a esse cliente
        lErro = CF("RelacionamentoClientes_Verifica_Cliente", objCliente.lCodigo, lCodRelacionamento)
        If lErro <> SUCESSO And lErro <> 102694 Then gError 102696
        
        'Se existir => erro
        If lErro = SUCESSO Then gError 102697
        
        lErro = CF("RotasPontos_Testa_Filial", objCliente.lCodigo, 0)
        If lErro <> SUCESSO Then gError 102696
        
        'Incluído por Luiz Nogueira em 27/10/03
        'Exclui os contatos relacionados ao cliente
        lErro = CF("ClienteContatos_Exclui_Cliente_EmTrans", objCliente.lCodigo)
        If lErro <> SUCESSO Then gError 102703
        
        'Exclui registros de MvDiaCli relacionados ao Cliente. Não deverá haver nenhum. Só para dar robustez na integridade referencial
        lErro = CF("MvDiaCli_Exclui_Cliente", lCodigoCliente)
        If lErro <> SUCESSO Then gError 6882
    
        'Exclui registros de MvPerCli relacionados ao Cliente. Não deverá haver nenhum. Só para dar robustez na integridade referencial
        lErro = CF("MvPerCli_Exclui_Cliente", lCodigoCliente)
        If lErro <> SUCESSO Then gError 6883
    
        'Exclui as Categorias das Filiais do Cliente
        lErro = Cliente_Exclui_Categorias(lCodigoCliente)
        If lErro <> SUCESSO Then gError 33043
    
    End If
    
    'Exclui as Filiais do Cliente
    lErro = CF("FiliaisClientes_Exclui", lCodigoCliente)
    If lErro <> SUCESSO Then gError 6880

    '################################################################
    'Inserido por Wagner
    lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_CLIENTE, lCodigoCliente)
    If lErro <> SUCESSO Then gError 6880
    '################################################################

    'Exclui o Cliente
    lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM Clientes", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 12476
    
    lErro = CF("Cliente_Exclui_Customizado", objCliente)
    If lErro <> SUCESSO Then gError 192969
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 12481

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Cliente_Exclui = SUCESSO

    Exit Function

Erro_Cliente_Exclui:

    Cliente_Exclui = gErr

    Select Case gErr

        Case 6877, 6880, 6882, 6883, 12447, 12454, 12521, 33043, 33936, 33942, 58101, 97906, 99930, 102696, 102703, 192969
            'Tratados nas rotinas chamadas

        Case 12270
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", gErr, lCodigoCliente)

        Case 12448
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_DEBITOS", gErr, lCodigoCliente)

        Case 12455
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_RECEB_ANTEC", gErr, lCodigoCliente)

        Case 12468
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 12472
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 12473, 12474
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 12475
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, lCodigoCliente)

        Case 12476
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTE", gErr, lCodigoCliente)

        Case 12481
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 33937
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_PED_VENDA", gErr, lCodigoCliente, iFilialEmpresa, lPedido)

        Case 33943
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_CHEQUE_PRE", gErr, lCodigoCliente)
        
        Case 110113
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTE_TRANSFERIDO", gErr, lCodigoCliente)
        
        Case 102697
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_RELACIONAMENTOCLIENTE", gErr, objCliente.lCodigo, lCodRelacionamento)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148850)

    End Select

    Call Transacao_Rollback

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Private Function Cliente_Exclui_Categorias(ByVal lCodCliente As Long) As Long
'Exclui as Categorias das Filiais do Cliente cujo código foi passado como parâmetro.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim lComando(2) As Long
Dim iIndice As Integer
Dim iFilial As Integer
Dim iFim_de_Arquivo As Integer

On Error GoTo Erro_Cliente_Exclui_Categorias

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 33044
    Next

    'Seleciona registros de FilialClienteCategorias a serem excluídos
    lErro = Comando_ExecutarPos(lComando(0), "SELECT Filial FROM FilialClienteCategorias WHERE Cliente =?", 0, iFilial, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 33045

    'Tenta ler
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 33046

    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Exclui registro de FilialClienteCategorias ligado ao Cliente
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM FilialClienteCategorias", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 33048

        'Traz o próximo registro
        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 33049

    Loop

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Cliente_Exclui_Categorias = SUCESSO

    Exit Function

Erro_Cliente_Exclui_Categorias:

    Cliente_Exclui_Categorias = Err

    Select Case Err

        Case 33044
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33045, 33046, 33049
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS2", Err, lCodCliente)

        Case 33048
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS", Err, lCodCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148851)

    End Select

    'Fecha os comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function MvDiaCli_Exclui_Cliente(ByVal lCodCliente As Long) As Long
'Exclui os registros da tabela MvDiaCli associados ao Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lComando(2) As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tMvDiaCli As typeMvDiaCli
Dim iIndice As Integer
    
On Error GoTo Erro_MvDiaCli_Exclui_Cliente

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 6890
    Next
    
    'Seleciona os Movimentos do Cliente a serem excluídos
    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa, Filial, Data FROM MvDiaCli WHERE Cliente =?", 0, tMvDiaCli.iFilialEmpresa, tMvDiaCli.iFilial, tMvDiaCli.dtData, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 6891

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6892

    'Exclui os Movimentos do Cliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6893
        
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM MvDiaCli", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6894

        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6895

    Loop
    
    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    MvDiaCli_Exclui_Cliente = SUCESSO

    Exit Function

Erro_MvDiaCli_Exclui_Cliente:

    MvDiaCli_Exclui_Cliente = Err

    Select Case Err

        Case 6890
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 6891, 6892, 6895
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACLI", Err)

        Case 6893
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVDIACLI", Err, tMvDiaCli.iFilialEmpresa, lCodCliente, tMvDiaCli.iFilial, tMvDiaCli.dtData)
        
        Case 6894
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVDIACLI", Err, tMvDiaCli.iFilialEmpresa, lCodCliente, tMvDiaCli.iFilial, tMvDiaCli.dtData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148852)

    End Select
    
    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Private Function Enderecos_Exclui_Cliente(ByVal objCliente As ClassCliente) As Long
'Exclui os Endereços do Cliente (principal, entrega e cobrança)
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
    
On Error GoTo Erro_Enderecos_Exclui_Cliente

    'Exclui o Endereço principal do Cliente
    lErro = CF("Endereco_Exclui", objCliente.lEndereco)
    If lErro <> SUCESSO Then Error 12578
    
    'Exclui o Endereço de entrega do Cliente
    lErro = CF("Endereco_Exclui", objCliente.lEnderecoEntrega)
    If lErro <> SUCESSO Then Error 12512
    
    'Exclui o Endereço de cobrança do Cliente
    lErro = CF("Endereco_Exclui", objCliente.lEnderecoCobranca)
    If lErro <> SUCESSO Then Error 12516

    Enderecos_Exclui_Cliente = SUCESSO

    Exit Function

Erro_Enderecos_Exclui_Cliente:

    Enderecos_Exclui_Cliente = Err

    Select Case Err

        Case 12512, 12516, 12578  'Tratados na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148853)

    End Select

    Exit Function

End Function

Function FiliaisClientes_Exclui(ByVal lCodCliente As Long) As Long
'Exclui as Filiais do Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lComando(2) As Long
Dim lErro As Long
Dim lEndereco(3) As Long
Dim iFim_de_Arquivo As Integer
Dim iIndice As Integer
Dim objFilialCliente As New ClassFilialCliente
    
On Error GoTo Erro_FiliaisClientes_Exclui

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 6881
    Next
    
    'Seleciona as Filiais do Cliente a serem excluídas
    lErro = Comando_ExecutarPos(lComando(0), "SELECT Endereco, EnderecoEntrega, EnderecoCobranca FROM FiliaisClientes WHERE CodCliente =?", 0, lEndereco(0), lEndereco(1), lEndereco(2), lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 12477

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 12478
    If iFim_de_Arquivo = AD_SQL_SEM_DADOS Then Error 12479

    'Exclui as Filiais do Cliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Exclui a Filial
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM FiliaisClientes", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 12480

        objFilialCliente.lEndereco = lEndereco(0)
        objFilialCliente.lEnderecoEntrega = lEndereco(1)
        objFilialCliente.lEnderecoCobranca = lEndereco(2)
        
        'Exclui os Endereços da Filial
        lErro = CF("Enderecos_Exclui_FilialCliente", objFilialCliente)
        
        'Traz a próxima Filial
        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6879

    Loop
    
    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    FiliaisClientes_Exclui = SUCESSO

    Exit Function

Erro_FiliaisClientes_Exclui:

    FiliaisClientes_Exclui = Err

    Select Case Err

        Case 6881
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 6879, 12477, 12478
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err)

        Case 12479
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", Err, lCodCliente)

        Case 12480
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIAISCLIENTES", Err, lCodCliente)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148854)

    End Select
    
    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function MvPerCli_Exclui_Cliente(ByVal lCodCliente As Long) As Long
'Exclui os registros da tabela MvPerCli associados ao Cliente cujo código foi passado.
'Chamada por Cliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lComando(2) As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tMvPerCli As typeMvPerCli
Dim iIndice As Integer
    
On Error GoTo Erro_MvPerCli_Exclui_Cliente

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 6884
    Next
    
    'Seleciona os Movimentos do Cliente a serem excluídos
    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa, Exercicio, Filial FROM MvPerCli WHERE Cliente =?", 0, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, tMvPerCli.iFilial, lCodCliente)
    If lErro <> AD_SQL_SUCESSO Then Error 6885

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6886

    'Exclui os Movimentos do Cliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6887
        
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM MvPerCli", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6888

        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6889

    Loop
    
    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    MvPerCli_Exclui_Cliente = SUCESSO

    Exit Function

Erro_MvPerCli_Exclui_Cliente:

    MvPerCli_Exclui_Cliente = Err

    Select Case Err

        Case 6884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 6885, 6886, 6889
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", Err)

        Case 6887
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCLI", Err, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, lCodCliente, tMvPerCli.iFilial)

        Case 6888
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCLI", Err, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, lCodCliente, tMvPerCli.iFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148855)

    End Select
    
    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function CategoriaCliente_Grava(ByVal objCategoriaCliente As ClassCategoriaCliente, ByVal colItensCategoria As Collection) As Long
'Inclui ou altera uma categoria de clientes e seus valores
'tabelas: CategoriaCliente e CategoriaClienteItem

Dim lErro As Long, iIndice As Integer, colItensCategoriaCopia As New Collection
Dim sCategoriaDescricao As String, sCategoriaItem As String, sCategoriaItemDescricao As String
Dim objCategoriaClienteItem As New ClassCategoriaClienteItem
Dim lTransacao As Long, alComando(1 To 6) As Long

On Error GoTo Erro_CategoriaCliente_Grava

    'Cria uma cópia "de trabalho" da coleção passada como parâmetro
    For Each objCategoriaClienteItem In colItensCategoria
        colItensCategoriaCopia.Add objCategoriaClienteItem
    Next

    'Abertura comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 28858

    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28859

    sCategoriaDescricao = String(STRING_CATEGORIACLIENTE_DESCRICAO, 0)

    'Pesquisa no BD a categoria em questão
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Descricao FROM CategoriaCliente WHERE Categoria = ?", 0, sCategoriaDescricao, objCategoriaCliente.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28860

    'Lê a categoria, se estiver no BD
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28861

    'Se a categoria existir
    If lErro = AD_SQL_SUCESSO Then

        'Se a descrição da categoria for diferente
        If sCategoriaDescricao <> objCategoriaCliente.sDescricao Then

             'Alterar a descrição da categoria
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CategoriaCliente SET Descricao = ?", alComando(1), objCategoriaCliente.sDescricao)
            If lErro <> AD_SQL_SUCESSO Then Error 28862

        End If

        sCategoriaItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
        sCategoriaItemDescricao = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)

        'Percorrer todos os itens atuais da categoria no BD
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Item, Descricao FROM CategoriaClienteItem WHERE Categoria = ?", 0, sCategoriaItem, sCategoriaItemDescricao, objCategoriaCliente.sCategoria)
        If lErro <> AD_SQL_SUCESSO Then Error 28863

        'Lê o item da categoria, se estiver no BD
        lErro = Comando_BuscarProximo(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28864

        Do While lErro <> AD_SQL_SEM_DADOS

            'Procura o item na coleção correspondente ao grid
            iIndice = CategoriaCliente_Grava_ItemNaColecao(colItensCategoriaCopia, sCategoriaItem)

            'Se (ainda) existir no grid uma linha com a mesma chave
            If iIndice <> 0 Then

                Set objCategoriaClienteItem = colItensCategoriaCopia.Item(iIndice)

                'Se a descrição do grid for <> da que estiver no BD
                If objCategoriaClienteItem.sDescricao <> sCategoriaItemDescricao Then

                    'Alterar a descrição do item
                    lErro = Comando_ExecutarPos(alComando(4), "UPDATE CategoriaClienteItem SET Descricao = ?", alComando(3), objCategoriaClienteItem.sDescricao)
                    If lErro <> AD_SQL_SUCESSO Then Error 28865

                End If

                'Excluir o item da coleção
                colItensCategoriaCopia.Remove (iIndice)

            Else

                lErro = CategoriaClienteItem_NaoUtilizado(objCategoriaCliente.sCategoria, sCategoriaItem)
                If lErro Then Error 28866

                'Excluir o item do BD
                lErro = Comando_ExecutarPos(alComando(5), "DELETE FROM CategoriaClienteItem", alComando(3))
                If lErro <> AD_SQL_SUCESSO Then Error 28867

            End If

            'Lê o item da categoria, se estiver no BD
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28868

        Loop

    Else
        'A Categoria não existe

        'Criá-la (inclui-la em CategoriaCliente)
        lErro = Comando_Executar(alComando(6), "INSERT INTO CategoriaCliente (Categoria,Descricao) VALUES(?,?)", objCategoriaCliente.sCategoria, objCategoriaCliente.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then Error 28869

    End If

    'Grava os itens que ainda não faziam parte da categoria
    lErro = CategoriaCliente_Grava_NovosItens(objCategoriaCliente, colItensCategoriaCopia)
    If lErro <> SUCESSO Then Error 28870

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28871

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    CategoriaCliente_Grava = SUCESSO

    Exit Function

Erro_CategoriaCliente_Grava:

    CategoriaCliente_Grava = Err

    Select Case Err

        Case 28858
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28859
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28860, 28861
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTE1", Err, objCategoriaCliente.sCategoria)

        Case 28862
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CATEGORIACLIENTE", Err, objCategoriaCliente.sCategoria)

        Case 28863, 28864, 28868
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM", Err, objCategoriaCliente.sCategoria)

        Case 28865
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CATEGORIACLIENTEITEM", Err, objCategoriaCliente.sCategoria)

        Case 28866, 28870

        Case 28867
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CATEGORIACLIENTEITEM", Err, objCategoriaCliente.sCategoria)

        Case 28869
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CATEGORIACLIENTE", Err, objCategoriaCliente.sCategoria)

        Case 28871
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148856)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function CategoriaCliente_Grava_NovosItens(ByVal objCategoriaCliente As ClassCategoriaCliente, ByVal colItensCategoria As Collection) As Long
'Percorrer os itens na colecao incluindo-os no BD

Dim lErro As Long, lComando As Long
Dim objCategoriaItem As New ClassCategoriaClienteItem

On Error GoTo Erro_CategoriaCliente_Grava_NovosItens

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28873

    'Percorrer a coleção
    For Each objCategoriaItem In colItensCategoria

        lErro = Comando_Executar(lComando, "INSERT INTO CategoriaClienteItem (Categoria, Item, Descricao) VALUES(?,?,?)", objCategoriaItem.sCategoria, objCategoriaItem.sItem, objCategoriaItem.sDescricao)
        If lErro <> AD_SQL_SUCESSO Then Error 28874

    Next

    'Fechamento comando
    Call Comando_Fechar(lComando)

    CategoriaCliente_Grava_NovosItens = SUCESSO

    Exit Function

Erro_CategoriaCliente_Grava_NovosItens:

    CategoriaCliente_Grava_NovosItens = Err

    Select Case Err

        Case 27119
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28874
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CATEGORIACLIENTEITEM", Err, objCategoriaItem.sCategoria)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148857)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CategoriaCliente_Exclui(ByVal objCategoriaCliente As ClassCategoriaCliente) As Long
'Excluir a tabela com todos os seus itens se não for usada

Dim lErro As Long
Dim sCategoriaDescricao As String
Dim lTransacao As Long
Dim lComandoSel As Long
Dim lComandoDel As Long
Dim lComando As Long
Dim iTipo As Integer

On Error GoTo Erro_CategoriaCliente_Exclui

    'Abertura comandos
    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then Error 28992

    lComandoDel = Comando_Abrir()
    If lComandoDel = 0 Then Error 28993

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28979

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28994

    sCategoriaDescricao = String(STRING_CATEGORIACLIENTE_DESCRICAO, 0)

    'Pesquisa no BD a Categoria
    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Descricao FROM CategoriaCliente WHERE Categoria = ?", 0, sCategoriaDescricao, objCategoriaCliente.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28995

    'Verifica se Categoria existe
    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28996

    'Não encontrou a Categoria ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 28897

    '"Lock" na Categoria
    lErro = Comando_LockExclusive(lComandoSel)
    If lErro <> AD_SQL_SUCESSO Then Error 28898

    'Testa uso, nas tabelas que têm relacionamentos com CategoriaCliente .
    lErro = CategoriaCliente_NaoUtilizado(objCategoriaCliente.sCategoria)
    If lErro <> SUCESSO Then Error 19327
    
    'Exclui os itens da categoria
    lErro = CategoriaCliente_ExcluiItens(objCategoriaCliente)
    If lErro <> SUCESSO Then Error 28900

    'Exclui a própria categoria
    lErro = Comando_ExecutarPos(lComandoDel, "DELETE FROM CategoriaCliente", lComandoSel)
    If lErro <> AD_SQL_SUCESSO Then Error 28901

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28902

    'Libera comandos
    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoDel)

    CategoriaCliente_Exclui = SUCESSO

    Exit Function

Erro_CategoriaCliente_Exclui:

    CategoriaCliente_Exclui = Err

    Select Case Err
        
        Case 19327
            
        Case 28992, 28993, 28979
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28994
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28995, 28996
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTE1", Err, objCategoriaCliente.sCategoria)

        Case 28897
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_NAO_CADASTRADA", Err, objCategoriaCliente.sCategoria)

        Case 28898
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIACLIENTE", Err, objCategoriaCliente.sCategoria)

        Case 28900

        Case 28901
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CATEGORIACLIENTE", Err, objCategoriaCliente.sCategoria)

        Case 28902
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148858)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoDel)

    Exit Function

End Function

Private Function CategoriaCliente_NaoUtilizado(ByVal sCategoria As String) As Long
'Retorna SUCESSO se não houver referências no bd sCategoria em outras tabelas (como Clientes, TiposDeCliente,...)

Dim lErro As Long
Dim lComando As Long
Dim lCliente As Long
Dim sCategoriaCliente As String
Dim iTipoCliente As Integer

On Error GoTo Erro_CategoriaCliente_NaoUtilizado

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28980

    'Pesquisa se a categoria está associado a algum Cliente(em FilialClienteCategorias)
    lErro = Comando_Executar(lComando, "SELECT Cliente FROM FilialClienteCategorias WHERE Categoria = ? ", lCliente, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28981

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28982

    'Se está associado a algum cliente => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 28899

    'Inicializando a String com o número de caracteres relativo ao do BD
    sCategoriaCliente = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    
    'Verifica se CategoriaCliente é usada na Tabela ICMSExcecoes.
    lErro = Comando_Executar(lComando, "SELECT CategoriaCliente FROM ICMSExcecoes WHERE CategoriaCliente = ?", sCategoriaCliente, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 19345
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19346
    
    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19347

    'Pesquisa CategoriaCliente é usada na tabela TipoDeClienteCategorias
    lErro = Comando_Executar(lComando, "SELECT TipoDeCliente FROM TipoDeClienteCategorias WHERE Categoria = ?", iTipoCliente, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 19348

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19349

    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19350
    
    'Verifica se CategoriaCliente é usada na Tabela IPIExcecoes.
    lErro = Comando_Executar(lComando, "SELECT CategoriaCliente FROM IPIExcecoes WHERE CategoriaCliente = ? ", sCategoriaCliente, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 19351
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19352
    
    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19353
    
    'Pesquisa se a categoria está associado a alguma PadrõesTribSaída
    lErro = Comando_Executar(lComando, "SELECT CategoriaFilialCliente FROM  PadroesTribSaida WHERE CategoriaFilialCliente = ? ", sCategoriaCliente, sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 58125

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58126

    'Se está associado a algum cliente => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 58127
    
    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    CategoriaCliente_NaoUtilizado = SUCESSO

    Exit Function

Erro_CategoriaCliente_NaoUtilizado:

    CategoriaCliente_NaoUtilizado = Err

    Select Case Err

        Case 19345, 19346
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMSEXCECOES", Err)
        
        Case 19347
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_ICMSEXCECOES", Err, sCategoria)
            
        Case 19348, 19349
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTECATEGORIAS", Err)
                        
        Case 19350
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_TIPOCLIENTECATEGORIAS", Err, sCategoria)
            
        Case 19351, 19352
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPIEXCECOES", Err)
                    
        Case 19353
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_IPIEXCECOES", Err, sCategoria)
        
        Case 28981, 28982
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS1", Err, sCategoria)
        
        Case 29890
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28899
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_UTILIZADA", Err, sCategoria, lCliente)
        
        Case 58125, 58126
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOTRIBSAIDA", Err)
        
        Case 58127
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTE_PADROESTRIBSAIDA", Err, sCategoriaCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148859)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function CategoriaCliente_ExcluiItens(ByVal objCategoriaCliente As ClassCategoriaCliente) As Long
'Exclui os itens de uma categoria

Dim lErro As Long, lComandoSel As Long, lComandoDel As Long
Dim sCategoriaItem As String, sCategoriaItemDescricao As String

On Error GoTo Erro_CategoriaCliente_ExcluiItens

    'Abertura comandos
    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then Error 28907

    lComandoDel = Comando_Abrir()
    If lComandoDel = 0 Then Error 28908

    sCategoriaItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
    sCategoriaItemDescricao = String(STRING_CATEGORIACLIENTEITEM_DESCRICAO, 0)

    'Percorrer todos os itens da categoria no BD
    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Item, Descricao FROM CategoriaClienteItem WHERE Categoria = ?", 0, sCategoriaItem, sCategoriaItemDescricao, objCategoriaCliente.sCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 28909

    'Lê o item da categoria, se estiver no BD
    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28910

    Do While lErro <> AD_SQL_SEM_DADOS

        'Excluir o item do BD
        lErro = Comando_ExecutarPos(lComandoDel, "DELETE FROM CategoriaClienteItem", lComandoSel)
        If lErro <> AD_SQL_SUCESSO Then Error 28912

        'Lê o item da categoria, se estiver no BD
        lErro = Comando_BuscarProximo(lComandoSel)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28913

    Loop

    'Libera comandos
    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoDel)

    CategoriaCliente_ExcluiItens = SUCESSO

    Exit Function

Erro_CategoriaCliente_ExcluiItens:

    CategoriaCliente_ExcluiItens = Err

    Select Case Err

        Case 28907, 28908
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28909, 28910, 28913
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM", Err, objCategoriaCliente.sCategoria)

        Case 28911

        Case 28912
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CATEGORIACLIENTEITEM", Err, objCategoriaCliente.sCategoria)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148860)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoDel)

    Exit Function

End Function

Private Function CategoriaClienteItem_NaoUtilizado(ByVal sCategoria As String, ByVal sItem As String) As Long
'Retorna SUCESSO se não houver referências no bd a este par (sCategoria, sItem) em outras tabelas (como Clientes, TiposDeCliente,...)

Dim lErro As Long, lComando As Long, lCliente As Long

Dim sCategoriaCliente As String
Dim iTipoCliente As Integer

On Error GoTo Erro_CategoriaClienteItem_NaoUtilizado

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28872

    'Pesquisa se o par (categoria,item) está associado a algum Cliente
    lErro = Comando_Executar(lComando, "SELECT Cliente FROM FilialClienteCategorias WHERE Categoria = ? AND Item = ?", lCliente, sCategoria, sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 28875

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28876

    'Se está associado a algum cliente => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 28877

    'Inicializando a String com o número de caracteres relativo ao do BD
    sCategoriaCliente = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    
    'Verifica se Categoria e/ou Item são usados na Tabela ICMSExcecoes.
    lErro = Comando_Executar(lComando, "SELECT CategoriaCliente FROM ICMSExcecoes WHERE CategoriaCliente = ? AND CategoriaClienteItem = ? ", sCategoriaCliente, sCategoria, sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 19335
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19336
    
    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19337

    'Pesquisa se o par (categoria,item) está associado a algum registro da tabela TipoDeClienteCategorias
    lErro = Comando_Executar(lComando, "SELECT TipoDeCliente FROM TipoDeClienteCategorias WHERE Categoria = ? AND Item = ?", iTipoCliente, sCategoria, sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 19339

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19340

    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19341
    
    'Verifica se Categoria e/ou Item são usados na Tabela IPIExcecoes.
    lErro = Comando_Executar(lComando, "SELECT CategoriaCliente FROM IPIExcecoes WHERE CategoriaCliente = ? AND CategoriaClienteItem = ?", sCategoriaCliente, sCategoria, sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 19342
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19343
    
    'Se está sendo usado => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 19344

    'Pesquisa se o par (categoria,item) está associado a PadroesTribSaida
    lErro = Comando_Executar(lComando, "SELECT CategoriaFilialCliente FROM PadroesTribSaida WHERE CategoriaFilialCliente = ? AND ItemCategoria = ?", sCategoriaCliente, sCategoria, sItem)
    If lErro <> AD_SQL_SUCESSO Then Error 58130

    'Verifica se existe
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58131

    'Se está associado a algum cliente => erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 58132

    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    CategoriaClienteItem_NaoUtilizado = SUCESSO

    Exit Function

Erro_CategoriaClienteItem_NaoUtilizado:

    CategoriaClienteItem_NaoUtilizado = Err

    Select Case Err

        Case 19335, 19336
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ICMSEXCECOES", Err)
        
        Case 19337
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_ICMSEXCECOES", Err, sCategoria, sItem)
            
        Case 19339, 19340
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTECATEGORIAS", Err)
                        
        Case 19341
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_TIPOCLIENTECATEGORIAS", Err, sCategoria, sItem)
            
        Case 19342, 19343
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IPIEXCECOES", Err)
                    
        Case 19344
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_IPIEXCECOES", Err, sCategoria, sItem)
        
        Case 19338, 28872
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28875, 28876
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS", Err, sCategoria, sItem)

        Case 28877
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_UTILIZADA", Err, lCliente, sItem, sCategoria)
        
        Case 58130, 58131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAOTRIBSAIDA", Err)
        
        Case 58132
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_PADROESTRIBSAIDA", Err, sCategoria, sItem)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148861)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function CategoriaCliente_Grava_ItemNaColecao(ByVal colItensCategoria As Collection, ByVal sItem As String) As Long
'Retorna o índice do item na coleção ou zero se não achá-lo (na coleção)

Dim iIndice As Integer
Dim objCategoriaItem As New ClassCategoriaClienteItem

    'Pesquisa o item na coleção
    For iIndice = 1 To colItensCategoria.Count

        Set objCategoriaItem = colItensCategoria.Item(iIndice)

        'Se achou o item na coleção
        If objCategoriaItem.sItem = sItem Then

            CategoriaCliente_Grava_ItemNaColecao = iIndice

            Exit Function

        End If

    Next

    CategoriaCliente_Grava_ItemNaColecao = SUCESSO

End Function

Function FiliaisClientes_Grava(ByVal objFilialCliente As ClassFilialCliente, ByVal colEndereco As Collection) As Long
'Grava a Filial Cliente e seus 3 tipos de Enderecos

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_FiliaisClientes_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 12624

    lErro = FiliaisClientes_Grava_EmTrans(objFilialCliente, colEndereco)
    If lErro <> SUCESSO Then Error 59486
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 12633

    FiliaisClientes_Grava = SUCESSO

    Exit Function

Erro_FiliaisClientes_Grava:

    FiliaisClientes_Grava = Err

    Select Case Err

        Case 12624
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 12633
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", Err)
            
        Case 59486
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148862)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NFiscal_Automatico(lNumIntNF As Long) As Long
'Retorna o número do próximo NumIntNF disponivel

    NFiscal_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_NOTA_FISCAL", lNumIntNF)
    
End Function

Function NFiscalItem_Automatico(lNumIntNFItem As Long) As Long
'Retorna o número do próximo NumIntNFItem disponivel

    NFiscalItem_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_ITEM_NOTA_FISCAL", lNumIntNFItem)
    
End Function

Function CategoriaClienteItem_Lock(ByVal sCategoria As String, ByVal sItemCategoria As String, ByVal lComando As Long) As Long
'faz lock shared de uma CategoriaClienteItem para garantir a sua existencia

Dim sCategoriaItem As String
Dim lErro As Long

On Error GoTo Erro_CategoriaClienteItem_Lock

    sCategoriaItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)
                
    'Pesquisa a Item da CategoriaCliente
    lErro = Comando_ExecutarLockado(lComando, "SELECT Item FROM CategoriaClienteItem WHERE Categoria = ? AND Item = ?", sCategoriaItem, sCategoria, sItemCategoria)
    If lErro <> AD_SQL_SUCESSO Then Error 40553

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40554

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 40555

    '"Lock" da CategoriaClienteItem
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 40556
    
    CategoriaClienteItem_Lock = SUCESSO
    
    Exit Function
    
Erro_CategoriaClienteItem_Lock:

    CategoriaClienteItem_Lock = Err
    
    Select Case Err
    
        Case 40553, 40554
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM", Err, sCategoria)
        
        Case 40555
        
        Case 40556
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIACLIENTEITEM2", Err, sCategoria, sItemCategoria)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148863)

    End Select

    Exit Function
    
End Function

Function TipoDocInfo_Lock_Sigla(ByVal sSiglaMovto As String, ByVal lComando As Long) As Long
'faz lock shared de um tipodocinfo pela sigla para garantir a sua existencia

Dim lErro As Long
Dim sSigla As String

On Error GoTo Erro_TipoDocInfo_Lock_Sigla

    sSigla = String(STRING_TIPO_DOC_INFO_SIGLA, 0)
            
    'Pesquisa o Tipo de Documento
    lErro = Comando_ExecutarLockado(lComando, "SELECT Sigla FROM TiposDocInfo WHERE Sigla = ?", sSigla, sSiglaMovto)
    If lErro <> AD_SQL_SUCESSO Then Error 40567

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40568

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 40569

    '"Lock" do Tipo Documento
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 40570

    TipoDocInfo_Lock_Sigla = SUCESSO
    
    Exit Function
    
Erro_TipoDocInfo_Lock_Sigla:

    TipoDocInfo_Lock_Sigla = Err
    
    Select Case Err
    
        Case 40567, 40568
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODOCINFO", Err, sSiglaMovto)

        Case 40569
            
        Case 40570
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDOCINFO", Err, sSiglaMovto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148864)

    End Select

    Exit Function
    
End Function

Function NaturezaOperacao_Lock(ByVal sNaturezaOperacao As String, ByVal lComando As Long) As Long
'faz lock shared de uma natureza de operacao para garantir a sua existencia

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_NaturezaOperacao_Lock

    sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)

    'Pesquisa a Natureza da Operação
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM NaturezaOP WHERE Codigo = ?", sCodigo, sNaturezaOperacao)
    If lErro <> AD_SQL_SUCESSO Then Error 40549

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40550

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 40551

    '"Lock" da NaturezaOP
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 40552
    
    NaturezaOperacao_Lock = SUCESSO
    
    Exit Function
    
Erro_NaturezaOperacao_Lock:

    NaturezaOperacao_Lock = Err

    Select Case Err

        Case 40549, 40550
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, sNaturezaOperacao)
        
        Case 40551
        
        Case 40552
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", Err, sNaturezaOperacao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148865)

    End Select

    Exit Function
    
End Function

Function NaturezaOperacao_Grava(ByVal objNaturezaOperacao As ClassNaturezaOp) As Long
'Grava a Natureza de Operação no banco de dados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NaturezaOperacao_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 17971
    
    lErro = NaturezaOperacao_Grava_EmTrans(objNaturezaOperacao)
    If lErro <> SUCESSO Then gError 97703

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 17976

    NaturezaOperacao_Grava = SUCESSO

    Exit Function

Erro_NaturezaOperacao_Grava:

    NaturezaOperacao_Grava = gErr

    Select Case gErr

        Case 97703
        
        Case 17969, 17970
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 17976
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148866)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function NaturezaOp_Testa_Vinculos(ByVal sCodigo As String) As Long
'Testa para cada tabela relacionada se a Natureza está sendo utilizada.

Dim lErro As Long
Dim iFilialPedVenda As Integer
Dim lCodPedVenda As Long
Dim iFilialPedVendaBaixado As Integer
Dim lCodPedVendaBaixado As Long
Dim iFilialNF As Integer
Dim lNumNF As Long
Dim sSerieNF As String
Dim sSigla As String

On Error GoTo Erro_NaturezaOp_Testa_Vinculos
        
    'Testa vinculo com Pedido de Venda
    lErro = CF("PedidosDeVenda_Le_NaturezaOp", sCodigo, iFilialPedVenda, lCodPedVenda)
    If lErro <> SUCESSO And lErro <> 49870 Then Error 49917
    
    'se encontrou ----> Erro
    If lErro = 49870 Then Error 49929
    
    'Testa Vinculo com Pedido de Venda Baixados
    lErro = CF("PedidosDeVendaBaixado_Le_NaturezaOp", sCodigo, iFilialPedVendaBaixado, lCodPedVendaBaixado)
    If lErro <> SUCESSO And lErro <> 49873 Then Error 49918
    
    'se encontrou -----> Erro
    If lErro = 49873 Then Error 49930
    
    'Testa Vinculo com NFiscal
    lErro = CF("NFiscal_Le_NaturezaOp", sCodigo, iFilialNF, lNumNF, sSerieNF)
    If lErro <> SUCESSO And lErro <> 49876 Then Error 49919
    
    'se encontrou -----> Erro
    If lErro = 49876 Then Error 49931
    
    'Testa Vinculo com PadraoTribSaida
    lErro = CF("PadraoTribSaida_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49881 Then Error 49921
    
    'se encontrou -----> Erro
    If lErro = 49881 Then Error 49933
    
    'Testa Vinculo com PadraoTribEntrada
    lErro = CF("PadraoTribEntrada_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49884 Then Error 49922
    
    'se encontrou -----> Erro
    If lErro = 49884 Then Error 49934
    
    'Testa Vinculo com Tipo Doc Info
    lErro = CF("TipoDocInfo_Le_NaturezaOp", sCodigo, sSigla)
    If lErro <> SUCESSO And lErro <> 49887 Then Error 49923
    
    'se encontrou -----> Erro
    If lErro = 49887 Then Error 49935
    
    'testa vinculo com Tributacao NF
    lErro = CF("TributacaoNF_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49890 Then Error 49924
    
    'se encontrou -----> Erro
    If lErro = 49890 Then Error 49936
    
    'Testa vinculo com Tributacao Item PV
    lErro = CF("TributacaoItemPV_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49903 Then Error 49925
    
    'se encontrou -----> Erro
    If lErro = 49903 Then Error 49937
    
    'Testa vinculo com Tributacao ComplNF
    lErro = CF("TributacaoComplNF_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49895 Then Error 49926
    
    'se encontrou -----> Erro
    If lErro = 49895 Then Error 49938
    
    'Testa vinculo com TributacaoComplPV
    lErro = CF("TributacaoComplPV_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49848 Then Error 49927
    
    'se encontrou -----> Erro
    If lErro = 49848 Then Error 49939
    
    'Testa vinculo com TributacaoItemNF
    lErro = CF("TributacaoItemNF_Le_NaturezaOp", sCodigo)
    If lErro <> SUCESSO And lErro <> 49901 Then Error 49928
    
    'se encontrou -----> Erro
    If lErro = 49901 Then Error 49940
    
    NaturezaOp_Testa_Vinculos = SUCESSO
    
    Exit Function

Erro_NaturezaOp_Testa_Vinculos:

NaturezaOp_Testa_Vinculos = Err

    Select Case Err
        
        Case 49917, 49918, 49919, 49920, 49921, 49922
        
        Case 49923, 49924, 49925, 49926, 49927, 49928
            
        Case 49929
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_PEDIDODEVENDA", Err, sCodigo, lCodPedVenda, iFilialPedVenda)
        
        Case 49930
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_PEDIDODEVENDABAIXADO", Err, sCodigo, lCodPedVendaBaixado, iFilialPedVendaBaixado)
        
        Case 49931
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_NFISCAL", Err, sCodigo, sSerieNF, lNumNF, iFilialNF)
        
        Case 49932
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_NFISCALBAIXADA", Err, sCodigo)
        
        Case 49933
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_PADRAOTRIBSAIDA", Err, sCodigo)
            
        Case 49934
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_PADRAOTRIBENTRADA", Err, sCodigo)
        
        Case 49935
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TIPODOCINFO", Err, sCodigo, sSigla)
        
        Case 49936
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TRIBUTACAONF", Err, sCodigo)
        
        Case 49937
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMPV", Err, sCodigo)
            
        Case 49938
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLNF", Err, sCodigo)
            
        Case 49939
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TRIBUTACAOCOMPLPV", Err, sCodigo)
            
        Case 49940
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_USADO_TRIBUTACAOITEMNF", Err, sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148867)
    End Select
            
    Exit Function
    
End Function

Function NaturezaOperacao_Exclui(ByVal sCodigo As String) As Long
'Exclui a Natureza de Operação do BD

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim sCodigo1 As String

On Error GoTo Erro_NaturezaOperacao_Exclui

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 17982

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 17983

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 17984

    sCodigo1 = String(STRING_NATUREZAOP_CODIGO, 0)

    'Pesquisa a Natureza de Operação no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM NaturezaOP WHERE Codigo = ?", 0, sCodigo1, sCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 17985

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 17986

    If lErro = AD_SQL_SEM_DADOS Then Error 17987
    
    'Lock da Natureza de Operação
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 6158
    
    'Verifica se tem relacao com alguma Tabela
    lErro = NaturezaOp_Testa_Vinculos(sCodigo)
    If lErro <> SUCESSO Then Error 49905
    
    'Exclui a natureza de operação do BD
    lErro = Comando_ExecutarPos(lComando1, "DELETE FROM NaturezaOp", lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 17989

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 17990

    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando1)

    NaturezaOperacao_Exclui = SUCESSO

    Exit Function

Erro_NaturezaOperacao_Exclui:

    NaturezaOperacao_Exclui = Err

    Select Case Err

        Case 17982, 17983
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17984
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 17985, 17986
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", Err, sCodigo)

        Case 17987
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, sCodigo)

        Case 17988
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", Err, sCodigo)

        Case 17989
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_NATUREZAOP", Err, sCodigo)

        Case 17990
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 49905
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148868)

    End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TituloReceber_Grava_Fatura(ByVal objTituloReceber As ClassTituloReceber, ByVal colNFiscalMarcado As Collection, ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal colcolDesconto As colcolDesconto, ByVal objContabil As ClassContabil) As Long
'Grava o titulo referente a uma fatura

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1) As Long
Dim lNumIntDoc As Long
Dim sNome As String
Dim colComissaoEmissao As New colComissao
Dim iIndice As Integer, dValorLivroAux As Double

On Error GoTo Erro_TituloReceber_Grava_Fatura

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36744
    Next
    
    'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 36745
    
    sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)

    'Procura no BD a FilialCliente fornecida
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Nome FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ? ", sNome, objTituloReceber.lCliente, objTituloReceber.iFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 36746

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36747

    'Se não encontrar --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 36748

    'Faz lock na FilialCliente
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then Error 36749

    'Cria um Novo Título a Receber
    lErro = CF("TituloReceber_Insere", objTituloReceber, colParcelaReceber, colComissaoEmissao, colcolComissao)
    If lErro <> SUCESSO Then Error 36753

    'TituloReceber_Insere deve retornar o valor de NumInDoc do Titulo em objTituloReceber.lNumIntDoc
    lErro = NFiscal_Grava_NumIntDocCpr(TITULO_RECEBER, objTituloReceber.lNumIntDoc, colNFiscalMarcado)
    If lErro <> SUCESSO Then Error 36754

    dValorLivroAux = Round(objTituloReceber.dValor - (objTituloReceber.dValorIRRF + objTituloReceber.dPISRetido + objTituloReceber.dISSRetido + objTituloReceber.dCOFINSRetido + objTituloReceber.dCSLLRetido) - IIf(objTituloReceber.iINSSRetido <> 0, objTituloReceber.dValorINSS, 0), 2)
    
    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objTituloReceber.lNumIntDoc, objTituloReceber.lCliente, objTituloReceber.iFilial, DATA_NULA, , 0, , -dValorLivroAux)
    If lErro <> SUCESSO Then Error 36751

    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 36752
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    TituloReceber_Grava_Fatura = SUCESSO
    
    Exit Function

Erro_TituloReceber_Grava_Fatura:

    TituloReceber_Grava_Fatura = Err
    
    Select Case Err
    
        Case 36743, 36750, 36751, 36753, 36754
    
        Case 36744
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
    
        Case 36745
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
        Case 36746, 36747
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", Err, objTituloReceber.lCliente, objTituloReceber.iFilial)
        
        Case 36748
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objTituloReceber.lCliente, objTituloReceber.iFilial)
        
        Case 36749
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCLIENTES", Err, objTituloReceber.lCliente, objTituloReceber.iFilial)
    
        Case 36752
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148869)
            
    End Select

    'Fecha a transacao e os comandos
    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function NFiscal_Grava_NumIntDocCpr(ByVal iClasseDocCPR As Integer, ByVal lNumIntDocCPR As Long, ByVal colNFiscalInfo As Collection) As Long
'Atualiza o campo NumIntDocCpr das notas fiscais passadas na colecao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim objNFiscalInfo As ClassNFiscalInfo
Dim lNumIntDocCPR1 As Long
Dim sSerie As String
Dim lNumNotaFiscal As Long
Dim iFilialEmpresa As Integer
Dim iIndice As Integer

On Error GoTo Erro_NFiscal_Grava_NumIntDocCpr

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 36755

    Next

    For Each objNFiscalInfo In colNFiscalInfo
            
        sSerie = String(STRING_SERIE, 0)
        
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDocCPR, Serie, NumNotaFiscal, FilialEmpresa FROM NFiscal WHERE NumIntDoc = ?", 0, lNumIntDocCPR1, sSerie, lNumNotaFiscal, iFilialEmpresa, objNFiscalInfo.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 36756
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 36757
    
        If lErro = AD_SQL_SEM_DADOS Then Error 36758
        
        If lNumIntDocCPR1 <> 0 Then Error 36759
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFiscal SET NumIntDocCPR = ?, ClasseDocCPR = ?", alComando(1), lNumIntDocCPR, iClasseDocCPR)
        If lErro <> AD_SQL_SUCESSO Then Error 36760
        
    Next

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Grava_NumIntDocCpr = SUCESSO

    Exit Function

Erro_NFiscal_Grava_NumIntDocCpr:

    NFiscal_Grava_NumIntDocCpr = Err

    Select Case Err

        Case 36755
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 36756, 36757
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTA_FISCAL_NUM_SERIE", Err, sSerie, lNumNotaFiscal, iFilialEmpresa)

        Case 36758
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NUM_SERIE_NAO_CADASTRADA", Err, sSerie, lNumNotaFiscal, iFilialEmpresa)

        Case 36759
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_NUMINTDOCCPR_NAO_ZERO", Err, sSerie, lNumNotaFiscal, iFilialEmpresa)
        
        Case 36760
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", Err, sSerie, lNumNotaFiscal, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148870)

    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FilialCliente_Exclui(ByVal objFilialCliente As ClassFilialCliente) As Long
'Exclui a Filial Cliente e seus Enderecos

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lComando5 As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim lPedido As Long
Dim iFilial As Integer
Dim lCodigoCliente As Long
Dim iAtivo As Integer
Dim lCodRelacionamento As Long

On Error GoTo Erro_FilialCliente_Exclui

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 12692

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 12693
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 112750
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then gError 112751
    
    lComando5 = Comando_Abrir()
    If lComando5 = 0 Then gError 112752

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 12694

    'Seleciona a Filial Cliente a ser excluída
    lErro = Comando_ExecutarPos(lComando1, "SELECT CodCliente FROM FiliaisClientes WHERE CodCliente =? AND CodFilial = ?", 0, lCodigo, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 12695

    'Lê a Filial
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12696
    
    If lErro = AD_SQL_SEM_DADOS Then gError 12697

    'Faz lock da FilialCliente
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 12329

    'Verifica se FilialCliente está relacionada à PedidoDeVenda ou a PedidoVendaBaixado.
    lErro = CF("PedidoVenda_Le_FilialCliente", objFilialCliente, lPedido, iFilial)
    If lErro <> SUCESSO And lErro <> 33947 Then gError 33948
    
    'A Filial do Cliente está relacionada com Pedido de Venda --> erro
    If lErro = SUCESSO Then gError 33949

    'Verifica se existem Notas Fiscais relacionadas à FilialCliente
    lErro = CF("NFs_FilialCliente_Nao_Relacionada", objFilialCliente)
    If lErro <> SUCESSO Then gError 6905

    'Verifica se existem Títulos a Receber (lançados ou baixadas) relacionados à FilialCliente
    lErro = CF("TitulosReceber_FilialCliente_Nao_Relacionada", objFilialCliente)
    If lErro <> SUCESSO Then gError 6906

    'Verifica se existe Cheque Pré relacionado com Filial Cliente
    lErro = CF("ChequePre_Le_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO Then gError 58139
    
    'Verifica se existem DebitosReceber relacionados com a FilialCliente
    lErro = CF("DebitosRecCli_Le_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12685 Then gError 12665

    'Verifica se FilialCliente está relacionada com Débitos
    If lErro = SUCESSO Then gError 12666

    'Preenche a coleção de códigos de Recebimentos Antecipados relacionados com a FilialCliente
    lErro = CF("RecebAntecipados_Le_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO And lErro <> 12690 Then gError 12667

    'Verifica se FilialCliente está relacionada com Recebimentos Antecipados
    If lErro = SUCESSO Then gError 12668

    'Inserido por Luiz Nogueira em 27/10/03
    'Verifica se existe algum relacionamento de cliente vinculado a essa filial do cliente
    lErro = CF("RelacionamentoClientes_Verifica_Cliente", objFilialCliente.lCodCliente, lCodRelacionamento, objFilialCliente.iCodFilial)
    If lErro <> SUCESSO And lErro <> 102694 Then gError 102714
    
    'Se existir => erro
    If lErro = SUCESSO Then gError 102715

    'Exclui registros de MvDiaCli relacionados à FilialCliente. Não deverá haver nenhum. Só para dar robustez na integridade referencial
    lErro = CF("MvDiaCli_Exclui_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO Then gError 6907

    'Exclui registros de MvPerCli relacionados à FilialCliente. Não deverá haver nenhum. Só para dar robustez na integridade referencial
    lErro = CF("MvPerCli_Exclui_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO Then gError 6908

    'Exclui as Categorias das Filiais do Cliente
    lErro = FilialCliente_Exclui_Categorias(objFilialCliente)
    If lErro <> SUCESSO Then gError 33103
    
    'Incluído por Luiz Nogueira em 27/10/03
    'Exclui os contatos relacionados ao cliente
    lErro = CF("ClienteContatos_Exclui_Cliente_EmTrans", objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> SUCESSO Then gError 102717
    
    lErro = CF("RotasPontos_Testa_Filial", objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> SUCESSO Then gError 102717
      
    'Exclui a Filial Cliente
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM FiliaisClientes", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 12698
    
    'tenta achar outra filial q esteja ativa--> se não achar, desativa o cliente
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        lErro = Comando_Executar(lComando3, "SELECT Ativo FROM FiliaisClientes WHERE CodClienteLoja = ? AND Ativo = ?", 0, iAtivo, objFilialCliente.lCodClienteLoja, CLIENTE_ATIVO)
    Else
        lErro = Comando_ExecutarPos(lComando3, "SELECT Ativo FROM FiliaisClientes WHERE CodCliente = ? AND Ativo = ?", 0, Ativo, objFilialCliente.lCodCliente, CLIENTE_ATIVO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 112745

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112746
    
    If lErro = AD_SQL_SEM_DADOS Then
        'Seleciona o Clientes no BD
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            lCodigoCliente = objFilialCliente.lCodClienteLoja
            lErro = Comando_ExecutarPos(lComando4, "SELECT Ativo FROM Clientes WHERE CodigoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialCliente.lCodClienteLoja, CLIENTE_ATIVO)
        Else
            lCodigoCliente = objFilialCliente.lCodCliente
            lErro = Comando_ExecutarPos(lComando4, "SELECT Ativo FROM Clientes WHERE Codigo = ? AND Ativo = ?", 0, Ativo, objFilialCliente.lCodCliente, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 112747

        lErro = Comando_BuscarPrimeiro(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112748
        
        If lErro = AD_SQL_SUCESSO Then
            'Atualiza a tabela de Clientes
            lErro = Comando_ExecutarPos(lComando5, "UPDATE Clientes SET Ativo = ?", lComando4, DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 112749
        End If
    End If
    
    'Exclui os Endereços da FilialCliente (principal, entrega e cobrança)
    lErro = CF("Enderecos_Exclui_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO Then gError 6910

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 12699

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)

    FilialCliente_Exclui = SUCESSO

    Exit Function

Erro_FilialCliente_Exclui:

    FilialCliente_Exclui = gErr

    Select Case gErr

        Case 6905, 6906, 6907, 6908, 6910, 12665, 12667, 33103, 33948, 58139, 102714, 102717
            'Tratados pela rotina chamada

        Case 12329
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCLIENTES", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12666
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_DEBITOS", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12668
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_RECEB_ANTEC", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12692, 12693, 112750, 112751, 112752
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 12694
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 12695, 12696, 112745, 112746
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12697
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12698
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCLIENTE", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 12699
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 33949
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_PED_VENDA", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, iFilial, lPedido)
        
        Case 112747, 112748
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
            
        Case 112749
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", gErr, lCodigoCliente)
            
        Case 102715
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_RELACIONAMENTOCLIENTE", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, lCodRelacionamento)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148871)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    
    Exit Function

End Function

Private Function FilialCliente_Exclui_Categorias(ByVal objFilialCliente As ClassFilialCliente) As Long
'Exclui as Categorias da Filial do Cliente cujo código da filial e do cliente foi passado como parâmetro.
'Chamada por FilialCliente_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim lComando(2) As Long
Dim iIndice As Integer
Dim iFilial As Integer
Dim iFim_de_Arquivo As Integer

On Error GoTo Erro_FilialCliente_Exclui_Categorias

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 33139
    Next

    'Seleciona as Categorias da Filial do Cliente a serem excluídas
    lErro = Comando_ExecutarPos(lComando(0), "SELECT Filial FROM FilialClienteCategorias WHERE Cliente =? AND Filial = ?", 0, iFilial, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 33140

    'Tenta ler
    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 33141

    'Exclui as Filiais do Cliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        'Exclui a Filial
        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM FilialClienteCategorias", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 33142

        'Traz a próxima Categoria da Filial do Cliente
        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 33143

    Loop

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    FilialCliente_Exclui_Categorias = SUCESSO

    Exit Function

Erro_FilialCliente_Exclui_Categorias:

    FilialCliente_Exclui_Categorias = Err

    Select Case Err

        Case 33139
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 33140, 33141, 33143
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCLIENTECATEGORIAS3", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 33142
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCLIENTECATEGORIAS1", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148872)

    End Select

    'Fecha os comandos
    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function TipoVendedor_Grava(ByVal objTipoVendedor As ClassTipoVendedor) As Long
'Atualiza ou insere um novo registro na tabela TiposDeVendedor

Dim lErro As Long
Dim iCodigo As Integer
Dim iVendedor As Integer
Dim lComando As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long

On Error GoTo Erro_TipoVendedor_Grava

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16228

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16229

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 16230

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16231

    'Pesquisa descrição dos outros Tipos de Vendedores no BD
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDeVendedor WHERE Codigo <> ? AND Descricao = ?", iVendedor, objTipoVendedor.iCodigo, objTipoVendedor.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 16232

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16233

    'Encontrou TipoVendedor com essa descrição
    If lErro = AD_SQL_SUCESSO Then Error 16234

    'Pesquisa Tipo de Vendedor no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM TiposDeVendedor WHERE Codigo = ? ", 0, iCodigo, objTipoVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16235

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16236

    If lErro = AD_SQL_SEM_DADOS Then
    'Tipo de Vendedor não pertence ao BD

        'Insere novo registro na tabela TiposDeVendedor
        lErro = Comando_Executar(lComando2, "INSERT INTO TiposDeVendedor (Codigo, Descricao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoSeguro, ComissaoICM, ComissaoIPI) VALUES (?,?,?,?,?,?,?,?,?,?)", objTipoVendedor.iCodigo, objTipoVendedor.sDescricao, objTipoVendedor.dPercComissao, objTipoVendedor.dPercComissaoBaixa, objTipoVendedor.dPercComissaoEmissao, objTipoVendedor.iComissaoSobreTotal, objTipoVendedor.iComissaoFrete, objTipoVendedor.iComissaoSeguro, objTipoVendedor.iComissaoICM, objTipoVendedor.iComissaoIPI)
        If lErro <> AD_SQL_SUCESSO Then Error 16237

    Else
    'Tipo de Vendedor existe no BD

        'Altera registro na tabela TiposDeVendedor
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TiposDeVendedor SET Descricao = ? , PercComissao = ? , PercComissaoBaixa = ? , PercComissaoEmissao = ? , ComissaoSobreTotal = ? , ComissaoFrete = ? , ComissaoSeguro = ? , ComissaoICM = ? , ComissaoIPI = ? ", lComando1, objTipoVendedor.sDescricao, objTipoVendedor.dPercComissao, objTipoVendedor.dPercComissaoBaixa, objTipoVendedor.dPercComissaoEmissao, objTipoVendedor.iComissaoSobreTotal, objTipoVendedor.iComissaoFrete, objTipoVendedor.iComissaoSeguro, objTipoVendedor.iComissaoICM, objTipoVendedor.iComissaoIPI)
        If lErro <> AD_SQL_SUCESSO Then Error 16238

    End If

    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16239

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    TipoVendedor_Grava = SUCESSO

    Exit Function

Erro_TipoVendedor_Grava:

    TipoVendedor_Grava = Err

        Select Case Err

        Case 16228, 16229, 16230
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16231
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16232, 16233, 16235, 16236
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16234
             Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TIPO_VENDEDOR_REPETIDA", Err, objTipoVendedor.iCodigo)

        Case 16237
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16238
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16239
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148873)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function TipoVendedor_Exclui(ByVal objTipoVendedor As ClassTipoVendedor) As Long
'exclui Tipo de Vendedor do BD

Dim lComando1 As Long
Dim lComando2 As Long
Dim lTransacao As Long
Dim lErro As Long
Dim sDescricao As String
Dim iCodVendedor As Integer

On Error GoTo Erro_TipoVendedor_Exclui

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 16251

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 16252

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 16253

    sDescricao = String(STRING_TIPO_DE_VENDEDOR_DESCRICAO, 0)

    'Pesquisa Tipo de vendedor no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Descricao FROM TiposDeVendedor WHERE Codigo = ? ", 0, sDescricao, objTipoVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16254

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16255

    If lErro = AD_SQL_SEM_DADOS Then Error 16256

    'Lock do Tipo de Vendedor
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 16257

    'Verifica se Tipo de Vendedor está relacionada com Vendedores
    lErro = CF("Vendedor_Le_Tipo", objTipoVendedor, iCodVendedor)
    If lErro <> SUCESSO And lErro <> 16250 Then Error 16243

    If lErro = SUCESSO Then Error 16244

    'Exclui Tipo de Vendedor da tabela TiposDeVendedor
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM TiposDeVendedor", lComando1)
    If lErro <> AD_SQL_SUCESSO Then Error 16258

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 16259

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    TipoVendedor_Exclui = SUCESSO

    Exit Function

Erro_TipoVendedor_Exclui:

    TipoVendedor_Exclui = Err

    Select Case Err

        Case 16243
            'Erro tratado na rotina chamada

        Case 16244
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEVENDEDOR_RELACIONADO_VENDEDOR", Err, objTipoVendedor.iCodigo, iCodVendedor)

        Case 16251, 16252
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16253
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 16254, 16255
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16256
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEVENDEDOR_NAO_CADASTRADO", Err, objTipoVendedor.iCodigo)

        Case 16257
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16258
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOSDEVENDEDOR", Err, objTipoVendedor.iCodigo)

        Case 16259
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148874)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Exit Function

End Function

Function Vendedor_Grava(ByVal objVendedor As ClassVendedor, ByVal objEndereco As ClassEndereco) As Long
'Atualiza ou insere um novo registro nas tabelas Vendedores

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Vendedor_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 16668

    lErro = CF("Vendedor_Grava_EmTrans", objVendedor, objEndereco)
    If lErro <> SUCESSO Then gError 76390
    
    'Confirma a Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 16686

    Vendedor_Grava = SUCESSO

    Exit Function

Erro_Vendedor_Grava:

    Vendedor_Grava = gErr

        Select Case gErr

        Case 16668
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 16686
             Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
     
        Case 76390
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148875)

    End Select

    Call Transacao_Rollback
    
    Exit Function

End Function

Function Vendedor_Grava_EmTrans(ByVal objVendedor As ClassVendedor, ByVal objEndereco As ClassEndereco) As Long
'Atualiza ou insere um novo registro nas tabelas Vendedores
'Função alterada por Mauricio Maciel em 10/04/03

Dim lErro As Long
Dim iIndice As Integer
Dim lCodigo As Long
Dim iCodigo1 As Integer
Dim iVendedor As Integer
Dim lEndereco As Long
Dim alComando(5) As Long

On Error GoTo Erro_Vendedor_Grava_EmTrans

    For iIndice = LBound(alComando) To UBound(alComando)
    
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 16664

    Next
    
    'Pesquisa nome reduzido de outros Vendedores no BD
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Vendedores WHERE Codigo <> ? AND NomeReduzido = ?", iVendedor, objVendedor.iCodigo, objVendedor.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 16669

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16670

    If lErro <> AD_SQL_SEM_DADOS Then gError 16671
    
    If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then
    
        'Lock do Tipo Vendedor
        If objVendedor.iTipo <> 0 Then
            
            lErro = CF("TipoVendedor_Lock", alComando(1), objVendedor.iTipo)
            If lErro <> SUCESSO Then gError 16684
                                
        End If
        
        'Lock da Região de Venda
        If objVendedor.iCodRegiao <> 0 Then
            
            lErro = CF("RegiaoVenda_Lock", alComando(2), objVendedor.iCodRegiao)
            If lErro <> SUCESSO Then gError 16685
                                
        End If
    
    End If
    'Pesquisa Vendedor no BD
    lErro = Comando_ExecutarPos(alComando(3), "SELECT Endereco FROM Vendedores WHERE Codigo = ? ", 0, lEndereco, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 16672

    'Verifica resultado da pesquisa
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16673

    If lErro = AD_SQL_SEM_DADOS Then
        'Vendedor não pertence ao BD
        
        'Gera novo código de endereço
        lErro = CF("Endereco_Automatico", lCodigo)
        If lErro <> SUCESSO Then gError 16661
    
        objVendedor.lEndereco = lCodigo
        objEndereco.lCodigo = lCodigo
        
        'Chama função de gravação de endereço do vendedor
        lErro = CF("Endereco_Grava", objEndereco)
        If lErro <> SUCESSO Then gError 16667
        
        'Insere novo registro na tabela Vendedores
        lErro = Comando_Executar(alComando(3), "INSERT INTO Vendedores (Codigo, Nome, NomeReduzido, Endereco, Tipo, Matricula, CodRegiao, PercComissao, PercComissaoBaixa, PercComissaoEmissao, ComissaoSobreTotal, ComissaoFrete, ComissaoSeguro, ComissaoICM, ComissaoIPI, Banco, Agencia, ContaCorrente, Ativo, Vinculo, CGC, InscricaoEstadual, RazaoSocial, Cargo, Superior, CodUsuario, RG) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objVendedor.iCodigo, objVendedor.sNome, objVendedor.sNomeReduzido, objVendedor.lEndereco, objVendedor.iTipo, objVendedor.sMatricula, objVendedor.iCodRegiao, objVendedor.dPercComissao, objVendedor.dPercComissaoBaixa, objVendedor.dPercComissaoEmissao, objVendedor.iComissaoSobreTotal, objVendedor.iComissaoFrete, objVendedor.iComissaoSeguro, objVendedor.iComissaoICM, objVendedor.iComissaoIPI, objVendedor.iBanco, objVendedor.sAgencia, objVendedor.sContaCorrente, objVendedor.iAtivo, objVendedor.iVinculo, objVendedor.sCgc, objVendedor.sInscricaoEstadual, objVendedor.sRazaoSocial, _
            objVendedor.iCargo, objVendedor.iSuperior, objVendedor.sCodUsuario, objVendedor.sRG)
        If lErro <> AD_SQL_SUCESSO Then gError 16674
        
    Else  'Vendedor existe no BD
        
        'Coloca código de endereço de Vendedor do BD em objEndereco
        objEndereco.lCodigo = lEndereco
        
        'Atualiza Endereco na tabela Enderecos
        lErro = CF("Endereco_Grava", objEndereco)
        If lErro <> SUCESSO Then gError 16665
        
        'Altera registro na tabela Vendedores
        lErro = Comando_ExecutarPos(alComando(4), "UPDATE Vendedores SET Nome = ?, NomeReduzido = ?, Tipo = ?, Matricula = ?, CodRegiao = ?, PercComissao = ?, PercComissaoBaixa = ?, PercComissaoEmissao = ?, ComissaoSobreTotal = ?, ComissaoFrete = ?, ComissaoSeguro=?, ComissaoICM = ?, ComissaoIPI = ?, Banco = ?, Agencia = ?, ContaCorrente = ?, Ativo = ?, Vinculo = ?, CGC = ?, InscricaoEstadual = ?, RazaoSocial = ?, Cargo = ?, Superior = ?, CodUsuario = ?, RG = ?", alComando(3), objVendedor.sNome, objVendedor.sNomeReduzido, objVendedor.iTipo, objVendedor.sMatricula, objVendedor.iCodRegiao, objVendedor.dPercComissao, objVendedor.dPercComissaoBaixa, objVendedor.dPercComissaoEmissao, objVendedor.iComissaoSobreTotal, objVendedor.iComissaoFrete, objVendedor.iComissaoSeguro, objVendedor.iComissaoICM, objVendedor.iComissaoIPI, objVendedor.iBanco, objVendedor.sAgencia, objVendedor.sContaCorrente, objVendedor.iAtivo, objVendedor.iVinculo, objVendedor.sCgc, objVendedor.sInscricaoEstadual, objVendedor.sRazaoSocial, _
            objVendedor.iCargo, objVendedor.iSuperior, objVendedor.sCodUsuario, objVendedor.sRG)
        If lErro <> AD_SQL_SUCESSO Then gError 16676

    End If

    lErro = CF("Vendedor_Grava_Customizacao", objVendedor)
    If lErro <> SUCESSO Then gError 183136

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Vendedor_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Vendedor_Grava_EmTrans:

    Vendedor_Grava_EmTrans = gErr

        Select Case gErr

        Case 16664
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 16669, 16670, 16672, 16673
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", gErr, objVendedor.iCodigo)

        Case 16671
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_VENDEDOR_REPETIDO", gErr, iVendedor)

        Case 16674
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_VENDEDOR", gErr, objVendedor.iCodigo)
        
        Case 16676
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_VENDEDOR", gErr, objVendedor.iCodigo)

        Case 16665, 16684, 16685, 16661, 16667
            'Erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148876)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function TipoVendedor_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'faz lock shared no Tipo Vendedor passado como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer

On Error GoTo Erro_TipoVendedor_Lock
    
    'Pesquisa Tipo Vendedor no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM TiposDeVendedor WHERE Codigo = ? ", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16678

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16679

    If lErro = AD_SQL_SEM_DADOS Then Error 16680
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16681
            
    TipoVendedor_Lock = SUCESSO

    Exit Function

Erro_TipoVendedor_Lock:

    TipoVendedor_Lock = Err

    Select Case Err

        Case 16678, 16679
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEVENDEDOR", Err, iCodigo)

        Case 16680
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEVENDEDOR_NAO_CADASTRADO", Err, iCodigo)

        Case 16681
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOSDEVENDEDOR", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148877)

    End Select

    Exit Function

End Function

Function RegiaoVenda_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'faz lock shared na Regiao de Venda passada como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer

On Error GoTo Erro_RegiaoVenda_Lock
    
    'pesquisa Regiao Venda no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM RegioesVendas WHERE Codigo = ? ", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16687

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16688

    If lErro = AD_SQL_SEM_DADOS Then Error 16689
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16690
            
    RegiaoVenda_Lock = SUCESSO

    Exit Function

Erro_RegiaoVenda_Lock:

    RegiaoVenda_Lock = Err

    Select Case Err

        Case 16687, 16688
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_REGIOESVENDAS", Err, iCodigo)

        Case 16689
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAO_VENDA_NAO_CADASTRADA", Err, iCodigo)

        Case 16690
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_REGIOESVENDAS", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148878)

    End Select

    Exit Function

End Function

Function Vendedor_Exclui(ByVal objVendedor As ClassVendedor) As Long
'Exclui Vendedor do BD

Dim lErro As Long
Dim iCodigo As Integer
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lTransacao As Long
Dim lNumIntCom As Long
Dim dtDataLog As Date

On Error GoTo Erro_Vendedor_Exclui

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 16700

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 16701
    
    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 112581
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then gError 112582

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 16702
    
    'Pesquisa Vendedor no BD
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo, DataLog FROM Vendedores WHERE Codigo = ? ", 0, iCodigo, dtDataLog, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16703

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 16704

    If lErro = AD_SQL_SEM_DADOS Then gError 16705

    'se o vendedor já foi transferido para o loja ==> nao pode excluir
    If dtDataLog <> DATA_NULA Then gError 126461

    'Lock em Vendedor
    lErro = Comando_LockExclusive(lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 16706

    lErro = CF("Vendedor_Objetos_Relacionados_Testa", objVendedor)
    If lErro <> SUCESSO Then gError 33965
    
    'Exclui Vendedor da tabela Vendedores
    lErro = Comando_ExecutarPos(lComando2, "DELETE FROM Vendedores", lComando1)
    If lErro <> AD_SQL_SUCESSO Then gError 16709
    
    'Pesquisa Vendedor no BD
    lErro = Comando_ExecutarPos(lComando3, "SELECT Vendedor FROM VendedorLoja WHERE Vendedor = ? ", 0, iCodigo, objVendedor.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112578

    lErro = Comando_BuscarPrimeiro(lComando3)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112579
    
    If lErro = AD_SQL_SUCESSO Then
        'Exclui Vendedor da tabela Vendedorloja
        lErro = Comando_ExecutarPos(lComando4, "DELETE FROM VendedorLoja", lComando3)
        If lErro <> AD_SQL_SUCESSO Then gError 112580
    End If
    
    'Exclui endereço do vendedor
    lErro = CF("Endereco_Exclui", objVendedor.lEndereco)
    If lErro <> SUCESSO Then gError 16699
        
    lErro = CF("Vendedor_Exclui_Customizado", objVendedor)
    If lErro <> SUCESSO Then gError 195891
        
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 16710

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Vendedor_Exclui = SUCESSO

    Exit Function

Erro_Vendedor_Exclui:

    Vendedor_Exclui = gErr

    Select Case gErr
    
        Case 16700, 16701, 112581, 112582
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 16702
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 16703, 16704, 112578, 112579
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", gErr, objVendedor.iCodigo)
        
        Case 16705
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case 16706
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_VENDEDOR", gErr)
            
        Case 16699, 33965, 195891


        Case 16709, 112580
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDOR", gErr, objVendedor.iCodigo)

        Case 16710
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 126461
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_DATALOG", gErr, objVendedor.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148879)

    End Select

    Call Transacao_Rollback
    
    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)

    Exit Function

End Function

Function NFs_FilialCliente_Nao_Relacionada(ByVal objFilialCliente As ClassFilialCliente) As Long
'Verifica se FilialCliente não está relacionada com Notas Fiscais (lançadas ou baixadas)
'Retorna SUCESSO se não estiver relacionada.
Dim lErro As Long, lComando As Long
Dim lNumIntDoc As Long
On Error GoTo Erro_NFs_FilialCliente_Nao_Relacionada

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27483

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE Cliente = ? AND FilialCli = ?", lNumIntDoc, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 27484

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27485
    
    'se achou alguma nota relacionada
    If lErro = AD_SQL_SUCESSO Then Error 27486
    
    lErro = Comando_Fechar(lComando)
    
    NFs_FilialCliente_Nao_Relacionada = SUCESSO

    Exit Function

Erro_NFs_FilialCliente_Nao_Relacionada:

    NFs_FilialCliente_Nao_Relacionada = Err

    Select Case Err

        Case 27483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 27484, 27485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NOTAS_FISCAIS_FIL_CLI", Err)
        
        Case 27486 'achou alguma nota relacionada
            Call Rotina_Erro(vbOKOnly, "ERRO_EXISTEM_NOTAS_FISCAIS_FIL_CLI", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148880)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TitulosReceber_FilialCliente_Nao_Relacionada(ByVal objFilialCliente As ClassFilialCliente) As Long
'Verifica se FilialCliente não está relacionada com Títulos a Receber (lançados ou baixados)
'Retorna SUCESSO se não estiver relacionada.

Dim colCodTitulosReceber As New Collection
Dim lErro As Long

On Error GoTo Erro_TitulosReceber_FilialCliente_Nao_Relacionada

    'Preenche a coleção de códigos de Títulos a Receber relacionados à FilialCliente
    lErro = CF("TitulosReceber_Le_FilialCliente", objFilialCliente, colCodTitulosReceber)
    If lErro <> SUCESSO And lErro <> 6984 Then Error 6919

    'Verifica se FilialCliente está relacionada com Títulos a Receber
    If lErro = SUCESSO Then Error 6920

    'Preenche a coleção de códigos de Títulos a Receber Baixados relacionados à FilialCliente
    lErro = CF("TitulosReceberBaixados_Le_FilialCliente", objFilialCliente, colCodTitulosReceber)
    If lErro <> SUCESSO And lErro <> 6989 Then Error 6921

    'Verifica se Cliente está relacionado com Títulos a Receber Baixado
    If lErro = SUCESSO Then Error 6922

    TitulosReceber_FilialCliente_Nao_Relacionada = SUCESSO

    Exit Function

Erro_TitulosReceber_FilialCliente_Nao_Relacionada:

    TitulosReceber_FilialCliente_Nao_Relacionada = Err

    Select Case Err

        Case 6919, 6921

        Case 6920
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_TITULOS_REC", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 6922
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_TITULOS_REC_BAIXADOS", Err, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148881)

    End Select

    Exit Function

End Function

Function DebitosRecCli_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente) As Long
'Le todos códigos internos de Debitos Receber relacionadas a FilialCliente

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_DebitosRecCli_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12682

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM DebitosRecCli WHERE Cliente = ? AND Filial = ? ORDER BY NumIntDoc", lNumIntDoc, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 12683

    'le o primeiro debito
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12684

    If lErro = AD_SQL_SEM_DADOS Then Error 12685

    lErro = Comando_Fechar(lComando)

    DebitosRecCli_Le_FilialCliente = SUCESSO

    Exit Function

Erro_DebitosRecCli_Le_FilialCliente:

    DebitosRecCli_Le_FilialCliente = Err

    Select Case Err

        Case 12682
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12683, 12684, 12686
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DEBITOSRECCLI", Err)

        Case 12685 'sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148882)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RecebAntecipados_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente) As Long
'Le todos os códigos internos de Recebimentos Antecipados relacionados a Filial Cliente
'Devolve em colCodRecAntec

Dim lComando As Long
Dim lNumIntRec As Long
Dim lErro As Long

On Error GoTo Erro_RecebAntecipados_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 12687

    lErro = Comando_Executar(lComando, "SELECT NumIntRec FROM RecebAntecipados WHERE Cliente = ? AND Filial_Cliente = ? ORDER BY NumIntRec", lNumIntRec, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 12688

    'le o primeiro movimento
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 12689

    If lErro = AD_SQL_SEM_DADOS Then Error 12690

    lErro = Comando_Fechar(lComando)

    RecebAntecipados_Le_FilialCliente = SUCESSO

    Exit Function

Erro_RecebAntecipados_Le_FilialCliente:

    RecebAntecipados_Le_FilialCliente = Err

    Select Case Err

        Case 12687
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 12688, 12689, 12691
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RECEB_ANTEC", Err)

        Case 12690 'sem dados

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148883)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function MvDiaCli_Exclui_FilialCliente(ByVal objFilialCliente As ClassFilialCliente) As Long
'Exclui os registros da tabela MvDiaCli associados à FilialCliente passada.
'Chamada DENTRO DE TRANSAÇÃO por FilialCliente_Exclui.

Dim lComando(2) As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tMvDiaCli As typeMvDiaCli
Dim iIndice As Integer

On Error GoTo Erro_MvDiaCli_Exclui_FilialCliente

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 6949
    Next

    'Seleciona os Movimentos da FilialCliente a serem excluídos
    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa, Data FROM MvDiaCli WHERE Cliente =? AND Filial=?", 0, tMvDiaCli.iFilialEmpresa, tMvDiaCli.dtData, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 6950

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6951

    'Exclui os Movimentos da FilialCliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6952

        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM MvDiaCli", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6953

        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6954

    Loop

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    MvDiaCli_Exclui_FilialCliente = SUCESSO

    Exit Function

Erro_MvDiaCli_Exclui_FilialCliente:

    MvDiaCli_Exclui_FilialCliente = Err

    Select Case Err

        Case 6949
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6950, 6951, 6954
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVDIACLI", Err)

        Case 6952
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVDIACLI", Err, tMvDiaCli.iFilialEmpresa, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, tMvDiaCli.dtData)

        Case 6953
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVDIACLI", Err, tMvDiaCli.iFilialEmpresa, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, tMvDiaCli.dtData)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148884)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function MvPerCli_Exclui_FilialCliente(ByVal objFilialCliente As ClassFilialCliente) As Long
'Exclui os registros da tabela MvPerCli associados à FilialCliente passada.
'Chamada DENTRO DE TRANSAÇÃO por FilialCliente_Exclui.

Dim lComando(2) As Long
Dim lErro As Long
Dim iFim_de_Arquivo As Integer
Dim tMvPerCli As typeMvPerCli
Dim iIndice As Integer

On Error GoTo Erro_MvPerCli_Exclui_FilialCliente

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 6955
    Next

    'Seleciona os Movimentos da FilialCliente a serem excluídos
    lErro = Comando_ExecutarPos(lComando(0), "SELECT FilialEmpresa, Exercicio FROM MvPerCli WHERE Cliente =? AND Filial=?", 0, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 6956

    iFim_de_Arquivo = Comando_BuscarPrimeiro(lComando(0))
    If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6957

    'Exclui os Movimentos da FilialCliente
    Do While iFim_de_Arquivo = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6958

        lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM MvPerCli", lComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 6959

        iFim_de_Arquivo = Comando_BuscarProximo(lComando(0))
        If iFim_de_Arquivo <> AD_SQL_SUCESSO And iFim_de_Arquivo <> AD_SQL_SEM_DADOS Then Error 6960

    Loop

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    MvPerCli_Exclui_FilialCliente = SUCESSO

    Exit Function

Erro_MvPerCli_Exclui_FilialCliente:

    MvPerCli_Exclui_FilialCliente = Err

    Select Case Err

        Case 6955
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6956, 6957, 6960
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MVPERCLI", Err)

        Case 6958
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MVPERCLI", Err, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case 6959
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MVPERCLI", Err, tMvPerCli.iFilialEmpresa, tMvPerCli.iExercicio, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148885)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function TitulosReceber_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente, ByVal colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber relacionados à FilialCliente passada
'Devolve em colCodTitulosReceber

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_TitulosReceber_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6981

    'Seleciona os códigos de Títulos a Receber relacionados à FilialCliente
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRec WHERE Cliente = ? AND Filial=? AND Status <> ? ORDER BY NumIntDoc", lNumIntDoc, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 6982

    'Lê o primeiro código de Título a Receber
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6983
    If lErro = AD_SQL_SEM_DADOS Then Error 6984

    Do While lErro <> AD_SQL_SEM_DADOS

        'Adiciona o código do Título a Receber lido na coleção
        colCodTitulosReceber.Add lNumIntDoc

        'Lê o próximo Título
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6985

    Loop

    lErro = Comando_Fechar(lComando)

    TitulosReceber_Le_FilialCliente = SUCESSO

    Exit Function

Erro_TitulosReceber_Le_FilialCliente:

    TitulosReceber_Le_FilialCliente = Err

    Select Case Err

        Case 6981
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6982, 6983, 6985
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC", Err)

        Case 6984 'Sem dados. A ser tratado na rotina chamadora.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148886)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TitulosReceberBaixados_Le_FilialCliente(ByVal objFilialCliente As ClassFilialCliente, ByVal colCodTitulosReceber As Collection) As Long
'Lê todos códigos de Títulos a Receber Baixados relacionados à FilialCliente passada
'Devolve em colCodTitulosReceber

Dim lComando As Long
Dim lNumIntDoc As Long
Dim lErro As Long

On Error GoTo Erro_TitulosReceberBaixados_Le_FilialCliente

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 6986

    'Seleciona os códigos de Títulos a Receber Baixados relacionados à FilialCliente
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRecBaixados WHERE Cliente = ? AND Filial=? ORDER BY NumIntDoc", lNumIntDoc, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 6987

    'Lê o primeiro código de Título a Receber
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6988
    If lErro = AD_SQL_SEM_DADOS Then Error 6989

    Do While lErro <> AD_SQL_SEM_DADOS

        'Adiciona o código do Título a Receber lido na coleção
        colCodTitulosReceber.Add lNumIntDoc

        'Lê o próximo Título
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 6990

    Loop

    lErro = Comando_Fechar(lComando)

    TitulosReceberBaixados_Le_FilialCliente = SUCESSO

    Exit Function

Erro_TitulosReceberBaixados_Le_FilialCliente:

    TitulosReceberBaixados_Le_FilialCliente = Err

    Select Case Err

        Case 6986
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 6987, 6988, 6990
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_REC_BAIXADOS", Err)

        Case 6989 'Sem dados. A ser tratado na rotina chamadora.

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148887)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloReceber_Grava(ByVal objTituloReceber As ClassTituloReceber, ByVal colComissaoEmissao As colComissao, ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal colcolDesconto As colcolDesconto, ByVal objContabil As ClassContabil) As Long
'Grava um Título à Receber

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TituloReceber_Grava

    'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 26225
    
    lErro = CF("TituloReceber_Grava_EmTrans", objTituloReceber, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
    If lErro <> SUCESSO Then gError 182503

    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 26240
    
    TituloReceber_Grava = SUCESSO
    
    Exit Function

Erro_TituloReceber_Grava:

    TituloReceber_Grava = gErr
    
    Select Case gErr

        Case 26225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 26240
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 182503
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148888)
            
    End Select

    'Fecha a transação
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Function TituloReceber_Grava_EmTrans(ByVal objTituloReceber As ClassTituloReceber, ByVal colComissaoEmissao As colComissao, ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal colcolDesconto As colcolDesconto, ByVal objContabil As ClassContabil) As Long
'Grava um Título à Receber

Dim lErro As Long
Dim lTransacao As Long, dValorLivroAux As Double
Dim lNumIntDoc As Long, lNumIntNF As Long

On Error GoTo Erro_TituloReceber_Grava_EmTrans

    'Abre a transacao
'    lTransacao = Transacao_Abrir()
'    If lTransacao = 0 Then gError 26225
        
    'Grava no BD
    lErro = TituloReceber_Grava_BD(objTituloReceber, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto)
    If lErro <> SUCESSO Then gError 26239
    
    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Grava", objTituloReceber.objProjetoInfo, PRJ_CR_TIPO_TITREC, objTituloReceber.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181706
    '############################################

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'se o documento foi originado fora de CPR nao pode ter a contabilidade alterada em CPR
        lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloReceber.sSiglaDocumento, objTituloReceber.lNumIntDoc, lNumIntNF)
        If lErro <> SUCESSO And lErro <> 41542 Then gError 59225
        
        'se nao é um docto gerado fora do cpr
        If lErro <> SUCESSO Then
            
            dValorLivroAux = Round(objTituloReceber.dValor - (objTituloReceber.dValorIRRF + objTituloReceber.dPISRetido + objTituloReceber.dISSRetido + objTituloReceber.dCOFINSRetido + objTituloReceber.dCSLLRetido) - IIf(objTituloReceber.iINSSRetido <> 0, objTituloReceber.dValorINSS, 0), 2)
            
            If Not (objContabil Is Nothing) Then
            
                'Grava os dados contábeis (contabilidade)
                lErro = objContabil.Contabil_Gravar_Registro(objTituloReceber.lNumIntDoc, objTituloReceber.lCliente, objTituloReceber.iFilial, DATA_NULA, , 0, , -dValorLivroAux)
                If lErro <> SUCESSO Then gError 20523
                
            End If

        End If
        
    End If
    
'    'Encerra transação
'    lErro = Transacao_Commit()
'    If lErro <> AD_SQL_SUCESSO Then gError 26240
    
    TituloReceber_Grava_EmTrans = SUCESSO
    
    Exit Function

Erro_TituloReceber_Grava_EmTrans:

    TituloReceber_Grava_EmTrans = gErr
    
    Select Case gErr
    
        Case 20523, 26239, 59225, 181706
    
        Case 26225
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 26240
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148888)
            
    End Select

    'Fecha a transação
'    Call Transacao_Rollback
    
    Exit Function
    
End Function

Private Function TituloReceber_Valida_Gravacao(ByVal objTituloReceber As ClassTituloReceber, ByVal objTituloReceberBD As ClassTituloReceber, ByVal colParcelaReceber As colParcelaReceber, ByVal colParcelaReceberBD As colParcelaReceber, ByVal colComissaoEmissao As colComissao, ByVal colComissaoEmissaoBD As colComissao, ByVal colcolDesconto As colcolDesconto) As Long
'Valida os dados de um Título a Receber para a Gravação

Dim iBaixado As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim iCamposAlterados As Integer
Dim iIndice As Integer
Dim dDesconto As Double
Dim lErro As Long, iMenorNumParc As Integer

On Error GoTo Erro_TituloReceber_Valida_Gravacao

    If objTituloReceber.iStatus = STATUS_PREVISAO Then
    
        If objTituloReceberBD.iStatus <> STATUS_PREVISAO Then
        
            'Exclui as Comissoes já cadastradas, inclusive as baixadas
            lErro = CF("Comissoes_Exclui", TIPO_TITULO_RECEBER, objTituloReceberBD.lNumIntDoc, 1)
            If lErro <> SUCESSO Then Error 26392
        
        End If
                    
    Else
        
        iBaixado = False
        'Verifica se alguma comissao foi baixada
        For iIndice = 1 To colComissaoEmissaoBD.Count
            'Se foi
            If colComissaoEmissaoBD(iIndice).iStatus = STATUS_BAIXADO Then
                
                iBaixado = True
                'Verifica se o Numero de comissões é o mesmo que o existente no BD
                If colComissaoEmissao.Count <> colComissaoEmissaoBD.Count Then
                    'Avisa que não é possível modificar o número de comissoes
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_COMISSOES_EMISSAO_NAO_ALTERAVEIS", objTituloReceberBD.lNumIntDoc)
                    If vbMsgRes = vbNo Then Error 26261
                End If
                'Verifica se algum  campo está diferente aos BD
                If colComissaoEmissao(iIndice).iCodVendedor <> colComissaoEmissaoBD(iIndice).iCodVendedor Or colComissaoEmissao(iIndice).dPercentual <> colComissaoEmissaoBD(iIndice).dPercentual Or colComissaoEmissao(iIndice).dValorBase <> colComissaoEmissaoBD(iIndice).dValorBase Or colComissaoEmissao(iIndice).dValor <> colComissaoEmissaoBD(iIndice).dValor Then
                    'Avisa que uma Comissao paga não pode ser alterada
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_COMISSAO_EMISSAO_PAGA", objTituloReceberBD.lNumIntDoc)
                    If vbMsgRes = vbNo Then Error 26262
                    Exit For
                End If
            End If
        Next
    
        'Se nenhuma comissão foi paga
        If iBaixado = False Then
            'Executa as alterações no BD
            lErro = CF("Comissoes_Altera", colComissaoEmissao, TIPO_TITULO_RECEBER, objTituloReceber.lNumIntDoc)
            If lErro <> SUCESSO Then Error 26392
        End If
        
    End If
    
    '#########################################
    'Inserido por Wagner 19/04/2005 pq o campo de retencao de iss nao está na tela
    objTituloReceber.dValorISS = objTituloReceberBD.dValorISS
    '#########################################
    
    'Flag para alterações
    iCamposAlterados = 0
    
    'Verifica se o Valor do Titulo e o valor dos Impostos foram alterados
    If Abs(objTituloReceberBD.dValor - objTituloReceber.dValor) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dValorIRRF - objTituloReceber.dValorIRRF) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dISSRetido - objTituloReceber.dISSRetido) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dValorISS - objTituloReceber.dValorISS) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.iINSSRetido - objTituloReceber.iINSSRetido) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dValorINSS - objTituloReceber.dValorINSS) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dPISRetido - objTituloReceber.dPISRetido) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dCOFINSRetido - objTituloReceber.dCOFINSRetido) > DELTA_VALORMONETARIO Or _
        Abs(objTituloReceberBD.dCSLLRetido - objTituloReceber.dCSLLRetido) > DELTA_VALORMONETARIO _
        Then Error 32278
    
    'Se nada ainda foi alterado
    If iCamposAlterados = 0 Then
        
        iMenorNumParc = IIf(colParcelaReceberBD.Count < colParcelaReceber.Count, colParcelaReceberBD.Count, colParcelaReceber.Count)
        
        'Para cada Parcela
        For iIndice = 1 To iMenorNumParc
            
            'Verificar se o Valor é igual ao do BD
            If colParcelaReceber(iIndice).dValor <> colParcelaReceberBD(iIndice).dValor Then
                iCamposAlterados = 1
            End If
        Next
    End If
    
    If iCamposAlterados = 1 Then
        
        'Pergunta se deseja prosseguir na alteração para os alteráveis
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATAVENCIMENTO_PARCELAS_ALTERAVEIS")
        If vbMsgRes = vbNo Then Error 26254
                    
    End If

    TituloReceber_Valida_Gravacao = SUCESSO
    
    Exit Function

Erro_TituloReceber_Valida_Gravacao:

    TituloReceber_Valida_Gravacao = Err
    
    Select Case Err
    
        Case 26254, 26261, 26262, 26263, 26392
        
        Case 32278
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TITREC_ALTERADO", Err)
        
        Case 26264
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_CHQPRE_DIFERENTE_DESCONTO", Err, iIndice, dDesconto)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148889)

    End Select
    
    Exit Function
    
End Function

Private Function TituloReceber_Grava_BD(ByVal objTituloReceber As ClassTituloReceber, ByVal colComissaoEmissao As colComissao, ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal colcolDesconto As colcolDesconto) As Long
'Grava o Título à Receber no Banco de Dados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim objTituloReceberBD As New ClassTituloReceber
Dim colParcelaReceberBD As New colParcelaReceber
Dim colComissaoEm As New colComissao
Dim tTituloReceber As typeTituloReceber
Dim colcolDescontoBD As New colcolDesconto
Dim colcolComissaoBD As New colcolComissao
Dim iIndice As Integer, iMenorNumParc As Integer
Dim objComissao As ClassComissao

Dim lComando2 As Long

On Error GoTo Erro_TituloReceber_Grava_BD

    'Abre os Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26241
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26242
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 108750
    
    If objTituloReceber.iFilialEmpresa = EMPRESA_TODA Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    tTituloReceber.sCcl = String(STRING_CCL, 0)
    tTituloReceber.sNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    'Verifica se o Titulo já existe
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte FROM TitulosRec WHERE NumTitulo = ? AND DataEmissao = ? AND SiglaDocumento = ? AND Status <> ? AND FilialEmpresa = ?", 0, tTituloReceber.lNumIntDoc, tTituloReceber.dValor, tTituloReceber.dValorIRRF, tTituloReceber.dValorISS, tTituloReceber.dISSRetido, tTituloReceber.dValorINSS, tTituloReceber.iINSSRetido, tTituloReceber.dPISRetido, tTituloReceber.dCOFINSRetido, tTituloReceber.dCSLLRetido, tTituloReceber.iMoeda, tTituloReceber.sCcl, tTituloReceber.sNatureza, tTituloReceber.lCpoGenerico1, tTituloReceber.iReajustePeriodicidade, tTituloReceber.dtReajusteBase, tTituloReceber.dtReajustadoAte, _
        objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, objTituloReceber.sSiglaDocumento, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 26243

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26244

    If lErro = AD_SQL_SUCESSO Then
            
        'Faz o Lock do Titulo a Receber
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 26386
               
        'Coloca o NumIntTitulo lido na coleção de parcelas
        For Each objParcelaReceber In colParcelaReceber
            objParcelaReceber.lNumIntTitulo = tTituloReceber.lNumIntDoc
        Next
        For Each objComissao In colComissaoEmissao
            objComissao.iTipoTitulo = TIPO_TITULO_RECEBER
            objComissao.lNumIntDoc = tTituloReceber.lNumIntDoc
            objComissao.iStatus = STATUS_LIBERADO
            objComissao.dtDataGeracao = objTituloReceber.dtDataEmissao
            objComissao.dtDataBaixa = DATA_NULA
        Next
        
        'Coloca o NumIntDoc lido no objTituloReceber
        objTituloReceber.lNumIntDoc = tTituloReceber.lNumIntDoc
        
        'Preenche objTituloReceberBD com alguns dados
        objTituloReceberBD.lNumIntDoc = tTituloReceber.lNumIntDoc
        objTituloReceberBD.dValor = tTituloReceber.dValor
        objTituloReceberBD.dValorIRRF = tTituloReceber.dValorIRRF
        objTituloReceberBD.dISSRetido = tTituloReceber.dISSRetido
        objTituloReceberBD.dValorISS = tTituloReceber.dValorISS
        objTituloReceberBD.iINSSRetido = tTituloReceber.iINSSRetido
        objTituloReceberBD.dValorINSS = tTituloReceber.dValorINSS
        objTituloReceberBD.dPISRetido = tTituloReceber.dPISRetido
        objTituloReceberBD.dCOFINSRetido = tTituloReceber.dCOFINSRetido
        objTituloReceberBD.dCSLLRetido = tTituloReceber.dCSLLRetido
        objTituloReceberBD.iMoeda = tTituloReceber.iMoeda
        objTituloReceberBD.sCcl = tTituloReceber.sCcl
        objTituloReceberBD.sNatureza = tTituloReceber.sNatureza
        objTituloReceberBD.lCpoGenerico1 = tTituloReceber.lCpoGenerico1
        objTituloReceberBD.iReajustePeriodicidade = tTituloReceber.iReajustePeriodicidade
        objTituloReceberBD.dtReajusteBase = tTituloReceber.dtReajusteBase
        objTituloReceberBD.dtReajustadoAte = tTituloReceber.dtReajustadoAte
            
        'Lê no BD as Parcelas referentes ao Título
        lErro = ParcelasReceber_Le_Locando(objTituloReceber, colParcelaReceberBD)
        If lErro <> SUCESSO Then gError 26246
        
        lErro = TituloReceber_Grava_Trata_NFe(objTituloReceber, colParcelaReceberBD)
        If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        iMenorNumParc = IIf(colParcelaReceberBD.Count < colParcelaReceber.Count, colParcelaReceberBD.Count, colParcelaReceber.Count)
                
        For iIndice = 1 To iMenorNumParc
            colParcelaReceber(iIndice).lNumIntDoc = colParcelaReceberBD(iIndice).lNumIntDoc
            colParcelaReceber(iIndice).iStatus = colParcelaReceberBD(iIndice).iStatus
            colParcelaReceber(iIndice).iProxSeqBaixa = colParcelaReceberBD(iIndice).iProxSeqBaixa
            colParcelaReceber(iIndice).iCobrador = colParcelaReceberBD(iIndice).iCobrador
            colParcelaReceber(iIndice).iCarteiraCobranca = colParcelaReceberBD(iIndice).iCarteiraCobranca
            
            For Each objComissao In colcolComissao(iIndice).colComissao
                
                objComissao.iTipoTitulo = TIPO_PARCELA
                objComissao.lNumIntDoc = colParcelaReceber(iIndice).lNumIntDoc
                objComissao.dtDataBaixa = DATA_NULA
                objComissao.dtDataGeracao = DATA_NULA
               
            Next
        
        Next
        
        'Lê no BD as Comissões referentes ao Título
        lErro = Comissoes_Le_Locando(tTituloReceber.lNumIntDoc, colComissaoEm, TIPO_TITULO_RECEBER)
        If lErro <> SUCESSO Then gError 26253
        
        'Valida os dados do Título para a gravação
        lErro = TituloReceber_Valida_Gravacao(objTituloReceber, objTituloReceberBD, colParcelaReceber, colParcelaReceberBD, colComissaoEmissao, colComissaoEm, colcolDesconto)
        If lErro <> SUCESSO Then gError 26390
        
        'Traz os dados particulares a cada Parcela do BD
        lErro = ParcelasReceber_Traz_DadosBD(colParcelaReceberBD, colcolDescontoBD, colcolComissaoBD)
        If lErro <> SUCESSO Then gError 26389
        
        'Valida os dados das Parcelas e faz as Alterações no BD
        lErro = CF("TituloReceber_Altera", colParcelaReceber, colParcelaReceberBD, colcolDesconto, colcolDescontoBD, colcolComissao, colcolComissaoBD)
        If lErro <> SUCESSO Then gError 26387
        
        'Atualiza o Número de Parcelas e outros dados em TitulosRec
        lErro = Comando_ExecutarPos(lComando2, "UPDATE TitulosRec SET NumParcelas=?, Moeda = ?, Ccl = ?, Natureza = ?, CpoGenerico1 = ?, ReajustePeriodicidade = ?, ReajusteBase = ?, CondicaoPagto = ?", lComando, objTituloReceber.iNumParcelas, objTituloReceber.iMoeda, objTituloReceber.sCcl, objTituloReceber.sNatureza, objTituloReceber.lCpoGenerico1, objTituloReceber.iReajustePeriodicidade, objTituloReceber.dtReajusteBase, objTituloReceber.iCondicaoPagto)
        If lErro <> SUCESSO Then gError 108751
         
    ElseIf lErro = AD_SQL_SEM_DADOS Then 'Não achou o Título à Receber
        
        If gcolModulo.Ativo(MODULO_FATURAMENTO) <> NAO_USA_MODULO And (objTituloReceber.sSiglaDocumento = TIPODOC_NF_FATURA_RECEBER Or objTituloReceber.sSiglaDocumento = TIPODOC_NF_FAT_SERVICO_RECEBER) Then gError 62418
        
        'Cria um Novo Título a Receber
        lErro = CF("TituloReceber_Insere", objTituloReceber, colParcelaReceber, colComissaoEmissao, colcolComissao)
        If lErro <> SUCESSO Then gError 26388
    
    End If
        
    'Fecha os comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
    TituloReceber_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Grava_BD:

    TituloReceber_Grava_BD = gErr
    
    Select Case gErr
    
        Case 26241, 26242, 108750
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 26243, 26244, 108751
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)
            
''        Case 26245
''            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOREC_FILIALEMPRESA_DIFERENTE", gErr, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
''
        Case 26246, 26253, 26387, 26388, 26389, 26390
        
        Case 26252
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_ALTERADO", gErr, colParcelaReceber.Count, colParcelaReceberBD.Count)
        
        Case 26386
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TITULOSREC", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)
            
        Case 62418
            Call Rotina_Erro(vbOKOnly, "ERRO_CRIACAO_NFR_COM_FATURAMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148890)
        
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    
End Function

Private Function ParcelasReceber_Le_Locando(ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelasRec As colParcelaReceber) As Long
'Lê e bloqueia as Parcelas a Receber do Titulo cujo NumIntDoc foi passado em objTituloReceber

Dim lErro As Long
Dim lComando As Long
Dim tParcelaReceber As typeParcelaReceber
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_ParcelasReceber_Le_Locando

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26248

    tParcelaReceber.sNumTitCobrador = String(STRING_NOSSO_NUMERO, 0)
    '###############################################
    'INSERIDO POR WAGNER
    tParcelaReceber.sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '###############################################
   
    
    'Seleciona Parcelas vinculadas ao Título
    tParcelaReceber.sAgenciaCheque = String(STRING_AGENCIA, 0)
    tParcelaReceber.sAprovacaoCartao = String(STRING_APROVACAO_CARTAO, 0)
    tParcelaReceber.sContaCorrenteCheque = String(STRING_CONTA_CORRENTE, 0)
    tParcelaReceber.sNumeroCartao = String(STRING_NUMERO_CARTAO, 0)
    
    '#############################################
    'ALTERADO POR WAGNER
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumParcela, Status, DataVencimento, DataVencimentoReal, Saldo, Valor, Cobrador, CarteiraCobranca, NumTitCobrador,ProxSeqBaixa, ProxSeqOcorr, Desconto1Codigo, Desconto1Ate, Desconto1Valor, Desconto2Codigo, Desconto2Ate, Desconto2Valor, Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, Previsao, Observacao, ValorOriginal, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao " & _
        "FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? ORDER BY DataVencimento", 0, tParcelaReceber.lNumIntDoc, tParcelaReceber.iNumParcela, tParcelaReceber.iStatus, tParcelaReceber.dtDataVencimento, tParcelaReceber.dtDataVencimentoReal, tParcelaReceber.dSaldo, tParcelaReceber.dValor, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.sNumTitCobrador, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iProxSeqOcorr, tParcelaReceber.iDesconto1Codigo, tParcelaReceber.dtDesconto1Ate, tParcelaReceber.dDesconto1Valor, _
        tParcelaReceber.iDesconto2Codigo, tParcelaReceber.dtDesconto2Ate, tParcelaReceber.dDesconto2Valor, tParcelaReceber.iDesconto3Codigo, tParcelaReceber.dtDesconto3Ate, tParcelaReceber.dDesconto3Valor, tParcelaReceber.iAceite, tParcelaReceber.iDescontada, tParcelaReceber.iPrevisao, tParcelaReceber.sObservacao, tParcelaReceber.dValorOriginal, _
        tParcelaReceber.iTipoPagto, tParcelaReceber.iCodConta, tParcelaReceber.dtDataCredito, tParcelaReceber.dtDataEmissaoCheque, tParcelaReceber.iBancoCheque, tParcelaReceber.sAgenciaCheque, tParcelaReceber.sContaCorrenteCheque, tParcelaReceber.lNumeroCheque, tParcelaReceber.dtDataDepositoCheque, tParcelaReceber.iAdmMeioPagto, tParcelaReceber.iParcelamento, tParcelaReceber.sNumeroCartao, tParcelaReceber.dtValidadeCartao, tParcelaReceber.sAprovacaoCartao, tParcelaReceber.dtDataTransacaoCartao, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 26249
    '#############################################
    
    'Tenta ler a primeira Parcela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18380
    
    'Se não achou Parcela --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 26250

    'Loop de adição de Parcelas à coleção
    Do While lErro <> AD_SQL_SEM_DADOS

        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26251

        'Adiciona Parcela na coleção
        '#############################################
        'ALTERADO POR WAGNER
        With tParcelaReceber
            Set objParcelaReceber = colParcelasRec.Add(.lNumIntDoc, objTituloReceber.lNumIntDoc, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, 0, 0, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValorOriginal)
            
            objParcelaReceber.iTipoPagto = .iTipoPagto
            objParcelaReceber.iCodConta = .iCodConta
            objParcelaReceber.dtDataCredito = .dtDataCredito
            objParcelaReceber.dtDataEmissaoCheque = .dtDataEmissaoCheque
            objParcelaReceber.iBancoCheque = .iBancoCheque
            objParcelaReceber.sAgenciaCheque = .sAgenciaCheque
            objParcelaReceber.sContaCorrenteCheque = .sContaCorrenteCheque
            objParcelaReceber.lNumeroCheque = .lNumeroCheque
            objParcelaReceber.dtDataDepositoCheque = .dtDataDepositoCheque
            objParcelaReceber.iAdmMeioPagto = .iAdmMeioPagto
            objParcelaReceber.iParcelamento = .iParcelamento
            objParcelaReceber.sNumeroCartao = .sNumeroCartao
            objParcelaReceber.dtValidadeCartao = .dtValidadeCartao
            objParcelaReceber.sAprovacaoCartao = .sAprovacaoCartao
            objParcelaReceber.dtDataTransacaoCartao = .dtDataTransacaoCartao
        End With
        '#############################################

        'Vai para a próxima Parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26252

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    ParcelasReceber_Le_Locando = SUCESSO

    Exit Function

Erro_ParcelasReceber_Le_Locando:

    ParcelasReceber_Le_Locando = Err

    Select Case Err

        Case 26248
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26249, 18380, 26252
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC1", Err, objTituloReceber.lNumIntDoc)

        Case 26250
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_SEM_PARCELAS", Err, objTituloReceber.lNumIntDoc)

        Case 26251
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148891)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Comissoes_Le_Locando(ByVal lNumIntDoc As Long, ByVal colComissao As colComissao, ByVal iTipoTitulo As Integer) As Long
'Deve ser chamada em transação

Dim lErro As Long
Dim lComando As Long
Dim tInfoComissao As typeInfoComissao

On Error GoTo Erro_Comissoes_Le_Locando

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26256

    'Lê as comissões do Documento com Tipo e o Número Interno passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntCom, Status, CodVendedor, DataBaixa, Percentual, ValorBase, Valor, DataGeracao FROM Comissoes WHERE TipoTitulo = ? AND NumIntDoc = ? AND FilialEmpresa = ? ORDER BY CodVendedor", 0, tInfoComissao.lNumIntCom, tInfoComissao.iStatus, tInfoComissao.iCodVendedor, tInfoComissao.dtDataBaixa, tInfoComissao.dPercentual, tInfoComissao.dValorBase, tInfoComissao.dValor, tInfoComissao.dtDataGeracao, iTipoTitulo, lNumIntDoc, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 26257

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26258

    'Loop de armazenamento das Comissões encontradas
    Do While lErro = AD_SQL_SUCESSO
        
        'Loca a Comissão
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26259

        With tInfoComissao
            colComissao.Add .lNumIntCom, giFilialEmpresa, .iStatus, iTipoTitulo, lNumIntDoc, .iCodVendedor, .dtDataBaixa, .dPercentual, .dValorBase, .dValor, .dtDataGeracao
        End With

        'Busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26260

    Loop

    'Libera o comando
    Call Comando_Fechar(lComando)

    Comissoes_Le_Locando = SUCESSO

    Exit Function

Erro_Comissoes_Le_Locando:

    Comissoes_Le_Locando = Err

    Select Case Err

        Case 26256
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26257, 26258, 26459, 26260
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES", Err, lNumIntDoc)
        
        Case 26259
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMISSOES", Err, lNumIntDoc, iTipoTitulo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148892)

    End Select

    'Libera o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CondicaoPagto_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared na Condicao de Pagamento passada como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer

On Error GoTo Erro_CondicaoPagto_Lock
    
    'Pesquisa Condicção de Pagamento no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM CondicoesPagto WHERE Codigo = ? ", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 16359

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16360

    'Não achou a Condição de Pagameto
    If lErro = AD_SQL_SEM_DADOS Then Error 16361
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 16362
            
    CondicaoPagto_Lock = SUCESSO

    Exit Function

Erro_CondicaoPagto_Lock:

    CondicaoPagto_Lock = Err

    Select Case Err

        Case 16359, 16360
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONDICAOPAGTO", Err)

        Case 16361
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, iCodigo)

        Case 16362
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONDICAOPAGTO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148893)

    End Select

    Exit Function

End Function

Function Comissoes_Exclui(ByVal iTipoTitulo As Integer, ByVal lNumIntDoc As Long, ByVal iInclusiveBaixadas As Integer) As Long
'Exclui as comissões do Documento com o Tipo e o Número Interno passados
'obs: nao verifica se as comissoes foram baixadas.

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lNumIntCom As Long, dtDataBaixa As Date
Dim objNFiscal As New ClassNFiscal
Dim objParcelaReceber As New ClassParcelaReceber
Dim objTituloReceber As New ClassTituloReceber
Dim objDebitoRecCli As New ClassDebitoRecCli

On Error GoTo Erro_Comissoes_Exclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26278
    
    lComando1 = Comando_Abrir()
    If lComando = 0 Then Error 26391
    
    'Lê as comissões com os dados passados
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntCom, DataBaixa FROM Comissoes WHERE TipoTitulo = ? AND NumIntDoc = ?", 0, lNumIntCom, dtDataBaixa, iTipoTitulo, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 26279

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26280

    'Loop de exclusão
    Do While lErro = AD_SQL_SUCESSO

        'loca a comissao
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26281
        
        'se não é para excluir comissoes baixadas e a comissao lida já foi baixada
        If iInclusiveBaixadas = 0 And dtDataBaixa <> DATA_NULA Then
        
            Select Case iTipoTitulo

                Case TIPO_NF

                    'Lê a Série, Número e FilialEmpresa da Nota Fiscal
                    objNFiscal.lNumIntDoc = lNumIntDoc

                    lErro = CF("NFiscal_Le", objNFiscal)
                    If lErro <> SUCESSO And lErro <> 31442 Then Error 58456
                                        
                    'Não encontrou  ---> Erro
                    If lErro = 31442 Then Error 58457
                    
                    Error 58458
                   
                Case TIPO_PARCELA
                    
                    'Lê o Número Interno do Título e o Número da Parcela
                    objParcelaReceber.lNumIntDoc = lNumIntDoc
                    
                    lErro = CF("ParcelaReceber_Le", objParcelaReceber)
                    If lErro <> SUCESSO And lErro <> 19147 Then Error 58459
                     
                    'Se não encontrou ---> Procura em Baixadas
                    If lErro = 19147 Then
                        
                        'Procura em Baixadas
                        lErro = CF("ParcelaReceber_Baixada_Le", objParcelaReceber)
                        If lErro <> SUCESSO And lErro <> 58559 Then Error 0

                        'Se não encontrou ---> Erro
                        If lErro = 58559 Then Error 58460
                        
                    End If
                    
                    
                    objTituloReceber.lNumIntDoc = objParcelaReceber.lNumIntTitulo
                    
                    'Lê o Número do Título
                    lErro = CF("TituloReceber_Le", objTituloReceber)
                    If lErro <> SUCESSO And lErro <> 26061 Then Error 58461
                    
                    'Se não encontrou ---> Procura em Baixados
                    If lErro = 26061 Then
                        
                        lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
                        If lErro <> SUCESSO And lErro <> 56570 Then Error 58461
                    
                        'Se não encontrou ---> Erro
                        If lErro = 56570 Then Error 58462
                    
                    End If
                    
                    Error 58463
                
                Case TIPO_DEBITO
                    
                    objDebitoRecCli.lNumIntDoc = lNumIntDoc
                    
                    'Lê os dados para informar o vinculo de Comissões com Débitos
                    lErro = CF("DebitoReceber_Le", objDebitoRecCli)
                    If lErro <> SUCESSO And lErro <> 17835 Then Error 58464

                    'Se não encontrou ---> Erro
                    If lErro = 17835 Then Error 58465
                                        
                    Error 58466
                
                Case TIPO_TITULO_RECEBER
                    
                    objTituloReceber.lNumIntDoc = lNumIntDoc
                    
                    'Lê o Número do Título
                    lErro = CF("TituloReceber_Le", objTituloReceber)
                    If lErro <> SUCESSO And lErro <> 26061 Then Error 58467
                    
                    'Se não encontrou ---> Procura em Baixados
                    If lErro = 26061 Then
                        
                        lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
                        If lErro <> SUCESSO And lErro <> 56570 Then Error 58461
                    
                        'Se não encontrou ---> Erro
                        If lErro = 56570 Then Error 58468
                    
                    End If
                    
                    Error 58469
                    
                Case Else
                    Error 56665
                End Select

        End If
        
        'Faz a exclusão no BD
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM Comissoes", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 26282
        
        'busca o próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26283
        
    Loop
    
    'Libera os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Comissoes_Exclui = SUCESSO
    
    Exit Function
    
Erro_Comissoes_Exclui:

    Comissoes_Exclui = Err
    
    Select Case Err
    
        Case 26278, 26391
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 26279, 26280, 26283
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES", Err, lNumIntDoc)
        
        Case 26281
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COMISSOES", Err, lNumIntDoc)
            
        Case 56665
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_BAIXADA", Err, lNumIntDoc, iTipoTitulo)
        
        Case 26282
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_COMISSOES", Err, lNumIntDoc, iTipoTitulo)
        
        Case 58456, 58459, 58461, 58464, 58467
        
        Case 58457
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTAFISCAL_NAO_CADASTRADO_COMISSOES", Err)
        
        Case 58458
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_BAIXADA_NFISCAL", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)
        
        Case 58460
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NAO_CADASTRADO_COMISSOES", Err)
        
        Case 58462, 58468
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOREC_NAO_CADASTRADO_COMISSOES", Err)
        
        Case 58463
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_BAIXADA_PARCELA", Err, objTituloReceber.lNumTitulo, objParcelaReceber.iNumParcela, objTituloReceber.iFilialEmpresa)
        
        Case 58465
            Call Rotina_Erro(vbOKOnly, "ERRO_DEBITOREC_NAO_CADASTRADO_COMISSOES", Err)
        
        Case 58466
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_BAIXADA_DEBITOS", Err, objDebitoRecCli.lNumTitulo, objDebitoRecCli.sSiglaDocumento, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial)
        
        Case 58469
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOES_BAIXADA_TITULO", Err, objTituloReceber.lNumTitulo, objTituloReceber.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148894)
            
    End Select
    
    'Libera os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
        
End Function

Private Function ParcelasReceber_Altera_CamposAlteraveis(ByVal colParcelaReceber As colParcelaReceber, ByVal iNumParcelas As Integer, ByVal bOrigemNF As Boolean) As Long
'Altera os campos Alteráveis das primeiras iNumParcelas parcelas da colecao
'As parcelas maiores que iNumParcelas, se houverem, serao novas e portanto serao incluidas e nao alteradas
'bOrigemNF indica se a alteracao está sendo feita da tela do FAT ou do CR. Se a origem for Nota fiscal nao poderao ser alterados os valores das parcelas

Dim lErro As Long, objParcRec As ClassParcelaReceber
Dim lComando As Long, dSaldo As Double, dValor As Double
Dim lComando1 As Long, objCobrador As New ClassCobrador, dSaldoChequePre As Double
Dim iIndice As Integer, iCobrador As Integer, iCarteiraCobranca As Integer
Dim lNumIntDoc As Long, dtDataVencimento As Date, dtDataVencimentoReal As Date
Dim objCarteiraCobrador As New ClassCarteiraCobrador, iNumParcela As Integer, bAlterouValor As Boolean
'###############################
'INSERIDO POR WAGNER
Dim iPrevisao As Integer
Dim sObservacao As String
'##############################

On Error GoTo Erro_ParcelasReceber_Altera_CamposAlteraveis

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26294
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26295
        
    '############################################################
    'INSERIDO POR WAGNER
    sObservacao = String(STRING_MAX_NOME_ARQUIVO, 0)
    '############################################################
        
    '##################################################
    'ALTERADO POR WAGNER
    'Seleciona as Parcelas a alterar
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumParcela, DataVencimento, DataVencimentoReal, Cobrador, CarteiraCobranca, Saldo, Valor, Previsao, Observacao FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? AND NumParcela <= ? ORDER BY NumParcela", 0, lNumIntDoc, iNumParcela, dtDataVencimento, dtDataVencimentoReal, iCobrador, iCarteiraCobranca, dSaldo, dValor, iPrevisao, sObservacao, colParcelaReceber(1).lNumIntTitulo, STATUS_EXCLUIDO, iNumParcelas)
    If lErro <> AD_SQL_SUCESSO Then gError 26296
    '##################################################
    
    'Loop de alteração das Parcelas
    For iIndice = 1 To iNumParcelas
        
        Set objParcRec = colParcelaReceber(iIndice)
        
        'Tenta ler a Parcela de número iIndice
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 26297
        
        If lErro = AD_SQL_SEM_DADOS Then
        
            'Verifica se está baixada.
            lErro = CF("ParcelaReceberBaixada_Le_SemNumIntDoc", objParcRec)
            If lErro <> SUCESSO And lErro <> 28565 Then gError 108740
            If lErro = 28565 Then gError 26297
            
            'Se estiver baixada => Altera o restante, menos as parcelas.
            Exit For
            
        End If
        
        bAlterouValor = Abs(objParcRec.dValor - dValor) > DELTA_VALORMONETARIO
        
        'se o vencimento foi alterado
        '#########################################
        'ALTERADO POR WAGNER
        If objParcRec.iPrevisao <> iPrevisao Or objParcRec.sObservacao <> sObservacao Or dtDataVencimento <> objParcRec.dtDataVencimento Or dtDataVencimentoReal <> objParcRec.dtDataVencimentoReal Or bAlterouValor Then
        '#########################################

            If dtDataVencimento <> objParcRec.dtDataVencimento Or dtDataVencimentoReal <> objParcRec.dtDataVencimentoReal Or bAlterouValor Then
            
                'verifica se a parcela está em cobranca eletronica
                If iCobrador <> COBRADOR_PROPRIA_EMPRESA Then
                
                    objCobrador.iCodigo = iCobrador
                    lErro = CF("Cobrador_Le", objCobrador)
                    If lErro <> SUCESSO And lErro <> 19294 Then gError 59200
                    If lErro <> SUCESSO Then gError 59201
                
                    If objCobrador.iCobrancaEletronica = 1 Then
                    
                        objCarteiraCobrador.iCobrador = iCobrador
                        objCarteiraCobrador.iCodCarteiraCobranca = iCarteiraCobranca
                        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
                        If lErro <> SUCESSO Then gError 59201
                        
                        If objCarteiraCobrador.iComRegistro <> 0 Then gError 59202
                        
                    End If
                    
                End If

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
                
                    lErro = ParcelaRec_ObtemSaldoChequePre(objParcRec.lNumIntDoc, dSaldoChequePre)
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
                        
                'Altera os campos alteráveis
                '#########################################
                'ALTERADO POR WAGNER
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET DataVencimento = ?, DataVencimentoReal =?, Valor = ?, Saldo =?, Previsao = ?, Observacao = ? ", lComando, objParcRec.dtDataVencimento, objParcRec.dtDataVencimentoReal, objParcRec.dValor, objParcRec.dSaldo, objParcRec.iPrevisao, objParcRec.sObservacao)
                If lErro <> AD_SQL_SUCESSO Then gError 26299
                '#########################################
                
            Else
            
                'Altera os campos alteráveis
                '#########################################
                'ALTERADO POR WAGNER
                lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET DataVencimento = ?, DataVencimentoReal =?, Previsao = ?, Observacao = ? ", lComando, objParcRec.dtDataVencimento, objParcRec.dtDataVencimentoReal, objParcRec.iPrevisao, objParcRec.sObservacao)
                If lErro <> AD_SQL_SUCESSO Then gError 26299
                '#########################################
            
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
        
        Case 59200, 81880, 81888, 108740
        
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148895)

    End Select

    'fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Private Function ParcelasReceber_Altera_Descontos(ByVal colParcelaReceber As colParcelaReceber, ByVal iNumParcelas As Integer) As Long
'Altera os campos de Desconto das primeiras iNumParcelas parcelas da colecao
'As parcelas maiores que iNumParcelas, se houverem, serao novas e portanto serao incluidas e nao alteradas

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iIndice As Integer
Dim iStatus As Integer

On Error GoTo Erro_ParcelasReceber_Altera_Descontos

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26301
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 26302
    
    'Lê as Parcelas a Receber no BD
    lErro = Comando_ExecutarPos(lComando, "SELECT Status FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? AND NumParcela <= ? ORDER BY NumParcela", 0, iStatus, colParcelaReceber(1).lNumIntTitulo, STATUS_EXCLUIDO, iNumParcelas)
    If lErro <> AD_SQL_SUCESSO Then Error 26303
        
    'Loop de alteração das Parcelas
    For iIndice = 1 To iNumParcelas
        
        'Pega o Próximo registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26304
        If lErro = AD_SQL_SEM_DADOS Then Error 26305

        'Verifica se a Parcela já está baixada
        If iStatus <> STATUS_BAIXADO Then
            'Atualiza os dados de Desconto no BD
            lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET Desconto1Codigo = ? , Desconto1Ate = ?, Desconto1Valor = ?,Desconto2Codigo = ? , Desconto2Ate = ?, Desconto2Valor = ?,Desconto3Codigo = ? , Desconto3Ate = ?, Desconto3Valor = ?", lComando, colParcelaReceber(iIndice).iDesconto1Codigo, colParcelaReceber(iIndice).dtDesconto1Ate, colParcelaReceber(iIndice).dDesconto1Valor, colParcelaReceber(iIndice).iDesconto2Codigo, colParcelaReceber(iIndice).dtDesconto2Ate, colParcelaReceber(iIndice).dDesconto2Valor, colParcelaReceber(iIndice).iDesconto3Codigo, colParcelaReceber(iIndice).dtDesconto3Ate, colParcelaReceber(iIndice).dDesconto3Valor)
            If lErro <> AD_SQL_SUCESSO Then Error 26306
        End If
    Next
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    ParcelasReceber_Altera_Descontos = SUCESSO
    
    Exit Function
    
Erro_ParcelasReceber_Altera_Descontos:

    ParcelasReceber_Altera_Descontos = Err
    
    Select Case Err
    
        Case 26301, 26302
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 26303, 26304, 26305
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC1", Err, colParcelaReceber(1).lNumIntTitulo)
        
        Case 26306
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", Err, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148896)
        
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function ComissoesParcelas_Altera(ByVal colParcelaReceber As colParcelaReceber, ByVal colcolComissao As colcolComissao, ByVal iNumParcelas As Integer) As Long
'Altera as comissões das primeiras iNumParcelas parcelas da colecao
'As parcelas maiores que iNumParcelas, se houverem, serao novas e portanto serao incluidas e nao alteradas

Dim lErro As Long
Dim lComando As Long
Dim iStatus As Integer
Dim iComissaoPaga As Integer
Dim colComissao As colComissao
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dtDataGer As Date, iComissaoLiberada As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ComissoesParcelas_Altera

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26328

    'Para cada coleção de comissões
    For iIndice = 1 To iNumParcelas

        Set colComissao = colcolComissao(iIndice).colComissao

        'Ler as comissões existentes no bd
        lErro = Comando_ExecutarPos(lComando, "SELECT Status, DataGeracao FROM Comissoes WHERE TipoTitulo = ? AND NumIntDoc = ? ORDER BY NumIntDoc", 0, iStatus, dtDataGer, TIPO_PARCELA, colParcelaReceber(iIndice).lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then Error 26330

        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26331

        'Flag de Status
        iComissaoPaga = False
        iComissaoLiberada = False

        Do While lErro = AD_SQL_SUCESSO

            If iStatus = STATUS_LIBERADO Then
                iComissaoLiberada = True
            End If
            
            'Verifica se a comissão esta baixada
            If iStatus = STATUS_BAIXADO Then
                iComissaoPaga = True
                Exit Do
            End If
            
            'pega o próximo registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26332

        Loop
        
        'Se nenhuma comissão foi paga
        If iComissaoPaga = False Then
        
            'Alterado para manter a liberação anterior, quer tenha sido na baixa ou na emissão
            If iComissaoLiberada Then
                For Each objComissao In colComissao
                    objComissao.dtDataGeracao = dtDataGer
                    objComissao.iStatus = STATUS_LIBERADO
                Next
            End If
            
            'Altera as Comissões no BD
            lErro = CF("Comissoes_Altera", colComissao, TIPO_PARCELA, colParcelaReceber(iIndice).lNumIntDoc)
            If lErro <> SUCESSO Then Error 26333

        End If

    Next

    'Fecha o comando
    Call Comando_Fechar(lComando)

    ComissoesParcelas_Altera = SUCESSO

    Exit Function

Erro_ComissoesParcelas_Altera:

    ComissoesParcelas_Altera = Err

    Select Case Err

        Case 26328
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 26330, 26331, 26332
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOES", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148897)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloReceber_Testa_Repeticao(ByVal objTituloReceber As ClassTituloReceber, Optional ByVal lComandoP As Long = 0) As Long
'Teste se há Titulo a Receber (lançado ou baixado)
'com os mesmos Número e DataEmissao dentro da faixa PERIODO_EMISSAO
'Função retorna SUCESSO se NÃO há ou se há e deseja prosseguir na INSERÇÃO de nova NotaFiscal

Dim lErro As Long
Dim lNumIntDoc As Long, dtDataEmissao As Date
Dim lComando As Long, sSiglaDocumento As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TituloReceber_Testa_Repeticao

    If lComandoP <> 0 Then
        lComando = lComandoP
    Else
        'Abre o comando
        lComando = Comando_Abrir()
        If lComando = 0 Then Error 26337
    End If

    If objTituloReceber.iFilialEmpresa = EMPRESA_TODA Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    'Seleciona TituloReceber Baixado com o mesmo Número e DataEmissao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRecBaixados WHERE SiglaDocumento = ? AND NumTitulo = ? AND DataEmissao = ? AND Status <> ? AND FilialEmpresa = ?", lNumIntDoc, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 59148

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59149
    If lErro = AD_SQL_SUCESSO Then Error 59150

    If objTituloReceber.dtDataEmissao <> DATA_NULA Then
    
        'Seleciona TituloReceber Baixado com o mesmo Número e DataEmissao dentro da faixa PERIODO_EMISSAO
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosRecBaixados WHERE SiglaDocumento = ? AND NumTitulo = ? AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND FilialEmpresa = ?", dtDataEmissao, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao - PERIODO_EMISSAO, objTituloReceber.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
        
    Else
    
        'Seleciona TituloReceber Baixado com o mesmo Número e DataEmissao independente de emissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosRecBaixados WHERE SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? AND FilialEmpresa = ?", dtDataEmissao, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 26344

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26345
    
    If lErro = AD_SQL_SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_TITULORECEBER_BAIXADO_MESMO_NUMERO", objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, dtDataEmissao)
        If vbMsgRes = vbNo Then Error 26346
    
    End If
    
    'Seleciona TituloReceber com o mesmo Número e tipo e data de emissao
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM TitulosRec WHERE SiglaDocumento = ? AND NumTitulo = ? AND DataEmissao = ? AND Status <> ? AND FilialEmpresa = ?", lNumIntDoc, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 59151

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59152
    If lErro = AD_SQL_SUCESSO Then Error 59153
    
    If objTituloReceber.dtDataEmissao <> DATA_NULA Then
        
        'Seleciona TituloReceber com o mesmo Número e DataEmissao dentro da faixa PERIODO_EMISSAO
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosRec WHERE SiglaDocumento = ? AND NumTitulo = ? AND DataEmissao > ? AND DataEmissao < ? AND Status <> ? AND FilialEmpresa = ?", dtDataEmissao, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao - PERIODO_EMISSAO, objTituloReceber.dtDataEmissao + PERIODO_EMISSAO, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
        
    Else
    
        'Seleciona TituloReceber com o mesmo Número independente de DataEmissao
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM TitulosRec WHERE SiglaDocumento = ? AND NumTitulo = ? AND Status <> ? AND FilialEmpresa = ?", dtDataEmissao, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, STATUS_EXCLUIDO, objTituloReceber.iFilialEmpresa)
        
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 26341

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 26342
    
    If lErro = AD_SQL_SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_TITULORECEBER_MESMO_NUMERO", objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, dtDataEmissao)
        If vbMsgRes = vbNo Then Error 26343
    
    End If
           
    'Fecha o comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)
    
    TituloReceber_Testa_Repeticao = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Testa_Repeticao:

    TituloReceber_Testa_Repeticao = Err

    Select Case Err
    
        Case 26337
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 26343, 26346
        
        Case 26341, 26342, 59151, 59152
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
        
        Case 26344, 26345, 59148, 59149
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
        
        Case 59150
            Call Rotina_Erro(vbOKOnly, "ERRO_TITRECBAIXA_JA_CADASTRADO", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
        
        Case 59153
            Call Rotina_Erro(vbOKOnly, "ERRO_TITREC_JA_CADASTRADO", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148898)
            
    End Select
    
    'Fecha o comando
    If lComandoP = 0 Then Call Comando_Fechar(lComando)
    
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
        
        'Insere a Parcela no BD
        With colParcelaReceber(iIndice)
        '######################################
        'ALTERADO POR WAGNER
            lErro = Comando_Executar(lComando, "INSERT INTO ParcelasRec (NumIntDoc,NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Saldo,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo, Desconto1Ate, Desconto1Valor,Desconto2Codigo, Desconto2Ate, Desconto2Valor,Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, Previsao, Observacao, ValorOriginal, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            .lNumIntDoc, .lNumIntTitulo, .iNumParcela, STATUS_ABERTO, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca, .sNumTitCobrador, .iProxSeqBaixa, .iProxSeqOcorr, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .iAceite, .iDescontada, .iPrevisao, .sObservacao, .dValorOriginal, .iTipoPagto, .iCodConta, .dtDataCredito, .dtDataEmissaoCheque, .iBancoCheque, .sAgenciaCheque, .sContaCorrenteCheque, .lNumeroCheque, .dtDataDepositoCheque, .iAdmMeioPagto, .iParcelamento, .sNumeroCartao, .dtValidadeCartao, .sAprovacaoCartao, .dtDataTransacaoCartao)
        '######################################
        End With
        If lErro <> AD_SQL_SUCESSO Then Error 26360

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148899)
            
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TituloReceber_Exclui(ByVal objTituloReceber As ClassTituloReceber, ByVal objContabil As ClassContabil) As Long
'Exclui o Título à Receber do Banco de Dados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TituloReceber_Exclui

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 26443

    lErro = CF("TituloReceber_Exclui_EmTransacao", objTituloReceber, objContabil)
    If lErro <> SUCESSO Then Error 62137
       
    'Encerra transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 26453
    
    TituloReceber_Exclui = SUCESSO
    
    Exit Function

Erro_TituloReceber_Exclui:

    TituloReceber_Exclui = Err
    
    Select Case Err
        
        Case 26443
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 62137
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148900)
            
    End Select

    'Encerra a transação e os comandos
    Call Transacao_Rollback
    
    Exit Function

End Function

Function ParcelasReceber_Exclui(ByVal objTituloReceber As ClassTituloReceber) As Long
'Exclui as parcelas relacionadas ao Título passado por parâmetro
'FUNÇÃO CHAMADA DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lNumero As Long
Dim iNumParcela As Integer
Dim tParcelaReceber As typeParcelaReceber
Dim objCarteiraCobrador As New ClassCarteiraCobrador

On Error GoTo Erro_ParcelasReceber_Exclui
    
    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26444
        
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26445
    
    If objTituloReceber.iExclusao = 1 Then
    
        'Seleciona as Parcelas referentes ao Título passado por ordem de Vencimento
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ProxSeqBaixa, Cobrador, CarteiraCobranca, Valor, Saldo FROM ParcelasRec WHERE NumIntTitulo = ? ORDER BY DataVencimento", 0, tParcelaReceber.lNumIntDoc, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.dValor, tParcelaReceber.dSaldo, objTituloReceber.lNumIntDoc)
        If lErro <> AD_SQL_SUCESSO Then gError 26455
    
    Else
        
        'Seleciona as Parcelas referentes ao Título passado por ordem de Vencimento
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, ProxSeqBaixa, Cobrador, CarteiraCobranca, Valor, Saldo FROM ParcelasRec WHERE NumIntTitulo = ? AND Status <> ? ORDER BY DataVencimento", 0, tParcelaReceber.lNumIntDoc, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.dValor, tParcelaReceber.dSaldo, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 26455
    
    End If
    
    iNumParcela = 1
    
    'Lê a primeira Parcela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26456
   
    'Loop de exclusão lógica das Parcelas
    Do While lErro = AD_SQL_SUCESSO
    
        'Faz o "lock" da Parcela
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 26447
        
        'Verifica se a parcela já sofreu alguma baixa
        If Abs(tParcelaReceber.dSaldo - tParcelaReceber.dValor) > DELTA_VALORMONETARIO Then gError 95377

        'a parcela tem que estar em cobranca na propria empresa. Isto é particularmente importante p/evitar "esquecer" uma parcela em cobranca eletronica.
        If tParcelaReceber.iCobrador <> COBRADOR_PROPRIA_EMPRESA Then gError 59218
        
        'Verifica se existe algum ChequePre relacionado a Parcela
        lErro = ParcelaRec_SemChequePre(tParcelaReceber.lNumIntDoc)
        If lErro <> SUCESSO Then gError 81887
                
        'atualizar o saldo da carteira devida ("em cheque-pre" ou "em carteira").
        objCarteiraCobrador.iCobrador = tParcelaReceber.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = tParcelaReceber.iCarteiraCobranca
        
        lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, -1, -tParcelaReceber.dValor)
        If lErro <> SUCESSO Then gError 15993
        
        'Exclui as comissões relacionadas a Parcela desde que nao estejam baixadas
        lErro = CF("Comissoes_Exclui", TIPO_PARCELA, tParcelaReceber.lNumIntDoc, 0)
        If lErro <> SUCESSO Then gError 26452
        
        '################################################################
        'Inserido por Wagner 26/05/2006
        lErro = CF("Exclusao_Valida", VALIDAEXCLUSAO_CODIGO_PARCELARECEBER, tParcelaReceber.lNumIntDoc)
        If lErro <> SUCESSO Then gError 177947
        '################################################################
        
        If objTituloReceber.iExclusao = 1 Then
                
            'Faz a exclusão lógica da Parcela
            lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasRec", lComando)
            If lErro <> AD_SQL_SUCESSO Then gError 83950
            
        Else
        
            'Faz a exclusão lógica da Parcela
            lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET Status = ? ", lComando, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 26453
        
        End If
        
        iNumParcela = iNumParcela + 1
        
        'Lê a próxima Parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26454
            
    Loop
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ParcelasReceber_Exclui = SUCESSO
    
    Exit Function
    
Erro_ParcelasReceber_Exclui:

    ParcelasReceber_Exclui = gErr
    
    Select Case gErr
    
        Case 15993, 26452, 177947
        
        Case 95377
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COM_BAIXA", gErr, iNumParcela, objTituloReceber.lNumTitulo)
        
        Case 26444, 26445, 26446
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                        
        Case 26455, 26456, 26454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC1", gErr, objTituloReceber.lNumIntDoc)
        
        Case 26447
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", gErr, objTituloReceber.lNumIntDoc, iNumParcela)
        
        Case 26449, 26450
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr, tParcelaReceber.lNumIntCheque)
        
        Case 26451
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUESPRE", gErr, tParcelaReceber.lNumIntCheque)
            
        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", gErr, iNumParcela, objTituloReceber.lNumTitulo)
        
        Case 59218
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COBRANCA_EMPRESA", gErr, objTituloReceber.lNumTitulo, iNumParcela)
        
        Case 83950
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASREC", gErr, iNumParcela, objTituloReceber.lNumTitulo)
        
        Case 81887

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148901)
            
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
        
    Exit Function
    
End Function

Function TituloReceber_Insere(ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelaReceber As colParcelaReceber, ByVal colComissaoEmissao As colComissao, ByVal colcolComissao As colcolComissao) As Long
'Insere o Título à Receber na Tabela de Titulos a Receber
    
Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim lComando1 As Long
Dim objComissao As ClassComissao
Dim objParcelaReceber As ClassParcelaReceber
Dim lNumeroFatura As Long

On Error GoTo Erro_TituloReceber_Insere

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 26393
        
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 26353
    
    'Se for Fatura a Receber
    If objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER Then
                
        'Gera o Número da Fatura
        lErro = CF("GeracaoFatura_Automatico_Trans", lNumeroFatura)
        If lErro <> SUCESSO Then Error 61464
        
        'Coloca no Número do Título
        objTituloReceber.lNumTitulo = lNumeroFatura
        
    End If
    
    'Verifica a existencia de um Título no BD com as mesmas características
    lErro = CF("TituloReceber_Testa_Repeticao", objTituloReceber)
    If lErro <> SUCESSO Then Error 26336
    
    lErro = CF("TitRec_Testa_Repeticao_Insert_Cust", objTituloReceber)
    If lErro <> SUCESSO Then Error 26336
    
    'Gera um novo código interno para o Título
    lErro = CF("TituloReceber_Automatico", lNumIntDoc)
    If lErro <> SUCESSO Then Error 26347
    
    objTituloReceber.lNumIntDoc = lNumIntDoc
    
    'Verifica se Cliente e Filial estão cadastrados
    lErro = CF("FilialCliente_Lock", objTituloReceber.lCliente, objTituloReceber.iFilial)
    If lErro <> SUCESSO And lErro <> 43050 Then Error 26229
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 43050 Then Error 26230
    
    If objTituloReceber.iCondicaoPagto <> 0 Then
        'Verifica se Condição de Pagamento está cadastrada
        lErro = CF("CondicaoPagto_Lock", lComando, objTituloReceber.iCondicaoPagto)
        If lErro <> SUCESSO Then Error 26354
    End If
    
    If objTituloReceber.iFilialEmpresa = 0 Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    'Insere o Título na tabela TitulosRec
    lErro = Comando_Executar(lComando1, "INSERT INTO TitulosRec (NumIntDoc, FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, DataEmissao, Status, Saldo, NumParcelas, Valor, ValorIRRF, ValorISS, ISSRetido, ValorINSS, INSSRetido, PercJurosDiarios, PercMulta, Observacao, CondicaoPagto, DataRegistro, PISRetido, COFINSRetido, CSLLRetido, Moeda, Ccl, Natureza, CpoGenerico1, ReajustePeriodicidade, ReajusteBase, ReajustadoAte) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        lNumIntDoc, objTituloReceber.iFilialEmpresa, objTituloReceber.lCliente, objTituloReceber.iFilial, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo, objTituloReceber.dtDataEmissao, STATUS_LANCADO, objTituloReceber.dSaldo, objTituloReceber.iNumParcelas, objTituloReceber.dValor, objTituloReceber.dValorIRRF, objTituloReceber.dValorISS, objTituloReceber.dISSRetido, objTituloReceber.dValorINSS, objTituloReceber.iINSSRetido, objTituloReceber.dPercJurosDiario, objTituloReceber.dPercMulta, objTituloReceber.sObservacao, objTituloReceber.iCondicaoPagto, objTituloReceber.dtDataRegistro, objTituloReceber.dPISRetido, objTituloReceber.dCOFINSRetido, objTituloReceber.dCSLLRetido, objTituloReceber.iMoeda, objTituloReceber.sCcl, objTituloReceber.sNatureza, objTituloReceber.lCpoGenerico1, objTituloReceber.iReajustePeriodicidade, objTituloReceber.dtReajusteBase, objTituloReceber.dtReajustadoAte)
    If lErro <> AD_SQL_SUCESSO Then Error 26394

    If objTituloReceber.iStatus <> STATUS_PREVISAO Then

        For Each objComissao In colComissaoEmissao
            
            objComissao.iTipoTitulo = TIPO_TITULO_RECEBER
            objComissao.lNumIntDoc = lNumIntDoc
            objComissao.iStatus = STATUS_LIBERADO
            objComissao.dtDataGeracao = objTituloReceber.dtDataEmissao
            objComissao.dtDataBaixa = DATA_NULA
    
        Next
    
        'Grava as comissões do Título
        lErro = CF("Comissoes_Grava", colComissaoEmissao)
        If lErro <> SUCESSO Then Error 26348
        
    End If
        
    For Each objParcelaReceber In colParcelaReceber
        objParcelaReceber.lNumIntTitulo = lNumIntDoc
    Next
    
    'Grava as Parcelas a Receber
    lErro = CF("ParcelasReceber_Insere", colParcelaReceber, colcolComissao)
    If lErro <> SUCESSO Then Error 26349
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    TituloReceber_Insere = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Insere:

    TituloReceber_Insere = Err
    
    Select Case Err
    
        Case 26230
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objTituloReceber.iFilial, objTituloReceber.lCliente)

        Case 26393, 26353
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 26229, 26336, 26347, 26348, 26349, 26354, 61464
        
        Case 26394
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TITULOS_REC", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148902)
            
    End Select
    
    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function

Public Function TitRec_Testa_Repeticao_Insert_Cust(ByVal objTitRec As ClassTituloReceber)
    TitRec_Testa_Repeticao_Insert_Cust = SUCESSO
End Function

Private Function ParcelasReceber_Traz_DadosBD(ByVal colParcelaReceberBD As colParcelaReceber, ByVal colcolDescontoBD As colcolDesconto, ByVal colcolComissaoBD As colcolComissao) As Long
'Traz do BD os Dados particulares das Parcela  a Receber

Dim lErro As Long
Dim objParcelaReceber As ClassParcelaReceber
Dim colDesconto As colDesconto
Dim colComissao As colComissao

On Error GoTo Erro_ParcelasReceber_Traz_DadosBD

    'Para cada Parcela
    For Each objParcelaReceber In colParcelaReceberBD
        
        Set colDesconto = New colDesconto
        Set colComissao = New colComissao
        
        'Para para coldesconto os descontos cadastrados
        If objParcelaReceber.iDesconto1Codigo > 0 Then colDesconto.Add objParcelaReceber.iDesconto1Codigo, objParcelaReceber.dtDesconto1Ate, objParcelaReceber.dtDesconto1Ate
        If objParcelaReceber.iDesconto2Codigo > 0 Then colDesconto.Add objParcelaReceber.iDesconto2Codigo, objParcelaReceber.dtDesconto2Ate, objParcelaReceber.dtDesconto2Ate
        If objParcelaReceber.iDesconto3Codigo > 0 Then colDesconto.Add objParcelaReceber.iDesconto3Codigo, objParcelaReceber.dtDesconto3Ate, objParcelaReceber.dtDesconto3Ate
        
        colcolDescontoBD.Add colDesconto
        
        'Lê as comissões das Parcelas Gravadas
        lErro = Comissoes_Le_Locando(objParcelaReceber.lNumIntDoc, colComissao, TIPO_PARCELA)
        If lErro <> SUCESSO Then Error 26265
        
        colcolComissaoBD.Add colComissao
        
    Next
    
    ParcelasReceber_Traz_DadosBD = SUCESSO
    
    Exit Function
    
Erro_ParcelasReceber_Traz_DadosBD:

    ParcelasReceber_Traz_DadosBD = Err

    Select Case Err
    
        Case 26265, 26266
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148903)
            
    End Select
    
    Exit Function

End Function

Function TituloReceber_Altera(ByVal colParcelaReceber As colParcelaReceber, ByVal colParcelaReceberBD As colParcelaReceber, ByVal colcolDesconto As colcolDesconto, ByVal colcolDescontoBD As colcolDesconto, ByVal colcolComissao As colcolComissao, ByVal colcolComissaoBD As colcolComissao) As Long
'Altera os dados do Título à Receber

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objParcelaReceber As ClassParcelaReceber
Dim vbMsgRes As VbMsgBoxResult
Dim iBaixado As Integer
Dim objComissao As ClassComissao, iMenorNumParc As Integer

On Error GoTo Erro_TituloReceber_Altera
    
    iMenorNumParc = IIf(colParcelaReceberBD.Count < colParcelaReceber.Count, colParcelaReceberBD.Count, colParcelaReceber.Count)
    
    iIndice = 0
    
    'Para cada Parcela
    For Each objParcelaReceber In colParcelaReceber
        
        iIndice = iIndice + 1
    
        'apenas para as parcelas que serao alteradas
        If iIndice <= iMenorNumParc Then
        
            'Pesquisa se ProxSeqBaixa da Parcela é maior que 1
            If objParcelaReceber.iProxSeqBaixa > 1 Then
            
                'Verifica se o número de descontos da Tela e do BD são iguais
                If colcolDesconto(iIndice).colDesconto.Count <> colcolDescontoBD(iIndice).colDesconto.Count Then
                    
                    'Se não tiverem avisa que parcela com baixa não pode ter número de descontos alterados. e Pergunta se desja prosseguir na alteração
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PARCELA_COM_BAIXA_NAO_ALTERAVEL", iIndice)
                    If vbMsgRes = vbNo Then Error 26267
                
                End If
                        
                For iIndice1 = 1 To colcolDesconto(iIndice).colDesconto.Count
                    
                    'Verifica se os dados de cada objDesconto são válidos
                    If colcolDesconto(iIndice).colDesconto(iIndice1).iCodigo <> colcolDescontoBD(iIndice).colDesconto(iIndice1).iCodigo Or _
                       colcolDesconto(iIndice).colDesconto(iIndice1).dValor <> colcolDescontoBD(iIndice).colDesconto(iIndice1).dValor Or _
                       colcolDesconto(iIndice).colDesconto(iIndice1).dtData <> colcolDescontoBD(iIndice).colDesconto(iIndice1).dtData Then
                       
                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PARCELA_COM_BAIXA_DESCONTO_INALTERAVEL", iIndice)
                        If vbMsgRes = vbNo Then Error 26268
                        
                    End If
                Next
            
            End If
            
            'Flag de Baixa
            iBaixado = False
            
            'Verifica se existe alguma comissão baixada
            For Each objComissao In colcolComissaoBD(iIndice).colComissao
                If objComissao.iStatus = STATUS_BAIXADO Then
                    iBaixado = True
                    Exit For
                End If
            Next
            
            'Se houve alguma baixa
            If iBaixado = True Then
                
                If colcolComissao(iIndice).colComissao.Count <> colcolComissaoBD(iIndice).colComissao.Count Then
                    
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_NUM_COMISSOES_NAO_ALTERAVEL")
                    If vbMsgRes = vbNo Then Error 26271
                
                End If
                
                'Verifica se os dados de colcolcomissao são os mesmos cadastrados no BD
                For iIndice1 = 1 To colcolComissao(iIndice).colComissao.Count
                    colcolComissao(iIndice).colComissao(iIndice1).iTipoTitulo = TIPO_PARCELA
                    colcolComissao(iIndice).colComissao(iIndice1).iTipoTitulo = colParcelaReceberBD(iIndice).lNumIntDoc
                    
                    If colcolComissao(iIndice).colComissao(iIndice1).iCodVendedor <> colcolComissaoBD(iIndice).colComissao(iIndice1).iCodVendedor Or colcolComissao(iIndice).colComissao(iIndice1).dPercentual <> colcolComissaoBD(iIndice).colComissao(iIndice1).dPercentual Or _
                      colcolComissao(iIndice).colComissao(iIndice1).dValorBase <> colcolComissaoBD(iIndice).colComissao(iIndice1).dValorBase Or colcolComissao(iIndice).colComissao(iIndice1).dValor <> colcolComissaoBD(iIndice).colComissao(iIndice1).dValor Then
                      
                        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_COMISSAO_PARCELA_PAGA", iIndice)
                        If vbMsgRes = vbNo Then Error 26272
                    End If
                Next
            End If
    
        End If
    
    Next
    
    'Faz as alterações dos campos possíveis na tabela de Parcelas
    lErro = ParcelasReceber_Altera_CamposAlteraveis(colParcelaReceber, iMenorNumParc, False)
    If lErro <> SUCESSO Then Error 26293
    
    'Faz a alteração nos descontos das parcelas
    lErro = ParcelasReceber_Altera_Descontos(colParcelaReceber, iMenorNumParc)
    If lErro <> SUCESSO Then Error 26300
                  
    'Faz as alterações no Bd das comissões das parcela
    lErro = CF("ComissoesParcelas_Altera", colParcelaReceber, colcolComissao, iMenorNumParc)
    If lErro <> SUCESSO Then Error 26327

    'excluir parcelas que estavam no bd e nao estao na colecao em memoria
    lErro = ParcelasRec_ExcluiAltTit(colParcelaReceberBD, colParcelaReceber.Count)
    If lErro <> SUCESSO Then Error 26327
    
    'se o numero de parcelas aumentou
    If colParcelaReceberBD.Count < colParcelaReceber.Count Then
        'incluir parcelas que nao estavam no bd e estao na colecao em memoria
        lErro = CF("ParcelasReceber_Insere", colParcelaReceber, colcolComissao, colParcelaReceberBD.Count + 1)
        If lErro <> SUCESSO Then Error 26327
    End If

    TituloReceber_Altera = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Altera:

    TituloReceber_Altera = Err
    
    Select Case Err
           
        Case 26267, 26268, 26269, 26271, 26272, 26293, 26300, 26307, 26326, 26327
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148904)
            
    End Select
    
    Exit Function
        
End Function

Function InstrCobranca_Lock(ByVal iInstrucao As Integer, ByVal lComando As Long) As Long
'faz lock shared de uma Instrucao de Cobranca  para garantir a sua existencia

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_InstrCobranca_Lock

    'Pesquisa a Instrucao de Cobranca
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM TiposInstrCobranca WHERE Codigo = ?", iCodigo, iInstrucao)
    If lErro <> AD_SQL_SUCESSO Then Error 40601

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 40602

    'Se não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 40603

    '"Lock" da Instrucao
    lErro = Comando_LockShared(lComando)
    If lErro <> SUCESSO Then Error 40604
    
    InstrCobranca_Lock = SUCESSO
    
    Exit Function
    
Erro_InstrCobranca_Lock:

    InstrCobranca_Lock = Err

    Select Case Err

        Case 40601, 40602
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOS_INSTRUCAO_COBRANCA", Err, iInstrucao)
        
        Case 40603
        
        Case 40604
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_INSTRUCAO_COBRANCA", Err, iInstrucao)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148905)

    End Select

    Exit Function
    
End Function

Function Parcela_Com_Desconto(ByVal colDesconto As colDesconto, ByVal dValorParcela As Double, ByVal dtDataDeposito As Date, ByVal dtDataVencimento As Date, dValorComDesconto As Double) As Long
'Retorna o Valor da Parcela com o desconto mais apropriado dos descontos contidos em colDescoltos, já aplicados

Dim lErro As Long
Dim objDesconto As New ClassDesconto
Dim iDias As Integer
Dim dtDataMelhor As Date
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dtDataUtil As Date

On Error GoTo Erro_Parcela_Com_Desconto

    'Inicializa o Valor dValorComDesconto com o valor da Parcela
    dValorComDesconto = dValorParcela
    
    'Inicializa a data melhor de desconto com uma data nula
    dtDataMelhor = DATA_NULA
    iIndice = 0
    
    'Para cada desconto da coleção
    For Each objDesconto In colDesconto
        
        iIndice = iIndice + 1
        If dtDataDeposito <= objDesconto.dtData Then
            If dtDataMelhor <> DATA_NULA Then
                If objDesconto.dtData < dtDataMelhor Then
                    dtDataMelhor = objDesconto.dtData
                    iIndice2 = iIndice
                End If
            Else
                dtDataMelhor = objDesconto.dtData
                iIndice2 = iIndice
            End If
        End If
                
    Next
    
    If dtDataMelhor <> DATA_NULA And dtDataMelhor < dtDataVencimento Then
    
        Set objDesconto = colDesconto(iIndice2)
    
        'Verifica qual o Tipo de Desconto
        Select Case objDesconto.iCodigo
        
            'Valor fixo até a data informada
            Case VALOR_FIXO
                dValorComDesconto = dValorComDesconto - objDesconto.dValor
            
            'Percentual até a data informada
            Case Percentual
                dValorComDesconto = dValorComDesconto - (dValorComDesconto * objDesconto.dValor / 100)
                
            'Valor por antecipação de dia corrido
            Case VALOR_ANT_DIA
                iDias = dtDataVencimento - dtDataDeposito
                dValorComDesconto = dValorComDesconto - iDias * objDesconto.dValor
            
            'Percentual sobre o valor por dia corrido
            Case PERC_ANT_DIA
                iDias = dtDataVencimento - dtDataDeposito
                dValorComDesconto = dValorComDesconto - (dValorComDesconto * iDias * objDesconto.dValor / 100)
                
            'Valor por antecipação de dia útil
            Case VALOR_ANT_DIA_UTIL
                iDias = dtDataVencimento - dtDataDeposito
                If iDias > 0 Then
                    lErro = CF("Dias_Uteis", dtDataDeposito, dtDataVencimento, iDias)
                    If lErro <> SUCESSO Then Error 26212
                
                    dValorComDesconto = dValorComDesconto - (iDias * objDesconto.dValor)
                
                End If
            
            'Percentual sobre o valor por dia útil
            Case PERC_ANT_DIA_UTIL
                iDias = dtDataVencimento - dtDataDeposito
                If iDias > 0 Then
                    lErro = CF("Dias_Uteis", dtDataDeposito, dtDataVencimento, iDias)
                    If lErro <> SUCESSO Then Error 26213
                
                    dValorComDesconto = dValorComDesconto - (dValorComDesconto * iDias * objDesconto.dValor / 100)
                
                End If

        End Select
        
    End If
    
    If dValorComDesconto <= 0 Then dValorComDesconto = 0

    Parcela_Com_Desconto = SUCESSO

    Exit Function

Erro_Parcela_Com_Desconto:

    Parcela_Com_Desconto = Err

    Select Case Err

        Case 26212, 26213

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148906)

    End Select
            
    Exit Function
            
End Function

Private Function ChequePre_Inserir(ByVal objChequePre As ClassChequePre) As Long
'Rotina responsavel pela insercao de Cheque Pre
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long, lCliente As Long
Dim lComando As Long
Dim lNumProxChequePre As Long, lNumIntCheque As Long
Dim lSequencialBack As Long

On Error GoTo Erro_ChequePre_Inserir

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 17715

    'Pega o numero do proximo cheque pre
    lErro = CF("ChequePre_Automatico", lNumProxChequePre)
    If lErro <> SUCESSO Then gError 59159

    'gera um novo número interno de itens de cupom fiscal
    lErro = CF("Config_ObterNumInt", "CPRConfig", "COD_PROX_CHEQUE_BACKOFFICE", lSequencialBack, 1)
    If lErro <> SUCESSO Then gError 126223

    objChequePre.lNumIntCheque = lNumProxChequePre
    
    'Grava o ChequePre na tabela de ChequePre
    lErro = Comando_Executar(lComando, "INSERT INTO ChequePre (NumIntCheque,FilialEmpresa, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, SequencialBack, DataEmissao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objChequePre.lNumIntCheque, objChequePre.iFilialEmpresa, objChequePre.lCliente, objChequePre.iFilial, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero, objChequePre.dtDataDeposito, objChequePre.dValor, objChequePre.lNumBordero, lSequencialBack, objChequePre.dtDataEmissao)
    If lErro <> AD_SQL_SUCESSO Then Error 17709
   
    'Libera comando
    Call Comando_Fechar(lComando)

    ChequePre_Inserir = SUCESSO

    Exit Function

Erro_ChequePre_Inserir:

    ChequePre_Inserir = Err

    Select Case Err

        Case 17715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 17709
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CHEQUESPRE", Err, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 59159
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148907)
        
    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ChequePre_Automatico(lNumIntDoc As Long) As Long
'Gera o próximo Número Interno de um ChequePre
'CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_ChequePre_Automatico

    'Gera o número automático
    lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_CHEQUE_PRE", lNumAuto)
    If lErro <> SUCESSO Then Error 26350

    lNumIntDoc = CLng(lNumAuto)
    
    ChequePre_Automatico = SUCESSO
    
    Exit Function
    
Erro_ChequePre_Automatico:

    ChequePre_Automatico = Err
    
    Select Case Err
    
        Case 26350
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148908)
            
    End Select

    Exit Function

End Function

Function PadraoCobranca_Automatico(sCodigo As String) As Long
'retorna o código do próximo PadraoCobranca disponivel

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_PadraoCobranca_Automatico

    'Obtém o número Automático do BD
    lErro = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_PADRAO_COBRANCA", "PadroesCobranca", "Codigo", lNumAuto)
    If lErro <> SUCESSO Then Error 40600
    
    sCodigo = CStr(lNumAuto)

    PadraoCobranca_Automatico = SUCESSO

    Exit Function

Erro_PadraoCobranca_Automatico:

    PadraoCobranca_Automatico = Err

    Select Case Err

        Case 40600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148909)

    End Select
    
    Exit Function

End Function

Function TituloReceber_Automatico(lNumIntDoc) As Long
'Gera o próximo Número Interno de um Título a Receber
'CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_TituloReceber_Automatico

    'Gera o número automático
    lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_TIT_REC", lNumAuto)
    If lErro <> SUCESSO Then Error 26352

    lNumIntDoc = CLng(lNumAuto)
    
    TituloReceber_Automatico = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Automatico:

    TituloReceber_Automatico = Err
    
    Select Case Err
    
        Case 26352
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148910)
            
    End Select

    Exit Function

End Function

Function ParcelaReceber_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Gera o próximo Número Interno a ser utilizado por uma Parcela à Receber
'CHAMADA EM TRANSAÇÃO

Dim lErro As Long
Dim lNumAuto As Long

On Error GoTo Erro_ParcelaReceber_Automatico

    'Gera o número automático
    lErro = CF("Config_ObterNumInt", "CPRConfig", "NUM_PROX_PARC_REC", lNumAuto, iQtde)
    If lErro <> SUCESSO Then Error 26351

    lNumIntDoc = CLng(lNumAuto)
    
    ParcelaReceber_Automatico = SUCESSO
    
    Exit Function
    
Erro_ParcelaReceber_Automatico:

    ParcelaReceber_Automatico = Err
    
    Select Case Err
    
        Case 26351
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148911)
            
    End Select

    Exit Function

End Function

Function CPR_Gera(ByVal objNFiscal As ClassNFiscal, iClasseDocCPR As Integer, lNumIntDocCPR As Long) As Long
'Gera conta a Pagar/Receber vinculada à Nota Fiscal.
'Devolve ClasseDoc e NumIntDoc de Conta a Pagar/Receber gravada.

Dim lErro As Long
Dim sTipoDocCPR As String
Dim lComando As Long

On Error GoTo Erro_CPR_Gera

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46191
    
    sTipoDocCPR = String(STRING_TIPO_DOC_SIGLA, 0)
    
    lErro = Comando_Executar(lComando, "SELECT TiposDeDocumento.Sigla, TiposDeDocumento.ClasseDocCPR FROM TiposDeDocumento, TiposDocInfo WHERE TiposDeDocumento.Sigla = TiposDocInfo.TipoDocCPR AND TiposDocInfo.Codigo = ?", sTipoDocCPR, iClasseDocCPR, objNFiscal.iTipoNFiscal)
    If lErro <> AD_SQL_SUCESSO Then Error 46192
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 46193
    
    If lErro = AD_SQL_SUCESSO Then
    
        Select Case iClasseDocCPR
        
            Case CPR_TITULO_PAGAR
                
                If ISSerieEletronica(objNFiscal.sSerie) Then
                    sTipoDocCPR = TIPODOC_NFE_FATURA_PAGAR
                End If
                
                lErro = CF("TituloPagar_Gera", objNFiscal, sTipoDocCPR, lNumIntDocCPR)
                If lErro <> SUCESSO Then Error 46194
                
            Case CPR_TITULO_RECEBER
            
                If ISSerieEletronica(objNFiscal.sSerie) Then
                    sTipoDocCPR = TIPODOC_NFE_FATURA_RECEBER
                End If
            
                lErro = CF("TituloReceber_Gera", objNFiscal, sTipoDocCPR, lNumIntDocCPR)
                If lErro <> SUCESSO Then Error 46195
                
            Case CPR_NF_PAGAR
            
                lErro = CF("NFPagar_Gera", objNFiscal, sTipoDocCPR, lNumIntDocCPR)
                If lErro <> SUCESSO Then Error 46196
                
            Case CPR_DEBITO_RECEBER
                lErro = CF("DebitoRecCli_Gera", objNFiscal, sTipoDocCPR, lNumIntDocCPR)
                If lErro <> SUCESSO Then Error 46538
                
            Case CPR_CREDITO_PAGAR
                lErro = CF("CreditoPagar_Gera", objNFiscal, sTipoDocCPR, lNumIntDocCPR)
                If lErro <> SUCESSO Then Error 51629
                
        End Select
    
    Else
        
        iClasseDocCPR = 0
        lNumIntDocCPR = 0
    
    End If
    
    Call Comando_Fechar(lComando)
    
    CPR_Gera = SUCESSO
    
    Exit Function
    
Erro_CPR_Gera:

    CPR_Gera = Err
    
    Select Case Err
    
        Case 46191
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 46192, 46193
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDEDOCUMENTO", Err)
        
        Case 46194 To 46196, 46538, 51629
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148912)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TituloPagar_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long
'registra nota fiscal fatura a pagar a partir dos dados em objNFiscal vinculada a nota fiscal emitida pelos modulos de faturamento/estoque
'retorna em lNumIntDocCPR o numintdoc do registro em TitulosPag

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar, objTitPagAux As New ClassTituloPagar
Dim colParcelas As New ColParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim objParcPag As ClassParcelaPagar, dValorPagar As Double

On Error GoTo Erro_TituloPagar_Gera

    objTituloPagar.dOutrasDespesas = objNFiscal.dValorOutrasDespesas
    objTituloPagar.dtDataEmissao = objNFiscal.dtDataEmissao
    objTituloPagar.dValorFrete = objNFiscal.dValorFrete
    objTituloPagar.dValorProdutos = objNFiscal.dValorProdutos
    objTituloPagar.dValorSeguro = objNFiscal.dValorSeguro
    objTituloPagar.dValorTotal = objNFiscal.dValorTotal
    objTituloPagar.iFilial = objNFiscal.iFilialForn
    objTituloPagar.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objTituloPagar.iNumParcelas = objNFiscal.ColParcelaPagar.Count
    objTituloPagar.iStatus = STATUS_LANCADO
    objTituloPagar.lFornecedor = objNFiscal.lFornecedor
    objTituloPagar.lNumTitulo = objNFiscal.lNumNotaFiscal
    objTituloPagar.sSiglaDocumento = sTipoDocCPR
    objTituloPagar.dtDataRegistro = gdtDataHoje
    objTituloPagar.iCondicaoPagto = objNFiscal.iCondicaoPagto
    objTituloPagar.sNatureza = objNFiscal.sNatureza
    
    If Not (objNFiscal.objInfoAdic Is Nothing) Then
        objTituloPagar.sCcl = objNFiscal.objInfoAdic.sCcl
    Else
        objTituloPagar.sCcl = ""
    End If
    
    objTituloPagar.dValorICMS = objNFiscal.objTributacaoNF.dICMSValor
    objTituloPagar.dValorICMSSubst = objNFiscal.objTributacaoNF.dICMSSubstValor
    objTituloPagar.dValorIPI = objNFiscal.objTributacaoNF.dIPIValor
    objTituloPagar.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
    objTituloPagar.iCreditoICMS = IIf(objNFiscal.objTributacaoNF.dICMSCredito > 0.01, 1, 0)
    objTituloPagar.iCreditoIPI = IIf(objNFiscal.objTributacaoNF.dIPICredito > 0.01, 1, 0)
    objTituloPagar.dValorINSS = objNFiscal.objTributacaoNF.dValorINSS
    objTituloPagar.iINSSRetido = objNFiscal.objTributacaoNF.iINSSRetido
    objTituloPagar.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
    objTituloPagar.dISSRetido = objNFiscal.objTributacaoNF.dISSRetido
    objTituloPagar.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
    objTituloPagar.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
    
    objTituloPagar.dTaxaMoeda = objNFiscal.dTaxaMoeda
    
'    objTituloPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
    
    For Each objParcelaPagar In objNFiscal.ColParcelaPagar
        With objParcelaPagar
            Set objParcPag = colParcelas.Add(.lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras, , .dValor, 0)
            Set objParcPag.objInfoUsu = objParcelaPagar.objInfoUsu
            dValorPagar = Arredonda_Moeda(dValorPagar + .dValor, 2)
        End With
    Next
    
    If objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCNT Or objNFiscal.iTipoNFiscal = DOCINFO_NFEEFCNT3 Or objNFiscal.iTipoNFiscal = DOCINFO_NFEFBFRS Then
        objTituloPagar.dValorTotal = dValorPagar + objTituloPagar.dCOFINSRetido + objTituloPagar.dCSLLRetido + objTituloPagar.dISSRetido + objTituloPagar.dPISRetido + objTituloPagar.dValorINSS + objTituloPagar.dValorIRRF
        objTituloPagar.dValorProdutos = dValorPagar - (objNFiscal.dValorTotal - objNFiscal.dValorProdutos)
    End If
    
    objTituloPagar.lNumIntDoc = 0
    
    'se está alterando uma nf
    If objNFiscal.lNumIntDocCPR < 0 Then
    
        objTitPagAux.lNumIntDoc = -objNFiscal.lNumIntDocCPR
        lErro = CF("TituloPagarBaixado_Le", objTitPagAux)
        If lErro <> SUCESSO And lErro <> 56661 Then gError 180336
    
        If lErro = SUCESSO Then
        
            If Abs(objTitPagAux.dValorTotal - objTituloPagar.dValorTotal) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dValorIRRF - objTituloPagar.dValorIRRF) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dPISRetido - objTituloPagar.dPISRetido) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dISSRetido - objTituloPagar.dISSRetido) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dCOFINSRetido - objTituloPagar.dCOFINSRetido) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dCSLLRetido - objTituloPagar.dCSLLRetido) > DELTA_VALORMONETARIO Or _
                Abs(objTitPagAux.dValorINSS - objTituloPagar.dValorINSS) > DELTA_VALORMONETARIO Or objTitPagAux.iINSSRetido <> objTituloPagar.iINSSRetido Then gError 180336
            
            objNFiscal.lNumIntDocCPR = -objNFiscal.lNumIntDocCPR
            objTituloPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
        
        Else
        
            'Desfaz os documentos associados
            lErro = CF("CPR_Exclui", objNFiscal.iClasseDocCPR, -objNFiscal.lNumIntDocCPR, 1)
            If lErro <> SUCESSO Then gError 46197
        
        End If
        
    End If
        
    If objTituloPagar.lNumIntDoc = 0 Then
    
        lErro = CF("NFFatPag_Grava_BD", objTituloPagar, colParcelas, False)
        If lErro <> SUCESSO And lErro <> 180335 Then gError 46197
        If lErro = 180335 Then gError 180336
    
    End If
    
    lNumIntDocCPR = objTituloPagar.lNumIntDoc
    
    TituloPagar_Gera = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Gera:

    TituloPagar_Gera = gErr
    
    Select Case gErr
    
        Case 46197
        
        Case 180336
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULOPAGAR_INSERCAO_MESMO_NUMERO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148913)
            
    End Select
    
    Exit Function

End Function

Function NFPagar_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long
'registra nota fiscal a pagar na tabela NFPag de ctas a pagar vinculada a nota fiscal emitida pelos modulos de faturamento/estoque.
'Retorna em lNumIntDocCPR o numintdoc da nf gerada.

Dim lErro As Long
Dim objNFPag As New ClassNFsPag, objNFPagAux As New ClassNFsPag

On Error GoTo Erro_NFPagar_Gera
    
    objNFPag.dOutrasDespesas = objNFiscal.dValorOutrasDespesas
    objNFPag.dtDataEmissao = objNFiscal.dtDataEmissao
    objNFPag.dtDataVencimento = objNFiscal.dtDataVencimento
    objNFPag.dValorFrete = objNFiscal.dValorFrete
    objNFPag.dValorProdutos = objNFiscal.dValorProdutos
    objNFPag.dValorSeguro = objNFiscal.dValorSeguro
    objNFPag.dValorTotal = objNFiscal.dValorTotal
    objNFPag.iFilial = objNFiscal.iFilialForn
    objNFPag.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objNFPag.iStatus = objNFiscal.iStatus
    objNFPag.lFornecedor = objNFiscal.lFornecedor
    objNFPag.lNumNotaFiscal = objNFiscal.lNumNotaFiscal
    
    If objNFiscal.iTipoDocInfo = TIPODOCINFO_CONHECIMENTOFRETE Then
        objNFPag.dValorICMS = objNFiscal.objConhecimentoFrete.dValorICMS
    Else
        objNFPag.dValorICMS = objNFiscal.objTributacaoNF.dICMSValor
        objNFPag.dValorICMSSubst = objNFiscal.objTributacaoNF.dICMSSubstValor
        objNFPag.dValorIPI = objNFiscal.objTributacaoNF.dIPIValor
        objNFPag.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
        objNFPag.iCreditoICMS = IIf(objNFiscal.objTributacaoNF.dICMSCredito > 0.01, 1, 0)
        objNFPag.iCreditoIPI = IIf(objNFiscal.objTributacaoNF.dIPICredito > 0.01, 1, 0)
        objNFPag.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
        objNFPag.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
        objNFPag.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
        objNFPag.dTaxaMoeda = objNFiscal.dTaxaMoeda
        objNFPag.dISSRetido = objNFiscal.objTributacaoNF.dISSRetido
        objNFPag.dValorINSS = objNFiscal.objTributacaoNF.dValorINSS
        objNFPag.iINSSRetido = objNFiscal.objTributacaoNF.iINSSRetido
        
    End If
    
    objNFPag.lNumIntDoc = 0
    
    'se está alterando uma nf
    If objNFiscal.lNumIntDocCPR < 0 Then
    
        objNFPagAux.lNumIntDoc = -objNFiscal.lNumIntDocCPR
        lErro = CF("NFPagBaixada_Le", objNFPagAux)
        If lErro <> SUCESSO And lErro <> 18257 Then Error 46198
        
        If lErro = SUCESSO Then
        
            If Abs(objNFPagAux.dValorTotal - objNFPag.dValorTotal) > DELTA_VALORMONETARIO Or _
                Abs(objNFPagAux.dValorIRRF - objNFPag.dValorIRRF) > DELTA_VALORMONETARIO Or _
                Abs(objNFPagAux.dPISRetido - objNFPag.dPISRetido) > DELTA_VALORMONETARIO Or _
                Abs(objNFPagAux.dISSRetido - objNFPag.dISSRetido) > DELTA_VALORMONETARIO Or _
                Abs(objNFPagAux.dValorINSS - objNFPag.dValorINSS) > DELTA_VALORMONETARIO Or _
                Abs(objNFPagAux.dCOFINSRetido - objNFPag.dCOFINSRetido) > DELTA_VALORMONETARIO Or objNFPagAux.iINSSRetido <> objNFPag.iINSSRetido Or _
                Abs(objNFPagAux.dCSLLRetido - objNFPag.dCSLLRetido) > DELTA_VALORMONETARIO Then Error 46197
            
            objNFiscal.lNumIntDocCPR = -objNFiscal.lNumIntDocCPR
            objNFPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
        
        Else
        
            'Desfaz os documentos associados
            lErro = CF("CPR_Exclui", objNFiscal.iClasseDocCPR, -objNFiscal.lNumIntDocCPR, 1)
            If lErro <> SUCESSO Then Error 46198
        
        End If
    
    End If
    
    If objNFPag.lNumIntDoc = 0 Then
    
        lErro = CF("NFPag_Grava_BD", objNFPag, NFPAG_NFINSERCAO)
        If lErro <> SUCESSO Then Error 46198
    
    End If
    
    lNumIntDocCPR = objNFPag.lNumIntDoc
    
    NFPagar_Gera = SUCESSO
    
    Exit Function
    
Erro_NFPagar_Gera:

    NFPagar_Gera = Err
    
    Select Case Err
    
        Case 46198
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148914)
            
    End Select
    
    Exit Function

End Function

Function TituloReceber_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long
'registra titulo a receber vinculado a nota fiscal emitida pelos modulos de faturamento/estoque

Dim lErro As Long, dValorLiquido As Double
Dim objTituloReceber As New ClassTituloReceber
Dim colParcelas As New colParcelaReceber
Dim objParcelaReceber As ClassParcelaReceber
Dim colComissaoEmissao As New colComissao
Dim colComissao As colComissao, colcolComissao As New colcolComissao
Dim colChequePre As New colChequePre
Dim dtDataVencimento As Date
Dim dtDataVencimentoReal As Date
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_TituloReceber_Gera

    lErro = CF("TituloReceber_Gera_Customizado2", objNFiscal)
    If lErro <> SUCESSO Then gError 198615

    objTituloReceber.dtDataEmissao = objNFiscal.dtDataEmissao
    objTituloReceber.dValor = objNFiscal.dValorTotal - objNFiscal.dValorDescontoTit
    objTituloReceber.iFilial = objNFiscal.iFilialCli
    objTituloReceber.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objTituloReceber.iNumParcelas = objNFiscal.colParcelaReceber.Count
    objTituloReceber.iStatus = STATUS_LANCADO
    objTituloReceber.lCliente = objNFiscal.lCliente
    objTituloReceber.lNumTitulo = objNFiscal.lNumNotaFiscal
    objTituloReceber.sSiglaDocumento = sTipoDocCPR
    objTituloReceber.iCondicaoPagto = objNFiscal.iCondicaoPagto
    objTituloReceber.dtDataRegistro = gdtDataHoje
    objTituloReceber.dPercJurosDiario = gobjCRFAT.dPercJurosDiario
    objTituloReceber.dPercMulta = gobjCRFAT.dPercMulta
    objTituloReceber.iMoeda = MOEDA_REAL
    If Not (objNFiscal.objInfoAdic Is Nothing) Then
        objTituloReceber.sCcl = objNFiscal.objInfoAdic.sCcl
    Else
        objTituloReceber.sCcl = ""
    End If
    
    objTituloReceber.sNatureza = objNFiscal.sNatureza
    objTituloReceber.lCpoGenerico1 = 0
    objTituloReceber.iReajustePeriodicidade = 0
    objTituloReceber.dtReajusteBase = DATA_NULA
    objTituloReceber.dtReajustadoAte = DATA_NULA
    
    'A nota de conhecimento de frete não tem tributação
    If (objNFiscal.iTipoNFiscal <> TIPODOCINFO_CONHECIMENTOFRETE_FATURA And objNFiscal.iTipoNFiscal <> TIPODOCINFO_CONHECIMENTOFRETE_FATURA_COMPLICMS) Then
        objTituloReceber.dValorISS = objNFiscal.objTributacaoNF.dISSValor
        objTituloReceber.dValorINSS = objNFiscal.objTributacaoNF.dValorINSS
        objTituloReceber.iINSSRetido = objNFiscal.objTributacaoNF.iINSSRetido
        objTituloReceber.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
        objTituloReceber.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
        objTituloReceber.dISSRetido = objNFiscal.objTributacaoNF.dISSRetido
        objTituloReceber.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
        objTituloReceber.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
    Else
        objTituloReceber.dValorISS = 0
        objTituloReceber.dValorINSS = objNFiscal.objConhecimentoFrete.dValorINSS
        objTituloReceber.iINSSRetido = objNFiscal.objConhecimentoFrete.iINSSRetido
        objTituloReceber.dValorIRRF = 0
        objTituloReceber.dPISRetido = 0
        objTituloReceber.dISSRetido = 0
        objTituloReceber.dCOFINSRetido = 0
        objTituloReceber.dCSLLRetido = 0
    End If
    
    For Each objParcelaReceber In objNFiscal.colParcelaReceber
        
        dtDataVencimento = objParcelaReceber.dtDataVencimento
        
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVencimentoReal)
        If lErro <> SUCESSO Then gError 94102
        
        objParcelaReceber.dtDataVencimentoReal = dtDataVencimentoReal
        
        With objParcelaReceber
        '######################################################
        'ALTERADO POR WAGNER
            Set objParcRec = colParcelas.Add(.lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iProxSeqBaixa, .iCarteiraCobranca, .iCobrador, .sNumTitCobrador, 0, 0, 0, 0, 0, 0, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .lNumIntCheque, .iAceite, .iDescontada, .iProxSeqOcorr, .iPrevisao, .sObservacao, .dValor)
            
            objParcRec.iTipoPagto = .iTipoPagto
            objParcRec.iCodConta = .iCodConta
            objParcRec.iBancoCheque = .iBancoCheque
            objParcRec.sAgenciaCheque = .sAgenciaCheque
            objParcRec.sContaCorrenteCheque = .sContaCorrenteCheque
            objParcRec.lNumeroCheque = .lNumeroCheque
            objParcRec.dtDataDepositoCheque = .dtDataDepositoCheque
            objParcRec.dtDataEmissaoCheque = .dtDataEmissaoCheque
            objParcRec.iAdmMeioPagto = .iAdmMeioPagto
            objParcRec.iParcelamento = .iParcelamento
            objParcRec.sNumeroCartao = .sNumeroCartao
            objParcRec.dtValidadeCartao = .dtValidadeCartao
            objParcRec.sAprovacaoCartao = .sAprovacaoCartao
            objParcRec.dtDataCredito = .dtDataCredito
            objParcRec.dtDataTransacaoCartao = .dtDataTransacaoCartao
            
            colChequePre.Add 0, 0, 0, 0, "", "", 0, DATA_NULA, 0, 0
        '######################################################
                
            dValorLiquido = dValorLiquido + .dValor
        
        End With
        
        Set colComissao = New colComissao
        colcolComissao.Add colComissao
        
    Next
    
    objTituloReceber.dSaldo = dValorLiquido
    
    'If objNFiscal.iTipoNFiscal = DOCINFO_NFFBFRS Then
        objTituloReceber.dValor = dValorLiquido + objTituloReceber.dCOFINSRetido + objTituloReceber.dCSLLRetido + objTituloReceber.dISSRetido + objTituloReceber.dPISRetido + objTituloReceber.dValorINSS + objTituloReceber.dValorIRRF
    'End If
    
    lErro = CF("ComissoesNF_Gera_ComissoesParc", objNFiscal, colcolComissao, objTituloReceber.dSaldo)
    If lErro <> SUCESSO Then gError 56857
    
    lErro = CF("TituloReceber_Insere", objTituloReceber, colParcelas, colComissaoEmissao, colcolComissao)
    If lErro <> SUCESSO Then gError 46199
    
    'mario
    lErro = CF("TituloReceber_TipoPagto", objTituloReceber, colParcelas)
    If lErro <> SUCESSO Then gError 183016
    
    lErro = CF("TituloReceber_Gera_Customizado", objNFiscal, objTituloReceber, colParcelas)
    If lErro <> SUCESSO Then gError 198615
    
    lNumIntDocCPR = objTituloReceber.lNumIntDoc
    
    TituloReceber_Gera = SUCESSO
    
    Exit Function

Erro_TituloReceber_Gera:

    TituloReceber_Gera = gErr
    
    Select Case gErr
    
        Case 46199, 56857, 94102, 183016, 198615
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148915)
            
    End Select
    
    Exit Function
    
 End Function

Function CarteiraCobrador_AtualizaSaldo(ByVal objCarteiraCobrador As ClassCarteiraCobrador, ByVal iQtdeParcelas As Integer, ByVal dValorParcelas As Double, Optional ByVal lComandoP1 As Long = 0, Optional ByVal lComandoP2 As Long = 0) As Long
'atualiza a qtde e valor de titulos na carteira de cobranca de um cobrador
'iQtdeParcelas e dValorParcelas devem ser positivos para entradas e negativos para saidas da carteira
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, lComandoSel As Long, lComandoMod As Long
Dim lQuantidadeAtual As Long, dSaldoAtual As Double, dNovoSaldo As Double

On Error GoTo Erro_CarteiraCobrador_AtualizaSaldo

    If lComandoP1 <> 0 Then
        lComandoSel = lComandoP1
    Else
        lComandoSel = Comando_Abrir()
        If lComandoSel = 0 Then Error 19441
    End If
    
    If lComandoP2 <> 0 Then
        lComandoMod = lComandoP2
    Else
        lComandoMod = Comando_Abrir()
        If lComandoMod = 0 Then Error 19442
    End If
    
    lErro = Comando_ExecutarPos(lComandoSel, "SELECT QuantidadeAtual, SaldoAtual FROM CarteirasCobrador WHERE Cobrador = ? AND CodCarteiraCobranca = ?", 0, lQuantidadeAtual, dSaldoAtual, objCarteiraCobrador.iCobrador, objCarteiraCobrador.iCodCarteiraCobranca)
    If lErro <> AD_SQL_SUCESSO Then Error 19443
    
    lErro = Comando_BuscarPrimeiro(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19444
    
    If lErro = AD_SQL_SEM_DADOS Then Error 19445
    
    'verificar se algum valor ficará menor que zero
'    If lQuantidadeAtual + iQtdeParcelas < 0 Then Error 19446
    dNovoSaldo = Format(CStr(dSaldoAtual + dValorParcelas), "0.00")
'    If dNovoSaldo < 0 Then Error 19447
    
    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE CarteirasCobrador SET QuantidadeAtual = QuantidadeAtual + ?, SaldoAtual = ?", lComandoSel, iQtdeParcelas, dNovoSaldo)
    If lErro <> AD_SQL_SUCESSO Then Error 19448
    
    If lComandoP1 = 0 Then Call Comando_Fechar(lComandoSel)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComandoMod)
    
    CarteiraCobrador_AtualizaSaldo = SUCESSO
     
    Exit Function
    
Erro_CarteiraCobrador_AtualizaSaldo:

    CarteiraCobrador_AtualizaSaldo = Err
     
    Select Case Err
          
        Case 19441, 19442
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 19443, 19444
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CARTEIRA_COBRADOR", Err, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)
        
        Case 19445
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRADOR_INEXISTENTE", Err)
        
        Case 19446
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRADOR_QTDE_NEG", Err, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)
        
        Case 19447
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRA_COBRADOR_SALDO_NEG", Err, objCarteiraCobrador.iCodCarteiraCobranca, objCarteiraCobrador.iCobrador)
        
        Case 19448
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CARTEIRAS_COBRADOR", Err)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148916)
     
    End Select
     
    If lComandoP1 = 0 Then Call Comando_Fechar(lComandoSel)
    If lComandoP2 = 0 Then Call Comando_Fechar(lComandoMod)
    
    Exit Function

End Function

Function TransfCartCobr_Grava1(ByVal objTransfCartCobr As ClassTransfCartCobr, ByVal iCobradorAnterior As Integer, ByVal iCarteiraAnterior As Integer, ByVal dValor As Double) As Long
'registra a transferencia de uma parcela de uma carteira de cobranca para outra
'TEM QUE SER CHAMADO DENTRO DE TRANSACAO
Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador
Dim lComando As Long, lNumIntDoc As Long

On Error GoTo Erro_TransfCartCobr_Grava1

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19450
    
    'retira da estatistica da carteira anterior
    objCarteiraCobrador.iCobrador = iCobradorAnterior
    objCarteiraCobrador.iCodCarteiraCobranca = iCarteiraAnterior
    lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, -1, -dValor)
    If lErro <> SUCESSO Then Error 19451
    
    'inclui na estatistica da carteira nova
    objCarteiraCobrador.iCobrador = objTransfCartCobr.iCobrador
    objCarteiraCobrador.iCodCarteiraCobranca = objTransfCartCobr.iCarteiraCobranca
    lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, 1, dValor)
    If lErro <> SUCESSO Then Error 19452
    
    If objTransfCartCobr.lNumIntDoc = 0 Then
    
        lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_PROX_TRANSFCARTCOBR", lNumIntDoc)
        If lErro <> SUCESSO Then Error 32211
        
        objTransfCartCobr.lNumIntDoc = lNumIntDoc
        
    End If
    
    'incluir registro na tabela de transferencias de carteira de cobranca
    lErro = Comando_Executar(lComando, "INSERT INTO TransfCartCobr (NumIntDoc, NumIntParc, NumSeqOcorr, Cobrador, CarteiraCobranca, Data, DataRegistro) VALUES (?,?,?,?,?,?,?)", objTransfCartCobr.lNumIntDoc, objTransfCartCobr.lNumIntParc, objTransfCartCobr.iNumSeqOcorr, objTransfCartCobr.iCobrador, objTransfCartCobr.iCarteiraCobranca, objTransfCartCobr.dtData, objTransfCartCobr.dtDataRegistro)
    If lErro <> AD_SQL_SUCESSO Then Error 19453
    
    Call Comando_Fechar(lComando)
    
    TransfCartCobr_Grava1 = SUCESSO
     
    Exit Function
    
Erro_TransfCartCobr_Grava1:

    TransfCartCobr_Grava1 = Err
     
    Select Case Err
          
        Case 19450
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 19451, 19452, 32211
        
        Case 19453
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRANSFCARTCOBR", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148917)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'mario
Function Comissoes_Gera(ByVal lComando As Long, ByVal lComando1 As Long, ByVal objNFiscal As ClassNFiscal, ByVal colComissoesEmissao As colComissao) As Long
'Gera as Comissões na Emissão com base nas comissões do grid comissão em coloca em colComissoesEmissao

Dim lErro As Long
Dim objComissaoNF As ClassComissaoNF
Dim iIndice As Integer
Dim dPercentual As Double
Dim lNumIntCom As Long

On Error GoTo Erro_Comissoes_Gera

    'Para cada comissão
    For Each objComissaoNF In objNFiscal.ColComissoesNF
    
        'Calcula o percentual de Emissão
        dPercentual = objComissaoNF.dPercentual * objComissaoNF.dPercentualEmissao
        
        'Guarda em colComissoesEmissao
        colComissoesEmissao.Add lNumIntCom, giFilialEmpresa, STATUS_LIBERADO, TIPO_NF, objNFiscal.lNumIntDoc, objComissaoNF.iCodVendedor, DATA_NULA, dPercentual, objComissaoNF.dValorBase, objComissaoNF.dValorEmissao, objNFiscal.dtDataEmissao

    Next

    Comissoes_Gera = SUCESSO
    
    Exit Function

Erro_Comissoes_Gera:

    Comissoes_Gera = Err
    
    Select Case Err
    
        Case 42444
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148918)
            
    End Select
    
    Exit Function

End Function

'mario
Function Comissao_Automatico1(ByVal lComando As Long, ByVal lComando1 As Long, lNumIntDoc As Long) As Long
'Gera o próximo Número Interno de uma Comissao
'AVISO: ESSA ROTINA DEVE SER CHAMADA DE DENTRO DE UMA TRANSACAO

Dim lErro As Long
Dim sConteudo As String

On Error GoTo Erro_Comissao_Automatico1

    sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)

    'Pesquisa qual o último numero usado
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM CPRConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, NUM_PROX_COMISSAO, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then Error 44270

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44271

    'Faz "lock"
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44272

    'Guarda o numero
    lNumIntDoc = CLng(sConteudo)

    'Incrementa o valor
    sConteudo = CStr(lNumIntDoc + 1)

    'Atualiza a tabela
    lErro = Comando_ExecutarPos(lComando1, "UPDATE CPRConfig SET Conteudo = ?", lComando, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 44273

    Comissao_Automatico1 = SUCESSO

    Exit Function

Erro_Comissao_Automatico1:

    Comissao_Automatico1 = Err

    Select Case Err

        Case 44270, 44271
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, NUM_PROX_COMISSAO)

        Case 44272
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CPRCONFIG", Err, NUM_PROX_COMISSAO)

        Case 44273
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", Err, NUM_PROX_COMISSAO)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148919)

    End Select

    Exit Function

End Function

'mario
Function Comissoes_Grava1(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal lComando3 As Long, ByVal colComissao As colComissao) As Long
'Grava as Comissões passadas em colComissao

Dim lErro As Long
Dim objComissao As ClassComissao
Dim iCodigo As Integer
Dim lNumIntCom As Long

On Error GoTo Erro_Comissoes_Grava1

    'Loop de gravação
    For Each objComissao In colComissao

        'Gera o próximo número interno para a nova Comissão
        lErro = CF("Comissao_Automatico1", lComando2, lComando3, lNumIntCom)
        If lErro <> SUCESSO Then Error 44278

        'Insere a comissao no BD
        lErro = Comando_Executar(lComando1, "INSERT INTO Comissoes (NumIntCom, FilialEmpresa, Status, TipoTitulo, NumIntDoc, CodVendedor, DataGeracao, DataBaixa, Percentual, ValorBase, Valor) VALUES (?,?,?,?,?,?,?,?,?,?,?)", lNumIntCom, giFilialEmpresa, objComissao.iStatus, objComissao.iTipoTitulo, objComissao.lNumIntDoc, objComissao.iCodVendedor, objComissao.dtDataGeracao, objComissao.dtDataBaixa, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor)
        If lErro <> AD_SQL_SUCESSO Then Error 44279

    Next

    Comissoes_Grava1 = SUCESSO

    Exit Function

Erro_Comissoes_Grava1:

    Comissoes_Grava1 = Err

    Select Case Err

        Case 44278, 59074

        Case 44279
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOES", Err, objComissao.iTipoTitulo, objComissao.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148920)

    End Select

    Exit Function

End Function

'mario
Function NFiscal_Automatico1(ByVal lComando As Long, ByVal lComando1 As Long, ByVal lComando2 As Long, lNumIntNF As Long) As Long
'Retorna o número do próximo NumIntNF disponivel

Dim lErro As Long
Dim lCodigo_NumIntNF As Long
Dim lcodigo_automatico As Long
Dim sConteudo As String
Dim sConteudo2 As String

On Error GoTo Erro_NFiscal_Automatico1

    sConteudo = String(STRING_CONTEUDO, 0)

    'Pesquisa o número do NumIntNF gerado automaticamente pelo sistema
    lErro = Comando_ExecutarPos(lComando, "SELECT Conteudo FROM CRFatConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, NUM_PROX_NOTA_FISCAL, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then Error 44263

    'le o registro com o número do NumIntNF automático
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44264

    'loca o registro que guarda o numero do NumIntNF gerado automaticamente
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 44265

    lcodigo_automatico = CLng(sConteudo)

    'Valor de retorno da funcao
    lNumIntNF = lcodigo_automatico

    lcodigo_automatico = lcodigo_automatico + 1

    sConteudo2 = CStr(lcodigo_automatico)

    'atualiza o numero do NumIntNF gerado automaticamente
    lErro = Comando_ExecutarPos(lComando2, "UPDATE CRFatConfig SET Conteudo = ?", lComando, sConteudo2)
    If lErro <> AD_SQL_SUCESSO Then Error 44269

    NFiscal_Automatico1 = SUCESSO

    Exit Function

Erro_NFiscal_Automatico1:

    NFiscal_Automatico1 = Err

    Select Case Err

    Case 44263, 44264
        Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRFATCONFIG", Err, NUM_PROX_NOTA_FISCAL)

    Case 44265
        Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CRFATCONFIG", Err, NUM_PROX_NOTA_FISCAL)

    Case 44266, 44267, 44268
        Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)

    Case 44269
        Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CRFATCONFIG", Err, NUM_PROX_NOTA_FISCAL)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148921)

    End Select

    Exit Function

End Function

'mario
Function ComissoesNF_Grava(ByVal lComando As Long, ByVal lComando1 As Long, ByVal objNFiscal As ClassNFiscal) As Long
'Insere no Banco de Dados as Comissões da Nota Fiscal passada em objNFiscal
'ATENÇÃO: Essa função deve ser chamada dentro de Transação.

'Alterado por Tulio, inclusao do campo direto/indireto

Dim lErro As Long
Dim iCodVendedor As Integer
Dim objComissaoNF As ClassComissaoNF
Dim iCodigo As Integer
Dim objComissaoPorItem As ClassComissaoPorItem

On Error GoTo Erro_ComissoesNF_Grava

    For Each objComissaoNF In objNFiscal.ColComissoesNF

        'Insere a Comissão no BD
        lErro = Comando_Executar(lComando1, "INSERT INTO ComissoesNF (NumIntDoc,CodVendedor,ValorBase,Percentual,Valor,PercentualEmissao,ValorEmissao,Indireta,Seq) VALUES (?,?,?,?,?,?,?,?,?)", objNFiscal.lNumIntDoc, objComissaoNF.iCodVendedor, objComissaoNF.dValorBase, objComissaoNF.dPercentual, objComissaoNF.dValor, objComissaoNF.dPercentualEmissao, objComissaoNF.dValorEmissao, objComissaoNF.iIndireta, objComissaoNF.iSeq)
        If lErro <> AD_SQL_SUCESSO Then Error 39017

        For Each objComissaoPorItem In objComissaoNF.colLinhas
        
            objComissaoPorItem.iTipoDoc = TIPO_NF
            objComissaoPorItem.iSeqComissao = objComissaoNF.iSeq
            
            If objComissaoPorItem.iLinha >= 1 And objNFiscal.colItens.Count >= objComissaoPorItem.iLinha Then objComissaoPorItem.lNumIntDocItem = objNFiscal.colItens.Item(objComissaoPorItem.iLinha).lNumIntDoc
            
            With objComissaoPorItem
                lErro = Comando_Executar(lComando, "INSERT INTO ComissoesPorItem (TipoDoc, NumIntDocItem, SeqComissao, ValorBase, Percentual, Valor, Linha, PercentualEmissao, ValorEmissao) VALUES (?,?,?,?,?,?,?,?,?)", _
                    .iTipoDoc, .lNumIntDocItem, .iSeqComissao, .dValorBase, .dPercentual, .dValor, .iLinha, .dPercentualEmissao, .dValorEmissao)
            End With
            If lErro <> AD_SQL_SUCESSO Then Error 39017
        
        Next
    
    Next

    ComissoesNF_Grava = SUCESSO

    Exit Function

Erro_ComissoesNF_Grava:

    ComissoesNF_Grava = Err

    Select Case Err

        Case 59072
        
        Case 39017
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESNF", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148922)

    End Select

    Exit Function

End Function

Function TiposDeCliente_Grava(ByVal objTipoCliente As ClassTipoCliente) As Long
'Inclui ou altera um Tipo de Cliente
'Tabelas:TiposDeCliente e TipoDeClienteCategoria

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TiposDeCliente_Grava

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 28946

    lErro = CF("TiposDeCliente_Grava_EmTrans", objTipoCliente)
    If lErro <> SUCESSO Then gError 130713
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 28961

    TiposDeCliente_Grava = SUCESSO

    Exit Function

Erro_TiposDeCliente_Grava:

    TiposDeCliente_Grava = gErr

    Select Case gErr

        Case 130713
        
        Case 28945
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 28961
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148923)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function TiposDeCliente_Grava_EmTrans(ByVal objTipoCliente As ClassTipoCliente) As Long
'Inclui ou altera um Tipo de Cliente
'Tabelas:TiposDeCliente e TipoDeClienteCategoria

Dim lErro As Long
Dim iIndice As Integer
Dim colCategoriaItemCopia As New Collection
Dim tTipoCliente As typeTipoCliente
Dim sCategoria As String, sItem As String
Dim objTipoClienteCategoria As ClassTipoClienteCategoria
Dim alComando(1 To 8) As Long
Dim iCodigo As Integer

On Error GoTo Erro_TiposDeCliente_Grava_EmTrans

    'Cria uma cópia "de trabalho" da coleção passada como parametro
    For Each objTipoClienteCategoria In objTipoCliente.colCategoriaItem
        colCategoriaItemCopia.Add objTipoClienteCategoria
    Next

    'Inicializa comandos
    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 28945

    Next

    tTipoCliente.sDescricao = String(STRING_TIPO_CLIENTE_DESCRICAO, 0)

    'Faz "locks" nas tabelas dos objetos relacionados ao Tipo de Cliente
    lErro = CF("TipoCliente_Le_Tabelas_Lock", objTipoCliente)
    If lErro <> SUCESSO Then Error 28947

    'Procura por uma outra Descricão já cadastrada
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM TiposDeCliente WHERE Codigo <> ? AND Descricao = ?", iCodigo, objTipoCliente.iCodigo, objTipoCliente.sDescricao)
    If lErro <> AD_SQL_SUCESSO Then Error 28948

    'Verifica se existe descrição duplicada
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28949

    'Se existir outra Descricao ==> erro
    If lErro <> AD_SQL_SEM_DADOS Then Error 28950

    'Pesquisa no BD o Tipo de Cliente em questão
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Descricao FROM TiposDeCliente WHERE Codigo = ?", 0, tTipoCliente.sDescricao, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28951

    'Lê o Tipo de Cliente, se estiver no BD
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28952

    'Se o Tipo de Cliente já existe - MODIFICAÇÃO
    If lErro = AD_SQL_SUCESSO Then

        'Alterar o Tipo de Cliente
        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TiposDeCliente SET Descricao = ?, LimiteCredito = ?, CondicaoPagto = ?, Desconto = ?, CodMensagem = ?, TabelaPreco = ?, Observacao = ?, ContaContabil = ?, Vendedor = ?, ComissaoVendas = ?, Regiao = ?, FreqVisitas = ?, CodTransportadora = ?, CodCobrador = ?, PadraoCobranca = ?", _
            alComando(2), objTipoCliente.sDescricao, objTipoCliente.dLimiteCredito, objTipoCliente.iCondicaoPagto, objTipoCliente.dDesconto, objTipoCliente.iCodMensagem, objTipoCliente.iTabelaPreco, objTipoCliente.sObservacao, objTipoCliente.sContaContabil, objTipoCliente.iVendedor, objTipoCliente.dComissaoVendas, objTipoCliente.iRegiao, objTipoCliente.iFreqVisitas, objTipoCliente.iCodTransportadora, objTipoCliente.iCodCobrador, objTipoCliente.iPadraoCobranca)
        If lErro <> AD_SQL_SUCESSO Then Error 28953

        sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
        sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

        'Percorre as Categorias do Tipo de Cliente no BD
        lErro = Comando_ExecutarPos(alComando(4), "SELECT Categoria, Item FROM TipoDeClienteCategorias WHERE TipoDeCliente = ?", 0, sCategoria, sItem, objTipoCliente.iCodigo)
        If lErro <> AD_SQL_SUCESSO Then Error 28954

        'Lê a Categoria do Tipo de Cliente, se estiver no BD
        lErro = Comando_BuscarProximo(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28955

        Do While lErro <> AD_SQL_SEM_DADOS

            'Procura a Categoria na coleção correspondente ao grid
            iIndice = TipoCliente_Grava_CategoriaNaColecao(colCategoriaItemCopia, sCategoria)

            If iIndice <> 0 Then

                Set objTipoClienteCategoria = colCategoriaItemCopia.Item(iIndice)

                'Se houver sido alterado o Item
                If objTipoClienteCategoria.sItem <> sItem Then

                    'verifica se existe e bloqueia
                    lErro = CF("CategoriaClienteItem_Lock", objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem, alComando(8))
                    If lErro <> AD_SQL_SUCESSO And lErro <> 40555 Then Error 41626
                    If lErro = 40555 Then Error 41630
                                        
                    'Altera a Tabela TipoDeClienteCategorias
                    lErro = Comando_ExecutarPos(alComando(5), "UPDATE TipoDeClienteCategorias SET Categoria = ?, Item = ?", alComando(4), objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem)
                    If lErro <> AD_SQL_SUCESSO Then Error 28956

                End If

                'Retira da coleção
                colCategoriaItemCopia.Remove (iIndice)

            Else

                'Excluir registro em TipoDeClienteCategorias
                lErro = Comando_ExecutarPos(alComando(6), "DELETE FROM TipoDeClienteCategorias", alComando(4))
                If lErro <> AD_SQL_SUCESSO Then Error 28957

            End If

            'Lê a Categoria do Tipo de Cliente, se estiver no BD
            lErro = Comando_BuscarProximo(alComando(4))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28958

        Loop

    Else 'O Tipo de Cliente não existe

        'Insere em TiposDeCliente, criando um novo Tipo
        lErro = Comando_Executar(alComando(7), "INSERT INTO TiposDeCliente (Codigo, Descricao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, CodTransportadora, CodCobrador, PadraoCobranca) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objTipoCliente.iCodigo, objTipoCliente.sDescricao, objTipoCliente.dLimiteCredito, objTipoCliente.iCondicaoPagto, objTipoCliente.dDesconto, objTipoCliente.iCodMensagem, objTipoCliente.iTabelaPreco, objTipoCliente.sObservacao, objTipoCliente.sContaContabil, objTipoCliente.iVendedor, objTipoCliente.dComissaoVendas, objTipoCliente.iRegiao, objTipoCliente.iFreqVisitas, objTipoCliente.iCodTransportadora, objTipoCliente.iCodCobrador, objTipoCliente.iPadraoCobranca)
        If lErro <> AD_SQL_SUCESSO Then Error 28959

    End If

    'Grava as Categorias que ainda não faziam parte do Tipo de Cliente
    lErro = TipoCliente_Grava_NovasCategorias(objTipoCliente, colCategoriaItemCopia)
    If lErro <> SUCESSO Then Error 28960

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TiposDeCliente_Grava_EmTrans = SUCESSO

    Exit Function

Erro_TiposDeCliente_Grava_EmTrans:

    TiposDeCliente_Grava_EmTrans = Err

    Select Case Err

        Case 28945
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28947, 28960, 41626

        Case 41630
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM1", Err, objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem)
        
        Case 28948, 28949
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDECLIENTE2", Err, objTipoCliente.iCodigo, objTipoCliente.sDescricao)

        Case 28950
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_DESCR_DUPLICATA", Err, objTipoCliente.sDescricao)

        Case 28951, 28952
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSDECLIENTE1", Err, objTipoCliente.iCodigo)

        Case 28953
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TIPOSDECLIENTE", Err, objTipoCliente.iCodigo)

        Case 28954, 28955, 28958
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODECLIENTECATEGORIAS", Err, objTipoCliente.iCodigo)

        Case 28956
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TIPODECLIENTECATEGORIAS", Err, objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem)

        Case 28957
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPODECLIENTECATEGORIAS", Err, objTipoCliente.iCodigo)

        Case 28959
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPOSDECLIENTE", Err, objTipoCliente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148923)

    End Select

    'Libera comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function CPRConfig_ObterNumInt(ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComandoSel As Long, lComandoMod As Long
Dim sConteudo As String

On Error GoTo Erro_CPRConfig_ObterNumInt

    sConteudo = String(STRING_CPRCONFIG_CONTEUDO, 0)

    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then Error 7363

    lComandoMod = Comando_Abrir()
    If lComandoMod = 0 Then Error 7364

    lErro = Comando_ExecutarPos(lComandoSel, "SELECT Conteudo FROM CPRConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, sConteudo, sCodigo, EMPRESA_TODA)
    If lErro <> AD_SQL_SUCESSO Then Error 7365

    lErro = Comando_BuscarProximo(lComandoSel)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 7366

    lNumInt = CLng(sConteudo)
    sConteudo = CStr(lNumInt + iQtde)

    lErro = Comando_ExecutarPos(lComandoMod, "UPDATE CPRConfig SET Conteudo = ?", lComandoSel, sConteudo)
    If lErro <> AD_SQL_SUCESSO Then Error 7368

    lErro = Comando_Fechar(lComandoSel)
    lErro = Comando_Fechar(lComandoMod)

    CPRConfig_ObterNumInt = SUCESSO

    Exit Function

Erro_CPRConfig_ObterNumInt:

    CPRConfig_ObterNumInt = Err

    Select Case Err

        Case 7363, 7364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 7365, 7366
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", Err, sCodigo)

        Case 7368
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CPRCONFIG", Err, sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148924)

    End Select

    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoMod)

    Exit Function

End Function

Function TabelaDePreco_Lock(ByVal iTabelaPreco As Integer) As Long
'Faz lock da Tabela de Preço passada como parâmetro
'Chamada DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaDePreco_Lock
    
    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43125
    
    'Seleciona a Tabela de Venda à ser lockado
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM TabelasDePreco WHERE Codigo =?", iCodigo, iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then Error 43126

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43127

    'Não encontrou a Tabela de Preço --> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43128

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43129

    'Libera comando
    Call Comando_Fechar(lComando)

    TabelaDePreco_Lock = SUCESSO

    Exit Function

Erro_TabelaDePreco_Lock:

    TabelaDePreco_Lock = Err

    Select Case Err
        
        Case 43125
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
        Case 43126, 43127
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO1", Err, iTabelaPreco)
            
        Case 43128
            'Erro tratado na rotina chamada

        Case 43129
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECO", Err, iTabelaPreco)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148925)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function PadraoCobranca_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared do Padrão de Cobrança passada como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer
Dim iInativo As Integer

On Error GoTo Erro_PadraoCobranca_Lock
    
    'Pesquisa o Padrão de Cobrança no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo, Inativo FROM PadroesCobranca WHERE Codigo = ?", iCodigo1, iInativo, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 43442

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43443

    'Se não encontrou o Padrão de Cobrança--> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43444
    
    'Se o Padrão de Cobrança for Inativo
    If iInativo = 1 Then Error 61510
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43445
            
    PadraoCobranca_Lock = SUCESSO

    Exit Function

Erro_PadraoCobranca_Lock:

    PadraoCobranca_Lock = Err

    Select Case Err

        Case 43442, 43443
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PADRAO_COBRANCA", Err, iCodigo)

        Case 43444
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_NAO_CADASTRADO1", Err, iCodigo)

        Case 43445
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PADRAO_COBRANCA", Err, iCodigo)
        
        Case 61510
            Call Rotina_Erro(vbOKOnly, "ERRO_PADRAO_COBRANCA_INATIVO", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148926)

    End Select

    Exit Function

End Function

Private Function TipoCliente_Grava_CategoriaNaColecao(ByVal colCategoriaItemCopia As Collection, ByVal sCategoria As String) As Long
'Retorna o Indice da Categoria na coleção ou Zero se não achá-la (na coleção)

Dim iIndice As Integer
Dim objTipoClienteCategoria As New ClassTipoClienteCategoria

    'Pesquisa a Categoria na coleção
    For iIndice = 1 To colCategoriaItemCopia.Count

        Set objTipoClienteCategoria = colCategoriaItemCopia.Item(iIndice)

        'Se achou a Catagoria na coleção
        If objTipoClienteCategoria.sCategoria = sCategoria Then

            TipoCliente_Grava_CategoriaNaColecao = iIndice

            Exit Function

        End If

    Next

    TipoCliente_Grava_CategoriaNaColecao = SUCESSO

End Function

Private Function TipoCliente_Grava_NovasCategorias(ByVal objTipoCliente As ClassTipoCliente, ByVal colCategoriaItem As Collection) As Long
'Percorre as Categorias na coleção incluindo-as no BD

Dim lErro As Long, lComando As Long, lComando1 As Long
Dim objTipoClienteCategoria As New ClassTipoClienteCategoria

On Error GoTo Erro_TipoCliente_Grava_NovasCategorias

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28962
    
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 41627

    'Percorre a coleção
    For Each objTipoClienteCategoria In colCategoriaItem

        lErro = CF("CategoriaClienteItem_Lock", objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem, lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> 40555 Then Error 41628
        If lErro <> AD_SQL_SUCESSO Then Error 41629
        
        lErro = Comando_Executar(lComando, "INSERT INTO TipoDeClienteCategorias (TipoDeCliente, Categoria, Item) VALUES(?,?,?)", objTipoClienteCategoria.iTipoDeCliente, objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem)
        If lErro <> AD_SQL_SUCESSO Then Error 28963

    Next

    'Libera comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    TipoCliente_Grava_NovasCategorias = SUCESSO

    Exit Function

Erro_TipoCliente_Grava_NovasCategorias:

    TipoCliente_Grava_NovasCategorias = Err

    Select Case Err

        Case 28962, 41627
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28963
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TIPODECLIENTECATEGORIAs", Err, objTipoCliente.iCodigo)

        Case 41628
        
        Case 41629
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIACLIENTEITEM1", Err, objTipoClienteCategoria.sCategoria, objTipoClienteCategoria.sItem)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148927)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TiposDeCliente_Exclui(ByVal objTipoCliente As ClassTipoCliente) As Long
'Exclui o TipoCliente

Dim lErro As Long
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim iCodigo As Integer
Dim iFim_de_Arquivo As Integer

On Error GoTo Erro_TiposDeCliente_Exclui

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 28969

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 28970

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then Error 28971

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28972

    'Seleciona algum Cliente com esse codigo de Tipo de Cliente
    lErro = Comando_ExecutarPos(lComando1, "SELECT Codigo FROM Clientes WHERE Tipo =?", 0, lCodigo, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 19179

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19180

    'Verifica se existe algum Cliente relacionado com esse Tipo
    If lErro <> AD_SQL_SEM_DADOS Then Error 19181

    'Seleciona o TipoCliente a ser excluido
    lErro = Comando_ExecutarPos(lComando2, "SELECT Codigo FROM TiposDeCliente WHERE Codigo =?", 0, iCodigo, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 19182

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19183

    If lErro = AD_SQL_SEM_DADOS Then Error 19184

    'Lock do TipoCliente
    lErro = Comando_LockExclusive(lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 19185

    'Exclui da Tabela TipoClienteCategorias
    lErro = TipoCliente_Exclui_Categorias(objTipoCliente)
    If lErro <> SUCESSO Then Error 28973

    'Deleta o TipoCliente em questao
    lErro = Comando_ExecutarPos(lComando3, "DELETE FROM TiposDeCliente", lComando2)
    If lErro <> AD_SQL_SUCESSO Then Error 19186

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 19187

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    TiposDeCliente_Exclui = SUCESSO

    Exit Function

Erro_TiposDeCliente_Exclui:

    TiposDeCliente_Exclui = Err

    Select Case Err

        Case 28969, 28970, 28971
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28972
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 19179, 19180
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 19181
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOCLIENTE_RELACIONADO", Err, lCodigo)

        Case 19182, 19183
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTE", Err)

        Case 19184
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOCLIENTE_INEXISTENTE", Err)

        Case 19185
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TIPOCLIENTE", Err, objTipoCliente.iCodigo)

        Case 19186
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPOCLIENTE", Err, objTipoCliente.iCodigo)

        Case 19187
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 28973

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148928)

    End Select

    Call Transacao_Rollback

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)

    Exit Function

End Function

Private Function TipoCliente_Exclui_Categorias(ByVal objTipoCliente As ClassTipoCliente) As Long
'Exclui as Categorias de um Tipo de Cliente

Dim lErro As Long, iIndice As Integer
Dim alComando(2) As Long
Dim sCategoria As String, sItem As String

On Error GoTo Erro_TipoCliente_Exclui_Categorias

    'Abre comandos
    For iIndice = 0 To 1
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 28974
    Next

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Percorre todas as Categorias no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Categoria, Item FROM TipoDeClienteCategorias WHERE TipoDeCliente = ?", 0, sCategoria, sItem, objTipoCliente.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28975

    'Lê a Categoria, se estiver no BD
    lErro = Comando_BuscarProximo(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28976

    Do While lErro <> AD_SQL_SEM_DADOS

        'Excluir a Categoria no BD
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM TipoDeClienteCategorias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 28977

        'Lê a Categoria do Tipo de Cliente, se estiver no BD
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28978

    Loop

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    TipoCliente_Exclui_Categorias = SUCESSO

    Exit Function

Erro_TipoCliente_Exclui_Categorias:

    TipoCliente_Exclui_Categorias = Err

    Select Case Err

        Case 28974
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28975, 28976, 28978
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPODECLIENTECATEGORIAS", Err, objTipoCliente.iCodigo)

        Case 28977
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TIPODECLIENTECATEGORIAS", Err, objTipoCliente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148929)

    End Select

    For iIndice = 0 To 1
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TransfCartCobr_Grava(ByVal objTransfCartCobr As ClassTransfCartCobr) As Long
'Executa a Transferencia "manual" de uma parcela a receber entre carteiras de cobranca
'Atualiza a tabela de ParcelasRec e Insere os dados na tabela de TransfCartCobr
'Tambem atualiza as estatisticas

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iCobradorAnterior As Integer, iCobradorTeste As Integer
Dim iCarteiraAnterior As Integer, objCobrador As New ClassCobrador
Dim dSaldo As Double, lNumIntDoc As Long
Dim tTransfCartCobr As TypeTransfCartCobr
Dim objContabAutomatica As ClassContabAutomatica

On Error GoTo Erro_TransfCartCobr_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 48981
        
    lComando = 0
    lComando1 = 0

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 48982

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 48983
        
    'faz o select em ParcelasRec para pegar os dados do ultimo cobrador e carteira e o saldo e Proxima ocorrencia
    lErro = Comando_ExecutarPos(lComando, "SELECT Cobrador, CarteiraCobranca, Saldo, ProxSeqOcorr FROM ParcelasRec WHERE NumIntDoc = ?", 0, iCobradorAnterior, iCarteiraAnterior, dSaldo, tTransfCartCobr.iNumSeqOcorr, objTransfCartCobr.lNumIntParc)
    If lErro <> AD_SQL_SUCESSO Then Error 48984
    
    'Tenta ler a Parcelarec para alteracao do cobrador e sua carteira
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 48985

    'Faz "lock"
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 48986
        
    'obter cobrador a ser testado
    If iCobradorAnterior = COBRADOR_PROPRIA_EMPRESA Then
        iCobradorTeste = objTransfCartCobr.iCobrador
    Else
        iCobradorTeste = iCobradorAnterior
    End If
    
    If iCobradorTeste <> COBRADOR_PROPRIA_EMPRESA Then
    
        'ler dados do cobrador
        objCobrador.iCodigo = iCobradorTeste
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then Error 59205
        If lErro <> SUCESSO Then Error 59206
        
        'nao pode transferir titulo de/para cobrador com cobranca eletronica
        'If objCobrador.iCobrancaEletronica = 1 Then Error 59207
     
    End If
    
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        lErro = CF("Config_ObterNumInt", "CRFATConfig", "NUM_PROX_TRANSFCARTCOBR", lNumIntDoc)
        If lErro <> SUCESSO Then Error 56538
        
        objTransfCartCobr.lNumIntDoc = lNumIntDoc
    
        Set objContabAutomatica = New ClassContabAutomatica
    
        lErro = objContabAutomatica.Inicializa_Contab(objTransfCartCobr, MODULO_CONTASARECEBER, "BCR", objTransfCartCobr.dtData, DATA_NULA)
        If lErro <> SUCESSO Then Error 56539
        
        objTransfCartCobr.dSaldo = dSaldo
        lErro = objContabAutomatica.GeraContabilizacao(objTransfCartCobr)
        If lErro <> SUCESSO Then Error 56540
        
        lErro = objContabAutomatica.Finaliza_Contab()
        If lErro <> SUCESSO Then Error 56541
    
    End If
    
    'Atualiza a tabela
    lErro = Comando_ExecutarPos(lComando1, "UPDATE ParcelasRec SET Cobrador =?, CarteiraCobranca =? , ProxSeqOcorr =?, NumTitCobrador = ? ", lComando, objTransfCartCobr.iCobrador, objTransfCartCobr.iCarteiraCobranca, tTransfCartCobr.iNumSeqOcorr + 1, "")
    If lErro <> AD_SQL_SUCESSO Then Error 48987
    
    objTransfCartCobr.iNumSeqOcorr = tTransfCartCobr.iNumSeqOcorr
    
    lErro = CF("TransfCartCobr_Grava1", objTransfCartCobr, iCobradorAnterior, iCarteiraAnterior, dSaldo)
    If lErro <> SUCESSO Then Error 48988
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 48989

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    TransfCartCobr_Grava = SUCESSO
    
    Exit Function
        
Erro_TransfCartCobr_Grava:

    TransfCartCobr_Grava = Err
    
        Select Case Err
        
            Case 48981
                Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
            
            Case 48982, 48983
                Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                
            Case 48984, 48985
                Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", Err)
            
            Case 48986
                Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", Err)
                
            Case 48987
                Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", Err, , objTransfCartCobr.lNumIntParc)
            
            Case 48988, 56538, 56539, 56540, 56541, 59205
            
            Case 48989
                Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
                
            Case 59206
                Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)
            
            Case 59207
                Call Rotina_Erro(vbOKOnly, "ERRO_TRANSF_MANUAL_COBR_ELETRONICA", Err)
                
            Case Else
            
        End Select
        
        Call Transacao_Rollback
        Call Comando_Fechar(lComando)
        Call Comando_Fechar(lComando1)

        Exit Function
        
End Function

Function Transportadora_Exclui(ByVal objTransportadora As ClassTransportadora) As Long
'Exclui a Transportadora

Dim lErro As Long
Dim lComando(2) As Long
Dim lTransacao As Long
Dim colCodNotasFiscais As New Collection
Dim lTotal As Long
Dim iIndice As Integer
Dim iCodigo As Integer

On Error GoTo Erro_Transportadora_Exclui

    'Abre comandos
    For iIndice = 0 To 1
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then Error 22080
    Next

    'Inicia a Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 22081

    'Seleciona a Transportadora a ser excluída
    lErro = Comando_ExecutarPos(lComando(0), "SELECT Codigo FROM Transportadoras WHERE Codigo =?", 0, iCodigo, objTransportadora.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 22082

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 22083
    
    If lErro = AD_SQL_SEM_DADOS Then Error 22084

    'Lock da Transportadora
    lErro = Comando_LockExclusive(lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 22109

    'Verifica se Transportadora esta relacionada com Notas Fiscais
    lErro = NFs_Transportadora_Nao_Relacionada(iCodigo)
    If lErro <> SUCESSO Then Error 22108

    'verifica se transportadora está ligada a pedido de venda
    lErro = PVs_Transportadora_Nao_Relacionada(iCodigo)
    If lErro <> SUCESSO Then Error 19361
    
    'Verifica se Transportadora esta relacionada com FiliaisClientes
    lErro = FilCli_Transportadora_Nao_Relacionada(iCodigo)
    If lErro <> SUCESSO Then Error 22110
    
    'Verifica se Transportadora está relacionada com Tipo Cliente
    lErro = CF("TipoCliente_Transportadora_Nao_Relacionada", iCodigo)
    If lErro <> SUCESSO Then Error 58109
    
    'Exclui a Transportadora
    lErro = Comando_ExecutarPos(lComando(1), "DELETE FROM Transportadoras", lComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 22086
    
    'Exclui o Endereço da Transportadora
    lErro = CF("Endereco_Exclui", objTransportadora.lEndereco)
    If lErro <> SUCESSO Then Error 22121

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 22085

    'Fecha os comandos
    For iIndice = 0 To 1
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Transportadora_Exclui = SUCESSO

    Exit Function

Erro_Transportadora_Exclui:

    Transportadora_Exclui = Err

    Select Case Err

        'Tratados nas rotinas chamadas
        Case 19361, 22108, 22110, 22121, 58109

        Case 22109
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TRANSPORTADORA", Err)

        Case 22080
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 22081
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 22082, 22083
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", Err)

        Case 22084
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case 22086
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TRANSPORTADORA", Err, objTransportadora.iCodigo)

        Case 22085
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148930)

    End Select

    Call Transacao_Rollback

    For iIndice = 0 To 1
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Private Function NFs_Transportadora_Nao_Relacionada(ByVal iCodTransportadora As Integer) As Long
'Verifica se Transportadora está relacionada com Notas Fiscais

Dim lErro As Long, lComando As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_NFs_Transportadora_Nao_Relacionada

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27495

    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE CodTransportadora = ? OR CodTranspRedesp = ?", lNumIntDoc, iCodTransportadora, iCodTransportadora)
    If lErro <> AD_SQL_SUCESSO Then Error 27496

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27497
    
    'se achou alguma nota relacionada
    If lErro = AD_SQL_SUCESSO Then Error 27498
    
    lErro = Comando_Fechar(lComando)
    
    NFs_Transportadora_Nao_Relacionada = SUCESSO

    Exit Function

Erro_NFs_Transportadora_Nao_Relacionada:

    NFs_Transportadora_Nao_Relacionada = Err

    Select Case Err

        Case 27495
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
                    
        Case 27496, 27497
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFS_TRANSPORTADORA", Err)
        
        Case 27498 'achou alguma nota relacionada
            Call Rotina_Erro(vbOKOnly, "ERRO_EXISTEM_NFS_TRANSPORTADORA", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148931)

    End Select

    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function PVs_Transportadora_Nao_Relacionada(ByVal iCodigo As Integer) As Long
'Verifica se Tranportadora está relacionada com  Pedido de Venda ou Pedidos de VendaBaixados

Dim lErro As Long
Dim lComando As Long
Dim lCodigo1 As Long

On Error GoTo Erro_PVs_Transportadora_Nao_Relacionada
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19354
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM PedidosDeVenda WHERE CodTransportadora = ? OR CodTranspRedesp = ?", lCodigo1, iCodigo, iCodigo)
    If lErro <> SUCESSO Then Error 19355
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19356
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 19357
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM PedidosDeVendaBaixados WHERE CodTransportadora = ? OR CodTranspRedesp = ?", lCodigo1, iCodigo, iCodigo)
    If lErro <> SUCESSO Then Error 19358
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19359
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 19360
    
    Call Comando_Fechar(lComando)
    
    PVs_Transportadora_Nao_Relacionada = SUCESSO
    
    Exit Function
    
Erro_PVs_Transportadora_Nao_Relacionada:
    
    PVs_Transportadora_Nao_Relacionada = Err
    
    Select Case Err
            
        Case 19354
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 19355, 19356
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDAS1", Err)
            
        Case 19357
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_PEDIDODEVENDAS", Err, lCodigo1)
                    
        Case 19358, 19359
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDASBAIXADOS", Err)
            
        Case 19360
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_PEDIDODEVENDASBAIXADOS", Err, iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148932)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function FilCli_Transportadora_Nao_Relacionada(ByVal iCodigo As Integer)
'verifica se a transportadora identificada por iCodigo está relacionada a alguma filial de cliente

Dim lErro As Long
Dim lComando As Long
Dim lCodigo1 As Long

On Error GoTo Erro_FilCli_Transportadora_Nao_Relacionada
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 19362

    lErro = Comando_Executar(lComando, "SELECT CodCliente FROM FiliaisClientes WHERE CodTransportadora = ? OR CodTranspRedesp = ?", lCodigo1, iCodigo, iCodigo)
    If lErro <> SUCESSO Then Error 19363
        
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 19364
    
    'se está sendo utilizada ==>Erro.
    If lErro = AD_SQL_SUCESSO Then Error 19365
    
    Call Comando_Fechar(lComando)
    
    FilCli_Transportadora_Nao_Relacionada = SUCESSO
    
    Exit Function
    
Erro_FilCli_Transportadora_Nao_Relacionada:

    FilCli_Transportadora_Nao_Relacionada = Err
    
    Select Case Err
        
        Case 19362
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 19363, 19364
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", Err)
                
        Case 19365
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_FILIAISCLIENTES", Err, iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148933)
        
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function TabelaPreco_Altera_Trans(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Atualiza Tabela de Preço existente

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lTransacao As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Altera_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28136

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 28137

    'verificar se o nome é utilizado por outra tabela de preco
    
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TabelasDePreco WHERE Descricao = ? AND Codigo <> ?", iCodigo, objTabelaPreco.sDescricao, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 42740

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 42741
    
    If lErro = AD_SQL_SUCESSO Then Error 42742

    'Pesquisa no BD a Tabela de Preço
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM TabelasDePreco WHERE Codigo = ?", 0, iCodigo, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28139

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28140

    'Se não encontrou a Tabela no BD ==> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 28141
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 41532
    
    'atualiza a Tabela de Preço
    lErro = Comando_ExecutarPos(lComando1, "UPDATE TabelasDePreco SET Descricao = ?, Ativo = ?, AjustaTabVlrCompCoefMaior = ?, VlrCompCoef = ?, MarkUp = ?, CargoMinimo = ?, Moeda = ?, Comissao = ?, Tipo = ?, DescontoLimitado = ?, DescontoMaximo = ?", lComando, objTabelaPreco.sDescricao, objTabelaPreco.iAtivo, objTabelaPreco.iAjustaTabVlrCompCoefMaior, objTabelaPreco.dVlrCompCoef, objTabelaPreco.dMarkUp, objTabelaPreco.iCargoMinimo, objTabelaPreco.iMoeda, objTabelaPreco.dComissao, objTabelaPreco.iTipo, objTabelaPreco.iDescontoLimitado, objTabelaPreco.dDescontoMaximo)
    If lErro <> AD_SQL_SUCESSO Then Error 28142
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    TabelaPreco_Altera_Trans = SUCESSO

    Exit Function

Erro_TabelaPreco_Altera_Trans:

    TabelaPreco_Altera_Trans = Err

    Select Case Err

        Case 28136, 28137
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28139, 28140, 42740, 42741
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO1", Err, objTabelaPreco.iCodigo)

        Case 28141
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)

        Case 28142
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELASDEPRECO", Err, objTabelaPreco.iCodigo)

        Case 28143
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 42742
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TABELAPRECO_JA_EXISTENTE", Err, objTabelaPreco.sDescricao)
        
        Case 41532
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELAPRECO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148934)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function TabelaPreco_Altera(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Atualiza Tabela de Preço existente

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TabelaPreco_Altera

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 28138

    lErro = CF("TabelaPreco_Altera_Trans", objTabelaPreco)
    If lErro <> SUCESSO Then gError 112583
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 28143
    
    TabelaPreco_Altera = SUCESSO

    Exit Function

Erro_TabelaPreco_Altera:

    TabelaPreco_Altera = gErr

    Select Case gErr

        Case 28138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 28143
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 112583
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148935)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function TabelaPreco_Cria_Trans(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Cria uma nova Tabela de Preço

Dim lErro As Long
Dim lComando As Long
Dim iCodigo As Integer

On Error GoTo Erro_TabelaPreco_Cria_Trans

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28118
    
    'Pesquisa no BD a Tabela de Preço
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM TabelasDePreco WHERE Codigo = ?", 0, iCodigo, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28120

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28121

    'Encontrou a Tabela no BD ==> Erro
    If lErro = AD_SQL_SUCESSO Then Error 28122

    'verificar se o nome é utilizado por outra tabela de preco
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TabelasDePreco WHERE Descricao = ? AND Codigo <> ?", iCodigo, objTabelaPreco.sDescricao, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 28036

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28037
    
    If lErro = AD_SQL_SUCESSO Then Error 28197
    
    'Se não encontrou, insere Tabela de Preço
    lErro = Comando_Executar(lComando, "INSERT INTO TabelasDePreco (Codigo, Descricao, Ativo, AjustaTabVlrCompCoefMaior, VlrCompCoef, MarkUp, CargoMinimo, Moeda, Comissao, Tipo, DescontoLimitado, DescontoMaximo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", objTabelaPreco.iCodigo, objTabelaPreco.sDescricao, 1, objTabelaPreco.iAjustaTabVlrCompCoefMaior, objTabelaPreco.dVlrCompCoef, objTabelaPreco.dMarkUp, objTabelaPreco.iCargoMinimo, objTabelaPreco.iMoeda, objTabelaPreco.dComissao, objTabelaPreco.iTipo, objTabelaPreco.iDescontoLimitado, objTabelaPreco.dDescontoMaximo)
    If lErro <> AD_SQL_SUCESSO Then Error 28123

    Call Comando_Fechar(lComando)

    TabelaPreco_Cria_Trans = SUCESSO

    Exit Function

Erro_TabelaPreco_Cria_Trans:

    TabelaPreco_Cria_Trans = Err

    Select Case Err

        Case 28036, 28037, 28120, 28121
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO1", Err, objTabelaPreco.iCodigo)

        Case 28118
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28122
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_JA_EXISTENTE", Err, objTabelaPreco.iCodigo)

        Case 28123
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELASDEPRECO", Err, objTabelaPreco.iCodigo)

        Case 28124
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 28197
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_TABELAPRECO_JA_EXISTENTE", Err, objTabelaPreco.sDescricao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148936)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function TabelaPreco_Cria(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Cria uma nova Tabela de Preço

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TabelaPreco_Cria
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 28119
    
    lErro = CF("TabelaPreco_Cria_Trans", objTabelaPreco)
    If lErro <> SUCESSO Then gError 112577
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 28124

    TabelaPreco_Cria = SUCESSO

    Exit Function

Erro_TabelaPreco_Cria:

    TabelaPreco_Cria = gErr

    Select Case gErr
        
        Case 28119
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 28124
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case 112577
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148937)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function TabelaPreco_Automatico(lNumProxCodigo As Long) As Long
'Gera o próximo código da Tabela de Preço automático.

Dim lErro As Long

On Error GoTo Erro_TabelaPreco_Automatico
    
    'Chama a rotina que gera o sequencial
    lErro = CF("Config_ObterAutomatico", "CRFATConfig", "NUM_PROX_TABELA_PRECO", "TabelasDePreco", "Codigo", lNumProxCodigo)
    If lErro <> SUCESSO Then Error 28165

    TabelaPreco_Automatico = SUCESSO

    Exit Function

Erro_TabelaPreco_Automatico:

    TabelaPreco_Automatico = Err

    Select Case Err

        Case 28165

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148938)

    End Select

    Exit Function

End Function

Function CRFATConfig_ObterNumInt1(ByVal sCodigo As String, lNumInt As Long, Optional ByVal iQtde As Integer = 1) As Long
'obter um numero automatico a partir de um contador identificado por uma chave
'iQtde identifica quantos numeros devem ser "liberados" a partir do lNumInt retornado
'TEM QUE SER CHAMADA DENTRO DE TRANSACAO

Dim lErro As Long
Dim lComandoSel As Long, lComandoMod As Long
Dim sConteudo As String

On Error GoTo Erro_CRFATConfig_ObterNumInt1

    sConteudo = String(STRING_CONTEUDO, 0)

    lComandoSel = Comando_Abrir()
    If lComandoSel = 0 Then Error 7363

    lComandoMod = Comando_Abrir()
    If lComandoMod = 0 Then Error 7364

    'Obtém o Número Interno dos itens da NF
    lErro = CF("CRFATConfig_ObterNumInt", lComandoSel, lComandoMod, sCodigo, lNumInt, iQtde)
    If lErro <> SUCESSO Then Error 45681

    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoMod)

    CRFATConfig_ObterNumInt1 = SUCESSO

    Exit Function

Erro_CRFATConfig_ObterNumInt1:

    CRFATConfig_ObterNumInt1 = Err

    Select Case Err

        Case 7363, 7364
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 45681
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148939)

    End Select

    Call Comando_Fechar(lComandoSel)
    Call Comando_Fechar(lComandoMod)

    Exit Function

End Function

Function TabelaPreco_Exclui(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Exclui a Tabela e Preço do BD

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim sDescricao As String
Dim iCodTabela As Integer

On Error GoTo Erro_TabelaPreco_Exclui

    'Aberturas comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 28039
    Next
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28041

    sDescricao = String(STRING_TABELAPRECO_DESCRICAO, 0)

    'Pesquisa a Tabela de Preço no BD
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Descricao FROM TabelasDePreco WHERE Codigo = ?", 0, sDescricao, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28042

    'Verifica se existe
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28043
    If lErro = AD_SQL_SEM_DADOS Then Error 28044 'Não encontrou a Tabela de Preço

    'Lock da Tabela de Preço
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 28045
    
    'Verifica os objetos relacionados com TabelaPreco
    lErro = TabelaDePreco_Testa_Vinculo(objTabelaPreco)
    If lErro <> SUCESSO Then Error 58285
    
    'Chama TabelaPrecoItensExclui
    lErro = CF("TabelaPrecoItens_Exclui", objTabelaPreco)
    If lErro <> SUCESSO Then Error 28052
    
    'Exclui Tabela de Preço
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM TabelasDePreco", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then Error 28148

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28053

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    TabelaPreco_Exclui = SUCESSO

    Exit Function

Erro_TabelaPreco_Exclui:

    TabelaPreco_Exclui = Err

    Select Case Err

        Case 28039
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28041
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28042, 28043
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO1", Err, objTabelaPreco.iCodigo)

        Case 28044
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", Err, objTabelaPreco.iCodigo)

        Case 28045
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECO", Err, objTabelaPreco.iCodigo)

        Case 28052, 58285 'Tratados na rotinas chamadas

        Case 28053
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)

        Case 28148
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TABELASDEPRECO", Err, objTabelaPreco.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148940)

    End Select

    'Fechamento transação
    Call Transacao_Rollback

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function TabelaPrecoItens_Exclui(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Rotina responsável pela exclusão de itens da Tabela de Preço relacionados à TabelasDePreco
'DEVE SE CHAMADA DENTRO DE UMA TRANSAÇÃO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim sCodProduto As String

On Error GoTo Erro_TabelaPrecoItens_Exclui

    sCodProduto = String(STRING_PRODUTO, 0)

    'Abertura comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28054

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 28055

    'Pesquisa no BD os itens da tabela de preço passado como parametro
    lErro = Comando_ExecutarPos(lComando, "SELECT CodProduto FROM TabelasDePrecoItens WHERE CodTabela = ?", 0, sCodProduto, objTabelaPreco.iCodigo)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28057

    'Verifica se existem itens
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28058

    Do While lErro = AD_SQL_SUCESSO

        'Exclui o item da tabela de TabelasDePrecoItens
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TabelasDePrecoItens", lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 28060

        'Busca o próximo ítem
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28061

    Loop

    'Fechamento comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    TabelaPrecoItens_Exclui = SUCESSO

    Exit Function

Erro_TabelaPrecoItens_Exclui:

    TabelaPrecoItens_Exclui = Err

    Select Case Err

        Case 28054, 28055
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28057, 28058, 28061
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", Err, objTabelaPreco.iCodigo)

        Case 28059
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECOITENS1", Err, objTabelaPreco.iCodigo)

        Case 28060
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TABELASDEPRECOITENS", Err, objTabelaPreco.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148941)

    End Select

    'Fechamento comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

   Exit Function

End Function

Function TabelaPreco_Exclui_Precos(ByVal objTabelaPreco As ClassTabelaPreco) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim iCodTabela As Integer

On Error GoTo Erro_TabelaPreco_Exclui_Precos

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 28046
    Next
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28047
    
    lErro = Comando_ExecutarPos(alComando(0), "SELECT CodTabela FROM TabelasDePrecoItens WHERE CodTabela = ? AND FilialEmpresa = ?", 0, iCodTabela, objTabelaPreco.iCodigo, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 28048
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then Error 28049
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE TabelasDePrecoItens SET Preco = 0", alComando(0))
        If lErro <> SUCESSO Then Error 28050
        
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 28063
        
    Loop
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28035
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    TabelaPreco_Exclui_Precos = SUCESSO
    
    Exit Function
    
Erro_TabelaPreco_Exclui_Precos:

    TabelaPreco_Exclui_Precos = Err
    
    Select Case Err
    
        Case 28035
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case 28046
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 28047
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)
        
        Case 28048, 28049, 28063
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS1", Err)
        
        Case 28050
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TABELASDEPRECOITENS", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148942)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    Exit Function
        
End Function

Function ParcelasRec_Inserir(ByVal objParcelaReceber As ClassParcelaReceber, ByVal lComandoParcIns As Long) As Long
'auxiliar ao cancelamento de baixa de parcela a receber

Dim lErro As Long

On Error GoTo Erro_ParcelasRec_Inserir
    
    With objParcelaReceber
        
        If .iCarteiraCobranca = 0 Then
            .iCarteiraCobranca = CARTEIRA_CARTEIRA
        End If
        
        If .iCobrador = 0 Then
            .iCobrador = COBRADOR_PROPRIA_EMPRESA
        End If
    
        '######################################
        'ALTERADO POR WAGNER
        lErro = Comando_Executar(lComandoParcIns, "INSERT INTO ParcelasRec (NumIntDoc,NumIntTitulo,NumParcela,Status,DataVencimento,DataVencimentoReal,Saldo,Valor,Cobrador,CarteiraCobranca,NumTitCobrador,NumIntCheque,ProxSeqBaixa,ProxSeqOcorr,Desconto1Codigo, Desconto1Ate, Desconto1Valor,Desconto2Codigo, Desconto2Ate, Desconto2Valor,Desconto3Codigo, Desconto3Ate, Desconto3Valor, Aceite, Descontada, Previsao, Observacao, ValorOriginal, TipoPagto, CodConta, DataCredito, DataEmissaoCheque, BancoCheque, AgenciaCheque, ContaCorrenteCheque, NumeroCheque, DataDepositoCheque, AdmMeioPagto, Parcelamento, NumeroCartao, ValidadeCartao, AprovacaoCartao, DataTransacaoCartao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
        .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iCobrador, .iCarteiraCobranca, .sNumTitCobrador, .lNumIntCheque, .iProxSeqBaixa, .iProxSeqOcorr, .iDesconto1Codigo, .dtDesconto1Ate, .dDesconto1Valor, .iDesconto2Codigo, .dtDesconto2Ate, .dDesconto2Valor, .iDesconto3Codigo, .dtDesconto3Ate, .dDesconto3Valor, .iAceite, .iDescontada, .iPrevisao, .sObservacao, .dValorOriginal, .iTipoPagto, .iCodConta, .dtDataCredito, .dtDataEmissaoCheque, .iBancoCheque, .sAgenciaCheque, .sContaCorrenteCheque, .lNumeroCheque, .dtDataDepositoCheque, .iAdmMeioPagto, .iParcelamento, .sNumeroCartao, .dtValidadeCartao, .sAprovacaoCartao, .dtDataTransacaoCartao)
        '######################################
    End With
    If lErro <> AD_SQL_SUCESSO Then Error 46306

    ParcelasRec_Inserir = SUCESSO

    Exit Function

Erro_ParcelasRec_Inserir:

    ParcelasRec_Inserir = Err

    Select Case Err

        Case 46306
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148943)

    End Select

    Exit Function

End Function

Private Function CRFATConfig_Grava_Trans(ByVal objCRFATConfig As ClassCRFATConfig, alComando() As Long) As Long
'grava registro em crfatconfig. Apenas altera o conteudo, se nao existir dá erro.
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, tCRFATConfig As typeCRFATConfig

On Error GoTo Erro_CRFATConfig_Grava_Trans

    'inicializar buffer
    tCRFATConfig.sConteudo = String(STRING_CRFATCONFIG_CONTEUDO, 0)
    'ler registo
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Conteudo FROM CRFATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, tCRFATConfig.sConteudo, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then Error 55012
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 55013
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then Error 55014
        
    'se o conteudo foi modificado
    If objCRFATConfig.sConteudo <> tCRFATConfig.sConteudo Then
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CRFATConfig SET Conteudo = ?", alComando(1), objCRFATConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then Error 55015
    End If
    
    Exit Function
    
Erro_CRFATConfig_Grava_Trans:

    CRFATConfig_Grava_Trans = Err
     
    Select Case Err
     
        Case 55012, 55013
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRFATCONFIG2", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
     
        Case 55014
            Call Rotina_Erro(vbOKOnly, "ERRO_CRFATCONFIG_INEXISTENTE", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case 55015
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CRFATCONFIG", Err, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148944)
     
    End Select
     
    Exit Function

End Function

Function CRFATConfig_Grava_Configs_Trans(ByVal colConfigs As ColCRFATConfig) As Long
'grava na tabela CRFATConfig uma colecao de registros de configuracao
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, iIndice As Integer
Dim alComando(1 To 2) As Long
Dim objCRFATConfig As ClassCRFATConfig

On Error GoTo Erro_CRFATConfig_Grava_Configs_Trans
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 55016
    Next
    
   'para cada elemento da colecao
    For Each objCRFATConfig In colConfigs
    
        lErro = CRFATConfig_Grava_Trans(objCRFATConfig, alComando())
        If lErro <> SUCESSO Then Error 55017
        
    Next
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    CRFATConfig_Grava_Configs_Trans = SUCESSO
    
    Exit Function
    
Erro_CRFATConfig_Grava_Configs_Trans:
    
    CRFATConfig_Grava_Configs_Trans = Err
    
    Select Case Err
    
        Case 55016
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 55017
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148945)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function CRFATConfig_Grava_Configs(ByVal colConfigs As ColCRFATConfig) As Long
'grava na tabela CRFATConfig uma colecao de registros de configuracao

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_CRFATConfig_Grava_Configs
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 56556

    lErro = CF("CRFATConfig_Grava_Configs_Trans", colConfigs)
    If lErro <> SUCESSO Then Error 56557

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 56558
          
    CRFATConfig_Grava_Configs = SUCESSO
    
    Exit Function
    
Erro_CRFATConfig_Grava_Configs:
    
    CRFATConfig_Grava_Configs = Err
    
    Select Case Err
    
        Case 56556
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 56557

        Case 56558
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148946)
            
    End Select
    
    Call Transacao_Rollback
    
    Exit Function
    
End Function

Private Function Clientes_Le_CondicaoPagto(ByVal iCondicaoPagto As Integer, lTotal As Long) As Long
'retorna a qtde de clientes associados a cond de pagto
'retorna 16461 se nao houver cliente associado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Clientes_Le_CondicaoPagto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16458

    'Pesquisa Clientes no BD
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Clientes WHERE CondicaoPagto = ? ", lTotal, iCondicaoPagto)
    If lErro <> AD_SQL_SUCESSO Then Error 16459

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16460

    If lTotal = 0 Then Error 16461

    Call Comando_Fechar(lComando)

    Clientes_Le_CondicaoPagto = SUCESSO

    Exit Function

Erro_Clientes_Le_CondicaoPagto:

    Clientes_Le_CondicaoPagto = Err

    Select Case Err

        Case 16458
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16459, 16460
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)

        Case 16461

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148947)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Fornecedores_Le_CondicaoPagto(ByVal iCondicaoPagto As Integer, lTotal As Long) As Long
'retorna a qtde de fornecedores associados a cond de pagto
'retorna 16465 se nao houver fornecedor associado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Fornecedores_Le_CondicaoPagto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16462

    'pesquisa Fornecedores no BD
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Fornecedores WHERE CondicaoPagto = ? ", lTotal, iCondicaoPagto)
    If lErro <> AD_SQL_SUCESSO Then Error 16463

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16464

    If lTotal = 0 Then Error 16465

    Call Comando_Fechar(lComando)

    Fornecedores_Le_CondicaoPagto = SUCESSO

    Exit Function

Erro_Fornecedores_Le_CondicaoPagto:

    Fornecedores_Le_CondicaoPagto = Err

    Select Case Err

        Case 16462
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16463, 16464
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORNECEDORES", Err)

        Case 16465

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148948)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TiposCliente_Le_CondicaoPagto(ByVal iCondicaoPagto As Integer, lTotal As Long) As Long
'retorna a qtde de tipos de clientes associados a cond de pagto
'retorna 16469 se nao houver tipos de cliente associado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TiposCliente_Le_CondicaoPagto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16466

    'Pesquisa Tipos de Cliente no BD
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM TiposDeCliente WHERE CondicaoPagto = ? ", lTotal, iCondicaoPagto)
    If lErro <> AD_SQL_SUCESSO Then Error 16467

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16468

    If lTotal = 0 Then Error 16469

    Call Comando_Fechar(lComando)

    TiposCliente_Le_CondicaoPagto = SUCESSO

    Exit Function

Erro_TiposCliente_Le_CondicaoPagto:

    TiposCliente_Le_CondicaoPagto = Err

    Select Case Err

        Case 16466
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16467, 16468
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOCLIENTE", Err)

        Case 16469

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148949)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TiposFornecedor_Le_CondicaoPagto(ByVal iCondicaoPagto As Integer, lTotal As Long) As Long
'retorna a qtde de tipos de fornecedor associados a cond de pagto
'retorna 16473 se nao houver tipos de fornecedor associado

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_TiposFornecedor_Le_CondicaoPagto

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 16470

    'Pesquisa Tipos de Fornecedor no BD
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM TiposDeFornecedor WHERE CondicaoPagto = ? ", lTotal, iCondicaoPagto)
    If lErro <> AD_SQL_SUCESSO Then Error 16471

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 16472

    If lTotal = 0 Then Error 16473

    Call Comando_Fechar(lComando)

    TiposFornecedor_Le_CondicaoPagto = SUCESSO

    Exit Function

Erro_TiposFornecedor_Le_CondicaoPagto:

    TiposFornecedor_Le_CondicaoPagto = Err

    Select Case Err

        Case 16470
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 16471, 16472
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPOSFORNECEDOR", Err)

        Case 16473

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148950)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function DocumentoCPR_OrigemNFiscal(ByVal sSigla As String, ByVal lNumIntDoc As Long, lNumIntNF As Long) As Long
'retorna SUCESSO se existe registro em NFiscal associado a um docto em CPR identificado pela sigla conforme a tabela TiposDeDocumento e lNumIntDoc na tabela em CPR preenchendo em lNumIntNF o NumIntDoc da tabela NFiscal
'e retorna 41542 se nao existir
'se existir isto quer dizer que o documento em CPR foi criado atraves do registro da nota fiscal nos modulos de faturamento ou estoque

Dim lErro As Long, lComando As Long
Dim objTipoDocumento As New ClassTipoDocumento
        
On Error GoTo Erro_DocumentoCPR_OrigemNFiscal

    'Abertura comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 43001

    objTipoDocumento.sSigla = sSigla
    
    lErro = CF("TipoDocumento_Le", objTipoDocumento)
    If lErro <> SUCESSO And lErro <> 17078 Then Error 18905
    
    'Não achou o Tipo de Documento
    If lErro = 17078 Then Error 18906
    
    'Lê a NFiscal
    lErro = Comando_Executar(lComando, "SELECT NumIntDoc FROM NFiscal WHERE ClasseDocCPR = ? AND NumIntDocCPR = ?", lNumIntNF, objTipoDocumento.iClasseDocCPR, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then Error 43002

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43003

    'Não encontrou ==> erro
    If lErro = AD_SQL_SEM_DADOS Then Error 41542

    'Fechamento comando
    Call Comando_Fechar(lComando)

    DocumentoCPR_OrigemNFiscal = SUCESSO
    
    Exit Function
    
Erro_DocumentoCPR_OrigemNFiscal:

    DocumentoCPR_OrigemNFiscal = Err

    Select Case Err

        Case 43001
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 43002, 43003
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL2", Err, objTipoDocumento.iClasseDocCPR, lNumIntDoc)

        Case 41542
            'Não encontrou a NFiscal
        
        Case 18905
        
        Case 18906
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, sSigla)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148951)

    End Select

    'Fechamento comando
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function
    
Function ChequePre_Grava(ByVal objChequePre As ClassChequePre, ByVal colInfoParcRec As Collection) As Long
'Rotina que inclui ou altera um Cheque Pre e as associacoes com parcelas de colInfoParcRec

Dim lErro As Long, lTransacao As Long

On Error GoTo Erro_ChequePre_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 17703

    'exclui o movto de cta corrente correspondente ao cheque se existir
    lErro = CF("ChequePre_Grava_Trans", objChequePre, colInfoParcRec)
    If lErro <> SUCESSO Then gError 183030

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 17714

    ChequePre_Grava = SUCESSO

    Exit Function

Erro_ChequePre_Grava:

    ChequePre_Grava = gErr

    Select Case gErr
        
        Case 17703
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 17714
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 183030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183031)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function ChequePre_Exclui(ByVal objChequePre As ClassChequePre) As Long
'Rotina responsavel pela exclusao de Cheque Pre

Dim lErro As Long
Dim alComando(10) As Long
Dim iIndice As Integer
Dim lTransacao As Long
Dim lNumProxBorderoChqPre As Long
Dim tChequePre As typeChequePre
Dim lNumIntCheque As Long
Dim lNumIntParcela1 As Long
Dim lNumIntDoc As Long
Dim sConteudo As String
Dim lNumIntParcela As Long, dSaldoParcela As Double
Dim lNumIntBaixa As Long, dtData As Date

On Error GoTo Erro_ChequePre_Exclui

    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 17739

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 17740

    Next

    'Pesquisa no BD o cheque pre passado como parametro
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntCheque, FilialEmpresa, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero, Localizacao FROM ChequePreN WHERE Banco = ? AND Agencia = ? AND ContaCorrente = ? AND Numero = ? AND Status <> ?", 0, _
        tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, tChequePre.iLocalizacao, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17747

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17748
    
    'Se nao encontrou --> Erro
    If lErro = AD_SQL_SEM_DADOS Then gError 17749
    
    If tChequePre.iLocalizacao <> CHEQUEPRE_LOCALIZACAO_BACKOFFICE Then gError 210924

    objChequePre.lNumIntCheque = tChequePre.lNumIntCheque

    'Alterado por Rafael Menezes em 11/10/2002
    'if tChequePre.lNumBordero <> 0 Then gError 17750
    'verifica se o cheque está vinculado a um bordero
'    lErro = Comando_Executar(alComando(8), "SELECT NumIntCheque From ChequePre WHERE NumIntCheque = ? AND NumIntCheque IN (SELECT NumIntCheque From ChequeBordero)", lNumIntCheque, tChequePre.lNumIntCheque)
'    If lErro <> AD_SQL_SUCESSO Then gError 109130
'
'    'busca o primeiro
'    lErro = Comando_BuscarPrimeiro(alComando(8))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109131
'
'    'se encontrou->erro. Cheque vinculado a bordero não pode ser alterado ou excluído
'    If lErro = AD_SQL_SUCESSO Then gError 109132
    'Fim Alterado por Rafael Menezes em 11/10/2002

    'verifica se o cheque está vinculado a uma baixa a pagar
    lErro = Comando_Executar(alComando(9), "SELECT NumIntBaixa, Data From BaixasPag WHERE NumIntCheque = ? AND FilialEmpresaCheque = ? AND Status <> 5 ", lNumIntBaixa, dtData, tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 210920
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(9))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210921
    
    'se encontrou->erro. Cheque vinculado a BaixasPag
    If lErro <> AD_SQL_SEM_DADOS Then gError 209263
    
    'verifica se o cheque está vinculado a uma baixa a receber
    lErro = Comando_Executar(alComando(10), "SELECT NumIntBaixa, Data From BaixasRec WHERE NumIntCheque = ? AND FilialEmpresaCheque = ? AND Status <> 5 ", lNumIntBaixa, dtData, tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 210922
    
    'busca o primeiro
    lErro = Comando_BuscarPrimeiro(alComando(10))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210923
    
    'se encontrou->erro. Cheque vinculado a BaixasRec
    If lErro <> AD_SQL_SEM_DADOS Then gError 209264

    If (tChequePre.lCliente <> objChequePre.lCliente) Or (tChequePre.iFilial <> objChequePre.iFilial) Then gError 17781

    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 17751

    'Pesquisa as parcelas que estavam vinculadas ao cheque pre
    lErro = Comando_ExecutarPos(alComando(5), "SELECT NumIntCheque, NumIntParcela FROM ChequeParcelaRec WHERE NumIntCheque = ?", 0, lNumIntCheque, lNumIntParcela1, tChequePre.lNumIntCheque)
    If lErro <> AD_SQL_SUCESSO Then gError 91345

    lErro = Comando_BuscarPrimeiro(alComando(5))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17753
    
    Do While lErro = AD_SQL_SUCESSO
    
        'exclui a vinculacao da parcela ao cheque
        lErro = ChequeParcelaRec_Exclui(lNumIntCheque, lNumIntParcela1)
        If lErro <> SUCESSO Then gError 91308
    
        lErro = Comando_BuscarProximo(alComando(5))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17755

    Loop
    
'    'Exclui o cheque pre da tabela de ChequePre
'    lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ChequePre", alComando(1))
'    If lErro <> AD_SQL_SUCESSO Then gError 17756
    
    'Exclui o cheque pre da tabela de ChequePre
    lErro = Comando_ExecutarPos(alComando(4), "UPDATE ChequePreN SET Status = ?", alComando(1), STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 17756
    
    'exclui o movto de cta corrente correspondente ao cheque se existir
    lErro = CF("ChequePre_MovCCI_Exclui", objChequePre)
    If lErro <> SUCESSO Then gError 126167
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 17758

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ChequePre_Exclui = SUCESSO

    Exit Function

Erro_ChequePre_Exclui:

    ChequePre_Exclui = gErr

    Select Case gErr

        Case 15582, 126167

        Case 17739
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 17740
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 17758
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 17747, 17748
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE1", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17749
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAO_ENCONTRADO", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

'        Case 17750
'            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17751
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CHEQUESPRE", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17752, 17753, 17755
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr, tChequePre.lNumIntCheque)

        Case 17754
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", gErr, objChequePre.lNumIntCheque)

        Case 17756
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUESPRE", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17778, 17779
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CPRCONFIG", gErr, "NUM_PROX_BORDERO_CHQPAG")

        Case 17780
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CPRCONFIG", gErr, "NUM_PROX_BORDERO_CHQPRE")

        Case 17781
            Call Rotina_Erro(vbOKOnly, "ERRO_CLINTE_FILIAL_NAO_CONFEREM", gErr, objChequePre.lCliente, objChequePre.iFilial)
        
        Case 91314, 91315, 91317, 91345
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPARCELAREC", gErr)
        
        Case 91316
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUEPARCELAREC", gErr)
        
        Case 109130, 109131
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEBORDERO", gErr)
        
        Case 109132
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
       
        Case 209263
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_UTILIZADO_BAIXAPAG", gErr, lNumIntBaixa, dtData)
        
        Case 209264
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_UTILIZADO_BAIXAREC", gErr, lNumIntBaixa, dtData)
       
       Case 210920, 210921
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXAPAG", gErr)
       
       Case 210922, 210923
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASREC", gErr)
            
        Case 210924
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAO_BACKOFFICE1", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148953)

    End Select

    Call Transacao_Rollback

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function DebitoRecCli_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long

Dim lErro As Long
Dim objDebitoRecCli As New ClassDebitoRecCli

On Error GoTo Erro_DebitoRecCli_Gera

    objDebitoRecCli.dOutrasDespesas = objNFiscal.dValorOutrasDespesas
    objDebitoRecCli.dSaldo = Round(objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido), 2)
    objDebitoRecCli.dtDataEmissao = objNFiscal.dtDataEmissao
    objDebitoRecCli.dValorBaixado = objNFiscal.dValorTotal
    objDebitoRecCli.dValorFrete = objNFiscal.dValorFrete
    objDebitoRecCli.dValorProdutos = objNFiscal.dValorProdutos
    objDebitoRecCli.dValorSeguro = objNFiscal.dValorSeguro
    objDebitoRecCli.dValorTotal = objNFiscal.dValorTotal
    objDebitoRecCli.iFilial = objNFiscal.iFilialCli
    objDebitoRecCli.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objDebitoRecCli.iStatus = objNFiscal.iStatus
    objDebitoRecCli.lCliente = objNFiscal.lCliente
    objDebitoRecCli.lNumTitulo = objNFiscal.lNumNotaFiscal
    objDebitoRecCli.sSiglaDocumento = sTipoDocCPR
    
    objDebitoRecCli.dValorICMS = objNFiscal.objTributacaoNF.dICMSValor
    objDebitoRecCli.dValorICMSSubst = objNFiscal.objTributacaoNF.dICMSSubstValor
    objDebitoRecCli.dValorIPI = objNFiscal.objTributacaoNF.dIPIValor
    objDebitoRecCli.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
    objDebitoRecCli.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
    objDebitoRecCli.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
    objDebitoRecCli.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
    
    lErro = CF("DebitoRecCli_Insere", objDebitoRecCli)
    If lErro <> SUCESSO Then Error 46539
    
    lNumIntDocCPR = objDebitoRecCli.lNumIntDoc
    
    DebitoRecCli_Gera = SUCESSO
    
    Exit Function

Erro_DebitoRecCli_Gera:

    DebitoRecCli_Gera = Err
    
    Select Case Err
    
        Case 46539
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148954)
            
    End Select
    
    Exit Function
    
End Function

Function DebitoRecCli_Insere(ByVal objDebitoRecCli As ClassDebitoRecCli) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumProxDebitoRecCli As Long

On Error GoTo Erro_DebitoRecCli_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46540
    
    'verifica se foi entrado antes
    lErro = DebitoRecCli_Testa_Repeticao(objDebitoRecCli)
    If lErro <> SUCESSO Then Error 59250
    
    'Gera o número automático do próximo Débito
    lErro = CF("CPRConfig_ObterNumInt", "NUM_PROX_DEBITO_RECEBER", lNumProxDebitoRecCli)
    If lErro <> SUCESSO Then Error 46541

    objDebitoRecCli.dSaldo = Round(objDebitoRecCli.dValorTotal - (objDebitoRecCli.dValorIRRF + objDebitoRecCli.dPISRetido + objDebitoRecCli.dCOFINSRetido + objDebitoRecCli.dCSLLRetido), 2)
    objDebitoRecCli.lNumIntDoc = lNumProxDebitoRecCli

    If objDebitoRecCli.iFilialEmpresa = 0 Then objDebitoRecCli.iFilialEmpresa = giFilialEmpresa
    
    'Insere novo registro na tabela DebitosRecCli
    lErro = Comando_Executar(lComando, "INSERT INTO DebitosRecCli (NumIntDoc, FilialEmpresa, Cliente, Filial, SiglaDocumento, NumTitulo, Status, DataEmissao, ValorTotal, Saldo, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, ValorIPI, ValorIRRF, Observacao, PISRetido, COFINSRetido, CSLLRetido) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objDebitoRecCli.lNumIntDoc, objDebitoRecCli.iFilialEmpresa, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, objDebitoRecCli.sSiglaDocumento, objDebitoRecCli.lNumTitulo, STATUS_LANCADO, objDebitoRecCli.dtDataEmissao, objDebitoRecCli.dValorTotal, objDebitoRecCli.dSaldo, objDebitoRecCli.dValorSeguro, objDebitoRecCli.dValorFrete, objDebitoRecCli.dOutrasDespesas, objDebitoRecCli.dValorProdutos, objDebitoRecCli.dValorICMS, objDebitoRecCli.dValorICMSSubst, objDebitoRecCli.dValorIPI, objDebitoRecCli.dValorIRRF, objDebitoRecCli.sObservacao, objDebitoRecCli.dPISRetido, objDebitoRecCli.dCOFINSRetido, objDebitoRecCli.dCSLLRetido)
    If lErro <> AD_SQL_SUCESSO Then Error 46542

    Call Comando_Fechar(lComando)
    
    DebitoRecCli_Insere = SUCESSO
    
    Exit Function
    
Erro_DebitoRecCli_Insere:

    DebitoRecCli_Insere = Err
    
    Select Case Err
    
        Case 46540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 46541, 59250
        
        Case 46542
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DEBITOSRECCLI", Err, lNumProxDebitoRecCli)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148955)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function CreditoPagar_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long
'Gera um Crédito a pagar

Dim lErro As Long
Dim objCreditoPagar As New ClassCreditoPagar
Dim iDebitaICMS As Integer
Dim iDebitaIPI As Integer
Dim lComando As Long

On Error GoTo Erro_CreditoPagar_Gera
        
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 51635
    
    objCreditoPagar.dOutrasDespesas = objNFiscal.dValorOutrasDespesas
    objCreditoPagar.dSaldo = Round(objNFiscal.dValorTotal - (objNFiscal.objTributacaoNF.dIRRFValor + objNFiscal.objTributacaoNF.dPISRetido + objNFiscal.objTributacaoNF.dCOFINSRetido + objNFiscal.objTributacaoNF.dCSLLRetido), 2)
    objCreditoPagar.dtDataEmissao = objNFiscal.dtDataEmissao
    objCreditoPagar.dValorBaixado = objNFiscal.dValorTotal
    objCreditoPagar.dValorFrete = objNFiscal.dValorFrete
    objCreditoPagar.dValorProdutos = objNFiscal.dValorProdutos
    objCreditoPagar.dValorSeguro = objNFiscal.dValorSeguro
    objCreditoPagar.dValorTotal = objNFiscal.dValorTotal
    objCreditoPagar.iFilial = objNFiscal.iFilialForn
    objCreditoPagar.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objCreditoPagar.iStatus = objNFiscal.iStatus
    objCreditoPagar.lFornecedor = objNFiscal.lFornecedor
    objCreditoPagar.lNumTitulo = objNFiscal.lNumNotaFiscal
    objCreditoPagar.sSiglaDocumento = sTipoDocCPR
    
    objCreditoPagar.dValorICMS = objNFiscal.objTributacaoNF.dICMSValor
    objCreditoPagar.dValorICMSSubst = objNFiscal.objTributacaoNF.dICMSSubstValor
    objCreditoPagar.dValorIPI = objNFiscal.objTributacaoNF.dIPIValor
    objCreditoPagar.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
    objCreditoPagar.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
    objCreditoPagar.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
    objCreditoPagar.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
    objCreditoPagar.iDebitoICMS = 0
    objCreditoPagar.iDebitoIPI = 0
    
    If (objCreditoPagar.dValorIPI <> 0 Or objCreditoPagar.dValorICMS <> 0) Then
        
        lErro = Comando_Executar(lComando, "SELECT ICMSCredita,IPICredita FROM TiposDeTributacaoMovto WHERE Tipo = ? ", iDebitaICMS, iDebitaIPI, objNFiscal.objTributacaoNF.iTipoTributacao)
        If lErro <> AD_SQL_SUCESSO Then Error 51632
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And AD_SQL_SEM_DADOS Then Error 51633
        If lErro <> AD_SQL_SUCESSO Then Error 51634
        
        If iDebitaICMS = TRIB_GERA_DEBITO Then objCreditoPagar.iDebitoICMS = vbChecked
        If iDebitaIPI = TRIB_GERA_DEBITO Then objCreditoPagar.iDebitoIPI = vbChecked
                
    End If
    
    
    'Insere na tabela o crédito a pagar geradao
    lErro = CreditoPagar_Insere(objCreditoPagar)
    If lErro <> SUCESSO Then Error 46539
    
    lNumIntDocCPR = objCreditoPagar.lNumIntDoc
    
    Call Comando_Fechar(lComando)
    
    CreditoPagar_Gera = SUCESSO
    
    Exit Function

Erro_CreditoPagar_Gera:

    CreditoPagar_Gera = Err
    
    Select Case Err
    
        Case 46539
        
        Case 51632, 51633
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TIPO_TRIBUTACAO", Err, objNFiscal.objTributacaoNF.iTipoTributacao)
        
        Case 51634
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_CADASTRADO", Err, objNFiscal.objTributacaoNF.iTipoTributacao)
        
        Case 51635
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148956)
            
    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Function CreditoPagar_Insere(ByVal objCreditoPagar As ClassCreditoPagar) As Long

Dim lErro As Long
Dim lComando As Long
Dim lNumProxCreditoPagar As Long

On Error GoTo Erro_CreditoPagar_Insere

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 46540
    
    lErro = CreditoPagar_Testa_Repeticao(objCreditoPagar)
    If lErro <> SUCESSO Then Error 59251
    
    'Gera o número automático do próximo crédito
    lErro = CF("CPRConfig_ObterNumInt", "NUM_PROX_CREDITO_PAGAR", lNumProxCreditoPagar)
    If lErro <> SUCESSO Then Error 46541

    objCreditoPagar.dSaldo = Round(objCreditoPagar.dValorTotal - (objCreditoPagar.dValorIRRF + objCreditoPagar.dPISRetido + objCreditoPagar.dCOFINSRetido + objCreditoPagar.dCSLLRetido), 2)
    objCreditoPagar.lNumIntDoc = lNumProxCreditoPagar

    If objCreditoPagar.iFilialEmpresa = 0 Then objCreditoPagar.iFilialEmpresa = giFilialEmpresa
    
    'Insere novo registro na tabela CreditosPagForn
    lErro = Comando_Executar(lComando, "INSERT INTO CreditosPagForn (NumIntDoc, FilialEmpresa, Fornecedor, Filial, SiglaDocumento, NumTitulo, Status, DataEmissao, ValorTotal, Saldo, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorICMS, ValorICMSSubst, ValorIPI, ValorIRRF, Observacao, PISRetido, COFINSRetido, CSLLRetido) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", objCreditoPagar.lNumIntDoc, objCreditoPagar.iFilialEmpresa, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo, STATUS_LANCADO, objCreditoPagar.dtDataEmissao, objCreditoPagar.dValorTotal, objCreditoPagar.dSaldo, objCreditoPagar.dValorSeguro, objCreditoPagar.dValorFrete, objCreditoPagar.dOutrasDespesas, objCreditoPagar.dValorProdutos, objCreditoPagar.dValorICMS, objCreditoPagar.dValorICMSSubst, objCreditoPagar.dValorIPI, objCreditoPagar.dValorIRRF, objCreditoPagar.sObservacao, objCreditoPagar.dPISRetido, objCreditoPagar.dCOFINSRetido, objCreditoPagar.dCSLLRetido)
    If lErro <> AD_SQL_SUCESSO Then Error 46542

    Call Comando_Fechar(lComando)
    
    CreditoPagar_Insere = SUCESSO
    
    Exit Function
    
Erro_CreditoPagar_Insere:

    CreditoPagar_Insere = Err
    
    Select Case Err
    
        Case 46540
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 46541, 59251
        
        Case 46542
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CREDITOSPAGFORN", Err, lNumProxCreditoPagar)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148957)
            
    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function

End Function


Function Mensagem_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared da Mensagem passada como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer

On Error GoTo Erro_Mensagem_Lock
    
    'Pesquisa a Mensagem no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo FROM Mensagens WHERE Codigo = ?", iCodigo1, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 43427

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43428

    'Se não encontrou a Mensagem --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43429
    
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43430
            
    Mensagem_Lock = SUCESSO

    Exit Function

Erro_Mensagem_Lock:

    Mensagem_Lock = Err

    Select Case Err

        Case 43427, 43428
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MENSAGEM", Err, iCodigo)

        Case 43429
            Call Rotina_Erro(vbOKOnly, "ERRO_MENSAGEM_NAO_CADASTRADA", Err, iCodigo)

        Case 43430
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_MENSAGEM", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148958)

    End Select

    Exit Function

End Function


Function Cobrador_Lock(ByVal lComando As Long, ByVal iCodigo As Integer) As Long
'Rotina executada dentro de uma transação
'Faz lock shared do Cobrador passado como parâmetro

Dim lErro As Long
Dim iCodigo1 As Integer
Dim iInativo As Integer

On Error GoTo Erro_Cobrador_Lock
    
    'Pesquisa o Cobrador no BD
    lErro = Comando_ExecutarLockado(lComando, "SELECT Codigo, Inativo FROM Cobradores WHERE Codigo = ?", iCodigo1, iInativo, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 43433

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 43434

    'Se não encontrou o Cobrador --> Erro
    If lErro = AD_SQL_SEM_DADOS Then Error 43435
    
    'Se o Cobrador for Inativo --> Erro
    If iInativo = 1 Then Error 61507

    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then Error 43436
            
    Cobrador_Lock = SUCESSO

    Exit Function

Erro_Cobrador_Lock:

    Cobrador_Lock = Err

    Select Case Err

        Case 43433, 43434
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COBRADOR", Err, iCodigo)

        Case 43435
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, iCodigo)

        Case 43436
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_COBRADOR", Err)
        
        Case 61507
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_INATIVO", Err, iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148959)

    End Select

    Exit Function

End Function

Function ComissoesNF_Gera_ComissoesParc(ByVal objNFiscal As ClassNFiscal, colcolComissao As colcolComissao, ByVal dValorLiqTitulo As Double) As Long
'Preenche a colecao de colecoes de comissoes a nivel de parcela de um titulo a receber associado a NFiscal

Dim lErro As Long, objComissao As New ClassComissao, iIndice As Integer
Dim colComissao As colComissao, objComissaoNF As ClassComissaoNF
Dim dtDataGeracao As Date, iMotivoBaixa As Integer, iStatusComissao As Integer

On Error GoTo Erro_ComissoesNF_Gera_ComissoesParc

    'Se a coleção de parcelas estiver preenchida
    If objNFiscal.colParcelaReceber.Count > 0 Then
    
        'Para cada comissão da Nota Fiscal
        For Each objComissaoNF In objNFiscal.ColComissoesNF
        
                With objComissao
                    'Recolhe o percentual , vendedor e filial empresa da comissão
                    .dPercentual = objComissaoNF.dPercentual * (1 - objComissaoNF.dPercentualEmissao)
                    .iCodVendedor = objComissaoNF.iCodVendedor
                    .iFilialEmpresa = objNFiscal.iFilialEmpresa
                    
                End With
                
                'Para cada parcela a receber ligada a nota
                For iIndice = 1 To objNFiscal.colParcelaReceber.Count
                    
                    
                    Set colComissao = colcolComissao.Item(iIndice).colComissao
                    dtDataGeracao = DATA_NULA
                    iStatusComissao = STATUS_PENDENTE
                    
                    If objNFiscal.colParcelaReceber(iIndice).iStatus = STATUS_BAIXADO Then
                        
                        lErro = CF("ParcelaRecBaixada_Le_DataBaixa", objNFiscal.colParcelaReceber(iIndice).lNumIntDoc, dtDataGeracao, iMotivoBaixa)
                        If lErro <> SUCESSO And lErro <> 86189 Then gError 86190
                        
                        If iMotivoBaixa <> MOTIVO_PERDA Then iStatusComissao = STATUS_LIBERADO
                    
                    End If
                    
                    objComissao.dtDataGeracao = dtDataGeracao
                    
                    
                    objComissao.dValorBase = objComissaoNF.dValorBase * objNFiscal.colParcelaReceber.Item(iIndice).dValor / dValorLiqTitulo
                    objComissao.dValorBase = CDbl(Format(objComissao.dValorBase, "0.00"))
                    objComissao.dValor = objComissao.dValorBase * objComissao.dPercentual
                    objComissao.dValor = CDbl(Format(objComissao.dValor, "0.00"))
                    
                    With objComissao
                        colComissao.Add 0, .iFilialEmpresa, iStatusComissao, TIPO_PARCELA, 0, .iCodVendedor, DATA_NULA, .dPercentual, .dValorBase, .dValor, .dtDataGeracao
                    End With
                Next
                
        Next
        
    End If
    
    ComissoesNF_Gera_ComissoesParc = SUCESSO
    
    Exit Function
    
Erro_ComissoesNF_Gera_ComissoesParc:

    ComissoesNF_Gera_ComissoesParc = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148960)

    End Select

    Exit Function

End Function



Function TabelaDePreco_Testa_Vinculo(ByVal objTabelaPreco As ClassTabelaPreco) As Long
'Verifica se tem algum objeto relacionado

Dim iCodigo As Integer
Dim lErro As Long

On Error GoTo Erro_TabelaDePreco_Testa_Vinculo
    
    iCodigo = objTabelaPreco.iCodigo
    
    'Verifica Cliente
    lErro = Cliente_Le_TabelaDePreco(iCodigo)
    If lErro <> SUCESSO Then Error 58279
    
    'Verifica NFiscal
    lErro = NFiscal_Le_TabelaDePreco(iCodigo)
    If lErro <> SUCESSO Then Error 58280
    
    'Verifica Pedido de Venda e Pedido de Venda Baixados
    lErro = PedidoDeVendas_le_TabelaDePreco(iCodigo)
    If lErro <> SUCESSO Then Error 58281
    
    'Verifica Tipo de Cliente
    lErro = TipoCliente_Le_TabelaDePreco(iCodigo)
    If lErro <> SUCESSO Then Error 58282
        
    TabelaDePreco_Testa_Vinculo = SUCESSO
        
    Exit Function
    
Erro_TabelaDePreco_Testa_Vinculo:

    TabelaDePreco_Testa_Vinculo = Err
    
    Select Case Err
        
        Case 58279, 58280, 58281, 58282
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148961)

    End Select

    Exit Function
        
End Function


'FERNANDO IR PARA O SELECT
Function Cliente_Le_TabelaDePreco(ByVal iCodigo As Integer) As Long
'Verifica se existe ligacao entre a Tabela de Preço Passada e cliente.

Dim lErro As Long
Dim lComando As Long
Dim lCodigoCliente As Long

On Error GoTo Erro_Cliente_Le_TabelaDePreco
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58260

    'Pesquisa Tabela de Preco na Tabela de Clientes
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM Clientes WHERE TabelaPreco= ? ", lCodigoCliente, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 58261

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58262
    
    'Existe Tabela de Preco relacionado com Clientes --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 58263
    
    lErro = Comando_Fechar(lComando)

    Cliente_Le_TabelaDePreco = SUCESSO
    
    Exit Function
    
Erro_Cliente_Le_TabelaDePreco:

    Cliente_Le_TabelaDePreco = Err

    Select Case Err
        
        Case 58260
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58261, 58262
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", Err)
                    
        Case 58263 'Encontrou a Tabela De Preco na Tabela de Cliente
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_RELACIONADA_CLIENTE", Err, iCodigo, lCodigoCliente)
                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148962)

    End Select

    lErro = Comando_Fechar(lComando)

    Exit Function
        
End Function

'FERNANDO IR PARA O SELECT

Function NFiscal_Le_TabelaDePreco(ByVal iCodigo As Integer) As Long
'Verifica se existe ligacao entre a Tabela de Preço Passada e NFiscal.

Dim lErro As Long
Dim lComando As Long
Dim sSerie As String
Dim lNumeroNF As Long
Dim iFilialNF As Integer

On Error GoTo Erro_NFiscal_Le_TabelaDePreco
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58264
    
    sSerie = String(STRING_SERIE, 0)
    
    'Pesquisa Tabela de Preco na Tabela de NFiscal
    lErro = Comando_Executar(lComando, "SELECT Serie, NumNotaFiscal, FilialEmpresa FROM NFiscal WHERE TabelaPreco= ? ", sSerie, lNumeroNF, iFilialNF, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 58265

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58266
    
    'Existe Tabela de Preco relacionado com NFiscal --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 58267
    
    lErro = Comando_Fechar(lComando)

    NFiscal_Le_TabelaDePreco = SUCESSO
    
    Exit Function
    
Erro_NFiscal_Le_TabelaDePreco:

    NFiscal_Le_TabelaDePreco = Err

    Select Case Err
        
        Case 58264
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58265, 58266
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL4", Err)
                    
        Case 58267 'Encontrou a Tabela De Preco na Tabela de NFiscal
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_RELACIONADA_NFISCAL", Err, iCodigo, sSerie, lNumeroNF, iFilialNF)
                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148963)

    End Select
    
    lErro = Comando_Fechar(lComando)

    Exit Function
        
End Function

'FERNANDO IR PARA O SELECT

Function PedidoDeVendas_le_TabelaDePreco(ByVal iCodigo As Integer) As Long
'Verifica se existe ligacao entre a Tabela de Preço Passada e o Ped. de Venda e Ped. Venda Baixado.

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lCodigoPV As Long
Dim iFilialEmpresaPV As Integer
Dim lCodigoPVBaixado As Long
Dim iFilialEmpresaPVBaixado As Integer

On Error GoTo Erro_PedidoDeVendas_le_TabelaDePreco
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58268
    
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 58290

    'Pesquisa Tabela de Preco na Tabela de PedidoVenda
    lErro = Comando_Executar(lComando, "SELECT Codigo, FilialEmpresa FROM PedidosDeVenda WHERE TabelaPreco= ? ", lCodigoPV, iFilialEmpresaPV, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 58269

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58270
    
    'Existe Tabela de Preco relacionado com Pedido de Venda --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 58271
    
    'Pesquisa Tabela de Preco na Tabela de Pedido de Venda Baixado
    lErro = Comando_Executar(lComando2, "SELECT Codigo, FilialEmpresa FROM PedidosDeVendaBaixados WHERE TabelaPreco= ? ", lCodigoPVBaixado, iFilialEmpresaPVBaixado, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 58272

    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58273
    
    'Existe Tabela de Preco relacionado com Pedido de Venda Baixado --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 58274
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)

    PedidoDeVendas_le_TabelaDePreco = SUCESSO
    
    Exit Function
    
Erro_PedidoDeVendas_le_TabelaDePreco:

    PedidoDeVendas_le_TabelaDePreco = Err

    Select Case Err
        
        Case 58268, 58290
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58269, 58270
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDA", Err)
                    
        Case 58271 'Encontrou a Tabela De Preco na Tabela de Pedido de Venda
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_RELACIONADA_PEDVENDA", Err, iCodigo, lCodigoPV, iFilialEmpresaPV)
                      
        Case 58272, 58273
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDODEVENDABAIXADOS", Err)
                    
        Case 58274 'Encontrou a Tabela De Preco na Tabela de Pedido de Venda Baixado
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_RELACIONADA_PEDVENDA_BAIXADO", Err, iCodigo, lCodigoPVBaixado, iFilialEmpresaPVBaixado)
                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148964)

    End Select
    
    lErro = Comando_Fechar(lComando)
    lErro = Comando_Fechar(lComando2)

    Exit Function
        
End Function

'FERNANDO IR PARA O SELECT

Function TipoCliente_Le_TabelaDePreco(ByVal iCodigo As Integer) As Long
'Verifica se existe ligacao entre a Tabela de Preço Passada e TipoCliente.

Dim lErro As Long
Dim iCodigoTipo As Integer
Dim lComando As Long

On Error GoTo Erro_TipoCliente_Le_TabelaDePreco
    
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 58275

    'Pesquisa Tabela de Preco na Tabela de TipoDeCliente
    lErro = Comando_Executar(lComando, "SELECT Codigo FROM TiposDeCliente WHERE TabelaPreco= ? ", iCodigoTipo, iCodigo)
    If lErro <> AD_SQL_SUCESSO Then Error 58276

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 58277
    
    'Existe Tabela de Preco relacionado com TipoDeCliente --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 58278
    
    Call Comando_Fechar(lComando)

    TipoCliente_Le_TabelaDePreco = SUCESSO
    
    Exit Function
    
Erro_TipoCliente_Le_TabelaDePreco:

    TipoCliente_Le_TabelaDePreco = Err

    Select Case Err
        
        Case 58275
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 58276, 58277
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL4", Err)
                    
        Case 58278 'Encontrou a Tabela De Preco na Tabela de TiposDeCliente
            Call Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_RELACIONADA_TIPOSDECLIENTE", Err, iCodigo, iCodigoTipo)
                      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148965)

    End Select
    
    Call Comando_Fechar(lComando)

    Exit Function
        
End Function

Private Function ChequePre_MesmaIdentificacao(ByVal objChequePre1 As ClassChequePre, ByVal objChequePre2 As ClassChequePre) As Boolean
'retorna true se os dois cheques tiverem o mesmo bco, agencia, cta e numero.

    ChequePre_MesmaIdentificacao = (objChequePre1.iBanco = objChequePre2.iBanco And objChequePre1.sAgencia = objChequePre2.sAgencia And objChequePre1.sContaCorrente = objChequePre2.sContaCorrente And objChequePre1.lNumero = objChequePre2.lNumero)
    
End Function

Private Function ChequePre_OutroTitulo(ByVal objChequePre As ClassChequePre, ByVal lNumIntTitulo As Long) As Long
'retorna SUCESSO se o cheque nao está cadastrado para outro titulo

Dim lErro As Long
Dim lComando As Long, lNumIntCheque As Long

On Error GoTo ERRO_CHEQUEPRE_OUTROTITULO

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 59160

    'Pesquisa cheque pré-datado em outro titulo em aberto
    lErro = Comando_Executar(lComando, "SELECT ChequePre.NumIntCheque FROM ChequePre, ParcelasRec,ChequeParcelaRec WHERE ParcelasRec.NumIntTitulo <> ? AND ChequePre.NumIntCheque = ChequeParcelaRec.NumIntCheque AND ChequeParcelaRec.NumIntParcela = ParcelasRec.NumIntDoc AND ChequePre.Banco = ? AND ChequePre.Agencia = ? AND ChequePre.ContaCorrente = ? AND ChequePre.Numero = ?", lNumIntCheque, lNumIntTitulo, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
    If lErro <> AD_SQL_SUCESSO Then Error 59161
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59162
    
    'Este cheque pré-datado é/foi usado em outro titulo --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 59163
    
    'verifica se o chequepre já foi depositado
    lErro = Comando_Executar(lComando, "SELECT NumIntCheque FROM ChequePre WHERE Banco = ? AND Agencia = ? AND ContaCorrente = ? AND Numero = ? AND NumBordero <> 0", lNumIntCheque, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
    If lErro <> AD_SQL_SUCESSO Then Error 59186

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59187
    
    'Este cheque pré-datado é/foi depositado --> ERRO
    If lErro = AD_SQL_SUCESSO Then Error 59188
    
    Call Comando_Fechar(lComando)
    
    ChequePre_OutroTitulo = SUCESSO
     
    Exit Function
    
ERRO_CHEQUEPRE_OUTROTITULO:

    ChequePre_OutroTitulo = Err
     
    Select Case Err
          
        Case 59160
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 59161, 59162, 59186, 59187
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE3", Err)
            
        Case 59163
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_OUTROTITULO", Err)
        
        Case 59188
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_DEPOSITADO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148966)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function CreditoPagar_Testa_Repeticao(ByVal objCreditoPagar As ClassCreditoPagar) As Long
'Pede confirmacao do usuario se tentar inserir novo registro que possivelmente tenha sido entrado antes
'Chamada para tentar evitar entradas em duplicidade
'Obs.: estou desconsiderando a sigla do docto

Dim lErro As Long, lComando As Long, dtDataEmissao As Date
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CreditoPagar_Testa_Repeticao

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 59241
    
    'Verifica se Data Emissão está preenchida
    If objCreditoPagar.dtDataEmissao <> DATA_NULA Then
        
        'Tenta lerCrédito a Pagar de mesmo forn e numero com data emissão dentro da Faixa ou DATA_NULA
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM CreditosPagForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND (( DataEmissao > ? AND DataEmissao < ? ) OR DataEmissao = ? ) AND Status <> ?", _
            dtDataEmissao, objCreditoPagar.iFilialEmpresa, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.lNumTitulo, (objCreditoPagar.dtDataEmissao - PERIODO_EMISSAO), (objCreditoPagar.dtDataEmissao + PERIODO_EMISSAO), DATA_NULA, STATUS_EXCLUIDO)
        
    Else
        
        'Tenta lerCrédito a Pagar de mesmo forn e numero independente de data emissão
        lErro = Comando_Executar(lComando, "SELECT DataEmissao FROM CreditosPagForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND Filial = ? AND NumTitulo = ? AND Status <> ?", _
            dtDataEmissao, objCreditoPagar.iFilialEmpresa, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.lNumTitulo, STATUS_EXCLUIDO)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then Error 59242
    
    'Lê o crédito a pagar, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 59243
    
    'Achou o Crédito Pagar no BD
    If lErro = AD_SQL_SUCESSO Then
    
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CREDITO_PAGAR_NUMERO_REPETIDO", objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo, objCreditoPagar.dtDataEmissao)
        If vbMsgRes = vbNo Then Error 59244
        
    End If
            
    Call Comando_Fechar(lComando)
    
    CreditoPagar_Testa_Repeticao = SUCESSO
     
    Exit Function
    
Erro_CreditoPagar_Testa_Repeticao:

    CreditoPagar_Testa_Repeticao = Err
     
    Select Case Err
          
        Case 59241
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
        
        Case 59242, 59243
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CREDITOSPAGFORN", Err, objCreditoPagar.lFornecedor, objCreditoPagar.iFilial, objCreditoPagar.sSiglaDocumento, objCreditoPagar.lNumTitulo, objCreditoPagar.dtDataEmissao)
        
        Case 59244 'desistiu
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148967)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function DebitoRecCli_Testa_Repeticao(ByVal objDebitoRecCli As ClassDebitoRecCli) As Long
'Pede confirmacao do usuario se tentar inserir novo registro que possivelmente tenha sido entrado antes
'Chamada para tentar evitar entradas em duplicidade
'Obs.: estou desconsiderando a sigla do docto e a filial da empresa

Dim lErro As Long, lComando As Long, dtDataEmissao As Date
Dim vbMsgRes As VbMsgBoxResult
Dim sSiglaDoc As String

On Error GoTo Erro_DebitoRecCli_Testa_Repeticao

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 59246
    
    sSiglaDoc = String(255, 0)
    
    'Verifica se Data Emissão está preenchida
    If objDebitoRecCli.dtDataEmissao <> DATA_NULA Then
        
        'Tenta docto de mesmo cliente e numero com data emissão dentro da Faixa ou DATA_NULA
        lErro = Comando_Executar(lComando, "SELECT DataEmissao, SiglaDocumento FROM DebitosRecCli WHERE Cliente = ? AND Filial = ? AND NumTitulo = ? AND (( DataEmissao > ? AND DataEmissao < ? ) OR DataEmissao = ? ) AND Status <> ?", _
            dtDataEmissao, sSiglaDoc, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, objDebitoRecCli.lNumTitulo, (objDebitoRecCli.dtDataEmissao - PERIODO_EMISSAO), (objDebitoRecCli.dtDataEmissao + PERIODO_EMISSAO), DATA_NULA, STATUS_EXCLUIDO)
        
    Else
        
        'Tenta ler docto de mesmo cliente e numero independente de data emissão
        lErro = Comando_Executar(lComando, "SELECT DataEmissao, SiglaDocumento FROM DebitosRecCli WHERE Cliente = ? AND Filial = ? AND NumTitulo = ? AND Status <> ?", _
            dtDataEmissao, sSiglaDoc, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, objDebitoRecCli.lNumTitulo, STATUS_EXCLUIDO)
    
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 59247
    
    'Lê o docto, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 59248
    
'    'Achou o docto no BD
'    If lErro = AD_SQL_SUCESSO Then gError 108790
            
    'Se encontrou
    If lErro = AD_SQL_SUCESSO Then
        
        'Pergunta se deseja cadastrar uma Fatura com o mesmo número
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DEBITORECCLI_JA_EXISTENTE", objDebitoRecCli.lNumTitulo, objDebitoRecCli.lCliente, objDebitoRecCli.iFilial, sSiglaDoc, dtDataEmissao)
        
        'Se não, erro
        If vbMsgRes = vbNo Then gError 108790
    
    End If
            
    Call Comando_Fechar(lComando)
    
    DebitoRecCli_Testa_Repeticao = SUCESSO
     
    Exit Function
    
Erro_DebitoRecCli_Testa_Repeticao:

    DebitoRecCli_Testa_Repeticao = gErr
     
    Select Case gErr
          
        Case 59246
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 59247, 59248
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DEBITOSRECCLI", gErr)
        
        Case 59249 'desistiu
        
        Case 108790
'            Call Rotina_Erro(vbOKOnly, "ERRO_DEBCLI_INSERCAO_MESMO_NUMERO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148968)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
End Function

Function TituloReceber_Exclui_EmTransacao(ByVal objTituloReceber As ClassTituloReceber, ByVal objContabil As ClassContabil) As Long
'Exclui o Título à Receber do Banco de Dados
'ANTENÇÃO: ESSA ROTINA DEVE SER CHAMADA DENTRO DE UMA TRANSACAO

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lNumIntDoc As Long, lNumIntNF As Long
Dim iClasse As Integer
Dim objNFiscal As New ClassNFiscal
Dim lNumTitulo As Long

On Error GoTo Erro_TituloReceber_Exclui_EmTransacao

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26444

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26445

    If objTituloReceber.iFilialEmpresa = 0 Then objTituloReceber.iFilialEmpresa = giFilialEmpresa
    
    If objTituloReceber.iExclusao = 1 Then

        'Seleciona o Título na Tabela de Títulos à Receber
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumTitulo FROM TitulosRec WHERE NumIntDoc = ? AND FilialEmpresa = ? AND SiglaDocumento = ? ", 0, lNumIntDoc, lNumTitulo, objTituloReceber.lNumIntDoc, objTituloReceber.iFilialEmpresa, objTituloReceber.sSiglaDocumento)
        If lErro <> AD_SQL_SUCESSO Then gError 26446

    Else

        'Seleciona o Título na Tabela de Títulos à Receber
        lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumTitulo FROM TitulosRec WHERE NumIntDoc = ? AND FilialEmpresa = ? AND Status <> ? AND SiglaDocumento = ? ", 0, lNumIntDoc, lNumTitulo, objTituloReceber.lNumIntDoc, objTituloReceber.iFilialEmpresa, STATUS_EXCLUIDO, objTituloReceber.sSiglaDocumento)
        If lErro <> AD_SQL_SUCESSO Then gError 26446

    End If
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26447
    
    'Não achou o Título à Receber --> erro
    If lErro = AD_SQL_SEM_DADOS Then gError 26448

    'Faz um "lock" no Título
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 26449
    
    objTituloReceber.lNumTitulo = lNumTitulo
    
    iClasse = TITULO_RECEBER

    If objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER Then
    
        lErro = FaturaReceber_Desassocia_NFiscal(lNumIntDoc)
        If lErro <> SUCESSO Then gError 62442
        
    Else

        'Verifica se existe alguma NFiscal correspondente ao Título
        lErro = CF("DocumentoCPR_OrigemNFiscal", objTituloReceber.sSiglaDocumento, lNumIntDoc, lNumIntNF)
        If lErro <> SUCESSO And lErro <> 41542 Then gError 43062
        
        If lErro = SUCESSO Then
        
            objNFiscal.lNumIntDoc = lNumIntNF
    
            lErro = CF("NFiscal_Le", objNFiscal)
            If lErro <> SUCESSO And lErro <> 31442 Then gError 62138
            
            'se a nota não estiver cadastrada
            If lErro <> SUCESSO Then gError 62139
            
            'se não for uma exclusão e a nota não estiver cancelada
            If objTituloReceber.iExclusao = 0 And objNFiscal.iStatus <> STATUS_CANCELADO Then gError 43063
            
        End If
        
    End If
    
    'exclui as comissoes desde que nao estejam baixadas
    lErro = CF("Comissoes_Exclui", TIPO_TITULO_RECEBER, objTituloReceber.lNumIntDoc, 0)
    If lErro <> SUCESSO Then gError 26450

    'Exclui as parcelas referentes ao Título passado
    lErro = CF("ParcelasReceber_Exclui", objTituloReceber)
    If lErro <> SUCESSO Then gError 26451
    
    If objTituloReceber.iExclusao = 1 Then
    
        'Faz a exclusão do Titulo(EXCLUSAO FISICA)
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM TitulosRec", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 83951
    
    Else
    
        'Faz a exclusão do Titulo(EXCLUSAO LÓGICA)
        lErro = Comando_ExecutarPos(lComando1, "UPDATE TitulosRec SET Status = ? ", lComando, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 26452
    
    End If
    
    lErro = CF("Trata_Exclusao_Do_Destino", objTituloReceber)
    If lErro <> SUCESSO Then gError 192057
    
    '############################################
    'Inserido por Wagner 22/03/2006
    lErro = CF("ProjetoInfo_Exclui", PRJ_CR_TIPO_TITREC, objTituloReceber.lNumIntDoc)
    If lErro <> SUCESSO Then gError 181764
    '############################################
    
    If Not (objContabil Is Nothing) Then
        If objTituloReceber.iExclusao = 1 Then
            'Exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
            lErro = objContabil.Contabil_Exclui(objTituloReceber.lNumIntDoc, CONTABIL_EXCLUI_EFETIVO, , , , objTituloReceber.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83949
        Else
            'Exclui os lançamentos contábeis (se houverem) associados ao titulo em questao
            lErro = objContabil.Contabil_Exclui(objTituloReceber.lNumIntDoc, , , , objTituloReceber.dtDataEstorno, objTituloReceber.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 20522
        End If
    End If
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    TituloReceber_Exclui_EmTransacao = SUCESSO
    
    Exit Function

Erro_TituloReceber_Exclui_EmTransacao:

    TituloReceber_Exclui_EmTransacao = gErr
    
    Select Case gErr
    
        Case 20522, 26450, 26451, 43062, 62138, 62442, 83949, 181764, 192057
    
        Case 26443
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 26444, 26445
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 26446, 26447
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC", gErr, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)

        Case 26448
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", gErr, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)

        Case 26449
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_TITULOS_REC", gErr)

        Case 26452, 83951
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_TITULOS_RECEBER", gErr)

        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 43063
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_REL_NFISCAL", gErr)
        
        Case 62138
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148969)
            
    End Select
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function GeracaoFatura_Automatico_Trans(lNumeroFatura As Long) As Long
'Tem que ser chamada dentro de Transação
'Gera o Próximo Numero da Fatura

Dim lErro As Long

On Error GoTo Erro_GeracaoFatura_Automatico_Trans

    lErro = CF("CRFATConfig_ObterNumInt1", "FATURA_NUM_PROX", lNumeroFatura)
    If lErro <> SUCESSO Then Error 61465
    
    GeracaoFatura_Automatico_Trans = SUCESSO
    
    Exit Function

Erro_GeracaoFatura_Automatico_Trans:

    GeracaoFatura_Automatico_Trans = Err

    Select Case Err
    
        Case 61465
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148970)
    
    End Select
    
    Exit Function

End Function

'????? Fernando Está funcao vai para CRFAT Select
Function Transportadora_Le_Total(iTotalTransp As Integer) As Long
'Le a quantidade de Transportadoras já gravados na Tabela de Transportadoras

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_Transportadora_Le_Total

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 64071

    'Seleciona a quantidade de Transportadoras Gravados
    lErro = Comando_Executar(lComando, "SELECT COUNT(*) FROM Transportadoras", iTotalTransp)
    If lErro <> AD_SQL_SUCESSO Then Error 64072
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 64073
            
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Transportadora_Le_Total = SUCESSO
    
    Exit Function
    
Erro_Transportadora_Le_Total:
    
    Transportadora_Le_Total = Err
    
    Select Case Err
    
        Case 64071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 64072, 64073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA2", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148971)
            
    End Select
        
    'Fecha comando
    Call Comando_Fechar(lComando)
        
    Exit Function
        
End Function

Function CRFATConfig_Grava(ByVal sCodigo As String, ByVal iFilialEmpresa As Integer, ByVal sConteudo As String) As Long
'grava na tabela CRFATConfig o Conteudo passado

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim objCRFATConfig As New ClassCRFATConfig
Dim iIndice As Integer

On Error GoTo Erro_CRFATConfig_Grava
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 61456

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 61457
    Next
    
    objCRFATConfig.iFilialEmpresa = iFilialEmpresa
    objCRFATConfig.sCodigo = sCodigo
    objCRFATConfig.sConteudo = sConteudo
    
    lErro = CRFATConfig_Grava_Trans(objCRFATConfig, alComando())
    If lErro <> SUCESSO Then Error 61458
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 61459
          
    CRFATConfig_Grava = SUCESSO
    
    Exit Function
    
Erro_CRFATConfig_Grava:
    
    CRFATConfig_Grava = Err
    
    Select Case Err
    
        Case 61456
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 61457
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 61458 'Tratado na Rotina chamada

        Case 61459
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148972)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function FaturaReceber_Desassocia_NFiscal(ByVal lNumIntFatura As Long) As Long
'Faz a desassociacao da fatura com as notas fiscais

Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim lErro As Long
Dim lNumNotaFiscal As Long

On Error GoTo Erro_FaturaReceber_Desassocia_NFiscal
    
    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then Error 62443
    Next
    
    'Busca as NFs associadas a Fatura
    lErro = Comando_ExecutarPos(alComando(0), "SELECT NumNotaFiscal FROM NFiscal WHERE ClasseDocCPR = ? AND NumIntDocCPR = ?", 0, lNumNotaFiscal, CPR_TITULO_RECEBER, lNumIntFatura)
    If lErro <> AD_SQL_SUCESSO Then Error 62444
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62445
    
    'Para cada NF encontrada
    Do While lErro <> AD_SQL_SEM_DADOS
        'Desassocia a NF
        lErro = Comando_ExecutarPos(alComando(1), "UPDATE NFiscal SET ClasseDocCPR = 0, NumIntDocCPR = 0", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then Error 62447
        
        'BUsca a próxima
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 62446
               
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    FaturaReceber_Desassocia_NFiscal = SUCESSO

    Exit Function
    
Erro_FaturaReceber_Desassocia_NFiscal:

    FaturaReceber_Desassocia_NFiscal = Err
    
    Select Case Err
            
        Case 62443
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)
            
        Case 62444, 62445, 62446
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL3", Err)
        
        Case 62447
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL1", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148973)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Function NFiscalNumAuto(ByVal objNFiscal As ClassNFiscal) As Long
'Lê o Proximo número na tabela de Série e Coloca no objNFiscal
'Faz Lock Exclusive e atualiza o Número na Tabela de Série

Dim lErro As Long
Dim tSerie As typeSerie
Dim lComando As Long
Dim lComando1 As Long

On Error GoTo Erro_NFiscalNumAuto

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 60436

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 60437

    lErro = Comando_ExecutarPos(lComando, "SELECT ProxNumNFiscal FROM Serie WHERE Serie = ? AND FilialEmpresa = ?", 0, tSerie.lProxNumNFiscal, objNFiscal.sSerie, objNFiscal.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 60438

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60439
    
    lErro = Comando_LockExclusive(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 60440
    
    objNFiscal.lNumNotaFiscal = tSerie.lProxNumNFiscal
    
    tSerie.lProxNumNFiscal = tSerie.lProxNumNFiscal + 1
    
    lErro = Comando_ExecutarPos(lComando1, "UPDATE Serie SET ProxNumNFiscal = ?", lComando, tSerie.lProxNumNFiscal)
    If lErro <> AD_SQL_SUCESSO Then gError 60441
    
    lErro = CF("NFiscal_Testa_Numero", objNFiscal)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    NFiscalNumAuto = SUCESSO
    
    Exit Function
    
Erro_NFiscalNumAuto:

    NFiscalNumAuto = gErr
    
    Select Case gErr
    
        Case 60436, 60437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 60438, 60439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SERIE1", gErr, objNFiscal.sSerie)

        Case 60440
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SERIE", gErr)
        
        Case 60441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SERIE", gErr, objNFiscal.sSerie)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148974)
            
    End Select
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function
    
End Function

Function Importa_Clientes() As Long
'Importa os dados da tabela ImportCli para criar clientes e suas filiais
'obs.: se o cliente nao tem filial=1 troca a primeira filial encontrada para filial1

Dim lErro As Long
Dim tImportCli As typeImportCli
Dim lComando As Long
Dim lTransacao As Long
Dim objFilialCliente As New ClassFilialCliente
Dim objCliente As New ClassCliente
Dim colEndereco As New Collection
Dim lCodigo As Long, lClienteAnterior As Long

On Error GoTo Erro_Importa_Clientes

    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 76329
    
    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 76342
    
    tImportCli.sBairro = String(STRING_BAIRRO, 0)
    tImportCli.sBairro1 = String(STRING_BAIRRO, 0)
    tImportCli.sBairro2 = String(STRING_BAIRRO, 0)
    tImportCli.sCEP = String(STRING_CEP, 0)
    tImportCli.sCEP1 = String(STRING_CEP, 0)
    tImportCli.sCEP2 = String(STRING_CEP, 0)
    tImportCli.sCidade = String(25, 0)
    tImportCli.sCidade1 = String(25, 0)
    tImportCli.sCidade2 = String(25, 0)
    tImportCli.sContato = String(STRING_CONTATO, 0)
    tImportCli.sContato1 = String(STRING_CONTATO, 0)
    tImportCli.sContato2 = String(STRING_CONTATO, 0)
    tImportCli.sEmail = String(STRING_EMAIL, 0)
    tImportCli.sEmail1 = String(STRING_EMAIL, 0)
    tImportCli.sEmail2 = String(STRING_EMAIL, 0)
    tImportCli.sEndereco = String(STRING_ENDERECO, 0)
    tImportCli.sEndereco1 = String(STRING_ENDERECO, 0)
    tImportCli.sEndereco2 = String(STRING_ENDERECO, 0)
    tImportCli.sFax = String(STRING_FAX, 0)
    tImportCli.sFax1 = String(STRING_FAX, 0)
    tImportCli.sFax2 = String(STRING_FAX, 0)
    tImportCli.sFilialCGC = String(STRING_CGC, 0)
    tImportCli.sFilialContaContabil = String(STRING_CONTA, 0)
    tImportCli.sFilialInscEstadual = String(STRING_INSCR_EST, 0)
    tImportCli.sFilialInscMunicipal = String(STRING_INSCR_MUN, 0)
    tImportCli.sFilialNome = String(STRING_FILIAL_NOME, 0)
    tImportCli.sFilialObservacao1 = String(STRING_CLIENTE_OBSERVACAO, 0)
    tImportCli.sFilialTipoFrete = String(1, 0)
    tImportCli.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tImportCli.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tImportCli.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tImportCli.sSiglaEstado = String(STRING_ESTADO_SIGLA, 0)
    tImportCli.sSiglaEstado1 = String(STRING_ESTADO_SIGLA, 0)
    tImportCli.sSiglaEstado2 = String(STRING_ESTADO_SIGLA, 0)
    tImportCli.sTelefone1 = String(25, 0)
    tImportCli.sTelefone11 = String(25, 0)
    tImportCli.sTelefone12 = String(25, 0)
    tImportCli.sTelefone2 = String(25, 0)
    tImportCli.sTelefone21 = String(25, 0)
    tImportCli.sTelefone22 = String(25, 0)
    
    lClienteAnterior = -1
    
    'Lê os registros da tabela ImportCli
    With tImportCli
    lErro = Comando_Executar(lComando, "SELECT Codigo,RazaoSocial,NomeReduzido,Tipo,Observacao,LimiteCredito,CondicaoPagto,Desconto,CodPadraoCobranca,CodMensagem,TabelaPreco,NumPagamentos,CodFilial," _
        & "FilialNome,FilialCGC,FilialInscricaoEstadual,FilialInscricaoMunicipal,FilialCodTransportadora,FilialObservacao1,FilialContaContabil,FilialVendedor,FilialComissaoVendas,FilialRegiao,FilialFreqVisitas,FilialDataUltVisita,FilialCodCobrador," _
        & "FilialICMSBaseCalculoComIPI,FilialRevendedor,FilialTipoFrete,Endereco,Bairro,Cidade,SiglaEstado,CodigoPais,CEP,Telefone1,Telefone2,Email,Fax,Contato,Endereco1,Bairro1,Cidade1,SiglaEstado1,CodigoPais1,CEP1,Telefone11,Telefone21,Email1,Fax1,Contato1," _
        & "Endereco2,Bairro2,Cidade2,SiglaEstado2,CodigoPais2,CEP2,Telefone12,Telefone22,Email2,Fax2,Contato2 FROM ImportCli ORDER BY Codigo,CodFilial", .lCodCliente, .sRazaoSocial, .sNomeReduzido, .iTipo, .sObservacao, .dLimiteCredito, _
        .iCondicaoPagto, .dDesconto, .iCodPadraoCobranca, .iCodMensagem, .iTabelaPreco, .lNumPagamentos, .iCodFilial, .sFilialNome, .sFilialCGC, .sFilialInscEstadual, .sFilialInscMunicipal, _
        .iFilialCodTransportadora, .sFilialObservacao1, .sFilialContaContabil, .iFilialVendedor, .dFilialComissaoVendas, .iFilialRegiao, .iFilialFreqVisitas, .dtFilialDataUltVisita, .iFilialCodCobrador, .iFilialICMSBaseCalculoIPI, _
        .lFilialRevendedor, .sFilialTipoFrete, .sEndereco, .sBairro, .sCidade, .sSiglaEstado, .iCodigoPais, tImportCli.sCEP, .sTelefone1, .sTelefone2, .sEmail, .sFax, .sContato, .sEndereco1, .sBairro1, _
        .sCidade1, .sSiglaEstado1, .iCodigoPais1, .sCEP1, .sTelefone11, .sTelefone21, .sEmail1, .sFax1, tImportCli.sContato1, .sEndereco2, .sBairro2, .sCidade2, .sSiglaEstado2, .iCodigoPais2, .sCEP2, .sTelefone12, _
        .sTelefone22, .sEmail2, .sFax2, .sContato2)
    End With
    If lErro <> AD_SQL_SUCESSO Then gError 76330
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76331
    
    Do While lErro = AD_SQL_SUCESSO
    
        objFilialCliente.lCodCliente = tImportCli.lCodCliente
        objFilialCliente.iCodFilial = tImportCli.iCodFilial
        
        'Verifica se já existe a Filial do Cliente lido na tabela FiliaisClientes
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 76332
    
        'Se não existe a filial na tabela FiliaisClientes
        If lErro = 12567 Then
        
            If lClienteAnterior <> objFilialCliente.lCodCliente Then
            
                objFilialCliente.iCodFilial = FILIAL_MATRIZ
                
            End If
            
            Set colEndereco = New Collection
                
            If objFilialCliente.iCodFilial = FILIAL_MATRIZ Then
            
                'Preenche objCliente e colEndereco a partir de tImportCli
                lErro = Cliente_PreencheObjetosImportacao(tImportCli, objCliente, colEndereco)
                If lErro <> SUCESSO Then gError 76333
                
                'Grava o Cliente
                lErro = CF("Cliente_Grava_EmTrans", objCliente, colEndereco)
                If lErro <> SUCESSO Then gError 76334

            Else
        
                'Preenche objFilialCliente e colEndereco a partir de tImportCli
                lErro = FilialCliente_PreencheObjetosImportacao(tImportCli, objFilialCliente, colEndereco)
                If lErro <> SUCESSO Then gError 76335
                
                'Grava a Filial Cliente
                lErro = FiliaisClientes_Grava_EmTrans(objFilialCliente, colEndereco)
                If lErro <> SUCESSO Then gError 76336
            
            End If
        
        End If
        
        lClienteAnterior = objFilialCliente.lCodCliente
        
        'Busca o proximo registro de ImportCli
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 76341
        
    Loop
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 76343
    
    Call Comando_Fechar(lComando)

    Importa_Clientes = SUCESSO
    
    Exit Function
    
Erro_Importa_Clientes:

    Importa_Clientes = gErr
    
    Select Case gErr
    
        Case 76329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 76330, 76331, 76341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_IMPORTCLI", gErr)
        
        Case 76332, 76333, 76334, 76335, 76336
            'Erros tratados nas rotinas chamadas
            
        Case 76342
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 76343
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148975)
            
    End Select
    
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

Private Function Cliente_PreencheObjetosImportacao(tImportCli As typeImportCli, ByVal objCliente As ClassCliente, ByVal colEndereco As Collection) As Long
'Preenche objCliente e colEndereco a partir dos dados existentes em tImportCli

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Cliente_PreencheObjetosImportacao
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        objCliente.lCodigoLoja = tImportCli.lCodCliente
    Else
        objCliente.lCodigo = tImportCli.lCodCliente
    End If
    objCliente.sRazaoSocial = tImportCli.sRazaoSocial
    objCliente.sNomeReduzido = tImportCli.sNomeReduzido
    
    'Todo cliente está recebendo Tipo=1
    objCliente.iTipo = 1
    
    objCliente.sObservacao = tImportCli.sObservacao
    objCliente.dLimiteCredito = tImportCli.dLimiteCredito
    objCliente.dDesconto = tImportCli.dDesconto
    objCliente.iTabelaPreco = tImportCli.iTabelaPreco
    
    'nao incluir cnodicao de pagto
    objCliente.iCondicaoPagto = tImportCli.iCondicaoPagto
    
    objCliente.iCodMensagem = tImportCli.iCodMensagem
    objCliente.lNumPagamentos = tImportCli.lNumPagamentos
    objCliente.iCodPadraoCobranca = tImportCli.iCodPadraoCobranca
    objCliente.iProxCodFilial = 2 '???
    
    lErro = FilialCliente_PreencheInfoImportacao(objCliente, tImportCli)
    If lErro <> SUCESSO Then gError 76337
    
    lErro = FilialCliente_PreencheEnderecosImportacao(tImportCli, colEndereco)
    If lErro <> SUCESSO Then gError 76338
        
    Cliente_PreencheObjetosImportacao = SUCESSO
    
    Exit Function
    
Erro_Cliente_PreencheObjetosImportacao:

    Cliente_PreencheObjetosImportacao = gErr
    
    Select Case gErr
    
        Case 76337, 76338
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148976)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheObjetosImportacao(tImportCli As typeImportCli, ByVal objFilialCliente As ClassFilialCliente, ByVal colEndereco As Collection) As Long
'Preenche objFilialCliente e colEndereco a partir dos dados existentes em tImportCli

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_FilialCliente_PreencheObjetosImportacao

    objFilialCliente.iCodFilial = tImportCli.iCodFilial
    objFilialCliente.sNome = tImportCli.sFilialNome
    objFilialCliente.lCodCliente = tImportCli.lCodCliente
    
    lErro = FilialCliente_PreencheInfoImportacao(objFilialCliente, tImportCli)
    If lErro <> SUCESSO Then gError 76339
    
    lErro = FilialCliente_PreencheEnderecosImportacao(tImportCli, colEndereco)
    If lErro <> SUCESSO Then gError 76340
    
    FilialCliente_PreencheObjetosImportacao = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_PreencheObjetosImportacao:

    FilialCliente_PreencheObjetosImportacao = gErr
    
    Select Case gErr
    
        Case 76339, 76340
            'Erros tratados nas rotinas chamadas
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148977)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheEnderecosImportacao(tImportCli As typeImportCli, ByVal colEndereco As Collection) As Long
'Preenche colEndereco a partir dos dados existentes em tImportCli

Dim lErro As Long
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_FilialCliente_PreencheEnderecosImportacao

    'Preenche colEndereco
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais
    objEndereco.sBairro = tImportCli.sBairro
    objEndereco.sCEP = tImportCli.sCEP
    
    'Verifica se o tamanho da string Cidade lida é maior que o permitido
    If Len(Trim(tImportCli.sCidade)) > 15 Then
        objEndereco.sCidade = "***"
    Else
        objEndereco.sCidade = tImportCli.sCidade
    End If
    
    objEndereco.sContato = tImportCli.sContato
    objEndereco.sEmail = tImportCli.sEmail
    objEndereco.sEndereco = tImportCli.sEndereco
    objEndereco.sFax = tImportCli.sFax
    objEndereco.sSiglaEstado = tImportCli.sSiglaEstado
    
    'Verifica se o tamanho da string Telefone lida é maior que o permitido
    If Len(Trim(tImportCli.sTelefone1)) > 18 Then
        objEndereco.sTelefone1 = "***"
    Else
        objEndereco.sTelefone1 = tImportCli.sTelefone1
    End If
    
    If Len(Trim(tImportCli.sTelefone2)) > 18 Then
        objEndereco.sTelefone2 = "***"
    Else
        objEndereco.sTelefone2 = tImportCli.sTelefone2
    End If
    
    colEndereco.Add objEndereco
    
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais1
    objEndereco.sBairro = tImportCli.sBairro1
    objEndereco.sCEP = tImportCli.sCEP1
    
    If Len(Trim(tImportCli.sCidade1)) > 15 Then
        objEndereco.sCidade = "***"
    Else
        objEndereco.sCidade = tImportCli.sCidade1
    End If
    
    objEndereco.sContato = tImportCli.sContato1
    objEndereco.sEmail = tImportCli.sEmail1
    objEndereco.sEndereco = tImportCli.sEndereco1
    objEndereco.sFax = tImportCli.sFax1
    objEndereco.sSiglaEstado = tImportCli.sSiglaEstado1
    
    If Len(Trim(tImportCli.sTelefone11)) > 18 Then
        objEndereco.sTelefone1 = "***"
    Else
        objEndereco.sTelefone1 = tImportCli.sTelefone11
    End If
    
    If Len(Trim(tImportCli.sTelefone21)) > 18 Then
        objEndereco.sTelefone2 = "***"
    Else
        objEndereco.sTelefone2 = tImportCli.sTelefone21
    End If
    
    colEndereco.Add objEndereco
    
    Set objEndereco = New ClassEndereco
    
    objEndereco.iCodigoPais = tImportCli.iCodigoPais2
    objEndereco.sBairro = tImportCli.sBairro2
    objEndereco.sCEP = tImportCli.sCEP2
    
    If Len(Trim(tImportCli.sCidade2)) > 15 Then
        objEndereco.sCidade = "***"
    Else
        objEndereco.sCidade = tImportCli.sCidade2
    End If
    
    objEndereco.sContato = tImportCli.sContato2
    objEndereco.sEmail = tImportCli.sEmail2
    objEndereco.sEndereco = tImportCli.sEndereco2
    objEndereco.sFax = tImportCli.sFax2
    objEndereco.sSiglaEstado = tImportCli.sSiglaEstado2
    
    If Len(Trim(tImportCli.sTelefone12)) > 18 Then
        objEndereco.sTelefone1 = "***"
    Else
        objEndereco.sTelefone1 = tImportCli.sTelefone12
    End If
    
    If Len(Trim(tImportCli.sTelefone22)) > 18 Then
        objEndereco.sTelefone2 = "***"
    Else
        objEndereco.sTelefone2 = tImportCli.sTelefone22
    End If
    
    colEndereco.Add objEndereco
    
    FilialCliente_PreencheEnderecosImportacao = SUCESSO
    
    Exit Function
    
Erro_FilialCliente_PreencheEnderecosImportacao:

    FilialCliente_PreencheEnderecosImportacao = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148978)
            
    End Select
    
    Exit Function
    
End Function

Private Function FilialCliente_PreencheInfoImportacao(ByVal objFilialCliente As Object, tImportCli As typeImportCli) As Long

Dim lErro As Long
Dim objRegiaoVenda As New ClassRegiaoVenda

On Error GoTo Erro_FilialCliente_PreencheInfoImportacao

    If Len(Trim(tImportCli.sFilialCGC)) > 11 And Len(Trim(tImportCli.sFilialCGC)) <> 14 Then
        objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000000")
    Else
        If Len(Trim(tImportCli.sFilialCGC)) > 8 And Len(Trim(tImportCli.sFilialCGC)) <> 11 Then
            objFilialCliente.sCgc = Format(Trim(tImportCli.sFilialCGC), "00000000000")
        Else
            objFilialCliente.sCgc = tImportCli.sFilialCGC
        End If
    End If
        
    objFilialCliente.sInscricaoEstadual = tImportCli.sFilialInscEstadual
    objFilialCliente.sInscricaoMunicipal = tImportCli.sFilialInscMunicipal
    objFilialCliente.sObservacao = tImportCli.sObservacao
    objFilialCliente.iVendedor = tImportCli.iFilialVendedor
    objFilialCliente.iRegiao = tImportCli.iFilialRegiao
    objFilialCliente.sContaContabil = tImportCli.sFilialContaContabil
    objFilialCliente.dComissaoVendas = tImportCli.dFilialComissaoVendas
    objFilialCliente.iCodCobrador = tImportCli.iFilialCodCobrador
    objFilialCliente.iFreqVisitas = tImportCli.iFilialFreqVisitas
    objFilialCliente.dtDataUltVisita = tImportCli.dtFilialDataUltVisita
    objFilialCliente.iCodTransportadora = tImportCli.iFilialCodTransportadora
    
    If tImportCli.sFilialTipoFrete = "F" Then
        objFilialCliente.iTipoFrete = TIPO_FOB
    ElseIf tImportCli.sFilialTipoFrete = "C" Then
        objFilialCliente.iTipoFrete = TIPO_CIF
    End If
    
    FilialCliente_PreencheInfoImportacao = SUCESSO
     
    Exit Function
    
Erro_FilialCliente_PreencheInfoImportacao:

    FilialCliente_PreencheInfoImportacao = gErr
     
    Select Case Err
          
        Case 76339, 76340, 76385
            'Erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148979)
     
    End Select
     
    Exit Function

End Function

Function FiliaisClientes_Grava_EmTrans(ByVal objFilialCliente As ClassFilialCliente, ByVal colEndereco As Collection) As Long
'Grava a Filial Cliente e seus 3 tipos de Enderecos

Dim lErro As Long
Dim iIndice As Integer
Dim lComando1 As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lComando4 As Long
Dim lComando5 As Long
Dim lComando6 As Long
Dim lCodigo As Long
Dim lCodEndereco As Long
Dim tFilialCliente As typeFilialCliente
Dim objEndereco As ClassEndereco
Dim lEndereco(1 To 3) As Long
Dim colCategoriaItemCopia As New Collection
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objClienteHistorico As New ClassClienteHistorico
Dim objPais As New ClassPais
Dim objCliente As New ClassCliente
Dim objLog As New ClassLog
Dim iAtivo As Integer
Dim lCodigoCliente As Long
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_FiliaisClientes_Grava_EmTrans

    tFilialCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialCliente.sCgc = String(STRING_CGC, 0)
    tFilialCliente.sRG = String(STRING_RG, 0)
    tFilialCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialCliente.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tFilialCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialCliente.sContaContabil = String(STRING_CONTA, 0)

    'Inicializa comandos
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 12621

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 12622

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 12623
    
    lComando4 = Comando_Abrir()
    If lComando4 = 0 Then gError 112734
    
    lComando5 = Comando_Abrir()
    If lComando5 = 0 Then gError 112735
    
    lComando6 = Comando_Abrir()
    If lComando6 = 0 Then gError 112744

    'Cria uma cópia "de trabalho" da coleção passada como parâmetro
    For Each objFilialCliCategoria In objFilialCliente.colCategoriaItem
        colCategoriaItemCopia.Add objFilialCliCategoria
    Next

    'Faz os locks nas tabelas ligadas ao cadastramento de uma nova Filial Cliente
    lErro = CF("FiliaisClientes_Le_Tabelas_Lock", objFilialCliente)
    If lErro <> SUCESSO Then gError 12630

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialCliente.iCodFilialLoja <> 0 Then

        If objFilialCliente.lCodCliente <> 0 Then

            'Procura por um outro Nome já cadastrado
            lErro = Comando_Executar(lComando1, "SELECT CodCliente FROM FiliaisClientes WHERE CodCliente = ? AND CodFilialLoja <> ? AND Nome = ? ", lCodigo, objFilialCliente.lCodCliente, objFilialCliente.iCodFilialLoja, objFilialCliente.sNome)
            If lErro <> AD_SQL_SUCESSO Then gError 117610

        Else
        
            'Procura por um outro Nome já cadastrado
            lErro = Comando_Executar(lComando1, "SELECT CodCliente FROM FiliaisClientes WHERE CodClienteLoja = ? AND CodFilialLoja <> ? AND Nome = ? ", lCodigo, objFilialCliente.lCodClienteLoja, objFilialCliente.iCodFilialLoja, objFilialCliente.sNome)
            If lErro <> AD_SQL_SUCESSO Then gError 12625
    
        End If
    
    Else
    
        'Procura por um outro Nome já cadastrado
        lErro = Comando_Executar(lComando1, "SELECT CodCliente FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial <> ? AND Nome = ? ", lCodigo, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, objFilialCliente.sNome)
        If lErro <> AD_SQL_SUCESSO Then gError 117613

    End If

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12626

    'Se existir outro Nome acusa erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 12627
    
    'Atualiza os dados para futura gravacao de histórico
    With objClienteHistorico
    
        .dtDataAtualizacao = gdtDataHoje
        .iFilialCliente = objFilialCliente.iCodFilial
        .lCodigo = objFilialCliente.lCodCliente
        .sCgc = objFilialCliente.sCgc
        .sInscricaoEstadual = objFilialCliente.sInscricaoEstadual
        .sInscricaoMunicipal = objFilialCliente.sInscricaoMunicipal
        .sInscricaoSuframa = objFilialCliente.sInscricaoSuframa
        .iAtivo = objFilialCliente.iAtivo
        .lCodigoLoja = objFilialCliente.lCodClienteLoja
        .iFilialEmpresaLoja = objFilialCliente.iFilialEmpresaLoja
        .iCodFilialLoja = objFilialCliente.iCodFilialLoja
        
        objCliente.lCodigoLoja = objFilialCliente.lCodClienteLoja
        objCliente.lCodigo = objFilialCliente.lCodCliente
            
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 97620
        
        'Verifica se Cliente nao esta cadastrado
        If lErro = 12293 Then gError 97621
        
        .sRazaoSocial = objCliente.sRazaoSocial
        
        .sNomeReduzido = objCliente.sNomeReduzido
        .iTipo = objFilialCliente.iTipo
        .dLimiteCredito = objFilialCliente.dLimiteCredito
        .sObservacao = objCliente.sObservacao
        .iCondicaoPagto = objCliente.iCondicaoPagto
        .iTabelaPreco = objCliente.iTabelaPreco
        .sUsuarioCobrador = objCliente.sUsuarioCobrador
        .sUsuRespCallCenter = objCliente.sUsuRespCallCenter
        .iCodTransportadora = objFilialCliente.iCodTransportadora
        .iVendedor = objFilialCliente.iVendedor
        .sRG = objFilialCliente.sRG
        .iRegimeTributario = objFilialCliente.iRegimeTributario
        .dDesconto = objCliente.dDesconto
        .dComissaoVendas = objFilialCliente.dComissaoVendas
        .iRegiao = objFilialCliente.iRegiao
        
        If objFilialCliente.iCodMensagem <> 0 Then
            objMensagem.iCodigo = objFilialCliente.iCodMensagem
            'Lê a Mensagem "default" p/ o Cliente
            lErro = CF("Mensagem_Le", objMensagem)
            If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
            
            .sMensagemNF = objMensagem.sDescricao
        End If
        
        If colEndereco.Count > 0 Then
        
            lErro = CF("Endereco_Ajusta_Dados", colEndereco.Item(1))
            If lErro <> SUCESSO And lErro <> 47876 Then gError 97622
        
            .sBairro = colEndereco.Item(1).sBairro
            .sCEP = colEndereco.Item(1).sCEP
            .sCidade = colEndereco.Item(1).sCidade
            .sEndereco = colEndereco.Item(1).sEndereco
            .sUF = colEndereco.Item(1).sSiglaEstado
            
            .sTelefone1 = colEndereco.Item(1).sTelefone1
            .sTelefone2 = colEndereco.Item(1).sTelefone2
            .sEmail = colEndereco.Item(1).sEmail
            .sEmail2 = colEndereco.Item(1).sEmail2
            
            'Le o nome do Pais com o codigo no obj
            objPais.iCodigo = colEndereco.Item(1).iCodigoPais
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 97622
        
            'Pais não está cadastrado - --- -> Erro
            If lErro = 47876 Then gError 97623
            
            .sPais = objPais.sNome
            
        End If
        
    End With
    
    'FAZ A GRAVACAO DO HISTORICO
    lErro = CF("ClienteHistorico_Grava", objClienteHistorico)
    If lErro <> SUCESSO Then gError 97624
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialCliente.iCodFilialLoja <> 0 Then
    
        If objFilialCliente.lCodCliente <> 0 Then
    
            'Pesquisa a Filial no BD
            lErro = Comando_ExecutarPos(lComando2, "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisClientes WHERE CodCliente = ? AND CodFilialLoja = ? AND FilialEmpresaLoja = ?", 0, _
                tFilialCliente.iCodFilial, tFilialCliente.sNome, tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, tFilialCliente.iCodTransportadora, tFilialCliente.sObservacao, tFilialCliente.sContaContabil, tFilialCliente.iVendedor, tFilialCliente.dComissaoVendas, tFilialCliente.iRegiao, tFilialCliente.iFreqVisitas, tFilialCliente.dtDataUltVisita, tFilialCliente.iCodCobrador, tFilialCliente.iTipoFrete, tFilialCliente.sInscricaoSuframa, objFilialCliente.lCodCliente, objFilialCliente.iCodFilialLoja, objFilialCliente.iFilialEmpresaLoja)
            If lErro <> AD_SQL_SUCESSO Then gError 117611
    
        Else
    
            'Pesquisa a Filial no BD
            lErro = Comando_ExecutarPos(lComando2, "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisClientes WHERE CodClienteLoja = ? AND CodFilialLoja = ? AND FilialEmpresaLoja = ?", 0, _
                tFilialCliente.iCodFilial, tFilialCliente.sNome, tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, tFilialCliente.iCodTransportadora, tFilialCliente.sObservacao, tFilialCliente.sContaContabil, tFilialCliente.iVendedor, tFilialCliente.dComissaoVendas, tFilialCliente.iRegiao, tFilialCliente.iFreqVisitas, tFilialCliente.dtDataUltVisita, tFilialCliente.iCodCobrador, tFilialCliente.iTipoFrete, tFilialCliente.sInscricaoSuframa, objFilialCliente.lCodClienteLoja, objFilialCliente.iCodFilialLoja, objFilialCliente.iFilialEmpresaLoja)
            If lErro <> AD_SQL_SUCESSO Then gError 12628
    
        End If
    
    Else
    
        'Pesquisa a Filial no BD
        lErro = Comando_ExecutarPos(lComando2, "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", 0, _
            tFilialCliente.iCodFilial, tFilialCliente.sNome, tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, tFilialCliente.iCodTransportadora, tFilialCliente.sObservacao, tFilialCliente.sContaContabil, tFilialCliente.iVendedor, tFilialCliente.dComissaoVendas, tFilialCliente.iRegiao, tFilialCliente.iFreqVisitas, tFilialCliente.dtDataUltVisita, tFilialCliente.iCodCobrador, tFilialCliente.iTipoFrete, tFilialCliente.sInscricaoSuframa, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 117612

    End If
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando2)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 12629
    
    'Filial Cliente não está cadastrada - GRAVACAO
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Grava os Três Novos Enderecos de Filiais
        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Automatico", lCodEndereco)
            If lErro <> SUCESSO Then gError 12616

            colEndereco.Item(iIndice).lCodigo = lCodEndereco

            lEndereco(iIndice) = lCodEndereco

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 12617

        Next

        '########################################################
        'Alterado por Wagner 27/09/05
        'Inserido CodMensagem
        'Insere a FilialCliente na tabela FiliaisClientes
        lErro = Comando_Executar(lComando3, "INSERT INTO FiliaisClientes (CodCliente,CodFilial,Nome,CGC,InscricaoEstadual,InscricaoMunicipal,Endereco,EnderecoEntrega,EnderecoCobranca,CodTransportadora,Observacao,ContaContabil,Vendedor,ComissaoVendas,Regiao,FreqVisitas,DataUltVisita,CodCobrador,TipoFrete,InscricaoSuframa,RG, CodFilialLoja, FilialEmpresaLoja, CodClienteLoja, Ativo, CodTranspRedesp, DetPagFrete, Guia, CodMensagem, DataUltAtualizacao, Exportado, RegimeTributario, IEIsento, IENaoContrib, DataRegistro) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objFilialCliente.lCodCliente, objFilialCliente.iCodFilial, objFilialCliente.sNome, objFilialCliente.sCgc, objFilialCliente.sInscricaoEstadual, objFilialCliente.sInscricaoMunicipal, lEndereco(1), lEndereco(2), lEndereco(3), objFilialCliente.iCodTransportadora, objFilialCliente.sObservacao, objFilialCliente.sContaContabil, objFilialCliente.iVendedor, objFilialCliente.dComissaoVendas, objFilialCliente.iRegiao, objFilialCliente.iFreqVisitas, objFilialCliente.dtDataUltVisita, objFilialCliente.iCodCobrador, objFilialCliente.iTipoFrete, objFilialCliente.sInscricaoSuframa, objFilialCliente.sRG, objFilialCliente.iCodFilialLoja, objFilialCliente.iFilialEmpresaLoja, objFilialCliente.lCodClienteLoja, objFilialCliente.iAtivo, objFilialCliente.iCodTranspRedesp, objFilialCliente.iDetPagFrete, objFilialCliente.sGuia, objFilialCliente.iCodMensagem, Date, 0, objFilialCliente.iRegimeTributario, objFilialCliente.iIEIsento, objFilialCliente.iIENaoContrib, Date)
        If lErro <> AD_SQL_SUCESSO Then gError 12631
        '########################################################

        'Insere as categorias da filial do cliente em FilialClientesCategorias
        lErro = CF("FilialCliente_Insere_Categorias", objFilialCliente, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 33101

    Else 'FilialCliente já existe - MODIFICAÇÃO

        'Grava os Três Enderecos de Filiais
        
        If colEndereco.Count > 0 Then colEndereco.Item(1).lCodigo = tFilialCliente.lEndereco
        If colEndereco.Count > 1 Then colEndereco.Item(2).lCodigo = tFilialCliente.lEnderecoEntrega
        If colEndereco.Count > 2 Then colEndereco.Item(3).lCodigo = tFilialCliente.lEnderecoCobranca

        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 12887

        Next

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
    
            'Atualiza a tabela de FiliaisClientes
            lErro = Comando_ExecutarPos(lComando3, "UPDATE FiliaisClientes SET Nome = ?, CGC = ?, RG = ?, Ativo = ?, DataUltAtualizacao = ?, Exportado = ? ", lComando2, _
                objFilialCliente.sNome, objFilialCliente.sCgc, objFilialCliente.sRG, objFilialCliente.iAtivo, Date, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 12632
    
        Else
    
            '########################################################
            'Alterado por Wagner 27/09/05
            'Inserido CodMensagem
            'Atualiza a tabela de FiliaisClientes
            lErro = Comando_ExecutarPos(lComando3, "UPDATE FiliaisClientes SET Nome = ?, CGC = ?, InscricaoEstadual = ?, InscricaoMunicipal = ?, Endereco = ?, EnderecoEntrega = ?, EnderecoCobranca = ?, CodTransportadora = ?, Observacao = ?, ContaContabil = ?, Vendedor = ?, ComissaoVendas = ?, Regiao = ?, FreqVisitas = ?, DataUltVisita = ?, CodCobrador = ?, TipoFrete = ?,InscricaoSuframa=?, RG=?, Ativo=?, CodTranspRedesp=?, DetPagFrete=?, Guia=?, CodMensagem = ? , DataUltAtualizacao = ?, Exportado = ?, RegimeTributario = ?, IEIsento = ?, IENaoContrib = ? ", lComando2, _
                objFilialCliente.sNome, objFilialCliente.sCgc, objFilialCliente.sInscricaoEstadual, objFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, objFilialCliente.iCodTransportadora, objFilialCliente.sObservacao, objFilialCliente.sContaContabil, objFilialCliente.iVendedor, objFilialCliente.dComissaoVendas, objFilialCliente.iRegiao, objFilialCliente.iFreqVisitas, objFilialCliente.dtDataUltVisita, objFilialCliente.iCodCobrador, objFilialCliente.iTipoFrete, objFilialCliente.sInscricaoSuframa, objFilialCliente.sRG, objFilialCliente.iAtivo, objFilialCliente.iCodTranspRedesp, objFilialCliente.iDetPagFrete, objFilialCliente.sGuia, objFilialCliente.iCodMensagem, Date, 0, objFilialCliente.iRegimeTributario, objFilialCliente.iIEIsento, objFilialCliente.iIENaoContrib)
            If lErro <> AD_SQL_SUCESSO Then gError 12632
            '########################################################
    
        End If

        'Atualiza as categorias da filial do cliente
        lErro = CF("FilialCliente_Atualiza_Categorias", objFilialCliente, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 33102

    End If
    
    'se ativo uma filial--> ativo o cliente desta filial
    If objFilialCliente.iAtivo = CLIENTE_ATIVO Then
        'Seleciona o Clientes no BD q está desativado e ativa
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialCliente.lCodCliente = 0 Then
            lCodigoCliente = objFilialCliente.lCodClienteLoja
            lErro = Comando_ExecutarPos(lComando4, "SELECT Ativo FROM Clientes WHERE CodigoLoja = ? AND Ativo <> ?", 0, iAtivo, objFilialCliente.lCodClienteLoja, CLIENTE_ATIVO)
        Else
            lCodigoCliente = objFilialCliente.lCodCliente
            lErro = Comando_ExecutarPos(lComando4, "SELECT Ativo FROM Clientes WHERE Codigo = ? AND Ativo <> ?", 0, Ativo, objFilialCliente.lCodCliente, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 112736

        lErro = Comando_BuscarPrimeiro(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112737
        
        If lErro = AD_SQL_SUCESSO Then
            'Atualiza a tabela de Clientes
            lErro = Comando_ExecutarPos(lComando5, "UPDATE Clientes SET Ativo = ?", lComando4, objFilialCliente.iAtivo)
            If lErro <> AD_SQL_SUCESSO Then gError 112738
        End If
    Else
        'tenta achar outra filial q esteja ativa--> se não achar, desativa o cliente
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialCliente.lCodCliente = 0 Then
            lErro = Comando_Executar(lComando4, "SELECT Ativo FROM FiliaisClientes WHERE CodClienteLoja = ? AND Ativo = ?", 0, iAtivo, objFilialCliente.lCodClienteLoja, CLIENTE_ATIVO)
        Else
            lErro = Comando_ExecutarPos(lComando4, "SELECT Ativo FROM FiliaisClientes WHERE CodCliente = ? AND Ativo = ?", 0, Ativo, objFilialCliente.lCodCliente, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 112739

        lErro = Comando_BuscarPrimeiro(lComando4)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112740
        
        If lErro = AD_SQL_SEM_DADOS Then
            'Seleciona o Clientes no BD
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialCliente.lCodCliente = 0 Then
                lCodigoCliente = objFilialCliente.lCodClienteLoja
                lErro = Comando_ExecutarPos(lComando5, "SELECT Ativo FROM Clientes WHERE CodigoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialCliente.lCodClienteLoja, CLIENTE_ATIVO)
            Else
                lCodigoCliente = objFilialCliente.lCodCliente
                lErro = Comando_ExecutarPos(lComando5, "SELECT Ativo FROM Clientes WHERE Codigo = ? AND Ativo = ?", 0, Ativo, objFilialCliente.lCodCliente, CLIENTE_ATIVO)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 112741
    
            lErro = Comando_BuscarPrimeiro(lComando5)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112742
            
            If lErro = AD_SQL_SUCESSO Then
                'Atualiza a tabela de Clientes
                lErro = Comando_ExecutarPos(lComando6, "UPDATE Clientes SET Ativo = ?", lComando5, objFilialCliente.iAtivo)
                If lErro <> AD_SQL_SUCESSO Then gError 112743
            End If
        End If
    End If
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
    
        'Monta a String de log que deverá ser gravada.
        Call FilialCliente_Carrega_Log(objLog, objFilialCliente, colEndereco)
    
        'Monta Registro Para a Tabela de Log
        lErro = CF("Log_Grava", objLog)
        If lErro <> SUCESSO Then gError 113793
        
    End If

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    Call Comando_Fechar(lComando6)
    
    FiliaisClientes_Grava_EmTrans = SUCESSO

    Exit Function

Erro_FiliaisClientes_Grava_EmTrans:

    FiliaisClientes_Grava_EmTrans = gErr

    Select Case gErr

        Case 12621, 12622, 12623, 112734, 112735, 112744
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 12625, 12626, 12629, 112739, 112740, 117610, 117613
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES2", gErr)

        Case 12627
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NOME_DUPLICADO", gErr, objFilialCliente.sNome)

        Case 12628
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objFilialCliente.lCodClienteLoja, objFilialCliente.iCodFilialLoja)

        Case 12630, 12616, 12617, 12887, 33101, 33102

        Case 12631
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIAISCLIENTES", gErr)

        Case 12632
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISCLIENTES", gErr)
            
        Case 97622, 97620, 97624
        
        Case 97623
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, objPais.iCodigo)
            
        Case 97621
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objFilialCliente.lCodCliente)
    
        Case 112736, 112737, 112741, 112742
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)
                    
        Case 112738, 112743
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", gErr, lCodigoCliente)
            
        Case 117611
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilialLoja)
            
        Case 117612
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148980)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    Call Comando_Fechar(lComando4)
    Call Comando_Fechar(lComando5)
    Call Comando_Fechar(lComando6)
    
    Exit Function

End Function

Private Sub FilialCliente_Carrega_Log(ByVal objLog As ClassLog, ByVal objFilialCliente As ClassFilialCliente, ByVal colEndereco As Collection)

Dim objEndereco As ClassEndereco

    With objFilialCliente
    
        'Função para Carregar o objLog
        objLog.sLog = CStr(.lCodCliente) & Chr(vbKeyEscape) & CStr(.iCodFilial) & Chr(vbKeyEscape) & CStr(.lCodClienteLoja) & Chr(vbKeyEscape) & CStr(.iCodFilialLoja) & Chr(vbKeyEscape) & CStr(.sCgc) & Chr(vbKeyEscape) & CStr(.sNome) & Chr(vbKeyEscape)
        objLog.sLog = objLog.sLog & CStr(.sRG) & Chr(vbKeyEscape) & CStr(.iFilialEmpresaLoja) & Chr(vbKeyEscape) & CStr(.iAtivo) & Chr(vbKeyEscape)
    
    End With
    
    Set objEndereco = colEndereco.Item(1)
    
    With objEndereco
    
        'Função para Carregar o objLog
        objLog.sLog = objLog.sLog & CStr(.iCodigoPais) & Chr(vbKeyEscape) & CStr(.lCodigo) & Chr(vbKeyEscape) & CStr(.sBairro) & Chr(vbKeyEscape) & CStr(.sCEP) & Chr(vbKeyEscape) & CStr(.sCidade) & Chr(vbKeyEscape) & CStr(.sContato) & Chr(vbKeyEscape) & CStr(.sEmail) & Chr(vbKeyEscape) & CStr(.sEndereco) & Chr(vbKeyEscape) & CStr(.sFax) & Chr(vbKeyEscape) & CStr(.sSiglaEstado) & Chr(vbKeyEscape) & CStr(.sTelefone1) & Chr(vbKeyEscape) & CStr(.sTelefone2) & Chr(vbKeyEscape)
        
    End With
    
    objLog.sLog = objLog.sLog & Chr(vbKeyEnd)

    Exit Sub

End Sub

Function Cliente_Grava_EmTrans(ByVal objCliente As ClassCliente, ByVal colEndereco As Collection) As Long
'Grava o Cliente passado como parâmetro

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(6) As Long
Dim lCodigo As Long
Dim lCodEndereco As Long
Dim tCliente As typeCliente
Dim tFilialCliente As typeFilialCliente
Dim colCategoriaItemCopia As New Collection
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objFilialCliente As New ClassFilialCliente
Dim objClienteHistorico As New ClassClienteHistorico
Dim objPais As New ClassPais
Dim objEndereco As New ClassEndereco
Dim objLog As New ClassLog
Dim lCodigoLoja As Long
Dim lCodigoCliente As Long
Dim iAtivo As Integer
Dim bClienteNovo As Boolean
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Cliente_Grava_EmTrans

    bClienteNovo = False

    'Inicializa os atributos STRING dos types Cliente e FilialCliente
    tCliente.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tCliente.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialCliente.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialCliente.sCgc = String(STRING_CGC, 0)
    tFilialCliente.sRG = String(STRING_RG, 0)
    tFilialCliente.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialCliente.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialCliente.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tFilialCliente.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialCliente.sContaContabil = String(STRING_CONTA, 0)

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 113772
    Next
    
    'guarda o código do cliente para fins de msg de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        lCodigoCliente = objCliente.lCodigoLoja
    Else
        lCodigoCliente = objCliente.lCodigo
    End If

    'Cria uma cópia "de trabalho" da coleção passada como parâmetro
    For Each objFilialCliCategoria In objCliente.colCategoriaItem
        colCategoriaItemCopia.Add objFilialCliCategoria
    Next

    'se estiver operando no Caixa Central
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
        
        'Procura por um outro Cliente com o mesmo Nome Reduzido
        lErro = Comando_Executar(alComando(0), "SELECT CodigoLoja FROM Clientes WHERE CodigoLoja <> ? AND NomeReduzido = ? ", lCodigoLoja, objCliente.lCodigoLoja, objCliente.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 113771
    
    Else
    
        'Procura por um outro Cliente com o mesmo Nome Reduzido
        lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Clientes WHERE Codigo <> ? AND NomeReduzido = ? ", lCodigo, objCliente.lCodigo, objCliente.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 113773
    
    End If
    
    'Tenta ler o código deste Cliente
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113774

    'Se existir outro Cliente com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 113775
        
    'Atualiza os dados para futura gravacao de histórico
    With objClienteHistorico
    
        .dtDataAtualizacao = gdtDataHoje
        .iFilialCliente = FILIAL_MATRIZ
        .lCodigo = objCliente.lCodigo
        .sCgc = objCliente.sCgc
        .sInscricaoEstadual = objCliente.sInscricaoEstadual
        .sInscricaoMunicipal = objCliente.sInscricaoMunicipal
        .sInscricaoSuframa = objCliente.sInscricaoSuframa
        .sRazaoSocial = objCliente.sRazaoSocial
        .iAtivo = objCliente.iAtivo
        .lCodigoLoja = objCliente.lCodigoLoja
        .iFilialEmpresaLoja = objCliente.iFilialEmpresaLoja
        
        .sNomeReduzido = objCliente.sNomeReduzido
        .iTipo = objCliente.iTipo
        .dLimiteCredito = objCliente.dLimiteCredito
        .sObservacao = objCliente.sObservacao
        .iCondicaoPagto = objCliente.iCondicaoPagto
        .iTabelaPreco = objCliente.iTabelaPreco
        .sUsuarioCobrador = objCliente.sUsuarioCobrador
        .sUsuRespCallCenter = objCliente.sUsuRespCallCenter
        .iCodTransportadora = objCliente.iCodTransportadora
        .iVendedor = objCliente.iVendedor
        .sRG = objCliente.sRG
        .iRegimeTributario = objCliente.iRegimeTributario

        .dDesconto = objCliente.dDesconto
        .dComissaoVendas = objCliente.dComissaoVendas
        .iRegiao = objCliente.iRegiao
        
        If objCliente.iCodMensagem <> 0 Then
            objMensagem.iCodigo = objCliente.iCodMensagem
            'Lê a Mensagem "default" p/ o Cliente
            lErro = CF("Mensagem_Le", objMensagem)
            If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
            
            .sMensagemNF = objMensagem.sDescricao
        End If
        
        If colEndereco.Count > 0 Then
        
            lErro = CF("Endereco_Ajusta_Dados", colEndereco.Item(1))
            If lErro <> SUCESSO And lErro <> 47876 Then gError 113776
            
            .sEndereco = colEndereco.Item(1).sEndereco
            .sBairro = colEndereco.Item(1).sBairro
            .sCEP = colEndereco.Item(1).sCEP
            .sCidade = colEndereco.Item(1).sCidade
            .sUF = colEndereco.Item(1).sSiglaEstado
            
            .sTelefone1 = colEndereco.Item(1).sTelefone1
            .sTelefone2 = colEndereco.Item(1).sTelefone2
            .sEmail = colEndereco.Item(1).sEmail
            .sEmail2 = colEndereco.Item(1).sEmail2
            
            'Le o nome do Pais com o codigo no obj
            objPais.iCodigo = colEndereco.Item(1).iCodigoPais
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError 113776
        
            'Pais não está cadastrado - --- -> Erro
            If lErro = 47876 Then gError 113777
            
            .sPais = objPais.sNome
            
        End If
        
    End With
    
    'FAZ A GRAVACAO DO HISTORICO
    lErro = CF("ClienteHistorico_Grava", objClienteHistorico)
    If lErro <> SUCESSO Then gError 113778
    
    'Faz "locks" nas tabelas dos objetos relacionados ao Cliente
    lErro = CF("Clientes_Le_Tabelas_Lock", objCliente)
    If lErro <> SUCESSO Then gError 113779
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
    
        'Seleciona o Cliente no BD por codigoloja
        lErro = Comando_ExecutarPos(alComando(1), "SELECT CodigoLoja, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Ativo  FROM Clientes WHERE CodigoLoja = ?", 0, _
            tCliente.lCodigoLoja, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iAtivo, objCliente.lCodigoLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 113794
    
    Else
    
        'Seleciona o Cliente no BD por codigo
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Ativo  FROM Clientes WHERE Codigo = ?", 0, _
            tCliente.lCodigo, tCliente.sRazaoSocial, tCliente.sNomeReduzido, tCliente.iTipo, tCliente.sObservacao, tCliente.dLimiteCredito, tCliente.iCondicaoPagto, tCliente.dDesconto, tCliente.iCodMensagem, tCliente.iTabelaPreco, tCliente.iAtivo, objCliente.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 113780
    
    End If

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113781
    
    'Cliente não está cadastrado - GRAVACAO
    If lErro = AD_SQL_SEM_DADOS Then
    
        bClienteNovo = True
        
        'Passa p/ o objLog o tipo de operação a ser realizada.
        objLog.iOperacao = INCLUSAO_CLIENTE
        
        'Grava os novos Enderecos (Principal, Entrega, Cobranca) de Cliente
        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Automatico", lCodEndereco)
            If lErro <> SUCESSO Then gError 113782

            colEndereco.Item(iIndice).lCodigo = lCodEndereco

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 113783

        Next

        If colEndereco.Count > 0 Then objCliente.lEndereco = colEndereco.Item(1).lCodigo
        If colEndereco.Count > 1 Then objCliente.lEnderecoEntrega = colEndereco.Item(2).lCodigo
        If colEndereco.Count > 2 Then objCliente.lEnderecoCobranca = colEndereco.Item(3).lCodigo

        '###########################
        'Alterado por Wagner - Inserido a coluna Bloqueado
        lErro = Comando_Executar(alComando(2), "INSERT INTO Clientes (Codigo,RazaoSocial,NomeReduzido,Tipo,Observacao,LimiteCredito,CondicaoPagto,Desconto,CodMensagem,TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, CodigoLoja,FilialEmpresaLoja, Bloqueado, UsuarioCobrador, UsuRespCallCenter, IgnoraRecebPadrao, TemFaixaReceb,PercentMaisReceb,PercentMenosReceb,RecebForaFaixa) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objCliente.lCodigo, objCliente.sRazaoSocial, objCliente.sNomeReduzido, objCliente.iTipo, objCliente.sObservacao, objCliente.dLimiteCredito, objCliente.iCondicaoPagto, objCliente.dDesconto, objCliente.iCodMensagem, objCliente.iTabelaPreco, FILIAL_MATRIZ + 1, objCliente.iCodPadraoCobranca, objCliente.iAtivo, objCliente.lCodigoLoja, objCliente.iFilialEmpresaLoja, objCliente.iBloqueado, objCliente.sUsuarioCobrador, objCliente.sUsuRespCallCenter, objCliente.iIgnoraRecebPadrao, objCliente.iTemFaixaReceb, objCliente.dPercentMaisReceb, objCliente.dPercentMenosReceb, objCliente.iRecebForaFaixa)
        If lErro <> AD_SQL_SUCESSO Then gError 113784
        '###########################

        lErro = Comando_Executar(alComando(3), "INSERT INTO FiliaisClientes (CodCliente,CodFilial,Nome,CGC,InscricaoEstadual,InscricaoMunicipal,Endereco,EnderecoEntrega,EnderecoCobranca,CodTransportadora,Observacao,ContaContabil,Vendedor,ComissaoVendas,Regiao,FreqVisitas,DataUltVisita,CodCobrador,InscricaoSuframa, RG, CodFilialLoja, FilialEmpresaLoja, CodClienteLoja, Ativo, CodTranspRedesp, DetPagFrete, Guia, TipoFrete, DataUltAtualizacao, Exportado, RegimeTributario, IEIsento, IENaoContrib, DataRegistro ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                 objCliente.lCodigo, FILIAL_MATRIZ, MATRIZ, objCliente.sCgc, objCliente.sInscricaoEstadual, objCliente.sInscricaoMunicipal, objCliente.lEndereco, objCliente.lEnderecoEntrega, objCliente.lEnderecoCobranca, objCliente.iCodTransportadora, objCliente.sObservacao2, objCliente.sContaContabil, objCliente.iVendedor, objCliente.dComissaoVendas, objCliente.iRegiao, objCliente.iFreqVisitas, objCliente.dtDataUltVisita, objCliente.iCodCobrador, objCliente.sInscricaoSuframa, objCliente.sRG, FILIAL_MATRIZ, objCliente.iFilialEmpresaLoja, objCliente.lCodigoLoja, objCliente.iAtivo, objCliente.iCodTranspRedesp, objCliente.iDetPagFrete, objCliente.sGuia, objCliente.iTipoFrete, Date, 0, objCliente.iRegimeTributario, objCliente.iIEIsento, objCliente.iIENaoContrib, Date)
        If lErro <> AD_SQL_SUCESSO Then gError 113785

        'Insere as categorias da filial do cliente em FilialClientesCategorias
        objFilialCliente.lCodCliente = objCliente.lCodigo

        objFilialCliente.iCodFilial = FILIAL_MATRIZ
        
        lErro = CF("FilialCliente_Insere_Categorias", objFilialCliente, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 113786

    Else 'Cliente já existe - MODIFICACAO
        
        'se o campo ativo do BD vai ser desativado -->desativa todas as filiais
        If tCliente.iAtivo <> objCliente.iAtivo And objCliente.iAtivo = DESMARCADO Then
            'Seleciona as filiais Clientes no BD por codigo(menos a matriz)
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM FiliaisClientes WHERE CodClienteLoja = ? AND CodFilialLoja <> ? AND Ativo = ?", 0, iAtivo, objCliente.lCodigoLoja, FILIAL_MATRIZ, CLIENTE_ATIVO)
            Else
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial <> ? AND Ativo = ?", 0, Ativo, objCliente.lCodigo, FILIAL_MATRIZ, CLIENTE_ATIVO)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 112730
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112731
            
            Do While lErro = AD_SQL_SUCESSO
                'Atualiza a tabela de FiliaisClientes
                lErro = Comando_ExecutarPos(alComando(6), "UPDATE FiliaisClientes SET Ativo = ?", alComando(5), objCliente.iAtivo)
                If lErro <> AD_SQL_SUCESSO Then gError 112732
                                
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 112733
            Loop
        End If
        
        'Passa p/ o objLog o tipo de operação a ser realizada.
        objLog.iOperacao = ALTERACAO_CLIENTE
    
        'Seleciona a tabela de Filiais Clientes (para modificar a Matriz e pegar os Codigos de Enderecos em tFilial.lEndereco)
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objCliente.lCodigo = 0 Then
            lErro = Comando_ExecutarPos(alComando(3), "SELECT CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador,InscricaoSuframa, RG FROM  FiliaisClientes WHERE CodClienteLoja = ? AND CodFilialLoja = ?", 0, _
                tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, tFilialCliente.iCodTransportadora, tFilialCliente.sObservacao, tFilialCliente.sContaContabil, tFilialCliente.iVendedor, tFilialCliente.dComissaoVendas, tFilialCliente.iRegiao, tFilialCliente.iFreqVisitas, tFilialCliente.dtDataUltVisita, tFilialCliente.iCodCobrador, tFilialCliente.sInscricaoSuframa, tFilialCliente.sRG, objCliente.lCodigoLoja, FILIAL_MATRIZ)
        Else
            lErro = Comando_ExecutarPos(alComando(3), "SELECT CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador,InscricaoSuframa, RG FROM  FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", 0, _
                tFilialCliente.sCgc, tFilialCliente.sInscricaoEstadual, tFilialCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, tFilialCliente.iCodTransportadora, tFilialCliente.sObservacao, tFilialCliente.sContaContabil, tFilialCliente.iVendedor, tFilialCliente.dComissaoVendas, tFilialCliente.iRegiao, tFilialCliente.iFreqVisitas, tFilialCliente.dtDataUltVisita, tFilialCliente.iCodCobrador, tFilialCliente.sInscricaoSuframa, tFilialCliente.sRG, objCliente.lCodigo, FILIAL_MATRIZ)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 113787

        'Verifica se filial cliente existe
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 113788

        'Faz a modificação dos três Endereços
        If colEndereco.Count > 0 Then colEndereco.Item(1).lCodigo = tFilialCliente.lEndereco
        If colEndereco.Count > 1 Then colEndereco.Item(2).lCodigo = tFilialCliente.lEnderecoEntrega
        If colEndereco.Count > 2 Then colEndereco.Item(3).lCodigo = tFilialCliente.lEnderecoCobranca

        'Atualiza os endereços
        For iIndice = 1 To colEndereco.Count

            If colEndereco.Item(iIndice).lCodigo = 0 Then

                lErro = CF("Endereco_Automatico", lCodEndereco)
                If lErro <> SUCESSO Then gError 126015
    
                colEndereco.Item(iIndice).lCodigo = lCodEndereco

            End If

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 113789

        Next

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            '#####################################
            'Alterado por Wagner
            'Atualiza a tabela de Clientes
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE Clientes SET RazaoSocial = ?, NomeReduzido = ?, Ativo = ?, Bloqueado=?, UsuarioCobrador = ?, UsuRespCallCenter = ?", alComando(1), _
                objCliente.sRazaoSocial, objCliente.sNomeReduzido, objCliente.iAtivo, objCliente.iBloqueado, objCliente.sUsuarioCobrador, objCliente.sUsuRespCallCenter)
            If lErro <> AD_SQL_SUCESSO Then gError 110111
            '#####################################
    
            'Atualiza a tabela de FiliaisClientes
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE FiliaisClientes SET CGC = ?, RG=?, Ativo = ?, Endereco = ?, DataUltAtualizacao = ?, Exportado = ?", alComando(3), _
                objCliente.sCgc, objCliente.sRG, objCliente.iAtivo, tFilialCliente.lEndereco, Date, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 110112

        Else
        
            '#####################################
            'Alterado por Wagner
            'Atualiza a tabela de Clientes
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE Clientes SET RazaoSocial = ?, NomeReduzido = ?, Tipo = ?, Observacao = ?, LimiteCredito = ?, CondicaoPagto = ?, Desconto = ?, CodMensagem = ?, TabelaPreco = ?, CodPadraoCobranca = ?, Ativo = ?, Bloqueado = ?, UsuarioCobrador = ?, UsuRespCallCenter = ?, IgnoraRecebPadrao = ?, TemFaixaReceb = ?,PercentMaisReceb = ?,PercentMenosReceb = ?,RecebForaFaixa = ?  ", alComando(1), _
                objCliente.sRazaoSocial, objCliente.sNomeReduzido, objCliente.iTipo, objCliente.sObservacao, objCliente.dLimiteCredito, objCliente.iCondicaoPagto, objCliente.dDesconto, objCliente.iCodMensagem, objCliente.iTabelaPreco, objCliente.iCodPadraoCobranca, objCliente.iAtivo, objCliente.iBloqueado, objCliente.sUsuarioCobrador, objCliente.sUsuRespCallCenter, objCliente.iIgnoraRecebPadrao, objCliente.iTemFaixaReceb, objCliente.dPercentMaisReceb, objCliente.dPercentMenosReceb, objCliente.iRecebForaFaixa)
            If lErro <> AD_SQL_SUCESSO Then gError 113790
            '#####################################
    
            'Atualiza a tabela de FiliaisClientes
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE FiliaisClientes SET CGC = ?, InscricaoEstadual = ?, InscricaoMunicipal = ?, Endereco = ?, EnderecoEntrega = ?, EnderecoCobranca = ?, CodTransportadora = ?, Observacao = ?, ContaContabil = ?, Vendedor = ?, ComissaoVendas = ?, Regiao = ?, FreqVisitas = ?, DataUltVisita = ?, CodCobrador = ?, TipoFrete = ?,InscricaoSuframa=?, RG=?, Ativo = ?, CodTranspRedesp=?, DetPagFrete=?, Guia=?, DataUltAtualizacao = ?, Exportado = ?, RegimeTributario=?, IEIsento = ?, IENaoContrib = ?", alComando(3), _
                objCliente.sCgc, objCliente.sInscricaoEstadual, objCliente.sInscricaoMunicipal, tFilialCliente.lEndereco, tFilialCliente.lEnderecoEntrega, tFilialCliente.lEnderecoCobranca, objCliente.iCodTransportadora, objCliente.sObservacao2, objCliente.sContaContabil, objCliente.iVendedor, objCliente.dComissaoVendas, objCliente.iRegiao, objCliente.iFreqVisitas, objCliente.dtDataUltVisita, objCliente.iCodCobrador, objCliente.iTipoFrete, objCliente.sInscricaoSuframa, objCliente.sRG, objCliente.iAtivo, objCliente.iCodTranspRedesp, objCliente.iDetPagFrete, objCliente.sGuia, Date, 0, objCliente.iRegimeTributario, objCliente.iIEIsento, objCliente.iIENaoContrib)
            If lErro <> AD_SQL_SUCESSO Then gError 113791

        End If

        objFilialCliente.lCodCliente = objCliente.lCodigo
        
        objFilialCliente.iCodFilial = FILIAL_MATRIZ
        
        lErro = CF("FilialCliente_Atualiza_Categorias", objFilialCliente, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 113792

    End If
    
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
    
        'Monta a String de log que deverá ser gravada.
        Call Cliente_Carrega_Log(objLog, objCliente, colEndereco)
    
        'Monta Registro Para a Tabela de Log
        lErro = CF("Log_Grava", objLog)
        If lErro <> SUCESSO Then gError 113793
        
    End If
    
    lErro = CF("Cliente_Grava_Customizado", objCliente)
    If lErro <> SUCESSO Then gError 192968
    
    lErro = CF("Cliente_Grava_Customizado2", objCliente, bClienteNovo)
    If lErro <> SUCESSO Then gError 192968
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Cliente_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Cliente_Grava_EmTrans:

    Cliente_Grava_EmTrans = gErr

    Select Case gErr

        Case 110111, 113790
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", gErr, lCodigoCliente)

        Case 110112, 113791, 112732
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISCLIENTES", gErr)
            
        Case 113772
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 113771, 113773, 113774, 113780, 113781, 113794
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES", gErr)

        Case 113775
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RED_DUPLICADO", gErr)

        Case 113779, 113782, 113783, 113789, 113786, 113792, 113793, 126015, 192968

        Case 113784
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLIENTES", gErr, lCodigoCliente)
        
        Case 113785
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIAISCLIENTES", gErr)

        Case 113787
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, lCodigoCliente)

        Case 113788, 112730, 112731, 112733
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr)

        Case 113776, 113778
        
        Case 113777
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, objPais.iCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148981)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Log_Grava(ByVal objLog As ClassLog, Optional ByVal colRegistro As Collection) As Long
'grava o registro de log

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim vRegistro As Variant
Dim lTamanho As Long
Dim iQtdeNumIntDoc As Integer
Dim iQtdeNumIntDoc1 As Integer
Dim iPosicao As Integer

On Error GoTo Erro_Log_Grava

    'Abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 104185

    If colRegistro Is Nothing Then
        Set colRegistro = New Collection
    End If

    If colRegistro.Count > 0 Then
    
        For Each vRegistro In colRegistro
        
            '1020 = 4 * 255 que é o tamanho do registro de log
            iQtdeNumIntDoc = iQtdeNumIntDoc + Len(vRegistro) \ MAX_REGISTRO_LOG
        
            If Len(vRegistro) Mod MAX_REGISTRO_LOG > 0 Then iQtdeNumIntDoc = iQtdeNumIntDoc + 1
        
        Next
        
    Else
    
        lTamanho = Len(objLog.sLog)
    
        '1020 = 4 * 255 que é o tamanho do registro de log
        iQtdeNumIntDoc = lTamanho \ MAX_REGISTRO_LOG
        
        If lTamanho Mod MAX_REGISTRO_LOG > 0 Then iQtdeNumIntDoc = iQtdeNumIntDoc + 1
    
    End If
    
    lErro = CF("Config_ObterNumInt", "LojaConfig", "NUM_PROX_LOG", lNumIntDoc, iQtdeNumIntDoc)
    If lErro <> SUCESSO Then gError 104192
    
    objLog.lNumIntDoc = lNumIntDoc
    
    If colRegistro.Count > 0 Then
    
        For Each vRegistro In colRegistro
    
            iPosicao = 1
    
            '1020 = 4 * 255 que é o tamanho do registro de log
            iQtdeNumIntDoc1 = Len(vRegistro) \ MAX_REGISTRO_LOG
            
            If Len(vRegistro) Mod MAX_REGISTRO_LOG > 0 Then iQtdeNumIntDoc1 = iQtdeNumIntDoc1 + 1
    
            'indica quantos registros no bd terão que ser lidos em seguida para compor o registro no arquivo
            objLog.iContinuacao = iQtdeNumIntDoc1 - 1
    
            Do While iPosicao <= Len(vRegistro)
    
                objLog.sLog = Mid(vRegistro, iPosicao, MAX_REGISTRO_LOG)
                
                'Chama a funcao que insere o registro no log
                lErro = Log_Grava_1(objLog, lComando)
                If lErro <> SUCESSO Then gError 110030
                
                objLog.lNumIntDoc = objLog.lNumIntDoc + 1
                objLog.iContinuacao = 0
                iPosicao = iPosicao + MAX_REGISTRO_LOG

            Loop
            
        Next
        
    Else
    
        objLog.lNumIntDoc = lNumIntDoc
    
        vRegistro = objLog.sLog
    
        iPosicao = 1
    
        '1020 = 4 * 255 que é o tamanho do registro de log
        iQtdeNumIntDoc1 = Len(vRegistro) \ MAX_REGISTRO_LOG
        
        If Len(vRegistro) Mod MAX_REGISTRO_LOG > 0 Then iQtdeNumIntDoc1 = iQtdeNumIntDoc1 + 1

        'indica quantos registros no bd terão que ser lidos em seguida para compor o registro no arquivo
        objLog.iContinuacao = iQtdeNumIntDoc1 - 1
    
        Do While iPosicao <= Len(vRegistro)

            objLog.sLog = Mid(vRegistro, iPosicao, MAX_REGISTRO_LOG)
            
            'Chama a funcao que insere o registro no log
            lErro = Log_Grava_1(objLog, lComando)
            If lErro <> SUCESSO Then gError 110030
            
            objLog.lNumIntDoc = objLog.lNumIntDoc + 1
            objLog.iContinuacao = 0
            iPosicao = iPosicao + MAX_REGISTRO_LOG

        Loop
    
    End If
    
    'vai permitir que outros registros guardem o identificador do log e numero de registros no bd que foram utilizados para armazenar a info.
    objLog.lNumIntDoc = lNumIntDoc
    objLog.iQuantLog = iQtdeNumIntDoc
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

    Log_Grava = SUCESSO

    Exit Function

Erro_Log_Grava:

    Log_Grava = gErr

    Select Case gErr

        Case 104185
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 104192, 110030, 110031
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148982)

    End Select

    'Fecha o comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function Log_Grava_1(ByVal objLog As ClassLog, lComando As Long) As Long
'insere o registro na tabela log

Dim lErro As Long
Dim sLog1 As String
Dim sLog2 As String
Dim sLog3 As String
Dim sLog4 As String

On Error GoTo Erro_Log_Grava_1

    objLog.dHora = CDbl(Time)
    objLog.dtData = gdtDataAtual
    
    sLog1 = Mid(objLog.sLog, 1, STRING_LOG)
    sLog2 = Mid(objLog.sLog, 256, STRING_LOG)
    sLog3 = Mid(objLog.sLog, 512, STRING_LOG)
    sLog4 = Mid(objLog.sLog, 768, STRING_LOG)
    
    lErro = Comando_Executar(lComando, "INSERT INTO Log (NumIntDoc, Operacao , Log1, Log2 , Log3, Log4, Data, Hora) VALUES (?,?,?,?,?,?,?,?)", objLog.lNumIntDoc, objLog.iOperacao, sLog1, sLog2, sLog3, sLog4, objLog.dtData, objLog.dHora)
    If lErro <> AD_SQL_SUCESSO Then gError 104186

    Log_Grava_1 = SUCESSO

    Exit Function

Erro_Log_Grava_1:

    Log_Grava_1 = gErr

    Select Case gErr

        Case 104186
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_LOG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148983)

    End Select

    Exit Function

End Function

Private Sub Cliente_Carrega_Log(ByVal objLog As ClassLog, ByVal objCliente As ClassCliente, ByVal colEndereco As Collection)

Dim objEndereco As ClassEndereco

    With objCliente
    
        'Função para Carregar o objLog
        objLog.sLog = CStr(.lCodigo) & Chr(vbKeyEscape) & CStr(.lCodigoLoja) & Chr(vbKeyEscape) & CStr(.sCgc) & Chr(vbKeyEscape) & CStr(.sNomeReduzido) & Chr(vbKeyEscape) & CStr(.sRG) & Chr(vbKeyEscape)
        objLog.sLog = objLog.sLog & CStr(.sRazaoSocial) & Chr(vbKeyEscape) & CStr(.iAtivo) & Chr(vbKeyEscape) & CStr(.iFilialEmpresaLoja) & Chr(vbKeyEscape)
        
    End With
    
    Set objEndereco = colEndereco.Item(1)

    With objEndereco
    
        'Função para Carregar o objLog
        objLog.sLog = objLog.sLog & CStr(.iCodigoPais) & Chr(vbKeyEscape) & CStr(.lCodigo) & Chr(vbKeyEscape) & CStr(.sBairro) & Chr(vbKeyEscape) & CStr(.sCEP) & Chr(vbKeyEscape) & CStr(.sCidade) & Chr(vbKeyEscape) & CStr(.sContato) & Chr(vbKeyEscape) & CStr(.sEmail) & Chr(vbKeyEscape) & CStr(.sEndereco) & Chr(vbKeyEscape) & CStr(.sFax) & Chr(vbKeyEscape) & CStr(.sSiglaEstado) & Chr(vbKeyEscape) & CStr(.sTelefone1) & Chr(vbKeyEscape) & CStr(.sTelefone2) & Chr(vbKeyEscape)
        
    End With
    
    objLog.sLog = objLog.sLog & Chr(vbKeyEnd)

    Exit Sub

End Sub

Public Function CondPagto_Critica_Compras(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
'Verifica se a Condição de Pagamento esta sendo utilizada no COMPRAS

Dim lErro As Long
Dim lComando As Long
Dim iCondicaoPagto As Integer

On Error GoTo Erro_CondPagto_Critica_Compras

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87307
        
    'Faz leitura na view PedidoCompraTodos com a chave CondicaoPagto
    lErro = Comando_Executar(lComando, "SELECT CondicaoPagto FROM PedidoCompraTodos WHERE CondicaoPagto = ?", iCondicaoPagto, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 87308
    
    'Busca o primeiro registro que satifaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87309
    
    'Se encontrou registro então dispara erro
    If lErro = AD_SQL_SUCESSO Then gError 87310
    
    'Faz leitura na view PedidoCotacaoTodos com a chave CondicaoPagto
    lErro = Comando_Executar(lComando, "SELECT CondPagtoPrazo FROM PedidoCotacaoTodos WHERE CondPagtoPrazo = ?", iCondicaoPagto, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 87311
        
    'Busca o primeiro registro que satifaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87312
    
    'Se encontrou registro então dispara erro
    If lErro = AD_SQL_SUCESSO Then gError 87313
    
    'Faz leitura na tabela OrçamentoVenda com a chave CondicaoPagto
    lErro = Comando_Executar(lComando, "SELECT CondicaoPagto FROM OrcamentoVenda WHERE CondicaoPagto = ?", iCondicaoPagto, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 87314
        
    'Busca o primeiro registro que satifaz a condição
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87315
    
    'Se encontrou registro então dispara erro
    If lErro = AD_SQL_SUCESSO Then gError 87316
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    CondPagto_Critica_Compras = SUCESSO

    Exit Function

Erro_CondPagto_Critica_Compras:

    CondPagto_Critica_Compras = gErr

    Select Case gErr
    
        Case 87307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 87308, 87309
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOMPRATODOS", gErr)

        Case 87310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_VINCULADA_PEDCOMPRAS", gErr, objCondicaoPagto.iCodigo)

        Case 87311, 87312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PEDIDOCOTACAOTODOS", gErr)
            
        Case 87313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_VINCULADO_PEDCOTACAO", gErr, objCondicaoPagto.iCodigo)

        Case 87314, 87315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORCAMENTOVENDA1", gErr)

       Case 87316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDPAGTO_VINCULADO_ORCAMENTOVENDA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148984)
            
    End Select
    
    'Fecha Comandos
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function FilialFornecedor_Grava_Estatisticas(ByVal objNFiscal As ClassNFiscal, ByVal iCancelamento As Integer, Optional ByVal objFilFornFilEmp As ClassFilFornFilEmp) As Long
'Tem que ser chamado dentro de Transação
'Atualiza a gravação da parte de estatística do Fornecedor

Dim lErro As Long
Dim objFilFornFilEmp1 As New ClassFilFornFilEmp
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim lMaiorAtraso As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_FilialFornecedor_Grava_Estatisticas

    If objFilFornFilEmp Is Nothing Then Set objFilFornFilEmp = New ClassFilFornFilEmp

    'Inicializa comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 52634

    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then gError 52635

    lComando3 = Comando_Abrir()
    If lComando3 = 0 Then gError 52636
    
    If Not objNFiscal Is Nothing Then
    
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 89129
    
        If lErro <> SUCESSO Then gError 89130
        
        'Trata-se de uma nota fiscal de compras
        If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Or objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then
    
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then
    
                objFilFornFilEmp.lNumeroCompras = 1
                objFilFornFilEmp.dValorAcumuladoCompras = objNFiscal.dValorTotal
            
            Else
                
                objFilFornFilEmp.lNumTotalDevolucoes = 1
                objFilFornFilEmp.dValorAcumuladoDevolucoes = objNFiscal.dValorTotal
                
            End If
            
            objFilFornFilEmp1.lCodFornecedor = objNFiscal.lFornecedor
            objFilFornFilEmp1.iFilialEmpresa = objNFiscal.iFilialEmpresa
            objFilFornFilEmp1.iCodFilial = objNFiscal.iFilialForn
        
        Else
        
            'Adicionado por Rafael em 22/08/2002
            objFilFornFilEmp.lCodFornecedor = objNFiscal.lFornecedor
            objFilFornFilEmp.iFilialEmpresa = objNFiscal.iFilialEmpresa
            objFilFornFilEmp.iCodFilial = objNFiscal.iFilialForn
            'Fim Adicionado por Rafael em 22/08/2002
            
            objFilFornFilEmp1.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
            objFilFornFilEmp1.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
            objFilFornFilEmp1.iCodFilial = objFilFornFilEmp.iCodFilial
            
        End If

    Else
    
        objFilFornFilEmp1.lCodFornecedor = objFilFornFilEmp.lCodFornecedor
        objFilFornFilEmp1.iFilialEmpresa = objFilFornFilEmp.iFilialEmpresa
        objFilFornFilEmp1.iCodFilial = objFilFornFilEmp.iCodFilial

    End If

    If (Not objNFiscal Is Nothing And objTipoDocInfo.iOrigem = DOCINFO_ORIGEM_FORNECEDOR) Or objNFiscal Is Nothing Then
        
        lErro = CF("FilFornFilEmp_Le_Lock", lComando, objFilFornFilEmp1)
        If lErro <> SUCESSO And lErro <> 52642 Then gError 52637

        'se não encontrou
        If lErro = 52642 Then

            'se é um cancelamento/exclusao de nota e não encontrou o registro na tabela FilFornFilEmp ==> erro
            'If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then gError 89418 'Comentado por Wagner

            If iCancelamento <> CANCELAMENTO_FILIALFORNFILEMP Then

                'Insere na tabela de FilialFornFilEmp
                lErro = Comando_Executar(lComando3, "INSERT INTO FilialFornFilEmp (FilialEmpresa, CodFornecedor, CodFilial, NumeroCompras, ValorAcumuladoCompras, NumTotalDevolucoes, ValorAcumuladoDevolucoes, PedidosEmAberto, AtrasoAcumulado, PedidosRecebidos, ItensPedidosRecebidos) VALUES (?,?,?,?,?,?,?,?,?,?,?)", objFilFornFilEmp1.iFilialEmpresa, objFilFornFilEmp1.lCodFornecedor, objFilFornFilEmp1.iCodFilial, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 52638
                
            End If

        'se encontrou
        Else
            
            If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then
            
                'Atualiza a tabela de FilialFornFilEmp subtraindo os valores
                lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialFornFilEmp SET NumeroCompras=NumeroCompras - ?, ValorAcumuladoCompras=ValorAcumuladoCompras-?, NumTotalDevolucoes=NumTotalDevolucoes-?, ValorAcumuladoDevolucoes=ValorAcumuladoDevolucoes-?, PedidosEmAberto=PedidosEmAberto-?, AtrasoAcumulado=AtrasoAcumulado-?, PedidosRecebidos=PedidosRecebidos-?, ItensPedidosRecebidos=ItensPedidosRecebidos-?", lComando, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 89419
            
            Else
            
                'Atualiza a tabela de FilialFornFilEmp somando os valores
                lErro = Comando_ExecutarPos(lComando2, "UPDATE FilialFornFilEmp SET NumeroCompras=NumeroCompras + ?, ValorAcumuladoCompras=ValorAcumuladoCompras+?, NumTotalDevolucoes=NumTotalDevolucoes+?, ValorAcumuladoDevolucoes=ValorAcumuladoDevolucoes+?, PedidosEmAberto=PedidosEmAberto+?, AtrasoAcumulado=AtrasoAcumulado+?, PedidosRecebidos=PedidosRecebidos+?, ItensPedidosRecebidos=ItensPedidosRecebidos+?", lComando, objFilFornFilEmp.lNumeroCompras, objFilFornFilEmp.dValorAcumuladoCompras, objFilFornFilEmp.lNumTotalDevolucoes, objFilFornFilEmp.dValorAcumuladoDevolucoes, objFilFornFilEmp.lPedidosEmAberto, objFilFornFilEmp.lAtrasoAcumulado, objFilFornFilEmp.lPedidosRecebidos, objFilFornFilEmp.lItensPedidosRecebidos)
                If lErro <> AD_SQL_SUCESSO Then gError 52639

            End If

        End If
 
        If Not objNFiscal Is Nothing Then
        
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then

                'Atualiza as tabelas SldMesForn e SldDiaForn
                lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, iCancelamento)
                If lErro <> SUCESSO Then gError 89345
        
            ElseIf objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then

                If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then

                    'Atualiza as tabelas SldMesForn e SldDiaForn
                    lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, CADASTRAMENTO_FILIALFORNFILEMP)
                    If lErro <> SUCESSO Then gError 89793
                    
                Else

                    'Atualiza as tabelas SldMesForn e SldDiaForn
                    lErro = CF("SldMesForn_Grava_Estatisticas", objNFiscal, CANCELAMENTO_FILIALFORNFILEMP)
                    If lErro <> SUCESSO Then gError 89794

                End If

            End If

        End If

    End If

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    FilialFornecedor_Grava_Estatisticas = SUCESSO

    Exit Function

Erro_FilialFornecedor_Grava_Estatisticas:

    FilialFornecedor_Grava_Estatisticas = gErr

    Select Case gErr

        Case 52634, 52635, 52636
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 52637, 89129, 89345, 89793, 89794

        Case 52638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERIR_FILIALFORNFILEMP", gErr)

        Case 52639, 89419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_FILIALFORNFILEMP", gErr)

        Case 89130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr)

        Case 89418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNFILEMP_NAO_CADASTRADO", gErr, objFilFornFilEmp1.lCodFornecedor, objFilFornFilEmp1.iCodFilial, objFilFornFilEmp1.iFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148985)

    End Select

    'Libera comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    Call Comando_Fechar(lComando3)
    
    Exit Function

End Function

Function SldMesForn_Grava_Estatisticas(ByVal objNFiscal As ClassNFiscal, ByVal iCancelamento As Integer) As Long
'Atualiza as tabelas SldMesForn e SldDiaForn
'Tem que ser chamado dentro de Transação

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lComando3 As Long
Dim tSldDiaForn As typeSldDiaForn
Dim tSldMesForn As typeSldMesForn
Dim objItemNF As ClassItemNF
Dim objSldDiaForn As New ClassSldDiaForn
Dim dFator As Double
Dim iIndice As Integer
Dim alComando(1 To 4) As Long

On Error GoTo Erro_SldMesForn_Grava_Estatisticas

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89341
    Next

    For Each objItemNF In objNFiscal.ColItensNF
        
        objSldDiaForn.iFilialEmpresa = objNFiscal.iFilialEmpresa
        objSldDiaForn.lFornecedor = objNFiscal.lFornecedor
        objSldDiaForn.iFilialForn = objNFiscal.iFilialForn
        objSldDiaForn.sProduto = objItemNF.sProduto
        objSldDiaForn.dtData = objNFiscal.dtDataEmissao
        
        'Calcula fator de conversão para UM_Venda
        lErro = CF("UM_Conversao_Trans", objItemNF.iClasseUM, objItemNF.sUnidadeMed, objItemNF.sUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 89342
    
        If iCancelamento = CANCELAMENTO_FILIALFORNFILEMP Then
        
            'guarda a quantidade na unidade de estoque
            objSldDiaForn.dQuantCompra = -objItemNF.dQuantidade * dFator
            objSldDiaForn.dValorCompra = -objItemNF.dValorTotal
    
        Else
        
            'guarda a quantidade na unidade de estoque
            objSldDiaForn.dQuantCompra = objItemNF.dQuantidade * dFator
            objSldDiaForn.dValorCompra = objItemNF.dValorTotal
    
        End If
        
        'Atualiza a tabela SldDiaForn utilizando os valores passados atraves de objSldDiaForn.
        lErro = SldDiaForn_Atualiza(alComando(1), alComando(2), objSldDiaForn)
        If lErro <> SUCESSO Then gError 89343

        'Atualiza a tabela SldMesForn utilizando os valores passados atraves de objSldDiaForn.
        lErro = SldMesForn_Atualiza(alComando(3), alComando(4), objSldDiaForn)
        If lErro <> SUCESSO Then gError 89344

    Next

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    SldMesForn_Grava_Estatisticas = SUCESSO

    Exit Function

Erro_SldMesForn_Grava_Estatisticas:

    SldMesForn_Grava_Estatisticas = gErr

    Select Case gErr

        Case 89341
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89342, 89343, 89344

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148986)

    End Select

    'Fecha os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function SldDiaForn_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objSldDiaForn As ClassSldDiaForn) As Long
'Atualiza a tabela SldDiaForn utilizando os valores passados atraves de objSldDiaForn.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim dtData As Date

On Error GoTo Erro_SldDiaForn_Atualiza

    'Le a tabela de SldDiaForn para a filial/Fornecedor/FilialForn/produto/data em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Data FROM SldDiaForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND Produto = ? AND Data = ?", 0, dtData, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, objSldDiaForn.dtData)
    If lErro <> AD_SQL_SUCESSO Then gError 89331

    'leitura do registro referente a filial/Fornecedor/FilialForn/produto/data em questão
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89332
    
    'se não encontrou um registro para a data em questão ==> insere um registro
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere um registro na tabela de SldDiaForn para a filial/Fornecedor/FilialForn/produto/data em questão
        lErro = Comando_Executar(lComando2, "INSERT INTO SldDiaForn (FilialEmpresa, Fornecedor, FilialForn, Produto, Data, QuantCompra, ValorCompra) VALUES (?,?,?,?,?,?,?)", _
        objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, objSldDiaForn.dtData, objSldDiaForn.dQuantCompra, objSldDiaForn.dValorCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 89333
    
    Else
        'se encontrou ==> atualiza o registro
    
        'lock do registro em questão
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 89334
    
        'Atualiza a tabela de SldDiaForn para a filial/Fornecedor/FilialForn/produto/data em questão
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SldDiaForn SET QuantCompra = QuantCompra + ?, ValorCompra = ValorCompra + ?", lComando1, objSldDiaForn.dQuantCompra, objSldDiaForn.dValorCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 89335
        
    End If

    SldDiaForn_Atualiza = SUCESSO

    Exit Function

Erro_SldDiaForn_Atualiza:

    SldDiaForn_Atualiza = gErr

    Select Case gErr

        Case 89331, 89332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDDIAFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(objSldDiaForn.dtData))

        Case 89333
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDDIAFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(objSldDiaForn.dtData))
        
        Case 89334
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDDIAFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(objSldDiaForn.dtData))

        Case 89335
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDDIAFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(objSldDiaForn.dtData))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148987)

    End Select

    Exit Function

End Function

Private Function SldMesForn_Atualiza(ByVal lComando1 As Long, ByVal lComando2 As Long, ByVal objSldDiaForn As ClassSldDiaForn) As Long
'Atualiza a tabela SldMesForn utilizando os valores passados atraves de objSldDiaForn.
'ATENCAO: Esta Funcao tem que ser chamada dentro de transacao

Dim lErro As Long
Dim iAno As Integer
Dim sMes As String

On Error GoTo Erro_SldMesForn_Atualiza

    sMes = CStr(Month(objSldDiaForn.dtData))

    'Le a tabela de SldMesForn para a filial/Fornecedor/FilialForn/produto/Ano em questão
    lErro = Comando_ExecutarPos(lComando1, "SELECT Ano FROM SldMesForn WHERE FilialEmpresa = ? AND Fornecedor = ? AND FilialForn = ? AND Produto = ? AND Ano = ?", 0, iAno, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, Year(objSldDiaForn.dtData))
    If lErro <> AD_SQL_SUCESSO Then gError 89336

    'leitura do registro referente a filial/Fornecedor/FilialForn/produto/data em questão
    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89337
    
    'se não encontrou um registro em questão ==> insere um registro
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere um registro na tabela de SldMesForn para a filial/Fornecedor/FilialForn/produto/Ano em questão
        lErro = Comando_Executar(lComando2, "INSERT INTO SldMesForn (FilialEmpresa, Fornecedor, FilialForn, Produto, Ano, QuantCompras" + sMes + ", ValorCompras" + sMes + ") VALUES (?,?,?,?,?,?,?)", _
        objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, Year(objSldDiaForn.dtData), objSldDiaForn.dQuantCompra, objSldDiaForn.dValorCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 89338
    
    Else
        'se encontrou ==> atualiza o registro
    
        'lock do registro em questão
        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 89339
    
        'Atualiza a tabela de SldMesForn para a filial/Fornecedor/FilialForn/produto/data em questão
        lErro = Comando_ExecutarPos(lComando2, "UPDATE SldMesForn SET QuantCompras" + sMes + " = QuantCompras" + sMes + " + ?, ValorCompras" + sMes + " = ValorCompras" + sMes + " + ?", lComando1, objSldDiaForn.dQuantCompra, objSldDiaForn.dValorCompra)
        If lErro <> AD_SQL_SUCESSO Then gError 89340
        
    End If

    SldMesForn_Atualiza = SUCESSO

    Exit Function

Erro_SldMesForn_Atualiza:

    SldMesForn_Atualiza = gErr

    Select Case gErr

        Case 89336, 89337
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_SLDMESFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(Year(objSldDiaForn.dtData)))

        Case 89338
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_SLDMESFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(Year(objSldDiaForn.dtData)))
        
        Case 89339
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_SLDMESFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(Year(objSldDiaForn.dtData)))

        Case 89340
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_SLDMESFORN", gErr, objSldDiaForn.iFilialEmpresa, objSldDiaForn.lFornecedor, objSldDiaForn.iFilialForn, objSldDiaForn.sProduto, CStr(Year(objSldDiaForn.dtData)))

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148988)

    End Select

    Exit Function

End Function

Function NFiscal_Grava_DataHora(ByVal objNFiscal As ClassNFiscal) As Long
'Atualiza os campos Data e Hora de entrada e saida da nota fiscal passada como parametro

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lNumNotaFiscal As Long
Dim iFilialEmpresa As Integer
Dim iIndice As Integer
Dim sSerie As String

On Error GoTo Erro_NFiscal_Grava_DataHora

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 89715

    Next

    sSerie = String(STRING_SERIE, 0)

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Serie, NumNotaFiscal, FilialEmpresa FROM NFiscal WHERE NumIntDoc = ?", 0, sSerie, lNumNotaFiscal, iFilialEmpresa, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 89716

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89717

    If lErro = AD_SQL_SEM_DADOS Then gError 89718
    
    'alterado por mario em 01/10/01 para permitir alterar os campos Origem, Destino, ValorContainer, ValorMercadoria e VolumeEspecie
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFiscal SET TaxaMoeda = ?, SemDataSaida = ?, NaoImpCobranca = ?, DataEmissao = ?, DataEntrada = ?, DataSaida = ?, HoraEntrada = ?, HoraSaida = ?, OrigemNF = ?, DestinoNF = ?, ValorContainer = ?, ValorMercadoria = ?, VolumeEspecie = ?, DataAtualizacao=?, MensagemNota = ?, MensagemCorpoNota = ?, RPS = ?", alComando(1), objNFiscal.dTaxaMoeda, objNFiscal.iSemDataSaida, objNFiscal.iNaoImpCobranca, objNFiscal.dtDataEmissao, objNFiscal.dtDataEntrada, objNFiscal.dtDataSaida, CDbl(objNFiscal.dtHoraEntrada), CDbl(objNFiscal.dtHoraSaida), objNFiscal.sOrigem, objNFiscal.sDestino, objNFiscal.dValorContainer, objNFiscal.dValorMercadoria, objNFiscal.lVolumeEspecie, gdtDataHoje, objNFiscal.sMensagemNota, objNFiscal.sMensagemCorpoNota, objNFiscal.iRecibo)
    If lErro <> AD_SQL_SUCESSO Then gError 89719
        
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    NFiscal_Grava_DataHora = SUCESSO

    Exit Function

Erro_NFiscal_Grava_DataHora:

    NFiscal_Grava_DataHora = gErr

    Select Case gErr

        Case 89715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 89716, 89717
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)

        Case 89718
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 89719
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFISCAL", gErr, sSerie, lNumNotaFiscal, iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148989)

    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function TituloPagar_Altera(ByVal objTituloPagar As ClassTituloPagar) As Long
'Lê o Titulo associado ao numintdoc e se existir altera a data de emissao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_TituloPagar_Altera
    
    If objTituloPagar.lNumIntDoc <> 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 89724
        Next
        
        'Seleciona TituloPagar a partir dos dados fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TitulosPag WHERE NumIntDoc = ? And Status <> ?", 0, _
            lNumIntDoc, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 89725
        
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89726
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPag SET DataEmissao = ?", alComando(1), objTituloPagar.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 89727
        
        'Se não encontrou --> pesquisa nos titulos baixados
        Else
        
            'Seleciona TituloPagarBaixado a partir dos dados fornecidos
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM TitulosPagBaixados WHERE NumIntDoc = ? AND Status <> ? ", 0, _
                lNumIntDoc, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 89728
        
            'Tenta ler
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89729
        
            If lErro = AD_SQL_SUCESSO Then
            
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosPagBaixados SET DataEmissao = ?", alComando(1), objTituloPagar.dtDataEmissao)
                If lErro <> AD_SQL_SUCESSO Then gError 89730
            
            'Se não encontrou --> pesquisa nos titulos baixados
            Else
        
                gError 89731
                
            End If
            
        End If
        
        'Fechamento comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    TituloPagar_Altera = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Altera:

    TituloPagar_Altera = gErr

    Select Case gErr
    
        Case 89724
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89725, 89726
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
        
        Case 89727
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAG", gErr, objTituloPagar.lNumIntDoc)
        
        Case 89728, 89729
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSPAGBAIXADOS", gErr, objTituloPagar.lNumIntDoc)
        
        Case 89730
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSPAGBAIXADOS", gErr, objTituloPagar.lNumIntDoc)
        
        Case 89731
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148990)
            
    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function TituloReceber_Altera1(ByVal objTituloReceber As ClassTituloReceber) As Long
'Lê o Titulo associado ao numintdoc e se existir altera a data de emissao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long, iNumParcelasBD As Integer

On Error GoTo Erro_TituloReceber_Altera1
    
    If objTituloReceber.lNumIntDoc <> 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 89741
        Next
        
        'Seleciona TituloReceber a partir dos dados fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumParcelas FROM TitulosRec WHERE NumIntDoc = ? And Status <> ?", 0, _
            lNumIntDoc, iNumParcelasBD, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 89742
        
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89743
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosRec SET DataEmissao = ?", alComando(1), objTituloReceber.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 89744
        
        'Se não encontrou --> pesquisa nos titulos baixados
        Else
        
            'Seleciona TituloReceberBaixado a partir dos dados fornecidos
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc, NumParcelas FROM TitulosRecBaixados WHERE NumIntDoc = ? AND Status <> ? ", 0, _
                lNumIntDoc, iNumParcelasBD, objTituloReceber.lNumIntDoc, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 89745
        
            'Tenta ler
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89746
        
            If lErro = AD_SQL_SUCESSO Then
            
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE TitulosRecBaixados SET DataEmissao = ?", alComando(1), objTituloReceber.dtDataEmissao)
                If lErro <> AD_SQL_SUCESSO Then gError 89747
            
            'Se não encontrou --> pesquisa nos titulos baixados
            Else
        
                gError 89748
                
            End If
            
        End If
        
        'Fechamento comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
        If iNumParcelasBD <> objTituloReceber.iNumParcelas Then gError 81882
            
    End If
    
    TituloReceber_Altera1 = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Altera1:

    TituloReceber_Altera1 = gErr

    Select Case gErr
    
        Case 89741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89742, 89743
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSREC1", gErr, objTituloReceber.lNumIntDoc)
        
        Case 89744
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSREC", gErr, objTituloReceber.lNumIntDoc)
        
        Case 89745, 89746
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOSRECBAIXADOS2", gErr, objTituloReceber.lNumTitulo)
        
        Case 89747
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_TITULOSRECBAIXADOS", gErr, objTituloReceber.lNumTitulo)
        
        Case 89748
        
        Case 81882
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_ALTERADO", Err, objTituloReceber.iNumParcelas, iNumParcelasBD)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148991)
            
    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function NFsPag_Altera(ByVal objNFsPag As ClassNFsPag) As Long
'Lê o Titulo associado ao numintdoc e se existir altera a data de emissao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_NFsPag_Altera
    
    If objNFsPag.lNumIntDoc <> 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 89749
        Next
        
        'Seleciona NFsPag a partir dos dados fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM NFsPag WHERE NumIntDoc = ? And Status <> ?", 0, _
            lNumIntDoc, objNFsPag.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 89750
        
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89751
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFsPag SET DataEmissao = ?", alComando(1), objNFsPag.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 89752
        
        'Se não encontrou --> pesquisa nos titulos baixados
        Else
        
            'Seleciona NFsPagBaixado a partir dos dados fornecidos
            lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM NFsPagBaixadas WHERE NumIntDoc = ? AND Status <> ? ", 0, _
                lNumIntDoc, objNFsPag.lNumIntDoc, STATUS_EXCLUIDO)
            If lErro <> AD_SQL_SUCESSO Then gError 89753
        
            'Tenta ler
            lErro = Comando_BuscarPrimeiro(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89754
        
            If lErro = AD_SQL_SUCESSO Then
            
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE NFsPagBaixadas SET DataEmissao = ?", alComando(1), objNFsPag.dtDataEmissao)
                If lErro <> AD_SQL_SUCESSO Then gError 89755
            
            'Se não encontrou --> pesquisa nos titulos baixados
            Else
        
                gError 89756
                
            End If
            
        End If
        
        'Fechamento comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    NFsPag_Altera = SUCESSO
    
    Exit Function
    
Erro_NFsPag_Altera:

    NFsPag_Altera = gErr

    Select Case gErr
    
        Case 89749
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89750, 89751
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG4", gErr, objNFsPag.lNumIntDoc)
        
        Case 89752
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFSPAG1", gErr, objNFsPag.lNumIntDoc)
        
        Case 89753, 89754
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAGBAIXADAS3", gErr, objNFsPag.lNumIntDoc)
        
        Case 89755
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NFSPAGBAIXADAS", gErr, objNFsPag.lNumIntDoc)
        
        Case 89756
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148992)
            
    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function DebitoRecCli_Altera(ByVal objDebitoRecCli As ClassDebitoRecCli) As Long
'Lê o Titulo associado ao numintdoc e se existir altera a data de emissao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_DebitoRecCli_Altera
    
    If objDebitoRecCli.lNumIntDoc <> 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 89757
        Next
        
        'Seleciona DebitoRecCli a partir dos dados fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM DebitosRecCli WHERE NumIntDoc = ? And Status <> ?", 0, _
            lNumIntDoc, objDebitoRecCli.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 89758
        
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89759
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE DebitosRecCli SET DataEmissao = ?", alComando(1), objDebitoRecCli.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 89760
        
        'Se não encontrou --> pesquisa nos titulos baixados
        Else
        
            gError 89761
                
        End If
        
        'Fechamento comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    DebitoRecCli_Altera = SUCESSO
    
    Exit Function
    
Erro_DebitoRecCli_Altera:

    DebitoRecCli_Altera = gErr

    Select Case gErr
    
        Case 89757
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89758, 89759
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DEBITOSRECCLI3", gErr, objDebitoRecCli.lNumIntDoc)
        
        Case 89760
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_DEBITOSRECCLI", gErr, objDebitoRecCli.lNumIntDoc)
        
        Case 89761
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148993)
            
    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function CreditoPagar_Altera(ByVal objCreditoPagar As ClassCreditoPagar) As Long
'Lê o Titulo associado ao numintdoc e se existir altera a data de emissao

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_CreditoPagar_Altera
    
    If objCreditoPagar.lNumIntDoc <> 0 Then
    
        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 89762
        Next
        
        'Seleciona CreditoPagar a partir dos dados fornecidos
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntDoc FROM CreditosPagForn WHERE NumIntDoc = ? And Status <> ?", 0, _
            lNumIntDoc, objCreditoPagar.lNumIntDoc, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 89763
        
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 89764
        
        If lErro = AD_SQL_SUCESSO Then
        
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE CreditosPagForn SET DataEmissao = ?", alComando(1), objCreditoPagar.dtDataEmissao)
            If lErro <> AD_SQL_SUCESSO Then gError 89765
        
        'Se não encontrou --> pesquisa nos titulos baixados
        Else
        
            gError 89766
                
        End If
        
        'Fechamento comandos
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
    
    End If
    
    CreditoPagar_Altera = SUCESSO
    
    Exit Function
    
Erro_CreditoPagar_Altera:

    CreditoPagar_Altera = gErr

    Select Case gErr
    
        Case 89762
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 89763, 89764
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CREDITOSPAGFORN1", gErr, objCreditoPagar.lNumIntDoc)
        
        Case 89765
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CREDITOSPAGFORN", gErr, objCreditoPagar.lNumIntDoc)
        
        Case 89766
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148994)
            
    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function SldMesForn_Altera_Estatisticas(ByVal objNFiscal As ClassNFiscal, ByVal ObjNFiscalBD As ClassNFiscal) As Long
'Atualiza as tabelas SldMesForn e SldDiaForn diminuindo a quantidade e o valor da data de emissao original e colocando na nova data
'Tem que ser chamado dentro de Transação

Dim lErro As Long
Dim dtDataEmissao As Date
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_SldMesForn_Altera_Estatisticas

    If objNFiscal.dtDataEmissao <> ObjNFiscalBD.dtDataEmissao Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 89129
    
        If lErro <> SUCESSO Then gError 89130

        If objTipoDocInfo.iOrigem = DOCINFO_ORIGEM_FORNECEDOR Then
        
            If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then
        
                'Cancela as estatisticas da data de emissao antiga
                lErro = CF("SldMesForn_Grava_Estatisticas", ObjNFiscalBD, CANCELAMENTO_FILIALFORNFILEMP)
                If lErro <> SUCESSO Then gError 89733
            
                dtDataEmissao = ObjNFiscalBD.dtDataEmissao
            
                ObjNFiscalBD.dtDataEmissao = objNFiscal.dtDataEmissao
            
                'Grava as estatisticas na nova data de emissao
                lErro = CF("SldMesForn_Grava_Estatisticas", ObjNFiscalBD, CADASTRAMENTO_FILIALFORNFILEMP)
                If lErro <> SUCESSO Then gError 89734
        
                ObjNFiscalBD.dtDataEmissao = dtDataEmissao

            ElseIf objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS_DEVOLUCAO Then

                'Cancela as estatisticas da data de emissao antiga
                lErro = CF("SldMesForn_Grava_Estatisticas", ObjNFiscalBD, CADASTRAMENTO_FILIALFORNFILEMP)
                If lErro <> SUCESSO Then gError 89733
            
                dtDataEmissao = ObjNFiscalBD.dtDataEmissao
            
                ObjNFiscalBD.dtDataEmissao = objNFiscal.dtDataEmissao
            
                'Grava as estatisticas na nova data de emissao
                lErro = CF("SldMesForn_Grava_Estatisticas", ObjNFiscalBD, CANCELAMENTO_FILIALFORNFILEMP)
                If lErro <> SUCESSO Then gError 89734
        
                ObjNFiscalBD.dtDataEmissao = dtDataEmissao

            End If


        End If

    End If

    SldMesForn_Altera_Estatisticas = SUCESSO
    
    Exit Function

Erro_SldMesForn_Altera_Estatisticas:

    SldMesForn_Altera_Estatisticas = gErr
    
    Select Case gErr

        Case 89733, 89734

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 148995)

    End Select
    
    Exit Function

End Function

Function CPR_Altera(ByVal objNFiscal As ClassNFiscal) As Long
'Altera a data de emissao do titulo relacionado à nota fiscal

Dim lErro As Long
Dim sTipoDocCPR As String
Dim lComando As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objTituloReceber As New ClassTituloReceber
Dim objNFsPag As New ClassNFsPag
Dim objDebitoRecCli As New ClassDebitoRecCli
Dim objCreditoPagar As New ClassCreditoPagar
Dim objParcelaReceber As ClassParcelaReceber
Dim dtDataVencimento As Date
Dim dtDataVencimentoReal As Date

On Error GoTo Erro_CPR_Altera

    Select Case objNFiscal.iClasseDocCPR
        
        Case CPR_TITULO_PAGAR
            
            objTituloPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
            objTituloPagar.dtDataEmissao = objNFiscal.dtDataEmissao
            objTituloPagar.sNatureza = objNFiscal.sNatureza
            
            lErro = TituloPagar_Altera(objTituloPagar)
            If lErro <> SUCESSO And lErro <> 89731 Then gError 89767
            
        Case CPR_TITULO_RECEBER
        
            objTituloReceber.lNumIntDoc = objNFiscal.lNumIntDocCPR
            objTituloReceber.dtDataEmissao = objNFiscal.dtDataEmissao
            objTituloReceber.iNumParcelas = objNFiscal.colParcelaReceber.Count
            objTituloReceber.sNatureza = objNFiscal.sNatureza
            
            lErro = TituloReceber_Altera1(objTituloReceber)
            If lErro <> SUCESSO And lErro <> 89748 Then gError 89768
            
            'alterado por mario em 01/10/01 para permitir alterar as datas de vencimento e data de vencimentoreal
            If objNFiscal.colParcelaReceber.Count > 0 Then
            
                For Each objParcelaReceber In objNFiscal.colParcelaReceber
                
                    dtDataVencimento = objParcelaReceber.dtDataVencimento
                    
                    lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVencimentoReal)
                    If lErro <> SUCESSO Then gError 92617
                    
                    objParcelaReceber.dtDataVencimentoReal = dtDataVencimentoReal
                
                    objParcelaReceber.lNumIntTitulo = objNFiscal.lNumIntDocCPR
                
                Next
            
                'Faz as alterações dos campos possíveis na tabela de Parcelas  a Receber
                lErro = ParcelasReceber_Altera_CamposAlteraveis(objNFiscal.colParcelaReceber, objNFiscal.colParcelaReceber.Count, True)
                If lErro <> SUCESSO Then gError 92616
            
            End If
            
        Case CPR_NF_PAGAR
        
            objNFsPag.lNumIntDoc = objNFiscal.lNumIntDocCPR
            objNFsPag.dtDataEmissao = objNFiscal.dtDataEmissao
        
            lErro = NFsPag_Altera(objNFsPag)
            If lErro <> SUCESSO And lErro <> 89756 Then gError 89769
            
        Case CPR_DEBITO_RECEBER
        
            objDebitoRecCli.lNumIntDoc = objNFiscal.lNumIntDocCPR
            objDebitoRecCli.dtDataEmissao = objNFiscal.dtDataEmissao
        
            lErro = DebitoRecCli_Altera(objDebitoRecCli)
            If lErro <> SUCESSO And lErro <> 89761 Then gError 89770
            
        Case CPR_CREDITO_PAGAR
            
            objCreditoPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
            objCreditoPagar.dtDataEmissao = objNFiscal.dtDataEmissao
        
            lErro = CreditoPagar_Altera(objCreditoPagar)
            If lErro <> SUCESSO And lErro <> 89766 Then gError 89771
            
    End Select
    
    CPR_Altera = SUCESSO
    
    Exit Function
    
Erro_CPR_Altera:

    CPR_Altera = gErr
    
    Select Case gErr
    
        Case 89767 To 89771, 92616, 92617
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148996)
            
    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ChequeParcelaRec_Exclui(ByVal lNumIntCheque As Long, ByVal lNumIntParcela As Long) As Long
'Exclui o registro referente à associacao de um Cheque Pre à uma parcela
'Se a parcela nao tiver outros cheque associados a ela, entao transfere a parcela de volta p/CARTEIRA_CARTEIRA

Dim lErro As Long
Dim alComando(0 To 4) As Long
Dim iIndice As Integer, lNumIntChequeOutro As Long
Dim lNumIntDoc As Long, bSemOutroCheque As Boolean
Dim dSaldoParcela As Double, objTransfCartCobr As New ClassTransfCartCobr
Dim iProxSeqOcorr As Integer

On Error GoTo Erro_ChequeParcelaRec_Exclui

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 17740

    Next

    'seleciona a parcela
    lErro = Comando_ExecutarPos(alComando(0), "SELECT ProxSeqOcorr, Saldo FROM ParcelasRec WHERE NumIntDoc = ?", 0, iProxSeqOcorr, dSaldoParcela, lNumIntParcela)
    If lErro <> AD_SQL_SUCESSO Then gError 17752

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17753
'    If lErro <> AD_SQL_SUCESSO Then gError 81690
    
    If lErro = AD_SQL_SUCESSO Then
    
        'bloqueia a parcela
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 81691
        
        'seleciona a associacao parcela cheque
        lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntCheque FROM ChequeParcelaRec WHERE NumIntParcela = ? AND NumIntCheque = ?", 0, lNumIntChequeOutro, lNumIntParcela, lNumIntCheque)
        If lErro <> AD_SQL_SUCESSO Then gError 91314
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 91315
        If lErro <> AD_SQL_SUCESSO Then gError 81692
        
        'pesquisar se ainda vai sobrar outro cheque associado à parcela
        lErro = Comando_Executar(alComando(2), "SELECT NumIntCheque FROM ChequeParcelaRec WHERE NumIntParcela = ? AND NumIntCheque <> ?", lNumIntChequeOutro, lNumIntParcela, lNumIntCheque)
        If lErro <> AD_SQL_SUCESSO Then gError 17752
    
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17753
        
        'se não há outro cheque associado à parcela
        If lErro = AD_SQL_SEM_DADOS Then
        
            'trazer a parcela de volta para CARTEIRA_CARTEIRA
            
            objTransfCartCobr.iCobrador = COBRADOR_PROPRIA_EMPRESA
            objTransfCartCobr.iCarteiraCobranca = CARTEIRA_CARTEIRA
            objTransfCartCobr.dtData = gdtDataAtual
            objTransfCartCobr.dtDataRegistro = gdtDataHoje
            objTransfCartCobr.iNumSeqOcorr = iProxSeqOcorr
            objTransfCartCobr.lNumIntParc = lNumIntParcela
            
            lErro = CF("TransfCartCobr_Grava1", objTransfCartCobr, COBRADOR_PROPRIA_EMPRESA, CARTEIRA_CHEQUEPRE, dSaldoParcela)
            If lErro <> SUCESSO Then gError 15582
        
            'Atualiza a tabela de parcelasrec
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE ParcelasRec SET Cobrador =?, CarteiraCobranca =?, ProxSeqOcorr =?", alComando(0), COBRADOR_PROPRIA_EMPRESA, CARTEIRA_CARTEIRA, iProxSeqOcorr + 1)
            If lErro <> AD_SQL_SUCESSO Then gError 17754
        
        End If
        
        'exclui a ligacao entre o cheque e a parcela
        lErro = Comando_ExecutarPos(alComando(4), "DELETE ChequeParcelaRec", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 91316
    
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ChequeParcelaRec_Exclui = SUCESSO

    Exit Function

Erro_ChequeParcelaRec_Exclui:

    ChequeParcelaRec_Exclui = gErr

    Select Case gErr

        Case 15582

        Case 17740
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 17752, 17753
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC_CHEQUE", gErr, lNumIntCheque)

        Case 17754
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", gErr, lNumIntCheque)

        Case 91314, 91315, 91317, 81692
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPARCELAREC", gErr)
        
        Case 91316
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUEPARCELAREC", gErr)

        Case 81690
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
        
        Case 81691
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148997)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ChequeParcelaRec_Inserir(ByVal objChequePre As ClassChequePre, ByVal colInfoParcRec As Collection) As Long
'Rotina que insere os registros correspondentes a associacao entre um cheque-pre e as parcelas a ele associadas

Dim lErro As Long, alComando(0 To 2) As Long
Dim iIndice As Integer, objTransfCartCobr As New ClassTransfCartCobr
Dim objInfoParcRec As ClassInfoParcRec, iCobrador As Integer, iCarteira As Integer, iProxSeqOcorr As Integer, lNumIntDoc As Long, dSaldo As Double

On Error GoTo Erro_ChequeParcelaRec_Inserir

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 91312
    Next

    For Each objInfoParcRec In colInfoParcRec

        'Grava o ChequePre na tabela de ChequeParcelaRec
        lErro = Comando_Executar(alComando(0), "INSERT INTO ChequeParcelaRec (NumIntCheque, NumIntParcela, ValorRecebto,Juros, Multa, Desconto) VALUES (?,?,?,?,?,?)", objChequePre.lNumIntCheque, objInfoParcRec.lNumIntParc, objInfoParcRec.dValorReceber, objInfoParcRec.dValorJuros, objInfoParcRec.dValorMulta, objInfoParcRec.dValorDesconto)
        If lErro <> AD_SQL_SUCESSO Then gError 91313
    
        'Pesquisa a parcela que será vinculada ao cheque pre
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Cobrador, CarteiraCobranca,ProxSeqOcorr, NumIntDoc, Saldo FROM ParcelasRec WHERE NumIntDoc = ?", 0, iCobrador, iCarteira, iProxSeqOcorr, lNumIntDoc, dSaldo, objInfoParcRec.lNumIntParc)
        If lErro <> AD_SQL_SUCESSO Then gError 17710

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17711
        If lErro <> AD_SQL_SUCESSO Then gError 81693
        
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 81694
        
        If Not (iCobrador = COBRADOR_PROPRIA_EMPRESA And (iCarteira = CARTEIRA_CARTEIRA Or iCarteira = CARTEIRA_CHEQUEPRE)) Then Error 19139

        'se a parcela estava em carteira
        If iCarteira = CARTEIRA_CARTEIRA Then
        
            'transferi-la p/carteira cheque_pre
            objTransfCartCobr.iCobrador = COBRADOR_PROPRIA_EMPRESA
            objTransfCartCobr.iCarteiraCobranca = CARTEIRA_CHEQUEPRE
            objTransfCartCobr.dtData = gdtDataAtual
            objTransfCartCobr.dtDataRegistro = gdtDataHoje
            objTransfCartCobr.iNumSeqOcorr = iProxSeqOcorr
            objTransfCartCobr.lNumIntParc = lNumIntDoc
            objTransfCartCobr.lNumIntDoc = 0
            
            lErro = CF("TransfCartCobr_Grava1", objTransfCartCobr, COBRADOR_PROPRIA_EMPRESA, CARTEIRA_CARTEIRA, dSaldo)
            If lErro <> SUCESSO Then gError 19140
            
            'Atualiza a tabela de parcelasrec
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE ParcelasRec SET NumIntDoc = ?, Cobrador =?, CarteiraCobranca =?, ProxSeqOcorr =?", alComando(1), objInfoParcRec.lNumIntParc, COBRADOR_PROPRIA_EMPRESA, CARTEIRA_CHEQUEPRE, iProxSeqOcorr + 1)
            If lErro <> AD_SQL_SUCESSO Then gError 17716
    
        End If
        
    Next
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ChequeParcelaRec_Inserir = SUCESSO

    Exit Function

Erro_ChequeParcelaRec_Inserir:

    ChequeParcelaRec_Inserir = gErr

    Select Case gErr
        
        Case 17710, 17711, 81693
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
        
        Case 81694
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", gErr)
        
        Case 17716
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELAS_REC2", gErr)
                
        Case 19139
            Call Rotina_Erro(vbOKOnly, "ERRO_PARC_NAO_CART_OU_CHEQUE", gErr)
                
        Case 19140
        
        Case 91312
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 91313
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CHEQUEPARCELAREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 148998)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FormacaoPreco_Grava(ByVal colFormacaoPreco As Collection) As Long
'grava a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim sItem As String
Dim iTipo As Integer
Dim dtData As Date
Dim objFormacaoPreco As ClassFormacaoPreco

On Error GoTo Erro_FormacaoPreco_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92189

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92213
    Next
    
    Set objFormacaoPreco = colFormacaoPreco.Item(1)
    
    If objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then

        sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Item FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", 0, sItem, CATEGORIA_PRODUTO_PRECO, objFormacaoPreco.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 92192

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92193

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92194
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92195
    
    ElseIf objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM Produtos WHERE Codigo = ?", 0, iTipo, objFormacaoPreco.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92196

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92197

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92198
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92199
    
    ElseIf objFormacaoPreco.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto =?", 0, dtData, giFilialEmpresa, objFormacaoPreco.iTabelaPreco, objFormacaoPreco.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92200

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92201

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92202
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92203
    
    End If
    
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Escopo FROM FormacaoPreco WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=?", 0, iEscopo, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92204
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92205
    
    'se o modelo estiver cadastrado
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 92208
    
        lErro = CF("FormacaoPreco_Exclui_Trans", objFormacaoPreco)
        If lErro <> SUCESSO Then gError 92265
    
    End If
    
    'Gerar o Número Interno da Nota Fiscal
    lErro = CF("FormacaoPreco_Automatico", lNumIntDoc, colFormacaoPreco.Count)
    If lErro <> SUCESSO Then gError 92210
    
    For Each objFormacaoPreco In colFormacaoPreco
    
        'o modelo não está cadastrado ==> inserir o modelo
        lErro = Comando_Executar(alComando(4), "INSERT INTO FormacaoPreco (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Expressao, Linha, Titulo) VALUES (?,?,?,?,?,?,?,?,?)", _
        objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco, lNumIntDoc, objFormacaoPreco.sExpressao, objFormacaoPreco.iLinha, objFormacaoPreco.sTitulo)
        If lErro <> AD_SQL_SUCESSO Then gError 92211

        lNumIntDoc = lNumIntDoc + 1

    Next
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92212

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FormacaoPreco_Grava = SUCESSO

    Exit Function

Erro_FormacaoPreco_Grava:

    FormacaoPreco_Grava = gErr

    Select Case gErr

        Case 92189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92192, 92193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", gErr, CATEGORIA_PRODUTO_PRECO, objFormacaoPreco.sItemCategoria)

        Case 92194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA", gErr, objFormacaoPreco.sItemCategoria, CATEGORIA_PRODUTO_PRECO)

        Case 92195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM1", gErr, CATEGORIA_PRODUTO_PRECO, objFormacaoPreco.sItemCategoria)

        Case 92196, 92197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objFormacaoPreco.sProduto)

        Case 92198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objFormacaoPreco.sProduto)

        Case 92199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS1", gErr, objFormacaoPreco.sProduto)

        Case 92200, 92201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS2", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iTabelaPreco, objFormacaoPreco.sProduto)

        Case 92202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objFormacaoPreco.iTabelaPreco, objFormacaoPreco.sProduto)

        Case 92203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECOITENS2", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iTabelaPreco, objFormacaoPreco.sProduto)

        Case 92204, 92205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92210, 92265

        Case 92211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)
          
        Case 92212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 92213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148999)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FormacaoPreco_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Retorna o número do próximo NumIntNF disponivel

    FormacaoPreco_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_FORMACAOPRECO", lNumIntDoc, iQtde)
    
End Function

Function FormacaoPreco_Exclui_Trans(ByVal objFormacaoPreco As ClassFormacaoPreco) As Long
'exclui a expressão de formação de preço
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_FormacaoPreco_Exclui_Trans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92225
    Next
   
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Escopo FROM FormacaoPreco WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=?", 0, iEscopo, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92226
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92227
    
    'If lErro = AD_SQL_SEM_DADOS Then gError 92228 'Comentado por Wagner (Para não dar erro ao excluir uma planilha que não tenha formação de preço)
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM FormacaoPreco", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92229
        
        'le o modelo em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92227
        
    Loop
        
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FormacaoPreco_Exclui_Trans = SUCESSO

    Exit Function

Erro_FormacaoPreco_Exclui_Trans:

    FormacaoPreco_Exclui_Trans = gErr

    Select Case gErr

        Case 92224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92226, 92227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_NAO_CADASTRADO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FORMACAOPRECO", gErr, objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco)

        Case 92230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149000)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FormacaoPreco_Exclui(ByVal objFormacaoPreco As ClassFormacaoPreco) As Long
'exclui a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_FormacaoPreco_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92224

    lErro = CF("FormacaoPreco_Exclui_Trans", objFormacaoPreco)
    If lErro <> SUCESSO Then gError 92266
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92230

    FormacaoPreco_Exclui = SUCESSO

    Exit Function

Erro_FormacaoPreco_Exclui:

    FormacaoPreco_Exclui = gErr

    Select Case gErr

        Case 92224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 92266
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149001)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MnemonicoFPreco_Grava(ByVal colMnemonicoFPreco As Collection) As Long
'grava a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim sItem As String
Dim iTipo As Integer
Dim dtData As Date
Dim objMnemonicoFPreco As ClassMnemonicoFPreco

On Error GoTo Erro_MnemonicoFPreco_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92303

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92304
    Next
    
    Set objMnemonicoFPreco = colMnemonicoFPreco.Item(1)
    
    If objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then

        sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Item FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", 0, sItem, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPreco.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 92305

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92306

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92307
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92308
    
    ElseIf objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM Produtos WHERE Codigo = ?", 0, iTipo, objMnemonicoFPreco.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92309

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92310

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92311
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92312
    
    ElseIf objMnemonicoFPreco.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto =?", 0, dtData, giFilialEmpresa, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92313

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92314

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92315
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92316
    
    End If
    
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Escopo FROM MnemonicoFPreco WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=?", 0, iEscopo, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92317
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92318
    
    'se o modelo estiver cadastrado
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 92319
    
        lErro = CF("MnemonicoFPreco_Exclui_Trans", objMnemonicoFPreco)
        If lErro <> SUCESSO Then gError 92320
    
    End If
    
    'Gerar o Número Interno da Nota Fiscal
    lErro = CF("MnemonicoFPreco_Automatico", lNumIntDoc, colMnemonicoFPreco.Count)
    If lErro <> SUCESSO Then gError 92321
    
    For Each objMnemonicoFPreco In colMnemonicoFPreco
    
        'o modelo não está cadastrado ==> inserir o modelo
        lErro = Comando_Executar(alComando(4), "INSERT INTO MnemonicoFPreco (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Mnemonico, MnemonicoDesc, Expressao, Tipo, NumParam) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
        objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco, lNumIntDoc, objMnemonicoFPreco.sMnemonico, objMnemonicoFPreco.sMnemonicoDesc, objMnemonicoFPreco.sExpressao, TIPO_NUMERICO, 0)
        If lErro <> AD_SQL_SUCESSO Then gError 92322

        lNumIntDoc = lNumIntDoc + 1

    Next
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92323

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MnemonicoFPreco_Grava = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Grava:

    MnemonicoFPreco_Grava = gErr

    Select Case gErr

        Case 92303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92305, 92306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", gErr, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPreco.sItemCategoria)

        Case 92307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA", gErr, objMnemonicoFPreco.sItemCategoria, CATEGORIA_PRODUTO_PRECO)

        Case 92308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM1", gErr, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPreco.sItemCategoria)

        Case 92309, 92310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objMnemonicoFPreco.sProduto)

        Case 92311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objMnemonicoFPreco.sProduto)

        Case 92312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS1", gErr, objMnemonicoFPreco.sProduto)

        Case 92313, 92314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS2", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.sProduto)

        Case 92315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.sProduto)

        Case 92316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECOITENS2", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.sProduto)

        Case 92317, 92318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPRECO2", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)

        Case 92319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MNEMONICOFPRECO", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)

        Case 92320, 92321

        Case 92322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MNEMONICOFPRECO", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)
          
        Case 92323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149002)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MnemonicoFPreco_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Retorna o número do próximo NumIntNF disponivel

    MnemonicoFPreco_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_MNEMONICOFPRECO", lNumIntDoc, iQtde)
    
End Function

Function MnemonicoFPreco_Exclui_Trans(ByVal objMnemonicoFPreco As ClassMnemonicoFPreco) As Long
'exclui a expressão de formação de preço
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_MnemonicoFPreco_Exclui_Trans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92324
    Next

    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Escopo FROM MnemonicoFPreco WHERE Funcao = 0 AND FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=?", 0, iEscopo, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)
    If lErro <> AD_SQL_SUCESSO Then gError 92325

    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92326

    'If lErro = AD_SQL_SEM_DADOS Then gError 92327 'Comentado por Wagner para não dar mensagem ao excluir uma planilha que não tenha formação de preço

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM MnemonicoFPreco", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92328

        'le o modelo em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92329

    Loop

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MnemonicoFPreco_Exclui_Trans = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Exclui_Trans:

    MnemonicoFPreco_Exclui_Trans = gErr

    Select Case gErr

        Case 92324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92325, 92326, 92329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPRECO2", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)

        Case 92327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOFPRECO_NAO_CADASTRADO", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)

        Case 92328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MNEMONICOFPRECO", gErr, objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149003)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function MnemonicoFPreco_Exclui(ByVal objMnemonicoFPreco As ClassMnemonicoFPreco) As Long
'exclui a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_MnemonicoFPreco_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92330

    lErro = CF("MnemonicoFPreco_Exclui_Trans", objMnemonicoFPreco)
    If lErro <> SUCESSO Then gError 92331

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92332

    MnemonicoFPreco_Exclui = SUCESSO

    Exit Function

Erro_MnemonicoFPreco_Exclui:

    MnemonicoFPreco_Exclui = gErr

    Select Case gErr

        Case 92330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92331

        Case 92332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149004)

    End Select

    Call Transacao_Rollback

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function


Function ConhecimentoFrete_Altera(ByVal objNFiscal As ClassNFiscal) As Long
'Atualiza o campo valor mercadoria

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lNumIntNFiscal As Long
Dim iIndice As Integer

On Error GoTo Erro_ConhecimentoFrete_Altera

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92625

    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntNFiscal FROM ConhecimentoFrete WHERE NumIntNFiscal = ?", 0, lNumIntNFiscal, objNFiscal.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 92626

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92627

    If lErro = AD_SQL_SEM_DADOS Then gError 92628
    
    'alterado por mario em 01/10/01 para permitir alterar os campos Origem, Destino, ValorContainer, ValorMercadoria e VolumeEspecie
    lErro = Comando_ExecutarPos(alComando(2), "UPDATE ConhecimentoFrete SET ValorMercadoria = ?", alComando(1), objNFiscal.objConhecimentoFrete.dValorMercadoria)
    If lErro <> AD_SQL_SUCESSO Then gError 92629
    
    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ConhecimentoFrete_Altera = SUCESSO

    Exit Function

Erro_ConhecimentoFrete_Altera:

    ConhecimentoFrete_Altera = gErr

    Select Case gErr

        Case 92625
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92626, 92627
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONHECIMENTO_FRETE", gErr)

        Case 92628
            Call Rotina_Erro(vbOKOnly, "ERRO_CONHECIMENTOFRETE_NAO_CADASTRADO", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 92629
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CONHECIMENTOFRETE", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.iFilialEmpresa)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149005)

    End Select

    'Fechamento comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ClienteHistorico_Grava(objClienteHistorico As ClassClienteHistorico) As Long
'Grava o Historico do Cliente a partir dos dados passados no objClienteHistorico

Dim lErro As Long
Dim iSituacao As Integer
Dim lComando As Long

On Error GoTo Erro_ClienteHistorico_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97600
    
    'Inicializa a Situacao
    iSituacao = -1
        
    'Verifica se é um novo registro ou uma alteracao dos campos da ClienteHistorico
    lErro = ClienteHistorico_Verifica_Alteracao(objClienteHistorico, iSituacao)
    If lErro <> SUCESSO Then gError 97601
    
    'Se foi alterado algum campo válido (vide ClienteHistorico) ou é um novo reg. => GRAVA
    If ((iSituacao = REGISTRO_INEXISTENTE) Or (iSituacao = REGISTRO_ALTERADO)) Then
    
        'Faz a Gravacao na Tabela de ClienteHistorico
        lErro = ClienteHistorico_Grava1(objClienteHistorico)
        If lErro <> SUCESSO Then gError 97602
            
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    ClienteHistorico_Grava = SUCESSO
    
    
    Exit Function

Erro_ClienteHistorico_Grava:

    ClienteHistorico_Grava = gErr
    
    Select Case gErr
    
        Case 97601, 97602
        
        Case 97600
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149006)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteHistorico_Verifica_Alteracao(objClienteHistorico As ClassClienteHistorico, iSituacao As Integer) As Long
'Verifica se está sendo feita alguma alteracao nos campos presentes da tabela ClienteHistorico

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tClienteHistorico As typeClienteHistorico
Dim iCodMensagem As Integer
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_ClienteHistorico_Verifica_Alteracao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97603

    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Clientes.Codigo, FiliaisClientes.CGC, FiliaisClientes.InscricaoEstadual,FiliaisClientes.InscricaoSuframa, FiliaisClientes.InscricaoMunicipal, Clientes.RazaoSocial, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Paises.Nome, Enderecos.CEP, Clientes.Ativo, Clientes.CodigoLoja, Clientes.FilialEmpresaLoja, " & _
                    "Clientes.NomeReduzido,Clientes.Tipo,Clientes.LimiteCredito,Clientes.Observacao,Clientes.CondicaoPagto,Clientes.TabelaPreco,Clientes.UsuarioCobrador, " & _
                    "Clientes.UsuRespCallCenter,FiliaisClientes.CodTransportadora,FiliaisClientes.Vendedor,FiliaisClientes.RG,FiliaisClientes.RegimeTributario, " & _
                    "Enderecos.Telefone1,Enderecos.Telefone2,Enderecos.Email,Enderecos.Email2, Clientes.CodMensagem, Clientes.Desconto, FiliaisClientes.ComissaoVendas, FiliaisClientes.Regiao FROM Clientes, FiliaisClientes, Enderecos, Paises " & _
                "WHERE Clientes.Codigo = FiliaisClientes.CodCliente AND FiliaisClientes.Endereco = Enderecos.Codigo AND Enderecos.CodigoPais = Paises.Codigo AND Clientes.Codigo=? AND FiliaisClientes.CodFilial=?"
    
    With tClienteHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        .sNomeReduzido = String(STRING_MAXIMO, 0)
        .sObservacao = String(STRING_MAXIMO, 0)
        .sUsuarioCobrador = String(STRING_MAXIMO, 0)
        .sUsuRespCallCenter = String(STRING_MAXIMO, 0)
        .sTelefone1 = String(STRING_MAXIMO, 0)
        .sTelefone2 = String(STRING_MAXIMO, 0)
        .sEmail = String(STRING_MAXIMO, 0)
        .sEmail2 = String(STRING_MAXIMO, 0)
        .sRG = String(STRING_MAXIMO, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .lCodigo, .sCgc, .sInscricaoEstadual, .sInscricaoSuframa, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, .iAtivo, .lCodigoLoja, .iFilialEmpresaLoja, _
            .sNomeReduzido, .iTipo, .dLimiteCredito, .sObservacao, .iCondicaoPagto, .iTabelaPreco, .sUsuarioCobrador, .sUsuRespCallCenter, .iCodTransportadora, .iVendedor, .sRG, .iRegimeTributario, .sTelefone1, .sTelefone2, .sEmail, .sEmail2, iCodMensagem, .dDesconto, .dComissaoVendas, .iRegiao, objClienteHistorico.lCodigo, objClienteHistorico.iFilialCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 97604
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97605
    
        'Se existe o registro => Verificar se houve alteracao considerável
        If lErro <> AD_SQL_SEM_DADOS Then
        
            If iCodMensagem <> 0 Then
                objMensagem.iCodigo = iCodMensagem
                'Lê a Mensagem "default" p/ o Cliente
                lErro = CF("Mensagem_Le", objMensagem)
                If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
                
                .sMensagemNF = objMensagem.sDescricao
            End If
        
            'Verifica se houve alguma alteracao nos dados já existentes
            Call ClienteHistorico_Verifica_Alteracao1(objClienteHistorico, tClienteHistorico, iSituacao)
               
        ElseIf lErro = AD_SQL_SEM_DADOS Then
            
            iSituacao = REGISTRO_INEXISTENTE
            
        End If
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    ClienteHistorico_Verifica_Alteracao = SUCESSO
    
    Exit Function

Erro_ClienteHistorico_Verifica_Alteracao:

    ClienteHistorico_Verifica_Alteracao = gErr
    
    Select Case gErr
    
        Case 97604, 97605
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)
        
        Case 97603
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149007)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Private Sub ClienteHistorico_Verifica_Alteracao1(objClienteHistorico As ClassClienteHistorico, tClienteHistorico As typeClienteHistorico, iSituacao As Integer)
'Verifica se houve alguma alteracao nos dados já existentes

On Error GoTo Erro_ClienteHistorico_Verifica_Alteracao1

    'Se houver alguma alteração => iSituacao = REGISTRO_ALTERADO
    With objClienteHistorico
    
        If Not String_Igual(.sCgc, tClienteHistorico.sCgc) Then gError 97606
        If Not String_Igual(.sInscricaoEstadual, tClienteHistorico.sInscricaoEstadual) Then gError 97606
        If Not String_Igual(.sInscricaoMunicipal, tClienteHistorico.sInscricaoMunicipal) Then gError 97606
        If Not String_Igual(.sInscricaoSuframa, tClienteHistorico.sInscricaoSuframa) Then gError 97606
        If Not String_Igual(.sRazaoSocial, tClienteHistorico.sRazaoSocial) Then gError 97606
        If Not String_Igual(.sEndereco, tClienteHistorico.sEndereco) Then gError 97606
        If Not String_Igual(.sBairro, tClienteHistorico.sBairro) Then gError 97606
        If Not String_Igual(.sCidade, tClienteHistorico.sCidade) Then gError 97606
        If Not String_Igual(.sUF, tClienteHistorico.sUF) Then gError 97606
        If Not String_Igual(.sPais, tClienteHistorico.sPais) Then gError 97606
        If Not String_Igual(.sCEP, tClienteHistorico.sCEP) Then gError 97606
        If .iAtivo <> tClienteHistorico.iAtivo Then gError 97606
        If .iFilialEmpresaLoja <> tClienteHistorico.iFilialEmpresaLoja Then gError 97606
        If .lCodigoLoja <> tClienteHistorico.lCodigoLoja Then gError 97606
        
        If Not String_Igual(.sNomeReduzido, tClienteHistorico.sNomeReduzido) Then gError 97606
        If Not String_Igual(.sObservacao, tClienteHistorico.sObservacao) Then gError 97606
        If Not String_Igual(.sUsuarioCobrador, tClienteHistorico.sUsuarioCobrador) Then gError 97606
        If Not String_Igual(.sUsuRespCallCenter, tClienteHistorico.sUsuRespCallCenter) Then gError 97606
        If Not String_Igual(.sTelefone1, tClienteHistorico.sTelefone1) Then gError 97606
        If Not String_Igual(.sTelefone2, tClienteHistorico.sTelefone2) Then gError 97606
        If Not String_Igual(.sEmail, tClienteHistorico.sEmail) Then gError 97606
        If Not String_Igual(.sEmail2, tClienteHistorico.sEmail2) Then gError 97606
        If Not String_Igual(.sMensagemNF, tClienteHistorico.sMensagemNF) Then gError 97606
        If .iTipo <> tClienteHistorico.iTipo Then gError 97606
        If Abs(.dLimiteCredito - tClienteHistorico.dLimiteCredito) > DELTA_VALORMONETARIO Then gError 97606
        If .iCondicaoPagto <> tClienteHistorico.iCondicaoPagto Then gError 97606
        If .iTabelaPreco <> tClienteHistorico.iTabelaPreco Then gError 97606
        If .iCodTransportadora <> tClienteHistorico.iCodTransportadora Then gError 97606
        If .iVendedor <> tClienteHistorico.iVendedor Then gError 97606
        If .iRegimeTributario <> tClienteHistorico.iRegimeTributario Then gError 97606

        If Abs(.dComissaoVendas - tClienteHistorico.dComissaoVendas) > DELTA_VALORMONETARIO Then gError 97606
        If Abs(.dDesconto - tClienteHistorico.dDesconto) > DELTA_VALORMONETARIO Then gError 97606
        If .iRegiao <> tClienteHistorico.iRegiao Then gError 97606
        
        iSituacao = REGISTRO_INALTERADO
            
    End With
    
    Exit Sub

Erro_ClienteHistorico_Verifica_Alteracao1:

    Select Case gErr
        
        Case 97606
            iSituacao = REGISTRO_ALTERADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149008)
            
    End Select
    
    Exit Sub

End Sub

Function ClienteHistorico_Grava1(objClienteHistorico As ClassClienteHistorico) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_ClienteHistorico_Grava1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97607
    
    'Busca o Novo Codigo para insercao
    lErro = CF("Config_ObterAutomatico_EmTrans", "CPRConfig", "NUM_PROX_CLIENTEHISTORICO", "ClienteHistorico", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97608
    
    With objClienteHistorico
    
        .lNumIntDoc = lNumIntDoc
        
        'Guarda a expressão SQL
        sInsert = "INSERT INTO ClienteHistorico (NumIntDoc, DataAtualizacao, Codigo, Filial, CPF_CGC, InscricaoEstadual, InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, CEP, Ativo, CodigoLoja, FilialEmpresaLoja,InscricaoSuframa, Usuario, DataReg, HoraReg,NomeReduzido,Tipo,LimiteCredito,Observacao,CondicaoPagto,TabelaPreco,UsuarioCobrador,UsuRespCallCenter,CodTransportadora,Vendedor,RG,RegimeTributario,Telefone1,Telefone2,Email,Email2,MensagemNF,VersaoControle,Desconto,ComissaoVendas,Regiao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
                
        'Grava historico na tabela de ClienteHistorico
        lErro = Comando_Executar(lComando, sInsert, .lNumIntDoc, .dtDataAtualizacao, .lCodigo, .iFilialCliente, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, left(.sRazaoSocial, 40), left(.sEndereco, 40), left(.sBairro, 12), left(.sCidade, 15), .sUF, .sPais, .sCEP, .iAtivo, .lCodigoLoja, .iFilialEmpresaLoja, .sInscricaoSuframa, gsUsuario, Date, CDbl(Time), .sNomeReduzido, .iTipo, .dLimiteCredito, .sObservacao, .iCondicaoPagto, .iTabelaPreco, .sUsuarioCobrador, .sUsuRespCallCenter, .iCodTransportadora, .iVendedor, .sRG, .iRegimeTributario, .sTelefone1, .sTelefone2, .sEmail, .sEmail2, .sMensagemNF, 1, .dDesconto, .dComissaoVendas, .iRegiao)
        If lErro <> AD_SQL_SUCESSO Then gError 97609
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    ClienteHistorico_Grava1 = SUCESSO
    
    Exit Function

Erro_ClienteHistorico_Grava1:

    ClienteHistorico_Grava1 = gErr
    
    Select Case gErr
    
        Case 97608
        
        Case 97609
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLIENTEHISTORICO", gErr)
            
        Case 97607
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149009)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Public Function ClienteHistorico_Grava_Retroativos(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Clientes já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela ClienteHistorico como data de gravação

'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasCRFAT.ClassCRFATGrava"),"ClienteHistorico_Grava_Retroativos",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ClienteHistorico_Grava_Retroativos

    'Abre o Comando
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94895
    
    'Chama a função de gravação
    lErro = CF("ClienteHistorico_Grava_Retroativos_EmTrans", dtDataAtualizacao)
    If lErro <> SUCESSO Then gError 94896
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 94897
    
    ClienteHistorico_Grava_Retroativos = SUCESSO
    
    Exit Function
    
Erro_ClienteHistorico_Grava_Retroativos:

    ClienteHistorico_Grava_Retroativos = gErr

    Select Case gErr
        
        Case 94896
        
        Case 94885
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94897
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149010)
    
    End Select

    'Faz o rollback da transação
    Call Transacao_Rollback
        
End Function

Public Function ClienteHistorico_Grava_Retroativos_EmTrans(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Clientes já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela ClienteHistorico como data de gravação
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tClienteHistorico As typeClienteHistorico
Dim objClienteHistorico As New ClassClienteHistorico

On Error GoTo Erro_ClienteHistorico_Grava_Retroativos_EmTrans

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97610
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Clientes.Codigo, FiliaisClientes.CodFilial , FiliaisClientes.CGC, FiliaisClientes.InscricaoEstadual, FiliaisClientes.InscricaoMunicipal, Clientes.RazaoSocial, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Paises.Nome, Enderecos.CEP, Clientes.Ativo, Clientes.CodigoLoja, Clientes.FilialEmpresaLoja FROM Clientes, FiliaisClientes, Enderecos, Paises WHERE Clientes.Codigo = FiliaisClientes.CodCliente AND FiliaisClientes.Endereco = Enderecos.Codigo AND Enderecos.CodigoPais = Paises.Codigo"
    
    With tClienteHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .lCodigo, .iFilialCliente, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, .iAtivo, .lCodigoLoja, .iFilialEmpresaLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 97612
    
    End With
        
    'Busca o Primeiro Registro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97613

    'Enquanto tem cliente => grava o registro
    Do While lErro <> AD_SQL_SEM_DADOS
                                    
        'Atualiza os dados para futura gravacao de histórico
        With objClienteHistorico
    
            .dtDataAtualizacao = dtDataAtualizacao
            .iFilialCliente = tClienteHistorico.iFilialCliente
            .lCodigo = tClienteHistorico.lCodigo
            .sBairro = tClienteHistorico.sBairro
            .sCEP = tClienteHistorico.sCEP
            .sCgc = tClienteHistorico.sCgc
            .sCidade = tClienteHistorico.sCidade
            .sEndereco = tClienteHistorico.sEndereco
            .sInscricaoEstadual = tClienteHistorico.sInscricaoEstadual
            .sInscricaoMunicipal = tClienteHistorico.sInscricaoMunicipal
            .sUF = tClienteHistorico.sUF
            .sRazaoSocial = tClienteHistorico.sRazaoSocial
            .sPais = tClienteHistorico.sPais
            .iAtivo = tClienteHistorico.iAtivo
            .lCodigoLoja = tClienteHistorico.lCodigoLoja
            .iFilialEmpresaLoja = tClienteHistorico.iFilialEmpresaLoja
            
        End With
    
        'FAZ A GRAVACAO DO HISTORICO
        lErro = ClienteHistorico_Grava1(objClienteHistorico)
        If lErro <> SUCESSO Then gError 97614
    
        'Busca o Próximo Registro
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97615
        
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    ClienteHistorico_Grava_Retroativos_EmTrans = SUCESSO
    
    Exit Function
    
Erro_ClienteHistorico_Grava_Retroativos_EmTrans:

    ClienteHistorico_Grava_Retroativos_EmTrans = gErr

    Select Case gErr
        
        Case 97610
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97612, 97613, 97615
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTEHISTORICO", gErr)
            
        Case 97614
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149011)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
End Function

Private Function NaturezaOPHistorico_Grava(objNaturezaOPHistorico As ClassNaturezaOPHistorico) As Long
'Grava o Historico do NaturezaOP a partir dos dados passados no objNaturezaOPHistorico

Dim lErro As Long
Dim iSituacao As Integer
Dim lComando As Long

On Error GoTo Erro_NaturezaOPHistorico_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97686
    
    'Inicializa a Situacao
    iSituacao = -1
        
    'Verifica se é um novo registro ou uma alteracao dos campos da NaturezaOPHistorico
    lErro = NaturezaOPHistorico_Verifica_Alteracao(objNaturezaOPHistorico, iSituacao)
    If lErro <> SUCESSO Then gError 97687
    
    'Se foi alterado algum campo válido (vide NaturezaOPHistorico) ou é um novo reg. => GRAVA
    If ((iSituacao = REGISTRO_INEXISTENTE) Or (iSituacao = REGISTRO_ALTERADO)) Then
    
        'Faz a Gravacao na Tabela de NaturezaOPHistorico
        lErro = NaturezaOPHistorico_Grava1(objNaturezaOPHistorico)
        If lErro <> SUCESSO Then gError 97688
            
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    NaturezaOPHistorico_Grava = SUCESSO
    
    Exit Function

Erro_NaturezaOPHistorico_Grava:

    NaturezaOPHistorico_Grava = gErr
    
    Select Case gErr
    
        Case 97687, 97688
        
        Case 97686
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149012)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function NaturezaOPHistorico_Verifica_Alteracao(objNaturezaOPHistorico As ClassNaturezaOPHistorico, iSituacao As Integer) As Long
'Verifica se está sendo feita alguma alteracao nos campos presentes da tabela NaturezaOPHistorico

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tNaturezaOPHistorico As typeNaturezaOPHistorico

On Error GoTo Erro_NaturezaOPHistorico_Verifica_Alteracao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97689

    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT NaturezaOP.Codigo, NaturezaOP.DescrNF FROM NaturezaOP WHERE NaturezaOP.Codigo=?"
    
    With tNaturezaOPHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)
        .sDescricao = String(STRING_NATUREZAOP_DESCRNF, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigo, .sDescricao, objNaturezaOPHistorico.sCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 97690
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97691
    
        'Se existe o registro => Verificar se houve alteracao considerável
        If lErro <> AD_SQL_SEM_DADOS Then
        
            'Verifica se houve alguma alteracao nos dados já existentes
            Call NaturezaOPHistorico_Verifica_Alteracao1(objNaturezaOPHistorico, tNaturezaOPHistorico, iSituacao)
               
        ElseIf lErro = AD_SQL_SEM_DADOS Then
            
            iSituacao = REGISTRO_INEXISTENTE
            
        End If
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    NaturezaOPHistorico_Verifica_Alteracao = SUCESSO
    
    Exit Function

Erro_NaturezaOPHistorico_Verifica_Alteracao:

    NaturezaOPHistorico_Verifica_Alteracao = gErr
    
    Select Case gErr
    
        Case 97690, 97691
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOPHISTORICO", gErr)
        
        Case 97689
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149013)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Private Sub NaturezaOPHistorico_Verifica_Alteracao1(objNaturezaOPHistorico As ClassNaturezaOPHistorico, tNaturezaOPHistorico As typeNaturezaOPHistorico, iSituacao As Integer)
'Verifica se houve alguma alteracao nos dados já existentes

On Error GoTo Erro_NaturezaOPHistorico_Verifica_Alteracao1

    'Se houver alguma alteração => iSituacao = REGISTRO_ALTERADO
    With objNaturezaOPHistorico
    
        If .sDescricao <> tNaturezaOPHistorico.sDescricao Then gError 97692
        
        iSituacao = REGISTRO_INALTERADO
            
    End With
    
    Exit Sub

Erro_NaturezaOPHistorico_Verifica_Alteracao1:

    Select Case gErr
        
        Case 97692
            iSituacao = REGISTRO_ALTERADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149014)
            
    End Select
    
    Exit Sub

End Sub

Function NaturezaOPHistorico_Grava1(objNaturezaOPHistorico As ClassNaturezaOPHistorico) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_NaturezaOPHistorico_Grava1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97693
    
    'Busca o Novo Codigo para insercao
    lErro = CF("Config_ObterAutomatico_EmTrans", "CRFATConfig", "NUM_PROX_NATUREZAOPHISTORICO", "NaturezaOPHistorico", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97694
    
    With objNaturezaOPHistorico
    
        .lNumIntDoc = lNumIntDoc
        
        'Guarda a expressão SQL
        sInsert = "INSERT INTO NaturezaOPHistorico (NumIntDoc, DataAtualizacao, CodNaturezaOP, DescNaturezaOP) VALUES (?,?,?,?)"
                
        'Grava historico na tabela de NaturezaOPHistorico
        lErro = Comando_Executar(lComando, sInsert, .lNumIntDoc, .dtDataAtualizacao, .sCodigo, .sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 97695
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    NaturezaOPHistorico_Grava1 = SUCESSO
    
    Exit Function

Erro_NaturezaOPHistorico_Grava1:

    NaturezaOPHistorico_Grava1 = gErr
    
    Select Case gErr
    
        Case 97694
        
        Case 97695
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NATUREZAOPHISTORICO", gErr)
            
        Case 97693
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149015)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Public Function NaturezaOPHistorico_Grava_Retroativos(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de NaturezaOP já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela NaturezaOPHistorico como data de gravação

'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasCRFAT.ClassCRFATGrava"),"NaturezaOPHistorico_Grava_Retroativos",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_NaturezaOPHistorico_Grava_Retroativos

    'Abre o Comando
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94898
        
    'Chama a função que faz a gravação
    lErro = CF("NaturezaOPHistorico_Grava_Retroativos_EmTrans ", dtDataAtualizacao)
    If lErro <> SUCESSO Then gError 94899
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 94900
    
    NaturezaOPHistorico_Grava_Retroativos = SUCESSO
    
    Exit Function
    
Erro_NaturezaOPHistorico_Grava_Retroativos:

    NaturezaOPHistorico_Grava_Retroativos = gErr

    Select Case gErr
        
        Case 94899
        
        Case 94898
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94900
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149016)
    
    End Select

    'Faz o rollback da transação
    Call Transacao_Rollback
        
End Function

Public Function NaturezaOPHistorico_Grava_Retroativos_EmTrans(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de NaturezaOP já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela NaturezaOPHistorico como data de gravação
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tNaturezaOPHistorico As typeNaturezaOPHistorico
Dim objNaturezaOPHistorico As New ClassNaturezaOPHistorico

On Error GoTo Erro_NaturezaOPHistorico_Grava_Retroativos_EmTrans

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97696
        
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT NaturezaOP.Codigo, NaturezaOP.DescrNF FROM NaturezaOP"
    
    With tNaturezaOPHistorico
        
        'Inicializa as Strings do Type
        .sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)
        .sDescricao = String(STRING_NATUREZAOP_DESCRNF, 0)
        
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .sCodigo, .sDescricao)
        If lErro <> AD_SQL_SUCESSO Then gError 97698
    
    End With
        
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97699
    
        'Enquanto tem NaturezaOP => grava o registro
        Do While lErro <> AD_SQL_SEM_DADOS
                                        
            'Atualiza os dados para futura gravacao de histórico
            With objNaturezaOPHistorico
        
                .dtDataAtualizacao = dtDataAtualizacao
                .sCodigo = tNaturezaOPHistorico.sCodigo
                .sDescricao = tNaturezaOPHistorico.sDescricao
                
            End With
        
            'FAZ A GRAVACAO DO HISTORICO
            lErro = NaturezaOPHistorico_Grava1(objNaturezaOPHistorico)
            If lErro <> SUCESSO Then gError 97700
        
            'Busca o Próximo Registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97701
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
    NaturezaOPHistorico_Grava_Retroativos_EmTrans = SUCESSO
    
    Exit Function
    
Erro_NaturezaOPHistorico_Grava_Retroativos_EmTrans:

    NaturezaOPHistorico_Grava_Retroativos_EmTrans = gErr

    Select Case gErr
        
        Case 97696
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97698, 97699, 97701
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOPHISTORICO", gErr)
            
        Case 97700
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149017)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)
        
End Function

Function ParcelaRecBaixada_Le_DataBaixa(ByVal lNumIntParcela As Long, dtDataBaixa As Date, iMotivoBaixa As Integer) As Long
Attribute ParcelaRecBaixada_Le_DataBaixa.VB_Description = "Função de leitura que retorna a Data da Baixa (em dtDataBaixa) e o Motivo da Baixa (em iMotivoBaixa) da parcela baixada que possui o número interno passado (em lNumIntParcela). Caso a parcela passada não tenha sido encontrada entre as baixadas o código de erro retonado pela função é o 86189"
'Lê a data da baixa para a parcela baixada com o NumIntDoc passado

Dim lErro As Long
Dim lComando As Long
Dim sComando As String
Dim dtData As Date
Dim iMotivo As Integer

On Error GoTo Erro_ParcelaRecBaixada_Le_DataBaixa

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 86186

    dtData = DATA_NULA
    
    sComando = "SELECT BaixasRec.Data, BaixasRec.Motivo FROM BaixasRec, BaixasParcRec WHERE BaixasRec.NumIntBaixa = BaixasParcRec.NumIntBaixa AND NumIntParcela = ? ORDER BY sequencial DESC "

    lErro = Comando_Executar(lComando, sComando, dtData, iMotivo, lNumIntParcela)
    If lErro <> AD_SQL_SUCESSO Then gError 86187

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 86188

    dtDataBaixa = dtData
    
    If lErro = AD_SQL_SEM_DADOS Then gError 86189
        
    Call Comando_Fechar(lComando)

    ParcelaRecBaixada_Le_DataBaixa = SUCESSO

    Exit Function

Erro_ParcelaRecBaixada_Le_DataBaixa:

    ParcelaRecBaixada_Le_DataBaixa = gErr

    Select Case gErr

        Case 86186
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 86187, 86188
           Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_BAIXASPARCREC", gErr)

        Case 86189


        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149018)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function TranspHistorico_Grava(objTranspHistorico As ClassTranspHistorico) As Long
'Grava o Historico do Transp a partir dos dados passados no objTranspHistorico

Dim lErro As Long
Dim iSituacao As Integer
Dim lComando As Long

On Error GoTo Erro_TranspHistorico_Grava

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97711
    
    'Inicializa a Situacao
    iSituacao = -1
        
    'Verifica se é um novo registro ou uma alteracao dos campos da TranspHistorico
    lErro = TranspHistorico_Verifica_Alteracao(objTranspHistorico, iSituacao)
    If lErro <> SUCESSO Then gError 97712
    
    'Se foi alterado algum campo válido (vide TranspHistorico) ou é um novo reg. => GRAVA
    If ((iSituacao = REGISTRO_INEXISTENTE) Or (iSituacao = REGISTRO_ALTERADO)) Then
    
        'Faz a Gravacao na Tabela de TranspHistorico
        lErro = TranspHistorico_Grava1(objTranspHistorico)
        If lErro <> SUCESSO Then gError 97713
            
    End If

    'Fecha o comando
    Call Comando_Fechar(lComando)
    
    TranspHistorico_Grava = SUCESSO
    
    Exit Function

Erro_TranspHistorico_Grava:

    TranspHistorico_Grava = gErr
    
    Select Case gErr
    
        Case 97712, 97713
        
        Case 97711
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149019)
    
    End Select
    
    'Fecha o comando
    Call Comando_Fechar(lComando)

End Function

Private Function TranspHistorico_Verifica_Alteracao(objTranspHistorico As ClassTranspHistorico, iSituacao As Integer) As Long
'Verifica se está sendo feita alguma alteracao nos campos presentes da tabela TranspHistorico

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tTranspHistorico As typeTranspHistorico

On Error GoTo Erro_TranspHistorico_Verifica_Alteracao

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97714

    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Transportadoras.Codigo, Transportadoras.CGC, Transportadoras.InscricaoEstadual, Transportadoras.InscricaoMunicipal, Transportadoras.Nome, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Paises.Nome, Enderecos.CEP FROM Transportadoras, Enderecos, Paises WHERE Transportadoras.Endereco = Enderecos.Codigo AND Enderecos.CodigoPais = Paises.Codigo AND Transportadoras.Codigo = ?"
    
    With tTranspHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_TRANSPORTADORA_NOME, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .iCodTransp, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP, objTranspHistorico.iCodTransp)
        If lErro <> AD_SQL_SUCESSO Then gError 97715
    
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97716
    
        'Se existe o registro => Verificar se houve alteracao considerável
        If lErro <> AD_SQL_SEM_DADOS Then
        
            'Verifica se houve alguma alteracao nos dados já existentes
            Call TranspHistorico_Verifica_Alteracao1(objTranspHistorico, tTranspHistorico, iSituacao)
               
        ElseIf lErro = AD_SQL_SEM_DADOS Then
            
            iSituacao = REGISTRO_INEXISTENTE
            
        End If
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    TranspHistorico_Verifica_Alteracao = SUCESSO
    
    Exit Function

Erro_TranspHistorico_Verifica_Alteracao:

    TranspHistorico_Verifica_Alteracao = gErr
    
    Select Case gErr
    
        Case 97715, 97716
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORAHISTORICO", gErr)
        
        Case 97714
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149020)
            
    End Select
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Private Sub TranspHistorico_Verifica_Alteracao1(objTranspHistorico As ClassTranspHistorico, tTranspHistorico As typeTranspHistorico, iSituacao As Integer)
'Verifica se houve alguma alteracao nos dados já existentes

On Error GoTo Erro_TranspHistorico_Verifica_Alteracao1

    'Se houver alguma alteração => iSituacao = REGISTRO_ALTERADO
    With objTranspHistorico
    
        If .sCgc <> tTranspHistorico.sCgc Then gError 97717
        If .sInscricaoEstadual <> tTranspHistorico.sInscricaoEstadual Then gError 97717
        If .sInscricaoMunicipal <> tTranspHistorico.sInscricaoMunicipal Then gError 97717
        If .sRazaoSocial <> tTranspHistorico.sRazaoSocial Then gError 97717
        If .sEndereco <> tTranspHistorico.sEndereco Then gError 97717
        If .sBairro <> tTranspHistorico.sBairro Then gError 97717
        If .sCidade <> tTranspHistorico.sCidade Then gError 97717
        If .sUF <> tTranspHistorico.sUF Then gError 97717
        If .sPais <> tTranspHistorico.sPais Then gError 97717
        If .sCEP <> tTranspHistorico.sCEP Then gError 97717
        
        iSituacao = REGISTRO_INALTERADO
            
    End With
    
    Exit Sub

Erro_TranspHistorico_Verifica_Alteracao1:

    Select Case gErr
        
        Case 97717
            iSituacao = REGISTRO_ALTERADO
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149021)
            
    End Select
    
    Exit Sub

End Sub

Function TranspHistorico_Grava1(objTranspHistorico As ClassTranspHistorico) As Long

Dim lErro As Long
Dim lNumIntDoc As Long
Dim lComando As Long
Dim sInsert As String

On Error GoTo Erro_TranspHistorico_Grava1
    
    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97718
    
    'Busca o Novo Codigo para insercao
    lErro = CF("Config_ObterAutomatico_EmTrans", "CRFATConfig", "NUM_PROX_TRANSPHISTORICO", "TransportadoraHistorico", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 97719
    
    With objTranspHistorico
    
        .lNumIntDoc = lNumIntDoc
        
        'Guarda a expressão SQL
        sInsert = "INSERT INTO TransportadoraHistorico (NumIntDoc, DataAtualizacao, CodTransp, CGC, InscricaoEstadual, InscricaoMunicipal, RazaoSocial, Endereco, Bairro, Cidade, UF, Pais, CEP) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"
                
        'Grava historico na tabela de TranspHistorico
        lErro = Comando_Executar(lComando, sInsert, .lNumIntDoc, .dtDataAtualizacao, .iCodTransp, .sCgc, left(.sInscricaoEstadual, 18), left(.sInscricaoMunicipal, 18), left(.sRazaoSocial, 40), left(.sEndereco, 40), left(.sBairro, 12), left(.sCidade, 15), .sUF, .sPais, .sCEP)
        If lErro <> AD_SQL_SUCESSO Then gError 97720
        
    End With
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)

    TranspHistorico_Grava1 = SUCESSO
    
    Exit Function

Erro_TranspHistorico_Grava1:

    TranspHistorico_Grava1 = gErr
    
    Select Case gErr
    
        Case 97719
        
        Case 97720
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRANSPHISTORICO", gErr)
            
        Case 97718
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149022)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

Public Function TranspHistorico_Grava_Retroativos(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Transportadoras já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela TranspHistorico como data de gravação
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!

'ATENÇÃO: NÃO RETIRE O COMENTÁRIO ABAIXO
'Para chamar esta funcao pelo immediate ...
'call callbyname(createobject("RotinasCRFAT.ClassCRFATGrava"),"TranspHistorico_Grava_Retroativos",VbMethod,cdate("01/01/02"))

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_TranspHistorico_Grava_Retroativos

    'Abre o Comando
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 94901
    
    'Chama a função que faz a gravação
    lErro = CF("TranspHistorico_Grava_Retroativos_EmTrans ", dtDataAtualizacao)
    If lErro <> SUCESSO Then gError 94902
    
    'Faz o commit da transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 94903
    
    TranspHistorico_Grava_Retroativos = SUCESSO
    
    Exit Function
    
Erro_TranspHistorico_Grava_Retroativos:

    TranspHistorico_Grava_Retroativos = gErr

    Select Case gErr
        
        Case 94902
        
        Case 94901
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 94903
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149023)
    
    End Select

    'Faz o rollback da transação
    Call Transacao_Rollback

End Function

Public Function TranspHistorico_Grava_Retroativos_EmTrans(ByVal dtDataAtualizacao As Date) As Long
'Grava o historico de Transportadoras já existentes no banco de dados
'dtDataAtualizacao (input): passa a data que será inserida na tabela TranspHistorico como data de gravação
'ATENÇÃO: ESTA FUNÇÃO DEVE SER CHAMADA EM TRANSAÇÃO !!

Dim lErro As Long
Dim lComando As Long
Dim sSelect As String
Dim tTranspHistorico As typeTranspHistorico
Dim objTranspHistorico As New ClassTranspHistorico

On Error GoTo Erro_TranspHistorico_Grava_Retroativos_EmTrans

    'Abre o Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 97721
    
    'Carrega sSelect com a Expressão SQL
    sSelect = "SELECT Transportadoras.Codigo, Transportadoras.CGC, Transportadoras.InscricaoEstadual, Transportadoras.InscricaoMunicipal, Transportadoras.Nome, Enderecos.Endereco, Enderecos.Bairro, Enderecos.Cidade, Enderecos.SiglaEstado, Paises.Nome, Enderecos.CEP FROM Transportadoras, Enderecos, Paises WHERE Transportadoras.Endereco = Enderecos.Codigo AND Enderecos.CodigoPais = Paises.Codigo"
    
    With tTranspHistorico
        
        'Inicializa as Strings do Type
        .sBairro = String(STRING_BAIRRO, 0)
        .sCEP = String(STRING_CEP, 0)
        .sCgc = String(STRING_CGC, 0)
        .sCidade = String(STRING_CIDADE, 0)
        .sEndereco = String(STRING_ENDERECO, 0)
        .sInscricaoEstadual = String(STRING_INSCR_EST, 0)
        .sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
        .sPais = String(STRING_PAISES_NOME, 0)
        .sRazaoSocial = String(STRING_TRANSPORTADORA_NOME, 0)
        .sUF = String(STRING_ESTADO_SIGLA, 0)
    
        'Executa o Comando SQL
        lErro = Comando_Executar(lComando, sSelect, .iCodTransp, .sCgc, .sInscricaoEstadual, .sInscricaoMunicipal, .sRazaoSocial, .sEndereco, .sBairro, .sCidade, .sUF, .sPais, .sCEP)
        If lErro <> AD_SQL_SUCESSO Then gError 97723
    
    End With
        
        'Busca o Primeiro Registro
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97724
    
        'Enquanto tem Transp => grava o registro
        Do While lErro <> AD_SQL_SEM_DADOS
                                        
            'Atualiza os dados para futura gravacao de histórico
            With objTranspHistorico
        
                .dtDataAtualizacao = dtDataAtualizacao
                .iCodTransp = tTranspHistorico.iCodTransp
                .sBairro = tTranspHistorico.sBairro
                .sCEP = tTranspHistorico.sCEP
                .sCgc = tTranspHistorico.sCgc
                .sCidade = tTranspHistorico.sCidade
                .sEndereco = tTranspHistorico.sEndereco
                .sInscricaoEstadual = tTranspHistorico.sInscricaoEstadual
                .sInscricaoMunicipal = tTranspHistorico.sInscricaoMunicipal
                .sUF = tTranspHistorico.sUF
                .sRazaoSocial = tTranspHistorico.sRazaoSocial
                .sPais = tTranspHistorico.sPais
                
            End With
        
            'FAZ A GRAVACAO DO HISTORICO
            lErro = TranspHistorico_Grava1(objTranspHistorico)
            If lErro <> SUCESSO Then gError 97725
        
            'Busca o Próximo Registro
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 97726
            
    Loop
    
    'Fecha o Comando
    Call Comando_Fechar(lComando)
    
    TranspHistorico_Grava_Retroativos_EmTrans = SUCESSO
    
    Exit Function
    
Erro_TranspHistorico_Grava_Retroativos_EmTrans:

    TranspHistorico_Grava_Retroativos_EmTrans = gErr

    Select Case gErr
        
        Case 97721
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 97723, 97724, 97726
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPHISTORICO", gErr)
            
        Case 97725
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149024)
    
    End Select

    'Fecha o Comando
    Call Comando_Fechar(lComando)

End Function

'****
'Incluida por Tulio em 20/3 na hora
'de subir as funcoes da tela PlanComissoesInpal

Public Function ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans(ByVal lNumIntDocOrigem As Long) As Long
'Exclui o registro de ComissoesRegras a partir do NumIntDocOrigem passado
'DEVE SER CHAMADA EM TRANSAÇÃO!
'lNumIntDocOrigem RECEBE (INPUT) o NumIntDocOrigem que identifica o registro a ser excluído

Dim lErro As Long
Dim lComando As Long
Dim lComando2 As Long
Dim lNumIntDoc As Long

On Error GoTo Erro_ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans

    'Abre comando
    lComando = Comando_Abrir()
    lComando2 = Comando_Abrir()
    If lComando = 0 Or lComando2 = 0 Then gError 98930
    
    'obtem um conjunto unitario de regras do bd a partir do NumIntDocOrigem
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc FROM ComissoesRegras WHERE NumIntDocOrigem = ?", 0, lNumIntDoc, lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 98931
    
    'se posiciona no primeiro registro do conjunto unitario obtido anteriormente
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 98932
    
'    'se o conjunto for vazio => erro
'    If lErro <> AD_SQL_SUCESSO Then gError 98933
        
    If lErro = AD_SQL_SUCESSO Then
    
        'exclui a regra
        lErro = Comando_ExecutarPos(lComando2, "DELETE FROM ComissoesRegras", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 98934
    
    End If
    
    'fechando comandos...
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans = SUCESSO

    Exit Function
    
Erro_ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans:

    ComissoesRegras_Exclui_NumIntDocOrigem_EmTrans = gErr

    Select Case gErr
    
        '??? botar erros no bd dps...
        Case 98930
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98931, 98932
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COMISSOESREGRAS", gErr)
        
        Case 98934
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COMISSOESREGRAS", gErr)
        
        Case 98933
            Call Rotina_Erro(vbOKOnly, "ERRO_COMISSOESREGRAS_NAO_ENCONTRADA", gErr, lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149025)
            
    End Select
    
    'fechando comandos...
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando2)
    
    Exit Function

End Function

Public Function ComissoesRegras_Grava_EmTrans(ByVal objComissoesRegras As ClassComissoesRegras) As Long
'Chama a funcao que ira inserir efetivamente os dados da regra
'DEVE SER CHAMADA EM TRANSACAO
'objComissoesRegras RECEBE (INPUT) os dados que serao gravados no BD!

Dim lErro As Long
Dim lComando As Long
Dim lNumIntDoc As Long
Dim iVendedor As Integer

On Error GoTo Erro_ComissoesRegras_Grava_EmTrans

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 98905
    
    'Obtem o NumIntDoc da regra que sera gravada
    lErro = CF("Config_ObterAutomatico_EmTrans", "FATConfig", "NUM_PROX_COMISSOESREGRAS", "ComissoesRegras", "NumIntDoc", lNumIntDoc)
    If lErro <> SUCESSO Then gError 98906

    'guarda o NumIntDoc no obj (O NumIntDoc tbm serve para a ordenacao)
    objComissoesRegras.lOrdenacao = lNumIntDoc
    
    'guarda o numintdoc
    objComissoesRegras.lNumIntDoc = lNumIntDoc
    
    'se vendedor estiver preenchido
    If objComissoesRegras.iVendedor > 0 Then
        
        'Faz o lock do vendedor
        lErro = CF("Vendedor_Lock", lComando, objComissoesRegras.iVendedor)
        If lErro <> SUCESSO Then gError 98907
    
    End If
    
    'Insere o registro no BD
    lErro = Comando_Executar(lComando, "INSERT INTO ComissoesRegras (NumIntDoc, Vendedor, Regra, Ordenacao, ValorBase, PercComissao, PercComissaoEmiss, VendedorIndireto, NumIntDocOrigem) VALUES (?,?,?,?,?,?,?,?,?)", objComissoesRegras.lNumIntDoc, objComissoesRegras.iVendedor, objComissoesRegras.sRegra, objComissoesRegras.lOrdenacao, objComissoesRegras.sValorBase, objComissoesRegras.sPercComissao, objComissoesRegras.dPercComissaoEmiss, objComissoesRegras.iVendedorIndireto, objComissoesRegras.lNumIntDocOrigem)
    If lErro <> AD_SQL_SUCESSO Then gError 98908

    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    ComissoesRegras_Grava_EmTrans = SUCESSO
    
    Exit Function

Erro_ComissoesRegras_Grava_EmTrans:

    ComissoesRegras_Grava_EmTrans = gErr
    
    Select Case gErr
        
        Case 98905
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 98906, 98907
        
        Case 98908
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_COMISSOESREGRAS", gErr)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 149026)
            
    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function ParcelasRec_ExcluiAltTit(colParcelaReceberBD As colParcelaReceber, ByVal iNovoNumParcRec As Integer) As Long
'excluir parcelas que estavam no bd e nao estao na colecao em memoria na alteracao de um titulo a receber

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim lNumero As Long
Dim iNumParcela As Integer
Dim tParcelaReceber As typeParcelaReceber
Dim objCarteiraCobrador As New ClassCarteiraCobrador

On Error GoTo Erro_ParcelasRec_ExcluiAltTit

    'Abre comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 26444
        
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 26445
    
    'Seleciona as Parcelas referentes ao Título passado por numero de parcela
    lErro = Comando_ExecutarPos(lComando, "SELECT NumIntDoc, NumParcela, ProxSeqBaixa, Cobrador, CarteiraCobranca, Valor, Saldo FROM ParcelasRec WHERE NumIntTitulo = ? AND NumParcela > ? ORDER BY NumParcela", 0, tParcelaReceber.lNumIntDoc, iNumParcela, tParcelaReceber.iProxSeqBaixa, tParcelaReceber.iCobrador, tParcelaReceber.iCarteiraCobranca, tParcelaReceber.dValor, tParcelaReceber.dSaldo, colParcelaReceberBD(1).lNumIntTitulo, iNovoNumParcRec)
    If lErro <> AD_SQL_SUCESSO Then gError 26455
    
    'Lê a primeira Parcela
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26456
   
    'Loop de exclusão lógica das Parcelas
    Do While lErro = AD_SQL_SUCESSO
    
        'Faz o "lock" da Parcela
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 26447
        
        'Verifica se a parcela já sofreu alguma baixa
        If Abs(tParcelaReceber.dSaldo - tParcelaReceber.dValor) > DELTA_VALORMONETARIO Then gError 95377

        'a parcela tem que estar em cobranca na propria empresa. Isto é particularmente importante p/evitar "esquecer" uma parcela em cobranca eletronica.
        If tParcelaReceber.iCobrador <> COBRADOR_PROPRIA_EMPRESA Then gError 105801
        
        'Verifica se existe algum ChequePre relacionado a Parcela
        lErro = ParcelaRec_SemChequePre(tParcelaReceber.lNumIntDoc)
        If lErro <> SUCESSO Then gError 81883
        
        'atualizar o saldo da carteira devida ("em cheque-pre" ou "em carteira").
        objCarteiraCobrador.iCobrador = tParcelaReceber.iCobrador
        objCarteiraCobrador.iCodCarteiraCobranca = tParcelaReceber.iCarteiraCobranca
        
        lErro = CF("CarteiraCobrador_AtualizaSaldo", objCarteiraCobrador, -1, -tParcelaReceber.dValor)
        If lErro <> SUCESSO Then gError 15993
        
        'Exclui as comissões relacionadas a Parcela desde que nao estejam baixadas
        lErro = CF("Comissoes_Exclui", TIPO_PARCELA, tParcelaReceber.lNumIntDoc, 0)
        If lErro <> SUCESSO Then gError 26452
        
        'Faz a exclusão lógica da Parcela
        lErro = Comando_ExecutarPos(lComando1, "DELETE FROM ParcelasRec", lComando)
        If lErro <> AD_SQL_SUCESSO Then gError 83950
            
        'Lê a próxima Parcela
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 26454
            
    Loop
    
    'Fecha os comandos
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    ParcelasRec_ExcluiAltTit = SUCESSO
     
    Exit Function
    
Erro_ParcelasRec_ExcluiAltTit:

    ParcelasRec_ExcluiAltTit = gErr
     
    Select Case gErr
          
        Case 15993, 26452
        
        Case 95377
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COM_BAIXA2", gErr, iNumParcela)
        
        Case 26444, 26445, 26446
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                        
        Case 26455, 26456, 26454
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC1", gErr, colParcelaReceberBD(1).lNumIntTitulo)
        
        Case 26447
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_PARCELAS_REC", gErr, colParcelaReceberBD(1).lNumIntTitulo, iNumParcela)
        
        Case 26448
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COM_BAIXA2", gErr, iNumParcela)
            
        Case 26449, 26450
'            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr, tParcelaReceber.lNumIntCheque)
        
        Case 26451
'            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUESPRE", gErr, tParcelaReceber.lNumIntCheque)
            
        Case 26453
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_PARCELASREC", gErr, iNumParcela, colParcelaReceberBD(1).lNumIntTitulo)
        
        Case 81883
        
        Case 83950
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PARCELASREC", gErr, iNumParcela, colParcelaReceberBD(1).lNumIntTitulo)
        
        Case 91343
            Call Rotina_Erro(vbOKOnly, "ERRO_EXISTEPARCELAVINCULADACHEQUEPRE", gErr)

        Case 105801
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_COBRANCA_EMPRESA2", gErr, iNumParcela, colParcelaReceberBD(1).lNumIntTitulo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149027)
     
    End Select
     
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function
    
Private Function ParcelaRec_SemChequePre(ByVal lNumIntDoc As Long) As Long
'retorna SUCESSO se NAO houver cheque pre NAO DEPOSITADO (numbordero = 0) associado à parcela identificada por lNumIntDoc

Dim lErro As Long, lNumIntCheque As Long
Dim lComando As Long

On Error GoTo Erro_ParcelaRec_SemChequePre

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81884
        
    lErro = Comando_Executar(lComando, "SELECT ChequeParcelaRec.NumIntCheque FROM ChequePre, ChequeParcelaRec WHERE ChequePre.NumIntCheque = ChequeParcelaRec.NumIntCheque AND ChequeParcelaRec.NumIntParcela = ? AND NumBordero = 0", _
        lNumIntCheque, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 81885
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81886
    
    If lErro = AD_SQL_SUCESSO Then gError 91343
    
    Call Comando_Fechar(lComando)
    
    ParcelaRec_SemChequePre = SUCESSO
     
    Exit Function
    
Erro_ParcelaRec_SemChequePre:

    ParcelaRec_SemChequePre = gErr
     
    Select Case gErr
          
        Case 81884
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81885
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr, lNumIntCheque)
        
        Case 91343
            Call Rotina_Erro(vbOKOnly, "ERRO_EXISTEPARCELAVINCULADACHEQUEPRE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149028)
     
    End Select
     
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Private Function ParcelaRec_ObtemSaldoChequePre(ByVal lNumIntDoc As Long, dSaldoChequePre As Double) As Long
'Obtem o valor total vinculado à parcela identificada por lNumIntDoc de cheques pre-datados nao depositados

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ParcelaRec_ObtemSaldoChequePre

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 81890
        
    dSaldoChequePre = 0
        
    lErro = Comando_Executar(lComando, "SELECT SUM(ValorRecebto-(Juros+Multa)+Desconto) FROM ChequePre, ChequeParcelaRec WHERE ChequePre.NumIntCheque = ChequeParcelaRec.NumIntCheque AND ChequeParcelaRec.NumIntParcela = ? AND NumBordero = 0", _
        dSaldoChequePre, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 81891
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81892
    
    Call Comando_Fechar(lComando)
    
    ParcelaRec_ObtemSaldoChequePre = SUCESSO
     
    Exit Function
    
Erro_ParcelaRec_ObtemSaldoChequePre:

    ParcelaRec_ObtemSaldoChequePre = gErr
     
    Select Case gErr
          
        Case 81890
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 81891, 81892
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE", gErr, lNumIntDoc)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149029)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'**********************
' Inicio da Tela de VendedoresFilial Sergio Ricardo dia 3/07/02


Function VendedorLoja_Filial_Grava(objVendedor As ClassVendedor) As Long
'Função de Gravação

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim objLog As New ClassLog
Dim iCodVendedor As Integer
Dim iTipoLog As Integer

On Error GoTo Erro_VendedorLoja_Filial_Grava

    'Abre a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 107515

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107514

    'Seleciona o Vendedor com o Código Passado
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM Vendedores WHERE Codigo =?", 0, iCodVendedor, objVendedor.iCodigo)
    If lErro <> SUCESSO Then gError 107533

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107516
    
    
    'Faz Lock no Vendedor Selecionado
    lErro = Comando_LockShared(lComando)
    If lErro <> AD_SQL_SUCESSO Then gError 107517
    
    'Função que Exclui o Vendedor na Tabela de Vendedor Loja
    lErro = CF("VendedorFilial_Exclui", objVendedor)
    If lErro <> SUCESSO Then gError 107518

    'Função que Inclui o Vendedor na Tabela de Vendedor Loja
    lErro = CF("VendedorFilial_Inclui", objVendedor)
    If lErro <> SUCESSO And lErro <> 107531 Then gError 107519
    
'    'Passa para o Log o Tipo da Operação
'    iTipoLog = ALTERACAO_VENDEDORLOJA_FILIAL
'
'    'Monta a String de log que deverá ser gravada.
'    Call VendedorFilial_Carrega_Log(objVendedor, objLog, iTipoLog)
'
'    'Monta Registro Para a Tabela de Log
'    lErro = CF("Log_Grava", objLog)
'    If lErro <> SUCESSO Then gError 107520
    
    'Confirma a Transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 107521

    'Fecha os comandos
    Call Comando_Fechar(lComando)

    VendedorLoja_Filial_Grava = SUCESSO

    Exit Function

Erro_VendedorLoja_Filial_Grava:

    VendedorLoja_Filial_Grava = gErr

    Select Case gErr

        Case 107514
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 107515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 107533, 107516
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDOR", gErr, objVendedor.iCodigo)
        
        Case 107517
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_VENDEDOR", gErr)
        
        Case 107518, 107519, 107520
            'Erro Tratado Dentro da Função Chamadora
        
        Case 107521
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149030)

    End Select

    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function VendedorFilial_Exclui(objVendedor As ClassVendedor) As Long
'Função que exclui no Banco de Dados Todos Registros Vinculados ao Vendedor passado como parâmetro

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim objFilialEmpresa As New AdmFiliais
Dim iCodVendedor As Integer

On Error GoTo Erro_VendedorFilial_Exclui

    'Abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 107522
    Next

    lErro = Comando_ExecutarPos(alComando(0), "SELECT Vendedor FROM VendedorLoja WHERE Vendedor = ?", 0, iCodVendedor, objVendedor.iCodigo)
    If lErro <> SUCESSO Then gError 107523
        
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107524
    
    Do While lErro = AD_SQL_SUCESSO
        
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM VendedorLoja", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 107525
    
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 107526
    
    Loop
    
    VendedorFilial_Exclui = SUCESSO
            
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
Erro_VendedorFilial_Exclui:
    
    VendedorFilial_Exclui = gErr

    Select Case gErr
        
        Case 107522
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107523, 107524, 107526
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_VENDEDORLOJA", gErr, objVendedor.iCodigo)
        
        Case 107525
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_VENDEDORLOJA", gErr, objVendedor.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149031)

    End Select
    
    'Fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function
                
Function VendedorFilial_Inclui(objVendedor As ClassVendedor) As Long
'Função que Inclui no Banco de Dados Todos Registros Vinculados ao Vendedor passado como parâmetro

Dim lErro As Long
Dim lComando As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_VendedorFilial_Inclui

    'Abre os comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 107527
    
    'Se a Coleção estiver vazia Então Desvia o Código
    If (objVendedor.colFiliaisLoja.Count = 0) Then gError 107531

    For Each objFilialEmpresa In objVendedor.colFiliaisLoja
        
        'Grava Registros na Tabela VendedorLoja
        lErro = Comando_Executar(lComando, "INSERT INTO VendedorLoja (Vendedor , FilialEmpresa) VALUES ( ?, ?)", objVendedor.iCodigo, objFilialEmpresa.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 107528
    
    Next
    
    VendedorFilial_Inclui = SUCESSO
            
    Call Comando_Fechar(lComando)
    
    Exit Function
    
Erro_VendedorFilial_Inclui:
    
    VendedorFilial_Inclui = gErr

    Select Case gErr
    
        Case 107531
            'Desvio de Código
        
        Case 107527
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 107528
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_VENDEDORLOJA_FILIAL", gErr, objVendedor.iCodigo, objFilialEmpresa.sNome)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149032)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Sub VendedorFilial_Carrega_Log(objVendedor As ClassVendedor, objLog As ClassLog, iTipoLog)

Dim objFilialEmpresa As New AdmFiliais

    'Função para Carregar o objLog
    objLog.sLog = CStr(objVendedor.iCodigo) & Chr(vbKeyEscape)

    For Each objFilialEmpresa In objVendedor.colFiliaisLoja
        
        'Função para Carregar o objLog
        objLog.sLog = objLog.sLog & Chr(vbKeyControl) & CStr(objFilialEmpresa.iCodFilial) & Chr(vbKeyEscape)
    
    Next
    
    objLog.sLog = objLog.sLog & Chr(vbKeyEnd)

    objLog.iOperacao = iTipoLog

    Exit Sub

End Sub

'criada por Mauricio Maciel em 31/03/2003
Function Planilhas_Grava(ByVal colPlanilhas As Collection) As Long
'grava a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim sItem As String
Dim iTipo As Integer
Dim dtData As Date
Dim objPlanilhas As New ClassPlanilhas
Dim objFormacaoPreco As ClassFormacaoPreco 'Inserido por Wagner

On Error GoTo Erro_Planilhas_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92189

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92213
    Next
    
    Set objPlanilhas = colPlanilhas.Item(1)
    
    If objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_CATEGORIA Then

        sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Item FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", 0, sItem, CATEGORIA_PRODUTO_PRECO, objPlanilhas.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 92192

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92193

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92194
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92195
    
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_PRODUTO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM Produtos WHERE Codigo = ?", 0, iTipo, objPlanilhas.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92196

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92197

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92198
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92199
    
    ElseIf objPlanilhas.iEscopo = FORMACAO_PRECO_ESCOPO_TABPRECO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto =?", 0, dtData, giFilialEmpresa, objPlanilhas.iTabelaPreco, objPlanilhas.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92200

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92201

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92202
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92203
    
    End If
    
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Escopo FROM Planilhas WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? AND TipoPlanilha = ?", 0, iEscopo, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco, objPlanilhas.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92204
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92205
    
    'se o modelo estiver cadastrado
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 92208
    
        lErro = CF("Planilhas_Exclui_Trans", objPlanilhas)
        If lErro <> SUCESSO Then gError 92265
    
    End If
    
    'Gerar o Número Interno da Nota Fiscal
    lErro = CF("Planilhas_Automatico", lNumIntDoc, colPlanilhas.Count)
    If lErro <> SUCESSO Then gError 92210
    
    iIndice = 0
    
    For Each objPlanilhas In colPlanilhas
    
        'o modelo não está cadastrado ==> inserir o modelo
        lErro = Comando_Executar(alComando(4), "INSERT INTO Planilhas (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Expressao, Linha, Titulo, TipoPlanilha) VALUES (?,?,?,?,?,?,?,?,?,?)", _
        objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco, lNumIntDoc, objPlanilhas.sExpressao, objPlanilhas.iLinha, objPlanilhas.sTitulo, objPlanilhas.iTipoPlanilha)
        If lErro <> AD_SQL_SUCESSO Then gError 92211

        '###############################################
        'Inserido por Wagner
        'Se o tipo de Planilha for geral grava também FormacaoPreco
                
        If objPlanilhas.iTipoPlanilha = 1 Then
        
            Set objFormacaoPreco = New ClassFormacaoPreco
            
            iIndice = iIndice + 1
        
            objFormacaoPreco.dValor = objPlanilhas.dValor
            objFormacaoPreco.iEscopo = objPlanilhas.iEscopo
            objFormacaoPreco.iFilialEmpresa = objPlanilhas.iFilialEmpresa
            objFormacaoPreco.iLinha = objPlanilhas.iLinha
            objFormacaoPreco.iTabelaPreco = objPlanilhas.iTabelaPreco
            objFormacaoPreco.lNumIntDoc = lNumIntDoc
            objFormacaoPreco.sExpressao = objPlanilhas.sExpressao
            objFormacaoPreco.sItemCategoria = objPlanilhas.sItemCategoria
            objFormacaoPreco.sProduto = objPlanilhas.sProduto
            objFormacaoPreco.sTitulo = objPlanilhas.sTitulo
            
            If iIndice = 1 Then
            
                lErro = FormacaoPreco_Exclui_Trans(objFormacaoPreco)
                If lErro <> SUCESSO Then gError 138294
            
            End If
        
            lErro = Comando_Executar(alComando(4), "INSERT INTO FormacaoPreco (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Expressao, Linha, Titulo) VALUES (?,?,?,?,?,?,?,?,?)", _
            objFormacaoPreco.iFilialEmpresa, objFormacaoPreco.iEscopo, objFormacaoPreco.sItemCategoria, objFormacaoPreco.sProduto, objFormacaoPreco.iTabelaPreco, objFormacaoPreco.lNumIntDoc, objFormacaoPreco.sExpressao, objFormacaoPreco.iLinha, objFormacaoPreco.sTitulo)
            If lErro <> SUCESSO Then gError 138295
        
        End If
        '###############################################

        lNumIntDoc = lNumIntDoc + 1

    Next
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92212

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Planilhas_Grava = SUCESSO

    Exit Function

Erro_Planilhas_Grava:

    Planilhas_Grava = gErr

    Select Case gErr

        Case 92189
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92192, 92193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", gErr, CATEGORIA_PRODUTO_PRECO, objPlanilhas.sItemCategoria)

        Case 92194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA", gErr, objPlanilhas.sItemCategoria, CATEGORIA_PRODUTO_PRECO)

        Case 92195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM1", gErr, CATEGORIA_PRODUTO_PRECO, objPlanilhas.sItemCategoria)

        Case 92196, 92197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objPlanilhas.sProduto)

        Case 92198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objPlanilhas.sProduto)

        Case 92199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS1", gErr, objPlanilhas.sProduto)

        Case 92200, 92201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS2", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iTabelaPreco, objPlanilhas.sProduto)

        Case 92202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objPlanilhas.iTabelaPreco, objPlanilhas.sProduto)

        Case 92203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECOITENS2", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iTabelaPreco, objPlanilhas.sProduto)

        Case 92204, 92205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92210, 92265, 138294, 138295 'Inserido por Wagner

        Case 92211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)
          
        Case 92212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 92213
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149033)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'criada por Mauricio Maciel em 31/03/2003
Function Planilhas_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Retorna o número do próximo NumIntNF disponivel

    Planilhas_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_FORMACAOPRECO", lNumIntDoc, iQtde)
    
End Function

'criada por Mauricio Maciel em 31/03/2003
Function Planilhas_Exclui(ByVal objPlanilhas As ClassPlanilhas) As Long
'exclui a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_Planilhas_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92224

    lErro = CF("Planilhas_Exclui_Trans", objPlanilhas)
    If lErro <> SUCESSO Then gError 92266
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92230

    Planilhas_Exclui = SUCESSO

    Exit Function

Erro_Planilhas_Exclui:

    Planilhas_Exclui = gErr

    Select Case gErr

        Case 92224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 92266
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149034)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'criada por Mauricio Maciel em 31/03/2003
Function Planilhas_Exclui_Trans(ByVal objPlanilhas As ClassPlanilhas) As Long
'exclui a expressão de formação de preço
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim objFormacaoPreco As New ClassFormacaoPreco

On Error GoTo Erro_Planilhas_Exclui_Trans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92225
    Next
   
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Escopo FROM Planilhas WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? AND TipoPlanilha = ?", 0, iEscopo, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco, objPlanilhas.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92226
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92227
    
    If lErro = AD_SQL_SEM_DADOS Then gError 92228
    
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM Planilhas", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92229
        
        'le o modelo em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92227
        
    Loop
    
    '#########################################################
    'Inserido por Wagner
    objFormacaoPreco.dValor = objPlanilhas.dValor
    objFormacaoPreco.iEscopo = objPlanilhas.iEscopo
    objFormacaoPreco.iFilialEmpresa = objPlanilhas.iFilialEmpresa
    objFormacaoPreco.iLinha = objPlanilhas.iLinha
    objFormacaoPreco.iTabelaPreco = objPlanilhas.iTabelaPreco
    objFormacaoPreco.lNumIntDoc = lNumIntDoc
    objFormacaoPreco.sExpressao = objPlanilhas.sExpressao
    objFormacaoPreco.sItemCategoria = objPlanilhas.sItemCategoria
    objFormacaoPreco.sProduto = objPlanilhas.sProduto
    objFormacaoPreco.sTitulo = objPlanilhas.sTitulo
    
    lErro = FormacaoPreco_Exclui_Trans(objFormacaoPreco)
    If lErro <> SUCESSO Then gError 138299
    '#########################################################
        
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Planilhas_Exclui_Trans = SUCESSO

    Exit Function

Erro_Planilhas_Exclui_Trans:

    Planilhas_Exclui_Trans = gErr

    Select Case gErr

        Case 92224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92226, 92227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92228
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORMACAOPRECO_NAO_CADASTRADO", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92229
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_PLANILHAS", gErr, objPlanilhas.iFilialEmpresa, objPlanilhas.iEscopo, objPlanilhas.sItemCategoria, objPlanilhas.sProduto, objPlanilhas.iTabelaPreco)

        Case 92230
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 138299 'Inserido por Wagner
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149035)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'função criada por Mauricio Maciel em 04/04/03
Function MnemonicoFPTipo_Automatico(lNumIntDoc As Long, Optional ByVal iQtde As Integer = 1) As Long
'Retorna o número do próximo NumIntNF disponivel

    MnemonicoFPTipo_Automatico = CF("Config_ObterNumInt", "CRFatConfig", "NUM_PROX_MNEMONICOFPRECO", lNumIntDoc, iQtde)
    
End Function

'Função Criada por Mauricio Maciel em 04/04/03
Function MnemonicoFPTipo_Exclui(ByVal objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long
'exclui a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long

On Error GoTo Erro_MnemonicoFPTipo_Exclui

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92330

    lErro = CF("MnemonicoFPTipo_Exclui_Trans", objMnemonicoFPTipo)
    If lErro <> SUCESSO Then gError 92331

    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92332

    MnemonicoFPTipo_Exclui = SUCESSO

    Exit Function

Erro_MnemonicoFPTipo_Exclui:

    MnemonicoFPTipo_Exclui = gErr

    Select Case gErr

        Case 92330
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92331

        Case 92332
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149036)

    End Select

    Call Transacao_Rollback

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'função criada por Mauricio Maciel em 04/04/03
Function MnemonicoFPTipo_Exclui_Trans(ByVal objMnemonicoFPTipo As ClassMnemonicoFPTipo) As Long
'exclui a expressão de formação de preço
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim objMnemonicoFPreco As New ClassMnemonicoFPreco

On Error GoTo Erro_MnemonicoFPTipo_Exclui_Trans

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92324
    Next

    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Escopo FROM MnemonicoFPTipo WHERE Funcao = 0 AND FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? AND TipoPlanilha =?", 0, iEscopo, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92325

    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92326

    'Comentado por Wagner
    'If lErro = AD_SQL_SEM_DADOS Then gError 92327

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM MnemonicoFPTipo", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92328

        'le o modelo em questão
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92329

    Loop
    
    '#####################################################
    'Inserido por Wagner
    objMnemonicoFPreco.iEscopo = objMnemonicoFPTipo.iEscopo
    objMnemonicoFPreco.iFilialEmpresa = objMnemonicoFPTipo.iFilialEmpresa
    objMnemonicoFPreco.iFuncao = objMnemonicoFPTipo.iFuncao
    objMnemonicoFPreco.iNumParam = objMnemonicoFPTipo.iNumParam
    objMnemonicoFPreco.iParam1 = objMnemonicoFPTipo.iParam1
    objMnemonicoFPreco.iParam2 = objMnemonicoFPTipo.iParam2
    objMnemonicoFPreco.iParam3 = objMnemonicoFPTipo.iParam3
    objMnemonicoFPreco.iTabelaPreco = objMnemonicoFPTipo.iTabelaPreco
    objMnemonicoFPreco.iTipo = objMnemonicoFPTipo.iTipo
    objMnemonicoFPreco.lNumIntDoc = lNumIntDoc
    objMnemonicoFPreco.sExpressao = objMnemonicoFPTipo.sExpressao
    objMnemonicoFPreco.sItemCategoria = objMnemonicoFPTipo.sItemCategoria
    objMnemonicoFPreco.sMnemonico = objMnemonicoFPTipo.sMnemonico
    objMnemonicoFPreco.sMnemonicoDesc = objMnemonicoFPTipo.sMnemonicoDesc
    objMnemonicoFPreco.sProduto = objMnemonicoFPTipo.sProduto
    
    lErro = MnemonicoFPreco_Exclui_Trans(objMnemonicoFPreco)
    If lErro <> SUCESSO Then gError 138307
    '######################################################

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MnemonicoFPTipo_Exclui_Trans = SUCESSO

    Exit Function

Erro_MnemonicoFPTipo_Exclui_Trans:

    MnemonicoFPTipo_Exclui_Trans = gErr

    Select Case gErr

        Case 92324
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92325, 92326, 92329
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPTIPO2", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)

        Case 92327
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MNEMONICOFPTIPO_NAO_CADASTRADO", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)

        Case 92328
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_MNEMONICOFPTIPO", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)

        Case 138307 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149037)

    End Select

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'função criada por Mauricio Maciel em 04/04/03
Function MnemonicoFPTipo_Grava(ByVal colMnemonicoFPTipo As Collection) As Long
'grava a expressão de formação de preço

Dim lErro As Long
Dim alComando(1 To 4) As Long
Dim iIndice As Integer
Dim iEscopo As Integer
Dim lNumIntDoc As Long
Dim lTransacao As Long
Dim sItem As String
Dim iTipo As Integer
Dim dtData As Date
Dim objMnemonicoFPTipo As ClassMnemonicoFPTipo
Dim objMnemonicoFPreco As ClassMnemonicoFPreco 'Inserido por Wagner

On Error GoTo Erro_MnemonicoFPTipo_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 92303

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 92304
    Next
    
    Set objMnemonicoFPTipo = colMnemonicoFPTipo.Item(1)
    
    If objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_CATEGORIA Then

        sItem = String(STRING_CATEGORIAPRODUTOITEM_ITEM, 0)

        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Item FROM CategoriaProdutoItem WHERE Categoria = ? AND Item = ?", 0, sItem, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPTipo.sItemCategoria)
        If lErro <> AD_SQL_SUCESSO Then gError 92305

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92306

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92307
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92308
    
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_PRODUTO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM Produtos WHERE Codigo = ?", 0, iTipo, objMnemonicoFPTipo.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92309

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92310

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92311
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92312
    
    ElseIf objMnemonicoFPTipo.iEscopo = MNEMONICOFPRECO_ESCOPO_TABPRECO Then
    
        'Executar comando SQL
        lErro = Comando_ExecutarPos(alComando(1), "SELECT DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto =?", 0, dtData, giFilialEmpresa, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.sProduto)
        If lErro <> AD_SQL_SUCESSO Then gError 92313

        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92314

        'Se não encontrou
        If lErro = AD_SQL_SEM_DADOS Then gError 92315
    
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 92316
    
    End If
    
    'seleciona o modelo segundo os parametros passados
    lErro = Comando_ExecutarPos(alComando(2), "SELECT Escopo FROM MnemonicoFPTipo WHERE FilialEmpresa=? AND Escopo=? AND ItemCategoria=? AND Produto=? AND TabelaPreco=? AND TipoPlanilha = ?", 0, iEscopo, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.iTipoPlanilha)
    If lErro <> AD_SQL_SUCESSO Then gError 92317
    
    'le o modelo em questão
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 92318
    
    'se o modelo estiver cadastrado
    If lErro = AD_SQL_SUCESSO Then
    
        lErro = Comando_LockExclusive(alComando(2))
        If lErro <> AD_SQL_SUCESSO Then gError 92319
    
        lErro = CF("MnemonicoFPTipo_Exclui_Trans", objMnemonicoFPTipo)
        If lErro <> SUCESSO Then gError 92320
    
    End If
    
    'Gerar o Número Interno da Nota Fiscal
    lErro = CF("MnemonicoFPTipo_Automatico", lNumIntDoc, colMnemonicoFPTipo.Count)
    If lErro <> SUCESSO Then gError 92321
    
    iIndice = 0
    
    For Each objMnemonicoFPTipo In colMnemonicoFPTipo
    
        'o modelo não está cadastrado ==> inserir o modelo
        lErro = Comando_Executar(alComando(4), "INSERT INTO MnemonicoFPTipo (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Mnemonico, MnemonicoDesc, Expressao, Tipo, NumParam, TipoPlanilha) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", _
        objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco, lNumIntDoc, objMnemonicoFPTipo.sMnemonico, objMnemonicoFPTipo.sMnemonicoDesc, objMnemonicoFPTipo.sExpressao, TIPO_NUMERICO, 0, objMnemonicoFPTipo.iTipoPlanilha)
        If lErro <> AD_SQL_SUCESSO Then gError 92322
        
        '########################################################
        'Inserido por Wagner
        If objMnemonicoFPTipo.iTipoPlanilha = 1 Then
        
            Set objMnemonicoFPreco = New ClassMnemonicoFPreco
            
            iIndice = iIndice + 1
                    
            objMnemonicoFPreco.iEscopo = objMnemonicoFPTipo.iEscopo
            objMnemonicoFPreco.iFilialEmpresa = objMnemonicoFPTipo.iFilialEmpresa
            objMnemonicoFPreco.iFuncao = objMnemonicoFPTipo.iFuncao
            objMnemonicoFPreco.iNumParam = objMnemonicoFPTipo.iNumParam
            objMnemonicoFPreco.iParam1 = objMnemonicoFPTipo.iParam1
            objMnemonicoFPreco.iParam2 = objMnemonicoFPTipo.iParam2
            objMnemonicoFPreco.iParam3 = objMnemonicoFPTipo.iParam3
            objMnemonicoFPreco.iTabelaPreco = objMnemonicoFPTipo.iTabelaPreco
            objMnemonicoFPreco.iTipo = objMnemonicoFPTipo.iTipo
            objMnemonicoFPreco.lNumIntDoc = lNumIntDoc
            objMnemonicoFPreco.sExpressao = objMnemonicoFPTipo.sExpressao
            objMnemonicoFPreco.sItemCategoria = objMnemonicoFPTipo.sItemCategoria
            objMnemonicoFPreco.sMnemonico = objMnemonicoFPTipo.sMnemonico
            objMnemonicoFPreco.sMnemonicoDesc = objMnemonicoFPTipo.sMnemonicoDesc
            objMnemonicoFPreco.sProduto = objMnemonicoFPTipo.sProduto
        
            'Se for o primeiro exclui Todos
            If iIndice = 1 Then
                
                lErro = MnemonicoFPreco_Exclui_Trans(objMnemonicoFPreco)
                If lErro <> SUCESSO Then gError 138292
            
            End If
        
            'o modelo não está cadastrado ==> inserir o modelo
            lErro = Comando_Executar(alComando(4), "INSERT INTO MnemonicoFPreco (FilialEmpresa, Escopo, ItemCategoria, Produto, TabelaPreco, NumIntDoc, Mnemonico, MnemonicoDesc, Expressao, Tipo, NumParam) VALUES (?,?,?,?,?,?,?,?,?,?,?)", _
            objMnemonicoFPreco.iFilialEmpresa, objMnemonicoFPreco.iEscopo, objMnemonicoFPreco.sItemCategoria, objMnemonicoFPreco.sProduto, objMnemonicoFPreco.iTabelaPreco, objMnemonicoFPreco.lNumIntDoc, objMnemonicoFPreco.sMnemonico, objMnemonicoFPreco.sMnemonicoDesc, objMnemonicoFPreco.sExpressao, TIPO_NUMERICO, 0)
            If lErro <> AD_SQL_SUCESSO Then gError 138293
                
        End If
        '########################################################

        lNumIntDoc = lNumIntDoc + 1

    Next
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 92323

   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    MnemonicoFPTipo_Grava = SUCESSO

    Exit Function

Erro_MnemonicoFPTipo_Grava:

    MnemonicoFPTipo_Grava = gErr

    Select Case gErr

        Case 92303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 92304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 92305, 92306
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CATEGORIAPRODUTOITEM2", gErr, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPTipo.sItemCategoria)

        Case 92307
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_CADASTRADA", gErr, objMnemonicoFPTipo.sItemCategoria, CATEGORIA_PRODUTO_PRECO)

        Case 92308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_CATEGORIAPRODUTOITEM1", gErr, CATEGORIA_PRODUTO_PRECO, objMnemonicoFPTipo.sItemCategoria)

        Case 92309, 92310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PRODUTOS", gErr, objMnemonicoFPTipo.sProduto)

        Case 92311
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objMnemonicoFPTipo.sProduto)

        Case 92312
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_PRODUTOS1", gErr, objMnemonicoFPTipo.sProduto)

        Case 92313, 92314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS2", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.sProduto)

        Case 92315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.sProduto)

        Case 92316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_TABELASDEPRECOITENS2", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iTabelaPreco, objMnemonicoFPTipo.sProduto)

        Case 92317, 92318
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_MNEMONICOFPTIPO2", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)

        Case 92319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOCK_MNEMONICOFPTIPO", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)

        Case 92320, 92321, 138292, 138293 'Inserido por Wagner

        Case 92322
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_MNEMONICOFPTIPO", gErr, objMnemonicoFPTipo.iFilialEmpresa, objMnemonicoFPTipo.iEscopo, objMnemonicoFPTipo.sItemCategoria, objMnemonicoFPTipo.sProduto, objMnemonicoFPTipo.iTabelaPreco)
          
        Case 92323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149038)

    End Select

    Call Transacao_Rollback
   
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function ClassificacaoFiscal_Grava(objClassificacaoFiscal As ClassClassificacaoFiscal, Optional ByVal bNaoAbreTrans As Boolean = False) As Long
'Grava no Banco de daos um novo registro de Classificação Fiscal ou Realiza um modificação

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim sCodigo As String
Dim iIndice As Integer
Dim lTransacao As Long

On Error GoTo Erro_ClassificacaoFiscal_Grava

    'abre a transacao
    If Not bNaoAbreTrans Then
        lTransacao = Transacao_Abrir()
        If lTransacao = 0 Then gError 123495
    End If
    
    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 123496
    Next

    'Inicializa as Strings
    sCodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)

    'Verifica se o Código já existe no Banco de Dados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM ClassificacaoFiscal WHERE Codigo=?", 0, sCodigo, objClassificacaoFiscal.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 123497

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 123498

    'Se retornou SUCESSO então realiza a atualização do registro
    If lErro = AD_SQL_SUCESSO Then

        'Loca o Código para realizar a modificação
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 123499

        'Realiza o UPDATE através do Código
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE ClassificacaoFiscal SET Descricao = ?, IIAliquota = ?, IPIAliquota = ?, PISAliquota = ?, COFINSAliquota = ?, ICMSAliquota = ?", alComando(1), objClassificacaoFiscal.sDescricao, objClassificacaoFiscal.dIIAliquota, objClassificacaoFiscal.dIPIAliquota, objClassificacaoFiscal.dPISAliquota, objClassificacaoFiscal.dCOFINSAliquota, objClassificacaoFiscal.dICMSAliquota)
        If lErro <> AD_SQL_SUCESSO Then gError 125000

    'Se não achou o registro então
    Else

        'Insere um novo registro no Banco
        lErro = Comando_Executar(alComando(3), "INSERT INTO ClassificacaoFiscal (Codigo, Descricao, IIAliquota,IPIAliquota,PISAliquota,COFINSAliquota,ICMSAliquota) VALUES (?,?,?,?,?,?,?)", objClassificacaoFiscal.sCodigo, objClassificacaoFiscal.sDescricao, objClassificacaoFiscal.dIIAliquota, objClassificacaoFiscal.dIPIAliquota, objClassificacaoFiscal.dPISAliquota, objClassificacaoFiscal.dCOFINSAliquota, objClassificacaoFiscal.dICMSAliquota)
        If lErro <> AD_SQL_SUCESSO Then gError 125001

    End If

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    If Not bNaoAbreTrans Then
        'Confirma a transação
        lErro = Transacao_Commit()
        If lErro <> AD_SQL_SUCESSO Then gError 125002
    End If
    
    ClassificacaoFiscal_Grava = SUCESSO

    Exit Function

Erro_ClassificacaoFiscal_Grava:

    ClassificacaoFiscal_Grava = gErr

    Select Case gErr

        Case 123495
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 123496
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 123497, 123498
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 123499
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125000
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125001
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLACIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125002
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149039)

    End Select

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    If Not bNaoAbreTrans Then Call Transacao_Rollback

    Exit Function

End Function

Public Function ClassificacaoFiscal_Exclui(objClassificacaoFiscal As ClassClassificacaoFiscal) As Long
'Exclui do Banco de Dados o registro com o código passado como parâmetro

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer
Dim sCodigo As String
Dim lTransacao As Long

On Error GoTo Erro_ClassificacaoFiscal_Exclui

    'abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 125003

    'abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = SUCESSO Then gError 125004
    Next

    'Inicializa as Strings
    sCodigo = String(STRING_PRODUTO_IPI_CODIGO, 0)

    'Verifica se o Código já existe no Banco de Dados
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo FROM ClassificacaoFiscal WHERE Codigo=?", 0, sCodigo, objClassificacaoFiscal.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 125005

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 125006

    If lErro <> SUCESSO Then gError 125007

    'Realiza a locação do Código para poder exclui-lo
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125008

    'Exclui o Código do Banco de Dados
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM ClassificacaoFiscal", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 125009

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 125010

    ClassificacaoFiscal_Exclui = SUCESSO

    Exit Function

Erro_ClassificacaoFiscal_Exclui:

    ClassificacaoFiscal_Exclui = gErr

    Select Case gErr

        Case 125003
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 125004
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 125005, 125006
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125007 'Código não encontrado no BD

        Case 125008
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125009
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLASSIFICACAOFISCAL", gErr, objClassificacaoFiscal.sCodigo)

        Case 125010
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149040)

    End Select

    'realiza o fechamento do comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Call Transacao_Rollback

    Exit Function

End Function

Private Function CondicaoPagto_ExcluiParcelas(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long

Dim lErro As Long, alComando(1 To 2) As Long, iIndice As Integer, iSeq As Integer

On Error GoTo Erro_CondicaoPagto_ExcluiParcelas

    'Abre os Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 124025
    Next

    lErro = Comando_ExecutarPos(alComando(1), "SELECT Seq FROM CondicoesPagtoParc WHERE Codigo = ?", 0, iSeq, objCondicaoPagto.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 124026
    
    lErro = Comando_BuscarProximo(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124027
        
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM CondicoesPagtoParc", alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 124028
                
        lErro = Comando_BuscarProximo(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124029
    
    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    CondicaoPagto_ExcluiParcelas = SUCESSO
     
    Exit Function
    
Erro_CondicaoPagto_ExcluiParcelas:

    CondicaoPagto_ExcluiParcelas = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149041)
     
    End Select
     
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function RelacionamentoClientes_Grava_Trans(ByVal objRelacionamentoClientes As ClassRelacClientes, Optional ByVal bGravaAtendente As Boolean = False, Optional ByVal sCodUsu As String = "", Optional ByVal bAlteracaoValida As Boolean = True) As Long

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoClientes_Grava_Trans

    If bGravaAtendente Then
    
        lErro = RelacionamentoClientes_Grava_Atendente(objRelacionamentoClientes, sCodUsu)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    'Se o Código não foi preemchido
    If objRelacionamentoClientes.lCodigo = 0 Then
    
        'Obtém o próximo código de relacionamento para giFilialEmpresa
        lErro = CF("Config_ObterAutomatico_EmTrans", "CRFATConfig", "NUM_PROX_RELACIONAMENTOCLIENTES", "RelacionamentoClientes", "Codigo", lCodigo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        objRelacionamentoClientes.lCodigo = lCodigo
        
    End If
    
    'Chama a função para efetuar a gravação do relacionamento
    lErro = RelacionamentoClientes_Grava_EmTrans(objRelacionamentoClientes, bAlteracaoValida)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    RelacionamentoClientes_Grava_Trans = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Grava_Trans:

    RelacionamentoClientes_Grava_Trans = gErr

    Select Case gErr
   
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149042)

    End Select
    
    Exit Function
    
End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Grava(ByVal objRelacionamentoClientes As ClassRelacClientes, Optional ByVal bGravaAtendente As Boolean = False, Optional ByVal sCodUsu As String = "") As Long
'Abre transação e chama a função para gravação de um relacionamento com cliente
'objRelacionamentoClientes recebe os dados que serão gravados no BD

Dim lErro As Long
Dim lTransacao As Long
Dim bAlteracaoValida As Boolean
Dim vbMsgRes As VbMsgBoxResult
'Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoClientes_Grava

    'Indica que a possível alteração do relacionamento é válida
    bAlteracaoValida = True
    
    If objRelacionamentoClientes.lCodigo <> 0 Then
    
        'Verifica se esse relacionamento é apontado por outro relacionamento como relacionamento anterior
        'para determinar se a data pode ser alterada ou não
        lErro = CF("RelacionamentoClientes_Valida_Alteracao", objRelacionamentoClientes, bAlteracaoValida)
        If lErro <> SUCESSO Then gError 102711
        
    End If
       
    'Se a alteração não foi considerada válida
    If Not bAlteracaoValida Then
    
        'Avisa ao usuário que a data e o cliente não serão alterados
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_RELACIONAMENTO_NAO_ALTERA_DADOS", objRelacionamentoClientes.lCodigo)
        
        'Se o usuário não concordou com a atualização => sai da função por erro
        If vbMsgRes = vbNo Then gError 102712
    
    End If
   
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 102546
    
'    If bGravaAtendente Then
'
'        lErro = RelacionamentoClientes_Grava_Atendente(objRelacionamentoClientes, sCodUsu)
'        If lErro <> SUCESSO Then gError 182340
'
'    End If
'
'    'Se o Código não foi preemchido
'    If objRelacionamentoClientes.lCodigo = 0 Then
'
'        'Obtém o próximo código de relacionamento para giFilialEmpresa
'        lErro = CF("Config_ObterAutomatico_EmTrans", "CRFATConfig", "NUM_PROX_RELACIONAMENTOCLIENTES", "RelacionamentoClientes", "Codigo", lCodigo)
'        If lErro <> SUCESSO Then gError 182341
'
'        objRelacionamentoClientes.lCodigo = lCodigo
'
'    End If
'
'    'Chama a função para efetuar a gravação do relacionamento
'    lErro = RelacionamentoClientes_Grava_EmTrans(objRelacionamentoClientes, bAlteracaoValida)
'    If lErro <> SUCESSO Then gError 102548

    lErro = CF("RelacionamentoClientes_Grava_Trans", objRelacionamentoClientes, bGravaAtendente, sCodUsu, bAlteracaoValida)
    If lErro <> SUCESSO Then gError 102548
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 102547
    
    RelacionamentoClientes_Grava = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Grava:

    RelacionamentoClientes_Grava = gErr

    Select Case gErr
   
        Case 102711, 102712, 182340, 182341
        
        Case 102546
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 102548
        
        Case 102547
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149042)

    End Select
    
    'Faz o rollback da transacao
    Call Transacao_Rollback
    
    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Grava_EmTrans(ByVal objRelacionamentoClientes As ClassRelacClientes, ByVal bAlteracaoValida As Boolean) As Long
'Faz o lock dos objetos relacionados, verifica a existência do relacionamento
'e insere um novo relacionamento ou altera o já existente
'objRelacionamentoClientes recebe os dados que serão gravados no BD

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim tRelacionamentoClientes As typeRelacionamentoClientes

On Error GoTo Erro_RelacionamentoClientes_Grava_EmTrans

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 102559
    Next
    
    If objRelacionamentoClientes.lNumIntParcRec <> 0 Then
        objRelacionamentoClientes.iTipoDoc = RELACCLI_TIPODOC_PARCELARECEBER
        objRelacionamentoClientes.lNumIntDocOrigem = objRelacionamentoClientes.lNumIntParcRec
    End If

    'Efetua os locks dos objetos relacionados
    lErro = CF("RelacionamentoClientes_Locks_Relacionados", objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102558
    
    'Verifica se já existe relacionamento com o código em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo, Data, Cliente FROM RelacionamentoClientes WHERE Codigo=? AND FilialEmpresa=?", 0, tRelacionamentoClientes.lCodigo, tRelacionamentoClientes.dtData, tRelacionamentoClientes.lCliente, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102565
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102566
    
    'Se não encontrou relacionamento com o código em questão
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere um novo relacionamento no BD
        With objRelacionamentoClientes
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelacionamentoClientes (Codigo,FilialEmpresa,Origem,Data,Hora,Tipo,Cliente,FilialCliente,Contato,Atendente,RelacionamentoAnt,Assunto1,Assunto2,Status,NumIntParcRec,DataProxCobr,DataPrevReceb, StatusCG, DataFim, HoraFim, Motivo, TipoDoc, NumIntDocOrigem,Satisfacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lCodigo, .iFilialEmpresa, .iOrigem, .dtData, CDbl(.dtHora), .lTipo, .lCliente, .iFilialCliente, .iContato, .iAtendente, .lRelacionamentoAnt, .sAssunto1, left(.sAssunto2, 255), .iStatus, .lNumIntParcRec, .dtDataProxCobr, .dtDataPrevReceb, .iStatusCG, .dtDataFim, CDbl(.dtHoraFim), .lMotivo, .iTipoDoc, .lNumIntDocOrigem, .lSatisfacao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 102567
    
    'Senão, ou seja, se encontrou o relacionamento
    Else
    
        'Se a alteração não é válida
        If Not bAlteracaoValida Then
        
            'joga no obj a data que já está no BD, ou seja, a atualização não fará diferença
            objRelacionamentoClientes.dtData = tRelacionamentoClientes.dtData
            
            'joga no obj o cliente que já está no BD, ou seja, a atualização não fará diferença
            objRelacionamentoClientes.lCliente = tRelacionamentoClientes.lCliente
        
        End If
        
        'Altera os dados do relacionamento
        With objRelacionamentoClientes
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE RelacionamentoClientes SET Origem=?, Data=?, Hora=?, Tipo=?, Cliente=?, FilialCliente=?, Contato=?, Atendente=?, RelacionamentoAnt=?, Assunto1=?, Assunto2=?, Status=?, NumIntParcRec = ?, DataProxCobr=?, DataPrevReceb = ?, StatusCG=?, DataFim = ?, HoraFim = ?, Motivo = ?, TipoDoc = ?, NumIntDocOrigem  =?, Satisfacao = ?", alComando(0), .iOrigem, .dtData, CDbl(.dtHora), .lTipo, .lCliente, .iFilialCliente, .iContato, .iAtendente, .lRelacionamentoAnt, .sAssunto1, left(.sAssunto2, 255), .iStatus, .lNumIntParcRec, .dtDataProxCobr, .dtDataPrevReceb, .iStatusCG, .dtDataFim, CDbl(.dtHoraFim), .lMotivo, .iTipoDoc, .lNumIntDocOrigem, .lSatisfacao)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 102568
    
    End If
    
    lErro = RelacCliTexto_Grava(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If objRelacionamentoClientes.iTipoDoc = RELACCLI_TIPODOC_OV And objRelacionamentoClientes.lNumIntDocOrigem <> 0 Then
    
        lErro = Comando_Executar(alComando(2), "UPDATE OrcamentoVenda SET DataPrevReceb = ?, DataProxCobr=?, Status=? WHERE NumIntDoc = ?", objRelacionamentoClientes.dtDataPrevReceb, objRelacionamentoClientes.dtDataProxCobr, objRelacionamentoClientes.lStatusTipoDoc, objRelacionamentoClientes.lNumIntDocOrigem)
        If lErro <> AD_SQL_SUCESSO Then gError 102568
    
    End If
        
    If objRelacionamentoClientes.iTipoDoc = RELACCLI_TIPODOC_SOLSRV And objRelacionamentoClientes.lNumIntDocOrigem <> 0 Then
        
        If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then
            'Vai mudar o status dos itens da solicitação
            lErro = CF("SolicitacaoSRV_Encerra", objRelacionamentoClientes.lNumIntDocOrigem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        End If
        
    End If
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacionamentoClientes_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Grava_EmTrans:

    RelacionamentoClientes_Grava_EmTrans = gErr

    Select Case gErr

        Case 102559
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102558

        Case 102565, 102566
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 102567
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELACIONAMENTOCLIENTES", gErr, objRelacionamentoClientes.lCodigo)
            
        Case 102568
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELACIONAMENTOCLIENTES", gErr, objRelacionamentoClientes.lCodigo)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149043)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Exclui(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Abre transação e chama a função responsável pela exclusão do relacionamento
'objRelacionamentoClientes RECEBE(Input) os dados do relacionamento a ser excluído

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long

On Error GoTo Erro_RelacionamentoClientes_Exclui

    'Executa a abertura da Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 102637
    
    'Exclui o relacionamento em questão
    lErro = RelacionamentoClientes_Exclui_EmTrans(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102639
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 102640

    RelacionamentoClientes_Exclui = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Exclui:

    RelacionamentoClientes_Exclui = gErr

    Select Case gErr

        Case 102637
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 102639

        Case 102640
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149044)

    End Select

    'Faz o rollback da transacao
    Call Transacao_Rollback

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoClientes_Exclui_EmTrans(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Exclui o relacionamento com código e filial empresa passados como parâmetro
'Um relacionamento que seja indicado como antecessor de outro, não poderá ser excluído
'objRelacionamentoClientes RECEBE(Input)

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoClientes_Exclui_EmTrans

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 102641
    Next
    
    'Lê o relacionamento em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM RelacionamentoClientes WHERE Codigo=? AND FilialEmpresa=?", 0, lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102642
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102643
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102644
    
    'Verifica no bd se existe algum relacionamento nessa filial que usa o relacionamento
    'a ser excluído como relacionamento anterior
    lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM RelacionamentoClientes WHERE RelacionamentoAnt=? AND FilialEmpresa=?", lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102645

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102646
    
    'Se encontrou => erro, pois se o relacionamento é apontado como anterior
    'a outro relacionamento, ele não pode ser excluído
    If lErro = AD_SQL_SUCESSO Then gError 102647
    
    'Exclui o relacionamento em questão
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RelacionamentoClientes", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 102648
    
    lErro = RelacCliTexto_Exclui(objRelacionamentoClientes)
    If lErro <> AD_SQL_SUCESSO Then gError 102648
    
    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacionamentoClientes_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Exclui_EmTrans:

    RelacionamentoClientes_Exclui_EmTrans = gErr

    Select Case gErr

        Case 102641
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 102642, 102643, 102645, 102646
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 102644
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo)
        
        Case 102647
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_ANTERIOR_OUTRO", gErr, objRelacionamentoClientes.lCodigo, lCodigo)
        
        Case 102648
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149045)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Grava(ByVal objAtendente As ClassAtendentes) As Long
'Grava o Atendente no BD

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim iAtivo As Integer

On Error GoTo Erro_Atendentes_Grava

    'Abre transação no SGEDados
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 102762
    
    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 102763
    Next

    'Seleciona o Atendente com o código passado
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Atendentes WHERE Codigo=? AND FilialEmpresa=?", iCodigo, objAtendente.iCodigo, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102764

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102765

    'Se encontrou o atendente => erro
    If lErro = AD_SQL_SUCESSO Then gError 102766
    
    '*** LOCK DO USUÁRIO NO DIC ****
    'Seleciona o usuario vinculado ao Atendente
    lErro = Comando_ExecutarLockado(alComando(1), "SELECT Ativo FROM Usuario WHERE CodUsuario=?", iAtivo, objAtendente.sCodUsuario)
    If lErro <> AD_SQL_SUCESSO Then gError 102767

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 102768
    
    'Se não encontrou => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 102769

    'Efetua lock do usuário
    lErro = Comando_LockShared(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 102770
    '****************************************

    'Insere o atendente no BD
    lErro = Comando_Executar(alComando(2), "INSERT INTO Atendentes (Codigo, CodUsuario, FilialEmpresa) VALUES (?,?,?)", objAtendente.iCodigo, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102771

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 102772

    Atendentes_Grava = SUCESSO

    Exit Function

Erro_Atendentes_Grava:

    Atendentes_Grava = gErr

    Select Case gErr

        Case 102762
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 102763
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102764, 102765
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr, objAtendente.iCodigo)

        Case 102766
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_ALTERAVEL", gErr)

        Case 102767, 102768
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", gErr)
        
        Case 102769
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr, objAtendente.sCodUsuario)
            
        Case 102770
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_USUARIO", gErr, objAtendente.sCodUsuario)
        
        Case 102771
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ATENDENTE", gErr)

        Case 102772
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149046)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Faz rollback da transação
    Call Transacao_Rollback
    
End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function Atendentes_Exclui(ByVal objAtendente As ClassAtendentes) As Long
'Exclui o atendente da tabela

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1) As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Atendentes_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 102777
    
    'Abertura de comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 102778
    Next

    'Seleciona o atendente no bd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM Atendentes WHERE CodUsuario=? AND FilialEmpresa=?", 0, iCodigo, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 102779

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 102780

    'Faz o lock do atendente
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 102781

    'Verifica se o atendente está vinculado a outras tabelas
    lErro = CF("Atendentes_Valida_Exclusao", objAtendente)
    If lErro <> SUCESSO Then gError 102782

    'Exclui o atendente do BD
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Atendentes", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 102783

    'Fechamento de comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 102784

    Atendentes_Exclui = SUCESSO

    Exit Function

Erro_Atendentes_Exclui:

    Atendentes_Exclui = gErr

    Select Case gErr

        Case 102782
        
        Case 102777
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 102778
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 102779, 102780
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)

        Case 102781
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ATENDENTES", gErr, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)

        Case 102783
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_ATENDENTES", gErr, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)

        Case 102784
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149047)

    End Select

    'Fechamento de comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Rollback da transação
    Call Transacao_Rollback
    
    Exit Function

End Function

Function Transportadora_Grava_EmTrans(ByVal objTransportadora As ClassTransportadora, ByVal objEndereco As ClassEndereco) As Long
'Grava A Transportadora no BD

Dim lComando(3) As Long
Dim lErro As Long
Dim lTransacao As Long
Dim lCodEndereco As Long
Dim iIndice As Integer
Dim tTransportadora As typeTransportadora
Dim iCodigo  As Integer
Dim iTotalTransp As Integer
Dim objTranspHistorico As New ClassTranspHistorico
Dim objPais As New ClassPais

On Error GoTo Erro_Transportadora_Grava_EmTrans

    'Abre os  comandos
    For iIndice = 0 To 2
        lComando(iIndice) = Comando_Abrir()
        If lComando(iIndice) = 0 Then gError 22065
    Next

    'Procura por um outra Transportadora com o mesmo Nome Reduzido
    lErro = Comando_Executar(lComando(0), "SELECT Codigo FROM Transportadoras WHERE Codigo <> ? AND NomeReduzido = ?", iCodigo, objTransportadora.iCodigo, objTransportadora.sNomeReduzido)
    If lErro <> AD_SQL_SUCESSO Then gError 22067

    'Tenta ler o código desta Transportadora
    lErro = Comando_BuscarPrimeiro(lComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 22069

    'Se existir outra Transportadora com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 22070

    'Atualiza os dados para futura gravacao de histórico
    With objTranspHistorico
    
        .dtDataAtualizacao = gdtDataHoje
        .iCodTransp = objTransportadora.iCodigo
        .sBairro = objEndereco.sBairro
        .sCEP = objEndereco.sCEP
        .sCgc = objTransportadora.sCgc
        .sCidade = objEndereco.sCidade
        .sEndereco = objEndereco.sEndereco
        .sInscricaoEstadual = objTransportadora.sInscricaoEstadual
        .sInscricaoMunicipal = objTransportadora.sInscricaoMunicipal
        .sUF = objEndereco.sSiglaEstado
        .sRazaoSocial = objTransportadora.sNome
                
        'Le o nome do Pais com o codigo no obj
        objPais.iCodigo = objEndereco.iCodigoPais
        lErro = CF("Paises_Le", objPais)
        If lErro <> SUCESSO And lErro <> 47876 Then gError 97728
    
        'Pais não está cadastrado - --- -> Erro
        If lErro = 47876 Then gError 97729
        
        .sPais = objPais.sNome
        
        'FAZ A GRAVACAO DO HISTORICO
        lErro = TranspHistorico_Grava(objTranspHistorico)
        If lErro <> SUCESSO Then gError 97730
                
    End With
    
    'Lê a Transportadora no BD
    lErro = Comando_ExecutarPos(lComando(1), "SELECT Codigo, Endereco FROM Transportadoras WHERE Codigo = ?", 0, _
        tTransportadora.iCodigo, tTransportadora.lEndereco, objTransportadora.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 22071
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 22072

    'Transportadora não existe - CRIACAO
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Se for Versão Light
        If giTipoVersao = VERSAO_LIGHT Then
            
            'Lê o Número Total de Transportadoras
            lErro = Transportadora_Le_Total(iTotalTransp)
            If lErro <> SUCESSO Then gError 64074
            
            'Se ultrapassou o número máximo de transportadoras ---> ERRO
            If iTotalTransp >= LIMITE_TRANSP_VLIGHT Then gError 64075
        
        End If

        'Grava o novo Endereco de Transportadora
        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 22076

        objEndereco.lCodigo = lCodEndereco

        lErro = CF("Endereco_Grava", objEndereco)
        If lErro <> SUCESSO Then gError 22077

        'Preenche ponteiro de endereço de objTransportadora
        objTransportadora.lEndereco = objEndereco.lCodigo
    
        'Insere a Transportadora na tabela Transportadora
        lErro = Comando_Executar(lComando(2), "INSERT INTO Transportadoras (Codigo, Nome, NomeReduzido, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, ViaTransporte, PesoMinimo, Guia, Observacao, IEIsento, IENaoContrib) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", objTransportadora.iCodigo, objTransportadora.sNome, objTransportadora.sNomeReduzido, objTransportadora.sCgc, objTransportadora.sInscricaoEstadual, objTransportadora.sInscricaoMunicipal, objTransportadora.lEndereco, objTransportadora.iViaTransporte, objTransportadora.dPesoMinimo, objTransportadora.sGuia, objTransportadora.sObservacao, objTransportadora.iIEIsento, objTransportadora.iIENaoContrib)
        If lErro <> AD_SQL_SUCESSO Then gError 22073

    Else 'Tranportadora existe - ALTERACAO

        objTransportadora.lEndereco = tTransportadora.lEndereco

        'Faz a modificação do Endereço
        objEndereco.lCodigo = objTransportadora.lEndereco

        'Atualiza o endereço
        
        lErro = CF("Endereco_Grava", objEndereco)
        If lErro <> SUCESSO Then gError 22074

        'Atualiza a tabela de Transportadora
        lErro = Comando_ExecutarPos(lComando(2), "UPDATE Transportadoras SET Nome=?, NomeReduzido=?, CGC=?, InscricaoEstadual=?, InscricaoMunicipal=?, ViaTransporte=?, PesoMinimo = ?, Guia = ?, Observacao=?, IEIsento = ?, IENaoContrib = ?", lComando(1), objTransportadora.sNome, objTransportadora.sNomeReduzido, objTransportadora.sCgc, objTransportadora.sInscricaoEstadual, objTransportadora.sInscricaoMunicipal, objTransportadora.iViaTransporte, objTransportadora.dPesoMinimo, objTransportadora.sGuia, objTransportadora.sObservacao, objTransportadora.iIEIsento, objTransportadora.iIENaoContrib)
        If lErro <> AD_SQL_SUCESSO Then gError 22075

    End If

    'Fecha os comandos
    For iIndice = 0 To 2
        lErro = Comando_Fechar(lComando(iIndice))
    Next

    Transportadora_Grava_EmTrans = SUCESSO

    Exit Function
    
Erro_Transportadora_Grava_EmTrans:

    Transportadora_Grava_EmTrans = gErr

    Select Case gErr

        Case 22065
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 22070
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSP_NOME_RED_DUPLICADO", gErr, iCodigo)

        Case 22067, 22069, 22071, 22072
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRANSPORTADORA", gErr)

        Case 22073
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TRANSPORTADORA", gErr)

        Case 22074, 22076, 22077, 64074

        Case 22075
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_TRANSPORTADORA", gErr)
        
        Case 64075
            Call Rotina_Erro(vbOKOnly, "ERRO_LIMITE_TRANSP_VLIGHT", gErr, LIMITE_TRANSP_VLIGHT)

        Case 97728, 97730
        
        Case 97729
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO", gErr, objPais.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149048)

    End Select

    For iIndice = 0 To 2
        Call Comando_Fechar(lComando(iIndice))
    Next

    Exit Function

End Function

Function NaturezaOperacao_Grava_EmTrans(ByVal objNaturezaOperacao As ClassNaturezaOp) As Long
'Grava a Natureza de Operação no banco de dados

Dim lErro As Long
Dim lComando As Long
Dim lComando1 As Long
Dim iStatus As Integer
Dim sCodigo As String
Dim objNaturezaOPHistorico As New ClassNaturezaOPHistorico

On Error GoTo Erro_NaturezaOperacao_Grava_EmTrans

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 17969

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 17970
    
    sCodigo = String(STRING_NATUREZAOP_CODIGO, 0)

    'Pesquisa no BD a Natureza de Operação em questão
    lErro = Comando_ExecutarPos(lComando, "SELECT Codigo FROM NaturezaOp WHERE Codigo = ?", 0, sCodigo, objNaturezaOperacao.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 17972

    'Lê a Natureza de Operação, se estiver no BD
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17973

    'Atualiza os dados para futura gravacao de histórico
    With objNaturezaOPHistorico
    
        .dtDataAtualizacao = gdtDataHoje
        .sCodigo = objNaturezaOperacao.sCodigo
        .sDescricao = objNaturezaOperacao.sDescrNF
        
    End With
    
    'Encontrou a Natureza ==> é uma alteração
    If lErro = AD_SQL_SUCESSO Then
    
        '"Lock" da Natureza de Operação
        lErro = Comando_LockExclusive(lComando)
        If lErro <> SUCESSO Then gError 29644
        
        'FAZ A GRAVACAO DO HISTORICO
        lErro = NaturezaOPHistorico_Grava(objNaturezaOPHistorico)
        If lErro <> SUCESSO Then gError 97703
        
        'Atualiza a Natureza de Operação
        lErro = Comando_ExecutarPos(lComando1, "UPDATE NaturezaOp SET Descricao = ?, DescrNF = ?", lComando, objNaturezaOperacao.sDescricao, objNaturezaOperacao.sDescrNF)
        If lErro <> AD_SQL_SUCESSO Then gError 17974

    'Não encontrou a Natureza de Operação ==>  é uma inserção
    Else

        'FAZ A GRAVACAO DO HISTORICO
        lErro = NaturezaOPHistorico_Grava(objNaturezaOPHistorico)
        If lErro <> SUCESSO Then gError 97703
        
        'Insere a Natureza de Operação no BD
        lErro = Comando_Executar(lComando1, "INSERT INTO NaturezaOp (Codigo, Descricao, DescrNF) VALUES (?,?,?)", objNaturezaOperacao.sCodigo, objNaturezaOperacao.sDescricao, objNaturezaOperacao.sDescrNF)
        If lErro <> AD_SQL_SUCESSO Then gError 17975

    End If

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    NaturezaOperacao_Grava_EmTrans = SUCESSO

    Exit Function

Erro_NaturezaOperacao_Grava_EmTrans:

    NaturezaOperacao_Grava_EmTrans = gErr

    Select Case gErr

        Case 97703
        
        Case 17969, 17970
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 17972, 17973
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", gErr, objNaturezaOperacao.sCodigo)

        Case 17974
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_NATUREZAOP", gErr, objNaturezaOperacao.sCodigo)

        Case 17975
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_NATUREZAOP", gErr, objNaturezaOperacao.sCodigo)

        Case 29644
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", gErr, objNaturezaOperacao.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 149049)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)

    Exit Function

End Function

'#####################################################################
'Inserido por Wagner 06/07/2006
Public Function Contato_Automatico(lCodigo As Long) As Long
'retorna o número do proximo Contato disponivel

    Contato_Automatico = CF("Config_ObterAutomatico", "CPRConfig", "NUM_PROX_CONTATO", "Contatos", "Codigo", lCodigo)

End Function

Function Contato_Grava(ByVal objContato As ClassContatos, ByVal colEndereco As Collection) As Long
'Grava o Contato passado como parâmetro

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Contato_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180430

    lErro = Contato_Grava_EmTrans(objContato, colEndereco)
    If lErro <> SUCESSO Then gError 180431
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180432

    Contato_Grava = SUCESSO

    Exit Function

Erro_Contato_Grava:

    Contato_Grava = gErr

    Select Case gErr

        Case 180430
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 180431

        Case 180432
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180433)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Contato_Grava_EmTrans(ByVal objContato As ClassContatos, ByVal colEndereco As Collection) As Long
'Grava o Contato passado como parâmetro

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(6) As Long
Dim lCodigo As Long
Dim lCodEndereco As Long
Dim tContato As typeContato
Dim tFilialContato As typeFilialContato
Dim colCategoriaItemCopia As New Collection
Dim objFilialConCategoria As ClassFilialConCategoria
Dim objFilialContato As New ClassFilialContato
Dim objEndereco As New ClassEndereco
Dim lCodigoLoja As Long
Dim lCodigoContato As Long
Dim iAtivo As Integer
Dim objCliente As New ClassCliente

On Error GoTo Erro_Contato_Grava_EmTrans

    'Inicializa os atributos STRING dos types Contato e FilialContato
    tContato.sRazaoSocial = String(STRING_CLIENTE_RAZAO_SOCIAL, 0)
    tContato.sNomeReduzido = String(STRING_CLIENTE_NOME_REDUZIDO, 0)
    tContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialContato.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialContato.sCgc = String(STRING_CGC, 0)
    tFilialContato.sRG = String(STRING_RG, 0)
    tFilialContato.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialContato.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialContato.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tFilialContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialContato.sContaContabil = String(STRING_CONTA, 0)

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180434
    Next
    
    'guarda o código do Contato para fins de msg de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lCodigoContato = objContato.lCodigoLoja
    Else
        lCodigoContato = objContato.lCodigo
    End If

    'Cria uma cópia "de trabalho" da coleção passada como parâmetro
    For Each objFilialConCategoria In objContato.colCategoriaItem
        colCategoriaItemCopia.Add objFilialConCategoria
    Next

    'se estiver operando no Caixa Central
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        
        'Procura por um outro Contato com o mesmo Nome Reduzido
        lErro = Comando_Executar(alComando(0), "SELECT CodigoLoja FROM Contatos WHERE CodigoLoja <> ? AND NomeReduzido = ? ", lCodigoLoja, objContato.lCodigoLoja, objContato.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 180435
    
    Else
    
        'Procura por um outro Contato com o mesmo Nome Reduzido
        lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Contatos WHERE Codigo <> ? AND NomeReduzido = ? ", lCodigo, objContato.lCodigo, objContato.sNomeReduzido)
        If lErro <> AD_SQL_SUCESSO Then gError 180436
    
    End If
    
    'Tenta ler o código deste Contato
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180437

    'Se existir outro Contato com o mesmo Nome Reduzido, erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 180438
    
    Call Contato_Transfere_Cliente(objContato, objCliente)
    
    'Faz "locks" nas tabelas dos objetos relacionados ao Cliente
    lErro = CF("Clientes_Le_Tabelas_Lock", objCliente)
    If lErro <> SUCESSO Then gError 180543
        
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
    
        'Seleciona o Contato no BD por codigoloja
        lErro = Comando_ExecutarPos(alComando(1), "SELECT CodigoLoja, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Ativo  FROM Contatos WHERE CodigoLoja = ?", 0, _
            tContato.lCodigoLoja, tContato.sRazaoSocial, tContato.sNomeReduzido, tContato.iTipo, tContato.sObservacao, tContato.dLimiteCredito, tContato.iCondicaoPagto, tContato.dDesconto, tContato.iCodMensagem, tContato.iTabelaPreco, tContato.iAtivo, objContato.lCodigoLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 180439
    
    Else
    
        'Seleciona o Contato no BD por codigo
        lErro = Comando_ExecutarPos(alComando(1), "SELECT Codigo, RazaoSocial, NomeReduzido, Tipo, Observacao, LimiteCredito, CondicaoPagto, Desconto, CodMensagem, TabelaPreco, Ativo  FROM Contatos WHERE Codigo = ?", 0, _
            tContato.lCodigo, tContato.sRazaoSocial, tContato.sNomeReduzido, tContato.iTipo, tContato.sObservacao, tContato.dLimiteCredito, tContato.iCondicaoPagto, tContato.dDesconto, tContato.iCodMensagem, tContato.iTabelaPreco, tContato.iAtivo, objContato.lCodigo)
        If lErro <> AD_SQL_SUCESSO Then gError 180440
    
    End If

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180441
    
    'Contato não está cadastrado - GRAVACAO
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Grava os novos Enderecos (Principal, Entrega, Cobranca) de Contato
        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Automatico", lCodEndereco)
            If lErro <> SUCESSO Then gError 180442

            colEndereco.Item(iIndice).lCodigo = lCodEndereco

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 180443

        Next

        If colEndereco.Count > 0 Then objContato.lEndereco = colEndereco.Item(1).lCodigo
        If colEndereco.Count > 1 Then objContato.lEnderecoEntrega = colEndereco.Item(2).lCodigo
        If colEndereco.Count > 2 Then objContato.lEnderecoCobranca = colEndereco.Item(3).lCodigo

        lErro = Comando_Executar(alComando(2), "INSERT INTO Contatos (Codigo,RazaoSocial,NomeReduzido,Tipo,Observacao,LimiteCredito,CondicaoPagto,Desconto,CodMensagem,TabelaPreco,ProxCodFilial, CodPadraoCobranca, Ativo, CodigoLoja,FilialEmpresaLoja, Bloqueado) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objContato.lCodigo, objContato.sRazaoSocial, objContato.sNomeReduzido, objContato.iTipo, objContato.sObservacao, objContato.dLimiteCredito, objContato.iCondicaoPagto, objContato.dDesconto, objContato.iCodMensagem, objContato.iTabelaPreco, FILIAL_MATRIZ + 1, objContato.iCodPadraoCobranca, objContato.iAtivo, objContato.lCodigoLoja, objContato.iFilialEmpresaLoja, objContato.iBloqueado)
        If lErro <> AD_SQL_SUCESSO Then gError 180444

        lErro = Comando_Executar(alComando(3), "INSERT INTO FiliaisContatos (CodContato,CodFilial,Nome,CGC,InscricaoEstadual,InscricaoMunicipal,Endereco,EnderecoEntrega,EnderecoCobranca,CodTransportadora,Observacao,ContaContabil,Vendedor,ComissaoVendas,Regiao,FreqVisitas,DataUltVisita,CodCobrador,InscricaoSuframa, RG, CodFilialLoja, FilialEmpresaLoja, CodContatoLoja, Ativo, CodTranspRedesp, DetPagFrete, Guia, TipoFrete) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                                 objContato.lCodigo, FILIAL_MATRIZ, MATRIZ, objContato.sCgc, objContato.sInscricaoEstadual, objContato.sInscricaoMunicipal, objContato.lEndereco, objContato.lEnderecoEntrega, objContato.lEnderecoCobranca, objContato.iCodTransportadora, objContato.sObservacao2, objContato.sContaContabil, objContato.iVendedor, objContato.dComissaoVendas, objContato.iRegiao, objContato.iFreqVisitas, objContato.dtDataUltVisita, objContato.iCodCobrador, objContato.sInscricaoSuframa, objContato.sRG, FILIAL_MATRIZ, objContato.iFilialEmpresaLoja, objContato.lCodigoLoja, objContato.iAtivo, objContato.iCodTranspRedesp, objContato.iDetPagFrete, objContato.sGuia, objContato.iTipoFrete)
        If lErro <> AD_SQL_SUCESSO Then gError 180445

        'Insere as categorias da filial do Contato em FilialContatosCategorias
        objFilialContato.lCodContato = objContato.lCodigo

        objFilialContato.iCodFilial = FILIAL_MATRIZ
        
        lErro = FilialContato_Insere_Categorias(objFilialContato, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 180446

    Else 'Contato já existe - MODIFICACAO
        
        'se o campo ativo do BD vai ser desativado -->desativa todas as filiais
        If tContato.iAtivo <> objContato.iAtivo And objContato.iAtivo = DESMARCADO Then
            'Seleciona as filiais Contatos no BD por codigo(menos a matriz)
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM FiliaisContatos WHERE CodContatoLoja = ? AND CodFilialLoja <> ? AND Ativo = ?", 0, iAtivo, objContato.lCodigoLoja, FILIAL_MATRIZ, CLIENTE_ATIVO)
            Else
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM FiliaisContatos WHERE CodContato = ? AND CodFilial <> ? AND Ativo = ?", 0, Ativo, objContato.lCodigo, FILIAL_MATRIZ, CLIENTE_ATIVO)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 180447
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180448
            
            Do While lErro = AD_SQL_SUCESSO
                'Atualiza a tabela de FiliaisContatos
                lErro = Comando_ExecutarPos(alComando(6), "UPDATE FiliaisContatos SET Ativo = ?", alComando(5), objContato.iAtivo)
                If lErro <> AD_SQL_SUCESSO Then gError 180449
                                
                lErro = Comando_BuscarProximo(alComando(5))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180450
            Loop
        End If
        
        'Seleciona a tabela de Filiais Contatos (para modificar a Matriz e pegar os Codigos de Enderecos em tFilial.lEndereco)
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
            lErro = Comando_ExecutarPos(alComando(3), "SELECT CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador,InscricaoSuframa, RG FROM  FiliaisContatos WHERE CodContatoLoja = ? AND CodFilialLoja = ?", 0, _
                tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.sInscricaoSuframa, tFilialContato.sRG, objContato.lCodigoLoja, FILIAL_MATRIZ)
        Else
            lErro = Comando_ExecutarPos(alComando(3), "SELECT CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador,InscricaoSuframa, RG FROM  FiliaisContatos WHERE CodContato = ? AND CodFilial = ?", 0, _
                tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.sInscricaoSuframa, tFilialContato.sRG, objContato.lCodigo, FILIAL_MATRIZ)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 180451

        'Verifica se filial Contato existe
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180452

        'Faz a modificação dos três Endereços
        If colEndereco.Count > 0 Then colEndereco.Item(1).lCodigo = tFilialContato.lEndereco
        If colEndereco.Count > 1 Then colEndereco.Item(2).lCodigo = tFilialContato.lEnderecoEntrega
        If colEndereco.Count > 2 Then colEndereco.Item(3).lCodigo = tFilialContato.lEnderecoCobranca

        'Atualiza os endereços
        For iIndice = 1 To colEndereco.Count

            If colEndereco.Item(iIndice).lCodigo = 0 Then

                lErro = CF("Endereco_Automatico", lCodEndereco)
                If lErro <> SUCESSO Then gError 180453
    
                colEndereco.Item(iIndice).lCodigo = lCodEndereco

            End If

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 180454

        Next

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then

            'Atualiza a tabela de Contatos
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE Contatos SET RazaoSocial = ?, NomeReduzido = ?, Ativo = ?, Bloqueado=?", alComando(1), _
                objContato.sRazaoSocial, objContato.sNomeReduzido, objContato.iAtivo, objContato.iBloqueado)
            If lErro <> AD_SQL_SUCESSO Then gError 180455
    
            'Atualiza a tabela de FiliaisContatos
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE FiliaisContatos SET CGC = ?, RG=?, Ativo = ?, Endereco = ?", alComando(3), _
                objContato.sCgc, objContato.sRG, objContato.iAtivo, tFilialContato.lEndereco)
            If lErro <> AD_SQL_SUCESSO Then gError 180456

        Else
        
            'Atualiza a tabela de Contatos
            lErro = Comando_ExecutarPos(alComando(2), "UPDATE Contatos SET RazaoSocial = ?, NomeReduzido = ?, Tipo = ?, Observacao = ?, LimiteCredito = ?, CondicaoPagto = ?, Desconto = ?, CodMensagem = ?, TabelaPreco = ?, CodPadraoCobranca = ?, Ativo = ?, Bloqueado = ? ", alComando(1), _
                objContato.sRazaoSocial, objContato.sNomeReduzido, objContato.iTipo, objContato.sObservacao, objContato.dLimiteCredito, objContato.iCondicaoPagto, objContato.dDesconto, objContato.iCodMensagem, objContato.iTabelaPreco, objContato.iCodPadraoCobranca, objContato.iAtivo, objContato.iBloqueado)
            If lErro <> AD_SQL_SUCESSO Then gError 180457
    
            'Atualiza a tabela de FiliaisContatos
            lErro = Comando_ExecutarPos(alComando(0), "UPDATE FiliaisContatos SET CGC = ?, InscricaoEstadual = ?, InscricaoMunicipal = ?, Endereco = ?, EnderecoEntrega = ?, EnderecoCobranca = ?, CodTransportadora = ?, Observacao = ?, ContaContabil = ?, Vendedor = ?, ComissaoVendas = ?, Regiao = ?, FreqVisitas = ?, DataUltVisita = ?, CodCobrador = ?, TipoFrete = ?,InscricaoSuframa=?, RG=?, Ativo = ?, CodTranspRedesp=?, DetPagFrete=?, Guia=?", alComando(3), _
                objContato.sCgc, objContato.sInscricaoEstadual, objContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, objContato.iCodTransportadora, objContato.sObservacao2, objContato.sContaContabil, objContato.iVendedor, objContato.dComissaoVendas, objContato.iRegiao, objContato.iFreqVisitas, objContato.dtDataUltVisita, objContato.iCodCobrador, objContato.iTipoFrete, objContato.sInscricaoSuframa, objContato.sRG, objContato.iAtivo, objContato.iCodTranspRedesp, objContato.iDetPagFrete, objContato.sGuia)
            If lErro <> AD_SQL_SUCESSO Then gError 180458

        End If

        objFilialContato.lCodContato = objContato.lCodigo
        
        objFilialContato.iCodFilial = FILIAL_MATRIZ
        
        lErro = FilialContato_Atualiza_Categorias(objFilialContato, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 180459

    End If
        
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Contato_Grava_EmTrans = SUCESSO

    Exit Function

Erro_Contato_Grava_EmTrans:

    Contato_Grava_EmTrans = gErr

    Select Case gErr
            
        Case 180434
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180435, 180436, 180437, 180440, 180441, 180439
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)

        Case 180438
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RED_DUPLICADO", gErr)

        Case 180442, 180443, 180446, 180453, 180454, 180459, 180543

        Case 180444
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CONTATOS", gErr, lCodigoContato)
        
        Case 180445
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIAISCONTATOS", gErr)

        Case 180447, 180448, 180450, 180452
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr)

        Case 180449, 180456, 180458
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISCONTATOS", gErr)

        Case 180451
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_SEM_FILIAL", gErr, lCodigoContato)

        Case 180455, 180457
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONTATO", gErr, lCodigoContato)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180460)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FilialContato_Insere_Categorias(ByVal objFilialContato As ClassFilialContato, ByVal colCategoriaItem As Collection) As Long
'Percorre as Categorias na coleção incluindo-as no BD
'Chamada DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(0 To 1) As Long
Dim sCategoria As String, sItem As String
Dim objFilialConCategoria As New ClassFilialConCategoria

On Error GoTo Erro_FilialContato_Insere_Categorias

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180461
    Next
    
    'Para cada registro da coleção
    For Each objFilialConCategoria In colCategoriaItem

        lErro = CategoriaClienteItem_Lock(objFilialConCategoria.sCategoria, objFilialConCategoria.sItem, alComando(0))
        If lErro <> SUCESSO Then gError 180462
        
        'Insere a filial Contato categoria
        lErro = Comando_Executar(alComando(1), "INSERT INTO FilialContatoCategorias (Contato, Filial, Categoria, Item) VALUES(?,?,?,?)", objFilialContato.lCodContato, objFilialContato.iCodFilial, objFilialConCategoria.sCategoria, objFilialConCategoria.sItem)
        If lErro <> AD_SQL_SUCESSO Then gError 180463

    Next

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FilialContato_Insere_Categorias = SUCESSO

    Exit Function

Erro_FilialContato_Insere_Categorias:

    FilialContato_Insere_Categorias = gErr

    Select Case gErr

        Case 180461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180462

        Case 180463
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIALCONTATOCATEGORIAS", gErr, objFilialContato.lCodContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180464)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FilialContato_Atualiza_Categorias(ByVal objFilialContato As ClassFilialContato, ByVal colCategoriaItem As Collection) As Long
'atualiza as categorias da filial Contato no bd
'Chamada DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim lContato As Long
Dim sCategoria As String
Dim sItem As String
Dim sCategoria1 As String
Dim sItem1 As String
Dim iIndice As Integer
Dim objFilialConCategoria As ClassFilialConCategoria

On Error GoTo Erro_FilialContato_Atualiza_Categorias

    'Abre os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180461
    Next

    sCategoria = String(STRING_CATEGORIACLIENTE_CATEGORIA, 0)
    sItem = String(STRING_CATEGORIACLIENTEITEM_ITEM, 0)

    'Percorre as categorias associadas a filial que já estão no bd
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Categoria, Item FROM FilialContatoCategorias WHERE Contato = ? AND Filial = ?", 0, sCategoria, sItem, objFilialContato.lCodContato, objFilialContato.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 180462

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180463

    Do While lErro <> AD_SQL_SEM_DADOS

        iIndice = FilialContato_Grava_CategoriaNaColecao(colCategoriaItem, sCategoria)

        If iIndice <> 0 Then
            
            Set objFilialConCategoria = colCategoriaItem.Item(iIndice)
            
            'Se houver sido alterado o Item
            If objFilialConCategoria.sItem <> sItem Then

                lErro = CategoriaClienteItem_Lock(sCategoria, objFilialConCategoria.sItem, alComando(1))
                If lErro <> SUCESSO Then gError 180464
                
                'Atualiza a categoria da filial do Contato
                lErro = Comando_ExecutarPos(alComando(2), "UPDATE FilialContatoCategorias SET Item = ?", alComando(0), objFilialConCategoria.sItem)
                If lErro <> AD_SQL_SUCESSO Then gError 180465

            End If

            'Retira da coleção
            colCategoriaItem.Remove (iIndice)

        Else
            
            'Excluir registro em FilialCliCategoria
            lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM FilialContatoCategorias", alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 180466
            
        End If

        'Tenta ler a próxima categoria do Contato da filial
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180467

    Loop

    'Chama FilialContato_Insere_Categorias
    lErro = FilialContato_Insere_Categorias(objFilialContato, colCategoriaItem)
    If lErro <> SUCESSO Then gError 180468

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FilialContato_Atualiza_Categorias = SUCESSO

    Exit Function

Erro_FilialContato_Atualiza_Categorias:

    FilialContato_Atualiza_Categorias = gErr

    Select Case gErr

        Case 180461
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180462, 180463, 180467
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCONTATOCATEGORIAS2", gErr, objFilialConCategoria.lContato)

        Case 180465
            Call Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_FILIALCONTATOCATEGORIAS", gErr, objFilialConCategoria.lContato)
        
        Case 180464, 180468

        Case 180466
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCONTATOCATEGORIAS1", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180469)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Contato_Exclui(ByVal objContato As ClassContatos, Optional ByVal bExcluiEnd As Boolean = True) As Long
'Exclui o Contato passado como parâmetro

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_Contato_Exclui

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180430

    lErro = Contato_Exclui_EmTrans(objContato, bExcluiEnd)
    If lErro <> SUCESSO Then gError 180431
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180432

    Contato_Exclui = SUCESSO

    Exit Function

Erro_Contato_Exclui:

    Contato_Exclui = gErr

    Select Case gErr

        Case 180430
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 180431

        Case 180432
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180433)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function Contato_Exclui_EmTrans(ByVal objContato As ClassContatos, Optional ByVal bExcluiEnd As Boolean = True) As Long
'Exclui o Contato e suas filiais

Dim lErro As Long
Dim alComando(2) As Long
Dim lCodigo As Long
Dim lPedido As Long
Dim lNumIntCheque As Long
Dim iIndice As Integer
Dim iFilialEmpresa As Integer
Dim objLog As New ClassLog
Dim colEndereco As New Collection
Dim lCodigoLoja As Long
Dim lCodigoContato As Long
Dim iTransferido As Integer
Dim lCodRelacionamento As Long

On Error GoTo Erro_Contato_Exclui_EmTrans

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180470
    Next
    
    'Inserido por Luiz Nogueira em 27/10/03
    'Verifica se existe algum relacionamento de cliente vinculado a esse cliente
    lErro = CF("RelacionamentoContatos_Verifica_Cliente", objContato.lCodigo, lCodRelacionamento)
    If lErro <> SUCESSO And lErro <> 181055 Then gError 102696
    
    'Se existir => erro
    If lErro = SUCESSO Then gError 102697
    
    'guarda o código para as msgs de erro
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lCodigoContato = objContato.lCodigoLoja
    Else
        lCodigoContato = objContato.lCodigo
    End If
    
    'Seleciona o Contato a ser excluído
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objContato.lCodigo = 0 Then
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Transferido FROM Contatos WHERE CodigoLoja =?", 0, iTransferido, objContato.lCodigoLoja)
    Else
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Transferido FROM Contatos WHERE Codigo =?", 0, iTransferido, objContato.lCodigo)
    End If
    
    If lErro <> AD_SQL_SUCESSO Then gError 180472
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180473

    If lErro = AD_SQL_SEM_DADOS Then gError 180474

    'se o registro foi transferido ==> nao pode excluir
    If iTransferido = TRANSFERIDO Then gError 180475

    'Lock do Contato
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 180476

    'Lock das Filiais
    lErro = FiliaisContatos_Lock(lCodigoContato)
    If lErro <> SUCESSO Then gError 180477

    If giLocalOperacao <> LOCALOPERACAO_CAIXA_CENTRAL Then
    
        lErro = ClienteFContatos_Exclui_Cliente_EmTrans(objContato.lCodigo)
        If lErro <> SUCESSO Then gError 181149
    
        'Exclui as Categorias das Filiais do Contato
        lErro = Contato_Exclui_Categorias(lCodigoContato)
        If lErro <> SUCESSO Then gError 180478
    
    End If
    
    'Exclui as Filiais do Contato
    lErro = FiliaisContatos_Exclui(lCodigoContato, bExcluiEnd)
    If lErro <> SUCESSO Then gError 180479

    'Exclui o Contato
    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM Contatos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 180480
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Contato_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_Contato_Exclui_EmTrans:

    Contato_Exclui_EmTrans = gErr

    Select Case gErr
        
        Case 102696
   
        Case 102697
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_REL_RELACIONAMENTOCLIENTE", gErr, objContato.lCodigo, lCodRelacionamento)
   
        Case 180470
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180472, 180473
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)

        Case 180474
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, lCodigoContato)
            
        Case 180475, 181149
        
        Case 180476
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTATOS", gErr, lCodigoContato)
        
        Case 180477, 180478, 180479
            'Tratados nas rotinas chamadas

        Case 180480
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CONTATO", gErr, lCodigoContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180482)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function Contato_Exclui_Categorias(ByVal lCodContato As Long) As Long
'Exclui as Categorias das Filiais do Contato cujo código foi passado como parâmetro.
'Chamada por Contato_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim iFilial As Integer

On Error GoTo Erro_Contato_Exclui_Categorias

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180483
    Next

    'Seleciona registros de FilialContatoCategorias a serem excluídos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Filial FROM FilialContatoCategorias WHERE Contato =?", 0, iFilial, lCodContato)
    If lErro <> AD_SQL_SUCESSO Then gError 180484

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180485

    Do While lErro = AD_SQL_SUCESSO

        'Exclui registro de FilialContatoCategorias ligado ao Contato
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM FilialContatoCategorias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 180486

        'Traz o próximo registro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180487

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Contato_Exclui_Categorias = SUCESSO

    Exit Function

Erro_Contato_Exclui_Categorias:

    Contato_Exclui_Categorias = gErr

    Select Case gErr

        Case 180483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180484, 180485, 180487
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCONTATOCATEGORIAS2", gErr, lCodContato)

        Case 180486
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCONTATOCATEGORIAS", gErr, lCodContato)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180488)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FilialContato_Exclui(ByVal objFilialContato As ClassFilialContato) As Long
'Exclui a Filial Contato e seus Enderecos

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 5) As Long
Dim lTransacao As Long
Dim lCodigo As Long
Dim lPedido As Long
Dim iFilial As Integer
Dim lCodigoContato As Long
Dim iAtivo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim lCodRelacionamento As Long

On Error GoTo Erro_FilialContato_Exclui

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180489
    Next

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180490

    'Inserido por Luiz Nogueira em 27/10/03
    'Verifica se existe algum relacionamento de cliente vinculado a essa filial do cliente
    lErro = CF("RelacionamentoContatos_Verifica_Cliente", objFilialContato.lCodContato, lCodRelacionamento, objFilialContato.iCodFilial)
    If lErro <> SUCESSO And lErro <> 181055 Then gError 102714
    
    'Se existir => erro
    If lErro = SUCESSO Then gError 102715

    'Seleciona a Filial Contato a ser excluída
    lErro = Comando_ExecutarPos(alComando(1), "SELECT CodContato FROM FiliaisContatos WHERE CodContato =? AND CodFilial = ?", 0, lCodigo, objFilialContato.lCodContato, objFilialContato.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 180491

    'Lê a Filial
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180492
    
    If lErro = AD_SQL_SEM_DADOS Then gError 180493

    'Faz lock da FilialContato
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 180494

    'Exclui as Categorias das Filiais do Contato
    lErro = FilialContato_Exclui_Categorias(objFilialContato)
    If lErro <> SUCESSO Then gError 180495
    
    'Exclui a Filial Contato
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM FiliaisContatos", alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 180496
    
    'tenta achar outra filial q esteja ativa--> se não achar, desativa o Contato
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
        lErro = Comando_Executar(alComando(3), "SELECT Ativo FROM FiliaisContatos WHERE CodContatoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialContato.lCodContatoLoja, CLIENTE_ATIVO)
    Else
        lErro = Comando_ExecutarPos(alComando(3), "SELECT Ativo FROM FiliaisContatos WHERE CodContato = ? AND Ativo = ?", 0, Ativo, objFilialContato.lCodContato, CLIENTE_ATIVO)
    End If
    If lErro <> AD_SQL_SUCESSO Then gError 180497

    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180498
    
    If lErro = AD_SQL_SEM_DADOS Then
        'Seleciona o Contatos no BD
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            lCodigoContato = objFilialContato.lCodContatoLoja
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Ativo FROM Contatos WHERE CodigoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialContato.lCodContatoLoja, CLIENTE_ATIVO)
        Else
            lCodigoContato = objFilialContato.lCodContato
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Ativo FROM Contatos WHERE Codigo = ? AND Ativo = ?", 0, Ativo, objFilialContato.lCodContato, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 180499

        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180500
        
        If lErro = AD_SQL_SUCESSO Then
            'Atualiza a tabela de Contatos
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE Contatos SET Ativo = ?", alComando(4), DESMARCADO)
            If lErro <> AD_SQL_SUCESSO Then gError 180501
        End If
    End If
    
    Call FilialContato_Transfere_FilialCliente(objFilialContato, objFilialCliente)
 
    'Exclui os Endereços da FilialContato (principal, entrega e cobrança)
    lErro = CF("Enderecos_Exclui_FilialCliente", objFilialCliente)
    If lErro <> SUCESSO Then gError 180502

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180503

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FilialContato_Exclui = SUCESSO

    Exit Function

Erro_FilialContato_Exclui:

    FilialContato_Exclui = gErr

    Select Case gErr

        Case 102714

        Case 102715
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_REL_RELACIONAMENTOCLIENTE", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial, lCodRelacionamento)
            
        Case 180489
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180490
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 180491, 180492, 180497, 180498
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180493
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCONTATO_NAO_CADASTRADA", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180494
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAISCONTATOS", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180496
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCONTATO", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180499, 180500
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)
            
        Case 180501
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONTATO", gErr, lCodigoContato)

        Case 180502
            'Tratados pela rotina chamada

        Case 180503
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180504)

    End Select

    Call Transacao_Rollback

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Private Function FilialContato_Exclui_Categorias(ByVal objFilialContato As ClassFilialContato) As Long
'Exclui as Categorias da Filial do Contato cujo código da filial e do Contato foi passado como parâmetro.
'Chamada por FilialContato_Exclui DENTRO DE TRANSAÇÃO.

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim iFilial As Integer

On Error GoTo Erro_FilialContato_Exclui_Categorias

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180505
    Next

    'Seleciona as Categorias da Filial do Contato a serem excluídas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Filial FROM FilialContatoCategorias WHERE Contato =? AND Filial = ?", 0, iFilial, objFilialContato.lCodContato, objFilialContato.iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then gError 180506

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180507

    'Exclui as Filiais do Contato
    Do While lErro = AD_SQL_SUCESSO

        'Exclui a Filial
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM FilialContatoCategorias", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 180508

        'Traz a próxima Categoria da Filial do Contato
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180509

    Loop

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    FilialContato_Exclui_Categorias = SUCESSO

    Exit Function

Erro_FilialContato_Exclui_Categorias:

    FilialContato_Exclui_Categorias = gErr

    Select Case gErr

        Case 180505
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180506, 180507, 180509
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALCONTATOCATEGORIAS3", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180508
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIALCONTATOCATEGORIAS1", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180510)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FilialContatos_Grava(ByVal objFilialContato As ClassFilialContato, ByVal colEndereco As Collection) As Long
'Grava a Filial Contato e seus 3 tipos de Enderecos

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_FilialContatos_Grava

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180511

    lErro = FilialContatos_Grava_EmTrans(objFilialContato, colEndereco)
    If lErro <> SUCESSO Then gError 180512
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180513

    FilialContatos_Grava = SUCESSO

    Exit Function

Erro_FilialContatos_Grava:

    FilialContatos_Grava = gErr

    Select Case gErr

        Case 180511
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 180512
        
        Case 180513
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180514)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function FilialContatos_Grava_EmTrans(ByVal objFilialContato As ClassFilialContato, ByVal colEndereco As Collection) As Long
'Grava a Filial Contato e seus 3 tipos de Enderecos

Dim lErro As Long
Dim iIndice As Integer
Dim alComando(1 To 6) As Long
Dim lCodigo As Long
Dim lCodEndereco As Long
Dim tFilialContato As typeFilialContato
Dim objEndereco As ClassEndereco
Dim lEndereco(1 To 3) As Long
Dim colCategoriaItemCopia As New Collection
Dim objFilialConCategoria As ClassFilialConCategoria
Dim objPais As New ClassPais
Dim objContato As New ClassContatos
Dim objLog As New ClassLog
Dim iAtivo As Integer
Dim lCodigoContato As Long

On Error GoTo Erro_FilialContatos_Grava_EmTrans

    tFilialContato.sNome = String(STRING_FILIAL_CLIENTE_NOME, 0)
    tFilialContato.sCgc = String(STRING_CGC, 0)
    tFilialContato.sRG = String(STRING_RG, 0)
    tFilialContato.sInscricaoEstadual = String(STRING_INSCR_EST, 0)
    tFilialContato.sInscricaoMunicipal = String(STRING_INSCR_MUN, 0)
    tFilialContato.sInscricaoSuframa = String(STRING_INSCR_SUF, 0)
    tFilialContato.sObservacao = String(STRING_CLIENTE_OBSERVACAO, 0)
    tFilialContato.sContaContabil = String(STRING_CONTA, 0)

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180515
    Next

    'Cria uma cópia "de trabalho" da coleção passada como parâmetro
    For Each objFilialConCategoria In objFilialContato.colCategoriaItem
        colCategoriaItemCopia.Add objFilialConCategoria
    Next
    
    'Faz os locks nas tabelas ligadas ao cadastramento de uma nova Filial Contato
    lErro = FiliaisContatos_Le_Tabelas_Lock(objFilialContato)
    If lErro <> SUCESSO Then gError 180516

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.iCodFilialLoja <> 0 Then

        If objFilialContato.lCodContato <> 0 Then
            'Procura por um outro Nome já cadastrado
            lErro = Comando_Executar(alComando(1), "SELECT CodContato FROM FiliaisContatos WHERE CodContato = ? AND CodFilialLoja <> ? AND Nome = ? ", lCodigo, objFilialContato.lCodContato, objFilialContato.iCodFilialLoja, objFilialContato.sNome)
            If lErro <> AD_SQL_SUCESSO Then gError 180517
        Else
            'Procura por um outro Nome já cadastrado
            lErro = Comando_Executar(alComando(1), "SELECT CodContato FROM FiliaisContatos WHERE CodContatoLoja = ? AND CodFilialLoja <> ? AND Nome = ? ", lCodigo, objFilialContato.lCodContatoLoja, objFilialContato.iCodFilialLoja, objFilialContato.sNome)
            If lErro <> AD_SQL_SUCESSO Then gError 180518
        End If
    
    Else
    
        'Procura por um outro Nome já cadastrado
        lErro = Comando_Executar(alComando(1), "SELECT CodContato FROM FiliaisContatos WHERE CodContato = ? AND CodFilial <> ? AND Nome = ? ", lCodigo, objFilialContato.lCodContato, objFilialContato.iCodFilial, objFilialContato.sNome)
        If lErro <> AD_SQL_SUCESSO Then gError 180519

    End If

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180520

    'Se existir outro Nome acusa erro
    If lErro <> AD_SQL_SEM_DADOS Then gError 180521
       
    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.iCodFilialLoja <> 0 Then
    
        If objFilialContato.lCodContato <> 0 Then
            'Pesquisa a Filial no BD
            lErro = Comando_ExecutarPos(alComando(2), "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisContatos WHERE CodContato = ? AND CodFilialLoja = ? AND FilialEmpresaLoja = ?", 0, _
                tFilialContato.iCodFilial, tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sInscricaoSuframa, objFilialContato.lCodContato, objFilialContato.iCodFilialLoja, objFilialContato.iFilialEmpresaLoja)
            If lErro <> AD_SQL_SUCESSO Then gError 180522
        Else
            'Pesquisa a Filial no BD
            lErro = Comando_ExecutarPos(alComando(2), "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisContatos WHERE CodContatoLoja = ? AND CodFilialLoja = ? AND FilialEmpresaLoja = ?", 0, _
                tFilialContato.iCodFilial, tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sInscricaoSuframa, objFilialContato.lCodContatoLoja, objFilialContato.iCodFilialLoja, objFilialContato.iFilialEmpresaLoja)
            If lErro <> AD_SQL_SUCESSO Then gError 180523
        End If
    
    Else
    
        'Pesquisa a Filial no BD
        lErro = Comando_ExecutarPos(alComando(2), "SELECT CodFilial, Nome, CGC, InscricaoEstadual, InscricaoMunicipal, Endereco, EnderecoEntrega, EnderecoCobranca, CodTransportadora, Observacao, ContaContabil, Vendedor, ComissaoVendas, Regiao, FreqVisitas, DataUltVisita, CodCobrador, TipoFrete,InscricaoSuframa FROM FiliaisContatos WHERE CodContato = ? AND CodFilial = ?", 0, _
            tFilialContato.iCodFilial, tFilialContato.sNome, tFilialContato.sCgc, tFilialContato.sInscricaoEstadual, tFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, tFilialContato.iCodTransportadora, tFilialContato.sObservacao, tFilialContato.sContaContabil, tFilialContato.iVendedor, tFilialContato.dComissaoVendas, tFilialContato.iRegiao, tFilialContato.iFreqVisitas, tFilialContato.dtDataUltVisita, tFilialContato.iCodCobrador, tFilialContato.iTipoFrete, tFilialContato.sInscricaoSuframa, objFilialContato.lCodContato, objFilialContato.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 180524

    End If
    
    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180525
    
    'Filial Contato não está cadastrada - GRAVACAO
    If lErro = AD_SQL_SEM_DADOS Then
        
        'Grava os Três Novos Enderecos de Filiais
        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Automatico", lCodEndereco)
            If lErro <> SUCESSO Then gError 180526

            colEndereco.Item(iIndice).lCodigo = lCodEndereco

            lEndereco(iIndice) = lCodEndereco

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 180527

        Next

        'Insere a FilialContato na tabela FiliaisContatos
        lErro = Comando_Executar(alComando(3), "INSERT INTO FiliaisContatos (CodContato,CodFilial,Nome,CGC,InscricaoEstadual,InscricaoMunicipal,Endereco,EnderecoEntrega,EnderecoCobranca,CodTransportadora,Observacao,ContaContabil,Vendedor,ComissaoVendas,Regiao,FreqVisitas,DataUltVisita,CodCobrador,TipoFrete,InscricaoSuframa,RG, CodFilialLoja, FilialEmpresaLoja, CodContatoLoja, Ativo, CodTranspRedesp, DetPagFrete, Guia, CodMensagem) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            objFilialContato.lCodContato, objFilialContato.iCodFilial, objFilialContato.sNome, objFilialContato.sCgc, objFilialContato.sInscricaoEstadual, objFilialContato.sInscricaoMunicipal, lEndereco(1), lEndereco(2), lEndereco(3), objFilialContato.iCodTransportadora, objFilialContato.sObservacao, objFilialContato.sContaContabil, objFilialContato.iVendedor, objFilialContato.dComissaoVendas, objFilialContato.iRegiao, objFilialContato.iFreqVisitas, objFilialContato.dtDataUltVisita, objFilialContato.iCodCobrador, objFilialContato.iTipoFrete, objFilialContato.sInscricaoSuframa, objFilialContato.sRG, objFilialContato.iCodFilialLoja, objFilialContato.iFilialEmpresaLoja, objFilialContato.lCodContatoLoja, objFilialContato.iAtivo, objFilialContato.iCodTranspRedesp, objFilialContato.iDetPagFrete, objFilialContato.sGuia, objFilialContato.iCodMensagem)
        If lErro <> AD_SQL_SUCESSO Then gError 180528

        'Insere as categorias da filial do Contato em FilialContatosCategorias
        lErro = FilialContato_Insere_Categorias(objFilialContato, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 180529

    Else 'FilialContato já existe - MODIFICAÇÃO

        'Grava os Três Enderecos de Filiais
        
        If colEndereco.Count > 0 Then colEndereco.Item(1).lCodigo = tFilialContato.lEndereco
        If colEndereco.Count > 1 Then colEndereco.Item(2).lCodigo = tFilialContato.lEnderecoEntrega
        If colEndereco.Count > 2 Then colEndereco.Item(3).lCodigo = tFilialContato.lEnderecoCobranca

        For iIndice = 1 To colEndereco.Count

            lErro = CF("Endereco_Grava", colEndereco.Item(iIndice))
            If lErro <> SUCESSO Then gError 180530

        Next

        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL Then
            'Atualiza a tabela de FiliaisContatos
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE FiliaisContatos SET Nome = ?, CGC = ?, RG = ?, Ativo = ?", alComando(2), _
                objFilialContato.sNome, objFilialContato.sCgc, objFilialContato.sRG, objFilialContato.iAtivo)
            If lErro <> AD_SQL_SUCESSO Then gError 180531
        Else
            'Atualiza a tabela de FiliaisContatos
            lErro = Comando_ExecutarPos(alComando(3), "UPDATE FiliaisContatos SET Nome = ?, CGC = ?, InscricaoEstadual = ?, InscricaoMunicipal = ?, Endereco = ?, EnderecoEntrega = ?, EnderecoCobranca = ?, CodTransportadora = ?, Observacao = ?, ContaContabil = ?, Vendedor = ?, ComissaoVendas = ?, Regiao = ?, FreqVisitas = ?, DataUltVisita = ?, CodCobrador = ?, TipoFrete = ?,InscricaoSuframa=?, RG=?, Ativo=?, CodTranspRedesp=?, DetPagFrete=?, Guia=?, CodMensagem = ?", alComando(2), _
                objFilialContato.sNome, objFilialContato.sCgc, objFilialContato.sInscricaoEstadual, objFilialContato.sInscricaoMunicipal, tFilialContato.lEndereco, tFilialContato.lEnderecoEntrega, tFilialContato.lEnderecoCobranca, objFilialContato.iCodTransportadora, objFilialContato.sObservacao, objFilialContato.sContaContabil, objFilialContato.iVendedor, objFilialContato.dComissaoVendas, objFilialContato.iRegiao, objFilialContato.iFreqVisitas, objFilialContato.dtDataUltVisita, objFilialContato.iCodCobrador, objFilialContato.iTipoFrete, objFilialContato.sInscricaoSuframa, objFilialContato.sRG, objFilialContato.iAtivo, objFilialContato.iCodTranspRedesp, objFilialContato.iDetPagFrete, objFilialContato.sGuia, objFilialContato.iCodMensagem)
            If lErro <> AD_SQL_SUCESSO Then gError 180532
        End If

        'Atualiza as categorias da filial do Contato
        lErro = FilialContato_Atualiza_Categorias(objFilialContato, colCategoriaItemCopia)
        If lErro <> SUCESSO Then gError 180533

    End If
    
    'se ativo uma filial--> ativo o Contato desta filial
    If objFilialContato.iAtivo = CLIENTE_ATIVO Then
        'Seleciona o Contatos no BD q está desativado e ativa
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato = 0 Then
            lCodigoContato = objFilialContato.lCodContatoLoja
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Ativo FROM Contatos WHERE CodigoLoja = ? AND Ativo <> ?", 0, iAtivo, objFilialContato.lCodContatoLoja, CLIENTE_ATIVO)
        Else
            lCodigoContato = objFilialContato.lCodContato
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Ativo FROM Contatos WHERE Codigo = ? AND Ativo <> ?", 0, Ativo, objFilialContato.lCodContato, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 180534

        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180535
        
        If lErro = AD_SQL_SUCESSO Then
            'Atualiza a tabela de Contatos
            lErro = Comando_ExecutarPos(alComando(5), "UPDATE Contatos SET Ativo = ?", alComando(4), objFilialContato.iAtivo)
            If lErro <> AD_SQL_SUCESSO Then gError 180536
        End If
    Else
        'tenta achar outra filial q esteja ativa--> se não achar, desativa o Contato
        If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato = 0 Then
            lErro = Comando_Executar(alComando(4), "SELECT Ativo FROM FiliaisContatos WHERE CodContatoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialContato.lCodContatoLoja, CLIENTE_ATIVO)
        Else
            lErro = Comando_ExecutarPos(alComando(4), "SELECT Ativo FROM FiliaisContatos WHERE CodContato = ? AND Ativo = ?", 0, Ativo, objFilialContato.lCodContato, CLIENTE_ATIVO)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 180537

        lErro = Comando_BuscarPrimeiro(alComando(4))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180538
        
        If lErro = AD_SQL_SEM_DADOS Then
            'Seleciona o Contatos no BD
            If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato = 0 Then
                lCodigoContato = objFilialContato.lCodContatoLoja
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM Contatos WHERE CodigoLoja = ? AND Ativo = ?", 0, iAtivo, objFilialContato.lCodContatoLoja, CLIENTE_ATIVO)
            Else
                lCodigoContato = objFilialContato.lCodContato
                lErro = Comando_ExecutarPos(alComando(5), "SELECT Ativo FROM Contatos WHERE Codigo = ? AND Ativo = ?", 0, Ativo, objFilialContato.lCodContato, CLIENTE_ATIVO)
            End If
            If lErro <> AD_SQL_SUCESSO Then gError 180539
    
            lErro = Comando_BuscarPrimeiro(alComando(5))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180540
            
            If lErro = AD_SQL_SUCESSO Then
                'Atualiza a tabela de Contatos
                lErro = Comando_ExecutarPos(alComando(6), "UPDATE Contatos SET Ativo = ?", alComando(5), objFilialContato.iAtivo)
                If lErro <> AD_SQL_SUCESSO Then gError 180541
            End If
        End If
    End If

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    FilialContatos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_FilialContatos_Grava_EmTrans:

    FilialContatos_Grava_EmTrans = gErr

    Select Case gErr

        Case 180515
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180516, 180526, 180527, 180529, 180530, 180533

        Case 180517, 180518, 180519, 180520, 180525, 180537, 180538
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS2", gErr)

        Case 180521
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCONTATO_NOME_DUPLICADO", gErr, objFilialContato.sNome)

       Case 180522
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilialLoja)

        Case 180523
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr, objFilialContato.lCodContatoLoja, objFilialContato.iCodFilialLoja)

        Case 180524
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr, objFilialContato.lCodContato, objFilialContato.iCodFilial)

        Case 180528
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_FILIAISContatoS", gErr)

        Case 180531, 180532
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISContatoS", gErr)

        Case 180534, 180535, 180539, 180540
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)
                    
        Case 180536, 180541
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONTATO", gErr, lCodigoContato)
             
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180542)

    End Select

    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Public Function FilialContato_Transfere_FilialCliente(ByVal objFilialContato As ClassFilialContato, ByVal objFilialCliente As ClassFilialCliente) As Long
'Grava a Filial Contato e seus 3 tipos de Enderecos

Dim lErro As Long
Dim objFilialCliCategoria As ClassFilialCliCategoria
Dim objFilialConCategoria As ClassFilialConCategoria

On Error GoTo Erro_FilialContato_Transfere_FilialCliente
    
    For Each objFilialConCategoria In objFilialContato.colCategoriaItem
    
        Set objFilialCliCategoria = New ClassFilialCliCategoria
        
        objFilialCliCategoria.sCategoria = objFilialConCategoria.sCategoria
        objFilialCliCategoria.sItem = objFilialConCategoria.sItem
        
        objFilialCliente.colCategoriaItem.Add objFilialCliCategoria

    Next
    
    objFilialCliente.dComissaoVendas = objFilialContato.dComissaoVendas
    objFilialCliente.dLimiteCredito = objFilialContato.dLimiteCredito
    objFilialCliente.dtDataUltVisita = objFilialContato.dtDataUltVisita
    objFilialCliente.iAtivo = objFilialContato.iAtivo
    objFilialCliente.iCodCobrador = objFilialContato.iCodCobrador
    objFilialCliente.iCodFilial = objFilialContato.iCodFilial
    objFilialCliente.iCodFilialLoja = objFilialContato.iCodFilialLoja
    objFilialCliente.iCodMensagem = objFilialContato.iCodMensagem
    objFilialCliente.iCodTransportadora = objFilialContato.iCodTransportadora
    objFilialCliente.iCodTranspRedesp = objFilialContato.iCodTranspRedesp
    objFilialCliente.iDetPagFrete = objFilialContato.iDetPagFrete
    objFilialCliente.iFilialEmpresaLoja = objFilialContato.iFilialEmpresaLoja
    objFilialCliente.iFreqVisitas = objFilialContato.iFreqVisitas
    objFilialCliente.iICMSBaseCalculoComIPI = objFilialContato.iICMSBaseCalculoComIPI
    objFilialCliente.iRegiao = objFilialContato.iRegiao
    objFilialCliente.iTipo = objFilialContato.iTipo
    objFilialCliente.iTipoFrete = objFilialContato.iTipoFrete
    objFilialCliente.iTransferido = objFilialContato.iTransferido
    objFilialCliente.iVendedor = objFilialContato.iVendedor
'    objFilialCliente.lCodCliente = objFilialContato.lCodCliente
'    objFilialCliente.lCodClienteLoja = objFilialContato.lCodClienteLoja
    objFilialCliente.lEndereco = objFilialContato.lEndereco
    objFilialCliente.lEnderecoCobranca = objFilialContato.lEnderecoCobranca
    objFilialCliente.lEnderecoEntrega = objFilialContato.lEnderecoEntrega
    objFilialCliente.lRevendedor = objFilialContato.lRevendedor
    objFilialCliente.sCgc = objFilialContato.sCgc
    objFilialCliente.sContaContabil = objFilialContato.sContaContabil
    objFilialCliente.sGuia = objFilialContato.sGuia
    objFilialCliente.sInscricaoEstadual = objFilialContato.sInscricaoEstadual
    objFilialCliente.sInscricaoMunicipal = objFilialContato.sInscricaoMunicipal
    objFilialCliente.sInscricaoSuframa = objFilialContato.sInscricaoSuframa
    objFilialCliente.sNome = objFilialContato.sNome
    objFilialCliente.sNomeReduzidoCli = objFilialContato.sNomeReduzidoCli
    objFilialCliente.sObservacao = objFilialContato.sObservacao
    objFilialCliente.sRazaoSocialCli = objFilialContato.sRazaoSocialCli
    objFilialCliente.sRG = objFilialContato.sRG
    objFilialCliente.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL

    FilialContato_Transfere_FilialCliente = SUCESSO

    Exit Function

Erro_FilialContato_Transfere_FilialCliente:

    FilialContato_Transfere_FilialCliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180514)

    End Select

    Exit Function

End Function

Public Function Contato_Transfere_Cliente(ByVal objContato As ClassContatos, ByVal objCliente As ClassCliente) As Long
'Grava a  Contato e seus 3 tipos de Enderecos

Dim lErro As Long

On Error GoTo Erro_Contato_Transfere_Cliente

    Set objCliente.colCategoriaItem = objContato.colCategoriaItem
    objCliente.dComissaoVendas = objContato.dComissaoVendas
    objCliente.dDesconto = objContato.dDesconto
    objCliente.dLimiteCredito = objContato.dLimiteCredito
    objCliente.dMediaCompra = objContato.dMediaCompra
    objCliente.dSaldoAtrasados = objContato.dSaldoAtrasados
    objCliente.dSaldoDuplicatas = objContato.dSaldoDuplicatas
    objCliente.dSaldoPedidosLiberados = objContato.dSaldoPedidosLiberados
    objCliente.dSaldoTitulos = objContato.dSaldoTitulos
    objCliente.dtDataPrimeiraCompra = objContato.dtDataPrimeiraCompra
    objCliente.dtDataUltChequeDevolvido = objContato.dtDataUltChequeDevolvido
    objCliente.dtDataUltimaCompra = objContato.dtDataUltimaCompra
    objCliente.dtDataUltimoProtesto = objContato.dtDataUltimoProtesto
    objCliente.dtDataUltVisita = objContato.dtDataUltVisita
    objCliente.dValorAcumuladoCompras = objContato.dValorAcumuladoCompras
    objCliente.dValPagtosAtraso = objContato.dValPagtosAtraso
    objCliente.iAtivo = objContato.iAtivo
    objCliente.iAtivoFilial = objContato.iAtivoFilial
    objCliente.iBloqueado = objContato.iBloqueado
    objCliente.iCodCobrador = objContato.iCodCobrador
    objCliente.iCodMensagem = objContato.iCodMensagem
    objCliente.iCodPadraoCobranca = objContato.iCodPadraoCobranca
    objCliente.iCodTransportadora = objContato.iCodTransportadora
    objCliente.iCodTranspRedesp = objContato.iCodTranspRedesp
    objCliente.iCondicaoPagto = objContato.iCondicaoPagto
    objCliente.iDetPagFrete = objContato.iDetPagFrete
    objCliente.iFilialEmpresaLoja = objContato.iFilialEmpresaLoja
    objCliente.iFreqVisitas = objContato.iFreqVisitas
    objCliente.iNumChequesDevolvidos = objContato.iNumChequesDevolvidos
    objCliente.iProxCodFilial = objContato.iProxCodFilial
    objCliente.iRegiao = objContato.iRegiao
    objCliente.iTabelaPreco = objContato.iTabelaPreco
    objCliente.iTipo = objContato.iTipo
    objCliente.iTipoFrete = objContato.iTipoFrete
    objCliente.iTransferido = objContato.iTransferido
    objCliente.iVendedor = objContato.iVendedor
'    objCliente.lCodigo = objContato.dComissaoVendas
'    objCliente.lCodigoLoja = objContato.dComissaoVendas
    objCliente.lEndereco = objContato.lEndereco
    objCliente.lEnderecoCobranca = objContato.lEnderecoCobranca
    objCliente.lEnderecoEntrega = objContato.lEnderecoEntrega
    objCliente.lMaiorAtraso = objContato.lMaiorAtraso
    objCliente.lMediaAtraso = objContato.lMediaAtraso
    objCliente.lNumeroCompras = objContato.lNumeroCompras
    objCliente.lNumPagamentos = objContato.lNumPagamentos
    objCliente.lNumTitulosProtestados = objContato.lNumTitulosProtestados
    objCliente.sCgc = objContato.sCgc
    objCliente.sContaContabil = objContato.sContaContabil
    objCliente.sGuia = objContato.sGuia
    objCliente.sInscricaoEstadual = objContato.sInscricaoEstadual
    objCliente.sInscricaoMunicipal = objContato.sInscricaoMunicipal
    objCliente.sInscricaoSuframa = objContato.sInscricaoSuframa
    objCliente.sNomeReduzido = objContato.sNomeReduzido
    objCliente.sObservacao = objContato.sObservacao
    objCliente.sObservacao2 = objContato.sObservacao2
    objCliente.sRazaoSocial = objContato.sRazaoSocial
    objCliente.sRG = objContato.sRG

    Contato_Transfere_Cliente = SUCESSO

    Exit Function

Erro_Contato_Transfere_Cliente:

    Contato_Transfere_Cliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180514)

    End Select

    Exit Function

End Function

Function FiliaisContatos_Lock(ByVal lCodigo As Long) As Long
'Faz lock das Filiais Contatos relacionadas ao Contato passado por parametro para garantir a exclusao de um Contato
'Chamada por Contato_Exclui DENTRO DE TRANSAÇÃO

Dim lErro As Long
Dim lComando1 As Long
Dim iCodFilial As Integer

On Error GoTo Erro_FiliaisContatos_Lock

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then gError 180553
    
    'Seleciona as Filiais do Contato a serem lockadas
    lErro = Comando_ExecutarLockado(lComando1, "SELECT CodFilial FROM FiliaisContatos WHERE CodContato =?", iCodFilial, lCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 180554

    lErro = Comando_BuscarPrimeiro(lComando1)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180555

    If lErro = AD_SQL_SEM_DADOS Then gError 180556

    Do While lErro = AD_SQL_SUCESSO

        lErro = Comando_LockExclusive(lComando1)
        If lErro <> AD_SQL_SUCESSO Then gError 180557

        lErro = Comando_BuscarProximo(lComando1)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180558

    Loop

    Call Comando_Fechar(lComando1)

    FiliaisContatos_Lock = SUCESSO

    Exit Function

Erro_FiliaisContatos_Lock:

    FiliaisContatos_Lock = gErr

    Select Case gErr
        
        Case 180553
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 180554, 180555, 180558
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISContatoS", gErr)

        Case 180556
            Call Rotina_Erro(vbOKOnly, "ERRO_Contato_SEM_FILIAL", gErr, lCodigo)

        Case 180557
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_FILIAIScontatoS", gErr, lCodigo, iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180559)

    End Select

    Call Comando_Fechar(lComando1)

    Exit Function

End Function

Function FiliaisContatos_Le_Tabelas_Lock(ByVal objFilialContato As ClassFilialContato) As Long
'Faz o lock das tabelas relacionadas a tabela FiliaisContatos para a gravação de novo Contato
'CHAMADA EM TRANSAÇÃO por FiliaisContatos_Grava

Dim alComando(6) As Long
Dim lCodigo As Long
Dim iCodigo As Integer
Dim iTipoConta As Integer
Dim iAtivo As Integer
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_FiliaisContatos_Le_Tabelas_Lock

    'Abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180560
    Next

    If giLocalOperacao = LOCALOPERACAO_CAIXA_CENTRAL And objFilialContato.lCodContato = 0 Then
        'Seleciona o Contato relacionado a Filial para ser lockado
        lErro = Comando_ExecutarLockado(alComando(0), "SELECT Codigo FROM Contatos WHERE FilialEmpresaLoja = ? AND CodigoLoja =?", lCodigo, objFilialContato.iFilialEmpresaLoja, objFilialContato.lCodContatoLoja)
        If lErro <> AD_SQL_SUCESSO Then gError 180561
    Else
        'Seleciona o Contato relacionado a Filial para ser lockado
        lErro = Comando_ExecutarLockado(alComando(0), "SELECT Codigo FROM Contatos WHERE Codigo =?", lCodigo, objFilialContato.lCodContato)
        If lErro <> AD_SQL_SUCESSO Then gError 180562
    End If

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180571
    If lErro = AD_SQL_SEM_DADOS Then gError 180563 'Contato não existe no BD

    'Lock do Contato
    lErro = Comando_LockShared(alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 180564

    'PlanoConta
    If Len(Trim(objFilialContato.sContaContabil)) > 0 Then
    
        lErro = CF("Conta_Analitica_Lock", objFilialContato.sContaContabil, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO Then gError 180565
        
    End If

    'Cobrador
    If objFilialContato.iCodCobrador > 0 Then
    
        lErro = CF("Cobrador_Lock", alComando(2), objFilialContato.iCodCobrador)
        If lErro <> SUCESSO Then gError 180566
    
    End If
    
    'Transportadora
    If objFilialContato.iCodTransportadora > 0 Then
    
        lErro = CF("Transportadora_Lock", alComando(3), objFilialContato.iCodTransportadora)
        If lErro <> SUCESSO Then gError 180567
        
    End If

    'Transportadora de redespacho
    If objFilialContato.iCodTranspRedesp > 0 Then
    
        lErro = CF("Transportadora_Lock", alComando(3), objFilialContato.iCodTranspRedesp)
        If lErro <> SUCESSO Then gError 180568
        
    End If

    'Vendedor
    If objFilialContato.iVendedor > 0 Then
    
        lErro = CF("Vendedor_Lock", alComando(4), objFilialContato.iVendedor)
        If lErro <> SUCESSO Then gError 180569
        
    End If

    'RegioesVendas
    If objFilialContato.iRegiao > 0 Then
    
        lErro = CF("RegiaoVenda_Lock", alComando(5), objFilialContato.iRegiao)
        If lErro <> SUCESSO Then gError 180570
        
    End If

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FiliaisContatos_Le_Tabelas_Lock = SUCESSO

    Exit Function

Erro_FiliaisContatos_Le_Tabelas_Lock:

    FiliaisContatos_Le_Tabelas_Lock = gErr

    Select Case gErr

        Case 180560
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 180561, 180562, 180571
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)

        Case 180563
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_NAO_CADASTRADO", gErr, objFilialContato.lCodContato)

        Case 180564
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTATOS", gErr, objFilialContato.lCodContato)

        Case 180565 To 180570
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180572)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FiliaisContatos_Exclui(ByVal lCodContato As Long, Optional bExcluiEnd As Boolean = True) As Long
'Exclui as Filiais do Contato cujo código foi passado.
'Chamada por Contato_Exclui DENTRO DE TRANSAÇÃO.

Dim alComando(2) As Long
Dim lErro As Long
Dim lEndereco(3) As Long
Dim iIndice As Integer
Dim objFilialCliente As New ClassFilialCliente
    
On Error GoTo Erro_FiliaisContatos_Exclui

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180573
    Next
    
    'Seleciona as Filiais do Contato a serem excluídas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Endereco, EnderecoEntrega, EnderecoCobranca FROM FiliaisContatos WHERE CodContato =?", 0, lEndereco(0), lEndereco(1), lEndereco(2), lCodContato)
    If lErro <> AD_SQL_SUCESSO Then gError 180574

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180575
    If lErro = AD_SQL_SEM_DADOS Then gError 180576

    'Exclui as Filiais do Contato
    Do While lErro = AD_SQL_SUCESSO

        'Exclui a Filial
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM FiliaisContatos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 180577

        If bExcluiEnd Then

            objFilialCliente.lEndereco = lEndereco(0)
            objFilialCliente.lEnderecoEntrega = lEndereco(1)
            objFilialCliente.lEnderecoCobranca = lEndereco(2)
            
            'Exclui os Endereços da Filial
            lErro = CF("Enderecos_Exclui_FilialCliente", objFilialCliente)
            If lErro <> SUCESSO Then gError 180578
            
        End If
        
        'Traz a próxima Filial
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 180579

    Loop
    
    'Fecha os comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        lErro = Comando_Fechar(alComando(iIndice))
    Next

    FiliaisContatos_Exclui = SUCESSO

    Exit Function

Erro_FiliaisContatos_Exclui:

    FiliaisContatos_Exclui = gErr

    Select Case gErr

        Case 180573
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 180574, 180575, 180579
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCONTATOS", gErr)

        Case 180576
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTATO_SEM_FILIAL", gErr, lCodContato)

        Case 180577
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_FILIAISContatoS", gErr, lCodContato)
            
        Case 180578

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180580)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function FiliaisContatos_Automatico(ByVal lCodContato As Long, iCodigo As Integer) As Long
'retorna o número da proxima Filial Contato disponivel em iCodigo em relacao ao numero de Contato passado em lCodigo

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_FiliaisContatos_Automatico

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 180844
    
    lErro = CF("FiliaisContatos_Automatico_EmTrans", lCodContato, iCodigo)
    If lErro <> SUCESSO Then gError 180845
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 180846

    FiliaisContatos_Automatico = SUCESSO

    Exit Function

Erro_FiliaisContatos_Automatico:

    FiliaisContatos_Automatico = gErr

    Select Case gErr

    Case 180844
        Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

    Case 180845

    Case 180846
        Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
    
    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180847)

    End Select

    Call Transacao_Rollback

    Exit Function

End Function

Function FiliaisContatos_Automatico_EmTrans(ByVal lCodContato As Long, iCodigo As Integer) As Long
'retorna o número da proxima Filial Contato disponivel em iCodigo em relacao ao numero de Contato passado em lCodigo

Dim alComando(1 To 3) As Long
Dim lErro As Long
Dim lTransacao As Long
Dim iCodigo_Automatico As Integer
Dim iCodFilial As Integer
Dim iIndice As Integer

On Error GoTo Erro_FiliaisContatos_Automatico_EmTrans

    'Abre comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 180848
    Next
    
    'Seleciona o Proximo Codigo da Filial , do respectivo Contato
    lErro = Comando_ExecutarPos(alComando(1), "SELECT ProxCodFilial FROM Contatos WHERE Codigo = ?", 0, iCodigo_Automatico, lCodContato)
    If lErro <> AD_SQL_SUCESSO Then gError 180849

    'le o registro com o número da próxima Filial
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 180850

    'loca o registro que guarda o numero da Filial Contato gerado automaticamente
    lErro = Comando_LockExclusive(alComando(1))
    If lErro <> AD_SQL_SUCESSO Then gError 180851

    'verifica se este número ainda não foi utilizado
    'Pesquisa as numerações de Filiais maiores ou iguais a numeração gerada automaticamente
    lErro = Comando_Executar(alComando(2), "SELECT CodFilial FROM FiliaisContatos WHERE CodContato = ? AND CodFilial >= ? ORDER BY CodFilial", iCodFilial, lCodContato, iCodigo_Automatico)
    If lErro <> AD_SQL_SUCESSO Then gError 180852

    'le o primeiro codigo >= ao codigo automatico
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 180853

    'pesquisa o proximo número de codigo disponível
    Do While lErro <> AD_SQL_SEM_DADOS

        If iCodFilial > iCodigo_Automatico Then Exit Do

        If iCodFilial = iCodigo_Automatico Then
            iCodigo_Automatico = iCodigo_Automatico + 1
        End If

        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SEM_DADOS And lErro <> AD_SQL_SUCESSO Then gError 180854

    Loop

    'Valor de retorno da funcao
    iCodigo = iCodigo_Automatico

    iCodigo_Automatico = iCodigo_Automatico + 1

    'atualiza o numero do Contato gerado automaticamente
    lErro = Comando_ExecutarPos(alComando(3), "UPDATE Contatos SET ProxCodFilial = ?", alComando(1), iCodigo_Automatico)
    If lErro <> AD_SQL_SUCESSO Then gError 180855
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    FiliaisContatos_Automatico_EmTrans = SUCESSO

    Exit Function

Erro_FiliaisContatos_Automatico_EmTrans:

    FiliaisContatos_Automatico_EmTrans = gErr

    Select Case gErr

        Case 180848
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 180849, 180850
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)
    
        Case 180851
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CONTATOS", gErr)
    
        Case 180852, 180853, 180854
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONTATOS", gErr)
    
        Case 180855
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CONTATO", gErr, lCodContato)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 180856)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Grava(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Abre transação e chama a função para gravação de um relacionamento com cliente
'objRelacionamentoClientes recebe os dados que serão gravados no BD

Dim lErro As Long
Dim lTransacao As Long
Dim bAlteracaoValida As Boolean
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_RelacionamentoContatos_Grava

    'Indica que a possível alteração do relacionamento é válida
    bAlteracaoValida = True
    
    'Verifica se esse relacionamento é apontado por outro relacionamento como relacionamento anterior
    'para determinar se a data pode ser alterada ou não
    lErro = CF("RelacionamentoContatos_Valida_Alteracao", objRelacionamentoClientes, bAlteracaoValida)
    If lErro <> SUCESSO Then gError 181061
    
    'Se a alteração não foi considerada válida
    If Not bAlteracaoValida Then
    
        'Avisa ao usuário que a data e o cliente não serão alterados
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_RELACIONAMENTO_NAO_ALTERA_DADOS", objRelacionamentoClientes.lCodigo)
        
        'Se o usuário não concordou com a atualização => sai da função por erro
        If vbMsgRes = vbNo Then gError 181062
    
    End If
   
    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181063
    
    'Chama a função para efetuar a gravação do relacionamento
    lErro = RelacionamentoContatos_Grava_EmTrans(objRelacionamentoClientes, bAlteracaoValida)
    If lErro <> SUCESSO Then gError 181064
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181065
    
    RelacionamentoContatos_Grava = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Grava:

    RelacionamentoContatos_Grava = gErr

    Select Case gErr
   
        Case 181061, 181062
        
        Case 181063
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181064
        
        Case 181065
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181066)

    End Select
    
    'Faz o rollback da transacao
    Call Transacao_Rollback

End Function

'Incluído por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Grava_EmTrans(ByVal objRelacionamentoClientes As ClassRelacClientes, ByVal bAlteracaoValida As Boolean) As Long
'Faz o lock dos objetos relacionados, verifica a existência do relacionamento
'e insere um novo relacionamento ou altera o já existente
'objRelacionamentoClientes recebe os dados que serão gravados no BD

Dim lErro As Long
Dim alComando(1) As Long
Dim iIndice As Integer
Dim tRelacionamentoClientes As typeRelacionamentoClientes

On Error GoTo Erro_RelacionamentoContatos_Grava_EmTrans

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181067
    Next

    'Efetua os locks dos objetos relacionados
    lErro = CF("RelacionamentoContatos_Locks_Relacionados", objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 181068
    
    'Verifica se já existe relacionamento com o código em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo, Data, Cliente FROM RelacionamentoContatos WHERE Codigo=? AND FilialEmpresa=?", 0, tRelacionamentoClientes.lCodigo, tRelacionamentoClientes.dtData, tRelacionamentoClientes.lCliente, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 181069
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181070
    
    'Se não encontrou relacionamento com o código em questão
    If lErro = AD_SQL_SEM_DADOS Then
    
        'Insere um novo relacionamento no BD
        With objRelacionamentoClientes
            lErro = Comando_Executar(alComando(1), "INSERT INTO RelacionamentoContatos (Codigo,FilialEmpresa,Origem,Data,Hora,Tipo,Cliente,FilialCliente,Contato,Atendente,RelacionamentoAnt,Assunto1,Assunto2,Status) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", .lCodigo, .iFilialEmpresa, .iOrigem, .dtData, CDbl(.dtHora), .lTipo, .lCliente, .iFilialCliente, .iContato, .iAtendente, .lRelacionamentoAnt, .sAssunto1, .sAssunto2, .iStatus)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 181071
    
    'Senão, ou seja, se encontrou o relacionamento
    Else
    
        'Se a alteração não é válida
        If Not bAlteracaoValida Then
        
            'joga no obj a data que já está no BD, ou seja, a atualização não fará diferença
            objRelacionamentoClientes.dtData = tRelacionamentoClientes.dtData
            
            'joga no obj o cliente que já está no BD, ou seja, a atualização não fará diferença
            objRelacionamentoClientes.lCliente = tRelacionamentoClientes.lCliente
        
        End If
        
        'Altera os dados do relacionamento
        With objRelacionamentoClientes
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE RelacionamentoContatos SET Origem=?, Data=?, Hora=?, Tipo=?, Cliente=?, FilialCliente=?, Contato=?, Atendente=?, RelacionamentoAnt=?, Assunto1=?, Assunto2=?, Status=?", alComando(0), .iOrigem, .dtData, CDbl(.dtHora), .lTipo, .lCliente, .iFilialCliente, .iContato, .iAtendente, .lRelacionamentoAnt, .sAssunto1, .sAssunto2, .iStatus)
        End With
        If lErro <> AD_SQL_SUCESSO Then gError 181072
    
    End If
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacionamentoContatos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Grava_EmTrans:

    RelacionamentoContatos_Grava_EmTrans = gErr

    Select Case gErr

        Case 181067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181068

        Case 181069, 181070
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 181071
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELACIONAMENTOCLIENTES", gErr, objRelacionamentoClientes.lCodigo)
            
        Case 181072
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_RELACIONAMENTOCLIENTES", gErr, objRelacionamentoClientes.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181073)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Exclui(ByVal objRelacionamentoClientes As ClassRelacClientes, Optional ByVal bValidaOutrosRelac As Boolean = True) As Long
'Abre transação e chama a função responsável pela exclusão do relacionamento
'objRelacionamentoClientes RECEBE(Input) os dados do relacionamento a ser excluído

Dim lErro As Long
Dim lTransacao As Long
Dim lComando As Long

On Error GoTo Erro_RelacionamentoContatos_Exclui

    'Executa a abertura da Transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181074
    
    'Exclui o relacionamento em questão
    lErro = RelacionamentoContatos_Exclui_EmTrans(objRelacionamentoClientes, bValidaOutrosRelac)
    If lErro <> SUCESSO Then gError 181075
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181076

    RelacionamentoContatos_Exclui = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Exclui:

    RelacionamentoContatos_Exclui = gErr

    Select Case gErr

        Case 181074
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181075

        Case 181076
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181077)

    End Select

    'Faz o rollback da transacao
    Call Transacao_Rollback

End Function

'Incluída por Luiz Nogueira em 27/10/03
Public Function RelacionamentoContatos_Exclui_EmTrans(ByVal objRelacionamentoClientes As ClassRelacClientes, Optional ByVal bValidaOutrosRelac As Boolean = True) As Long
'Exclui o relacionamento com código e filial empresa passados como parâmetro
'Um relacionamento que seja indicado como antecessor de outro, não poderá ser excluído
'objRelacionamentoClientes RECEBE(Input)

Dim lErro As Long
Dim alComando(2) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_RelacionamentoContatos_Exclui_EmTrans

    'Executa a abertura do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181078
    Next
    
    'Lê o relacionamento em questão
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM RelacionamentoContatos WHERE Codigo=? AND FilialEmpresa=?", 0, lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 181079
    
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181080
    
    'Se não encontrou o relacionamento => erro
    If lErro = AD_SQL_SEM_DADOS Then gError 181081
    
    If bValidaOutrosRelac Then
    
        'Verifica no bd se existe algum relacionamento nessa filial que usa o relacionamento
        'a ser excluído como relacionamento anterior
        lErro = Comando_Executar(alComando(1), "SELECT Codigo FROM RelacionamentoContatos WHERE RelacionamentoAnt=? AND FilialEmpresa=?", lCodigo, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 181082
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181083
        
        'Se encontrou => erro, pois se o relacionamento é apontado como anterior
        'a outro relacionamento, ele não pode ser excluído
        If lErro = AD_SQL_SUCESSO Then gError 181084
        
    End If
    
    'Exclui o relacionamento em questão
    lErro = Comando_ExecutarPos(alComando(2), "DELETE FROM RelacionamentoContatos", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 181085
    
    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacionamentoContatos_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_RelacionamentoContatos_Exclui_EmTrans:

    RelacionamentoContatos_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181078
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181079, 181080, 181082, 181083
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELACIONAMENTOCLIENTES", gErr)
        
        Case 181081
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo)
        
        Case 181084
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_ANTERIOR_OUTRO", gErr, objRelacionamentoClientes.lCodigo, lCodigo)
        
        Case 181085
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RELACIONAMENTOCLIENTES", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181086)

    End Select

    'Executa o fechamento do Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function ClienteFContatos_Exclui(ByVal objClienteFContatos As ClassClienteContatos) As Long
'Abre transação e chama a função para exclusão de contatos de um cliente
'objClienteFContatos recebe o código do cliente e da filial que terão seus contatos excluídos

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ClienteFContatos_Exclui

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181138
    
    'Chama a função para efetuar a exclusão do relacionamento
    lErro = ClienteFContatos_Exclui_EmTrans(objClienteFContatos)
    If lErro <> SUCESSO Then gError 181139
    
    'Verifica se não foi eliminado algum contato que esteja sendo usada em alguma outra tabela relacionada
    lErro = CF("ClienteFContatos_Valida_Exclusao", objClienteFContatos)
    If lErro <> SUCESSO Then gError 181140
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181141

    ClienteFContatos_Exclui = SUCESSO

    Exit Function

Erro_ClienteFContatos_Exclui:

    ClienteFContatos_Exclui = gErr

    Select Case gErr

        Case 181138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181139, 181140
        
        Case 181141
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181142)

    End Select
    
    'Faz o rollback da transação
    Call Transacao_Rollback

End Function

Public Function ClienteFContatos_Exclui_EmTrans(ByVal objClienteFContatos As ClassClienteContatos) As Long
'Exclui os contatos do cliente/filial passados como parâmetros
'objClienteFContatos recebe o código do cliente e da filial que terão seus contatos excluídos

Dim lErro As Long
Dim alComando(1) As Long
Dim iIndice As Integer
Dim lCodigo As Long

On Error GoTo Erro_ClienteFContatos_Exclui_EmTrans

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181143
    Next
    
    'Lê os contatos que serão excluídos
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM ClienteFContatos WHERE Cliente=? AND FilialCliente=?", 0, lCodigo, objClienteFContatos.lCliente, objClienteFContatos.iFilialCliente)
    If lErro <> AD_SQL_SUCESSO Then gError 181144
    
    'Posiciona no primeiro contato encontrado
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181145
    
    'Enquando houverem contatos para exclusão
    Do While lErro = AD_SQL_SUCESSO
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ClienteFContatos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 181146
        
        'Posiciona no próximo contato encontrado
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181147
    
    Loop

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ClienteFContatos_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_ClienteFContatos_Exclui_EmTrans:

    ClienteFContatos_Exclui_EmTrans = gErr

    Select Case gErr

        Case 181143
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181144, 181145, 181147
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTECONTATOS", gErr)
        
        Case 181146
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTECONTATOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181148)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function ClienteFContatos_Exclui_Cliente_EmTrans(ByVal lCliente As Long, Optional ByVal iFilialCliente As Integer) As Long
'Exclui todos os contatos para o cliente passado como parâmetro
'ou apenas os contatos para a filial do cliente passada como parâmetro
'lCliente recebe o código do cliente que deve ter seus contatos excluídos
'iFilialCliente indica se deve excluir os contatos de todas as filiais ou de uma filial específica
'Se for passado um valor em iFilialCliente, ou seja, iFilialCliente<>0, significa que devem ser
'excluídos apenas os contatos da filial passada

Dim lErro As Long
Dim alComando(1) As Long
Dim iIndice As Integer
Dim lCodContato As Long

On Error GoTo Erro_ClienteFContatos_Exclui_Cliente_EmTrans

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181150
    Next
    
    'Se foi passada uma filial específica do cliente
    If iFilialCliente <> 0 Then
    
        'Lê os contatos para a filial do cliente passada como parâmetro
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM ClienteFContatos WHERE Cliente=? AND FilialCliente=?", 0, lCodContato, lCliente, iFilialCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 181151
    
    'Senão
    Else
    
        'Lê os contatos de todas as filiais do cliente passado como parâmetro
        lErro = Comando_ExecutarPos(alComando(0), "SELECT Codigo FROM ClienteFContatos WHERE Cliente=?", 0, lCodContato, lCliente)
        If lErro <> AD_SQL_SUCESSO Then gError 181152
    
    End If
    
    'Posiciona no primeiro contato encontrado
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181153
    
    'Enquanto houver contatos para excluir
    Do While lErro = AD_SQL_SUCESSO
    
        'Exclui o contato
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM ClienteFContatos", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 181154
    
        'Busca o próximo contato
        lErro = Comando_BuscarPrimeiro(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181155
    
    Loop

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ClienteFContatos_Exclui_Cliente_EmTrans = SUCESSO

    Exit Function

Erro_ClienteFContatos_Exclui_Cliente_EmTrans:

    ClienteFContatos_Exclui_Cliente_EmTrans = gErr

    Select Case gErr

        Case 181150
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181151, 181152, 181153, 181155
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTECONTATOS", gErr)
        
        Case 181154
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CLIENTECONTATOS", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181156)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function

Public Function ClienteFContatos_Grava(ByVal colClienteFContatos As Collection) As Long
'Abre transação e chama a função para gravação de contatos
'colClienteFContatos recebe os contatos que serão gravados

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_ClienteFContatos_Grava

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181157
    
    'Efetua
    'Chama a função para efetuar a gravação do relacionamento
    lErro = ClienteFContatos_Grava_EmTrans(colClienteFContatos)
    If lErro <> SUCESSO Then gError 181158
    
    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181159

    ClienteFContatos_Grava = SUCESSO

    Exit Function

Erro_ClienteFContatos_Grava:

    ClienteFContatos_Grava = gErr

    Select Case gErr

        Case 181157
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181158
        
        Case 181159
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181160)

    End Select

    'Faz o rollback da transação
    Call Transacao_Rollback

End Function

Public Function ClienteFContatos_Grava_EmTrans(ByVal colClienteFContatos As Collection) As Long
'Faz o lock dos objetos relacionados, exclui os contatos gravados atualmente,
'insere os novos contatos e verifica se não ficou nenhum contato inconsistente no BD
'colClienteFContatos recebe os contatos a serem gravados

Dim lErro As Long
Dim lComando As Long
Dim iIndice As Integer
Dim objClienteFContatos As New ClassClienteContatos
Dim objClienteFContatosAux As New ClassClienteContatos

On Error GoTo Erro_ClienteFContatos_Grava_EmTrans

    'Executa a abertura dos Comandos
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 181161
    
    'Verifica se Cliente e Filial estão cadastrados e faz o lock deles
    lErro = CF("FilialContato_Lock", colClienteFContatos(1).lCliente, colClienteFContatos(1).iFilialCliente)
    If lErro <> SUCESSO And lErro <> 181036 Then gError 181162
    
    'Não encontrou a Filial Cliente --> erro
    If lErro = 181036 Then gError 181163
    
    'Guarda código do cliente e da filial que terão os contatos excluídos
    'A exclusão é feita para limpar o BD e posteriormente ser feita a gravação dos
    'novos contatos
    objClienteFContatos.lCliente = colClienteFContatos(1).lCliente
    objClienteFContatos.iFilialCliente = colClienteFContatos(1).iFilialCliente
    
    'Exclui os contatos existentes para o Cliente/Filial
    lErro = ClienteFContatos_Exclui_EmTrans(objClienteFContatos)
    If lErro <> SUCESSO Then gError 181164
    
    'Para cada contato a ser gravado
    For Each objClienteFContatosAux In colClienteFContatos
    
        'Insere o contato na tabela
        lErro = Comando_Executar(lComando, "INSERT INTO ClienteFContatos (Cliente, FilialCliente, Codigo, Contato, Setor, Cargo, Telefone, Fax, Email, Padrao) VALUES (?,?,?,?,?,?,?,?,?,?)", objClienteFContatosAux.lCliente, objClienteFContatosAux.iFilialCliente, objClienteFContatosAux.iCodigo, objClienteFContatosAux.sContato, objClienteFContatosAux.sSetor, objClienteFContatosAux.sCargo, objClienteFContatosAux.sTelefone, objClienteFContatosAux.sFax, objClienteFContatosAux.sEmail, objClienteFContatosAux.iPadrao)
        If lErro <> AD_SQL_SUCESSO Then gError 181165
    
    Next
    
    'Verifica se não foi eliminado algum contato que esteja sendo usada em alguma outra tabela relacionada
    lErro = CF("ClienteFContatos_Valida_Exclusao", objClienteFContatos)
    If lErro <> SUCESSO Then gError 181166

    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

    ClienteFContatos_Grava_EmTrans = SUCESSO

    Exit Function

Erro_ClienteFContatos_Grava_EmTrans:

    ClienteFContatos_Grava_EmTrans = gErr

    Select Case gErr

        Case 181161
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 181162, 181164, 181166

        Case 181163
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objClienteFContatos.iFilialCliente, objClienteFContatos.lCliente)
        
        Case 181165
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_CLIENTECONTATOS", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 149978)

    End Select

    'Executa o fechamento dos Comandos
    Call Comando_Fechar(lComando)

End Function

Public Function Contato_Transforma_Cliente(ByVal objContato As ClassContatos, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim lTransacao As Long
Dim iIndice As Integer
Dim objFilialCliente As ClassFilialCliente
Dim objFilialContato As ClassFilialContato
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim colEnderecos As New Collection
Dim lCodigo As Long
Dim objClienteContatos As ClassClienteContatos
Dim colClienteContatos As Collection
Dim objRelacCli As ClassRelacClientes
Dim colRelacCli As New Collection
Dim lCodigoRel As Long
Dim colCatItens As Collection
Dim iAux As Integer

On Error GoTo Erro_Contato_Transforma_Cliente:

    'Abre transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 181142
    
    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 181198
    Next
    
    'Le o Contato
    lErro = CF("Contato_Le", objContato)
    If lErro <> SUCESSO And lErro <> 180548 Then gError 181157
    
    'Cria o objCliente cm base no objContato
    lErro = Contato_Transfere_Cliente(objContato, objCliente)
    If lErro <> SUCESSO Then gError 181143
    
    'Obtém o próximo código do cliente
    lErro = CF("Config_ObterAutomatico_EmTrans", "CPRConfig", "NUM_PROX_CLIENTE", "Clientes", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 181144
    
    'Acerta o código do Cliente
    objCliente.lCodigo = lCodigo
    
    'Grava o Cliente
    lErro = CF("Cliente_Grava_EmTrans", objCliente, colEnderecos)
    If lErro <> SUCESSO Then gError 181145
    
    'Le as filiais do Contato
    lErro = CF("FiliaisContatos_Le_Contato", objContato, colCodigoNome)
    If lErro <> SUCESSO Then gError 181146
    
    'Para cada filial
    For Each objCodigoNome In colCodigoNome
    
        Set objFilialContato = New ClassFilialContato
        Set objFilialCliente = New ClassFilialCliente
        Set objClienteContatos = New ClassClienteContatos
        Set colClienteContatos = New Collection
        Set objRelacCli = New ClassRelacClientes
        Set colCatItens = New Collection
        
        objFilialContato.iCodFilial = objCodigoNome.iCodigo
        objFilialContato.lCodContato = objContato.lCodigo
        
        'Le todos os dados da Filial
        lErro = CF("FilialContato_Le", objFilialContato)
        If lErro <> SUCESSO And lErro <> 180801 Then gError 181147
        
        lErro = CF("FilialContato_Le_Categorias", objFilialContato, colCatItens)
        If lErro <> SUCESSO Then gError 181166
        
        Set objFilialContato.colCategoriaItem = colCatItens
        
        'Cria o objFilialCliente com base no objFilialContato
        lErro = FilialContato_Transfere_FilialCliente(objFilialContato, objFilialCliente)
        If lErro <> SUCESSO Then gError 181148
        
        'Muda a referência para o Cliente
        objFilialCliente.lCodCliente = objCliente.lCodigo
        
        'Grava a filial passando colEnderecos vazio para não alterar a referencia aos endereços já existentes
        lErro = CF("FiliaisClientes_Grava_EmTrans", objFilialCliente, colEnderecos)
        If lErro <> SUCESSO Then gError 181149
        
        'Pesquisa a Filial no BD
        lErro = Comando_ExecutarPos(alComando(1), "SELECT 1 FROM FiliaisClientes WHERE CodCliente = ? AND CodFilial = ?", 0, iAux, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> AD_SQL_SUCESSO Then gError 181205
    
        'Tenta ler
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 181206
        
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE FiliaisClientes SET Endereco = ?,EnderecoEntrega = ?,EnderecoCobranca = ?", alComando(1), _
            objFilialCliente.lEndereco, objFilialCliente.lEnderecoEntrega, objFilialCliente.lEnderecoCobranca)
        If lErro <> AD_SQL_SUCESSO Then gError 181207
        
        objClienteContatos.iFilialCliente = objCodigoNome.iCodigo
        objClienteContatos.lCliente = objContato.lCodigo
        
        'Le todos os contatos relacionados a clientes futuros
        lErro = CF("ClienteFContatos_Le_Todos", colClienteContatos, objClienteContatos)
        If lErro <> SUCESSO And lErro <> 181086 Then gError 181150
        
        'Para cada contato
        For Each objClienteContatos In colClienteContatos
            'Acerta a referência para o Cliente
            objClienteContatos.lCliente = objCliente.lCodigo
        Next
        
        If colClienteContatos.Count > 0 Then
            'Grava os Contatos para o Cliente
            lErro = CF("ClienteContatos_Grava_EmTrans", colClienteContatos)
            If lErro <> SUCESSO Then gError 181151
        End If
            
    Next
   
    'Le todos os relacionamentos com os Clientes
    lErro = CF("RelacionamentoContatos_Le_Contato", objContato, colRelacCli)
    If lErro <> SUCESSO Then gError 181152
    
    'Para cada relacionamento
    For Each objRelacCli In colRelacCli

        lErro = RelacionamentoContatos_Exclui_EmTrans(objRelacCli, False)
        If lErro <> SUCESSO Then gError 181158

        'Muda a referencia para o Cliente
        objRelacCli.lCliente = objCliente.lCodigo
        
        lErro = CF("Config_ObterAutomatico_EmTrans", "CRFATConfig", "NUM_PROX_RELACIONAMENTOCONTATOS", "RelacionamentoClientes", "Codigo", lCodigoRel)
        If lErro <> SUCESSO Then gError 181165
        
        objRelacCli.lCodigo = lCodigoRel
    
        'Grava o Relacionamento
        lErro = RelacionamentoClientes_Grava_EmTrans(objRelacCli, False)
        If lErro <> SUCESSO Then gError 181153
               
    Next
    
    lErro = Contato_Exclui_EmTrans(objContato, False)
    If lErro <> SUCESSO Then gError 181154

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fecha a transação
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 181155

    Contato_Transforma_Cliente = SUCESSO

    Exit Function

Erro_Contato_Transforma_Cliente:

    Contato_Transforma_Cliente = gErr
    
    Select Case gErr
    
        Case 181142
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
            
        Case 181143 To 181154, 181157, 181158
        
        Case 181155
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
            
        Case 181198
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 181205, 181206
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES", gErr, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        
        Case 181207
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_FILIAISCLIENTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 181156)

    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
   
    'Faz o rollback da transação
    Call Transacao_Rollback
    
End Function

Private Function FilialContato_Grava_CategoriaNaColecao(ByVal colCategoriaItem As Collection, ByVal sCategoria As String) As Long
'Retorna o Indice da Categoria na coleção ou Zero se não achá-la (na coleção)

Dim iIndice As Integer
Dim objFilailCliCategoria As New ClassFilialConCategoria

    'Pesquisa a Categoria na coleção
    For iIndice = 1 To colCategoriaItem.Count

        Set objFilailCliCategoria = colCategoriaItem.Item(iIndice)

        'Se achou a categoria na coleção
        If objFilailCliCategoria.sCategoria = sCategoria Then

            FilialContato_Grava_CategoriaNaColecao = iIndice

            Exit Function

        End If

    Next

    FilialContato_Grava_CategoriaNaColecao = 0

End Function
'#####################################################################

'#####################################################################
'Inserido por Wagner 11/09/2006
Public Function RelacionamentoClientes_Grava_Atendente(ByVal objRelacionamentoClientes As ClassRelacClientes, ByVal sCodUsu As String) As Long

Dim lErro As Long
Dim alComando(0 To 2) As Long
Dim iIndice As Integer
Dim iCodigo As Integer
Dim objAtendente As New ClassAtendentes
Dim iAtivo As Integer

On Error GoTo Erro_RelacionamentoClientes_Grava_Atendente

    'Executa a abertura dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 182287
    Next

    'Seleciona o Atendente com o código de usuário passado
    lErro = Comando_Executar(alComando(0), "SELECT Codigo FROM Atendentes WHERE CodUsuario = ? AND FilialEmpresa = ?", iCodigo, sCodUsu, objRelacionamentoClientes.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 182288

    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182289

    'Se não encontrou o atendente => Cadastra
    If lErro <> AD_SQL_SUCESSO Then
    
        'Seleciona o usuario vinculado ao Atendente
        lErro = Comando_ExecutarLockado(alComando(1), "SELECT Ativo FROM Usuario WHERE CodUsuario=?", iAtivo, sCodUsu)
        If lErro <> AD_SQL_SUCESSO Then gError 182290
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182291
        
        'Se não encontrou => erro
        If lErro = AD_SQL_SEM_DADOS Then gError 182292
    
        'Efetua lock do usuário
        lErro = Comando_LockShared(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 182293
        
        'Obtém o próximo código de relacionamento para giFilialEmpresa
        lErro = CF("Config_Obter_Inteiro_Automatico_EmTrans", "CRFATConfig", "NUM_PROX_ATENDENTE", "Atendentes", "Codigo", iCodigo)
        If lErro <> SUCESSO Then gError 182294
        
        objAtendente.iCodigo = iCodigo
        objAtendente.iFilialEmpresa = objRelacionamentoClientes.iFilialEmpresa
        objAtendente.sCodUsuario = sCodUsu
    
        'Insere o atendente no BD
        lErro = Comando_Executar(alComando(2), "INSERT INTO Atendentes (Codigo, CodUsuario, FilialEmpresa) VALUES (?,?,?)", objAtendente.iCodigo, objAtendente.sCodUsuario, objAtendente.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 182295
    
    End If
        
    objRelacionamentoClientes.iAtendente = iCodigo
        
    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacionamentoClientes_Grava_Atendente = SUCESSO

    Exit Function

Erro_RelacionamentoClientes_Grava_Atendente:

    RelacionamentoClientes_Grava_Atendente = gErr

    Select Case gErr

        Case 182287
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 182288, 182289
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ATENDENTES", gErr)
        
        Case 182290, 182291
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_USUARIO", gErr)
        
        Case 182292
            Call Rotina_Erro(vbOKOnly, "ERRO_USUARIO_NAO_ENCONTRADO", gErr, sCodUsu)
            
        Case 182293
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_USUARIO", gErr, sCodUsu)
        
        Case 182294
        
        Case 182295
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_ATENDENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182296)

    End Select

    'Executa o fechamento dos Comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

End Function
'#################################################################

Function ChequePre_Grava_Trans(ByVal objChequePre As ClassChequePre, ByVal colInfoParcRec As Collection) As Long
'Rotina que inclui ou altera um Cheque Pre e as associacoes com parcelas de colInfoParcRec

Dim lErro As Long, alComando(8) As Long
Dim iIndice As Integer, bEncontrou As Boolean
Dim objInfoParcRec As ClassInfoParcRec, objInfoParcRec2 As ClassInfoParcRec
Dim tInfoParcRec As typeInfoParcRec, tChequePre As typeChequePre
Dim colInfoParcRecTemp As New Collection
Dim lNumIntCheque As Long
Dim lNumIntBaixa As Long, dtData As Date

On Error GoTo Erro_ChequePre_Grava_Trans

    For iIndice = LBound(alComando) To UBound(alComando)

        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 17700

    Next

    tChequePre.sAgencia = String(STRING_AGENCIA, 0)
    tChequePre.sContaCorrente = String(STRING_CONTA_CORRENTE, 0)

    'Pesquisa no BD o cheque pre passado como parametro
    lErro = Comando_ExecutarPos(alComando(1), "SELECT NumIntCheque, FilialEmpresa, Cliente, Filial, Banco, Agencia, ContaCorrente, Numero, DataDeposito, Valor, NumBordero FROM ChequePreN WHERE Banco = ? AND Agencia = ? AND ContaCorrente = ? AND Numero = ? AND Status <> 5", 0, _
        tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa, tChequePre.lCliente, tChequePre.iFilial, tChequePre.iBanco, tChequePre.sAgencia, tChequePre.sContaCorrente, tChequePre.lNumero, tChequePre.dtDataDeposito, tChequePre.dValor, tChequePre.lNumBordero, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17747

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 17748
    
    'Se nao encontrou --> insercao
    If lErro = AD_SQL_SEM_DADOS Then

        'insere o ChequeParcelaPre no BD
        lErro = ChequePre_Inserir(objChequePre)
        If lErro <> SUCESSO Then gError 17713
       
        'Rotina encarregada de gravar o cheque pre na tabela ChequeParcelaRec
        lErro = ChequeParcelaRec_Inserir(objChequePre, colInfoParcRec)
        If lErro <> SUCESSO Then gError 91308

        'prepara e insere o movto de cta corrente correspondente ao cheque
        lErro = CF("ChequePre_MovCCI_Grava", objChequePre)
        If lErro <> SUCESSO Then gError 126155

    Else
    
        'bloqueia o cheque
        lErro = Comando_LockExclusive(alComando(1))
        If lErro <> AD_SQL_SUCESSO Then gError 17751
        
        'verifica se o cheque está vinculado a um bordero
        lErro = Comando_Executar(alComando(6), "SELECT NumIntCheque From ChequesPre_Disp WHERE NumIntCheque = ? AND FilialEmpresa = ?", lNumIntCheque, tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 109127
        
        'busca o primeiro
        lErro = Comando_BuscarPrimeiro(alComando(6))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109128
        
        'se não encontrou->erro. Cheque vinculado a bordero não pode ser alterado ou excluído
        If lErro = AD_SQL_SEM_DADOS Then gError 109129
        'Fim Alterado por Rafael Menezes em 11/10/2002
        
        'verifica se o cheque está vinculado a uma baixa a pagar
        lErro = Comando_Executar(alComando(7), "SELECT NumIntBaixa, Data From BaixasPag WHERE NumIntCheque = ? AND FilialEmpresaCheque = ? AND Status <> 5 ", lNumIntBaixa, dtData, tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 109127
        
        'busca o primeiro
        lErro = Comando_BuscarPrimeiro(alComando(7))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109128
        
        'se encontrou->erro. Cheque vinculado a BaixasPag
        If lErro <> AD_SQL_SEM_DADOS Then gError 209263
        
        'verifica se o cheque está vinculado a uma baixa a pagar
        lErro = Comando_Executar(alComando(8), "SELECT NumIntBaixa, Data From BaixasPag WHERE NumIntCheque = ? AND FilialEmpresaCheque = ? AND Status <> 5 ", lNumIntBaixa, dtData, tChequePre.lNumIntCheque, tChequePre.iFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then gError 109127
        
        'busca o primeiro
        lErro = Comando_BuscarPrimeiro(alComando(8))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 109128
        
        'se encontrou->erro. Cheque vinculado a BaixasRec
        If lErro <> AD_SQL_SEM_DADOS Then gError 209264
        
        If (tChequePre.lCliente <> objChequePre.lCliente) Or (tChequePre.iFilial <> objChequePre.iFilial) Then gError 17781

        If tChequePre.iFilialEmpresa <> objChequePre.iFilialEmpresa Then gError 81684

        objChequePre.lNumIntCheque = tChequePre.lNumIntCheque
    
        lErro = Comando_ExecutarPos(alComando(5), "UPDATE ChequePreN SET DataDeposito = ?, Valor = ?, DataEmissao = ?", alComando(1), objChequePre.dtDataDeposito, objChequePre.dValor, objChequePre.dtDataEmissao)
        If lErro <> AD_SQL_SUCESSO Then gError 81683
        
        'cria uma copia de trabalho da colecao de parcelas
        For Each objInfoParcRec In colInfoParcRec
            Call colInfoParcRecTemp.Add(objInfoParcRec)
        Next
        
        'selecionar os dados gravados anteriormente das associacoes do cheque com parcelas
        lErro = Comando_ExecutarPos(alComando(2), "SELECT NumIntParcela, ValorRecebto, Juros, Multa, Desconto FROM ChequeParcelaRec WHERE NumIntCheque = ?", 0, tInfoParcRec.lNumIntParc, tInfoParcRec.dValorRecebto, tInfoParcRec.dValorJuros, tInfoParcRec.dValorMulta, tInfoParcRec.dValorDesconto, objChequePre.lNumIntCheque)
        If lErro <> AD_SQL_SUCESSO Then gError 17751
        
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81685
        
        Do While lErro <> AD_SQL_SEM_DADOS
        
            'verifica se a parcela continuou vinculada ao cheque
            bEncontrou = False
            
            For Each objInfoParcRec In colInfoParcRec
            
                If objInfoParcRec.lNumIntParc = tInfoParcRec.lNumIntParc Then
                
                    'a parcela permaneceu vinculada ao cheque
                    bEncontrou = True
                    
                    'se algum atributo da associacao foi alterado entao alterar o registro de chequeparcelarec
                    If (objInfoParcRec.dValorReceber <> tInfoParcRec.dValorRecebto) Or (objInfoParcRec.dValorJuros <> tInfoParcRec.dValorJuros) Or (objInfoParcRec.dValorMulta <> tInfoParcRec.dValorMulta) Or (objInfoParcRec.dValorDesconto <> tInfoParcRec.dValorDesconto) Then
                    
                        lErro = Comando_ExecutarPos(alComando(3), "UPDATE ChequeParcelaRec SET ValorRecebto = ?, Juros = ?, Multa = ?, Desconto = ?", alComando(2), objInfoParcRec.dValorReceber, objInfoParcRec.dValorJuros, objInfoParcRec.dValorMulta, objInfoParcRec.dValorDesconto)
                        If lErro <> AD_SQL_SUCESSO Then gError 81686
                        
                    End If
                    
                    'exclui o objeto da colecao temporaria
                    iIndice = 0
                    For Each objInfoParcRec2 In colInfoParcRecTemp
                        iIndice = iIndice + 1
                        If objInfoParcRec2.lNumIntParc = tInfoParcRec.lNumIntParc Then
                            Call colInfoParcRecTemp.Remove(iIndice)
                        End If
                    Next
                    
                    Exit For
                    
                End If
                                    
            Next
        
            'se a parcela nao está mais vinculada ao cheque
            If bEncontrou = False Then
                
                'excluir a associacao
                lErro = Comando_ExecutarPos(alComando(4), "DELETE FROM ChequeParcelaRec", alComando(2))
                If lErro <> AD_SQL_SUCESSO Then gError 81687
            
            End If
                
            lErro = Comando_BuscarProximo(alComando(2))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 81688
        
        Loop
    
        'inclui em ChequeParcelaRec os dados das parcelas que passaram a estar vinculadas ao cheque
        'que sao as que sobraram em colInfoParcRecTemp
        lErro = ChequeParcelaRec_Inserir(objChequePre, colInfoParcRecTemp)
        If lErro <> SUCESSO Then gError 81689

        'exclui o movto de cta corrente correspondente ao cheque se existir
        lErro = CF("ChequePre_MovCCI_Exclui", objChequePre)
        If lErro <> SUCESSO Then gError 126161

        'prepara e insere o movto de cta corrente correspondente ao cheque
        lErro = CF("ChequePre_MovCCI_Grava", objChequePre)
        If lErro <> SUCESSO Then gError 126162


    End If
    
    'verificar se há alguma parcela vinculada ao cheque que tem mais valor vinculado (com este cheque ou outro) do que seu saldo aberto
    '???
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ChequePre_Grava_Trans = SUCESSO

    Exit Function

Erro_ChequePre_Grava_Trans:

    ChequePre_Grava_Trans = gErr

    Select Case gErr
        
        Case 17747, 17748
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEPRE1", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

'        Case 17750
'            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17751, 81685, 81688
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CHEQUESPRE", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)

        Case 17781
            Call Rotina_Erro(vbOKOnly, "ERRO_CLINTE_FILIAL_NAO_CONFEREM", gErr, objChequePre.lCliente, objChequePre.iFilial)
        
        Case 17700
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 17713, 91308, 81689, 126155, 126161, 126162

        Case 81683
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CHEQUESPRE2", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
        
        Case 81684
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUESPRE_OUTRA_FILIALEMP", gErr)
        
        Case 81686
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CHEQUEPARCREC", gErr)
        
        Case 81687
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_CHEQUEPARCELAREC", gErr)
            
        Case 109127, 109128
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CHEQUEBORDERO", gErr)
        
        Case 109129
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMBORDERO_CHEQUEPRE_DEPOSITADO", gErr, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero)
        
        Case 209263
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_UTILIZADO_BAIXAPAG", gErr, lNumIntBaixa, dtData)
        
        Case 209264
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_UTILIZADO_BAIXAREC", gErr, lNumIntBaixa, dtData)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183029)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function Vendedor_Grava_Customizacao(ByVal objVendedor As ClassVendedor) As Long
'customizacoes da rotina de gravacao de vendedor

    Vendedor_Grava_Customizacao = SUCESSO

End Function

Function Vendedor_Exclui_Customizado(ByVal objVendedor As ClassVendedor) As Long
'customizacoes da rotina de exclusao de vendedor

    Vendedor_Exclui_Customizado = SUCESSO

End Function

Public Function TituloReceber_Gera_Customizado(ByVal objNFiscal As ClassNFiscal, ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelas As colParcelaReceber) As Long
    TituloReceber_Gera_Customizado = SUCESSO
End Function

Public Function TituloReceber_Gera_Customizado2(ByVal objNFiscal As ClassNFiscal) As Long
    TituloReceber_Gera_Customizado2 = SUCESSO
End Function

Function TabelaPreco_Valida_Venda(ByVal objNF As ClassNFiscal) As Long
    TabelaPreco_Valida_Venda = SUCESSO
End Function

Function TabelaPreco_Valida_Compra(ByVal objNF As ClassNFiscal, ByVal objTipoDocInfo As ClassTipoDocInfo) As Long
'atualiza (ou não) tabelas de preço em função do registro da compra de um material

Dim lErro As Long
Dim alComando(1 To 5) As Long
Dim iIndice As Integer, bAumentarPreco As Double
Dim objItemNF As ClassItemNF
Dim iCodTabela As Integer
Dim dCoef As Double, dQtdeCompra As Double
Dim dPreco As Double, dNovoPreco As Double
Dim dtDataVigencia As Date, dFator As Double, sFiltro As String
Dim bContinuaLoop As Boolean, objProduto As New ClassProduto

On Error GoTo Erro_TabelaPreco_Valida_Compra

    If objTipoDocInfo.iCompras = TIPODOCINFO_COMPRAS Then

        For iIndice = LBound(alComando) To UBound(alComando)
            alComando(iIndice) = Comando_Abrir()
            If alComando(iIndice) = 0 Then gError 200247
        Next
    
'        lErro = CF("FilialEmpresa_Le_Filtro", sFiltro)
'        If lErro <> SUCESSO Then gError 201036
        
        'Le as tabelas de preço que devem ser reajustadas
        lErro = Comando_Executar(alComando(1), "SELECT Codigo, VlrCompCoef FROM TabelasDePreco WHERE AjustaTabVlrCompCoefMaior = ?", iCodTabela, dCoef, MARCADO)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200248
    
        lErro = Comando_BuscarPrimeiro(alComando(1))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200249
        
        Do While lErro = AD_SQL_SUCESSO
        
            For Each objItemNF In objNF.ColItensNF
        
                objProduto.sCodigo = objItemNF.sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 201030
                If lErro <> SUCESSO Then gError 201035
                
                dNovoPreco = objItemNF.dValorTotal
                If objItemNF.objTributacaoItemNF.dIPICredito = 0 Then
                    dNovoPreco = dNovoPreco + objItemNF.objTributacaoItemNF.dIPIValor
                End If
                dQtdeCompra = objItemNF.dQuantidade
                If dQtdeCompra = 0 Then dQtdeCompra = 1
                                
                lErro = CF("UM_Conversao", objProduto.iClasseUM, objProduto.sSiglaUMVenda, objItemNF.sUnidadeMed, dFator)
                If lErro <> SUCESSO Then gError 201031
                                
                dNovoPreco = Arredonda_Moeda(dCoef * dNovoPreco * dFator / dQtdeCompra)
                
                dtDataVigencia = DATA_NULA
                 
                'Pesquisa no BD Ítens de Tabela de Preço com vigência a partir de hoje
                lErro = Comando_ExecutarPos(alComando(2), "SELECT Preco, DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia >= ? " & sFiltro & " ORDER BY DataVigencia DESC", 0, dPreco, dtDataVigencia, objNF.iFilialEmpresa, iCodTabela, objItemNF.sProduto, gdtDataAtual)
                If lErro <> AD_SQL_SUCESSO Then gError 200250
            
                'Verifica se existe preço
                lErro = Comando_BuscarPrimeiro(alComando(2))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200251
                
                Do While lErro = AD_SQL_SUCESSO
                
                    If (dNovoPreco - dPreco) > DELTA_VALORMONETARIO Then
                    
                         'Loca o registro
                        lErro = Comando_LockExclusive(alComando(2))
                        If lErro <> AD_SQL_SUCESSO Then gError 200252
                    
                        'Atualiza o preço
                        lErro = Comando_ExecutarPos(alComando(3), "UPDATE TabelasDePrecoItens SET Preco = ? ", alComando(2), dNovoPreco)
                        If lErro <> AD_SQL_SUCESSO Then gError 200253
                    
                    End If
                    
                    lErro = Comando_BuscarProximo(alComando(2))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200254
    
                Loop
                
                'se ainda nao existe preço com a data de hoje
                If dtDataVigencia <> gdtDataAtual Then
                
                    'Pesquisa no BD o preço atual
                    lErro = Comando_ExecutarPos(alComando(4), "SELECT Preco, DataVigencia FROM TabelasDePrecoItens WHERE FilialEmpresa = ? AND CodTabela = ? AND CodProduto = ? AND DataVigencia < ? ORDER BY DataVigencia DESC", 0, dPreco, dtDataVigencia, objNF.iFilialEmpresa, iCodTabela, objItemNF.sProduto, gdtDataAtual)
                    If lErro <> AD_SQL_SUCESSO Then gError 201032
                
                    'Verifica se existe preço
                    lErro = Comando_BuscarPrimeiro(alComando(4))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201033
                
                    If lErro <> AD_SQL_SUCESSO Or ((dNovoPreco - dPreco) > DELTA_VALORMONETARIO) Then
                    
                        lErro = Comando_Executar(alComando(5), "INSERT INTO TabelasDePrecoItens (CodTabela, FilialEmpresa, CodProduto, DataVigencia, Preco, Ativo, DataLog, Observacao) VALUES (?,?,?,?,?,?,?,?)", _
                            iCodTabela, objNF.iFilialEmpresa, objItemNF.sProduto, gdtDataAtual, dNovoPreco, 0, DATA_NULA, "Atualização automática. NF " & CStr(objNF.lNumNotaFiscal) & " de " & CStr(objNF.dtDataEmissao))
                        If lErro <> AD_SQL_SUCESSO Then gError 201034
                    
                    End If
                
                End If
        
            Next
        
            lErro = Comando_BuscarProximo(alComando(1))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200255
        
        Loop
     
        For iIndice = LBound(alComando) To UBound(alComando)
            Call Comando_Fechar(alComando(iIndice))
        Next
        
    End If

    TabelaPreco_Valida_Compra = SUCESSO

    Exit Function

Erro_TabelaPreco_Valida_Compra:

    TabelaPreco_Valida_Compra = gErr

    Select Case gErr

        Case 201030, 201031, 201036
        
        Case 201035
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_ENCONTRADO", gErr, objProduto.sCodigo)
        
        Case 200247
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 200248, 200249, 200255
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECO", gErr)
            
        Case 200250, 200251, 200254, 201032, 201033
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TABELASDEPRECOITENS", gErr)

        Case 201034
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_TABELASDEPRECOITENS", gErr)
        
        Case 200252
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_TABELASDEPRECO", gErr)
        
        Case 200253
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_TABELASDEPRECO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 200256)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'EXCLUSÃO
Public Function DASAliquotas_Exclui(ByVal objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_DASAliquotas_Exclui

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200912

    lErro = CF("DASAliquotas_Exclui_EmTrans", objDASAliquotas)
    If lErro <> SUCESSO Then gError 200913

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200914

    DASAliquotas_Exclui = SUCESSO

    Exit Function

Erro_DASAliquotas_Exclui:

    DASAliquotas_Exclui = gErr

    Select Case gErr

        Case 200912
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200913

        Case 200914
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200966)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function DASAliquotas_Exclui_EmTrans(ByVal objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_DASAliquotas_Exclui_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200915
    Next

    'Le a tabelaDASAliquotas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Ano FROM DASAliquotas WHERE Ano= ?  AND Mes= ? AND FilialEmpresa = ?", _
                0, iAux, objDASAliquotas.iAno, objDASAliquotas.iMes, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 200916

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200917

    'Se não existir => ERRO
    If lErro = AD_SQL_SEM_DADOS Then gError 200918

    'Trava registro contra alterações/Leituras
    lErro = Comando_LockExclusive(alComando(0))
    If lErro <> SUCESSO Then gError 200919

    lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM DASAliquotas", alComando(0))
    If lErro <> AD_SQL_SUCESSO Then gError 200920

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DASAliquotas_Exclui_EmTrans = SUCESSO

    Exit Function

Erro_DASAliquotas_Exclui_EmTrans:

    DASAliquotas_Exclui_EmTrans = gErr

    Select Case gErr

        Case 200915
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200916, 200917
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DASALIQUOTAS", gErr)

        Case 200918
            Call Rotina_Erro(vbOKOnly, "ERRO_DASALIQUOTAS_NAO_CADASTRADO", gErr, objDASAliquotas.iAno, objDASAliquotas.iMes)

        Case 200919
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DASALIQUOTAS", gErr)

        Case 200920
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_DASALIQUOTAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200965)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function DASAliquotas_Grava(ByVal objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_DASAliquotas_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 200921

    lErro = CF("DASAliquotas_Grava_EmTrans", objDASAliquotas)
    If lErro <> SUCESSO Then gError 200922

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 200923

    DASAliquotas_Grava = SUCESSO

    Exit Function

Erro_DASAliquotas_Grava:

    DASAliquotas_Grava = gErr

    Select Case gErr

        Case 200921
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 200922

        Case 200923
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200964)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function DASAliquotas_Grava_EmTrans(ByVal objDASAliquotas As ClassDASAliquotas) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_DASAliquotas_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 200924
    Next

    'Le a tabelaDASAliquotas
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Ano FROM DASAliquotas WHERE Ano= ?  AND Mes= ? AND FilialEmpresa = ? ", _
                0, iAux, objDASAliquotas.iAno, objDASAliquotas.iMes, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 200925

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200926

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 200927

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE DASAliquotas SET AliquotaICMSServ= ? ,AliquotaICMS= ? , AliquotaTotal= ? ", alComando(0), _
                objDASAliquotas.dAliquotaICMSServ, objDASAliquotas.dAliquotaICMS, objDASAliquotas.dAliquotaTotal)
        If lErro <> AD_SQL_SUCESSO Then gError 200928

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO DASAliquotas( Ano, Mes, AliquotaICMSServ, AliquotaICMS, AliquotaTotal, FilialEmpresa) VALUES (?,?,?,?,?,?)", _
                objDASAliquotas.iAno, objDASAliquotas.iMes, objDASAliquotas.dAliquotaICMSServ, objDASAliquotas.dAliquotaICMS, objDASAliquotas.dAliquotaTotal, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 200929

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    DASAliquotas_Grava_EmTrans = SUCESSO

    Exit Function

Erro_DASAliquotas_Grava_EmTrans:

    DASAliquotas_Grava_EmTrans = gErr

    Select Case gErr

        Case 200924
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 200925, 200926
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_DASALIQUOTAS", gErr)

        Case 200927
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_DASALIQUOTAS", gErr)

        Case 200928
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_DASALIQUOTAS", gErr)

        Case 200929
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_DASALIQUOTAS", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200963)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function NFE_Grava1(ByVal lNumIntNF As Long, ByVal objTipoDocInfo As ClassTipoDocInfo, Optional ByVal colNumIntNF As Collection) As Long

Dim lErro As Long
Dim lLote As Long
Dim objNFeFedLote As New ClassNFeFedLote
Dim colNFeFedLote As New Collection
Dim vbMsg As VbMsgBoxResult
Dim sDiretorio As String
Dim lRetorno As Long
Dim vValor As Variant
Dim bBloqueioFilial As Boolean, iFilialComPermissao As Integer
Dim iFilialEmpresa As Integer
Dim objVersao As New ClassVersaoNFe
Dim objNFOficial As New ClassNFiscal
Dim objNF As New ClassNFiscal, iScan As Integer
Dim sFiltro As String

On Error GoTo Erro_NFE_Grava1

    iFilialEmpresa = giFilialEmpresa
    If giFilialEmpresa > 50 Then 'giFilialEmpresa = giFilialEmpresa - 50
        Call CF("FilialEmpresa_Le_Filtro", sFiltro)
        If Len(Trim(sFiltro)) > 0 Then giFilialEmpresa = giFilialEmpresa - 50
    End If
    
    lErro = CF("Verifica_Bloqueios_Filial", giFilialEmpresa, bBloqueioFilial, iFilialComPermissao)
    If lErro <> SUCESSO Then gError 199991

    If gobjCRFAT.iUsaNFe = MARCADO And gobjCRFAT.iGravaNFe = MARCADO And (Not (bBloqueioFilial)) Then

        If objTipoDocInfo.iTipo = 1 Or objTipoDocInfo.iTipo = 2 Then
           
            'pede confirmacao
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ENVIO_NFE")
    
            If vbMsg = vbYes Then
    
                lErro = CF("Usuario_Configura_Aviso_NFe", gsUsuario, USUCONFIG_AVISO_NFE_SEM_AUTO)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
                'obtem o nº automatico
                lErro = CF("Config_ObterNumInt_Trans", "FATConfig", "NUM_PROX_LOTE_NFE", lLote)
                If lErro <> SUCESSO Then gError 199991
        
                If lNumIntNF <> 0 Then
                    If iFilialEmpresa > 50 And iFilialEmpresa <> giFilialEmpresa Then
                        objNFeFedLote.lNumIntNF = lNumIntNF - 1
                        
                        
                        'verifica se a nota oficial existe e é o clone da real
                        objNFOficial.lNumIntDoc = lNumIntNF - 1
                        
                        lErro = CF("NFiscal_Le", objNFOficial)
                        If lErro <> SUCESSO And lErro <> 31442 Then gError 207651
                        
                        If lErro <> SUCESSO Then gError 207652
                        
                        objNF.lNumIntDoc = lNumIntNF
                        
                        lErro = CF("NFiscal_Le", objNF)
                        If lErro <> SUCESSO And lErro <> 31442 Then gError 207653
                        
                        If lErro <> SUCESSO Then gError 207654
                        
                        If objNF.lNumNotaFiscal <> objNFOficial.lNumNotaFiscal Then gError 207655
                        
                    Else
                        objNFeFedLote.lNumIntNF = lNumIntNF
                    End If
                    objNFeFedLote.lLote = lLote
                            
                    colNFeFedLote.Add objNFeFedLote
                Else
                    For Each vValor In colNumIntNF
                        Set objNFeFedLote = New ClassNFeFedLote
                        
                        objNFeFedLote.lNumIntNF = vValor
                        objNFeFedLote.lLote = lLote
                                
                        colNFeFedLote.Add objNFeFedLote
                    Next
                    
                End If
                
                For Each objNFeFedLote In colNFeFedLote
                    lErro = CF("NFe_Valida_Envio_Cust", objNFeFedLote)
                    If lErro <> SUCESSO Then gError 199992
                Next
                
                lErro = CF("NFeFedScan_Verifica_Contingencia", giFilialEmpresa, Date, iScan)
                If lErro <> SUCESSO Then gError 199992
                
                lErro = CF("NFeFedLote_Grava", giFilialEmpresa, colNFeFedLote, IIf(iScan = MARCADO, gobjCRFAT.sNFeSistemaContingencia, ""))
                If lErro <> SUCESSO Then gError 199992
            
                objVersao.iCodigo = gobjCRFAT.iVersaoNFE
                
                lErro = CF("VersaoNFe_Le", objVersao)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 199992
            
                sDiretorio = String(255, 0)
                lRetorno = GetPrivateProfileString("Forprint", "DirBin", "c:\sge\programa\", sDiretorio, 255, NOME_ARQUIVO_ADM)
                sDiretorio = left(sDiretorio, lRetorno)
                        
                lErro = WinExec(sDiretorio & objVersao.sProgramaEnvio & " Envio " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & CStr(lLote) & " " & IIf(iScan = MARCADO, gobjCRFAT.sNFeSistemaContingencia, ""), SW_NORMAL)
     
                Call Rotina_Aviso(vbOK, "AVISO_INICIO_ENVIONFE", lLote)
    
                lErro = CF("NFE_Trata_Nota_Denegada")
                If lErro <> SUCESSO Then gError 207390
    
            End If
    
        End If
    
    End If

    giFilialEmpresa = iFilialEmpresa
    
    NFE_Grava1 = SUCESSO

    Exit Function

Erro_NFE_Grava1:

    NFE_Grava1 = gErr

    giFilialEmpresa = iFilialEmpresa
    
    Select Case gErr

        Case 199991, 199992, 207390, 207651, 207653
        
        Case 207652, 207654
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr)

        Case 207655
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INVALIDA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199993)

    End Select

    Exit Function

End Function

Function NFSE_Grava1(ByVal lNumIntNF As Long, ByVal objTipoDocInfo As ClassTipoDocInfo, Optional ByVal colNumIntNF As Collection) As Long

Dim lErro As Long
Dim lLote As Long
Dim objNFeFedLote As New ClassNFeFedLote
Dim colNFeFedLote As New Collection
Dim vbMsg As VbMsgBoxResult
Dim vValor As Variant
Dim bBloqueioFilial As Boolean, iFilialComPermissao As Integer
Dim sNFSEEXE As String

On Error GoTo Erro_NFSE_Grava1

    lErro = CF("Verifica_Bloqueios_Filial", giFilialEmpresa, bBloqueioFilial, iFilialComPermissao)
    If lErro <> SUCESSO Then gError 207054

    If gobjCRFAT.iUsaNFSE = MARCADO And gobjCRFAT.iGravaNFSE = MARCADO And (Not (bBloqueioFilial)) Then

            'pede confirmacao
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_ENVIO_NFSE")
    
            If vbMsg = vbYes Then
    
                'obtem o nº automatico
                lErro = CF("Config_ObterNumInt_Trans", "FATConfig", "NUM_PROX_LOTE_RPSWEB", lLote)
                If lErro <> SUCESSO Then gError 207055
        
                If lNumIntNF <> 0 Then
                    objNFeFedLote.lNumIntNF = lNumIntNF
                    objNFeFedLote.lLote = lLote
                            
                    colNFeFedLote.Add objNFeFedLote
                Else
                    For Each vValor In colNumIntNF
                        Set objNFeFedLote = New ClassNFeFedLote
                        
                        objNFeFedLote.lNumIntNF = vValor
                        objNFeFedLote.lLote = lLote
                                
                        colNFeFedLote.Add objNFeFedLote
                    Next
                    
                End If
                        
                        
                lErro = CF("RPSWEBLote_Grava", giFilialEmpresa, colNFeFedLote)
                If lErro <> SUCESSO Then gError 207056
            
                lErro = CF("NFSE_Obter_EXE", giFilialEmpresa, sNFSEEXE)
                If lErro <> SUCESSO Then gError 201186
                
                lErro = WinExec(sNFSEEXE & " Envio " & CStr(glEmpresa) & " " & CStr(giFilialEmpresa) & " " & CStr(lLote), SW_NORMAL)
    
                Call Rotina_Aviso(vbOK, "AVISO_INICIO_ENVIONFE", lLote)
    
            End If
    
    
    End If

    NFSE_Grava1 = SUCESSO

    Exit Function

Erro_NFSE_Grava1:

    NFSE_Grava1 = gErr

    Select Case gErr

        Case 207054 To 207056, 201186

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207057)

    End Select

    Exit Function

End Function

Public Function NFE_Trata_Nota_Denegada() As Long

Dim lErro As Long, lTransacao As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntNF As Long
Dim objNFiscal As ClassNFiscal
Dim dtDataCancelamento As Date, objTipoDocInfo As ClassTipoDocInfo

On Error GoTo Erro_NFE_Trata_Nota_Denegada

    'Abre a transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 207382

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207381
    Next

    'Le as NFeFedDenegada ainda nao processadas
    lErro = Comando_Executar(alComando(0), "SELECT NumIntNF FROM NFeFedDenegada WHERE Processada= 0", lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 207383

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207384

    Do While lErro = SUCESSO

        Set objNFiscal = New ClassNFiscal
        objNFiscal.lNumIntDoc = lNumIntNF
    
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 207384
        
        If lErro <> SUCESSO Then gError 210481
        
        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 210476
        
        dtDataCancelamento = gdtDataHoje
        objNFiscal.sMotivoCancel = "Nota Fiscal denegada."
        
        'Lê o tipodocinfo da nota fiscal
        Set objTipoDocInfo = New ClassTipoDocInfo
        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 207384
        
        If lErro <> SUCESSO Then gError 210482
        
        'Cancela a Nota Fiscal
        If objTipoDocInfo.iTipo = DOCINFO_NF_INT_SAIDA Then
            lErro = CF("NotaFiscalSaida_Cancelar_EmTrans1", objNFiscal, dtDataCancelamento)
            If lErro <> SUCESSO Then gError 207385
        Else
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscalEntrada_Cancelar_EmTrans1", objNFiscal, dtDataCancelamento)
            If lErro <> SUCESSO Then gError 207385
        End If
        
        lErro = Comando_Executar(alComando(1), "UPDATE NFeFedDenegada SET Processada= 1 WHERE NumIntNF = ?", lNumIntNF)
        If lErro <> AD_SQL_SUCESSO Then gError 207386
    
        'Busca Primeiro
        lErro = Comando_BuscarProximo(alComando(0))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207387
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    'Confirma a transacao
    lErro = Transacao_Commit()
    If lErro <> SUCESSO Then gError 207388

    NFE_Trata_Nota_Denegada = SUCESSO

    Exit Function

Erro_NFE_Trata_Nota_Denegada:

    NFE_Trata_Nota_Denegada = gErr

    Select Case gErr

        Case 207381
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207382
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 207383, 207384, 207387
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFEFEDDENEGADA", gErr)

        Case 207385, 210476

        Case 207388
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case 210481
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 210482
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOCINFO_NAO_CADASTRADO2", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207389)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next


    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

'ROTINAS CRIADAS AUTOMATICAMENTE PELA TELA BROWSECRIA
'GRAVAÇÃO
Public Function RetiradaEntrega_Grava(ByVal objRetiradaEntrega As ClassRetiradaEntrega) As Long

Dim lErro As Long
Dim lTransacao As Long

On Error GoTo Erro_RetiradaEntrega_Grava

    'Abertura de transação
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 207741

    lErro = CF("RetiradaEntrega_Grava_EmTrans", objRetiradaEntrega)
    If lErro <> SUCESSO Then gError 207742

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 207743

    RetiradaEntrega_Grava = SUCESSO

    Exit Function

Erro_RetiradaEntrega_Grava:

    RetiradaEntrega_Grava = gErr

    Select Case gErr

        Case 207741
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 207742

        Case 207743
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207744)

    End Select

    'Desfaz Transação
    Call Transacao_Rollback

    Exit Function

End Function

Public Function InfoAdicionais_Grava(ByVal objInfoAdic As ClassInfoAdic) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, sNaturezaAux As String

On Error GoTo Erro_InfoAdicionais_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 201509
    Next
    
    If gobjCRFAT.iInfoAdicNatObrigatorio = MARCADO And Len(Trim(objInfoAdic.sNatureza)) = 0 Then gError 213145
    If gobjCRFAT.iInfoAdicCclObrigatorio = MARCADO And Len(Trim(objInfoAdic.sCcl)) = 0 Then gError 213146

    'Le a tabela InfoAdicDoc
    sNaturezaAux = String(STRING_NATMOVCTA_CODIGO, 0)
    lErro = Comando_ExecutarPos(alComando(0), "SELECT Natureza FROM InfoAdicDoc WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, sNaturezaAux, objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 201510

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201511

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 201512
        
        If Len(Trim(objInfoAdic.sNatureza)) <> 0 Or Len(Trim(objInfoAdic.sCcl)) <> 0 Then
        
            lErro = Comando_ExecutarPos(alComando(1), "UPDATE InfoAdicDoc SET Natureza = ?, Ccl = ? ", alComando(0), objInfoAdic.sNatureza, objInfoAdic.sCcl)
            If lErro <> AD_SQL_SUCESSO Then gError 201513
    
        Else
        
            lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InfoAdicDoc", alComando(0))
            If lErro <> AD_SQL_SUCESSO Then gError 201515
        
        End If

    Else
    
        If Len(Trim(objInfoAdic.sNatureza)) <> 0 Or Len(Trim(objInfoAdic.sCcl)) <> 0 Then
        
            lErro = Comando_Executar(alComando(1), "INSERT INTO InfoAdicDoc( TipoDoc, NumIntDoc, Natureza, Ccl) VALUES (?,?,?,?)", _
                    objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc, objInfoAdic.sNatureza, objInfoAdic.sCcl)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201514
        
        End If
        
    End If
    
    If Not (objInfoAdic.objRetEnt Is Nothing) Then
        objInfoAdic.objRetEnt.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objRetEnt.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = RetiradaEntrega_Grava_EmTrans(objInfoAdic.objRetEnt)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Else
        lErro = RetiradaEntrega_Exclui(objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If

    If Not (objInfoAdic.objCompra Is Nothing) Then
        objInfoAdic.objCompra.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objCompra.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = InfoAdicCompra_Grava(objInfoAdic.objCompra)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Else
        lErro = InfoAdicCompra_Exclui(objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    If Not (objInfoAdic.objExportacao Is Nothing) Then
        objInfoAdic.objExportacao.iTipoDoc = objInfoAdic.iTipoDoc
        objInfoAdic.objExportacao.lNumIntDoc = objInfoAdic.lNumIntDoc
        lErro = InfoAdicExportacao_Grava(objInfoAdic.objExportacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Else
        lErro = InfoAdicExportacao_Exclui(objInfoAdic.iTipoDoc, objInfoAdic.lNumIntDoc)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    End If
    
    Call CF("InfoAdic_AjustaAnexo", objInfoAdic)
    lErro = CF("Anexos_Grava", objInfoAdic.objAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InfoAdicionais_Grava = SUCESSO

    Exit Function

Erro_InfoAdicionais_Grava:

    InfoAdicionais_Grava = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201509
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 201510, 201511
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICDOC", gErr)

        Case 201512
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICDOC", gErr)

        Case 201513
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INFOADICDOC", gErr)

        Case 201514
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INFOADICDOC", gErr)
            
        Case 213145
            Call Rotina_Erro(vbOKOnly, "ERRO_INFOADICDOC_NAT_OBRIGATORIO", gErr)
        
        Case 213146
            Call Rotina_Erro(vbOKOnly, "ERRO_INFOADICDOC_CCL_OBRIGATORIO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209194)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function RetiradaEntrega_Grava_EmTrans(ByVal objRetiradaEntrega As ClassRetiradaEntrega) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lEnderecoRet As Long, lEnderecoEnt As Long
Dim lCodigo As Long
Dim lCodEndereco As Long

On Error GoTo Erro_RetiradaEntrega_Grava_EmTrans

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207715
    Next

    'Le a tabelaRetiradaEntrega
    lErro = Comando_ExecutarPos(alComando(0), "SELECT EnderecoRet, EnderecoEnt  FROM RetiradaEntrega WHERE TipoDoc = ? AND NumIntDoc = ?", _
                0, lEnderecoRet, lEnderecoEnt, objRetiradaEntrega.iTipoDoc, objRetiradaEntrega.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 207716

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207717

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 207718

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE RetiradaEntrega SET CNPJCPFRet= ? , " & _
                "CNPJCPFEnt= ? , ClienteRet= ? , FornecedorRet= ? , FilialCliRet= ? , FilialFornRet= ? , " & _
                "ClienteEnt= ? , FornecedorEnt= ? , FilialCliEnt= ? , FilialFornEnt= ? ", alComando(0), _
                objRetiradaEntrega.sCNPJCPFRet, objRetiradaEntrega.sCNPJCPFEnt, objRetiradaEntrega.lClienteRet, objRetiradaEntrega.lFornecedorRet, objRetiradaEntrega.iFilialCliRet, objRetiradaEntrega.iFilialFornRet, _
                objRetiradaEntrega.lClienteEnt, objRetiradaEntrega.lFornecedorEnt, objRetiradaEntrega.iFilialCliEnt, objRetiradaEntrega.iFilialFornEnt)
        If lErro <> AD_SQL_SUCESSO Then gError 207719

        objRetiradaEntrega.objEnderecoEnt.lCodigo = lEnderecoEnt

        'Grava o Endereco do Fornecedor
        lErro = CF("Endereco_Grava", objRetiradaEntrega.objEnderecoEnt)
        If lErro <> SUCESSO Then gError 207734

        objRetiradaEntrega.objEnderecoRet.lCodigo = lEnderecoRet

        'Grava o Endereco do Fornecedor
        lErro = CF("Endereco_Grava", objRetiradaEntrega.objEnderecoRet)
        If lErro <> SUCESSO Then gError 207735

    Else

        'Gera codigo para o novo endereco
        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 207736

        objRetiradaEntrega.objEnderecoEnt.lCodigo = lCodEndereco
        objRetiradaEntrega.lEnderecoEnt = lCodEndereco
        
        'Grava o Endereco do Fornecedor
        lErro = CF("Endereco_Grava", objRetiradaEntrega.objEnderecoEnt)
        If lErro <> SUCESSO Then gError 207737

        'Gera codigo para o novo endereco
        lErro = CF("Endereco_Automatico", lCodEndereco)
        If lErro <> SUCESSO Then gError 207738

        objRetiradaEntrega.objEnderecoRet.lCodigo = lCodEndereco
        objRetiradaEntrega.lEnderecoRet = lCodEndereco
        
        'Grava o Endereco do Fornecedor
        lErro = CF("Endereco_Grava", objRetiradaEntrega.objEnderecoRet)
        If lErro <> SUCESSO Then gError 207739

        lErro = CF("CRFATConfig_ObterNumInt1", "NUM_PROX_RETIRADAENTREGA", lCodigo)
        If lErro <> SUCESSO Then gError 207740

        lErro = Comando_Executar(alComando(1), "INSERT INTO RetiradaEntrega( TipoDoc, NumIntDoc, EnderecoRet, EnderecoEnt, CNPJCPFRet, " & _
                "CNPJCPFEnt, ClienteRet, FornecedorRet, FilialCliRet, FilialFornRet, " & _
                "ClienteEnt, FornecedorEnt, FilialCliEnt, FilialFornEnt) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
                objRetiradaEntrega.iTipoDoc, objRetiradaEntrega.lNumIntDoc, objRetiradaEntrega.lEnderecoRet, objRetiradaEntrega.lEnderecoEnt, objRetiradaEntrega.sCNPJCPFRet, _
                objRetiradaEntrega.sCNPJCPFEnt, objRetiradaEntrega.lClienteRet, objRetiradaEntrega.lFornecedorRet, objRetiradaEntrega.iFilialCliRet, objRetiradaEntrega.iFilialFornRet, _
                objRetiradaEntrega.lClienteEnt, objRetiradaEntrega.lFornecedorEnt, objRetiradaEntrega.iFilialCliEnt, objRetiradaEntrega.iFilialFornEnt)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207720

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RetiradaEntrega_Grava_EmTrans = SUCESSO

    Exit Function

Erro_RetiradaEntrega_Grava_EmTrans:

    RetiradaEntrega_Grava_EmTrans = gErr

    Select Case gErr

        Case 207715
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207716, 207717
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RETIRADAENTREGA", gErr)

        Case 207718
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_RETIRADAENTREGA", gErr)

        Case 207719
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_RETIRADAENTREGA", gErr)

        Case 207720
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RETIRADAENTREGA", gErr)
            
        Case 207734 To 207740

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207721)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function InfoAdicionais_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer, objInfoAdic As New ClassInfoAdic

On Error GoTo Erro_InfoAdicionais_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209201
    Next

    'Le a tabela InfoAdicDoc
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoDoc FROM InfoAdicDoc WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, iAux, iTipoDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209202

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209203

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209204
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InfoAdicDoc", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 209205
        
    End If

    lErro = RetiradaEntrega_Exclui(iTipoDoc, lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = InfoAdicCompra_Exclui(iTipoDoc, lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = InfoAdicExportacao_Exclui(iTipoDoc, lNumIntDoc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call CF("InfoAdic_AjustaAnexo", objInfoAdic)
    lErro = CF("Anexos_Exclui", objInfoAdic.objAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    InfoAdicionais_Exclui = SUCESSO

    Exit Function

Erro_InfoAdicionais_Exclui:

    InfoAdicionais_Exclui = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209201
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209202, 209203
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICDOC", gErr)

        Case 209204
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICDOC", gErr)

        Case 209205
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INFOADICDOC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209194)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Public Function RetiradaEntrega_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim lEnderecoEnt As Long, lEnderecoRet As Long

On Error GoTo Erro_RetiradaEntrega_Exclui
    
    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 207726
    Next
    
    'Le a tabelaRetiradaEntrega
    lErro = Comando_ExecutarPos(alComando(0), "SELECT EnderecoEnt, EnderecoRet  FROM RetiradaEntrega WHERE TipoDoc = ? AND NumIntDoc = ?", _
                0, lEnderecoEnt, lEnderecoRet, iTipoDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 207727

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 207728

    'Se não existir => ERRO
    If lErro <> AD_SQL_SEM_DADOS Then 'gError 207729

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 207730
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM RetiradaEntrega", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 207731
    
        lErro = CF("Endereco_Exclui", lEnderecoEnt)
        If lErro <> SUCESSO Then gError 207746
    
        lErro = CF("Endereco_Exclui", lEnderecoRet)
        If lErro <> SUCESSO Then gError 207747
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
        
    RetiradaEntrega_Exclui = SUCESSO

    Exit Function

Erro_RetiradaEntrega_Exclui:

    RetiradaEntrega_Exclui = gErr

    Select Case gErr

        Case 207726
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 207727, 207728
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RETIRADAENTREGA", gErr)

        Case 207729, 207746, 207747

        Case 207730
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_RETIRADAENTREGA", gErr)

        Case 207731
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_RETIRADAENTREGA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 207732)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function InfoAdicExportacao_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_InfoAdicExportacao_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209195
    Next

    'Le a tabelaInfoAdicExportacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoDoc FROM InfoAdicExportacao WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, iAux, iTipoDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209196

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209197

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209198
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InfoAdicExportacao", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 209199
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InfoAdicExportacao_Exclui = SUCESSO

    Exit Function

Erro_InfoAdicExportacao_Exclui:

    InfoAdicExportacao_Exclui = gErr

    Select Case gErr

        Case 209195
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209196, 209197
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICEXPORTACAO", gErr)

        Case 209198
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICEXPORTACAO", gErr)

        Case 209199
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INFOADICEXPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209200)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Public Function InfoAdicCompra_Exclui(ByVal iTipoDoc As Integer, ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer

On Error GoTo Erro_InfoAdicCompra_Exclui

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209201
    Next

    'Le a tabelaInfoAdicCompra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoDoc FROM InfoAdicCompra WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, iAux, iTipoDoc, lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209202

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209203

    If lErro <> AD_SQL_SEM_DADOS Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209204
    
        lErro = Comando_ExecutarPos(alComando(1), "DELETE FROM InfoAdicCompra", alComando(0))
        If lErro <> AD_SQL_SUCESSO Then gError 209205
        
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InfoAdicCompra_Exclui = SUCESSO

    Exit Function

Erro_InfoAdicCompra_Exclui:

    InfoAdicCompra_Exclui = gErr

    Select Case gErr

        Case 209201
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209202, 209203
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICCOMPRA", gErr)

        Case 209204
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICCOMPRA", gErr)

        Case 209205
            Call Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_INFOADICCOMPRA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209206)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function InfoAdicCompra_Grava(ByVal objInfoAdicCompra As ClassInfoAdicCompra) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_InfoAdicCompra_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209207
    Next

    'Le a tabelaInfoAdicCompra
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoDoc FROM InfoAdicCompra WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, iAux, objInfoAdicCompra.iTipoDoc, objInfoAdicCompra.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209208

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209209

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209210

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE InfoAdicCompra SET NotaEmpenho= ? , Pedido= ? , Contrato= ? ", alComando(0), _
                objInfoAdicCompra.sNotaEmpenho, objInfoAdicCompra.sPedido, objInfoAdicCompra.sContrato)
        If lErro <> AD_SQL_SUCESSO Then gError 209211

    Else
        lErro = Comando_Executar(alComando(1), "INSERT INTO InfoAdicCompra( TipoDoc, NumIntDoc, NotaEmpenho, Pedido, Contrato) VALUES (?,?,?,?,?)", _
                objInfoAdicCompra.iTipoDoc, objInfoAdicCompra.lNumIntDoc, objInfoAdicCompra.sNotaEmpenho, objInfoAdicCompra.sPedido, objInfoAdicCompra.sContrato)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209212
    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InfoAdicCompra_Grava = SUCESSO

    Exit Function

Erro_InfoAdicCompra_Grava:

    InfoAdicCompra_Grava = gErr

    Select Case gErr

        Case 209207
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209208, 209209
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICCOMPRA", gErr)

        Case 209210
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICCOMPRA", gErr)

        Case 209211
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INFOADICCOMPRA", gErr)

        Case 209212
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INFOADICCOMPRA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209213)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function InfoAdicExportacao_Grava(ByVal objInfoAdicExportacao As ClassInfoAdicExportacao) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer
Dim iAux As Integer
Dim lNumIntDoc As Long

On Error GoTo Erro_InfoAdicExportacao_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209214
    Next

    'Le a tabelaInfoAdicExportacao
    lErro = Comando_ExecutarPos(alComando(0), "SELECT TipoDoc FROM InfoAdicExportacao WHERE TipoDoc= ?  AND NumIntDoc= ? ", _
                0, iAux, objInfoAdicExportacao.iTipoDoc, objInfoAdicExportacao.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 209215

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(0))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209216

    'Se existir => UPDATE, senão => INSERT
    If lErro = AD_SQL_SUCESSO Then

        'Trava registro contra alterações/Leituras
        lErro = Comando_LockExclusive(alComando(0))
        If lErro <> SUCESSO Then gError 209217

        lErro = Comando_ExecutarPos(alComando(1), "UPDATE InfoAdicExportacao SET UFEmbarque= ? , LocalEmbarque= ?, NumIntDE = ?, NumRE = ? ", alComando(0), _
                objInfoAdicExportacao.sUFEmbarque, objInfoAdicExportacao.sLocalEmbarque, objInfoAdicExportacao.lNumIntDE, objInfoAdicExportacao.sNumRE)
        If lErro <> AD_SQL_SUCESSO Then gError 209218

    Else

        lErro = Comando_Executar(alComando(1), "INSERT INTO InfoAdicExportacao( TipoDoc, NumIntDoc, UFEmbarque, LocalEmbarque, NumIntDE, NumRE) VALUES (?,?,?,?,?,?)", _
                objInfoAdicExportacao.iTipoDoc, objInfoAdicExportacao.lNumIntDoc, objInfoAdicExportacao.sUFEmbarque, objInfoAdicExportacao.sLocalEmbarque, objInfoAdicExportacao.lNumIntDE, objInfoAdicExportacao.sNumRE)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209219

    End If

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    InfoAdicExportacao_Grava = SUCESSO

    Exit Function

Erro_InfoAdicExportacao_Grava:

    InfoAdicExportacao_Grava = gErr

    Select Case gErr

        Case 209214
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209215, 209216
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_INFOADICEXPORTACAO", gErr)

        Case 209217
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCKEXCLUSIVE_INFOADICEXPORTACAO", gErr)

        Case 209218
            Call Rotina_Erro(vbOKOnly, "ERRO_UPDATE_INFOADICEXPORTACAO", gErr)

        Case 209219
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_INFOADICEXPORTACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209220)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Private Function RelacCliTexto_Grava(ByVal objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim alComando(0 To 1) As Long
Dim iIndice As Integer, sAux As String, sTexto As String
Dim iSeq As Integer

On Error GoTo Erro_RelacCliTexto_Grava

    'Abertura de Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209818
    Next
    
    lErro = Comando_Executar(alComando(0), "DELETE RelacCliTexto WHERE Codigo = ? AND FilialEmpresa = ?", objRelacCli.lCodigo, objRelacCli.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209819
    
    sAux = objRelacCli.sAssunto1 & objRelacCli.sAssunto2
    iSeq = 0
    
    Do While Len(Trim(sAux)) > 0
    
        sTexto = left(sAux, 250)
        sAux = Mid(sAux, 251)
        iSeq = iSeq + 1
    
        lErro = Comando_Executar(alComando(1), "INSERT INTO RelacCliTexto (Codigo, FilialEmpresa, Seq, Texto) VALUES (?,?,?,?)", objRelacCli.lCodigo, objRelacCli.iFilialEmpresa, iSeq, sTexto)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209820
    
    Loop
    
    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RelacCliTexto_Grava = SUCESSO

    Exit Function

Erro_RelacCliTexto_Grava:

    RelacCliTexto_Grava = gErr

    Select Case gErr

        Case 209818
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209819
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_RELACIONAMENTOCLIENTES", gErr)
        
        Case 209820
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_RELACIONAMENTOCLIENTES", gErr, objRelacCli.lCodigo)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209821)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Private Function RelacCliTexto_Exclui(ByVal objRelacCli As ClassRelacClientes) As Long

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_RelacCliTexto_Exclui

    'Abertura de Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 209818
    
    lErro = Comando_Executar(lComando, "DELETE RelacCliTexto WHERE Codigo = ? AND FilialEmpresa = ?", objRelacCli.lCodigo, objRelacCli.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209819
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    RelacCliTexto_Exclui = SUCESSO

    Exit Function

Erro_RelacCliTexto_Exclui:

    RelacCliTexto_Exclui = gErr

    Select Case gErr

        Case 209818
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209819
            Call Rotina_Erro(vbOKOnly, "ERRO_DELETE_RELACIONAMENTOCLIENTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209821)

    End Select

    'Fecha Comando
    Call Comando_Fechar(lComando)

    Exit Function
    
End Function

Public Function CRFATConfig_Grava2(ByVal objCRFATConfig As ClassCRFATConfig) As Long
'grava na tabela CRFATConfig

Dim lTransacao As Long
Dim lErro As Long
Dim alComando(1 To 2) As Long
Dim iIndice As Integer, iTipo As Integer

On Error GoTo Erro_CRFATConfig_Grava2
    
    'Abrir transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 209831

    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 209832
    Next
    
    lErro = Comando_ExecutarPos(alComando(1), "SELECT Tipo FROM CRFATConfig WHERE Codigo = ? AND FilialEmpresa = ?", 0, iTipo, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 209833
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 209834
    
    'Se não encontrou o registro
    If lErro = AD_SQL_SEM_DADOS Then
        lErro = Comando_Executar(alComando(2), "INSERT INTO CRFATConfig (Codigo, FilialEmpresa, Descricao, Tipo, Conteudo, ConteudoInicial) VALUES (?,?,?,?,?,?) ", _
                                               objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa, objCRFATConfig.sDescricao, objCRFATConfig.iTipo, objCRFATConfig.sConteudo, objCRFATConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 209835
    Else
        lErro = Comando_ExecutarPos(alComando(2), "UPDATE CRFATConfig SET Conteudo = ?", alComando(1), objCRFATConfig.sConteudo)
        If lErro <> AD_SQL_SUCESSO Then gError 209836
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    'Fechar transacao
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 209837
          
    CRFATConfig_Grava2 = SUCESSO
    
    Exit Function
    
Erro_CRFATConfig_Grava2:
    
    CRFATConfig_Grava2 = gErr
    
    Select Case gErr
    
        Case 209831
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 209832
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 209833, 209834
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CRFATCONFIG2", gErr, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
     
        Case 209835
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_CRFATCONFIG", gErr, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)
        
        Case 209836
            Call Rotina_Erro(vbOKOnly, "ERRO_ATUALIZACAO_CRFATCONFIG", gErr, objCRFATConfig.sCodigo, objCRFATConfig.iFilialEmpresa)

        Case 209837
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209838)
            
    End Select
    
    Call Transacao_Rollback
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

Private Function TituloReceber_Grava_Trata_NFe(ByVal objTituloReceber As ClassTituloReceber, ByVal colParcelasRec As colParcelaReceber) As Long
'Se o titulo for de uma nfe autorizada, grava uma copia de como era a cobranca antes da alteracao que está sendo feita

Dim lErro As Long
Dim alComando(1 To 3) As Long
Dim iIndice As Integer, lNumIntDocNF As Long
Dim iParcela As Integer, schNFe As String
Dim objParcRec As ClassParcelaReceber

On Error GoTo Erro_TituloReceber_Grava_Trata_NFe
    
    'abrir comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 211722
    Next
    
    'Verifica se está vinculada a uma NFe autorizada
    schNFe = String(STRING_NFE_CHNFE, 0)
    lErro = Comando_Executar(alComando(1), "SELECT NFNumIntDoc, chNFe FROM Danfe WHERE NFClasseDocCPR = ? AND NFNumIntDocCPR = ?", lNumIntDocNF, schNFe, CPR_TITULO_RECEBER, objTituloReceber.lNumIntDoc)
    If lErro <> AD_SQL_SUCESSO Then gError 211723
    
    'Lê o primeiro registro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211724
    
    'Se encontrou o registro é porque a NFe já foi autorizada, então grava o histórico se não existir
    If lErro <> AD_SQL_SEM_DADOS Then
    
        'Verifica se já existem os dados da parcela que foram na NFe
        lErro = Comando_Executar(alComando(2), "SELECT NumParcela FROM DanfeCobranca WHERE chNFe = ?", iParcela, schNFe)
        If lErro <> AD_SQL_SUCESSO Then gError 211725
        
        'Lê o primeiro registro
        lErro = Comando_BuscarPrimeiro(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 211726
    
        'Se ainda não tem os dados vinculados a NFe gravados, grava
        If lErro = AD_SQL_SEM_DADOS Then
        
            For Each objParcRec In colParcelasRec
    
                lErro = Comando_Executar(alComando(3), "INSERT INTO DanfeCobranca (chNFe, NumParcela, DataVencimento, Valor) VALUES (?,?,?,?) ", _
                                                       schNFe, objParcRec.iNumParcela, objParcRec.dtDataVencimento, objParcRec.dValor)
                If lErro <> AD_SQL_SUCESSO Then gError 211727
                
            Next
        
        End If
        
    End If
    
   'fechar comandos
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
          
    TituloReceber_Grava_Trata_NFe = SUCESSO
    
    Exit Function
    
Erro_TituloReceber_Grava_Trata_NFe:
    
    TituloReceber_Grava_Trata_NFe = gErr
    
    Select Case gErr
        
        Case 211722
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 211723, 211724
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL", gErr)
        
        Case 211725, 211726
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
     
        Case 211727
            Call Rotina_Erro(vbOKOnly, "ERRO_INSERCAO_PARCELAS_REC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 211728)
            
    End Select
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function
    
End Function

