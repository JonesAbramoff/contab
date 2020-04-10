Attribute VB_Name = "PrincECF"
Option Explicit

'#Const TESTE_SEM_IMPRESSORA = 1

Sub Main()

Dim lSistema As Long, lErro As Long, sCodTela As String, sProjeto As String, sClasse As String
Dim objUsuarioEmpresa As New ClassUsuarioEmpresa
Dim bSplashFormLoaded As Boolean
Dim objFilialEmpresa As New ClassFilialEmpresa
Dim objAux As Object
Dim objOperador As ClassOperador
Dim sRetorno As String
Dim lTamanho As Long
Dim lSequencial As String
Dim iStatusCaixa As Integer
Dim objObject As Object
Dim objTela As Object
Dim objLojaArqFisMestre As New ClassLojaArqFisMestre
Dim colLojaArqFisAnalitico As New Collection
Dim objUltimaReducao As New ClassUltimaReducao
Dim objAliquota As ClassAliquotaICMS, dtUltimoMovto As Date

'Dim iRZPendenteDiaAnterior As Integer
Dim iRZPendente As Integer
Dim bBackup As Boolean

On Error GoTo Erro_Main
    
    bBackup = False
    If InStr(UCase(Command$), UCase("-Backup")) <> 0 Then
        bBackup = True
    End If

    If (Not bBackup) And App.PrevInstance = True Then
        MsgBox "O SGEECF já está sendo executado...."
        Exit Sub
    End If
    
    gbVPN = False
    
    giLocalOperacao = LOCALOPERACAO_ECF
    
    gsNomePrinc = "SGEECF"
    
    Call Inicializa_Tamanhos_String
        
    frmSplashECF.Show
    DoEvents
    
    bSplashFormLoaded = True

    App.HelpFile = App.Path & "\sgeprinc.hlp"
    
    DoEvents
    
'    'para permitir acessar o dicionario de dados
'    lSistema = Sistema_Abrir()
'    If lSistema = 0 Then gError 126112
'
'    'faz login utilizando o codigo do usuario e a senha supervisor
'
'    objUsuarioEmpresa.sCodUsuario = "supervisor"
'    objUsuarioEmpresa.sSenha = "24137134"
'
'    lErro = Usuario_Executa_Login(objUsuarioEmpresa.sCodUsuario, objUsuarioEmpresa.sSenha)
'    If lErro <> SUCESSO Then gError 126113

    Call ECF_Grava_Log("Abertura do Sistema.")
    
    'apenas para agilizar cargas futuras de telas
    Call Tela_ObterFuncao(sCodTela, sProjeto, sClasse)
    
    lTamanho = 150
    sRetorno = String(lTamanho, 0)
    
    'verifica se o caixa possui ECF
    Call GetPrivateProfileString(APLICACAO_CAIXA, "Debug", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    
    giDebug = StrParaInt(sRetorno)
    If InStr(UCase(Command$), UCase("-Debug")) <> 0 Then giDebug = 1
    
    gdtDataAnterior = DATA_NULA
    gdtDataAtualizacaoDadosCCC = DATA_NULA
    gdtDataHoraFimPapel = DATA_NULA
    gdtUltimaReducao = DATA_NULA
    
    Set gcolModulo = New AdmColModulo
            
    Dim sMD5 As String
    Dim sNomeCompleto As String
    
    lErro = Carrega_Sistema()
    If lErro <> SUCESSO Then gError 99597
    
    lErro = Inicializa_ECF()
    If lErro <> SUCESSO Then gError 99870
        
    lErro = TestaVersaoPgm()
    If lErro <> SUCESSO Then gError 99870
        
    lErro = TestaVersaoECF()
    If lErro <> SUCESSO Then gError 99870
    
    lErro = Caixa_Verifica_Consistencia
    If lErro <> SUCESSO Then gError 99870
      
    lErro = CF_ECF("Requisito_IX")
    If lErro <> SUCESSO Then gError 204860

    lErro = CF_ECF("Requisito_XI")
    If lErro <> SUCESSO Then gError 210241

    lErro = CF_ECF("SaldoEmDinheiro_Trata_Carga")
    If lErro <> SUCESSO Then gError 210241

    lErro = Inicializa_LeitoraCodBarras()
    If lErro <> SUCESSO Then gError 117683

    lErro = AFRAC_Carrega_Dados_UltimaReducao(objUltimaReducao, objLojaArqFisMestre, colLojaArqFisAnalitico)
    If lErro <> SUCESSO Then gError 133665
    
    
    If giDebug = 1 Then MsgBox ("Programar Aliquotas 1")
    
    'programa as aliquotas
    lErro = CF_ECF("Programar_Aliquotas")
    If lErro <> SUCESSO Then gError 214515
    
    If giDebug = 1 Then MsgBox ("Programar Aliquotas 2")
                    
    
    lErro_Chama_Tela = SUCESSO
    
    If giDebug = 1 Then MsgBox ("32")
    
    'carrega a tela principal
    ECF.Show
    
    If giDebug = 1 Then MsgBox ("33")
    
    If lErro_Chama_Tela <> SUCESSO Then Unload ECF
    
    Unload frmSplashECF
    
    lTamanho = 20
    sRetorno = String(lTamanho, 0)
    
'    Call GetPrivateProfileString(APLICACAO_ECF, "DataUltimoArquivo", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    If sRetorno <> String(lTamanho, 0) Then dtDataUltArq = StrParaDate(sRetorno)

    If giDebug = 1 Then MsgBox ("34")

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
        Set objTela = ECF

        lErro = Verifica_Arquivo_TEF_PAYGO(objTela)
        If lErro <> SUCESSO And lErro <> 53 And lErro <> 214580 Then gError 109781
        
'        lErro = Verifica_Arquivo_Loja(dtDataUltArq)
'            lErro = CF_ECF("Verifica_Arquivo_Loja")
'            If lErro <> SUCESSO And lErro <> 53 Then gError 109782
    
        If Not AFRAC_ImpressoraCFe(giCodModeloECF) Then
    
            lTamanho = 10
            sRetorno = String(lTamanho, 0)
            Call GetPrivateProfileString(APLICACAO_ECF, "CupomAberto", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
            
            If giDebug = 1 Then MsgBox ("51")
            
            'se o cupom não foi encerrado
            If CInt(sRetorno) <> 0 Then Call CancelaCupom
        
            If giDebug = 1 Then MsgBox ("52")
        
            If giStatusCaixa = STATUS_CAIXA_FECHADO And giStatusSessao <> SESSAO_ENCERRADA Then
            
                'Função que Faz o Encerramento da Sessão
                lErro = CF_ECF("Operacoes_Sessao_Executa_Encerramento")
                If lErro <> SUCESSO Then gError 214013
    
            End If
        
            lErro = Verifica_Caixa_Fechado
            If lErro <> SUCESSO Then gError 112620
            
        Else

            lErro = MovimentoCaixa_DataUltMovto(dtUltimoMovto)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If dtUltimoMovto < Date Then
            
                If dtUltimoMovto <> DATA_NULA Then
                
                    'fechar o dia anterior
                    lErro = CF_ECF("Caixa_Executa_Fechamento")
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                End If
                
                'abrir o novo
                lErro = CF_ECF("Caixa_Executa_Abertura")
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            End If
                    
            gdtDataAnterior = Date
    
        End If
    
    End If
    
    If giDebug = 1 Then MsgBox ("82")
    
    'Atualiza o arquivo
    lErro = WritePrivateProfileString(APLICACAO_ECF, "CupomAberto", "0", NOME_ARQUIVO_CAIXA)
    If lErro = 0 Then gError 117536
    
    If giDebug = 1 Then MsgBox ("83")
    
    Call Status_Caixa
        
    If giDebug = 1 Then MsgBox ("84")
        
    If Not bBackup Then
        lErro = CF_ECF("Inicializa_Sessao", AFRAC_ImpressoraCFe(giCodModeloECF))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    Else
        lErro = CF_ECF("Backup_Executa_Direto", 1)
        gError 99870
    End If
        
    If giDebug = 1 Then MsgBox ("85")
            
    gobjLojaECF.dtData = Date
            
    If gobjLojaECF.lIntervaloTrans > 0 Then

        Set objObject = gobjLojaECF
        If giDebug = 1 Then MsgBox ("86")

        lErro = CF_ECF("Rotina_FTP_Envio_Caixa", objObject)
        If lErro <> SUCESSO Then gError 133390

        If giDebug = 1 Then MsgBox ("87")

    End If
            
'    lErro = AFRAC_RZPendenteDiaAnterior(iRZPendenteDiaAnterior)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Reducao Z Pendente Dia Anteior")
'        If lErro <> SUCESSO Then gError 210919
'    End If

    lErro = AFRAC_RZPendente(iRZPendente)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Reducao Z Pendente")
        If lErro <> SUCESSO Then gError 214009
    End If
            
    If giDebug = 1 Then MsgBox ("88")
               
    If giCodModeloECF = IMPRESSORA_NFCE And gobjNFeInfo.iEmContingencia <> 0 And gsUF = "SP" Then
        giCodModeloECF = IMPRESSORA_SAT_2_5_15
    End If
    
    Exit Sub
    
Erro_Main:

    If bSplashFormLoaded Then Unload frmSplashECF
    
    Select Case gErr
            
        Case 53, 99597, 99870, 99830, 99898, 99907, 99908, 99935, 105839, 109537, 109538, 109781, 109782, 112620, 117538, 117543, 126112, 126113, 117683, 133390, 204860, 207967, 210196, 210241, 214013, 214515, ERRO_SEM_MENSAGEM
        
        Case 117536
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "CupomAberto", NOME_ARQUIVO_CAIXA)
        
        Case 126133
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_FECHADO_SESSAO_NAO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165187)

    End Select
    
    Unload ECF
    
    Exit Sub

End Sub

Private Sub CancelaCupom()
    
Dim lErro As Long
Dim iCodigo As Integer
Dim lNumItens As Long
Dim iIndice As Integer
Dim objItens As New ClassItemCupomFiscal
Dim objAliquota As New ClassAliquotaICMS
Dim objVenda As New ClassVenda
Dim sRetorno As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CancelaCupom
    
    'COO atual
    lErro = AFRAC_LerInformacaoImpressora("023", sRetorno)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 112061
    End If
        
    'cancelar o Cupom de Venda
    lErro = AFRAC_CancelarCupom(Nothing, Nothing)
    If lErro <> SUCESSO Then
        Call CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Cancelar Cupom")
        If lErro <> SUCESSO Then gError 99610
    End If
    For iIndice = gcolVendas.Count To 1 Step -1
        Set objVenda = gcolVendas.Item(iIndice)
        If objVenda.iTipo = OPTION_CF Then
            'se o último númeor de cupom é o da última venda executada
            If sRetorno = FormataCpoNum(objVenda.objCupomFiscal.lNumero, 6) Then
                lErro = Alteracoes_CancelamentoCupom(objVenda)
                If lErro <> SUCESSO Then gError 112078
                gcolVendas.Remove (iIndice)
                Exit For
            End If
         End If
    Next
    
    Exit Sub

Erro_CancelaCupom:

    Select Case gErr
                
        Case 99610, 112078, 112061
                    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165188)

    End Select
    
    Exit Sub
        
End Sub

Private Function Alteracoes_CancelamentoCupom(objVenda As ClassVenda) As Long

Dim objMovCaixa As ClassMovimentoCaixa
Dim objCheque As ClassChequePre
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim objCarne As ClassCarne
Dim lErro As Long
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim lSequencial As Long
Dim objAliquota As New ClassAliquotaICMS
Dim objItens As ClassItemCupomFiscal
Dim iIndice1 As Integer
Dim sLog As String
Dim colRegistro As New Collection

On Error GoTo Erro_Alteracoes_CancelamentoCupom

    For Each objItens In objVenda.objCupomFiscal.colItens
        For Each objAliquota In gcolAliquotasTotal
            If objItens.dAliquotaICMS = objAliquota.dAliquota Then
                objAliquota.dValorTotalizadoLoja = objAliquota.dValorTotalizadoLoja - ((objItens.dPrecoUnitario * objItens.dQuantidade) * objAliquota.dAliquota)
                Exit For
            End If
        Next
    Next
    
    For iIndice = gcolMovimentosCaixa.Count To 1 Step -1
        Set objMovCaixa = gcolMovimentosCaixa.Item(iIndice)
        If objMovCaixa.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada movimento da venda
    For Each objMovCaixa In objVenda.colMovimentosCaixa
        '??? 24/08/2016 If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro - objMovCaixa.dValor
        'Se for de cartao de crédito ou débito especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO Or objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto <> 0 Then
            'Busca em gcolCartão a ocorrencia de Cartão nao especificado
            For iIndice = gcolCartao.Count To 1 Step -1
                Set objAdmMeioPagtoCondPagto = gcolCartao.Item(iIndice)
                'Se encontrou
                If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iParcelamento = objMovCaixa.iParcelamento And objAdmMeioPagtoCondPagto.iTipoCartao = objMovCaixa.iTipoCartao Then
                    'Atualiza o saldo do cartão
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolCartao.Remove (iIndice)
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de crédito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAOCREDITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CDEBITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
        'Se o omvimento for de cartão de débito não especificado
        If (objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_CARTAODEBITO) And objMovCaixa.iAdmMeioPagto = 0 Then
            'inclui na col como não especificado
            For Each objAdmMeioPagtoCondPagto In gcolCartao
                If objAdmMeioPagtoCondPagto.sNomeAdmMeioPagto = STRING_NAO_DETALHADO_CCREDITO Then
                    'Atualiza o saldo de não especificado
                    objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    Exit For
                End If
            Next
        End If
    Next
    
    'Para cada movimento
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o movimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for um recebimento em ticket
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada obj de ticket da coleção global de tickets
                For Each objAdmMeioPagtoCondPagto In gcolTicket
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo de não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Ticket da coleção global
                For iIndice1 = gcolTicket.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolTicket.Item(iIndice1)
                    'Se encontrou o ticket/parcelamento
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolTicket.Remove (iIndice1)
                        'Sinaliza que encontrou
                        Exit For
                    End If
                Next
            End If
        End If
    Next
    
    Set objAdmMeioPagtoCondPagto = New ClassAdmMeioPagtoCondPagto
    
    'Verifica se já existe movimentos de Outros\
    'Para cada MOvimento de Outros
    For iIndice = objVenda.colMovimentosCaixa.Count To 1 Step -1
        'Pega o MOvimento
        Set objMovCaixa = objVenda.colMovimentosCaixa.Item(iIndice)
        'Se for do tipo outros
        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_OUTROS Then
            'Se for não especificado
            If objMovCaixa.iAdmMeioPagto = 0 Then
                'Para cada pagamento em outros na coleção global
                For Each objAdmMeioPagtoCondPagto In gcolOutros
                    'Se for o não especificado
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = 0 Then
                        'Atualiza o saldo não especificado
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                    End If
                Next
            'Se for especificado
            Else
                'Para cada Pagamento em outros na col global
                For iIndice1 = gcolOutros.Count To 1 Step -1
                    Set objAdmMeioPagtoCondPagto = gcolOutros.Item(iIndice1)
                    'Se for do mesmo tipo que o atual
                    If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovCaixa.iAdmMeioPagto Then
                        'Atualiza o saldo
                        objAdmMeioPagtoCondPagto.dSaldo = objAdmMeioPagtoCondPagto.dSaldo - objMovCaixa.dValor
                        If objAdmMeioPagtoCondPagto.dSaldo = 0 Then gcolOutros.Remove (iIndice1)
                        Exit For
                    End If
                Next
            End If
        End If
    Next
        
    'remove o Carne na col global
    If objVenda.objCarne.colParcelas.Count > 0 Then
        For iIndice = 1 To gcolCarne.Count
            Set objCarne = gcolCarne.Item(iIndice)
            If objCarne.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCarne.Remove (iIndice)
        Next
    End If
    
    'remove o Cheque na col global
    If objVenda.colCheques.Count > 0 Then
        For iIndice = gcolCheque.Count To 1 Step -1
            Set objCheque = gcolCheque.Item(iIndice)
            If objCheque.lCupomFiscal = objVenda.objCupomFiscal.lNumero Then gcolCheque.Remove (iIndice)
        Next
    End If
    
    'Abre a Transação
    lErro = CF_ECF("Caixa_Transacao_Abrir", lSequencial)
    If lErro <> SUCESSO Then gError 99952

    'Joga na string para gravar
    sLog = TIPOREGISTROECF_EXCLUSAOCUPOM & Chr(vbKeyControl) & objVenda.objCupomFiscal.lNumero & Chr(vbKeyEnd)
    
    colRegistro.Add sLog
    
    'Grava no arquivo
    lErro = CF_ECF("MovimentoCaixaECF_Grava", colRegistro)
    If lErro <> SUCESSO Then gError 99901
    
    Set colRegistro = New Collection
    
    lSequencial = glSeqTransacaoAberta
    
    'Fecha a Transação
    lErro = CF_ECF("Caixa_Transacao_Fechar", lSequencial)
    If lErro <> SUCESSO Then gError 99953
    
    Alteracoes_CancelamentoCupom = SUCESSO
    
    Exit Function
    
Erro_Alteracoes_CancelamentoCupom:
    
    Alteracoes_CancelamentoCupom = gErr
    
    Select Case gErr
    
        Case 99901, 99953, 99952
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 165189)

    End Select
        
    lSequencial = glSeqTransacaoAberta
            
    'Rollback na transação
    lErro = CF_ECF("Caixa_Transacao_Rollback", glTransacaoPAFECF)
        
    Exit Function
    
End Function

Private Function Verifica_Caixa_Fechado() As Long
    
Dim lTamanho As Long
Dim sRetorno As String
Dim iStatus As Integer
Dim dtDataUltimaReducaoZCaixaConfig As Date
Dim lErro As Long
Dim iNaoRedZ As Integer
Dim objVenda As ClassVenda
Dim colOrcamento As New Collection
Dim lNumero As Long
'Dim iRZPendenteDiaAnterior As Integer
Dim dtUltimaReducaoECF As Date
Dim sData As String
Dim sHora As String
Dim iRZPendente As Integer
Dim dtData As Date
Dim sDataHora As String
Dim dtDataECF As Date
Dim dtHoraECF As Date
Dim dtDataUltimoMovimento As Date

On Error GoTo Erro_Verifica_Caixa_Fechado
        
        
    If giDebug = 1 Then MsgBox ("57")
        
    lTamanho = 10
    sRetorno = String(lTamanho, 0)
    Call GetPrivateProfileString(APLICACAO_CAIXA, "StatusCaixa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
    If sRetorno <> String(lTamanho, 0) Then iStatus = StrParaInt(sRetorno)
    
    If giDebug = 1 Then MsgBox ("Verifica_Caixa_Fechado - sRetorno = " & sRetorno)
    
    If giDebug = 1 Then MsgBox ("Verifica_Caixa_Fechado - iSatus = " & CStr(iStatus))
    
    'data do movimento da ultima reducao Z, nao necessariamente a data em que foi feita.
    lErro = AFRAC_DataReducao(dtUltimaReducaoECF)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao Z")
    If lErro <> SUCESSO Then gError 214010
            
    If giDebug = 1 Then MsgBox ("AFRAC_DataReduca")
            
    If giDebug = 1 Then MsgBox ("dtUltimaReducaoECF = " & CStr(dtUltimaReducaoECF))
            
            
    'le a data/hora do ecf
    lErro = AFRAC_LerInformacaoImpressora("017", sDataHora)
    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
    If lErro <> SUCESSO Then gError 207235

    If giDebug = 1 Then MsgBox ("AFRAC_LerInformacaoImpressora 17")

    dtDataECF = StrParaDate(left(sDataHora, 2) & "/" & Mid(sDataHora, 3, 2) & "/" & Mid(sDataHora, 5, 2))
    
    dtHoraECF = StrParaDate(Mid(sDataHora, 7, 2) & ":" & Mid(sDataHora, 9, 2) & ":" & right(sDataHora, 2))
    
    If giDebug = 1 Then MsgBox ("AFRAC_LerInformacaoImpressora 17 dtDataECF = " & CStr(dtDataECF) & " HoraECF = " & CStr(dtHoraECF))
    
    lErro = AFRAC_DataMovimentoProximaReducao(sData)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "DataMovimentoProximaReducao")
        If lErro <> SUCESSO Then gError 214024
    End If

    If giDebug = 1 Then MsgBox ("AFRAC_DataMovimentoProximaReucao")


    dtDataUltimoMovimento = StrParaDate(left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4))
    
    If giDebug = 1 Then MsgBox ("dtDataUltimoMovimento = " & CStr(dtDataUltimoMovimento))
    
    lErro = AFRAC_RZPendente(iRZPendente)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Reducao Z Pendente")
        If lErro <> SUCESSO Then gError 214023
    End If
    
     If giDebug = 1 Then MsgBox ("AFRAC_RZPendente iRZPendente = " & CStr(iRZPendente))
   
    If gdtDataAnterior = DATA_NULA Then
    
        If iRZPendente = 1 Then
    
            gdtDataAnterior = dtDataUltimoMovimento
        
        Else
        
            gdtDataAnterior = dtDataECF
        
        End If
    
    End If
    
    If giDebug = 1 Then MsgBox ("gdtUltimaReducao = " & CStr(gdtUltimaReducao))
    
    'se o caixa está aberto e a data da ultima reducao registrada pelo sistema é menor q a data da ultima reducao do ECF--> fecha o caixa
    'ou se o caixa esta aberto e a data do ultimo movimento é menor q a data do ecf
    If iStatus <> CAIXA_STATUS_FECHADO And (gdtUltimaReducao < dtUltimaReducaoECF Or (dtDataUltimoMovimento <> DATA_NULA And (dtDataECF > dtDataUltimoMovimento + 1 Or (dtDataECF = dtDataUltimoMovimento + 1 And Hour(dtHoraECF) > 2)))) Then
'    If iRZPendente = 1 And iStatus <> CAIXA_STATUS_FECHADO And (gdtUltimaReducao < dtUltimaReducaoECF Or ((dtDataECF > dtDataUltimoMovimento + 1 Or (dtDataECF = dtDataUltimoMovimento + 1 And Hour(dtHoraECF) > 2)))) Then
        
'      (gdtUltimaReducao < dtUltimaReducaoECF Then
        
        'se nao for a Elgin, pois a reducao é feita automaticamente
'        If giCodModeloECF <> 7 Then
'            iNaoRedZ = NAO_FAZ_REDUCAOZ
'        Else
        
        If giDebug = 1 Then MsgBox ("58")
        
        'ele fecha o caixa mas nao comanda a impressao da reducao Z pois o ECF automaticamente imprime
        lErro = CF_ECF("Caixa_Executa_Fechamento")
        If lErro <> SUCESSO Then gError 112623
        
        If giDebug = 1 Then MsgBox ("79")
        
    End If
    
    
    lErro = AFRAC_RZPendente(iRZPendente)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Reducao Z Pendente")
        If lErro <> SUCESSO Then gError 214023
    End If
    
    If iRZPendente = 1 Then

        gdtDataAnterior = dtDataUltimoMovimento
    
    Else
    
        gdtDataAnterior = dtDataECF
    
    End If
    
    If giDebug = 1 Then MsgBox ("80")

'    'Função Que le os orcamentos
'    lErro = CF_ECF("OrcamentoECF_Le", colOrcamento)
'    If lErro <> SUCESSO Then gError 204221
'
'    'se  a data do orcamento é de DataAtual-2  ==>
'    'imprime o cupom como cancelado - PAFECF
'    For Each objVenda In colOrcamento
'
'        If objVenda.objCupomFiscal.dtDataEmissao + 2 <= Date Then
'
'            lErro = CF_ECF("Imprime_Orcamento", objVenda)
'            If lErro <> SUCESSO Then gError 204219
'        End If
'
'    Next
    
    
    
    If giDebug = 1 Then MsgBox ("81")
    
    Verifica_Caixa_Fechado = SUCESSO
    
    Exit Function
    
Erro_Verifica_Caixa_Fechado:
    
    Verifica_Caixa_Fechado = gErr
    
    Select Case gErr
            
        Case 112623, 204219, 204221, 210436, 214023, 214024
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165190)

    End Select
    
End Function

Private Function MovimentoCaixa_DataUltMovto(dtUltimoMovto As Date) As Long

Dim lErro As Long
Dim dtData1 As Date
Dim lComando As Long

On Error GoTo Erro_MovimentoCaixa_DataUltMovto

    dtUltimoMovto = DATA_NULA
    
    lComando = Comando_AbrirExt(glConexaoPAFECF)
    If lComando = 0 Then gError 210425

    lErro = Comando_Executar(lComando, "SELECT Data FROM MovimentoCaixa WHERE Msg LIKE '24_88%' ORDER BY Data DESC ", dtData1)
    If lErro <> AD_SQL_SUCESSO Then gError 210426
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210427
    
    If lErro = AD_SQL_SUCESSO Then dtUltimoMovto = dtData1
    
    Call Comando_Fechar(lComando)
    
    MovimentoCaixa_DataUltMovto = SUCESSO
    
    Exit Function
    
Erro_MovimentoCaixa_DataUltMovto:

    MovimentoCaixa_DataUltMovto = gErr

    Select Case gErr

        Case 210425
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 210426 To 210430
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MOVIMENTOCAIXA, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165191)

    End Select

    Call Comando_Fechar(lComando)

End Function

'Function Verifica_Arquivo_Loja() As Long
''faz a leitura do ultimo arquivo, mesmo que o dia tenha mudado para cuidar do caixa (suprimento) que ainda sobrou
'
'Dim lErro As Long
'Dim sTipo As String
'Dim iPos As Integer
'Dim iPosEsc As Integer
'Dim iIndice As Integer
'Dim iIndice1 As Integer
'Dim sNomeArq As String
'Dim iPosInicio As Integer
'Dim bAchou As Boolean
'Dim bAchou1 As Boolean
'Dim iPosInicio1 As Integer
'Dim bAchouAbert As Boolean
'Dim colRegistroTemp As New Collection
'Dim sSeqAbert As String
'Dim sSeqFech As String
'Dim sNum As String
'Dim sRegistro As String
'Dim colRegistro As New Collection
'Dim iRegistro As Integer
'Dim iNum As Integer
'Dim iSeqArq As Integer
'Dim lTamanho As Long
'Dim sRetorno As String
'Dim sNomeArq1 As String
'Dim sNomeArq2 As String
'Dim sRet As String
'Dim iCodEmpresa As Integer
'Dim sArquivo As String
'Dim colVendas As New Collection
'Dim objVenda As New ClassVenda
'Dim objVendedor As ClassVendedor
'Dim objProduto As ClassProduto
'Dim objItens As ClassItemCupomFiscal
'Dim sNomeArq0 As String
'Dim iStatus As Integer
'Dim bNaoPulaCodigo As Boolean
'Dim vbMsg As VbMsgBoxResult
'Dim dtData As Date
'Dim dtData1 As Date
'Dim alComando(1 To 3) As Long
'Dim lNumMovto As Long
'Dim iTipo As Integer
'Dim dHora As Double
'Dim iSeq As Integer
'Dim sMsg As String
'Dim lErro1 As Long
'
'On Error GoTo Erro_Verifica_Arquivo_Loja
'
'    bNaoPulaCodigo = True
'    bAchouAbert = True
'    bAchou = False
'
'    iPosInicio = 1
'
'    If giDebug = 1 Then MsgBox ("46")
'
'    'Verifica se tem alguma venda de Loja em aberta
''    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(dtData, "ddmmyy")) & (".txt")
'
'    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
''    sNomeArq0 = sNomeArq
'
''    sRet = Dir(sNomeArq)
'
''    If Len(sRet) = 0 Then
''
''        vbMsg = Rotina_AvisoECF(vbYesNo, AVISO_ARQUIVO_NAO_ENCONTRADO, sNomeArq)
''
''        If vbMsg = vbNo Then
''            gError 105693
''        Else
''            bNaoPulaCodigo = False
''        End If
''    End If
''
''    If giDebug = 1 Then MsgBox ("47")
''
''    If bNaoPulaCodigo Then
''
''        Open sNomeArq For Input As #1
''
''        iIndice = 0
''
''        Do While Not EOF(1)
'
'   'Abre os  comandos
'    For iIndice = LBound(alComando) To UBound(alComando)
'        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
'        If alComando(iIndice) = 0 Then gError 210425
'    Next
'
'    sRegistro = String(STRING_MOVIMENTOCAIXA_MSG, 0)
'
'    lErro = Comando_Executar(alComando(1), "SELECT Data  FROM MovimentoCaixa  ORDER BY Data DESC ", dtData1)
'    If lErro <> AD_SQL_SUCESSO Then gError 210426
'
'    lErro = Comando_BuscarPrimeiro(alComando(1))
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 210427
'
'    If lErro = AD_SQL_SUCESSO Then
'
'        sMsg = String(STRING_MOVIMENTOCAIXA_MSG, 0)
'
'        lErro = Comando_Executar(alComando(2), "SELECT NumMovto, Data, Tipo, Msg, Hora, Seq  FROM MovimentoCaixa  WHERE Data = ? ORDER BY NumMovto, Seq  ", lNumMovto, dtData, iTipo, sMsg, dHora, iSeq, dtData1)
'        If lErro <> AD_SQL_SUCESSO Then gError 210428
'
'        lErro1 = Comando_BuscarPrimeiro(alComando(2))
'        If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 210429
'
'
'        Do While lErro1 = AD_SQL_SUCESSO
'
'            colRegistro.Add sMsg
'
'            lErro1 = Comando_BuscarProximo(alComando(2))
'            If lErro1 <> AD_SQL_SUCESSO And lErro1 <> AD_SQL_SEM_DADOS Then gError 210430
'
'            If iSeq = 1 Or lErro1 = AD_SQL_SEM_DADOS Then
'
'                lErro = CF_ECF("Registro_ECF_CC", colRegistro, sArquivo, iRegistro, iIndice, colVendas)
'                If lErro <> SUCESSO Then gError 110002
'
'                Set colRegistro = New Collection
'
'            End If
'
'
''            If sRegistro <> "" Then
''
''                iIndice = iIndice + 1
''
''                'Procura o Primeiro Control para saber o tipo do registro
''                iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
''
''                sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
''
''                Select Case sTipo
''
''                    'verifica se é de abertura
''                    Case TIPOREGISTROECF_ABRESEQ
''                        'se encontrei um sequencial aberto antes do encerramento do anterior-->desprezo o anterior
''                        If bAchouAbert Then Set colRegistroTemp = New Collection
''                        'Recolhe o sequencial
''                        sSeqAbert = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos + 1))
''                        bAchouAbert = True
''                            'Procura o end
''                        iPosEsc = InStr(iPos, sRegistro, Chr(vbKeyEnd))
''                        'achando o último sequencial de abertura ...
''                        sNum = Mid(sRegistro, iPos + 1, iPosEsc - (iPos + 1))
''                    'verifica se é de fechamento
''                    Case TIPOREGISTROECF_ENCERRASEQ
''                        'Recolhe o sequencial
''                        sSeqFech = Mid(sRegistro, iPos + 1, Len(sRegistro) - (iPos + 1))
''                        If sSeqAbert = sSeqFech Then
''                         bAchou = True
''
''                    Case Else
''                        colRegistroTemp.Add sMsg
''
''                End Select
'
''                If bAchou Then
''
''                    For iIndice1 = 1 To colRegistroTemp.Count
''
''                        sRegistro = colRegistroTemp.Item(iIndice1)
''
''                        iRegistro = iRegistro + 1
''
''                        colRegistro.Add sRegistro
''
''                        If Mid(sRegistro, Len(sRegistro), 1) = Chr(vbKeyEnd) Then
''
''                            lErro = CF_ECF("Registro_ECF_CC", colRegistro, sArquivo, iRegistro, iIndice, colVendas)
''                            If lErro <> SUCESSO Then gError 110002
''
''                            Set colRegistro = New Collection
''                            iRegistro = 0
''                        End If
''                    Next
''                    Set colRegistroTemp = New Collection
''                    bAchouAbert = False
''                    bAchou = False
''                'End If
''            End If
''
''            colRegistroTemp.Add sMsg
'
'
'        Loop
'
'
'
'        If giDebug = 1 Then MsgBox ("48")
'
''        Close 1
'
'        'Primeira Posição
'        iPosInicio = 1
'
'    '    If iIndice > 0 Then
'    '        'Procura o Primeiro Control para saber o tipo do registro
'    '        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
'    '
'    '        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
'    '
'    '        'se o último registro não for um encerramento ou um rollback ou um registro de arquivo
'    '        If sTipo <> TIPOREGISTROECF_ENCERRASEQ And sTipo <> TIPOREGISTROECF_ARQUIVO Then
'    '            'grava
'    '            lErro = CF_ECF("Caixa_Transacao_Rollback", StrParaLong(sNum))
'    '            If lErro <> SUCESSO Then gError 109817
'    '        'se for do tipo arquivo...
'    '        ElseIf sTipo = TIPOREGISTROECF_ARQUIVO Then
'    '            lTamanho = 15
'    '            sRetorno = String(lTamanho, 0)
'    '
'    '            'pega o último sequencial do arquivo
'    '            Call GetPrivateProfileString(APLICACAO_ARQUIVO, "SeqArquivo", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    '            If sRetorno <> String(lTamanho, 0) Then iSeqArq = StrParaInt(sRetorno)
'    '
'    '            iNum = StrParaInt(Mid(sRegistro, iPos + 1, Len(sRegistro) - iPos - 1))
'    '            'se não chegou a atualizar o .ini --> atualiza
'    '            If iNum <> iSeqArq Then
'    '                'pega o código da empresa
'    '                sRetorno = String(lTamanho, 0)
'    '                Call GetPrivateProfileString(APLICACAO_EMPRESA, "CodEmpresa", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
'    '                If sRetorno <> String(lTamanho, 0) Then iCodEmpresa = StrParaInt(sRetorno)
'    '
'    '                'verifica se o arquivo foi renomeado...
'    '                sNomeArq1 = iCodEmpresa & "_" & giFilialEmpresa & "_" & giCodCaixa & "_" & iNum
'    '                sNomeArq2 = sNomeArq1 & ".tmp"
'    '                'Verifica o arquivo intpos.001
'    '
'    '                'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    '                sNomeArq0 = sNomeArq2
'    '
'    '                sRet = Dir(sNomeArq2, vbNormal)
'    '                'se encontrou-->renomeia
'    '                If sRet <> "" Then
'    '                    sNomeArq1 = sNomeArq1 & ".ccc"
'    '
'    '                    'esta atribuicao é feita pois caso haja algum erro será mostrado o nome contido em sNomeArq0
'    '                    sNomeArq0 = sNomeArq1
'    '
'    '                    'renomeando os arquivos
'    '                    Name sNomeArq2 As sNomeArq1
'    '                End If
'    '
'    '                'Atualiza o sequencial de arquivo
'    '                Call WritePrivateProfileString(APLICACAO_ARQUIVO, "SeqArquivo", CStr(iNum), NOME_ARQUIVO_CAIXA)
'    '
'    '            End If
'    '        End If
'    '    End If
'
'        For Each objVenda In colVendas
'
'            bAchou = False
'            'se for um orçamneto
'            If objVenda.objCupomFiscal.iTipo = OPTION_DAV Or objVenda.objCupomFiscal.iTipo = OPTION_PREVENDA Then
'                'verifica se o vendedor ainda existe-> se naão existir não impacta na memória
'                If objVenda.objCupomFiscal.iVendedor <> 0 Then
'
'                    For Each objVendedor In gcolVendedores
'                        If objVenda.objCupomFiscal.iVendedor = objVendedor.iCodigo Then
'                            bAchou = True
'                            Exit For
'                        End If
'                    Next
'
'                Else
'
'                    bAchou = True
'                End If
'
'                If bAchou Then
'                    'verifica se os produtos ainda existe-> se naão existir não impacta na memória
''                    For Each objItens In objVenda.objCupomFiscal.colItens
''                        bAchou = False
''                        For iIndice = 1 To gaobjProdutosNome.Count
''                            Set objProduto = gaobjProdutosNome.Item(iIndice)
''                            If objItens.sProduto = objProduto.sCodigo Then
''                                objItens.sProdutoNomeRed = objProduto.sNomeReduzido
''                                bAchou = True
''                                Exit For
''                            End If
''                        Next
''                        If Not (bAchou) Then Exit For
''                    Next
'
'
'                    For Each objItens In objVenda.objCupomFiscal.colItens
'                        bAchou = False
'
'                        lErro = CF_ECF("Produtos_Le", objItens.sProduto, objProduto)
'                        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 214854
'
'                        If lErro = SUCESSO Then
'                            objItens.sProdutoNomeRed = objProduto.sNomeReduzido
'                            bAchou = True
'                        Else
'                            Exit For
'
'                        End If
'
'
'                    Next
'
'                End If
'            Else
'                bAchou = True
'            End If
'
''            If gcolVendas.Count = 0 Then Set gcolCheque = New Collection
'
'            If bAchou Then
'                If objVenda.objCupomFiscal.lNumero <> 0 Then
'                    'Atualiza os Movimentos nas coleções globais
'                    Call CF_ECF("Atualiza_Movimentos_Memoria", objVenda)
'                    Call Atualiza_Movimentos(objVenda)
'                ElseIf objVenda.objCupomFiscal.iStatus = STATUS_BAIXADO Then
'                    'Atualiza os Movimentos nas coleções globais
'                    Call CF_ECF("Atualiza_Movimentos_Memoria1", objVenda)
'                    Call Atualiza_Movimentos(objVenda)
'                End If
'
'                'Atribui para a coleção global o objvenda
'                gcolVendas.Add objVenda
'            End If
'        Next
'
'    End If
'
'    If giDebug = 1 Then MsgBox ("49")
'
'    'mario
'    'se foi lido um arquivo de outra data ==> limpa as colecoes
'    If dtData1 <> Date Then
'        Set gcolMovimentosCaixa = New Collection
'        Set gcolVendas = New Collection
'    End If
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'    If giDebug = 1 Then MsgBox ("50")
'
'
'    Verifica_Arquivo_Loja = SUCESSO
'
'    Exit Function
'
'Erro_Verifica_Arquivo_Loja:
'
'    Verifica_Arquivo_Loja = gErr
'
'    Close 1
'
'    Select Case gErr
'
'        Case 53
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO, gErr, sNomeArq0)
'
'        Case 58
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_EXISTENTE, gErr, sNomeArq0)
'
'        Case 105693, 110002, 214854
'
'        Case 210425
'            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)
'
'        Case 210426 To 210430
'            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MOVIMENTOCAIXA, gErr)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165191)
'
'    End Select
'
'    For iIndice = LBound(alComando) To UBound(alComando)
'        Call Comando_Fechar(alComando(iIndice))
'    Next
'
'End Function

'Private Function Verifica_Arquivo_Orc_Loja() As Long
''se for o primeiro ecf da loja a faz a leitura do ultimo arquivo orcamento atual
'
'Dim lErro As Long
'Dim colRegistro As New Collection
'Dim colOrcamento As New Collection
'Dim objArqSeq As New ClassArqSeq
'
'On Error GoTo Erro_Verifica_Arquivo_Orc_Loja
'
'
'    If giDebug = 1 Then MsgBox ("54")
'
'    'le o arquivo ArqSeq para descobrir a ultima data que foi gravado o arquivo de orcamento
'    lErro = CF_ECF("ArqSeq_Le", objArqSeq)
'    If lErro <> SUCESSO Then gError 105877
'
'    If giDebug = 1 Then MsgBox ("55")
'
'    'se o arquivo de orcamento antigo existir  ==>
'    'transfere os dados do arquivo antigo para o atual
'    If objArqSeq.dtData <> Date Then
'
'        'le os registros do orcamento antigo, se houver
'        lErro = CF_ECF("OrcamentoECF_Le_Lock", colRegistro, colOrcamento)
'        If lErro <> SUCESSO Then gError 105850
'
'        'Grava os registros validos no orcamento atual (retirando as exclusoes e os que tiverem a data de validade vencida)
'        lErro = CF_ECF("OrcamentoECF_Grava_Novo", colRegistro)
'        If lErro <> SUCESSO Then gError 105876
'
'        Call CF_ECF("OrcamentoECF_Fechar")
'
'    End If
'
'    If giDebug = 1 Then MsgBox ("56")
'
'    Verifica_Arquivo_Orc_Loja = SUCESSO
'
'    Exit Function
'
'Erro_Verifica_Arquivo_Orc_Loja:
'
'    Verifica_Arquivo_Orc_Loja = gErr
'    lErro = gErr
'
'    Call CF_ECF("OrcamentoECF_Fechar")
'
'    Select Case lErro
'
'        Case 105850, 105876, 105877
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, lErro, Error$, 165192)
'
'    End Select
'
'End Function

Private Function Verifica_Arquivo_TEF(objTela As Object) As Long

Dim sTipo As String
Dim sTipo1 As String
Dim sNum As String
Dim sNum1 As String
Dim iPos As Integer
Dim iPos1 As Integer
Dim iPosEsc As Integer
Dim iPosEsc1 As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim bCont As Boolean
Dim asReg1() As String
Dim asReg() As String
Dim sReg As String
Dim sNomeArq As String
Dim iPosInicio As Integer
Dim bAchou As Boolean
Dim dPag As Double
Dim sRede As String
Dim sNSU As String
Dim sFim As String
Dim dValorTotal As Double
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim bCon As Boolean
Dim lNum As Long
Dim sRet As String
Dim objTiposMeiosPagtos As New ClassTMPLoja
Dim sDescricao As String
Dim iLoop As Integer
Dim lNumero As Long
Dim sTxt As String
Dim vbMsg As VbMsgBoxResult
Dim sRegistro As String
Dim sRegistro1 As String

Dim lErro As Long
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivoINTPOS As String
Dim objFormMsg As Object
Dim objVenda As New ClassVenda

On Error GoTo Erro_Verifica_Arquivo_TEF
        
    If giTEF = CAIXA_ACEITA_TEF Then
        
        If giDebug = 1 Then MsgBox ("35")
        
        'cria o diretorio que vai abrigar o backup do arquivo ARQUIVO_TEF_RESP2
        sRet = Dir(Dir_Tef_Resp2_Backup_Prop, vbDirectory)
            
        'se o diretorio nao existir ==> cria
        If sRet = "" Then MkDir (Dir_Tef_Resp2_Backup_Prop)
        
        If giTipoTEF = TIPOTEF_SITEF Then
            'cria o diretorio que vai abrigar o backup do arquivo de confirmacao de transacao quando o sitef nao esta ativo
            sRet = Dir(Dir_Tef_Req_Backup1_Prop, vbDirectory)
                
            'se o diretorio nao existir ==> cria
            If sRet = "" Then MkDir (Dir_Tef_Req_Backup1_Prop)
        End If
        
        lErro = 1
            
        'fica em loop ate o gerenciador padrao estar ativo
        Do While lErro <> SUCESSO
            
            lErro = CF_ECF("TEF_Gerenciador_Padrao_PAYGO")
            If lErro <> SUCESSO And lErro <> 133756 Then gError 133739
            
        Loop
        
        If giDebug = 1 Then MsgBox ("36")
        
        Set objTela = ECF
        Set objFormMsg = MsgTEF1
        
        lTamanho = 10
        sRetorno = String(lTamanho, 0)
    
        'Indica o status do TEF quando foi interrompido o processo
        Call GetPrivateProfileString(APLICACAO_ECF, "StatusTEF", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
        If sRetorno <> String(lTamanho, 0) Then
        
            sRetorno = StringZ(sRetorno)
            
            If sRetorno = TEF_STATUS_VENDA Or sRetorno = TEF_STATUS_VENDA_GERENCIAL Or sRetorno = TEF_STATUS_VENDA_VINCULADA Then
            
                sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
                    
                If Len(sArquivoINTPOS) = 0 Then
            
                    sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Backup_Prop)
                    
                    If Len(sArquivoINTPOS) <> 0 Then
                    
                        FileCopy Arquivo_Tef_Resp2_Backup_Prop, Arquivo_Tef_Resp2_Prop
            
                    End If
                    
                End If
                
                If giDebug = 1 Then MsgBox ("37")
                
                sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
                
                If Len(sArquivoINTPOS) <> 0 Then
                
'                    lErro = CF_ECF("TEF_NaoConfirma_Transacao", objTela)
'                    If lErro <> SUCESSO Then gError 133708
                
                    lTamanho = 10
                    sRetorno = String(lTamanho, 0)
                
                    'Pega o ultimo COO de multiplos cartoes e veja se existe necessidade de cancelamento
                    Call GetPrivateProfileString(APLICACAO_ECF, "COO", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
                    
                    If sRetorno <> String(lTamanho, 0) Then
                    
                        sRetorno = StringZ(sRetorno)
                    
                        objVenda.objCupomFiscal.lNumero = StrParaLong(sRetorno)
                    
                        If giDebug = 1 Then MsgBox ("38")
                    
                        lErro = CF_ECF("TEF_NaoConfirma_Transacao1", objTela, objVenda)
                        If lErro <> SUCESSO Then gError 126586
                        
                        If giDebug = 1 Then MsgBox ("39")
                        
                        lErro = CF_ECF("TEF_CNF_NCN_Pendente")
                        If lErro <> SUCESSO Then gError 126587
                    
                        
                        If giDebug = 1 Then MsgBox ("40")
                        'cancela os cartoes ja confirmados e nao confirma o ultimo, se houverem
                        lErro = CF_ECF("TEF_CNC", objVenda, objFormMsg, objTela)
                        If lErro <> SUCESSO Then gError 133803
                
                        If giDebug = 1 Then MsgBox ("41")
                
'                        'imprime os comprovantes de cancelamento se houverem
'                        lErro = CF_ECF("TEF_Imprime_CNC", objVenda, objFormMsg, objTela)
'                        If lErro <> SUCESSO Then gError 133804
                    
                    End If
                
                Else
                
                
                    If giDebug = 1 Then MsgBox ("42")
                    
                    lErro = CF_ECF("TEF_NaoConfirma_Transacao", objTela)
                    If lErro <> SUCESSO Then gError 133708
                    
                    If giDebug = 1 Then MsgBox ("43")
                    
                    lErro = CF_ECF("TEF_Libera_Impressora", objTela)
                    If lErro <> SUCESSO Then gError 133707
                    
                    If giDebug = 1 Then MsgBox ("44")
                    
                    'Atualiza o arquivo
                    lErro = WritePrivateProfileString(APLICACAO_ECF, "StatusTEF", "", NOME_ARQUIVO_CAIXA)
                    If lErro = 0 Then gError 133709
                
                End If
                
            End If
                
        End If
            
    End If
    
    
    If giDebug = 1 Then MsgBox ("45")
    
    Verifica_Arquivo_TEF = SUCESSO
    
    Exit Function
    
Erro_Verifica_Arquivo_TEF:
    
    Verifica_Arquivo_TEF = gErr
    
    Select Case gErr
    
        Case 126586, 126587, 133704, 133706, 133708, 133739, 133802, 133803, 133804
        
        Case 133705, 133707, 133709
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "StatusTEF", NOME_ARQUIVO_CAIXA)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165193)

    End Select
    
    
'    Close 1
'
'    Select Case gErr
'
'        Case 109699, 53, 99821, 109412
'
'        Case 99803
'            Call Rotina_ErroECF(vbOKOnly, ERRO_IMPRESSORA_NAO_RESPONDE, gErr)
'
'        Case 109566, 112395, 99803, 109413
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165194)
'
'    End Select
    
    Exit Function
    
End Function

Private Sub Status_Caixa()

Dim objOperador As ClassOperador
Dim sStatusSessao As String
Dim sStatusCaixa As String
    
    'Verifica o Status do Caixa
    If giStatusCaixa = 0 Then
        sStatusCaixa = "Caixa Fechado"
    ElseIf giStatusCaixa = 1 Then
        sStatusCaixa = "Caixa Aberto"
    End If
    
    
    
    
    'Verifica o Status da Sessão
    If giStatusSessao = SESSAO_ENCERRADA Then
        sStatusSessao = "Sessao Fechada"
    Else
        If giStatusSessao = SESSAO_ABERTA Then

            'Função que Executa a Suspenção da Sessão
            Call CF_ECF("Sessao_Executa_Suspensao")
        
            sStatusSessao = "Sessao Suspensa"
        
        Else
            sStatusSessao = "Sessao Suspensa"
        End If
    End If
    
    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then

        GL_objMDIForm.Caption = "SESSAO_STATUS : " & sStatusSessao
    
    Else
    
        GL_objMDIForm.Caption = "CAIXA_STATUS : " & sStatusCaixa & " / SESSAO_STATUS : " & sStatusSessao
        
    End If
    
    For Each objOperador In gcolOperadores
        If objOperador.iCodigo = giCodOperador Then
            'Indica o Nome do Operador dessa Sessão.
            GL_objMDIForm.Caption = GL_objMDIForm.Caption & " / OPERADOR : " & objOperador.sNome
            Exit For
        End If
    Next

End Sub
   
Private Function Inicializa_ECF() As Long

Dim lErro As Long
Dim lErro1 As Long
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iMetodo As Integer
Dim sRetorno As String
Dim sRetorno1 As String
Dim sMarcaECF As String
Dim sTipoECF As String
Dim sModeloECF As String
Dim sMFAdicional As String
Dim sData As String
Dim sHora As String
Dim dtDataHoraSistema As Date
Dim dtDataHoraECF As Date
Dim dtDataECF As Date
Dim dtHoraECF As Date
Dim objLojaArqFisMestre As New ClassLojaArqFisMestre
Dim colLojaArqFisAnalitico As New Collection
Dim objUltimaReducao As New ClassUltimaReducao
Dim lCOOUltRZ As Long
Dim lTamanho As Long
Dim sCOOFinal1 As String
Dim dtUltimaReducaoECF As Date

On Error GoTo Erro_inicializa_ECF

    If giDebug = 1 Then MsgBox ("18")

    lErro = AFRAC_AbrePortaSerial()
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abre Porta Serial")
        If lErro <> SUCESSO Then Call AFRAC_FechaPortaSerial
        If lErro <> SUCESSO Then gError 126793
    End If
        
    If giDebug = 1 Then MsgBox ("19")
        
    'verifica se a impressora esta ligada
    lErro = AFRAC_VerificaImpressoraLigada()
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 126792
    End If
        
    If giDebug = 1 Then MsgBox ("20")
        
    'nº de série no Afrac
    sRetorno = String(15, 0)
    lErro = AFRAC_LerInformacaoImpressora("002", sRetorno)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
        If lErro <> SUCESSO Then gError 112060
    End If
    
    If InStr(sRetorno, Chr(0)) <> 0 Then
        gsNumSerie = Mid(sRetorno, 1, InStr(sRetorno, Chr(0)) - 1)
    Else
        gsNumSerie = sRetorno
    End If
    
    If giDebug = 1 Then MsgBox ("21")
    
    
    
    lErro = AFRAC_ConfigurarLinhasEntreCupons(gobjLojaECF.iLinhasEntreCupons)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Configurar Linhas Entre Cupons")
        If lErro <> SUCESSO Then gError 99876
    End If
    
    If giDebug = 1 Then MsgBox ("25")
    
    
    lErro = AFRAC_ConfigurarEspacoEntreLinhas(gobjLojaECF.lEspacoEntreLinhas)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Configurar Linhas Entre Cupons")
        If lErro <> SUCESSO Then gError 210911
    End If
    
    If giDebug = 1 Then MsgBox ("26")
    
    
    For Each objTiposMeiosPagtos In gcolTiposMeiosPagtos
        lErro = AFRAC_ProgramarFormasDePagamento(objTiposMeiosPagtos.iTipo, objTiposMeiosPagtos.sDescricao, False)
        If lErro <> SUCESSO Then
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "")
            If lErro <> SUCESSO Then gError 99878
        End If
    Next
            
'    sRetorno = ""
'    lErro = AFRAC_InformarRazaoSocial(sRetorno)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "")
'        If lErro <> SUCESSO Then gError 99914
'    End If
'    If Len(Trim(sRetorno)) > 0 Then gsNomeEmpresa = sRetorno
    
    
'    sRetorno = ""
'       lErro = AFRAC_InformarCNPJIE(sRetorno, sRetorno1)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "")
'        If lErro <> SUCESSO Then gError 99915
'    End If
'
'    If Len(Trim(sRetorno)) > 0 Then
'        gsCNPJ = Trim(sRetorno)
'        gsInscricaoEstadual = Trim(sRetorno1)
'    End If
    
    If giDebug = 1 Then MsgBox ("28")
    
    
    
'    sRetorno = ""
'    lErro = AFRAC_InformarEndereco(sRetorno)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "")
'        If lErro <> SUCESSO Then gError 99917
'    End If
'    If Len(Trim(sRetorno)) > 0 Then gsEndereco = sRetorno
    
    lErro = AFRAC_InformarMensagemCupom(gsMensagemCupom)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "")
        If lErro <> SUCESSO Then gError 99918
    End If
    
    If giDebug = 1 Then MsgBox ("29")
    
'    sRetorno = String(840, 0)
'    lErro = AFRAC_LerTotalizadoresNSICMS(sRetorno)
'    lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Totalizadores")
'    If lErro <> SUCESSO Then gError 112416
'
    
    If giDebug = 1 Then MsgBox ("30.1")
    
    lErro = AFRAC_RetornaTipoECF(sTipoECF)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna MarcaECF")
        If lErro <> SUCESSO Then gError 204372
    End If
    
    gsTipoECF = Trim(sTipoECF)
    
    If giDebug = 1 Then MsgBox ("30.2")
    
    lErro = AFRAC_RetornaMarcaECF(sMarcaECF)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna MarcaECF")
        If lErro <> SUCESSO Then gError 204373
    End If
    
    gsMarcaECF = Trim(sMarcaECF)
    
    If giDebug = 1 Then MsgBox ("30.3")
    
    lErro = AFRAC_RetornaModeloECF(sModeloECF)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Retorna ModeloECF")
        If lErro <> SUCESSO Then gError 204374
    End If
    
    gsModeloECF = Trim(sModeloECF)
    
    
    If giDebug = 1 Then MsgBox ("30.4")
    
    lErro = AFRAC_LerInformacaoImpressora("050", sMFAdicional)
'    If lErro <> SUCESSO Then
'        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Informação Impressora")
'        If lErro <> SUCESSO Then gError 207192
'    End If
    
    gsMFAdicional = sMFAdicional
    
    If giDebug = 1 Then MsgBox ("30.5")
    
    dtDataHoraSistema = Now
    
    'verifica se ha diferenca de data, hora > 1 hora - PAFECF
    lErro = AFRAC_DataHoraImpressora(sData, sHora)
    If lErro <> SUCESSO Then
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data/Hora Impressora")
        If lErro <> SUCESSO Then gError 210780
    End If


    If giUsaImpressoraFiscal = 1 Then
    
        dtDataECF = left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 2)

'        If dtDataECF <> dtDataSistema Then gError 210781

        dtHoraECF = left(sHora, 2) & ":" & Mid(sHora, 3, 2) & ":" & right(sHora, 2)
        
        dtDataHoraECF = dtDataECF + dtHoraECF
        

        If Abs(DateDiff("n", dtDataHoraSistema, dtDataHoraECF)) > 60 Then gError 210782
        
        If gdtUltimaReducao = DATA_NULA Then
        
            'data do movimento da ultima reducao Z, nao necessariamente a data em que foi feita.
            lErro = AFRAC_DataReducao(dtUltimaReducaoECF)
            lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Data Ultima Reducao Z")
            If lErro <> SUCESSO Then gError 214010
        
'            lErro = AFRAC_Carrega_Dados_UltimaReducao(objUltimaReducao, objLojaArqFisMestre, colLojaArqFisAnalitico)
'            If lErro <> SUCESSO Then gError 214019
        
'            gdtUltimaReducao = objUltimaReducao.dtDataMovimento
            
            gdtUltimaReducao = dtUltimaReducaoECF
            
        End If
        
        lTamanho = 15
        sRetorno = String(lTamanho, 0)
        
        
        'pega o COO da ultima Reducao Z
        Call GetPrivateProfileString(APLICACAO_ECF, "COOultRZ", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        If sRetorno = String(lTamanho, 0) Then gError 214020
            
        lCOOUltRZ = StrParaLong(sRetorno)
                            
        If lCOOUltRZ = 0 Then
        
            If objUltimaReducao.lCOOReducaoZ = 0 Then
            
                lErro = AFRAC_Carrega_Dados_UltimaReducao(objUltimaReducao, objLojaArqFisMestre, colLojaArqFisAnalitico)
                If lErro <> SUCESSO Then gError 214022
        
            End If
            
            sCOOFinal1 = objUltimaReducao.lCOOReducaoZ
            
            'Atualiza o arquivo
            lErro = WritePrivateProfileString(APLICACAO_ECF, "COOultRZ", sCOOFinal1, NOME_ARQUIVO_CAIXA)
            If lErro = SUCESSO Then gError 214021
        
        End If
        
        
    End If
    
    
    If giDebug = 1 Then MsgBox ("31")
    
    Inicializa_ECF = SUCESSO
    
    Exit Function
    
Erro_inicializa_ECF:
    
    Inicializa_ECF = gErr
    
    Select Case gErr
                
        Case 99873 To 99878, 99914 To 99918, 109700, 109701, 109702, 112060, 112071, 113061, 126792, 126793, 133671, 204372 To 204374, 207192, 210780, 210911, 210912, 214019, 214022
                                
        Case 210781
'            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_DIFEREM, gErr, dtDataECF, dtDataSistema)
            
        Case 210782
            Call Rotina_ErroECF(vbOKOnly, ERRO_HORAS_DIFEREM, gErr, dtDataHoraECF, dtDataHoraSistema)
            
        Case 214020
            Call Rotina_ErroECF(vbOKOnly, ERRO_PREENCHIMENTO_ARQUIVO_CONFIG, gErr, "COOultRZ", APLICACAO_ECF, NOME_ARQUIVO_CAIXA)
        
        Case 214021
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "COOultRZ", NOME_ARQUIVO_CAIXA)
                                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165195)

    End Select
    
    Exit Function
    
End Function

Private Function Carrega_Sistema() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Sistema

    lErro = CF_ECF("Carrega_Caixa_Config")
    If lErro <> SUCESSO Then gError 99598
    
    lErro = CF_ECF("AbreBDs_PAFECF")
    If lErro <> SUCESSO Then gError 210413
    
    lErro = CF_ECF("Carrega_Arquivo_FonteDados")
    If lErro <> SUCESSO Then gError 99599
        
#If TESTE_SEM_IMPRESSORA = 1 Then
    
    giDebug = 1
    giUsaImpressoraFiscal = 0
    giCodModeloECF = IMPRESSORA_NAO_FISCAL
    
#End If

    'se em orcamento.ini esta indicado que usa impressora fiscal e o modelo for diferente de 4 (4 é o modelo nao fiscal) ==> erro
    If giUsaImpressoraFiscal = 1 And giCodModeloECF = IMPRESSORA_NAO_FISCAL Then gError 204297
    
    'se em orcamento.ini esta indicado que nao usa impressora fiscal e o modelo for  4 (4 é o modelo nao fiscal) ==> erro
    If giUsaImpressoraFiscal = 0 And giCodModeloECF <> IMPRESSORA_NAO_FISCAL Then gError 204298
        
    If AFRAC_ImpressoraCFe(giCodModeloECF) Then
        giPreVenda = 0
    End If
        
'    lErro = CF_ECF("Inicializa_Dados")
'    If lErro <> SUCESSO Then gError 99776
    
    giCodModeloECFConfig = giCodModeloECF
    
    Carrega_Sistema = SUCESSO
    
    Exit Function
    
Erro_Carrega_Sistema:
    
    Carrega_Sistema = gErr
    
    Select Case gErr
    
        Case 99598, 99599, 99776, 204860, 210196, 210241, 210413
        
        Case 204297
            Call Rotina_ErroECF(vbOKOnly, ERRO_USAIMPRESSORAFISCAL_INVALIDO, gErr)
        
        Case 204298
            Call Rotina_ErroECF(vbOKOnly, ERRO_USAIMPRESSORANAOFISCAL_INVALIDO, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165196)

    End Select
    
    Exit Function

End Function





'Private Function Carrega_Dados_ProdutosCodBarras(sRegistro As String, colTabPreco As Collection, colProdutoDesconto As Collection) As Long
'
'Dim iPos As Integer
'Dim objProduto As ClassProduto
'Dim objProduto2 As ClassProduto
'Dim iIndice As Integer
'Dim iPosInicio As Integer
'Dim iPosFim As Integer
'Dim iPosMeio As Integer
'Dim iPosColInicio As Integer
'Dim sProduto As String
'Dim objTabPreco As New ClassTabelaPrecoItem
'Dim lErro As Long
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Carrega_Dados_ProdutosCodBarras
'
'    Set objProduto = New ClassProduto
'
'    'Primeira Posição
'    iPosInicio = 1
'
'    'Procura o Primeiro Control para saber onde começa a string
'    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
'
'    'Procura o Primeiro Escape dentro da String
'    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'    'Pega última posição e guarda
'    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
'
'    iIndice = 0
'
'    Do While iPosMeio <> 0
'
'       iIndice = iIndice + 1
'
'        Select Case iIndice
'
'            Case 1: objProduto.sCodigo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 2: objProduto.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 3: objProduto.sSiglaUMVenda = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 4: objProduto.sReferencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 5: objProduto.sFigura = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 6: objProduto.sSituacaoTribECF = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 7: objProduto.sICMSAliquota = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 8: objProduto.sCodigoBarras = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 9: objProduto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 10
'                objProduto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                If objProduto.iAtivo = PRODUTO_INATIVO Then gError 112613
'            Case Else:
'
'        End Select
'
'        'Atualiza as Posições
'        iPosInicio = iPosMeio + 1
'        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'        If iPosInicio <= iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
'
'    Loop
'
'    bAchou = False
'
'    For Each objTabPreco In colTabPreco
'        If UCase(objTabPreco.sCodProduto) = UCase(objProduto.sCodigo) Then
'            objProduto.dPrecoLoja = objTabPreco.dPreco
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    If bAchou Then
'        For Each objProduto2 In colProdutoDesconto
'            If objProduto2.sCodigo = objProduto.sCodigo Then
'                objProduto.dPercentMenosReceb = objProduto2.dPercentMenosReceb
'                Exit For
'            End If
'        Next
'
'        gaobjProdutosCodBarras.Add objProduto
'    End If
'
'    Carrega_Dados_ProdutosCodBarras = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Dados_ProdutosCodBarras:
'
'    Carrega_Dados_ProdutosCodBarras = gErr
'
'    Select Case gErr
'
'        Case 112613
'
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 165197)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function Carrega_Dados_ProdutosReferencia(sRegistro As String, colTabPreco As Collection, colProdutoDesconto As Collection) As Long
'
'Dim iPos As Integer
'Dim objProduto As ClassProduto
'Dim objProduto2 As ClassProduto
'Dim iIndice As Integer
'Dim iPosInicio As Integer
'Dim iPosFim As Integer
'Dim iPosMeio As Integer
'Dim iPosColInicio As Integer
'Dim sProduto As String
'Dim objTabPreco As New ClassTabelaPrecoItem
'Dim lErro As Long
'Dim bAchou As Boolean
'
'On Error GoTo Erro_Carrega_Dados_ProdutosReferencia
'
'    Set objProduto = New ClassProduto
'
'    'Primeira Posição
'    iPosInicio = 1
'
'    'Procura o Primeiro Control para saber onde começa a string
'    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1
'
'    'Procura o Primeiro Escape dentro da String
'    iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'    'Pega última posição e guarda
'    iPosFim = (InStr(iPosInicio, sRegistro, Chr(vbKeyEnd)))
'
'    iIndice = 0
'
'    Do While iPosMeio <> 0
'
'       iIndice = iIndice + 1
'
'        Select Case iIndice
'
'            Case 1: objProduto.sCodigo = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 2: objProduto.sNomeReduzido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 3: objProduto.sSiglaUMVenda = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 4: objProduto.sReferencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 5: objProduto.sFigura = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 6: objProduto.sSituacaoTribECF = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 7: objProduto.sICMSAliquota = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 8: objProduto.sDescricao = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
'            Case 9
'                objProduto.iAtivo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
'                If objProduto.iAtivo = PRODUTO_INATIVO Then gError 112615
'            Case Else:
'
'        End Select
'
'        'Atualiza as Posições
'        iPosInicio = iPosMeio + 1
'        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))
'
'        If iPosInicio <= iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim
'
'    Loop
'
'    bAchou = False
'
'    For Each objTabPreco In colTabPreco
'        If UCase(objTabPreco.sCodProduto) = UCase(objProduto.sCodigo) Then
'            objProduto.dPrecoLoja = objTabPreco.dPreco
'            bAchou = True
'            Exit For
'        End If
'    Next
'
'    If bAchou Then
'        For Each objProduto2 In colProdutoDesconto
'            If objProduto2.sCodigo = objProduto.sCodigo Then
'                objProduto.dPercentMenosReceb = objProduto2.dPercentMenosReceb
'                Exit For
'            End If
'        Next
'
'        If Len(Trim(objProduto.sReferencia)) > 0 Then
'            gaobjProdutosReferencia.Add objProduto
'        End If
'
'        gaobjProdutosNome.Add objProduto
'
'        If objProduto.colCodBarras > 0 Then
'            gaobjProdutosCodBarras.Add objProduto
'        End If
'
'    End If
'
'    Carrega_Dados_ProdutosReferencia = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Dados_ProdutosReferencia:
'
'    Carrega_Dados_ProdutosReferencia = gErr
'
'    Select Case gErr
'
'        Case 112615
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, 165198)
'
'    End Select
'
'    Exit Function
'
'End Function




Private Function Rotina_Configuracao_Empresa(bConfigurouEmpresa As Boolean) As Long
'faz a configuracao a nivel de empresa

Dim colModuloFilEmp As New Collection
Dim objModuloFilEmp As ClassModuloFilEmp
Dim iConfigurarSGE As Integer
Dim objConfiguraADM As New ClassConfiguraADM
Dim colModuloFilial As New Collection
Dim lErro As Long
Dim objFiliais As AdmFiliais

On Error GoTo Erro_Rotina_Configuracao_Empresa

    'le todos os objetos ModuloFilEmp para a empresa em questão e coloca-os em colModuloFilEmp
    lErro = ModuloFilEmp_Le_EmpresaFilial(glEmpresa, EMPRESA_TODA, colModuloFilEmp)
    If lErro <> SUCESSO Then Error 44858
    
    iConfigurarSGE = True
    
    'pesquisa se há algum módulo a configurar que necessita passar pela tela de configuração
    For Each objModuloFilEmp In colModuloFilEmp
        If objModuloFilEmp.iConfigurado = NAO_CONFIGURADO Then
                objConfiguraADM.colModulosConfigurar.Add objModuloFilEmp.sSiglaModulo
        End If
        'pesquisa se há algum modulo da empresa já configurado ==> significa que a configuração geral da empresa já foi feita
        If objModuloFilEmp.iConfigurado = CONFIGURADO Then
            iConfigurarSGE = False
        End If
    Next
    
    If iConfigurarSGE = True Then objConfiguraADM.colModulosConfigurar.Add SISTEMA_SGE
    
    If objConfiguraADM.colModulosConfigurar.Count > 0 Then
    
        Call Carrega_ColFiliais_EmpresaToda
        
        'carrega o wizard de configuração da empresa
        objConfiguraADM.iConfiguracaoOK = False
    
        Call Chama_Tela("frmWizardEmpresa", objConfiguraADM)
    
        If objConfiguraADM.iConfiguracaoOK = False Then Error 44859
    
        lErro = CF("Retorna_ColFiliais")
        If lErro <> SUCESSO Then Error 44944
    
        bConfigurouEmpresa = True
    
    End If
    
    Rotina_Configuracao_Empresa = SUCESSO
    
    Exit Function
    
Erro_Rotina_Configuracao_Empresa:

    Rotina_Configuracao_Empresa = Err
    
    Select Case Err
    
        Case 44858, 44859, 44944
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, 165199)

    End Select
    
    Exit Function
    
End Function

Private Function Rotina_Configuracao_Filial(bConfigurouFilial As Boolean) As Long
'faz a configuracao a nivel de filial

Dim colModuloFilEmp As New Collection
Dim objModuloFilEmp As ClassModuloFilEmp
Dim objConfiguraADM As New ClassConfiguraADM
Dim colModuloFilial As New Collection
Dim lErro As Long

On Error GoTo Erro_Rotina_Configuracao_Filial
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        'le todos os objetos ModuloFilEmp para a filial em questão e coloca-os em colModuloFilEmp
        lErro = ModuloFilEmp_Le_EmpresaFilial(glEmpresa, giFilialEmpresa, colModuloFilEmp)
        If lErro <> SUCESSO Then Error 44860
    
        'pesquisa se há algum módulo a configurar que necessita passar pela tela de configuração
        For Each objModuloFilEmp In colModuloFilEmp
            If objModuloFilEmp.iConfigurado = NAO_CONFIGURADO Then
                objConfiguraADM.colModulosConfigurar.Add objModuloFilEmp.sSiglaModulo
                'seleciona os módulos que necessitam passar por tela de configuração.
                If objModuloFilEmp.sSiglaModulo = MODULO_ESTOQUE Then
                    colModuloFilial.Add objModuloFilEmp.sSiglaModulo
                End If
                    
            End If
        Next
    
        If colModuloFilial.Count > 0 Then
            
            Call Carrega_ColFiliais_Filial(objConfiguraADM)
            
            objConfiguraADM.iConfiguracaoOK = False
        
            'carrega o wizard de configuração da filial
            Call Chama_Tela("frmWizardFilial", objConfiguraADM)
        
            If objConfiguraADM.iConfiguracaoOK = False Then Error 44861
            
            lErro = CF("Retorna_ColFiliais")
            If lErro <> SUCESSO Then Error 44946
        
            bConfigurouFilial = True
        
        ElseIf objConfiguraADM.colModulosConfigurar.Count > 0 Then
        
            Call Carrega_ColFiliais_Filial(objConfiguraADM)
        
            lErro = Gravar_Registro(objConfiguraADM.colModulosConfigurar)
            If lErro <> SUCESSO Then Error 44875
        
            lErro = CF("Retorna_ColFiliais")
            If lErro <> SUCESSO Then Error 44947
        
            bConfigurouFilial = True
        
        End If
    
    End If

    Rotina_Configuracao_Filial = SUCESSO
    
    Exit Function
    
Erro_Rotina_Configuracao_Filial:

    Rotina_Configuracao_Filial = Err
    
    Select Case Err
    
        Case 44860, 44861, 44875, 44946, 44947
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, 165200)

    End Select
    
    Exit Function
    
End Function

Private Sub Carrega_ColFiliais_EmpresaToda()

Dim objFiliais As New AdmFiliais

    'coloca gcolFiliais como uma coleção de filiais composta somente pela empresa toda
    Set gcolFiliais = New Collection
    
    objFiliais.sNome = gsNomeEmpresa
    objFiliais.iCodFilial = EMPRESA_TODA

    'coloca a filial lida na coleção
    gcolFiliais.Add objFiliais
    
    Exit Sub

End Sub

Private Sub Carrega_ColFiliais_Filial(objConfiguraADM As ClassConfiguraADM)

Dim objFiliais As New AdmFiliais

    'coloca gcolFiliais como uma coleção de filiais composta somente pela empresa toda
    Set gcolFiliais = New Collection
    
    objFiliais.sNome = gsNomeFilialEmpresa
    objFiliais.iCodFilial = giFilialEmpresa
    Set objFiliais.colModulos = objConfiguraADM.colModulosConfigurar

    'coloca a filial lida na coleção
    gcolFiliais.Add objFiliais

End Sub

Private Function Valida_Step(sModulo As String, colModulosConfigurar As Collection) As Long

Dim vModulo As Variant

    For Each vModulo In colModulosConfigurar

        If sModulo = vModulo Then
            Valida_Step = SUCESSO
            Exit Function
        End If
        
    Next
    
    Valida_Step = 44870

End Function

Private Function Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim lTransacao As Long
Dim lTransacaoDic As Long
Dim lConexao As Long

On Error GoTo Erro_Gravar_Registro
    
    lConexao = GL_lConexaoDic
    
    'Inicia a Transacao
    lTransacaoDic = Transacao_AbrirExt(lConexao)
    If lTransacaoDic = 0 Then Error 44963
    
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 44871
    
    lErro = CTB_Exercicio_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 44872
    
    lErro = CR_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41927
    
    lErro = EST_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41928
    
    lErro = FAT_Filial_Gravar_Registro(colModulosConfigurar)
    If lErro <> SUCESSO Then Error 41929
    
    lErro = CF("ModuloFilEmp_Atualiza_Configurado", glEmpresa, giFilialEmpresa, colModulosConfigurar)
    If lErro <> SUCESSO Then Error 44956
    
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 44873
    
    lErro = Transacao_CommitExt(lTransacaoDic)
    If lErro <> AD_SQL_SUCESSO Then Error 44964
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = Err
    
    Select Case Err

        Case 44871
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_TRANSACAO1, Err)

        Case 44872, 44956, 44963, 44964, 41927, 41928, 41929

        Case 44873
            Call Rotina_ErroECF(vbOKOnly, ERRO_COMMIT_TRANSACAO1, Err)
            
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 165201)

    End Select

    If Err <> 44964 Then Call Transacao_Rollback
    Call Transacao_RollbackExt(lTransacaoDic)

    Exit Function
    
End Function

Private Function CTB_Exercicio_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long

On Error GoTo Erro_CTB_Exercicio_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTABILIDADE, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("Exercicio_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 44874
        
    End If
    
    CTB_Exercicio_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CTB_Exercicio_Gravar_Registro:
    
    CTB_Exercicio_Gravar_Registro = Err
    
    Select Case Err

        Case 44874

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error$, 165202)

    End Select

    Exit Function
    
End Function

Private Function CR_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_CR_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_CONTASARECEBER, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("CR_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41913
        
    End If
    
    CR_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_CR_Filial_Gravar_Registro:
    
    CR_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41913

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error$, 165203)

    End Select

    Exit Function
    
End Function

Private Function EST_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection
Dim sIntervaloProducao As String

On Error GoTo Erro_EST_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_ESTOQUE, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        sIntervaloProducao = "0"
        lErro = CF("EST_Instalacao_Filial", giFilialEmpresa, sIntervaloProducao)
        If lErro <> SUCESSO Then Error 41914
        
    End If
    
    EST_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_EST_Filial_Gravar_Registro:
    
    EST_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41914

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error$, 165204)

    End Select

    Exit Function
    
End Function

Private Function FAT_Filial_Gravar_Registro(colModulosConfigurar As Collection) As Long

Dim lErro As Long
Dim colSegmentos As Collection

On Error GoTo Erro_FAT_Filial_Gravar_Registro

    lErro = Valida_Step(MODULO_FATURAMENTO, colModulosConfigurar)

    If lErro = SUCESSO Then
        
        lErro = CF("FAT_Instalacao_Filial", giFilialEmpresa)
        If lErro <> SUCESSO Then Error 41915
        
    End If
    
    FAT_Filial_Gravar_Registro = SUCESSO
       
    Exit Function
    
Erro_FAT_Filial_Gravar_Registro:
    
    FAT_Filial_Gravar_Registro = Err
    
    Select Case Err

        Case 41915

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error$, 165205)

    End Select

    Exit Function
    
End Function

Public Function Empresa_Filial_Configura(objFilialEmpresa As ClassFilialEmpresa) As Long

Dim bConfigurouEmpresa As Boolean
Dim bConfigurouFilial As Boolean
Dim lErro As Long

On Error GoTo Erro_Empresa_Filial_Configura

    'seleciona a Empresa e filial
    lErro = Sistema_DefEmpresa(objFilialEmpresa.sNomeEmpresa, objFilialEmpresa.lCodEmpresa, objFilialEmpresa.sNomeFilial, objFilialEmpresa.iCodFilial)
    If lErro <> AD_BOOL_TRUE Then Error 41619
    
    glEmpresa = objFilialEmpresa.lCodEmpresa
    
    bConfigurouEmpresa = False
    
    lErro = Rotina_Configuracao_Empresa(bConfigurouEmpresa)
    If lErro <> SUCESSO Then Error 44876
    
    bConfigurouFilial = False
    
    lErro = Rotina_Configuracao_Filial(bConfigurouFilial)
    If lErro <> SUCESSO Then Error 44877
    
    'se houve configuracao de modulo
    If bConfigurouEmpresa = True Or bConfigurouFilial = True Then
    
        'força a reinicializacao dos modulos, por exemplo para pegar a nova mascara de conta contabil
        If Sistema_Inicializa_Modulos <> SUCESSO Then Error 56601

    End If
    
    Empresa_Filial_Configura = SUCESSO
    
    Exit Function
    
Erro_Empresa_Filial_Configura:

    Empresa_Filial_Configura = Err
    
    Select Case Err
    
        Case 41619, 44876, 44877, 56601  'tratado na rotina chamada
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, 165206)

    End Select
    
    Exit Function
 
End Function

'Private Sub Atualiza_Movimentos(objVenda As ClassVenda)
'
'Dim objMovCaixa As ClassMovimentoCaixa
'Dim objCheque As ClassChequePre
'Dim lErro As Long
'
'On Error GoTo Erro_Atualiza_Movimentos
'
'   'Jogo todos os cheques na col global
'    For Each objCheque In objVenda.colCheques
'        'Atualiza o saldos de cheques
'        gdSaldocheques = gdSaldocheques + objCheque.dValor
'    Next
'
'    'Para cada movimento da venda
'    For Each objMovCaixa In objVenda.colMovimentosCaixa
'        'Se for de cartao de crédito ou débito especificado
'        If objMovCaixa.iTipo = MOVIMENTOCAIXA_RECEB_DINHEIRO Then gdSaldoDinheiro = gdSaldoDinheiro + objMovCaixa.dValor
'    Next
'
'    Exit Sub
'
'Erro_Atualiza_Movimentos:
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 165207)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Function Desmembra_MovimentosCaixa(objMovimentosCaixa As ClassMovimentoCaixa, sRegistro As String, colImfCompl As Collection) As Long
'Função que Desmembra o Movimentos de Caixa e Carrega em gcolMovimentosCaixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_MovimentosCaixa

    'Instancia o Obj da ClassMovimentoCaixa
    Set objMovimentosCaixa = New ClassMovimentoCaixa
    
    iInicial = 1
         
    'Primeira Posição
    iPosInicio = 1
    
    'Inicializa a variavel
    iIndice = 0
    
    'Posição Final
    iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
    
    'Procura o Primeiro Control para saber o tipo do registro
    iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
    
    'Verifica a Posição do Shifth
    iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
        
    Do While iPosInicio < (iPosEnd - 1)

        iIndice = iIndice + 1
        
        'acerta os Ponteiros
        If iIndice = 1 Then
        
            iPosInicio = iPos + 1
            iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
            
        End If
            
        'Recolhe os Dados do arquivo de movimentosCaixa e Coloca no obj
        
        Select Case iIndice
            
            Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 8: objMovimentosCaixa.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 9: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 10: objMovimentosCaixa.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 11: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 12: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 13: objMovimentosCaixa.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 14: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 15: objMovimentosCaixa.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 16: objMovimentosCaixa.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 17: objMovimentosCaixa.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 18: objMovimentosCaixa.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 19: objMovimentosCaixa.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 20: objMovimentosCaixa.sFavorecido = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
            Case 21: objMovimentosCaixa.sHistorico = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
            Case 22: Exit Do
        
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPos + 1
        iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))

    Loop
    
    'atualiza as posições
    iPosInicioShift = iPosShift + 1
    iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
    
    'Verifica se Existe na String o caracterShift se exitir executa o LOOP
    If iPosShift <> 0 Then
    
        'Enquanto for menor que a penultima posição
        Do While iPosInicioShift <= (iPosEnd - 1)
        
            'Adciona em lConteudo
            lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
            
            'Adciona na Coleção
            colImfCompl.Add lConteudo
            
            'Atualiza as posições
            iPosInicioShift = iPosShift + 1
            iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
                    
        Loop
    
    End If
    
    'Verifica se Chegou na Penultima posição se Chegou então
    If iPosInicioShift = (iPosEnd - 1) Then
     
         'Procura a marcação do flag end
         iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyEnd))
         'adciona na variável
         lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
          
         'Adciona na Coleção
         colImfCompl.Add lConteudo
    
    End If
    
    Desmembra_MovimentosCaixa = SUCESSO
       
    Exit Function

Erro_Desmembra_MovimentosCaixa:
    
    Desmembra_MovimentosCaixa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165208)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Private Function Desmembrar_ECF(ByVal colRegistro As Collection, ByVal sArquivo As String, ByVal iRegistro As Integer) As Long

Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim objVenda As New ClassVenda
Dim sRegistro As String
Dim iRegistroCol As Integer
Dim lErro As Long

On Error GoTo Erro_Desmembrar_ECF

    sRegistro = colRegistro.Item(1)

    iPosAtual = InStr(sRegistro, Chr(vbKeyControl))
    iPosAtual = iPosAtual + 1
    iPosFimAtual = InStr(iPosAtual, sRegistro, Chr(vbKeySeparator))

    objVenda.iTipo = StrParaInt(Mid(sRegistro, iPosAtual, iPosFimAtual - iPosAtual))
    objVenda.objCupomFiscal.iTipo = objVenda.iTipo
    iRegistroCol = 1
   
    'guarda as infos de carne se houverem
    lErro = CF("Vendas_Carne", iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.objCarne)
    If lErro <> SUCESSO Then gError 110006

    'pula o segundo separador
    iPosFimAtual = iPosFimAtual + 1

    'guarda as infos de cupom
    lErro = CF("Vendas_Cupom", iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.objCupomFiscal)
    If lErro <> SUCESSO Then gError 110007

    'guarda as infos de movimento de caixa
    lErro = CF("Vendas_Movcx", iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colMovimentosCaixa)
    If lErro <> SUCESSO Then gError 110008
    
    iPosFimAtual = iPosFimAtual + 1
    
    'guarda as infos de cheque se houverem
    lErro = CF("Vendas_Cheque", iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colCheques)
    If lErro <> SUCESSO Then gError 110009

    'guarda as infos de troca se houverem
    lErro = CF("Vendas_Troca", iPosAtual, iPosFimAtual, colRegistro, iRegistroCol, objVenda.colTroca)
    If lErro <> SUCESSO Then gError 110010

    lErro = CF("Venda_Gravar_CC", objVenda)
    If lErro <> SUCESSO Then gError 110011
    
    Desmembrar_ECF = SUCESSO

    Exit Function

Erro_Desmembrar_ECF:

    Desmembrar_ECF = gErr

    Select Case gErr

        Case 110006, 110007, 110008, 110009, 110010, 110011

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165209)

    End Select

    Exit Function

End Function


'Function MovimentoCheque_Atualiza_Memoria1(objMovimentosCaixa As ClassMovimentoCaixa) As Long
''Função que Limpa a Coleção Global a Tela apos a Função de Gravação
'
'Dim lErro As Long
'Dim objCheque As New ClassChequePre
'
'On Error GoTo Erro_MovimentoCheque_Atualiza_Memoria1
'
'    For Each objCheque In gcolCheque
'
'        'Verifica se é a Mensma administardora , parcelamento , Cartão
'        If objCheque.lNumMovtoSangria = objMovimentosCaixa.lNumMovto Then
'
'            gdSaldocheques = gdSaldocheques + objMovimentosCaixa.dValor
'
'        End If
'
'    Next
'
'    MovimentoCheque_Atualiza_Memoria1 = SUCESSO
'
'    Exit Function
'
'Erro_MovimentoCheque_Atualiza_Memoria1:
'
'    MovimentoCheque_Atualiza_Memoria1 = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error, 165210)
'
'    End Select
'
'    Exit Function
'
'End Function


Public Function Desmembra_Dados_TransfCaixa(sRegistro As String, objTransfCaixa As ClassTransfCaixa, objChequePara As ClassChequePre) As Long

Dim iPos As Integer
Dim objMCDe As New ClassMovimentoCaixa
Dim objMCPara As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim iPosInicio As Integer
Dim iPosFim As Integer
Dim iPosMeio As Integer

    'Primeira Posição
    iPosInicio = 1

    'Procura o Primeiro Control para saber onde começa a string
    iPosInicio = InStr(iPosInicio, sRegistro, Chr(vbKeyControl)) + 1

    'Procura o Primeiro Escape dentro da String
    iPosMeio = InStr(iPosInicio, sRegistro, Chr(vbKeyEscape))

    'Pega última posição e guarda
    iPosFim = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))

    iIndice = 0

    Do While iPosMeio <> 0

       iIndice = iIndice + 1

        Select Case iIndice

            'desmembra o Movimento de Origem

            Case 1: objMCDe.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 2: objMCDe.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 3: objMCDe.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 4: objMCDe.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 5: objMCDe.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 6: objMCDe.iCodConta = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 7: objMCDe.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 8: objMCDe.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 9: objMCDe.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 10: objMCDe.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 11: objMCDe.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 12: objMCDe.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 13: objMCDe.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 14: objMCDe.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 15: objMCDe.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 16: objMCDe.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 17: objMCDe.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 18: objMCDe.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 19: objMCDe.lNumRefInterna = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 20: objMCDe.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 21: objMCDe.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 22: objMCDe.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 23: objMCDe.sFavorecido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 24: objMCDe.sHistorico = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)

            'desmembra o Movimento de Destino

            Case 25: objMCPara.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 26: objMCPara.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 27: objMCPara.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 28: objMCPara.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 29: objMCPara.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 30: objMCPara.iCodConta = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 31: objMCPara.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 32: objMCPara.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 33: objMCPara.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 34: objMCPara.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 35: objMCPara.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 36: objMCPara.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 37: objMCPara.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 38: objMCPara.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 39: objMCPara.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 40: objMCPara.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 41: objMCPara.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 42: objMCPara.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 43: objMCPara.lNumRefInterna = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 44: objMCPara.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 45: objMCPara.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 46: objMCPara.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 47: objMCPara.sFavorecido = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 48: objMCPara.sHistorico = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            'desmembra o Cheque de Destino
            
            Case 49: objChequePara.dtDataDeposito = StrParaDate(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 50: objChequePara.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 51: objChequePara.iAprovado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 52: objChequePara.iBanco = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 53: objChequePara.iChequeSel = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 54: objChequePara.iECF = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 55: objChequePara.iFilial = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 56: objChequePara.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 57: objChequePara.iFilialEmpresaLoja = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 58: objChequePara.iNaoEspecificado = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 59: objChequePara.iStatus = StrParaInt(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 60: objChequePara.lCliente = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 61: objChequePara.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 62: objChequePara.lNumBordero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 63: objChequePara.lNumBorderoLoja = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 64: objChequePara.lNumero = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 65: objChequePara.lNumIntCheque = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 66: objChequePara.lNumMovtoCaixa = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 67: objChequePara.lNumMovtoSangria = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 68: objChequePara.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 69: objChequePara.lSequencialBack = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 70: objChequePara.lSequencialLoja = StrParaLong(Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio))
            Case 71: objChequePara.sAgencia = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 72: objChequePara.sContaCorrente = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            Case 73: objChequePara.sCPFCGC = Mid(sRegistro, iPosInicio, iPosMeio - iPosInicio)
            
        End Select

        'Atualiza as Posições
        iPosInicio = iPosMeio + 1
        iPosMeio = (InStr(iPosInicio, sRegistro, Chr(vbKeyEscape)))

        If iPosInicio < iPosFim And iPosMeio = 0 Then iPosMeio = iPosFim

    Loop

    'Joga na coleção
    gcolMovimentosCaixa.Add objMCDe
    gcolMovimentosCaixa.Add objMCPara

    If objChequePara.lNumIntCheque > 0 Then gcolCheque.Add objChequePara
    
    Set objTransfCaixa.objMovCaixaDe = objMCDe
    Set objTransfCaixa.objMovCaixaPara = objMCPara
    
    Desmembra_Dados_TransfCaixa = SUCESSO

End Function

Private Function BuscaChequePorCupom(objCheque As ClassChequePre, dValor As Double, Optional bRetornaCheque As Boolean = False) As Boolean
'busca o cheque nao especificado com o mesmo valor do cupom fiscal contido em
'objCheque (parametro de input e output) e retorna o valor de tal cheque nao especificado...
'dValor eh parametro de output

Dim objChequeAux As ClassChequePre

    'seta inicialmente o retorno como nao achado
    BuscaChequePorCupom = False

    'para cada cheque na colecao GLOBAL
    For Each objChequeAux In gcolCheque

        'verifica se o cheque eh nao especificado e tem o cupom em questao
        With objChequeAux

           If .iNaoEspecificado = 1 Then 'precisa fazer constante pra isso?????

               If .lCupomFiscal = objCheque.lCupomFiscal Then

                    'se achou... indica que achou, retorna o valor e sai
                    BuscaChequePorCupom = True
                    dValor = objChequeAux.dValor

                    'se for pra retornar o cheque no obj passado...
                    If bRetornaCheque Then

                        Set objCheque = objChequeAux

                    End If

                    Exit Function

               End If

           End If

        End With

    Next

End Function

Private Function Busca_Chq_NumIntCheque(lNumIntCheque As Long, objCheque As ClassChequePre) As Boolean
'busca o cheque atraves do NumIntCheque (parametro de input)
'e retorna o cheque em objCheque(output)

Dim objChequeAux As ClassChequePre

   Busca_Chq_NumIntCheque = False

   For Each objChequeAux In gcolCheque

      If lNumIntCheque = objChequeAux.lSequencialCaixa Then

         Set objCheque = objChequeAux

         Busca_Chq_NumIntCheque = True

         Exit Function

      End If

   Next

End Function


Function Desmembra_OperacoesECF(sRegistro As String, objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que Desmembra o Movimentos de Caixa e Carrega em gcolMovimentosCaixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_OperacoesECF

    iInicial = 1
         
    'Primeira Posição
    iPosInicio = 1
    
    'Inicializa a variavel
    iIndice = 0
    
    'Instancia o Obj da ClassMovimentoCaixa
    Set objMovimentosCaixa = New ClassMovimentoCaixa
        
    'Posição Final
    iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
    
    'Procura o Primeiro Control para saber o tipo do registro
    iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
    'Verifica a Posição do Shifth
    iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
    
    Do While iPosInicio <= iPosEnd

        iIndice = iIndice + 1
        
        'acerta os Ponteiros
        If iIndice = 1 Then
        
            iPosInicio = iPos + 1
            iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
            
        End If
            
        'Recolhe os Dados do arquivo de movimentosCaixa e Coloca no obj
        
        Select Case iIndice
            
            Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 8: objMovimentosCaixa.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 9: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 10: objMovimentosCaixa.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 11: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 12: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 13: objMovimentosCaixa.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 14: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 15: objMovimentosCaixa.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 16: objMovimentosCaixa.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 17: objMovimentosCaixa.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 18: objMovimentosCaixa.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 19: objMovimentosCaixa.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 20: objMovimentosCaixa.sFavorecido = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
            Case 21: objMovimentosCaixa.sHistorico = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
            Case 22: Exit Do
    
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPos + 1
        iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        If iPos = 0 Then iPos = iPosEnd
    Loop

    Desmembra_OperacoesECF = SUCESSO
       
    Exit Function

Erro_Desmembra_OperacoesECF:
    
    Desmembra_OperacoesECF = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165211)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Function Desmembra_MovimentosCheque(objMovimentosCaixa As ClassMovimentoCaixa, sRegistro As String, colImfCompl As Collection) As Long
'Função que Desmembra o Movimentos de Caixa e Carrega em gcolMovimentosCaixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_MovimentosCheque

    'Instancia o Obj da ClassMovimentoCaixa
    Set objMovimentosCaixa = New ClassMovimentoCaixa
    
    iInicial = 1
         
    'Primeira Posição
    iPosInicio = 1
    
    'Inicializa a variavel
    iIndice = 0
    
    'Posição Final
    iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
    
    'Procura o Primeiro Control para saber o tipo do registro
    iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
    
    'Verifica a Posição do Shifth
    iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
        
    Do While iPosInicio < (iPosEnd - 1)

        iIndice = iIndice + 1
        
        'acerta os Ponteiros
        If iIndice = 1 Then
        
            iPosInicio = iPos + 1
            iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
            
        End If
            
        'Recolhe os Dados do arquivo de movimentosCaixa e Coloca no obj
        
        Select Case iIndice
            
            Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 8: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 9: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 10: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 11: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
            Case 12: Exit Do
        
        End Select
        
        'Atualiza as Posições
        iPosInicio = iPos + 1
        iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))

    Loop
    
    'atualiza as posições
    iPosInicioShift = iPosShift + 1
    iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
    
    'Verifica se Existe na String o caracterShift se exitir executa o LOOP
    If iPosShift <> 0 Then
    
        'Enquanto for menor que a penultima posição
        Do While iPosInicioShift <= (iPosEnd - 1)
        
            'Adciona em lConteudo
            lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
            
            'Adciona na Coleção
            colImfCompl.Add lConteudo
            
            'Atualiza as posições
            iPosInicioShift = iPosShift + 1
            iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyShift))
                    
        Loop
    
    End If
    
    'Verifica se Chegou na Penultima posição se Chegou então
    If iPosInicioShift = (iPosEnd - 1) Then
     
         'Procura a marcação do flag end
         iPosShift = InStr(iPosShift + 1, sRegistro, Chr(vbKeyEnd))
         'adciona na variável
         lConteudo = StrParaLong(Mid(sRegistro, iPosInicioShift, iPosShift - iPosInicioShift))
          
         'Adciona na Coleção
         colImfCompl.Add lConteudo
    
    End If
    
    Desmembra_MovimentosCheque = SUCESSO
       
    Exit Function

Erro_Desmembra_MovimentosCheque:
    
    Desmembra_MovimentosCheque = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165212)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function

Private Function Inicializa_LeitoraCodBarras() As Long

Dim objAliquota As ClassAliquotaICMS
Dim sTributacao As String
Dim iIndice As Integer
Dim lErro As Long
Dim lErro1 As Long
Dim objTiposMeiosPagtos As ClassTMPLoja
Dim iMetodo As Integer
Dim sRetorno As String
Dim sRetorno1 As String
Dim iPos As Integer
Dim sAliquota As String
Dim bAchou As Boolean
Dim sTrib As String
Dim iInicio As Integer
Dim iFim As Integer

On Error GoTo Erro_Inicializa_LeitoraCodBarras

    
    Inicializa_LeitoraCodBarras = SUCESSO
    
    Exit Function
    
Erro_Inicializa_LeitoraCodBarras:
    
    Inicializa_LeitoraCodBarras = gErr
    
    Select Case gErr
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165213)

    End Select
    
    Exit Function
    
End Function


Function Teste()
  Dim X As String
  
  If 1 = 1 Then
    MsgBox X
  End If

End Function

Private Function TestaVersaoPgm() As Long

Dim lErro As Long, lComando As Long
Dim sConteudo As String

On Error GoTo Erro_TestaVersaoPgm

    If InStr(fOSMachineName, "JONES-STI") = 0 And InStr(fOSMachineName, "GHEINER1-PC") = 0 And InStr(fOSMachineName, "W02-PC") = 0 Then
    
        If CORPORATOR_ECF_VERSAO_PGM <> "" Then
    
            lComando = Comando_AbrirExt(glConexaoPAFECF)
            If lComando = 0 Then gError ERRO_SEM_MENSAGEM
            
            sConteudo = String(255, 0)
            lErro = Comando_Executar(lComando, "SELECT Conteudo FROM Controle WHERE Codigo = '1001'", sConteudo)
            If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = AD_SQL_SUCESSO Then
                If UCase(Trim(sConteudo)) <> UCase(Trim(CORPORATOR_ECF_VERSAO_PGM)) Then gError 201229
            End If
            
            Call Comando_Fechar(lComando)
            
        End If
    
    End If
    
    TestaVersaoPgm = SUCESSO
    
    Exit Function
    
Erro_TestaVersaoPgm:

    TestaVersaoPgm = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201229
            Call Rotina_ErroECF(vbOKOnly, ERRO_VERSAO_PGM_INCOMPATIVEL_PGM, gErr, UCase(Trim(CORPORATOR_ECF_VERSAO_PGM)), UCase(Trim(sConteudo)))
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 201228)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function TestaVersaoECF() As Long

Dim lErro As Long, lComando As Long
Dim sConteudo As String, lIdAtualizacao As Long

On Error GoTo Erro_TestaVersaoECF

    If InStr(fOSMachineName, "JONES-STI") = 0 And InStr(fOSMachineName, "GHEINER1-PC") = 0 And InStr(fOSMachineName, "W02-PC") = 0 Then
    
        If CORPORATOR_ECF_VERSAO_BD_ECF <> "" Then
    
            lComando = Comando_AbrirExt(glConexaoPAFECF)
            If lComando = 0 Then gError ERRO_SEM_MENSAGEM
            
            sConteudo = String(255, 0)
            lErro = Comando_Executar(lComando, "SELECT MAX(IdAtualizacao) FROM VersaoBD", lIdAtualizacao)
            If lErro <> AD_SQL_SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = Comando_BuscarProximo(lComando)
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro = AD_SQL_SUCESSO Then
                sConteudo = CStr(lIdAtualizacao)
                If UCase(Trim(sConteudo)) <> UCase(Trim(CORPORATOR_ECF_VERSAO_BD_ECF)) Then gError 201229
            End If
            
            Call Comando_Fechar(lComando)
            
        End If
    
    End If
    
    TestaVersaoECF = SUCESSO
    
    Exit Function
    
Erro_TestaVersaoECF:

    TestaVersaoECF = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 201229
            Call Rotina_ErroECF(vbOKOnly, ERRO_VERSAO_PGM_INCOMPATIVEL_BD_ECF, gErr, UCase(Trim(CORPORATOR_ECF_VERSAO_BD_ECF)), UCase(Trim(sConteudo)))
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 201228)

    End Select
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Private Function Verifica_Arquivo_TEF_PAYGO(objTela As Object) As Long

Dim sTipo As String
Dim sTipo1 As String
Dim sNum As String
Dim sNum1 As String
Dim iPos As Integer
Dim iPos1 As Integer
Dim iPosEsc As Integer
Dim iPosEsc1 As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim bCont As Boolean
Dim asReg1() As String
Dim asReg() As String
Dim sReg As String
Dim sNomeArq As String
Dim iPosInicio As Integer
Dim bAchou As Boolean
Dim dPag As Double
Dim sRede As String
Dim sNSU As String
Dim sFim As String
Dim dValorTotal As Double
Dim iPosAtual As Integer
Dim iPosFimAtual As Integer
Dim bCon As Boolean
Dim lNum As Long
Dim sRet As String
Dim objTiposMeiosPagtos As New ClassTMPLoja
Dim sDescricao As String
Dim iLoop As Integer
Dim lNumero As Long
Dim sTxt As String
Dim vbMsg As VbMsgBoxResult
Dim sRegistro As String
Dim sRegistro1 As String
Dim lSequencial As Long

Dim lErro As Long
Dim lTamanho As Long
Dim sRetorno As String
Dim sArquivoINTPOS As String
Dim sArquivoINTPOSBACK As String
Dim objFormMsg As Object
Dim objVenda As New ClassVenda
Dim lErro1 As Long
Dim objTEF As New ClassTEF
Dim objMsg As Object
Dim sArquivoSTS As String


On Error GoTo Erro_Verifica_Arquivo_TEF_PAYGO
        
    If giTEF = CAIXA_ACEITA_TEF Then
        
        If giDebug = 1 Then MsgBox ("35")
        
        lErro = 1
            
        'fica em loop ate o gerenciador padrao estar ativo
        Do While lErro <> SUCESSO
            
            lErro = CF_ECF("TEF_Gerenciador_Padrao_PAYGO")
            If lErro <> SUCESSO And lErro <> 133756 Then gError 133739
            
        Loop
        
        If giDebug = 1 Then MsgBox ("36")
        
        Set objTela = ECF
        Set objMsg = MsgTEF1
        
        
        If giDebug = 1 Then MsgBox ("37")
        
        sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
        
        If Len(sArquivoINTPOS) = 0 Then
            sArquivoINTPOSBACK = Dir(Arquivo_Tef_Resp2_Backup_Prop)
            
            If Len(sArquivoINTPOSBACK) <> 0 Then
                FileCopy Arquivo_Tef_Resp2_Backup_Prop, Arquivo_Tef_Resp2_Prop
                
                sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
            End If
            
        End If
        
        lTamanho = 10
        sRetorno = String(lTamanho, 0)
    
        'Pega o ultimo COO de multiplos cartoes e veja se existe necessidade de cancelamento
        Call GetPrivateProfileString(APLICACAO_ECF, "COO", CONSTANTE_ERRO, sRetorno, lTamanho, NOME_ARQUIVO_CAIXA)
        
        If sRetorno <> String(lTamanho, 0) Then
        
            sRetorno = StringZ(sRetorno)
        
            objVenda.objCupomFiscal.lNumero = StrParaLong(sRetorno)
    
        End If
        
        
        sNomeArq = Dir(gsDirMVTEF & "TEF_" & objVenda.objCupomFiscal.lNumero & "_*.txt")
    
        
        If Len(sArquivoINTPOS) <> 0 And Len(sNomeArq) > 0 Then
    
            
            vbMsg = Rotina_AvisoECF(vbYesNo, "Existe TEF a ser impresso, deseja imprimir?")

            If vbMsg = vbYes Then

                lErro = CF_ECF("TEF_Imprime_Gerencial_PAYGO", objMsg, objTela, objVenda)

                vbMsg = vbYes

                'se nao conseguiu imprimir os comprovantes e quer continuar tentando
                Do While vbMsg = vbYes And lErro <> SUCESSO

                    vbMsg = Rotina_AvisoECF(vbYesNo, "Impressora não responde, tentar novamente?")

                    If vbMsg = vbYes Then

                        'imprime gerencial
                        lErro = CF_ECF("TEF_Imprime_Gerencial_PAYGO", objMsg, objTela, objVenda)

                    End If

                Loop

                lErro1 = CF_ECF("TEF_Trata_Resp1", objTEF, Arquivo_Tef_Resp2_Prop)
                If lErro1 <> SUCESSO Then gError 214572


                'se imprimiu corretamente
                If lErro = SUCESSO Then

                    '<> 1 significa que requer confirmação =1 significa que nao precisa de confirmacao
                    If objTEF.iStatusConfirmacao <> 1 Then

                        lErro = CF_ECF("TEF_Confirma_Transacao1_PAYGO", objVenda)
                        If lErro <> SUCESSO Then gError 214573

                    Else

                        'Atualiza o arquivo
                        lErro = WritePrivateProfileString(APLICACAO_ECF, "StatusTEF", "", NOME_ARQUIVO_CAIXA)
                        If lErro = 0 Then gError 214574

                    End If


                    sArquivoSTS = Dir(Arquivo_Tef_Resp1_Prop)
                    If sArquivoSTS <> "" Then Kill Arquivo_Tef_Resp1_Prop

                    sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
                    If sArquivoINTPOS <> "" Then Kill Arquivo_Tef_Resp2_Prop

                    sArquivoINTPOSBACK = Dir(Arquivo_Tef_Resp2_Backup_Prop)
                    If sArquivoINTPOSBACK <> "" Then Kill Arquivo_Tef_Resp2_Backup_Prop


                Else



                    'se houve falha de impressao e nao quer tentar novamente

                    lErro = CF_ECF("TEF_NaoConfirma_Transacao1_PAYGO", objTela, objVenda)
                    If lErro <> SUCESSO Then gError 214575

                    'cancela os cartoes ja confirmados e nao confirma o ultimo
                    lErro = CF_ECF("TEF_CNC_PAYGO", objVenda, objMsg, objTela)
                    If lErro <> SUCESSO Then gError 214576

                    sArquivoSTS = Dir(Arquivo_Tef_Resp1_Prop)
                    If sArquivoSTS <> "" Then Kill Arquivo_Tef_Resp1_Prop

                    sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
                    If sArquivoINTPOS <> "" Then Kill Arquivo_Tef_Resp2_Prop

                    sArquivoINTPOSBACK = Dir(Arquivo_Tef_Resp2_Backup_Prop)
                    If sArquivoINTPOSBACK <> "" Then Kill Arquivo_Tef_Resp2_Backup_Prop

                    gError 214577

                End If


            Else
            
            
                'se houve falha de impressao e nao quer tentar novamente

                lErro = CF_ECF("TEF_NaoConfirma_Transacao1_PAYGO", objTela, objVenda)
                If lErro <> SUCESSO Then gError 214578

                'cancela os cartoes ja confirmados e nao confirma o ultimo
                lErro = CF_ECF("TEF_CNC_PAYGO", objVenda, objMsg, objTela)
                If lErro <> SUCESSO Then gError 214579

                sArquivoSTS = Dir(Arquivo_Tef_Resp1_Prop)
                If sArquivoSTS <> "" Then Kill Arquivo_Tef_Resp1_Prop

                sArquivoINTPOS = Dir(Arquivo_Tef_Resp2_Prop)
                If sArquivoINTPOS <> "" Then Kill Arquivo_Tef_Resp2_Prop
            
                sArquivoINTPOSBACK = Dir(Arquivo_Tef_Resp2_Backup_Prop)
                If sArquivoINTPOSBACK <> "" Then Kill Arquivo_Tef_Resp2_Backup_Prop
            
            
'                If Len(objTEF.sNSU) > 0 Then
'
'                    gError 214580
'
'                End If
            
            End If
            
        End If
            
    End If
    
    If giDebug = 1 Then MsgBox ("45")
    
    Verifica_Arquivo_TEF_PAYGO = SUCESSO
    
    Exit Function
    
Erro_Verifica_Arquivo_TEF_PAYGO:
    
    Verifica_Arquivo_TEF_PAYGO = gErr
    
    Select Case gErr
    
        Case 133739, 214572, 214573, 214575, 214576, 214578, 214579
    
        Case 214574
            Call Rotina_ErroECF(vbOKOnly, ERRO_ARQUIVO_NAO_ENCONTRADO1, gErr, APLICACAO_ECF, "StatusTEF", NOME_ARQUIVO_CAIXA)

        Case 214577, 214580
            Call Rotina_ErroECF(vbOKOnly, ERRO_TEF_FALHA_IMPRESSAO, gErr, objTEF.sRede, objTEF.sNSU, objTEF.dValorTotal)
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 214581)

    End Select
    
    
'    Close 1
'
'    Select Case gErr
'
'        Case 109699, 53, 99821, 109412
'
'        Case 99803
'            Call Rotina_ErroECF(vbOKOnly, ERRO_IMPRESSORA_NAO_RESPONDE, gErr)
'
'        Case 109566, 112395, 99803, 109413
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 165194)
'
'    End Select
    
    Exit Function
    
End Function


Private Function Caixa_Verifica_Consistencia()
'O objetivo dessa função é garantir a que
'1-Os dados de venda cadastrados no BD sejam da mesma empresa\filial\caixa (evitar erros de reinstalação)
'2-Que o sistema não entre em caso de data do windows errada para que não ocorram vendas fora da ordem cronológica
'3-Que não reinstalem um caixa com o BD defasado ou zerado
'OBS-> Só vai valer para dados registrados após a mexida (AND Empresa <> 0), ou seja, ignorará erros passados
Dim lErro As Long, vbResult As VbMsgBoxResult
Dim dtData As Date, dtDataSGEECF As Date
Dim alComando(1 To 5) As Long, iIndice As Integer
Dim iEmp As Integer, iFil As Integer, iCx As Integer

On Error GoTo Erro_Caixa_Verifica_Consistencia
   
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_AbrirExt(glConexaoPAFECF)
        If alComando(iIndice) = 0 Then gError 216248
    Next

    'Verifica se existem vendas de outra empresa\filial\caixa
    lErro = Comando_Executar(alComando(1), "SELECT Data, Empresa, FilialEmpresa, CodCaixa FROM MovimentoCaixa WHERE Tipo = ? AND Empresa <> 0 AND (Empresa <> ? OR FilialEmpresa <> ? OR CodCaixa <> ?)", dtData, iEmp, iFil, iCx, TIPOREGISTROECF_VENDAS, giCodEmpresa, giFilialEmpresa, giCodCaixa)
    If lErro <> AD_SQL_SUCESSO Then gError 216249
    
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216250

    If lErro = AD_SQL_SUCESSO Then gError 216251 'BD com vendas de outra empresa\filial\caixa
    
    If Year(Date) < 2017 Then gError 216252 'Data retroativa
    
    dtDataSGEECF = FileDateTime(App.Path & "\admlib2.dll")
    
    If Now < dtDataSGEECF Then gError 216253 'Data retroativa baseada no admlib2.dll

    'Verifica a última venda
    lErro = Comando_Executar(alComando(2), "SELECT Data FROM MovimentoCaixa WHERE Tipo = ? AND Empresa <> 0 ORDER BY Data DESC", dtData, TIPOREGISTROECF_VENDAS)
    If lErro <> AD_SQL_SUCESSO Then gError 216254
    
    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216255
    
    If lErro = AD_SQL_SUCESSO Then
        If Date < dtData Then gError 216256 'Data retroativa baseada na última venda
        
        If Abs(DateDiff("d", Date, dtData)) > 7 Then
            'Se a venda anterior ocorreu há mais de 7 dias avisa pois o BD pode ser fruto de uma restauração errada
            vbResult = Rotina_AvisoECF(vbYesNo, "A última venda registrada no banco de dados é de " & Format(dtData, "dd/mm/yyyy") & " só prossiga caso isso esteja correto. Deseja prosseguir ?")
            If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
        End If
    End If
    
    'Verifica a última venda
    lErro = Comando_Executar(alComando(3), "SELECT Data FROM MovimentoCaixa WHERE Tipo = ? ", dtData, TIPOREGISTROECF_VENDAS)
    If lErro <> AD_SQL_SUCESSO Then gError 216257
    
    lErro = Comando_BuscarPrimeiro(alComando(3))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 216258
    
    If lErro = AD_SQL_SEM_DADOS Then
        'Confirma se é um BD novo (sem vendas) ou se reinstalaram errado
        vbResult = Rotina_AvisoECF(vbYesNo, "Não existem vendas cadastradas para esse caixa, somente prossiga caso isso esteja correto. Deseja prosseguir ?")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    End If
    
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Caixa_Verifica_Consistencia = SUCESSO
    
    Exit Function
    
Erro_Caixa_Verifica_Consistencia:

    Caixa_Verifica_Consistencia = gErr

    Select Case gErr

        Case 216248
            Call Rotina_ErroECF(vbOKOnly, ERRO_ABERTURA_COMANDO, gErr)

        Case 216249, 216250, 216254, 216255, 216257, 216258
            Call Rotina_ErroECF(vbOKOnly, ERRO_LEITURA_MOVIMENTOCAIXA, gErr)
            
        Case 216251
            Call Rotina_ErroECF(vbOKOnly, "Erro de configuração do caixa. Existe vendas para Empresa: " & CStr(iEmp) & " - Filial: " & CStr(iFil) & " - Caixa: " & CStr(iCx) & " em " & Format(dtData, "dd/mm/yyyy") & " e o CaixaConfig está configurado atualmente para Empresa: " & CStr(giCodEmpresa) & " - Filial: " & CStr(giFilialEmpresa) & " - Caixa: " & CStr(giCodCaixa), gErr)

        Case 216252
            Call Rotina_ErroECF(vbOKOnly, "Erro de Data (" & Format(Date, "dd/mm/yyyy") & "). Favor verificar a data exibida no seu sistema operacional", gErr)

        Case 216253
            Call Rotina_ErroECF(vbOKOnly, "Erro de Data (" & Format(Date, "dd/mm/yyyy") & "). Data do executável " & Format(dtDataSGEECF, "dd/mm/yyyy"), gErr)

        Case 216256
            Call Rotina_ErroECF(vbOKOnly, "Erro de Data (" & Format(Date, "dd/mm/yyyy") & "). Última venda cadastrada em " & Format(dtData, "dd/mm/yyyy"), gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB, gErr, Error$, 216260)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
End Function

