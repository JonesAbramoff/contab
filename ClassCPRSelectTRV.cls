VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPRSelectTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

Function CoinfoNumCCred_Le_NumIntParc(ByVal lNumIntDocParc As Long, sNumCCred As String, lNumVou As Long) As Long
'Lê dados da baixa de uma parcela a receber de um titulo baixado ou nao

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String

On Error GoTo Erro_CoinfoNumCCred_Le_NumIntParc

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 182833
    
    sNumCCred = String(255, 0)
    
    sSQL = "SELECT NumVou, RIGHT(NumCCred,6) FROM TRVVouchers AS V, ParcelasRec AS P WHERE V.TipoDocDestino = 3 AND V.NumIntDocDestino = P.NumIntTitulo AND V.Cartao = 1 AND P.NumIntDoc = ?"
    'sSQL = "SELECT NumVou, Cartao FROM CoinfoItemFaturaImport2 WHERE NumIntParc = ?"

    lErro = Comando_Executar(lComando, sSQL, lNumVou, sNumCCred, lNumIntDocParc)
    If lErro <> AD_SQL_SUCESSO Then gError 182834

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 182835
    
    Call Comando_Fechar(lComando)

    CoinfoNumCCred_Le_NumIntParc = SUCESSO

    Exit Function

Erro_CoinfoNumCCred_Le_NumIntParc:

    CoinfoNumCCred_Le_NumIntParc = gErr

    Select Case gErr

        Case 182833
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 182834, 182835
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_COINFOITEMFATURA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182836)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function RecebAntecipados_Le_FilialCliente_Baixa_TRV(ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal iConta As Integer, ByVal colRecebAntecipado As Collection) As Long
'Lê os Recebimentos antecipados com Filial e Cliente passados em objFilialCliente, Excluido<>EXCLUIDO, SaldoNaoApropriado>0, ordenado por DataMovimento e de acordo com giFilialEmpresa.

Dim lComando As Long
Dim lErro As Long
Dim tAntecipRec As typeAntecipRec
Dim tMovCtaCorrente As typeMovContaCorrente
Dim tContaCorrente As typeContaCorrenteInt
Dim objAntecipRec As ClassAntecipReceb
Dim iContador As Integer
Dim bValido As Boolean

On Error GoTo Erro_RecebAntecipados_Le_FilialCliente_Baixa_TRV

    'Inicializa comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 29138

    tMovCtaCorrente.sHistorico = String(STRING_HISTORICOMOVCONTA, 0)
    tMovCtaCorrente.sNumRefExterna = String(STRING_NUMREFEXTERNA, 0)
    
    'Se alguma Filial está selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then

        lErro = Comando_Executar(lComando, "SELECT RecebAntecipados.Cliente, RecebAntecipados.Filial_Cliente, RecebAntecipados.NumIntRec, RecebAntecipados.NumMovto, RecebAntecipados.SaldoNaoApropriado, MovimentosContaCorrente.FilialEmpresa, MovimentosContaCorrente.CodConta, MovimentosContaCorrente.Sequencial, MovimentosContaCorrente.TipoMeioPagto, MovimentosContaCorrente.Numero, MovimentosContaCorrente.DataMovimento, MovimentosContaCorrente.Valor, MovimentosContaCorrente.Historico, MovimentosContaCorrente.Portador, MovimentosContaCorrente.NumRefExterna FROM RecebAntecipados, MovimentosContaCorrente WHERE MovimentosContaCorrente.NumMovto = RecebAntecipados.NumMovto AND RecebAntecipados.Excluido <> ? AND RecebAntecipados.SaldoNaoApropriado > ? AND MovimentosContaCorrente.FilialEmpresa = ? ORDER BY MovimentosContaCorrente.DataMovimento ", _
        tAntecipRec.lCliente, tAntecipRec.iFilial_Cliente, tAntecipRec.lNumIntRec, tAntecipRec.lNumMovto, tAntecipRec.dSaldoNaoApropriado, tContaCorrente.iFilialEmpresa, tContaCorrente.iCodigo, tMovCtaCorrente.lSequencial, tMovCtaCorrente.iTipoMeioPagto, tMovCtaCorrente.lNumero, tMovCtaCorrente.dtDataMovimento, tMovCtaCorrente.dValor, tMovCtaCorrente.sHistorico, tMovCtaCorrente.iPortador, tMovCtaCorrente.sNumRefExterna, EXCLUIDO, DELTA_VALORMONETARIO, giFilialEmpresa)
        If lErro <> AD_SQL_SUCESSO Then Error 29139

    'Se Empresa Toda está selecionada
    Else

        lErro = Comando_Executar(lComando, "SELECT RecebAntecipados.Cliente, RecebAntecipados.Filial_Cliente,RecebAntecipados.NumIntRec, RecebAntecipados.NumMovto, RecebAntecipados.SaldoNaoApropriado, MovimentosContaCorrente.FilialEmpresa, MovimentosContaCorrente.CodConta, MovimentosContaCorrente.Sequencial, MovimentosContaCorrente.TipoMeioPagto, MovimentosContaCorrente.Numero, MovimentosContaCorrente.DataMovimento, MovimentosContaCorrente.Valor, MovimentosContaCorrente.Historico, MovimentosContaCorrente.Portador, MovimentosContaCorrente.NumRefExterna FROM RecebAntecipados, MovimentosContaCorrente WHERE MovimentosContaCorrente.NumMovto = RecebAntecipados.NumMovto AND RecebAntecipados.Excluido <> ? AND RecebAntecipados.SaldoNaoApropriado > ? ORDER BY MovimentosContaCorrente.DataMovimento", _
         tAntecipRec.lCliente, tAntecipRec.iFilial_Cliente, tAntecipRec.lNumIntRec, tAntecipRec.lNumMovto, tAntecipRec.dSaldoNaoApropriado, tContaCorrente.iFilialEmpresa, tContaCorrente.iCodigo, tMovCtaCorrente.lSequencial, tMovCtaCorrente.iTipoMeioPagto, tMovCtaCorrente.lNumero, tMovCtaCorrente.dtDataMovimento, tMovCtaCorrente.dValor, tMovCtaCorrente.sHistorico, tMovCtaCorrente.iPortador, tMovCtaCorrente.sNumRefExterna, EXCLUIDO, DELTA_VALORMONETARIO)
        If lErro <> AD_SQL_SUCESSO Then Error 29140

    End If

    'Lê o primeiro Recebimento antecipado
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 29141

    'Se não existe Recebimento antecipado com os dados desejados
    If lErro = AD_SQL_SEM_DADOS Then Error 29142
    
    iContador = 0

    Do While lErro <> AD_SQL_SEM_DADOS 'And iContador < NUM_MAX_RECEB_ANTECIPADOS

        bValido = True
        
        If dtDataDe <> DATA_NULA Then
            If tMovCtaCorrente.dtDataMovimento < dtDataDe Then bValido = False
        End If
        If dtDataAte <> DATA_NULA Then
            If tMovCtaCorrente.dtDataMovimento > dtDataAte Then bValido = False
        End If
        If iConta <> 0 Then
            If tContaCorrente.iCodigo <> iConta Then bValido = False
        End If

        If bValido Then

            iContador = iContador + 1
            
            Set objAntecipRec = New ClassAntecipReceb
    
            'Passa os dados do Recebimento Antecipado para o Obj
            objAntecipRec.dSaldoNaoApropriado = tAntecipRec.dSaldoNaoApropriado
            objAntecipRec.dtData = tMovCtaCorrente.dtDataMovimento
            objAntecipRec.dValor = tMovCtaCorrente.dValor
            objAntecipRec.iCodConta = tContaCorrente.iCodigo
            objAntecipRec.iExcluido = NAO_EXCLUIDO
            objAntecipRec.iFilial = tAntecipRec.iFilial_Cliente
            objAntecipRec.iFilialEmpresa = tContaCorrente.iFilialEmpresa
            objAntecipRec.iTipoMeioPagto = tMovCtaCorrente.iTipoMeioPagto
            objAntecipRec.lCliente = tAntecipRec.lCliente
            objAntecipRec.lNumIntRec = tAntecipRec.lNumIntRec
            objAntecipRec.lNumMovto = tAntecipRec.lNumMovto
            objAntecipRec.lSequencial = tMovCtaCorrente.lSequencial
            objAntecipRec.sHistorico = tMovCtaCorrente.sHistorico
            objAntecipRec.sNumRefExterna = tMovCtaCorrente.sNumRefExterna
    
            'Coloca o Recebimento antecipado lido na coleção
            colRecebAntecipado.Add objAntecipRec
            
        End If

        'Lê o próximo Recebimento antecipado
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 29143

    Loop

    'Libera comando
    Call Comando_Fechar(lComando)

    RecebAntecipados_Le_FilialCliente_Baixa_TRV = SUCESSO

    Exit Function

Erro_RecebAntecipados_Le_FilialCliente_Baixa_TRV:

    RecebAntecipados_Le_FilialCliente_Baixa_TRV = Err

    Select Case Err

        Case 29138
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 29139, 29140, 29141, 29143
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RECEB_ANTEC", Err)

        Case 29142

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148630)

    End Select

    'Libera comando
    Call Comando_Fechar(lComando)

    Exit Function

End Function

Function ParcelasRec_Le_BaixaManualSel_TRV(ByVal lCodCli As Long, ByVal iFilialCli As Integer, ByVal dtEmissaoInic As Date, ByVal dtEmissaoFim As Date, ByVal dtVencInic As Date, ByVal dtVencFim As Date, ByVal lTituloInic As Long, ByVal lTituloFim As Long, ByVal dValorDe As Double, ByVal dValorAte As Double, ByVal sTipoDoc As String, ByVal colInfoParcRec As Collection) As Long
'preenche a colecao de parcelas a receber "candidatas" para baixa manual por selecao
'datas nao preenchidas devem ser passadas como DATA_NULA
'numeros de titulo nao preenchidos devem ser passados como zero
'IMPORTANTE: nao estou preenchendo a razao social e o nome reduzido do cliente nos objetos armazenados na colecao
Dim sSelect As String, lErro As Long
Dim lComando As Long, iStatusAberto As Integer
'buffers p/registros pesquisados
Dim lTitulosRecNumTitulo As Long, iTitulosRecFilialEmpresa As Integer, sTitulosRecSiglaDocumento As String
Dim lParcelasRecNumIntDoc As Long, iParcelasRecNumParcela As Integer, dtParcelasRecDataVencimento As Date, dParcelasRecSaldo As Double, dParcelasRecValor As Double, iParcelasRecCobrador As Integer, iParcelasRecCarteiraCobranca As Integer
Dim lCliente As Long
Dim iFilial As Integer, dtDataEmissao As Date

On Error GoTo Erro_ParcelasRec_Le_BaixaManualSel_TRV

    iStatusAberto = STATUS_ABERTO

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 27067

    sTitulosRecSiglaDocumento = String(STRING_SIGLA_DOCUMENTO, 0)

    'p/montagem do comando SQL SELECT dinamicamente
    Call ParcelasRec_Le_BaixaManualSel1(sSelect, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, lCodCli, iFilialCli)

    'executa a preparacao da parte fixa do SELECT
    lErro = ParcelasRec_Le_BaixaManualSel2(lComando, sSelect, lTitulosRecNumTitulo, iTitulosRecFilialEmpresa, sTitulosRecSiglaDocumento, lParcelasRecNumIntDoc, iParcelasRecNumParcela, dtParcelasRecDataVencimento, dParcelasRecSaldo, dParcelasRecValor, iParcelasRecCobrador, iParcelasRecCarteiraCobranca, iStatusAberto, lCliente, iFilial, dtDataEmissao)
    If lErro <> SUCESSO Then Error 27068

    'complementa a passagem dos parametros que variam de acordo com a selecao do usuario
    'e executa o SELECT p/obtencao das parcelas
    lErro = ParcelasRec_Le_BaixaManualSel3(lComando, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, lCodCli, iFilialCli)
    If lErro <> SUCESSO Then Error 27069

    'processa todas as parcelas retornadas pelo SELECT executado acima
    lErro = ParcelasRec_Le_BaixaManualSel_TRV4(lComando, lTitulosRecNumTitulo, iTitulosRecFilialEmpresa, sTitulosRecSiglaDocumento, lParcelasRecNumIntDoc, iParcelasRecNumParcela, dtParcelasRecDataVencimento, dParcelasRecSaldo, dParcelasRecValor, iParcelasRecCobrador, iParcelasRecCarteiraCobranca, lCliente, iFilial, dValorDe, dValorAte, sTipoDoc, colInfoParcRec, dtDataEmissao)
    If lErro <> SUCESSO Then Error 27070
    
    lErro = Comando_Fechar(lComando)

    ParcelasRec_Le_BaixaManualSel_TRV = SUCESSO

    Exit Function

Erro_ParcelasRec_Le_BaixaManualSel_TRV:

    ParcelasRec_Le_BaixaManualSel_TRV = Err

    Select Case Err

        Case 27068, 27069, 27070

        Case 27067
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148589)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Function ParcelasRec_Le_BaixaManualSel2(ByVal lComando As Long, sSelect As String, vlTitulosRecNumTitulo As Variant, viTitulosRecFilialEmpresa As Variant, vsTitulosRecSiglaDocumento As Variant, vlParcelasRecNumIntDoc As Variant, viParcelasRecNumParcela As Variant, vdtParcelasRecDataVencimento As Variant, vdParcelasRecSaldo As Variant, vdParcelasRecValor As Variant, viParcelasRecCobrador As Variant, viParcelasRecCarteiraCobranca As Variant, viStatusAberto As Variant, vlCliente As Variant, viFilial As Variant, vdtDataEmissao As Variant) As Long
'isola a preparacao da parte fixa do SELECT, serve apenas para diminuir o tamanho de ParcelasRec_Le_BaixaManualSel
Dim ret As Integer, lErro As Long
On Error GoTo Erro_ParcelasRec_Le_BaixaManualSel2

    ret = Comando_PrepararInt(lComando, sSelect)
    If (ret <> AD_SQL_SUCESSO) Then Error 27046
    
    ret = Comando_BindVarInt(lComando, vdtDataEmissao)
    If (ret <> AD_SQL_SUCESSO) Then Error 27047

    ret = Comando_BindVarInt(lComando, vlTitulosRecNumTitulo)
    If (ret <> AD_SQL_SUCESSO) Then Error 27047

    ret = Comando_BindVarInt(lComando, viTitulosRecFilialEmpresa)
    If (ret <> AD_SQL_SUCESSO) Then Error 27161
    
    ret = Comando_BindVarInt(lComando, vsTitulosRecSiglaDocumento)
    If (ret <> AD_SQL_SUCESSO) Then Error 29014
    
    ret = Comando_BindVarInt(lComando, vlCliente)
    If (ret <> AD_SQL_SUCESSO) Then Error 27065

    ret = Comando_BindVarInt(lComando, viFilial)
    If (ret <> AD_SQL_SUCESSO) Then Error 27066

    ret = Comando_BindVarInt(lComando, vlParcelasRecNumIntDoc)
    If (ret <> AD_SQL_SUCESSO) Then Error 27048

    ret = Comando_BindVarInt(lComando, viParcelasRecNumParcela)
    If (ret <> AD_SQL_SUCESSO) Then Error 27049

    ret = Comando_BindVarInt(lComando, vdtParcelasRecDataVencimento)
    If (ret <> AD_SQL_SUCESSO) Then Error 27050

    ret = Comando_BindVarInt(lComando, vdParcelasRecSaldo)
    If (ret <> AD_SQL_SUCESSO) Then Error 27051
    
    ret = Comando_BindVarInt(lComando, vdParcelasRecValor)
    If (ret <> AD_SQL_SUCESSO) Then Error 17723
    
    ret = Comando_BindVarInt(lComando, viParcelasRecCobrador)
    If (ret <> AD_SQL_SUCESSO) Then Error 29020

    ret = Comando_BindVarInt(lComando, viParcelasRecCarteiraCobranca)
    If (ret <> AD_SQL_SUCESSO) Then Error 29020

    ret = Comando_BindVarInt(lComando, viStatusAberto)
    If (ret <> AD_SQL_SUCESSO) Then Error 27052
'
'    ret = Comando_BindVarInt(lComando, vlCliente)
'    If (ret <> AD_SQL_SUCESSO) Then Error 27065
'
'    ret = Comando_BindVarInt(lComando, viFilial)
'    If (ret <> AD_SQL_SUCESSO) Then Error 27066

    ParcelasRec_Le_BaixaManualSel2 = SUCESSO

    Exit Function

Erro_ParcelasRec_Le_BaixaManualSel2:

    ParcelasRec_Le_BaixaManualSel2 = Err

    Select Case Err

        Case 17723, 27046 To 27052, 27065, 27066, 27161, 29014, 29020
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148590)

    End Select

    Exit Function

End Function

Public Function ParcelasRec_Le_BaixaManualSel3(ByVal lComando As Long, dtEmissaoInic As Date, dtEmissaoFim As Date, dtVencInic As Date, dtVencFim As Date, lTituloInic As Long, lTituloFim As Long, ByVal lCodCli As Long, ByVal iFilialCli As Integer) As Long
'auxiliar a ParcelasRec_Le_BaixaManualSel
'complementa a passagem dos parametros que variam de acordo com a selecao do usuario e executa o SELECT p/obtencao das parcelas
Dim lErro As Long, ret As Integer, vFilialEmpresa As Variant, vTituloInic As Variant
Dim vTituloFim As Variant, vEmissaoInic As Variant, vEmissaoFim As Variant, vVencInic As Variant, vVencFim As Variant
Dim vlCliente As Variant
Dim viFilial As Variant
Dim vlClienteAux As Variant
Dim vlClienteAux2 As Variant
Dim vlClienteAux3 As Variant

On Error GoTo Erro_ParcelasRec_Le_BaixaManualSel3

    '############################################
    'Inserido por Wagner
    If lCodCli <> 0 Then
    
        vlClienteAux = CLng(0)
        vlClienteAux2 = CLng(0)
        vlClienteAux3 = CLng(0)
        vlCliente = lCodCli
        Select Case lCodCli
        
            Case 109660
                'vlClienteAux = CLng(1120961)
                vlCliente = CLng(109664)
                vlClienteAux = CLng(1120965)
                vlClienteAux2 = CLng(109662)
                vlClienteAux3 = CLng(1120963)
            Case 1120961
                'vlClienteAux = CLng(109660)
                vlCliente = CLng(109664)
                vlClienteAux = CLng(1120965)
                vlClienteAux2 = CLng(109662)
                vlClienteAux3 = CLng(1120963)
            Case 109664
                'vlClienteAux = CLng(1120965)
                'vlClienteAux2 = CLng(114654)
                'vlClienteAux3 = CLng(114652)
                vlCliente = CLng(109660)
                vlClienteAux = CLng(1120961)
                vlClienteAux2 = CLng(109662)
                vlClienteAux3 = CLng(1120963)
            Case 1120965
                'vlClienteAux = CLng(109664)
                'vlClienteAux2 = CLng(114654)
                'vlClienteAux3 = CLng(114652)
                vlCliente = CLng(109660)
                vlClienteAux = CLng(1120961)
                vlClienteAux2 = CLng(109662)
                vlClienteAux3 = CLng(1120963)
            Case 109662
                'vlClienteAux = CLng(1120963)
                vlCliente = CLng(109664)
                vlClienteAux = CLng(1120965)
                vlClienteAux2 = CLng(109660)
                vlClienteAux3 = CLng(1120961)
            Case 1120963
                'vlClienteAux = CLng(109662)
                vlCliente = CLng(109664)
                vlClienteAux = CLng(1120965)
                vlClienteAux2 = CLng(109660)
                vlClienteAux3 = CLng(1120961)
        
        End Select

        ret = Comando_BindVarInt(lComando, vlCliente)
        If (ret <> AD_SQL_SUCESSO) Then gError 140287
        
        If vlClienteAux <> 0 Then
            ret = Comando_BindVarInt(lComando, vlClienteAux)
            If (ret <> AD_SQL_SUCESSO) Then gError 140287
        End If
        If vlClienteAux2 <> 0 Then
            ret = Comando_BindVarInt(lComando, vlClienteAux2)
            If (ret <> AD_SQL_SUCESSO) Then gError 140287
        End If
        If vlClienteAux3 <> 0 Then
            ret = Comando_BindVarInt(lComando, vlClienteAux3)
            If (ret <> AD_SQL_SUCESSO) Then gError 140287
        End If
    End If
    
    If iFilialCli <> 0 Then
        viFilial = iFilialCli
        ret = Comando_BindVarInt(lComando, viFilial)
        If (ret <> AD_SQL_SUCESSO) Then gError 140288
    End If
    '############################################
    
    'Se entrar como empresa toda pode pegar titulos de todas as filiais, senao pegar titulos somente da filial corrente
    If giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        vFilialEmpresa = giFilialEmpresa
        ret = Comando_BindVarInt(lComando, vFilialEmpresa)
        If (ret <> AD_SQL_SUCESSO) Then gError 27054
    End If

    'Se titulo inicial preenchido
    If (lTituloInic <> 0) Then
        vTituloInic = lTituloInic
        ret = Comando_BindVarInt(lComando, vTituloInic)
        If (ret <> AD_SQL_SUCESSO) Then gError 27055
    End If

    'Se titulo final preenchido
    If (lTituloFim <> 0) Then
        vTituloFim = lTituloFim
        ret = Comando_BindVarInt(lComando, vTituloFim)
        If (ret <> AD_SQL_SUCESSO) Then gError 27056
    End If

    'Se o limite inicial de data de emissao de titulo estiver preenchido
    If (dtEmissaoInic <> DATA_NULA) Then
        vEmissaoInic = dtEmissaoInic
        ret = Comando_BindVarInt(lComando, vEmissaoInic)
        If (ret <> AD_SQL_SUCESSO) Then gError 27057
    End If

    'Se o limite final de data de emissao de titulo estiver preenchido
    If (dtEmissaoFim <> DATA_NULA) Then
        vEmissaoFim = dtEmissaoFim
        ret = Comando_BindVarInt(lComando, vEmissaoFim)
        If (ret <> AD_SQL_SUCESSO) Then gError 27058
    End If

    'Se o limite inicial de data de vencimento de parcela estiver preenchido
    If (dtVencInic <> DATA_NULA) Then
        vVencInic = dtVencInic
        ret = Comando_BindVarInt(lComando, vVencInic)
        If (ret <> AD_SQL_SUCESSO) Then gError 27059
    End If

    'Se o limite final de data de vencimento de parcela estiver preenchido
    If (dtVencFim <> DATA_NULA) Then
        vVencFim = dtVencFim
        ret = Comando_BindVarInt(lComando, vVencFim)
        If (ret <> AD_SQL_SUCESSO) Then gError 27060
    End If

    ret = Comando_ExecutarInt(lComando)
    If (ret <> AD_SQL_SUCESSO) Then gError 27061

    ParcelasRec_Le_BaixaManualSel3 = SUCESSO

    Exit Function
    
Erro_ParcelasRec_Le_BaixaManualSel3:

    ParcelasRec_Le_BaixaManualSel3 = gErr

    Select Case gErr

        Case 27054 To 27061
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", gErr)
            
        Case 140287, 140288

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 148591)

    End Select

    Exit Function

End Function

Private Function ParcelasRec_Le_BaixaManualSel_TRV4(ByVal lComando As Long, lTitulosRecNumTitulo As Long, iTitulosRecFilialEmpresa As Integer, sTitulosRecSiglaDocumento As String, lParcelasRecNumIntDoc As Long, iParcelasRecNumParcela As Integer, dtParcelasRecDataVencimento As Date, dParcelasRecSaldo As Double, dParcelasRecValor As Double, iParcelasRecCobrador As Integer, iParcelasRecCarteiraCobranca As Integer, lCliente As Long, iFilial As Integer, ByVal dValorDe As Double, ByVal dValorAte As Double, ByVal sTipoDoc As String, ByVal colInfoParcRec As Collection, dtDataEmissao As Date) As Long
'auxiliar a ParcelasRec_Le_BaixaManualSel
'processa os registros lidos
Dim lErro As Long
Dim objInfoParcRec As ClassInfoParcRec
Dim bValido As Boolean

On Error GoTo Erro_ParcelasRec_Le_BaixaManualSel_TRV4

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27062

    If lErro = AD_SQL_SEM_DADOS Then Error 27063

    Do While lErro = AD_SQL_SUCESSO
    
        bValido = True
        
        If dValorDe <> 0 Then
            If dParcelasRecSaldo - dValorDe < QTDE_ESTOQUE_DELTA Then bValido = False
        End If
        
        If dValorAte <> 0 Then
            If dParcelasRecSaldo - dValorAte > QTDE_ESTOQUE_DELTA Then bValido = False
        End If
        
        If Len(Trim(sTipoDoc)) <> 0 Then
            If sTitulosRecSiglaDocumento <> sTipoDoc Then bValido = False
        End If
        
        If bValido Then

            Set objInfoParcRec = New ClassInfoParcRec
    
            objInfoParcRec.iFilialEmpresa = iTitulosRecFilialEmpresa
            objInfoParcRec.dtVencimento = dtParcelasRecDataVencimento
            objInfoParcRec.dValor = dParcelasRecSaldo
            objInfoParcRec.dValorOriginal = dParcelasRecValor
            objInfoParcRec.iFilialCliente = iFilial
            objInfoParcRec.iNumParcela = iParcelasRecNumParcela
            objInfoParcRec.lCliente = lCliente
            objInfoParcRec.lNumIntParc = lParcelasRecNumIntDoc
            objInfoParcRec.lNumTitulo = lTitulosRecNumTitulo
            objInfoParcRec.sSiglaDocumento = sTitulosRecSiglaDocumento
            objInfoParcRec.iCobrador = iParcelasRecCobrador
            objInfoParcRec.iCarteiraCobrador = iParcelasRecCarteiraCobranca
            'objInfoParcRec.sNomeRedCliente = sClientesNomeReduzido
            'objInfoParcRec.sRazaoSocialCli = sClientesRazaoSocial
            objInfoParcRec.dtDataEmissao = dtDataEmissao
            
            colInfoParcRec.Add objInfoParcRec
            
        End If
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 27064

    Loop

    ParcelasRec_Le_BaixaManualSel_TRV4 = SUCESSO
    
    Exit Function

Erro_ParcelasRec_Le_BaixaManualSel_TRV4:

    ParcelasRec_Le_BaixaManualSel_TRV4 = Err

    Select Case Err

        Case 27063
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_PARCELAS_REC_SEL", Err)

        Case 27062, 27064
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELAS_REC", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148592)

    End Select

    Exit Function

End Function
'######################################################################

Function RegiaoVenda_Le_Customizado(ByVal objRegiaoVenda As ClassRegiaoVenda) As Long
    
Dim lErro As Long
Dim objRegiaoVendaTRV As New ClassRegiaoVendaTRV
Dim iCodRegiaoPai As Integer
Dim alComando(1 To 1) As Long
Dim iIndice As Integer

On Error GoTo Erro_RegiaoVenda_Le_Customizado

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 197187
    Next

    'Alocação de espaço no buffer

    'Le a tabela VendedorTRV
    lErro = Comando_Executar(alComando(1), "SELECT CodRegiaoPai FROM TRVRegiaoVenda WHERE Codigo = ? ", iCodRegiaoPai, objRegiaoVenda.iCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 197188

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 197189

    objRegiaoVendaTRV.iCodRegiaoPai = iCodRegiaoPai
    
    Set objRegiaoVenda.objInfoUsu = objRegiaoVendaTRV

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    RegiaoVenda_Le_Customizado = SUCESSO

    Exit Function

Erro_RegiaoVenda_Le_Customizado:

    RegiaoVenda_Le_Customizado = gErr

    Select Case gErr

        Case 197187
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 197188, 197189
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TRVREGIAOVENDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197190)

    End Select

    'Fecha Comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function
    
End Function

Public Function CobrancaPorEmail_NFsPag_Le(ByVal objCobrancaSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String
Dim lNumIntDocNFPag As Long
Dim objNfsPag As ClassNFsPag
Dim colTipoDoc As New Collection
Dim objCobrSelTRV As ClassCobrPorEmailSelTRV

On Error GoTo Erro_CobrancaPorEmail_NFsPag_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185991
    
    If Not (objCobrancaSel.objInfoUsu Is Nothing) Then
        Set objCobrSelTRV = objCobrancaSel.objInfoUsu
        Set colTipoDoc = objCobrSelTRV.colTiposDoc
    End If
    
    Call CobrancaPorEmail_NFsPag_PreparaSQL(objCobrancaSel, sSQL, colTipoDoc)

    lErro = CobrancaPorEmail_NFsPag_PreparaInt(lComando, lNumIntDocNFPag, _
                        objCobrancaSel.dtDataEmissaoAte, objCobrancaSel.dtDataEmissaoDe, _
                        objCobrancaSel.dSaldoAte, objCobrancaSel.dSaldoDe, _
                        objCobrancaSel.lClienteAte, objCobrancaSel.lClienteDe, _
                        objCobrancaSel.lTituloAte, objCobrancaSel.lTituloDe, STATUS_BAIXADO, STATUS_EXCLUIDO, sSQL, colTipoDoc)
    If lErro <> SUCESSO Then gError 185992

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185993
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objNfsPag = New ClassNFsPag
    
        objNfsPag.lNumIntDoc = lNumIntDocNFPag
    
        lErro = CF("NFPag_Le_Todos2", objNfsPag, MARCADO)
        If lErro <> SUCESSO And lErro <> 18257 Then gError 185994
    
        objCobrancaSel.colParcelas.Add objNfsPag
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185995
    
    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    CobrancaPorEmail_NFsPag_Le = SUCESSO

    Exit Function

Erro_CobrancaPorEmail_NFsPag_Le:

    CobrancaPorEmail_NFsPag_Le = gErr

    Select Case gErr

        Case 185991
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 185992, 185994

        Case 185993, 185995
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185996)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)


    Exit Function

End Function

Private Function CobrancaPorEmail_NFsPag_PreparaSQL(ByVal objCobrancaSel As ClassCobrancaPorEmailSel, sSQL As String, ByVal colTipoDoc As Collection) As Long

Dim lErro As Long
Dim sOrderBy As String
Dim sWhere As String
Dim sWhereAux As String
Dim vValor As Variant

On Error GoTo Erro_CobrancaPorEmail_NFsPag_PreparaSQL

    sSQL = "SELECT NumIntDoc FROM NFsPag "

    sWhere = "WHERE Status <> ? AND Status <> ? AND ValorTotal > 0 AND NumIntTitPag = 0 "

    If objCobrancaSel.dSaldoAte <> 0 Then
        sWhere = sWhere & "AND Valor <= ? "
    End If
    
    If objCobrancaSel.dSaldoDe <> 0 Then
        sWhere = sWhere & "AND Valor >= ? "
    End If
    
    If objCobrancaSel.lClienteAte <> 0 Then
        sWhere = sWhere & "AND Fornecedor <= ? "
    End If
    
    If objCobrancaSel.lClienteDe <> 0 Then
        sWhere = sWhere & "AND Fornecedor >= ? "
    End If
    
    If objCobrancaSel.lTituloAte <> 0 Then
        sWhere = sWhere & "AND NumNotaFiscal <= ? "
    End If
    
    If objCobrancaSel.lTituloDe <> 0 Then
        sWhere = sWhere & "AND NumNotaFiscal >= ? "
    End If
    
    If objCobrancaSel.dtDataEmissaoAte <> DATA_NULA Then
        sWhere = sWhere & "AND DataVencimento <= ? "
    End If
    
    If objCobrancaSel.dtDataEmissaoDe <> DATA_NULA Then
        sWhere = sWhere & "AND DataVencimento >= ? "
    End If
    
    sWhereAux = ""
    
    sWhereAux = " AND ( (1 <> 1) "
       
    For Each vValor In colTipoDoc
        If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " OR "
        If vValor <> 0 Then
            sWhereAux = sWhereAux & " (EXISTS (SELECT D.NumIntDoc FROM DocFaturadosRapido AS D WHERE D.TipoDocDestino = 5 AND D.NumIntDocDestino = NFsPag.NumIntDoc AND D.TipoDoc = ?)) "
            If vValor = TRV_TIPODOC_CMCC Then
                sWhereAux = sWhereAux & " OR (EXISTS (SELECT T.NumIntDocTitRec FROM TitulosRecTRV AS T WHERE T.NumIntDocNFPagComi = NfsPag.NumIntDoc)) "
                sWhereAux = sWhereAux & " OR (EXISTS (SELECT C.NumIntDocCorporator FROM CoinfoComissao AS C WHERE C.NumIntDocCorporator = NfsPag.NumIntDoc )) "
            End If
        Else
            sWhereAux = sWhereAux & " ((NOT EXISTS (SELECT D.NumIntDoc FROM DocFaturadosRapido AS D WHERE D.TipoDocDestino = 5 AND D.NumIntDocDestino = NFsPag.NumIntDoc)) AND ( NOT EXISTS (SELECT T.NumIntDocTitRec FROM TitulosRecTRV AS T WHERE T.NumIntDocNFPagComi = NfsPag.NumIntDoc))) "
        End If
    Next
    
    sWhere = sWhere & sWhereAux & " ) "
    
    sOrderBy = "ORDER BY NumNotaFiscal, DataEmissao "
    
    sSQL = sSQL & sWhere & sOrderBy

    CobrancaPorEmail_NFsPag_PreparaSQL = SUCESSO

    Exit Function

Erro_CobrancaPorEmail_NFsPag_PreparaSQL:

    CobrancaPorEmail_NFsPag_PreparaSQL = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185975)

    End Select


    Exit Function

End Function

Private Function CobrancaPorEmail_NFsPag_PreparaInt(ByVal lComando As Long, vlNumIntDoc As Variant, ByVal vdtDataAte As Variant, ByVal vdtDataDe As Variant, ByVal vlFornecedorAte As Variant, ByVal vlFornecedorDe As Variant, ByVal vdValorAte As Variant, ByVal vdValorDe As Variant, ByVal vlNumNFAte As Variant, ByVal vlNumNFDe As Variant, ByVal viStatus1 As Variant, ByVal viStatus2 As Variant, ByVal sSQL As String, ByVal colTipoDoc As Collection) As Long
'Passa as variáveis de retorno e filtro do SELECT que monta os dados do Grid
 
Dim lErro As Long
Dim iIndice As Integer
Dim vValor(1 To 100) As Variant
Dim vValorAux As Variant

On Error GoTo Erro_CobrancaPorEmail_NFsPag_PreparaInt
 
    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185976
                       
    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185977

    lErro = Comando_BindVarInt(lComando, viStatus1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185978

    lErro = Comando_BindVarInt(lComando, viStatus2)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185979
  
    If vdValorAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vdValorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185982
    End If
    If vdValorDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vdValorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185983
    End If
    If vlFornecedorAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185984
    End If
    If vlFornecedorDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185985
    End If
    If vlNumNFAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlNumNFAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185986
    End If
    If vlNumNFDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlNumNFDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185987
    End If
    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185980
    End If
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185981
    End If
    
    iIndice = 0
    For Each vValorAux In colTipoDoc
        If vValorAux <> 0 Then
                   
            iIndice = iIndice + 1
                    
            Select Case vValorAux
                
                Case TRV_TIPODOC_OCR
                    vValor(iIndice) = TRV_TIPODOC_OCR_TEXTO
                
                Case TRV_TIPODOC_NVL
                    vValor(iIndice) = TRV_TIPODOC_NVL_TEXTO
                
                Case TRV_TIPODOC_VOU
                    vValor(iIndice) = TRV_TIPODOC_VOU_TEXTO
                
                Case TRV_TIPODOC_CMC
                    vValor(iIndice) = TRV_TIPODOC_CMC_TEXTO
                
                Case TRV_TIPODOC_CMR
                    vValor(iIndice) = TRV_TIPODOC_CMR_TEXTO
                
                Case TRV_TIPODOC_OVER
                    vValor(iIndice) = TRV_TIPODOC_OVER_TEXTO
                
                Case TRV_TIPODOC_CMCC
                    vValor(iIndice) = TRV_TIPODOC_CMCC_TEXTO
                
            End Select
        
            lErro = Comando_BindVarInt(lComando, vValor(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then gError 185981
        
        End If
    Next
    
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185989
 
   CobrancaPorEmail_NFsPag_PreparaInt = SUCESSO
 
    Exit Function
 
Erro_CobrancaPorEmail_NFsPag_PreparaInt:
 
    CobrancaPorEmail_NFsPag_PreparaInt = gErr
 
    Select Case gErr
 
        Case 185976 To 185989
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185990)
 
    End Select
 
    Exit Function
 
End Function

Public Function CobrancaPorEmail_ParcPagBaixados_Le(ByVal objCobrancaSel As ClassCobrancaPorEmailSel) As Long

Dim lErro As Long
Dim lComando As Long
Dim sSQL As String
Dim lNumIntDocParcela As Long
Dim objParcelaPag As ClassParcelaPagar
Dim dtDataIni As Date
Dim dtDataFim As Date
Dim dtDataUltimaBaixa As Date
Dim colTipoDoc As New Collection
Dim objCobrSelTRV As ClassCobrPorEmailSelTRV

On Error GoTo Erro_CobrancaPorEmail_ParcPagBaixados_Le

    'Executa a abertura do Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 185991
    
    If Not (objCobrancaSel.objInfoUsu Is Nothing) Then
        Set objCobrSelTRV = objCobrancaSel.objInfoUsu
        Set colTipoDoc = objCobrSelTRV.colTiposDoc
    End If
    
    Call CobrancaPorEmail_ParcPagBaixados_PreparaSQL(objCobrancaSel, sSQL, colTipoDoc)

    lErro = CobrancaPorEmail_ParcPagBaixados_PreparaInt(lComando, lNumIntDocParcela, dtDataUltimaBaixa, _
                        objCobrancaSel.dtDataBaixaAte, objCobrancaSel.dtDataBaixaDe, _
                        objCobrancaSel.lClienteAte, objCobrancaSel.lClienteDe, _
                        objCobrancaSel.lTituloAte, objCobrancaSel.lTituloDe, objCobrancaSel.sTipo, STATUS_EXCLUIDO, _
                        sSQL, colTipoDoc)
    If lErro <> SUCESSO Then gError 185992

    'Busca Primeiro
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185993
    
    Do While lErro <> AD_SQL_SEM_DADOS
    
        Set objParcelaPag = New ClassParcelaPagar
    
        objParcelaPag.lNumIntDoc = lNumIntDocParcela
    
        lErro = CF("ParcelaPagar_Le", objParcelaPag)
        If lErro <> SUCESSO And lErro <> 60479 Then gError 185994
        
        If lErro <> SUCESSO Then
        
            lErro = CF("ParcelaPagarBaixada_Le", objParcelaPag)
            If lErro <> SUCESSO And lErro <> 123509 Then gError 185994
        
        End If
    
        objParcelaPag.dtDataUltimaBaixa = dtDataUltimaBaixa
    
        objCobrancaSel.colParcelas.Add objParcelaPag
    
        'Busca Próximo
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 185995
    
    Loop
    
    'Fecha Comando
    Call Comando_Fechar(lComando)

    CobrancaPorEmail_ParcPagBaixados_Le = SUCESSO

    Exit Function

Erro_CobrancaPorEmail_ParcPagBaixados_Le:

    CobrancaPorEmail_ParcPagBaixados_Le = gErr

    Select Case gErr

        Case 185991
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 185992, 185994

        Case 185993, 185995
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_PARCELASREC", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185996)

    End Select
    
    'Fecha Comando
    Call Comando_Fechar(lComando)


    Exit Function

End Function

Private Function CobrancaPorEmail_ParcPagBaixados_PreparaSQL(ByVal objCobrancaSel As ClassCobrancaPorEmailSel, sSQL As String, ByVal colTipoDoc As Collection) As Long

Dim lErro As Long
Dim sOrderBy As String
Dim sWhere As String
Dim sWhereAux As String
Dim vValor As Variant

On Error GoTo Erro_CobrancaPorEmail_ParcPagBaixados_PreparaSQL

    sSQL = "SELECT DISTINCT P.NumIntDoc, PB.DataBaixa FROM ParcelasPagTodas P, TitulosPagTodos T, BaixasParcPag as BP, BaixasPag as B, ParcelasPagDataBaixa AS PB "

    sWhere = "WHERE P.Status <> ? AND P.NumIntTitulo = T.NumIntDoc AND BP.NumIntParcela = P. NumIntDoc AND P.Saldo = 0 AND BP.NumIntBaixa = B.NumIntBaixa AND PB.NumIntDoc = P.NumIntDoc "
                    
    If objCobrancaSel.dtDataBaixaAte <> DATA_NULA Then
        sWhere = sWhere & "AND PB.DataBaixa <= ? "
    End If
    
    If objCobrancaSel.dtDataBaixaDe <> DATA_NULA Then
        sWhere = sWhere & "AND PB.DataBaixa >= ? "
    End If
   
    If objCobrancaSel.lClienteAte <> 0 Then
        sWhere = sWhere & "AND T.Fornecedor <= ? "
    End If
    
    If objCobrancaSel.lClienteDe <> 0 Then
        sWhere = sWhere & "AND T.Fornecedor >= ? "
    End If
    
    If objCobrancaSel.lTituloAte <> 0 Then
        sWhere = sWhere & "AND T.NumTitulo <= ? "
    End If
    
    If objCobrancaSel.lTituloDe <> 0 Then
        sWhere = sWhere & "AND T.NumTitulo >= ? "
    End If

    If Len(Trim(objCobrancaSel.sTipo)) <> 0 Then
        sWhere = sWhere & "AND T.SiglaDocumento = ? "
    End If
    
    sWhereAux = " AND ( (1 <> 1) "
    
    For Each vValor In colTipoDoc
        If Len(Trim(sWhereAux)) > 0 Then sWhereAux = sWhereAux & " OR "
        If vValor <> 0 Then
            If vValor = TRV_TIPODOC_CMCC Or vValor = TRV_TIPODOC_CMR Or vValor = TRV_TIPODOC_CMC Then
                sWhereAux = sWhereAux & " (EXISTS (SELECT D.NumIntDoc FROM DocFaturadosRapido AS D, NFsPag_Todas AS N WHERE D.TipoDocDestino = 5 AND D.NumIntDocDestino = N.NumIntDoc AND N.NumIntTitPag = T.NumIntDoc AND D.TipoDoc = ?)) "
                If vValor = TRV_TIPODOC_CMCC Then
                    sWhereAux = sWhereAux & " OR  (EXISTS (SELECT T1.NumIntDocTitRec FROM TitulosRecTRV AS T1, NFsPag_Todas AS N WHERE T1.NumIntDocNFPagComi = N.NumIntDoc AND T1.NumIntDocTitRec = T.NumIntDoc)) "
                    sWhereAux = sWhereAux & " OR  (EXISTS (SELECT N.NumIntDoc FROM CoinfoComissao AS C, NFsPag_Todas AS N WHERE C.NumIntDocCorporator = N.NumIntDoc AND N.NumIntTitPag = T.NumIntDoc)) "
                End If
            Else
                sWhereAux = sWhereAux & " (EXISTS (SELECT D.NumIntDoc FROM DocFaturadosRapido AS D WHERE D.TipoDocDestino = 4 AND D.NumIntDocDestino = T.NumIntDoc AND D.TipoDoc = ?)) "
                If vValor = TRV_TIPODOC_OVER Then
                    sWhereAux = sWhereAux & " OR (SiglaDocumento = 'OVER')"
                End If
            End If
        Else
            sWhereAux = sWhereAux & " (SiglaDocumento NOT IN ('FP', 'FSCP', 'OVER'))"
        End If
    Next
    
    sWhere = sWhere & sWhereAux & " ) "

    sOrderBy = "ORDER BY P.NumIntDoc "
    
    sSQL = sSQL & sWhere & sOrderBy

    CobrancaPorEmail_ParcPagBaixados_PreparaSQL = SUCESSO

    Exit Function

Erro_CobrancaPorEmail_ParcPagBaixados_PreparaSQL:

    CobrancaPorEmail_ParcPagBaixados_PreparaSQL = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185975)

    End Select


    Exit Function

End Function

Private Function CobrancaPorEmail_ParcPagBaixados_PreparaInt(ByVal lComando As Long, vlNumIntDoc As Variant, vdtDataUltimaBaixa As Variant, ByVal vdtDataAte As Variant, ByVal vdtDataDe As Variant, ByVal vlFornecedorAte As Variant, ByVal vlFornecedorDe As Variant, ByVal vlTituloAte As Variant, ByVal vlTituloDe As Variant, ByVal vsTipo As Variant, ByVal viStatus1 As Variant, ByVal sSQL As String, ByVal colTipoDoc As Collection) As Long
'Passa as variáveis de retorno e filtro do SELECT que monta os dados do Grid
 
Dim lErro As Long
Dim iIndice As Integer
Dim vValor(1 To 100) As Variant
Dim vValorAux As Variant

On Error GoTo Erro_CobrancaPorEmail_ParcPagBaixados_PreparaInt
 
    lErro = Comando_PrepararInt(lComando, sSQL)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185976
                       
    lErro = Comando_BindVarInt(lComando, vlNumIntDoc)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185977

    lErro = Comando_BindVarInt(lComando, vdtDataUltimaBaixa)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185977

    lErro = Comando_BindVarInt(lComando, viStatus1)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185978

    If vdtDataAte <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185981
    End If
    
    If vdtDataDe <> DATA_NULA Then
        lErro = Comando_BindVarInt(lComando, vdtDataDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185980
   End If
   
    If vlFornecedorAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185984
    End If
    If vlFornecedorDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlFornecedorDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185985
    End If
    If vlTituloAte <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlTituloAte)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185986
    End If
    If vlTituloDe <> 0 Then
        lErro = Comando_BindVarInt(lComando, vlTituloDe)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185987
    End If
    If Len(Trim(vsTipo)) <> 0 Then
        lErro = Comando_BindVarInt(lComando, vsTipo)
        If (lErro <> AD_SQL_SUCESSO) Then gError 185988
    End If
    
    iIndice = 0
    For Each vValorAux In colTipoDoc
        If vValorAux <> 0 Then
                   
            iIndice = iIndice + 1
                    
            Select Case vValorAux
                
                Case TRV_TIPODOC_OCR
                    vValor(iIndice) = TRV_TIPODOC_OCR_TEXTO
                
                Case TRV_TIPODOC_NVL
                    vValor(iIndice) = TRV_TIPODOC_NVL_TEXTO
                
                Case TRV_TIPODOC_VOU
                    vValor(iIndice) = TRV_TIPODOC_VOU_TEXTO
                
                Case TRV_TIPODOC_CMC
                    vValor(iIndice) = TRV_TIPODOC_CMC_TEXTO
                
                Case TRV_TIPODOC_CMR
                    vValor(iIndice) = TRV_TIPODOC_CMR_TEXTO
                
                Case TRV_TIPODOC_OVER
                    vValor(iIndice) = TRV_TIPODOC_OVER_TEXTO
                
                Case TRV_TIPODOC_CMCC
                    vValor(iIndice) = TRV_TIPODOC_CMCC_TEXTO
                
            End Select
        
            lErro = Comando_BindVarInt(lComando, vValor(iIndice))
            If (lErro <> AD_SQL_SUCESSO) Then gError 185981
        
        End If
    Next
     
    lErro = Comando_ExecutarInt(lComando)
    If (lErro <> AD_SQL_SUCESSO) Then gError 185989
 
   CobrancaPorEmail_ParcPagBaixados_PreparaInt = SUCESSO
 
    Exit Function
 
Erro_CobrancaPorEmail_ParcPagBaixados_PreparaInt:
 
    CobrancaPorEmail_ParcPagBaixados_PreparaInt = gErr
 
    Select Case gErr
 
        Case 185976 To 185989
            Call Rotina_Erro(vbOKOnly, "ERRO_PREPARACAO_COMANDO_SQL", gErr, sSQL)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 185990)
 
    End Select
 
    Exit Function
 
End Function

Public Function Fatura_Preenche_Cobranca_Cust(ByVal objGridParcelas As AdmGrid, ByVal objComboTipoCobr As Object, ByVal iLinha As Integer, ByVal iColuna As Integer, ByVal sNomeRedForn As String)
   
Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iTipoCobranca As Integer, sConteudo As String

Dim Erro_Fatura_Preenche_Cobranca_Cust:

    If Len(Trim(sNomeRedForn)) > 0 Then
    
        objFornecedor.sNomeReduzido = sNomeRedForn
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 198694
        
        If lErro <> SUCESSO Then gError 198695
        
        lErro = CF("TRVConfig_Le", TRVCONFIG_BANCO_PADRAO_PAGTO, EMPRESA_TODA, sConteudo)
        If lErro <> SUCESSO Then gError 192369
        
        If objFornecedor.iBanco <> StrParaInt(sConteudo) Then
            iTipoCobranca = TIPO_COBRANCA_DOC
        Else
            iTipoCobranca = TIPO_COBRANCA_DEP_CONTA
        End If
        
        Call Combo_Seleciona_ItemData(objComboTipoCobr, iTipoCobranca)
        
        objGridParcelas.objGrid.TextMatrix(iLinha, iColuna) = objComboTipoCobr.Text
        
    End If

    Fatura_Preenche_Cobranca_Cust = SUCESSO
 
    Exit Function
 
Erro_Fatura_Preenche_Cobranca_Cust:
 
    Fatura_Preenche_Cobranca_Cust = gErr
 
    Select Case gErr
 
        Case 198694
        
        Case 198695
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198696)
 
    End Select

End Function

Public Function ContatoCli_Selecao_Le_Endereco(objFilial As ClassFilialCliente, lEndereco As Long)
    lEndereco = objFilial.lEnderecoEntrega
    ContatoCli_Selecao_Le_Endereco = SUCESSO
End Function

Public Function ContatoCli_Selecao_PreparaSQL(ByVal objContatoCliSel As ClassContatoCliSel, sSQL As String) As Long

Dim lErro As Long
Dim sWhere As String
Dim sOrderBy As String
Dim sSubWhere As String
Dim vValor As Variant

On Error GoTo Erro_ContatoCli_Selecao_PreparaSQL

    sSQL = "SELECT C.CodCliente, C.CodFilial, C.DataUltimaCompra, C.DataPrimeiraCompra, C.ValorTotalComprado, C.DataUltimoContato FROM ContatoCli_Selecao AS C "
    sOrderBy = "ORDER BY C.CodCliente, C.CodFilial "
    
    If objContatoCliSel.iMarca = TRV_EMPRESA_AMBOS Then
        sWhere = "WHERE 1=1 "
    ElseIf objContatoCliSel.iMarca = TRV_EMPRESA_MY Then
        sWhere = "WHERE C.CodCliente >= 1000000 "
    Else
        sWhere = "WHERE C.CodCliente < 1000000 "
    End If
                
    If objContatoCliSel.sRespCallCenter <> "" Then
        sWhere = sWhere & "AND C.RespCallCenter = ? "
    End If
    
    If objContatoCliSel.dtDataPCompDe <> DATA_NULA Then
        sWhere = sWhere & "AND C.DataPrimeiraCompra >= ? "
    End If
    
    If objContatoCliSel.dtDataPCompAte <> DATA_NULA Then
        sWhere = sWhere & "AND (C.DataPrimeiraCompra <= ? OR C.DataPrimeiraCompra IS NULL)"
    End If

    If objContatoCliSel.dtDataUCompDe <> DATA_NULA Then
        sWhere = sWhere & "AND C.DataUltimaCompra >= ? "
    End If
    
    If objContatoCliSel.dtDataUCompAte <> DATA_NULA Then
        sWhere = sWhere & "AND (C.DataUltimaCompra <= ? OR C.DataUltimaCompra IS NULL)"
    End If
    
    If objContatoCliSel.dtDataContDe <> DATA_NULA Then
        sWhere = sWhere & "AND C.DataUltimoContato >= ? "
    End If
    
    If objContatoCliSel.dtDataContAte <> DATA_NULA Then
        sWhere = sWhere & "AND (C.DataUltimoContato <= ? OR C.DataUltimoContato IS NULL) "
    End If
        
    sSubWhere = ""
    If Len(Trim(objContatoCliSel.sCategoria)) > 0 Then
        If Len(Trim(objContatoCliSel.sCategoriaDe)) > 0 Then
            sSubWhere = sSubWhere & "AND F.Item >= ? "
        End If
        If Len(Trim(objContatoCliSel.sCategoriaAte)) > 0 Then
            sSubWhere = sSubWhere & "AND F.Item <= ? "
        End If
        sWhere = sWhere & " AND EXISTS ( SELECT F.Cliente FROM FilialClienteCategorias AS F WHERE C.CodCliente= F.Cliente AND C.CodFilial = F.Filial AND " & _
        "F.Categoria = ? " & sSubWhere & ") "
    End If
    
    For Each vValor In objContatoCliSel.colTiposNaoConsiderar
        sWhere = sWhere & "AND C.TipoCliente <> ? "
    Next
    
    sSubWhere = ""
    If objContatoCliSel.dtDataPContDe <> DATA_NULA Then
        sSubWhere = sSubWhere & "AND R.DataProxCobr >= ? "
    End If
    
    If objContatoCliSel.dtDataPContAte <> DATA_NULA Then
        sSubWhere = sSubWhere & "AND R.DataProxCobr <= ? "
    End If
    
    If Len(Trim(sSubWhere)) > 0 Then
        sWhere = sWhere & "AND EXISTS (SELECT R.Codigo FROM RelacionamentoClientes AS R WHERE R.Cliente = C.CodCliente " & sSubWhere & " ) "
    End If
    
    sSubWhere = ""
    If objContatoCliSel.dtDataNCompDe <> DATA_NULA Then
        sSubWhere = sSubWhere & "AND U.Data >= ? "
    End If
    
    If objContatoCliSel.dtDataNCompAte <> DATA_NULA Then
        sSubWhere = sSubWhere & "AND U.Data <= ? "
    End If
    
    If Len(Trim(sSubWhere)) > 0 Then
        sWhere = sWhere & "AND NOT EXISTS (SELECT U.Data FROM ContatoCli_Compras AS U WHERE U.Cliente = C.CodCliente " & sSubWhere & " ) "
    End If
    
    If objContatoCliSel.iSoAtivos = MARCADO Then
        sWhere = sWhere & "AND C.Ativo = ?  "
    End If
            
    sSQL = sSQL & sWhere & sOrderBy
    
    ContatoCli_Selecao_PreparaSQL = SUCESSO
     
    Exit Function
    
Erro_ContatoCli_Selecao_PreparaSQL:

    ContatoCli_Selecao_PreparaSQL = gErr
     
    Select Case gErr

        Case 200225

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200011)
     
    End Select
     
    Exit Function

End Function

Function NfsPag_Le_FilialFornecedor_Desvinculadas(ByVal objFilialFornecedor As ClassFilialFornecedor, ByVal colNFPagAberta As ColNFsPag) As Long
'Lê em NfsPag Notas Fiscais Vinculadas à Filial de Fornecedor passada
'mas desvinculadas de qualquer Titulo a Pagar.

Dim lErro As Long
Dim tNFPag As typeNFsPag
Dim alComando(1 To 3) As Long
Dim sFiltro As String
Dim iFilialEmpresa As Integer
Dim iFilialFornec As Integer
Dim bTodosOVER As Boolean
Dim lFornAux As Long
Dim iIndice As Integer

On Error GoTo Erro_NfsPag_Le_FilialFornecedor_Desvinculadas

    'Abre o comando
    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 202158
    Next

    Call CF("FilialEmpresa_Le_Filtro_Filial", sFiltro)
    
    bTodosOVER = False
    
    lErro = Comando_Executar(alComando(1), "SELECT Fornecedor FROM FilialFornecedorCategorias WHERE Fornecedor = ? AND Categoria = ? AND Item = ?", lFornAux, objFilialFornecedor.lCodFornecedor, TRV_CATEGORIA_FATURAOVER, TRV_CATEGORIA_FATURAOVER_MOSTRATODOS)
    If lErro <> AD_SQL_SUCESSO Then gError 202159

    lErro = Comando_BuscarPrimeiro(alComando(1))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202160
    
    If lErro = AD_SQL_SUCESSO Then bTodosOVER = True

    'Lê as Notas Fiscais com Fornecedor e Filial passados e desvinculadas
    lErro = Comando_Executar(alComando(2), "SELECT FilialEmpresa, Fornecedor, Filial, NumIntDoc, NumNotaFiscal, DataEmissao, Status, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NFsPag WHERE Fornecedor = ? AND NumIntTitPag = 0 AND Status <> ? " & sFiltro & " ORDER BY NumNotaFiscal", tNFPag.iFilialEmpresa, tNFPag.lFornecedor, tNFPag.iFilial, tNFPag.lNumIntDoc, tNFPag.lNumNotaFiscal, tNFPag.dtDataEmissao, tNFPag.iStatus, tNFPag.dtDataVencimento, tNFPag.dValorTotal, tNFPag.dValorSeguro, tNFPag.dValorFrete, tNFPag.dOutrasDespesas, tNFPag.dValorProdutos, tNFPag.dValorIRRF, tNFPag.dValorICMS, tNFPag.dValorICMSSubst, tNFPag.iCreditoICMS, tNFPag.dValorIPI, tNFPag.iCreditoIPI, tNFPag.dPISRetido, tNFPag.dCOFINSRetido, tNFPag.dCSLLRetido, tNFPag.dTaxaMoeda, objFilialFornecedor.lCodFornecedor, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then gError 202161

    lErro = Comando_BuscarPrimeiro(alComando(2))
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202162
    
    'Loop de NFs
    Do While lErro = AD_SQL_SUCESSO

        'Armazena em colNFPagAberta as NF's desvinculadas encontradas
        With tNFPag
            colNFPagAberta.Add .lNumIntDoc, .iFilialEmpresa, .lFornecedor, .iFilial, .lNumNotaFiscal, .dtDataEmissao, .iStatus, .lNumIntTitPag, .dtDataVencimento, .dValorTotal, .dValorSeguro, .dValorFrete, .dOutrasDespesas, .dValorProdutos, .dValorIRRF, .dValorICMS, .dValorICMSSubst, .iCreditoICMS, .dValorIPI, .iCreditoIPI, .dPISRetido, .dCOFINSRetido, .dCSLLRetido
        End With

        'Tenta ler a próxima NF
        lErro = Comando_BuscarProximo(alComando(2))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202163

    Loop

    If bTodosOVER Then
    
        'Lê as Notas Fiscais de OVER e desvinculadas
        lErro = Comando_Executar(alComando(3), "SELECT N.FilialEmpresa, N.Fornecedor, N.Filial, N.NumIntDoc, N.NumNotaFiscal, N.DataEmissao, N.Status, N.DataVencimento, N.ValorTotal, N.ValorSeguro, N.ValorFrete, N.OutrasDespesas, N.ValorProdutos, N.ValorIRRF, N.ValorICMS, N.ValorICMSSubst, N.CreditoICMS, N.ValorIPI, N.CreditoIPI, N.PISRetido, N.COFINSRetido, N.CSLLRetido, N.TaxaMoeda FROM NFsPag AS N, TRVTitulos AS T " & _
                                                "WHERE N.NumIntDoc = T.NumIntDocOrigem AND T.TipoDocOrigem = 5 AND T.TipoDoc = ? AND NumIntTitPag = 0 AND Status <> ? " & sFiltro & " ORDER BY NumNotaFiscal", tNFPag.iFilialEmpresa, tNFPag.lFornecedor, tNFPag.iFilial, tNFPag.lNumIntDoc, tNFPag.lNumNotaFiscal, tNFPag.dtDataEmissao, tNFPag.iStatus, tNFPag.dtDataVencimento, tNFPag.dValorTotal, tNFPag.dValorSeguro, tNFPag.dValorFrete, tNFPag.dOutrasDespesas, tNFPag.dValorProdutos, tNFPag.dValorIRRF, tNFPag.dValorICMS, tNFPag.dValorICMSSubst, tNFPag.iCreditoICMS, tNFPag.dValorIPI, tNFPag.iCreditoIPI, tNFPag.dPISRetido, tNFPag.dCOFINSRetido, tNFPag.dCSLLRetido, tNFPag.dTaxaMoeda, TRV_TIPODOC_OVER_TEXTO, STATUS_EXCLUIDO)
        If lErro <> AD_SQL_SUCESSO Then gError 202164
    
        lErro = Comando_BuscarPrimeiro(alComando(3))
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202165
        
        'Loop de NFs
        Do While lErro = AD_SQL_SUCESSO
    
            'Armazena em colNFPagAberta as NF's desvinculadas encontradas
            With tNFPag
                colNFPagAberta.Add .lNumIntDoc, .iFilialEmpresa, .lFornecedor, .iFilial, .lNumNotaFiscal, .dtDataEmissao, .iStatus, .lNumIntTitPag, .dtDataVencimento, .dValorTotal, .dValorSeguro, .dValorFrete, .dOutrasDespesas, .dValorProdutos, .dValorIRRF, .dValorICMS, .dValorICMSSubst, .iCreditoICMS, .dValorIPI, .iCreditoIPI, .dPISRetido, .dCOFINSRetido, .dCSLLRetido
            End With
    
            'Tenta ler a próxima NF
            lErro = Comando_BuscarProximo(alComando(3))
            If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 202166
    
        Loop
        
    End If

    'Fecha comando
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    NfsPag_Le_FilialFornecedor_Desvinculadas = SUCESSO

    Exit Function

Erro_NfsPag_Le_FilialFornecedor_Desvinculadas:

    NfsPag_Le_FilialFornecedor_Desvinculadas = gErr

    Select Case gErr

        Case 202158
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 202159, 202160
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIALFORNECEDORCATEGORIAS", gErr)

        Case 202161 To 202166
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202167)

    End Select

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next
    
    Exit Function

End Function

Function NfsPag_Le_FaturaPagar(ByVal objTituloPagar As ClassTituloPagar, ByVal colNFPagVinculada As ColNFsPag) As Long
'Lê as Notas Fiscais vinculadas ao Título passado por parâmetro

Dim lErro As Long
Dim lComando As Long
Dim tNFPag As typeNFsPag

On Error GoTo Erro_NfsPag_Le_FaturaPagar

    'Abre comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 18614

    'Seleciona as NF's vinculadas ao Título passado
    lErro = Comando_Executar(lComando, "SELECT FilialEmpresa, Fornecedor, Filial, NumIntDoc, NumNotaFiscal, DataEmissao, Status, DataVencimento, ValorTotal, ValorSeguro, ValorFrete, OutrasDespesas, ValorProdutos, ValorIRRF, ValorICMS, ValorICMSSubst, CreditoICMS, ValorIPI, CreditoIPI, PISRetido, COFINSRetido, CSLLRetido, TaxaMoeda FROM NFsPag WHERE NumIntTitPag = ? AND Status <> ? ORDER BY DataEmissao", tNFPag.iFilialEmpresa, tNFPag.lFornecedor, tNFPag.iFilial, tNFPag.lNumIntDoc, tNFPag.lNumNotaFiscal, tNFPag.dtDataEmissao, tNFPag.iStatus, tNFPag.dtDataVencimento, tNFPag.dValorTotal, tNFPag.dValorSeguro, tNFPag.dValorFrete, tNFPag.dOutrasDespesas, tNFPag.dValorProdutos, tNFPag.dValorIRRF, tNFPag.dValorICMS, tNFPag.dValorICMSSubst, tNFPag.iCreditoICMS, tNFPag.dValorIPI, tNFPag.iCreditoIPI, tNFPag.dPISRetido, tNFPag.dCOFINSRetido, tNFPag.dCSLLRetido, tNFPag.dTaxaMoeda, objTituloPagar.lNumIntDoc, STATUS_EXCLUIDO)
    If lErro <> AD_SQL_SUCESSO Then Error 18620

    'Tenta ler
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18622
    
    If lErro = AD_SQL_SEM_DADOS Then Error 26020

    'Loop de Notas Fiscais
    Do While lErro = AD_SQL_SUCESSO

        'Armazena NF em colNFPagVinculada
        With tNFPag
            colNFPagVinculada.Add .lNumIntDoc, .iFilialEmpresa, .lFornecedor, .iFilial, .lNumNotaFiscal, .dtDataEmissao, .iStatus, objTituloPagar.lNumIntDoc, .dtDataVencimento, .dValorTotal, .dValorSeguro, .dValorFrete, .dOutrasDespesas, .dValorProdutos, .dValorIRRF, .dValorICMS, .dValorICMSSubst, .iCreditoICMS, .dValorIPI, .iCreditoIPI, .dPISRetido, .dCOFINSRetido, .dCSLLRetido
        End With

        'Busca a próxima NF vinculada
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 18621

    Loop

    'Fecha comando
    Call Comando_Fechar(lComando)

    NfsPag_Le_FaturaPagar = SUCESSO

    Exit Function

Erro_NfsPag_Le_FaturaPagar:

    Select Case Err

        Case 18614
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 18620, 18621, 18622
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFSPAG", Err)

        Case 26020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148598)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Sub ParcelasRec_Le_BaixaManualSel1(sSelect As String, ByVal dtEmissaoInic As Date, ByVal dtEmissaoFim As Date, ByVal dtVencInic As Date, ByVal dtVencFim As Date, ByVal lTituloInic As Long, ByVal lTituloFim As Long, ByVal lCodCli As Long, ByVal iFilialCli As Integer, Optional ByVal sTipoDoc As String = "")
'auxiliar a ParcelasRec_Le_BorderoCobranca
'monta o SELECT dinamicamente
Dim sFrom As String, sWhere As String, sFields As String, sOrderBy As String

    sFields = "TitulosRec.DataEmissao, TitulosRec.NumTitulo, TitulosRec.FilialEmpresa, TitulosRec.SiglaDocumento, TitulosRec.Cliente, TitulosRec.Filial"
    sFields = sFields & ", ParcelasRec.NumIntDoc, ParcelasRec.NumParcela, ParcelasRec.DataVencimento, ParcelasRec.Saldo, ParcelasRec.Valor, ParcelasRec.Cobrador, ParcelasRec.CarteiraCobranca"

    sFrom = " FROM ParcelasRec, TitulosRec"
'    sWhere = " WHERE ParcelasRec.Status = ? AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND TitulosRec.Cliente = ? AND TitulosRec.Filial = ? AND NOT (ParcelasRec.Cobrador = 1 AND ParcelasRec.CarteiraCobranca = 2)"

    '######################################
    'Inserido por Wagner
'    sWhere = " WHERE ParcelasRec.Status = ? AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc AND NOT (ParcelasRec.Cobrador = 1 AND ParcelasRec.CarteiraCobranca = 2)"
    sWhere = " WHERE ParcelasRec.Status = ? AND ParcelasRec.NumIntTitulo = TitulosRec.NumIntDoc"

    If lCodCli <> 0 Then
        If lCodCli = 109660 Or lCodCli = 1120961 Then ' Amex
            'sWhere = sWhere & " AND (TitulosRec.Cliente = ? OR TitulosRec.Cliente = ?)"
            sWhere = sWhere & " AND (TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ?)"
        ElseIf lCodCli = 109664 Or lCodCli = 1120965 Then 'Visanet 109664 e 1120965
            'sWhere = sWhere & " AND (TitulosRec.Cliente = ? OR TitulosRec.Cliente = ? OR TitulosRec.Cliente = ? OR TitulosRec.Cliente = ?)"
            sWhere = sWhere & " AND (TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ?)"
        ElseIf lCodCli = 109662 Or lCodCli = 1120963 Then 'Redecard 109662 e 1120963
            'sWhere = sWhere & " AND (TitulosRec.Cliente = ? OR TitulosRec.Cliente = ?)"
            sWhere = sWhere & " AND (TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ? OR TitulosRec.Cliente <> ?)"
        Else
            sWhere = sWhere & " AND TitulosRec.Cliente = ? "
        End If
    End If
    
    If iFilialCli <> 0 Then
        sWhere = sWhere & " AND TitulosRec.Filial = ?"
    End If
    '######################################

    'Se entrar como empresa toda pode pegar titulos de todas as filiais, senao pegar titulos somente da filial corrente
    If giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        sWhere = sWhere & " AND TitulosRec.FilialEmpresa = ?"
    End If

    'Se titulo inicial preenchido
    If (lTituloInic <> 0) Then
       sWhere = sWhere & " AND TitulosRec.NumTitulo >= ?"
    End If

    'Se titulo final preenchido
    If (lTituloFim <> 0) Then
       sWhere = sWhere & " AND TitulosRec.NumTitulo <= ?"
    End If

    'Se o limite inicial de data de emissao de titulo estiver preenchido
    If (dtEmissaoInic <> DATA_NULA) Then
       sWhere = sWhere & " AND TitulosRec.DataEmissao >= ?"
    End If

    'Se o limite final de data de emissao de titulo estiver preenchido
    If (dtEmissaoFim <> DATA_NULA) Then
       sWhere = sWhere & " AND TitulosRec.DataEmissao <= ?"
    End If

    'Se o limite inicial de data de vencimento de parcela estiver preenchido
    If (dtVencInic <> DATA_NULA) Then
       sWhere = sWhere & " AND ParcelasRec.DataVencimento >= ?"
    End If

    'Se o limite final de data de vencimento de parcela estiver preenchido
    If (dtVencFim <> DATA_NULA) Then
       sWhere = sWhere & " AND ParcelasRec.DataVencimento <= ?"
    End If
    
    If (sTipoDoc <> "") Then
       sWhere = sWhere & " AND TitulosRec.SiglaDocumento = ?"
    End If

    sOrderBy = " ORDER BY TitulosRec.NumTitulo, ParcelasRec.NumParcela"
    
    sSelect = "SELECT " & sFields & sFrom & sWhere & sOrderBy

End Sub

Public Function AdmExtFin_Obtem_Datas(ByVal lNumIntDocMov As Long, ByVal dtData As Date, dtDataMov As Date, dtDataBaixa As Date, dtDataCTB As Date) As Long
    dtDataMov = dtData
    dtDataBaixa = gdtDataAtual
    dtDataCTB = dtData
    AdmExtFin_Obtem_Datas = SUCESSO
End Function
