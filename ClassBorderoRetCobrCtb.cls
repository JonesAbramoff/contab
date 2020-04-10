VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBorderoRetCobrCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'??? pendencia: rever codigos de erros nao utilizados e numeracao de erros em geral

Private mvarobjBorderoRetCobrCart As ClassBorderoRetCobrCart

Private mvarobjFormCTBAux As Object

Private bContabilizandoTarifas As Boolean

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloRec As ClassTituloReceber
Private gobjParcelaRec As ClassParcelaReceber
Private gobjBaixaParcRec As ClassBaixaParcRec
Private gobjBaixaReceber As ClassBaixaReceber
Private gsContaCtaCorrente As String 'conta contabil da conta corrente dos depositos
Private gsContaFilDep As String 'conta contabil da filial recebedora
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o deposito
Private gsBancoNomeRed As String 'nome reduzido do banco do bordero

Private Const DATA_BAIXA As String = "Data_Baixa"
Private Const DATA_CREDITO As String = "Data_Credito"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const NUM_TITULO As String = "Numero_Titulo"
Private Const PARCELA1 As String = "Parcela"
Private Const VALOR_BAIXAR As String = "Valor_Baixar"
Private Const VALOR_DESCONTO As String = "Valor_Desconto"
Private Const VALOR_MULTA As String = "Valor_Multa"
Private Const VALOR_JUROS As String = "Valor_Juros"
Private Const VALOR_RECEBIDO As String = "Valor_Recebido"
Private Const CTACARTEIRACOBRADOR As String = "Cta_CarteiraCobrador"
Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CONTA_DET As String = "FilCli_Conta_Ctb_Det"

Public Property Get Controls() As Object
    
    If mvarobjFormCTBAux Is Nothing Then
    
        Set mvarobjFormCTBAux = New FormAuxCTB
        
    End If
    
    Set Controls = mvarobjFormCTBAux.Controls
    
End Property

Public Property Set objBorderoRetCobrCart(ByVal vData As ClassBorderoRetCobrCart)
    Set mvarobjBorderoRetCobrCart = vData
End Property

Public Property Get objBorderoRetCobrCart() As ClassBorderoRetCobrCart
    Set objBorderoRetCobrCart = mvarobjBorderoRetCobrCart
End Property

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcrec e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    bContabilizandoTarifas = False
    
    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcRec = vParams(0)
    Set gobjParcelaRec = vParams(1)
    Set gobjTituloRec = vParams(2)
    Set gobjBaixaReceber = vParams(3)

    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = objBorderoRetCobrCart.objMovCCI.iCodConta
    
        lErro = CF("ContaCorrenteInt_Le", iConta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32243
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial onde vai ser feito o deposito
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32244
    
    'se contabiliza parcela p/parcela
    If gobjCR.iContabSemDet = 0 Then
    
        dValorLivroAux = Round(gobjBaixaParcRec.dValorRecebido + gobjBaixaParcRec.dValorDesconto - gobjBaixaParcRec.dValorJuros - gobjBaixaParcRec.dValorMulta, 2)
    
        'se a filial onde vai ser feito o deposito é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If giFilialEmpresaConta <> gobjTituloRec.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial onde vai ser feito o deposito
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoRetCobrFilDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
            If lErro <> SUCESSO Then Error 32245
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloRec.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32246
        
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoRetCobrFilNaoDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloRec.iFilialEmpresa, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32247
        
        Else
        
            'grava a contabilizacao na filial da cta (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoRetCobr", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32248
        
        End If
    
    Else
    
        Controls("GridParcelas").Tag = gobjBaixaReceber.colBaixaParcRec.Count
            
        'grava a contabilizacao na filial da cta (a mesma do titulo)
        lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoRetCobrRes", gobjBaixaReceber.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
        If lErro <> SUCESSO Then Error 32248
            
    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32243 To 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144657)
     
    End Select
     
    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objCliente As New ClassCliente, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor
Dim iLinha As Integer, iConta As Integer, sContaContabil As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, sContaTela As String
Dim objCarteiraCobrador As New ClassCarteiraCobrador, dValor As Double, bLer As Boolean
Dim objBaixaParcRec As ClassBaixaParcRec, objInfoParcRec As ClassInfoParcRec
Dim objFilial As New ClassFilialCliente

On Error GoTo Erro_Calcula_Mnemonico

    If bContabilizandoTarifas = False Then
    
        Select Case objMnemonicoValor.sMnemonico
    
            Case FILIAL_CONTA_DET
                For Each objInfoParcRec In objBorderoRetCobrCart.colInfoParcRec
                    
                    objFilial.lCodCliente = objInfoParcRec.lCliente
                    objFilial.iCodFilial = objInfoParcRec.iFilialCliente
                    lErro = CF("FilialCliente_Le", objFilial)
                    If lErro <> SUCESSO Then gError 39586
                    
                    If objFilial.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 56500
                    
                    Else
                    
                        sContaTela = ""
                        
                    End If
                    
                    objMnemonicoValor.colValor.Add sContaTela
                    
                Next
            
            Case "Cliente_Codigo_Det"
                For Each objInfoParcRec In objBorderoRetCobrCart.colInfoParcRec
                    objMnemonicoValor.colValor.Add objInfoParcRec.lCliente
                Next
            
            Case "Cta_CartCobr_Det"
                            
                For Each objInfoParcRec In objBorderoRetCobrCart.colInfoParcRec
            
                    objCarteiraCobrador.iCobrador = objInfoParcRec.iCobrador
                    objCarteiraCobrador.iCodCarteiraCobranca = objInfoParcRec.iCarteiraCobrador
                    
                    lErro = CartCobr_ObtemCtaTela(objCarteiraCobrador, sContaTela)
                    If lErro <> SUCESSO Then Error 32274
                    
                    objMnemonicoValor.colValor.Add sContaTela
                
                Next
            
            Case "Valor_Recebido_Det"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    objMnemonicoValor.colValor.Add objBaixaParcRec.dValorRecebido
                Next
            
            Case "Valor_Baixar_Det"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    objMnemonicoValor.colValor.Add objBaixaParcRec.dValorBaixado
                Next
            
            Case "Valor_Desconto_Det"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    objMnemonicoValor.colValor.Add objBaixaParcRec.dValorDesconto
                Next
            
            Case "Valor_Juros_Det"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    objMnemonicoValor.colValor.Add objBaixaParcRec.dValorJuros
                Next
            
            Case "Valor_Multa_Det"
                For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                    objMnemonicoValor.colValor.Add objBaixaParcRec.dValorMulta
                Next
            
            Case "Numero_Titulo_Det"
                
                For Each objInfoParcRec In objBorderoRetCobrCart.colInfoParcRec
                    
                    objMnemonicoValor.colValor.Add objInfoParcRec.lNumTitulo
                    
                Next
            
            Case "Parcela_Det"
            
                For Each objInfoParcRec In objBorderoRetCobrCart.colInfoParcRec
                    
                    objMnemonicoValor.colValor.Add objInfoParcRec.iNumParcela
                    
                Next
                        
            Case "Banco_NomeRed", "Conta_Contabil_Conta", "FilDep_Cta_Transf", "FilNaoDep_Cta_Transf", DATA_BAIXA, DATA_CREDITO, CTACARTEIRACOBRADOR
                lErro = Calcula_Mnemonico_Comum(objMnemonicoValor)
                If lErro <> SUCESSO Then Error 32262
                
            Case "Valor_Tarifas"
                objMnemonicoValor.colValor.Add 0
            
            Case "Valor_Custas"
                objMnemonicoValor.colValor.Add 0
            
            Case "Valor_Recebido"
            
                dValor = Round(gobjBaixaParcRec.dValorBaixado - gobjBaixaParcRec.dValorDesconto + gobjBaixaParcRec.dValorJuros + gobjBaixaParcRec.dValorMulta, 2)
                objMnemonicoValor.colValor.Add dValor
            
            Case "Valor_Baixar"
            
                objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorBaixado
            
            Case "Valor_Desconto"
            
                objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorDesconto
            
            Case "Valor_Juros"
            
                objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorJuros
            
            Case "Valor_Multa"
            
                objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorMulta
            
            Case "Numero_Titulo"
            
                objMnemonicoValor.colValor.Add gobjTituloRec.lNumTitulo
                
            Case "Parcela"
                objMnemonicoValor.colValor.Add gobjParcelaRec.iNumParcela
            
            Case "Cliente_Codigo"
                
                objMnemonicoValor.colValor.Add gobjTituloRec.lCliente
            
            Case "Cliente_Nome"
            
                objCliente.lCodigo = gobjTituloRec.lCliente
                
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO Then Error 32253
                
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
            
            Case "FilialCli_Codigo"
            
                objMnemonicoValor.colValor.Add gobjTituloRec.iFilial
            
            Case FILIAL_CONTA
                
                objFilial.lCodCliente = gobjTituloRec.lCliente
                objFilial.iCodFilial = gobjTituloRec.iFilial
                lErro = CF("FilialCliente_Le", objFilial)
                If lErro <> SUCESSO Then gError 39586
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 56500
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Case Else
                Error 39695
    
        End Select

    Else
    
        Select Case objMnemonicoValor.sMnemonico
    
            Case "Banco_NomeRed", "Conta_Contabil_Conta", "FilDep_Cta_Transf", "FilNaoDep_Cta_Transf", DATA_BAIXA, DATA_CREDITO, CTACARTEIRACOBRADOR
                lErro = Calcula_Mnemonico_Comum(objMnemonicoValor)
                If lErro <> SUCESSO Then Error 32263
                
            Case "Valor_Tarifas"
                objMnemonicoValor.colValor.Add objBorderoRetCobrCart.dValorTarifas
            
            Case "Valor_Custas"
                objMnemonicoValor.colValor.Add objBorderoRetCobrCart.dValorCustas
            
            Case "Valor_Recebido", "Valor_Baixar", "Valor_Desconto", "Valor_Juros", "Valor_Multa", "Parcela", "Cliente_Codigo", "FilialCli_Codigo"
            
                objMnemonicoValor.colValor.Add 0
                    
            Case "Numero_Titulo", "Cliente_Nome"
            
                objMnemonicoValor.colValor.Add ""
                
            Case Else
                Error 39695
    
        End Select
        
    End If
    
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 32253, 56527, 56528, 56529, 56530, 56803, 32262, 32263, 32274, 39586, 56500
        
        Case 56798, 56802
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56804
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144658)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao_Outras(objContabAutomatica As ClassContabAutomatica) As Long
'para gerar a contabilizacao referente à custas e tarifas

Dim lErro As Long, lDoc As Long, iConta As Integer
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao_Outras

    bContabilizandoTarifas = True
    
    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = objBorderoRetCobrCart.objMovCCI.iCodConta
    
        lErro = CF("ContaCorrenteInt_Le", iConta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32243
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial onde vai ser feito o deposito
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32244
    
    'grava a contabilizacao na filial da cta
    lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoRetCobrTarifa", objBorderoRetCobrCart.objMovCCI.lNumMovto, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32248
    
    GeraContabilizacao_Outras = SUCESSO
    
    Exit Function
    
Erro_GeraContabilizacao_Outras:

    GeraContabilizacao_Outras = Err
    
    Select Case Err

        Case 32243, 32244, 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144659)

    End Select

    Exit Function

End Function

Function Calcula_Mnemonico_Comum(objMnemonicoValor As ClassMnemonicoValor) As Long
'codigo em comum do Calcula_Mnemonico entre o processamento de baixas de titulos e o processamento de tarifas e custas

Dim lErro As Long, objCliente As New ClassCliente, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor
Dim iLinha As Integer, iConta As Integer, sContaContabil As String, sBancoNomeRed As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, sContaTela As String
Dim objCarteiraCobrador As New ClassCarteiraCobrador, dValor As Double, objBanco As New ClassBanco

On Error GoTo Erro_Calcula_Mnemonico_Comum

    Select Case objMnemonicoValor.sMnemonico

        Case "Banco_NomeRed"
            If gsBancoNomeRed = "" Then
            
                iConta = objBorderoRetCobrCart.objMovCCI.iCodConta
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 56546
                
                If objContaCorrenteInt.iCodBanco <> 0 Then
                
                    objBanco.iCodBanco = objContaCorrenteInt.iCodBanco
                    lErro = CF("Banco_Le", objBanco)
                    If lErro <> SUCESSO And lErro <> 16091 Then Error 32280
                    If lErro = SUCESSO Then sBancoNomeRed = objBanco.sNomeReduzido
                
                Else
                
                    sBancoNomeRed = ""
                    
                End If
                
                gsBancoNomeRed = sBancoNomeRed
                
            End If

            objMnemonicoValor.colValor.Add gsBancoNomeRed
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                iConta = objBorderoRetCobrCart.objMovCCI.iCodConta
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 56546
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56547
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                
        Case "FilDep_Cta_Transf" 'conta de transferencia da filial do deposito

            If gsContaFilDep = "" Then
            
                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(giFilialEmpresaConta, sContaContabil)
                If lErro <> SUCESSO Then Error 56548
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56549
                    
                Else
                
                    sContaTela = ""
                    
                End If
            
                gsContaFilDep = sContaTela
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaFilDep
        
        Case "FilNaoDep_Cta_Transf" 'conta de transferencia da filial da parcela

                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloRec.iFilialEmpresa, sContaContabil)
                If lErro <> SUCESSO Then Error 56550
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56551
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
        
        Case DATA_BAIXA
            objMnemonicoValor.colValor.Add objBorderoRetCobrCart.objMovCCI.dtDataBaixa

        Case DATA_CREDITO
            objMnemonicoValor.colValor.Add objBorderoRetCobrCart.dtDataCredito
            
        Case CTACARTEIRACOBRADOR
        
            objCarteiraCobrador.iCobrador = objBorderoRetCobrCart.iCobrador
            objCarteiraCobrador.iCodCarteiraCobranca = objBorderoRetCobrCart.iCarteiraCobr
            
            If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then
            
                Select Case objCarteiraCobrador.iCodCarteiraCobranca
                
                    Case CARTEIRA_CARTEIRA
                        sCampoGlobal = "CtaReceberCarteira"
                    
                    Case CARTEIRA_CHEQUEPRE
                        sCampoGlobal = "CtaRecChequePre"
                        
                    Case CARTEIRA_JURIDICO
                        sCampoGlobal = "CtaJuridico"
                    
                    Case Else
                        Error 56802
                        
                End Select
                
                objMnemonico.sMnemonico = sCampoGlobal
                lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
                If lErro <> SUCESSO And lErro <> 39690 Then Error 56803
                If lErro <> SUCESSO Then Error 56804
                
                sContaTela = objMnemonico.sValor
                
            Else
            
                lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
                If lErro <> SUCESSO And lErro <> 23551 Then Error 56528
                If lErro <> SUCESSO Then Error 56798
                
                If objCarteiraCobrador.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56529
                
                Else
                
                    sContaTela = ""
                    
                End If
        
            End If
            
            objMnemonicoValor.colValor.Add sContaTela
        
        Case Else
            Error 39695

    End Select
    
    Calcula_Mnemonico_Comum = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico_Comum:

    Calcula_Mnemonico_Comum = Err

    Select Case Err

        Case 32280, 32253, 56527, 56528, 56529, 56530, 56803
        
        Case 56798, 56802
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56804
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico_Comum = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144660)

    End Select

    Exit Function

End Function

'Private Function Carrega_TitParcRec(lNumIntParcela As Long) As Long
''Preenche gobjTituloRec e gobjParcelaRec correspondentes a lNumIntParcela (numintdoc da tabela de parcelasrec/parcelasrecbaixadas)
'
'Dim lErro As Long
'
'On Error GoTo Erro_Carrega_TitParcRec
'
'    gobjParcelaRec.lNumIntDoc = lNumIntParcela
'
'    lErro = CF("ParcelaReceber_Baixada_Le",gobjParcelaRec)
'    If lErro <> SUCESSO And lErro <> 58559 Then gError
'    If lErro <> SUCESSO Then
'
'        lErro = CF("ParcelaReceber_Le",gobjParcelaRec)
'        If lErro <> SUCESSO And lErro <> 19147 Then gError
'        If lErro <> SUCESSO Then gError
'
'        lErro = CF("TituloReceber_Le",gobjTituloRec)
'        If lErro <> SUCESSO And lErro <> 26059 Then gError
'        If lErro <> SUCESSO Then gError
'
'    Else
'
'        lErro = CF("TituloReceberBaixado_Le",gobjTituloRec)
'        If lErro <> SUCESSO And lErro <> 56570 Then gError
'        If lErro <> SUCESSO Then gError
'
'    End If
'
'    Carrega_TitParcRec = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_TitParcRec:
'
'    Carrega_TitParcRec = gErr
'
'    Select Case gErr
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144661)
'
'    End Select
'
'    Exit Function
'
'End Function

'??? já existe em ctbaixareccancelar.cls
Private Function CartCobr_ObtemCtaTela(objCarteiraCobrador As ClassCarteiraCobrador, sContaTela As String) As Long

Dim lErro As Long, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_CartCobr_ObtemCtaTela

    If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then

        Select Case objCarteiraCobrador.iCodCarteiraCobranca

            Case CARTEIRA_CARTEIRA
                sCampoGlobal = "CtaReceberCarteira"

            Case CARTEIRA_CHEQUEPRE
                sCampoGlobal = "CtaRecChequePre"

            Case CARTEIRA_JURIDICO
                sCampoGlobal = "CtaJuridico"

            Case Else
                Error 56799

        End Select

        objMnemonico.sMnemonico = sCampoGlobal
        lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
        If lErro <> SUCESSO And lErro <> 39690 Then Error 56800
        If lErro <> SUCESSO Then Error 56801

        sContaTela = objMnemonico.sValor

    Else

        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 49726
        If lErro <> SUCESSO Then Error 56797

        If objCarteiraCobrador.sContaContabil <> "" Then

            lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
            If lErro <> SUCESSO Then Error 56526

        End If

    End If

    CartCobr_ObtemCtaTela = SUCESSO
     
    Exit Function
    
Erro_CartCobr_ObtemCtaTela:

    CartCobr_ObtemCtaTela = gErr
     
    Select Case gErr
          
        Case 49726, 56526, 56800
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144662)
     
    End Select
     
    Exit Function

End Function



Private Sub Class_Terminate()
    Set mvarobjFormCTBAux = Nothing
End Sub
