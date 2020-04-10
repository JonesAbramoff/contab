VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoDescChq3Ocx 
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   LockControls    =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   3615
   Begin VB.CommandButton BotaoIntGeracao 
      Caption         =   "Interromper Geração"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   232
      TabIndex        =   1
      Top             =   2070
      Width           =   3105
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   412
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2610
      Width           =   2700
      Begin VB.CommandButton BotaoGerar 
         Caption         =   "Gerar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "BorderoDescChq3Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   90
         Picture         =   "BorderoDescChq3Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   885
      End
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   360
      Left            =   217
      TabIndex        =   5
      Top             =   1560
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   495
      Left            =   -20000
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Geração do Bordero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   577
      TabIndex        =   10
      Top             =   135
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Cheques:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   435
      TabIndex        =   9
      Top             =   645
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cheques Processados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1095
      Width           =   1950
   End
   Begin VB.Label ChequesProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2040
      TabIndex        =   7
      Top             =   1065
      Width           =   1350
   End
   Begin VB.Label TotalCheques 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2055
      TabIndex        =   6
      Top             =   615
      Width           =   1350
   End
End
Attribute VB_Name = "BorderoDescChq3Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim giCancelaBatch As Integer
Dim giExecutando As Integer
Dim gobjBorderoDescChq As ClassBorderoDescChq

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloRec As ClassTituloReceber
Private gobjParcelaRec As ClassParcelaReceber
Private gobjBaixaParcRec As ClassBaixaParcRec
Private gobjBaixaReceber As ClassBaixaReceber
Private gsContaCtaCorrente As String 'conta contabil da conta corrente onde foram depositados os cheques
Private gsContaFilDep As String 'conta contabil da filial recebedora
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o deposito

Public Sub Form_Load()

On Error GoTo Erro_Form_Load

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143728)
    
    End Select
    
    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objBorderoDescChq As ClassBorderoDescChq) As Long

On Error GoTo Erro_Trata_Parametros

    If objBorderoDescChq Is Nothing Then gError 109237
    
    'ínicializa as variáveis que controlarão o batch
    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO
    
    'aponta para o objBorderoDescChq recebido por parâmetro
    Set gobjBorderoDescChq = objBorderoDescChq
    
    'delimita a barra de progresso
    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    'preenche a quantidade de cheques recebida
    TotalCheques.Caption = gobjBorderoDescChq.iQuantChequesSel

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 109237
            Call Rotina_Erro(vbOKOnly, "ERRO_OBJBORDERODESCCHQ_NAO_CRIADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143729)
            
    End Select
    
    Exit Function

End Function

Private Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    'habilita o botão de interromper o processamento
    BotaoIntGeracao.Enabled = True
    
    'se não há um pedido de cancelamento de batch
    If giCancelaBatch <> CANCELA_BATCH Then
    
        'define o estado do processo como em andamento
        giExecutando = ESTADO_ANDAMENTO
        
        gobjBorderoDescChq.objTelaAtualizacao = Me
        
        'chama função de atualizar o bordero de desconto de cheques
        lErro = CF("BorderoDescChq_Atualizar", gobjBorderoDescChq)
        
        'terminou a execução
        giExecutando = ESTADO_PARADO
        
        'desabilita o botão de interromper
        BotaoIntGeracao.Enabled = False
        
        'se não houve sucesso na atualização e a mesma foi interrompida-> erro
        If lErro <> SUCESSO And lErro <> 109282 Then gError 109238
        
        'se a atualização foi interrompida-> erro
        If lErro = 109282 Then gError 109239
        
        Call Chama_Tela("BorderoDescChq4", gobjBorderoDescChq)
        
        Unload Me
        
    End If

    Exit Sub
    
Erro_BotaoGerar_Click:
    
    Select Case gErr
    
        Case 109238
        
        Case 109239
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143730)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    
On Error GoTo Erro_BotaoFechar_Click
    
    'se tentar fechar com o o batch em execução, dá erro
    If giExecutando = ESTADO_ANDAMENTO Then gError 109234
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoFechar_Click:
    
    Select Case gErr
    
        Case 109234
            Call Rotina_Erro(vbOKOnly, "ERRO_BATCH_EM_ANDAMENTO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143731)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoIntGeracao_Click()

On Error GoTo Erro_BotaoIntGeracao_Click

    If giExecutando = ESTADO_ANDAMENTO Then
        giCancelaBatch = CANCELA_BATCH
    Else
        gError 109240
    End If
    
    Exit Sub
    
Erro_BotaoIntGeracao_Click:
    
    Select Case gErr
    
        Case 109240
            Call Rotina_Erro(vbOKOnly, "ERRO_BATCH_PARADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143732)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'libera as globais
    Set gobjBorderoDescChq = Nothing

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    'fecha, mas cancela o batch primeiro
    If giExecutando = ESTADO_ANDAMENTO Then
        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If

End Sub

Public Function Mostra_Evolucao(iCancela As Integer, iNumProc As Integer) As Long

Dim lErro As Long
Dim iEventos As Integer
Dim iProcessados As Integer
Dim iTotal As Integer

On Error GoTo Erro_Mostra_Evolucao

    iEventos = DoEvents()
    
    If giCancelaBatch = CANCELA_BATCH Then
        
        iCancela = CANCELA_BATCH
        giExecutando = ESTADO_PARADO
        
    Else
        
        'atualiza dados da tela ( registros atualizados e a barra )
        iProcessados = StrParaInt(ChequesProcessados.Caption)
        iTotal = gobjBorderoDescChq.iQuantChequesSel
        If iTotal = 0 Then iTotal = 1
        TotalCheques.Caption = CStr(iTotal)
        
        iProcessados = iProcessados + iNumProc
        ChequesProcessados.Caption = CStr(iProcessados)
        
        BarraProgresso.Value = CInt((iProcessados / iTotal) * 100)
        
        giExecutando = ESTADO_ANDAMENTO
        
    End If
    
    Mostra_Evolucao = SUCESSO

    Exit Function
    
Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err
    
    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143733)
            
    End Select
        
    giCancelaBatch = CANCELA_BATCH
    
    Exit Function
    
End Function

Private Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoDescChq2", gobjBorderoDescChq)
    
    'Fecha a tela
    Unload Me

End Sub

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcrec e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iConta As Integer, dValorLivroAux As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcRec = vParams(0)
    Set gobjParcelaRec = vParams(1)
    Set gobjTituloRec = vParams(2)
    Set gobjBaixaReceber = vParams(3)

    'se ainda nao obtive a filial empresa onde vai ser feito o deposito
    If giFilialEmpresaConta = 0 Then
    
        iConta = gobjBorderoDescChq.iContaCorrente
    
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
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoDescChequeFilDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
            If lErro <> SUCESSO Then Error 32245
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloRec.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32246
        
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoDescChequeFilNaoDep", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloRec.iFilialEmpresa, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32247
        
        Else
        
            'grava a contabilizacao na filial da cta (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoDescCheque", gobjBaixaParcRec.lNumIntDoc, gobjTituloRec.lCliente, gobjTituloRec.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta, , , dValorLivroAux)
            If lErro <> SUCESSO Then Error 32248
        
        End If
    
    Else
    
        GridParcelas.Tag = gobjBaixaReceber.colBaixaParcRec.Count
    
        'grava a contabilizacao na filial da cta
        lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoDescChequeRes", gobjBaixaReceber.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
        If lErro <> SUCESSO Then Error 32248
    
    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32243 To 32248
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143645)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long
Dim lErro As Long, sContaContabil As String, dValor As Double, iConta As Integer
Dim objFilial As New ClassFilialCliente, sContaTela As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objInfoParcRec As ClassInfoParcRec, objCarteiraCobrador As New ClassCarteiraCobrador
Dim objBaixaParcRec As ClassBaixaParcRec
Dim dTotalRecebido As Double, sContaMascarada As String
Dim sMsgErro As String
Dim iIndice As Integer

On Error GoTo Erro_Calcula_Mnemonico

    If gobjBorderoDescChq.colInfoParcRec.Count <> gobjBaixaReceber.colBaixaParcRec.Count Then
        sMsgErro = "Mais parcelas que baixas. " & CStr(gobjBorderoDescChq.colInfoParcRec.Count) & " contra " & CStr(gobjBaixaReceber.colBaixaParcRec.Count)
    End If

    iIndice = 0
    For Each objInfoParcRec In gobjBorderoDescChq.colInfoParcRec
        iIndice = iIndice + 1
        Set objBaixaParcRec = gobjBaixaReceber.colBaixaParcRec.Item(iIndice)
        If objInfoParcRec.lNumIntParc <> objBaixaParcRec.lNumIntParcela Then
            If sMsgErro <> "" Then sMsgErro = sMsgErro & vbNewLine
            sMsgErro = sMsgErro & "Parcela da linha " & CStr(iIndice) & " não corresponde a baixa. Número interno " & CStr(objInfoParcRec.lNumIntParc) & " contra " & CStr(objBaixaParcRec.lNumIntParcela)
            Exit For
        End If
    Next
    If sMsgErro <> "" Then gError 99999

    Select Case objMnemonicoValor.sMnemonico

        Case "FilialCli_Conta_Ctb"
            For Each objInfoParcRec In gobjBorderoDescChq.colInfoParcRec

                objFilial.lCodCliente = objInfoParcRec.lCliente
                objFilial.iCodFilial = objInfoParcRec.iFilialCliente
                
                lErro = CF("FilialCliente_Le", objFilial)
                If lErro <> SUCESSO Then gError 186138
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 186139
                
                Else
                
                    sContaMascarada = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaMascarada
            Next
        
        
        'Início do trecho adicionado Rafael Menezes em 15/10/2002
        Case "Valor_Recebido_Total"
            objMnemonicoValor.colValor.Add gobjBorderoDescChq.dValorChequesSel
        
        Case "Valor_Recebido_Loja"
            For Each objBaixaParcRec In gobjBaixaReceber.colBaixaParcRec
                dTotalRecebido = dTotalRecebido + objBaixaParcRec.dValorRecebido
            Next
            objMnemonicoValor.colValor.Add (gobjBorderoDescChq.dValorChequesSel - dTotalRecebido)
        
        'Fim do trecho adicionado Rafael Menezes em 15/10/2002
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
            For Each objInfoParcRec In gobjBorderoDescChq.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.lNumTitulo
            Next
        
        Case "Parcela_Det"
            For Each objInfoParcRec In gobjBorderoDescChq.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.iNumParcela
            Next
        
        Case "Cliente_Codigo_Det"
            For Each objInfoParcRec In gobjBorderoDescChq.colInfoParcRec
                objMnemonicoValor.colValor.Add objInfoParcRec.lCliente
            Next
        
        Case "Valor_Pago"
        
            dValor = gobjBaixaParcRec.dValorRecebido
            objMnemonicoValor.colValor.Add dValor
        
        Case "Valor_Baixado"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorBaixado
        
        Case "Valor_Desconto"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorDesconto
        
        Case "Valor_Juros"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorJuros
        
        Case "Valor_Multa"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcRec.dValorMulta
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                iConta = gobjBorderoDescChq.iContaCorrente
                lErro = CF("ContaCorrenteInt_Le", iConta, objContaCorrenteInt)
                If lErro <> SUCESSO Then gError 56546
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 56547
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                
        Case "FilDep_Cta_Transf" 'conta de transferencia da filial do deposito

            If gsContaFilDep = "" Then
            
                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(giFilialEmpresaConta, sContaContabil)
                If lErro <> SUCESSO Then gError 56548
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 56549
                    
                Else
                
                    sContaTela = ""
                    
                End If
            
                gsContaFilDep = sContaTela
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaFilDep
        
        Case "FilNaoDep_Cta_Transf" 'conta de transferencia da filial da parcela

                lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloRec.iFilialEmpresa, sContaContabil)
                If lErro <> SUCESSO Then gError 56550
                
                If sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 56551
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
        
        Case "Num_Titulo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.lNumTitulo
            
        Case "Parcela"
            objMnemonicoValor.colValor.Add gobjParcelaRec.iNumParcela
        
        Case "Cliente_Codigo"
            
            objMnemonicoValor.colValor.Add gobjTituloRec.lCliente
                    
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloRec.iFilial
        
        Case Else
        
            Error 56552

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 56546 To 56551, 32273
        
        Case 56552
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 99999
            Call Rotina_Erro(vbOKOnly, sMsgErro, Err, Error)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143644)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    ' ???Parent.HelpContextID = IDH_BORDERO_DESCCHQ1
    Set Form_Load_Ocx = Me
    Caption = "Borderô de Desconto de Cheques - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoDescChq3"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******
