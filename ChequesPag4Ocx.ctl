VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ChequesPag4Ocx 
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   ScaleHeight     =   3975
   ScaleWidth      =   3900
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   615
      Left            =   -20000
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   540
      ScaleHeight     =   495
      ScaleWidth      =   2760
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3270
      Width           =   2820
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   75
         Picture         =   "ChequesPag4Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   885
      End
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   1050
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2250
         Picture         =   "ChequesPag4Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoIntAtualiza 
      Caption         =   "Interromper Atualização"
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
      Left            =   398
      TabIndex        =   0
      Top             =   2730
      Width           =   3105
   End
   Begin MSComctlLib.ProgressBar BarraProgresso 
      Height          =   345
      Left            =   210
      TabIndex        =   4
      Top             =   2160
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label TotalTitulos 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1335
      Width           =   1365
   End
   Begin VB.Label TitulosProcessados 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1695
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "Atualização de Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   443
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parcelas Processadas:"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1695
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Parcelas:"
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
      Left            =   615
      TabIndex        =   10
      Top             =   1350
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Left            =   615
      TabIndex        =   11
      Top             =   585
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   12
      Top             =   975
      Width           =   1950
   End
   Begin VB.Label ChequesProcessados 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label TotalCheques 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   570
      Width           =   1365
   End
End
Attribute VB_Name = "ChequesPag4Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giExecutando As Integer
Dim giCancelaBatch As Integer
Dim gobjChequesPag As ClassChequesPag

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloPagar As New ClassTituloPagar
Private gobjParcelaPagar As New ClassParcelaPagar
Private gobjBaixaParcPagar As New ClassBaixaParcPagar
Private gobjBaixaPagar As New ClassBaixaPagar
Private gsContaCtaCorrente As String 'conta contabil da conta corrente
Private gsContaFilPag As String 'conta contabil da filial pagadora
Private gsContaFornecedores As String

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    BotaoAtualizar.Enabled = False
    BotaoVoltar.Enabled = False
    BotaoIntAtualiza.Enabled = True
    
    If giCancelaBatch <> CANCELA_BATCH Then
    
        giExecutando = ESTADO_ANDAMENTO
        gobjChequesPag.objTelaAtualizacao = Me
        lErro = CF("ChequesPag_AtualizarBD", gobjChequesPag)
        giExecutando = ESTADO_PARADO
        
        BotaoIntAtualiza.Enabled = False
        
        If lErro <> SUCESSO And lErro <> 30626 And lErro <> 26429 Then Error 41425
            
        If lErro = 30626 Or lErro = 26429 Then Error 41426 'interrompeu
                
    End If
    
    Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case Err
        
        Case 41426
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
            Unload Me
            
        Case 41425
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144571)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    If giExecutando = ESTADO_ANDAMENTO Then Error 41427
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoFechar_Click:
        
    Select Case Err
    
        Case 41427
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_IMPOSSIVEL_CONTINUAR")
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144572)
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoIntAtualiza_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    End If
    
    'Fecha a tela
    Unload Me

End Sub

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long

Dim lErro As Long, objInfoParcPag As ClassInfoParcPag
Dim iCount As Integer

On Error GoTo Erro_Trata_Parametros

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO
    
    Set gobjChequesPag = objChequesPag
    
    Set gobjChequesPag.objEvolucao = Me
    
    'Passa para a tela os dados dos Títulos selecionados
    
    TotalCheques.Caption = CStr(gobjChequesPag.ColInfoChequePag.Count)
    
    iCount = 0
    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
    
        If objInfoParcPag.iSeqCheque <> 0 Then iCount = iCount + 1
    
    Next
    
    TotalTitulos.Caption = CStr(iCount)
    
    TitulosProcessados.Caption = "0"
    ChequesProcessados.Caption = "0"
    
    BarraProgresso.Min = 0
    BarraProgresso.Max = 100
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144573)
            
    End Select

    giCancelaBatch = CANCELA_BATCH
        
    Exit Function

End Function

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
        
        iProcessados = CInt(TitulosProcessados.Caption)
        iTotal = CInt(TotalTitulos.Caption)
        
        iProcessados = iProcessados + iNumProc
        TitulosProcessados.Caption = CStr(iProcessados)
        
        BarraProgresso.Value = CInt((iProcessados / iTotal) * 100)
        
        giExecutando = ESTADO_ANDAMENTO
        
    End If
    
    Mostra_Evolucao = SUCESSO

    Exit Function
    
Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144574)
            
    End Select
        
    giCancelaBatch = CANCELA_BATCH
    
    Exit Function
    
End Function

Public Function Mostra_Evolucao1(iCancela As Integer, iNumProc As Integer) As Long

Dim lErro As Long
Dim iEventos As Integer
Dim iProcessados As Integer

On Error GoTo Erro_Mostra_Evolucao1

    iEventos = DoEvents()
    
    If giCancelaBatch = CANCELA_BATCH Then
        
        iCancela = CANCELA_BATCH
        giExecutando = ESTADO_PARADO
        
    Else
        'atualiza dados da tela ( registros atualizados e a barra )
        
        iProcessados = CInt(ChequesProcessados.Caption)
        iProcessados = iProcessados + iNumProc
        ChequesProcessados.Caption = CStr(iProcessados)
        
        giExecutando = ESTADO_ANDAMENTO
        
    End If
    
    Mostra_Evolucao1 = SUCESSO

    Exit Function
    
Erro_Mostra_Evolucao1:

    Mostra_Evolucao1 = Err
    
    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144575)
            
    End Select
        
    giCancelaBatch = CANCELA_BATCH
    
    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If giExecutando = ESTADO_ANDAMENTO Then
        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjContabAutomatica = Nothing
    Set gobjTituloPagar = Nothing
    Set gobjParcelaPagar = Nothing
    Set gobjBaixaParcPagar = Nothing
    Set gobjBaixaPagar = Nothing
    gsContaCtaCorrente = ""
    gsContaFilPag = ""
    
    Set gobjChequesPag = Nothing
    
End Sub

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaContabil As String, dValor As Double, iIndice As Integer
Dim objFilial As New ClassFilialFornecedor, sContaTela As String, iMotivoDiferenca As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, bAchou As Boolean
Dim objBaixaParcPagarDet As ClassBaixaParcPagar
Dim objInfoParcPag As ClassInfoParcPag, objFornecedor As New ClassFornecedor
Dim objInfoChequePag As ClassInfoChequePag
Dim sContaFormatada As String, iContaPreenchida As Integer, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico

    bAchou = True
    
    'tratar mnemonicos comuns a contab parcela a parcela e contab p/cheque c/um todo
    Select Case objMnemonicoValor.sMnemonico

        Case "Numero_Cheque"
            objMnemonicoValor.colValor.Add gobjChequesPag.ColInfoChequePag.Item(gobjChequesPag.iIndiceChequeProc).lNumRealCheque
        
        Case "Valor_Pago"
        
            dValor = gobjBaixaParcPagar.dValorBaixado - gobjBaixaParcPagar.dValorDesconto + gobjBaixaParcPagar.dValorJuros + gobjBaixaParcPagar.dValorMulta + gobjBaixaParcPagar.dValorDiferenca
            objMnemonicoValor.colValor.Add dValor
        
        Case "Valor_Baixado"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcPagar.dValorBaixado
        
        Case "Valor_Desconto"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcPagar.dValorDesconto
        
        Case "Valor_Juros"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcPagar.dValorJuros
        
        Case "Valor_Multa"
        
            objMnemonicoValor.colValor.Add gobjBaixaParcPagar.dValorMulta
        
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
                
            If gsContaCtaCorrente = "" Then
                
                lErro = CF("ContaCorrenteInt_Le", gobjChequesPag.iCta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 41905
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41906
                                        
                Else
                
                    sContaTela = ""
                    
                End If
                
                gsContaCtaCorrente = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add gsContaCtaCorrente
                        
        Case Else
            bAchou = False
                    
    End Select
    
    If bAchou = False Then
    
        'se contabiliza o cheque como um todo
        If gobjCP.iContabSemDet = 1 Then
        
            Select Case objMnemonicoValor.sMnemonico
        
                Case "Valor_Forn_Aux"
                    
                    Set objInfoChequePag = gobjChequesPag.ColInfoChequePag(gobjChequesPag.iIndiceChequeProc)
                    
                    iIndice = 1
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
                
                        'pular parcelas de outros cheques
                        If objInfoParcPag.iSeqCheque = objInfoChequePag.iSeqCheque Then

                            If gsContaFornecedores = "" Then
                            
                                objMnemonico.sMnemonico = "CtaFornecedores"
                                lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
                                If lErro <> SUCESSO And lErro <> 39690 Then Error 56808
                                If lErro <> SUCESSO Then Error 56809
                                
                                lErro = CF("Conta_Formata", objMnemonico.sValor, sContaFormatada, iContaPreenchida)
                                If lErro <> SUCESSO Then Error 56808
                                
                                If iContaPreenchida = CONTA_PREENCHIDA Then gsContaFornecedores = sContaFormatada
        
                            End If
                            
                            If objInfoParcPag.sContaFilForn = "" Or objInfoParcPag.sContaFilForn = gsContaFornecedores Then
                            
                                Set objBaixaParcPagarDet = gobjChequesPag.objColBaixaParcPagar(iIndice)
                                dValor = Round(dValor + objBaixaParcPagarDet.dValorBaixado - objBaixaParcPagarDet.dValorDiferenca, 2)
                                
                            End If
                            
                            iIndice = iIndice + 1
        
                        End If
        
                    Next
                    
                    objMnemonicoValor.colValor.Add dValor
                
                Case "Valor_Pago_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        With objBaixaParcPagarDet
                            dValor = Round(.dValorBaixado - .dValorDesconto + .dValorJuros + .dValorMulta + .dValorDiferenca, 2)
                        End With
                        objMnemonicoValor.colValor.Add dValor
                    Next
                        
                Case "Valor_Baixado_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorBaixado
                    Next
                
                Case "Valor_Desconto_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorDesconto
                    Next
                
                Case "Valor_Juros_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorJuros
                    Next
                
                Case "Valor_Multa_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorMulta
                    Next
                
                Case "Valor_Diferenca_Det"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    For Each objBaixaParcPagarDet In gobjChequesPag.objColBaixaParcPagar
                        If iMotivoDiferenca = 0 Or iMotivoDiferenca = objBaixaParcPagarDet.iMotivoDiferenca Then
                            dValor = objBaixaParcPagarDet.dValorDiferenca
                        Else
                            dValor = 0
                        End If
                        objMnemonicoValor.colValor.Add dValor
                    Next
                
                Case "FilialForn_Conta_Det" 'conta contabil da filial do fornecedor da parcela
                    
                    Set objInfoChequePag = gobjChequesPag.ColInfoChequePag(gobjChequesPag.iIndiceChequeProc)
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
                
                        'pular parcelas de outros cheques
                        If objInfoParcPag.iSeqCheque = objInfoChequePag.iSeqCheque Then

                            If objInfoParcPag.sContaFilForn <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objInfoParcPag.sContaFilForn, sContaTela)
                                If lErro <> SUCESSO Then Error 41904
                                
                                objMnemonicoValor.colValor.Add sContaTela
                            Else
                                objMnemonicoValor.colValor.Add ""
                            End If
        
                        End If
        
                    Next
                
                Case "Num_Titulo_Det"
                
                    Set objInfoChequePag = gobjChequesPag.ColInfoChequePag(gobjChequesPag.iIndiceChequeProc)
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
                
                        'pular parcelas de outros cheques
                        If objInfoParcPag.iSeqCheque = objInfoChequePag.iSeqCheque Then

                            objMnemonicoValor.colValor.Add objInfoParcPag.lNumTitulo
                    
                        End If
        
                    Next
                    
                Case "Fornec_Codigo_Det"
                    
                    Set objInfoChequePag = gobjChequesPag.ColInfoChequePag(gobjChequesPag.iIndiceChequeProc)
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
                
                        'pular parcelas de outros cheques
                        If objInfoParcPag.iSeqCheque = objInfoChequePag.iSeqCheque Then

                            objMnemonicoValor.colValor.Add objInfoParcPag.lFornecedor
                    
                        End If
        
                    Next
                                    
'?????
                Case "Valor_Diferenca"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    If iMotivoDiferenca = 0 Then
                        dValor = gobjBaixaParcPagar.dValorDiferenca
                    Else
                        If iMotivoDiferenca <= gobjChequesPag.adValorDiferenca_UBound Then dValor = gobjChequesPag.adValorDiferenca(iMotivoDiferenca)
                    End If
                    objMnemonicoValor.colValor.Add dValor
                
                Case Else
                
                    Error 41911
        
            End Select
        
        Else 'se contabiliza parcela a parcela
        
            Select Case objMnemonicoValor.sMnemonico
        
                Case "Valor_Diferenca"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    If iMotivoDiferenca = gobjBaixaParcPagar.iMotivoDiferenca Then dValor = gobjBaixaParcPagar.dValorDiferenca
                    objMnemonicoValor.colValor.Add dValor
                
                Case "Num_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lNumTitulo
                    
                Case "Fornecedor_Codigo"
                    
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lFornecedor
                            
                Case "Fornecedor_Nome"
                
                    objFornecedor.lCodigo = gobjTituloPagar.lFornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO Then Error 41903
                
                    objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
                
                Case "Fornecedor_NomeRed"
                
                    objFornecedor.lCodigo = gobjTituloPagar.lFornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO Then Error 41903
                
                    objMnemonicoValor.colValor.Add objFornecedor.sNomeReduzido
                
                Case "Data_Emissao_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.dtDataEmissao
                
                Case "FilialForn_Conta" 'conta contabil da filial do fornecedor da parcela
                    
                    objFilial.lCodFornecedor = gobjTituloPagar.lFornecedor
                    objFilial.iCodFilial = gobjTituloPagar.iFilial
                    
                    lErro = CF("FilialFornecedor_Le", objFilial)
                    If lErro <> SUCESSO Then Error 41903
                    
                    If objFilial.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 41904
                        
                        objMnemonicoValor.colValor.Add sContaTela
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
        
                Case "FilPag_Cta_Transf" 'conta de transferencia da filial pagadora
        
                    If gsContaFilPag = "" Then
                    
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjChequesPag.iFilialEmpresaCta, sContaContabil)
                        If lErro <> SUCESSO Then Error 41907
                        
                        If sContaContabil <> "" Then
                        
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then Error 41908
                            
                        Else
                        
                            sContaTela = ""
                            
                        End If
                    
                        gsContaFilPag = sContaTela
                        
                    End If
                    
                    objMnemonicoValor.colValor.Add gsContaFilPag
                
                Case "FilNaoPag_Cta_Transf" 'conta de transferencia da filial da parcela
        
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloPagar.iFilialEmpresa, sContaContabil)
                        If lErro <> SUCESSO Then Error 41909
                        
                        If sContaContabil <> "" Then
                        
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then Error 41910
                        
                        Else
                        
                            sContaTela = ""
                            
                        End If
                        
                        objMnemonicoValor.colValor.Add sContaTela
                
                                
                Case "Tipo_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.sSiglaDocumento
         
                
                Case Else
                
                    Error 41911
        
            End Select

        End If
        
    End If
            
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56808
        
        Case 56809
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 41903 To 41910
        
        Case 41911
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144576)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcpag e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iIndice As Integer, dValorDiferenca As Double
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjBaixaParcPagar = vParams(0)
    Set gobjParcelaPagar = vParams(1)
    Set gobjTituloPagar = vParams(2)
    Set gobjBaixaPagar = vParams(3)
    
    'obtem numero de doc para a filial pagadora
    lErro = objContabAutomatica.Obter_Doc(lDoc, gobjChequesPag.iFilialEmpresaCta)
    If lErro <> SUCESSO Then Error 32184
    
    'se contabiliza parcela p/parcela
    If gobjCP.iContabSemDet = 0 Then
    
        'se a filial pagadora é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If gobjChequesPag.iFilialEmpresaCta <> gobjTituloPagar.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial pagadora
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeAutoFilPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPag.iFilialEmpresaCta)
            If lErro <> SUCESSO Then Error 32185
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloPagar.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32186
         
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeAutoFilNaoPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloPagar.iFilialEmpresa, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32187
        
        Else
        
            'grava a contabilizacao na filial pagadora (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeAuto", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPag.iFilialEmpresaCta, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32188
        
        End If
    
    Else 'se contabiliza o cheque como um todo
    
        GridParcelas.Tag = gobjChequesPag.objColBaixaParcPagar.Count
        
        Set gobjBaixaParcPagar = New ClassBaixaParcPagar
        
        With gobjBaixaParcPagar
            .dValorBaixado = gobjChequesPag.dValorBaixado
            .dValorDesconto = gobjChequesPag.dValorDesconto
            .dValorJuros = gobjChequesPag.dValorJuros
            .dValorMulta = gobjChequesPag.dValorMulta
        End With
        
        For iIndice = gobjChequesPag.adValorDiferenca_LBound To gobjChequesPag.adValorDiferenca_UBound
        
            dValorDiferenca = Round(dValorDiferenca + gobjChequesPag.adValorDiferenca(iIndice), 2)
        
        Next
        
        gobjBaixaParcPagar.dValorDiferenca = dValorDiferenca
        
        'grava a contabilizacao na filial pagadora
        lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeAutoRes", gobjBaixaPagar.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPag.iFilialEmpresaCta)
        If lErro <> SUCESSO Then Error 32188
    
    End If
        
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32183 To 32188
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144577)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144578)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_IMPRESSAO_CHEQUES_P4
    Set Form_Load_Ocx = Me
    Caption = "Impressão de Cheques - Passo 4"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequesPag4"
    
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

Private Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("ChequesPag3", gobjChequesPag)

    'Fecha a tela
    Unload Me
    
End Sub

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



Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub ChequesProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ChequesProcessados, Source, X, Y)
End Sub

Private Sub ChequesProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ChequesProcessados, Button, Shift, X, Y)
End Sub

Private Sub TotalCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalCheques, Source, X, Y)
End Sub

Private Sub TotalCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalCheques, Button, Shift, X, Y)
End Sub

