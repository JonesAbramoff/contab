VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ChequePagAvulso4Ocx 
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ScaleHeight     =   3390
   ScaleWidth      =   3630
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   495
      Left            =   -20000
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
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
      Left            =   255
      TabIndex        =   0
      Top             =   2085
      Width           =   3105
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   413
      ScaleHeight     =   495
      ScaleWidth      =   2745
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2625
      Width           =   2805
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   105
         Picture         =   "ChequePagAvulso4Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2205
         Picture         =   "ChequePagAvulso4Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
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
         Left            =   1072
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1050
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   173
      TabIndex        =   4
      Top             =   1530
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label TotalTitulos 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2070
      TabIndex        =   6
      Top             =   645
      Width           =   1365
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2070
      TabIndex        =   7
      Top             =   1110
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
      Left            =   308
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   90
      TabIndex        =   9
      Top             =   1125
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
      Left            =   465
      TabIndex        =   10
      Top             =   660
      Width           =   1575
   End
End
Attribute VB_Name = "ChequePagAvulso4Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giCancelaBatch As Integer
Dim giExecutando As Integer ' 0: nao está executando, 1: em andamento

Dim gobjChequesPagAvulso As ClassChequesPagAvulso

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloPagar As New ClassTituloPagar
Private gobjParcelaPagar As New ClassParcelaPagar
Private gobjBaixaParcPagar As New ClassBaixaParcPagar
Private gobjBaixaPagar As New ClassBaixaPagar
Private gsContaFornecedores As String

Function Trata_Parametros(Optional objChequesPagAvulso As ClassChequesPagAvulso) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO

    If (objChequesPagAvulso Is Nothing) Then
        Error 7787
    Else
        Set gobjChequesPagAvulso = objChequesPagAvulso
    End If

    Set gobjChequesPagAvulso.objEvolucao = Me

    'Passa para a tela os dados dos Títulos selecionados
    TotalTitulos.Caption = CStr(gobjChequesPagAvulso.iQtdeParcelasSelecionadas)
    TitulosProcessados.Caption = "0"

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 7787
            giCancelaBatch = CANCELA_BATCH

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144492)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        giCancelaBatch = CANCELA_BATCH
        BotaoFechar.Enabled = False
        Exit Sub
    End If

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    BotaoAtualizar.Enabled = False
    BotaoVoltar.Enabled = False
    BotaoIntAtualiza.Enabled = True
    
    If giCancelaBatch <> CANCELA_BATCH Then

        giExecutando = ESTADO_ANDAMENTO
        gobjChequesPagAvulso.objTelaAtualizacao = Me
        lErro = CF("ChequesPagAvulso_AtualizarBD", gobjChequesPagAvulso)
        giExecutando = ESTADO_PARADO

        BotaoIntAtualiza.Enabled = False

        If lErro <> SUCESSO And lErro <> 30627 Then Error 7999

        If lErro = 30627 Then Error 41428 'interrompeu

    End If

    Exit Sub

Erro_BotaoAtualizar_Click:

    Select Case Err

        Case 41428
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
            Unload Me

        Case 7999

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144493)

    End Select

    Exit Sub

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
        iProcessados = CInt(TitulosProcessados.Caption)
        iTotal = CInt(TotalTitulos.Caption)

        iProcessados = iProcessados + iNumProc
        TitulosProcessados.Caption = CStr(iProcessados)

        ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)

        giExecutando = ESTADO_ANDAMENTO

    End If

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144494)

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

    Set gobjChequesPagAvulso = Nothing

End Sub

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaContabil As String, dValor As Double, iIndice As Integer
Dim objFilial As New ClassFilialFornecedor, sContaTela As String, iMotivoDiferenca As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, bAchou As Boolean
Dim objBaixaParcPagarDet As ClassBaixaParcPagar
Dim objInfoParcPag As ClassInfoParcPag, objFornecedor As New ClassFornecedor
Dim sContaFormatada As String, iContaPreenchida As Integer, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico

    bAchou = True
    
    'tratar mnemonicos comuns a contab parcela a parcela e contab p/cheque c/um todo
    Select Case objMnemonicoValor.sMnemonico

        Case "Numero_Cheque"
            objMnemonicoValor.colValor.Add gobjChequesPagAvulso.lNumCheque

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

            lErro = CF("ContaCorrenteInt_Le", gobjChequesPagAvulso.iCta, objContaCorrenteInt)
            If lErro <> SUCESSO Then Error 32191

            If objContaCorrenteInt.sContaContabil <> "" Then

                lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                If lErro <> SUCESSO Then Error 32192

            Else

                sContaTela = ""

            End If

            objMnemonicoValor.colValor.Add sContaTela

        Case "Conta_Cheque_Pre" 'conta contabil associada a conta corrente utilizada p/o cheques pre-datados

            lErro = CF("ContaCorrenteInt_Le", gobjChequesPagAvulso.iCta, objContaCorrenteInt)
            If lErro <> SUCESSO Then Error 32191

            If objContaCorrenteInt.sContaContabilChqPre <> "" Then

                lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabilChqPre, sContaTela)
                If lErro <> SUCESSO Then Error 32192

            Else

                sContaTela = ""

            End If

            objMnemonicoValor.colValor.Add sContaTela

        Case Else
            bAchou = False
                    
    End Select
    
    If bAchou = False Then
    
        'se contabiliza o cheque como um todo
        If gobjCP.iContabSemDet = 1 Then
        
            Select Case objMnemonicoValor.sMnemonico
        
                Case "Valor_Forn_Aux"
                    
                    iIndice = 1
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPagAvulso.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

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
                            
                                Set objBaixaParcPagarDet = gobjChequesPagAvulso.objColBaixaParcPagar(iIndice)
                                dValor = Round(dValor + objBaixaParcPagarDet.dValorBaixado - objBaixaParcPagarDet.dValorDiferenca, 2)
                                
                            End If
                            
                            iIndice = iIndice + 1
        
                        End If
        
                    Next
                    
                    objMnemonicoValor.colValor.Add dValor
                
                Case "Valor_Pago_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        With objBaixaParcPagarDet
                            dValor = Round(.dValorBaixado - .dValorDesconto + .dValorJuros + .dValorMulta + .dValorDiferenca, 2)
                        End With
                        objMnemonicoValor.colValor.Add dValor
                    Next
                        
                Case "Valor_Baixado_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorBaixado
                    Next
                
                Case "Valor_Desconto_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorDesconto
                    Next
                
                Case "Valor_Juros_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorJuros
                    Next
                
                Case "Valor_Multa_Det"
                
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorMulta
                    Next
                
                Case "Valor_Diferenca_Det"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    For Each objBaixaParcPagarDet In gobjChequesPagAvulso.objColBaixaParcPagar
                        If iMotivoDiferenca = 0 Or iMotivoDiferenca = objBaixaParcPagarDet.iMotivoDiferenca Then
                            dValor = objBaixaParcPagarDet.dValorDiferenca
                        Else
                            dValor = 0
                        End If
                        objMnemonicoValor.colValor.Add dValor
                    Next
                
                Case "FilialForn_Conta_Det" 'conta contabil da filial do fornecedor da parcela
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPagAvulso.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

                            If objInfoParcPag.sContaFilForn <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objInfoParcPag.sContaFilForn, sContaTela)
                                If lErro <> SUCESSO Then Error 32190
                                
                                objMnemonicoValor.colValor.Add sContaTela
                            Else
                                objMnemonicoValor.colValor.Add ""
                            End If
        
                        End If
        
                    Next
                
                Case "Num_Titulo_Det"
                
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPagAvulso.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

                            objMnemonicoValor.colValor.Add objInfoParcPag.lNumTitulo
                    
                        End If
        
                    Next
                    
                Case "Fornec_Codigo_Det"
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjChequesPagAvulso.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

                            objMnemonicoValor.colValor.Add objInfoParcPag.lFornecedor
                    
                        End If
        
                    Next
                                    
'?????
                Case "Valor_Diferenca"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    If iMotivoDiferenca = 0 Then
                        dValor = gobjBaixaParcPagar.dValorDiferenca
                    Else
                        If iMotivoDiferenca <= gobjChequesPagAvulso.adValorDiferenca_UBound Then dValor = gobjChequesPagAvulso.adValorDiferenca(iMotivoDiferenca)
                    End If
                    objMnemonicoValor.colValor.Add dValor
                
                Case Else
                
                    Error 32197
        
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
                    If lErro <> SUCESSO Then Error 32189
                
                    objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
                
                Case "Fornecedor_NomeRed"
                
                    objFornecedor.lCodigo = gobjTituloPagar.lFornecedor
                    lErro = CF("Fornecedor_Le", objFornecedor)
                    If lErro <> SUCESSO Then Error 32189
                
                    objMnemonicoValor.colValor.Add objFornecedor.sNomeReduzido
                
                Case "Data_Emissao_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.dtDataEmissao
                
                Case "FilialForn_Conta" 'conta contabil da filial do fornecedor da parcela
        
                    objFilial.lCodFornecedor = gobjTituloPagar.lFornecedor
                    objFilial.iCodFilial = gobjTituloPagar.iFilial
        
                    lErro = CF("FilialFornecedor_Le", objFilial)
                    If lErro <> SUCESSO Then Error 32189
        
                    If objFilial.sContaContabil <> "" Then
        
                        lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 32190
        
                        objMnemonicoValor.colValor.Add sContaTela
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
        
                Case "FilPag_Cta_Transf" 'conta de transferencia da filial pagadora
        
                    lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjChequesPagAvulso.iFilialEmpresaCta, sContaContabil)
                    If lErro <> SUCESSO Then Error 32193
        
                    If sContaContabil <> "" Then
        
                        lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 32194
        
                    Else
        
                        sContaTela = ""
        
                    End If
        
                    objMnemonicoValor.colValor.Add sContaTela
        
                Case "FilNaoPag_Cta_Transf" 'conta de transferencia da filial da parcela
        
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloPagar.iFilialEmpresa, sContaContabil)
                        If lErro <> SUCESSO Then Error 32195
        
                        If sContaContabil <> "" Then
        
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then Error 32196
        
                        Else
        
                            sContaTela = ""
        
                        End If
        
                        objMnemonicoValor.colValor.Add sContaTela
        
                Case "Tipo_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.sSiglaDocumento
        
        
                Case Else
        
                    Error 32197
        
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
        
        Case 32189 To 32196

        Case 32197
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144495)

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
    lErro = objContabAutomatica.Obter_Doc(lDoc, gobjChequesPagAvulso.iFilialEmpresaCta)
    If lErro <> SUCESSO Then Error 32198

    'se contabiliza parcela p/parcela
    If gobjCP.iContabSemDet = 0 And gobjChequesPagAvulso.dtBomPara = DATA_NULA Then
    
        'se a filial pagadora é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If gobjChequesPagAvulso.iFilialEmpresaCta <> gobjTituloPagar.iFilialEmpresa And giContabCentralizada = 0 Then
    
            'grava a contabilizacao na filial pagadora
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeManualFilPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPagAvulso.iFilialEmpresaCta)
            If lErro <> SUCESSO Then Error 32199
    
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloPagar.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32200
    
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeManualFilNaoPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloPagar.iFilialEmpresa, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32201
    
        Else
    
            'grava a contabilizacao na filial pagadora (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeManual", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPagAvulso.iFilialEmpresaCta, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32202
    
        End If

    Else 'se contabiliza o cheque como um todo
    
        GridParcelas.Tag = gobjChequesPagAvulso.objColBaixaParcPagar.Count
        
        Set gobjBaixaParcPagar = New ClassBaixaParcPagar
        
        With gobjBaixaParcPagar
            .dValorBaixado = gobjChequesPagAvulso.dValorBaixado
            .dValorDesconto = gobjChequesPagAvulso.dValorDesconto
            .dValorJuros = gobjChequesPagAvulso.dValorJuros
            .dValorMulta = gobjChequesPagAvulso.dValorMulta
        End With
        
        For iIndice = gobjChequesPagAvulso.adValorDiferenca_LBound To gobjChequesPagAvulso.adValorDiferenca_UBound
        
            dValorDiferenca = Round(dValorDiferenca + gobjChequesPagAvulso.adValorDiferenca(iIndice), 2)
        
        Next
        
        gobjBaixaParcPagar.dValorDiferenca = dValorDiferenca
        
        'grava a contabilizacao na filial pagadora
        If gobjChequesPagAvulso.dtBomPara = DATA_NULA Then
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequeManualRes", gobjBaixaPagar.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPagAvulso.iFilialEmpresaCta)
        Else
            lErro = objContabAutomatica.Gravar_Registro(Me, "ChequePreManualRes", gobjChequesPagAvulso.lNumIntDocChequePre, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjChequesPagAvulso.iFilialEmpresaCta)
        End If
        If lErro <> SUCESSO Then Error 32202
    
    End If
    
    GeraContabilizacao = SUCESSO

    Exit Function

Erro_GeraContabilizacao:

    GeraContabilizacao = Err

    Select Case Err

        Case 32198 To 32202

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144496)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144497)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUE_MANUAL_P4
    Set Form_Load_Ocx = Me
    Caption = "Cheque Manual - Passo 4"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequePagAvulso4"
    
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

    'Chama a tela do passo Anterior
    Call Chama_Tela("ChequePagAvulso3", gobjChequesPagAvulso)
       
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

Private Sub BotaoIntAtualiza_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    
    End If
    
    'Fecha a tela
    Unload Me

End Sub


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

