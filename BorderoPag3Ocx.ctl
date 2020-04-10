VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoPag3Ocx 
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   ScaleHeight     =   3330
   ScaleWidth      =   4065
   Begin MSFlexGridLib.MSFlexGrid GridParcelas 
      Height          =   735
      Left            =   3720
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1296
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
      Left            =   465
      TabIndex        =   0
      Top             =   2055
      Width           =   3105
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   637
      ScaleHeight     =   495
      ScaleWidth      =   2730
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2595
      Width           =   2790
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   90
         Picture         =   "BorderoPag3Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
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
         Height          =   330
         Left            =   1065
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2220
         Picture         =   "BorderoPag3Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   457
      TabIndex        =   4
      Top             =   1545
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
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
      Left            =   660
      TabIndex        =   6
      Top             =   630
      Width           =   1575
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
      Left            =   285
      TabIndex        =   7
      Top             =   1080
      Width           =   1950
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
      Left            =   540
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2272
      TabIndex        =   9
      Top             =   1050
      Width           =   1350
   End
   Begin VB.Label TotalTitulos 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2272
      TabIndex        =   10
      Top             =   600
      Width           =   1350
   End
End
Attribute VB_Name = "BorderoPag3Ocx"
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

Public gobjBorderoPagEmissao As ClassBorderoPagEmissao

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjTituloPagar As New ClassTituloPagar
Private gobjParcelaPagar As New ClassParcelaPagar
Private gobjBaixaParcPagar As New ClassBaixaParcPagar
Private gobjBaixaPagar As New ClassBaixaPagar
Private gsContaCtaCorrente As String 'conta contabil da conta corrente
Private gsContaFilPag As String 'conta contabil da filial pagadora
Private giFilialEmpresaConta As Integer 'filial empresa possuidora da conta corrente utilizada p/o pagto
Private gsContaFornecedores As String

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
        gobjBorderoPagEmissao.objTelaAtualizacao = Me
        lErro = CF("BorderoPagto_Criar", gobjBorderoPagEmissao)
        giExecutando = ESTADO_PARADO

        BotaoIntAtualiza.Enabled = False

        If lErro <> SUCESSO And lErro <> 26428 Then Error 7783

        If lErro = 26428 Then Error 41424 'interrompeu

        'Chama a tela do passo seguinte
        Call Chama_Tela("BorderoPag4", gobjBorderoPagEmissao)

        'Fecha a tela
        Unload Me
    
    End If

    Exit Sub

Erro_BotaoAtualizar_Click:

    Select Case Err

        Case 41424
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
            'Unload Me

        Case 7783
            'Unload Me

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143807)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objBorderoPagEmissao As ClassBorderoPagEmissao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO

    Set gobjBorderoPagEmissao = objBorderoPagEmissao

    Set gobjBorderoPagEmissao.objEvolucao = Me

    'Passa para a tela os dados dos Títulos selecionados
    TotalTitulos.Caption = CStr(gobjBorderoPagEmissao.iQtdeParcelasSelecionadas)
    TitulosProcessados.Caption = "0"

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143808)

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

        ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)

        giExecutando = ESTADO_ANDAMENTO

    End If

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143809)

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
    giFilialEmpresaConta = 0

    Set gobjBorderoPagEmissao = Nothing
    
End Sub

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, sContaContabil As String, dValor As Double
Dim objFilial As New ClassFilialFornecedor, sContaTela As String, iMotivoDiferenca As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas, bAchou As Boolean, iIndice As Integer
Dim objBaixaParcPagarDet As ClassBaixaParcPagar
Dim objInfoParcPag As ClassInfoParcPag
Dim sContaFormatada As String, iContaPreenchida As Integer, objMnemonico As New ClassMnemonicoCTBValor
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Calcula_Mnemonico

    bAchou = True
    
    'tratar mnemonicos comuns a contab parcela a parcela e contab p/bordero c/um todo
    Select Case objMnemonicoValor.sMnemonico

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
                
        Case "Numero_Bordero"
            objMnemonicoValor.colValor.Add gobjBorderoPagEmissao.lNumero
            
        Case "Conta_Contabil_Conta" 'conta contabil associada a conta corrente utilizada p/o pagto
            'calcula-la apenas uma vez e deixa-la guardada
                
            If gsContaCtaCorrente = "" Then
                
                lErro = CF("ContaCorrenteInt_Le", gobjBorderoPagEmissao.iCta, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 41897
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41898
                                        
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
    
        'se contabiliza o bordero como um todo
        If gobjCP.iContabSemDet = 1 Then
        
            Select Case objMnemonicoValor.sMnemonico
        
                Case "Valor_Forn_Aux"
                    
                    iIndice = 1
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
                
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
                            
                                Set objBaixaParcPagarDet = gobjBorderoPagEmissao.objColBaixaParcPagar(iIndice)
                                dValor = Round(dValor + objBaixaParcPagarDet.dValorBaixado - objBaixaParcPagarDet.dValorDiferenca, 2)
                                
                            End If
                            
                            iIndice = iIndice + 1
        
                        End If
        
                    Next
                    
                    objMnemonicoValor.colValor.Add dValor
                
                Case "Valor_Pago_Det"
                
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        With objBaixaParcPagarDet
                            dValor = Round(.dValorBaixado - .dValorDesconto + .dValorJuros + .dValorMulta + .dValorDiferenca, 2)
                        End With
                        objMnemonicoValor.colValor.Add dValor
                    Next
                        
                Case "Valor_Baixado_Det"
                
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorBaixado
                    Next
                
                Case "Valor_Desconto_Det"
                
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorDesconto
                    Next
                
                Case "Valor_Juros_Det"
                
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorJuros
                    Next
                
                Case "Valor_Multa_Det"
                
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        objMnemonicoValor.colValor.Add objBaixaParcPagarDet.dValorMulta
                    Next
                
                Case "Valor_Diferenca_Det"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    For Each objBaixaParcPagarDet In gobjBorderoPagEmissao.objColBaixaParcPagar
                        If iMotivoDiferenca = 0 Or iMotivoDiferenca = objBaixaParcPagarDet.iMotivoDiferenca Then
                            dValor = objBaixaParcPagarDet.dValorDiferenca
                        Else
                            dValor = 0
                        End If
                        objMnemonicoValor.colValor.Add dValor
                    Next
                
                Case "FilialForn_Conta_Det" 'conta contabil da filial do fornecedor da parcela
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

                            If objInfoParcPag.sContaFilForn <> "" Then
                            
                                lErro = Mascara_RetornaContaTela(objInfoParcPag.sContaFilForn, sContaTela)
                                If lErro <> SUCESSO Then Error 41896
                                
                                objMnemonicoValor.colValor.Add sContaTela
                            Else
                                objMnemonicoValor.colValor.Add ""
                            End If
        
                        End If
        
                    Next
                
                Case "Num_Titulo_Det"
                
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
                
                        'pular parcelas com iseqcheque = 0
                        If objInfoParcPag.iSeqCheque <> 0 Then

                            objMnemonicoValor.colValor.Add objInfoParcPag.lNumTitulo
                    
                        End If
        
                    Next
                    
                Case "Fornec_Codigo_Det"
                    
                    'percorrer a colecao de parcelas a pagar
                    For Each objInfoParcPag In gobjBorderoPagEmissao.colInfoParcPag
                
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
                        If iMotivoDiferenca <= gobjBorderoPagEmissao.adValorDiferenca_UBound Then dValor = gobjBorderoPagEmissao.adValorDiferenca(iMotivoDiferenca)
                    End If
                    objMnemonicoValor.colValor.Add dValor
                
                Case Else
                
                    Error 41652
        
            End Select
        
        Else 'se contabiliza parcela a parcela
        
            Select Case objMnemonicoValor.sMnemonico
        
                Case "Valor_Diferenca"
                    iMotivoDiferenca = objMnemonicoValor.vParam(1)
                    If iMotivoDiferenca = gobjBaixaParcPagar.iMotivoDiferenca Then dValor = gobjBaixaParcPagar.dValorDiferenca
                    objMnemonicoValor.colValor.Add dValor
                
                Case "FilialForn_Conta" 'conta contabil da filial do fornecedor da parcela
                    
                    objFilial.lCodFornecedor = gobjTituloPagar.lFornecedor
                    objFilial.iCodFilial = gobjTituloPagar.iFilial
                    
                    lErro = CF("FilialFornecedor_Le", objFilial)
                    If lErro <> SUCESSO Then Error 41895
                    
                    If objFilial.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 41896
                        
                        objMnemonicoValor.colValor.Add sContaTela
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
        
                Case "FilPag_Cta_Transf" 'conta de transferencia da filial pagadora
        
                    If gsContaFilPag = "" Then
                    
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(giFilialEmpresaConta, sContaContabil)
                        If lErro <> SUCESSO Then Error 41899
                        
                        If sContaContabil <> "" Then
                        
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then Error 41900
                            
                        Else
                        
                            sContaTela = ""
                            
                        End If
                    
                        gsContaFilPag = sContaTela
                        
                    End If
                    
                    objMnemonicoValor.colValor.Add gsContaFilPag
                
                Case "FilNaoPag_Cta_Transf" 'conta de transferencia da filial da parcela
        
                        lErro = gobjContabAutomatica.Obter_ContaContabilTransferencia(gobjTituloPagar.iFilialEmpresa, sContaContabil)
                        If lErro <> SUCESSO Then Error 41901
                        
                        If sContaContabil <> "" Then
                        
                            lErro = Mascara_RetornaContaTela(sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then Error 41902
                        
                        Else
                        
                            sContaTela = ""
                            
                        End If
                        
                        objMnemonicoValor.colValor.Add sContaTela
                
                Case "Num_Titulo"
                
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lNumTitulo
                    
                Case "Fornecedor_Codigo"
                    
                    objMnemonicoValor.colValor.Add gobjTituloPagar.lFornecedor
                            
                Case "Tipo_Titulo"
                    objMnemonicoValor.colValor.Add gobjTituloPagar.sSiglaDocumento
                            
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
                            
                Case Else
                
                    Error 41652
        
            End Select
                
        End If
        
    End If
    
    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 41895 To 41902
        
        Case 56808
        
        Case 56809
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 41652
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143810)

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

    'se ainda nao obtive a filial empresa pagadora
    If giFilialEmpresaConta = 0 Then
    
        lErro = CF("ContaCorrenteInt_Le", gobjBorderoPagEmissao.iCta, objContasCorrentesInternas)
        If lErro <> SUCESSO Then Error 32181
    
        giFilialEmpresaConta = objContasCorrentesInternas.iFilialEmpresa
        
    End If
    
    'obtem numero de doc para a filial pagadora
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresaConta)
    If lErro <> SUCESSO Then Error 32160
    
    'se contabiliza parcela p/parcela
    If gobjCP.iContabSemDet = 0 Then
    
        'se a filial pagadora é diferente da do titulo
        'e a contabilidade é descentralizada por filiais
        If giFilialEmpresaConta <> gobjTituloPagar.iFilialEmpresa And giContabCentralizada = 0 Then
                        
            'grava a contabilizacao na filial pagadora
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoPagtoFilPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
            If lErro <> SUCESSO Then Error 32161
        
            'obtem numero de doc para a filial do titulo
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjTituloPagar.iFilialEmpresa)
            If lErro <> SUCESSO Then Error 32162
        
            'grava a contabilizacao na filial do titulo
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoPagtoFilNaoPag", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjTituloPagar.iFilialEmpresa, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32163
        
        Else
        
            'grava a contabilizacao na filial pagadora (a mesma do titulo)
            lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoPagto", gobjBaixaParcPagar.lNumIntDoc, gobjTituloPagar.lFornecedor, gobjTituloPagar.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta, , , -gobjBaixaParcPagar.dValorBaixado)
            If lErro <> SUCESSO Then Error 32164
        
        End If
    
    Else 'se contabiliza o bordero como um todo
    
        GridParcelas.Tag = gobjBorderoPagEmissao.objColBaixaParcPagar.Count
        
        Set gobjBaixaParcPagar = New ClassBaixaParcPagar
        
        With gobjBaixaParcPagar
            .dValorBaixado = gobjBorderoPagEmissao.dValorBaixado
            .dValorDesconto = gobjBorderoPagEmissao.dValorDesconto
            .dValorJuros = gobjBorderoPagEmissao.dValorJuros
            .dValorMulta = gobjBorderoPagEmissao.dValorMulta
        End With
        
        For iIndice = gobjBorderoPagEmissao.adValorDiferenca_LBound To gobjBorderoPagEmissao.adValorDiferenca_UBound
        
            dValorDiferenca = Round(dValorDiferenca + gobjBorderoPagEmissao.adValorDiferenca(iIndice), 2)
        
        Next
        
        gobjBaixaParcPagar.dValorDiferenca = dValorDiferenca
        
        lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoPagtoRes", gobjBaixaPagar.lNumIntBaixa, 0, 0, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresaConta)
        If lErro <> SUCESSO Then Error 32164
    
    End If
    
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32161, 32163, 32164
            '??? retirar a msgbox e usar rotina erro
            'Call Rotina_Erro(vbOKOnly, "ERRO_CONTABILIZACAO_TIT_BORDPAG", Err, gobjTituloPagar.)
            MsgBox ("Erro na contabilização do título " & CStr(gobjTituloPagar.lNumTitulo) & " do fornecedor " & CStr(gobjTituloPagar.lFornecedor) & ". Verifique se está preenchida a conta contábil no cadastro de fornecedores. ")
        
        Case 32160, 32162, 32181
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143811)
     
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143812)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_PAGT_P3
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Pagamento - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoPag3"
    
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
    Call Chama_Tela("BorderoPag2", gobjBorderoPagEmissao)
    
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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

