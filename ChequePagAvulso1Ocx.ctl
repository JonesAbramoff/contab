VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ChequePagAvulso1Ocx 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   6150
   Begin VB.Frame FrameBomPara 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   435
      Left            =   2310
      TabIndex        =   28
      Top             =   4470
      Width           =   2505
      Begin MSComCtl2.UpDown UpDownDataBomPara 
         Height          =   300
         Left            =   2190
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBomPara 
         Height          =   300
         Left            =   1035
         TabIndex        =   30
         Top             =   90
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelBomPara 
         AutoSize        =   -1  'True
         Caption         =   "Bom Para:"
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
         Left            =   105
         TabIndex        =   31
         Top             =   135
         Width           =   885
      End
   End
   Begin VB.CheckBox CheckChequePre 
      Caption         =   "Cheque Pré-Datado"
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
      Left            =   300
      TabIndex        =   27
      Top             =   4590
      Width           =   2040
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   2250
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5025
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1125
         Picture         =   "ChequePagAvulso1Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   90
         Picture         =   "ChequePagAvulso1Ocx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.ComboBox ContaCorrente 
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   135
      Width           =   1965
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecionar os Títulos que Atendam às Seguintes Condições"
      Height          =   3255
      Left            =   270
      TabIndex        =   16
      Top             =   1170
      Width           =   5610
      Begin VB.Frame Frame3 
         Caption         =   "Portador"
         Height          =   675
         Left            =   465
         TabIndex        =   17
         Top             =   2355
         Width           =   4584
         Begin VB.OptionButton OptionApenas 
            Caption         =   "Apenas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1305
            TabIndex        =   9
            Top             =   300
            Width           =   990
         End
         Begin VB.OptionButton OptionQualquer 
            Caption         =   "Qualquer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   108
            TabIndex        =   8
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox Portador 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2316
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   264
            Width           =   2130
         End
      End
      Begin VB.TextBox Parcela 
         Height          =   285
         Left            =   3495
         TabIndex        =   7
         Top             =   1950
         Width           =   750
      End
      Begin VB.TextBox Titulo 
         Height          =   285
         Left            =   1650
         TabIndex        =   6
         Top             =   1950
         Width           =   945
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   1650
         TabIndex        =   5
         Top             =   1406
         Width           =   2604
      End
      Begin MSComCtl2.UpDown UpDownDataVencimento 
         Height          =   300
         Left            =   2805
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   300
         Left            =   1650
         TabIndex        =   3
         Top             =   345
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fornecedor 
         Height          =   300
         Left            =   1650
         TabIndex        =   4
         Top             =   870
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento Até:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   375
         Width           =   1410
      End
      Begin VB.Label LabelParcela 
         AutoSize        =   -1  'True
         Caption         =   "Parcela:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   20
         Top             =   1995
         Width           =   720
      End
      Begin VB.Label LabelTitulo 
         AutoSize        =   -1  'True
         Caption         =   "Título:"
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
         Left            =   990
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   1995
         Width           =   585
      End
      Begin VB.Label LabelFornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor:"
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
         Left            =   540
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   " Filial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1050
         TabIndex        =   23
         Top             =   1470
         Width           =   525
      End
   End
   Begin MSComCtl2.UpDown UpDownDataEmissao 
      Height          =   300
      Left            =   5685
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   135
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataEmissao 
      Height          =   300
      Left            =   4530
      TabIndex        =   1
      Top             =   135
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownDataContabil 
      Height          =   300
      Left            =   5685
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox DataContabil 
      Height          =   300
      Left            =   4530
      TabIndex        =   2
      Top             =   660
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Contábil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3690
      TabIndex        =   24
      Top             =   705
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3675
      TabIndex        =   25
      Top             =   180
      Width           =   765
   End
   Begin VB.Label LabelContaCorrente 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   345
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   26
      Top             =   195
      Width           =   570
   End
End
Attribute VB_Name = "ChequePagAvulso1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private iChequeAlterado As Integer

Private gobjChequesPagAvulso As ClassChequesPagAvulso
Private WithEvents objEventoContaCorrente As AdmEvento
Attribute objEventoContaCorrente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoTitulo As AdmEvento
Attribute objEventoTitulo.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoSeguir_Click()

Dim lErro As Long
Dim iCodConta As Integer
Dim iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_BotaoSeguir_Click

    'Verifica se a ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 15696

'    'Verifica se o ProxCheque está preenchido
'    If Len(Trim(ProxCheque.Text)) = 0 Then Error 15700

    'Verifica se a DataEmissao está preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Error 15701

    'Verifica se a DataVencimento está preenchida
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Error 15703

    If CheckChequePre.Value = vbChecked Then
        If Len(Trim(DataBomPara.ClipText)) = 0 Then Error 15700
    End If

    If gobjChequesPagAvulso.colInfoParcPag.Count = 0 Or iChequeAlterado = REGISTRO_ALTERADO Then
        Set gobjChequesPagAvulso = New ClassChequesPagAvulso
    End If
    
    'Extrai o Código da Conta Corrente que está na tela
    iCodConta = Codigo_Extrai(ContaCorrente.Text)

    'Passa o Código da Conta para o Obj
    objContaCorrenteInt.iCodigo = iCodConta

    'Lê os dados da Conta
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 15697

    'Se a Conta não estiver cadastrada --> erro
    If lErro = 11807 Then Error 15698

    'Se a Conta não é Bancária
    If objContaCorrenteInt.iCodBanco = 0 Then Error 15699

    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then

        'Se a Conta não é da Filial selecionada
        If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15732

    End If

    'Se o Portador não for Qualquer
    If OptionQualquer.Value = False Then

        'Verifica se o Portador está preenchido
        If Len(Trim(Portador.Text)) = 0 Then Error 15710

    End If

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        'Verifica se a DataContabil está preenchida
        If Len(Trim(DataContabil.ClipText)) = 0 Then Error 15702

        'Verifica se a DataContabil é maior ou igual que a DataEmissao
        If CDate(DataContabil.Text) < CDate(DataEmissao.Text) Then Error 15704

        gobjChequesPagAvulso.dtContabil = CDate(DataContabil.Text)

    Else

        gobjChequesPagAvulso.dtContabil = gdtDataAtual

    End If

    gobjChequesPagAvulso.lNumImpressao = 0

    'Passa os dados da tela para o Obj global
    gobjChequesPagAvulso.dtEmissao = CDate(DataEmissao.Text)
    gobjChequesPagAvulso.dtVencto = CDate(DataVencimento.Text)
    gobjChequesPagAvulso.iCta = iCodConta

    If CheckChequePre.Value = vbChecked Then
        gobjChequesPagAvulso.dtBomPara = StrParaDate(DataBomPara.Text)
    Else
        gobjChequesPagAvulso.dtBomPara = DATA_NULA
    End If
    
    'Se Filial foi preenchida
    If Len(Trim(Filial.Text)) <> 0 Then

        gobjChequesPagAvulso.iFilial = Codigo_Extrai(Filial.Text)

    'Se Filial não foi preenchida
    Else

        'Inicializa a Filial no Obj global
        gobjChequesPagAvulso.iFilial = 0

    End If

    'Se Título estiver preenchido
    If Len(Trim(Titulo.Text)) <> 0 Then

        'Se a Filial não foi preenchida
        If Len(Trim(Filial.Text)) = 0 Then Error 15853

        gobjChequesPagAvulso.lNumTitulo = CLng(Titulo.Text)

    'Se Título não estiver preenchido
    Else

        'Inicializa o Número do Título no Obj global
        gobjChequesPagAvulso.lNumTitulo = 0

    End If

    'Se Parcela foi preenchida
    If Len(Trim(Parcela.Text)) <> 0 Then

        'Se o Título não foi preenchido
        If Len(Trim(Titulo.Text)) = 0 Then Error 15854

        gobjChequesPagAvulso.iNumParcela = CInt(Parcela.Text)

    'Se Parcela não foi preenchida
    Else

        'Inicializa o Número de parcelas no Obj global
        gobjChequesPagAvulso.iNumParcela = 0

    End If

    'Se a Opção Apenas um portador estiver selecionada
    If OptionApenas.Value = True Then gobjChequesPagAvulso.iPortador = Codigo_Extrai(Portador.Text)

    'Se a Opção Qualquer portador estiver selecionada
    If OptionQualquer.Value = True Then gobjChequesPagAvulso.iQualquerPortador = QUALQUER_PORTADOR
    
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê os dados do Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 15695
        
        'Se não encontrou --> erro
        If lErro = 6681 Then Error 61126
            
    End If
    
    gobjChequesPagAvulso.lFornecedor = objFornecedor.lCodigo
''    gobjChequesPagAvulso.lNumCheque = CLng(ProxCheque.Text)
    
    If gobjChequesPagAvulso.colInfoParcPag.Count = 0 Or iChequeAlterado = REGISTRO_ALTERADO Then
        'Preenche coleção de parcelas a pagar
        lErro = CF("ParcelasPagar_Le_ChequesPagAvulso", gobjChequesPagAvulso)
        If lErro <> SUCESSO Then Error 15794
    End If
    
    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequePagAvulso2", gobjChequesPagAvulso)

    'Fecha a tela
    Unload Me

    Exit Sub

Erro_BotaoSeguir_Click:

    Select Case Err

        Case 15696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)

        Case 15697, 15794

        Case 15698
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, iCodConta)

        Case 15699
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

''        Case 15700
''            lErro = Rotina_Erro(vbOKOnly, "ERRO_PROXCHEQUE_NAO_INFORMADO", Err)

        Case 15701, 15702, 15703, 15700
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case 15704
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATACONTABIL_MENOR_DATAEMISSAO", Err)

        Case 15710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_INFORMADO", Err)

        Case 15732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case 15853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_INFORMADO_SEM_FILIAL", Err)

        Case 15854
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_INFORMADA_SEM_TITULO", Err)
    
        Case 61126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144460)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objChequesPagAvulso As ClassChequesPagAvulso) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    Set gobjChequesPagAvulso = objChequesPagAvulso

    'Se há Cheques Avulsos
    If (objChequesPagAvulso.colInfoParcPag.Count > 0) Then
    
        If objChequesPagAvulso.iCta <> 0 Then
        
            ContaCorrente.Text = objChequesPagAvulso.iCta
            Call ContaCorrente_Validate(bSGECancelDummy)
        
        End If
        
        DataEmissao.Text = Format(objChequesPagAvulso.dtEmissao, "dd/mm/yy")
        DataVencimento.Text = Format(objChequesPagAvulso.dtVencto, "dd/mm/yy")
        
        If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
                DataContabil.Text = Format(objChequesPagAvulso.dtContabil, "dd/mm/yy")
        End If
        
        If gobjChequesPagAvulso.lFornecedor <> 0 Then
            
            Fornecedor.Text = gobjChequesPagAvulso.lFornecedor
            Call Fornecedor_Validate(bSGECancelDummy)
        
        End If
        
        If gobjChequesPagAvulso.iFilial <> 0 Then
            
            Filial.Text = gobjChequesPagAvulso.iFilial
            Call Filial_Validate(bSGECancelDummy)
        
        End If
        
        If gobjChequesPagAvulso.lNumTitulo <> 0 Then
            
            Titulo.Text = gobjChequesPagAvulso.lNumTitulo
            Call Titulo_Validate(bSGECancelDummy)
        
        End If
        
        If gobjChequesPagAvulso.iNumParcela <> 0 Then
            
            Parcela.Text = gobjChequesPagAvulso.iNumParcela
            Call Parcela_Validate(bSGECancelDummy)
        
        End If
        
        'Se a Opção Qualquer portador estiver selecionada
        If gobjChequesPagAvulso.iQualquerPortador = QUALQUER_PORTADOR Then
            OptionQualquer.Value = True
            OptionApenas.Value = False
        ElseIf gobjChequesPagAvulso.iPortador > 0 Then
            
            OptionQualquer.Value = False
            OptionApenas.Value = True
            
            For iIndice = 0 To Portador.ListCount - 1
            
                If Portador.ItemData(iIndice) = gobjChequesPagAvulso.iPortador Then
                    Portador.ListIndex = iIndice
                    Exit For
                End If
            Next
            
        End If
    
        If objChequesPagAvulso.dtBomPara <> DATA_NULA Then
            CheckChequePre.Value = vbChecked
            Call DateParaMasked(DataBomPara, objChequesPagAvulso.dtBomPara)
        End If
    
    End If
    
    iChequeAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144461)

    End Select

    Exit Function

End Function

Private Sub CheckChequePre_Click()
    
    iChequeAlterado = REGISTRO_ALTERADO
    If CheckChequePre.Value = vbChecked Then
        FrameBomPara.Enabled = True
        LabelBomPara.Enabled = True
    Else
        Call DateParaMasked(DataBomPara, DATA_NULA)
        FrameBomPara.Enabled = False
        LabelBomPara.Enabled = False
    End If
    
End Sub

Private Sub ContaCorrente_Change()

    iChequeAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaCorrente_Click()

    iChequeAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se a Conta existe na Combo. Se existir, seleciona
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15726

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Passa o Código da Conta para o Obj
        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 15727

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 15728

        'Se a Conta não é Bancária
        If objContaCorrenteInt.iCodBanco = 0 Then Error 15729

        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Se a Conta não é da Filial selecionada
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 15731

        End If

        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 15730

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 15726, 15727

        Case 15728
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 15729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_BANCARIA", Err, ContaCorrente.Text)

        Case 15730
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)

        Case 15731
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, ContaCorrente.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144462)

    End Select

    Exit Sub

End Sub

Private Sub DataContabil_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataContabil_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataContabil)

End Sub

Private Sub DataContabil_Validate(Cancel As Boolean)
'Critica a DataContabil

Dim lErro As Long

On Error GoTo Erro_DataContabil_Validate

    'Verifica se a DataContabil está vazia
    If Len(DataContabil.ClipText) > 0 Then

        'Verifica se a DataContabil é válida
        lErro = Data_Critica(DataContabil.Text)
        If lErro <> SUCESSO Then Error 15711

    End If

    Exit Sub

Erro_DataContabil_Validate:

    Cancel = True


    Select Case Err

        Case 15711

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144463)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissao_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
'Critica a DataEmissao

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a DataEmissao está vazia
    If Len(DataEmissao.ClipText) > 0 Then

        'Verifica se a DataEmissao é válida
        lErro = Data_Critica(DataEmissao.Text)
        If lErro <> SUCESSO Then Error 15712

    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 15712

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144464)

    End Select

    Exit Sub

End Sub

Private Sub DataVencimento_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVencimento_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataVencimento)

End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
'Critica a DataVencimento

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a DataVencimento está vazia
    If Len(DataVencimento.ClipText) > 0 Then

        'Verifica se a DataVencimento é válida
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 15713

    End If

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case Err

        Case 15713

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144465)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim iCodFilial As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Se a Filial não foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Se é a Filial selecionada na Combo
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se a Filial existe na Combo. Se existir, seleciona
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 15714

    'Se a Filial(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Verifica se o Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 15715

        'Lê os dados do Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15716

        'Passada os Códigos de Fornecedor e Filial para o Obj
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código em questão
        lErro = CF("FilialFornecedor_Le", objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 12929 Then Error 15717

        'Se não existe Filial com o Código em questão
        If lErro = 12929 Then Error 15718

        'Coloca a Filial na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Se a Filial(STRING) não existe na Combo
    If lErro = 6731 Then Error 15719

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 15714, 15717

        Case 15715
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 15716

        Case 15718
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_FILIALFORNECEDOR_INEXISTENTE", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
            End If

        Case 15719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144466)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()
'Preencher a combo de ctas c/ctas de bcos em que o layout p/cheque esta preenchido.

Dim lErro As Long
Dim colCodigoNomeConta As New AdmColCodigoNome
Dim objCodigoNomeConta As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoContaCorrente = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoTitulo = New AdmEvento
    Set objEventoParcela = New AdmEvento

    'Carrega a Coleção de Contas
    lErro = CF("ContasCorrentes_Bancarias_Le_CodigosNomesRed", colCodigoNomeConta)
    If lErro <> SUCESSO Then Error 15673

    'Preenche a ComboBox CodConta com os objetos da coleção de Contas
    For Each objCodigoNomeConta In colCodigoNomeConta

        ContaCorrente.AddItem CStr(objCodigoNomeConta.iCodigo) & SEPARADOR & objCodigoNomeConta.sNome
        ContaCorrente.ItemData(ContaCorrente.NewIndex) = objCodigoNomeConta.iCodigo

    Next

    'Seleciona uma das Contas
    If ContaCorrente.ListCount <> 0 Then ContaCorrente.Text = ContaCorrente.List(0)

    'Carrega a ComboBox Portador
    lErro = Portador_Carrega(Portador)
    If lErro <> SUCESSO Then Error 15674

    'Preenche as Datas com a data corrente do sistema
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataVencimento.Text = Format(gdtDataAtual, "dd/mm/yy")

    'Mostra qualquer Portador como default
    OptionQualquer.Value = True
    OptionApenas.Value = False
    Portador.Enabled = False

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) <> MODULO_ATIVO Then
        Label10.Enabled = False
        DataContabil.Enabled = False
        UpDownDataContabil.Enabled = False
    Else
        DataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
    End If

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 15673, 15674

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144467)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoContaCorrente = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoTitulo = Nothing
    Set objEventoParcela = Nothing

    Set gobjChequesPagAvulso = Nothing

End Sub

Private Sub Fornecedor_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Limpa a Combo de Filiais
    Filial.Clear

    'Se Fornecedor não está preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then Exit Sub

    'Lê os dados do Fornecedor
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> AD_SQL_SUCESSO Then Error 15676

    'Lê os dados da Filial do Fornecedor
    lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
    If lErro <> AD_SQL_SUCESSO Then Error 15677

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    'Seleciona filial na Combo Filial
    Call CF("Filial_Seleciona", Filial, iCodFilial)

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True


    Select Case Err

        Case 15676, 15677

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144468)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o Fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama a Tela de com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub LabelContaCorrente_Click()
'Chamada do Browse de Contas

Dim colSelecao As Collection
Dim objConta As New ClassContasCorrentesInternas
Dim objChequePagAvulso As New ClassChequesPagAvulso

    'Se a ContaCorrente não está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then

        'Passa o Código da Conta que está na tela para o Obj
        objConta.iCodigo = 0

    'Se a ContaCorrente está preenchida
    Else

        'Passa o Código da Conta que está na tela para o Obj
        objConta.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    End If

    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then

        'Chama a tela com a lista das contas correntes bancarias
        Call Chama_Tela("CtaCorrBancariaLista", colSelecao, objConta, objEventoContaCorrente)


    Else
        'Chama a tela com a lista de todas as contas correntes bancarias
        Call Chama_Tela("CtaCorrBancariaTodasLista", colSelecao, objConta, objEventoContaCorrente)

    End If


    Exit Sub

End Sub

Private Sub LabelParcela_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objParcelaPagar As New ClassParcelaPagar
Dim objFornecedor As New ClassFornecedor
Dim iFilial As Integer

On Error GoTo Erro_LabelParcela_Click

    If Len(Trim(Titulo.Text)) = 0 Then Error 57461

    If Len(Trim(Fornecedor.Text)) = 0 Then Error 57477
    
    If Len(Trim(Filial.Text)) = 0 Then Error 57479

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 57478

    'Se não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then Error 57480

    iFilial = Codigo_Extrai(Filial.Text)

    colSelecao.Add objFornecedor.lCodigo
    colSelecao.Add iFilial
    colSelecao.Add StrParaLong(Titulo.Text)
    
    'Chama a tela
    Call Chama_Tela("ParcelasPagLista", colSelecao, objParcelaPagar, objEventoParcela)

    Exit Sub

Erro_LabelParcela_Click:

    Select Case Err

        Case 57461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 57477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
            
        Case 57478
            'Tratado na rotina chamada
            
        Case 57479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 57480
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144469)

    End Select

    Exit Sub

End Sub

Private Sub LabelTitulo_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelTitulo_Click

    'Se Forncedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 57458

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 57459

    objFornecedor.sNomeReduzido = Fornecedor.Text
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 57460

    objTituloPagar.lFornecedor = objFornecedor.lCodigo

    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)

    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial

    'Chama Tela TitulosPagLista
    Call Chama_Tela("TitulosPagLista", colSelecao, objTituloPagar, objEventoTitulo)

    Exit Sub

Erro_LabelTitulo_Click:

    Select Case Err

        Case 57458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 57459
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 57460
            'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144470)

    End Select

    Exit Sub


End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objParcelaPagar As ClassParcelaPagar

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaPagar = obj1

    If Not (objParcelaPagar Is Nothing) Then
        Parcela.Text = CStr(objParcelaPagar.iNumParcela)
        Call Parcela_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144471)

    End Select

    Exit Sub

End Sub

Private Sub objEventoTitulo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

    Set objTituloPagar = obj1

    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(bSGECancelDummy)

    Filial.Text = objTituloPagar.iFilial
    Call Filial_Validate(bSGECancelDummy)

    Titulo.Text = CStr(objTituloPagar.lNumTitulo)

    Me.Show

    Exit Sub


End Sub

Private Sub OptionApenas_Click()

    'Se a Opção Apenas um portador está selecionada, habilita a Combo de Portador
    If OptionApenas.Value = True Then Portador.Enabled = True

    'Se a Opção Apenas um portador não está selecionada
    If OptionApenas.Value = False Then

        'Desabilita a Combo de Portador
        Portador.ListIndex = -1
        Portador.Enabled = False

    End If

End Sub

Private Sub OptionQualquer_Click()

    'Se a Opção qualquer portador estiver selecionada
    If OptionQualquer.Value = True Then

        'Desabilita a Combo de Portador
        Portador.ListIndex = -1
        Portador.Enabled = False

    End If

    'Se a Opção qualquer portador não estiver selecionada, habilita a Combo de Portador
    If OptionQualquer.Value = False Then Portador.Enabled = True

End Sub

Private Sub Parcela_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Se Parcela não está preenchida
    If Len(Trim(Parcela.Text)) = 0 Then Exit Sub

    'Verifica se a Parcela que está na tela é um número inteiro
    lErro = Inteiro_Critica(Parcela.Text)
    If lErro <> SUCESSO Then Error 15693

    Exit Sub

Erro_Parcela_Validate:

    Cancel = True


    Select Case Err

        Case 15693

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144472)

    End Select

    Exit Sub

End Sub

Private Sub Portador_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Portador_Click()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Titulo_Change()

    iChequeAlterado = REGISTRO_ALTERADO

End Sub

''Private Sub ProxCheque_Validate(Cancel As Boolean)
''
''Dim lErro As Long
''
''On Error GoTo Erro_ProxCheque_Validate
''
''    'Verifica se o ProxCheque está preenchido
''    If Len(Trim(ProxCheque.Text)) = 0 Then Exit Sub
''
''    'Verifica se o ProxCheque que está na tela é um número do tipo Long
''    lErro = Long_Critica(ProxCheque.Text)
''    If lErro <> SUCESSO Then Error 15772
''
''    Exit Sub
''
''Erro_ProxCheque_Validate:

''    Cancel = True

''
''    Select Case Err
''
''        Case 15772
''
''        Case Else
''             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144473)
''
''    End Select
''
''    Exit Sub
''
''End Sub
''
Private Sub Titulo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer

On Error GoTo Erro_Titulo_Validate

    'Se Título não está preenchido
    If Len(Trim(Titulo.Text)) = 0 Then Exit Sub

    'Verifica se Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15678

    'Verifica se Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Error 15679

    'Verifica se o Titulo que está na tela é um número do tipo Long
    lErro = Long_Critica(Titulo.Text)
    If lErro <> SUCESSO Then Error 15687

    Exit Sub

Erro_Titulo_Validate:

    Cancel = True


    Select Case Err

        Case 15678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 15679
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 15680, 15687

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144474)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaCorrente_evSelecao(obj1 As Object)

Dim objConta As ClassContasCorrentesInternas

    'Retorna os dados do Obj para a tela
    Set objConta = obj1

    ContaCorrente.Text = CStr(objConta.iCodigo)
    Call ContaCorrente_Validate(bSGECancelDummy)

    'Exibe a tela
    Me.Show

    Exit Sub

End Sub

Private Sub UpDownDataContabil_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_DownClick

    'Diminui a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15720

    Exit Sub

Erro_UpDownDataContabil_DownClick:

    Select Case Err

        Case 15720

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144475)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataContabil_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataContabil_UpClick

    'Aumenta a DataContabil em 1 dia
    lErro = Data_Up_Down_Click(DataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15721

    Exit Sub

Erro_UpDownDataContabil_UpClick:

    Select Case Err

        Case 15721

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144476)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Diminui a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15722

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case Err

        Case 15722

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144477)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aumenta a DataEmissao em 1 dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15723

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case Err

        Case 15723

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144478)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_DownClick

    'Diminui a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15724

    Exit Sub

Erro_UpDownDataVencimento_DownClick:

    Select Case Err

        Case 15724

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144479)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencimento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVencimento_UpClick

    'Aumenta a DataVencimento em 1 dia
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15725

    Exit Sub

Erro_UpDownDataVencimento_UpClick:

    Select Case Err

        Case 15725

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144480)

    End Select

    Exit Sub

End Sub

Private Function Portador_Carrega(objComboBox As ComboBox) As Long
'Carrega a ComboBox de Portadores

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As New AdmCodigoNome

On Error GoTo Erro_Portador_Carrega

    'Lê Codigos e NomesReduzidos de Portadores
    lErro = CF("Portadores_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 15632

    'Preenche a ComboBox Portadores
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    Portador_Carrega = SUCESSO

    Exit Function

Erro_Portador_Carrega:

    Portador_Carrega = Err

    Select Case Err

        Case 15632

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144481)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CHEQUE_MANUAL_P1
    Set Form_Load_Ocx = Me
    Caption = "Cheque Manual - Passo 1"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ChequePagAvulso1"

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is ContaCorrente Then
            Call LabelContaCorrente_Click
        ElseIf Me.ActiveControl Is Fornecedor Then
            Call LabelFornecedor_Click
        ElseIf Me.ActiveControl Is Titulo Then
            Call LabelTitulo_Click
        ElseIf Me.ActiveControl Is Parcela Then
            Call LabelParcela_Click
        End If

    End If

End Sub





Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub LabelParcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelParcela, Source, X, Y)
End Sub

Private Sub LabelParcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelParcela, Button, Shift, X, Y)
End Sub

Private Sub LabelTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTitulo, Source, X, Y)
End Sub

Private Sub LabelTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTitulo, Button, Shift, X, Y)
End Sub

Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
End Sub

Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub

Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContaCorrente, Source, X, Y)
End Sub

Private Sub LabelContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub UpDownDataBomPara_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBomPara_DownClick

    'Diminui a DataBomPara em 1 dia
    lErro = Data_Up_Down_Click(DataBomPara, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15724

    Exit Sub

Erro_UpDownDataBomPara_DownClick:

    Select Case Err

        Case 15724

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144479)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataBomPara_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBomPara_UpClick

    'Aumenta a DataBomPara em 1 dia
    lErro = Data_Up_Down_Click(DataBomPara, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15725

    Exit Sub

Erro_UpDownDataBomPara_UpClick:

    Select Case Err

        Case 15725

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144480)

    End Select

    Exit Sub

End Sub

