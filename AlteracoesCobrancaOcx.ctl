VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl AlteracoesCobrancaOcx 
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   KeyPreview      =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   7530
   Begin VB.CommandButton BotaoInstrucoes 
      Caption         =   "Instruções Cadastradas"
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
      Left            =   120
      TabIndex        =   45
      Top             =   5655
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5250
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "AlteracoesCobrancaOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "AlteracoesCobrancaOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "AlteracoesCobrancaOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "AlteracoesCobrancaOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Cobrança"
      Height          =   2310
      Left            =   135
      TabIndex        =   20
      Top             =   3255
      Width           =   7290
      Begin VB.Frame Frame4 
         Caption         =   "Vencimento"
         Height          =   600
         Left            =   210
         TabIndex        =   21
         Top             =   600
         Width           =   6915
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   300
            Left            =   6585
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   210
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox NovoVcto 
            Height          =   300
            Left            =   5490
            TabIndex        =   9
            Top             =   210
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Original: "
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
            TabIndex        =   24
            Top             =   270
            Width           =   705
         End
         Begin VB.Label LabelVctoOriginal 
            Height          =   195
            Left            =   1140
            TabIndex        =   25
            Top             =   270
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Atual: "
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
            Left            =   2430
            TabIndex        =   26
            Top             =   240
            Width           =   570
         End
         Begin VB.Label LabelVctoAtual 
            Height          =   195
            Left            =   3060
            TabIndex        =   27
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Novo: "
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
            Left            =   4830
            TabIndex        =   28
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.ComboBox Instrucao1 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   1283
         Width           =   3570
      End
      Begin VB.ComboBox Instrucao2 
         Height          =   315
         Left            =   1320
         TabIndex        =   12
         Top             =   1838
         Width           =   3570
      End
      Begin VB.ComboBox Ocorrencia 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   233
         Width           =   3570
      End
      Begin MSMask.MaskEdBox Juros 
         Height          =   300
         Left            =   5925
         TabIndex        =   8
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DiasdeProtesto1 
         Height          =   300
         Left            =   6495
         TabIndex        =   11
         Top             =   1290
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DiasdeProtesto2 
         Height          =   300
         Left            =   6510
         TabIndex        =   13
         Top             =   1845
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Instrução 1:"
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
         Left            =   195
         TabIndex        =   29
         Top             =   1343
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Juros:"
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
         Left            =   5340
         TabIndex        =   30
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Dias p/Protesto:"
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
         Left            =   5025
         TabIndex        =   31
         Top             =   1350
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência:"
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
         Left            =   225
         TabIndex        =   32
         Top             =   293
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Dias p/Protesto:"
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
         Left            =   5025
         TabIndex        =   33
         Top             =   1905
         Width           =   1410
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Instrução 2:"
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
         Left            =   195
         TabIndex        =   34
         Top             =   1898
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Situação Atual"
      Height          =   720
      Left            =   120
      TabIndex        =   23
      Top             =   2460
      Width           =   7275
      Begin VB.Label Carteira 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4725
         TabIndex        =   35
         Top             =   330
         Width           =   2280
      End
      Begin VB.Label Cobrador 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1080
         TabIndex        =   36
         Top             =   330
         Width           =   2370
      End
      Begin VB.Label label1007 
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
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
         Left            =   3855
         TabIndex        =   37
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cobrador:"
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
         Left            =   135
         TabIndex        =   38
         Top             =   330
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Identificação"
      Height          =   1740
      Left            =   120
      TabIndex        =   19
      Top             =   660
      Width           =   7275
      Begin VB.CommandButton BotaoProxNum 
         Height          =   300
         Left            =   5010
         Picture         =   "AlteracoesCobrancaOcx.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Numeração Automática"
         Top             =   1320
         Width           =   300
      End
      Begin VB.ComboBox Tipo 
         Height          =   315
         ItemData        =   "AlteracoesCobrancaOcx.ctx":0A7E
         Left            =   915
         List            =   "AlteracoesCobrancaOcx.ctx":0A80
         TabIndex        =   2
         Top             =   840
         Width           =   2190
      End
      Begin VB.ComboBox Filial 
         Height          =   315
         Left            =   4605
         TabIndex        =   1
         Top             =   345
         Width           =   2355
      End
      Begin MSMask.MaskEdBox Cliente 
         Height          =   315
         Left            =   915
         TabIndex        =   0
         Top             =   345
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   300
         Left            =   4035
         TabIndex        =   3
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Parcela 
         Height          =   300
         Left            =   900
         TabIndex        =   4
         Top             =   1320
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Sequencial 
         Height          =   300
         Left            =   4605
         TabIndex        =   5
         Top             =   1320
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "99"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissao 
         Height          =   300
         Left            =   6840
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   810
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   300
         Left            =   5745
         TabIndex        =   47
         Top             =   810
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   4935
         TabIndex        =   48
         Top             =   855
         Width           =   750
      End
      Begin VB.Label LabelSequencial 
         AutoSize        =   -1  'True
         Caption         =   "Seq.Instrução:"
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
         Left            =   3270
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label LabelTipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   375
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   40
         Top             =   900
         Width           =   450
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   41
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4005
         TabIndex        =   42
         Top             =   405
         Width           =   525
      End
      Begin VB.Label NumeroLabel 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   3240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   43
         Top             =   900
         Width           =   720
      End
      Begin VB.Label ClienteLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   44
         Top             =   405
         Width           =   660
      End
   End
End
Attribute VB_Name = "AlteracoesCobrancaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFilialAlterada As Integer
Dim iTipoAlterado As Integer
Dim iNumeroAlterado As Integer
Dim iParcelaAlterada As Integer
Dim iSequencialAlterado As Integer
Private iDataEmissaoAlterada As Integer

Dim glNumIntParc As Long

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1
Private WithEvents objEventoSequencial As AdmEvento
Attribute objEventoSequencial.VB_VarHelpID = -1
Private WithEvents objEventoTipoDoc As AdmEvento
Attribute objEventoTipoDoc.VB_VarHelpID = -1

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjOcorrRemParcRec As ClassOcorrRemParcRec
Private gobjTituloReceber As New ClassTituloReceber
Private gobjParcelaReceber As New ClassParcelaReceber

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iSequencial As Integer

On Error GoTo Erro_BotaoProxNum_Click

    If glNumIntParc = 0 Then Error 59193
    
    'Gera o próximo Sequencial
    lErro = CF("OcorrRemParcRec_Automatico", glNumIntParc, iSequencial)
    If lErro <> SUCESSO Then Error 57542

    Call Limpa_Tela_Campos_AlteracoesCobranca
    
    'Mostra na tela o próximo Sequencial
    Sequencial.PromptInclude = False
    Sequencial.Text = iSequencial
    Sequencial.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57542
        
        Case 59193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142764)
    
    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Or _
    Len(Trim(Numero.Text)) = 0 Or Len(Trim(Parcela.Text)) = 0 Or Len(Trim(Sequencial.Text)) = 0 Then Error 28580

    If glNumIntParc = 0 Then Error 28582

    'Preenche objOcorrRemParcRec
    objOcorrRemParcRec.lNumIntParc = glNumIntParc
    objOcorrRemParcRec.iNumSeqOcorr = CInt(Sequencial.Text)
    objOcorrRemParcRec.iFilialEmpresa = giFilialEmpresa
    
    'Pesquisa a Ocorrência no BD
    lErro = CF("OcorrRemParcRec_Le", objOcorrRemParcRec)
    If lErro <> SUCESSO And lErro <> 28535 Then Error 28593

    'Se não encontrou a Ocorrência ==> Erro
    If lErro <> SUCESSO Then Error 28594

    'Se a Instrução já estiver sido enviada p/ o Banco --> Erro
    If objOcorrRemParcRec.lNumBordero <> 0 Then Error 28647
    
    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_OCORRENCIA")

    If vbMsgRes = vbYes Then

        'Chama OcorrRemParcRec_Exclui
        lErro = CF("OcorrRemParcRec_Exclui", objOcorrRemParcRec)
        If lErro <> SUCESSO Then Error 28595

        'Limpa a tela
        Call Limpa_Tela_AlteracoesCobranca

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 28580
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", Err)

        Case 28582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", Err)

        Case 28595, 28593

        Case 28594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORR_REM_COBR_NAO_CADASTRADA", Err, objOcorrRemParcRec.iNumSeqOcorr)

        Case 28647
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_ENVIADA_BANCO", Err, objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142765)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 28552

    'Limpa a Tela
    Call Limpa_Tela_AlteracoesCobranca
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 28552

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 142766)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 28551

    'Limpa a Tela
    Call Limpa_Tela_AlteracoesCobranca

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 28551

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142767)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_AlteracoesCobranca()

    'Limpa os Campos da Tela
    Call Limpa_Tela(Me)
    
    'Limpa os campos não limpos em Limpa_Tela
    Sequencial.PromptInclude = False
    Sequencial.Text = ""
    Sequencial.PromptInclude = True
    Filial.Clear
    Tipo.Text = ""
    Cobrador.Caption = ""
    Carteira.Caption = ""
    Ocorrencia.Text = ""
    LabelVctoAtual.Caption = ""
    LabelVctoOriginal.Caption = ""
    Instrucao1.Text = ""
    Instrucao2.Text = ""
    
    'Zera glNumIntParc
    glNumIntParc = 0

    iAlterado = 0

End Sub

Private Sub Limpa_Tela_Campos_AlteracoesCobranca()

    'Limpa os campos
    Ocorrencia.Text = ""
    Juros.Text = ""
    LabelVctoOriginal.Caption = ""
    NovoVcto.PromptInclude = False
    NovoVcto.Text = ""
    NovoVcto.PromptInclude = True
    Instrucao1.Text = ""
    DiasDeProtesto1.Text = ""
    Instrucao2.Text = ""
    DiasDeProtesto1.Text = ""
    
End Sub

Private Sub Cliente_Change()

   iClienteAlterado = 1
   iAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then Error 28538

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 28539

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
             Call CF("Filial_Seleciona", Filial, iCodFilial)

        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear

        End If
        
        'Se Cliente foi alterado zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 28538
            
        Case 28539

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142768)

    End Select

    Exit Sub

End Sub

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub DiasDeProtesto1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasDeProtesto1_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDeProtesto1_Validate

    'Verifica se algum Dia foi digitado
    If Len(Trim(DiasDeProtesto1.ClipText)) = 0 Then Exit Sub

    'Critica o Dia
    lErro = Inteiro_Critica(DiasDeProtesto1.Text)
    If lErro <> SUCESSO Then Error 28639

    Exit Sub

Erro_DiasDeProtesto1_Validate:

    Cancel = True


    Select Case Err

        Case 28639

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142769)

    End Select

    Exit Sub

End Sub

Private Sub DiasDeProtesto2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasDeProtesto2_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiasDeProtesto2_Validate

    'Verifica se algum Dia foi digitado
    If Len(Trim(DiasDeProtesto2.ClipText)) = 0 Then Exit Sub

    'Critica o Dia
    lErro = Inteiro_Critica(DiasDeProtesto2.Text)
    If lErro <> SUCESSO Then Error 28640

    Exit Sub

Erro_DiasDeProtesto2_Validate:

    Cancel = True


    Select Case Err

        Case 28640

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142770)

    End Select

    Exit Sub
End Sub

Private Sub NovoVcto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NovoVcto_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NovoVcto, iAlterado)

End Sub

Private Sub NovoVcto_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NovoVcto_Validate

    'Verifica se a data foi digitada
    If Len(Trim(NovoVcto.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(NovoVcto.Text)
    If lErro <> SUCESSO Then Error 28537

    Exit Sub

Erro_NovoVcto_Validate:

    Cancel = True


    Select Case Err

        Case 28537

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142771)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()

    iFilialAlterada = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO

    If Filial.ListIndex = -1 Then Exit Sub

    'Se a Filial foi alterada zera glNumIntParc
    glNumIntParc = 0
    
    Call Filial_Validate(bSGECancelDummy)

    Exit Sub

Erro_Filial_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142772)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    If iFilialAlterada Then

        'Verifica se a filial foi preenchida
        If Len(Trim(Filial.Text)) = 0 Then Exit Sub

        'Verifica se é uma filial selecionada
        If Filial.Text = Filial.List(Filial.ListIndex) Then
            Call Verifica_Alteracao
            Exit Sub
        End If
        
        'Tenta selecionar na combo
        lErro = Combo_Seleciona(Filial, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 28540

        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then

            'Verifica se o Cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then Error 28541

            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo

            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then Error 28542

            If lErro = 17660 Then Error 28543

            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        End If

        'Não encontrou a STRING
        If lErro = 6731 Then Error 28544
        
        'Se Filial foi alterado zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iFilialAlterada = 0

    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 28540, 28542

        Case 28541
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 28543
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 28544
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142773)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'Carrega os Tipos de Documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 28520

    'Carrega Ocorrências
    lErro = Carrega_Ocorrencia()
    If lErro <> SUCESSO Then Error 28521

    'Carrega Instruções
    lErro = Carrega_Instrucao()
    If lErro <> SUCESSO Then Error 28522

    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoParcela = New AdmEvento
    Set objEventoSequencial = New AdmEvento
    Set objEventoTipoDoc = New AdmEvento
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28520, 28521, 28522

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142774)

    End Select
    
     iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Lê os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then Error 28523

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = Err

    Select Case Err

        Case 28523

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142775)

    End Select

    Exit Function

End Function

Private Function Carrega_Ocorrencia() As Long

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Ocorrencia

    'Lê o código e a descrição de todas as Ocorrências
    lErro = CF("Cod_Nomes_Le", "TiposDeOcorRemCobr", "Codigo", "Descricao", STRING_TIPOINSTRCOBR_DESCRICAO, colCodigoNome)
    If lErro <> SUCESSO Then Error 28536
    
    For Each objCodigoNome In colCodigoNome

        If objCodigoNome.iCodigo <> COBRANCA_OCORR_INC_TITULO Then
        
            'Adiciona novo ítem na List da ComboBox Ocorrencia
            Ocorrencia.AddItem CInt(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            Ocorrencia.ItemData(Ocorrencia.NewIndex) = objCodigoNome.iCodigo

        End If
        
    Next

    Carrega_Ocorrencia = SUCESSO

    Exit Function

Erro_Carrega_Ocorrencia:

    Carrega_Ocorrencia = Err

    Select Case Err

        Case 28536

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142776)

    End Select

    Exit Function

End Function

Private Function Carrega_Instrucao() As Long
'Carrega a ComboBox de Instrução

Dim lErro As Long
Dim colTiposInstrCobranca As New Collection
Dim objTiposInstrCobranca As New ClassTipoInstrCobr

On Error GoTo Erro_Carrega_Instrucao

    'Lê o código e a descrição de todas as Ocorrências
    lErro = CF("TiposInstrCobranca_Le_Todos", colTiposInstrCobranca)
    If lErro <> SUCESSO Then Error 28561

    For Each objTiposInstrCobranca In colTiposInstrCobranca

        'Adiciona novo ítem na List da ComboBox Instrução
        Instrucao1.AddItem CInt(objTiposInstrCobranca.iCodigo) & SEPARADOR & objTiposInstrCobranca.sDescricao
        Instrucao1.ItemData(Instrucao1.NewIndex) = objTiposInstrCobranca.iRequerDias
        Instrucao2.AddItem CInt(objTiposInstrCobranca.iCodigo) & SEPARADOR & objTiposInstrCobranca.sDescricao
        Instrucao2.ItemData(Instrucao2.NewIndex) = objTiposInstrCobranca.iRequerDias
    
    Next

    'Desabilita os campos DiasDeProtesto
    DiasDeProtesto1.Enabled = False
    DiasDeProtesto2.Enabled = False

    'Desabilita Instrução2
    Instrucao2.Enabled = False

    Carrega_Instrucao = SUCESSO

    Exit Function

Erro_Carrega_Instrucao:

    Carrega_Instrucao = Err

    Select Case Err

        Case 28561

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142777)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoParcela = Nothing
    Set objEventoSequencial = Nothing
    Set objEventoTipoDoc = Nothing

    Set gobjContabAutomatica = Nothing
    Set gobjOcorrRemParcRec = Nothing
    Set gobjTituloReceber = Nothing
    Set gobjParcelaReceber = Nothing

End Sub

Private Sub Instrucao1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Instrucao1_Click()

Dim lErro As Long

On Error GoTo Erro_Instrucao1_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se Instrução1 foi informada
    If Instrucao1.ListIndex = -1 Then
    
        DiasDeProtesto1.Text = ""
        DiasDeProtesto1.Enabled = False
        Instrucao2.ListIndex = -1
        Instrucao2.Enabled = False
        
    Else
    
        Instrucao2.Enabled = True
    
        'Verifica se requer dias
        If Instrucao1.ItemData(Instrucao1.ListIndex) = INSTR_COBR_REQUER_DIAS Then
            DiasDeProtesto1.Enabled = True
        Else
            DiasDeProtesto1.Text = ""
            DiasDeProtesto1.Enabled = False
        End If
        
    End If
    
    Exit Sub

Erro_Instrucao1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142778)

    End Select

    Exit Sub

End Sub

Private Sub Instrucao1_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoInstrCobr As New ClassTipoInstrCobr

On Error GoTo Erro_Instrucao1_Validate

    'Verifica se Instrucao1 foi preenchida
    If Len(Trim(Instrucao1.Text)) = 0 Then
        DiasDeProtesto1.Text = ""
        Exit Sub
    End If
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
    If Instrucao1.Text = Instrucao1.List(Instrucao1.ListIndex) Then Exit Sub
    
    If IsNumeric(Instrucao1.Text) Then
        objTipoInstrCobr.iCodigo = CInt(Instrucao1.Text)
        'Lê o Tipo de Instrução de Cobrança
        lErro = CF("TipoInstrCobranca_Le", objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 16549 Then Error 43542
    
        'Se não achou o Tipo de Instrução de Cobrança --> Erro
        If lErro <> SUCESSO Then Error 43543
        
        'Mostra na tela a Inscrição
        Instrucao1.Text = objTipoInstrCobr.iCodigo & SEPARADOR & objTipoInstrCobr.sDescricao
        
    Else
        
        objTipoInstrCobr.sDescricao = Instrucao1.Text
        'Lê o Tipo de Instrução de Cobrança
        lErro = CF("TipoInstrCobranca_Le_Descricao", objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 43163 Then Error 43544
    
        'Se não achou o Tipo de Instrução de Cobrança --> Erro
        If lErro <> SUCESSO Then Error 43545
        
        'Mostra na tela a Inscrição
        Instrucao1.Text = objTipoInstrCobr.iCodigo & SEPARADOR & objTipoInstrCobr.sDescricao
        
    End If
    
    'Seleciona na Combo
    lErro = Combo_Item_Igual(Instrucao1)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 43546

    'Se não encontrar -> erro
    If lErro = 12253 Then Error 43547
                
    Exit Sub

Erro_Instrucao1_Validate:

    Cancel = True


    Select Case Err

        Case 43542, 43544, 43546

        Case 43543
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_CADASTRADA", Err, objTipoInstrCobr.iCodigo)

        Case 43545
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_CADASTRADA1", Err, Instrucao1.Text)

        Case 43547
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_SELECIONADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142779)

    End Select

    Exit Sub

End Sub

Private Sub Instrucao2_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Instrucao2_Click()

Dim lErro As Long

On Error GoTo Erro_Instrucao2_Click

    iAlterado = REGISTRO_ALTERADO

    'Verifica se Instrução2 foi informada
    If Instrucao2.ListIndex = -1 Then
        
        DiasDeProtesto2.Text = ""
        DiasDeProtesto2.Enabled = False

    Else
    
        'Verifica se requer dias
        If Instrucao2.ItemData(Instrucao2.ListIndex) = INSTR_COBR_REQUER_DIAS Then
            DiasDeProtesto2.Enabled = True
        Else
            DiasDeProtesto2.Text = ""
            DiasDeProtesto2.Enabled = False
        End If

    End If
    
    Exit Sub

Erro_Instrucao2_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142780)

    End Select

    Exit Sub

End Sub

Private Sub Instrucao2_Validate(Cancel As Boolean)
'chamar combo_seleciona ou algo parecido e remover codigo inutil
'obs.: Não posso usar Combo_Seleciona porque no Item data não estou guardando
'o código e sim se a Instrução requer dias

Dim lErro As Long
Dim objTipoInstrCobr As New ClassTipoInstrCobr

On Error GoTo Erro_Instrucao2_Validate

    'Verifica se Instrucao2 foi preenchida
    If Len(Trim(Instrucao2.Text)) = 0 Then
        DiasDeProtesto2.Text = ""
        Exit Sub
    End If
    
    'Verifica se está preenchida com o ítem selecionado na ComboBox TabelaPreco
    If Instrucao2.Text = Instrucao2.List(Instrucao2.ListIndex) Then Exit Sub

    If IsNumeric(Instrucao2.Text) Then
    
        objTipoInstrCobr.iCodigo = CInt(Instrucao2.Text)
        'Lê o Tipo de Instrução de Cobrança
        lErro = CF("TipoInstrCobranca_Le", objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 16549 Then Error 43548
    
        'Se não achou o Tipo de Instrução de Cobrança --> Erro
        If lErro <> SUCESSO Then Error 43549
        
        'Mostra na tela a Inscrição
        Instrucao2.Text = objTipoInstrCobr.iCodigo & SEPARADOR & objTipoInstrCobr.sDescricao
        
    Else
        objTipoInstrCobr.sDescricao = Instrucao2.Text
        'Lê o Tipo de Instrução de Cobrança
        lErro = CF("TipoInstrCobranca_Le_Descricao", objTipoInstrCobr)
        If lErro <> SUCESSO And lErro <> 43163 Then Error 43550
    
        'Se não achou o Tipo de Instrução de Cobrança --> erro
        If lErro <> SUCESSO Then Error 43551
        
        'Mostra na tela a Inscrição
        Instrucao2.Text = objTipoInstrCobr.iCodigo & SEPARADOR & objTipoInstrCobr.sDescricao
        
    End If
    
    'Seleciona na Combo
    lErro = Combo_Item_Igual(Instrucao2)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 43552

    'Se não encontrar -> Erro
    If lErro <> SUCESSO Then Error 43553
    
    Exit Sub

Erro_Instrucao2_Validate:

    Cancel = True


    Select Case Err

        Case 43548, 43550, 43552

        Case 43549
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_CADASTRADA", Err, objTipoInstrCobr.iCodigo)

        Case 43551
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_CADASTRADA1", Err, Instrucao2.Text)

        Case 43553
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_NAO_SELECIONADA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142781)

    End Select

    Exit Sub

End Sub

Private Sub Juros_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Juros_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Juros_Validate

    'Verifica se algum juros foi digitado
    If Len(Trim(Juros.ClipText)) = 0 Then Exit Sub

    'Critica o Juros
    lErro = Valor_NaoNegativo_Critica(Juros.Text)
    If lErro <> SUCESSO Then Error 28550

    'Põe o Juros formatado na tela
    Juros.Text = Format(Juros.Text, "Fixed")

    Exit Sub

Erro_Juros_Validate:

    Cancel = True


    Select Case Err

        Case 28550

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142782)

    End Select

    Exit Sub

End Sub

Private Sub LabelParcela_Click()
'Lista as parcelas do titulo selecionado

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objParcelaReceber As ClassParcelaReceber
Dim colSelecao As New Collection

On Error GoTo Erro_LabelParcela_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then Error 43148
    If Len(Trim(Filial.Text)) = 0 Then Error 43149
    If Len(Trim(Tipo.Text)) = 0 Then Error 43150
    If Len(Trim(Numero.ClipText)) = 0 Then Error 43151
    
    objCliente.sNomeReduzido = Cliente.Text
    'Lê o Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 43152
    
    'Se não achou o Cliente --> erro
    If lErro <> SUCESSO Then Error 43153
    
    colSelecao.Add objCliente.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)
    colSelecao.Add SCodigo_Extrai(Tipo.Text)
    colSelecao.Add StrParaLong(Numero.Text)
    
    'Chama a tela
    Call Chama_Tela("ParcelasRecLista", colSelecao, objParcelaReceber, objEventoParcela)
    
    Exit Sub
    
Erro_LabelParcela_Click:

    Select Case Err
    
        Case 43148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
    
        Case 43149
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 43150
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)
            
        Case 43151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)
        
        Case 43152
    
        Case 43153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142783)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelSequencial_Click()

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim colSelecao As New Collection

On Error GoTo Erro_LabelSequencial_Click

    If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Or _
    Len(Trim(Numero.Text)) = 0 Or Len(Trim(Parcela.Text)) = 0 Then Error 43562
    
    If glNumIntParc = 0 Then Error 43160

    objOcorrRemParcRec.lNumIntParc = glNumIntParc
    
    colSelecao.Add objOcorrRemParcRec.lNumIntParc
        
    'Chama a tela
    Call Chama_Tela("InstrRemParcRecLista", colSelecao, objOcorrRemParcRec, objEventoSequencial)
    
    Exit Sub
    
Erro_LabelSequencial_Click:

    Select Case Err
    
        Case 43160
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", Err)
    
        Case 43562
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142784)
            
    End Select
    
    Exit Sub

End Sub

Private Sub LabelTipo_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)
    
    'Chama a tela TipoDocTituloRecLista
    Call Chama_Tela("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDoc)

End Sub

Private Sub Numero_Change()

    iNumeroAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Numero_GotFocus()
Dim iNumAux As Integer
    
    iNumAux = iNumeroAlterado
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
    iNumeroAlterado = iNumAux
    
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Número está preenchido
    If Len(Trim(Numero.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(Numero.ClipText)
    If lErro <> SUCESSO Then Error 28547

    If iNumeroAlterado Then
        
        'Se Número foi alterado zera glNUmIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iNumeroAlterado = 0

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True


    Select Case Err

        Case 28547

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142785)

    End Select

    Exit Sub

End Sub

Private Sub NumeroLabel_Click()

Dim objTituloReceber As New ClassTituloReceber
Dim objCliente As New ClassCliente
Dim colSelecao As New Collection
Dim lErro As Long
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_NumeroLabel_Click

    If Len(Trim(Cliente.Text)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 87206
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 87207

    End If
    
    'Guarda o código no objTituloReceber
    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)

    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If

    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If

    If Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND SiglaDocumento = ?"
        Else
            iPreenchido = 1
            sSelecao = "SiglaDocumento = ?"
        End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)
    End If

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 87206

        Case 87207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142786)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1
    
    lErro = Traz_TitReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then gError 87215
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 87215
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142787)

    End Select

    Exit Sub

End Sub

Private Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long, bCancela As Boolean
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objParcelaReceber = obj1

    If Not (objParcelaReceber Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objParcelaReceber.iNumParcela)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bCancela)
    End If

    Me.Show

    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142788)

    End Select

    Exit Sub

End Sub

Private Sub objEventoSequencial_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcorrRemParcRec As ClassOcorrRemParcRec
Dim objTitRec As New ClassTituloReceber
Dim objParcRec As New ClassParcelaReceber

On Error GoTo Erro_objEventoSequencial_evSelecao

    Set objOcorrRemParcRec = obj1

    'Limpa a Tela
    Call Limpa_Tela_AlteracoesCobranca

    objOcorrRemParcRec.iFilialEmpresa = giFilialEmpresa
    
    'Pesquisa a Ocorrência no BD
    lErro = CF("OcorrRemParcRec_Le", objOcorrRemParcRec)
    If lErro <> SUCESSO And lErro <> 28535 Then gError ERRO_SEM_MENSAGEM
   
    objParcRec.lNumIntDoc = objOcorrRemParcRec.lNumIntParc
    
    lErro = CF("ParcelaReceber_Le", objParcRec)
    If lErro <> SUCESSO And lErro = 19147 Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> SUCESSO Then
    
        lErro = CF("ParcelaReceber_Baixada_Le", objParcRec)
        If lErro <> SUCESSO And lErro = 58559 Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    objTitRec.lNumIntDoc = objParcRec.lNumIntTitulo
    
    lErro = Traz_TitReceber_Tela(objTitRec)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Parcela.PromptInclude = False
    Parcela.Text = objParcRec.iNumParcela
    Parcela.PromptInclude = True
    Call Parcela_Validate(bSGECancelDummy)
    
    Sequencial.PromptInclude = False
    Sequencial.Text = CInt(objOcorrRemParcRec.iNumSeqOcorr)
    Sequencial.PromptInclude = True
    
    lErro = Traz_Ocorrencia_Tela(objOcorrRemParcRec)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoSequencial_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142789)
     
     End Select
     
     Exit Sub

End Sub

Private Sub objEventoTipoDoc_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142790)
     
     End Select
     
     Exit Sub

End Sub

Private Sub Ocorrencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ocorrencia_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Ocorrencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Ocorrencia_Validate

    'Verifica se Ocorrência foi preenchida
    If Len(Trim(Ocorrencia.Text)) = 0 Then Exit Sub

    'Verifica se Ocorrência foi selecionada
    If Ocorrencia.Text = Ocorrencia.List(Ocorrencia.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Ocorrencia, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 28568

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then Error 28569

    'Se não encontrou a STRING
    If lErro = 6731 Then Error 28570

    Exit Sub

Erro_Ocorrencia_Validate:

    Cancel = True


    Select Case Err

        Case 28568

        Case 28569
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_CADASTRADA", Err, iCodigo)

        Case 28570
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_NAO_CADASTRADA1", Err, Ocorrencia.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142791)

    End Select

    Exit Sub

End Sub

Private Sub Parcela_Change()

    iParcelaAlterada = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Parcela_GotFocus()
Dim iParcelaAux As Integer
    
    iParcelaAux = iParcelaAlterada
    Call MaskEdBox_TrataGotFocus(Parcela, iAlterado)
    iParcelaAlterada = iParcelaAux

End Sub

Private Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    'Verifica se está preenchido
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Valor_Positivo_Critica(Parcela.ClipText)
    If lErro <> SUCESSO Then Error 28620

    If iParcelaAlterada Then
    
        'Se Parcela foi alterada zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iParcelaAlterada = 0
    
    End If

    Exit Sub

Erro_Parcela_Validate:

    Select Case Err

        Case 28620
            Cancel = True

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142792)

    End Select

    Exit Sub

End Sub

Private Sub Sequencial_Change()

    iSequencialAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Sequencial_GotFocus()

Dim iSeqAux As Integer
    
    iSeqAux = iSequencialAlterado
    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)
    iSequencialAlterado = iSeqAux
    
End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec

On Error GoTo Erro_Sequencial_Validate
        
    If Len(Trim(Sequencial.Text)) > 0 Then
        
        lErro = Valor_Positivo_Critica(Sequencial.Text)
        If lErro <> SUCESSO Then Error 57998
    
    End If

    If glNumIntParc <> 0 And Len(Trim(Sequencial.Text)) Then
            
        'Preenche objOcorrRemParcRec
        objOcorrRemParcRec.lNumIntParc = glNumIntParc
        objOcorrRemParcRec.iNumSeqOcorr = CInt(Sequencial.Text)
        objOcorrRemParcRec.iFilialEmpresa = giFilialEmpresa
        
        'Lê a Ocorrência
        lErro = CF("OcorrRemParcRec_Le", objOcorrRemParcRec)
        If lErro <> SUCESSO And lErro <> 28535 Then Error 28576

        'Não encontrou a Ocorrência
        If lErro <> SUCESSO Then
            Call Limpa_Tela_Campos_AlteracoesCobranca
            Exit Sub
        End If

        'Verifica se Ocorrência não faz parte de um bordero de cobrança
        If objOcorrRemParcRec.lNumBordero <> 0 Then Error 28630
        
        'Se a Ocorrência estiver cadastrada mostra os dados na tela
        Call Traz_Ocorrencia_Tela(objOcorrRemParcRec)

    End If
        
    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True


    Select Case Err

        Case 28575, 57998

        Case 28576

        Case 28630
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_OCORRENCIA_BORDERO_COBRANCA", Err, objOcorrRemParcRec.lNumIntParc, objOcorrRemParcRec.iNumSeqOcorr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142793)

    End Select

    Exit Sub

End Sub

Private Function Traz_Ocorrencia_Tela(objOcorrRemParcRec As ClassOcorrRemParcRec) As Long

    If objOcorrRemParcRec.iCodOcorrencia <> 0 Then
        Ocorrencia.Text = CStr(objOcorrRemParcRec.iCodOcorrencia)
        Call Ocorrencia_Validate(bSGECancelDummy)
    Else
        Ocorrencia.Text = ""
    End If
    
    If objOcorrRemParcRec.dJuros <> 0 Then
        Juros.Text = Format(objOcorrRemParcRec.dJuros, "Fixed")
    Else
        Juros.Text = ""
    End If
    
    If objOcorrRemParcRec.iCodOcorrencia <> COBRANCA_OCORR_ALT_VCTO Then
        LabelVctoOriginal.Caption = ""
    Else
        LabelVctoOriginal.Caption = Format(objOcorrRemParcRec.dtData, "dd/mm/yyyy")
    End If
    
    If objOcorrRemParcRec.iInstrucao1 <> 0 Then
        Instrucao1.Text = CStr(objOcorrRemParcRec.iInstrucao1)
        Call Instrucao1_Validate(bSGECancelDummy)
    Else
        Instrucao1.Text = ""
    End If
    If objOcorrRemParcRec.iDiasDeProtesto1 <> 0 Then
        DiasDeProtesto1.Text = CStr(objOcorrRemParcRec.iDiasDeProtesto1)
    Else
        DiasDeProtesto1.Text = ""
    End If
    If objOcorrRemParcRec.iInstrucao2 <> 0 Then
        Instrucao2.Text = CStr(objOcorrRemParcRec.iInstrucao2)
        Call Instrucao2_Validate(bSGECancelDummy)
    Else
        Instrucao2.Text = ""
    End If
    If objOcorrRemParcRec.iDiasDeProtesto2 <> 0 Then
        DiasDeProtesto2.Text = CStr(objOcorrRemParcRec.iDiasDeProtesto2)
    Else
        DiasDeProtesto2.Text = ""
    End If
    If objOcorrRemParcRec.dtNovaDataVcto <> DATA_NULA And objOcorrRemParcRec.dtNovaDataVcto <> 0 Then
        NovoVcto.PromptInclude = False
        NovoVcto.Text = objOcorrRemParcRec.dtNovaDataVcto
        NovoVcto.PromptInclude = True
    Else
        NovoVcto.PromptInclude = False
        NovoVcto.Text = ""
        NovoVcto.PromptInclude = True
    End If
    
    iAlterado = 0

End Function

Private Sub Tipo_Change()

    iTipoAlterado = 1
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Click()

Dim lErro As Long

On Error GoTo Erro_Tipo_Click

    iAlterado = REGISTRO_ALTERADO
    
    If Tipo.ListIndex = -1 Then Exit Sub

    'Se o Tipo foi alterado zera glNumIntParc
    glNumIntParc = 0
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_Tipo_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142794)

    End Select

    Exit Sub

End Sub

Private Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then
        Call Verifica_Alteracao
        Exit Sub
    End If
    
    'Tenta localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 28545

    'Se não encontrar -> Erro
    If lErro = 60483 Then Error 28546

    If iTipoAlterado Then
        
        'Se Tipo foi alterado zera glNumIntParc
        glNumIntParc = 0

        Call Verifica_Alteracao

        iTipoAlterado = 0

    End If

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 28545

        Case 28546
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142795)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDown1_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(NovoVcto, DIMINUI_DATA)
    If lErro Then Error 28548

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case Err

        Case 28548

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142796)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown1_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(NovoVcto, AUMENTA_DATA)
    If lErro Then Error 28549

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case Err

        Case 28549

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142797)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Valida os dados para gravação

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(Tipo.Text)) = 0 Or _
    Len(Trim(Numero.Text)) = 0 Or Len(Trim(Parcela.Text)) = 0 Then Error 28648
    
    'Verifica se glNumIntParc global ao modulo é <> 0
    If glNumIntParc = 0 Then Error 28581
    
    'Verifica preenchimento de Sequencial
    If Len(Trim(Sequencial.Text)) = 0 Then Error 28583
    
    'Verifica preenchimento de Ocorrência
    If Len(Trim(Ocorrencia.Text)) = 0 Then Error 28665

    'Verifica preenchimento de Instrução1
    If Len(Trim(Instrucao1.Text)) > 0 Then
    
        'Verifica se Instrução1 Requer Dias
        If Instrucao1.ItemData(Instrucao1.ListIndex) = INSTR_COBR_REQUER_DIAS And Len(Trim(DiasDeProtesto1.Text)) = 0 Then Error 28649
        
    End If
    
    'Verifica preenchimento de Instrução2
    If Len(Trim(Instrucao2.Text)) > 0 Then
        'Verifica se Instrução2 Requer Dias
        If Instrucao2.ItemData(Instrucao2.ListIndex) = INSTR_COBR_REQUER_DIAS And Len(Trim(DiasDeProtesto2.Text)) = 0 Then Error 28650
    End If
    
    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objOcorrRemParcRec)
    If lErro <> SUCESSO Then Error 28559

    'se está gravando uma alteracao de vcto o campo de novo vcto tem que estar preenchido
    If objOcorrRemParcRec.iCodOcorrencia = COBRANCA_OCORR_ALT_VCTO And objOcorrRemParcRec.dtNovaDataVcto = DATA_NULA Then Error 59203
    
    'se nao está gravando uma alteracao de vcto o campo de novo vcto tem que estar vazio
    If objOcorrRemParcRec.iCodOcorrencia <> COBRANCA_OCORR_ALT_VCTO And objOcorrRemParcRec.dtNovaDataVcto <> DATA_NULA Then Error 59204
     
    'Chama OcorrRemParcRec_Grava
    objOcorrRemParcRec.objTelaAtualizacao = Me
    lErro = CF("OcorrRemParcRec_Grava", objOcorrRemParcRec)
    If lErro <> SUCESSO Then Error 28560

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 28559, 28560

        Case 28581
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_CADASTRADA", Err)

        Case 28583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", Err)

        Case 28648
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_INFORMADA1", Err)

        Case 28649, 28650
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIAS_PROTESTO_NAO_PREENCHIDO", Err)

        Case 28665
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORRENCIA_COBRANCA_NAO_PREENCHIDA", Err)

        Case 28666
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INSTRUCAO_COBRANCA_NAO_PREENCHIDA", Err)

        Case 59203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORR_ALT_DATA_VCTO_SEM_DATA", Err)
        
        Case 59204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_OCORR_VCTO_ALT_ERRADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142798)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objOcorrRemParcRec As ClassOcorrRemParcRec) As Long
'Recolhe os dados da tela para a memória

Dim lErro As Long
Dim objCobrador As New ClassCobrador

On Error GoTo Erro_Move_Tela_Memoria

    objCobrador.sNomeReduzido = Cobrador.Caption
    'Lê o Cobrador
    lErro = CF("Cobrador_Le_NomeReduzido", objCobrador)
    If lErro <> SUCESSO And lErro <> 43557 Then Error 43560
    
    'Se não encontrou o Cobrador --> Erro
    If lErro <> SUCESSO Then Error 43561
    
    'Preenche objOcorrRemParcRec
    objOcorrRemParcRec.lNumIntParc = glNumIntParc
    If Len(Trim(Sequencial.ClipText)) > 0 Then objOcorrRemParcRec.iNumSeqOcorr = CInt(Sequencial.ClipText)
    objOcorrRemParcRec.iFilialEmpresa = giFilialEmpresa
    objOcorrRemParcRec.iCobrador = objCobrador.iCodigo
    If Len(Trim(Ocorrencia.Text)) > 0 Then objOcorrRemParcRec.iCodOcorrencia = Codigo_Extrai(Ocorrencia.Text)
    objOcorrRemParcRec.dtDataRegistro = gdtDataHoje
    
    'se nao está gravando uma alteracao de vcto
    If objOcorrRemParcRec.iCodOcorrencia <> COBRANCA_OCORR_ALT_VCTO Then
    
        objOcorrRemParcRec.dtData = gdtDataAtual
        
    Else
    
        If Len(Trim(LabelVctoOriginal.Caption)) <> 0 Then
            objOcorrRemParcRec.dtData = CDate(LabelVctoOriginal.Caption)
        Else
            If Len(Trim(LabelVctoAtual.Caption)) <> 0 Then
                objOcorrRemParcRec.dtData = CDate(LabelVctoAtual.Caption)
            Else
                objOcorrRemParcRec.dtData = DATA_NULA
            End If
        End If
        
    End If
    
    If Len(Trim(NovoVcto.ClipText)) > 0 Then
        objOcorrRemParcRec.dtNovaDataVcto = CDate(NovoVcto.Text)
    Else
        objOcorrRemParcRec.dtNovaDataVcto = DATA_NULA
    End If
    If Len(Trim(Juros.Text)) > 0 Then objOcorrRemParcRec.dJuros = CDbl(Juros.Text)
    If Len(Trim(Instrucao1.Text)) > 0 Then objOcorrRemParcRec.iInstrucao1 = Codigo_Extrai(Instrucao1.Text)
    If Len(Trim(DiasDeProtesto1.Text)) > 0 Then objOcorrRemParcRec.iDiasDeProtesto1 = CInt(DiasDeProtesto1.Text)
    If Len(Trim(Instrucao2.Text)) > 0 Then objOcorrRemParcRec.iInstrucao2 = Codigo_Extrai(Instrucao2.Text)
    If Len(Trim(DiasDeProtesto2.Text)) > 0 Then objOcorrRemParcRec.iDiasDeProtesto2 = CInt(DiasDeProtesto2.Text)

    objOcorrRemParcRec.dValorCobrado = 0

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 43560
        
        Case 43561
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO1", Err, objCobrador.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142799)

    End Select

    Exit Function

End Function

Private Function Verifica_Alteracao() As Long
'tenta obter o NumInt da parcela e trazer seus dados para a tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objCobrador As New ClassCobrador
Dim objCarteiraCobranca As New ClassCarteiraCobranca
Dim vbMsgRet As VbMsgBoxResult

On Error GoTo Erro_Verifica_Alteracao

    'Verifica preenchimento de Cliente
    If Len(Trim(Cliente.Text)) = 0 Then Exit Function

    'Verifica preenchimento de Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Function

    'Verifica preenchimento do Tipo
    If Len(Trim(Tipo.Text)) = 0 Then Exit Function

    'Verifica preenchimento de NumTítulo
    If Len(Trim(Numero.Text)) = 0 Then Exit Function

    'Verifica preenchimento da Parcela
    If Len(Trim(Parcela.ClipText)) = 0 Then Exit Function

    objCliente.sNomeReduzido = Cliente.Text

    'Lê Cliente
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 28611

    'Se não encontrou o Cliente --> Erro
    If lErro <> SUCESSO Then Error 28612

   'Preenche objTituloReceber
    gobjTituloReceber.iFilialEmpresa = giFilialEmpresa
    gobjTituloReceber.lCliente = objCliente.lCodigo
    gobjTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    gobjTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    gobjTituloReceber.lNumTitulo = CLng(Numero.Text)
    gobjTituloReceber.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Pesquisa no BD o Título Receber
    lErro = CF("TituloReceber_Le_SemNumIntDoc", gobjTituloReceber)
    If lErro <> SUCESSO And lErro <> 28574 Then Error 28613

    'Se não encontrou o Título --> Erro
    If lErro <> SUCESSO Then Error 28614

    'Preenche objParcelaReceber
    gobjParcelaReceber.lNumIntTitulo = gobjTituloReceber.lNumIntDoc
    gobjParcelaReceber.iNumParcela = CInt(Parcela.Text)

    'Pesquisa no BD a Parcela
    lErro = CF("ParcelaReceber_Le_SemNumIntDoc", gobjParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28590 Then Error 28615

    'Se não encontrou a Parcela --> Erro
    If lErro <> SUCESSO Then Error 28616

    'Verifica se é uma Parcela Baixada
    lErro = CF("ParcelaReceberBaixada_Le_SemNumIntDoc", gobjParcelaReceber)
    If lErro <> SUCESSO And lErro <> 28567 Then Error 28645

    'Se encontrou a Parcela Receber Baixada --> Erro
    If lErro = SUCESSO Then Error 28646
    
    If Len(Trim(gobjParcelaReceber.sNumTitCobrador)) = 0 Then
        vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_PARCREC_SEM_CONF_ENTRADA")
        If vbMsgRet = vbNo Then Error 28611
    End If
    
    'Preenche objCobrador
    objCobrador.iCodigo = gobjParcelaReceber.iCobrador

    'Lê Cobrador
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 28617
    
    'Se não encontrou o Cobrador --> Erro
    If lErro <> SUCESSO Then Error 28578

    If objCobrador.iCobrancaEletronica <> COBRANCA_ELETRONICA Then Error 28618
    
    'Preenche Cobrador da tela
    If objCobrador.sNomeReduzido <> "" Then
        Cobrador.Caption = objCobrador.sNomeReduzido
    Else
        Cobrador.Caption = ""
    End If

    'Preenche objCarteiraCobranca
    objCarteiraCobranca.iCodigo = gobjParcelaReceber.iCarteiraCobranca

    'Lê CarteiraCobranca
    lErro = CF("CarteiraDeCobranca_Le", objCarteiraCobranca)
    If lErro <> SUCESSO And lErro <> 23413 Then Error 28619
    
    'Se não encontrou a Carteira Cobrança --> Erro
    If lErro <> SUCESSO Then Error 28579

    'Preenche Carteira da tela
    If objCarteiraCobranca.sDescricao <> "" Then
        Carteira.Caption = objCarteiraCobranca.sDescricao
    Else
        Carteira.Caption = ""
    End If

    LabelVctoAtual.Caption = Format(gobjParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
    
    If Len(Trim(Sequencial.Text)) = 0 Then

        glNumIntParc = gobjParcelaReceber.lNumIntDoc

        iSequencialAlterado = 0

    Else
        If gobjParcelaReceber.lNumIntDoc <> 0 Then
            glNumIntParc = gobjParcelaReceber.lNumIntDoc
            Call Sequencial_Validate(bSGECancelDummy)
        End If
    End If

    Verifica_Alteracao = SUCESSO

    Exit Function

Erro_Verifica_Alteracao:

    Verifica_Alteracao = Err

    Select Case Err

        Case 28578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_CADASTRADO", Err, objCobrador.iCodigo)

        Case 28579
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err, objCarteiraCobranca.iCodigo)

        Case 28611, 28613, 28615, 28617, 28619, 28645

        Case 28612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)

        Case 28614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO2", Err, gobjTituloReceber.iFilialEmpresa, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, gobjTituloReceber.sSiglaDocumento, gobjTituloReceber.lNumTitulo)

        Case 28616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_NAO_CADASTRADA", Err, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case 28618
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_COBR_ELETRONICA", Err)

        Case 28646
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAREC_NUMINT_BAIXADA", Err, gobjParcelaReceber.lNumIntTitulo, gobjParcelaReceber.iNumParcela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142800)

    End Select

    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador, sContaTela As String, sContaCobr As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case "Valor_Cobrado"

            objMnemonicoValor.colValor.Add gobjOcorrRemParcRec.dSaldo

        Case "Cliente_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.lCliente
        
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.iFilial
        
        Case "Titulo_Numero"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.lNumTitulo
        
        Case "Titulo_Filial"
        
            objMnemonicoValor.colValor.Add gobjTituloReceber.iFilialEmpresa
        
        Case "Parcela_Numero"
        
            objMnemonicoValor.colValor.Add gobjParcelaReceber.iNumParcela
                
        Case "CartCobrOrigem_Conta"
            
            objCarteiraCobrador.iCobrador = gobjParcelaReceber.iCobrador
            objCarteiraCobrador.iCodCarteiraCobranca = gobjParcelaReceber.iCarteiraCobranca
            
            lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
            If lErro <> SUCESSO And lErro <> 23551 Then Error 56811
            If lErro <> SUCESSO Then Error 56812
            
            sContaTela = ""
            sContaCobr = IIf(objCarteiraCobrador.iCodCarteiraCobranca <> CARTEIRA_DESCONTADA, objCarteiraCobrador.sContaContabil, objCarteiraCobrador.sContaDuplDescontadas)
            
            If sContaCobr <> "" Then
            
                lErro = Mascara_RetornaContaTela(sContaCobr, sContaTela)
                If lErro <> SUCESSO Then Error 56813
            
                sContaCobr = sContaTela
                
            End If

            objMnemonicoValor.colValor.Add sContaTela
        
        Case Else

            Error 56542

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56811, 56813
                
        Case 56542
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 56812
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142801)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de parcela e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjOcorrRemParcRec = vParams(0)
        
    'obtem numero de doc
    lErro = objContabAutomatica.Obter_Doc(lDoc, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 32241

    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, "InstrCobrEletr", gobjOcorrRemParcRec.lNumIntDoc, gobjTituloReceber.lCliente, gobjTituloReceber.iFilial, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, giFilialEmpresa)
    If lErro <> SUCESSO Then Error 32242
            
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32241, 32242
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 142802)
     
    End Select
     
    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_INSTRUCOES_PARA_TITULO_COBRANCA_ELETRONICA
    Set Form_Load_Ocx = Me
    Caption = "Instruções para Títulos em Cobrança Eletrônica"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "AlteracoesCobranca"
    
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
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Parcela Then
            Call LabelParcela_Click
        ElseIf Me.ActiveControl Is Sequencial Then
            Call LabelSequencial_Click
        End If
    
    End If
    
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub LabelVctoOriginal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVctoOriginal, Source, X, Y)
End Sub

Private Sub LabelVctoOriginal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVctoOriginal, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelVctoAtual_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVctoAtual, Source, X, Y)
End Sub

Private Sub LabelVctoAtual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVctoAtual, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label15_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label15, Source, X, Y)
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label15, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Carteira_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Carteira, Source, X, Y)
End Sub

Private Sub Carteira_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Carteira, Button, Shift, X, Y)
End Sub

Private Sub Cobrador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Cobrador, Source, X, Y)
End Sub

Private Sub Cobrador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Cobrador, Button, Shift, X, Y)
End Sub

Private Sub label1007_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(label1007, Source, X, Y)
End Sub

Private Sub label1007_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(label1007, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub LabelSequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSequencial, Source, X, Y)
End Sub

Private Sub LabelSequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSequencial, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub LabelParcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelParcela, Source, X, Y)
End Sub

Private Sub LabelParcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelParcela, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub NumeroLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroLabel, Source, X, Y)
End Sub

Private Sub NumeroLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroLabel, Button, Shift, X, Y)
End Sub

Private Sub ClienteLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ClienteLabel, Source, X, Y)
End Sub

Private Sub ClienteLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ClienteLabel, Button, Shift, X, Y)
End Sub

Function Traz_TitReceber_Tela(objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_TitReceber_Tela
    
    'Lê o Título à Receber
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 87211

    'Não encontrou o Título à Receber --> erro
    If lErro = 26061 Then gError 87212
    
    'Coloca o Cliente na Tela
    Cliente.Text = objTituloReceber.lCliente
    Call Cliente_Validate(bSGECancelDummy)

    'Coloca a Filial na Tela
    Filial.Text = objTituloReceber.iFilial
    Call Filial_Validate(bSGECancelDummy)
    
    'Coloca o Tipo na tela
    Tipo.Text = objTituloReceber.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)

    If objTituloReceber.lNumTitulo = 0 Then
        Numero.PromptInclude = False
        Numero.Text = ""
        Numero.PromptInclude = True
    Else
        Numero.PromptInclude = False
        Numero.Text = CStr(objTituloReceber.lNumTitulo)
        Numero.PromptInclude = True
    End If

    Call Numero_Validate(bSGECancelDummy)
    
    Call DateParaMasked(DataEmissao, objTituloReceber.dtDataEmissao)
    Call DataEmissao_Validate(bSGECancelDummy)
    
    Me.Show

Traz_TitReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TitReceber_Tela:

    Traz_TitReceber_Tela = gErr

    Select Case gErr

        Case 87211
        
        Case 87212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142803)

    End Select

    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134029

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134029

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142804)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoInstrucoes_Click()

Dim lErro As Long
Dim objOcorrRemParcRec As New ClassOcorrRemParcRec
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoInstrucoes_Click
       
    'Chama a tela
    Call Chama_Tela("InstrRemParcRec2Lista", colSelecao, objOcorrRemParcRec, objEventoSequencial)
    
    Exit Sub
    
Erro_BotaoInstrucoes_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142784)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub DataEmissao_GotFocus()
Dim iDataAux As Integer
    
    iDataAux = iDataEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iDataEmissaoAlterada = iDataAux

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iDataEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate
    
    'se a data não foi alterada ------> Fim
    If iDataEmissaoAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 26140
        
    'Se Número foi alterado zera glNUmIntParc
    glNumIntParc = 0

    Call Verifica_Alteracao
    
    iDataEmissaoAlterada = 0
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case Err

        Case 26140
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175191)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 26141

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 26141

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175195)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 26142

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 26142

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175196)

    End Select

    Exit Sub

End Sub

