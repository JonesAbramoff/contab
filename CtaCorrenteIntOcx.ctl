VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CtaCorrenteIntOcx 
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9240
   KeyPreview      =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   9240
   Begin VB.Frame Frame4 
      Caption         =   "Borderô de pagamento"
      Height          =   630
      Left            =   120
      TabIndex        =   48
      Top             =   6240
      Width           =   8895
      Begin VB.TextBox DirArqBordPagto 
         Height          =   285
         Left            =   2865
         MaxLength       =   250
         TabIndex        =   50
         Top             =   210
         Width           =   5490
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8355
         TabIndex        =   49
         Top             =   180
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Diretório padrão dos arquivos:"
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
         Index           =   3
         Left            =   195
         TabIndex        =   51
         Top             =   240
         Width           =   2730
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6885
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   165
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CtaCorrenteIntOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CtaCorrenteIntOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CtaCorrenteIntOcx.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CtaCorrenteIntOcx.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Banco"
      Height          =   2835
      Left            =   120
      TabIndex        =   27
      Top             =   1305
      Width           =   5625
      Begin VB.TextBox ConvenioPagto 
         Height          =   300
         Left            =   1815
         MaxLength       =   20
         TabIndex        =   45
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox CodBanco 
         Height          =   315
         Left            =   1215
         TabIndex        =   4
         Top             =   285
         Width           =   2115
      End
      Begin MSMask.MaskEdBox Agencia 
         Height          =   300
         Left            =   1230
         TabIndex        =   5
         Top             =   720
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DVAgencia 
         Height          =   300
         Left            =   1950
         TabIndex        =   6
         Top             =   720
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NumConta 
         Height          =   300
         Left            =   3150
         TabIndex        =   7
         Top             =   705
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DVNumConta 
         Height          =   300
         Left            =   4560
         TabIndex        =   8
         Top             =   705
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DVAgConta 
         Height          =   300
         Left            =   5145
         TabIndex        =   9
         Top             =   705
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Contato 
         Height          =   300
         Left            =   1215
         TabIndex        =   10
         Top             =   1170
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone1 
         Height          =   300
         Left            =   1230
         TabIndex        =   11
         Top             =   1605
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Telefone2 
         Height          =   300
         Left            =   2820
         TabIndex        =   12
         Top             =   1620
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CreditoRotativo 
         Height          =   300
         Left            =   1815
         TabIndex        =   47
         Top             =   2475
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Crédito Rotativo:"
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
         Index           =   2
         Left            =   285
         TabIndex        =   46
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Convênio p/Pagto:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   2100
         Width           =   1620
      End
      Begin VB.Label e 
         AutoSize        =   -1  'True
         Caption         =   "e"
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
         Left            =   2625
         TabIndex        =   28
         Top             =   1650
         Width           =   120
      End
      Begin VB.Label lblCodBanco 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   29
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
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
         Index           =   3
         Left            =   375
         TabIndex        =   30
         Top             =   750
         Width           =   765
      End
      Begin VB.Label lblLabels 
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
         Height          =   195
         Index           =   5
         Left            =   2535
         TabIndex        =   31
         Top             =   750
         Width           =   570
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
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
         Index           =   8
         Left            =   405
         TabIndex        =   32
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Telefones:"
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
         Index           =   9
         Left            =   225
         TabIndex        =   33
         Top             =   1650
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saldo Inicial"
      Height          =   765
      Left            =   135
      TabIndex        =   24
      Top             =   4200
      Width           =   5625
      Begin MSMask.MaskEdBox DataSaldoInicial 
         Height          =   300
         Left            =   1155
         TabIndex        =   13
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
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   3480
         TabIndex        =   14
         Top             =   345
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown SpinData 
         Height          =   315
         Left            =   2310
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   615
         TabIndex        =   34
         Top             =   375
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   2895
         TabIndex        =   35
         Top             =   375
         Width           =   510
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificação"
      Height          =   1185
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   5625
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1800
         Picture         =   "CtaCorrenteIntOcx.ctx":0994
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox Nome 
         Height          =   300
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   2
         Top             =   345
         Width           =   2295
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   1215
         TabIndex        =   0
         Top             =   345
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Descricao 
         Height          =   300
         Left            =   1215
         TabIndex        =   3
         Top             =   765
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Index           =   0
         Left            =   2250
         TabIndex        =   36
         Top             =   375
         Width           =   555
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   375
         Width           =   660
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Index           =   1
         Left            =   225
         TabIndex        =   38
         Top             =   795
         Width           =   930
      End
   End
   Begin VB.Frame FrameContabilidade 
      Caption         =   "Contabilidade"
      Height          =   1125
      Left            =   135
      TabIndex        =   26
      Top             =   5025
      Width           =   5625
      Begin MSMask.MaskEdBox ContaContabil 
         Height          =   315
         Left            =   1395
         TabIndex        =   15
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label ContaContabilLabel 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   675
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   39
         Top             =   285
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
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
         Left            =   315
         TabIndex        =   40
         Top             =   735
         Width           =   945
      End
      Begin VB.Label DescContaContabil 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1380
         TabIndex        =   41
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.ListBox ListaCodConta 
      Height          =   4950
      IntegralHeight  =   0   'False
      Left            =   6015
      TabIndex        =   16
      Top             =   1200
      Width           =   3015
   End
   Begin MSComctlLib.TreeView TvwContas 
      Height          =   4305
      Left            =   6015
      TabIndex        =   17
      Top             =   1200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7594
      _Version        =   393217
      Indentation     =   453
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Plano de Contas"
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
      Left            =   5955
      TabIndex        =   42
      Top             =   945
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   " Caixa/Conta Corrente Interna"
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
      Left            =   5955
      TabIndex        =   43
      Top             =   945
      Width           =   2640
   End
End
Attribute VB_Name = "CtaCorrenteIntOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoBancos As AdmEvento
Attribute objEventoBancos.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

Private Sub Agencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Agencia, iAlterado)

End Sub

Private Sub Agencia_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Agencia_Validate
        
    'Verifica  se  a Agencia  foi  informada
    If Len(Trim(Agencia.Text)) <> 0 Then
        
        'O código do banco deve ser preenchido
        If Len(Trim(CodBanco.Text)) = 0 Then Error 43372
        
    End If
        
    Exit Sub
        
Erro_Agencia_Validate:

    Cancel = True


    Select Case Err

        Case 43372
             lErro = Rotina_Erro(vbOKOnly, "ERRO_CODBANCO_NAO_INFORMADO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155305)

    End Select

    Exit Sub
       
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera número automático para o código de Conta
    lErro = CF("Conta_Automatica", iCodigo)
    If lErro <> SUCESSO Then Error 57701

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57701
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155306)
    
    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Verifica se codigo é menor que um
    If CInt(Codigo.Text) < 1 Then Error 55957
    
    objContaCorrenteInt.iCodigo = CInt(Codigo.Text)
    
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 43521
    
    If lErro = 11807 Then Exit Sub
    
    'Se alguma Filial tiver sido selecionada
    If giFilialEmpresa <> EMPRESA_TODA Then

        If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43522

    End If
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case Err

        Case 43521

        Case 43522
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, Codigo.Text, giFilialEmpresa)

        Case 55957
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INVALIDO1", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155307)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoBancos = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    
    'Carrega a lista de contas correntes
    lErro = Carga_Lista_ContasCorrentes()
    If lErro <> SUCESSO Then Error 11777
    
    'Carrega  a combo de bancos
    lErro = Preenche_Combo_Bancos()
    If lErro <> SUCESSO Then Error 11778
            
    'Carrega na tela os dados relativos à contabilidade
    lErro = Carrega_Contabilidade()
    If lErro <> SUCESSO Then Error 11779

    Call Limpa_Tela_CtaCorrenteInt
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err
    
    Select Case Err
    
        Case 11777, 11778, 11779

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155308)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Agencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Limpa_Tela_CtaCorrenteInt()
'Limpa a tela e carrega a data inicial da conta como sendo a data atual do sistema
    
    Call Limpa_Tela(Me)
    
    Call Limpa_Parte_Contabil
    
    Codigo.Text = ""
    
    DataSaldoInicial.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    CodBanco.ListIndex = -1
    
    CodBanco.Text = ""
    
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 55942
    
    Call Limpa_Tela_CtaCorrenteInt
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 55942
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155309)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim sNomeReduzido As String

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 43370
    
    objContaCorrenteInt.iCodigo = CInt(Codigo.Text)
    
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 43523
    
    If lErro <> SUCESSO Then Error 43524
    
    'Pedido de Confirmação de Exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CONTACORRENTE", objContaCorrenteInt.iCodigo)
    
    If vbMsgRes = vbYes Then
        
        'Rotina encarregada de excluir a conta do banco de dados
        lErro = CF("ContasCorrentesInt_Exclui", objContaCorrenteInt)
        If lErro <> SUCESSO Then Error 43371
        
        'Exclui a conta da lista de contas
        Call ListaCodConta_Exclui(objContaCorrenteInt.iCodigo)
        
        Call Limpa_Tela_CtaCorrenteInt
                
        Codigo.SetFocus
    
        iAlterado = 0
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 43370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
        
        Case 43371, 43523
    
        Case 43524
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155310)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
    
On Error GoTo Erro_BotaoLimpar_Click
    
    'Pergunta se dejesa salvar as mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80456
    
    Call Limpa_Tela_CtaCorrenteInt
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
    
        Case 80456

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155311)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CodBanco_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Codigo_GotFocus()
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        Label5.Visible = False
        TvwContas.Visible = False
        Label6.Visible = True
        ListaCodConta.Visible = True
    End If
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataSaldoInicial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataSaldoInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataSaldoInicial, iAlterado)

End Sub

Private Sub DataSaldoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaldoInicial_Validate

    'verifica se a data final está vazia
    If Len(DataSaldoInicial.ClipText) = 0 Then Error 11784

    'verifica se a data final é válida
    lErro = Data_Critica(DataSaldoInicial.Text)
    If lErro <> SUCESSO Then Error 11783

    Exit Sub

Erro_DataSaldoInicial_Validate:

    Cancel = True


    Select Case Err

        Case 11783

        Case 11784
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155312)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVAgConta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVAgConta_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DVAgConta_Validate

    If Len(Trim(DVAgConta.Text)) = 0 Then Exit Sub
    
    'Verifica se o numero da conta foi preenchido
    If Len(Trim(NumConta.Text)) = 0 Then Error 11785
    
    'Verifica se a agencia foi preenchida
    If Len(Trim(Agencia.Text)) = 0 Then Error 11786
    
    Exit Sub
    
Erro_DVAgConta_Validate:

    Cancel = True


    Select Case Err
    
        Case 11785
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
        
        Case 11786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_NAO_PREENCHIDA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155313)
        
    End Select
    
    Exit Sub
    
End Sub

Private Sub DVAgencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DVAgencia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DVAgencia_Validate

    If Len(Trim(DVAgencia.Text)) = 0 Then Exit Sub
    
    'Verifica se a agencia foi preenchida
    If Len(Trim(Agencia.Text)) = 0 Then Error 11787
    
    Exit Sub

Erro_DVAgencia_Validate:

    Cancel = True


    Select Case Err
    
        Case 11787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_NAO_PREENCHIDA", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155314)
    
    End Select
    
    Exit Sub

End Sub

Private Sub DVNumConta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carga_Lista_ContasCorrentes() As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carga_Lista_ContasCorrentes

   'leitura dos codigos e descricoes das Contas no BD
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 11788

    'preenche listbox com descricao das contas
    For iIndice = 1 To colCodigoNomeRed.Count
        Set objCodigoNome = colCodigoNomeRed(iIndice)
        ListaCodConta.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        ListaCodConta.ItemData(ListaCodConta.NewIndex) = objCodigoNome.iCodigo
    Next

    Carga_Lista_ContasCorrentes = SUCESSO

    Exit Function

Erro_Carga_Lista_ContasCorrentes:

    Carga_Lista_ContasCorrentes = Err

    Select Case Err

        Case 11788

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155315)

    End Select

    Exit Function
    
End Function

Private Function Preenche_Combo_Bancos() As Long

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Combo_Bancos

    'leitura dos codigos e descricoes das ListaCodConta de venda no BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then Error 11789
   
   'preenche ComboBox com código e nome dos CodBancoes
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        CodBanco.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodBanco.ItemData(CodBanco.NewIndex) = objCodigoNome.iCodigo
    Next

    Preenche_Combo_Bancos = SUCESSO
   
    Exit Function
    
Erro_Preenche_Combo_Bancos:

    Preenche_Combo_Bancos = Err
    
    Select Case Err
     
        Case 11789
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155316)
            
    End Select
    
    Exit Function
    
End Function

Function Trata_Parametros(Optional objContaCorrenteInt As ClassContasCorrentesInternas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    'Verifica se alguma conta foi passada como parametro
    If Not (objContaCorrenteInt Is Nothing) Then
        
        lErro = Traz_Dados_Tela(objContaCorrenteInt)
        If lErro <> SUCESSO Then Error 11790
            
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err
    
        Case 11790
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155317)
    
    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Tela_Extrai

    sTabela = "ContasCorrentesInternas"

    If Len(Trim(Codigo.ClipText)) > 0 Then
        objContaCorrenteInt.iCodigo = CInt(Codigo.Text)
    Else
        objContaCorrenteInt.iCodigo = 0
    End If

    objContaCorrenteInt.sNomeReduzido = Nome.Text
    objContaCorrenteInt.sDescricao = Descricao.Text

    If Len(CodBanco.Text) > 0 Then

        objContaCorrenteInt.iCodBanco = Codigo_Extrai(CodBanco.Text)
        objContaCorrenteInt.sAgencia = Agencia.Text
        objContaCorrenteInt.sDVAgencia = DVAgencia.Text
        objContaCorrenteInt.sNumConta = NumConta.Text
        objContaCorrenteInt.sDVNumConta = DVNumConta.Text
        objContaCorrenteInt.sDVAgConta = DVAgConta.Text
    Else
        objContaCorrenteInt.iCodBanco = 0
    End If


    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "Codigo", objContaCorrenteInt.iCodigo, 0, "Codigo"
    colCampoValor.Add "NomeReduzido", objContaCorrenteInt.sNomeReduzido, STRING_NOME, "NomeReduzido"
    colCampoValor.Add "Descricao", objContaCorrenteInt.sDescricao, STRING_CONTA_CORRENTE_DESCRICAO, "Descricao"
    colCampoValor.Add "CodBanco", objContaCorrenteInt.iCodBanco, 0, "CodBanco"
    colCampoValor.Add "Agencia", objContaCorrenteInt.sAgencia, STRING_AGENCIA, "Agencia"
    colCampoValor.Add "DVAgencia", objContaCorrenteInt.sDVAgencia, STRING_DV, "DVAgencia"
    colCampoValor.Add "NumConta", objContaCorrenteInt.sNumConta, STRING_NUMCONTA, "NumConta"
    colCampoValor.Add "DVNumConta", objContaCorrenteInt.sDVNumConta, STRING_DV, "DVNumConta"
    colCampoValor.Add "DVAgConta", objContaCorrenteInt.sDVAgConta, STRING_DV, "DVAgConta"
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155318)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Tela_Preenche

    objContaCorrenteInt.iCodigo = colCampoValor.Item("Codigo").vValor

    If objContaCorrenteInt.iCodigo <> 0 Then
        
        objContaCorrenteInt.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objContaCorrenteInt.sDescricao = colCampoValor.Item("Descricao").vValor
            
        lErro = Traz_Dados_Tela(objContaCorrenteInt)
        If lErro <> SUCESSO Then Error 34673
        
        iAlterado = 0
        
    End If

    Exit Sub
    
Erro_Tela_Preenche:

    Select Case Err
    
        Case 34673
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155319)
            
    End Select

    Exit Sub
    
End Sub

Private Function Traz_Dados_Tela(objContaCorrenteInt As ClassContasCorrentesInternas) As Long
'Traz os dados da conta corrente passada como parametro

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sContaMascarada As String

On Error GoTo Erro_Traz_Dados_Tela

    'Le os dados da conta passada como parametro
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 7691

    'Verifica sua existencia
    If lErro = 11807 Then
    
        Codigo.Text = objContaCorrenteInt.iCodigo
    
    Else
        'Coloca na tela os dados da conta passada por parametro
        Codigo.Text = CStr(objContaCorrenteInt.iCodigo)
        Nome.Text = objContaCorrenteInt.sNomeReduzido
        Descricao.Text = objContaCorrenteInt.sDescricao
        Agencia.Text = objContaCorrenteInt.sAgencia
        DVAgencia.Text = objContaCorrenteInt.sDVAgencia
        NumConta.Text = objContaCorrenteInt.sNumConta
        DVAgConta.Text = objContaCorrenteInt.sDVAgConta
        DVNumConta.Text = objContaCorrenteInt.sDVNumConta
        Contato.Text = objContaCorrenteInt.sContato
        ConvenioPagto.Text = objContaCorrenteInt.sConvenioPagto
        Telefone1.Text = objContaCorrenteInt.sTelefone
        Telefone2.Text = objContaCorrenteInt.sFax
        Valor.Text = CStr(objContaCorrenteInt.dSaldoInicial)
        CreditoRotativo.Text = CStr(objContaCorrenteInt.dRotativo)
        
        Call DateParaMasked(DataSaldoInicial, objContaCorrenteInt.dtDataInicial)
        
        lErro = Traz_Contabilidade_Tela(objContaCorrenteInt)
        If lErro <> SUCESSO Then Error 11792
        
        For iIndice = 0 To CodBanco.ListCount - 1
        
            If CodBanco.ItemData(iIndice) = objContaCorrenteInt.iCodBanco Then
                CodBanco.ListIndex = iIndice
                Exit For
            End If
        Next
        If iIndice = CodBanco.ListCount Then CodBanco.ListIndex = -1
        
        DirArqBordPagto.Text = objContaCorrenteInt.sDirArqBordPagto
    
    End If
    
    Traz_Dados_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Dados_Tela:

    Traz_Dados_Tela = Err
    
    Select Case Err
    
        Case 7691, 11792
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155320)
    
    End Select
            
    Exit Function
            
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objEventoBancos = Nothing
    Set objEventoContaCorrenteInt = Nothing
    Set objEventoContaContabil = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
End Sub

Private Sub lblCodBanco_Click()

Dim objBanco As New ClassBanco
Dim colSelecao As Collection

    If Len(CodBanco.Text) = 0 Then
        objBanco.iCodBanco = 0
    Else
        objBanco.iCodBanco = CodBanco.ItemData(CodBanco.ListIndex)
    End If

    Call Chama_Tela("BancoLista", colSelecao, objBanco, objEventoBancos)

End Sub

Private Sub lblCodigo_Click()

Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim colSelecao As Collection

    If Len(Codigo.Text) = 0 Then
        objContaCorrenteInt.iCodigo = 0
    Else
        objContaCorrenteInt.iCodigo = CInt(Codigo.Text)
    End If

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContaCorrenteInt, objEventoContaCorrenteInt)

End Sub

Private Sub ListaCodConta_DblClick()

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_ListaCodConta_DblClick

    'Pega o codigo da conta corrente selecionada
    objContaCorrenteInt.iCodigo = ListaCodConta.ItemData(ListaCodConta.ListIndex)
    
    'Traz para a tela os dados da conta selecionada
    lErro = Traz_Dados_Tela(objContaCorrenteInt)
    If lErro <> SUCESSO Then Error 11793
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    iAlterado = 0
    
    Exit Sub
    
Erro_ListaCodConta_DblClick:

    Select Case Err
    
        Case 11793
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155321)
    End Select
    
    Exit Sub
    
End Sub

Private Sub Nome_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Nome_GotFocus()

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        Label5.Visible = False
        TvwContas.Visible = False
        Label6.Visible = True
        ListaCodConta.Visible = True
    End If

End Sub

Private Sub NumConta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumConta_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumConta, iAlterado)

End Sub

Private Sub NumConta_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_NumConta_Validate

    'Verifica  se  o numero da conta  foi  informada
    If Len(Trim(NumConta.Text)) <> 0 Then
        
        'A Agencia deve ser preenchida
        If Len(Trim(Agencia.Text)) = 0 Then Error 22036
        
    End If
        
    Exit Sub
        
Erro_NumConta_Validate:

    Cancel = True


    Select Case Err

        Case 22036
             lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_NAO_PREENCHIDA", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155322)

    End Select

    Exit Sub
       
End Sub

Private Sub ContaContabilLabel_Click()
'Chama browse de plano de contas

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String
Dim lErro As Long

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then Error 57748

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case Err

    Case 57748

    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155323)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)
'Retorno do browse do plano de contas

Dim objPlanoConta As ClassPlanoConta
Dim lErro As Long
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then

        ContaContabil.Text = ""

    Else

        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 57749

        ContaContabil.Text = sContaEnxuta

        ContaContabil.PromptInclude = True

    End If

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case Err

        Case 57749
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155324)

    End Select

    Exit Sub

End Sub

Private Sub objEventoBancos_evSelecao(obj1 As Object)
    
Dim objBanco As ClassBanco
Dim iIndice As Integer
    
    Set objBanco = obj1

    For iIndice = 0 To CodBanco.ListCount - 1

        If CodBanco.ItemData(iIndice) = objBanco.iCodBanco Then
            CodBanco.ListIndex = iIndice
            Exit For
        End If
    Next
    
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show
    
End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

Dim lErro As Long
    
On Error GoTo Erro_objEventoContaCorrenteInt_evSelecao
    
    Set objContaCorrenteInt = obj1
    
    'Traz para tela os dados da conta selecionada
    lErro = Traz_Dados_Tela(objContaCorrenteInt)
    If lErro <> SUCESSO Then Error 11794
    
    iAlterado = 0
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoContaCorrenteInt_evSelecao:

    Select Case Err
    
        Case 11794
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155325)
            
    End Select
        
    Exit Sub

End Sub

Private Sub SpinData_DownClick()
Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    DataSaldoInicial.SetFocus

    'verifica se a data foi preenchida
    If Len(Trim(DataSaldoInicial.ClipText)) > 0 Then

        sData = DataSaldoInicial.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 11795

        DataSaldoInicial.PromptInclude = False
        DataSaldoInicial.Text = sData
        DataSaldoInicial.PromptInclude = True
        
    End If

    Exit Sub

Erro_SpinData_DownClick:

    Select Case Err

        Case 11795

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155326)

    End Select

    Exit Sub

End Sub

Private Sub Telefone1_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Telefone2_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)

Dim lErro As Long

On Error GoTo Erro_TvwContas_Expand

    If objNode.Tag <> NETOS_NA_ARVORE Then
    
        'move os dados do plano de contas do banco de dados para a arvore colNodes.
        lErro = CF("Carga_Arvore_Conta_Modulo1", objNode, TvwContas.Nodes, MODULO_TESOURARIA)
        If lErro <> SUCESSO Then Error 40808
        
    End If
    
    Exit Sub
    
Erro_TvwContas_Expand:

    Select Case Err
    
        Case 40808
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155327)
        
    End Select
        
    Exit Sub
    
End Sub

Private Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodBanco_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objBanco As New ClassBanco
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_CodBanco_Validate

    'verifica se foi preenchido o CodBanco
    If Len(Trim(CodBanco.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox CodBanco
    If CodBanco.Text = CodBanco.List(CodBanco.ListIndex) Then Exit Sub

    'tenta Selecionar o banco com aquele codigo
    lErro = Combo_Seleciona(CodBanco, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 11799
    
    If lErro = 6730 Then
    
        objBanco.iCodBanco = iCodigo
        
        'Verifica se o banco esta no BD
        lErro = CF("Banco_Le", objBanco)
        If lErro <> SUCESSO And lErro <> 16091 Then Error 11800
        
        If lErro = 16091 Then Error 18206
        
        CodBanco.AddItem CStr(objBanco.iCodBanco) & SEPARADOR & objBanco.sNomeReduzido
        CodBanco.ItemData(CodBanco.NewIndex) = objBanco.iCodBanco
        
        CodBanco.ListIndex = CodBanco.NewIndex
    
    End If
    
    If lErro = 6731 Then Error 11798
        
    Exit Sub

Erro_CodBanco_Validate:

    Cancel = True

    Select Case Err

        Case 11798
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", Err, CodBanco.Text)
         
        Case 11799, 11800

        Case 18206
            'Se o banco nao estiver no BD pergunta se quer criar
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODBANCO_INEXISTENTE", objBanco.iCodBanco)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("Bancos", objBanco)
            End If
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155328)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    DataSaldoInicial.SetFocus

    'Verifica de a data foi preenchida
    If Len(Trim(DataSaldoInicial.ClipText)) > 0 Then

        sData = DataSaldoInicial.Text

        'aumenta a data de um dia
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 11802

        DataSaldoInicial.PromptInclude = False
        DataSaldoInicial.Text = sData
        DataSaldoInicial.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_UpClick:

    Select Case Err

        Case 11802

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155329)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi informado
    If Len(Trim(Codigo.Text)) = 0 Then Error 11809
    
    'Verifica se o nome foi informado
    If Len(Trim(Nome.Text)) = 0 Then Error 11810
    
    'Verifica se a data de saldo inicial foi informada
    If Len(Trim(DataSaldoInicial.ClipText)) = 0 Then Error 11811
    
    'Verifica se o banco foi informado
    If Len(Trim(CodBanco.Text)) <> 0 Then
    
           'A Agência deve estar preenchida
           If Len(Trim(Agencia.Text)) = 0 Then Error 22033
           
           'O Número da conta deve estar preenchido
           If Len(Trim(NumConta.Text)) = 0 Then Error 22034
    
    End If
    
    'Recolhe os dados da Tela
    lErro = Move_Tela_Memoria(objContaCorrenteInt)
    If lErro <> SUCESSO Then Error 11812
        
    lErro = Trata_Alteracao(objContaCorrenteInt, objContaCorrenteInt.iCodigo)
    If lErro <> SUCESSO Then Error 43373
                
    'Grava no BD a nova conta corrente
    lErro = CF("ContasCorrentesInt_Grava", objContaCorrenteInt)
    If lErro <> SUCESSO Then Error 11813
    
    'Exclui ( se existir) da lista de contas correntes
    Call ListaCodConta_Exclui(objContaCorrenteInt.iCodigo)
    
    'Adiciona na lista de codigo de contas
    Call ListaCodConta_Adiciona(objContaCorrenteInt)
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 11809
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)
        
        Case 11810
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)
            
        Case 11811
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", Err)
            
        Case 11812, 11813, 43373
        
        Case 22033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_NAO_PREENCHIDA", Err)
        
        Case 22034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMCONTA_NAO_PREENCHIDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155330)
    
    End Select

    Exit Function

End Function

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    If Len(Valor.Text) > 0 Then
    
        'Faza critica do valor do saldo Inicial
        lErro = Valor_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 11814
                
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err
    
        Case 11814
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155331)
            
    End Select
        
    Exit Sub

End Sub

Private Sub DVNumConta_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DVNumConta_Validate

    If Len(Trim(DVNumConta.Text)) = 0 Then Exit Sub
    
    'Verifica de o numero da conta foi preenchido
    If Len(Trim(NumConta.Text)) = 0 Then Error 11815
    
    Exit Sub

Erro_DVNumConta_Validate:

    Cancel = True


    Select Case Err
    
        Case 11815
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMCONTA_NAO_PREENCHIDO", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155332)
    
    End Select
    
    Exit Sub

End Sub

Private Function Move_Tela_Memoria(objContaCorrenteInt As ClassContasCorrentesInternas) As Long
'Move os dados da Tela para a memoria

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objContaCorrenteInt.iCodigo = Codigo.Text
    objContaCorrenteInt.sNomeReduzido = Nome.Text
    objContaCorrenteInt.sDescricao = Descricao.Text
    objContaCorrenteInt.iCodBanco = Codigo_Extrai(CodBanco.Text)
    objContaCorrenteInt.sAgencia = Agencia.Text
    objContaCorrenteInt.sDVAgencia = DVAgencia.Text
    objContaCorrenteInt.sNumConta = NumConta.Text
    objContaCorrenteInt.sDVNumConta = DVNumConta.Text
    objContaCorrenteInt.sDVAgConta = DVAgConta.Text
    objContaCorrenteInt.sContato = Contato.Text
    objContaCorrenteInt.sConvenioPagto = ConvenioPagto.Text
    objContaCorrenteInt.sTelefone = Telefone1.Text
    objContaCorrenteInt.sFax = Telefone2.Text
    objContaCorrenteInt.dtDataInicial = CDate(DataSaldoInicial.Text)
    objContaCorrenteInt.sDirArqBordPagto = DirArqBordPagto.Text
    
    If Len(Trim(Valor.Text)) = 0 Then
        objContaCorrenteInt.dSaldoInicial = 0
    Else
        objContaCorrenteInt.dSaldoInicial = CDbl(Valor.Text)
    End If
    
    If Len(Trim(CreditoRotativo.Text)) = 0 Then
        objContaCorrenteInt.dRotativo = 0
    Else
        objContaCorrenteInt.dRotativo = CDbl(CreditoRotativo.Text)
    End If
    
    
    'Recolhe da tela os dados relativos a contabilidade
    lErro = Move_Contabil_Memoria(objContaCorrenteInt)
    If lErro <> SUCESSO Then Error 11842

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = SUCESSO
    
    Select Case Err
    
        Case 11842
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155333)

    End Select
    
    Exit Function

End Function

Private Function Move_Contabil_Memoria(objContaCorrenteInt As ClassContasCorrentesInternas)

Dim lErro As Long
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Move_Contabil_Memoria
 
    'Verifica se o modulo de contabilidade esta ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
            'Guarda a conta corrente
            lErro = CF("Conta_Formata", ContaContabil.Text, sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then Error 11843
            
            objContaCorrenteInt.sContaContabil = sContaFormatada
            
        End If
        
    Else
        objContaCorrenteInt.sContaContabil = ""
        
    End If

    Move_Contabil_Memoria = SUCESSO

    Exit Function
    
Erro_Move_Contabil_Memoria:

    Move_Contabil_Memoria = Err
    
    Select Case Err
    
        Case 11843
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155334)
    
    End Select
    
    Exit Function
    
End Function

Function Carrega_Contabilidade() As Long
'faz inicializacoes da arvore de plano de contas e da mascara da conta contabil

Dim lErro As Long, sMascaraConta As String

On Error GoTo Erro_Carrega_Contabilidade

    'verifica se o modulo relativo a contabilidade esta ativo
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then

        'carrega a arvore de contas contabeis
        lErro = CF("Carga_Arvore_Conta_Modulo", TvwContas.Nodes, MODULO_TESOURARIA)
        If lErro <> SUCESSO Then Error 11858
    
        'Inicializa a mascara do campo de conta contabil
         
        'obtem a mascara da conta contabil
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 11816
        
        ContaContabil.Mask = sMascaraConta
    
    Else
    
        'Incluido a inicialização da máscara para não dar erro na gravação de clientes com conta mas que o módulo de contabilidade foi desabilitado
        lErro = MascaraConta(sMascaraConta)
        If lErro <> SUCESSO Then Error 11816
        
        ContaContabil.Mask = sMascaraConta
        
        ContaContabil.Enabled = False
        
    End If

    Carrega_Contabilidade = SUCESSO
    
    Exit Function
    
Erro_Carrega_Contabilidade:

    Carrega_Contabilidade = Err
    
    Select Case Err
    
        Case 11816, 11858
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155335)
            
    End Select
    
    Exit Function

End Function

Private Sub Limpa_Parte_Contabil()

    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        DescContaContabil.Caption = ""
    End If

End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaContabil_Validate

    If Len(Trim(ContaContabil.ClipText)) > 0 Then
    
        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then Error 39801
        
        If lErro = SUCESSO Then
        
            sContaFormatada = objPlanoConta.sConta
            
            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)
            
            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then Error 39802
            
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaMascarada
            ContaContabil.PromptInclude = True
        
        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then
    
            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
            If lErro <> SUCESSO And lErro <> 5700 Then Error 11859
    
            'conta não cadastrada
            If lErro = 5700 Then Error 11860
             
        End If
        
        DescContaContabil.Caption = objPlanoConta.sDescConta

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True

    Select Case Err

        Case 11859, 39801

        Case 11860
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                Call Chama_Tela("PlanoConta", objPlanoConta)
            End If
        
        Case 39802
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objPlanoConta.sConta)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155336)

    End Select

    Exit Sub

End Sub

Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ContaContabil_GotFocus()
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        Label6.Visible = False
        ListaCodConta.Visible = False
        Label5.Visible = True
        TvwContas.Visible = True
    End If
End Sub

Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

Dim sConta As String
Dim sCaracterInicial As String
Dim lPosicaoSeparador As Long
Dim lErro As Long
Dim sContaEnxuta As String
Dim sContaMascarada As String
Dim cControl As Control
Dim iLinha As Integer

On Error GoTo Erro_TvwContas_NodeClick

    sCaracterInicial = left(Node.Key, 1)

    If sCaracterInicial = "A" Then

        sConta = right(Node.Key, Len(Node.Key) - 1)

        sContaEnxuta = String(STRING_CONTA, 0)

        'volta mascarado apenas os caracteres preenchidos
        lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
        If lErro <> SUCESSO Then Error 11864

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True

        'Preenche a Descricao da Conta
        lPosicaoSeparador = InStr(Node.Text, SEPARADOR)
        DescContaContabil.Caption = Mid(Node.Text, lPosicaoSeparador + 1)

    End If

    Exit Sub

Erro_TvwContas_NodeClick:

    Select Case Err

        Case 11864
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155337)

    End Select

    Exit Sub

End Sub

Private Function Traz_Contabilidade_Tela(objContaCorrenteInt As ClassContasCorrentesInternas)
    
Dim sContaMascarada As String
Dim lErro As Long
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Contabilidade_Tela
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
        
        If objContaCorrenteInt.sContaContabil <> "" Then
            sContaMascarada = String(STRING_CONTA, 0)
        
            lErro = Mascara_RetornaContaEnxuta(objContaCorrenteInt.sContaContabil, sContaMascarada)
            If lErro <> SUCESSO Then Error 11865
        Else
            sContaMascarada = ""
        End If
        
        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True
        
        Call ContaContabil_Validate(bCancel)
        
    End If
    
    Traz_Contabilidade_Tela = SUCESSO
            
    Exit Function

Erro_Traz_Contabilidade_Tela:

    Traz_Contabilidade_Tela = Err
    
    Select Case Err
    
        Case 11865
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, objContaCorrenteInt.sContaContabil)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155338)
            
    End Select

    Exit Function

End Function

Private Sub ListaCodConta_Exclui(iCodConta As Integer)

Dim iIndice As Integer

    For iIndice = 0 To ListaCodConta.ListCount - 1
    
        If ListaCodConta.ItemData(iIndice) = iCodConta Then
        
            ListaCodConta.RemoveItem (iIndice)
            Exit For
        
        End If
    
    Next

End Sub

Private Sub ListaCodConta_Adiciona(objContaCorrenteInt As ClassContasCorrentesInternas)
        
Dim sListBoxItem As String
Dim iIndice As Integer
    
    For iIndice = 0 To ListaCodConta.ListCount - 1
        
        If ListaCodConta.ItemData(iIndice) > objContaCorrenteInt.iCodigo Then Exit For
        
    Next
    
    'Concatena o código com a descrição do Histórico
    sListBoxItem = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido
    ListaCodConta.AddItem sListBoxItem, iIndice
    ListaCodConta.ItemData(iIndice) = objContaCorrenteInt.iCodigo
          
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_PLANO_CONTAS
    Set Form_Load_Ocx = Me
    Caption = "Caixa / Conta Corrente Interna"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CtaCorrenteInt"
    
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
        
        If Me.ActiveControl Is Codigo Then
            Call lblCodigo_Click
        ElseIf Me.ActiveControl Is CodBanco Then
            Call lblCodBanco_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        End If
    
    End If
    
End Sub

Private Sub lblLabels_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(lblLabels(Index), Source, X, Y)
End Sub

Private Sub lblLabels_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblLabels(Index), Button, Shift, X, Y)
End Sub


Private Sub e_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(e, Source, X, Y)
End Sub

Private Sub e_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(e, Button, Shift, X, Y)
End Sub

Private Sub lblCodBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(lblCodBanco, Source, X, Y)
End Sub

Private Sub lblCodBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblCodBanco, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub lblCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(lblCodigo, Source, X, Y)
End Sub

Private Sub lblCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(lblCodigo, Button, Shift, X, Y)
End Sub

Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub DescContaContabil_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescContaContabil, Source, X, Y)
End Sub

Private Sub DescContaContabil_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescContaContabil, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub ConvenioPagto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização física dos arquivos .html"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = String(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        DirArqBordPagto.Text = sBuffer
        Call DirArqBordPagto_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub
  
End Sub

Private Sub DirArqBordPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPos As Integer

On Error GoTo Erro_DirArqBordPagto_Validate

    If Len(Trim(DirArqBordPagto.Text)) = 0 Then Exit Sub
    
    If right(DirArqBordPagto.Text, 1) <> "\" And right(DirArqBordPagto.Text, 1) <> "/" Then
        iPos = InStr(1, DirArqBordPagto.Text, "/")
        If iPos = 0 Then
            DirArqBordPagto.Text = DirArqBordPagto.Text & "\"
        Else
            DirArqBordPagto.Text = DirArqBordPagto.Text & "/"
        End If
    End If

    If Len(Trim(Dir(DirArqBordPagto.Text, vbDirectory))) = 0 Then gError 192327

    Exit Sub

Erro_DirArqBordPagto_Validate:

    Cancel = True

    Select Case gErr

        Case 192327, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, DirArqBordPagto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192328)

    End Select

    Exit Sub

End Sub

Private Sub CreditoRotativo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CreditoRotativo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CreditoRotativo_Validate

    If Len(CreditoRotativo.Text) > 0 Then
    
        'Faza critica do valor do saldo Inicial
        lErro = Valor_Critica(CreditoRotativo.Text)
        If lErro <> SUCESSO Then gError 197838
                
    End If

    Exit Sub

Erro_CreditoRotativo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 197838
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197839)
            
    End Select
        
    Exit Sub

End Sub

