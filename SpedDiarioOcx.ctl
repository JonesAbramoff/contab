VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl SpedDiarioOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CheckBox EmpToda 
      Caption         =   "Empresa Toda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7725
      TabIndex        =   50
      Top             =   690
      Value           =   1  'Checked
      Width           =   1725
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   900
      Left            =   120
      TabIndex        =   20
      Top             =   30
      Width           =   2505
      Begin MSComCtl2.UpDown UpDownPeriodoDe 
         Height          =   330
         Left            =   1665
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   165
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoDe 
         Height          =   315
         Left            =   675
         TabIndex        =   0
         Top             =   165
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownPeriodoAte 
         Height          =   330
         Left            =   1665
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   495
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox PeriodoAte 
         Height          =   330
         Left            =   690
         TabIndex        =   1
         Top             =   495
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelPeriodoAte 
         Appearance      =   0  'Flat
         Caption         =   "Até:"
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
         Height          =   255
         Left            =   270
         TabIndex        =   25
         Top             =   540
         Width           =   450
      End
      Begin VB.Label LabelPeriodoDe 
         Appearance      =   0  'Flat
         Caption         =   "De:"
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
         Height          =   255
         Left            =   270
         TabIndex        =   24
         Top             =   195
         Width           =   390
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Localização do arquivo"
      Height          =   900
      Left            =   2670
      TabIndex        =   30
      Top             =   30
      Width           =   5010
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
         Height          =   345
         Left            =   4530
         TabIndex        =   23
         Top             =   210
         Width           =   360
      End
      Begin VB.TextBox NomeDiretorio 
         Height          =   285
         Left            =   900
         TabIndex        =   2
         Top             =   225
         Width           =   3615
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   285
         Left            =   915
         MaxLength       =   100
         TabIndex        =   3
         Top             =   540
         Width           =   3615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Diretório:"
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
         TabIndex        =   32
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo:"
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
         Left            =   180
         TabIndex        =   31
         Top             =   585
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Demais Informações"
      Height          =   4935
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   9285
      Begin VB.Frame Frame1 
         Caption         =   "Versão 3 ou superior"
         Height          =   2085
         Left            =   90
         TabIndex        =   41
         Top             =   2745
         Width           =   9105
         Begin VB.ComboBox TipoECD 
            Height          =   315
            ItemData        =   "SpedDiarioOcx.ctx":0000
            Left            =   1635
            List            =   "SpedDiarioOcx.ctx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   225
            Width           =   5310
         End
         Begin VB.Frame FrameTipoECD 
            Caption         =   "Identificação da SCP"
            Height          =   1410
            Index           =   2
            Left            =   90
            TabIndex        =   43
            Top             =   555
            Width           =   8925
            Begin VB.TextBox OBS01 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   3795
               MultiLine       =   -1  'True
               TabIndex        =   46
               Text            =   "SpedDiarioOcx.ctx":009C
               Top             =   165
               Width           =   5040
            End
            Begin MSMask.MaskEdBox CNPJSCP 
               Height          =   315
               Left            =   1530
               TabIndex        =   16
               Top             =   255
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               _Version        =   393216
               AllowPrompt     =   -1  'True
               MaxLength       =   14
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   " "
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "CNPJ/Código:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   315
               TabIndex        =   45
               Top             =   315
               Width           =   1200
            End
         End
         Begin VB.Frame FrameTipoECD 
            Caption         =   "Dados das SCPs"
            Height          =   1410
            Index           =   1
            Left            =   90
            TabIndex        =   44
            Top             =   555
            Width           =   8925
            Begin VB.TextBox NomeSCPGrid 
               BorderStyle     =   0  'None
               Height          =   275
               Left            =   2745
               MaxLength       =   250
               TabIndex        =   47
               Top             =   705
               Width           =   6135
            End
            Begin MSMask.MaskEdBox CNPJSCPGrid 
               Height          =   270
               Left            =   1095
               TabIndex        =   48
               Top             =   690
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   476
               _Version        =   393216
               BorderStyle     =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   14
               Format          =   "00\.000\.000\/0000-00; ; ; "
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridSCPs 
               Height          =   1095
               Left            =   105
               TabIndex        =   15
               Top             =   225
               Width           =   8715
               _ExtentX        =   15372
               _ExtentY        =   1931
               _Version        =   393216
               Rows            =   7
               Cols            =   4
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
            End
         End
         Begin VB.Label Forma 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   6960
            TabIndex        =   49
            Top             =   210
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "Ind.Tipo ECD:"
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
            Height          =   210
            Left            =   390
            TabIndex        =   42
            Top             =   285
            Width           =   1305
         End
      End
      Begin VB.Frame FrameVersao2 
         Caption         =   "Versão 2 ou superior"
         Height          =   1545
         Left            =   90
         TabIndex        =   36
         Top             =   1185
         Width           =   9105
         Begin VB.CheckBox EmpGrandePorte 
            Caption         =   "Empresa de grande porte"
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
            Height          =   255
            Left            =   3645
            TabIndex        =   13
            Top             =   1200
            Width           =   2490
         End
         Begin VB.CheckBox PossuiNIRE 
            Caption         =   "Possui registro na junta comercial (possui NIRE)"
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
            Height          =   255
            Left            =   1635
            TabIndex        =   9
            Top             =   135
            Width           =   4470
         End
         Begin VB.ComboBox Finalidade 
            Height          =   315
            ItemData        =   "SpedDiarioOcx.ctx":0108
            Left            =   1650
            List            =   "SpedDiarioOcx.ctx":0118
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   420
            Width           =   7410
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Empresa possui registro na junta comercial"
            Height          =   210
            Left            =   3960
            TabIndex        =   37
            Top             =   255
            Width           =   1845
         End
         Begin MSMask.MaskEdBox HashEscrSubst 
            Height          =   315
            Left            =   1650
            TabIndex        =   11
            Top             =   780
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NIRESubst 
            Height          =   315
            Left            =   1635
            TabIndex        =   12
            Top             =   1155
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "NIRE Subst:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   540
            TabIndex        =   40
            Top             =   1215
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hash Escr.Subst:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   135
            TabIndex        =   39
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label7 
            Caption         =   "Finalidade:"
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
            Height          =   210
            Left            =   690
            TabIndex        =   38
            Top             =   480
            Width           =   900
         End
      End
      Begin VB.ComboBox Versao 
         Height          =   315
         Left            =   8085
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   900
         Width           =   1080
      End
      Begin VB.ComboBox IndIniPer 
         Height          =   315
         ItemData        =   "SpedDiarioOcx.ctx":0179
         Left            =   1740
         List            =   "SpedDiarioOcx.ctx":0189
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   7425
      End
      Begin VB.ComboBox SituacaoEspecial 
         Height          =   315
         ItemData        =   "SpedDiarioOcx.ctx":0255
         Left            =   1740
         List            =   "SpedDiarioOcx.ctx":0268
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   540
         Width           =   7425
      End
      Begin MSMask.MaskEdBox NumOrd 
         Height          =   300
         Left            =   1740
         TabIndex        =   6
         Top             =   900
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ContaOutros 
         Height          =   315
         Left            =   4605
         TabIndex        =   7
         Top             =   900
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Versão:"
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
         Left            =   7395
         TabIndex        =   35
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número do Diário:"
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
         Left            =   135
         TabIndex        =   34
         Top             =   945
         Width           =   1545
      End
      Begin VB.Label LabelConta 
         AutoSize        =   -1  'True
         Caption         =   "Conta Outros:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   3360
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   930
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ind. Iní. Período:"
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
         Left            =   180
         TabIndex        =   29
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Situação Especial:"
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
         Left            =   90
         TabIndex        =   28
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.PictureBox Picture9 
      Height          =   555
      Left            =   7740
      ScaleHeight     =   495
      ScaleWidth      =   1620
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   105
      Width           =   1680
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   1110
         Picture         =   "SpedDiarioOcx.ctx":02AF
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   607
         Picture         =   "SpedDiarioOcx.ctx":042D
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGerar 
         Height          =   345
         Left            =   105
         Picture         =   "SpedDiarioOcx.ctx":095F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gera o arquivo"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "SpedDiarioOcx"
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

Dim iListIndexDefault As Integer

Private WithEvents objEventoConta As AdmEvento
Attribute objEventoConta.VB_VarHelpID = -1
Private gobjFilialEmpresa As AdmFiliais

Dim objGridSCPs As AdmGrid
Dim iGrid_CNPJSCP_Col As Integer
Dim iGrid_NomeSCP_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGerar_Click()

Dim lErro As Long, iIndice As Integer
Dim sDiretorio As String
Dim dtData As Date
Dim lNumOrd As Long
Dim sNomeArqParam As String
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim iIndIniPer As Integer, iFinalidade As Integer
Dim iTipoECD As Integer, sCodSCP As String, colSCPs As New Collection
Dim objFilial As AdmFiliais, iFilEmp As Integer

On Error GoTo Erro_BotaoGerar_Click
    
    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 203084
    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 203085
    If Len(Trim(NumOrd.Text)) = 0 Then gError 203111
    
    If right(NomeDiretorio.Text, 1) = "\" Or right(NomeDiretorio.Text, 1) = "/" Then
        sDiretorio = NomeDiretorio.Text & NomeArquivo.Text
    Else
        sDiretorio = NomeDiretorio.Text & "\" & NomeArquivo.Text
    End If
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatorios estao preenchidos
    If Len(Trim(PeriodoDe.ClipText)) = 0 Then gError 203086
    If Len(Trim(PeriodoAte.ClipText)) = 0 Then gError 203087
    
    'data inicial não pode ser maior que a data final
    If Len(Trim(PeriodoDe.ClipText)) <> 0 And Len(Trim(PeriodoAte.ClipText)) <> 0 Then

         If StrParaDate(PeriodoDe.Text) > StrParaDate(PeriodoAte.Text) Then gError 203088

    End If
    
    lNumOrd = StrParaLong(NumOrd.Text)
    
    If Len(ContaOutros.ClipText) > 0 Then
        'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaOutros.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 207367
                
        'conta não cadastrada
        If lErro = 5700 Then gError 207368
    
        If objPlanoConta.iNaturezaSped <> 9 Then gError 207371
    End If
    
    iIndIniPer = IndIniPer.ItemData(IndIniPer.ListIndex)
    iFinalidade = Codigo_Extrai(Finalidade.Text)
    
    If PossuiNIRE.Value = vbChecked And Len(Trim(gobjFilialEmpresa.sJucerja)) = 0 Then gError 213684
    If (iFinalidade = 1 Or iFinalidade = 2 Or iFinalidade = 3) And Len(Trim(HashEscrSubst.ClipText)) = 0 Then gError 213685
    If iFinalidade = 3 And Len(Trim(NIRESubst.ClipText)) = 0 Then gError 213686
    
    If Len(Trim(gobjFilialEmpresa.sSignatarioCTB)) = 0 Or Len(Trim(gobjFilialEmpresa.sCPFSignatarioCTB)) = 0 Or Len(Trim(gobjFilialEmpresa.sCodQualiSigCTB)) = 0 Then gError 213687
    If Len(Trim(gobjFilialEmpresa.sContador)) = 0 Or Len(Trim(gobjFilialEmpresa.sCPFContador)) = 0 Then gError 213688
    
    iTipoECD = Codigo_Extrai(TipoECD.Text)
    
    If iTipoECD = 1 Then
    
        'Para cada linha existente do Grid
        For iIndice = 1 To objGridSCPs.iLinhasExistentes
            Set objFilial = New AdmFiliais
            objFilial.sCgc = Replace(Replace(Replace(Replace(GridSCPs.TextMatrix(iIndice, iGrid_CNPJSCP_Col), "-", ""), ".", ""), "/", ""), "\", "")
            objFilial.sNome = GridSCPs.TextMatrix(iIndice, iGrid_NomeSCP_Col)
            
            If Len(Trim(objFilial.sCgc)) = 0 Then gError 213840
            'If Len(Trim(objFilial.sNome)) = 0 Then gError 99999
            
            colSCPs.Add objFilial
        Next
        
        If colSCPs.Count = 0 Then gError 213841
        
    ElseIf iTipoECD = 2 Then
    
        sCodSCP = CNPJSCP.ClipText
        
        If Len(Trim(sCodSCP)) = 0 Then gError 213842
    
    End If
    
    If EmpToda.Value = vbChecked Then
        iFilEmp = EMPRESA_TODA
    Else
        iFilEmp = giFilialEmpresa
    End If
    
    lErro = Sistema_Preparar_Batch(sNomeArqParam)
    If lErro <> SUCESSO Then gError 203089
    
    lErro = CF("Rotina_Sped_Contabil_Diario", sNomeArqParam, iFilEmp, sDiretorio, StrParaDate(PeriodoDe.Text), StrParaDate(PeriodoAte.Text), lNumOrd, sContaFormatada, iIndIniPer, Codigo_Extrai(SituacaoEspecial.Text), Versao.ItemData(Versao.ListIndex), IIf(PossuiNIRE.Value = vbChecked, MARCADO, DESMARCADO), iFinalidade, HashEscrSubst.ClipText, NIRESubst.ClipText, IIf(EmpGrandePorte.Value = vbChecked, MARCADO, DESMARCADO), iTipoECD, sCodSCP, colSCPs)
    If lErro <> SUCESSO Then gError 203253
        
    Call BotaoLimpar_Click
   
    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub
    
Erro_BotaoGerar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 203084
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_INFORMADO", gErr)
        
        Case 203085
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_INFORMADO", gErr)
        
        Case 203086
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIO_PERIODO_VAZIA1", gErr)
        
        Case 203087
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FINAL_PERIODO_VAZIA", gErr)
        
        Case 203088
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case 203089, 203253, 207367
        
        Case 203111
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_DIARIO_NAO_INFORMADO", gErr)
        
        Case 207368
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr)
        
        Case 207371
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NATUREZA_SPED_OUTROS", gErr)
            
        Case 213684
            Call Rotina_Erro(vbOKOnly, "ERRO_FILEMP_NIRE_NAO_PREENCHIDO", gErr)
        
        Case 213685
            Call Rotina_Erro(vbOKOnly, "ERRO_HASH_SUBSTIT_NAO_PREENCHIDO", gErr)
        
        Case 213686
            Call Rotina_Erro(vbOKOnly, "ERRO_NIRE_SUBSTIT_NAO_PREENCHIDA", gErr)
        
        Case 213687
            Call Rotina_Erro(vbOKOnly, "ERRO_DADOS_SIGNATARIO_ESCRITURACAO_INCOMPLETOS", gErr)
        
        Case 213688
            Call Rotina_Erro(vbOKOnly, "ERRO_DADOS_CONTADOR_INCOMPLETOS", gErr)
            
        Case 213840
            Call Rotina_Erro(vbOKOnly, "ERRO_IDENTIFICACAO_SCP_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 213841
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_SCP_CADASTRADA", gErr)
        
        Case 213842
            Call Rotina_Erro(vbOKOnly, "ERRO_IDENTIFICACAO_SCP_NAO_PREENCHIDO", gErr)
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203090)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    Call Limpa_Tela(Me)

    'Fecha comando de setas
    Call ComandoSeta_Fechar(Me.Name)
   
    NomeDiretorio.Text = CurDir
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
           Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203091)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoConta = Nothing
    Set gobjFilialEmpresa = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sMascaraConta As String
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Form_Load
    
'    'Inicializa a Lista de Plano de Contas
'    lErro = CF("Carga_Arvore_Conta", TvwContas.Nodes)
'    If lErro <> SUCESSO Then gError 207369

    Set objGridSCPs = New AdmGrid
        
    lErro = Inicializa_Grid_SCP(objGridSCPs)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Inicializa a máscara de Conta
    sMascaraConta = String(STRING_CONTA, 0)
    
    'le a mascara das contas
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TipoECD.ListIndex = 0
    
    ContaOutros.Mask = sMascaraConta
    
    Set objEventoConta = New AdmEvento
    
    lErro = Carrega_Versao
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    
    If objFilialEmpresa.iCodFilial = EMPRESA_TODA Then objFilialEmpresa.iCodFilial = 1
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
    
    Set gobjFilialEmpresa = objFilialEmpresa
    
    If Len(Trim(objFilialEmpresa.sJucerja)) <> 0 Then
        PossuiNIRE.Value = vbChecked
    Else
        PossuiNIRE.Value = vbUnchecked
    End If
    
    IndIniPer.ListIndex = 0
    Finalidade.ListIndex = 0
    
'    iListIndexDefault = Drive1.ListIndex
    
'    If Len(Trim(CurDir)) > 0 Then
'        Dir1.Path = CurDir
'        Drive1.Drive = Left(CurDir, 2)
'    End If
'
'    NomeDiretorio.Text = Dir1.Path
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203092)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    'Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    'gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203093)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long

On Error GoTo Erro_Tela_Preenche

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203094)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = 0
    Set Form_Load_Ocx = Me
    Caption = "Sped Diário"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "SpedDiario"

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


Private Sub BotaoProcurar_Click()

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

On Error GoTo Erro_BotaoProcurar_Click

    szTitle = "Localização do arquivos"
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        NomeDiretorio.Text = sBuffer
        Call NomeDiretorio_Validate(bSGECancelDummy)
  
    End If
  
    Exit Sub

Erro_BotaoProcurar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192326)

    End Select

    Exit Sub

End Sub

Private Sub NumOrd_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumOrd, iAlterado)
End Sub

Private Sub TipoECD_Click()
    Select Case Codigo_Extrai(TipoECD.Text)
        Case 0
            FrameTipoECD(1).Visible = False
            FrameTipoECD(2).Visible = False
            Forma.Caption = "G-Livro Diário"
        Case 1
            FrameTipoECD(1).Visible = True
            FrameTipoECD(2).Visible = False
            Forma.Caption = "G-Livro Diário"
        Case 2
            FrameTipoECD(1).Visible = False
            FrameTipoECD(2).Visible = True
            Forma.Caption = "S-Escrituração SCP"
    End Select
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
   ' Parent.UnloadDoFilho

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

Function Trata_Parametros(Optional obj1 As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203095)

    End Select

    Exit Function

End Function

'Private Sub Dir1_Change()
'
'     NomeDiretorio.Text = Dir1.Path
'
'End Sub

'Private Sub Drive1_Change()
'
'On Error GoTo Erro_Drive1_Change
'
'    Dir1.Path = Drive1.Drive
'
'    Exit Sub
'
'Erro_Drive1_Change:
'
'    Select Case Err
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 203096)
'
'    End Select
'
'    Drive1.ListIndex = iListIndexDefault
'
'    Exit Sub
'
'End Sub
'
'Private Sub Drive1_GotFocus()
'
'    iListIndexDefault = Drive1.ListIndex
'
'End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 203097

'    Drive1.Drive = Mid(NomeDiretorio.Text, 1, 2)
'
'    Dir1.Path = NomeDiretorio.Text

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 76, 203097
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203098)

    End Select

    Exit Sub

End Sub


Private Sub PeriodoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoDe_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoDe.Text)
    If lErro <> SUCESSO Then gError 203099

    Exit Sub

Erro_PeriodoDe_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 203099
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203100)
            
    End Select
    
    Exit Sub

End Sub

Private Sub PeriodoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PeriodoAte_Validate

    'Critica o valor data
    lErro = Data_Critica(PeriodoAte.Text)
    If lErro <> SUCESSO Then gError 203101

    Exit Sub

Erro_PeriodoAte_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 203101
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203102)
            
    End Select
    
    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 203103

    End If

    Exit Sub

Erro_UpDownPeriodoDe_DownClick:

    Select Case gErr

        Case 203103

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203104)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoDe_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoDe.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoDe, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 203105

    End If

    Exit Sub

Erro_UpDownPeriodoDe_UpClick:

    Select Case gErr

        Case 203105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203106)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_DownClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 203107

    End If

    Exit Sub

Erro_UpDownPeriodoAte_DownClick:

    Select Case gErr

        Case 203107

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203108)

    End Select

    Exit Sub

End Sub

Private Sub UpDownPeriodoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPeriodoAte_UpClick

    'Se a data está preenchida
    If Len(Trim(PeriodoAte.ClipText)) > 0 Then

        'Diminui a data em um dia
        lErro = Data_Up_Down_Click(PeriodoAte, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 203109

    End If

    Exit Sub

Erro_UpDownPeriodoAte_UpClick:

    Select Case gErr

        Case 203109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 203110)

    End Select

    Exit Sub

End Sub

Private Sub ContaOutros_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaOutros_Validate

    'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
    lErro = CF("Conta_Critica", ContaOutros.Text, sContaFormatada, objPlanoConta, MODULO_CONTABILIDADE)
    If lErro <> SUCESSO And lErro <> 5700 Then gError 207363
            
    'conta não cadastrada
    If lErro = 5700 Then gError 207364
    
    If Len(ContaOutros.ClipText) > 0 Then
        If objPlanoConta.iNaturezaSped <> 9 Then gError 207366
    End If
    
    Exit Sub

Erro_ContaOutros_Validate:

    Cancel = True
    
    ContaOutros.SetFocus

    Select Case gErr
    
        Case 207363
        
        Case 207364
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaOutros.Text)
            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada
                Call Chama_Tela("PlanoConta", objPlanoConta)
            End If
            
        Case 207366
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NATUREZA_SPED_OUTROS", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207365)
    
    End Select
    
    Exit Sub

End Sub

'Private Sub TvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
'
'Dim sConta As String
'Dim sCaracterInicial As String
'Dim lPosicaoSeparador As Long
'Dim lErro As Long
'Dim sContaEnxuta As String
'Dim sContaMascarada As String
'Dim cControl As Control
'Dim iLinha As Integer
'
'On Error GoTo Erro_TvwContas_NodeClick
'
'    sCaracterInicial = left(Node.Key, 1)
'
'    If sCaracterInicial <> "A" Then Error 20299
'
'    sConta = right(Node.Key, Len(Node.Key) - 1)
'
'    sContaEnxuta = String(STRING_CONTA, 0)
'
'    'volta mascarado apenas os caracteres preenchidos
'    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
'    If lErro <> SUCESSO Then Error 20300
'
'    ContaOutros.PromptInclude = False
'    ContaOutros.Text = sContaEnxuta
'    ContaOutros.PromptInclude = True
'
'    Exit Sub
'
'Erro_TvwContas_NodeClick:
'
'    Select Case Err
'
'        Case 20299
'
'        Case 20300
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", Err, sConta)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143122)
'
'    End Select
'
'    Exit Sub
'
'End Sub
'
'Private Sub TvwContas_Expand(ByVal objNode As MSComctlLib.Node)
'
'Dim lErro As Long
'
'On Error GoTo Erro_TvwContas_Expand
'
'    If objNode.Tag <> NETOS_NA_ARVORE Then
'
'        'move os dados do plano de contas do banco de dados para a arvore colNodes.
'        lErro = CF("Carga_Arvore_Conta1", objNode, TvwContas.Nodes)
'        If lErro <> SUCESSO Then Error 40798
'
'    End If
'
'    Exit Sub
'
'Erro_TvwContas_Expand:
'
'    Select Case Err
'
'        Case 40798
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143123)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub LabelConta_Click()

Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim sContaOrigem As String
Dim iContaPreenchida As Integer
Dim lErro As Long

On Error GoTo Erro_LabelConta_Click

    If Len(Trim(ContaOutros.ClipText)) > 0 Then
    
        lErro = CF("Conta_Formata", ContaOutros.Text, sContaOrigem, iContaPreenchida)
        If lErro <> SUCESSO Then gError 200998

        If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sContaOrigem
    Else
        objPlanoConta.sConta = ""
    End If
           
    'Chama a tela que lista os vendedores
    Call Chama_Tela("PlanoContaLista", colSelecao, objPlanoConta, objEventoConta)

    Exit Sub
    
Erro_LabelConta_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200999)
            
    End Select

    Exit Sub
    
    
End Sub

Private Sub objEventoConta_evSelecao(obj1 As Object)
    
Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sConta As String
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoConta_evSelecao
    
    Set objPlanoConta = obj1
    
    sConta = objPlanoConta.sConta
    
    sContaEnxuta = String(STRING_CONTA, 0)

    lErro = Mascara_RetornaContaEnxuta(sConta, sContaEnxuta)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ContaOutros.PromptInclude = False
    ContaOutros.Text = sContaEnxuta
    ContaOutros.PromptInclude = True
    Call ContaOutros_Validate(bSGECancelDummy)

    Me.Show
    
    Exit Sub
    
Erro_objEventoConta_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202001)
        
    End Select

    Exit Sub

End Sub

Private Function Carrega_Versao() As Long

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCodigoNome As AdmCodigoNome
Dim iIndice As Integer
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Carrega_Versao
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("Cod_Nomes_Le", "SpedCtbVersaoLeiaute", "Codigo", "Versao", STRING_MAXIMO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iIndice = -1
    For Each objCodigoNome In colCodigoDescricao
    
        iIndice = iIndice + 1
        
        'coloca na combo
        Versao.AddItem objCodigoNome.sNome
        Versao.ItemData(Versao.NewIndex) = objCodigoNome.iCodigo

    Next
    
    Versao.ListIndex = iIndice

    Carrega_Versao = SUCESSO

    Exit Function

Erro_Carrega_Versao:

    Carrega_Versao = gErr

    Select Case gErr

        'Erro já tratado
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207403)

    End Select

    Exit Function

End Function

Private Sub HashEscrSubst_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer, sCaracter As String

On Error GoTo Erro_HashEscrSubst_Validate

    HashEscrSubst.Text = UCase(HashEscrSubst.Text)

    For iIndice = 1 To Len(Trim(HashEscrSubst.Text))
        sCaracter = Mid(Trim(HashEscrSubst.Text), iIndice, 1)
        Select Case sCaracter
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"
            
            Case Else
                gError 203099
        End Select
    Next

    Exit Sub

Erro_HashEscrSubst_Validate:
    
    Cancel = True

    Select Case gErr
    
        Case 203099
            Call Rotina_Erro(vbOKOnly, "ERRO_HASH_SUBSTIT_INVALIDO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 203100)
            
    End Select
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_SCP(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("CNPJ/Código")
    objGridInt.colColuna.Add ("Nome")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (CNPJSCPGrid.Name)
    objGridInt.colCampo.Add (NomeSCPGrid.Name)

    'Colunas do Grid
    iGrid_CNPJSCP_Col = 1
    iGrid_NomeSCP_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridSCPs

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 101

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridSCPs.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_SCP = SUCESSO

    Exit Function

End Function

Public Sub GridSCPs_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSCPs, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSCPs, iAlterado)
    End If

End Sub

Public Sub GridSCPs_EnterCell()

    Call Grid_Entrada_Celula(objGridSCPs, iAlterado)

End Sub

Public Sub GridSCPs_GotFocus()

    Call Grid_Recebe_Foco(objGridSCPs)

End Sub

Public Sub GridSCPs_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridSCPs)

End Sub

Public Sub GridSCPs_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridSCPs, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSCPs, iAlterado)
    End If

End Sub

Public Sub GridSCPs_LeaveCell()

    Call Saida_Celula(objGridSCPs)

End Sub

Public Sub GridSCPs_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGridSCPs)

End Sub

Public Sub GridSCPs_RowColChange()

    Call Grid_RowColChange(objGridSCPs)

End Sub

Public Sub GridSCPs_Scroll()

    Call Grid_Scroll(objGridSCPs)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridSCPs.Name

                Select Case GridSCPs.Col
        
                    Case iGrid_CNPJSCP_Col
        
                        'Chama SaidaCelula de Categoria
                        lErro = Saida_Celula_Padrao(objGridInt, CNPJSCPGrid, True)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                    Case iGrid_NomeSCP_Col
        
                        'Chama SaidaCelula de Valor
                        lErro = Saida_Celula_Padrao(objGridInt, NomeSCPGrid)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
                End Select
                                
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 33025

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 33025
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155604)

    End Select

    Exit Function

End Function

Public Sub CNPJSCPGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CNPJSCPGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSCPs)
End Sub

Public Sub CNPJSCPGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSCPs)
End Sub

Public Sub CNPJSCPGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridSCPs.objControle = CNPJSCPGrid
    lErro = Grid_Campo_Libera_Foco(objGridSCPs)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub NomeSCPGrid_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NomeSCPGrid_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSCPs)
End Sub

Public Sub NomeSCPGrid_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSCPs)
End Sub

Public Sub NomeSCPGrid_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridSCPs.objControle = NomeSCPGrid
    lErro = Grid_Campo_Libera_Foco(objGridSCPs)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        Case CNPJSCPGrid.Name
        
            If Len(Trim(GridSCPs.TextMatrix(iLinha, iGrid_CNPJSCP_Col))) = 0 Then
                CNPJSCPGrid.Enabled = True
            Else
                CNPJSCPGrid.Enabled = False
            End If
            
        Case Else
            objControl.Enabled = True
            
        End Select
            
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156347)

    End Select

    Exit Sub

End Sub

