VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl OperacoesECF 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   5040
   Begin VB.Frame Frame3 
      Caption         =   "DAV Emitidos - Arquivo Eletrônico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   5130
      TabIndex        =   34
      Top             =   5640
      Width           =   4935
      Begin VB.Frame Frame4 
         Caption         =   "Intervalo de Datas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   210
         TabIndex        =   36
         Top             =   360
         Width           =   4380
         Begin MSMask.MaskEdBox DataDeDAVArquivo 
            Height          =   420
            Left            =   585
            TabIndex        =   37
            Top             =   495
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   435
            Left            =   4080
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   480
            Width           =   180
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAteDAVArquivo 
            Height          =   420
            Left            =   2790
            TabIndex        =   39
            Top             =   480
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   435
            Left            =   1920
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   495
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2265
            TabIndex        =   42
            Top             =   540
            Width           =   510
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   41
            Top             =   525
            Width           =   435
         End
      End
      Begin VB.CommandButton BotaoDAVArquivo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1260
         Picture         =   "OperacoesECF.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1770
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DAV Emitidos - Relatório Gerencial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   90
      TabIndex        =   25
      Top             =   5640
      Width           =   4815
      Begin VB.CommandButton BotaoDAVRelGer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1260
         Picture         =   "OperacoesECF.ctx":3642
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1770
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Intervalo de Datas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   135
         TabIndex        =   26
         Top             =   360
         Width           =   4380
         Begin MSMask.MaskEdBox DataDeDavRelGer 
            Height          =   420
            Left            =   585
            TabIndex        =   27
            Top             =   495
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAteDAV1 
            Height          =   435
            Left            =   4080
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   480
            Width           =   180
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAteDAVRelGer 
            Height          =   420
            Left            =   2790
            TabIndex        =   29
            Top             =   480
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataDeDAV1 
            Height          =   435
            Left            =   1920
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   495
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   31
            Top             =   525
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2265
            TabIndex        =   30
            Top             =   540
            Width           =   510
         End
      End
   End
   Begin VB.Frame FrameMemoriaFiscal 
      Caption         =   "Memória Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   5145
      TabIndex        =   14
      Top             =   90
      Width           =   4935
      Begin VB.CommandButton Sair 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   2760
         Picture         =   "OperacoesECF.ctx":6C84
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4635
         Width           =   1935
      End
      Begin VB.OptionButton Completa 
         Caption         =   "Completa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Value           =   -1  'True
         Width           =   2655
      End
      Begin VB.CommandButton BotaoLeituraMemoriaFiscal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   405
         Picture         =   "OperacoesECF.ctx":AA06
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4620
         Width           =   1935
      End
      Begin VB.OptionButton MemoriaFiscalReducoes 
         Caption         =   "Intervalo de Reduções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2835
         Width           =   3180
      End
      Begin VB.Frame FrameReducoes 
         Caption         =   "Intervalo de Reduções"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   375
         TabIndex        =   20
         Top             =   3210
         Width           =   4380
         Begin MSMask.MaskEdBox ReducaoDe 
            Height          =   420
            Left            =   840
            TabIndex        =   10
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ReducaoAte 
            Height          =   420
            Left            =   3000
            TabIndex        =   11
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "#####"
            PromptChar      =   " "
         End
         Begin VB.Label LabelReducaoDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   285
            TabIndex        =   22
            Top             =   540
            Width           =   435
         End
         Begin VB.Label LabelReducaoAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2400
            TabIndex        =   21
            Top             =   540
            Width           =   510
         End
      End
      Begin VB.OptionButton MemoriaFiscalDatas 
         Caption         =   "Intervalo de Datas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   870
         Width           =   2760
      End
      Begin VB.Frame FrameDatas 
         Caption         =   "Intervalo de Datas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   375
         TabIndex        =   15
         Top             =   1230
         Width           =   4380
         Begin MSComCtl2.UpDown UpDownDataDe 
            Height          =   435
            Left            =   1935
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   480
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataDe 
            Height          =   420
            Left            =   585
            TabIndex        =   8
            Top             =   495
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataAte 
            Height          =   435
            Left            =   4080
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   180
            _ExtentX        =   450
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataAte 
            Height          =   420
            Left            =   2790
            TabIndex        =   9
            Top             =   480
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   741
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataAte 
            AutoSize        =   -1  'True
            Caption         =   "Até:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2265
            TabIndex        =   19
            Top             =   540
            Width           =   510
         End
         Begin VB.Label LabelDataDe 
            AutoSize        =   -1  'True
            Caption         =   "De:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   18
            Top             =   525
            Width           =   435
         End
      End
   End
   Begin VB.Frame FrameSessao 
      Caption         =   "Sessão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   90
      TabIndex        =   13
      Top             =   1560
      Width           =   4830
      Begin VB.CommandButton BotaoSessaoInicia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1110
         Picture         =   "OperacoesECF.ctx":E048
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   375
         Width           =   1935
      End
      Begin VB.CommandButton BotaoSessaoEncerra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1110
         Picture         =   "OperacoesECF.ctx":1276A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2175
         Width           =   1935
      End
      Begin VB.CommandButton BotaoSessaoSuspende 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1110
         Picture         =   "OperacoesECF.ctx":16364
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1260
         Width           =   1935
      End
   End
   Begin VB.Frame FrameCaixa 
      Caption         =   "Caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4815
      Begin VB.CommandButton BotaoCaixaAbertura 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         Picture         =   "OperacoesECF.ctx":1A222
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   450
         Width           =   1935
      End
      Begin VB.CommandButton BotaoCaixaFechamento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2340
         Picture         =   "OperacoesECF.ctx":1D744
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   450
         Width           =   1935
      End
   End
End
Attribute VB_Name = "OperacoesECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'mario

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Operações de Caixa"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "OperacoesECF"

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


Private Sub BotaoDAVArquivo_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_BotaoDAVRelGer_Click

        
    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDeDAVArquivo.ClipText)) = 0 Or Len(Trim(DataAteDAVArquivo.ClipText)) = 0 Then gError 204378
    
    dtDataDe = DataDeDAVArquivo.Text
    dtDataAte = DataAteDAVArquivo.Text

    If dtDataDe > dtDataAte Then gError 204381

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("DAVEmitidos_Grava", dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError 204379
    
    'Limpa a Tela
    Call Limpa_Tela_Operacoes
    
    Exit Sub
    
Erro_BotaoDAVRelGer_Click:

    
    Select Case gErr

        Case 204378
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_NAO_PREENCHIDAS, gErr)

        Case 204379
            'Erro Tratado Dentro da Função Chamada
            
        Case 204381
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204380)

    End Select

    Exit Sub

End Sub

Private Sub DataAteDAVArquivo_GotFocus()
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAteDAVArquivo)
End Sub

Private Sub DataAteDAVArquivo_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAteDAVArquivo_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataAteDAVArquivo.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataAteDAVArquivo.Text)
    If lErro <> SUCESSO Then gError 204391

    Exit Sub

Erro_DataAteDAVArquivo_Validate:

    Cancel = True

    Select Case gErr

        Case 204391
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204392)

    End Select

    Exit Sub

End Sub

Private Sub DataAteDAVRelGer_GotFocus()
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAteDAVRelGer)

End Sub

Private Sub DataAteDAVRelGer_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAteDAVRelGer_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataAteDAVRelGer.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataAteDAVRelGer.Text)
    If lErro <> SUCESSO Then gError 204389

    Exit Sub

Erro_DataAteDAVRelGer_Validate:

    Cancel = True

    Select Case gErr

        Case 204389
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204390)

    End Select

    Exit Sub

End Sub


Private Sub DataDeDAVArquivo_GotFocus()
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDeDAVArquivo)

End Sub

Private Sub DataDeDAVArquivo_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDeDAVArquivo_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataDeDAVArquivo.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataDeDAVArquivo.Text)
    If lErro <> SUCESSO Then gError 204385

    Exit Sub

Erro_DataDeDAVArquivo_Validate:

    Cancel = True

    Select Case gErr

        Case 204385
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204386)

    End Select

    Exit Sub

End Sub

Private Sub DataDeDavRelGer_GotFocus()
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDeDavRelGer)

End Sub

Private Sub DataDeDAVRelGer_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDeDAVRelGer_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataDeDavRelGer.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataDeDavRelGer.Text)
    If lErro <> SUCESSO Then gError 204387

    Exit Sub

Erro_DataDeDAVRelGer_Validate:

    Cancel = True

    Select Case gErr

        Case 204387
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 204388)

    End Select

    Exit Sub

End Sub



'Incluído por Luiz Nogueira em 24/03/04
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyF2
            Call BotaoCaixaAbertura_Click
        Case vbKeyF3
            Call BotaoCaixaFechamento_Click
        Case vbKeyF4
            
            Call BotaoCaixaLeituraX_Click
        Case vbKeyF5
'            Call BotaoCaixaReducaoZ_Click
        
        Case vbKeyF7
            Call BotaoLeituraMemoriaFiscal_Click
        Case vbKeyF8
            Call Sair_Click
            
        Case vbKeyF9
            Call BotaoSessaoInicia_Click
        Case vbKeyF10
            Call BotaoSessaoSuspende_Click
        Case vbKeyF11
            Call BotaoSessaoEncerra_Click
    
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

Public Sub Form_Unload(Cancel As Integer)

End Sub

Public Sub Form_Load()
'Inicialização

    
    
    'Chama a Funçãop que Limpa a Tela
    Call Limpa_Tela_Operacoes

    'Deixa os Frames desabilitados
    FrameDatas.Enabled = False
    FrameReducoes.Enabled = False

    If AFRAC_ImpressoraCFe(giCodModeloECF) Then FrameCaixa.Enabled = False

    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

End Sub

Sub Limpa_Tela_Operacoes()
'Função que Limpa a Tela no Caso de marcada a Opção Completa Completa

    'Limpa a tela
    Call Limpa_Tela(Me)

    Completa.Value = True
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objOperador As ClassOperador) As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub BotaoCaixaAbertura_Click()
'Botão de Abertura de Caixa
Dim lErro As Long
Dim sStatusSessao As String

On Error GoTo Erro_BotaoCaixaAbertura_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207982

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111317
        
        
        'Verifica se a Caxa esta Aberto se ja Estive Erro
        If giStatusCaixa = STATUS_CAIXA_ABERTO Then gError 107538
    
        'Se caixa não estiver Aberto então realizar a Abertura de Caixa
        lErro = CF_ECF("Caixa_Executa_Abertura")
        If lErro <> SUCESSO Then gError 107539
    
'???????********NESTE PONTO A SESSAO TEM QUE ESTAR ENCERRADA SENAO É ERRO ******* Mario. 25/11/02
   
    
        lErro = AFRAC_AbrirDia(Date)
        lErro = CF_ECF("Retorna_MSGErro_AFRAC", lErro, "Abertura do Dia")
        If lErro <> SUCESSO Then gError 112054
    
        Call Rotina_AvisoECF(vbOKOnly, AVISO_CAIXA_ABERTA)
    
    Else
    
        Call CF_ECF("Trata_Caixa_So_Orcamento")
    
    End If
    
    Exit Sub

Erro_BotaoCaixaAbertura_Click:

    Select Case gErr

        Case 107538
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_ABERTO, gErr, giCodCaixa)

        Case 107539, 112054, 207982
                'Erro Tratado Dentro da Função que foi Chamada

        Case 111317
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163647)

    End Select

    Exit Sub

End Sub

Private Sub BotaoCaixaLeituraX_Click()
'Rotina que Executa a LeituraX


    Call CF_ECF("Executa_LeituraX")

End Sub

'Private Sub BotaoCaixaReducaoZ_Click()
''Botão que Inicia a ReduçãoZ
'
'Dim lErro As Long
'Dim vbMsgRes As VbMsgBoxResult
'
'On Error GoTo Erro_BotaoCaixaReducaoZ_Click
'
'    lErro = CF_ECF("Requisito_XXII")
'    If lErro <> SUCESSO Then gError 207984
'
'
'    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105911
'
'    'Verifica se já foi executa a redução z para a data de hoje
'    If gdtUltimaReducao = gdtDataHoje Then gError 111320
'
'    If giStatusCaixa = STATUS_CAIXA_ABERTO Then gError 126488
'
'    'Função Que Prepara para Executar a Redução Z
'    lErro = CF_ECF("Caixa_Executa_ReducaoZ")
'    If lErro <> SUCESSO Then gError 107569
'
'    Exit Sub
'
'Erro_BotaoCaixaReducaoZ_Click:
'
'     Select Case gErr
'
'        Case 105911
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)
'
'        Case 107569, 207984
'            'Erro Tratado Dentro da Função que Foi Chamada
'
'        Case 111320
'            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(gdtDataHoje, "dd/mm/yyyy"))
'
'        Case 126488
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_ABERTO, gErr, giCodCaixa)
'
'        Case Else
'            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163649)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub Completa_Click()
'Executa a Leitura da Memória Fiscal Completa sem Intertícios de datas

Dim lErro As Long

On Error GoTo Erro_Completa_Click

    'Limpa Tuda a Tela
    Call Limpa_Tela_Operacoes

    'Desabilita os Frames
    FrameDatas.Enabled = False
    FrameReducoes.Enabled = False

    Exit Sub

Erro_Completa_Click:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163650)

    End Select

    Exit Sub

End Sub



Private Sub MemoriaFiscalDatas_Click()
'Função que Habilita o Frame Data

Dim lErro As Long

On Error GoTo Erro_MemoriaFiscalDatas_Click

    'Habilita o Frame Data
    FrameDatas.Enabled = True
    
    'Desabilita o Frame de Reduções
    FrameReducoes.Enabled = False
   
   Exit Sub

Erro_MemoriaFiscalDatas_Click:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163651)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_GotFocus()
'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataAte_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataAte)

    Exit Sub

Erro_DataAte_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163652)

    End Select

    Exit Sub

End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataAte_Validate

    'Verifica se Data Até esta Preenchida se não sai do Validate
    If Len(Trim(DataAte.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataAte.Text)
    If lErro <> SUCESSO Then gError 107583

    Exit Sub

Erro_DataAte_Validate:

    Cancel = True

    Select Case gErr

        Case 107583
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163653)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    'Trata A entrada em algum Campo

Dim lErro As Long

On Error GoTo Erro_DataDe_GotFocus
    
    'Função que Coloca o Cursor da Data no Inicio do Campo
    Call MaskEdBox_TrataGotFocus(DataDe)

    Exit Sub

Erro_DataDe_GotFocus:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163654)

    End Select

    Exit Sub


End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Valida os Dados do Campo de Data

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_DataDe_Validate

    'Verifica se Data De esta Preenchida se não sai do Validate
    If Len(Trim(DataDe.ClipText)) = 0 Then Exit Sub

    'Função que Serve para Verificar se a Data é Valida
    lErro = Data_Critica(DataDe.Text)
    If lErro <> SUCESSO Then gError 107582

    Exit Sub

Erro_DataDe_Validate:

    Cancel = True

    Select Case gErr

        Case 107582
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163655)

    End Select

    Exit Sub

End Sub

Private Sub MemoriaFiscalReducoes_Click()
'Função que Habilita o Frame Reduções

Dim lErro As Long

On Error GoTo Erro_MemoriaFiscalDatas_Click

    'Habilita o Frame Data
    FrameReducoes.Enabled = True
    'Desabilita o Frame de Datas
    FrameDatas.Enabled = False
    
    Exit Sub

Erro_MemoriaFiscalDatas_Click:

     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163656)

    End Select

    Exit Sub

End Sub

Private Sub Sair_Click()

    Unload Me

End Sub

Private Sub UpDownDataDe_DownClick()
'Função que serve para decrementar a Data
Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_DownClick

    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 107584

    Exit Sub

Erro_UpDownDataDe_DownClick:

    Select Case gErr

        Case 107584
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163657)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataDe_UpClick()
'Função que serve para imcrementar a Data
Dim lErro As Long

On Error GoTo Erro_UpDownDataDe_UpClick

    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 107585

    Exit Sub

Erro_UpDownDataDe_UpClick:

    Select Case gErr

        Case 107585
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163658)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_DownClick()
'Função que serve para decrementar a Data
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_DownClick

    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 107586

    Exit Sub

Erro_UpDownDataAte_DownClick:

    Select Case gErr

        Case 107586
            'Erro Tratado Dentro da Função Chamadora

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163659)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAte_UpClick()
'Função que serve para imcrementar a Data
Dim lErro As Long

On Error GoTo Erro_UpDownDataAte_UpClick

    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 107587

    Exit Sub

Erro_UpDownDataAte_UpClick:

    Select Case gErr

        Case 107587
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163660)

    End Select

    Exit Sub

End Sub

Private Sub ReducaoDe_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de Redução

Dim lErro As Long

On Error GoTo Erro_ReducaoDe_Validate

    'Verifica se o Intervalo de redução não está preenchido sai do validate
    If Len(Trim(ReducaoDe.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(ReducaoDe.Text)
    If lErro <> SUCESSO Then gError 107591

    Exit Sub

Erro_ReducaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 107591
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163661)

    End Select

    Exit Sub

End Sub

Private Sub ReducaoAte_Validate(Cancel As Boolean)
'Valida os Dados no Intervalo de Redução

Dim lErro As Long

On Error GoTo Erro_ReducaoAte_Validate

    'Verifica se o Intervalo de redução não está preenchido sai do validate
    If Len(Trim(ReducaoAte.Text)) = 0 Then Exit Sub

    'Função que valida se Intervalode Redução é Positivo
    lErro = Valor_Positivo_Critica(ReducaoAte.Text)
    If lErro <> SUCESSO Then gError 107592

    Exit Sub

Erro_ReducaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 107592
            'Erro Tratado Dentro da Função Chamada

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163662)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLeituraMemoriaFiscal_Click()
'Chama a Função que Vai Chama a Função da Afrac que Vai Executar a Leitura da Memoria Fiscal

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim sDe As String
Dim sAte As String
Dim iTipo As Integer
Dim iArquivo As Integer

On Error GoTo Erro_BotaoLeituraMemoriaFiscal_Click

    If giOrcamentoECF = CAIXA_SO_ORCAMENTO Then gError 105912

    iArquivo = 0

    If Completa Then
        iTipoLeitura = LEITURA_COMPLETA
        iTipo = LEITURA_COMPLETA
    
    ElseIf MemoriaFiscalDatas Then
        
        iTipoLeitura = LEITURA_DATAS
        iTipo = LEITURA_COMPLETA
        
        'Verificar se as Datas Estão Preenchidas se Erro
        If Len(Trim(DataDe.ClipText)) = 0 Or Len(Trim(DataAte.ClipText)) = 0 Then gError 107614
        
        If Len(Trim(DataDe.ClipText)) > 0 Then sDe = DataDe.Text
        If Len(Trim(DataAte.ClipText)) > 0 Then sAte = DataAte.Text

        If CDate(sDe) > CDate(sAte) Then gError 204383

    Else
        iTipoLeitura = LEITURA_REDUCOES
        iTipo = LEITURA_COMPLETA
        
        sDe = ReducaoDe.Text
        sAte = ReducaoAte.Text
        
    End If

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("MemoriaFiscal_Executa_Leitura", iTipoLeitura, sDe, sAte, iTipo, iArquivo)
    If lErro <> SUCESSO Then gError 107621
    
    'Limpa a Tela
    Call Limpa_Tela_Operacoes
    
    Exit Sub
    
Erro_BotaoLeituraMemoriaFiscal_Click:

    Select Case gErr

        Case 105912
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_SO_ORCAMENTO, gErr)

        Case 107614
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_MEMORIAFISCAL_NAO_PREENCHIDA, gErr)

        Case 107621
            'Erro Tratado Dentro da Função Chamada

        Case 204383
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163663)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoSessaoInicia_Click()

'Função que inicia a Sessão
Dim objOperador As New ClassOperador

Dim lErro As Long

On Error GoTo Erro_BotaoSessaoInicia_Click
    
    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207985

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then
    
        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111321
        
    End If
        
    'Função que Faz a Abertura de Sessão
    lErro = CF_ECF("Sessao_Executa_Abertura")
    If lErro <> SUCESSO Then gError 107554

    Exit Sub

Erro_BotaoSessaoInicia_Click:

    Select Case gErr

        Case 107554, 207985

        Case 111321
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163664)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSessaoSuspende_Click()
'Função que Suspende a Sessão

Dim objOperador As New ClassOperador
Dim iCogGerente As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoSessaoSuspende_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207986


    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111322
        
    End If
        
    'Se a Sessão Estiver Fechada então gera Erro
    If giStatusSessao = SESSAO_ENCERRADA Then gError 107605

    'Se Sessão estiver Suspensa
    If giStatusSessao = SESSAO_SUSPENSA Then gError 107606

    'Função que Executa a Suspenção da Sessão
    lErro = CF_ECF("Sessao_Executa_Suspensao")
    If lErro <> SUCESSO Then gError 107607

    'funcao que executa o termino da suspensao se a senha for digitada.
    lErro = CF_ECF("Sessao_Executa_Termino_Susp")
    If lErro <> SUCESSO Then gError 117542

    Exit Sub

Erro_BotaoSessaoSuspende_Click:

    Select Case gErr

        Case 107605
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 107606
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_SUSPENSA, gErr, giCodCaixa)

        Case 107607, 117542, 207986

        Case 111322
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163665)

    End Select

    Exit Sub

End Sub

Private Sub BotaoSessaoEncerra_Click()
'Botão que Termina a Sessão do Dia de Trabalho
Dim lErro As Long
Dim sStatusCaixa As String

On Error GoTo Erro_BotaoSessaoEncerra_Click

    lErro = CF_ECF("Requisito_XXII")
    If lErro <> SUCESSO Then gError 207987

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then

        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111323
        
    End If
        
    'Verificar se a Sessão não esta Aberta se Estiver Fechado Então Erro
    If giStatusSessao = SESSAO_ENCERRADA Then gError 107567

    'Função que Faz o Encerramento da Sessão
    lErro = CF_ECF("Operacoes_Sessao_Executa_Encerramento")
    If lErro <> SUCESSO Then gError 107568
    
    Exit Sub

Erro_BotaoSessaoEncerra_Click:

    Select Case gErr

        Case 107567
            Call Rotina_ErroECF(vbOKOnly, ERRO_SESSAO_ABERTA_INEXISTENTE, gErr, giCodCaixa)

        Case 107568, 207987
            'Erro Tratado Dentro da Função Chamadora

        Case 111323
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163666)

    End Select

    Exit Sub

End Sub


Private Sub BotaoCaixaFechamento_Click()
'Função Que Inicia o Fechamento do Caixa

Dim lErro As Long
Dim sRet As String
Dim sReg As String
Dim vbMsg As VbMsgBoxResult
Dim bOk As Boolean
Dim objAdmMeioPagto As ClassAdmMeioPagto

On Error GoTo Erro_BotaoCaixaFechamento_Click

    If giOrcamentoECF <> CAIXA_SO_ORCAMENTO Then


        lErro = CF_ECF("Requisito_XXII")
        If lErro <> SUCESSO Then gError 207983



        'Verifica se já foi executa a redução z para a data de hoje
        If gdtUltimaReducao = Date Then gError 111318
    
    
        'Verifica se o Caixa esta Fechado
        If giStatusCaixa = STATUS_CAIXA_FECHADO Then gError 107573
            
        'Função que vai Executar o Fechamento do Caixa
        lErro = CF_ECF("Caixa_Executa_Fechamento")
        If lErro <> SUCESSO Then gError 107574
    
    Else
    
        Call CF_ECF("Trata_Caixa_So_Orcamento")
    
    End If
        
    Exit Sub

Erro_BotaoCaixaFechamento_Click:

    Select Case gErr

        Case 107573
            Call Rotina_ErroECF(vbOKOnly, ERRO_CAIXA_FECHADO, gErr, giCodCaixa)

        Case 107574, 112055, 207983
            'Erro Tratado Dentro dsa Função Chamadora

        Case 111318
            Call Rotina_ErroECF(vbOKOnly, ERRO_REDUCAO_JA_EXECUTADA, gErr, Format(Date, "dd/mm/yyyy"))
                        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163667)

    End Select

    Exit Sub

End Sub

Function Desmembra_MovimentosCaixa(objMovimentosCaixa As ClassMovimentoCaixa) As Long
'Função que Desmembra o Movimentos de Caixa e Carrega em gcolMovimentosCaixa

Dim lErro As Long
Dim iPosInicio As Integer
Dim sRegistro As String
Dim iPos As Integer
Dim iPosFinal As Integer
Dim iPosicao3  As Integer
Dim iPosShift As Integer
Dim iIndice As Integer
Dim sNomeArq As String
Dim sTipo As String
Dim iPosEnd As Integer
Dim iCont As Integer
Dim iPosInicioShift As Integer
Dim lConteudo As Long
Dim iInicial As Integer

On Error GoTo Erro_Desmembra_MovimentosCaixa
    
    'Instancia o Obj da ClassMovimentoCaixa
    Set objMovimentosCaixa = New ClassMovimentoCaixa
    
    'Desmembra o Arquivo do dia Corrente
    sNomeArq = gsDirMVTEF & "MV" & CStr(Format(Date, "ddmmyy")) & (".txt")
    
    Open sNomeArq For Input As #1
    
    Do While Not EOF(1)
         
        iInicial = 1
             
        'Primeira Posição
        iPosInicio = 1
        
        'Inicializa a variavel
        iIndice = 0
       
        'Busca o próximo registro do arquivo
        Line Input #1, sRegistro
        
        'Posição Final
        iPosEnd = InStr(iPosInicio, sRegistro, Chr(vbKeyEnd))
        
        'Procura o Primeiro Control para saber o tipo do registro
        iPos = InStr(iPosInicio, sRegistro, Chr(vbKeyControl))
        
        'String para verificar se o tipo de registro é do tipo MovimentoCaixa
        sTipo = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
        
        'Verifica a Posição do Shifth
        iPosShift = InStr(iInicial, sRegistro, Chr(vbKeyShift))
        
        'Se for então
        If sTipo = TIPOREGISTROECF_MOVIMENTOCAIXA Then
        
            Do While iPosInicio < (iPosEnd - 1)
        
                iIndice = iIndice + 1
                
                'acerta os Ponteiros
                If iIndice = 1 Then
                
                    iPosInicio = iPos + 1
                    iPos = InStr(iPos + 1, sRegistro, Chr(vbKeyControl))
                    
                End If
                    
                'Recolhe os Dados do arquivo de movimentosCaixa e Coloca no obj
                
                Select Case iIndice
                    
                    Case 1: objMovimentosCaixa.iTipo = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 2: objMovimentosCaixa.dHora = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 3: objMovimentosCaixa.dtDataMovimento = StrParaDate(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 4: objMovimentosCaixa.iGerente = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 5: objMovimentosCaixa.iCodOperador = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 6: objMovimentosCaixa.lSequencial = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 7: objMovimentosCaixa.iFilialEmpresa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 8: objMovimentosCaixa.lTransferencia = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 9: objMovimentosCaixa.dValor = StrParaDbl(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 10: objMovimentosCaixa.lNumMovto = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 11: objMovimentosCaixa.iAdmMeioPagto = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 12: objMovimentosCaixa.iParcelamento = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 13: objMovimentosCaixa.iExcluiu = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 14: objMovimentosCaixa.iCaixa = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 15: objMovimentosCaixa.iTipoCartao = StrParaInt(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 16: objMovimentosCaixa.lCupomFiscal = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 17: objMovimentosCaixa.lMovtoEstorno = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 18: objMovimentosCaixa.lMovtoTransf = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 19: objMovimentosCaixa.lSequencialConta = StrParaLong(Mid(sRegistro, iPosInicio, iPos - iPosInicio))
                    Case 20: objMovimentosCaixa.sFavorecido = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 21: objMovimentosCaixa.sHistorico = Mid(sRegistro, iPosInicio, iPos - iPosInicio)
                    Case 22: Exit Do
            
                End Select
                
                'Atualiza as Posições
                iPosInicio = iPos + 1
                iPos = (InStr(iPosInicio, sRegistro, Chr(vbKeyControl)))
        
            Loop
            
        End If
    Loop
    
    'Fecha o Arquivo
    Close #1
    
    Desmembra_MovimentosCaixa = SUCESSO
       
    Exit Function

Erro_Desmembra_MovimentosCaixa:
    
    Desmembra_MovimentosCaixa = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163668)

    End Select
    
    'Fecha o Arquivo em caso de Erro
    Close #1

    Exit Function

End Function



Private Sub Desmembrar_Click()
'ESSA FUNÇÃO VAI SER APAGADA DA TELA ELA _
SÓ EXISTE PARA TESTAR A FUNÇÃO QUE DESMEMBRA _
AS INFORMAÇÕES GRAVADAS EM ARQUIVO

Dim objMovimentosCaixa As New ClassMovimentoCaixa
Dim lErro As Long
Dim sStatusSessao As String
Dim sStatusCaixa As String

On Error GoTo Erro_Desmembrar_Click

    lErro = Desmembra_MovimentosCaixa(objMovimentosCaixa)
    If lErro <> SUCESSO Then gError 109806
    
    Select Case objMovimentosCaixa.iTipo
    
        Case MOVIMENTO_CAIXA_ABERTURA, MOVIMENTO_CAIXA_LEITURA_X
        
            gcolMovimentosCaixa.Add objMovimentosCaixa
            'Verifica o Status da Sessão
            If giStatusSessao = SESSAO_ABERTA Then
                sStatusSessao = "Sessão Aberta"
            ElseIf giStatusSessao = SESSAO_SUSPENSA Then
                sStatusSessao = "Sessão Suspensa"
            ElseIf giStatusSessao = SESSAO_ENCERRADA Then
                sStatusSessao = "Sessão Encerrada"
            End If
            
            'Alterar o Caption do ECF
            GL_objMDIForm.Caption = "CAIXA_STATUS : Caixa Aberto / SESSAO_STATUS : " & sStatusSessao
            
            'Altera o Status do Caixa, Instancia a Varialvel Global
            giStatusCaixa = STATUS_CAIXA_ABERTO
    
        Case MOVIMENTO_CAIXA_FECHAMENTO
        
            'Indica o Nome do Operador dessa Sessão.
            GL_objMDIForm.Caption = "CAIXA_STATUS : Caixa Fechado"
            'Caixa Muda o Status
            giStatusCaixa = STATUS_CAIXA_FECHADO
        
        Case MOVIMENTO_CAIXA_REDUCAO_Z
            'Atualizar na Variável Global a Data da Ultima Redução Z
            gdtUltimaReducao = objMovimentosCaixa.dtDataMovimento
            
        Case MOVIMENTO_CAIXA_SESSAO_ABERTURA
            
            'Verifica o Status da Sessão
            If giStatusCaixa = 0 Then
                sStatusCaixa = "Caixa Fechada"
            ElseIf giStatusCaixa = 1 Then
                sStatusCaixa = "Sessão Aberta"
            End If
        
            'Indica o Nome do Operador dessa Sessão.
            GL_objMDIForm.Caption = "CAIXA_STATUS : " & sStatusCaixa & " / SESSAO_STATUS :  Sessão Aberta " & ""
            
        Case MOVIMENTO_CAIXA_SESSAO_FECHAMENTO
            
            'Indica o Novo status da Sessão Suspensa
            giStatusSessao = SESSAO_SUSPENSA
            
            'Verifica o Status da Sessão
            If giStatusCaixa = 0 Then
                sStatusCaixa = "Caixa Fechado"
            ElseIf giStatusCaixa = 1 Then
                sStatusCaixa = "Caixa Aberto"
            End If
        
            'Verifica o Status da Sessão
            If giStatusSessao = SESSAO_ABERTA Then
                sStatusSessao = "Sessão Aberta"
            ElseIf giStatusSessao = SESSAO_SUSPENSA Then
                sStatusSessao = "Sessão Suspensa"
            ElseIf giStatusSessao = SESSAO_ENCERRADA Then
                sStatusSessao = "Sessão Encerrada"
            End If
            
            'Muda a Caption do ECF para indicar que a Sessão está Suspensa
            GL_objMDIForm.Caption = "CAIXA_STATUS : " & sStatusCaixa & " / SESSAO_STATUS : " & sStatusSessao
            
    End Select
        
    Exit Sub

Erro_Desmembrar_Click:
    
     Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163669)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoDAVRelGer_Click()

Dim lErro As Long
Dim iTipoLeitura As Integer
Dim dtDataDe As Date
Dim dtDataAte As Date

On Error GoTo Erro_BotaoDAVRelGer_Click

        
    'Verificar se as Datas Estão Preenchidas se Erro
    If Len(Trim(DataDeDavRelGer.ClipText)) = 0 Or Len(Trim(DataAteDAVRelGer.ClipText)) = 0 Then gError 204350
    
    dtDataDe = DataDeDavRelGer.Text
    dtDataAte = DataAteDAVRelGer.Text

    If dtDataDe > dtDataAte Then gError 204382

    'Função que Vai Chamar Função da Afrac que Vai Executar a Leitura da Memoria Fiscal
    lErro = CF_ECF("DAV_Executa_RelGer", dtDataDe, dtDataAte)
    If lErro <> SUCESSO Then gError 204377
    
    'Limpa a Tela
    Call Limpa_Tela_Operacoes
    
    Exit Sub
    
Erro_BotaoDAVRelGer_Click:

    
    Select Case gErr

        Case 204350
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATAS_NAO_PREENCHIDAS, gErr)

        Case 204377
            'Erro Tratado Dentro da Função Chamada

        Case 204382
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_INICIAL_MAIOR1, gErr)

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 163663)

    End Select

    Exit Sub
    
End Sub

