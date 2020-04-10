VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelacionamentoContatosOcx 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8325
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   8325
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4500
      Index           =   1
      Left            =   480
      TabIndex        =   39
      Top             =   960
      Width           =   7455
      Begin VB.CheckBox FixarDados 
         Caption         =   "Fixar dados para próximo registro"
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
         Left            =   3840
         TabIndex        =   22
         ToolTipText     =   $"RelacionamentoContatos.ctx":0000
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Frame FramePrincipal 
         Caption         =   "Principal"
         Height          =   4450
         Left            =   120
         TabIndex        =   40
         Top             =   0
         Width           =   7095
         Begin VB.Frame FrameContato 
            Caption         =   "Contato"
            Height          =   1695
            Left            =   240
            TabIndex        =   43
            Top             =   240
            Width           =   6615
            Begin VB.ComboBox Origem 
               Height          =   315
               ItemData        =   "RelacionamentoContatos.ctx":00DA
               Left            =   4080
               List            =   "RelacionamentoContatos.ctx":00E4
               Style           =   2  'Dropdown List
               TabIndex        =   4
               ToolTipText     =   "Selecione quem originou o relacionamento: o seu cliente ou a sua empresa."
               Top             =   240
               Width           =   1215
            End
            Begin VB.ComboBox Tipo 
               Height          =   315
               ItemData        =   "RelacionamentoContatos.ctx":00FA
               Left            =   1200
               List            =   "RelacionamentoContatos.ctx":00FC
               TabIndex        =   11
               Text            =   "Tipo"
               ToolTipText     =   "Selecione o tipo de relacionamento com o cliente. Para cadastrar novos tipos, use a tela Campos Genéricos."
               Top             =   1245
               Width           =   4095
            End
            Begin VB.CommandButton BotaoProxNum 
               Height          =   285
               Left            =   2280
               Picture         =   "RelacionamentoContatos.ctx":00FE
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Pressione esse botão para gerar um código automático para o relacionamento."
               Top             =   255
               Width           =   300
            End
            Begin MSComCtl2.UpDown UpDownData 
               Height          =   300
               Left            =   2160
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   765
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox Data 
               Height          =   300
               Left            =   1200
               TabIndex        =   6
               ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
               Top             =   765
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Codigo 
               Height          =   315
               Left            =   1155
               TabIndex        =   1
               ToolTipText     =   "Informe o código do relacionamento."
               Top             =   240
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Mask            =   "999999999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox Hora 
               Height          =   315
               Left            =   4080
               TabIndex        =   9
               Top             =   765
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "hh:mm:ss"
               Mask            =   "##:##:##"
               PromptChar      =   " "
            End
            Begin VB.Label LabelHora 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Left            =   3540
               TabIndex        =   8
               Top             =   825
               Width           =   480
            End
            Begin VB.Label LabelOrigem 
               AutoSize        =   -1  'True
               Caption         =   "Origem:"
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
               Left            =   3360
               TabIndex        =   3
               Top             =   300
               Width           =   660
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
               Left            =   600
               TabIndex        =   10
               Top             =   1305
               Width           =   450
            End
            Begin VB.Label LabelData 
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
               TabIndex        =   5
               Top             =   825
               Width           =   480
            End
            Begin VB.Label LabelCodigo 
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
               Height          =   255
               Left            =   480
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   0
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Frame FrameAtendente 
            Caption         =   "Atendente"
            Height          =   855
            Left            =   240
            TabIndex        =   42
            Top             =   3480
            Width           =   3255
            Begin VB.ComboBox Atendente 
               Height          =   315
               Left            =   1250
               TabIndex        =   21
               ToolTipText     =   "Digite o código, o nome do atendente ou aperte F3 para consulta. Para cadastrar novos tipos, use a tela Campos Genéricos."
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label LabelAtendente 
               AutoSize        =   -1  'True
               Caption         =   "Atendente:"
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
               Left            =   240
               TabIndex        =   20
               Top             =   420
               Width           =   945
            End
         End
         Begin VB.Frame FrameCliente 
            Caption         =   "Cliente"
            Height          =   1335
            Left            =   240
            TabIndex        =   41
            Top             =   2040
            Width           =   6615
            Begin VB.ComboBox FilialCliente 
               Height          =   315
               Left            =   4920
               TabIndex        =   15
               ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
               Top             =   300
               Width           =   1380
            End
            Begin VB.ComboBox Contato 
               Height          =   315
               Left            =   1380
               TabIndex        =   17
               ToolTipText     =   $"RelacionamentoContatos.ctx":01E8
               Top             =   825
               Width           =   2175
            End
            Begin VB.TextBox Cliente 
               Height          =   315
               Left            =   1380
               TabIndex        =   13
               ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
               Top             =   300
               Width           =   2175
            End
            Begin MSMask.MaskEdBox Telefone 
               Height          =   315
               Left            =   4920
               TabIndex        =   19
               ToolTipText     =   "Informe o código do relacionamento."
               Top             =   840
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               PromptChar      =   " "
            End
            Begin VB.Label LabelTelefone 
               AutoSize        =   -1  'True
               Caption         =   "Telefone:"
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
               Left            =   3960
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   18
               Top             =   900
               Width           =   825
            End
            Begin VB.Label LabelContato 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   630
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   16
               Top             =   900
               Width           =   735
            End
            Begin VB.Label LabelCliente 
               AutoSize        =   -1  'True
               Caption         =   "Cliente Futuro:"
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
               TabIndex        =   12
               Top             =   360
               Width           =   1260
            End
            Begin VB.Label LabelFilialCliente 
               AutoSize        =   -1  'True
               Caption         =   "Filial:"
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
               Left            =   4320
               TabIndex        =   14
               Top             =   360
               Width           =   465
            End
         End
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   2
      Left            =   480
      TabIndex        =   36
      Top             =   960
      Width           =   7335
      Begin VB.Frame FrameAssunto 
         Caption         =   "Assunto"
         Height          =   4450
         Left            =   120
         TabIndex        =   37
         Top             =   0
         Width           =   7095
         Begin VB.CheckBox Encerrado 
            Caption         =   "Encerrado"
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
            Left            =   360
            TabIndex        =   35
            Top             =   4080
            Width           =   1215
         End
         Begin VB.TextBox Assunto 
            Height          =   1815
            Left            =   360
            MaxLength       =   510
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   2160
            Width           =   6255
         End
         Begin VB.Frame FrameContatoAnterior 
            Caption         =   "Contato Anterior"
            Height          =   1620
            Left            =   360
            TabIndex        =   38
            Top             =   240
            Width           =   6255
            Begin MSMask.MaskEdBox RelacionamentoAnt 
               Height          =   315
               Left            =   1080
               TabIndex        =   56
               ToolTipText     =   "Informe o código do relacionamento."
               Top             =   240
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   12
               Mask            =   "999999999999"
               PromptChar      =   " "
            End
            Begin VB.Label TipoContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1080
               TabIndex        =   32
               Top             =   1200
               Width           =   2655
            End
            Begin VB.Label HoraContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4320
               TabIndex        =   30
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label DataContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1080
               TabIndex        =   28
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label OrigemContatoAnt 
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   4320
               TabIndex        =   26
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label LabelTipoContatoAnt 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   480
               TabIndex        =   31
               Top             =   1260
               Width           =   450
            End
            Begin VB.Label LabelOrigemContatoAnt 
               AutoSize        =   -1  'True
               Caption         =   "Origem:"
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
               Left            =   3600
               TabIndex        =   25
               Top             =   300
               Width           =   660
            End
            Begin VB.Label LabelDataContatoAnt 
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
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   480
               TabIndex        =   27
               Top             =   780
               Width           =   480
            End
            Begin VB.Label LabelHoraContatoAnt 
               AutoSize        =   -1  'True
               Caption         =   "Hora:"
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
               Left            =   3780
               TabIndex        =   29
               Top             =   780
               Width           =   480
            End
            Begin VB.Label LabelCodContatoAnt 
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
               Left            =   360
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   24
               Top             =   300
               Width           =   660
            End
         End
         Begin VB.Label LabelAssunto 
            AutoSize        =   -1  'True
            Caption         =   "Assunto:"
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
            TabIndex        =   33
            Top             =   1950
            Width           =   750
         End
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4500
      Index           =   3
      Left            =   480
      TabIndex        =   44
      Top             =   960
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame FrameComplemento 
         Caption         =   "Complemento"
         Height          =   4455
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   7095
         Begin VB.TextBox CampoValor 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   2160
            TabIndex        =   48
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox Campo 
            BorderStyle     =   0  'None
            Height          =   315
            Left            =   480
            TabIndex        =   47
            Top             =   1320
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3495
            Left            =   240
            TabIndex        =   46
            Top             =   480
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   6165
            _Version        =   393216
         End
      End
   End
   Begin VB.CheckBox ImprimeGravacao 
      Caption         =   "Imprimir ao gravar"
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
      Left            =   240
      TabIndex        =   55
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   5400
      ScaleHeight     =   450
      ScaleWidth      =   2685
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   120
      Width           =   2745
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   120
         Picture         =   "RelacionamentoContatos.ctx":0270
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   630
         Picture         =   "RelacionamentoContatos.ctx":0372
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1140
         Picture         =   "RelacionamentoContatos.ctx":04CC
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1650
         Picture         =   "RelacionamentoContatos.ctx":0656
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "RelacionamentoContatos.ctx":0B88
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4950
      Left            =   240
      TabIndex        =   23
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8731
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Principal"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Assunto"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "RelacionamentoContatosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'??? implementar criação de registro em crfatconfig ao inserir nova filial para relacionamentoclientes e para atendentes

'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoContato As AdmEvento
Attribute objEventoContato.VB_VarHelpID = -1
Private WithEvents objEventoTelefone As AdmEvento
Attribute objEventoTelefone.VB_VarHelpID = -1
Private WithEvents objEventoAtendente As AdmEvento
Attribute objEventoAtendente.VB_VarHelpID = -1
Private WithEvents objEventoRelacionamentoAnt As AdmEvento
Attribute objEventoRelacionamentoAnt.VB_VarHelpID = -1

Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iTelefoneAlterado As Integer
Dim iFilialCliAlterada As Integer

Dim giFrameAtual As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    
    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoAtendente = New AdmEvento
    Set objEventoTelefone = New AdmEvento
    Set objEventoRelacionamentoAnt = New AdmEvento
    
    'Carrega a combo Tipo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    If lErro <> SUCESSO Then gError 102499
    
    'Carrega a combo AtendenteAte
    lErro = CF("Carrega_Atendentes", Atendente)
    If lErro <> SUCESSO Then gError 102523
    
    'Coloca data atual como padrão
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Coloca origem empresa como padrão
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 102499
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166585)
    
    End Select
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Function

Public Function Trata_Parametros(Optional ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um relacionamento
    If Not (objRelacionamentoClientes Is Nothing) Then
    
        'Se o código do relacionamento está preenchido,
        'significa que é uma consulta de um relacionamento gravado
        If objRelacionamentoClientes.lCodigo > 0 Then
        
            'Lê e traz os dados do relacionamento para a tela
            lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoClientes)
            If lErro <> SUCESSO Then gError 102500
        
        'Senão,
        'significa que é a criação de um novo contato com dados recebidos pelo obj
        Else
        
            'Apenas traz para a tela os dados do relacionamento
            'Isso acontece quando o usuário utiliza um relacionamento já cadastrado
            'para gerar um novo relacionamento
            lErro = Traz_RelacionamentoClientes_Tela1(objRelacionamentoClientes)
            If lErro <> SUCESSO Then gError 102501
            
            'Cria automaticamente o código para o contato
            Call BotaoProxNum_Click
        
        End If
    
    End If
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 102500, 102501
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166586)
    
    End Select
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoCliente = Nothing
    Set objEventoContato = Nothing
    Set objEventoTelefone = Nothing
    Set objEventoAtendente = Nothing
    Set objEventoRelacionamentoAnt = Nothing

    Call ComandoSeta_Liberar(Me.Name)
    
End Sub
'*** FECHAMENTO DA TELA - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - INÍCIO****

'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - INÍCIO ***
Private Sub Codigo_GotFocus()
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
End Sub

Private Sub RelacionamentoAnt_GotFocus()
    Call MaskEdBox_TrataGotFocus(RelacionamentoAnt, iAlterado)
End Sub
'*** EVENTO GOTFOCUS DOS CONTROLES MASCARADOS - FIM ***

'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***
Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código do relacionamento não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102978
    
    'Dispara função para imprimir relacionamento
    lErro = RelacionamentoClientes_Imprime(StrParaInt(Codigo.Text))
    If lErro <> SUCESSO Then gError 102979
    
    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 102979
        
        Case 102978
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166587)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 102529

    'Limpa a Tela
    Call Limpa_RelacionamentoCliente1

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 102529

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 166588)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objRelacionamentoClientes As New ClassRelacClientes
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 102633

    'Guarda no obj, código do relacionamento e filial empresa
    'Essas informações são necessárias para excluir o relacionamento
    objRelacionamentoClientes.lCodigo = StrParaLong(Codigo.Text)
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa

    'Lê o relacionamento com os filtros passados
    lErro = CF("RelacionamentoContatos_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 181031 Then gError 102634
    
    'Se não encontrou => erro
    If lErro = 181031 Then gError 102635
    
    'Se o relacionamento está com status encerrado, a msg de confirmação
    'deve explicitar esse detalhe
    If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then
        sAviso = "AVISO_CONFIRMA_EXCLUSAO_RELACIONAMENTOCLIENTES1"
    Else
        sAviso = "AVISO_CONFIRMA_EXCLUSAO_RELACIONAMENTOCLIENTES"
    End If
    
    'Pede a confirmação da exclusão do relacionamento com cliente
    vbMsgRes = Rotina_Aviso(vbYesNo, sAviso, objRelacionamentoClientes.lCodigo)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If

    'Faz a exclusão do Orcamento de Venda
    lErro = CF("RelacionamentoContatos_Exclui", objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102636

    'Limpa a Tela de Orcamento de Venda
    Call Limpa_RelacionamentoCliente
    
    'fecha o comando de setas
    Call ComandoSeta_Fechar(Me.Name)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 102633
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102634, 102636

        Case 102635
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 166589)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 102527

    'Limpa a Tela
    Call Limpa_RelacionamentoCliente
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 102527

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166590)

    End Select

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Obtém o próximo código de relacionamento para giFilialEmpresa
    lErro = CF("Config_ObterAutomatico", "CRFATConfig", "NUM_PROX_RELACIONAMENTOCONTATOS", "RelacionamentoContatos", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 102509
    
    'Exibe o código obtido
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 102509
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166591)

    End Select

End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        FrameTab(TabStrip1.SelectedItem.Index).Visible = True
        FrameTab(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166592)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 102526

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 102526

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166593)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 102525

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 102525

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166594)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection

    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    Call Chama_Tela("RelacionamentoContatos_Lista", colSelecao, objRelacionamentoCli, objEventoCodigo)
    
End Sub

Private Sub LabelCliente_Click()

Dim objContato As New ClassContatos
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objContato.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objContato.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ContatosLista", colSelecao, objContato, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166595)
    
    End Select
    
End Sub

Private Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Click()

Dim lErro As Long
Dim objFilialContato As New ClassFilialContato

On Error GoTo Erro_FilialCliente_Click

    'Se nenhuma filial foi selecionada => sai da função
    If FilialCliente.ListIndex = -1 Then Exit Sub
    
    objFilialContato.iCodFilial = Codigo_Extrai(FilialCliente.Text)
    
    'Lê a filial e obtém o telefone e os contatos da mesma
    lErro = Obtem_Contatos_FilialCliente(objFilialContato)
    If lErro <> SUCESSO Then gError 102631
    
    Exit Sub
    
Erro_FilialCliente_Click:

    Select Case gErr
    
        Case 102631
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166596)
    
    End Select

End Sub

Private Sub Contato_Click()

Dim lErro As Long
Dim lCliente As Long
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Contato_Click

    'Se o campo contato não foi preenchido => sai da função
    If Contato.ListIndex = -1 Then Exit Sub
    
    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102628

    'Guarda o código do cliente e da filial no obj
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objClienteContatos.iCodigo = Codigo_Extrai(Contato.Text)

    'Lê o contato no BD
    lErro = CF("ClienteFContatos_Le", objClienteContatos)
    If lErro <> SUCESSO And lErro <> 181075 Then gError 102655
    
    'Se não encontrou o contato => erro
    If lErro = 181075 Then gError 102687
    
    'Exibe o telefone cadastrado para o contato selecionado
    Telefone.Text = objClienteContatos.sTelefone
    
    iTelefoneAlterado = 0
    
    Exit Sub
    
Erro_Contato_Click:

    Select Case gErr

        Case 102628, 102655
        
        Case 102687
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166597)

    End Select
    

End Sub

Private Sub Atendente_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelCodContatoAnt_Click()

Dim objRelacionamentoCli As New ClassRelacClientes
Dim colSelecao As New Collection
Dim sSelecao As String

On Error GoTo Erro_LabelCodContatoAnt_Click

    'Se o cliente não foi preenchido => erro, pois não é possível exibir uma lista
    'de relacionamentos anteriores sem saber qual o cliente do relacionamento atual
    If Len(Trim(Cliente.Text)) = 0 Then gError 102704
    
    'Se a data não foi preenchida => erro, pois não é possível exibir uma lista
    'de relacionamentos anteriores sem saber qual a data do relacionamento atual
    If Len(Trim(Data.ClipText)) = 0 Then gError 102705
    
    'Passa para o obj o código do relacionamento, onde o registro deve tentar se posicionar
    objRelacionamentoCli.lCodigo = StrParaDbl(Codigo.Text)
    
    'Filtra os registro no browser, pois um relacionamento anterior obrigatoriamente
    'tem que pertencer ao cliente do relacionamento atual e tem que ter data menor que a ]
    'data atual
    sSelecao = "ClienteNomeReduzido=? AND Data<=? AND CodRelacionamento<>?"
    
    'Passa os valores para os filtros acima
    colSelecao.Add Trim(Cliente.Text)
    colSelecao.Add StrParaDate(Data.Text)
    colSelecao.Add StrParaDbl(Codigo.Text)
    
    'Chama o browser
    Call Chama_Tela("RelacionamentoContatos_Lista", colSelecao, objRelacionamentoCli, objEventoRelacionamentoAnt, sSelecao)
    
    Exit Sub

Erro_LabelCodContatoAnt_Click:

    Select Case gErr
    
        Case 102704
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_RELAC_ATUAL_NAO_PREENCHIDO", gErr, Error)
            
        Case 102705
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_RELAC_ATUAL_NAO_PREENCHIDO", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166598)
        
    End Select
    
End Sub

Private Sub Encerrado_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CLICK DOS CONTROLES - INÍCIO ***

'*** EVENTO CHANGE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Origem_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Hora_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Tipo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
End Sub
Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
End Sub
Private Sub Contato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Atendente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Telefone_Change()
    iAlterado = REGISTRO_ALTERADO
    iTelefoneAlterado = REGISTRO_ALTERADO
End Sub
Private Sub RelacionamentoAnt_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub Assunto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
'*** EVENTO CHANGE DOS CONTROLES - FIM ***

'*** EVENTO VALIDATE DOS CONTROLES - INÍCIO ***
Private Sub Codigo_Validate(Cancel As Boolean)

On Error GoTo Erro_Codigo_Validate

    'Se o código do relacionamento atual foi preenchido
    'e o código do relacionamento anterior também
    'e forem iguais => limpa o código de relacionamento anterior, pois ele não é válido
    If (StrParaDbl(Codigo.Text) > 0) And (StrParaDbl(RelacionamentoAnt.Text) > 0) And (StrParaDbl(Codigo.Text) = StrParaDbl(RelacionamentoAnt.Text)) Then Call Limpa_Frame_RelacionamentoAnt
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166599)
    
    End Select
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lCliente As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Data_Validate

    'Se a data não foi preenchida => sai da função
    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 102510

    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102661
    
    'Guarda no obj os dados necessários para validar o código do relacionamento anterior
    objRelacionamentoClientes.lCliente = lCliente
    objRelacionamentoClientes.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objRelacionamentoClientes.dtData = StrParaDate(Data.Text)
    
    'Verifica se o código do relacionamento anterior é válido
    'para o cliente/filial em questão
    lErro = Trata_RelacionamentoAnterior(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102662
    
    Exit Sub
    
Erro_Data_Validate:

    Cancel = True

    Select Case gErr
    
        Case 102510, 102661, 102662
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166600)
        
    End Select

End Sub

Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora de saida foi digitada
    If Len(Trim(Hora.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(Hora.Text)
    If lErro <> SUCESSO Then gError 102511

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 102511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166601)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Valida o tipo de relacionamento selecionado pelo cliente
    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo, "AVISO_CRIAR_TIPORELACIONAMENTOCLIENTES")
    If lErro <> SUCESSO Then gError 102512
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102512
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166602)

    End Select

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    'Faz a validação do cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102674
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102674
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166603)

    End Select

End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    'Faz a validação da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102680
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 102680
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166604)

    End Select

End Sub

Private Sub Contato_Validate(Cancel As Boolean)
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objClienteContatos As New ClassClienteContatos
Dim iCodigo As Integer
Dim lCliente As Long

On Error GoTo Erro_Contato_Validate

    'Se o contato foi preenchido
    If Len(Trim(Contato.Text)) > 0 Then
    
        'Se o contato foi selecionado na própria combo => sai da função
        If Contato.Text = Contato.List(Contato.ListIndex) Then Exit Sub
        
        'Se o cliente não foi preenchido => erro
        If Len(Trim(Cliente.Text)) = 0 Then gError 102682
        
        'Se a filial do cliente não foi preenchido => erro
        If Len(Trim(FilialCliente.Text)) = 0 Then gError 102683
    
        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(Contato, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102684
    
        'Se não encontrou o contato na combo, mas retornou um código
        If lErro = 6730 Then
        
            'Obtém o código do cliente
            lErro = Obtem_CodCliente(lCliente)
            If lErro <> SUCESSO Then gError 102686
            
            'Guarda os dados necessários para tentar ler o contato
            objClienteContatos.lCliente = lCliente
            objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
            objClienteContatos.iCodigo = iCodigo
            
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteFContatos_Le", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 181075 Then gError 102681
            
            'Se não encontrou o contato
            If lErro = 181075 Then gError 102685
            
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Text = objClienteContatos.sTelefone
        
        End If
        
        'Se foi digitado o nome do contato
        'e esse nome não foi encontrado na combo => erro
        If lErro = 6731 Then
        
            'Obtém o código do cliente
            lErro = Obtem_CodCliente(lCliente)
            If lErro <> SUCESSO Then gError 102686
            
            'Guarda os dados necessários para tentar ler o contato
            objClienteContatos.lCliente = lCliente
            objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
            objClienteContatos.sContato = Contato.Text
        
            'Lê o contato a partir dos dados passados
            lErro = CF("ClienteFContatos_Le_Nome", objClienteContatos)
            If lErro <> SUCESSO And lErro <> 178440 Then gError 178442
            
            'Se não encontrou o contato
            If lErro = 178440 Then gError 102687
        
            'Exibe o contato na tela
            Contato.Text = objClienteContatos.iCodigo & SEPARADOR & objClienteContatos.sContato
            
            'Exibe o telefone do contato
            Telefone.Text = objClienteContatos.sTelefone
        
        End If
    'Senão
    Else
    
        'Limpa o campo telefone
        Telefone.Text = ""
    
    End If
    
    iTelefoneAlterado = 0
    
    Exit Sub

Erro_Contato_Validate:

    Cancel = True

    Select Case gErr

        Case 102681, 102684, 102686
        
        Case 102682
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102683
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 102685
            
            'Verifica se o usuário deseja criar um novo contato
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CLIENTECONTATO", Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))

            'Se o usuário respondeu sim
            If vbMsgRes = vbYes Then
                'Chama a tela para cadastro de contatos
                Call Chama_Tela("ClienteFContatos", objClienteContatos)
            End If
        
        Case 102687
'            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO", gErr, Trim(Contato.Text), Trim(Cliente.Text), Trim(FilialCliente.Text))
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166605)

    End Select

    iTelefoneAlterado = 0
    
    Exit Sub

End Sub

Private Sub Telefone_Validate(Cancel As Boolean)

Dim lErro As Long
Dim colClienteContatos As New Collection
Dim colSelecao As New Collection
Dim sSelecao As String
Dim sTelefone As String
Dim objClienteContatos As New ClassClienteContatos
Dim lCliente As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer

On Error GoTo Erro_Telefone_Validate

    'Se o telefone não foi alterado => sai da função
    If iTelefoneAlterado = 0 Then Exit Sub
    
    'Se o telefone foi preenchido
    If Len(Trim(Telefone.Text)) > 0 Then
    
        'Guarda o telefone que deve ser usado para pesquisa
        sTelefone = Format(Telefone.Text, "####-####")
        
        'Pesquisa contas de clientes pelo número de telefone
        lErro = CF("ClienteFContatos_Le_Telefone", sTelefone, colClienteContatos)
        If lErro <> SUCESSO And lErro <> 181080 Then gError 102672
        
        'Se não encontrou => erro
        If lErro = 181080 Then gError 102673
        
        'Se encontrou mais de 1 cliente com o mesmo telefone
        If colClienteContatos.Count > 1 Then
        
            'Monta uma seleção que garanta que o browser só exibirá os
            'contatos com o mesmo telefone
            sSelecao = "ContatoTelefone=?"
            colSelecao.Add sTelefone
    
            'Chama a tela de consulta de cliente
            Call Chama_Tela("ClienteContatos_Lista", colSelecao, objClienteContatos, objEventoTelefone, sSelecao)
        
        Else
        
            'Joga na tela o cliente pertecente ao contato encontrado
            Cliente.Text = colClienteContatos(1).lCliente
            lErro = Valida_Cliente()
            If lErro <> SUCESSO Then gError 102677
            
            'Joga na tela a filial do cliente pertencente ao contato encontrado
            FilialCliente.Text = colClienteContatos(1).iFilialCliente
            lErro = Valida_FilialCliente()
            If lErro <> SUCESSO Then gError 102678
            
            'Joga na tela o contato ao qual pertence o telefone pesquisado
            Contato.Text = colClienteContatos(1).iCodigo & SEPARADOR & colClienteContatos(1).sContato
            'Call Contato_Validate(bSGECancelDummy)
        
            For iIndice = 0 To Contato.ListCount - 1
                If Contato.List(iIndice) = colClienteContatos(1).iCodigo & SEPARADOR & colClienteContatos(1).sContato Then
                    Contato.ListIndex = iIndice
                    Exit For
                End If
        
            Next
            
        End If
    
    'senão foi preenchido
    Else
    
        'limpa o campo contato
        Contato.Text = ""
    
    End If
    
    Exit Sub
    
Erro_Telefone_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102672, 102677
        
        Case 102673
            If Len(Trim(Cliente.Text)) = 0 Or Len(Trim(FilialCliente.Text)) = 0 Then
                Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTECONTATO_NAO_ENCONTRADO1", gErr, sTelefone)
            Else
                'Verifica se o usuário deseja cadastrar este telefone
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CADASTRAR_TELEFONECONTATO", Trim(Telefone.Text))
    
                'Se o usuário respondeu sim
                If vbMsgRes = vbYes Then
                    
                    'Obtém o código do cliente
                    Call Obtem_CodCliente(lCliente)
                    
                    'Guarda os dados necessários para tentar ler o contato
                    objClienteContatos.lCliente = lCliente
                    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
                    
                    'Chama a tela para cadastro de contatos
                    Call Chama_Tela("ClienteFContatos", objClienteContatos)
                End If
            End If
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166606)

    End Select
    
End Sub

Public Sub Atendente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Atendente_Validate

    'Valida o atendente selecionado pelo cliente
    lErro = CF("Atendente_Validate", Atendente)
    If lErro <> SUCESSO Then gError 102524
    
    Exit Sub

Erro_Atendente_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102524
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166607)

    End Select

End Sub

Private Sub RelacionamentoAnt_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes
Dim objContato As New ClassContatos
Dim objCamposGenericosValores As New ClassCamposGenericosValores

On Error GoTo Erro_RelacionamentoAnt_Validate

    'Se o campo está preenchido
    If StrParaDbl(RelacionamentoAnt.Text) > 0 Then
    
        'Se o usuário digitou como relacionamento anterior
        'o mesmo código desse relacionamento => erro
        If StrParaLong(RelacionamentoAnt.Text) = StrParaLong(Codigo.Text) Then gError 102691
        
        'Guarda no obj código e filialempresa onde
        objRelacionamentoClientes.lCodigo = StrParaLong(RelacionamentoAnt.Text)
        objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
        
        'Lê o relacionamento com os filtros passados
        lErro = CF("RelacionamentoContatos_Le", objRelacionamentoClientes)
        If lErro <> SUCESSO And lErro <> 181031 Then gError 102517
        
        'Se não encontrou o relacionamento => erro
        If lErro = 181031 Then gError 102518
        
        'Guarda em objContato o nome reduzido do cliente
        objContato.sNomeReduzido = Trim(Cliente.Text)
                
        'Lê o cliente a partir do nome reduzido
        'O objetivo dessa leitura é obter o código do cliente para compará-lo com
        'o código do cliente do relacionamento anterior
        lErro = CF("Contato_Le_NomeReduzido", objContato)
        If lErro <> SUCESSO And lErro <> 180745 Then gError 102519
        
        'Se não encontrou o cliente => erro
        If lErro = 180745 Then gError 102520
        
        'Se o cliente do relacionamento anterior não é o mesmo cliente
        'do relacionamento atual => erro
        If objRelacionamentoClientes.lCliente <> objContato.lCodigo Then gError 102521
        
        'Se a data do relacionamento anterior é maior do que a data do relacionamento atual => erro
        If objRelacionamentoClientes.dtData > Data.Text Then gError 102522
        
        '*** EXIBE OS DADOS DO RELACIONAMENTO ANTERIOR ***
        'Exibe os dados do relacionamento anterior
        'Origem
        'Se o relacionamento foi originado por cliente
        If objRelacionamentoClientes.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
            OrigemContatoAnt.Caption = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO
        'Senão
        Else
            OrigemContatoAnt.Caption = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO
        End If

        'Data
        DataContatoAnt.Caption = objRelacionamentoClientes.dtData
        
        'Hora
        'Se a hora foi gravada no BD => exibe-a na tela
        If objRelacionamentoClientes.dtHora <> 0 Then HoraContatoAnt.Caption = Format(objRelacionamentoClientes.dtHora, "hh:mm:ss")

        'LEITURA DO TIPO DE RELACIONAMENTO
        'Guarda no obj os dados necessários para ler o tipo de relacionamento
        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES
        objCamposGenericosValores.lCodValor = objRelacionamentoClientes.lTipo
        
        'Lê o tipo de contato para obter a descrição
        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
        If lErro <> SUCESSO And lErro <> 102399 Then gError 102659
        
        'Se não encontrou => erro
        If lErro = 102399 Then gError 102660
        
        'Exibe na tela o tipo do contato anterior
        TipoContatoAnt.Caption = objCamposGenericosValores.lCodValor & SEPARADOR & objCamposGenericosValores.sValor
        '*************************************************
        
    End If
        
    Exit Sub
    
Erro_RelacionamentoAnt_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 102517, 102519, 102659
        
        Case 102691
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_INVALIDO", gErr, Trim(RelacionamentoAnt.Text))
            
        Case 102518
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case 102520
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objContato.lCodigo)
            
        Case 102521
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_CLIENTE_DIFERENTE", gErr, objRelacionamentoClientes.lCodigo, objContato.sNomeReduzido)
                    
        Case 102522
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTOANT_DATA_INVALIDA", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.dtData)
        
        Case 102660
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPORELACIONAMENTOCLI_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lTipo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166608)
    
    End Select
    
End Sub

'*** EVENTO VALIDATE DOS CONTROLES - FIM ***

'*** TRATAMENTO DOS CONTROLES DA TELA - FIM ****

'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        ElseIf Me.ActiveControl Is RelacionamentoAnt Then
            Call LabelCodContatoAnt_Click
        ElseIf Me.ActiveControl Is Contato Then
            Call LabelContato_Click
        ElseIf Me.ActiveControl Is Telefone Then
            Call LabelTelefone_Click
        End If
    
    End If

End Sub
'*** TRATAMENTO DO EVENTO KEYDOWN  - FIM ***

'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***
Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As New ClassRelacClientes
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objRelacionamentoCli = obj1
    
    'Traz para a tela o relacionamento com código passado pelo browser
    lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102528
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 102528
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166609)
    
    End Select

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objContato As ClassContatos
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objContato = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objContato.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102675

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 102675
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166610)
    
    End Select

End Sub

Private Sub objEventoTelefone_evSelecao(obj1 As Object)

Dim objClienteContatos As ClassClienteContatos
Dim bCancel As Boolean

    Set objClienteContatos = obj1
    
    'Preenche o cliente
    Cliente.Text = objClienteContatos.lCliente
    Call Valida_Cliente
    
    'preenche a filial do cliente
    FilialCliente.Text = objClienteContatos.iFilialCliente
    Call Valida_FilialCliente
    
    Contato.Text = objClienteContatos.iCodigo
    
    Call Contato_Validate(bCancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoAtendente_evSelecao(obj1 As Object)

Dim objCamposGenericosValores As ClassCamposGenericosValores
Dim bCancel As Boolean

    Set objCamposGenericosValores = obj1
    
    Atendente.Text = objCamposGenericosValores.lCodValor
    
    Call Atendente_Validate(bCancel)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoRelacionamentoAnt_evSelecao(obj1 As Object)

Dim objRelacionamentoCli As ClassRelacClientes

    Set objRelacionamentoCli = obj1
    
    RelacionamentoAnt.Text = objRelacionamentoCli.lCodigo
    
    Call RelacionamentoAnt_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub
'*** TRATAMENTO DOS EVENTOS DE BROWSER - INÍCIO ***

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objRelacionamentoClientes As New ClassRelacClientes
Dim objCampoValor As AdmCampoValor
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "RelacionamentoContatos_Consulta"

    'Guarda no obj os dados que serão usados para identifica o registro a ser exibido
    objRelacionamentoClientes.lCodigo = StrParaDbl(Trim(Codigo.Text))
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 102649
    
    objRelacionamentoClientes.lCliente = lCliente
    objRelacionamentoClientes.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    objRelacionamentoClientes.dtData = StrParaDate(Data.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "CodRelacionamento", objRelacionamentoClientes.lCodigo, 0, "CodRelacionamento"
    colCampoValor.Add "FilialRelacionamento", objRelacionamentoClientes.iFilialEmpresa, 0, "FilialRelacionamento"
    colCampoValor.Add "ClienteNomeReduzido", Trim(Cliente.Text), STRING_CLIENTE_NOME_REDUZIDO, "ClienteNomeReduzido"
    colCampoValor.Add "CodFilialCliente", objRelacionamentoClientes.iFilialCliente, 0, "CodFilialCliente"
    colCampoValor.Add "Data", objRelacionamentoClientes.dtData, 0, "Data"
    
    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case 102649
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166611)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Tela_Preenche

    'Guarda o código do campo em questão no obj
    objRelacionamentoClientes.lCodigo = colCampoValor.Item("CodRelacionamento").vValor
    objRelacionamentoClientes.iFilialEmpresa = colCampoValor.Item("FilialRelacionamento").vValor

    'Preenche a tela com os valores para o campo em questão
    lErro = Traz_RelacionamentoClientes_Tela(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102650
    
    iAlterado = 0
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 102650
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166612)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub
'**** FIM DO TRATAMENTO DO SISTEMA DE SETAS ****

'*** FUNÇÕES DE APOIO À TELA - INÍCIO ***
Private Function Traz_RelacionamentoClientes_Tela(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'Traz pra tela os dados do relacionamento passado como parâmetro
'objRelacionamentoClientes RECEBE(Input) os dados que servirão para identificar o relacionamento a ser trazido para a tela

Dim lErro As Long

On Error GoTo Erro_Traz_RelacionamentoClientes_Tela

    'Limpa a tela
    Call Limpa_RelacionamentoCliente
    
    'Lê no BD os dados do relacionamento a ser lido
    lErro = CF("RelacionamentoContatos_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 181031 Then gError 102502
    
    'Se não encontrou o relacionamento => erro
    If lErro = 181031 Then gError 102503
    
    'Chama a função que traz para a tela os dados lidos
    lErro = Traz_RelacionamentoClientes_Tela1(objRelacionamentoClientes)
    If lErro <> SUCESSO Then gError 102504

    Traz_RelacionamentoClientes_Tela = SUCESSO

    Exit Function

Erro_Traz_RelacionamentoClientes_Tela:

    Traz_RelacionamentoClientes_Tela = gErr

    Select Case gErr

        Case 102502, 102504
        
        Case 102503
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166613)

    End Select

End Function

Private Function Traz_RelacionamentoClientes_Tela1(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long
'objRelacionamentoClientes RECEBE(Input) os dados que devem ser exibidos na tela

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_RelacionamentoClientes_Tela1

    'Exibe os dados do obj na tela
    
    'Código
    Codigo.PromptInclude = False
    Codigo.Text = objRelacionamentoClientes.lCodigo
    Codigo.PromptInclude = True
    
    'Origem
    'Se o relacionamento foi originado por cliente
    If objRelacionamentoClientes.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE Then
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE - 1
    'Senão
    Else
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    End If
    
    'Data
    'se a data foi preenchida
    If objRelacionamentoClientes.dtData > 0 Then
        Data.PromptInclude = False
        Data.Text = Format(objRelacionamentoClientes.dtData, "dd/mm/yy")
        Data.PromptInclude = True
    End If
    
    'Hora
    'Se a hora foi gravada no BD
    If objRelacionamentoClientes.dtHora <> 0 Then
        Hora.PromptInclude = False
        Hora.Text = Format(objRelacionamentoClientes.dtHora, "hh:mm:ss")
        Hora.PromptInclude = True
    End If
    
    'Tipo
    For iIndice = 1 To Tipo.ListCount
        
        If objRelacionamentoClientes.lTipo = Tipo.ItemData(iIndice - 1) Then
            Tipo.ListIndex = iIndice - 1
            Exit For
        End If
    Next
    
    
    'Cliente
    Cliente.Text = objRelacionamentoClientes.lCliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 102676
    
    'FilialCliente
    FilialCliente.Text = objRelacionamentoClientes.iFilialCliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 102679
    
    'Contato
    For iIndice = 1 To Contato.ListCount
        
        If objRelacionamentoClientes.iContato = Contato.ItemData(iIndice - 1) Then
            Contato.ListIndex = iIndice - 1
            Exit For
        End If
    Next
    
    'Atendente
    'Se o atendente foi informado
    If objRelacionamentoClientes.iAtendente > 0 Then
        Atendente.Text = objRelacionamentoClientes.iAtendente
        Call Atendente_Validate(bSGECancelDummy)
    End If
    
    'Se o código do relacionamento anterior foi preenchido
    If objRelacionamentoClientes.lRelacionamentoAnt > 0 Then
        'RelacionamentoAnterior
        RelacionamentoAnt.Text = objRelacionamentoClientes.lRelacionamentoAnt
        Call RelacionamentoAnt_Validate(bSGECancelDummy)
    End If
    
    'Assunto
    Assunto.Text = objRelacionamentoClientes.sAssunto1 & objRelacionamentoClientes.sAssunto2
    
    'Status
    If objRelacionamentoClientes.iStatus = RELACIONAMENTOCLIENTES_STATUS_ENCERRADO Then Encerrado.Value = vbChecked
    
    Traz_RelacionamentoClientes_Tela1 = SUCESSO

    Exit Function

Erro_Traz_RelacionamentoClientes_Tela1:

    Traz_RelacionamentoClientes_Tela1 = gErr

    Select Case gErr

        Case 102676, 102679
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166614)

    End Select

End Function

Private Sub Limpa_RelacionamentoCliente()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'Coloca data atual como padrão
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    'Limpa a origem
    Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
    
    'Limpa a combo tipo
    Tipo.ListIndex = -1
    
    'Limpa a combo filial
    FilialCliente.Clear
    
    'Limpa a combo contatos
    Contato.Clear
    
    'Limpa a combo de atendentes
    Atendente.ListIndex = -1
    
    'Seleciona o atendente padrão. Atendente padrão é o atendente vinculado ao usuário ativo
    'Para cada atendente da combo AtendenteDe
    For iIndice = 0 To Atendente.ListCount - 1
    
        'Se o conteúdo do atendente for igual ao seu código + "-" + nome reduzido do usuário ativo
        If Atendente.List(iIndice) = Atendente.ItemData(iIndice) & SEPARADOR & gsUsuario Then
        
            'Significa que achou o atendente "default"
            'Seleciona o atendente na combo
            Atendente.ListIndex = iIndice
            
            'Sai do For
            Exit For
        End If
    Next
    
    'Recarrega a combo Tipo e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)
    
    'Limpa o frame Relacionamento Anterior
    Call Limpa_Frame_RelacionamentoAnt
    
    'Desmarca a opção 'encerrado'
    Encerrado.Value = vbUnchecked
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Sub

Private Sub Limpa_RelacionamentoCliente1()

    'Se não é para manter os dados do cliente
    If FixarDados.Value = vbUnchecked Then
    
        'Limpa toda a tela
        Call Limpa_RelacionamentoCliente
    
    'Senão
    Else
        
        'Limpa todos os controles, exceto os controles que envolvem cliente e atendente
        Codigo.PromptInclude = False
        Codigo.Text = ""
        Codigo.PromptInclude = True
        
        Origem.ListIndex = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA - 1
        
        Data.PromptInclude = False
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
        Data.PromptInclude = True
        
        Hora.PromptInclude = False
        Hora.Text = ""
        Hora.PromptInclude = True
        
        'Recarrega a combo Tipo e seleciona a opção padrão
        'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
        Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPORELACIONAMENTOCLIENTES, Tipo)

        RelacionamentoAnt.Text = ""
        Assunto.Text = ""
        Encerrado.Value = vbUnchecked
    
    End If
    
    iAlterado = 0
    iClienteAlterado = 0
    iFilialCliAlterada = 0
    iTelefoneAlterado = 0
    
End Sub

Private Sub Limpa_Frame_RelacionamentoAnt()

RelacionamentoAnt.Text = ""
OrigemContatoAnt.Caption = ""
DataContatoAnt.Caption = ""
HoraContatoAnt.Caption = ""
TipoContatoAnt.Caption = ""

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objRelacionamentoCli As New ClassRelacClientes

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos obrigatórios estão preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 102530

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_RelacionamentoClientes_Memoria(objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102531

    'Verifica se esse relacionamento já existe no BD
    'e, em caso positivo, alerta ao usuário que está sendo feita uma alteração
    lErro = Trata_Alteracao(objRelacionamentoCli, objRelacionamentoCli.lCodigo, objRelacionamentoCli.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 102656
    
    'Grava no BD
    lErro = CF("RelacionamentoContatos_Grava", objRelacionamentoCli)
    If lErro <> SUCESSO Then gError 102532

    'Se for para imprimir o relacionamento depois da gravação
    If ImprimeGravacao.Value = vbChecked Then

        'Dispara função para imprimir orçamento
        lErro = RelacionamentoClientes_Imprime(objRelacionamentoCli.lCodigo)
        If lErro <> SUCESSO Then gError 102533

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 102530, 102531, 102532, 102656
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166615)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela são válidos para a gravação do registro

Dim lErro As Long

On Error GoTo Erro_Valida_Gravacao

    'Se o código não estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 102534
    
    'Se a origem não estiver preenchida => erro
    If Len(Trim(Origem.Text)) = 0 Then gError 102535
    
    'Se a data não estiver preenchida => erro
    If Len(Trim(Data.ClipText)) = 0 Then gError 102536
    
    'Se o tipo não estiver preenchido => erro
    If Len(Trim(Tipo.Text)) = 0 Then gError 102537
    
    'Se o cliente não estiver preenchido => erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 102538
    
    'Se a filial do cliente não estiver preenchida => erro
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 102539
    
    'Se o atendente não estiver preenchido => erro
    If Len(Trim(Atendente.Text)) = 0 Then gError 102540

    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 102534
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 102535
            Call Rotina_Erro(vbOKOnly, "ERRO_ORIGEMRELACCLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102536
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 102537
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_RELACCLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 102538
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 102539
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 102540
            Call Rotina_Erro(vbOKOnly, "ERRO_ATENDENTE_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166616)

    End Select

End Function

Private Function Move_RelacionamentoClientes_Memoria(ByVal objRelacionamentoCli As ClassRelacClientes) As Long
'Guarda os dados da tela na memória
'objRelacionamentoCli devolve os dados da tela

Dim lErro As Long
Dim objContato As New ClassContatos

On Error GoTo Erro_Move_RelacionamentoClientes_Memoria

    'Guarda o código do relacionamento
    objRelacionamentoCli.lCodigo = StrParaDbl(Trim(Codigo.Text))
    
    'Guarda a filial empresa do relacionamento
    objRelacionamentoCli.iFilialEmpresa = giFilialEmpresa
    
    'Se o relacionamento foi originado pelo cliente
    If Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE_TEXTO Then
    
        'Indica que é um relacionamento originado por cliente
        objRelacionamentoCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_CLIENTE
    
    'Se o relacionamento foi originado pela empresa
    ElseIf Origem.Text = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA_TEXTO Then
    
        'Indica que é um relacionamento originado pela empresa
        objRelacionamentoCli.iOrigem = RELACIONAMENTOCLIENTES_ORIGEM_EMPRESA
    
    End If
    
    'Guarda a data no obj
    objRelacionamentoCli.dtData = MaskedParaDate(Data)
    
    'Se é um relacionamento com a data atual e a hora não foi preenchida
    If CDate(Data.Text) = gdtDataHoje And Len(Trim(Hora.ClipText)) = 0 Then
        
        'Guarda no obj a hora atual
        objRelacionamentoCli.dtHora = Time
    
    'Senão, verifica se a hora está preenchida
    ElseIf Len(Trim(Hora.ClipText)) > 0 Then
    
        'Guarda no obj a hora informada pelo usuário
        objRelacionamentoCli.dtHora = StrParaDate(Hora.Text)
    
    End If
    
    'Guarda no obj, o tipo do relacionamento
    objRelacionamentoCli.lTipo = LCodigo_Extrai(Tipo.Text)
        
    '*** Leitura do cliente a partir do nome reduzido para obter o seu código ***
    
    'Guarda o nome reduzido do cliente
    objContato.sNomeReduzido = Trim(Cliente.Text)
    
    'Faz a leitura do cliente
    lErro = CF("Contato_Le_NomeReduzido", objContato)
    If lErro <> SUCESSO And lErro <> 180745 Then gError 102543
    
    'Se não encontrou o cliente => erro
    If lErro = 180745 Then gError 102544
    
    'Guarda no obj o código do cliente
    objRelacionamentoCli.lCliente = objContato.lCodigo
    
    '*** Fim da leitura de cliente ***
    
    'Guarda no obj o código da filial do cliente
    objRelacionamentoCli.iFilialCliente = Codigo_Extrai(FilialCliente.Text)
    
    'Guarda no obj o código do contato
    objRelacionamentoCli.iContato = Codigo_Extrai(Contato.Text)
    
    'Guarda no obj, o atendente do relacionamento
    objRelacionamentoCli.iAtendente = LCodigo_Extrai(Atendente.Text)
    
    'Guarda o código do relacionamento anterior
    objRelacionamentoCli.lRelacionamentoAnt = StrParaDbl(Trim(RelacionamentoAnt.Text))
    
    'Guarda no obj a primeira parte do assunto
    objRelacionamentoCli.sAssunto1 = Left(Assunto.Text, STRING_BUFFER_MAX_TEXTO - 1)
    
    'Guarda no obj a segunda parte do assunto
    objRelacionamentoCli.sAssunto2 = Mid(Assunto.Text, STRING_BUFFER_MAX_TEXTO)
    
    'Guarda no obj, o status do relacionamento
    objRelacionamentoCli.iStatus = Encerrado.Value
        
    Move_RelacionamentoClientes_Memoria = SUCESSO

    Exit Function

Erro_Move_RelacionamentoClientes_Memoria:

    Move_RelacionamentoClientes_Memoria = gErr

    Select Case gErr

        Case 102543
        
        Case 102544
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objContato.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166617)

    End Select

End Function

Private Function Obtem_CodCliente(lCliente As Long) As Long
'Obtém o código do cliente e da filial que estão na tela e guarda-os no objClienteContatos

Dim lErro As Long
Dim objContato As New ClassContatos

On Error GoTo Erro_Obtem_CodCliente

    'Se o cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
    
        '*** Leitura do cliente a partir do nome reduzido para obter o seu código ***
        
        'Guarda o nome reduzido do cliente
        objContato.sNomeReduzido = Trim(Cliente.Text)
        
        'Faz a leitura do cliente
        lErro = CF("Contato_Le_NomeReduzido", objContato)
        If lErro <> SUCESSO And lErro <> 180745 Then gError 102618
        
        'Se não encontrou o cliente => erro
        If lErro = 180745 Then gError 102619
        
        'Devolve o código do cliente
        lCliente = objContato.lCodigo
        
        '*** Fim da leitura de cliente ***
        
    End If

    Obtem_CodCliente = SUCESSO

    Exit Function

Erro_Obtem_CodCliente:

    Obtem_CodCliente = gErr

    Select Case gErr

        Case 102618

        Case 102619
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objContato.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166618)

    End Select

End Function

Public Function Obtem_Contatos_FilialCliente(objFilialContato As ClassFilialContato) As Long

Dim lErro As Long
Dim lCliente As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objContato As New ClassContatos

On Error GoTo Erro_Obtem_Contatos_FilialCliente

    'Verifica se foi preenchido o Cliente
    If Len(Trim(Cliente.Text)) = 0 Then gError 102516

    objContato.sNomeReduzido = Cliente.Text

    'Faz a leitura do cliente
    lErro = CF("Contato_Le_NomeReduzido", objContato)
    If lErro <> SUCESSO And lErro <> 180745 Then gError 102517
    
    objFilialContato.lCodContato = objContato.lCodigo

    'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
    lErro = CF("FilialContato_Le", objFilialContato)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 102517

    'Se não existe a Filial
    If lErro = 17660 Then gError 102518
    
    'Obtém o telefone e os contatos da filial
    lErro = Obtem_Contatos_Cliente(objFilialContato)
    If lErro <> SUCESSO Then gError 102629

    Obtem_Contatos_FilialCliente = SUCESSO

    Exit Function

Erro_Obtem_Contatos_FilialCliente:

    Obtem_Contatos_FilialCliente = gErr

    Select Case gErr

        Case 102517, 102629
        
        Case 102516
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 102518
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCONTATO", FilialCliente.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisContatos", objFilialContato)
            Else
                'Segura o foco
            End If

        Case 102519
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166619)

    End Select

End Function

Public Function Obtem_Contatos_Cliente(objFilialContato As ClassFilialContato) As Long

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objClienteContatos As New ClassClienteContatos

On Error GoTo Erro_Obtem_Contatos_Cliente

    '*** CARGA DA COMBO DE CONTATOS ***
    'Guarda no objClienteContatos, o código do cliente e da
    objClienteContatos.lCliente = objFilialContato.lCodContato
    objClienteContatos.iFilialCliente = objFilialContato.iCodFilial
    
    'Carrega a combo de contatos
    lErro = CF("Carrega_ClienteFContatos", Contato, objClienteContatos)
    If lErro <> SUCESSO And lErro <> 181070 Then gError 102627
    '***********************************
    
    'Se selecionou o contato padrão =>
    If Len(Trim(Contato.Text)) > 0 Then
    
        'traz o telefone do contato
        Call Contato_Click
    
    Else
    
        'Limpa o campo telefon
        Telefone.Text = ""
    End If
    
    Obtem_Contatos_Cliente = SUCESSO

    Exit Function

Erro_Obtem_Contatos_Cliente:

    Obtem_Contatos_Cliente = gErr

    Select Case gErr

        Case 102625, 102627, 102658
        
        Case 102626
            Call Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO1", gErr, objFilialContato.iCodFilial, Trim(Cliente.Text))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166620)

    End Select

End Function

Public Function Trata_RelacionamentoAnterior(ByVal objRelacionamentoClientes As ClassRelacClientes) As Long

Dim lErro As Long
Dim objRelacionamentoAnt As ClassRelacClientes

On Error GoTo Erro_Trata_RelacionamentoAnterior

    '*** VALIDAÇÃO DO CÓDIGO DO RELACIONAMENTO ANTERIOR ***
    'Se o código do relacionamento anterior foi preenchido
    If StrParaDbl(RelacionamentoAnt.Text) > 0 Then
    
        'Instancia o obj
        Set objRelacionamentoAnt = New ClassRelacClientes
        
        'Guarda no obj o código e a filialempresa do relacionamento anterior
        objRelacionamentoAnt.lCodigo = StrParaDbl(RelacionamentoAnt.Text)
        objRelacionamentoAnt.iFilialEmpresa = giFilialEmpresa
        
        'Lê o relacionamento com os filtros passados
        lErro = CF("RelacionamentoContatos_Le", objRelacionamentoAnt)
        If lErro <> SUCESSO And lErro <> 181031 Then gError 102657
        
        'Se não encontrou o relacionamento
        'ou se esse relacionamento não é válido para
        'o cliente, a filial e a data em questão
        If (lErro = 181031) Or (objRelacionamentoAnt.lCliente <> objRelacionamentoClientes.lCliente) Or (objRelacionamentoAnt.iFilialCliente <> objRelacionamentoClientes.iFilialCliente) Or (objRelacionamentoAnt.dtData > objRelacionamentoClientes.dtData) Then
        
            'Limpa o frame Relacionamento Anterior
            Call Limpa_Frame_RelacionamentoAnt
        
        End If
    
    End If
    '******************************************************
    
    Trata_RelacionamentoAnterior = SUCESSO

    Exit Function

Erro_Trata_RelacionamentoAnterior:

    Trata_RelacionamentoAnterior = gErr

    Select Case gErr

        Case 102657
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166621)

    End Select

End Function

Private Function Valida_Cliente() As Long
'Faz a validação do cliente

Dim lErro As Long
Dim objContato As New ClassContatos
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialContato As New ClassFilialContato
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Valida_Cliente

    'Se o campo cliente não foi alterado => sai da função
    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Contato_Le(Cliente, objContato, iCodFilial)
        If lErro <> SUCESSO Then gError 102513

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisContatos_Le_Contato", objContato, colCodigoNome)
        If lErro <> SUCESSO Then gError 102514

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
        
        'Guarda no obj o código do cliente e da filial para efetuar a leitura dos contatos
        objFilialContato.lCodContato = objContato.lCodigo
        objFilialContato.iCodFilial = iCodFilial
        
        'Guarda no obj o código do endereço que será lido
        objFilialContato.lEndereco = objContato.lEndereco
        
        'Guarda no obj os dados necessários para validar o código do relacionamento anterior
        objRelacionamentoClientes.lCliente = objFilialContato.lCodContato
        objRelacionamentoClientes.iFilialCliente = objFilialContato.iCodFilial
        
        'Verifica se o código do relacionamento anterior é válido
        'para o cliente/filial em questão
        lErro = Trata_RelacionamentoAnterior(objRelacionamentoClientes)
        If lErro <> SUCESSO Then gError 102658
    
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear
        
        'Limpa a combo de contatos
        Contato.Clear
        
        'Limpa o telefone
        Telefone.Text = ""
        
    End If
    
    iClienteAlterado = 0
    
    Valida_Cliente = SUCESSO

    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 102513, 102514, 102630
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166622)

    End Select

    Exit Function

End Function

Private Function Valida_FilialCliente() As Long
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialContato As New ClassFilialContato
Dim iCodigo As Integer

On Error GoTo Erro_Valida_FilialCliente

    'Se a filial de cliente não foi alterada => sai da função
    If iFilialCliAlterada = 0 Then Exit Function
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) > 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialCliente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 102515
    
        'Se foi digitado o nome da filial
        'e esse nome não foi encontrado na combo => erro
        If lErro = 6731 Then gError 102519
        
        'Mesmo que tenha encontrado a filial na combo, é preciso fazer a leitura para
        'obter o telefone da mesma
        
        'Passa o Código da Filial que está na tela para o Obj
        objFilialContato.iCodFilial = iCodigo
        
        'Lê a filial e obtém o telefone e os contatos da mesma
        lErro = Obtem_Contatos_FilialCliente(objFilialContato)
        If lErro <> SUCESSO Then gError 102631
        
        'Encontrou Filial no BD, coloca no Text da Combo
        FilialCliente.Text = CStr(objFilialContato.iCodFilial) & SEPARADOR & objFilialContato.sNome
    
    'se não foi preenchida
    Else
    
        'Limpa a combo de contatos
        Contato.Clear
        
        'Limpa o campo telefone
        Telefone.Text = ""
    
    End If
    
    iFilialCliAlterada = 0
    
    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 102515, 102517, 102625, 102627, 102631

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166623)

    End Select

    Exit Function

End Function

'Incluído por Luiz Nogueira em 04/06/03
Private Function RelacionamentoClientes_Imprime(ByVal lCodRelacionamento As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_RelacionamentoClientes_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código do relacionamento passado como parâmetro
    objRelacionamentoClientes.lCodigo = lCodRelacionamento
    
    'Guarda a FilialEmpresa ativa como filial do relacionamento
    objRelacionamentoClientes.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados do relacionamento para verificar se o mesmo existe no BD
    lErro = CF("RelacionamentoContatos_Le", objRelacionamentoClientes)
    If lErro <> SUCESSO And lErro <> 181031 Then gError 102975

    'Se não encontrou => erro, pois não é possível imprimir um relacionamento inexistente
    If lErro = 181031 Then gError 102976
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Relacionamento Contatos", "Codigo>=@NCODINI E Codigo<=@NCODFIM", 1, "rlclifut", "NCODINI", CStr(lCodRelacionamento), "NCODFIM", CStr(lCodRelacionamento))
    If lErro <> SUCESSO Then gError 102977

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    RelacionamentoClientes_Imprime = SUCESSO
    
    Exit Function

Erro_RelacionamentoClientes_Imprime:

    RelacionamentoClientes_Imprime = gErr
    
    Select Case gErr
    
        Case 102975, 102977
        
        Case 102976
            Call Rotina_Erro(vbOKOnly, "ERRO_RELACIONAMENTO_NAO_ENCONTRADO", gErr, objRelacionamentoClientes.lCodigo, objRelacionamentoClientes.iFilialEmpresa)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 166624)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function
'*** FUNÇÕES DE APOIO À TELA - FIM ***

'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Relacionamento com clientes futuros"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "RelacionamentoContatos"
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

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
'''    m_Caption = New_Caption
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
'***************************************************
'Fim Trecho de codigo comum as telas
'***************************************************

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub LabelOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigem, Source, X, Y)
End Sub

Private Sub LabelOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigem, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub LabelHora_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHora, Source, X, Y)
End Sub

Private Sub LabelHora_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHora, Button, Shift, X, Y)
End Sub

Private Sub LabelTipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipo, Source, X, Y)
End Sub

Private Sub LabelTipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipo, Button, Shift, X, Y)
End Sub

Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub

Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelFilialCliente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilialCliente, Source, X, Y)
End Sub

Private Sub LabelFilialCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilialCliente, Button, Shift, X, Y)
End Sub

Private Sub LabelContato_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelContato, Source, X, Y)
End Sub

Private Sub LabelContato_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelContato, Button, Shift, X, Y)
End Sub

Private Sub LabelTelefone_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTelefone, Source, X, Y)
End Sub

Private Sub LabelTelefone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTelefone, Button, Shift, X, Y)
End Sub

Private Sub LabelAtendente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAtendente, Source, X, Y)
End Sub

Private Sub LabelAtendente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAtendente, Button, Shift, X, Y)
End Sub

Private Sub LabelCodContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodContatoAnt, Source, X, Y)
End Sub

Private Sub LabelCodContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelOrigemContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelOrigemContatoAnt, Source, X, Y)
End Sub

Private Sub LabelOrigemContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelOrigemContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelDataContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDataContatoAnt, Source, X, Y)
End Sub

Private Sub LabelDataContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDataContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelHoraContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelHoraContatoAnt, Source, X, Y)
End Sub

Private Sub LabelHoraContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelHoraContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelTipoContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTipoContatoAnt, Source, X, Y)
End Sub

Private Sub LabelTipoContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTipoContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub LabelAssunto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAssunto, Source, X, Y)
End Sub

Private Sub LabelAssunto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAssunto, Button, Shift, X, Y)
End Sub

Private Sub OrigemContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(OrigemContatoAnt, Source, X, Y)
End Sub

Private Sub OrigemContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(OrigemContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub DataContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataContatoAnt, Source, X, Y)
End Sub

Private Sub DataContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub HoraContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(HoraContatoAnt, Source, X, Y)
End Sub

Private Sub HoraContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(HoraContatoAnt, Button, Shift, X, Y)
End Sub

Private Sub TipoContatoAnt_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoContatoAnt, Source, X, Y)
End Sub

Private Sub TipoContatoAnt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoContatoAnt, Button, Shift, X, Y)
End Sub
'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***

Private Sub LabelContato_Click()

Dim objClienteContatos As New ClassClienteContatos
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_LabelContato_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 178445
    
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 178446

    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 178443
    
    'Guarda os dados necessários para tentar ler o contato
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)

    Call Chama_Tela("ClienteFContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelContato_Click:

    Select Case gErr

        Case 178443
        
        Case 178445
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 178446
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178444)

    End Select

    Exit Sub
    
End Sub

Private Sub LabelTelefone_Click()

Dim objClienteContatos As New ClassClienteContatos
Dim lErro As Long
Dim lCliente As Long

On Error GoTo Erro_LabelTelefone_Click

    If Len(Trim(Cliente.Text)) = 0 Then gError 178447
    
    If Len(Trim(FilialCliente.Text)) = 0 Then gError 178448

    'Obtém o código do cliente
    lErro = Obtem_CodCliente(lCliente)
    If lErro <> SUCESSO Then gError 178449
    
    'Guarda os dados necessários para tentar ler o contato
    objClienteContatos.lCliente = lCliente
    objClienteContatos.iFilialCliente = Codigo_Extrai(FilialCliente.Text)

    Call Chama_Tela("ClienteFContatos", objClienteContatos)
    
    Exit Sub

Erro_LabelTelefone_Click:

    Select Case gErr

        Case 178447
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
            
        Case 178448
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 178449
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 178450)

    End Select

    Exit Sub


End Sub
