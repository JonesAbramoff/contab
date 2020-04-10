VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl GarantiaOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5265
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   660
      Width           =   9255
      Begin VB.Frame Frame3 
         Caption         =   "Fabricante do Produto"
         Height          =   675
         Left            =   60
         TabIndex        =   61
         Top             =   4485
         Width           =   9120
         Begin VB.TextBox Cliente 
            Height          =   315
            Left            =   1215
            TabIndex        =   13
            ToolTipText     =   "Digite código, nome reduzido, cgc do cliente ou pressione F3 para consulta."
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox FilialCliente 
            Height          =   315
            Left            =   4875
            TabIndex        =   14
            ToolTipText     =   "Digite o nome ou o código da filial do cliente com quem foi feito o relacionamento."
            Top             =   240
            Width           =   1380
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4350
            TabIndex        =   63
            Top             =   300
            Width           =   465
         End
         Begin VB.Label LabelCliente 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   62
            Top             =   285
            Width           =   660
         End
      End
      Begin VB.Frame FrameLote 
         Caption         =   "Rastreamento por Lote ou OP"
         Height          =   675
         Left            =   60
         TabIndex        =   51
         Top             =   2745
         Width           =   9150
         Begin VB.Frame FrameOP 
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   2550
            TabIndex        =   55
            Top             =   195
            Width           =   4320
            Begin VB.ComboBox FilialOP 
               Height          =   315
               Left            =   900
               TabIndex        =   56
               Top             =   90
               Width           =   2805
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Filial OP:"
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
               Left            =   90
               TabIndex        =   57
               Top             =   150
               Width           =   780
            End
         End
         Begin MSMask.MaskEdBox Lote 
            Height          =   300
            Left            =   1260
            TabIndex        =   8
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            PromptChar      =   " "
         End
         Begin VB.Label LoteLabel 
            AutoSize        =   -1  'True
            Caption         =   "Lote:"
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
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   52
            Top             =   330
            Width           =   450
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Identificação"
         Height          =   2670
         Left            =   60
         TabIndex        =   43
         Top             =   45
         Width           =   5820
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2220
            Picture         =   "GarantiaOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   270
            Width           =   300
         End
         Begin VB.CheckBox Ativo 
            Caption         =   "Ativo"
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
            Left            =   3240
            TabIndex        =   2
            Top             =   300
            Value           =   1  'Checked
            Width           =   795
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   1260
            TabIndex        =   3
            Top             =   750
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataVenda 
            Height          =   300
            Left            =   2250
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1740
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataVenda 
            Height          =   300
            Left            =   1260
            TabIndex        =   5
            ToolTipText     =   "Informe a data quando ocorreu o relacionamento. Em caso de agendamento, informe a data de quando ocorrerá."
            Top             =   1755
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
            Left            =   1260
            TabIndex        =   0
            Top             =   255
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Quantidade 
            Height          =   315
            Left            =   1260
            TabIndex        =   7
            Top             =   2250
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
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
            Height          =   165
            Left            =   255
            TabIndex        =   50
            Top             =   1275
            Width           =   930
         End
         Begin VB.Label DescricaoProduto 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            Top             =   1230
            Width           =   4470
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
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
            Left            =   150
            TabIndex        =   49
            Top             =   2325
            Width           =   1050
         End
         Begin VB.Label LabelGarantia 
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
            Left            =   540
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   48
            Top             =   300
            Width           =   660
         End
         Begin VB.Label ProdutoLabel 
            AutoSize        =   -1  'True
            Caption         =   "Produto:"
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
            Height          =   165
            Left            =   450
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   47
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data Venda:"
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
            TabIndex        =   46
            Top             =   1785
            Width           =   1080
         End
         Begin VB.Label UM 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3450
            TabIndex        =   45
            Top             =   2250
            Width           =   780
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "U.M.:"
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
            Left            =   2925
            TabIndex        =   44
            Top             =   2325
            Width           =   480
         End
      End
      Begin VB.Frame FrameSerie 
         Caption         =   "Rastreamento por Números de Série"
         Height          =   2670
         Left            =   6090
         TabIndex        =   41
         Top             =   45
         Width           =   3120
         Begin VB.TextBox NumSerie 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   645
            MaxLength       =   20
            TabIndex        =   42
            Top             =   780
            Width           =   2190
         End
         Begin MSFlexGridLib.MSFlexGrid GridNumSerie 
            Height          =   2250
            Left            =   90
            TabIndex        =   15
            Top             =   225
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   3969
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Nota Fiscal de Compra do Material"
         Height          =   915
         Index           =   5
         Left            =   60
         TabIndex        =   33
         Top             =   3450
         Width           =   9150
         Begin VB.Frame FrameForn 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   3660
            TabIndex        =   58
            Top             =   480
            Width           =   5460
            Begin VB.ComboBox FilialFornecedor 
               Height          =   315
               Left            =   3825
               TabIndex        =   12
               Top             =   15
               Width           =   1590
            End
            Begin MSMask.MaskEdBox Fornecedor 
               Height          =   315
               Left            =   1215
               TabIndex        =   11
               Top             =   15
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin VB.Label FilialFornecedorLabel 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3270
               TabIndex        =   60
               Top             =   75
               Width           =   465
            End
            Begin VB.Label FornecedorLabel 
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   135
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   59
               Top             =   45
               Width           =   1035
            End
         End
         Begin VB.OptionButton NFExterna 
            Caption         =   "Externa"
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
            Left            =   2760
            TabIndex        =   54
            Top             =   285
            Width           =   1305
         End
         Begin VB.OptionButton NFInterna 
            Caption         =   "Interna"
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
            Left            =   1200
            TabIndex        =   53
            Top             =   270
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.ComboBox SerieNFiscalOriginal 
            Height          =   315
            Left            =   1215
            TabIndex        =   9
            Top             =   495
            Width           =   765
         End
         Begin MSMask.MaskEdBox NFiscalOriginal 
            Height          =   300
            Left            =   2790
            TabIndex        =   10
            Top             =   510
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin VB.Label SerieNFOriginalLabel 
            AutoSize        =   -1  'True
            Caption         =   "Série:"
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
            Left            =   600
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   35
            Top             =   540
            Width           =   510
         End
         Begin VB.Label NFiscalOriginalLabel 
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
            Height          =   195
            Left            =   2025
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   34
            Top             =   540
            Width           =   720
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5025
      Index           =   2
      Left            =   90
      TabIndex        =   36
      Top             =   750
      Visible         =   0   'False
      Width           =   9240
      Begin VB.Frame Frame7 
         Caption         =   "Serviços/Peças"
         Height          =   4305
         Left            =   405
         TabIndex        =   39
         Top             =   645
         Width           =   8460
         Begin VB.CommandButton BotaoServicos 
            Caption         =   "Serviços/Peças"
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
            Left            =   315
            TabIndex        =   24
            Top             =   3885
            Width           =   1740
         End
         Begin MSMask.MaskEdBox PrazoValidade 
            Height          =   315
            Left            =   4935
            TabIndex        =   23
            Top             =   1200
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.TextBox DescricaoServico 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1890
            MaxLength       =   250
            TabIndex        =   22
            Top             =   1200
            Width           =   3645
         End
         Begin MSMask.MaskEdBox Servico 
            Height          =   315
            Left            =   630
            TabIndex        =   21
            Top             =   1125
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridServicos 
            Height          =   1875
            Left            =   270
            TabIndex        =   20
            Top             =   705
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   3307
            _Version        =   393216
         End
         Begin VB.CheckBox GarantiaTotal 
            Caption         =   "Todos c/exceção dos listados abaixo"
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
            Left            =   420
            TabIndex        =   18
            Top             =   285
            Width           =   3690
         End
         Begin MSMask.MaskEdBox GarantiaTotalPrazo 
            Height          =   315
            Left            =   7650
            TabIndex        =   19
            Top             =   300
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Prazo (em dias):"
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
            Left            =   6195
            TabIndex        =   40
            Top             =   345
            Width           =   1380
         End
      End
      Begin MSMask.MaskEdBox TipoGarantia 
         Height          =   315
         Left            =   1755
         TabIndex        =   17
         Top             =   210
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label DescTipoGarantia 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2385
         TabIndex        =   38
         Top             =   210
         Width           =   3015
      End
      Begin VB.Label LblTipoGarantia 
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
         Left            =   1245
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   37
         Top             =   270
         Width           =   450
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
      Left            =   4770
      TabIndex        =   25
      Top             =   135
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6795
      ScaleHeight     =   450
      ScaleWidth      =   2625
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   15
      Width           =   2685
      Begin VB.CommandButton BotaoImprimir 
         Height          =   345
         Left            =   105
         Picture         =   "GarantiaOcx.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Imprimir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   615
         Picture         =   "GarantiaOcx.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   345
         Left            =   1110
         Picture         =   "GarantiaOcx.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Excluir"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   345
         Left            =   1620
         Picture         =   "GarantiaOcx.ctx":04D0
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2130
         Picture         =   "GarantiaOcx.ctx":0A02
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5685
      Left            =   45
      TabIndex        =   32
      Top             =   270
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   10028
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Serviços/Peças Garantidos"
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
Attribute VB_Name = "GarantiaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iTipoAlterado As Integer
Dim iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFilialCliAlterada As Integer

'Eventos de browser
Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoSerieNFiscalOriginal As AdmEvento
Attribute objEventoSerieNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoNFiscalOriginal As AdmEvento
Attribute objEventoNFiscalOriginal.VB_VarHelpID = -1
Private WithEvents objEventoServico As AdmEvento
Attribute objEventoServico.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1

Dim objGridServico As AdmGrid
Dim objGridNumSerie As AdmGrid

Dim iGrid_Servico_Col As Integer
Dim iGrid_ServicoDesc_Col As Integer
Dim iGrid_PrazoValidade_Col As Integer

Dim iGrid_NumSerie_Col As Integer

Dim giFrameAtual As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Private Function Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    giFrameAtual = 1
    
    'Inicializa eventos de browser
    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoLote = New AdmEvento
    Set objEventoSerieNFiscalOriginal = New AdmEvento
    Set objEventoNFiscalOriginal = New AdmEvento
    Set objEventoServico = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCliente = New AdmEvento
    
    Set objGridServico = New AdmGrid
    Set objGridNumSerie = New AdmGrid
    
    Call Inicializa_Grid_Servico(objGridServico)
    Call Inicializa_Grid_NumSerie(objGridNumSerie)
    
    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 183819
    
    'Carrega as Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 183819
    
    'Inicializa a Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 183820
    
    'Inicializa a Máscara de Servico
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Servico)
    If lErro <> SUCESSO Then gError 183821
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Function
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr
    
    Select Case gErr
    
        Case 183819 To 183821
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183822)
    
    End Select
    
End Function

Private Function Inicializa_Grid_Servico(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Serviço")
    objGridInt.colColuna.Add ("Desc. Serviço")
    objGridInt.colColuna.Add ("Prazo Validade")

    objGridInt.colCampo.Add (Servico.Name)
    objGridInt.colCampo.Add (DescricaoServico.Name)
    objGridInt.colCampo.Add (PrazoValidade.Name)

    'Controles que participam do Grid
    iGrid_Servico_Col = 1
    iGrid_ServicoDesc_Col = 2
    iGrid_PrazoValidade_Col = 3

    objGridInt.objGrid = GridServicos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_GARANTIA_SERVICOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridServicos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Servico = SUCESSO

End Function

Private Function Inicializa_Grid_NumSerie(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Número de Série")

    objGridInt.colCampo.Add (NumSerie.Name)

    'Controles que participam do Grid
    iGrid_NumSerie_Col = 1

    objGridInt.objGrid = GridNumSerie

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_NUM_SERIE + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridNumSerie.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_NumSerie = SUCESSO

End Function

Public Function Trata_Parametros(Optional ByVal objGarantia As ClassGarantia) As Long
'Trata os parametros passados para a tela..

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se recebeu um objeto com dados de um relacionamento
    If Not (objGarantia Is Nothing) Then
    
        'Lê e traz os dados do relacionamento para a tela
        lErro = Traz_Garantia_Tela(objGarantia)
        If lErro <> SUCESSO Then gError 183823
        
    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr
    
    Select Case gErr
    
        Case 183823
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183824)
    
    End Select
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoLote = Nothing
    Set objEventoSerieNFiscalOriginal = Nothing
    Set objEventoNFiscalOriginal = Nothing
    Set objEventoServico = Nothing
    Set objEventoTipo = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCliente = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Sub BotaoExibirDados_Click()

End Sub

Private Sub Codigo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then gError 183825
    
    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 183825
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183826)

    End Select

    Exit Sub

End Sub


Private Sub GarantiaTotal_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GarantiaTotalPrazo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub GarantiaTotalPrazo_GotFocus()
    Call MaskEdBox_TrataGotFocus(GarantiaTotalPrazo, iAlterado)
End Sub

Private Sub GarantiaTotalPrazo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_GarantiaTotalPrazo_Validate

    If Len(Trim(GarantiaTotalPrazo.ClipText)) = 0 Then Exit Sub

    lErro = Inteiro_Critica(GarantiaTotalPrazo.Text)
    If lErro <> SUCESSO Then gError 186120
    
    Exit Sub

Erro_GarantiaTotalPrazo_Validate:

    Cancel = True

    Select Case gErr

        Case 186120
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186121)

    End Select

    Exit Sub

End Sub

Private Sub LblTipoGarantia_Click()

Dim objTipoGarantia As New ClassTipoGarantia
Dim colSelecao As New Collection

    objTipoGarantia.lCodigo = StrParaLong(Codigo.Text)
    
    Call Chama_Tela("TipoGarantiaLista", colSelecao, objTipoGarantia, objEventoTipo)

End Sub

Private Sub NFExterna_Click()
    If NFExterna.Value Then
        FrameForn.Enabled = True
        NFiscalOriginalLabel.MousePointer = vbArrowQuestion
    Else
        FrameForn.Enabled = False
        Fornecedor.Text = ""
        NFiscalOriginalLabel.MousePointer = vbNormal
    End If
End Sub

Private Sub NFInterna_Click()
    If NFInterna.Value Then
        FrameForn.Enabled = False
        Fornecedor.Text = ""
        NFiscalOriginalLabel.MousePointer = vbNormal
    Else
        FrameForn.Enabled = True
        NFiscalOriginalLabel.MousePointer = vbArrowQuestion
    End If
End Sub


Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim objTipoGarantia As ClassTipoGarantia
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoGarantia = obj1
    
    'Lê o tipo
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 186123
    
    'Se não encontrar --> Erro
    If lErro <> SUCESSO Then gError 186124
        
    TipoGarantia.Text = objTipoGarantia.lCodigo
    
    lErro = Exibe_Dados_TipoGarantia(objTipoGarantia)
    If lErro <> SUCESSO Then gError 186125
        
    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case gErr
    
        Case 186123, 186125
        
        Case 186124
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOGARANTIA_NAO_CADASTRADA", gErr, objTipoGarantia.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186126)
    
    End Select

End Sub


Public Sub NFiscalOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_NFiscalOriginalLabel_Click

    If NFInterna.Value Then

        'Guarda a Serie e o Número da Nota Fiscal Original da Tela
        objNFiscal.sSerie = SerieNFiscalOriginal.Text
        If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
            objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
        Else
            objNFiscal.lNumNotaFiscal = 0
        End If
        
        'Formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 183900
        
        colSelecao.Add sProduto
    
        'Chama a Tela NFiscalInternaSaidalLista
        Call Chama_Tela("NFiscalInternaSaidaLista", colSelecao, objNFiscal, objEventoNFiscalOriginal, "NumIntDoc IN (SELECT NumIntNF FROM ItensNFiscal WHERE Produto = ?)")

    End If
    
    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 183900

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183831)

    End Select

    Exit Sub
    
End Sub

Public Sub objEventoNFiscalOriginal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal Original
    SerieNFiscalOriginal.Text = objNFiscal.sSerie
    NFiscalOriginal.Text = objNFiscal.lNumNotaFiscal

    Me.Show

End Sub

Public Sub SerieNFOriginalLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As New Collection

    'Recolhe a Série da Nota Fiscal Original da tela
    objSerie.sSerie = SerieNFiscalOriginal.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerieNFiscalOriginal)

End Sub

Public Sub objEventoSerieNFiscalOriginal_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal Original na tela
    SerieNFiscalOriginal.Text = objSerie.sSerie

    Me.Show

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Produto_Validate

    'Se produto não estiver preenchido --> limpa descrição e unidade de medida
    If Len(Trim(Produto.ClipText)) = 0 Then
        DescricaoProduto.Caption = ""
        UM.Caption = ""
    Else
        'Caso esteja preenchido
        lErro = CF("Produto_Critica", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 183829

        If lErro = 25041 Then gError 183830

        DescricaoProduto.Caption = ""
        UM.Caption = ""

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            DescricaoProduto.Caption = objProduto.sDescricao
            UM.Caption = objProduto.sSiglaUMEstoque
        End If
        
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
            FrameLote.Enabled = True
            FrameSerie.Enabled = False
            FrameOP.Enabled = False
            Call Grid_Limpa(objGridNumSerie)
            FilialOP.ListIndex = -1
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
            FrameLote.Enabled = True
            FrameSerie.Enabled = False
            FrameOP.Enabled = True
            Call Grid_Limpa(objGridNumSerie)
        ElseIf objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
            FrameLote.Enabled = False
            FrameSerie.Enabled = True
            FrameOP.Enabled = False
            Lote.Text = ""
            FilialOP.ListIndex = -1
        Else
            FrameLote.Enabled = False
            FrameSerie.Enabled = False
            FrameOP.Enabled = False
            Call Grid_Limpa(objGridNumSerie)
            Lote.Text = ""
            FilialOP.ListIndex = -1
        End If
        
    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 183829

        Case 183830
            DescricaoProduto.Caption = ""
            UM.Caption = ""

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                Call Chama_Tela("Produto", objProduto)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183831)

    End Select

    Exit Sub

End Sub


Public Sub DataVenda_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVenda_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataVenda, iAlterado)
End Sub

Public Sub DataVenda_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVenda_Validate

    'Verifica se a Data foi digitada
    If Len(Trim(DataVenda.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataVenda.Text)
    If lErro <> SUCESSO Then gError 183827

    Exit Sub

Erro_DataVenda_Validate:

    Cancel = True

    Select Case gErr

        Case 183827

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183828)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub

    'Critica a Quantidade
    lErro = Valor_Positivo_Critica(Quantidade.Text)
    If lErro <> SUCESSO Then gError 183834

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 183834

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183835)

    End Select

    Exit Sub

End Sub

Private Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Lote_GotFocus()

    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Private Sub FilialOP_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_FilialOP_Validate

    'Se não estiver preenchida ou alterada pula a crítica
    If Len(Trim(FilialOP.Text)) = 0 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialOP, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 183836

    'Nao encontrou o item com o código informado
    If lErro = 6730 Then gError 183837

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 183838

    Exit Sub

Erro_FilialOP_Validate:

    Cancel = True

    Select Case gErr

        Case 183836

        Case 183837
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 183838
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183839)

    End Select

    Exit Sub

End Sub

Public Sub FilialOP_Click()

Dim lErro As Long

On Error GoTo Erro_FilialOP_Click

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_FilialOP_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183840)

    End Select

    Exit Sub

End Sub

Public Sub SerieNFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SerieNFiscalOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SerieNFiscalOriginal_Validate

    'Verififca se está preenchida
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then Exit Sub

    'Verifica se foi alguma Série selecionada
    If SerieNFiscalOriginal.Text = SerieNFiscalOriginal.List(SerieNFiscalOriginal.ListIndex) Then Exit Sub

    'Tenta achar a Série na combo
    lErro = Combo_Item_Igual(SerieNFiscalOriginal)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 183841

    'Não encontrou a Série
    If lErro = 12253 Then gError 183842

    Exit Sub

Erro_SerieNFiscalOriginal_Validate:

    Cancel = True


    Select Case gErr

        Case 183841

        Case 183842
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, SerieNFiscalOriginal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183843)

    End Select

    Exit Sub

End Sub


Private Sub LabelGarantia_Click()

Dim objGarantia As New ClassGarantia
Dim colSelecao As New Collection

    objGarantia.lCodigo = StrParaLong(Codigo.Text)
    
    Call Chama_Tela("GarantiaLista", colSelecao, objGarantia, objEventoCodigo)
    
End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objGarantia As ClassGarantia
Dim bCancel As Boolean
Dim lErro As Long

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objGarantia = obj1
    
    'Traz para a tela o relacionamento com código passado pelo browser
    lErro = Traz_Garantia_Tela(objGarantia)
    If lErro <> SUCESSO Then gError 183868
        
    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr
    
        Case 183868
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183869)
    
    End Select

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 183832

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 183832

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183833)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 183870

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 183871

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then gError 183872

    'Preenche unidade de medida com SiglaUMEstoque
    UM.Caption = objProduto.sSiglaUMEstoque

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 183870, 183872

        Case 183871
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183873)

    End Select

    Exit Sub

End Sub

Private Sub LoteLabel_Click()

Dim colSelecao As New Collection
Dim objRastroLote As New ClassRastreamentoLote
Dim sProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sSelecao As String
Dim lErro As Long

On Error GoTo Erro_LoteLabel_Click

    objRastroLote.sCodigo = Lote.Text

    sProduto = Produto.Text

    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 186127

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'Selecao
        colSelecao.Add sProdutoFormatado

        sSelecao = "Produto = ?"

        'Chama tela de Browse de RastreamentoLote
        Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLote, sSelecao)

    Else
    
        'Chama tela de Browse de RastreamentoLote
        Call Chama_Tela("RastroLoteLista1", colSelecao, objRastroLote, objEventoLote)

    End If
    
    Exit Sub

Erro_LoteLabel_Click:

    Select Case gErr

        Case 186127

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186128)

    End Select

    Exit Sub

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_objEventoLote_evSelecao

    Set objRastroLote = obj1

    Lote.Text = objRastroLote.sCodigo

    Me.Show

    Exit Sub

Erro_objEventoLote_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183874)

    End Select

    Exit Sub

End Sub

Public Sub NFiscalOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NFiscalOriginal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscalOriginal, iAlterado)

End Sub

Public Sub TipoGarantia_Change()

    iAlterado = REGISTRO_ALTERADO
    iTipoAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoGarantia_GotFocus()

Dim iTipoGarantiaAux As Integer

    iTipoGarantiaAux = iTipoAlterado
    Call MaskEdBox_TrataGotFocus(TipoGarantia, iAlterado)
    iTipoAlterado = iTipoGarantiaAux

End Sub

Public Sub TipoGarantia_Validate(Cancel As Boolean)
'Se mudar o tipo trazer dele os defaults para os campos da tela

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoGarantia As New ClassTipoGarantia
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoGarantia_Validate

    'Verifica se o Tipo foi alterado
    If iTipoAlterado = 0 Then Exit Sub
    
    'Verifica se o Tipo está preenchido
    If Len(Trim(TipoGarantia.Text)) = 0 Then
        DescTipoGarantia.Caption = ""
        iTipoAlterado = 0
        Exit Sub
    End If

    'Critica o valor
    lErro = Long_Critica(TipoGarantia.Text)
    If lErro <> SUCESSO Then gError 183875

    objTipoGarantia.lCodigo = StrParaLong(TipoGarantia.Text)

    'Lê o tipo
    lErro = CF("TipoGarantia_Le", objTipoGarantia)
    If lErro <> SUCESSO And lErro <> 183849 Then gError 183876
    
    'Se não encontrar --> Erro
    If lErro <> SUCESSO Then gError 183877
        
    lErro = Exibe_Dados_TipoGarantia(objTipoGarantia)
    If lErro <> SUCESSO Then gError 183878
    
    iTipoAlterado = 0
    
    Exit Sub

Erro_TipoGarantia_Validate:

    Cancel = True

    Select Case gErr

        Case 183875, 183876, 183878

        Case 183877
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOGARANTIA", objTipoGarantia.lCodigo)

            If vbMsgRes = vbYes Then

                Call Chama_Tela("TipoGarantia", objTipoGarantia)

            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183879)

    End Select

    Exit Sub

End Sub

Function Exibe_Dados_TipoGarantia(objTipoGarantia As ClassTipoGarantia) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim sProduto As String

On Error GoTo Erro_Exibe_Dados_TipoGarantia

    'Coloca a Descrição na Tela
    DescTipoGarantia.Caption = objTipoGarantia.sDescricao
    
    GarantiaTotal.Value = objTipoGarantia.iGarantiaTotal
    If objTipoGarantia.iGarantiaTotalPrazo <> 0 Then
        GarantiaTotalPrazo.Text = objTipoGarantia.iGarantiaTotalPrazo
    Else
        GarantiaTotalPrazo.Text = ""
    End If

    Call Grid_Limpa(objGridServico)

    'Exibe os dados da coleção na tela
    For iIndice = 1 To objTipoGarantia.colTipoGarantiaProduto.Count

        objProduto.sCodigo = objTipoGarantia.colTipoGarantiaProduto.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 183880

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError 183880
        
        Servico.PromptInclude = False
        Servico.Text = sProduto
        Servico.PromptInclude = True

        'Insere no Grid Categoria
        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text
        
        GridServicos.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objProduto.sDescricao
        GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col) = objTipoGarantia.colTipoGarantiaProduto.Item(iIndice).iPrazo
 
    Next

    objGridServico.iLinhasExistentes = objTipoGarantia.colTipoGarantiaProduto.Count
    
    Exibe_Dados_TipoGarantia = SUCESSO

    Exit Function
    
Erro_Exibe_Dados_TipoGarantia:

    Exibe_Dados_TipoGarantia = gErr

    Select Case gErr
    
        Case 183880
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183881)
            
    End Select
     
    Exit Function
    
End Function

Public Sub BotaoServicos_Click()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoServicos_Click

    If Me.ActiveControl Is Servico Then
    
        sProduto1 = Servico.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridServicos.Row = 0 Then gError 183881

        sProduto1 = GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 183882
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    'preenche o codigo do produto
    objProduto.sCodigo = sProduto
    
    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoServico)

    Exit Sub
        
Erro_BotaoServicos_Click:
    
    Select Case gErr
        
        Case 183881
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 183882
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183883)

    End Select

    Exit Sub

End Sub

Private Sub objEventoServico_evSelecao(obj1 As Object)

Dim objProduto As ClassProduto
Dim sProduto As String
Dim lErro As Long

On Error GoTo Erro_objEventoServico_evSelecao

    Set objProduto = obj1

    'Verifica se alguma linha está selecionada
    If GridServicos.Row < 1 Then Exit Sub

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
    If lErro <> SUCESSO Then gError 183884

    Servico.PromptInclude = False
    Servico.Text = sProduto
    Servico.PromptInclude = True

    If Not (Me.ActiveControl Is Servico) Then
    
        GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = Servico.Text
    
        'Faz o Tratamento do produto
        lErro = Traz_Servico_Tela()
        If lErro <> SUCESSO Then gError 183885

    End If
    
    Me.Show

    Exit Sub

Erro_objEventoServico_evSelecao:

    Select Case gErr

        Case 183884
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 183885
            GridServicos.TextMatrix(GridServicos.Row, iGrid_Servico_Col) = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183886)

    End Select

    Exit Sub

End Sub

Private Function Traz_Servico_Tela() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Traz_Servico_Tela

    'Critica o Produto
    lErro = CF("Produto_Critica_Filial", Servico.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 183887
    
    If lErro = 51381 Then gError 183888

    'Descricao Servico
    GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao

    'Acrescenta uma linha no Grid se for o caso
    If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
        
        objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1

    End If

    Traz_Servico_Tela = SUCESSO

    Exit Function

Erro_Traz_Servico_Tela:

    Traz_Servico_Tela = gErr

    Select Case gErr

        Case 183887

        Case 183888
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, Produto.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183889)

    End Select

    Exit Function

End Function

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 183890

    'Limpa a Tela
    Call Limpa_Garantia

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 183890

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 183891)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim objGarantia As New ClassGarantia
Dim lErro As Long
Dim sAviso As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    'Se o código não foi preenchido => erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 183968

    'Guarda no obj, código da garantia e filial empresa
    'Essas informações são necessárias para excluir a garantia
    objGarantia.lCodigo = StrParaLong(Codigo.Text)
    objGarantia.iFilialEmpresa = giFilialEmpresa

    'Lê a garantia
    lErro = CF("Garantia_Le", objGarantia)
    If lErro <> SUCESSO And lErro <> 183568 Then gError 183969
    
    'Se não encontrou => erro
    If lErro <> SUCESSO Then gError 183970
    
    'Pede a confirmação da exclusão da garantia
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_GARANTIA")
    
    If vbMsgRes = vbYes Then

        'Faz a exclusão da Solicitacao
        lErro = CF("Garantia_Exclui", objGarantia)
        If lErro <> SUCESSO Then gError 183971
    
        'Limpa a Tela de Orcamento de Venda
        Call Limpa_Garantia
        
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 183968
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183969, 183971

        Case 183970
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_NAO_ENCONTRADA", gErr, objGarantia.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 183972)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 183986

    'Limpa a Tela
    Call Limpa_Garantia
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 183986

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183987)

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
    lErro = CF("Config_ObterAutomatico", "SRVConfig", "NUM_PROX_GARANTIA_1", "Garantia", "Codigo", lCodigo)
    If lErro <> SUCESSO Then gError 183988
    
    'Exibe o código obtido
    Codigo.PromptInclude = False
    Codigo.Text = lCodigo
    Codigo.PromptInclude = True
    
    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 183988
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183989)

    End Select

End Sub

Private Sub TabStrip1_Click()

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual
    If TabStrip1.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        Frame1(giFrameAtual).Visible = False

        'Armazena novo valor de giFrameAtual
        giFrameAtual = TabStrip1.SelectedItem.Index
       
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183990)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVenda_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataVenda_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataVenda, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 183991

    Exit Sub

Erro_UpDownDataVenda_DownClick:

    Select Case gErr

        Case 183991

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183992)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVenda_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVenda_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataVenda, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 183993

    Exit Sub

Erro_UpDownDataVenda_UpClick:

    Select Case gErr

        Case 183994

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183995)

    End Select

    Exit Sub

End Sub


'*** TRATAMENTO DO EVENTO KEYDOWN  - INÍCIO ***
Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call LabelGarantia_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call LoteLabel_Click
        ElseIf Me.ActiveControl Is NFiscalOriginal Then
            Call NFiscalOriginalLabel_Click
        ElseIf Me.ActiveControl Is SerieNFiscalOriginal Then
            Call SerieNFOriginalLabel_Click
        ElseIf Me.ActiveControl Is Servico Then
            Call BotaoServicos_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub


'***************************************************
'Trecho de codigo comum as telas
'***************************************************

Public Function Form_Load_Ocx() As Object
'    ??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Garantia"
    Call Form_Load
End Function

Public Function Name() As String
    Name = "Garantia"
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

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - INÍCIO ***
Private Sub LabelGarantia_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelGarantia, Source, X, Y)
End Sub

Private Sub LabelGarantia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelGarantia, Button, Shift, X, Y)
End Sub

'*** TRATAMENTO DE DRAG AND DROP / MOUSEDOWN DOS LABELS - FIM ***


Private Function Traz_Garantia_Tela(ByVal objGarantia As ClassGarantia) As Long
'Traz pra tela os dados da garantia passado como parâmetro

Dim lErro As Long
Dim bCancel As Boolean
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim objTipoGarantia As New ClassTipoGarantia
Dim iAchou As Integer

On Error GoTo Erro_Traz_Garantia_Tela

    'Limpa a tela
    Call Limpa_Garantia
    
    'Lê no BD os dados da garantia em questao
    lErro = CF("Garantia_Le", objGarantia)
    If lErro <> SUCESSO And lErro <> 183568 Then gError 183858
    
    'Se não encontrou a garantia => erro
    If lErro <> SUCESSO Then gError 183859
    
    Codigo.PromptInclude = False
    Codigo.Text = objGarantia.lCodigo
    Codigo.PromptInclude = True

    If objGarantia.dtDataVenda <> DATA_NULA Then
        DataVenda.PromptInclude = False
        DataVenda.Text = Format(objGarantia.dtDataVenda, "dd/mm/yy")
        DataVenda.PromptInclude = True
    End If
    
    If objGarantia.lFornecedor <> 0 Then
        Call Fornecedor_Formata(objGarantia.lFornecedor)
        Call Filial_Formata(FilialFornecedor, objGarantia.iFilialFornecedor)
        NFExterna.Value = True
    Else
        NFInterna.Value = True
    End If
    Call NFInterna_Click
    
    'Se o código do cliente está preenchido
    If objGarantia.lCliFabr <> 0 Then
    
        Call Cliente_Formata(objGarantia.lCliFabr)

        'Se a filial do cliente está preenchida
        If objGarantia.iFilialCliFabr <> 0 Then

            Call Filial_FormataCli(FilialCliente, objGarantia.iFilialCliFabr)
            
        End If
        
    End If
        
    objProduto.sCodigo = objGarantia.sProduto
    
    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 183860
    
    lErro = CF("Traz_Produto_MaskEd", objGarantia.sProduto, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then gError 183861
    Call Produto_Validate(bSGECancelDummy)
    
    'Formata quantidade antes de colocar na Tela
    Quantidade.Text = Formata_Estoque(objGarantia.dQuantidade)
    
    UM.Caption = objProduto.sSiglaUMEstoque
    
    Lote.Text = objGarantia.sLote
    
    iAchou = 0

    'Se o Rastreamento possui FilialOP (Rastro Por Ordem de Produção)
    If objGarantia.iFilialOP <> 0 Then

        For iIndice = 0 To FilialOP.ListCount - 1
            If FilialOP.ItemData(iIndice) = objGarantia.iFilialOP Then
                iAchou = 1
                FilialOP.ListIndex = iIndice
                Exit For
            End If
        Next
        
        'se a filial não foi encontrada ==> erro
        If iAchou = 0 Then gError 183862

    End If
    
    
    SerieNFiscalOriginal.Text = objGarantia.sSerie
    
    If objGarantia.lNumNotaFiscal <> 0 Then
        NFiscalOriginal.Text = objGarantia.lNumNotaFiscal
    End If
    
    If objGarantia.lTipoGarantia <> 0 Then
    
        TipoGarantia.Text = objGarantia.lTipoGarantia
        
        objTipoGarantia.lCodigo = objGarantia.lTipoGarantia
        
        lErro = CF("TipoGarantia_Le", objTipoGarantia)
        If lErro <> SUCESSO And lErro <> 183849 Then gError 183863
        
        If lErro = SUCESSO Then
            DescTipoGarantia.Caption = objTipoGarantia.sDescricao
        End If
    End If
    
    If objGarantia.iGarantiaTotal = MARCADO Then
        GarantiaTotal.Value = vbChecked
    Else
        GarantiaTotal.Value = vbUnchecked
    End If
    
    If objGarantia.iGarantiaTotalPrazo <> 0 Then GarantiaTotalPrazo.Text = objGarantia.iGarantiaTotalPrazo
    
    lErro = Carrega_Grid_Servicos(objGarantia)
    If lErro <> SUCESSO Then gError 183864
    
    lErro = Carrega_Grid_NumSerie(objGarantia)
    If lErro <> SUCESSO Then gError 183865
    
    iAlterado = 0
    
    Traz_Garantia_Tela = SUCESSO

    Exit Function

Erro_Traz_Garantia_Tela:

    Traz_Garantia_Tela = gErr

    Select Case gErr

        Case 183858, 183860, 183861, 183863, 183864, 183865
        
        Case 183859
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_NAO_ENCONTRADA", gErr, objGarantia.lCodigo)
            
        Case 183862
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objGarantia.iFilialOP)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183866)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_Servicos(objGarantia As ClassGarantia) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim objGarantiaProduto As ClassGarantiaProduto
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Servicos

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridServico)

    For iIndice = 1 To objGarantia.colProduto.Count
       
        Set objGarantiaProduto = objGarantia.colProduto(iIndice)
       
        lErro = Mascara_RetornaProdutoEnxuto(objGarantiaProduto.sProduto, sServicoEnxuto)
        If lErro <> SUCESSO Then gError 183867

        'Mascara o produto enxuto
        Servico.PromptInclude = False
        Servico.Text = sServicoEnxuto
        Servico.PromptInclude = True

        GridServicos.TextMatrix(iIndice, iGrid_Servico_Col) = Servico.Text
        
        objProduto.sCodigo = objGarantiaProduto.sProduto
        
        'Lê o Servico
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 183868
        
        If lErro = SUCESSO Then
            GridServicos.TextMatrix(iIndice, iGrid_ServicoDesc_Col) = objProduto.sDescricao
        End If
        
        If objGarantiaProduto.iPrazo <> 0 Then GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col) = objGarantiaProduto.iPrazo
        
    Next

    'Atualiza o número de linhas existentes
    objGridServico.iLinhasExistentes = objGarantia.colProduto.Count

    Carrega_Grid_Servicos = SUCESSO

    Exit Function

Erro_Carrega_Grid_Servicos:

    Carrega_Grid_Servicos = gErr

    Select Case gErr

        Case 183867, 183868
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183869)

    End Select

    Exit Function

End Function

Private Function Carrega_Grid_NumSerie(objGarantia As ClassGarantia) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sServicoEnxuto As String
Dim objGarantiaNumSerie As ClassGarantiaNumSerie

On Error GoTo Erro_Carrega_Grid_NumSerie

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNumSerie)

    For iIndice = 1 To objGarantia.colNumSerie.Count
       
        Set objGarantiaNumSerie = objGarantia.colNumSerie(iIndice)
       
        GridNumSerie.TextMatrix(iIndice, iGrid_NumSerie_Col) = objGarantiaNumSerie.sNumSerie
        
    Next

    'Atualiza o número de linhas existentes
    objGridNumSerie.iLinhasExistentes = objGarantia.colNumSerie.Count

    Carrega_Grid_NumSerie = SUCESSO

    Exit Function

Erro_Carrega_Grid_NumSerie:

    Carrega_Grid_NumSerie = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183908)

    End Select

    Exit Function

End Function

Private Sub GridServicos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)

    End If

End Sub

Private Sub GridServicos_EnterCell()

    Call Grid_Entrada_Celula(objGridServico, iAlterado)

End Sub

Private Sub GridServicos_GotFocus()

    Call Grid_Recebe_Foco(objGridServico)

End Sub

Private Sub GridServicos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridServico, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridServico, iAlterado)
    End If


End Sub

Private Sub GridServicos_LeaveCell()

    Call Saida_Celula(objGridServico)

End Sub

Private Sub GridServicos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridServico)

End Sub

Private Sub GridServicos_Scroll()

    Call Grid_Scroll(objGridServico)

End Sub

Private Sub GridServicos_RowColChange()

    Call Grid_RowColChange(objGridServico)

End Sub

Private Sub GridServicos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridServico)

End Sub

Public Sub Servico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Servico_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub Servico_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub Servico_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = Servico
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrazoValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrazoValidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridServico)

End Sub

Public Sub PrazoValidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridServico)

End Sub

Public Sub PrazoValidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridServico.objControle = PrazoValidade
    lErro = Grid_Campo_Libera_Foco(objGridServico)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    
    If lErro = SUCESSO Then

        If objGridInt.objGrid.Name = GridServicos.Name Then

            'Verifica qual a coluna atual do Grid
            Select Case objGridInt.objGrid.Col
    
                'Se for a de Servico
                Case iGrid_Servico_Col
                    lErro = Saida_Celula_Servico(objGridInt)
                    If lErro <> SUCESSO Then gError 183996
        
                'Se for a de Prazo
                Case iGrid_PrazoValidade_Col
                    lErro = Saida_Celula_PrazoValidade(objGridInt)
                    If lErro <> SUCESSO Then gError 183997
            
            End Select
    
    
        Else
            
            Select Case objGridInt.objGrid.Col
            
                'Se for a de Servico
                Case iGrid_NumSerie_Col
                    'lErro = Saida_Celula_NumSerie(objGridInt)
                    lErro = Saida_Celula_Padrao(objGridInt, NumSerie, True)
                    If lErro <> SUCESSO Then gError 186009
            
    
            End Select

        End If
            

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 183998
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 183996 To 183998, 186009

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183999)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Servico(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Servico

    Set objGridInt.objControle = Servico

    If Len(Trim(Servico.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica", Servico.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 Then gError 186000

        'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
        If lErro <> SUCESSO Then gError 186001
                
    Else
        
        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = ""
        GridServicos.TextMatrix(GridServicos.Row, iGrid_PrazoValidade_Col) = ""
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186002

    If Len(Trim(Servico.ClipText)) <> 0 Then

        GridServicos.TextMatrix(GridServicos.Row, iGrid_ServicoDesc_Col) = objProduto.sDescricao
    
        If GridServicos.Row - GridServicos.FixedRows = objGridServico.iLinhasExistentes Then
            
            objGridServico.iLinhasExistentes = objGridServico.iLinhasExistentes + 1
    
        End If
    
    End If

    Saida_Celula_Servico = SUCESSO

    Exit Function

Erro_Saida_Celula_Servico:

    Saida_Celula_Servico = gErr

    Select Case gErr

        Case 186000, 186002
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 186001
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Servico.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Servico.Text
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186003)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PrazoValidade(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Garantia está deixando de ser a corrente

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_Saida_Celula_PrazoValidade

    Set objGridInt.objControle = PrazoValidade

    If Len(Trim(PrazoValidade.Text)) > 0 Then

        lErro = Inteiro_Critica(PrazoValidade.Text)
        If lErro <> SUCESSO Then gError 186004
        
    End If

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186005
    
    Saida_Celula_PrazoValidade = SUCESSO

    Exit Function

Erro_Saida_Celula_PrazoValidade:

    Saida_Celula_PrazoValidade = gErr

    Select Case gErr

        Case 186004, 186005
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186006)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NumSerie(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Quantidadeque está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_NumSerie

    Set objGridInt.objControle = NumSerie

    'Passa quantidade para o grid (p/ usar PrecoTotal_Calcula)
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 186010
    
    Saida_Celula_NumSerie = SUCESSO

    Exit Function

Erro_Saida_Celula_NumSerie:

    Saida_Celula_NumSerie = gErr

    Select Case gErr

        Case 186010
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186011)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 186007

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 186007

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 186008)

    End Select

    Exit Function

End Function

Private Sub Limpa_Garantia()

Dim iIndice As Integer

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    FilialOP.ListIndex = -1
    SerieNFiscalOriginal.ListIndex = -1
    SerieNFiscalOriginal.Text = ""
    
    DescTipoGarantia.Caption = ""
    DescricaoProduto.Caption = ""
    UM.Caption = ""
    
    FilialFornecedor.Clear
    NFInterna.Value = True
    
    FilialCliente.Clear
    
    Call Grid_Limpa(objGridServico)
    Call Grid_Limpa(objGridNumSerie)
    
    iAlterado = 0
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos obrigatórios estão preenchidos
    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 183910

    'Move os dados da tela para o objRelacionamentoClie
    lErro = Move_Garantia_Memoria(objGarantia)
    If lErro <> SUCESSO Then gError 183911

    'Verifica se essa solicitação já existe no BD
    'e, em caso positivo, alerta ao usuário que está sendo feita uma alteração
    lErro = Trata_Alteracao(objGarantia, objGarantia.iFilialEmpresa, objGarantia.lCodigo)
    If lErro <> SUCESSO Then gError 183912
    
    'Grava no BD
    lErro = CF("Garantia_Grava", objGarantia)
    If lErro <> SUCESSO Then gError 183913

    'Se for para imprimir o relacionamento depois da gravação
    If ImprimeGravacao.Value = vbChecked Then

        'Dispara função para imprimir orçamento
        lErro = Garantia_Imprime(objGarantia.lCodigo)
        If lErro <> SUCESSO Then gError 183914

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 183910 To 183914
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183915)

    End Select

    Exit Function

End Function

Private Function Valida_Gravacao() As Long
'Verifica se os dados da tela são válidos para a gravação do registro

Dim lErro As Long
Dim iIndice As Integer
Dim dQuantidade As Double

On Error GoTo Erro_Valida_Gravacao

    'Se o código não estiver preenchido => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 183892
    
    'Se o produto não estiver preenchido => erro
    If Len(Trim(Produto.ClipText)) = 0 Then gError 183893
    
    'Se a data de venda não estiver preenchida => erro
    If Len(Trim(DataVenda.ClipText)) = 0 Then gError 183894
    
    'Se a quantidade não estiver preenchido => erro
    If Len(Trim(Quantidade.Text)) = 0 Then gError 183895
    
    If StrParaDbl(Quantidade.Text) = 0 Then gError 183898
    
    If NFExterna.Value Then
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 129411
        If Codigo_Extrai(FilialFornecedor.Text) = 0 Then gError 129412
    End If
    
    If GarantiaTotal.Value = 1 And Len(Trim(GarantiaTotalPrazo.Text)) = 0 Then gError 207376
    
    For iIndice = 1 To objGridServico.iLinhasExistentes

        If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_Servico_Col))) = 0 Then gError 183896
        
        If GarantiaTotal.Value = 0 Then

            If Len(Trim(GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col))) = 0 Then gError 183897
        
        End If
        
    Next
    
    Valida_Gravacao = SUCESSO

    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 129411
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 129412
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFornecedor_NAO_INFORMADA", gErr)
    
        Case 183892
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 183893
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            
        Case 183894
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENDA_NAO_PREENCHIDA", gErr)
            
        Case 183895
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr)
        
        Case 183896
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case 183897
            Call Rotina_Erro(vbOKOnly, "ERRO_PRAZO_NAO_PREENCHIDO_GRID", gErr, iIndice)
        
        Case 183898
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_ZERADA", gErr)
        
        Case 207376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRAZO_GARANTIATOTAL_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183899)

    End Select

End Function

Private Function Move_Garantia_Memoria(objGarantia As ClassGarantia) As Long
'Move os dados da tela para objGarantia

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim iPreenchido As Integer
Dim sProduto As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Garantia_Memoria

    objGarantia.iFilialEmpresa = giFilialEmpresa
    
    objGarantia.lCodigo = StrParaLong(Codigo.Text)

    objGarantia.iAtivo = Ativo.Value
    
    If Len(Trim(Cliente.Text)) > 0 Then
    
        objcliente.sNomeReduzido = Cliente.Text
    
        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 183623
    
        'Se não achou o Cliente --> erro
        If lErro = 12348 Then gError 183624
    
        'Guarda código do Cliente em objPedidoVenda
        objGarantia.lCliFabr = objcliente.lCodigo
    
        objGarantia.iFilialCliFabr = Codigo_Extrai(FilialCliente.Text)
    
    End If

    'Formata o produto
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 183900

    If iPreenchido = PRODUTO_VAZIO Then gError 183901

    objGarantia.sProduto = sProduto

    objGarantia.dtDataVenda = StrParaDate(DataVenda.Text)
    
    objGarantia.dQuantidade = StrParaDbl(Quantidade.Text)
    
    objGarantia.sLote = Lote.Text
    
    objGarantia.iFilialOP = Codigo_Extrai(FilialOP.Text)

    objGarantia.sSerie = SerieNFiscalOriginal.Text
    
    objGarantia.lNumNotaFiscal = StrParaLong(NFiscalOriginal.Text)
    
    objGarantia.lTipoGarantia = StrParaLong(TipoGarantia.Text)
    
    If GarantiaTotal.Value = vbChecked Then
        objGarantia.iGarantiaTotal = MARCADO
    Else
        objGarantia.iGarantiaTotal = DESMARCADO
    End If

    objGarantia.iGarantiaTotalPrazo = StrParaInt(GarantiaTotalPrazo.Text)
    
    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) > 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o Fornecedor através do Nome Reduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 183902

        If lErro = SUCESSO Then objGarantia.lFornecedor = objFornecedor.lCodigo
                            
    End If
    objGarantia.iFilialFornecedor = Codigo_Extrai(FilialFornecedor.Text)

    Set objGarantia.objTela = Me

    'Move Grid Itens para memória
    lErro = Move_GridServico_Memoria(objGarantia)
    If lErro <> SUCESSO Then gError 183902

    lErro = Move_GridNumSerie_Memoria(objGarantia)
    If lErro <> SUCESSO Then gError 183903

    Move_Garantia_Memoria = SUCESSO

    Exit Function

Erro_Move_Garantia_Memoria:

    Move_Garantia_Memoria = gErr

    Select Case gErr

        Case 183623, 183900, 183902, 183903

        Case 183624
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case 183901
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183904)

    End Select

    Exit Function

End Function

Private Function Move_GridServico_Memoria(objGarantia As ClassGarantia) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim sProduto As String
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim objGarantiaProduto As ClassGarantiaProduto
Dim iIndice As Integer

On Error GoTo Erro_Move_GridServico_Memoria

    For iIndice = 1 To objGridServico.iLinhasExistentes

        Set objGarantiaProduto = New ClassGarantiaProduto
    
        'Formata o produto
        lErro = CF("Produto_Formata", GridServicos.TextMatrix(iIndice, iGrid_Servico_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 183905
    
        If iPreenchido = PRODUTO_VAZIO Then gError 183906
    
        objGarantiaProduto.sProduto = sProduto
        objGarantiaProduto.iPrazo = StrParaInt(GridServicos.TextMatrix(iIndice, iGrid_PrazoValidade_Col))
    
        objGarantia.colProduto.Add objGarantiaProduto
    
    Next
    
    Move_GridServico_Memoria = SUCESSO

    Exit Function

Erro_Move_GridServico_Memoria:

    Move_GridServico_Memoria = gErr

    Select Case gErr

        Case 183905

        Case 183906
            Call Rotina_Erro(vbOKOnly, "ERRO_SERVICO_NAO_PREENCHIDO_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183907)

    End Select

    Exit Function

End Function

Private Function Move_GridNumSerie_Memoria(objGarantia As ClassGarantia) As Long
'Recolhe do Grid os dados

Dim lErro As Long
Dim sProduto As String
Dim sServico As String
Dim iPreenchido As Integer
Dim dQuantidade As Double
Dim objGarantiaNumSerie As ClassGarantiaNumSerie
Dim iIndice As Integer

On Error GoTo Erro_Move_GridNumSerie_Memoria

    For iIndice = 1 To objGridNumSerie.iLinhasExistentes

        Set objGarantiaNumSerie = New ClassGarantiaNumSerie
    
        objGarantiaNumSerie.sNumSerie = GridNumSerie.TextMatrix(iIndice, iGrid_NumSerie_Col)
    
        objGarantia.colNumSerie.Add objGarantiaNumSerie
    
    Next
    
    Move_GridNumSerie_Memoria = SUCESSO

    Exit Function

Erro_Move_GridNumSerie_Memoria:

    Move_GridNumSerie_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183909)

    End Select

    Exit Function


End Function

'**** TRATAMENTO DO SISTEMA DE SETAS - INÍCIO ****
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim objGarantia As New ClassGarantia
Dim objCampoValor As AdmCampoValor
Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Garantia"

    'Guarda no obj os dados que serão usados para identifica o registro a ser exibido
    objGarantia.lCodigo = StrParaLong(Trim(Codigo.Text))
    objGarantia.iFilialEmpresa = giFilialEmpresa
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objGarantia.lCodigo, 0, "Codigo"
    colCampoValor.Add "FilialEmpresa", objGarantia.iFilialEmpresa, 0, "FilialEmpresa"
    
    'Filtro
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186012)

    End Select

    Exit Sub
    
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objGarantia As New ClassGarantia

On Error GoTo Erro_Tela_Preenche

    'Guarda o código do campo em questão no obj
    objGarantia.lCodigo = colCampoValor.Item("Codigo").vValor
    objGarantia.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor

    lErro = Traz_Garantia_Tela(objGarantia)
    If lErro <> SUCESSO Then gError 186013

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
    
        Case 186013
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186014)

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

Public Sub BotaoImprimir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoImprimir_Click

    'Se o código da Solicitacao não foi informado => erro
    If Len(Trim(Codigo.Text)) = 0 Then gError 186015

    'Dispara função para imprimir relacionamento
    lErro = Garantia_Imprime(StrParaLong(Codigo.Text))
    If lErro <> SUCESSO Then gError 186016

    Exit Sub

Erro_BotaoImprimir_Click:

    Select Case gErr

        Case 186015
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 186016

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 186017)

    End Select

    Exit Sub

End Sub

Private Function Garantia_Imprime(ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objGarantia As New ClassGarantia
Dim lNumIntRel As Long

On Error GoTo Erro_Garantia_Imprime

    'Transforma o ponteiro do mouse em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Guarda no obj o código da solicitacao passado como parâmetro
    objGarantia.lCodigo = lCodigo
    
    'Guarda a FilialEmpresa ativa como filial do relacionamento
    objGarantia.iFilialEmpresa = giFilialEmpresa
    
    'Lê os dados da solicitacao para verificar se o mesmo existe no BD
    lErro = CF("Garantia_Le", objGarantia)
    If lErro <> SUCESSO And lErro <> 183568 Then gError 186018

    'Se não encontrou => erro, pois não é possível imprimir uma solicitacao inexistente
    If lErro <> SUCESSO Then gError 186019
    
    lErro = CF("RelGarantia_Prepara", objGarantia, lNumIntRel)
    If lErro <> SUCESSO Then gError 186018
    
    'Dispara a impressão do relatório
    lErro = objRelatorio.ExecutarDireto("Garantia", "", 1, "Garantia", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 186020

    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault
    
    Garantia_Imprime = SUCESSO
    
    Exit Function

Erro_Garantia_Imprime:

    Garantia_Imprime = gErr
    
    Select Case gErr
    
        Case 186018, 186020
        
        Case 186019
            Call Rotina_Erro(vbOKOnly, "ERRO_GARANTIA_NAO_ENCONTRADA", gErr, objGarantia.lCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 186020)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Private Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 35698

    'Carrega na combo
    For Each objSerie In colSerie
        SerieNFiscalOriginal.AddItem objSerie.sSerie
    Next
            
    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 35698

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156182)

    End Select

    Exit Function

End Function

Private Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_FornecedorLabel_Click

    'Se é possível extrair o código do Fornecedor do conteúdo do controle
    If LCodigo_Extrai(Fornecedor.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objFornecedor.lCodigo = LCodigo_Extrai(Fornecedor.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do Fornecedor
    Else
        
        'Prenche o Nome Reduzido do Fornecedor com o Fornecedor da Tela
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        sOrdenacao = "Nome Reduzido + Código"
    
    End If
    
    'Chama a tela de consulta de Fornecedor
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor, "", sOrdenacao)

    Exit Sub
    
Erro_FornecedorLabel_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155185)
    
    End Select
    
End Sub

Private Sub Fornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 129422

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError 129423

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialFornecedor, iCodFilial)
        
    'Se Fornecedor não está preenchido
    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialFornecedor.Clear

    End If
       
    Exit Sub

Erro_Fornecedor_Validate:
        
    Cancel = True

    Select Case gErr
    
        Case 129422, 129423
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155186)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_FilialFornecedor_Validate
        
    If Len(Trim(FilialFornecedor.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o item selecionado na ComboBox Filial
    If FilialFornecedor.Text = FilialFornecedor.List(FilialFornecedor.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(FilialFornecedor, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 129418

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 129419

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 129420

        'Se não existe a Filial
        If lErro = 18272 Then gError 129421

        'Encontrou Filial no BD, coloca no Text da Combo
        FilialFornecedor.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 129505
    
    Exit Sub
    
Erro_FilialFornecedor_Validate:

    Select Case gErr

        Case 129418, 129420

        Case 129419
            Call Rotina_Erro(vbOKOnly, "ERRO_Fornecedor_NAO_PREENCHIDO", gErr)

        Case 129421
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 129505
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, FilialFornecedor.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155187)

    End Select

    Exit Sub

End Sub

Private Sub FilialFornecedor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FilialFornecedor_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor
Dim bCancel As Boolean

    Set objFornecedor = obj1

    'Preenche o Fornecedor com o Fornecedor selecionado
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Dispara o Validate de Fornecedor
    Call Fornecedor_Validate(bCancel)

    Exit Sub

End Sub

Public Sub Fornecedor_Formata(lFornecedor As Long)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Formata

    Fornecedor.Text = lFornecedor

    'Busca o Fornecedor no BD
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then gError 129500

    lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
    If lErro <> SUCESSO Then gError 129501

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilialFornecedor, colCodigoNome)
    
    Exit Sub

Erro_Fornecedor_Formata:

    Select Case gErr

        Case 129500, 129501

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155174)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sFornecedor = Fornecedor.Text
    objFilialFornecedor.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
    If lErro <> SUCESSO And lErro <> 18272 Then gError 129498

    If lErro = 18272 Then gError 129499

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialFornecedor.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 129498

        Case 129499
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, objFilial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155175)

    End Select

    Exit Sub

End Sub

Private Sub GridNumSerie_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNumSerie, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)

    End If

End Sub

Private Sub GridNumSerie_EnterCell()

    Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)

End Sub

Private Sub GridNumSerie_GotFocus()

    Call Grid_Recebe_Foco(objGridNumSerie)

End Sub

Private Sub GridNumSerie_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNumSerie, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNumSerie, iAlterado)
    End If


End Sub

Private Sub GridNumSerie_LeaveCell()

    Call Saida_Celula(objGridNumSerie)

End Sub

Private Sub GridNumSerie_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridNumSerie)

End Sub

Private Sub GridNumSerie_Scroll()

    Call Grid_Scroll(objGridNumSerie)

End Sub

Private Sub GridNumSerie_RowColChange()

    Call Grid_RowColChange(objGridNumSerie)

End Sub

Private Sub GridNumSerie_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridNumSerie)

End Sub

Public Sub NumSerie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumSerie_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridNumSerie)

End Sub

Public Sub NumSerie_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNumSerie)

End Sub

Public Sub NumSerie_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNumSerie.objControle = NumSerie
    lErro = Grid_Campo_Libera_Foco(objGridNumSerie)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Cliente_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Cliente_Validate

    'Faz a validação do cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 183285
    
    Exit Sub
    
Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 183285
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183286)

    End Select

End Sub

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 183283

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 183283

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183284)

    End Select
    
    Exit Sub

End Sub

Private Sub FilialCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iFilialCliAlterada = REGISTRO_ALTERADO
End Sub

Private Sub FilialCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FilialCliente_Validate

    'Faz a validação da filial do cliente
    lErro = Valida_FilialCliente()
    If lErro <> SUCESSO Then gError 183287
    
    Exit Sub
    
Erro_FilialCliente_Validate:

    Cancel = True

    Select Case gErr

        Case 183287
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183288)

    End Select

End Sub

Private Function Valida_FilialCliente() As Long
'Faz a validação da filial do cliente

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sCliente As String
Dim objcliente As New ClassCliente

On Error GoTo Erro_Valida_FilialCliente

    'Se a filial de cliente não foi alterada => sai da função
    If iFilialCliAlterada = 0 Then Exit Function
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialCliente.Text)) > 0 Then

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialCliente, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 183289
    
        'Se não encontrou o CÓDIGO
        If lErro = 6730 Then
    
            'Verifica se o cliente foi digitado
            If Len(Trim(Cliente.Text)) = 0 Then gError 183290
    
            sCliente = Cliente.Text
            objFilialCliente.iCodFilial = iCodigo
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 183291
            
            If lErro = 17660 Then
                
                'Lê o Cliente
                objcliente.sNomeReduzido = sCliente
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO And lErro <> 12348 Then gError 183292
                
                'Não encontrou Cliente
                If lErro = 12348 Then gError 183293
                
                objFilialCliente.lCodCliente = objcliente.lCodigo
            
                gError 183294
                
            End If
    
            'Coloca na tela a Filial lida
            FilialCliente.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
    
        'Não encontrou a STRING
        ElseIf lErro = 6731 Then
            gError 183295
    
        End If

    End If
    
    iFilialCliAlterada = 0
    
    Valida_FilialCliente = SUCESSO
    
    Exit Function

Erro_Valida_FilialCliente:

    Valida_FilialCliente = gErr

    Select Case gErr

        Case 183289, 183291, 183292, 183293

        Case 183290
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 183294
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 183295
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialCliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183296)

    End Select

    Exit Function

End Function


Private Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection
Dim sOrdenacao As String

On Error GoTo Erro_LabelCliente_Click

    'Se é possível extrair o código do cliente do conteúdo do controle
    If LCodigo_Extrai(Cliente.Text) <> 0 Then

        'Guarda o código para ser passado para o browser
        objcliente.lCodigo = LCodigo_Extrai(Cliente.Text)

        sOrdenacao = "Codigo"

    'Senão, ou seja, se está digitado o nome do cliente
    Else
        
        'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
        objcliente.sNomeReduzido = Cliente.Text
        
        sOrdenacao = "Nome Reduzido"
    
    End If
    
    'Chama a tela de consulta de cliente
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente, "", sOrdenacao)

    Exit Sub
    
Erro_LabelCliente_Click:

    Select Case gErr
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183312)
    
    End Select
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim lErro As Long

On Error GoTo Erro_objEventoCliente_evSelecao

    Set objcliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objcliente.sNomeReduzido

    'Dispara o Validate de Cliente
    lErro = Valida_Cliente()
    If lErro <> SUCESSO Then gError 183311

    Me.Show

    Exit Sub

Erro_objEventoCliente_evSelecao:

    Select Case gErr
    
        Case 183311
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183313)
    
    End Select

End Sub

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente
    
    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
    If lErro <> SUCESSO Then gError 183269

    lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 183270

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr
    
        Case 183269, 183270
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183271)

    End Select

    Exit Sub

End Sub

Public Sub Filial_FormataCli(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_FormataCli

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 183272

    If lErro = 17660 Then gError 183273

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_FormataCli:

    Select Case gErr

        Case 183272
        
        Case 183273
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183274)

    End Select

    Exit Sub

End Sub

Private Function Valida_Cliente() As Long
'Faz a validação do cliente

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialCliente As New ClassFilialCliente
Dim objRelacionamentoClientes As New ClassRelacClientes

On Error GoTo Erro_Valida_Cliente

    'Se o campo cliente não foi alterado => sai da função
    If iClienteAlterado = 0 Then Exit Function

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 183308

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 183309

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", FilialCliente, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", FilialCliente, iCodFilial)
        
    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        FilialCliente.Clear
        
    End If
    
    iClienteAlterado = 0
    
    Valida_Cliente = SUCESSO

    Exit Function

Erro_Valida_Cliente:

    Valida_Cliente = gErr
    
    Select Case gErr

        Case 183308, 183309
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183310)

    End Select

    Exit Function

End Function
