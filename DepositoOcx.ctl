VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DepositoOcx 
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9495
   KeyPreview      =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   870
      Width           =   8985
      Begin VB.Frame FrameRecibo 
         Caption         =   "Recibo"
         Height          =   705
         Left            =   15
         TabIndex        =   80
         Top             =   4410
         Width           =   2835
         Begin VB.CheckBox CheckImprimirRecibo 
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
            Height          =   270
            Left            =   165
            TabIndex        =   82
            Top             =   285
            Width           =   1935
         End
         Begin VB.CommandButton BotaoImprimirRecibo 
            Height          =   345
            Left            =   2250
            Picture         =   "DepositoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Imprimir"
            Top             =   225
            Width           =   420
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Complemento"
         Height          =   1650
         Left            =   30
         TabIndex        =   39
         Top             =   2655
         Width           =   8940
         Begin VB.ComboBox Historico 
            Height          =   315
            Left            =   1680
            TabIndex        =   9
            Top             =   225
            Width           =   7170
         End
         Begin MSMask.MaskEdBox NumRefExterna 
            Height          =   300
            Left            =   1680
            TabIndex        =   10
            Top             =   660
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ContaContabil 
            Height          =   315
            Left            =   7215
            TabIndex        =   11
            Top             =   660
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Natureza 
            Height          =   315
            Left            =   1680
            TabIndex        =   76
            Top             =   1140
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Ccl 
            Height          =   300
            Left            =   3795
            TabIndex        =   83
            Top             =   660
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            PromptChar      =   " "
         End
         Begin VB.Label CclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Ccl:"
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
            Left            =   3375
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   84
            Top             =   705
            Width           =   345
         End
         Begin VB.Label LabelNaturezaDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3015
            TabIndex        =   78
            Top             =   1140
            Width           =   5850
         End
         Begin VB.Label LabelNatureza 
            AutoSize        =   -1  'True
            Caption         =   "Natureza:"
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
            Left            =   750
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   77
            Top             =   1185
            Width           =   840
         End
         Begin VB.Label ContaContabilLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta Ctb Crédito:"
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
            Left            =   5595
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   72
            ToolTipText     =   "Conta Contábil de Aplicação"
            Top             =   705
            Width           =   1560
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
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
            Left            =   810
            TabIndex        =   45
            Top             =   315
            Width           =   825
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Externo:"
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
            Left            =   450
            TabIndex        =   46
            Top             =   720
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Recebimento"
         Height          =   765
         Left            =   15
         TabIndex        =   44
         Top             =   1815
         Width           =   8955
         Begin VB.ComboBox TipoMeioPagto 
            Height          =   315
            Left            =   1695
            TabIndex        =   8
            Top             =   315
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Forma:"
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
            Left            =   1050
            TabIndex        =   47
            Top             =   375
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados Principais"
         Height          =   1740
         Left            =   15
         TabIndex        =   38
         Top             =   15
         Width           =   8970
         Begin VB.CommandButton BotaoProjetos 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   75
            Top             =   1245
            Width           =   495
         End
         Begin VB.ComboBox Etapa 
            Height          =   315
            Left            =   5175
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1245
            Width           =   2475
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   6285
            Picture         =   "DepositoOcx.ctx":0102
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin VB.ComboBox CodContaCorrente 
            Height          =   315
            Left            =   1785
            TabIndex        =   1
            Top             =   330
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Data 
            Height          =   300
            Left            =   1800
            TabIndex        =   4
            Top             =   810
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Sequencial 
            Height          =   300
            Left            =   5190
            TabIndex        =   2
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#########"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown SpinData 
            Height          =   300
            Left            =   2940
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   810
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Valor 
            Height          =   300
            Left            =   5175
            TabIndex        =   5
            Top             =   810
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Projeto 
            Height          =   300
            Left            =   1800
            TabIndex        =   6
            Top             =   1245
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Etapa:"
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
            Index           =   62
            Left            =   4530
            TabIndex        =   74
            Top             =   1305
            Width           =   570
         End
         Begin VB.Label LabelProjeto 
            AutoSize        =   -1  'True
            Caption         =   "Projeto:"
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
            Left            =   1050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   73
            Top             =   1305
            Width           =   675
         End
         Begin VB.Label LblConta 
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
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
            TabIndex        =   48
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label LabelSequencial 
            AutoSize        =   -1  'True
            Caption         =   "Sequencial:"
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
            Left            =   4050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   49
            Top             =   345
            Width           =   1020
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   1215
            TabIndex        =   50
            Top             =   840
            Width           =   480
         End
         Begin VB.Label Label14 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   4575
            TabIndex        =   51
            Top             =   810
            Width           =   510
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4530
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   870
      Visible         =   0   'False
      Width           =   8985
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4695
         TabIndex        =   79
         Tag             =   "1"
         Top             =   2490
         Width           =   870
      End
      Begin VB.CommandButton CTBBotaoImprimir 
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7710
         TabIndex        =   17
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6300
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   870
         Width           =   2700
      End
      Begin VB.CommandButton CTBBotaoLimparGrid 
         Caption         =   "Limpar Grid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6270
         TabIndex        =   16
         Top             =   30
         Width           =   1245
      End
      Begin VB.CommandButton CTBBotaoModeloPadrao 
         Caption         =   "Modelo Padrão"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6270
         TabIndex        =   18
         Top             =   390
         Width           =   2700
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   3105
         TabIndex        =   28
         Top             =   2490
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   27
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6345
         TabIndex        =   30
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   40
         Top             =   3465
         Width           =   5895
         Begin VB.Label CTBCclLabel 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Custo:"
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
            Left            =   330
            TabIndex        =   52
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label CTBLabel7 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1215
            TabIndex        =   53
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1875
            TabIndex        =   54
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1875
            TabIndex        =   55
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin VB.CheckBox CTBLancAutomatico 
         Caption         =   "Recalcula Automaticamente"
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
         Left            =   3480
         TabIndex        =   21
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   22
         Top             =   1860
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDebito 
         Height          =   225
         Left            =   3435
         TabIndex        =   25
         Top             =   1890
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCredito 
         Height          =   225
         Left            =   2280
         TabIndex        =   24
         Top             =   1830
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBCcl 
         Height          =   225
         Left            =   1545
         TabIndex        =   23
         Top             =   1875
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin MSComCtl2.UpDown CTBUpDown 
         Height          =   300
         Left            =   1650
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   555
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   15
         Top             =   555
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBLote 
         Height          =   300
         Left            =   5565
         TabIndex        =   14
         Top             =   150
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTBDocumento 
         Height          =   300
         Left            =   3780
         TabIndex        =   13
         Top             =   165
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   5
         Mask            =   "#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid CTBGridContabil 
         Height          =   1860
         Left            =   15
         TabIndex        =   29
         Top             =   1215
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3281
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2985
         Left            =   6345
         TabIndex        =   31
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2985
         Left            =   6345
         TabIndex        =   32
         Top             =   1530
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5265
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label CTBLabel1 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   6300
         TabIndex        =   19
         Top             =   660
         Width           =   690
      End
      Begin VB.Label CTBLabel21 
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
         Height          =   255
         Left            =   45
         TabIndex        =   56
         Top             =   210
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   765
         TabIndex        =   57
         Top             =   195
         Width           =   1530
      End
      Begin VB.Label CTBLabel14 
         Caption         =   "Período:"
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
         Left            =   4245
         TabIndex        =   58
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5025
         TabIndex        =   59
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   60
         Top             =   555
         Width           =   1185
      End
      Begin VB.Label CTBLabel13 
         Caption         =   "Exercício:"
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
         Left            =   1995
         TabIndex        =   61
         Top             =   585
         Width           =   870
      End
      Begin VB.Label CTBLabel5 
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
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
         Left            =   45
         TabIndex        =   62
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label CTBLabelHistoricos 
         Caption         =   "Históricos"
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
         Left            =   6345
         TabIndex        =   63
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label CTBLabelContas 
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
         Height          =   255
         Left            =   6345
         TabIndex        =   64
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelCcl 
         Caption         =   "Centros de Custo / Lucro"
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
         Left            =   6345
         TabIndex        =   65
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label CTBLabelTotais 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totais:"
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
         Height          =   225
         Left            =   1800
         TabIndex        =   66
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   67
         Top             =   3075
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   68
         Top             =   3090
         Width           =   1155
      End
      Begin VB.Label CTBLabel8 
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
         Height          =   195
         Left            =   45
         TabIndex        =   69
         Top             =   600
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento:"
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
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   70
         Top             =   210
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   71
         Top             =   210
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7185
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "DepositoOcx.ctx":01EC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "DepositoOcx.ctx":0346
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "DepositoOcx.ctx":04D0
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "DepositoOcx.ctx":0A02
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5715
      Left            =   135
      TabIndex        =   43
      Top             =   405
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   10081
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contabilização"
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
Attribute VB_Name = "DepositoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil
Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoContaCorrenteInt As AdmEvento
Attribute objEventoContaCorrenteInt.VB_VarHelpID = -1

Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

Private Const CONTACORRENTE1 As String = "Conta_Corrente"
Private Const SEQUENCIAL1 As String = "Sequencial"
Private Const VALOR1 As String = "Valor"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const HISTORICO1 As String = "Historico"
Private Const CTACONTABIL_CREDITO As String = "CtaContabil_Cred"
Private Const DOC_EXTERNO As String = "Doc_Externo"
Private Const CTACONTACORRENTE As String = "Cta_Conta_Corrente"

'fim contabilidade

Dim iFrameAtual As Integer
Public iAlterado As Integer

Private WithEvents objEventoCCI As AdmEvento
Attribute objEventoCCI.VB_VarHelpID = -1
Private WithEvents objEventoDeposito As AdmEvento
Attribute objEventoDeposito.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lSeq As Long
Dim iConta As Integer

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 57572
    
    'Pega no itemdata o codigo da conta selecionada
    iConta = Codigo_Extrai(CodContaCorrente.Text)

    'Retorna o proximo sequencial disponivel para a conta em questao
    lErro = CF("CtaCorrente_Sequencial_Transacao", iConta, lSeq)
    If lErro <> SUCESSO Then Error 57547

    Sequencial.Text = CStr(lSeq)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57547
        
        Case 57572
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158899)
    
    End Select

    Exit Sub

End Sub

Private Sub CodContaCorrente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Tela_Extrai

    sTabela = "MovimentosContaCorrente"

    If Len(Trim(CodContaCorrente.Text)) > 0 Then
        objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    Else
        objMovContaCorrente.iCodConta = 0
    End If

    If Len(Trim(Sequencial.Text)) > 0 Then
        objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    Else
        objMovContaCorrente.lSequencial = 0
    End If

    If Len(Trim(Data.ClipText)) > 0 Then objMovContaCorrente.dtDataMovimento = CDate(Data.Text)
    If Len(Trim(Valor.Text)) > 0 Then objMovContaCorrente.dValor = CDbl(Valor.Text)

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo

    colCampoValor.Add "CodConta", objMovContaCorrente.iCodConta, 0, "CodConta"
    colCampoValor.Add "Sequencial", objMovContaCorrente.lSequencial, 0, "Sequencial"
    colCampoValor.Add "NumMovto", 0, 0, "NumMovto"
    
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Tipo", OP_IGUAL, MOVCCI_DEPOSITO
    colSelecao.Add "Excluido", OP_IGUAL, MOVCONTACORRENTE_NAO_EXCLUIDO

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158900)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Tela_Preenche
    
    objMovContaCorrente.iCodConta = colCampoValor.Item("CodConta").vValor
    objMovContaCorrente.lSequencial = colCampoValor.Item("Sequencial").vValor
    
    If objMovContaCorrente.iCodConta > 0 And objMovContaCorrente.lSequencial > 0 Then
        
        'Carrega objMovContaCorrente com os dados passados em colCampoValor
        objMovContaCorrente.lNumMovto = colCampoValor.Item("NumMovto").vValor
        
        lErro = Traz_Deposito_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 49602
   
        iAlterado = 0
    
    End If
    
    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 49602

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158901)

    End Select

    Exit Sub

End Sub

Private Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long
Dim sMascaraConta As String, sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
    If giTipoVersao = VERSAO_LIGHT Then
        Opcao.Visible = False
    End If
    
    Set objEventoCCI = New AdmEvento
    Set objEventoDeposito = New AdmEvento
    Set objEventoContaCorrenteInt = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoCcl = New AdmEvento
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Inicializa propriedade Mask de ContaContabil
    lErro = MascaraConta(sMascaraConta)
    If lErro <> SUCESSO Then gError 64390

    ContaContabil.Mask = sMascaraConta
    
    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 64390

    Ccl.Mask = sMascaraCcl
    
    'Carrega a combo dos Tipos de meio de Pagamento
    lErro = Carrega_TipoMeioPagto()
    If lErro <> SUCESSO Then gError 18050
    
    'Carrega a combo com os codigoa e nomes das contass correntes
    lErro = Carrega_CodContaCorrente()
    If lErro <> SUCESSO Then gError 18051

    'carrega a combo de historico
    lErro = Carrega_Historico()
    If lErro <> SUCESSO Then gError 18052

    iFrameAtual = 1
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_TESOURARIA)
    If lErro <> SUCESSO Then gError 36469
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 122835
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 18050, 18051, 18052, 36469, 64390, 122835
                        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158902)
    
    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Private Sub Historico_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LblConta_Click()
'chama browse de conta corrente

Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim colSelecao As New Collection

    If Len(Trim(CodContaCorrente.Text)) > 0 Then objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)

    Call Chama_Tela("CtaCorrenteLista", colSelecao, objContasCorrentesInternas, objEventoContaCorrenteInt)

End Sub

Private Sub objEventoContaCorrenteInt_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas

    Set objContaCorrenteInt = obj1
    
    CodContaCorrente.Text = objContaCorrenteInt.iCodigo
    Call CodContaCorrente_Validate(bSGECancelDummy)
    
    Me.Show

End Sub

Private Sub NumRefExterna_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCCI_evSelecao(obj1 As Object)

Dim objContaCorrenteInt As ClassContasCorrentesInternas
Dim lErro As Long

On Error GoTo Erro_objEventoCCI_evSelecao

    Set objContaCorrenteInt = obj1

    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 19564

    If lErro = 11807 Then Error 19565
        
    CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    Call CodContaCorrente_Validate(bSGECancelDummy)
    
    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoCCI_evSelecao:

    Select Case Err
    
        Case 19565
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodContaCorrente.Text)
            CodContaCorrente.SetFocus
            
        Case 19564
            CodContaCorrente.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158903)

    End Select

    Exit Sub

End Sub

Private Sub objEventoDeposito_evSelecao(obj1 As Object)

Dim objMovContaCorrente As ClassMovContaCorrente
Dim iIndex As Integer
Dim lErro As Long

On Error GoTo Erro_objEventoDeposito_evSelecao

    Set objMovContaCorrente = obj1
    
    lErro = Traz_Deposito_Tela(objMovContaCorrente)
    If lErro <> SUCESSO Then Error 18233
    
    iAlterado = 0
    
    Me.Show

    Exit Sub
    
Erro_objEventoDeposito_evSelecao:

    Select Case Err
    
        Case 18233
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158904)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Function Carrega_TipoMeioPagto() As Long
'Carrega na Combo TipoMeioPagto os tipo de meio de pagamento ativos

Dim lErro As Long
Dim colTipoMeioPagto As New Collection
Dim objTipoMeioPagto As ClassTipoMeioPagto

On Error GoTo Erro_Carrega_TipoMeioPagto

    'Le todos os tipo de pagamento
    lErro = CF("TipoMeioPagto_Le_Todos", colTipoMeioPagto)
    If lErro <> SUCESSO Then Error 18053

    For Each objTipoMeioPagto In colTipoMeioPagto
        
        'Verifica se estao ativos
        If objTipoMeioPagto.iInativo = TIPOMEIOPAGTO_ATIVO Then

            'coloca na combo
            TipoMeioPagto.AddItem CStr(objTipoMeioPagto.iTipo) & SEPARADOR & objTipoMeioPagto.sDescricao
            TipoMeioPagto.ItemData(TipoMeioPagto.NewIndex) = objTipoMeioPagto.iTipo

        End If

    Next

    Carrega_TipoMeioPagto = SUCESSO

    Exit Function

Erro_Carrega_TipoMeioPagto:

    Carrega_TipoMeioPagto = Err

    Select Case Err

        Case 18053

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158905)

    End Select

    Exit Function

End Function

Private Function Carrega_CodContaCorrente() As Long
'Carrega as contas correntes na combo de contas correntes

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome

On Error GoTo Erro_Carrega_CodContaCorrente

    'Le o nome e o codigo de todas a contas correntes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 18054

    For Each objCodigoNome In colCodigoNomeRed
    
        'Insere na combo de contas correntes
        CodContaCorrente.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        CodContaCorrente.ItemData(CodContaCorrente.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_CodContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_CodContaCorrente:

    Carrega_CodContaCorrente = Err

    Select Case Err

        Case 18054

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158906)

    End Select

    Exit Function

End Function

Private Function Carrega_Historico() As Long
'Carrega a combo de historicos com os historicos da tabela "HistPadraMovConta"

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome

On Error GoTo Erro_Carrega_Historico

    'le o Codigo e a descricao de todos os historicos
    lErro = CF("Cod_Nomes_Le", "HistPadraoMovConta", "Codigo", "Descricao", STRING_NOME, colCodigoNome)
    If lErro <> SUCESSO Then Error 18055

    For Each objCodigoNome In colCodigoNome

        'Insere na Combo de historicos
        Historico.AddItem objCodigoNome.sNome
        Historico.ItemData(Historico.NewIndex) = objCodigoNome.iCodigo

    Next

    Carrega_Historico = SUCESSO

    Exit Function

Erro_Carrega_Historico:

    Carrega_Historico = Err

    Select Case Err

        Case 18055

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158907)

    End Select
    
    Exit Function

End Function

Function Trata_Parametros(Optional objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se algum movimento foi passado por parametro
    If Not (objMovContaCorrente Is Nothing) Then
    
        'Traz os dados do movimento passado por parametro
        lErro = Traz_Deposito_Tela(objMovContaCorrente)
        If lErro <> SUCESSO Then Error 18056
        
    Else
        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 18056
              
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158908)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Traz_Deposito_Tela(objMovContaCorrente As ClassMovContaCorrente) As Long
'Coloca na Tela os dados do Deposito passado como parametro

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sNaturezaEnxuta As String
Dim sCclMascarado As String

On Error GoTo Erro_Traz_Deposito_Tela

    'Le o movimento passado como parametro
    lErro = CF("MovContaCorrente_Le", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 18058

    If lErro = 11893 Then gError 18059
    
    'Verificxa se o mivimento esta cadastrado como sendo um deposito
    If objMovContaCorrente.iTipo <> MOVCCI_DEPOSITO Then gError 18064
    
    'verifica se o movimento está excluído
    If objMovContaCorrente.iExcluido = MOVCONTACORRENTE_EXCLUIDO Then gError 18225
    
    'passa os dados para a Tela
    Historico.Text = objMovContaCorrente.sHistorico
    NumRefExterna.Text = objMovContaCorrente.sNumRefExterna
    Data.Text = Format(objMovContaCorrente.dtDataMovimento, "dd/MM/yy")
    Valor.Text = (objMovContaCorrente.dValor)
    
    lErro = CF("ContaCorrenteInt_Le", objMovContaCorrente.iCodConta, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 18057
    
    If lErro = 11807 Then gError 18062
    
    For iIndice = 0 To CodContaCorrente.ListCount - 1
        If CodContaCorrente.List(iIndice) = CStr(objMovContaCorrente.iCodConta) & SEPARADOR & objContaCorrenteInt.sNomeReduzido Then
            CodContaCorrente.ListIndex = iIndice
            Exit For
        End If
    Next
            
    Sequencial.Text = objMovContaCorrente.lSequencial
            
    objTipoMeioPagto.iTipo = objMovContaCorrente.iTipoMeioPagto
    
    lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
    If lErro <> SUCESSO And lErro <> 11909 Then gError 18061
    
    If lErro = 11909 Then gError 18063
    
    TipoMeioPagto.Text = CStr(objMovContaCorrente.iTipoMeioPagto) & SEPARADOR & objTipoMeioPagto.sDescricao

    If Len(Trim(objMovContaCorrente.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objMovContaCorrente.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    If Len(Trim(objMovContaCorrente.sCcl)) <> 0 Then
    
        sCclMascarado = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCclEnxuta(objMovContaCorrente.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 122833
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    Else
        
        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True
    
    End If
    
    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objMovContaCorrente.lNumMovto)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 36470
    
    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objMovContaCorrente.objProjetoInfo, objMovContaCorrente.lNumMovto, PRJ_CR_TIPO_DEPOSITO)
    If lErro <> SUCESSO Then gError 181676
    '###################################################
    
    Traz_Deposito_Tela = SUCESSO

    Exit Function

Erro_Traz_Deposito_Tela:

    Traz_Deposito_Tela = gErr

    Select Case gErr

        Case 18057, 18058, 18061, 36470, 181676, 122833
        
        Case 18059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", Err)
                
        Case 18062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objMovContaCorrente.iCodConta)

        Case 18063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objMovContaCorrente.iTipoMeioPagto)
        
        Case 18064
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_MOVIMENTO_NAO_DEPOSITO", Err, objMovContaCorrente.lSequencial, objMovContaCorrente.iCodConta)
        
        Case 18225
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVCONTACORRENTE_EXCLUIDO", Err, objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158909)

    End Select

    Exit Function

End Function

Private Sub Sequencial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Sequencial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)

End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)

On Error GoTo Erro_Sequencial_Validate

    'Verifica preenchimento do sequencial
    If Len(Trim(Sequencial.Text)) > 0 Then

        'Verifica se o sequencial é numérico
        If Not IsNumeric(Sequencial.Text) Then Error 55958

        'Verifica se codigo é menor que um
        If CLng(Sequencial.Text) < 1 Then Error 55959

    End If

    Exit Sub

Erro_Sequencial_Validate:

    Cancel = True

    Select Case Err

        Case 55958, 55959
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_INVALIDO", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158910)

    End Select

    Exit Sub

End Sub

Private Sub TipoMeioPagto_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoMeioPagto_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMeioPagto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_TipoMeioPagto_Validate

    'verifica se foi preenchido o TipoMeioPagto
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox TipoMeioPagto
    If TipoMeioPagto.Text = TipoMeioPagto.List(TipoMeioPagto.ListIndex) Then Exit Sub
    
    If iAlterado = REGISTRO_ALTERADO Then
    
        'Tenta selecionar o TipoMeioPagto com o codigo digitado
        lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18124
    
        If lErro = 6730 Then
    
            objTipoMeioPagto.iTipo = iCodigo
        
            'Pesquisa no BD a existencia do tipo passado por parametro
            lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
            If lErro <> SUCESSO And lErro <> 11909 Then Error 18125
            
            'Se existir coloca na Tela senão  ---> Erro
            If lErro = 11909 Then Error 18126
                
            TipoMeioPagto.Text = Format(objTipoMeioPagto.iTipo, "@@@@") & SEPARADOR & objTipoMeioPagto.sDescricao
        End If
    Else
    
        'Tenta selecionar o TipoMeioPagto com o codigo digitado
        lErro = Combo_Seleciona(TipoMeioPagto, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 40765
    
        If lErro = 6730 Then
    
            objTipoMeioPagto.iTipo = iCodigo
        
            'Pesquisa no BD a existencia do tipo passado por parametro
            lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
            If lErro <> SUCESSO And lErro <> 11909 Then Error 40766
            
            'Se existir coloca na Tela senão  ---> Erro
            If lErro = 11909 Then Error 40767
                
            TipoMeioPagto.Text = Format(objTipoMeioPagto.iTipo, "@@@@") & SEPARADOR & objTipoMeioPagto.sDescricao
        End If
        iAlterado = 0
    End If
    
    If lErro = 6731 Then Error 18123
    
    Exit Sub

Erro_TipoMeioPagto_Validate:

    Cancel = True


    Select Case Err

        Case 18123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, TipoMeioPagto.Text)
            
        Case 18124, 18125, 40765, 40766
        
        Case 18126, 40767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158911)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SpinData_DownClick()
Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_DownClick

    Data.SetFocus

    'Verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        'Diminui a data
        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then Error 18070

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_DownClick:

    Select Case Err

        Case 18070

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158912)

    End Select

    Exit Sub

End Sub

Private Sub SpinData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_SpinData_UpClick

    Data.SetFocus

    'verifica se a data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text
        
        'Aumenta a data
        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then Error 18071

        Data.PromptInclude = False
        Data.Text = sData
        Data.PromptInclude = True

    End If

    Exit Sub

Erro_SpinData_UpClick:

    Select Case Err

        Case 18071

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158913)

    End Select

    Exit Sub

End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub
    
    'verifica se a data é válida
    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then Error 18073

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case Err
        
        Case 18073

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158914)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Verifica se há um valor digitado
    If Len(Trim(Valor.Text)) > 0 Then
    
        'Critiva o valor digitado
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 18074
                
    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 18074

        Case 18080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NEGATIVO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158915)

    End Select

    Exit Sub

End Sub

Private Sub Historico_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iTamanho As Integer
Dim iCodigo As Integer
Dim iIndice As Integer
Dim sDescricao As Long
Dim objHistMovCta As New ClassHistMovCta

On Error GoTo Erro_Historico_Validate

    'Verifica o tamanho do texto em historico
    iTamanho = Len(Trim(Historico.Text))

    If iTamanho = 0 Then Exit Sub
    
    'Verifica se é maior que o tamanho maximo
    If iTamanho > STRING_HISTORICOMOVCONTA Then Error 18128

    'Verifica se o que foi digitado é numerico
    If Not (IsNumeric(Trim(Historico.Text))) Then Exit Sub
    
   lErro = Valor_Inteiro_Critica(Trim(Historico.Text))
   If lErro <> SUCESSO Then Error 40710
    
   objHistMovCta.iCodigo = CInt(Trim(Historico.Text))
                
   lErro = CF("HistMovCta_Le", objHistMovCta)
   If lErro <> SUCESSO And lErro <> 15011 Then Error 40737
    
   If lErro = 15011 Then Error 40742
        
   Historico.Text = objHistMovCta.sDescricao
      
   Exit Sub

Erro_Historico_Validate:

    Cancel = True


    Select Case Err

        Case 18128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_HISTORICOMOVCONTA", Err)
            
        Case 40710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INTEIRO", Err, Historico.Text)
                
        Case 40737
        
        Case 40742
            lErro = Rotina_Erro(vbOKOnly, "ERRO_HISTMOVCTA_NAO_CADASTRADO", Err, objHistMovCta.iCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158916)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_Deposito()

Dim lErro As Long

    Call Limpa_Tela(Me)

    Sequencial.Text = ""
    CodContaCorrente.Text = ""
    TipoMeioPagto.Text = ""
    Historico.Text = ""
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    
    ContaContabil.PromptInclude = False
    ContaContabil.Text = ""
    ContaContabil.PromptInclude = True
    
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValor As Double
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 18076

    If Len(Trim(Sequencial.Text)) = 0 Then gError 18077

    If Len(Trim(Data.ClipText)) = 0 Then gError 18078

    If Len(Trim(Valor.Text)) = 0 Then gError 18079

    If Len(Trim(TipoMeioPagto.Text)) = 0 Then gError 18227

    'Passa os dados da Tela para objMovcontacorrente
    lErro = Move_Tela_Memoria(objMovContaCorrente)
    If lErro <> SUCESSO Then gError 18081

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 20831

    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    '###############################################

    'Rotina encarregada de gravar o Deposito
    lErro = CF("MovCCI_Grava_Deposito", objMovContaCorrente, objContabil)
    If lErro <> SUCESSO Then gError 18082

    If CheckImprimirRecibo.Value = vbChecked Then
        
        lErro = ImprimirRecibo(objMovContaCorrente)
        If lErro <> SUCESSO Then gError 184451
    
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = gErr

    Select Case gErr

        Case 18076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 18077
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 18078
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 18079
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DEPOSITO_NAO_PREENCHIDO", gErr)

        Case 18081, 18082, 20831, 181786, 184451

        Case 18227
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158917)

    End Select

    Exit Function
    
End Function

Function Move_Tela_Memoria(objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objProjetoInfo As Object
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim sCclFormatada As String, iCclPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Move os dados da tela para objmovcontacorrente
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    objMovContaCorrente.dtDataMovimento = CDate(Data.Text)
    objMovContaCorrente.dValor = CDbl(Valor.Text)
    If Len(Trim(TipoMeioPagto.Text)) = 0 Then Error 18248
    objMovContaCorrente.iTipoMeioPagto = Codigo_Extrai(TipoMeioPagto.Text)
    objMovContaCorrente.sHistorico = Historico.Text
    objMovContaCorrente.sNumRefExterna = NumRefExterna.Text
    objMovContaCorrente.iFilialEmpresa = giFilialEmpresa
    objMovContaCorrente.iExcluido = MOVCONTACORRENTE_NAO_EXCLUIDO
    objMovContaCorrente.iPortador = 0
    objMovContaCorrente.iConciliado = 0
    objMovContaCorrente.iFavorecido = 0
    objMovContaCorrente.lNumRefInterna = 0
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 122825

    If iCclPreenchida = CCL_PREENCHIDA Then
        objMovContaCorrente.sCcl = sCclFormatada
    Else
        objMovContaCorrente.sCcl = ""
    End If
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_DEPOSITO)
    If lErro <> SUCESSO Then gError 181761
    
    Set objMovContaCorrente.objProjetoInfo = objProjetoInfo
    '##################################################
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 122825
    
    objMovContaCorrente.sNatureza = sNaturezaFormatada
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 18248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_NAO_INFORMADO", gErr)
            
        Case 181761
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158918)

    End Select

    Exit Function

End Function

Private Sub LabelSequencial_Click()

Dim objMovContaCorrente As New ClassMovContaCorrente
Dim colSelecao As New Collection

    If Len(Trim(CodContaCorrente.Text)) = 0 Then
        Set objMovContaCorrente = Nothing
        Call Chama_Tela("DepositoLista", colSelecao, objMovContaCorrente, objEventoDeposito)
    
    Else
        colSelecao.Add Codigo_Extrai(CodContaCorrente.Text)
        
        If Len(Sequencial.Text) = 0 Then
            objMovContaCorrente.lSequencial = 0
        Else
            objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
        End If
        
        Call Chama_Tela("DepositoLista_Conta", colSelecao, objMovContaCorrente, objEventoDeposito)
    End If
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro  As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a contaCorrente foi informada
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Error 18112
    
    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)

    'Verifica se o sequencial foi informado
    If Len(Trim(Sequencial.Text)) = 0 Then Error 18113
    
    objMovContaCorrente.lSequencial = CLng(Sequencial)
    
    'Pede a confirmacao da exclusao
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DEPOSITO", objMovContaCorrente.iCodConta, objMovContaCorrente.lSequencial)

    If vbMsgRes = vbYes Then
    
        objMovContaCorrente.iTipo = MOVCCI_DEPOSITO

        'Chama a rotina de exclusao (contabilidade)
        lErro = CF("MovCCI_Exclui", objMovContaCorrente, objContabil)
        If lErro <> SUCESSO Then Error 18114

        Call Limpa_Tela_Deposito

        iAlterado = 0

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 18112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", Err)
        
        Case 18113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", Err)
        
        Case 18114

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158919)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18115
    
    Call Limpa_Tela_Deposito

    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 18115
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158920)

    End Select
    
    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18111

    'Limpa a tela
    Call Limpa_Tela_Deposito

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 18111

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158921)

    End Select

    Exit Sub

End Sub

Private Sub CodContaCorrente_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_CodContaCorrente_Validate
    
    If Len(Trim(CodContaCorrente.Text)) = 0 Then Exit Sub

    'verifica se esta preenchida com o item selecionado na ComboBox CodContacOrrente
    If CodContaCorrente.Text = CodContaCorrente.List(CodContaCorrente.ListIndex) Then Exit Sub

    'Tenta selecionar a conta corrente na combo
    lErro = Combo_Seleciona(CodContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 18120

    If lErro = 6730 Then
    
        'Pega o codigo que estana combo
        objContaCorrenteInt.iCodigo = iCodigo
        
        'Procura no BD
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 18121
    
        'Se nao estiver cadastrada --> Erro
        If lErro = 11807 Then Error 18122
        
        'Se alguma Filial tiver sido selecionada
        If giFilialEmpresa <> EMPRESA_TODA Then

            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43533

        End If
        
        'Se estiver cadastrada põe na tela
        CodContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    ElseIf lErro = 6731 Then
    
        Error 18119

    End If
    
    Exit Sub

Erro_CodContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 18119
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, CodContaCorrente.Text)
            
        Case 18120, 18121
        
        Case 18122
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)
        
            If vbMsgRes = vbYes Then
                'Lembrar de manter na tela o numero passado como parametro
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
                
            Else
            End If
        
        Case 43533
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_PERTENCE_FILIAL", Err, CodContaCorrente.Text, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158922)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_UnLoad(Cancel As Integer)
    
Dim lErro As Long

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set objEventoCCI = Nothing
    Set objEventoDeposito = Nothing
    Set objEventoNatureza = Nothing
    Set objEventoCcl = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoContaCorrenteInt = Nothing
    
End Sub

Private Sub Historico_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_DEPOSITO_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_DEPOSITO_CONTABILIZACAO
                        
        End Select
    
    End If

End Sub

'Inclusão de Código 12/05/00 Raphael
Private Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sContaFormatada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim sContaMascarada As String

On Error GoTo Erro_ContaContabil_Validate

    'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
    lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_TESOURARIA)
    If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 64377

    If lErro = SUCESSO Then

        sContaFormatada = objPlanoConta.sConta

        'mascara a conta
        sContaMascarada = String(STRING_CONTA, 0)

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
        If lErro <> SUCESSO Then gError 64378

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaMascarada
        ContaContabil.PromptInclude = True


    'Se não encontrou a conta simples
    ElseIf lErro = 44096 Or lErro = 44098 Then

        'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
        lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_TESOURARIA)
        If lErro <> SUCESSO And lErro <> 5700 Then gError 64379

        'Conta não cadastrada
        If lErro = 5700 Then gError 64380

    End If

    Exit Sub

Erro_ContaContabil_Validate:

    Cancel = True


    Select Case gErr

        Case 64377, 64379
    
        Case 64378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            
        Case 64380
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_INEXISTENTE", gErr, ContaContabil.Text)
            
            'Pergunta se o usuário deseja cadastrar a conta
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            'Se sim
            If vbMsgRes = vbYes Then
                'Preenche o parâmetro que será passado para a tela
                objPlanoConta.sConta = sContaFormatada
                'Chama a tela de plano de contas
                Call Chama_Tela("PlanoConta", objPlanoConta)
            End If
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158923)
    
    End Select

    Exit Sub
    
End Sub

Public Sub ContaContabilLabel_Click()
'BROWSE PLANO_CONTA :

Dim lErro As Long
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection
Dim iContaPreenchida As Integer
Dim sConta As String

On Error GoTo Erro_ContaContabilLabel_Click

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 64381

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    Call Chama_Tela("PlanoContaTESLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_ContaContabilLabel_Click:

    Select Case gErr

        Case 64381
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158924)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta = "" Then
        ContaContabil.Text = ""
    Else
        ContaContabil.PromptInclude = False

        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 64382

        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 64382
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158925)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 36471

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 36472

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 36471

        Case 36472
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Private Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Private Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Private Sub CTBGridContabil_LeaveCell()

    Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Private Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Private Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Private Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Private Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Private Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Private Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Private Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Private Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Private Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Private Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Private Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Private Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Private Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Private Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Private Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Private Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Private Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub


Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Private Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Private Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Private Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Private Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Private Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Private Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Private Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Private Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Private Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Private Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Private Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Private Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Private Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Private Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objContasCorrentesInternas As New ClassContasCorrentesInternas
Dim sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case DOC_EXTERNO
            If Len(Trim(NumRefExterna.Text)) > 0 Then
                objMnemonicoValor.colValor.Add NumRefExterna.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CTACONTACORRENTE
            If Len(CodContaCorrente.Text) > 0 Then
                
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64402
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64403
                
                If objContasCorrentesInternas.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContasCorrentesInternas.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 64445
                    
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case CONTACORRENTE1
            If Len(CodContaCorrente.Text) > 0 Then
            
                objContasCorrentesInternas.iCodigo = Codigo_Extrai(CodContaCorrente.Text)
                
                'Procura a conta no BD
                lErro = CF("ContaCorrenteInt_Le", objContasCorrentesInternas.iCodigo, objContasCorrentesInternas)
                If lErro <> SUCESSO And lErro <> 11807 Then gError 64404
            
                'Se nao estiver cadastrada --> Erro
                If lErro = 11807 Then gError 64405
                
                objMnemonicoValor.colValor.Add objContasCorrentesInternas.sNomeReduzido
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case SEQUENCIAL1
            If Len(Sequencial.Text) > 0 Then
                objMnemonicoValor.colValor.Add CLng(Sequencial.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR1
            If Len(Valor.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Valor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case FORMA1
            If Len(TipoMeioPagto.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoMeioPagto.ItemData(TipoMeioPagto.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case CTACONTABIL_CREDITO
            If Len(ContaContabil.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add ContaContabil.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case Else
            gError 39547

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39547
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 64402, 64404, 64445
        
        Case 64403, 64405
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContasCorrentesInternas.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158926)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_DEPOSITO_ID
    Set Form_Load_Ocx = Me
    Caption = "Depósito"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Deposito"
    
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
        
        If Me.ActiveControl Is Sequencial Then
            Call LabelSequencial_Click
        ElseIf Me.ActiveControl Is CodContaCorrente Then
            Call LblConta_Click
        ElseIf Me.ActiveControl Is ContaContabil Then
            Call ContaContabilLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        End If
    
    End If
    
End Sub



Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub LblConta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblConta, Source, X, Y)
End Sub

Private Sub LblConta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblConta, Button, Shift, X, Y)
End Sub

Private Sub LabelSequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelSequencial, Source, X, Y)
End Sub

Private Sub LabelSequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelSequencial, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub CTBCclLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclLabel, Source, X, Y)
End Sub

Private Sub CTBCclLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclLabel, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel7, Source, X, Y)
End Sub

Private Sub CTBLabel7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel7, Button, Shift, X, Y)
End Sub

Private Sub CTBContaDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBContaDescricao, Source, X, Y)
End Sub

Private Sub CTBContaDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBContaDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBCclDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBCclDescricao, Source, X, Y)
End Sub

Private Sub CTBCclDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBCclDescricao, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel21, Source, X, Y)
End Sub

Private Sub CTBLabel21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel21, Button, Shift, X, Y)
End Sub

Private Sub CTBOrigem_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBOrigem, Source, X, Y)
End Sub

Private Sub CTBOrigem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBOrigem, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel14, Source, X, Y)
End Sub

Private Sub CTBLabel14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel14, Button, Shift, X, Y)
End Sub

Private Sub CTBPeriodo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBPeriodo, Source, X, Y)
End Sub

Private Sub CTBPeriodo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBPeriodo, Button, Shift, X, Y)
End Sub

Private Sub CTBExercicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBExercicio, Source, X, Y)
End Sub

Private Sub CTBExercicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBExercicio, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel13, Source, X, Y)
End Sub

Private Sub CTBLabel13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel13, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel5, Source, X, Y)
End Sub

Private Sub CTBLabel5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel5, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelHistoricos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelHistoricos, Source, X, Y)
End Sub

Private Sub CTBLabelHistoricos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelHistoricos, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelContas_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelContas, Source, X, Y)
End Sub

Private Sub CTBLabelContas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelContas, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelCcl_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelCcl, Source, X, Y)
End Sub

Private Sub CTBLabelCcl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelCcl, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel1, Source, X, Y)
End Sub

Private Sub CTBLabel1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel1, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelTotais_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelTotais, Source, X, Y)
End Sub

Private Sub CTBLabelTotais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelTotais, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalDebito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalDebito, Source, X, Y)
End Sub

Private Sub CTBTotalDebito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalDebito, Button, Shift, X, Y)
End Sub

Private Sub CTBTotalCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBTotalCredito, Source, X, Y)
End Sub

Private Sub CTBTotalCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBTotalCredito, Button, Shift, X, Y)
End Sub

Private Sub CTBLabel8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabel8, Source, X, Y)
End Sub

Private Sub CTBLabel8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabel8, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelDoc_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelDoc, Source, X, Y)
End Sub

Private Sub CTBLabelDoc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelDoc, Button, Shift, X, Y)
End Sub

Private Sub CTBLabelLote_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CTBLabelLote, Source, X, Y)
End Sub

Private Sub CTBLabelLote_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CTBLabelLote, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub


Private Sub ContaContabilLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaContabilLabel, Source, X, Y)
End Sub

Private Sub ContaContabilLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaContabilLabel, Button, Shift, X, Y)
End Sub

'#####################################
'Inserido por Wagner 03/08/2006
Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "GR"
    objItensPRJCR.sDescricao = "Geral"

    objItensPRJCR.dValorOriginal = StrParaDbl(Valor.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_DEPOSITO
    
    colItensPRJ.Add objItensPRJCR
    
    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
End Function
'#####################################

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Private Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_RECEBIMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub

Private Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_RECEBIMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Private Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Sub BotaoImprimirRecibo_Click()

Dim lErro As Long, lNumIntRel As Long
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_BotaoImprimirRecibo_Click
    
    'Verifica se os campos essencias da tela foram preenchidos
    If Len(Trim(CodContaCorrente.Text)) = 0 Then gError 11912
    If Len(Trim(Sequencial.Text)) = 0 Then gError 11913

    objMovContaCorrente.iCodConta = Codigo_Extrai(CodContaCorrente.Text)
    objMovContaCorrente.lSequencial = CLng(Sequencial.Text)
    
    lErro = CF("MovContaCorrente_ObterNumMovto", objMovContaCorrente)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 184443
    If lErro <> SUCESSO Then gError 184444
    
    lErro = ImprimirRecibo(objMovContaCorrente)
    If lErro <> SUCESSO Then gError 184449
    
    Exit Sub
    
Erro_BotaoImprimirRecibo_Click:
    
    Select Case gErr

        Case 11912
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_NAO_INFORMADA", gErr)

        Case 11913
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)
    
        Case 184443, 184449
        
        Case 184444
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184445)

    End Select

    Exit Sub
    
End Sub

Private Function ImprimirRecibo(ByVal objMovContaCorrente As ClassMovContaCorrente) As Long

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_ImprimirRecibo

    lErro = CF("RelReciboContabil_Prepara", objMovContaCorrente.lNumMovto, lNumIntRel)
    If lErro <> SUCESSO Then gError 184446
    
    lErro = objRelatorio.ExecutarDireto("Recibos Com Contabilização", "", 1, "", "NNUMINTREL", CStr(lNumIntRel))
    If lErro <> SUCESSO Then gError 184447
    
    ImprimirRecibo = SUCESSO
    
    Exit Function
    
Erro_ImprimirRecibo:

    ImprimirRecibo = gErr

    Select Case gErr

        Case 184446, 184447
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184448)

    End Select
    
    Exit Function

End Function

Public Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela_Modal("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'Preenche Ccl

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175259)

    End Select

    Exit Sub

End Sub

Private Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(Ccl.Text)) <> 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 31558
    
        If lErro = 5703 Then gError 31559
        
    End If

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True

    Select Case gErr

        Case 31558

        Case 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175260)

    End Select

    Exit Sub

End Sub
