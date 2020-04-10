VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ConhecimentoFreteOcx 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9480
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   5
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4800
         TabIndex        =   193
         Tag             =   "1"
         Top             =   2520
         Width           =   870
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
         Left            =   6360
         TabIndex        =   128
         Top             =   405
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
         Height          =   270
         Left            =   6360
         TabIndex        =   126
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   930
         Width           =   2700
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
         Height          =   270
         Left            =   7815
         TabIndex        =   127
         Top             =   90
         Width           =   1245
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
         Left            =   3465
         TabIndex        =   116
         Top             =   915
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   115
         Top             =   2025
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   114
         Top             =   1635
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2790
         Left            =   6375
         TabIndex        =   113
         Top             =   1590
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   108
         Top             =   3495
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
            Height          =   195
            Left            =   240
            TabIndex        =   112
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
            Height          =   195
            Left            =   1125
            TabIndex        =   111
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   110
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   109
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4665
         TabIndex        =   117
         Top             =   1335
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
         TabIndex        =   118
         Top             =   1320
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
         TabIndex        =   119
         Top             =   1350
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
         TabIndex        =   120
         Top             =   1290
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
         TabIndex        =   121
         Top             =   1335
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
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   570
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   585
         TabIndex        =   125
         Top             =   570
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
         Left            =   5580
         TabIndex        =   124
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
         Left            =   3795
         TabIndex        =   123
         Top             =   135
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
         Left            =   0
         TabIndex        =   131
         Top             =   1185
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
         Left            =   6360
         TabIndex        =   132
         Top             =   1575
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
         TabIndex        =   133
         Top             =   1560
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
         Left            =   6420
         TabIndex        =   129
         Top             =   720
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
         Height          =   255
         Left            =   45
         TabIndex        =   149
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   795
         TabIndex        =   148
         Top             =   150
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
         Left            =   4230
         TabIndex        =   147
         Top             =   615
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   146
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   145
         Top             =   570
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
         TabIndex        =   144
         Top             =   600
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
         TabIndex        =   143
         Top             =   960
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
         TabIndex        =   142
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
         TabIndex        =   141
         Top             =   1275
         Width           =   2340
      End
      Begin VB.Label CTBLabelTotais 
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
         Height          =   225
         Left            =   1800
         TabIndex        =   140
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   139
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   138
         Top             =   3045
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
         TabIndex        =   137
         Top             =   570
         Width           =   480
      End
      Begin VB.Label CTBLabelDoc 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   136
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label CTBLabelLote 
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
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   135
         Top             =   180
         Width           =   450
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
         Left            =   6360
         TabIndex        =   134
         Top             =   1290
         Visible         =   0   'False
         Width           =   2490
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   9120
      Begin VB.Frame FrameEndereco 
         Caption         =   "Dados Destinatário"
         Height          =   2175
         Index           =   1
         Left            =   225
         TabIndex        =   58
         Top             =   2475
         Width           =   8595
         Begin VB.CommandButton BotaoLimpaDestinatario 
            Height          =   330
            Left            =   8055
            Picture         =   "ConhecimentoFreteOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   156
            ToolTipText     =   "Limpar"
            Top             =   315
            Width           =   390
         End
         Begin VB.ComboBox UFDestinatario 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4485
            TabIndex        =   66
            Top             =   1290
            Width           =   630
         End
         Begin VB.TextBox EnderecoDestinatario 
            Height          =   315
            Left            =   1440
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   810
            Width           =   6345
         End
         Begin MSMask.MaskEdBox CidadeDestinatario 
            Height          =   315
            Left            =   1440
            TabIndex        =   64
            Top             =   1290
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPDestinatario 
            Height          =   315
            Left            =   6795
            TabIndex        =   68
            Top             =   1290
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGCDestinatario 
            Height          =   315
            Left            =   1455
            TabIndex        =   70
            Top             =   1755
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Destinatario 
            Height          =   315
            Left            =   1440
            TabIndex        =   60
            Top             =   330
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscEstDestinatario 
            Height          =   315
            Left            =   5445
            TabIndex        =   72
            Top             =   1755
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   465
            TabIndex        =   61
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   690
            TabIndex        =   63
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   3720
            TabIndex        =   65
            Top             =   1350
            Width           =   675
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CGC/CPF:"
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
            Left            =   495
            TabIndex        =   69
            Top             =   1815
            Width           =   885
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual:"
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
            Left            =   4125
            TabIndex        =   71
            Top             =   1815
            Width           =   1290
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6330
            TabIndex        =   67
            Top             =   1350
            Width           =   465
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Destinatário:"
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
            Index           =   3
            Left            =   285
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame FrameEndereco 
         Caption         =   "Dados Remetente"
         Height          =   2280
         Index           =   0
         Left            =   225
         TabIndex        =   43
         Top             =   135
         Width           =   8595
         Begin VB.CommandButton BotaoLimpaRemetente 
            Height          =   330
            Left            =   8055
            Picture         =   "ConhecimentoFreteOcx.ctx":0532
            Style           =   1  'Graphical
            TabIndex        =   155
            ToolTipText     =   "Limpar"
            Top             =   315
            Width           =   390
         End
         Begin VB.TextBox EnderecoRemetente 
            Height          =   315
            Left            =   1440
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   825
            Width           =   6345
         End
         Begin VB.ComboBox UFRemetente 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4485
            TabIndex        =   51
            Top             =   1320
            Width           =   630
         End
         Begin MSMask.MaskEdBox CidadeRemetente 
            Height          =   315
            Left            =   1440
            TabIndex        =   49
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEPRemetente 
            Height          =   315
            Left            =   6810
            TabIndex        =   53
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGCRemetente 
            Height          =   315
            Left            =   1440
            TabIndex        =   55
            Top             =   1800
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Remetente 
            Height          =   315
            Left            =   1440
            TabIndex        =   45
            Top             =   330
            Width           =   6330
            _ExtentX        =   11165
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscEstRemetente 
            Height          =   315
            Left            =   5460
            TabIndex        =   57
            Top             =   1800
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Remetente:"
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
            Left            =   390
            TabIndex        =   44
            Top             =   390
            Width           =   990
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            Caption         =   "CEP:"
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
            Left            =   6330
            TabIndex        =   52
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            Caption         =   "Insc. Estadual:"
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
            Left            =   4140
            TabIndex        =   56
            Top             =   1860
            Width           =   1290
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF:"
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
            Left            =   375
            TabIndex        =   54
            Top             =   1845
            Width           =   990
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
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
            Left            =   3720
            TabIndex        =   50
            Top             =   1380
            Width           =   675
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Cidade:"
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
            Left            =   690
            TabIndex        =   48
            Top             =   1380
            Width           =   690
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            Caption         =   "Endereço:"
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
            Left            =   465
            TabIndex        =   46
            Top             =   885
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4815
      Index           =   4
      Left            =   135
      TabIndex        =   165
      Top             =   750
      Visible         =   0   'False
      Width           =   9195
      Begin VB.CheckBox ComissaoAutomatica 
         Caption         =   "Calcula comissão automaticamente"
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
         Left            =   270
         TabIndex        =   185
         Top             =   150
         Value           =   1  'Checked
         Width           =   3360
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Comissões"
         Height          =   3870
         Index           =   0
         Left            =   90
         TabIndex        =   166
         Top             =   495
         Width           =   9045
         Begin VB.ComboBox DiretoIndireto 
            Height          =   315
            ItemData        =   "ConhecimentoFreteOcx.ctx":0A64
            Left            =   6600
            List            =   "ConhecimentoFreteOcx.ctx":0A6E
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Frame SSFrame4 
            Caption         =   "Totais - Comissões"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   1
            Left            =   120
            TabIndex        =   168
            Top             =   2910
            Width           =   6855
            Begin VB.Label TotalValorBase 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   1200
               TabIndex        =   174
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total:"
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
               Index           =   18
               Left            =   120
               TabIndex        =   173
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label TotalValorComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   5520
               TabIndex        =   172
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label TotalPercentualComissao 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Left            =   3840
               TabIndex        =   171
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Percentual:"
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
               Index           =   19
               Left            =   2760
               TabIndex        =   170
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   20
               Left            =   4920
               TabIndex        =   169
               Top             =   360
               Width           =   615
            End
         End
         Begin VB.CommandButton BotaoVendedores 
            Caption         =   "Vendedores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   7560
            Picture         =   "ConhecimentoFreteOcx.ctx":0A84
            Style           =   1  'Graphical
            TabIndex        =   167
            Top             =   3030
            Width           =   1380
         End
         Begin MSMask.MaskEdBox ValorComissao 
            Height          =   225
            Left            =   3675
            TabIndex        =   176
            Top             =   360
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox ValorBase 
            Height          =   225
            Left            =   2490
            TabIndex        =   177
            Top             =   375
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox PercentualComissao 
            Height          =   225
            Left            =   1800
            TabIndex        =   178
            Top             =   390
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   225
            Left            =   435
            TabIndex        =   179
            Top             =   375
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorEmissao 
            Height          =   225
            Left            =   5505
            TabIndex        =   180
            Top             =   390
            Width           =   1080
            _ExtentX        =   1905
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
         Begin MSMask.MaskEdBox PercentualEmissao 
            Height          =   225
            Left            =   4830
            TabIndex        =   181
            Top             =   375
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorBaixa 
            Height          =   225
            Left            =   7365
            TabIndex        =   182
            Top             =   345
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
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
         Begin MSMask.MaskEdBox PercentualBaixa 
            Height          =   225
            Left            =   6675
            TabIndex        =   183
            Top             =   360
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissoes 
            Height          =   1845
            Left            =   75
            TabIndex        =   184
            Top             =   240
            Width           =   8910
            _ExtentX        =   15716
            _ExtentY        =   3254
            _Version        =   393216
            Rows            =   7
            Cols            =   5
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9195
      Begin VB.Frame Frame14 
         Caption         =   "Identificação"
         Height          =   2385
         Left            =   225
         TabIndex        =   5
         Top             =   165
         Width           =   8820
         Begin VB.CommandButton BotaoLimparNF 
            Height          =   300
            Left            =   5895
            Picture         =   "ConhecimentoFreteOcx.ctx":102E
            Style           =   1  'Graphical
            TabIndex        =   154
            ToolTipText     =   "Limpar o Número"
            Top             =   840
            Width           =   345
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5160
            TabIndex        =   17
            Top             =   1410
            Width           =   1860
         End
         Begin VB.ComboBox Serie 
            Height          =   315
            Left            =   1995
            TabIndex        =   11
            Top             =   825
            Width           =   765
         End
         Begin MSMask.MaskEdBox NatOpInterna 
            Height          =   300
            Left            =   1995
            TabIndex        =   7
            Top             =   307
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   3090
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1935
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   1995
            TabIndex        =   19
            Top             =   1935
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1995
            TabIndex        =   15
            Top             =   1410
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Index           =   0
            Left            =   4620
            TabIndex        =   16
            Top             =   1440
            Width           =   465
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
            Left            =   1260
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   14
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data de Emissão:"
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
            Index           =   2
            Left            =   420
            TabIndex        =   18
            Top             =   1995
            Width           =   1500
         End
         Begin VB.Label NFiscalLabel 
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
            Height          =   255
            Left            =   4410
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   12
            Top             =   870
            Width           =   720
         End
         Begin VB.Label SerieLabel 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1410
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   10
            Top             =   900
            Width           =   510
         End
         Begin VB.Label LblNatOpInterna 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Operação:"
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
            Left            =   195
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   6
            Top             =   360
            Width           =   1725
         End
         Begin VB.Label NFiscal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5160
            TabIndex        =   13
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Status 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5160
            TabIndex        =   9
            Top             =   307
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Left            =   4500
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Composição do Frete"
         Height          =   2160
         Left            =   210
         TabIndex        =   21
         Top             =   2610
         Width           =   8820
         Begin VB.CheckBox PedagioIncluso 
            Caption         =   "Incluir Pedágio na Base Cálculo"
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
            Left            =   2070
            TabIndex        =   162
            Top             =   435
            UseMaskColor    =   -1  'True
            Value           =   1  'Checked
            Width           =   3045
         End
         Begin VB.Frame SSFrame6 
            Caption         =   "INSS"
            Height          =   555
            Left            =   5370
            TabIndex        =   158
            Top             =   165
            Width           =   3150
            Begin VB.CheckBox INSSRetido 
               Caption         =   "Retido"
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
               Left            =   2100
               TabIndex        =   159
               Top             =   225
               Width           =   900
            End
            Begin MSMask.MaskEdBox ValorINSS 
               Height          =   300
               Left            =   705
               TabIndex        =   160
               Top             =   195
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   15
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label30 
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
               Height          =   210
               Index           =   14
               Left            =   135
               TabIndex        =   161
               Top             =   240
               Width           =   510
            End
         End
         Begin VB.CheckBox ICMSIncluso 
            Caption         =   "ICMS Incluso"
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
            Left            =   255
            TabIndex        =   22
            Top             =   420
            UseMaskColor    =   -1  'True
            Width           =   1560
         End
         Begin MSMask.MaskEdBox FretePeso 
            Height          =   300
            Left            =   240
            TabIndex        =   24
            Top             =   1140
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SEC 
            Height          =   300
            Left            =   3645
            TabIndex        =   28
            Top             =   1140
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Despacho 
            Height          =   300
            Left            =   5355
            TabIndex        =   30
            Top             =   1140
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Pedagio 
            Height          =   300
            Left            =   7050
            TabIndex        =   32
            Top             =   1125
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Aliquota 
            Height          =   300
            Left            =   1860
            TabIndex        =   36
            Top             =   1725
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorICMS 
            Height          =   300
            Left            =   3675
            TabIndex        =   38
            Top             =   1725
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FreteValor 
            Height          =   300
            Left            =   1860
            TabIndex        =   26
            Top             =   1140
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox OutrosValores 
            Height          =   300
            Left            =   255
            TabIndex        =   34
            Top             =   1725
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox BaseCalculo 
            Height          =   300
            Left            =   5370
            TabIndex        =   40
            Top             =   1725
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Frete Peso"
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
            Left            =   510
            TabIndex        =   23
            Top             =   900
            Width           =   930
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sec/Cat"
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
            Left            =   4020
            TabIndex        =   27
            Top             =   900
            Width           =   720
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Despacho"
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
            Left            =   5640
            TabIndex        =   29
            Top             =   915
            Width           =   870
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Pedágio"
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
            Left            =   7440
            TabIndex        =   31
            Top             =   915
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota"
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
            Left            =   2235
            TabIndex        =   35
            Top             =   1515
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "ICMS"
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
            Left            =   4155
            TabIndex        =   37
            Top             =   1515
            Width           =   465
         End
         Begin VB.Label LabelTotais 
            AutoSize        =   -1  'True
            Caption         =   "Base Cálculo"
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
            Left            =   5520
            TabIndex        =   39
            Top             =   1515
            Width           =   1125
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Frete Valor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Left            =   2130
            TabIndex        =   25
            Top             =   900
            Width           =   945
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Outros Valores"
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
            Left            =   345
            TabIndex        =   33
            Top             =   1515
            Width           =   1260
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   7560
            TabIndex        =   41
            Top             =   1500
            Width           =   450
         End
         Begin VB.Label ValorTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   7080
            TabIndex        =   42
            Top             =   1725
            Width           =   1470
         End
      End
      Begin VB.Frame FrameOculto 
         Caption         =   "Frame3"
         Height          =   1065
         Left            =   3075
         TabIndex        =   186
         Top             =   3300
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox ValorSeguro 
            Height          =   315
            Left            =   600
            TabIndex        =   190
            Top             =   255
            Width           =   435
         End
         Begin VB.TextBox ValorDespesas 
            Height          =   315
            Left            =   600
            TabIndex        =   189
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox IPIValor1 
            Height          =   315
            Left            =   165
            TabIndex        =   188
            Top             =   570
            Width           =   435
         End
         Begin VB.TextBox ValorFrete 
            Height          =   315
            Left            =   165
            TabIndex        =   187
            Top             =   255
            Width           =   435
         End
         Begin MSMask.MaskEdBox ISSValor 
            Height          =   300
            Left            =   1050
            TabIndex        =   191
            Top             =   585
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label ICMSValor1 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1050
            TabIndex        =   192
            Top             =   255
            Width           =   405
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4815
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   9195
      Begin VB.Frame Frame6 
         Caption         =   "Dados Veículo"
         Height          =   780
         Index           =   1
         Left            =   105
         TabIndex        =   99
         Top             =   3600
         Width           =   8955
         Begin VB.ComboBox PlacaUF 
            Height          =   315
            Left            =   7755
            TabIndex        =   107
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox LocalVeiculo 
            Height          =   315
            Left            =   5545
            MaxLength       =   15
            TabIndex        =   105
            Top             =   300
            Width           =   1290
         End
         Begin VB.TextBox Placa 
            Height          =   315
            Left            =   3335
            MaxLength       =   10
            TabIndex        =   103
            Top             =   300
            Width           =   1290
         End
         Begin VB.TextBox MarcaVeiculo 
            Height          =   315
            Left            =   1125
            MaxLength       =   10
            TabIndex        =   101
            Top             =   300
            Width           =   1290
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "UF:"
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
            Left            =   7410
            TabIndex        =   106
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Local:"
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
            Left            =   4995
            TabIndex        =   104
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Placa:"
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
            Index           =   4
            Left            =   2745
            TabIndex        =   102
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Index           =   6
            Left            =   495
            TabIndex        =   100
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mercadoria(s) Transportada(s)"
         Height          =   3375
         Left            =   105
         TabIndex        =   73
         Top             =   120
         Width           =   8925
         Begin VB.ComboBox VolumeMarca 
            Height          =   315
            Left            =   5415
            TabIndex        =   164
            Top             =   345
            Width           =   1335
         End
         Begin VB.ComboBox VolumeEspecie 
            Height          =   315
            Left            =   3195
            TabIndex        =   163
            Top             =   345
            Width           =   1335
         End
         Begin VB.TextBox NaturezaCarga 
            Height          =   300
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   81
            Top             =   825
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Caption         =   "Locais"
            Height          =   900
            Left            =   480
            TabIndex        =   92
            Top             =   2250
            Width           =   8025
            Begin VB.TextBox CalculadoAte 
               Height          =   300
               Left            =   6420
               MaxLength       =   20
               TabIndex        =   98
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox Coleta 
               Height          =   300
               Left            =   930
               MaxLength       =   20
               TabIndex        =   94
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox Entrega 
               Height          =   300
               Left            =   3315
               MaxLength       =   20
               TabIndex        =   96
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Calculado Até :"
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
               Index           =   12
               Left            =   5055
               TabIndex        =   97
               Top             =   413
               Width           =   1320
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Coleta:"
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
               Left            =   195
               TabIndex        =   93
               Top             =   413
               Width           =   615
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Entrega:"
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
               Index           =   10
               Left            =   2505
               TabIndex        =   95
               Top             =   413
               Width           =   735
            End
         End
         Begin VB.TextBox VolumeNumero 
            Height          =   300
            Left            =   7335
            MaxLength       =   20
            TabIndex        =   79
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Observacao 
            Height          =   300
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   91
            Top             =   1815
            Width           =   6480
         End
         Begin VB.TextBox NotasFiscais 
            Height          =   300
            Left            =   1590
            MaxLength       =   20
            TabIndex        =   89
            Top             =   1350
            Width           =   4710
         End
         Begin MSMask.MaskEdBox VolumeQuant 
            Height          =   300
            Left            =   1590
            TabIndex        =   75
            Top             =   360
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PesoMercadoria 
            Height          =   300
            Left            =   3675
            TabIndex        =   83
            Top             =   810
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorMercadoria 
            Height          =   300
            Left            =   7350
            TabIndex        =   87
            Top             =   825
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Natureza Carga:"
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
            Index           =   13
            Left            =   150
            TabIndex        =   80
            Top             =   885
            Width           =   1395
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Valor Mercadoria:"
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
            Index           =   11
            Left            =   5820
            TabIndex        =   86
            Top             =   870
            Width           =   1515
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5355
            TabIndex        =   85
            Top             =   810
            Width           =   90
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kg ou m"
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
            Left            =   4665
            TabIndex        =   84
            Top             =   870
            Width           =   705
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   3165
            TabIndex        =   82
            Top             =   885
            Width           =   495
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
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
            Left            =   4770
            TabIndex        =   77
            Top             =   420
            Width           =   600
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
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
            Left            =   2385
            TabIndex        =   76
            Top             =   420
            Width           =   750
         End
         Begin VB.Label Label30 
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
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   74
            Top             =   420
            Width           =   1050
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Nº :"
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
            Left            =   6960
            TabIndex        =   78
            Top             =   420
            Width           =   345
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
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
            Left            =   465
            TabIndex        =   90
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Notas Fiscais:"
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
            Index           =   7
            Left            =   330
            TabIndex        =   88
            Top             =   1380
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   525
      Left            =   7395
      ScaleHeight     =   465
      ScaleWidth      =   1935
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   60
      Width           =   1995
      Begin VB.CommandButton BotaoExcluir 
         Height          =   330
         Left            =   555
         Picture         =   "ConhecimentoFreteOcx.ctx":1560
         Style           =   1  'Graphical
         TabIndex        =   157
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   330
         Left            =   1485
         Picture         =   "ConhecimentoFreteOcx.ctx":16EA
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   330
         Left            =   1020
         Picture         =   "ConhecimentoFreteOcx.ctx":1868
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   390
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   330
         Left            =   90
         Picture         =   "ConhecimentoFreteOcx.ctx":1D9A
         Style           =   1  'Graphical
         TabIndex        =   151
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   390
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5220
      Left            =   90
      TabIndex        =   0
      Top             =   405
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   9208
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Principais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Transporte"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complemento"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissões"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label ICMSSubstValor1 
      Height          =   270
      Left            =   390
      TabIndex        =   194
      Top             =   45
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "ConhecimentoFreteOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFretePesoAlterado As Integer
Dim iFreteValorAlterado As Integer
Dim iSECAlterado As Integer
Dim iDespachoAlterado As Integer
Dim iPedagioAlterado As Integer
Dim iOutrosAlterado As Integer
Dim iAliquotaAlterada As Integer
Dim iICMSAlterada As Integer
Dim iBaseCalculoAlterada As Integer
Dim iClienteAlterado As Integer
Dim iValorINSSAlterado As Integer

'Flag que indica se a tela está sendo preenchida.
Public gbCarregandoTela As Boolean

Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoNaturezaInterna As AdmEvento
Attribute objEventoNaturezaInterna.VB_VarHelpID = -1

'************** TRATAMENTO COMISSÕES ******************
'inicia objeto associado a GridComissoes
Public objTabComissoes As New ClassTabComissoes

Public objGridComissoes As AdmGrid
Public WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

'Declarações utilizadas para o cálculo das regras de comissões
Public objMnemonicoComissCalc As ClassMnemonicoComissCalc
Public objMnemonicoComissCalcAux As ClassMnemonicoComissCalcAux
Public objComissoesRegrasCalc As Object

'Indica que foi alterado pelo menos um campo na tela que seja utilizado para calcular as comissões
Public iComissoesAlterada As Integer

'******************************************************


'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_DADOSTRANSPORTE = 2
Private Const TAB_Complemento = 3
Private Const TAB_Comissoes = 4
Private Const TAB_Contabilizacao = 5

Private Const NATUREZAOP_PADRAO_CONHECIMENTOFRETE = "562"

'Mnemônicos da Contabilidade
Const FRETE_PESO As String = "Frete_Peso"
Const FRETE_VALOR As String = "Frete_Valor"
Const VALOR_TOTAL As String = "Valor_Total"
Const SECCAT As String = "Sec_Cat"
Const DESPACHO1 As String = "Despacho"
Const PEDAGIO1 As String = "Pedagio"
Const OUTROS_VALORES As String = "Outros_Valores"
Const ICMS As String = "ICMS"
Const ICMS_INCLUSO As String = "ICMS_Incluso"
Const BASE_CALCULO As String = "Base_Calculo"
Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Const CLIENTE_RAZAOSOCIAL  As String = "Cliente_Razao_Social"
Const CTA_VENDAS As String = "CtaVendas"

Public iAlterado As Integer
Dim iFrameAtual  As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
       
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    NatOpInterna.Text = NaturezaOp_Conv34(NATUREZAOP_PADRAO_CONHECIMENTOFRETE, gdtDataAtual)
    
    'Inicializa as Variáveis de browse
    Set objEventoSerie = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    
'************ TRATAMENTO COMISSOES *****************

        Set objGridComissoes = New AdmGrid
        Set objEventoVendedor = New AdmEvento
        
        Set objTabComissoes.objTela = Me
        
        '*** TRATAMENTO PARA CÁLCULO DAS REGRAS DE COMISSÕES ***
        Set objMnemonicoComissCalc = New ClassMnemonicoComissCalc
        Set objMnemonicoComissCalcAux = New ClassMnemonicoComissCalcAux
        Set objComissoesRegrasCalc = CreateObject("RotinasContab.ClassComissoesRegrasCalc")
        Set objMnemonicoComissCalc.objTela = Me
        Set objMnemonicoComissCalcAux.objTela = Me
        '********************************************************
        
        'Inicializa o Grid de Comissões
        lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
        If lErro <> SUCESSO Then gError 126985
        
        'o grid está sobreposto por uma frame
        objGridComissoes.iLinhasExistentes = objGridComissoes.iLinhasExistentes - 2
        Call Grid_Inicializa(objGridComissoes)
        
'***************************************************
    
    'Carrega as Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 62832
    
    'Carrega os Estados
    lErro = Carrega_PlacaUF()
    If lErro <> SUCESSO Then gError 62833
       
    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeEspecie
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)
    If lErro <> SUCESSO Then gError 62833

    'Incluído por Luiz Nogueira em 21/08/03
    'Carrega a combo VolumeMarca
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    If lErro <> SUCESSO Then gError 62833
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
    If lErro <> SUCESSO Then gError 62834
        
    iAlterado = 0
    iComissoesAlterada = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 62832, 62833, 62834, 126985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154790)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0
    
    Exit Sub

End Sub

Private Sub Aliquota_Change()
    iAlterado = REGISTRO_ALTERADO
    iAliquotaAlterada = REGISTRO_ALTERADO
End Sub

Private Sub Aliquota_GotFocus()
    iAliquotaAlterada = 0
End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Aliquota_Validate

    If iAliquotaAlterada = 0 Then Exit Sub
    
    If Len(Aliquota.Text) > 0 Then
        'Testa o valor
        lErro = Porcentagem_Critica2(Aliquota.Text)
        If lErro <> SUCESSO Then gError 62862
    Else
        ValorICMS.Text = ""
    End If

    Call ValorTotal_Calcula

    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True

    Select Case gErr

        Case 62862

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154791)

    End Select

    Exit Sub

End Sub


Private Sub BaseCalculo_Change()

    iAlterado = REGISTRO_ALTERADO
    iBaseCalculoAlterada = REGISTRO_ALTERADO
    
End Sub

Private Sub BaseCalculo_GotFocus()
    iBaseCalculoAlterada = 0
End Sub

Private Sub BaseCalculo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_BaseCalculo_Validate

    If iBaseCalculoAlterada = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(BaseCalculo.ClipText)) > 0 Then
        
        'Critica se é valor não negativo
        lErro = Valor_Positivo_Critica(BaseCalculo.Text)
        If lErro <> SUCESSO Then gError 62863
   
        'Põe o valor formatado na tela
        BaseCalculo.Text = Format(BaseCalculo.Text, "Fixed")
    
        lErro = BaseCalculo_Calula_ValorTotal
        If lErro <> SUCESSO Then gError 62864
    
    End If
        
    Exit Sub

Erro_BaseCalculo_Validate:

    Cancel = True

    Select Case gErr

        Case 62863, 62864

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154792)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objcliente As New ClassCliente

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Serie.Text)) = 0 Then gError 89137
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 89138

    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.sSerie = Serie.Text
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)

    lErro = TP_Cliente_Le(Cliente, objcliente, giFilialEmpresa, 0)
    If lErro <> SUCESSO Then gError 109000
    
    objNFiscal.lCliente = objcliente.lCodigo

    'Verifica se a existe nota fiscal está cadastrada
    lErro = CF("NFiscal_Le_1", objNFiscal)
    If lErro <> SUCESSO And lErro <> 83971 Then gError 89139
    
    'se a nota não está cadastrada ==> erro
    If lErro = 83971 Then gError 89140

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 81871

        'Faz o cancelamento de uma nota fiscal de Saida
        lErro = CF("NotaFiscalSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 89141
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscal1()
        If lErro <> SUCESSO Then gError 89142
    
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 109000

        Case 89137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 89138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 81871, 89139, 89141, 89142
       
        Case 89140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154793)
        
    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoLimpaDestinatario_Click()

    Call Limpa_Destinatario

End Sub

Private Sub BotaoLimpaRemetente_Click()

    Call Limpa_Remetente

End Sub

Private Sub CalculadoAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Coleta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub Despacho_Change()
    iAlterado = REGISTRO_ALTERADO
    iDespachoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Despacho_GotFocus()
    iDespachoAlterado = 0
End Sub

Private Sub Despacho_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Despacho_Validate

    If iDespachoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(Despacho.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(Despacho.Text)
        If lErro <> SUCESSO Then gError 62865
    
        'Põe o valor formatado na tela
        Despacho.Text = Format(Despacho.Text, "Fixed")
        
    End If
        
    Call ValorTotal_Calcula
    
    Exit Sub

Erro_Despacho_Validate:

    Cancel = True
    
    Select Case gErr

        Case 62865

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154794)

    End Select

    Exit Sub

End Sub

Private Sub Entrega_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FretePeso_Change()
    iAlterado = REGISTRO_ALTERADO
    iFretePesoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FretePeso_GotFocus()
    iFretePesoAlterado = 0
End Sub

Private Sub FretePeso_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FretePeso_Validate

    If iFretePesoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(FretePeso.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(FretePeso.Text)
        If lErro <> SUCESSO Then gError 62866
    
        'Põe o valor formatado na tela
        FretePeso.Text = Format(FretePeso.Text, "Fixed")
    
    End If
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_FretePeso_Validate:

    Cancel = True

    Select Case gErr

        Case 62866

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154795)

    End Select

    Exit Sub

End Sub

Private Sub FreteValor_Change()
    iAlterado = REGISTRO_ALTERADO
    iFreteValorAlterado = REGISTRO_ALTERADO
End Sub

Private Sub FreteValor_GotFocus()
    iFreteValorAlterado = 0
End Sub

Private Sub FreteValor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_FreteValor_Validate

    If iFreteValorAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(FreteValor.ClipText)) > 0 Then

        'Critica se é valor Positivo
        lErro = Valor_Positivo_Critica_Double(FreteValor.Text)
        If lErro <> SUCESSO Then gError 62867
    
        'Põe o valor formatado na tela
        FreteValor.Text = Format(FreteValor.Text, "Fixed")
    End If

    Call ValorTotal_Calcula
    
    Exit Sub

Erro_FreteValor_Validate:

    Cancel = True

    Select Case gErr

        Case 62867

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154796)

    End Select

    Exit Sub

End Sub

Private Sub ICMSIncluso_Click()

    iAlterado = REGISTRO_ALTERADO
    Call ValorTotal_Calcula

End Sub

Private Sub LblNatOpInterna_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    objNaturezaOp.sCodigo = NatOpInterna.Text

    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        dtDataRef = MaskedParaDate(DataEmissao)
    Else
        dtDataRef = DATA_NULA
    End If
        
    colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
    colSelecao.Add NATUREZA_SAIDA_COD_FINAL
            
    'Chama a Tela de browse de NaturezaOp p/naturezas de entrada
    Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, objEventoNatureza, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))

End Sub

Private Sub LocalVeiculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub MarcaVeiculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatOpInterna_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NatOpInterna_Validate(Cancel As Boolean)

Dim lErro As Long, objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NatOpInterna_Validate

    If Len(Trim(NatOpInterna.Text)) > 0 Then
    
        objNaturezaOp.sCodigo = NatOpInterna
        
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then gError 62868
        
        'se nao encontrou a natureza
        If lErro <> SUCESSO Then gError 62869
        
        If objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then gError 62885

        If Natop_ErroTamanho(MaskedParaDate(DataEmissao), objNaturezaOp.sCodigo) Then gError 108930
                       
    End If
        
    Exit Sub

Erro_NatOpInterna_Validate:

    Cancel = True
    
    Select Case gErr

        Case 108930
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", gErr)
        
        Case 62868
        
        Case 62869
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, NatOpInterna.Text)
        
        Case 62870
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", gErr, NatOpInterna.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154797)

    End Select
    
    Exit Sub

End Sub

Private Sub NaturezaCarga_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NotasFiscais_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim objNatOp As New ClassNaturezaOp

    Set objNatOp = obj1
    
    NatOpInterna.Text = objNatOp.sCodigo
    
    Me.Show
    
    Exit Sub

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub OutrosValores_Change()
    iAlterado = REGISTRO_ALTERADO
    iOutrosAlterado = REGISTRO_ALTERADO
End Sub

Private Sub OutrosValores_GotFocus()
    iOutrosAlterado = 0
End Sub

Private Sub OutrosValores_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OutrosValores_Validate

    If iOutrosAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(OutrosValores.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(OutrosValores.Text)
        If lErro <> SUCESSO Then gError 62871
        
        'Põe o valor formatado na tela
        OutrosValores.Text = Format(OutrosValores.Text, "Fixed")

    End If
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_OutrosValores_Validate:

    Cancel = True

    Select Case gErr

        Case 62871

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154798)

    End Select

    Exit Sub

End Sub

Private Sub Pedagio_Change()
    iAlterado = REGISTRO_ALTERADO
    iPedagioAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Pedagio_GotFocus()
    iPedagioAlterado = 0
End Sub

Private Sub Pedagio_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Pedagio_Validate

    If iPedagioAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(Pedagio.ClipText)) > 0 Then
        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(Pedagio.Text)
        If lErro <> SUCESSO Then gError 62872
    
        'Põe o valor formatado na tela
        Pedagio.Text = Format(Pedagio.Text, "Fixed")
        
    End If
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_Pedagio_Validate:

    Cancel = True


    Select Case gErr

        Case 62872

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154799)

    End Select

    Exit Sub

End Sub

Private Sub PedagioIncluso_Click()
    iAlterado = REGISTRO_ALTERADO
    Call ValorTotal_Calcula
End Sub

Private Sub PesoMercadoria_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub PesoMercadoria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PesoMercadoria_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(PesoMercadoria.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(PesoMercadoria.Text)
    If lErro <> SUCESSO Then gError 62873

    'Põe o valor formatado na tela
    PesoMercadoria.Text = Format(PesoMercadoria.Text, "Fixed")

    Exit Sub

Erro_PesoMercadoria_Validate:

    Cancel = True

    Select Case gErr

        Case 62873

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154800)

    End Select

    Exit Sub

End Sub

Private Sub SEC_Change()
    iAlterado = REGISTRO_ALTERADO
    iSECAlterado = REGISTRO_ALTERADO
End Sub

Private Sub SEC_GotFocus()
    iSECAlterado = 0
End Sub

Private Sub SEC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SEC_Validate

    If iSECAlterado = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(SEC.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(SEC.Text)
        If lErro <> SUCESSO Then gError 62874
    
        'Põe o valor formatado na tela
        SEC.Text = Format(SEC.Text, "Fixed")
    
    End If
    
    Call ValorTotal_Calcula

    Exit Sub

Erro_SEC_Validate:

    Cancel = True

    Select Case gErr

        Case 62874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154801)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Serie_Validate

    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Tenta selecionar a serie
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62875
    
    'Se não está na combo
    If lErro <> SUCESSO Then
    
        objSerie.sSerie = Serie.Text
        'Busca a série no BD
        lErro = CF("Serie_Le", objSerie)
        If lErro <> SUCESSO And lErro <> 22202 Then gError 62876
        If lErro <> SUCESSO Then gError 35750 'Se não encontrou
    
    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case gErr

        Case 35750
            'Pergunta se deseja criar nova série
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_SERIE", Serie.Text)
            'Se a resposta for afirmativa
            If vbMsgRes = vbYes Then
                'Chama a tela de cadastro de séries
                Call Chama_Tela("SerieNFiscal", objSerie)
            End If
            'segura o foco na série

        Case 62875, 62876

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154802)

    End Select
    
    Exit Sub

End Sub

Public Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série na Tela
    Serie.Text = objSerie.sSerie

    Me.Show

    Exit Sub

End Sub
Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long
Dim colSelecao As New Collection

    'Recolhe os dados da Nota Fiscal
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then Exit Sub
    
    colSelecao.Add TIPODOCINFO_CONHECIMENTOFRETE

    'Chama a Tela NFConhecimentoFreteLista
    Call Chama_Tela("NFConhecimentoFreteLista", colSelecao, objNFiscal, objEventoNFiscal)

    Exit Sub

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_Conhecimento_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 62877

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 62877

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154803)

    End Select

    Exit Sub

End Sub

Public Sub ClienteLabel_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Cliente.Text)) > 0 Then objcliente.sNomeReduzido = Cliente.Text

    'Chama Tela ClienteLista
    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO
    
    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o cliente foi alterado
    If iClienteAlterado = 0 Then Exit Sub
    'Se op cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 62878

        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 62879

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If Not gbCarregandoTela Then
            'Seleciona filial na Combo Filial
            If iCodFilial = FILIAL_MATRIZ Then
                Filial.ListIndex = 0
            Else
                Call CF("Filial_Seleciona", Filial, iCodFilial)
                
            End If
                            
        End If
        
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        Filial.Clear
    End If

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 62878, 62879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154804)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 62880

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case gErr

        Case 62880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154805)

    End Select

    Exit Sub

End Sub


Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 62881

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 62881

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154806)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 62882

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 62882

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154807)

    End Select

    Exit Sub

End Sub
Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    iComissoesAlterada = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se algo foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Faz o tratamento da Filial selecionada
    lErro = Trata_FilialCliente()
    If lErro <> SUCESSO Then gError 62883

    Exit Sub

Erro_Filial_Click:

    Select Case gErr

        Case 62883

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154808)

    End Select

    Exit Sub

End Sub


Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 62884

    'Se nao encontra o item com o código informado
    If lErro = 6730 Then

        'Verifica de o Cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 62885

        sCliente = Cliente.Text

        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 62886

        If lErro = 17660 Then gError 62887

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        lErro = Trata_FilialCliente()
        If lErro <> SUCESSO Then gError 62888

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 62889

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 62884, 62886

        Case 62887
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 62885
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 62888, 62889
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154809)

    End Select

    Exit Sub

End Sub

Private Function Carrega_PlacaUF() As Long
'Lê as Siglas dos Estados e alimenta a list da Combobox PlacaUF

Dim lErro As Long
Dim colSiglasUF As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_PlacaUF

    Set colSiglasUF = gcolUFs
    
    'Adiciona na Combo PlacaUF
    For iIndice = 1 To colSiglasUF.Count
        PlacaUF.AddItem colSiglasUF.Item(iIndice)
        UFRemetente.AddItem colSiglasUF.Item(iIndice)
        UFDestinatario.AddItem colSiglasUF.Item(iIndice)
    Next

    Carrega_PlacaUF = SUCESSO

    Exit Function

Erro_Carrega_PlacaUF:

    Carrega_PlacaUF = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154810)

    End Select

End Function

Public Sub PlacaUF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PlacaUF_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub UFRemetente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFRemetente_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(UFRemetente.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(UFRemetente)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62890

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62891

    Exit Sub

Erro_UFRemetente_Validate:

    Cancel = True


    Select Case gErr

        Case 62890

        Case 62891
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFRemetente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154811)

    End Select

    Exit Sub

End Sub
Public Sub UFRemetente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFRemetente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFDestinatario_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_UFDestinatario_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(UFDestinatario.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(UFDestinatario)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62892

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62893

    Exit Sub

Erro_UFDestinatario_Validate:

    Cancel = True


    Select Case gErr

        Case 62892

        Case 62893
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, UFDestinatario.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154812)

    End Select

    Exit Sub

End Sub
Public Sub UFDestinatario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UFDestinatario_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PlacaUF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PlacaUF_Validate

    'verifica se tem alguma Coisa preenchida
    If Len(Trim(PlacaUF.Text)) = 0 Then Exit Sub

    'Verifica se existe o item na combo
    lErro = Combo_Item_Igual(PlacaUF)
    If lErro <> SUCESSO And lErro <> 12253 Then gError 62894

    'Se não encontrar --> Erro
    If lErro = 12253 Then gError 62895

    Exit Sub

Erro_PlacaUF_Validate:

    Cancel = True

    Select Case gErr

        Case 62894

        Case 62895
            Call Rotina_Erro(vbOKOnly, "ERRO_UF_NAO_CADASTRADA", gErr, PlacaUF.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154813)

    End Select

    Exit Sub

End Sub


Public Sub CGCRemetente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGCRemetente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGCRemetente_Validate
    
    'Se CGCRemetente/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGCRemetente.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGCRemetente.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGCRemetente.Text)
            If lErro <> SUCESSO Then gError 62896
            
            'Formata e coloca na Tela
            CGCRemetente.Format = "000\.000\.000-00; ; ; "
            CGCRemetente.Text = CGCRemetente.Text

        Case STRING_CGC 'CGCRemetente
            
            'Critica CGCRemetente
            lErro = Cgc_Critica(CGCRemetente.Text)
            If lErro <> SUCESSO Then gError 62897
            
            'Formata e Coloca na Tela
            CGCRemetente.Format = "00\.000\.000\/0000-00; ; ; "
            CGCRemetente.Text = CGCRemetente.Text

        Case Else
                
            gError 62898

    End Select

    Exit Sub

Erro_CGCRemetente_Validate:

    Cancel = True

    Select Case gErr

        Case 62896, 62897

        Case 62898
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154814)

    End Select


    Exit Sub

End Sub
Public Sub CGCRemetente_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGCRemetente, iAlterado)

End Sub

Public Sub CGCDestinatario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CGCDestinatario_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGCDestinatario_Validate
    
    'Se CGCDestinatario/CPF não foi preenchido -- Exit Sub
    If Len(Trim(CGCDestinatario.Text)) = 0 Then Exit Sub
    
    Select Case Len(Trim(CGCDestinatario.Text))

        Case STRING_CPF 'CPF
            
            'Critica Cpf
            lErro = Cpf_Critica(CGCDestinatario.Text)
            If lErro <> SUCESSO Then gError 62899
            
            'Formata e coloca na Tela
            CGCDestinatario.Format = "000\.000\.000-00; ; ; "
            CGCDestinatario.Text = CGCDestinatario.Text

        Case STRING_CGC 'CGC
            
            'Critica CGCDestinatario
            lErro = Cgc_Critica(CGCDestinatario.Text)
            If lErro <> SUCESSO Then gError 62950
            
            'Formata e Coloca na Tela
            CGCDestinatario.Format = "00\.000\.000\/0000-00; ; ; "
            CGCDestinatario.Text = CGCDestinatario.Text

        Case Else
                
            gError 62951

    End Select

    Exit Sub

Erro_CGCDestinatario_Validate:

    Cancel = True


    Select Case gErr

        Case 62899, 62950

        Case 62951
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154815)

    End Select


    Exit Sub

End Sub
Public Sub CGCDestinatario_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGCDestinatario, iAlterado)

End Sub
Private Function Carrega_Serie() As Long
'Carrega as combos de Série e serie de NF original com as séries lidas do BD

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 62952

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next
            
    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 62952

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154816)

    End Select

    Exit Function

End Function

Public Sub BotaoLimparNF_Click()

    NFiscal.Caption = ""
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    'Libera as variáveis globais da tela
    Set objEventoSerie = Nothing
    Set objEventoCliente = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoNatureza = Nothing
       
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    '************ TRATAMENTO COMISSOES *************
    Set objGridComissoes = Nothing
    Set objTabComissoes = Nothing
        
    '*** TRATAMENTO PARA CÁLCULO DAS REGRAS DE COMISSÕES ***
    Set objMnemonicoComissCalc = Nothing
    Set objMnemonicoComissCalcAux = Nothing
    Set objComissoesRegrasCalc = Nothing
    '********************************************************



    'Fecha o Comando de Setas
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalConhecimentoFrete"

    'Lê os dados da Tela NFiscal
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 62953

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, STRING_BUFFER_MAX_TEXTO, "VolumeEspecie"
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, STRING_BUFFER_MAX_TEXTO, "VolumeMarca"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "Observacao", objNFiscal.sObservacao, STRING_NFISCAL_OBSERVACAO, "Observacao"
    colCampoValor.Add "INSSRetido", objNFiscal.objConhecimentoFrete.iINSSRetido, 0, "INSSRetido" 'William
    colCampoValor.Add "ValorINSS", objNFiscal.objConhecimentoFrete.dValorINSS, 0, "ValorINSS" 'William
    colCampoValor.Add "IncluiPedagio", objNFiscal.objConhecimentoFrete.iIncluiPedagio, 0, "IncluiPedagio" 'William
        
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "TipoNFiscal", OP_IGUAL, TIPODOCINFO_CONHECIMENTOFRETE
    
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 62953

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154817)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche
    
    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete

    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objNFiscal.lNumIntDoc <> 0 Then

        'Carrega objNFiscal com os dados passados em colCampoValor
        objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor
        objNFiscal.iTipoNFiscal = colCampoValor.Item("TipoNFiscal").vValor
        objNFiscal.sNaturezaOp = colCampoValor.Item("NaturezaOP").vValor
        objNFiscal.lCliente = colCampoValor.Item("Cliente").vValor
        objNFiscal.iFilialCli = colCampoValor.Item("FilialCli").vValor
        objNFiscal.sSerie = colCampoValor.Item("Serie").vValor
        objNFiscal.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
        objNFiscal.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objNFiscal.sPlaca = colCampoValor.Item("Placa").vValor
        objNFiscal.sPlacaUF = colCampoValor.Item("PlacaUF").vValor
        objNFiscal.lVolumeQuant = colCampoValor.Item("VolumeQuant").vValor
        objNFiscal.lVolumeEspecie = colCampoValor.Item("VolumeEspecie").vValor
        objNFiscal.lVolumeMarca = colCampoValor.Item("VolumeMarca").vValor
        objNFiscal.sVolumeNumero = colCampoValor.Item("VolumeNumero").vValor
        objNFiscal.iStatus = colCampoValor.Item("Status").vValor
        objNFiscal.sObservacao = colCampoValor.Item("Observacao").vValor
        objNFiscal.objConhecimentoFrete.iINSSRetido = colCampoValor.Item("INSSRetido").vValor 'William
        objNFiscal.objConhecimentoFrete.dValorINSS = colCampoValor.Item("ValorINSS").vValor 'William
        objNFiscal.objConhecimentoFrete.iIncluiPedagio = colCampoValor.Item("IncluiPedagio").vValor

        'Coloca os dados da NFiscal na tela
        lErro = Traz_Conhecimento_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 62954

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 62954

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154818)

    End Select

    Exit Sub

End Sub


Public Function Move_Conhecimento_Memoria(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente

On Error GoTo Erro_Move_Conhecimento_Memoria

    
'================== DADOS IDENTIFICACAO =====================
    If Len(Trim(Cliente.ClipText)) > 0 Then
        objcliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 62955
        
        'Não encontrou p Cliente --> erro
        If lErro = 12348 Then gError 62956

        objNFiscal.lCliente = objcliente.lCodigo
        
    End If
    
    'Verifica se Está Preenchido o Numero da Nota no Label
    objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Caption)
    objNFiscal.iFilialCli = Codigo_Extrai(Filial.Text)
    objNFiscal.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objNFiscal.sSerie = Serie.Text
    objNFiscal.sNaturezaOp = NatOpInterna.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.iTipoNFiscal = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.iTipoDocInfo = TIPODOCINFO_CONHECIMENTOFRETE
    objNFiscal.dtDataRegistro = gdtDataHoje

'============= DADOS COMPOSICAO FRETE ===================
    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete


    With objNFiscal.objConhecimentoFrete
    
        .iICMSIncluso = ICMSIncluso.Value
        .dFretePeso = StrParaDbl(FretePeso.Text)
        .dFreteValor = StrParaDbl(FreteValor.Text)
        .dSEC = StrParaDbl(SEC.Text)
        .dDespacho = StrParaDbl(Despacho.Text)
        .dPedagio = StrParaDbl(Pedagio.Text)
        .dOutrosValores = StrParaDbl(OutrosValores.Text)
        .dAliquotas = PercentParaDbl(Aliquota.FormattedText)
        .dValorICMS = StrParaDbl(ValorICMS.Text)
        .dBaseCalculo = StrParaDbl(BaseCalculo.Text)
        objNFiscal.dValorTotal = StrParaDbl(ValorTotal.Caption)
        .dValorINSS = StrParaDbl(ValorINSS.Text)
        .iINSSRetido = INSSRetido.Value
        .iIncluiPedagio = PedagioIncluso.Value
    End With

'===================== DADOS TRANSPORTE =======================
    With objNFiscal.objConhecimentoFrete
    
        .sRemetente = Remetente.Text
        .sEnderecoRemetente = EnderecoRemetente.Text
        .sMunicipioRemetente = CidadeRemetente.Text
        .sUFRemetente = UFRemetente.Text
        .sCepRemetente = CEPRemetente.ClipText
        .sCGCRemetente = CGCRemetente.ClipText
        .sInscEstadualRemetente = InscEstRemetente.Text

        .sDestinatario = Destinatario.Text
        .sEnderecoDestinatario = EnderecoDestinatario.Text
        .sMunicipioDestinatario = CidadeDestinatario.Text
        .sUFDestinatario = UFDestinatario.Text
        .sCepDestinatario = CEPDestinatario.ClipText
        .sCGCDestinatario = CGCDestinatario.ClipText
        .sInscEstadualDestinatario = InscEstDestinatario.Text
    
    End With
    
'================== DADOS COMPLEMENTARES ====================
    
    objNFiscal.lVolumeQuant = StrParaLong(VolumeQuant.Text)
    If Len(Trim(VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(VolumeEspecie.Text)
    If Len(Trim(VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(VolumeMarca.Text)
    objNFiscal.sVolumeNumero = VolumeNumero.Text
    
    With objNFiscal.objConhecimentoFrete
    
        .sNaturezaCarga = NaturezaCarga.Text
        .dPesoMercadoria = StrParaDbl(PesoMercadoria.Text)
        .dValorMercadoria = StrParaDbl(ValorMercadoria.Text)
        .sNotasFiscais = NotasFiscais.Text
        objNFiscal.sObservacao = Observacao.Text
        .sColeta = Coleta.Text
        .sEntrega = Entrega.Text
        .sCalculadoAte = CalculadoAte.Text
        .sMarcaVeiculo = MarcaVeiculo.Text
        .sLocalVeiculo = LocalVeiculo.Text
        objNFiscal.sPlaca = Placa.Text
        objNFiscal.sPlacaUF = PlacaUF.Text
    
    End With
    
'=========== INICIALIZANDO DADOS NAO PRESENTES NA TELA =========

    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataReferencia = DATA_NULA
    objNFiscal.dtDataSaida = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.lNumIntDoc = 0
    
'********************* TRATAMENTO COMISSOES *************************
    'Chama Move_GridComissoes_Memoria
    lErro = objTabComissoes.Move_TabComissoes_Memoria(objNFiscal, NOTA_FISCAL)
    If lErro <> SUCESSO Then gError 126986
'********************************************************************
    
    Move_Conhecimento_Memoria = SUCESSO

    Exit Function

Erro_Move_Conhecimento_Memoria:

    Move_Conhecimento_Memoria = gErr

    Select Case gErr

        Case 62955, 126986

        Case 62956
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154819)

    End Select

    Exit Function

End Function

Public Sub Placa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click
    
    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
         Frame1(TabStrip1.SelectedItem.Index).Visible = True
         'Torna Frame atual visivel
         Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao


        Select Case iFrameAtual
        
            Case TAB_DadosPrincipais
''''                Parent.HelpContextID = IDH_NF_SAIDA_DADOS_PRINCIPAIS
            
            Case TAB_Complemento
''''                Parent.HelpContextID = IDH_NF_SAIDA_COMPLEMENTO
                        
            Case TAB_Contabilizacao
''''                Parent.HelpContextID = IDH_NF_SAIDA_CONTABILIZACAO
            
        End Select
    
    End If

    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154820)

    End Select

    Exit Sub

End Sub
Private Sub ValorICMS_Change()
    
    iAlterado = REGISTRO_ALTERADO
    iICMSAlterada = REGISTRO_ALTERADO

End Sub

Private Sub ValorICMS_GotFocus()
    iICMSAlterada = 0
End Sub

Private Sub ValorICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Pedagio_Validate

    If iICMSAlterada = 0 Then Exit Sub

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMS.ClipText)) > 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorICMS.Text)
        If lErro <> SUCESSO Then gError 62957
    
        'Põe o valor formatado na tela
        ValorICMS.Text = Format(ValorICMS.Text, "Fixed")
    End If
    
    
    lErro = ValorTotal_Calcula(True)

    Exit Sub

Erro_Pedagio_Validate:

    Cancel = True

    Select Case gErr

        Case 62957

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154821)

    End Select

    Exit Sub

End Sub

Private Sub ValorMercadoria_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorMercadoria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorMercadoria_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorMercadoria.Text)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorMercadoria.Text)
    If lErro <> SUCESSO Then gError 62958

    'Põe o valor formatado na tela
    ValorMercadoria.Text = Format(ValorMercadoria.Text, "Fixed")

    Exit Sub

Erro_ValorMercadoria_Validate:

    Cancel = True

    Select Case gErr

        Case 62958

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154822)

    End Select

    Exit Sub

End Sub

Public Sub VolumeEspecie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeMarca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeNumero_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VolumeQuant_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Function Traz_Conhecimento_Tela(objNFiscal As ClassNFiscal) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim objNFiscalOriginal As New ClassNFiscal
Dim bCancel As Boolean
Dim bAlterouCT As Boolean

On Error GoTo Erro_Traz_Conhecimento_Tela
    
    'Limpa a tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 62860
    
    Set objNFiscal.objConhecimentoFrete = New ClassConhecimentoFrete
    objNFiscal.objConhecimentoFrete.lNumIntNFiscal = objNFiscal.lNumIntDoc
        
    lErro = ConhecimentoFrete_Le(objNFiscal.objConhecimentoFrete)
    If lErro <> SUCESSO And lErro <> 62857 Then gError 62959
    If lErro <> SUCESSO Then gError 62860
        
'********** TRATAMENTO COMISSOES ***********
    'Lê as Comissões da Nota Fiscal
    lErro = CF("ComissoesNF_Le", objNFiscal)
    If lErro <> SUCESSO Then gError 126987
'*******************************************
        
'============ DADOS IDENTIFICACAO =====================
    
    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        Status.Caption = STRING_STATUS_CANCELADO
    End If
    
    NatOpInterna.Text = objNFiscal.sNaturezaOp
    NFiscal.Caption = objNFiscal.lNumNotaFiscal
    Serie.Text = objNFiscal.sSerie
    
    'Preenche o Cliente
    Cliente.Text = objNFiscal.lCliente
    Call Cliente_Validate(bCancel)
    
    'Preenche a Filial do Cliente
    Call Filial_Formata(Filial, objNFiscal.iFilialCli)

    Call DateParaMasked(DataEmissao, objNFiscal.dtDataEmissao)
    
'========= DADOS COMPOSICAO FRETE =============
    With objNFiscal.objConhecimentoFrete
        
        ICMSIncluso.Value = .iICMSIncluso
        PedagioIncluso.Value = .iIncluiPedagio
        If .dFretePeso > 0 Then FretePeso.Text = Format(.dFretePeso, "Standard")
        FreteValor.Text = Format(.dFreteValor, "Standard")
        If .dSEC > 0 Then SEC.Text = Format(.dSEC, "Standard")
        If .dDespacho > 0 Then Despacho.Text = Format(.dDespacho, "Standard")
        If .dPedagio > 0 Then Pedagio.Text = Format(.dPedagio, "Standard")
        If .dOutrosValores > 0 Then OutrosValores.Text = Format(.dOutrosValores, "Standard")
        If .dAliquotas > 0 Then Aliquota.Text = Format(.dAliquotas * 100, "Fixed")
        If .dValorICMS > 0 Then ValorICMS.Text = Format(.dValorICMS, "Standard")
        BaseCalculo.Text = Format(.dBaseCalculo, "Standard")
        
        Call ValorTotal_Calcula(True)
    
'============= DADOS TRANSPORTE =========================
    
        Remetente.Text = .sRemetente
        EnderecoRemetente.Text = .sEnderecoRemetente
        CidadeRemetente.Text = .sMunicipioRemetente
        UFRemetente.Text = .sUFRemetente
        CEPRemetente.PromptInclude = False
        CEPRemetente.Text = .sCepRemetente
        CEPRemetente.PromptInclude = True
        
        CGCRemetente.Text = .sCGCRemetente
        Call CGCRemetente_Validate(False)
        InscEstRemetente.Text = .sInscEstadualRemetente
        
        Destinatario.Text = .sDestinatario
        EnderecoDestinatario.Text = .sEnderecoDestinatario
        CidadeDestinatario.Text = .sMunicipioDestinatario
        UFDestinatario.Text = .sUFDestinatario
        CEPDestinatario.PromptInclude = False
        CEPDestinatario.Text = .sCepDestinatario
        CEPDestinatario.PromptInclude = True
        
        CGCDestinatario.Text = .sCGCDestinatario
        Call CGCDestinatario_Validate(False)
        InscEstDestinatario.Text = .sInscEstadualDestinatario
    End With
'============= DADOS COMPLEMENTARES =======================

    If objNFiscal.lVolumeQuant > 0 Then VolumeQuant.Text = objNFiscal.lVolumeQuant
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call VolumeEspecie_Validate(bSGECancelDummy)
    Else
        VolumeEspecie.Text = ""
    End If
    
    'Alterado por Luiz Nogueira em 21/08/03
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call VolumeMarca_Validate(bSGECancelDummy)
    Else
        VolumeMarca.Text = ""
    End If
    
    VolumeNumero = objNFiscal.sVolumeNumero

    With objNFiscal.objConhecimentoFrete
    
        NaturezaCarga.Text = .sNaturezaCarga
        If .dPesoMercadoria > 0 Then PesoMercadoria.Text = Format(.dPesoMercadoria, "Fixed")
        If .dValorMercadoria > 0 Then ValorMercadoria.Text = Format(.dValorMercadoria, "Fixed")
        NotasFiscais.Text = .sNotasFiscais
        Observacao.Text = objNFiscal.sObservacao
        Coleta.Text = .sColeta
        Entrega.Text = .sEntrega
        CalculadoAte.Text = .sCalculadoAte
        MarcaVeiculo.Text = .sMarcaVeiculo
        Placa.Text = objNFiscal.sPlaca
        PlacaUF.Text = objNFiscal.sPlacaUF
        LocalVeiculo.Text = .sLocalVeiculo
        ValorINSS.Text = Format(.dValorINSS, "Standard")
        INSSRetido.Value = IIf(.iINSSRetido <> 0, vbChecked, vbUnchecked)
        
    End With
    
    Call ValorTotal_Calcula(True)
        
'***************** TRATAMENTO COMISSOES ************************
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objNFiscal)
    If lErro <> SUCESSO Then gError 126988
'****************************************************************
        
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 62861

    iAlterado = 0
     
    gbCarregandoTela = False
     
    Traz_Conhecimento_Tela = SUCESSO

    Exit Function

Erro_Traz_Conhecimento_Tela:

    gbCarregandoTela = False

    Traz_Conhecimento_Tela = gErr

    Select Case gErr

        Case 62959, 62861, 126987, 126988

        Case 62860
            Call Rotina_Erro(vbOKOnly, "ERRO_CONHECIMENTOFRETE_NAO_CADASTRADO", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154823)

    End Select

    Exit Function

End Function

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 62962
        If lErro <> SUCESSO Then gError 62963
        
        'Traz a nota para a tela
        lErro = Traz_Conhecimento_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 62964

    End If

    iAlterado = 0
    iComissoesAlterada = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 62962, 62964, 64003

        Case 62963
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154824)

    End Select
    
    iAlterado = 0
    iComissoesAlterada = 0

    Exit Function

End Function

Private Function Trata_FilialCliente() As Long

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente, objcliente As New ClassCliente
Dim objVendedor As New ClassVendedor
Dim objComissaoNF As New ClassComissaoNF
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Trata_FilialCliente

    objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
    'Lê a FilialCliente
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilialCliente)
    If lErro <> SUCESSO Then gError 62965

    'Preenche o endereço do Fornecedor e Destinatario
    Call Preenche_Destinatario_Remetente(objFilialCliente)
   
    If ComissaoAutomatica.Value = 1 Then

        Call Grid_Limpa(objGridComissoes)
        
'********************** TRATAMENTO COMISSOES ********************
        lErro = objTabComissoes.Comissao_Automatica_FilialCli_Exibe(objFilialCliente)
        If lErro <> SUCESSO Then gError 126988
'****************************************************************

    End If
   
    Trata_FilialCliente = SUCESSO

    Exit Function

Erro_Trata_FilialCliente:

    Trata_FilialCliente = gErr

    Select Case gErr

        Case 62965, 126988

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154825)
            
    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 62966
        
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

'********************* TRATAMENTO COMISSOES ********************
            'Se for o GridComissoes
            Case GridComissoes.Name

                lErro = objTabComissoes.Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then gError 126991
'****************************************************************

        End Select
        
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 62967

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 62966, 62967, 126991

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154826)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 62968

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal1()
    If lErro <> SUCESSO Then gError 62969
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 62968, 62969

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 154827)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dValorTotal As Double, dValorINSSRetido As Double
Dim objNFiscal As New ClassNFiscal
Dim dFator As Double, dValorIRRF As Double
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(NatOpInterna.Text)) = 0 Then gError 62970
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 62971
    If Len(Trim(Filial.Text)) = 0 Then gError 62972
    If Len(Trim(Serie.Text)) = 0 Then gError 62973
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 62974
    If Len(Trim(FreteValor.Text)) = 0 And Len(Trim(FretePeso.Text)) = 0 Then gError 62976
    If Len(Trim(BaseCalculo.Text)) = 0 Then gError 62977
    If Len(Trim(Remetente.Text)) = 0 Then gError 62975
    If Len(Trim(Destinatario.Text)) = 0 Then gError 62978
    
    dValorTotal = StrParaDbl(ValorTotal.Caption)

'************ TRATAMENTO COMISSOES ****************
    'Valida os dados do grid de comissões
    lErro = objTabComissoes.Valida_Grid_Comissoes()
    If lErro <> SUCESSO Then gError 126989
'**************************************************

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 62979
           
    If INSSRetido.Value = vbChecked Then
        dValorINSSRetido = StrParaDbl(ValorINSS.Text)
    End If
    If dValorTotal < dValorINSSRetido Then gError 81608
    
    'Recolhe os dados da tela
    lErro = Move_Conhecimento_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 62980
    
    If Len(Trim(NFiscal.Caption)) = 0 Then
        'verifica se o cliente tem crédito.
        lErro = CF("NFiscal_Testa_Credito", objNFiscal)
        If lErro <> SUCESSO Then gError 62981
    End If
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataEmissao.Text))
    If lErro <> SUCESSO Then gError 92039
    
    'Grava a Nota Fiscal de Saída(inclusive os dados contábeis)
    lErro = ConhecimentoFrete_Grava(objNFiscal, objContabil)
    If lErro <> SUCESSO Then gError 62982
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
            
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 62099, 126989
                
        Case 62970
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)

        Case 62971
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 62972
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 62973
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 62974
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 62975
            Call Rotina_Erro(vbOKOnly, "ERRO_REMETENTE_NAO_PREENCHIDO", gErr)
        
        Case 62976
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORFRETE_NAO_PREENCHIDO", gErr)
        
        Case 62977
            Call Rotina_Erro(vbOKOnly, "ERRO_BASECALCULO_NAO_PREENCHIDA", gErr)
        
        Case 62978
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINATARIO_NAO_PREENCHIDO", gErr)

        Case 62979, 81608
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)
                            
        Case 62980, 62982, 62981, 92039

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154828)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_NFiscal() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Limpa_Tela_NFiscal

    Call Limpa_Tela(Me)
    
    Status.Caption = ""
    Serie.Text = ""
   
    NFiscal.Caption = ""
    Filial.Clear

    ValorTotal.Caption = ""
    PlacaUF.Text = ""
    UFDestinatario.Text = ""
    UFRemetente.Text = ""

'********* TRATAMENTO COMISSOES ***********
    Call Grid_Limpa(objGridComissoes)
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorBase.Caption = ""
'******************************************

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    NatOpInterna.Text = NaturezaOp_Conv34(NATUREZAOP_PADRAO_CONHECIMENTOFRETE, gdtDataAtual)

    iAlterado = 0
    iClienteAlterado = 0
    iComissoesAlterada = 0
   
    'Incluído por Luiz Nogueira em 21/08/03
    VolumeMarca.Text = ""
    VolumeEspecie.Text = ""
    
    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie)

    'Incluído por Luiz Nogueira em 21/08/03
    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca)
    
    'Fecha o Sistema de Setas
    lErro = ComandoSeta_Fechar(Me.Name)
        
    Exit Function

Erro_Limpa_Tela_NFiscal:
    
    Limpa_Tela_NFiscal = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154829)
    
    End Select
    
    Exit Function
    
End Function

Private Function Limpa_Tela_NFiscal1() As Long
'Limpa a Tela NFiscalEntrada, mas mantém a natureza e o tipo
    
Dim sNatureza As String
Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_NFiscal1

    sNatureza = NatOpInterna.Text
    
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 62983
    
    NatOpInterna.Text = sNatureza
        
    Exit Function

Erro_Limpa_Tela_NFiscal1:
    
    Limpa_Tela_NFiscal1 = gErr
    
    Select Case gErr
    
        Case 62983
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154830)
    
    End Select
    
    Exit Function
    
End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 62984

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscal()
    If lErro <> SUCESSO Then gError 62985
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 62984, 62985

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154831)

    End Select

    Exit Sub

End Sub

'Início contabilidade
Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

    If giTipoVersao = VERSAO_LIGHT Then
        Call objContabil.Contabil_GridContabil_Consulta_Click
    End If

End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Public Sub CTBGridContabil_LeaveCell()

    Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub
Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

'****
Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

'****
Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

'****
Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'Traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim sContaAux As String
Dim sContaMascarada As String
Dim objcliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objMnemonicoCTBValor As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        'Retorna o valor do campo Frete Peso
        Case FRETE_PESO
            'Se o campo foi preenchido
            If Len(Trim(FretePeso.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(FretePeso.Text)
            Else
                'Guarda o valor 0 na coleção
                objMnemonicoValor.colValor.Add 0
            End If
            
        'Retorna o valor do campo Frete Valor
        Case FRETE_VALOR
            'Se o campo foi preenchido
            If Len(Trim(FreteValor.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(FreteValor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo Valor Total
        Case VALOR_TOTAL
            'Se o campo foi preenchido
            If Len(Trim(ValorTotal.Caption)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(ValorTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo Sec/Cat
        Case SECCAT
            'Se o campo foi preenchido
            If Len(Trim(SEC.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(SEC.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo Despacho
        Case DESPACHO1
            'Se o campo foi preenchido
            If Len(Trim(Despacho.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(Despacho.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo Pedágio
        Case PEDAGIO1
            'Se o campo foi preenchido
            If Len(Trim(Pedagio.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(Pedagio.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo Outros Valores
        Case OUTROS_VALORES
            'Se o campo foi preenchido
            If Len(Trim(OutrosValores.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(OutrosValores.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo ICMS
        Case ICMS
            'Se o campo foi preenchido
            If Len(Trim(ValorICMS.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(ValorICMS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o valor do campo ICMS Incluso
        Case ICMS_INCLUSO
            'Guarda na coleção o valor da check box ICMSIncluso
            objMnemonicoValor.colValor.Add ICMSIncluso.Value
        
        'Retorna o valor do campo Base Cálculo
        Case BASE_CALCULO
            'Se o campo foi preenchido
            If Len(Trim(BaseCalculo.Text)) > 0 Then
                'Guarda o valor do campo na coleção
                objMnemonicoValor.colValor.Add StrParaDbl(BaseCalculo.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        'Retorna o código do Cliente selecionado
        Case CLIENTE_CODIGO
            'Se o Cliente foi selecionado
            If Len(Trim(Cliente.Text)) > 0 Then
                
                'Guarda no obj o parâmetro que será passado para Cliente_Le_NomeReduzido
                objcliente.sNomeReduzido = Cliente.Text
                
                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO Then gError 79974
                
                'Guarda na coleção o código do Cliente
                objMnemonicoValor.colValor.Add objcliente.lCodigo
            
            'Se não selecionou o cliente
            Else
                'Adiciona o valor 0
                objMnemonicoValor.colValor.Add 0
            End If
            
        'Retorna o nome / razão social do Cliente selecionado
        Case CLIENTE_RAZAOSOCIAL
            'Se o Cliente foi preenchido
            If Len(Trim(Cliente.Text)) > 0 Then
                
                'Guarda no obj o parâmetro que será passado para Cliente_Le_NomeReduzido
                objcliente.sNomeReduzido = Cliente.Text
                
                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO Then gError 79980
                
                'Guarda na coleção o nome /  razão social do cliente
                objMnemonicoValor.colValor.Add objcliente.sRazaoSocial
            
            'Se não selecionou o cliente
            Else
                'Adiciona uma string vazia
                objMnemonicoValor.colValor.Add ""
            End If
        
        'Tenta retornar a conta contábil de vendas do clientes, caso não encontre retorna a conta de vendas global
        Case CTA_VENDAS
                    
                'Guarda no objCliente o parâmetro que será passado para Cliente_Le_NomeReduzido
                objcliente.sNomeReduzido = Cliente.Text
                
                'Lê os dados do cliente a partir do nome reduzido
                lErro = CF("Cliente_Le_NomeReduzido", objcliente)
                If lErro <> SUCESSO Then gError 79981
                
                'Guarda em objFilialCliente o parâmetro que será passado para
                objFilialCliente.lCodCliente = objcliente.lCodigo
                objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
                
                'Lê os dados da Filial do cliente a partir do codigo do cliente e da filial
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO Then gError 79982
                
                'Se a filial do cliente possui Conta Contábil de Vendas
                If Len(objFilialCliente.sContaContabil) > 0 Then
                
                    'Inicializa a variável que receberá a conta
                    sContaMascarada = String(STRING_CONTA, 0)

                    'Aplica o formato à conta que foi encontrada
                    lErro = Mascara_MascararConta(objFilialCliente.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 79983
                    
                    'Guarda a conta na coleção
                    objMnemonicoValor.colValor.Add sContaMascarada
                
                'Senão => procura a conta de vendas nos campos globais
                Else
                                        
                    'Guarda no obj o parâmetro que será passado para MnemonicoCTBValor_Le
                    objMnemonicoCTBValor.sMnemonico = CTA_VENDAS
                    
                    'Lê no BD os dados do Mnemonico Global
                    lErro = CF("MnemonicoCTBValor_Le", objMnemonicoCTBValor)
                    If lErro <> SUCESSO Then gError 79984
                
                    'Se encontrou uma conta de global de vendas = aplica o formato e guarda a conta na coleção
                    If Len(objMnemonicoCTBValor.sValor) > 0 Then
                    
                        'Guarda no obj a conta encontrada com seu formato
                        objMnemonicoValor.colValor.Add objMnemonicoCTBValor.sValor
                    'Se não encontrou
                    Else
                        'Guarda no obj uma string vazia
                        objMnemonicoValor.colValor.Add ""
                    End If
                
                End If
                
        'Se o mnemônico não foi tratado
        Case Else
            gError 79985
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:
    
    Calcula_Mnemonico = gErr
    
    Select Case gErr
    
        Case 79980, 79981, 79982, 79983, 79984
        
        Case 79985
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154832)
            
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_NF_SAIDA_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Conhecimento de Transporte"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConhecimentoFrete"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is NatOpInterna Then
''''            Call LblNatOpInterna_Click
        ElseIf Me.ActiveControl Is Serie Then
            Call SerieLabel_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFiscalLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call ClienteLabel_Click
        End If
    
    End If

End Sub

Public Sub VolumeQuant_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(VolumeQuant, iAlterado)

End Sub

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    objFilial.Text = CStr(iFilial)
    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then gError 62986

    If lErro = 17660 Then gError 62987

    'Coloca na tela a Filial lida
    objFilial.Text = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case gErr

        Case 62986
        
        Case 62987
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, objFilial.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154833)

    End Select

    Exit Sub

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

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
'***** fim do trecho a ser copiado ******


Private Function ValorTotal_Calcula(Optional bValidateICMSValor As Boolean = False) As Long
'Calcula o total de acordo com os valores informados na tela
'e a opção de ICMS Incluso ou não
Dim dFretePeso As Double
Dim dFreteValor As Double
Dim dSEC As Double
Dim dDespacho As Double
Dim dPedagio As Double
Dim dOutros As Double
Dim dValorBase As Double
Dim dAliquota As Double
Dim dValorTotal As Double
Dim dValorICMS As Double
Dim lErro As Long
    
On Error GoTo Erro_ValorTotal_Calcula
    
    'Recolhe os valores da tela
    dFretePeso = StrParaDbl(FretePeso.Text)
    dFreteValor = StrParaDbl(FreteValor.Text)
    dSEC = StrParaDbl(SEC.Text)
    dDespacho = StrParaDbl(Despacho.Text)
    dPedagio = IIf(PedagioIncluso.Value, StrParaDbl(Pedagio.Text), 0)
    dOutros = StrParaDbl(OutrosValores.Text)
    dValorICMS = StrParaDbl(ValorICMS.Text)
    dValorBase = StrParaDbl(BaseCalculo.Text)
    dAliquota = PercentParaDbl(Aliquota.FormattedText)
    
    'Calcula o Subtotal sem imposto
    dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dPedagio + dOutros
    
    If bValidateICMSValor Then dAliquota = 0
    
    'Se o imposto não é incluido
    If ICMSIncluso.Value = vbChecked Then
        'Se a aliquota estiver preenchida
        If dAliquota > 0 Then
            'Calcula o Valor ICMS com base na aliquota informada
            dValorICMS = dValorTotal * dAliquota
        'Se o VAlor estiver preenchido
        ElseIf dValorICMS > 0 Then
            'se o SUbtotal for positivo
            If dValorTotal > 0 Then
                'Calcula a alíquota
                dAliquota = dValorICMS / dValorTotal
            End If
        End If
    'Se o imposto é incluido
    Else
        'Se a aliquota estiver preenchida
        If dAliquota > 0 Then
            'Calcula o valor ICMS Com base na aliquota informada
            dValorICMS = dValorTotal / ((1 / dAliquota) - 1)
            dValorICMS = StrParaDbl(Format(dValorICMS, "#.00"))
        'Se o ValorICMS estiver preenchido
        ElseIf dValorICMS > 0 Then
            'Se o SUbtotal for positivo
            If dValorTotal > 0 Then
                'Calcula a alíquota
                dAliquota = dValorICMS / (dValorTotal + dValorICMS)
            End If
        End If
        'Inclui no total o ICMS
        dValorTotal = dValorTotal + dValorICMS
    
    End If
    
    If dValorICMS > dValorTotal Then gError 62989
    
    'Coloca na tela os valores calculados
    ValorTotal.Caption = Format(dValorTotal, "Standard")
    If dValorTotal > 0 Then BaseCalculo.Text = Format(dValorTotal, "Standard")
    If dValorICMS > 0 Then ValorICMS.Text = Format(dValorICMS, "Standard")
    If dAliquota > 0 Then Aliquota.Text = Format(dAliquota * 100, "Fixed")
    
    If Not gbCarregandoTela Then
    
'****************** TRATAMENTO COMISSOES *********************
        'Faz o cálculo automático das comissões
        lErro = objTabComissoes.Comissoes_Calcula_Padrao()
        If lErro <> SUCESSO Then gError 126990
'*************************************************************
        
    End If
    
    
    ValorTotal_Calcula = SUCESSO
    
    Exit Function
    
Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = Err
    
    Select Case gErr
    
        Case 62989
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORICMS_MAIOR_TOTAL", gErr, dValorICMS, dValorTotal)

        Case 126990

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154834)
            
    End Select
    
    Exit Function
    
End Function

Function BaseCalculo_Calula_ValorTotal() As Long
'CaLcula o total de acordo com os valores informados na tela
'e a opção de ICMS Incluso ou não
Dim dFretePeso As Double
Dim dFreteValor As Double
Dim dSEC As Double
Dim dDespacho As Double
Dim dPedagio As Double
Dim dOutros As Double
Dim dValorBase As Double
Dim dAliquota As Double
Dim dValorTotal As Double
Dim dValorICMS As Double

On Error GoTo Erro_BaseCalculo_Calula_ValorTotal

    'Recolhe os valores da tela
    dFretePeso = StrParaDbl(FretePeso.Text)
    dFreteValor = StrParaDbl(FreteValor.Text)
    dSEC = StrParaDbl(SEC.Text)
    dDespacho = StrParaDbl(Despacho.Text)
    dPedagio = IIf(PedagioIncluso.Value, StrParaDbl(Pedagio.Text), 0)
    dOutros = StrParaDbl(OutrosValores.Text)
    dValorICMS = StrParaDbl(ValorICMS.Text)
    dValorBase = StrParaDbl(BaseCalculo.Text)
    dAliquota = PercentParaDbl(Aliquota.FormattedText)
    
    'Calcula o Subtotal sem imposto
    dValorTotal = dFretePeso + dFreteValor + dSEC + dDespacho + dPedagio + dOutros
    
    If dValorBase < dValorTotal Then gError 62988
    
    dValorICMS = dValorBase - dValorTotal
    
    If dValorICMS > 0 Then
        ValorICMS.Text = Format(dValorICMS, "Standard")
    Else
        ValorICMS.Text = ""
    End If
    
    Call ValorTotal_Calcula(True)
    
    BaseCalculo_Calula_ValorTotal = SUCESSO
    
    Exit Function
    
Erro_BaseCalculo_Calula_ValorTotal:

    BaseCalculo_Calula_ValorTotal = Err
    
    Select Case gErr
    
        Case 62988
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_MENOR_SUBTOTAL", gErr, dValorBase, dValorTotal)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154835)
    
    End Select
    
    Exit Function

End Function

Function ConhecimentoFrete_Grava(objNFiscal As ClassNFiscal, objContabil As ClassContabil) As Long
'grava uma nota fiscal

Dim lErro As Long
Dim lTransacao As Long
Dim alComando(1 To 18) As Long
Dim iIndice As Integer

On Error GoTo Erro_ConhecimentoFrete_Grava

    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 62845

    For iIndice = LBound(alComando) To UBound(alComando)
        alComando(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 62846
    Next
    
    lErro = NFiscal_Lock_Gravacao(alComando(), objNFiscal)
    If lErro <> SUCESSO Then gError 62847
    
        
    lErro = ConhecimentoFrete_Grava_BD(alComando, objNFiscal)
    If lErro <> SUCESSO And lErro <> 62840 Then gError 62848
    
    'Se a Nota é nova
    If lErro = SUCESSO Then
        
        'verifica se o cliente possui o crédito para faturar a nota fiscal.
        'Se tiver atualiza as tabelas de cliente e estatistica de liberacao do usuario
        lErro = Processa_NFiscal_Credito(objNFiscal)
        If lErro <> SUCESSO Then gError 62849
        
        'Grava a Estatística do Cliente
        lErro = CF("FilialCliente_Grava_Estatistica", objNFiscal)
        If lErro <> SUCESSO Then gError 62850
        
        'Verifica se o modulo de Livros Fiscais está Ativo
        If gcolModulo.Ativo(MODULO_LIVROSFISCAIS) = MODULO_ATIVO Then
                            
            'Grava o Livro Fiscal a partir da Nota Fiscal
            lErro = CF("NotaFiscal_Grava_Fis", objNFiscal)
            If lErro <> SUCESSO Then gError 62851
        
        End If

    End If

    'Grava os dados contábeis (contabilidade)
    lErro = objContabil.Contabil_Gravar_Registro(objNFiscal.lNumIntDoc, objNFiscal.lCliente, objNFiscal.iFilialCli, DATA_NULA, NAO_AVALIA_PELO_CUSTO_REAL_PRODUCAO, NAO_AVISA_LANCAMENTOS_CONTABILIZADOS, objNFiscal.lNumNotaFiscal)
    If lErro <> SUCESSO Then gError 62852

    'Confirma Transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 62853

    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    ConhecimentoFrete_Grava = SUCESSO

    Exit Function

Erro_ConhecimentoFrete_Grava:

    ConhecimentoFrete_Grava = gErr

    Select Case gErr

        Case 62845
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)

        Case 62846
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 62847, 62848, 62849, 62850, 62851, 62852

        Case 62853
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT_TRANSACAO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154836)

    End Select

    Call Transacao_Rollback
    For iIndice = LBound(alComando) To UBound(alComando)
        Call Comando_Fechar(alComando(iIndice))
    Next

    Exit Function

End Function

Function ConhecimentoFrete_Grava_BD(alComando() As Long, objNFiscal As ClassNFiscal) As Long

Dim lNumIntDoc As Long
Dim lErro As Long
Dim iClasseDocCPR As Integer
Dim lNumIntDocCPR As Long
Dim alComando2(1 To 6) As Long
Dim colComissoesEmissao As New colComissao
Dim iIndice As Integer

On Error GoTo Erro_ConhecimentoFrete_Grava_BD

    'copia alguns alComando para alComando2
    For iIndice = LBound(alComando2) To UBound(alComando2)
        alComando2(iIndice) = Comando_Abrir()
        If alComando(iIndice) = 0 Then gError 126999
    Next

    If objNFiscal.lNumNotaFiscal = 0 Then
    
        'Gera o Número para a Nota Fiscal e atualiza a Tabela de Serie
        lErro = CF("NFiscalNumAuto", objNFiscal)
        If lErro <> SUCESSO Then gError 62835
    
        'verifica se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo
        lErro = CF("NFiscal_Testa_Existencia", alComando(10), alComando(11), objNFiscal)
        If lErro <> SUCESSO And lErro <> 42417 Then gError 62836
    
        If lErro = 42417 Then gError 62837
    
    Else
    
        'verifica se a nota fiscal já está cadastrada ou se já existe uma nota com os mesmos dados em um dado periodo
        lErro = CF("NFiscal_Testa_Existencia", alComando(10), alComando(11), objNFiscal)
        If lErro <> SUCESSO And lErro <> 42417 Then gError 62838
        If lErro <> 42417 Then gError 62839
    
        'Tratar a gravação da contabilidade na rotina chamadora
        gError 62840
        
    End If
    
    'Busca iClasseDocCPR e lNumIntDocCPR
    lErro = CF("CPR_Gera", objNFiscal, iClasseDocCPR, lNumIntDocCPR)
    If lErro <> SUCESSO Then gError 62841

    objNFiscal.iClasseDocCPR = iClasseDocCPR
    objNFiscal.lNumIntDocCPR = lNumIntDocCPR
    
    'Obtem o Número Interno da nova Nota Fiscal
    lErro = CF("NFiscal_Automatico1", alComando(12), alComando(13), alComando(14), lNumIntDoc)
    If lErro <> SUCESSO Then gError 62842

    objNFiscal.lNumIntDoc = lNumIntDoc

    With objNFiscal

        'Insere a nova Nota Fiscal de Saida no BD
        lErro = Comando_Executar(alComando(15), "INSERT INTO NFiscal (DataAtualizacao, NumIntDoc, Status, FilialEmpresa, Serie, NumNotaFiscal, Cliente, FilialCli, FilialEntrega, Fornecedor, FilialForn, DataEmissao, DataEntrada, DataSaida, DataVencimento, DataReferencia,NumPedidoVenda, NumPedidoTerc, ClasseDocCPR, NumIntDocCPR, ValorTotal, ValorProdutos, ValorFrete, ValorSeguro, ValorOutrasDespesas, ValorDesconto, CodTransportadora, MensagemNota, TabelaPreco, TipoNFiscal, NaturezaOp, PesoLiq, PesoBruto, NumIntTrib, Placa, PlacaUF, VolumeQuant, VolumeEspecie, VolumeMarca, Canal, NumIntNotaOriginal,FilialPedido, VolumeNumero, FreteRespons,Observacao) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", _
            gdtDataHoje, .lNumIntDoc, .iStatus, .iFilialEmpresa, .sSerie, .lNumNotaFiscal, .lCliente, .iFilialCli, .iFilialEntrega, .lFornecedor, .iFilialForn, .dtDataEmissao, .dtDataEntrada, .dtDataSaida, .dtDataVencimento, .dtDataReferencia, .lNumPedidoVenda, .sNumPedidoTerc, .iClasseDocCPR, .lNumIntDocCPR, .dValorTotal, .dValorProdutos, .dValorFrete, .dValorSeguro, .dValorOutrasDespesas, .dValorDesconto, .iCodTransportadora, .sMensagemNota, .iTabelaPreco, .iTipoNFiscal, .sNaturezaOp, .dPesoLiq, .dPesoBruto, .lNumIntTrib, .sPlaca, .sPlacaUF, .lVolumeQuant, .lVolumeEspecie, .lVolumeMarca, .iCanal, .lNumIntNotaOriginal, .iFilialPedido, .sVolumeNumero, .iFreteRespons, .sObservacao)
        If lErro <> AD_SQL_SUCESSO Then gError 62843

        objNFiscal.objConhecimentoFrete.lNumIntNFiscal = .lNumIntDoc
    End With

    With objNFiscal.objConhecimentoFrete

        lErro = Comando_Executar(alComando(16), "INSERT INTO ConhecimentoFrete (NumIntNFiscal,FretePeso,FreteValor,SEC,Despacho,Pedagio,OutrosValores,Aliquota,ValorICMS,BaseCalculo,PesoMercadoria,ValorMercadoria,NotasFiscais,Coleta,Entrega,CalculadoAte,NaturezaCarga,MarcaVeiculo,LocalVeiculo,Remetente,EnderecoRemetente,MunicipioRemetente,UFRemetente,CepRemetente,CGCRemetente,InscEstadualRemetente,Destinatario,EnderecoDestinatario,MunicipioDestinatario,UFDestinatario,CepDestinatario,CGCDestinatario,InscEstadualDestinatario,ICMSIncluso,INSSRetido,ValorINSS,IncluiPedagio) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ", _
        .lNumIntNFiscal, .dFretePeso, .dFreteValor, .dSEC, .dDespacho, .dPedagio, .dOutrosValores, .dAliquotas, .dValorICMS, .dBaseCalculo, .dPesoMercadoria, .dValorMercadoria, .sNotasFiscais, .sColeta, .sEntrega, .sCalculadoAte, .sNaturezaCarga, .sMarcaVeiculo, .sLocalVeiculo, .sRemetente, .sEnderecoRemetente, .sMunicipioRemetente, .sUFRemetente, .sCepRemetente, .sCGCRemetente, .sInscEstadualRemetente, .sDestinatario, .sEnderecoDestinatario, .sMunicipioDestinatario, .sUFDestinatario, .sCepDestinatario, .sCGCDestinatario, .sInscEstadualDestinatario, .iICMSIncluso, .iINSSRetido, .dValorINSS, .iIncluiPedagio) 'William
        If lErro <> AD_SQL_SUCESSO Then gError 62844

    End With

    lErro = CF("ComissoesNF_Grava", alComando(17), alComando(18), objNFiscal)
    If lErro <> SUCESSO Then gError 86178
    
    'Gera as Comissões na Emissão com base nas comissões armazenadas em objNFiscal e coloca-os em colComissoesEmissao
    lErro = CF("Comissoes_Gera", alComando2(1), alComando2(2), objNFiscal, colComissoesEmissao)
    If lErro <> SUCESSO Then gError 126997

    'Grava as Comissões passadas em colComissao
    lErro = CF("Comissoes_Grava1", alComando2(3), alComando2(4), alComando2(5), alComando2(6), colComissoesEmissao)
    If lErro <> SUCESSO Then gError 126998
    
    For iIndice = LBound(alComando2) To UBound(alComando2)
        Call Comando_Fechar(alComando2(iIndice))
    Next
       
    ConhecimentoFrete_Grava_BD = SUCESSO
    
    Exit Function
    
Erro_ConhecimentoFrete_Grava_BD:

    ConhecimentoFrete_Grava_BD = gErr
    
    Select Case gErr

        Case 62835, 62836, 62838, 62841, 62842, 86178, 126997, 126998

        Case 62837, 62839
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL_NUMAUTO", gErr)
                
        Case 62840
        
        Case 62843
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_NFISCAL", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
            
        Case 62844
            Call Rotina_Erro(vbOKOnly, "ERRO_INCLUSAO_CONHECIMENTOFRETE", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case 126999
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154837)
            
    End Select
    
    Exit Function

End Function

Function ConhecimentoFrete_Le(objConhecimentoFrete As ClassConhecimentoFrete) As Long

Dim lComando As Long
Dim lErro As Long
Dim tConhecimentoFrete As typeConhecimentoFrete

On Error GoTo Erro_ConhecimentoFrete_Le
    
    'Abre o comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 62854

    With tConhecimentoFrete
        'Inicializa as variáveis string
        .sCalculadoAte = String(STRING_CONHECIMENTOFRETE_CALCULOATE, 0)
        .sCepDestinatario = String(STRING_CEP, 0)
        .sCepRemetente = String(STRING_CEP, 0)
        .sCGCDestinatario = String(STRING_CGC, 0)
        .sCGCRemetente = String(STRING_CGC, 0)
        .sColeta = String(STRING_CONHECIMENTOFRETE_COLETA, 0)
        .sDestinatario = String(STRING_CONHECIMENTOFRETE_DESTINATARIO, 0)
        .sEnderecoDestinatario = String(STRING_ENDERECO, 0)
        .sEnderecoRemetente = String(STRING_ENDERECO, 0)
        .sEntrega = String(STRING_CONHECIMENTOFRETE_ENTREGA, 0)
        .sInscEstadualDestinatario = String(STRING_INSCR_EST, 0)
        .sInscEstadualRemetente = String(STRING_INSCR_EST, 0)
        .sLocalVeiculo = String(STRING_CIDADE, 0)
        .sMarcaVeiculo = String(STRING_CONHECIMENTOFRETE_MARCA, 0)
        .sMunicipioDestinatario = String(STRING_CIDADE, 0)
        .sMunicipioRemetente = String(STRING_CIDADE, 0)
        .sNaturezaCarga = String(STRING_CONHECIMENTOFRETE_NATUREZACARGA, 0)
        .sNotasFiscais = String(STRING_CONHECIMENTOFRETE_NOTAS, 0)
        .sRemetente = String(STRING_CONHECIMENTOFRETE_REMETENTE, 0)
        .sUFDestinatario = String(STRING_ESTADO_SIGLA, 0)
        .sUFRemetente = String(STRING_ESTADO_SIGLA, 0)

        'Busca no BD o conhecimento passado
        lErro = Comando_Executar(lComando, "SELECT FretePeso,FreteValor,SEC,Despacho,Pedagio,OutrosValores,Aliquota,ValorICMS,BaseCalculo,PesoMercadoria,ValorMercadoria,NotasFiscais,Coleta,Entrega,CalculadoAte,NaturezaCarga,MarcaVeiculo,LocalVeiculo,Remetente,EnderecoRemetente,MunicipioRemetente,UFRemetente,CepRemetente,CGCRemetente,InscEstadualRemetente,Destinatario,EnderecoDestinatario,MunicipioDestinatario,UFDestinatario,CepDestinatario,CGCDestinatario,InscEstadualDestinatario,ICMSIncluso,IncluiPedagio FROM ConhecimentoFrete WHERE NumIntNFiscal = ?", _
        .dFretePeso, .dFreteValor, .dSEC, .dDespacho, .dPedagio, .dOutrosValores, .dAliquotas, .dValorICMS, .dBaseCalculo, .dPesoMercadoria, .dValorMercadoria, .sNotasFiscais, .sColeta, .sEntrega, .sCalculadoAte, .sNaturezaCarga, .sMarcaVeiculo, .sLocalVeiculo, .sRemetente, .sEnderecoRemetente, .sMunicipioRemetente, .sUFRemetente, .sCepRemetente, .sCGCRemetente, .sInscEstadualRemetente, .sDestinatario, .sEnderecoDestinatario, .sMunicipioDestinatario, .sUFDestinatario, .sCepDestinatario, .sCGCDestinatario, .sInscEstadualDestinatario, .iICMSIncluso, .iIncluiPedagio, objConhecimentoFrete.lNumIntNFiscal)
        If lErro <> AD_SQL_SUCESSO Then gError 62855
        
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 62856
        If lErro <> AD_SQL_SUCESSO Then gError 62857 'Não encontrou
        
        'Carrega o obj com os dados lidos
        objConhecimentoFrete.dAliquotas = .dAliquotas
        objConhecimentoFrete.dBaseCalculo = .dBaseCalculo
        objConhecimentoFrete.dDespacho = .dDespacho
        objConhecimentoFrete.dFretePeso = .dFretePeso
        objConhecimentoFrete.dFreteValor = .dFreteValor
        objConhecimentoFrete.dOutrosValores = .dOutrosValores
        objConhecimentoFrete.dPedagio = .dPedagio
        objConhecimentoFrete.dPesoMercadoria = .dPesoMercadoria
        objConhecimentoFrete.dSEC = .dSEC
        objConhecimentoFrete.dValorICMS = .dValorICMS
        objConhecimentoFrete.dValorMercadoria = .dValorMercadoria
        objConhecimentoFrete.iICMSIncluso = .iICMSIncluso
        objConhecimentoFrete.sCalculadoAte = .sCalculadoAte
        objConhecimentoFrete.sCepDestinatario = .sCepDestinatario
        objConhecimentoFrete.sCepRemetente = .sCepRemetente
        objConhecimentoFrete.sCGCDestinatario = .sCGCDestinatario
        objConhecimentoFrete.sCGCRemetente = .sCGCRemetente
        objConhecimentoFrete.sColeta = .sColeta
        objConhecimentoFrete.sDestinatario = .sDestinatario
        objConhecimentoFrete.sEnderecoDestinatario = .sEnderecoDestinatario
        objConhecimentoFrete.sEnderecoRemetente = .sEnderecoRemetente
        objConhecimentoFrete.sEntrega = .sEntrega
        objConhecimentoFrete.sInscEstadualDestinatario = .sInscEstadualDestinatario
        objConhecimentoFrete.sInscEstadualRemetente = .sInscEstadualRemetente
        objConhecimentoFrete.sLocalVeiculo = .sLocalVeiculo
        objConhecimentoFrete.sMarcaVeiculo = .sMarcaVeiculo
        objConhecimentoFrete.sMunicipioDestinatario = .sMunicipioDestinatario
        objConhecimentoFrete.sMunicipioRemetente = .sMunicipioRemetente
        objConhecimentoFrete.sNaturezaCarga = .sNaturezaCarga
        objConhecimentoFrete.sNotasFiscais = .sNotasFiscais
        objConhecimentoFrete.sRemetente = .sRemetente
        objConhecimentoFrete.sUFDestinatario = .sUFDestinatario
        objConhecimentoFrete.sUFRemetente = .sUFRemetente
        objConhecimentoFrete.iIncluiPedagio = .iIncluiPedagio

    End With
        
    Call Comando_Fechar(lComando)
    
    ConhecimentoFrete_Le = SUCESSO
    
    Exit Function
    
Erro_ConhecimentoFrete_Le:

    ConhecimentoFrete_Le = gErr
    
    Select Case gErr
        
        Case 62854
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
    
        Case 62855, 62856
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CONHECIMENTO_FRETE", gErr)
    
        Case 62857
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154838)
    
    End Select
    
    Call Comando_Fechar(lComando)
        
    Exit Function
    
End Function

'????????? JÁ EXITE EM ROTINAS FAT
Friend Function Processa_NFiscal_Credito(objNFiscal As ClassNFiscal) As Long
'verifica se o cliente possui o crédito para faturar a nota fiscal.
'Se tiver atualiza as tabelas de cliente e estatistica de liberacao do usuario
'IMPORTANTE: TEM QUE SER CHAMADO DENTRO DE TRANSACAO

Dim lErro As Long
Dim objLiberacaoCredito As New ClassLiberacaoCredito
Dim objValorLiberadoCredito As New ClassValorLiberadoCredito
Dim lComando As Long
Dim lComando1 As Long
Dim tCliente As typeCliente
Dim sCodUsuario As String
Dim dValor As Double
Dim dtData As Date
Dim objClienteEstatistica As New ClassFilialClienteEst
Dim bNFPedido As Boolean
Dim iCreditoAprovado As Integer
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Processa_NFiscal_Credito

    lComando = Comando_Abrir()
    If lComando = 0 Then Error 44482

    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 44483

    'Verifica se NFiscal é oriunda de Pedido
    If objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV Or objNFiscal.iTipoNFiscal = DOCINFO_NFISVPV Then
        bNFPedido = True
    Else
        bNFPedido = False
    End If
    
    'Se for testa se Pedido tem crédito aprovado
    If bNFPedido Then
    
        objPedidoVenda.lCodigo = objNFiscal.lNumPedidoVenda
        objPedidoVenda.iFilialEmpresa = objNFiscal.iFilialPedido

        'verifica se o pedido tem credito aprovado
        lErro = CF("BloqueiosPV_Credito_Aprovado_Testa", objPedidoVenda, iCreditoAprovado)
        If lErro <> SUCESSO Then Error 25740

    End If
    
    'se NF não for oriunda de PV ou se crédito não está liberado
    If (Not bNFPedido) Or iCreditoAprovado <> BLOQUEIO_CREDITO_LIBERADO Then
    
        'Lê os saldos e o limite de credito do Cliente
        lErro = Comando_ExecutarLockado(lComando, "SELECT LimiteCredito FROM Clientes WHERE Codigo = ?", tCliente.dLimiteCredito, objNFiscal.lCliente)
        If lErro <> AD_SQL_SUCESSO Then Error 44484
    
        lErro = Comando_BuscarPrimeiro(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 44485
        
        'se não encontrou os dados do cliente
        If lErro = AD_SQL_SEM_DADOS Then Error 44486
    
        'loca o cliente
        lErro = Comando_LockExclusive(lComando)
        If lErro <> AD_SQL_SUCESSO Then Error 44487
            
        'Passa chave de objClienteEstatistica
        objClienteEstatistica.lCodCliente = objNFiscal.lCliente
        objClienteEstatistica.iFilialEmpresa = objNFiscal.iFilialEmpresa
        
        'Le dinamicamente o Saldo dos Titulos e dos Pedidos Liberados
        lErro = CF("Cliente_Le_Estatistica_Credito", objClienteEstatistica)
        If lErro <> SUCESSO Then Error 52955
            
        'Verifica se a soma dos creditos em Pedidos, Titulos e NFs ultrapassa o limite de Credito do Cliente
        If tCliente.dLimiteCredito < (objClienteEstatistica.dSaldoTitulos + objClienteEstatistica.dSaldoPedidosLiberados + objClienteEstatistica.dValorNFsNaoFaturadas + IIf(bNFPedido, 0, objNFiscal.dValorTotal)) Then
        
            'se um usuário não autorizou o credito ==> erro
            If Len(objNFiscal.sCodUsuario) = 0 Then Error 44488
                
                objLiberacaoCredito.sCodUsuario = objNFiscal.sCodUsuario
            
            If giTipoVersao = VERSAO_FULL Then
            
                'verificar se o usuário tem autorizacao para liberar o valor
                lErro = CF("LiberacaoCredito_Lock", objLiberacaoCredito)
                If lErro <> SUCESSO And lErro <> 44479 Then Error 44489
                
                'se não foi encontrado autorização para o usuario liberar credito
                If lErro = 44479 Then Error 44490
                
                'se o valor da nota ultrapassar o limite de credito que o usuario pode conceder por operacao
                If objNFiscal.dValorTotal > objLiberacaoCredito.dLimiteOperacao Then Error 44491
                
                objValorLiberadoCredito.sCodUsuario = objNFiscal.sCodUsuario
                objValorLiberadoCredito.iAno = Year(gdtDataAtual)
                
                'Lê a estatistica de liberação de credito de um usuario em um determinado ano
                lErro = CF("ValorLiberadoCredito_Lock", objValorLiberadoCredito)
                If lErro <> SUCESSO And lErro <> 44470 Then Error 44492
                
                'se o valor da nota ultrapassar o valor mensal que o usuario tem capacidade de liberar
                If objNFiscal.dValorTotal > objLiberacaoCredito.dLimiteMensal - objValorLiberadoCredito.adValorLiberado(Month(gdtDataAtual)) Then Error 44493
            
                sCodUsuario = objValorLiberadoCredito.sCodUsuario
            
            ElseIf giTipoVersao = VERSAO_LIGHT Then
            
                sCodUsuario = objNFiscal.sCodUsuario
            
            End If
            
            dValor = objNFiscal.dValorTotal
            dtData = gdtDataAtual
            
            'Atualiza a estatistica de liberação de credito do usuario
            lErro = CF("ValorLiberadoCredito_Grava", sCodUsuario, dValor, dtData)
            If lErro <> SUCESSO Then Error 44494
            
        End If
        
    End If
    
    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
        
    Processa_NFiscal_Credito = SUCESSO
    
    Exit Function
    
Erro_Processa_NFiscal_Credito:

    Processa_NFiscal_Credito = Err
    
    Select Case Err

        Case 44482, 44483
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 44484, 44485
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_CLIENTES1", Err, objNFiscal.lCliente)

        Case 44486
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objNFiscal.lCliente)

        Case 44487
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_CLIENTES", Err, objNFiscal.lCliente)
        
        Case 44488
           Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_CREDITO", Err, objNFiscal.lCliente)

        Case 25740, 44489, 44492, 44494, 52955

        Case 44490
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_INEXISTENTE", Err, objLiberacaoCredito.sCodUsuario)
        
        Case 44491
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEOPERACAO", Err, objLiberacaoCredito.sCodUsuario)
    
        Case 44493
            Call Rotina_Erro(vbOKOnly, "ERRO_LIBERACAOCREDITO_LIMITEMENSAL", Err, objLiberacaoCredito.sCodUsuario)
        
        Case 44495
            Call Rotina_Erro(vbOKOnly, "ERRO_MODIFICACAO_CLIENTE", Err, objNFiscal.lCliente)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154839)

    End Select

    Call Comando_Fechar(lComando)
    Call Comando_Fechar(lComando1)
    
    Exit Function

End Function
Private Function NFiscal_Lock_Gravacao(alComando() As Long, objNFiscal As ClassNFiscal) As Long
'executa o lock dos registros associados a nota fiscal.

Dim lErro As Long
Dim iIndice As Long
Dim sNome As String
Dim sDescricao As String
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iFilialEmpresa As Integer
Dim objItemNF As ClassItemNF
Dim objProduto As New ClassProduto
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim iAlmoxarifadoLocado As Integer
Dim objItemAloc As ClassItemNFAlocacao
Dim objComissaoNF As ClassComissaoNF

On Error GoTo Erro_NFiscal_Lock_Gravacao

    lErro = CF("NaturezaOperacao_Lock", objNFiscal.sNaturezaOp, alComando(1))
    If lErro <> SUCESSO And lErro <> 40551 Then gError 35974
    
    If lErro = 40551 Then gError 35975
   
    lErro = CF("FilialCliente_Lock1", alComando(2), objNFiscal.lCliente, objNFiscal.iFilialCli)
    If lErro <> SUCESSO And lErro <> 43050 Then gError 35977
    
    'Não achou a Filial do Cliente --> erro
    If lErro = 43050 Then gError 43052

    lErro = CF("Serie_Lock", alComando(3), objNFiscal.sSerie)
    If lErro <> SUCESSO And lErro <> 43032 Then gError 35978
    
    'Não encontrou a Série --> erro
    If lErro = 43032 Then gError 43034
    
    If objNFiscal.iTabelaPreco > 0 Then
        lErro = CF("NFiscal_TabelaPreco_Lock", alComando(9), objNFiscal)
        If lErro <> SUCESSO Then gError 44292
    End If
    
    'Se a Transportadora estiver preenchida
    If objNFiscal.iCodTransportadora > 0 Then
        
        lErro = CF("Transportadora_Lock", alComando(4), objNFiscal.iCodTransportadora)
        If lErro <> SUCESSO And lErro <> 43038 Then gError 35979
        
        'Não encontrou a Transportadora --> erro
        If lErro = 43038 Then gError 43040
        
    End If
    
    'Se o Canal de Venda estiver preenchido
    If objNFiscal.iCanal <> 0 Then
    
        lErro = CF("CanalVenda_Lock", alComando(5), objNFiscal.iCanal)
        If lErro <> SUCESSO And lErro <> 43026 Then gError 43022
        
        'Não encontrou o Canal de Venda --> erro
        If lErro = 43026 Then gError 43028
        
    End If
    
    For Each objComissaoNF In objNFiscal.ColComissoesNF
        lErro = CF("Vendedor_Lock", alComando(8), objComissaoNF.iCodVendedor)
        If lErro <> SUCESSO Then gError 51581
    Next
    
    'Lock shared nos Produtos e Almoxarifados
    For Each objItemNF In objNFiscal.ColItensNF
        
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Lock_Shared", alComando(6), objProduto)
        If lErro <> SUCESSO Then gError 35990

        objItemNF.iControleEstoque = objProduto.iControleEstoque
        objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
        objItemNF.sUMVenda = objProduto.sSiglaUMVenda
        objItemNF.iClasseUM = objProduto.iClasseUM
        objItemNF.iApropriacaoProd = objProduto.iApropriacaoCusto

        If objItemNF.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            For Each objItemAloc In objItemNF.colAlocacoes
    
                iAlmoxarifadoLocado = False
                
                If colCodigoNome.Count > 0 Then
                    For Each objCodigoNome In colCodigoNome
                        If objCodigoNome.iCodigo = objItemAloc.iAlmoxarifado Then
                            iAlmoxarifadoLocado = True
                        End If
                    Next
                End If
                    
                If iAlmoxarifadoLocado = False Then
                
                    colCodigoNome.Add objItemAloc.iAlmoxarifado, objItemAloc.sAlmoxarifado
                
                    lErro = Comando_ExecutarLockado(alComando(7), "SELECT FilialEmpresa FROM Almoxarifado WHERE Codigo = ? AND FilialEmpresa = ? ", iFilialEmpresa, objItemAloc.iAlmoxarifado, giFilialEmpresa)
                    If lErro <> AD_SQL_SUCESSO Then gError 35991
        
                    lErro = Comando_BuscarPrimeiro(alComando(7))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 35994
                    If lErro = AD_SQL_SEM_DADOS Then gError 35992
        
                    lErro = Comando_LockShared(alComando(7))
                    If lErro <> AD_SQL_SUCESSO Then gError 35993

                End If
            Next
        End If
    Next

    NFiscal_Lock_Gravacao = SUCESSO

    Exit Function

Erro_NFiscal_Lock_Gravacao:

    NFiscal_Lock_Gravacao = Err

    Select Case gErr

        Case 35974, 35975
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NATUREZAOP", gErr, objNFiscal.sNaturezaOp)

        Case 35976
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", gErr, objNFiscal.sNaturezaOp)

        Case 35977
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_NATUREZAOP", gErr, objNFiscal.sNaturezaOp)

        Case 35979
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_FILIAISCLIENTES1", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case 35990, 43022, 35978, 44292, 51581

        Case 35991, 35994
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ALMOXARIFADO", gErr, objItemNF.iAlmoxarifado)

        Case 35992
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objItemAloc.iAlmoxarifado)

        Case 35993
            Call Rotina_Erro(vbOKOnly, "ERRO_LOCK_ALMOXARIFADO1", gErr, objItemNF.iAlmoxarifado)

        Case 43028
            Call Rotina_Erro(vbOKOnly, "ERRO_CANALVENDA_NAO_CADASTRADO", gErr, objNFiscal.iCanal)

        Case 43034
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objNFiscal.sSerie)

        Case 43040
            Call Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objNFiscal.iCodTransportadora)

        Case 43052
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", gErr, objNFiscal.lCliente, objNFiscal.iFilialCli)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154840)

    End Select

    Exit Function

End Function

Private Sub Limpa_Remetente()

    Remetente.Text = ""
    EnderecoRemetente.Text = ""
    CidadeRemetente.Text = ""
    UFRemetente.Text = ""
    CEPRemetente.PromptInclude = False
    CEPRemetente.Text = ""
    CEPRemetente.PromptInclude = True
    CGCRemetente.PromptInclude = False
    CGCRemetente.Text = ""
    CGCRemetente.PromptInclude = True
    InscEstRemetente.PromptInclude = False
    InscEstRemetente.Text = ""
    InscEstRemetente.PromptInclude = True

End Sub

Private Sub Limpa_Destinatario()

    Destinatario.Text = ""
    EnderecoDestinatario.Text = ""
    CidadeDestinatario.Text = ""
    UFDestinatario.Text = ""
    CEPDestinatario.PromptInclude = False
    CEPDestinatario.Text = ""
    CEPDestinatario.PromptInclude = True
    CGCDestinatario.PromptInclude = False
    CGCDestinatario.Text = ""
    CGCDestinatario.PromptInclude = True
    InscEstDestinatario.PromptInclude = False
    InscEstDestinatario.Text = ""
    InscEstDestinatario.PromptInclude = True

End Sub

Public Function Preenche_Destinatario_Remetente(objFilialCliente As ClassFilialCliente) As Long
'Preenche o endereço do remetente e destinatario

Dim lErro As Long
Dim objEndereco As New ClassEndereco
Dim objcliente As New ClassCliente

On Error GoTo Erro_Preenche_Destinatario_Remetente

    'Atribui código ao objCliente
    objcliente.lCodigo = objFilialCliente.lCodCliente

    'Faz leitura na tabela Cliente afim de extrair o nome
    lErro = CF("Cliente_Le", objcliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 87323

    'Se não achou cliente então fornece erro
    If lErro = 12293 Then gError 87329
    
    lErro = CF("FilialCliente_Le", objFilialCliente)
    If lErro <> SUCESSO Then gError 87321
    
    If objFilialCliente.lEnderecoEntrega > 0 Then
        objEndereco.lCodigo = objFilialCliente.lEnderecoEntrega
    
        'Faz leitura na tabela de endereços
        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError 87322

        'Se não achou endereço - erro
        If lErro = 12309 Then gError 87330

        If Len(Trim(objEndereco.sEndereco)) > 0 Then

            'Traz os dados do destinatário para tela
            Call Traz_Destinatario_Tela(objEndereco, objcliente)

            'Traz o dados do remetente para a tela
            Call Traz_Remetente_Tela(objEndereco, objcliente)

        Else

            If objFilialCliente.lEndereco > 0 Then
                objEndereco.lCodigo = objFilialCliente.lEndereco
    
                'Le o endereço com o código passado
                lErro = CF("Endereco_Le", objEndereco)
                If lErro <> SUCESSO And lErro <> 12309 Then gError 87325

                'Se não achou endereço - erro
                If lErro = 12309 Then gError 87328

                'Traz os dados do destinatário para tela
                Call Traz_Destinatario_Tela(objEndereco, objcliente)

                'Traz o dados do remetente para a tela
                Call Traz_Remetente_Tela(objEndereco, objcliente)

            End If

        End If

    End If
    
Preenche_Destinatario_Remetente = SUCESSO

    Exit Function
    
Erro_Preenche_Destinatario_Remetente:

    Preenche_Destinatario_Remetente = gErr
    
    Select Case gErr
    
        Case 87321, 87322, 87323, 87325
        'Erros tratado na rotina
        
        Case 87328, 87329, 87330
        'não fornecer mensagem ao usuário, caso não é obrigatório
        'a existencia de um endereço cadastrado
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154841)

        End Select
        
    Exit Function

End Function


Public Function Traz_Destinatario_Tela(objEndereco As ClassEndereco, objcliente As ClassCliente) As Long

    Call Limpa_Destinatario

    With objEndereco

        If Len(Trim(.sEndereco)) > 0 Then
            EnderecoDestinatario.Text = .sEndereco
        End If
        
        If Len(Trim(.sCidade)) > 0 Then
            CidadeDestinatario.Text = .sCidade
        End If
        
        If Len(Trim(.sSiglaEstado)) > 0 Then
            UFDestinatario.Text = .sSiglaEstado
        End If
        
        If Len(Trim(.sCEP)) > 0 Then
            CEPDestinatario.PromptInclude = False
            CEPDestinatario.Text = .sCEP
            CEPDestinatario.PromptInclude = True
        End If
        
    End With

    With objcliente

        Destinatario.Text = .sRazaoSocial
        
        If Len(Trim(.sCgc)) > 0 Then
            CGCDestinatario.Text = .sCgc
            Call CGCDestinatario_Validate(bSGECancelDummy)
        End If
        
        If Len(Trim(.sInscricaoEstadual)) > 0 Then
            InscEstDestinatario.PromptInclude = False
            InscEstDestinatario.Text = .sInscricaoEstadual
            InscEstDestinatario.PromptInclude = True
        End If

    End With

End Function

Public Function Traz_Remetente_Tela(objEndereco As ClassEndereco, objcliente As ClassCliente) As Long

    Call Limpa_Remetente

    With objEndereco

        If Len(Trim(.sEndereco)) > 0 Then
            EnderecoRemetente.Text = .sEndereco
        End If
        
        If Len(Trim(.sCidade)) > 0 Then
            CidadeRemetente.Text = .sCidade
        End If
        
        If Len(Trim(.sSiglaEstado)) > 0 Then
            UFRemetente.Text = .sSiglaEstado
        End If
        
        If Len(Trim(.sCEP)) > 0 Then
            CEPRemetente.PromptInclude = False
            CEPRemetente.Text = .sCEP
            CEPRemetente.PromptInclude = True
        End If
        
    End With

    With objcliente

        Remetente.Text = .sRazaoSocial
        
        If Len(Trim(.sCgc)) > 0 Then
            CGCRemetente.Text = .sCgc
            Call CGCRemetente_Validate(bSGECancelDummy)
        End If
        
        If Len(Trim(.sInscricaoEstadual)) > 0 Then
            InscEstRemetente.PromptInclude = False
            InscEstRemetente.Text = .sInscricaoEstadual
            InscEstRemetente.PromptInclude = True
        End If

    End With

End Function


Private Sub ValorINSS_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorINSS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ValorINSS_Validate

    If iValorINSSAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(ValorINSS.ClipText)) <> 0 Then
    
        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorINSS.Text)
        If lErro <> SUCESSO Then Error 26144
    
        'Põe o valor formatado na tela
        ValorINSS.Text = Format(ValorINSS.Text, "Standard")
        
    End If
    
    Exit Sub

Erro_ValorINSS_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 49752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154842)

    End Select

    Exit Sub

End Sub

Private Sub INSSRetido_Click()

On Error GoTo Erro_INSSRetido_Click

    If gbCarregandoTela Then Exit Sub
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_INSSRetido_Click:

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154843)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub ValorINSS_GotFocus()
    iValorINSSAlterado = 0
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeMarca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeMarca_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEMARCA, VolumeMarca, "AVISO_CRIAR_VOLUMEMARCA")
    If lErro <> SUCESSO Then gError 102417
    
    Exit Sub

Erro_VolumeMarca_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102417
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154844)

    End Select

End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

'Incluído por Luiz Nogueira em 21/08/03
Public Sub VolumeEspecie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_VolumeEspecie_Validate

    lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_VOLUMEESPECIE, VolumeEspecie, "AVISO_CRIAR_VOLUMEESPECIE")
    If lErro <> SUCESSO Then gError 102416
    
    Exit Sub

Erro_VolumeEspecie_Validate:

    Cancel = True
    
    Select Case gErr

        Case 102416
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154845)

    End Select

End Sub


'***************** TRATAMENTO COMISSOES ********************
Public Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    'Se alguma comissao foi excluída
    If KeyCode = vbKeyDelete Then
        'atualiza os painéis totalizadores
        Call objTabComissoes.Soma_Percentual
        Call objTabComissoes.Soma_Valor
    End If

End Sub
Public Sub ComissaoAutomatica_Click()

Dim lErro As Long

On Error GoTo Erro_ComissaoAutomatica_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se a comissão automática estiver selecionada
    If ComissaoAutomatica.Value = vbChecked Then
        'Recalcula as comissoes
        lErro = objTabComissoes.Comissoes_Calcula_Padrao()
        If lErro <> SUCESSO Then gError 51616
        
    End If
    
    Exit Sub

Erro_ComissaoAutomatica_Click:

    Select Case gErr
    
        Case 51616
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154846)
            
    End Select
    
    Exit Sub
End Sub
Public Sub PercentualComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PercentualComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub PercentualComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PercentualEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub
Public Sub PercentualEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub PercentualEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorEmissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub


Public Sub BotaoVendedores_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoVendedores_Click
    
    'Chama a tela de browse de Vendedores
    lErro = objTabComissoes.BotaoVendedores_Click()
    If lErro <> SUCESSO Then gError 43696
    
    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case gErr
    
        Case 43696
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154847)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long

On Error GoTo Erro_objEventoVendedor_evSelecao
    
    Set objVendedor = obj1
    
    If GridComissoes.Row > 0 Then
    
        'Preenche a linha de Vendedor com dados default
        lErro = objTabComissoes.VendedorLinha_Preenche(objVendedor)
        If lErro <> SUCESSO Then gError 51617
        
    End If


    Me.Show
    
    Exit Sub

Erro_objEventoVendedor_evSelecao:

    Select Case gErr

        Case 51617  'tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154848)

    End Select

    Exit Sub

End Sub
Public Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Public Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

End Sub

Public Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Public Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub

'******************************************
'4 eventos do controle do Grid de Comissoes: DiretoIndireto
'******************************************

'Alterado por Tulio em 02/04

Private Sub DiretoIndireto_Change()

    '*** 19/04/02 - Tulio ***
    'Desmarca o cálculo automático de comissões
    ComissaoAutomatica.Value = vbUnchecked
    '************************************
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiretoIndireto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Private Sub DiretoIndireto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Private Sub DiretoIndireto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = DiretoIndireto
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'*************************************************************

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objcliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objcliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objcliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134011

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154849)

    End Select
    
    Exit Sub

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




