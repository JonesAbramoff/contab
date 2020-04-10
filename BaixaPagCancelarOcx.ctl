VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaPagCancelarOcx 
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   9405
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   2
      Left            =   180
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4560
         TabIndex        =   141
         Tag             =   "1"
         Top             =   1440
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
         Left            =   6330
         TabIndex        =   22
         Top             =   375
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
         Left            =   6330
         TabIndex        =   19
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   900
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
         Left            =   7770
         TabIndex        =   21
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   29
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
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   31
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4230
         MaxLength       =   150
         TabIndex        =   30
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   33
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   87
         Top             =   3450
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
            Left            =   240
            TabIndex        =   90
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
            Left            =   1125
            TabIndex        =   88
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   89
            Top             =   300
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   91
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
         Left            =   3495
         TabIndex        =   20
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   25
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
         Left            =   3450
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   18
         Top             =   540
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
         Left            =   5595
         TabIndex        =   17
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
         TabIndex        =   16
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
         Left            =   15
         TabIndex        =   32
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
         Left            =   6330
         TabIndex        =   34
         Top             =   1515
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
         Left            =   6330
         TabIndex        =   35
         Top             =   1515
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
         Left            =   6360
         TabIndex        =   23
         Top             =   690
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
         TabIndex        =   73
         Top             =   180
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   74
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
         TabIndex        =   81
         Top             =   615
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   82
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   83
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
         TabIndex        =   120
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
         TabIndex        =   92
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
         TabIndex        =   119
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
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
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1800
         TabIndex        =   84
         Top             =   3060
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   86
         Top             =   3045
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   85
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
         TabIndex        =   77
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2700
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   75
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5100
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   76
         Top             =   180
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   1
      Left            =   150
      TabIndex        =   0
      Top             =   705
      Width           =   9105
      Begin VB.Frame FramePagamento 
         Caption         =   "Cheque Pré de Terceiros"
         Height          =   1770
         Index           =   3
         Left            =   2550
         TabIndex        =   126
         Top             =   2985
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label Label20 
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
            Height          =   255
            Left            =   105
            TabIndex        =   140
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label ClienteCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   795
            TabIndex        =   139
            Top             =   1290
            Width           =   3255
         End
         Begin VB.Label NumeroCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   825
            TabIndex        =   138
            Top             =   795
            Width           =   1455
         End
         Begin VB.Label AgenciaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2580
            TabIndex        =   137
            Top             =   315
            Width           =   1125
         End
         Begin VB.Label FilialEmpresaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5640
            TabIndex        =   136
            Top             =   1290
            Width           =   525
         End
         Begin VB.Label ValorCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4440
            TabIndex        =   135
            Top             =   795
            Width           =   1740
         End
         Begin VB.Label ContaCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4425
            TabIndex        =   134
            Top             =   315
            Width           =   1740
         End
         Begin VB.Label BancoCT 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   810
            TabIndex        =   133
            Top             =   315
            Width           =   825
         End
         Begin VB.Label Label19 
            Caption         =   "Filial Empresa:"
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
            Left            =   4305
            TabIndex        =   132
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label18 
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
            Left            =   3840
            TabIndex        =   131
            Top             =   855
            Width           =   510
         End
         Begin VB.Label Label15 
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
            Left            =   60
            TabIndex        =   130
            Top             =   855
            Width           =   705
         End
         Begin VB.Label Label50 
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
            Height          =   255
            Left            =   3810
            TabIndex        =   129
            Top             =   345
            Width           =   600
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   1800
            TabIndex        =   128
            Top             =   345
            Width           =   810
         End
         Begin VB.Label Label1 
            Caption         =   "Banco:"
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
            Left            =   165
            TabIndex        =   127
            Top             =   345
            Width           =   660
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Crédito"
         Height          =   1770
         Index           =   2
         Left            =   2550
         TabIndex        =   93
         Top             =   2985
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label SaldoCredito 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
            Height          =   300
            Left            =   1980
            TabIndex        =   105
            Top             =   1305
            Width           =   1080
         End
         Begin VB.Label NumTitulo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero"
            Height          =   300
            Left            =   1950
            TabIndex        =   104
            Top             =   840
            Width           =   720
         End
         Begin VB.Label FilialEmpresaCR 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FilEmpr"
            Height          =   300
            Left            =   4560
            TabIndex        =   103
            Top             =   1305
            Width           =   525
         End
         Begin VB.Label ValorCredito 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Valor"
            Height          =   300
            Left            =   4560
            TabIndex        =   102
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label SiglaDocumentoCR 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SiglaDoc"
            Height          =   300
            Left            =   4560
            TabIndex        =   101
            Top             =   382
            Width           =   705
         End
         Begin VB.Label DataEmissaoCred 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DataEmissao"
            Height          =   300
            Left            =   1950
            TabIndex        =   100
            Top             =   382
            Width           =   1095
         End
         Begin VB.Label Label40 
            Caption         =   "Filial Empresa:"
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
            Left            =   3210
            TabIndex        =   99
            Top             =   1335
            Width           =   1245
         End
         Begin VB.Label Label39 
            Caption         =   "Saldo:"
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
            Left            =   1260
            TabIndex        =   98
            Top             =   1335
            Width           =   555
         End
         Begin VB.Label Label38 
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
            Left            =   3945
            TabIndex        =   97
            Top             =   870
            Width           =   510
         End
         Begin VB.Label Label37 
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
            Left            =   1110
            TabIndex        =   96
            Top             =   870
            Width           =   705
         End
         Begin VB.Label Label48 
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
            Height          =   255
            Left            =   4005
            TabIndex        =   95
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label34 
            Caption         =   "Data Emissão:"
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
            Left            =   570
            TabIndex        =   94
            Top             =   405
            Width           =   1245
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Adiantamento à Fornecedor"
         Height          =   1770
         Index           =   1
         Left            =   2550
         TabIndex        =   60
         Top             =   2985
         Visible         =   0   'False
         Width           =   6375
         Begin VB.Label Label5 
            Caption         =   "Data Movimto:"
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
            Left            =   390
            TabIndex        =   61
            Top             =   420
            Width           =   1245
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   3225
            TabIndex        =   63
            Top             =   420
            Width           =   1350
         End
         Begin VB.Label Label7 
            Caption         =   "Meio Pagto:"
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
            Left            =   600
            TabIndex        =   65
            Top             =   885
            Width           =   1035
         End
         Begin VB.Label Label16 
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
            Left            =   3870
            TabIndex        =   67
            Top             =   885
            Width           =   705
         End
         Begin VB.Label Label23 
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
            Left            =   1125
            TabIndex        =   69
            Top             =   1350
            Width           =   510
         End
         Begin VB.Label Label33 
            Caption         =   "Filial Empresa:"
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
            Left            =   3330
            TabIndex        =   71
            Top             =   1350
            Width           =   1245
         End
         Begin VB.Label DataMovimento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DataMovto"
            Height          =   300
            Left            =   1740
            TabIndex        =   62
            Top             =   390
            Width           =   1095
         End
         Begin VB.Label MeioPagtoDescricao 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MeioPagto"
            Height          =   300
            Left            =   1740
            TabIndex        =   66
            Top             =   855
            Width           =   960
         End
         Begin VB.Label ValorPA 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ValorPagtoAnt"
            Height          =   300
            Left            =   1740
            TabIndex        =   70
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Label FilialEmpresaPA 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FilEmpr"
            Height          =   300
            Left            =   4695
            TabIndex        =   72
            Top             =   1320
            Width           =   525
         End
         Begin VB.Label CCIntNomeReduzido 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CCorrente"
            Height          =   300
            Left            =   4695
            TabIndex        =   64
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label NumeroMP 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero"
            Height          =   300
            Left            =   4695
            TabIndex        =   68
            Top             =   855
            Width           =   720
         End
      End
      Begin VB.Frame FramePagamento 
         Caption         =   "Dados do Pagamento"
         Height          =   1770
         Index           =   0
         Left            =   2550
         TabIndex        =   106
         Top             =   2985
         Width           =   6375
         Begin VB.Frame Frame10 
            Caption         =   "Meio Pagamento"
            Height          =   585
            Left            =   2655
            TabIndex        =   107
            Top             =   195
            Width           =   3648
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Borderô"
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
               Height          =   204
               Index           =   1
               Left            =   1272
               TabIndex        =   13
               Top             =   288
               Width           =   996
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Cheque"
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   255
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton TipoMeioPagto 
               Caption         =   "Dinheiro"
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
               Height          =   255
               Index           =   2
               Left            =   2532
               TabIndex        =   14
               Top             =   285
               Width           =   1035
            End
         End
         Begin VB.Label Portador 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4500
            TabIndex        =   117
            Top             =   1350
            Width           =   1695
         End
         Begin VB.Label NumOuSequencial 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4500
            TabIndex        =   116
            Top             =   885
            Width           =   735
         End
         Begin VB.Label Historico 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   960
            TabIndex        =   115
            Top             =   1350
            Width           =   1845
         End
         Begin VB.Label ValorPagoPagto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   960
            TabIndex        =   114
            Top             =   885
            Width           =   1590
         End
         Begin VB.Label ContaCorrente 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   960
            TabIndex        =   113
            Top             =   405
            Width           =   1590
         End
         Begin VB.Label Label32 
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
            Height          =   240
            Left            =   3720
            TabIndex        =   112
            Top             =   915
            Width           =   750
         End
         Begin VB.Label Label25 
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
            Left            =   405
            TabIndex        =   111
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Local de Pagto:"
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
            Left            =   3105
            TabIndex        =   110
            Top             =   1380
            Width           =   1365
         End
         Begin VB.Label Label13 
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
            Height          =   255
            Left            =   90
            TabIndex        =   109
            Top             =   1380
            Width           =   810
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   345
            TabIndex        =   108
            Top             =   435
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Baixa"
         Height          =   1770
         Left            =   45
         TabIndex        =   59
         Top             =   2985
         Width           =   2385
         Begin VB.OptionButton Pagamento 
            Caption         =   "Cheque de Terceiros"
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
            Height          =   210
            Index           =   3
            Left            =   120
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   1455
            Value           =   -1  'True
            Width           =   2115
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Pagamento"
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
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   285
            Width           =   1290
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Adiantamento à Forn."
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
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   680
            Width           =   2145
         End
         Begin VB.OptionButton Pagamento 
            Caption         =   "Crédito/Devolução"
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
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1075
            Width           =   1950
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dados da Baixa"
         Height          =   1170
         Left            =   30
         TabIndex        =   46
         Top             =   1815
         Width           =   8880
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Valor Baixado:"
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
            Left            =   5610
            TabIndex        =   51
            Top             =   375
            Width           =   1245
         End
         Begin VB.Label ValorBaixado 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6930
            TabIndex        =   52
            Top             =   322
            Width           =   945
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desconto:"
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
            TabIndex        =   53
            Top             =   795
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Multa:"
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
            Left            =   3660
            TabIndex        =   55
            Top             =   795
            Width           =   540
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
            Left            =   6360
            TabIndex        =   57
            Top             =   795
            Width           =   525
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Valor Pago:"
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
            Left            =   3195
            TabIndex        =   49
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label Desconto 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   54
            Top             =   742
            Width           =   945
         End
         Begin VB.Label Multa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   56
            Top             =   735
            Width           =   945
         End
         Begin VB.Label Juros 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   6930
            TabIndex        =   58
            Top             =   742
            Width           =   945
         End
         Begin VB.Label ValorPago 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4260
            TabIndex        =   50
            Top             =   322
            Width           =   945
         End
         Begin VB.Label Label17 
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
            Left            =   900
            TabIndex        =   47
            Top             =   375
            Width           =   480
         End
         Begin VB.Label DataBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1470
            TabIndex        =   48
            Top             =   322
            Width           =   1125
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificador da Baixa de Parcela"
         Height          =   1770
         Left            =   30
         TabIndex        =   37
         Top             =   45
         Width           =   8880
         Begin VB.CommandButton BotaoTrazer 
            Caption         =   "Trazer Parcela"
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
            Left            =   6510
            TabIndex        =   8
            Top             =   1268
            Width           =   1650
         End
         Begin VB.ComboBox SiglaDocumento 
            Height          =   315
            Left            =   1275
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   795
            Width           =   2190
         End
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   4395
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   315
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Sequencial 
            Height          =   300
            Left            =   4395
            TabIndex        =   7
            Top             =   1290
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   300
            Left            =   4395
            TabIndex        =   5
            Top             =   795
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   9
            Mask            =   "999999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Parcela 
            Height          =   300
            Left            =   1260
            TabIndex        =   6
            Top             =   1290
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   2
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1275
            TabIndex        =   1
            Top             =   315
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownEmissao 
            Height          =   300
            Left            =   8430
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataEmissao 
            Height          =   300
            Left            =   7350
            TabIndex        =   3
            Top             =   315
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Emissao:"
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
            Left            =   6495
            TabIndex        =   42
            Top             =   375
            Width           =   765
         End
         Begin VB.Label SiglaDocumentoLabel 
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
            Left            =   735
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   870
            Width           =   450
         End
         Begin VB.Label SequencialLabel 
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
            Left            =   3315
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   45
            Top             =   1350
            Width           =   1020
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
            Left            =   495
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   44
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label NumeroTitulo 
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
            Left            =   3615
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   41
            Top             =   855
            Width           =   720
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   180
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   38
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label12 
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
            Left            =   3870
            TabIndex        =   39
            Top             =   375
            Width           =   465
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   6510
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   45
      Width           =   2745
      Begin VB.CommandButton BotaoConsultaTitPag 
         Height          =   390
         Left            =   120
         Picture         =   "BaixaPagCancelarOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Consulta de título a pagar"
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   390
         Left            =   2220
         Picture         =   "BaixaPagCancelarOcx.ctx":0F0A
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   390
         Left            =   1715
         Picture         =   "BaixaPagCancelarOcx.ctx":1088
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Limpar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   390
         Left            =   1210
         Picture         =   "BaixaPagCancelarOcx.ctx":15BA
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5280
      Left            =   90
      TabIndex        =   118
      Top             =   360
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9313
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
Attribute VB_Name = "BaixaPagCancelarOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBaixaPagCancelar
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoConsultaTitPag_Click()
    Call objCT.BotaoConsultaTitPag_Click
End Sub

Private Sub LabelParcela_Click()
    Call objCT.LabelParcela_Click
End Sub

Private Sub UserControl_Initialize()
    Set objCT = New CTBaixaPagCancelar
    Set objCT.objUserControl = Me
End Sub

Private Sub BotaoTrazer_Click()
     Call objCT.BotaoTrazer_Click
End Sub

Private Sub DataEmissao_GotFocus()
     Call objCT.DataEmissao_GotFocus
End Sub

Private Sub DataEmissao_Validate(Cancel As Boolean)
     Call objCT.DataEmissao_Validate(Cancel)
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub Fornecedor_Change()
     Call objCT.Fornecedor_Change
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)
     Call objCT.Fornecedor_Validate(Cancel)
End Sub

Private Sub FornecedorLabel_Click()
     Call objCT.FornecedorLabel_Click
End Sub

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
    Call objCT.Numero_Validate(Cancel)
End Sub

Private Sub NumeroTitulo_Click()
     Call objCT.NumeroTitulo_Click
End Sub

Private Sub BotaoFechar_Click()
     Call objCT.BotaoFechar_Click
End Sub

Private Sub BotaoGravar_Click()
     Call objCT.BotaoGravar_Click
End Sub

Private Sub BotaoLimpar_Click()
     Call objCT.BotaoLimpar_Click
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Pagamento_Click(Index As Integer)
     Call objCT.Pagamento_Click(Index)
End Sub

Private Sub Parcela_GotFocus()
     Call objCT.Parcela_GotFocus
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
    Call objCT.Parcela_Validate(Cancel)
End Sub

Private Sub Sequencial_GotFocus()
     Call objCT.Sequencial_GotFocus
End Sub

Private Sub Sequencial_Validate(Cancel As Boolean)
    Call objCT.Sequencial_Validate(Cancel)
End Sub

Private Sub SequencialLabel_Click()
     Call objCT.SequencialLabel_Click
End Sub

Private Sub SiglaDocumento_Validate(Cancel As Boolean)
     Call objCT.SiglaDocumento_Validate(Cancel)
End Sub

Private Sub SiglaDocumentoLabel_Click()
     Call objCT.SiglaDocumentoLabel_Click
End Sub

Private Sub UpDownEmissao_DownClick()
     Call objCT.UpDownEmissao_DownClick
End Sub

Private Sub UpDownEmissao_UpClick()
     Call objCT.UpDownEmissao_UpClick
End Sub

Function Trata_Parametros() As Long
     Trata_Parametros = objCT.Trata_Parametros()
End Function

Private Sub CTBBotaoModeloPadrao_Click()
     Call objCT.CTBBotaoModeloPadrao_Click
End Sub

Private Sub CTBModelo_Click()
     Call objCT.CTBModelo_Click
End Sub

Private Sub CTBGridContabil_Click()
     Call objCT.CTBGridContabil_Click
End Sub

Private Sub CTBGridContabil_EnterCell()
     Call objCT.CTBGridContabil_EnterCell
End Sub

Private Sub CTBGridContabil_GotFocus()
     Call objCT.CTBGridContabil_GotFocus
End Sub

Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
     Call objCT.CTBGridContabil_KeyPress(KeyAscii)
End Sub

Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.CTBGridContabil_KeyDown(KeyCode, Shift)
End Sub

Private Sub CTBGridContabil_LeaveCell()
     Call objCT.CTBGridContabil_LeaveCell
End Sub

Private Sub CTBGridContabil_Validate(Cancel As Boolean)
     Call objCT.CTBGridContabil_Validate(Cancel)
End Sub

Private Sub CTBGridContabil_RowColChange()
     Call objCT.CTBGridContabil_RowColChange
End Sub

Private Sub CTBGridContabil_Scroll()
     Call objCT.CTBGridContabil_Scroll
End Sub

Private Sub CTBConta_Change()
     Call objCT.CTBConta_Change
End Sub

Private Sub CTBConta_GotFocus()
     Call objCT.CTBConta_GotFocus
End Sub

Private Sub CTBConta_KeyPress(KeyAscii As Integer)
     Call objCT.CTBConta_KeyPress(KeyAscii)
End Sub

Private Sub CTBConta_Validate(Cancel As Boolean)
     Call objCT.CTBConta_Validate(Cancel)
End Sub

Private Sub CTBCcl_Change()
     Call objCT.CTBCcl_Change
End Sub

Private Sub CTBCcl_GotFocus()
     Call objCT.CTBCcl_GotFocus
End Sub

Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCcl_KeyPress(KeyAscii)
End Sub

Private Sub CTBCcl_Validate(Cancel As Boolean)
     Call objCT.CTBCcl_Validate(Cancel)
End Sub

Private Sub CTBCredito_Change()
     Call objCT.CTBCredito_Change
End Sub

Private Sub CTBCredito_GotFocus()
     Call objCT.CTBCredito_GotFocus
End Sub

Private Sub CTBLancAutomatico_Click()
    Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
    Call objCT.CTBAglutina_Click
End Sub

Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBCredito_KeyPress(KeyAscii)
End Sub

Private Sub CTBCredito_Validate(Cancel As Boolean)
     Call objCT.CTBCredito_Validate(Cancel)
End Sub

Private Sub CTBDebito_Change()
     Call objCT.CTBDebito_Change
End Sub

Private Sub CTBDebito_GotFocus()
     Call objCT.CTBDebito_GotFocus
End Sub

Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
     Call objCT.CTBDebito_KeyPress(KeyAscii)
End Sub

Private Sub CTBDebito_Validate(Cancel As Boolean)
     Call objCT.CTBDebito_Validate(Cancel)
End Sub

Private Sub CTBSeqContraPartida_Change()
     Call objCT.CTBSeqContraPartida_Change
End Sub

Private Sub CTBSeqContraPartida_GotFocus()
     Call objCT.CTBSeqContraPartida_GotFocus
End Sub

Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
     Call objCT.CTBSeqContraPartida_KeyPress(KeyAscii)
End Sub

Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
     Call objCT.CTBSeqContraPartida_Validate(Cancel)
End Sub

Private Sub CTBHistorico_Change()
     Call objCT.CTBHistorico_Change
End Sub

Private Sub CTBHistorico_GotFocus()
     Call objCT.CTBHistorico_GotFocus
End Sub

Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
     Call objCT.CTBHistorico_KeyPress(KeyAscii)
End Sub

Private Sub CTBHistorico_Validate(Cancel As Boolean)
     Call objCT.CTBHistorico_Validate(Cancel)
End Sub

Private Sub CTBAglutina_GotFocus()
     Call objCT.CTBAglutina_GotFocus
End Sub

Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
     Call objCT.CTBAglutina_KeyPress(KeyAscii)
End Sub

Private Sub CTBAglutina_Validate(Cancel As Boolean)
     Call objCT.CTBAglutina_Validate(Cancel)
End Sub

Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_NodeClick(Node)
End Sub

Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwContas_Expand(Node)
End Sub

Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.CTBTvwCcls_NodeClick(Node)
End Sub

Private Sub CTBListHistoricos_DblClick()
     Call objCT.CTBListHistoricos_DblClick
End Sub

Private Sub CTBBotaoLimparGrid_Click()
     Call objCT.CTBBotaoLimparGrid_Click
End Sub

Private Sub CTBLote_Change()
     Call objCT.CTBLote_Change
End Sub

Private Sub CTBLote_GotFocus()
     Call objCT.CTBLote_GotFocus
End Sub

Private Sub CTBLote_Validate(Cancel As Boolean)
     Call objCT.CTBLote_Validate(Cancel)
End Sub

Private Sub CTBDataContabil_Change()
     Call objCT.CTBDataContabil_Change
End Sub

Private Sub CTBDataContabil_GotFocus()
     Call objCT.CTBDataContabil_GotFocus
End Sub

Private Sub CTBDocumento_Change()
     Call objCT.CTBDocumento_Change
End Sub

Private Sub CTBDocumento_GotFocus()
     Call objCT.CTBDocumento_GotFocus
End Sub

Private Sub CTBBotaoImprimir_Click()
     Call objCT.CTBBotaoImprimir_Click
End Sub

Private Sub CTBUpDown_DownClick()
     Call objCT.CTBUpDown_DownClick
End Sub

Private Sub CTBUpDown_UpClick()
     Call objCT.CTBUpDown_UpClick
End Sub

Private Sub CTBLabelDoc_Click()
     Call objCT.CTBLabelDoc_Click
End Sub

Private Sub CTBLabelLote_Click()
     Call objCT.CTBLabelLote_Click
End Sub

Public Function Form_Load_Ocx() As Object

    Call objCT.Form_Load_Ocx
    Set Form_Load_Ocx = Me

End Function

Public Sub Form_Unload(Cancel As Integer)
        
    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
            Set objCT.objUserControl = Nothing
            Set objCT = Nothing
        End If
    End If
        
End Sub

Private Sub objCT_Unload()
   RaiseEvent Unload
End Sub

Public Function Name() As String
    Name = objCT.Name
End Function

Public Sub Show()
    Call objCT.Show
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = objCT.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    objCT.Caption = New_Caption
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.UserControl_KeyDown(KeyCode, Shift)
End Sub



Private Sub Portador_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Portador, Source, X, Y)
End Sub

Private Sub Portador_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Portador, Button, Shift, X, Y)
End Sub

Private Sub NumOuSequencial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumOuSequencial, Source, X, Y)
End Sub

Private Sub NumOuSequencial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumOuSequencial, Button, Shift, X, Y)
End Sub

Private Sub Historico_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Historico, Source, X, Y)
End Sub

Private Sub Historico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Historico, Button, Shift, X, Y)
End Sub

Private Sub ValorPagoPagto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPagoPagto, Source, X, Y)
End Sub

Private Sub ValorPagoPagto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPagoPagto, Button, Shift, X, Y)
End Sub

Private Sub ContaCorrente_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ContaCorrente, Source, X, Y)
End Sub

Private Sub ContaCorrente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ContaCorrente, Button, Shift, X, Y)
End Sub

Private Sub Label32_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label32, Source, X, Y)
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label32, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
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

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label23, Source, X, Y)
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label23, Button, Shift, X, Y)
End Sub

Private Sub Label33_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label33, Source, X, Y)
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label33, Button, Shift, X, Y)
End Sub

Private Sub DataMovimento_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataMovimento, Source, X, Y)
End Sub

Private Sub DataMovimento_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataMovimento, Button, Shift, X, Y)
End Sub

Private Sub MeioPagtoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MeioPagtoDescricao, Source, X, Y)
End Sub

Private Sub MeioPagtoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MeioPagtoDescricao, Button, Shift, X, Y)
End Sub

Private Sub ValorPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPA, Source, X, Y)
End Sub

Private Sub ValorPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPA, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaPA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaPA, Source, X, Y)
End Sub

Private Sub FilialEmpresaPA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaPA, Button, Shift, X, Y)
End Sub

Private Sub CCIntNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CCIntNomeReduzido, Source, X, Y)
End Sub

Private Sub CCIntNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CCIntNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub NumeroMP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroMP, Source, X, Y)
End Sub

Private Sub NumeroMP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroMP, Button, Shift, X, Y)
End Sub

Private Sub Label31_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label31, Source, X, Y)
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label31, Button, Shift, X, Y)
End Sub

Private Sub ValorBaixado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorBaixado, Source, X, Y)
End Sub

Private Sub ValorBaixado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorBaixado, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Desconto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Desconto, Source, X, Y)
End Sub

Private Sub Desconto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Desconto, Button, Shift, X, Y)
End Sub

Private Sub Multa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Multa, Source, X, Y)
End Sub

Private Sub Multa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Multa, Button, Shift, X, Y)
End Sub

Private Sub Juros_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Juros, Source, X, Y)
End Sub

Private Sub Juros_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Juros, Button, Shift, X, Y)
End Sub

Private Sub ValorPago_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorPago, Source, X, Y)
End Sub

Private Sub ValorPago_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorPago, Button, Shift, X, Y)
End Sub

Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub

Private Sub DataBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataBaixa, Source, X, Y)
End Sub

Private Sub DataBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label30_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label30, Source, X, Y)
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label30, Button, Shift, X, Y)
End Sub

Private Sub SiglaDocumentoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SiglaDocumentoLabel, Source, X, Y)
End Sub

Private Sub SiglaDocumentoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SiglaDocumentoLabel, Button, Shift, X, Y)
End Sub

Private Sub SequencialLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SequencialLabel, Source, X, Y)
End Sub

Private Sub SequencialLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SequencialLabel, Button, Shift, X, Y)
End Sub

Private Sub LabelParcela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelParcela, Source, X, Y)
End Sub

Private Sub LabelParcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelParcela, Button, Shift, X, Y)
End Sub

Private Sub NumeroTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumeroTitulo, Source, X, Y)
End Sub

Private Sub NumeroTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumeroTitulo, Button, Shift, X, Y)
End Sub

Private Sub FornecedorLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FornecedorLabel, Source, X, Y)
End Sub

Private Sub FornecedorLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FornecedorLabel, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

Private Sub SaldoCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoCredito, Source, X, Y)
End Sub

Private Sub SaldoCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoCredito, Button, Shift, X, Y)
End Sub

Private Sub NumTitulo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(NumTitulo, Source, X, Y)
End Sub

Private Sub NumTitulo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(NumTitulo, Button, Shift, X, Y)
End Sub

Private Sub FilialEmpresaCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FilialEmpresaCR, Source, X, Y)
End Sub

Private Sub FilialEmpresaCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FilialEmpresaCR, Button, Shift, X, Y)
End Sub

Private Sub ValorCredito_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorCredito, Source, X, Y)
End Sub

Private Sub ValorCredito_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorCredito, Button, Shift, X, Y)
End Sub

Private Sub SiglaDocumentoCR_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SiglaDocumentoCR, Source, X, Y)
End Sub

Private Sub SiglaDocumentoCR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SiglaDocumentoCR, Button, Shift, X, Y)
End Sub

Private Sub DataEmissaoCred_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataEmissaoCred, Source, X, Y)
End Sub

Private Sub DataEmissaoCred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataEmissaoCred, Button, Shift, X, Y)
End Sub

Private Sub Label40_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label40, Source, X, Y)
End Sub

Private Sub Label40_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label40, Button, Shift, X, Y)
End Sub

Private Sub Label39_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label39, Source, X, Y)
End Sub

Private Sub Label39_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label39, Button, Shift, X, Y)
End Sub

Private Sub Label38_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label38, Source, X, Y)
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label38, Button, Shift, X, Y)
End Sub

Private Sub Label37_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label37, Source, X, Y)
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label37, Button, Shift, X, Y)
End Sub

Private Sub Label48_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label48, Source, X, Y)
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label48, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
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


Private Sub CTBGerencial_Click()
    Call objCT.CTBGerencial_Click
End Sub

Private Sub CTBGerencial_GotFocus()
    Call objCT.CTBGerencial_GotFocus
End Sub

Private Sub CTBGerencial_KeyPress(KeyAscii As Integer)
    Call objCT.CTBGerencial_KeyPress(KeyAscii)
End Sub

Private Sub CTBGerencial_Validate(Cancel As Boolean)
    Call objCT.CTBGerencial_Validate(Cancel)
End Sub

