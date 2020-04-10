VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl BaixaAntecipCredFornecedorOcx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9390
   ScaleHeight     =   5790
   ScaleWidth      =   9390
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   3
      Left            =   150
      TabIndex        =   64
      Top             =   792
      Visible         =   0   'False
      Width           =   9105
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4920
         TabIndex        =   107
         Tag             =   "1"
         Top             =   1440
         Width           =   870
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   74
         Top             =   3330
         Width           =   5895
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   78
            Top             =   645
            Visible         =   0   'False
            Width           =   3720
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   77
            Top             =   285
            Width           =   3720
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
            TabIndex        =   76
            Top             =   300
            Width           =   570
         End
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
            TabIndex        =   75
            Top             =   660
            Visible         =   0   'False
            Width           =   1440
         End
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   73
         Top             =   2190
         Width           =   1770
      End
      Begin VB.CheckBox CTBAglutina 
         Height          =   210
         Left            =   4470
         TabIndex        =   72
         Top             =   2565
         Width           =   870
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
         Left            =   3450
         TabIndex        =   71
         Top             =   945
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2850
         IntegralHeight  =   0   'False
         Left            =   6330
         TabIndex        =   70
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
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
         TabIndex        =   68
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
         Left            =   6270
         TabIndex        =   67
         Top             =   60
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6330
         Style           =   2  'Dropdown List
         TabIndex        =   66
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
         Left            =   7725
         TabIndex        =   65
         Top             =   60
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4680
         TabIndex        =   69
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
         TabIndex        =   79
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
         TabIndex        =   80
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
         TabIndex        =   81
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
         TabIndex        =   82
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
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   585
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   84
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
         TabIndex        =   85
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
         TabIndex        =   86
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
         Height          =   1770
         Left            =   15
         TabIndex        =   87
         Top             =   1230
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3122
         _Version        =   393216
         Rows            =   7
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin MSComctlLib.TreeView CTBTvwContas 
         Height          =   2850
         Left            =   6330
         TabIndex        =   88
         Top             =   1530
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView CTBTvwCcls 
         Height          =   2850
         Left            =   6330
         TabIndex        =   89
         Top             =   1530
         Visible         =   0   'False
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   5027
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
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
         TabIndex        =   106
         Top             =   180
         Width           =   450
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
         TabIndex        =   105
         Top             =   180
         Width           =   1035
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
         TabIndex        =   104
         Top             =   600
         Width           =   480
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   103
         Top             =   3015
         Width           =   1155
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3885
         TabIndex        =   102
         Top             =   3015
         Width           =   1155
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
         Height          =   345
         Left            =   1980
         TabIndex        =   101
         Top             =   3030
         Width           =   615
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
         TabIndex        =   100
         Top             =   1275
         Visible         =   0   'False
         Width           =   2490
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
         TabIndex        =   99
         Top             =   1275
         Width           =   2340
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
         TabIndex        =   98
         Top             =   1275
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   97
         Top             =   990
         Width           =   1140
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
         TabIndex        =   96
         Top             =   630
         Width           =   870
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   95
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   94
         Top             =   615
         Width           =   1185
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
         TabIndex        =   93
         Top             =   645
         Width           =   735
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   92
         Top             =   135
         Width           =   1530
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
         Left            =   30
         TabIndex        =   91
         Top             =   180
         Width           =   720
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
         Left            =   6330
         TabIndex        =   90
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   1
      Left            =   150
      TabIndex        =   27
      Top             =   792
      Width           =   9105
      Begin VB.Frame FrameFiltrosAdiantamento 
         Caption         =   "Conta Corrente"
         Height          =   1575
         Left            =   2880
         TabIndex        =   44
         Top             =   2400
         Width           =   5535
         Begin VB.ComboBox ContaCorrenteSeleciona 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   43
            Top             =   960
            Width           =   2448
         End
         Begin VB.OptionButton CtaCorrenteTodas 
            Caption         =   "Todas"
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
            TabIndex        =   46
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton CtaCorrenteApenas 
            Caption         =   "Apenas:"
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
            TabIndex        =   45
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame FrameFiltrosCredito 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   2772
         TabIndex        =   47
         Top             =   2400
         Visible         =   0   'False
         Width           =   5895
         Begin VB.Frame FrameTipoCredito 
            Caption         =   "Tipo de Crédito"
            Height          =   735
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   5535
            Begin VB.ComboBox TipoCreditoSeleciona 
               Height          =   315
               Left            =   3360
               TabIndex        =   56
               Top             =   280
               Width           =   1815
            End
            Begin VB.OptionButton TipoCreditoApenas 
               Caption         =   "Apenas:"
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
               Left            =   2280
               TabIndex        =   55
               Top             =   310
               Width           =   1095
            End
            Begin VB.OptionButton TipoCreditoTodos 
               Caption         =   "Todos"
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
               Left            =   960
               TabIndex        =   54
               Top             =   310
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Frame FrameNumCredito 
            Caption         =   "Nº Crédito / Devolução"
            Height          =   735
            Left            =   120
            TabIndex        =   48
            Top             =   0
            Width           =   5535
            Begin MSMask.MaskEdBox NumCreditoInic 
               Height          =   300
               Left            =   1320
               TabIndex        =   49
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox NumCreditoFim 
               Height          =   300
               Left            =   3615
               TabIndex        =   50
               Top             =   285
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelNumCreditoAte 
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
               Height          =   255
               Left            =   3195
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   52
               Top             =   315
               Width           =   375
            End
            Begin VB.Label LabelNumCreditoDe 
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
               Height          =   255
               Left            =   960
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   51
               Top             =   315
               Width           =   375
            End
         End
      End
      Begin VB.Frame FrameDataEmissao 
         Caption         =   "Data da Emissão"
         Height          =   1575
         Left            =   360
         TabIndex        =   36
         Top             =   2400
         Width           =   2175
         Begin MSComCtl2.UpDown UpDownBaixaInic 
            Height          =   300
            Left            =   1710
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   435
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoInic 
            Height          =   300
            Left            =   630
            TabIndex        =   38
            Top             =   450
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownBaixaFim 
            Height          =   300
            Left            =   1725
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox EmissaoFim 
            Height          =   300
            Left            =   636
            TabIndex        =   40
            Top             =   936
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataEmissaoDe 
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
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   375
         End
         Begin VB.Label LabelDataEmissaoAte 
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
            Height          =   255
            Left            =   195
            TabIndex        =   41
            Top             =   990
            Width           =   375
         End
      End
      Begin VB.Frame FrameFornecedor 
         Caption         =   "Fornecedor"
         Height          =   960
         Left            =   360
         TabIndex        =   31
         Top             =   1200
         Width           =   8055
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5712
            TabIndex        =   32
            Top             =   396
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Fornecedor 
            Height          =   300
            Left            =   1536
            TabIndex        =   33
            Top             =   396
            Width           =   3348
            _ExtentX        =   5927
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
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
            ForeColor       =   &H00000080&
            Height          =   192
            Left            =   432
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   35
            Top             =   432
            Width           =   1032
         End
         Begin VB.Label LabelFilial 
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
            Height          =   252
            Left            =   5088
            TabIndex        =   34
            Top             =   432
            Width           =   612
         End
      End
      Begin VB.Frame FrameTipoBaixas 
         Caption         =   "Selecione os documento que serão baixados"
         Height          =   720
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   8055
         Begin VB.OptionButton TipoBaixaAdiantamento 
            Caption         =   "Adiantamentos"
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
            Left            =   1320
            TabIndex        =   30
            Top             =   276
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton TipoBaixaCreditos 
            Caption         =   "Créditos / Devoluções"
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
            Left            =   4440
            TabIndex        =   29
            Top             =   276
            Width           =   2415
         End
      End
   End
   Begin VB.Frame FramePrincipal 
      BorderStyle     =   0  'None
      Height          =   4860
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   792
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame FrameCreditos 
         Caption         =   " Créditos / Devoluções "
         Height          =   2985
         Left            =   360
         TabIndex        =   25
         Top             =   60
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CheckBox CreditoSeleciona 
            Height          =   255
            Left            =   840
            TabIndex        =   63
            Top             =   1080
            Width           =   615
         End
         Begin MSMask.MaskEdBox NumCredito 
            Height          =   225
            Left            =   4200
            TabIndex        =   57
            Top             =   1560
            Width           =   705
            _ExtentX        =   1244
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
            Mask            =   "999999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataCredito 
            Height          =   240
            Left            =   2160
            TabIndex        =   58
            Top             =   1560
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox SaldoCredito 
            Height          =   225
            Left            =   6240
            TabIndex        =   59
            Top             =   1560
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox ValorCredito 
            Height          =   225
            Left            =   5040
            TabIndex        =   60
            Top             =   1560
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox TipoCredito 
            Height          =   225
            Left            =   3480
            TabIndex        =   61
            Top             =   1560
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialCredito 
            Height          =   225
            Left            =   840
            TabIndex        =   62
            Top             =   1560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridCreditos 
            Height          =   1935
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
            _Version        =   393216
         End
      End
      Begin VB.Frame FrameBaixa 
         Caption         =   "Baixa"
         Height          =   1575
         Left            =   330
         TabIndex        =   6
         Top             =   3135
         Width           =   8295
         Begin VB.TextBox Historico 
            Height          =   320
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1175
            Width           =   5655
         End
         Begin VB.ComboBox ContaCorrente 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Top             =   750
            Width           =   1815
         End
         Begin MSComCtl2.UpDown UpDownDataBaixa 
            Height          =   300
            Left            =   3015
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   300
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataBaixa 
            Height          =   300
            Left            =   1920
            TabIndex        =   10
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelDataBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Data da Baixa:"
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
            TabIndex        =   15
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label LabelHistorico 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   14
            Top             =   1238
            Width           =   825
         End
         Begin VB.Label LabelCtaCorrente 
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
            Left            =   480
            TabIndex        =   13
            Top             =   810
            Width           =   1350
         End
         Begin VB.Label LabelValorBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Valor da Baixa:"
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
            Left            =   4920
            TabIndex        =   12
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label LabelValorBaixa2 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   6360
            TabIndex        =   11
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame FrameAdiantamentos 
         Caption         =   " Adiantamentos "
         Height          =   2985
         Left            =   360
         TabIndex        =   16
         Top             =   60
         Width           =   8295
         Begin VB.CheckBox AdiantamentoSeleciona 
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   960
            Width           =   615
         End
         Begin MSMask.MaskEdBox FilialAdiantamento 
            Height          =   225
            Left            =   720
            TabIndex        =   17
            Top             =   1440
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox SaldoAdiantamento 
            Height          =   225
            Left            =   6720
            TabIndex        =   18
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox ValorAdiantamento 
            Height          =   225
            Left            =   5640
            TabIndex        =   19
            Top             =   1440
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox MeioPagtoAdiantamento 
            Height          =   225
            Left            =   4200
            TabIndex        =   20
            Top             =   1440
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox CtaCorrenteAdiantamento 
            Height          =   225
            Left            =   2880
            TabIndex        =   21
            Top             =   1440
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataAdiantamento 
            Height          =   225
            Left            =   1560
            TabIndex        =   22
            Top             =   1440
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid GridAdiantamentos 
            Height          =   1935
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   8
            Cols            =   7
         End
      End
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7440
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaAntecipCredFornecedor.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   607
         Picture         =   "BaixaAntecipCredFornecedor.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "BaixaAntecipCredFornecedor.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5250
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   9260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtros"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adiantamentos / Créditos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "BaixaAntecipCredFornecedorOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'mnemonicos
Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const CONTA_COD As String = "Conta_Codigo"
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta"
Private Const DATA1 As String = "DataBaixa"
Private Const VALOR1 As String = "ValorBaixa"
Private Const HISTORICO1 As String = "Historico"

'Retirar quando passar para o Rotinas
Private Declare Function Comando_BindVarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_BindVar" (ByVal lComando As Long, lpVar As Variant) As Long
Private Declare Function Comando_PrepararInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Preparar" (ByVal lComando As Long, ByVal lpSQLStmt As String) As Long
Private Declare Function Comando_ExecutarInt Lib "ADSQLMN.DLL" Alias "AD_Comando_Executar" (ByVal lComando As Long) As Long

'Eventos para chamada de browser
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCreditoDe As AdmEvento
Attribute objEventoCreditoDe.VB_VarHelpID = -1
Private WithEvents objEventoCreditoAte As AdmEvento
Attribute objEventoCreditoAte.VB_VarHelpID = -1

'Obj Global com os Dados do Adiantamento ou do Credito
Dim gobjBaixaAntecipCredFornecedor As ClassBaixaAntecipCredForn

'Definições do Grid de Adiantementos
Dim objGridAdiantamentos As AdmGrid

Dim iGrid_AdiantamentoSeleciona_Col As Integer
Dim iGrid_AdiantamentoFilial_Col As Integer
Dim iGrid_AdiantamentoData_Col As Integer
Dim iGrid_AdiantamentoCtaCorrente_Col As Integer
Dim iGrid_AdiantamentoMeioPagto_Col As Integer
Dim iGrid_AdiantamentoValor_Col As Integer
Dim iGrid_AdiantamentoSaldo_Col As Integer

'Definições do Grid de Créditos
Dim objGridCreditos As AdmGrid

Dim iGrid_CreditoSeleciona_Col As Integer
Dim iGrid_CreditoFilial_Col As Integer
Dim iGrid_CreditoData_Col As Integer
Dim iGrid_CreditoTipo_Col As Integer
Dim iGrid_CreditoNumero_Col As Integer
Dim iGrid_CreditoValor_Col As Integer
Dim iGrid_CreditoSaldo_Col As Integer

'Variáveis de Controle das Alterações feitas na Tela
Public iAlterado As Integer
Dim iFramePrincipalAtual As Integer
Dim iFramePrincipalAlterado As Integer
Dim iUltTipoBaixa As Integer
Dim iFornecedorAlterado As Integer

'Definições dos TAB´s
Const TAB_FILTROS = 1
Const TAB_ADIANTAMENTOSCREDITOS = 2
Const TAB_Contabilizacao = 3

'Definições dos tipos de baixa
Const TIPOBAIXA_ADIANTAMENTO = 1
Const TIPOBAIXA_CREDITOS = 2

'Definições das Constantes da Tela
Const NUM_MAX_ADIANTAMENTOS = 100
Const NUM_MAX_CREDITOS = 100
Const CONTACORRENTE_TODAS = 1
Const CONTACORRENTE_APENAS = 2
Const TIPOCREDITO_TODOS = 1
Const TIPOCREDITO_APENAS = 2
Const TELA_TODA = 1

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    '??? criar IDH Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Adiantamentos / Créditos Fornecedores"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaAntecipCredFornecedor"
    
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

Private Sub AdiantamentoSeleciona_Click()

    'Se o usuário selecionou a parcela
    If GridAdiantamentos.TextMatrix(GridAdiantamentos.Row, iGrid_AdiantamentoSeleciona_Col) = True Then
        
        'Guarda no obj a informação de que a parcela está selecionada
        gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor(GridAdiantamentos.Row).iSelecionado = MARCADO
    
    Else
        
        'Guarda no obj global a informação de que a parcela não está selecionada
        gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor(GridAdiantamentos.Row).iSelecionado = DESMARCADO
    
    End If
    
    'Calcula o valor total de adiantamentos a serem baixados
    Call Calcula_Total_Baixar(TIPOBAIXA_ADIANTAMENTO)
    
    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridAdiantamentos)

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Calcula_Total_Baixar(iTipoBaixa As Integer)

Dim iIndice As Integer
Dim dTotalSelecionado As Double
Dim objGrid As AdmGrid
Dim ControleGrid As Object
Dim iGrid_Seleciona_Col As Integer
Dim iGrid_Saldo_Col As Integer

    dTotalSelecionado = 0
    
    'Verifica o tipo de baixa passado como parâmetro
    Select Case iTipoBaixa
    
        'Se for baixa de adiantamentos => o grid utilizado para o cálculo deve ser o de adiantamentos
        Case TIPOBAIXA_ADIANTAMENTO
        
            'Instancia o obj e controle com os seus correspondentes para o grid Adiantamentos
            Set objGrid = objGridAdiantamentos
            Set ControleGrid = GridAdiantamentos
            iGrid_Seleciona_Col = iGrid_AdiantamentoSeleciona_Col
            iGrid_Saldo_Col = iGrid_AdiantamentoSaldo_Col
        
        'Se for baixa de créditos => o grid utilizado para o cálculo deve ser o de créditos
        Case TIPOBAIXA_CREDITOS
            
            'Instancia o obj e controle com os seus correspondentes para o grid Débitos
            Set objGrid = objGridCreditos
            Set ControleGrid = GridCreditos
            iGrid_Seleciona_Col = iGrid_CreditoSeleciona_Col
            iGrid_Saldo_Col = iGrid_CreditoSaldo_Col
            
    End Select
    
    'Para cada linha do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        'Se a linha está selecionada
        If StrParaInt(ControleGrid.TextMatrix(iIndice, iGrid_Seleciona_Col)) <> 0 Then
            
            'Adiciona o valor da linha ao valor total das linhas selecionadas
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(ControleGrid.TextMatrix(iIndice, iGrid_Saldo_Col))
            
        End If
        
    Next
    
    'Exibe na tela o valor encontrado acima
    LabelValorBaixa2.Caption = Format(dTotalSelecionado, "Standard")
    
    'Guarda no obj global o valor encontrado
    gobjBaixaAntecipCredFornecedor.dValorBaixa = dTotalSelecionado
    
End Sub

Private Sub BotaoFechar_Click()
    'Fecha a tela
    Unload Me
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava a Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 92715

    Call Limpa_Tela_BaixaAntecipCredFornecedor
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 92715

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143185)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objBaixaAntecipCredFornecedor As New ClassBaixaAntecipCredForn

On Error GoTo Erro_Gravar_Registro

    lErro = Valida_Gravacao()
    If lErro <> SUCESSO Then gError 92716
    
    lErro = Move_Tela_Memoria(objBaixaAntecipCredFornecedor)
    If lErro <> SUCESSO Then gError 92782
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataBaixa.Text))
    If lErro <> SUCESSO Then gError 92791
    
    Select Case objBaixaAntecipCredFornecedor.iTipoBaixa
    
        Case TIPOBAIXA_ADIANTAMENTO
            
            lErro = CF("BaixaAntecipCredForn_Grava_Adiantamentos", objBaixaAntecipCredFornecedor, objContabil)
            If lErro <> SUCESSO Then gError 92717
        
        Case TIPOBAIXA_CREDITOS
    
            lErro = CF("BaixaAntecipCredForn_Grava_Creditos", objBaixaAntecipCredFornecedor, objContabil)
            If lErro <> SUCESSO Then gError 92718
    
    End Select
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
    
    Gravar_Registro = gErr
    
    Select Case gErr
    
        Case 92716 To 92718, 92782, 92791
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143186)
            
    End Select
    
    Exit Function
    
End Function

Private Function Valida_Gravacao() As Long

Dim iLinha As Integer

On Error GoTo Erro_Valida_Gravacao

    'Se o fornecedor não está preenchido => erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 92719
    
    'Se a filial do fornecedor não está preenchida => erro
    If Len(Trim(Filial.Text)) = 0 Then gError 92720
    
    'Se o valor da baixa não foi informado = > erro
    If StrParaDbl(LabelValorBaixa2.Caption) = 0 Then gError 92721
    
    'Se a conta corrente não foi informada => erro
    If Len(Trim(ContaCorrente.Text)) = 0 Then gError 92722
    
    'se for adiantamento
    If TipoBaixaAdiantamento.Value = True Then
    
        For iLinha = 1 To objGridAdiantamentos.iLinhasExistentes
        
            If GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoSeleciona_Col) = MARCADO Then
            
                If CDate(GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoData_Col)) > CDate(DataBaixa.Text) Then gError 92723
                
            End If
            
        Next
            
    Else
    
        For iLinha = 1 To objGridCreditos.iLinhasExistentes
        
            If GridCreditos.TextMatrix(iLinha, iGrid_CreditoSeleciona_Col) = MARCADO Then
            
                If Len(GridCreditos.TextMatrix(iLinha, iGrid_CreditoData_Col)) > 0 Then
            
                    If CDate(GridCreditos.TextMatrix(iLinha, iGrid_CreditoData_Col)) > CDate(DataBaixa.Text) Then gError 92779
                
                End If
                
            End If
            
        Next
    
    End If
    
    Valida_Gravacao = SUCESSO
    
    Exit Function

Erro_Valida_Gravacao:

    Valida_Gravacao = gErr
    
    Select Case gErr
    
        Case 92719
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 92720
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_PREENCHIDA", gErr)
            
        Case 92721
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_ANTECIPCREFOR_NAO_INFORMADO", gErr)

        Case 92722
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_ANTECIPCREFOR_NAO_INFORMADA", gErr)
        
        Case 92723
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADOCUMENTO_MAIOR_DATABAIXA", gErr, GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoData_Col), iLinha, DataBaixa.Text)
            
        Case 92779
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADOCUMENTO_MAIOR_DATABAIXA", gErr, GridCreditos.TextMatrix(iLinha, iGrid_CreditoData_Col), iLinha, DataBaixa.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143187)
            
    End Select
    
    Exit Function

End Function

Private Function Move_Tela_Memoria(objBaixaAntecipCredFornecedor As ClassBaixaAntecipCredForn) As Long

Dim iLinha As Integer
Dim objFornecedor As New ClassFornecedor
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Se o Fornecedor estiver preenchido lê o código no BD
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        'Lê o fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 92780
        
        If lErro = 6681 Then gError 92781 'Se nao encontrar --> erro

        objBaixaAntecipCredFornecedor.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) > 0 Then objBaixaAntecipCredFornecedor.iFilialFornecedor = Codigo_Extrai(Filial.Text)

    If Len(Trim(LabelValorBaixa2.Caption)) > 0 Then objBaixaAntecipCredFornecedor.dValorBaixa = CDbl(LabelValorBaixa2.Caption)
    
    If Len(Trim(DataBaixa.ClipText)) > 0 Then
        objBaixaAntecipCredFornecedor.dtDataBaixa = CDate(DataBaixa.Text)
    Else
        objBaixaAntecipCredFornecedor.dtDataBaixa = DATA_NULA
    End If

    objBaixaAntecipCredFornecedor.iCtaCorrenteBaixa = Codigo_Extrai(ContaCorrente.Text)
    
    objBaixaAntecipCredFornecedor.sHistoricoBaixa = Historico.Text

    Set objBaixaAntecipCredFornecedor.colAntecipCredFornecedor = New Collection

    'se for adiantamento
    If TipoBaixaAdiantamento.Value = True Then
    
        objBaixaAntecipCredFornecedor.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO
    
        For iLinha = 1 To objGridAdiantamentos.iLinhasExistentes
        
            If GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoSeleciona_Col) = MARCADO Then
            
                objBaixaAntecipCredFornecedor.colAntecipCredFornecedor.Add gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor.Item(iLinha)
                
            End If
            
        Next
            
    Else
    
        objBaixaAntecipCredFornecedor.iTipoBaixa = TIPOBAIXA_CREDITOS
    
        For iLinha = 1 To objGridCreditos.iLinhasExistentes
        
            If GridCreditos.TextMatrix(iLinha, iGrid_CreditoSeleciona_Col) = MARCADO Then
            
                objBaixaAntecipCredFornecedor.colAntecipCredFornecedor.Add gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor.Item(iLinha)
                
            End If
            
        Next
    
    End If
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 92780
    
        Case 92781
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, Fornecedor.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143188)
            
    End Select
    
    Exit Function

End Function

Private Sub ContaCorrente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)

Dim sContaCorrenteSeleciona As String
Dim iCodigo As Integer
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaCorrente_Validate

    'Indica que não existe conta corrente selecionada (só existirá uma conta selecionada depois da execução do validate)
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteBaixa = COMBO_INDICE

    'Se a ComboBox ContaCorrente não está preenchida => sai da função
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ContaCorrente
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then
        gobjBaixaAntecipCredFornecedor.iCtaCorrenteBaixa = ContaCorrente.ItemData(ContaCorrente.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 92774

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Lê os dados da conta corrente a partir do código passado como parâmetro
        objContaCorrente = iCodigo
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 92775

        'Se não encontrou a conta => o número de erro recebe tratamento para perguntar se deseja criar a conta
        If lErro = 11807 Then gError 92776

        'Coloca os dados da conta na combo
        ContaCorrente.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then
    
        'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
        objContaCorrente.sNomeReduzido = ContaCorrente.Text
        lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
        If lErro <> SUCESSO And lErro <> 15981 Then gError 92777

        'Se não encontrou a conta => Erro
        If lErro = 15981 Then gError 92778
        
        'Coloca os dados da conta na combo
        ContaCorrente.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
        
    End If
        
    'Guarda no obj global a tela o código da conta corrente selecionada
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteBaixa = objContaCorrente.iCodigo

    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True

    Select Case gErr

        Case 92774, 93775, 92777

        Case 92776
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteSeleciona.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            Else
                'Segura o foco
            End If

        Case 92778
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143189)

    End Select

    Exit Sub

End Sub

Private Sub ContaCorrenteSeleciona_Click()

    If ContaCorrenteSeleciona.ListIndex <> -1 Then
        iAlterado = REGISTRO_ALTERADO
        iFramePrincipalAlterado = REGISTRO_ALTERADO
        gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = ContaCorrenteSeleciona.ItemData(ContaCorrenteSeleciona.ListIndex)
    End If
End Sub

Private Sub CreditoSeleciona_Click()

    'Se o usuário selecionou a parcela
    If GridCreditos.TextMatrix(GridCreditos.Row, iGrid_CreditoSeleciona_Col) = True Then
        
        'Guarda no obj a informação de que a parcela está selecionada
        gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor(GridCreditos.Row).iSelecionado = MARCADO
    
    'Senão
    Else
        
        'Guarda no obj global a informação de que a parcela não está selecionada
        gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor(GridCreditos.Row).iSelecionado = DESMARCADO
    
    End If
    
    Call Calcula_Total_Baixar(TIPOBAIXA_CREDITOS)

    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridCreditos)
    
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataBaixa_Change()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    'Se a data DataBaixa está preenchida
    If Len(DataBaixa.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then gError 92773

    End If

    Exit Sub

Erro_DataBaixa_Validate:

    Cancel = True

    Select Case gErr

        Case 92773

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143190)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoFim_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub EmissaoInic_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Historico_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Historico_Validate(Cancel As Boolean)
    gobjBaixaAntecipCredFornecedor.sHistoricoBaixa = Historico.Text
End Sub

Private Sub NumCreditoFim_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumCreditoFim, iAlterado)
End Sub

Private Sub NumCreditoFim_Validate(Cancel As Boolean)
    gobjBaixaAntecipCredFornecedor.lNumCreditoAte = StrParaLong(NumCreditoFim.Text)
End Sub

Private Sub NumCreditoInic_GotFocus()
    Call MaskEdBox_TrataGotFocus(NumCreditoInic, iAlterado)
End Sub

Private Sub NumCreditoInic_Validate(Cancel As Boolean)
    gobjBaixaAntecipCredFornecedor.lNumCreditoDe = StrParaLong(NumCreditoInic.Text)
End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1
    
    Fornecedor.Text = CStr(objFornecedor.lCodigo)
    Call Fornecedor_Validate(Cancel)
    
    iAlterado = 0
    
    Me.Show

    Exit Sub

End Sub

Private Sub TipoCreditoSeleciona_Click()

Dim sTipoCredito As String
Dim lErro As Long

On Error GoTo Erro_TipoCreditoSeleciona_Click

    'Guarda na variável o conteúdo que será passado para a função
    'que obterá a sigla do Tipo
    sTipoCredito = TipoCreditoSeleciona.Text

    'Se foi selecionado um tipo
    If Len(Trim(sTipoCredito)) > 0 Then

        'Extrai a sigla do Tipo de credito selecionado
         sTipoCredito = SCodigo_Extrai(sTipoCredito)

        'Guarda no obj a sigla do tipo de credito selecionado
        gobjBaixaAntecipCredFornecedor.sTipoCredito = sTipoCredito

    'Senão
    Else

        'Indica que nenhum tipo de credito foi selecionado
        gobjBaixaAntecipCredFornecedor.sTipoCredito = " "

    End If

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_TipoCreditoSeleciona_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143191)

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
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    'm_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim lErro As Long

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_UnLoad(Cancel As Integer)

    Set objEventoFornecedor = Nothing
    Set objEventoCreditoDe = Nothing
    Set objEventoCreditoAte = Nothing
    Set objGridAdiantamentos = Nothing
    Set objGridCreditos = Nothing
    Set gobjBaixaAntecipCredFornecedor = Nothing
    
    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
End Sub

Public Sub Form_Load()
    
Dim lErro As Long
    
On Error GoTo Erro_Form_Load
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        TabStrip1.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    'Controle dos frames
    iFramePrincipalAtual = TAB_FILTROS
    iFramePrincipalAlterado = 0

    Call Inicializa_gobjBaixaAntecipCredFornecedor(TELA_TODA)

    FramePrincipal(1).Visible = True
    FramePrincipal(2).Visible = False
    FramePrincipal(3).Visible = False

    'Inicializa o Obj
    Set objGridAdiantamentos = New AdmGrid
    Set objGridCreditos = New AdmGrid
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCreditoDe = New AdmEvento
    Set objEventoCreditoAte = New AdmEvento
    
    'Preenche o campo data da baixa com a data atual
    DataBaixa.PromptInclude = False
    DataBaixa.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataBaixa.PromptInclude = True
    
    'Carrega ContaCorrenteSeleciona
    lErro = Carrega_Combos_ContaCorrente(ContaCorrenteSeleciona)
    If lErro <> SUCESSO Then gError 95384
    
    'Carrega ContaCorrente
    lErro = Carrega_Combos_ContaCorrente(ContaCorrente)
    If lErro <> SUCESSO Then gError 95391
    
    'Carrega TipoCredito
    lErro = TipoCreditoSeleciona_Carrega(TipoCreditoSeleciona)
    If lErro <> SUCESSO Then gError 95385
    
    'Inicializa o Grid de Adiantamentos
    lErro = Inicializa_GridAdiantamentos(objGridAdiantamentos)
    If lErro <> SUCESSO Then gError 95386
    
    'Inicializa o grid de Creditos
    lErro = Inicializa_GridCreditos(objGridCreditos)
    If lErro <> SUCESSO Then gError 95387
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 92787
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 92787, 95384 To 95387, 95391
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143192)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143193)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Carrega_Combos_ContaCorrente(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_Carrega_Combos_ContaCorrente

    'Lê Codigos e NomesReduzidos das Contas Correntes Internas
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError 95388

    'Preeche a Combo Passada
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next
    
    Carrega_Combos_ContaCorrente = SUCESSO

    Exit Function

Erro_Carrega_Combos_ContaCorrente:

    Carrega_Combos_ContaCorrente = gErr

    Select Case gErr

        Case 95388

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143194)

    End Select

    Exit Function

End Function

Private Function TipoCreditoSeleciona_Carrega(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colTipoCredito As New colTipoDocumento
Dim objTipoCredito As ClassTipoDocumento

On Error GoTo Erro_TipoCreditoSeleciona_Carrega

    'Lê os Tipos de Documentos de Creditos/Pagar
    lErro = CF("TiposDocumento_Le_CredPagar", colTipoCredito)
    If lErro <> SUCESSO Then gError 95389

    'Preenche a Combo passada
    For Each objTipoCredito In colTipoCredito
        objComboBox.AddItem objTipoCredito.sSigla & SEPARADOR & objTipoCredito.sDescricaoReduzida
    Next

    TipoCreditoSeleciona_Carrega = SUCESSO

    Exit Function

Erro_TipoCreditoSeleciona_Carrega:

    TipoCreditoSeleciona_Carrega = gErr

    Select Case gErr

        Case 95389

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143195)

    End Select

    Exit Function

End Function

Function Inicializa_GridAdiantamentos(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("C. Corrente")
    objGridInt.colColuna.Add ("Meio Pagto.")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
        
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AdiantamentoSeleciona.Name)
    objGridInt.colCampo.Add (FilialAdiantamento.Name)
    objGridInt.colCampo.Add (DataAdiantamento.Name)
    objGridInt.colCampo.Add (CtaCorrenteAdiantamento.Name)
    objGridInt.colCampo.Add (MeioPagtoAdiantamento.Name)
    objGridInt.colCampo.Add (ValorAdiantamento.Name)
    objGridInt.colCampo.Add (SaldoAdiantamento.Name)

    'Colunas do Grid
    iGrid_AdiantamentoSeleciona_Col = 1
    iGrid_AdiantamentoFilial_Col = 2
    iGrid_AdiantamentoData_Col = 3
    iGrid_AdiantamentoCtaCorrente_Col = 4
    iGrid_AdiantamentoMeioPagto_Col = 5
    iGrid_AdiantamentoValor_Col = 6
    iGrid_AdiantamentoSaldo_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridAdiantamentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ADIANTAMENTOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridAdiantamentos.ColWidth(0) = 0

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    'Proibido Excluir/Incluir linhas do Grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridAdiantamentos = SUCESSO

    Exit Function

End Function

Function Inicializa_GridCreditos(objGridInt As AdmGrid) As Long

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
        
    'Controles que participam do Grid
    objGridInt.colCampo.Add (CreditoSeleciona.Name)
    objGridInt.colCampo.Add (FilialCredito.Name)
    objGridInt.colCampo.Add (DataCredito.Name)
    objGridInt.colCampo.Add (TipoCredito.Name)
    objGridInt.colCampo.Add (NumCredito.Name)
    objGridInt.colCampo.Add (ValorCredito.Name)
    objGridInt.colCampo.Add (SaldoCredito.Name)

    'Colunas do Grid
    iGrid_CreditoSeleciona_Col = 1
    iGrid_CreditoFilial_Col = 2
    iGrid_CreditoData_Col = 3
    iGrid_CreditoTipo_Col = 4
    iGrid_CreditoNumero_Col = 5
    iGrid_CreditoValor_Col = 6
    iGrid_CreditoSaldo_Col = 7
    
    'Grid do GridInterno
    objGridInt.objGrid = GridCreditos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_CREDITOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridCreditos.ColWidth(0) = 0

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Proibido Excluir/Incluir linhas do Grid
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridCreditos = SUCESSO

    Exit Function

End Function

Private Sub TabStrip1_Click()

Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFramePrincipalAtual Then

        If TabStrip_PodeTrocarTab(iFramePrincipalAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub
        
        'Torna o frame clicado visível
        FramePrincipal(TabStrip1.SelectedItem.Index).Visible = True
    
        'Torna o frame atual invisível
        FramePrincipal(iFramePrincipalAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFramePrincipalAtual = TabStrip1.SelectedItem.Index
    
        'Se Frame selecionado foi o de Filtros
        If TabStrip1.SelectedItem.Index = TAB_FILTROS Then
        
            iFramePrincipalAlterado = 0
                    
        'Se Frame selecionado foi o de Adiantamento / Créditos
        ElseIf TabStrip1.SelectedItem.Index = TAB_ADIANTAMENTOSCREDITOS And iUltTipoBaixa <> gobjBaixaAntecipCredFornecedor.iTipoBaixa And iFramePrincipalAlterado <> 0 Then
            
'            'Exibe e oculta os controles específicos para o tipo de baixa selecionado
'            Select Case gobjBaixaAntecipCredFornecedor.iTipoBaixa
'
'                Case TIPOBAIXA_ADIANTAMENTO
'
'                    'Torna visível o frame que contém o grid que exibirá os adiantamentos
'                    FrameAdiantamentos.Visible = True
'
'                    'Torna invisível o frame que contém o grid que exibirá os creditos / devoluções
'                    FrameCreditos.Visible = False
'
'                    'Exibe os controles que são exclusivos para a baixa de adiantamentos
'                    ContaCorrente.Visible = True
'                    LabelCtaCorrente.Visible = True
'
'                    'Reposiciona os controles na tela
'                    LabelHistorico.Top = 1235
'                    Historico.Top = 1175
'
'                Case TIPOBAIXA_CREDITOS
'
'                    'Torna visível o frame que contém o grid que exibirá os creditos / devoluções
'                    FrameCreditos.Visible = True
'
'                    'Torna invisível o frame que contém o grid que exibirá os adiantamentos
'                    FrameAdiantamentos.Visible = False
'
'                    'Oculta os controles que são exclusivos para a baixa de adiantamentos
'                    ContaCorrente.Visible = False
'                    LabelCtaCorrente.Visible = False
'
'                    'Reposiciona os controles na tela
'                    LabelHistorico.Top = 975
'                    Historico.Top = 925
'
'            End Select
            
            'Carrega o Tab de Adiantamentos/Creditos
            lErro = Carrega_Tab_AdiantamentosCreditos()
            If lErro <> SUCESSO Then gError 95406
            
            iFramePrincipalAlterado = 0

        End If
    
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If TabStrip1.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
    
    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr
 
        Case 95406
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143196)

    End Select

    Exit Sub

End Sub

'Inicio do Tratamento do Grid de Adiantamentos
Public Sub GridAdiantamentos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridAdiantamentos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)
    End If

End Sub

Public Sub GridAdiantamentos_EnterCell()

    Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)

End Sub

Public Sub GridAdiantamentos_GotFocus()

    Call Grid_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub GridAdiantamentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAdiantamentos, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAdiantamentos, iAlterado)
    End If

End Sub

Public Sub GridAdiantamentos_LeaveCell()

    Call Saida_Celula(objGridAdiantamentos)

End Sub

Public Sub GridAdiantamentos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridAdiantamentos)

End Sub

Public Sub GridAdiantamentos_RowColChange()

    Call Grid_RowColChange(objGridAdiantamentos)

End Sub

Public Sub GridAdiantamentos_Scroll()

    Call Grid_Scroll(objGridAdiantamentos)

End Sub
'Fim do Tratamento do Grid de Adiantamentos

'Inicio do Tratamento do Grid de Créditos
Public Sub GridCreditos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCreditos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCreditos, iAlterado)
    End If

End Sub

Public Sub GridCreditos_EnterCell()

    Call Grid_Entrada_Celula(objGridCreditos, iAlterado)

End Sub

Public Sub GridCreditos_GotFocus()

    Call Grid_Recebe_Foco(objGridCreditos)

End Sub

Public Sub GridCreditos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCreditos, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCreditos, iAlterado)
    End If

End Sub

Public Sub GridCreditos_LeaveCell()

    Call Saida_Celula(objGridCreditos)

End Sub

Public Sub GridCreditos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCreditos)

End Sub

Public Sub GridCreditos_RowColChange()

    Call Grid_RowColChange(objGridCreditos)

End Sub

Public Sub GridCreditos_Scroll()

    Call Grid_Scroll(objGridCreditos)

End Sub
'Fim do Tratamento do Grid de Créditos


'Inicio do Tratamento dos Controles do Grid de Adiantamentos
Public Sub FilialAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub FilialAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub FilialAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = FilialAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub DataAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub DataAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = DataAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CtaCorrenteAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CtaCorrenteAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub CtaCorrenteAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub CtaCorrenteAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = CtaCorrenteAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub MeioPagtoAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MeioPagtoAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub MeioPagtoAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub MeioPagtoAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = MeioPagtoAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub ValorAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub ValorAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = ValorAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub SaldoAdiantamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SaldoAdiantamento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridAdiantamentos)

End Sub

Public Sub SaldoAdiantamento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAdiantamentos)

End Sub

Public Sub SaldoAdiantamento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridAdiantamentos.objControle = SaldoAdiantamento
    lErro = Grid_Campo_Libera_Foco(objGridAdiantamentos)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'Fim do Tratamento dos Controles do Grid de Adiantamentos


'Inicio do Tratamento dos Controles do Grid de Créditos
Public Sub FilialCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub FilialCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub FilialCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = FilialCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub DataCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub DataCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = DataCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TipoCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub TipoCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub TipoCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = TipoCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub NumCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub NumCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub NumCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = NumCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub ValorCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub ValorCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = ValorCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub SaldoCredito_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub SaldoCredito_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridCreditos)

End Sub

Public Sub SaldoCredito_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCreditos)

End Sub

Public Sub SaldoCredito_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridCreditos.objControle = SaldoCredito
    lErro = Grid_Campo_Libera_Foco(objGridCreditos)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'Fim do Tratamento dos Controles do Grid de Créditos

Private Sub TipoBaixaAdiantamento_Click()

    'Torna visível o frame de adiantamentos
    FrameFiltrosAdiantamento.Visible = True
    
    'Torna invisível o frame de créditos
    FrameFiltrosCredito.Visible = False
    
    'Guarda o tipo de baixa anterior
    iUltTipoBaixa = gobjBaixaAntecipCredFornecedor.iTipoBaixa
    
    'Guarda o tipo de baixa que está selecionada
    gobjBaixaAntecipCredFornecedor.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBaixaCreditos_Click()
    
    'Torna invisível o frame de créditos
    FrameFiltrosCredito.Visible = True
    
    'Torna visível o frame de adiantamentos
    FrameFiltrosAdiantamento.Visible = False
    
    'Guarda o tipo de baixa anterior
    iUltTipoBaixa = gobjBaixaAntecipCredFornecedor.iTipoBaixa
    
    'Guarda o tipo de baixa que está selecionada
    gobjBaixaAntecipCredFornecedor.iTipoBaixa = TIPOBAIXA_CREDITOS
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CtaCorrenteApenas_Click()

    'Habilita a combo CtaCorrente
    ContaCorrenteSeleciona.Enabled = True
    
    'Guarda que o filtro só será feito por uma conta
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteTodas = CONTACORRENTE_APENAS
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
   
End Sub

Private Sub CtaCorrenteTodas_Click()

    'Desabilita a combo CtaCorrente
    ContaCorrenteSeleciona.Enabled = False
    
    'Guarda que não haverá filtro por uma conta específica
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteTodas = CONTACORRENTE_TODAS
    
    'Limpar a seleção feita na combo
    ContaCorrenteSeleciona.ListIndex = -1
    
    'Limpar a seleção feita em gobjBaixaAntecipCredFornecedor
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = -1
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
   
End Sub

Private Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Private Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoInic_Validate

    'Se estiver preenchido => critica
    If Len(Trim(EmissaoInic.ClipText)) <> 0 Then
    
        lErro = Data_Critica(EmissaoInic.Text)
        If lErro <> SUCESSO Then gError 92714
        
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoDe = EmissaoInic.Text
        
    'Se não estiver preenchido => Limpa no obj
    Else
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoDe = DATA_NULA
    End If
    
    Exit Sub
    
Erro_EmissaoInic_Validate:
    
    Cancel = True
    
    Select Case gErr
 
        Case 92714
 
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143197)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoFim_Validate
    
    'Se estiver preenchido => critica e guarda no obj
    If Len(Trim(EmissaoFim.ClipText)) <> 0 Then
    
        'verifica se a data é valida
        lErro = Data_Critica(EmissaoFim.Text)
        If lErro <> SUCESSO Then gError 92713
        
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoAte = EmissaoFim.Text
        
    'Se não estiver preenchido => Limpa no obj
    Else
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoAte = DATA_NULA
    End If
    
    Exit Sub
    
Erro_EmissaoFim_Validate:
    
    Cancel = True
    
    Select Case gErr
 
        Case 92713
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143198)

    End Select

    Exit Sub
    
End Sub

Private Sub ContaCorrenteSeleciona_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumCreditoInic_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumCreditoFim_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoCreditoApenas_Click()

    'Habilita a combo Tipo de Credito
    TipoCreditoSeleciona.Enabled = True
    
    'Guarda que o filtro só será feito por um tipo de crédito
    gobjBaixaAntecipCredFornecedor.iTipoCreditoTodos = TIPOCREDITO_APENAS
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub TipoCreditoTodos_Click()

    'Desabilita a combo Tipo de Credito
    TipoCreditoSeleciona.Enabled = False
    
    'Guarda que não haverá filtro por um um tipo de crédito
    gobjBaixaAntecipCredFornecedor.iTipoCreditoTodos = TIPOCREDITO_TODOS
    
    'Limpar a seleção feita na combo
    TipoCreditoSeleciona.ListIndex = COMBO_INDICE
    
    'Guarda no obj a informação de que não há um tipo de débito específico selecionado
    gobjBaixaAntecipCredFornecedor.sTipoCredito = ""
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO
   
End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_Fornecedor_Validate

    'Se nao foi alterado => sai da funcao
    If iFornecedorAlterado = 0 Then Exit Sub

    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) <> 0 Then

        'Le o Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then gError 95392

        'Le a Filial do Fornecedor
'        lErro = CF("FilialFornecedor_Le",objFilialFornecedor)
'        If lErro <> SUCESSO Then gError 95393

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then gError 95393

        'Preenche Filial
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se Fornecedor não está preenchido
    Else
        'Limpa a Filial
        Filial.Clear

    End If
    
    'Guarda o codigo atual do Fornecedor/Filial
    gobjBaixaAntecipCredFornecedor.lFornecedor = objFornecedor.lCodigo
    gobjBaixaAntecipCredFornecedor.iFilialFornecedor = iCodFilial
    
    iFornecedorAlterado = 0

    Cancel = False

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 95392, 95393

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143199)

    End Select

    Exit Sub

End Sub

Private Sub LabelFornecedor_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Guarda o Nome reduzido
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    'Chama o browser de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

End Sub

Private Sub ContaCorrenteSeleciona_Validate(Cancel As Boolean)
'Faz a validação da ContaCorrenteSeleciona

Dim lErro As Long
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaCorrenteSeleciona_Validate

    'Limpa qualquer lixo que esteja no obj
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = -1

    'Se a ComboBox ContaCorrenteSeleciona não está preenchida => sai da função
    If Len(Trim(ContaCorrenteSeleciona.Text)) = 0 Then Exit Sub

    'Se estiver preenchida com o item selecionado na ComboBox => preenche o obj e sai da função.
    If ContaCorrenteSeleciona.Text = ContaCorrenteSeleciona.List(ContaCorrenteSeleciona.ListIndex) Then
        gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = ContaCorrenteSeleciona.ItemData(ContaCorrenteSeleciona.ListIndex)
        Exit Sub
    End If
    
    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrenteSeleciona, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 95394

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Guarda o Codigo da Conta Corrente
        objContaCorrente.iCodigo = iCodigo
        
        'Lê os dados da conta corrente a partir do código passado como parâmetro
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 95395

        'Se não encontrou a conta => Pergunta se Quer criar
        If lErro = 11807 Then gError 95396
        
        'Coloca os dados da conta na combo
        ContaCorrenteSeleciona.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido

    'Não existe o ítem com a STRING na List da ComboBox
    ElseIf lErro = 6731 Then
    
        'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
        objContaCorrente.sNomeReduzido = ContaCorrenteSeleciona.Text
        lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
        If lErro <> SUCESSO And lErro <> 15981 Then gError 95397

        'Se não encontrou a conta => Erro
        If lErro = 15981 Then gError 95398
        
        'Coloca os dados da conta na combo
        ContaCorrenteSeleciona.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
        
    End If
    
    'Guarda o código da conta no obj global
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = objContaCorrente.iCodigo
    
    Exit Sub
    
Erro_ContaCorrenteSeleciona_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 95394, 95395, 95397
                
        Case 95398 '??? Luiz: erro 95398 tem q exibir msg indicando q nao encontrou!
                   'Daniel: OK !!!
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, iCodigo)
        
        Case 95396
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteSeleciona.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de ContaCorrente
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143200)

    End Select

    Exit Sub

End Sub


Private Sub Filial_Validate(Cancel As Boolean)
'Faz a validação da ContaCorrenteSeleciona

Dim lErro As Long
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sNomeRed As String
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Limpa qualquer lixo que esteja no obj
    gobjBaixaAntecipCredFornecedor.iFilialFornecedor = -1

    'Se não houver conta selecionada => sai da função
    If Filial.ListIndex = -1 Then Exit Sub
    
    'Se o Fornecedor nao estiver preenchido => sai da funcao
    If Len(Trim(Fornecedor.Text)) = 0 Then
        Exit Sub
    Else
        'Guarda o nome do fornecedor
        sNomeRed = Fornecedor.Text
    End If
    
    'Se estiver preenchida com o item selecionado na ComboBox => preenche o obj e sai da função.
    If Filial.Text = Filial.List(Filial.ListIndex) Then
        gobjBaixaAntecipCredFornecedor.iFilialFornecedor = Filial.ItemData(Filial.ListIndex)
        Exit Sub
    End If
    
    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 95400

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        '??? Luiz: passar o codigo da filial dentro do obj
        'Guarda o Codigo da Filial do Fornecedor
        objFilialFornecedor.iCodFilial = iCodigo
        
        'Le o Codigo da Filial do Fornecedor
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 95401

        'Se não encontrou a filial => Pergunta se Quer criar
        If lErro = 18272 Then gError 95402
        
        'Coloca os dados da filial na combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    '??? Luiz: não é para ser feita leitura. Apenas dá erro direto
    'Daniel: OK !!!
    'Não existe o ítem com a STRING na List da ComboBox
    ElseIf lErro = 6731 Then
    
        'Le o NomeReduzido da Filial do Fornecedor
        objFilialFornecedor.sNome = sNomeRed
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 95403

        'Se não encontrou a filial => Erro
        If lErro = 18272 Then gError 95404
        
        'Coloca os dados da filial na combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Guarda o código da filial no obj global
    gobjBaixaAntecipCredFornecedor.iFilialFornecedor = objFilialFornecedor.iCodFilial
    
    Exit Sub
    
Erro_Filial_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 95400, 95401, 95403
                
        Case 95404
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_FORNECEDOR_INEXISTENTE", gErr, objFilialFornecedor.iCodFilial)
        
        Case 95402
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", sNomeRed, CStr(objFilialFornecedor.iCodFilial))

            If vbMsgRes = vbYes Then
                'Chama a tela de FilialFornecedor
                '??? Luiz: mudar o nome da tela
                'Daniel: OK !!!
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143201)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_AdiantamentosCreditos() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_AdiantamentosCreditos

    'Inicializa a colecao
    Set gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor = Nothing

    'Carrega o grid de acordo com o tipo da baixa previamente selecionado na tela
    Select Case gobjBaixaAntecipCredFornecedor.iTipoBaixa
    
        Case TIPOBAIXA_ADIANTAMENTO
        
            Call Limpa_Tab_AdiantamentosCreditos
        
            'Chama a funcao para fazer a leitura no BD
            lErro = CF("BaixaAntecipCredForn_Le_Adiantamentos", gobjBaixaAntecipCredFornecedor)
            If lErro <> SUCESSO And lErro <> 92708 Then gError 95433
            
            'se não houverem adiantamentos ==> erro
            If lErro = 92708 Then gError 92709
            
            'Torna o Frame com o grid de Adiantamentos visível
            FrameAdiantamentos.Visible = True
            
            'Torna o Frame com o grid de Créditos invisível
            FrameCreditos.Visible = False
            
            'Traz os dados para a Tela
            Call Traz_Adiantamento_Tela

        Case TIPOBAIXA_CREDITOS
        
            Call Limpa_Tab_AdiantamentosCreditos
        
            'Chama a funcao para fazer a leitura no BD
            lErro = CF("BaixaAntecipCredForn_Le_Creditos", gobjBaixaAntecipCredFornecedor)
            If lErro <> SUCESSO And lErro <> 92677 Then gError 95434

            'se não houverem créditos ==> erro
            If lErro = 92677 Then gError 92710

            'Torna o Frame com o grid de Adiantamentos invisível
            FrameAdiantamentos.Visible = False
            
            'Torna o Frame com o grid de Créditos visível
            FrameCreditos.Visible = True
            
            'Traz os dados para a Tela
            Call Traz_Creditos_Tela
        
    End Select

    Carrega_Tab_AdiantamentosCreditos = SUCESSO
    
    Exit Function

Erro_Carrega_Tab_AdiantamentosCreditos:

    Carrega_Tab_AdiantamentosCreditos = gErr
    
    Select Case gErr
            
        Case 92709
            Call Rotina_Erro(vbOKOnly, "ERRO_ANTECIPPAG_INEXISTENTE3", gErr)
        
        Case 92710
            Call Rotina_Erro(vbOKOnly, "ERRO_CREDITOPAGFORN_INEXISTENTE", gErr)
                    
        Case 95433, 95434
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143202)
    
    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92790
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 95435

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 92790

        Case 95435
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143203)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_TabStrip1_BeforeClick

    If TabStrip1.SelectedItem.Index = TAB_FILTROS Then

        lErro = Valida_Filtros()
        If lErro <> SUCESSO Then gError 95440

    End If

    Exit Sub

Erro_TabStrip1_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 95440

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143204)

    End Select

    Exit Sub

End Sub

Private Function Valida_Filtros() As Long

On Error GoTo Erro_Valida_Filtros

    'Se o fornecedor não foi preenchido => erro
    If gobjBaixaAntecipCredFornecedor.lFornecedor = 0 Then gError 95441

    'Se a filial do fornecedor não foi preenchida => erro
    If gobjBaixaAntecipCredFornecedor.iFilialFornecedor = 0 Then gError 95442

    'Se as datas de emissão (De e Até) foram preenchidas => verifica se data Até não é menor do que a data De
    If gobjBaixaAntecipCredFornecedor.dtDataEmissaoAte <> DATA_NULA And gobjBaixaAntecipCredFornecedor.dtDataEmissaoDe <> DATA_NULA Then

        'Se a data Até for menor do que a Data De => erro
        If gobjBaixaAntecipCredFornecedor.dtDataEmissaoAte < gobjBaixaAntecipCredFornecedor.dtDataEmissaoDe Then gError 95443

    End If

    Select Case gobjBaixaAntecipCredFornecedor.iTipoBaixa

        Case TIPOBAIXA_ADIANTAMENTO

            'Se o filtro de seleção por apenas uma conta corrente foi selecionado => Verifica se foi selecionada uma conta
            If gobjBaixaAntecipCredFornecedor.iCtaCorrenteTodas = CONTACORRENTE_APENAS Then

                'Se a Conta Corrente que será usada para filtro não foi informada => erro
                If gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = COMBO_INDICE Then gError 95444

            End If

        Case TIPOBAIXA_CREDITOS

        'Se os números de creditos(De e Até) foram preenchidos => verifica se o número Até não é menor do que o número De
        If gobjBaixaAntecipCredFornecedor.lNumCreditoDe <> 0 And gobjBaixaAntecipCredFornecedor.lNumCreditoAte <> 0 Then

            'Se a data Até for menor do que a Data De => erro
            If gobjBaixaAntecipCredFornecedor.lNumCreditoAte < gobjBaixaAntecipCredFornecedor.lNumCreditoDe Then gError 95445

        End If

        'Se o filtro de seleção por apenas um tipo de débito foi selecionado => verifica se foi selecionado um tipo
        If gobjBaixaAntecipCredFornecedor.iTipoCreditoTodos = TIPOCREDITO_APENAS Then
        
            'Se o tipo credito que será usado para filtro não foi informado => erro
            If Len(Trim(gobjBaixaAntecipCredFornecedor.sTipoCredito)) = 0 Then gError 92712
        
        End If

    End Select



    Valida_Filtros = SUCESSO

    Exit Function

Erro_Valida_Filtros:

    Valida_Filtros = gErr

    Select Case gErr

        Case 95441
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 95442
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_INFORMADA", gErr)

        Case 95443
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)

        Case 95444
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)

        Case 95445
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", gErr)

        Case 92712

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143205)

    End Select

    Exit Function

End Function

Sub UpDownEmissaoInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_UpClick

    lErro = Data_Up_Down_Click(EmissaoInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95447
    
    Exit Sub
    
Erro_UpDownEmissaoInic_UpClick:

    Select Case gErr
            
        Case 95447
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143206)
            
    End Select

End Sub

Sub UpDownEmissaoInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_DownClick

    lErro = Data_Up_Down_Click(EmissaoInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95448
    
    Exit Sub
    
Erro_UpDownEmissaoInic_DownClick:

    Select Case gErr
            
        Case 95448
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143207)
            
    End Select

End Sub

Sub UpDownEmissaoFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_UpClick

    lErro = Data_Up_Down_Click(EmissaoFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95449
    
    Exit Sub
    
Erro_UpDownEmissaoFim_UpClick:

    Select Case gErr
            
        Case 95449
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143208)
            
    End Select

End Sub

Sub UpDownEmissaoFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_DownClick

    lErro = Data_Up_Down_Click(EmissaoFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95450
    
    Exit Sub
    
Erro_UpDownEmissaoFim_DownClick:

    Select Case gErr
            
        Case 95450
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143209)
            
    End Select

End Sub

Sub UpDownDataBaixa_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_UpClick

    lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 95451
    
    Exit Sub
    
Erro_UpDownDataBaixa_UpClick:

    Select Case gErr
            
        Case 95451
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143210)
            
    End Select

End Sub

Sub UpDownDataBaixa_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_DownClick

    lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 95451
    
    Exit Sub
    
Erro_UpDownDataBaixa_DownClick:

    Select Case gErr
            
        Case 95451
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143211)
            
    End Select

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
    
    'Verifica se tem alteracao na tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 95452

    'Limpa a tela
    Call Limpa_Tela_BaixaAntecipCredFornecedor

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
            
        Case 95452
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143212)
            
    End Select
    
    Exit Sub

End Sub

Sub Limpa_Tela_BaixaAntecipCredFornecedor()

    'Limpa a tela
    Call Limpa_Tela(Me)
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    'Selecionar as options button defaults
    TipoBaixaAdiantamento.Value = True
    CtaCorrenteTodas.Value = True
    
    'Limpar os campos de data
    EmissaoInic.PromptInclude = False
    EmissaoInic.Text = ""
    EmissaoInic.PromptInclude = True
    
    EmissaoFim.PromptInclude = False
    EmissaoFim.Text = ""
    EmissaoFim.PromptInclude = True
    
    'Limpar as combos do frame principal
    Filial.ListIndex = -1
    ContaCorrenteSeleciona.ListIndex = -1
    
    'Limpa o TAB
    Call Limpa_Tab_AdiantamentosCreditos
    
    'Setar os campos iClienteAlterado, iFramePrincipalAlterado e iAlterado
    iFornecedorAlterado = 0
    iFramePrincipalAlterado = 0
    
    Call Inicializa_gobjBaixaAntecipCredFornecedor(TELA_TODA)
    
    iAlterado = 0

End Sub

Private Sub Inicializa_gobjBaixaAntecipCredFornecedor(Optional iTelaToda As Integer = 0)

    If iTelaToda = TELA_TODA Then
    
        'Inicializa o obj que serve de "espelho" à tela
        Set gobjBaixaAntecipCredFornecedor = New ClassBaixaAntecipCredForn
    
        'Guarda no obj global a informação de que não existem cliente e filial selecionados
        gobjBaixaAntecipCredFornecedor.lFornecedor = 0
        gobjBaixaAntecipCredFornecedor.iFilialFornecedor = -1
    
        'Guarda no obj global a informação do tipo de baixa selecionado
        gobjBaixaAntecipCredFornecedor.iTipoBaixa = TIPOBAIXA_ADIANTAMENTO
    
        'Guarda na variável a informação do último de baixa que foi selecionado ao se clicar no tabstrip
        iUltTipoBaixa = TIPOBAIXA_CREDITOS
    
        'Guarda no obj global a informação que todas as contas correntes estão selecionadas
        gobjBaixaAntecipCredFornecedor.iCtaCorrenteTodas = CONTACORRENTE_TODAS
    
        'Guarda no obj global a informação que não existe uma conta específica selecionada
        gobjBaixaAntecipCredFornecedor.iCtaCorrenteSelecionada = COMBO_INDICE
    
        'Guarda no obj global a informação de que não foram informadas datas de emissão
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoAte = DATA_NULA
        gobjBaixaAntecipCredFornecedor.dtDataEmissaoDe = DATA_NULA
    
        'Guarda no obj global a informação de que não foram informados números de débitos
        gobjBaixaAntecipCredFornecedor.lNumCreditoDe = 0
        gobjBaixaAntecipCredFornecedor.lNumCreditoAte = 0
        
        'Guarda no obj global a informação de que todos os tipos de débito estão selecionados
        gobjBaixaAntecipCredFornecedor.iTipoCreditoTodos = TIPOCREDITO_TODOS
        
        'Guarda no obj global a informação que não existe um tipo específico selecionado
        gobjBaixaAntecipCredFornecedor.sTipoCredito = ""
    
    End If
    
    'Guarda no obj global a data da baixa
    gobjBaixaAntecipCredFornecedor.dtDataBaixa = gdtDataHoje
    
    'Guarda no obj global a informação de que o valor das parcelas selecionadas é zero
    gobjBaixaAntecipCredFornecedor.dValorBaixa = 0
    
    'Guarda no obj global a informação de que ainda não existe conta selecionada para o movimento de baixa
    gobjBaixaAntecipCredFornecedor.iCtaCorrenteBaixa = COMBO_INDICE
    
    'Guarda no obj global a informação de que não há histórico informado
    gobjBaixaAntecipCredFornecedor.sHistoricoBaixa = ""

End Sub

Sub Limpa_Tab_AdiantamentosCreditos()

    'Limpa o Grid de Adiantamentos
    Call Grid_Limpa(objGridAdiantamentos)
    
    'Limpa o Grid de Creditos
    Call Grid_Limpa(objGridCreditos)
    
    DataBaixa.PromptInclude = False
    DataBaixa.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataBaixa.PromptInclude = True
    
    ContaCorrente.ListIndex = COMBO_INDICE
    Historico.Text = ""
    LabelValorBaixa2.Caption = ""
    
    Call Inicializa_gobjBaixaAntecipCredFornecedor
        
End Sub

Private Sub Traz_Adiantamento_Tela()

Dim lErro As Long
Dim objInfoBaixaAntecipCreForn As New ClassInfoBaixaAntecipCreForn
Dim iLinha As Integer

On Error GoTo Erro_Traz_Adiantamento_Tela

    For Each objInfoBaixaAntecipCreForn In gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor
        
        With objInfoBaixaAntecipCreForn
            
        'Incrementa a linha
        iLinha = iLinha + 1
        
        GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoFilial_Col) = .iFilial
        GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoData_Col) = Format(.dtDataEmissao, "dd/mm/yyyy")
        GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoCtaCorrente_Col) = .iCodConta & SEPARADOR & .sNomeReduzidoConta
        
        Select Case .iTipoMeioPagto
        
            Case DINHEIRO
                GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoMeioPagto_Col) = "Dinheiro"
                                        
            Case Cheque
                GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoMeioPagto_Col) = "Cheque"
            
            Case BORDERO
                GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoMeioPagto_Col) = "Borderô"
        
        End Select
        
        GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoValor_Col) = Format(.dValor, "standard")
        GridAdiantamentos.TextMatrix(iLinha, iGrid_AdiantamentoSaldo_Col) = Format(.dSaldoNaoApropriado, "standard")
                
        End With
        
    Next
    
    'Atualiza o numero de Linhas Existentes no grid
    objGridAdiantamentos.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridAdiantamentos)
    
    Exit Sub

Erro_Traz_Adiantamento_Tela:

    Select Case gErr
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143213)
            
    End Select
    
End Sub

Private Sub Traz_Creditos_Tela()

Dim objAntecipCredForn As New ClassInfoBaixaAntecipCreForn
Dim iLinha As Integer

On Error GoTo Erro_Traz_Creditos_Tela
    
    For Each objAntecipCredForn In gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor
    
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
       With objAntecipCredForn
        
            'Exibe a filial do crébito no grid
            GridCreditos.TextMatrix(iLinha, iGrid_CreditoFilial_Col) = .iFilial
            
            If .dtDataEmissao <> DATA_NULA Then
            
                'Exibe a data do crébito
                GridCreditos.TextMatrix(iLinha, iGrid_CreditoData_Col) = Format(.dtDataEmissao, "dd/mm/yyyy")
            
            End If
            
            'Exibe a sigla do tipo do crédito
            GridCreditos.TextMatrix(iLinha, iGrid_CreditoTipo_Col) = .sSiglaDocumento
            
            'Exibe o número do crébito
            GridCreditos.TextMatrix(iLinha, iGrid_CreditoNumero_Col) = .lNumero
            
            'Exibe o valor total do crébito no formato moeda
            GridCreditos.TextMatrix(iLinha, iGrid_CreditoValor_Col) = Format(.dValor, "Standard")
            
            'Exibe o saldo ainda não utilizado do crébito no formato moeda
            GridCreditos.TextMatrix(iLinha, iGrid_CreditoSaldo_Col) = Format(.dSaldoNaoApropriado, "Standard")
        
        End With
    
    Next
        
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridCreditos.iLinhasExistentes = gobjBaixaAntecipCredFornecedor.colAntecipCredFornecedor.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridCreditos)
    
    Exit Sub

Erro_Traz_Creditos_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143214)
        
    End Select
    
    Exit Sub
    
End Sub

'inicio contabilidade

Private Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Private Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Private Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

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

Private Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

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
Dim objFornecedor As New ClassFornecedor, sContaTela As String
Dim objFilial As New ClassFilialFornecedor, objConta As New ClassContasCorrentesInternas

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 92797
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 92798
            
                objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 92799
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 92800
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 92801
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 92802
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case CONTA_COD
        
            If ContaCorrente.ListIndex <> -1 Then
                objMnemonicoValor.colValor.Add ContaCorrente.ItemData(ContaCorrente.ListIndex)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case CONTA_CONTABIL_CONTA
        
            If ContaCorrente.ListIndex <> -1 Then
            
                objConta.iCodigo = ContaCorrente.ItemData(ContaCorrente.ListIndex)
                lErro = CF("ContaCorrenteInt_Le", objConta.iCodigo, objConta)
                If lErro <> SUCESSO Then gError 92803
                
                If objConta.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objConta.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 92804
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case VALOR1
            If Len(Trim(LabelValorBaixa2.Caption)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(LabelValorBaixa2.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case HISTORICO1
            If Len(Historico.Text) > 0 Then
                objMnemonicoValor.colValor.Add Historico.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case DATA1
            If Len(Trim(DataBaixa.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataBaixa.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case Else
            gError 92805

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 92797, 92798, 92799, 92800, 92801, 92802, 92803, 92804
        
        Case 92805
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143215)

    End Select

    Exit Function

End Function

Private Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134050

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134050

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143216)

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


