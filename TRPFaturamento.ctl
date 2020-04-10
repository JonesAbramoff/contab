VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPFaturamento 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   1
      Left            =   135
      TabIndex        =   43
      Top             =   660
      Width           =   9210
      Begin VB.Frame FrameS 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4725
         Index           =   1
         Left            =   165
         TabIndex        =   70
         Top             =   405
         Width           =   8970
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Index           =   12
            Left            =   7200
            Picture         =   "TRPFaturamento.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3555
            Width           =   780
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Index           =   12
            Left            =   8040
            Picture         =   "TRPFaturamento.ctx":101A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3555
            Width           =   780
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Index           =   11
            Left            =   7185
            Picture         =   "TRPFaturamento.ctx":21FC
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   2310
            Width           =   780
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Index           =   11
            Left            =   8025
            Picture         =   "TRPFaturamento.ctx":3216
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2310
            Width           =   780
         End
         Begin VB.Frame Frame11 
            Caption         =   "Desconsiderar os Vouchers abaixo"
            Height          =   2130
            Left            =   15
            TabIndex        =   108
            Top             =   2550
            Width           =   4440
            Begin MSMask.MaskEdBox ExcVouData 
               Height          =   240
               Left            =   2670
               TabIndex        =   112
               Top             =   615
               Width           =   1005
               _ExtentX        =   1773
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   8
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcVouNum 
               Height          =   240
               Left            =   1920
               TabIndex        =   111
               Top             =   615
               Width           =   770
               _ExtentX        =   1349
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcVouSerie 
               Height          =   240
               Left            =   1200
               TabIndex        =   110
               Top             =   660
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   1
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ExcVouTipo 
               Height          =   240
               Left            =   480
               TabIndex        =   109
               Top             =   675
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   1
               PromptChar      =   " "
            End
            Begin VB.CommandButton BotaoExcVou 
               Caption         =   "Vouchers"
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
               Left            =   225
               TabIndex        =   4
               Top             =   1740
               Width           =   1470
            End
            Begin MSFlexGridLib.MSFlexGrid GridExcVou 
               Height          =   1065
               Left            =   255
               TabIndex        =   3
               Top             =   270
               Width           =   3990
               _ExtentX        =   7038
               _ExtentY        =   1879
               _Version        =   393216
               Rows            =   6
               Cols            =   3
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
            End
         End
         Begin VB.ListBox TipoFaturamento 
            Columns         =   3
            Height          =   960
            ItemData        =   "TRPFaturamento.ctx":43F8
            Left            =   4500
            List            =   "TRPFaturamento.ctx":43FA
            Style           =   1  'Checkbox
            TabIndex        =   11
            Top             =   3555
            Width           =   2640
         End
         Begin VB.CheckBox optIndividual 
            Caption         =   "Faturar cada vou. individualmente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7170
            TabIndex        =   15
            Top             =   4230
            Width           =   1845
         End
         Begin VB.ListBox FiliaisEmpresa 
            Height          =   1410
            ItemData        =   "TRPFaturamento.ctx":43FC
            Left            =   4500
            List            =   "TRPFaturamento.ctx":4412
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   645
            Width           =   2640
         End
         Begin VB.ListBox TipoDoc 
            Columns         =   3
            Height          =   960
            ItemData        =   "TRPFaturamento.ctx":44AF
            Left            =   4500
            List            =   "TRPFaturamento.ctx":44B1
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   2310
            Width           =   2640
         End
         Begin VB.CommandButton BotaoDesmarcarTodos 
            Height          =   480
            Index           =   1
            Left            =   8010
            Picture         =   "TRPFaturamento.ctx":44B3
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   630
            Width           =   780
         End
         Begin VB.CommandButton BotaoMarcarTodos 
            Height          =   480
            Index           =   1
            Left            =   7170
            Picture         =   "TRPFaturamento.ctx":5695
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   630
            Width           =   780
         End
         Begin VB.Frame Frame3 
            Caption         =   "Desconsiderar os clientes abaixo"
            Height          =   1830
            Left            =   30
            TabIndex        =   71
            Top             =   555
            Width           =   4440
            Begin VB.CommandButton BotaoClientes 
               Caption         =   "Clientes"
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
               Left            =   225
               TabIndex        =   2
               Top             =   1440
               Width           =   1470
            End
            Begin MSMask.MaskEdBox ExcCliente 
               Height          =   240
               Left            =   870
               TabIndex        =   72
               Top             =   585
               Width           =   2970
               _ExtentX        =   5239
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AllowPrompt     =   -1  'True
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridExcCliente 
               Height          =   1065
               Left            =   225
               TabIndex        =   1
               Top             =   255
               Width           =   3990
               _ExtentX        =   7038
               _ExtentY        =   1879
               _Version        =   393216
               Rows            =   6
               Cols            =   3
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               FocusRect       =   2
               HighLight       =   0
            End
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   315
            Left            =   885
            TabIndex        =   0
            Top             =   45
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   556
            _Version        =   393216
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Faturamento"
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
            Left            =   4515
            TabIndex        =   106
            Top             =   3345
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
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
            Left            =   4515
            TabIndex        =   76
            Top             =   2100
            Width           =   1680
         End
         Begin VB.Label DescCliente 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3165
            TabIndex        =   75
            Top             =   45
            Width           =   5655
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
            Height          =   195
            Left            =   105
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   74
            Top             =   75
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filiais Empresas"
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
            Left            =   4530
            TabIndex        =   73
            Top             =   420
            Width           =   1365
         End
      End
      Begin VB.Frame FrameS 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   4740
         Index           =   2
         Left            =   135
         TabIndex        =   77
         Top             =   360
         Visible         =   0   'False
         Width           =   8985
         Begin VB.Frame Frame12 
            Caption         =   "Padrão"
            Height          =   690
            Left            =   90
            TabIndex        =   115
            Top             =   -15
            Width           =   8820
            Begin VB.CommandButton BotaoVencAplicar 
               Caption         =   "Apl."
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
               Left            =   8160
               TabIndex        =   124
               Top             =   210
               Width           =   495
            End
            Begin VB.CommandButton BotaoEmiVouAplicar 
               Caption         =   "Apl."
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
               Left            =   2670
               TabIndex        =   118
               Top             =   210
               Width           =   495
            End
            Begin VB.CommandButton BotaoEmiAplicar 
               Caption         =   "Apl."
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
               Left            =   5550
               TabIndex        =   121
               Top             =   225
               Width           =   495
            End
            Begin MSComCtl2.UpDown UpDownDataEmiPadrao 
               Height          =   300
               Left            =   5280
               TabIndex        =   120
               Top             =   240
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmiPadrao 
               Height          =   300
               Left            =   4230
               TabIndex        =   119
               Top             =   240
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataEmiVouPadrao 
               Height          =   300
               Left            =   2400
               TabIndex        =   117
               Top             =   225
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmiVouPadrao 
               Height          =   300
               Left            =   1350
               TabIndex        =   116
               Top             =   225
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataVencPadrao 
               Height          =   300
               Left            =   7890
               TabIndex        =   123
               Top             =   225
               Width           =   240
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataVencPadrao 
               Height          =   300
               Left            =   6840
               TabIndex        =   122
               Top             =   225
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Venc.:"
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
               Left            =   6225
               TabIndex        =   127
               Top             =   270
               Width           =   570
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Emis. Vou até:"
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
               Left            =   105
               TabIndex        =   126
               Top             =   270
               Width           =   1245
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Emis. Fat:"
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
               Left            =   3375
               TabIndex        =   125
               Top             =   285
               Width           =   855
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Condições de Pagamento"
            Height          =   3975
            Left            =   90
            TabIndex        =   78
            Top             =   735
            Width           =   8820
            Begin VB.CheckBox CPSelecionado 
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
               Left            =   2625
               TabIndex        =   84
               Top             =   2595
               Width           =   540
            End
            Begin MSMask.MaskEdBox CondPagto 
               Height          =   315
               Left            =   1170
               TabIndex        =   83
               Top             =   675
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CPDataEmissaoAte 
               Height          =   315
               Left            =   4395
               TabIndex        =   82
               Top             =   1275
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CPDataEmissaoDe 
               Height          =   315
               Left            =   2115
               TabIndex        =   81
               Top             =   1185
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CPDataEmissao 
               Height          =   315
               Left            =   3900
               TabIndex        =   80
               Top             =   2010
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox CPDataVencimento 
               Height          =   315
               Left            =   2205
               TabIndex        =   79
               Top             =   1800
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSFlexGridLib.MSFlexGrid GridCondPagto 
               Height          =   3420
               Left            =   90
               TabIndex        =   18
               Top             =   270
               Width           =   8640
               _ExtentX        =   15240
               _ExtentY        =   6033
               _Version        =   393216
               Rows            =   15
               Cols            =   8
               BackColorSel    =   -2147483643
               ForeColorSel    =   -2147483640
               AllowBigSelection=   0   'False
               Enabled         =   -1  'True
               FocusRect       =   2
            End
         End
      End
      Begin MSComctlLib.TabStrip TabStrip2 
         Height          =   5175
         Left            =   60
         TabIndex        =   69
         Top             =   0
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   9128
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Inicial"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Condições de Pagamento"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   5
      Left            =   105
      TabIndex        =   51
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoVoucher 
         Caption         =   "Voucher ..."
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
         Left            =   7755
         TabIndex        =   32
         Top             =   4620
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Itens (Vouchers, Ocorrências e Ocorrências de inativação)"
         Height          =   4095
         Left            =   60
         TabIndex        =   52
         Top             =   -15
         Width           =   9165
         Begin MSMask.MaskEdBox VouValorL 
            Height          =   255
            Left            =   4290
            TabIndex        =   114
            Top             =   3240
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouTipoV 
            Height          =   255
            Left            =   195
            TabIndex        =   105
            Top             =   1035
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouValorAporte 
            Height          =   255
            Left            =   6600
            TabIndex        =   96
            Top             =   3105
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouValorComi 
            Height          =   255
            Left            =   7455
            TabIndex        =   95
            Top             =   3075
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouValor 
            Height          =   255
            Left            =   5415
            TabIndex        =   67
            Top             =   3180
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouDataEmissao 
            Height          =   255
            Left            =   5265
            TabIndex        =   66
            Top             =   1695
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouNumero 
            Height          =   255
            Left            =   5130
            TabIndex        =   65
            Top             =   2640
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouSerie 
            Height          =   255
            Left            =   4290
            TabIndex        =   64
            Top             =   2790
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouTipo 
            Height          =   255
            Left            =   4410
            TabIndex        =   63
            Top             =   2055
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouFatSeq 
            Height          =   255
            Left            =   2445
            TabIndex        =   62
            Top             =   2370
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox VouCliente 
            Height          =   255
            Left            =   1275
            TabIndex        =   61
            Top             =   1485
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin VB.CheckBox VouSelecionado 
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
            Left            =   1380
            TabIndex        =   53
            Top             =   2490
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid GridVoucher 
            Height          =   510
            Left            =   30
            TabIndex        =   29
            Top             =   210
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   900
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   5
         Left            =   1785
         Picture         =   "TRPFaturamento.ctx":66AF
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4620
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   5
         Left            =   75
         Picture         =   "TRPFaturamento.ctx":7891
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4620
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Total Neto Vouchers:"
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
         Left            =   5655
         TabIndex        =   129
         Top             =   4230
         Width           =   1875
      End
      Begin VB.Label TotalVou 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   7560
         TabIndex        =   128
         Top             =   4185
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5220
      Index           =   4
      Left            =   120
      TabIndex        =   45
      Top             =   660
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Index           =   4
         Left            =   165
         Picture         =   "TRPFaturamento.ctx":88AB
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4620
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Index           =   4
         Left            =   1826
         Picture         =   "TRPFaturamento.ctx":98C5
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4605
         Width           =   1425
      End
      Begin VB.Frame Frame5 
         Caption         =   "Faturas a serem geradas"
         Height          =   4545
         Left            =   165
         TabIndex        =   49
         Top             =   15
         Width           =   8940
         Begin MSMask.MaskEdBox FatValorB 
            Height          =   255
            Left            =   7425
            TabIndex        =   131
            Top             =   870
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatValorDesc 
            Height          =   255
            Left            =   5595
            TabIndex        =   130
            Top             =   870
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatCondPagto 
            Height          =   255
            Left            =   4800
            TabIndex        =   97
            Top             =   1185
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatEmissao 
            Height          =   255
            Left            =   2685
            TabIndex        =   68
            Top             =   870
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatValor 
            Height          =   255
            Left            =   6510
            TabIndex        =   60
            Top             =   855
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatDataVenc 
            Height          =   255
            Left            =   3720
            TabIndex        =   59
            Top             =   870
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatSeq 
            Height          =   255
            Left            =   8325
            TabIndex        =   58
            Top             =   870
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FatCliente 
            Height          =   255
            Left            =   1080
            TabIndex        =   57
            Top             =   855
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin VB.CheckBox FatSelecionado 
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
            Left            =   615
            TabIndex        =   47
            Top             =   855
            Width           =   480
         End
         Begin MSFlexGridLib.MSFlexGrid GridFatura 
            Height          =   555
            Left            =   45
            TabIndex        =   26
            Top             =   255
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   979
            _Version        =   393216
            Rows            =   15
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5220
      Index           =   6
      Left            =   105
      TabIndex        =   98
      Top             =   675
      Visible         =   0   'False
      Width           =   9255
      Begin VB.Frame Frame10 
         Caption         =   "Opções de armazenamento"
         Height          =   990
         Left            =   120
         TabIndex        =   102
         Top             =   405
         Width           =   8955
         Begin VB.TextBox NomeDiretorio 
            Height          =   285
            Left            =   3420
            TabIndex        =   33
            Top             =   420
            Width           =   4770
         End
         Begin VB.CommandButton BotaoProcurar 
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
            Height          =   360
            Left            =   8250
            TabIndex        =   34
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Localização física dos arquivos html:"
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
            Height          =   300
            Index           =   2
            Left            =   195
            TabIndex        =   103
            Top             =   450
            Width           =   3225
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Opções de geração"
         Height          =   2505
         Left            =   135
         TabIndex        =   99
         Top             =   1605
         Width           =   8955
         Begin VB.CheckBox AbrirFatHtml 
            Caption         =   "Abrir faturas HTML"
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
            Left            =   645
            TabIndex        =   107
            Top             =   1980
            Width           =   2460
         End
         Begin VB.CommandButton BotaoModeloFat 
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
            Height          =   330
            Left            =   8220
            TabIndex        =   39
            Top             =   1875
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox ModeloFat 
            Height          =   315
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   1905
            Visible         =   0   'False
            Width           =   4770
         End
         Begin VB.OptionButton OptGerarEnviar 
            Caption         =   "Gerar e enviar por email"
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
            Left            =   3345
            TabIndex        =   36
            Top             =   315
            Width           =   2520
         End
         Begin VB.OptionButton OptSoGerar 
            Caption         =   "Somente gerar"
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
            Left            =   585
            TabIndex        =   35
            Top             =   315
            Value           =   -1  'True
            Width           =   2490
         End
         Begin VB.Frame Frame9 
            Caption         =   "Opções de envio de email"
            Height          =   1020
            Left            =   3390
            TabIndex        =   100
            Top             =   705
            Width           =   5445
            Begin VB.ComboBox Modelo 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "TRPFaturamento.ctx":AAA7
               Left            =   1860
               List            =   "TRPFaturamento.ctx":AAB1
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   390
               Width           =   3450
            End
            Begin VB.Label LabelModelo 
               Caption         =   "Modelo de email:"
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
               Left            =   345
               TabIndex        =   101
               Top             =   450
               Width           =   1455
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Modelo da Fatura em html:"
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
            Height          =   270
            Index           =   1
            Left            =   135
            TabIndex        =   104
            Top             =   1935
            Visible         =   0   'False
            Width           =   3225
         End
      End
      Begin VB.CommandButton botaoGerar 
         Caption         =   "Gerar Faturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   135
         TabIndex        =   40
         Top             =   4350
         Width           =   1860
      End
      Begin VB.CommandButton BotaoItemFat 
         Caption         =   "Itens a serem faturados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   6
         Left            =   7200
         TabIndex        =   41
         Top             =   4350
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5220
      Index           =   2
      Left            =   165
      TabIndex        =   85
      Top             =   660
      Visible         =   0   'False
      Width           =   9165
      Begin VB.CommandButton BotaoItemFat 
         Caption         =   "Itens a serem faturados por filial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   2
         Left            =   6945
         TabIndex        =   20
         Top             =   4635
         Width           =   2115
      End
      Begin VB.Frame Frame6 
         Caption         =   "Filiais"
         Height          =   4545
         Left            =   15
         TabIndex        =   86
         Top             =   45
         Width           =   9060
         Begin MSMask.MaskEdBox FilialNumVou 
            Height          =   255
            Left            =   3960
            TabIndex        =   87
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialValorSel 
            Height          =   255
            Left            =   6435
            TabIndex        =   88
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialNumFat 
            Height          =   255
            Left            =   2715
            TabIndex        =   89
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialEmpresa 
            Height          =   255
            Left            =   390
            TabIndex        =   90
            Top             =   1110
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridFilialEmpresa 
            Height          =   390
            Left            =   45
            TabIndex        =   19
            Top             =   240
            Width           =   8970
            _ExtentX        =   15822
            _ExtentY        =   688
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox FilialNumVouSel 
            Height          =   255
            Left            =   5190
            TabIndex        =   91
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox FilialValor 
            Height          =   255
            Left            =   7680
            TabIndex        =   92
            Top             =   1110
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5205
      Index           =   3
      Left            =   105
      TabIndex        =   44
      Top             =   660
      Visible         =   0   'False
      Width           =   9225
      Begin VB.CommandButton BotaoItemFat 
         Caption         =   "Itens a serem faturados por cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   3
         Left            =   7005
         TabIndex        =   25
         Top             =   4605
         Width           =   2115
      End
      Begin VB.CommandButton BotaoCliente 
         Caption         =   "Cliente ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5445
         TabIndex        =   24
         Top             =   4605
         Width           =   1455
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   555
         Index           =   3
         Left            =   150
         Picture         =   "TRPFaturamento.ctx":AAE6
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4605
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   555
         Index           =   3
         Left            =   1698
         Picture         =   "TRPFaturamento.ctx":BB00
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4605
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Caption         =   "Lista de Clientes"
         Height          =   4545
         Left            =   150
         TabIndex        =   50
         Top             =   -30
         Width           =   8985
         Begin MSMask.MaskEdBox CliFE 
            Height          =   255
            Left            =   1605
            TabIndex        =   113
            Top             =   2445
            Width           =   400
            _ExtentX        =   714
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CliValorSel 
            Height          =   255
            Left            =   6705
            TabIndex        =   93
            Top             =   2325
            Width           =   950
            _ExtentX        =   1667
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CliNumVouSel 
            Height          =   255
            Left            =   6180
            TabIndex        =   94
            Top             =   2850
            Width           =   950
            _ExtentX        =   1667
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CliValor 
            Height          =   255
            Left            =   3720
            TabIndex        =   56
            Top             =   1635
            Width           =   950
            _ExtentX        =   1667
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CliNumVou 
            Height          =   255
            Left            =   3735
            TabIndex        =   55
            Top             =   2385
            Width           =   950
            _ExtentX        =   1667
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CliCliente 
            Height          =   255
            Left            =   2025
            TabIndex        =   54
            Top             =   1065
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   450
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
            PromptChar      =   " "
         End
         Begin VB.CheckBox CliSelecionado 
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
            Left            =   885
            TabIndex        =   46
            Top             =   1095
            Width           =   510
         End
         Begin MSFlexGridLib.MSFlexGrid GridCliente 
            Height          =   4140
            Left            =   150
            TabIndex        =   21
            Top             =   240
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   7303
            _Version        =   393216
            Rows            =   16
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   7710
      ScaleHeight     =   480
      ScaleWidth      =   1605
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   45
      Width           =   1665
      Begin VB.CommandButton BotaoAtualizar 
         Height          =   360
         Left            =   75
         Picture         =   "TRPFaturamento.ctx":CCE2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Atualizar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   585
         Picture         =   "TRPFaturamento.ctx":D134
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1095
         Picture         =   "TRPFaturamento.ctx":D666
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5595
      Left            =   75
      TabIndex        =   42
      Top             =   330
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filiais"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Faturas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Itens"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Geração"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "TRPFaturamento"
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

Private Const TAB_Selecao = 1
Private Const TAB_FilialEmpresa = 2
Private Const TAB_CLIENTE = 3
Private Const TAB_FATURA = 4
Private Const TAB_VOUCHER = 5
Private Const TAB_GERACAO = 6

'Variáveis Globais
Dim iFrameAtual As Integer
Dim iFrameAtualS As Integer
Dim iAlterado As Integer
Dim iFrameSelecaoAlterado As Integer

Dim iTelaDesatualizada As Integer

Dim bDesabilitaCmdGridAux As Boolean

Dim gobjFaturamento As New ClassTRPFaturamento
Dim gcolCondPagto As New Collection

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoExcCliente As AdmEvento
Attribute objEventoExcCliente.VB_VarHelpID = -1
Private WithEvents objEventoVoucher As AdmEvento
Attribute objEventoVoucher.VB_VarHelpID = -1

'GridExcCliente
Dim objGridExcCliente As AdmGrid
Dim iGrid_ExcCliente_Col As Integer

'GridCliente
Dim objGridCliente As AdmGrid
Dim iGrid_CliSelecionado_Col As Integer
Dim iGrid_CliCliente_Col As Integer
Dim iGrid_CliFE_Col As Integer
Dim iGrid_CliValorFat_Col As Integer
Dim iGrid_CliNumVou_Col As Integer
Dim iGrid_CliValorFatSel_Col As Integer
Dim iGrid_CliNumVouSel_Col As Integer

'GridFatura
Dim objGridFatura As AdmGrid
Dim iGrid_FatSelecionado_Col As Integer
Dim iGrid_FatCliente_Col As Integer
Dim iGrid_FatSeq_Col As Integer
Dim iGrid_FatValor_Col As Integer
Dim iGrid_FatValorB_Col As Integer
Dim iGrid_FatValorDesc_Col As Integer
Dim iGrid_FatDataVenc_Col As Integer
Dim iGrid_FatEmissao_Col As Integer
Dim iGrid_FatCondPagto_Col As Integer

'GridVoucher
Dim objGridVoucher As AdmGrid
Dim iGrid_VouSelecionado_Col As Integer
Dim iGrid_VouCliente_Col As Integer
Dim iGrid_VouFatSeq_Col As Integer
Dim iGrid_VouTipo_Col As Integer
Dim iGrid_VouTipoV_Col As Integer
Dim iGrid_VouSerie_Col As Integer
Dim iGrid_VouDataEmissao_Col As Integer
Dim iGrid_VouNumero_Col As Integer
Dim iGrid_VouValor_Col As Integer
Dim iGrid_VouValorC_Col As Integer
Dim iGrid_VouValorA_Col As Integer
Dim iGrid_VouValorL_Col As Integer

'GridCondPagto
Dim objGridCondPagto As AdmGrid
Dim iGrid_CPSelecionado_Col As Integer
Dim iGrid_CondPagto_Col As Integer
Dim iGrid_CPDataEmissaoDe_Col As Integer
Dim iGrid_CPDataEmissaoAte_Col As Integer
Dim iGrid_CPDataEmissao_Col As Integer
Dim iGrid_CPDataVencimento_Col As Integer

'GridFilialEmpresa
Dim objGridFilialEmpresa As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_FilialNumFat_Col As Integer
Dim iGrid_FilialNumVou_Col As Integer
Dim iGrid_FilialValor_Col As Integer
Dim iGrid_FilialNumVouSel_Col As Integer
Dim iGrid_FilialValorSel_Col As Integer

'GridExcvou
Dim objGridExcVou As AdmGrid
Dim iGrid_ExcVouTipo_Col As Integer
Dim iGrid_ExcVouSerie_Col As Integer
Dim iGrid_ExcVouNum_Col As Integer
Dim iGrid_ExcVouData_Col As Integer

Const COR_CAMPO_OBRIGATORIO = &H80&
Const COR_CAMPO_NAO_OBRIGATORIO = &H80000012

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = TAB_Selecao
    iFrameAtualS = 1
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Set objEventoCliente = New AdmEvento
    Set objEventoExcCliente = New AdmEvento
    Set objEventoVoucher = New AdmEvento

    Set objGridCliente = New AdmGrid
    Set objGridExcCliente = New AdmGrid
    Set objGridFatura = New AdmGrid
    Set objGridVoucher = New AdmGrid
    Set objGridCondPagto = New AdmGrid
    Set objGridFilialEmpresa = New AdmGrid
    Set objGridExcVou = New AdmGrid
    
    bDesabilitaCmdGridAux = False

    'Inicializa o GridExcCliente
    lErro = Inicializa_Grid_ExcCliente(objGridExcCliente)
    If lErro <> SUCESSO Then gError 192085

    'Inicializa o GridExcCliente
    lErro = Inicializa_Grid_ExcVou(objGridExcVou)
    If lErro <> SUCESSO Then gError 192085

    'Inicializa o GridCliente
    lErro = Inicializa_Grid_Cliente(objGridCliente)
    If lErro <> SUCESSO Then gError 192086

    'Inicializa o GridFatura
    lErro = Inicializa_Grid_Fatura(objGridFatura)
    If lErro <> SUCESSO Then gError 192087

    'Inicializa o GridVoucher
    lErro = Inicializa_Grid_Voucher(objGridVoucher)
    If lErro <> SUCESSO Then gError 192088
    
    'Inicializa o GridCondPagto
    lErro = Inicializa_Grid_CondPagto(objGridCondPagto)
    If lErro <> SUCESSO Then gError 192089
    
    'Inicializa o GridFilialEmpresa
    lErro = Inicializa_Grid_FilialEmpresa(objGridFilialEmpresa)
    If lErro <> SUCESSO Then gError 192090
    
    'Lê as filiais empresas
    lErro = Carrega_FilialEmpresa
    If lErro <> SUCESSO Then gError 192091
    
'    lErro = Carrega_Grid_CondPagto
    lErro = Carrega_CategoriaClienteItem
    If lErro <> SUCESSO Then gError 192092

    lErro = CF("Carrega_Combo_TipoDoc", TipoDoc)
    If lErro <> SUCESSO Then gError 192093
       
    lErro = Carrega_Combo_Modelo
    If lErro <> SUCESSO Then gError 192094
    
    Call Default_Tela
    
    iTelaDesatualizada = DESMARCADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case 192085 To 192094

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192095)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    iAlterado = 0
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objGridCliente = Nothing
    Set objGridExcCliente = Nothing
    Set objGridFatura = Nothing
    Set objGridVoucher = Nothing
    Set objGridCondPagto = Nothing
    Set objGridFilialEmpresa = Nothing

    Set objEventoCliente = Nothing
    Set objEventoExcCliente = Nothing
    Set objEventoVoucher = Nothing
    
    Set gobjFaturamento = Nothing
    Set gcolCondPagto = Nothing

End Sub

Private Function Inicializa_Grid_Cliente(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Cliente

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("F.E.")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("N.Itens")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("N.Itens.S.")
    objGridInt.colColuna.Add ("Valor S.")

    'campos de edição do grid
    objGridInt.colCampo.Add (CliSelecionado.Name)
    objGridInt.colCampo.Add (CliFE.Name)
    objGridInt.colCampo.Add (CliCliente.Name)
    objGridInt.colCampo.Add (CliNumVou.Name)
    objGridInt.colCampo.Add (CliValor.Name)
    objGridInt.colCampo.Add (CliNumVouSel.Name)
    objGridInt.colCampo.Add (CliValorSel.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_CliSelecionado_Col = 1
    iGrid_CliFE_Col = 2
    iGrid_CliCliente_Col = 3
    iGrid_CliNumVou_Col = 4
    iGrid_CliValorFat_Col = 5
    iGrid_CliNumVouSel_Col = 6
    iGrid_CliValorFatSel_Col = 7

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCliente
    
    'Largura da primeira coluna
    GridCliente.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 1000 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 13

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Cliente = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Cliente:

    Inicializa_Grid_Cliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192096)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_FilialEmpresa(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid ItensRequisicoes

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_FilialEmpresa

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("N.Itens")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("N.Fat.S.")
    objGridInt.colColuna.Add ("N.Itens.S.")
    objGridInt.colColuna.Add ("Valor S.")

    'campos de edição do grid
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (FilialNumVou.Name)
    objGridInt.colCampo.Add (FilialValor.Name)
    objGridInt.colCampo.Add (FilialNumFat.Name)
    objGridInt.colCampo.Add (FilialNumVouSel.Name)
    objGridInt.colCampo.Add (FilialValorSel.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_FilialEmpresa_Col = 1
    iGrid_FilialNumVou_Col = 2
    iGrid_FilialValor_Col = 3
    iGrid_FilialNumFat_Col = 4
    iGrid_FilialNumVouSel_Col = 5
    iGrid_FilialValorSel_Col = 6

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridFilialEmpresa
    
    'Largura da primeira coluna
    GridFilialEmpresa.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 14

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_FilialEmpresa = SUCESSO

    Exit Function

Erro_Inicializa_Grid_FilialEmpresa:

    Inicializa_Grid_FilialEmpresa = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192097)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Fatura(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Fatura

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("C.Pagto")
    objGridInt.colColuna.Add ("Seq")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Venc.")
    objGridInt.colColuna.Add ("Bruto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Líquido")

    'campos de edição do grid
    objGridInt.colCampo.Add (FatSelecionado.Name)
    objGridInt.colCampo.Add (FatCliente.Name)
    objGridInt.colCampo.Add (FatCondPagto.Name)
    objGridInt.colCampo.Add (FatSeq.Name)
    objGridInt.colCampo.Add (FatEmissao.Name)
    objGridInt.colCampo.Add (FatDataVenc.Name)
    objGridInt.colCampo.Add (FatValorB.Name)
    objGridInt.colCampo.Add (FatValorDesc.Name)
    objGridInt.colCampo.Add (FatValor.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_FatSelecionado_Col = 1
    iGrid_FatCliente_Col = 2
    iGrid_FatCondPagto_Col = 3
    iGrid_FatSeq_Col = 4
    iGrid_FatEmissao_Col = 5
    iGrid_FatDataVenc_Col = 6
    iGrid_FatValorB_Col = 7
    iGrid_FatValorDesc_Col = 8
    iGrid_FatValor_Col = 9
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridFatura
    
    'Largura da primeira coluna
    GridFatura.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 1000 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 14

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Fatura = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Fatura:

    Inicializa_Grid_Fatura = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192098)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Voucher(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos2

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_Voucher

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("FAT")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("T")
    objGridInt.colColuna.Add ("S")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Vlr L.")
    objGridInt.colColuna.Add ("Vlr B.")
    objGridInt.colColuna.Add ("Vlr C.")
    objGridInt.colColuna.Add ("Vlr A.")

    'campos de edição do grid
    objGridInt.colCampo.Add (VouSelecionado.Name)
    objGridInt.colCampo.Add (VouCliente.Name)
    objGridInt.colCampo.Add (VouFatSeq.Name)
    objGridInt.colCampo.Add (VouTipo.Name)
    objGridInt.colCampo.Add (VouTipoV.Name)
    objGridInt.colCampo.Add (VouSerie.Name)
    objGridInt.colCampo.Add (VouNumero.Name)
    objGridInt.colCampo.Add (VouDataEmissao.Name)
    objGridInt.colCampo.Add (VouValorL.Name)
    objGridInt.colCampo.Add (VouValor.Name)
    objGridInt.colCampo.Add (VouValorComi.Name)
    objGridInt.colCampo.Add (VouValorAporte.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_VouSelecionado_Col = 1
    iGrid_VouCliente_Col = 2
    iGrid_VouFatSeq_Col = 3
    iGrid_VouTipo_Col = 4
    iGrid_VouTipoV_Col = 5
    iGrid_VouSerie_Col = 6
    iGrid_VouNumero_Col = 7
    iGrid_VouDataEmissao_Col = 8
    iGrid_VouValorL_Col = 9
    iGrid_VouValor_Col = 10
    iGrid_VouValorC_Col = 11
    iGrid_VouValorA_Col = 12

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridVoucher
    
    'Largura da primeira coluna
    GridVoucher.ColWidth(0) = 500

    'Linhas do grid
    objGridInt.objGrid.Rows = 1000 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 12

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Voucher = SUCESSO

    Exit Function

Erro_Inicializa_Grid_Voucher:

    Inicializa_Grid_Voucher = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192099)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ExcCliente(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ExcCliente

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cliente")

    'campos de edição do grid
    objGridInt.colCampo.Add (ExcCliente.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_ExcCliente_Col = 1
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridExcCliente

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3
    
    'Largura da primeira coluna
    GridExcCliente.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExcCliente = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ExcCliente:

    Inicializa_Grid_ExcCliente = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192100)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_CondPagto(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_CondPagto

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Cond. Pagto")
    objGridInt.colColuna.Add ("Vouchers emitidos De")
    objGridInt.colColuna.Add ("Até")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Data Vencimento")

    'campos de edição do grid
    objGridInt.colCampo.Add (CPSelecionado.Name)
    objGridInt.colCampo.Add (CondPagto.Name)
    objGridInt.colCampo.Add (CPDataEmissaoDe.Name)
    objGridInt.colCampo.Add (CPDataEmissaoAte.Name)
    objGridInt.colCampo.Add (CPDataEmissao.Name)
    objGridInt.colCampo.Add (CPDataVencimento.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_CPSelecionado_Col = 1
    iGrid_CondPagto_Col = 2
    iGrid_CPDataEmissaoDe_Col = 3
    iGrid_CPDataEmissaoAte_Col = 4
    iGrid_CPDataEmissao_Col = 5
    iGrid_CPDataVencimento_Col = 6
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridCondPagto

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9
    
    'Largura da primeira coluna
    GridCondPagto.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_CondPagto = SUCESSO

    Exit Function

Erro_Inicializa_Grid_CondPagto:

    Inicializa_Grid_CondPagto = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192101)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

    Exit Function

End Function

Sub Limpa_Tela_Faturamento()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Faturamento

    bDesabilitaCmdGridAux = False

    Call Limpa_Tela(Me)

    'Limpa os Grids da tela
    Call Grid_Limpa(objGridExcCliente)
    Call Grid_Limpa(objGridExcVou)

    Call Limpa_Tela_Faturamento_Aux
    
    Set gobjFaturamento = New ClassTRPFaturamento
    
    DescCliente.Caption = ""
    TotalVou.Caption = ""
    
    Call Default_Tela
    
    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_Selecao
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    Call TabStrip1_Click

    Exit Sub

Erro_Limpa_Tela_Faturamento:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192102)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_Faturamento_Aux()
    
    'Limpa os Grids da tela
    Call Grid_Limpa(objGridCliente)
    Call Grid_Limpa(objGridFatura)
    Call Grid_Limpa(objGridVoucher)
    Call Grid_Limpa(objGridFilialEmpresa)
    
    Call Ordenacao_Limpa(objGridCliente)
    Call Ordenacao_Limpa(objGridFatura)
    Call Ordenacao_Limpa(objGridVoucher)
    Call Ordenacao_Limpa(objGridFilialEmpresa)

End Sub

Private Sub BotaoLimpar_Click()
'Limpa a tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 192103

    'Limpa o restante da tela
    Call Limpa_Tela_Faturamento

    iAlterado = 0
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 192103
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192104)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAtualizar_Click()
'Atualiza a tela

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    lErro = CF("TRPFaturamento_Atualiza", gobjFaturamento)
    If lErro <> SUCESSO Then gError 192105
        
    Call Limpa_Tela_Faturamento_Aux

    lErro = Traz_Faturamento_Tela2(gobjFaturamento)
    If lErro <> SUCESSO Then gError 192106
    
    iTelaDesatualizada = DESMARCADO

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoAtualizar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192105, 192106
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192107)

    End Select

    Exit Sub

End Sub

Private Sub BotaoModeloFat_Click()

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    
    On Error GoTo Erro_BotaoModeloFat_Click
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|Html Files" & _
    "(*.html)|*.html"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

    ModeloFat.Text = CommonDialog1.FileName
    
    Exit Sub

Erro_BotaoModeloFat_Click:

    'User pressed the Cancel button
    Exit Sub
    
End Sub

Private Sub Cliente_Change()

    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
    
    Call Cliente_Preenche(Cliente)

End Sub

Private Sub FiliaisEmpresa_Click()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub optIndividual_Click()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub Cliente_Preenche(objControle As Object)

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
    
On Error GoTo Erro_Cliente_Preenche
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objControle, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 192108

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 192108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192109)

    End Select
    
    Exit Sub

End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    'Se o frame anterior foi o de Seleção e ele foi alterado
    If iFrameAtual <> TAB_Selecao And iFrameSelecaoAlterado = REGISTRO_ALTERADO Then

        DoEvents

        lErro = Traz_Faturamento_Tela
        If lErro <> SUCESSO Then gError 192110

        iFrameSelecaoAlterado = 0

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case 192110

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192111)

    End Select

    Exit Sub

End Sub

Private Sub TabStrip2_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip2)
End Sub

Private Sub TabStrip2_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip2_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index = iFrameAtualS Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtualS, TabStrip2, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    FrameS(TabStrip2.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    FrameS(iFrameAtualS).Visible = False
    'Armazena novo valor de iFrameAtualS
    iFrameAtualS = TabStrip2.SelectedItem.Index

    Exit Sub

Erro_TabStrip2_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192112)

    End Select

    Exit Sub

End Sub

Function Move_Selecao_Memoria(ByVal objFaturamento As ClassTRPFaturamento) As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim iLinha As Integer
Dim objCliente As ClassCliente
Dim objFatCondPagto  As ClassTRPFatCondPagto
Dim objCondPagto  As ClassCondicaoPagto
Dim objVoucher As ClassTRPVouchers
Dim iCount As Integer

On Error GoTo Erro_Move_Selecao_Memoria
   
    objFaturamento.lCliente = LCodigo_Extrai(Cliente.Text)
    
    If optIndividual.Value = vbChecked Then
        objFaturamento.iFaturarIndividualmente = MARCADO
    Else
        objFaturamento.iFaturarIndividualmente = DESMARCADO
    End If

    For iLinha = 1 To objGridExcCliente.iLinhasExistentes
    
        Set objCliente = New ClassCliente
    
        objCliente.lCodigo = LCodigo_Extrai(GridExcCliente.TextMatrix(iLinha, iGrid_ExcCliente_Col))
    
        objFaturamento.colExcClientes.Add objCliente
    
    Next
    
    For iLinha = 1 To objGridExcVou.iLinhasExistentes
    
        Set objVoucher = New ClassTRPVouchers
    
        objVoucher.sTipVou = GridExcVou.TextMatrix(iLinha, iGrid_ExcVouTipo_Col)
        objVoucher.sSerie = GridExcVou.TextMatrix(iLinha, iGrid_ExcVouSerie_Col)
        objVoucher.lNumVou = LCodigo_Extrai(GridExcVou.TextMatrix(iLinha, iGrid_ExcVouNum_Col))
    
        If Len(Trim(objVoucher.sTipVou)) = 0 Or Len(Trim(objVoucher.sSerie)) = 0 Or objVoucher.lNumVou = 0 Then gError 194427
    
        objFaturamento.colExcVouchers.Add objVoucher
    
    Next
    
    iCount = 0
    For iLinha = 0 To FiliaisEmpresa.ListCount - 1
        
        If Not FiliaisEmpresa.Selected(iLinha) Then
            objFaturamento.colFiliais.Add FiliaisEmpresa.ItemData(iLinha)
        Else
            iCount = iCount + 1
        End If
    
    Next
    
    If iCount = 0 Then gError 192215

    iCount = 0
    For iLinha = 0 To TipoDoc.ListCount - 1
        If Not TipoDoc.Selected(iLinha) Then
            objFaturamento.colTiposDoc.Add TipoDoc.List(iLinha)
        Else
            iCount = iCount + 1
        End If
    Next
    
    If iCount = 0 Then gError 192216
    
    iCount = 0
    For iLinha = 0 To TipoFaturamento.ListCount - 1
        If Not TipoFaturamento.Selected(iLinha) Then
            objFaturamento.colTiposFat.Add TipoFaturamento.List(iLinha)
        Else
            iCount = iCount + 1
        End If
    Next
    
    If iCount = 0 Then gError 192216
    
    For iLinha = 1 To objGridCondPagto.iLinhasExistentes
    
        If StrParaInt(GridCondPagto.TextMatrix(iLinha, iGrid_CPSelecionado_Col)) = MARCADO Then
    
            Set objFatCondPagto = New ClassTRPFatCondPagto
           
            For Each objCondPagto In gcolCondPagto
                If objCondPagto.iCodigo = Codigo_Extrai(GridCondPagto.TextMatrix(iLinha, iGrid_CondPagto_Col)) Then
                    Exit For
                End If
            Next
        
            Set objFatCondPagto.objCondPagtos = objCondPagto
        
            objFatCondPagto.dtDataEmissao = StrParaDate(GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissao_Col))
            objFatCondPagto.dtDataVencimento = StrParaDate(GridCondPagto.TextMatrix(iLinha, iGrid_CPDataVencimento_Col))
            objFatCondPagto.dtDataVouAte = StrParaDate(GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissaoAte_Col))
            objFatCondPagto.dtDataVouDe = StrParaDate(GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissaoDe_Col))
        
            objFaturamento.colCondPagtos.Add objFatCondPagto
            
        End If
    
    Next

    If objFaturamento.colCondPagtos.Count = 0 Then gError 192217

    Move_Selecao_Memoria = SUCESSO

    Exit Function

Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr

    Select Case gErr
    
        Case 192215
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_FILIAL_SELECIONADA", gErr)

        Case 192216
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_TIPODOC_SELECIONADA", gErr)

        Case 192217
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUMA_CONDPAGTO_SELECIONADA", gErr)
            
        Case 192218
        
        Case 194427
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_PREENCHIDO_GRID", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192113)

    End Select

    Exit Function

End Function

Function Carrega_Grid_CondPagto() As Long
'Recolhe dados do TAB de Seleção

Dim lErro As Long
Dim iLinha As Integer
Dim objCondPagto  As ClassCondicaoPagto
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome
Dim colCondPagto As New Collection
Dim objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_Carrega_Grid_CondPagto

    Call Grid_Limpa(objGridCondPagto)

    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
    lErro = CF("CondicoesPagto_Le_Recebimento", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 192115

    iLinha = 0
    For Each objCodigoDescricao In colCodigoDescricao
        iLinha = iLinha + 1
        
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        'Tem que calcular até quando deve pegar os vouchers emitidos
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
        
        Set objCondPagto = New ClassCondicaoPagto
        
        objCondPagto.iCodigo = objCodigoDescricao.iCodigo
        objCondPagto.sDescReduzida = objCodigoDescricao.sNome
        
        colCondPagto.Add objCondPagto

        lErro = CF("CondicaoPagto_Le", objCondPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 192116
        
        lErro = CF("CondicaoPagto_Le_Parcelas", objCondPagto)
        If lErro <> SUCESSO Then gError 192116
                
        objCondPagto.dValorTotal = 100
        objCondPagto.dtDataEmissao = gdtDataAtual
        objCondPagto.dtDataEntrega = gdtDataAtual
        objCondPagto.dtDataRef = gdtDataAtual
        
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondPagto, True, True, False)
        If lErro <> SUCESSO Then gError 192116
              
        GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissaoAte_Col) = Format(DateAdd("d", -1, gdtDataAtual), "dd/mm/yyyy")
        GridCondPagto.TextMatrix(iLinha, iGrid_CPSelecionado_Col) = CStr(MARCADO)
        GridCondPagto.TextMatrix(iLinha, iGrid_CondPagto_Col) = CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        
        If objCondPagto.colParcelas.Count = 0 Then
            GridCondPagto.TextMatrix(iLinha, iGrid_CPDataVencimento_Col) = Format(DateAdd("d", objCondPagto.iDiasParaPrimeiraParcela, gdtDataAtual), "dd/mm/yyyy")
        Else
            Set objParc = objCondPagto.colParcelas.Item(1)
            GridCondPagto.TextMatrix(iLinha, iGrid_CPDataVencimento_Col) = Format(objParc.dtVencimento, "dd/mm/yyyy")
        End If
        
        GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissao_Col) = Format(gdtDataAtual, "dd/mm/yyyy")

    Next
    
    objGridCondPagto.iLinhasExistentes = colCodigoDescricao.Count
    
    Call Grid_Refresh_Checkbox(objGridCondPagto)
    
    Set gcolCondPagto = colCondPagto

    Carrega_Grid_CondPagto = SUCESSO

    Exit Function

Erro_Carrega_Grid_CondPagto:

    Carrega_Grid_CondPagto = gErr

    Select Case gErr
    
        Case 192115, 192116

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192117)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Faturamento"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRPFaturamento"

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

Private Sub TipoDoc_Click()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
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
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Private Function Traz_Faturamento_Tela() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objFaturamento As New ClassTRPFaturamento
Dim objInfoFilial As ClassTRPFATInfoFilial
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoCliente As ClassTRPFATInfoCliente
Dim objFatCondPagto As ClassTRPFatCondPagto
Dim objCondPagto As ClassCondicaoPagto

On Error GoTo Erro_Traz_Faturamento_Tela

    GL_objMDIForm.MousePointer = vbHourglass

    Call Limpa_Tela_Faturamento_Aux
    
    lErro = Move_Selecao_Memoria(objFaturamento)
    If lErro <> SUCESSO Then gError 192118
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("TRPFaturamento_Le_Dados", objFaturamento)
    If lErro <> SUCESSO Then gError 192119
    
    If objFaturamento.colInfoVouchers.Count = 0 Then gError 192120
    
    Set gobjFaturamento = objFaturamento
        
    lErro = Traz_Faturamento_Tela2(objFaturamento)
    If lErro <> SUCESSO Then gError 192121
                
    GL_objMDIForm.MousePointer = vbDefault
                
    Traz_Faturamento_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Faturamento_Tela:

    GL_objMDIForm.MousePointer = vbDefault

    Traz_Faturamento_Tela = gErr
    
    Select Case gErr

        Case 192118, 192119, 192121
              
        Case 192120
            Call Rotina_Erro(vbOKOnly, "ERRO_SELECAO_NENHUM_VOUCHER", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192122)

    End Select

End Function

Private Function Traz_Faturamento_Tela2(ByVal objFaturamento As ClassTRPFaturamento) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iCol As Integer
Dim objInfoFilial As ClassTRPFATInfoFilial
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoCliente As ClassTRPFATInfoCliente
Dim objFatCondPagto As ClassTRPFatCondPagto
Dim objCondPagto As ClassCondicaoPagto
Dim vValor(2 To 6) As Variant
Dim sValor As String

On Error GoTo Erro_Traz_Faturamento_Tela2
    
    Call Ordenacao_Limpa(objGridFilialEmpresa)
    
    If objFaturamento.colInfoFiliais.Count >= objGridFilialEmpresa.objGrid.Rows Then
        Call Refaz_Grid(objGridFilialEmpresa, objFaturamento.colInfoFiliais.Count)
    End If

    iIndice = 0
    For Each objInfoFilial In objFaturamento.colInfoFiliais
    
        iIndice = iIndice + 1
    
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialEmpresa_Col) = objInfoFilial.iFilialEmpresa & SEPARADOR & objInfoFilial.sNome
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialNumFat_Col) = CStr(objInfoFilial.iNumFatS)
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialNumVou_Col) = CStr(objInfoFilial.iNumItens)
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialNumVouSel_Col) = CStr(objInfoFilial.iNumItensS)
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialValor_Col) = Format(objInfoFilial.dValor, "STANDARD")
        GridFilialEmpresa.TextMatrix(iIndice, iGrid_FilialValorSel_Col) = Format(objInfoFilial.dValorS, "STANDARD")
        
        vValor(iGrid_FilialNumFat_Col) = vValor(iGrid_FilialNumFat_Col) + objInfoFilial.iNumFatS
        vValor(iGrid_FilialNumVou_Col) = vValor(iGrid_FilialNumVou_Col) + objInfoFilial.iNumItens
        vValor(iGrid_FilialNumVouSel_Col) = vValor(iGrid_FilialNumVouSel_Col) + objInfoFilial.iNumItensS
        vValor(iGrid_FilialValor_Col) = vValor(iGrid_FilialValor_Col) + objInfoFilial.dValor
        vValor(iGrid_FilialValorSel_Col) = vValor(iGrid_FilialValorSel_Col) + objInfoFilial.dValorS
    
    Next
    
    bDesabilitaCmdGridAux = True
    GridFilialEmpresa.Col = iGrid_FilialEmpresa_Col
    GridFilialEmpresa.Row = iIndice + 1
    GridFilialEmpresa.CellFontBold = True
    GridFilialEmpresa.TextMatrix(iIndice + 1, iGrid_FilialEmpresa_Col) = "TOTAL"
    For iCol = 2 To 6
        If iCol <> iGrid_FilialValorSel_Col And iCol <> iGrid_FilialValor_Col Then
            sValor = vValor(iCol)
        Else
            sValor = Format(vValor(iCol), "STANDARD")
        End If
        GridFilialEmpresa.Col = iCol
        GridFilialEmpresa.Row = iIndice + 1
        GridFilialEmpresa.CellFontBold = True
        GridFilialEmpresa.TextMatrix(iIndice + 1, iCol) = sValor
    Next
    bDesabilitaCmdGridAux = False
    
    objGridFilialEmpresa.iLinhasExistentes = objFaturamento.colInfoFiliais.Count

    Call Ordenacao_Limpa(objGridCliente)

    If objFaturamento.colInfoClientes.Count >= objGridCliente.objGrid.Rows Then
        Call Refaz_Grid(objGridCliente, objFaturamento.colInfoClientes.Count)
    End If
    
    iIndice = 0
    For Each objInfoCliente In objFaturamento.colInfoClientes
    
        iIndice = iIndice + 1
    
        GridCliente.TextMatrix(iIndice, iGrid_CliFE_Col) = objInfoCliente.iFilialEmpresa
        GridCliente.TextMatrix(iIndice, iGrid_CliCliente_Col) = objInfoCliente.lCliente & SEPARADOR & objInfoCliente.sNome
        GridCliente.TextMatrix(iIndice, iGrid_CliNumVou_Col) = CStr(objInfoCliente.iNumItens)
        GridCliente.TextMatrix(iIndice, iGrid_CliNumVouSel_Col) = CStr(objInfoCliente.iNumItensS)
        GridCliente.TextMatrix(iIndice, iGrid_CliSelecionado_Col) = CStr(objInfoCliente.iMarcado)
        GridCliente.TextMatrix(iIndice, iGrid_CliValorFat_Col) = Format(objInfoCliente.dValor, "STANDARD")
        GridCliente.TextMatrix(iIndice, iGrid_CliValorFatSel_Col) = Format(objInfoCliente.dValorS, "STANDARD")
    
    Next
    
    objGridCliente.iLinhasExistentes = objFaturamento.colInfoClientes.Count
    
    Call Grid_Refresh_Checkbox(objGridCliente)
    
    Call Ordenacao_Limpa(objGridFatura)
    
    If objFaturamento.colInfoFaturas.Count >= objGridFatura.objGrid.Rows Then
        Call Refaz_Grid(objGridFatura, objFaturamento.colInfoFaturas.Count)
    End If
    
    iIndice = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas
    
        iIndice = iIndice + 1
        
        For Each objInfoCliente In objFaturamento.colInfoClientes
            If objInfoCliente.lCliente = objInfoFatura.lCliente Then
                Exit For
            End If
        Next
        
        For Each objFatCondPagto In objFaturamento.colCondPagtos
            If objFatCondPagto.objCondPagtos.iCodigo = objInfoFatura.iCondPagto Then
                Exit For
            End If
        Next
    
        GridFatura.TextMatrix(iIndice, iGrid_FatCliente_Col) = objInfoCliente.lCliente & SEPARADOR & objInfoCliente.sNome
        GridFatura.TextMatrix(iIndice, iGrid_FatCondPagto_Col) = objFatCondPagto.objCondPagtos.iCodigo & SEPARADOR & objFatCondPagto.objCondPagtos.sDescReduzida
        GridFatura.TextMatrix(iIndice, iGrid_FatDataVenc_Col) = Format(objInfoFatura.dtDataVencimento, "dd/mm/yyyy")
        GridFatura.TextMatrix(iIndice, iGrid_FatEmissao_Col) = Format(objInfoFatura.dtDataEmissao, "dd/mm/yyyy")
        GridFatura.TextMatrix(iIndice, iGrid_FatSelecionado_Col) = CStr(objInfoFatura.iMarcado)
        GridFatura.TextMatrix(iIndice, iGrid_FatSeq_Col) = CStr(objInfoFatura.lFatura)
        GridFatura.TextMatrix(iIndice, iGrid_FatValorB_Col) = Format(objInfoFatura.dValor + objInfoFatura.dValorAporte + objInfoFatura.dValorAporteCred, "STANDARD")
        GridFatura.TextMatrix(iIndice, iGrid_FatValorDesc_Col) = Format(objInfoFatura.dValorAporte + objInfoFatura.dValorAporteCred, "STANDARD")
        GridFatura.TextMatrix(iIndice, iGrid_FatValor_Col) = Format(objInfoFatura.dValor, "STANDARD")
    
    Next
    
    objGridFatura.iLinhasExistentes = objFaturamento.colInfoFaturas.Count
    
    Call Grid_Refresh_Checkbox(objGridFatura)
    
    Call Ordenacao_Limpa(objGridVoucher)
    
    If objFaturamento.colInfoVouchers.Count >= objGridVoucher.objGrid.Rows Then
        Call Refaz_Grid(objGridVoucher, objFaturamento.colInfoVouchers.Count)
    End If
    
    iIndice = 0
    For Each objInfoVoucher In objFaturamento.colInfoVouchers
    
        iIndice = iIndice + 1
        
        For Each objInfoCliente In objFaturamento.colInfoClientes
            If objInfoCliente.lCliente = objInfoVoucher.lCliente Then
                Exit For
            End If
        Next
    
        GridVoucher.TextMatrix(iIndice, iGrid_VouCliente_Col) = objInfoCliente.lCliente & SEPARADOR & objInfoCliente.sNome
        GridVoucher.TextMatrix(iIndice, iGrid_VouDataEmissao_Col) = Format(objInfoVoucher.dtDataEmissao, "dd/mm/yyyy")
        GridVoucher.TextMatrix(iIndice, iGrid_VouFatSeq_Col) = CStr(objInfoVoucher.lFatura)
        GridVoucher.TextMatrix(iIndice, iGrid_VouNumero_Col) = CStr(objInfoVoucher.lNumVou)
        GridVoucher.TextMatrix(iIndice, iGrid_VouSelecionado_Col) = CStr(objInfoVoucher.iMarcado)
        GridVoucher.TextMatrix(iIndice, iGrid_VouSerie_Col) = objInfoVoucher.sSerie
        GridVoucher.TextMatrix(iIndice, iGrid_VouTipo_Col) = objInfoVoucher.sTipoDoc
        GridVoucher.TextMatrix(iIndice, iGrid_VouTipoV_Col) = objInfoVoucher.sTipoVou
        GridVoucher.TextMatrix(iIndice, iGrid_VouValor_Col) = Format(objInfoVoucher.dValorBruto, "STANDARD")
        GridVoucher.TextMatrix(iIndice, iGrid_VouValorC_Col) = Format(objInfoVoucher.dValorComissao, "STANDARD")
        GridVoucher.TextMatrix(iIndice, iGrid_VouValorL_Col) = Format(objInfoVoucher.dValorBruto - objInfoVoucher.dValorComissao, "STANDARD")
        GridVoucher.TextMatrix(iIndice, iGrid_VouValorA_Col) = Format(objInfoVoucher.dValorAporte, "#,##0.00##")
    
    Next
    
    objGridVoucher.iLinhasExistentes = objFaturamento.colInfoVouchers.Count
    
    Call Grid_Refresh_Checkbox(objGridVoucher)
    
    Call Calcula_TotalVou
            
    Traz_Faturamento_Tela2 = SUCESSO
    
    Exit Function
    
Erro_Traz_Faturamento_Tela2:

    bDesabilitaCmdGridAux = False

    Traz_Faturamento_Tela2 = gErr
    
    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192123)

    End Select

End Function

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_FilialEmpresa

    FiliaisEmpresa.Clear

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 192124

    iIndice = 0
    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
    
        'If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            FiliaisEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FiliaisEmpresa.ItemData(FiliaisEmpresa.NewIndex) = objCodigoNome.iCodigo
            FiliaisEmpresa.Selected(iIndice) = True
        
            iIndice = iIndice + 1
        'End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 192124

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192125)

    End Select

    Exit Function


End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    
    objGridInt.objGrid.Rows = iNumLinhas + 1
    
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
    
End Sub

Private Sub GridCliente_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer
Dim objObjeto As Object

    Call Grid_Click(objGridCliente, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCliente, iAlterado)
    End If
    
    colcolColecao.Add gobjFaturamento.colInfoClientes
    
    Call Ordenacao_ClickGrid(objGridCliente, , colcolColecao)
    
    iLinha = 0
    For Each objObjeto In gobjFaturamento.colInfoClientes
        iLinha = iLinha + 1
        objObjeto.iLinha = iLinha
    Next

End Sub

Private Sub GridCliente_GotFocus()
    Call Grid_Recebe_Foco(objGridCliente)
End Sub

Private Sub GridCliente_EnterCell()
    Call Grid_Entrada_Celula(objGridCliente, iAlterado)
End Sub

Private Sub GridCliente_LeaveCell()
    Call Saida_Celula(objGridCliente)
End Sub

Private Sub GridCliente_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCliente, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCliente, iAlterado)
    End If

End Sub

Private Sub GridCliente_RowColChange()
    Call Grid_RowColChange(objGridCliente)
End Sub

Private Sub GridCliente_Scroll()
    Call Grid_Scroll(objGridCliente)
End Sub

Private Sub GridCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCliente)

End Sub

Private Sub GridCliente_LostFocus()
    Call Grid_Libera_Foco(objGridCliente)
End Sub


Private Sub GridCondPagto_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCondPagto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCondPagto, iAlterado)
    End If
   
    Call Ordenacao_ClickGrid(objGridCondPagto)

End Sub

Private Sub GridCondPagto_GotFocus()
    Call Grid_Recebe_Foco(objGridCondPagto)
End Sub

Private Sub GridCondPagto_EnterCell()
    Call Grid_Entrada_Celula(objGridCondPagto, iAlterado)
End Sub

Private Sub GridCondPagto_LeaveCell()
    Call Saida_Celula(objGridCondPagto)
End Sub

Private Sub GridCondPagto_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCondPagto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCondPagto, iAlterado)
    End If

End Sub

Private Sub GridCondPagto_RowColChange()
    Call Grid_RowColChange(objGridCondPagto)
End Sub

Private Sub GridCondPagto_Scroll()
    Call Grid_Scroll(objGridCondPagto)
End Sub

Private Sub GridCondPagto_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridCondPagto)

End Sub

Private Sub GridCondPagto_LostFocus()
    Call Grid_Libera_Foco(objGridCondPagto)
End Sub


Private Sub GridExcCliente_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridExcCliente, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcCliente, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridExcCliente)

End Sub

Private Sub GridExcCliente_GotFocus()
    Call Grid_Recebe_Foco(objGridExcCliente)
End Sub

Private Sub GridExcCliente_EnterCell()
    Call Grid_Entrada_Celula(objGridExcCliente, iAlterado)
End Sub

Private Sub GridExcCliente_LeaveCell()
    Call Saida_Celula(objGridExcCliente)
End Sub

Private Sub GridExcCliente_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridExcCliente, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcCliente, iAlterado)
    End If

End Sub

Private Sub GridExcCliente_RowColChange()
    Call Grid_RowColChange(objGridExcCliente)
End Sub

Private Sub GridExcCliente_Scroll()
    Call Grid_Scroll(objGridExcCliente)
End Sub

Private Sub GridExcCliente_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridExcCliente)

End Sub

Private Sub GridExcCliente_LostFocus()
    Call Grid_Libera_Foco(objGridExcCliente)
End Sub


Private Sub GridFatura_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer
Dim objObjeto As Object

    Call Grid_Click(objGridFatura, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFatura, iAlterado)
    End If
    
    colcolColecao.Add gobjFaturamento.colInfoFaturas
    
    Call Ordenacao_ClickGrid(objGridFatura, , colcolColecao)
    
    iLinha = 0
    For Each objObjeto In gobjFaturamento.colInfoFaturas
        iLinha = iLinha + 1
        objObjeto.iLinha = iLinha
    Next

End Sub

Private Sub GridFatura_GotFocus()
    Call Grid_Recebe_Foco(objGridFatura)
End Sub

Private Sub GridFatura_EnterCell()
    Call Grid_Entrada_Celula(objGridFatura, iAlterado)
End Sub

Private Sub GridFatura_LeaveCell()
    Call Saida_Celula(objGridFatura)
End Sub

Private Sub GridFatura_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFatura, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFatura, iAlterado)
    End If

End Sub

Private Sub GridFatura_RowColChange()
    Call Grid_RowColChange(objGridFatura)
End Sub

Private Sub GridFatura_Scroll()
    Call Grid_Scroll(objGridFatura)
End Sub

Private Sub GridFatura_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFatura)

End Sub

Private Sub GridFatura_LostFocus()
    Call Grid_Libera_Foco(objGridFatura)
End Sub


Private Sub GridFilialEmpresa_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection

    Call Grid_Click(objGridFilialEmpresa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilialEmpresa, iAlterado)
    End If
    
    colcolColecao.Add gobjFaturamento.colInfoFiliais
    
    Call Ordenacao_ClickGrid(objGridFilialEmpresa, , colcolColecao)

End Sub

Private Sub GridFilialEmpresa_GotFocus()
    Call Grid_Recebe_Foco(objGridFilialEmpresa)
End Sub

Private Sub GridFilialEmpresa_EnterCell()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_Entrada_Celula(objGridFilialEmpresa, iAlterado)
    End If
End Sub

Private Sub GridFilialEmpresa_LeaveCell()
    If Not bDesabilitaCmdGridAux Then
        Call Saida_Celula(objGridFilialEmpresa)
    End If
End Sub

Private Sub GridFilialEmpresa_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridFilialEmpresa, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridFilialEmpresa, iAlterado)
    End If

End Sub

Private Sub GridFilialEmpresa_RowColChange()
    If Not bDesabilitaCmdGridAux Then
        Call Grid_RowColChange(objGridFilialEmpresa)
    End If
End Sub

Private Sub GridFilialEmpresa_Scroll()
    Call Grid_Scroll(objGridFilialEmpresa)
End Sub

Private Sub GridFilialEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridFilialEmpresa)

End Sub

Private Sub GridFilialEmpresa_LostFocus()
    Call Grid_Libera_Foco(objGridFilialEmpresa)
End Sub

Private Sub GridVoucher_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection
Dim iLinha As Integer
Dim objObjeto As Object

    Call Grid_Click(objGridVoucher, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVoucher, iAlterado)
    End If
    
    colcolColecao.Add gobjFaturamento.colInfoVouchers
    
    Call Ordenacao_ClickGrid(objGridVoucher, , colcolColecao)
    
    iLinha = 0
    For Each objObjeto In gobjFaturamento.colInfoVouchers
        iLinha = iLinha + 1
        objObjeto.iLinha = iLinha
    Next

End Sub

Private Sub GridVoucher_GotFocus()
    Call Grid_Recebe_Foco(objGridVoucher)
End Sub

Private Sub GridVoucher_EnterCell()
    Call Grid_Entrada_Celula(objGridVoucher, iAlterado)
End Sub

Private Sub GridVoucher_LeaveCell()
    Call Saida_Celula(objGridVoucher)
End Sub

Private Sub GridVoucher_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridVoucher, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVoucher, iAlterado)
    End If

End Sub

Private Sub GridVoucher_RowColChange()
    Call Grid_RowColChange(objGridVoucher)
End Sub

Private Sub GridVoucher_Scroll()
    Call Grid_Scroll(objGridVoucher)
End Sub

Private Sub GridVoucher_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridVoucher)

End Sub

Private Sub GridVoucher_LostFocus()
    Call Grid_Libera_Foco(objGridVoucher)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridExcCliente.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_ExcCliente_Col
                
                    lErro = Saida_Celula_Cliente(objGridInt, ExcCliente)
                    If lErro <> SUCESSO Then gError 192126
                     
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridExcVou.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_ExcVouTipo_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, ExcVouTipo)
                    If lErro <> SUCESSO Then gError 192126
                     
                Case iGrid_ExcVouSerie_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, ExcVouSerie)
                    If lErro <> SUCESSO Then gError 192126
                    
                Case iGrid_ExcVouNum_Col
                
                    lErro = Saida_Celula_ExcVouNum(objGridInt)
                    If lErro <> SUCESSO Then gError 192126
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridCondPagto.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CPSelecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, CPSelecionado)
                    If lErro <> SUCESSO Then gError 192127
                
                Case iGrid_CPDataEmissao_Col

                    lErro = Saida_Celula_Data(objGridInt, CPDataEmissao)
                    If lErro <> SUCESSO Then gError 192128

                Case iGrid_CPDataEmissaoDe_Col
                
                    lErro = Saida_Celula_Data(objGridInt, CPDataEmissaoDe)
                    If lErro <> SUCESSO Then gError 192129
                    
                Case iGrid_CPDataEmissaoAte_Col
                
                    lErro = Saida_Celula_Data(objGridInt, CPDataEmissaoAte)
                    If lErro <> SUCESSO Then gError 192130
                    
                Case iGrid_CPDataVencimento_Col
                
                    lErro = Saida_Celula_Data(objGridInt, CPDataVencimento)
                    If lErro <> SUCESSO Then gError 192131
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridCliente.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_CliSelecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, CliSelecionado)
                    If lErro <> SUCESSO Then gError 192132
                    
            End Select
            
        ElseIf objGridInt.objGrid.Name = GridFatura.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_FatSelecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, FatSelecionado)
                    If lErro <> SUCESSO Then gError 192133
                
                Case iGrid_FatEmissao_Col

                    lErro = Saida_Celula_Data(objGridInt, FatEmissao)
                    If lErro <> SUCESSO Then gError 192134
                    
                Case iGrid_FatDataVenc_Col
                
                    lErro = Saida_Celula_Data(objGridInt, FatDataVenc)
                    If lErro <> SUCESSO Then gError 192135
                    
            End Select

            
        ElseIf objGridInt.objGrid.Name = GridVoucher.Name Then

            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_VouSelecionado_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, VouSelecionado)
                    If lErro <> SUCESSO Then gError 192136
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 192137

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 192126 To 192136

        Case 192137
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192138)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long

On Error GoTo Erro_Rotina_Grid_Enable
              
    Select Case objControl.Name
    
        'Grid ExcCliente
        Case ExcCliente.Name
            objControl.Enabled = True
        
        'Grid CondPagto
        Case CPSelecionado.Name, CPDataEmissaoDe.Name, CPDataEmissaoAte.Name, CPDataEmissao.Name, CPDataVencimento.Name
            objControl.Enabled = True

        'Grid Cliente
        Case CliSelecionado.Name
            objControl.Enabled = True
        
        'Grid Fatura
        Case FatSelecionado.Name, FatEmissao.Name, FatDataVenc.Name
            objControl.Enabled = True
        
        'Grid Voucher
        Case VouSelecionado.Name
            objControl.Enabled = True
            
        Case ExcVouTipo.Name, ExcVouSerie.Name
            If Len(Trim(GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouNum_Col))) = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case ExcVouNum.Name
            If Len(Trim(GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouTipo_Col))) <> 0 And Len(Trim(GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouSerie_Col))) <> 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        
        Case Else
            objControl.Enabled = False
            
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192139)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192140

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 192140
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192141)

    End Select

    Exit Function

End Function

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 192142
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192143

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 192142 To 192143
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192144)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Function Saida_Celula_Cliente(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Saida_Celula_Cliente

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        lErro = TP_Cliente_Le2(objControle, objCliente, 0)
        If lErro <> SUCESSO Then gError 192145
        
        'verifica se precisa preencher o grid com uma nova linha
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192146

    Saida_Celula_Cliente = SUCESSO

    Exit Function

Erro_Saida_Celula_Cliente:

    Saida_Celula_Cliente = gErr

    Select Case gErr

        Case 192145 To 192146
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192147)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub ExcCliente_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcCliente_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcCliente)
End Sub

Public Sub ExcCliente_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcCliente)
End Sub

Public Sub ExcCliente_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridExcCliente.objControle = ExcCliente
    lErro = Grid_Campo_Libera_Foco(objGridExcCliente)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CPSelecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CPSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCondPagto)
End Sub

Public Sub CPSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCondPagto)
End Sub

Public Sub CPSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCondPagto.objControle = CPSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridCondPagto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CPDataEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CPDataEmissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCondPagto)
End Sub

Public Sub CPDataEmissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCondPagto)
End Sub

Public Sub CPDataEmissao_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCondPagto.objControle = CPDataEmissao
    lErro = Grid_Campo_Libera_Foco(objGridCondPagto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CPDataVencimento_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CPDataVencimento_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCondPagto)
End Sub

Public Sub CPDataVencimento_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCondPagto)
End Sub

Public Sub CPDataVencimento_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCondPagto.objControle = CPDataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridCondPagto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CPDataEmissaoDe_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CPDataEmissaoDe_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCondPagto)
End Sub

Public Sub CPDataEmissaoDe_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCondPagto)
End Sub

Public Sub CPDataEmissaoDe_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCondPagto.objControle = CPDataEmissaoDe
    lErro = Grid_Campo_Libera_Foco(objGridCondPagto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CPDataEmissaoAte_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CPDataEmissaoAte_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCondPagto)
End Sub

Public Sub CPDataEmissaoAte_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCondPagto)
End Sub

Public Sub CPDataEmissaoAte_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCondPagto.objControle = CPDataEmissaoAte
    lErro = Grid_Campo_Libera_Foco(objGridCondPagto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CliSelecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    iTelaDesatualizada = MARCADO
    Call Marca_Desmarca_Cliente(StrParaInt(GridCliente.TextMatrix(GridCliente.Row, iGrid_CliSelecionado_Col)), GridCliente.Row)
End Sub

Public Sub CliSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridCliente)
End Sub

Public Sub CliSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridCliente)
End Sub

Public Sub CliSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridCliente.objControle = CliSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridCliente)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FatSelecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    iTelaDesatualizada = MARCADO
    Call Marca_Desmarca_Fatura(StrParaInt(GridFatura.TextMatrix(GridFatura.Row, iGrid_FatSelecionado_Col)), GridFatura.Row)
End Sub

Public Sub FatSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFatura)
End Sub

Public Sub FatSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFatura)
End Sub

Public Sub FatSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridFatura.objControle = FatSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridFatura)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FatEmissao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FatEmissao_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFatura)
End Sub

Public Sub FatEmissao_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFatura)
End Sub

Public Sub FatEmissao_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridFatura.objControle = FatEmissao
    lErro = Grid_Campo_Libera_Foco(objGridFatura)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FatDataVenc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FatDataVenc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridFatura)
End Sub

Public Sub FatDataVenc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridFatura)
End Sub

Public Sub FatDataVenc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridFatura.objControle = FatDataVenc
    lErro = Grid_Campo_Libera_Foco(objGridFatura)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub VouSelecionado_Click()
    iAlterado = REGISTRO_ALTERADO
    iTelaDesatualizada = MARCADO
    Call Marca_Desmarca_Voucher(StrParaInt(GridVoucher.TextMatrix(GridVoucher.Row, iGrid_VouSelecionado_Col)), GridVoucher.Row)
End Sub

Public Sub VouSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridVoucher)
End Sub

Public Sub VouSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridVoucher)
End Sub

Public Sub VouSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridVoucher.objControle = VouSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridVoucher)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Cliente_Validate

    DescCliente.Caption = ""

    If Len(Trim(Cliente.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(Cliente, objCliente, 0)
        If lErro <> SUCESSO Then gError 192148
        
        'Cliente.Text = objCliente.sNomeReduzido
        DescCliente.Caption = objCliente.sRazaoSocial

    End If
        
    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 192148
            'erro tratado na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192149)

    End Select

End Sub

'Private Sub Marca_Desmarca(ByVal objGrid As AdmGrid, ByVal iColuna As Integer, ByVal iMarcado As Integer, Optional ByVal iTabIndex As Integer = 0)
''Marca todos os bloqueios do Grid
'
'Dim iLinha As Integer
'
'    'Percorre todas as linhas do Grid
'    For iLinha = 1 To objGrid.iLinhasExistentes
'
'        objGrid.objGrid.TextMatrix(iLinha, iColuna) = CStr(iMarcado)
'
'        Select Case iTabIndex
'            Case TAB_VOUCHER
'                Call Marca_Desmarca_Voucher(iMarcado, iLinha)
'
'            Case TAB_CLIENTE
'                Call Marca_Desmarca_Cliente(iMarcado, iLinha)
'
'            Case TAB_FATURA
'                Call Marca_Desmarca_Fatura(iMarcado, iLinha)
'        End Select
'
'    Next
'
'    'Atualiza na tela os checkbox marcados
'    Call Grid_Refresh_Checkbox(objGrid)
'
'End Sub

Private Sub Marca_Desmarca(ByVal iMarcado As Integer)
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objObjeto As Object

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridCliente.iLinhasExistentes
        GridCliente.TextMatrix(iLinha, iGrid_CliSelecionado_Col) = CStr(iMarcado)
    Next
    For iLinha = 1 To objGridVoucher.iLinhasExistentes
        GridVoucher.TextMatrix(iLinha, iGrid_VouSelecionado_Col) = CStr(iMarcado)
    Next
    For iLinha = 1 To objGridFatura.iLinhasExistentes
        GridFatura.TextMatrix(iLinha, iGrid_FatSelecionado_Col) = CStr(iMarcado)
    Next
    
    For Each objObjeto In gobjFaturamento.colInfoClientes
        objObjeto.iMarcado = iMarcado
    Next
    For Each objObjeto In gobjFaturamento.colInfoFaturas
        objObjeto.iMarcado = iMarcado
    Next
    For Each objObjeto In gobjFaturamento.colInfoVouchers
        objObjeto.iMarcado = iMarcado
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridCliente)
    Call Grid_Refresh_Checkbox(objGridVoucher)
    Call Grid_Refresh_Checkbox(objGridFatura)
    
    Call Calcula_TotalVou
    
End Sub


Private Sub BotaoMarcarTodos_Click(Index As Integer)

Dim iLinha As Integer

On Error GoTo Erro_BotaoMarcarTodos_Click

    GL_objMDIForm.MousePointer = vbHourglass

    Select Case Index

        Case TAB_Selecao
            
            For iLinha = 0 To FiliaisEmpresa.ListCount - 1
                FiliaisEmpresa.Selected(iLinha) = True
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO
            
        Case TAB_VOUCHER, TAB_CLIENTE, TAB_FATURA
            Call Marca_Desmarca(MARCADO)
            
        Case 11
            
            For iLinha = 0 To TipoDoc.ListCount - 1
                TipoDoc.Selected(iLinha) = True
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO
            
        Case 12
            
            For iLinha = 0 To TipoFaturamento.ListCount - 1
                TipoFaturamento.Selected(iLinha) = True
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO

'        Case TAB_VOUCHER
'
'            Call Marca_Desmarca(objGridVoucher, iGrid_VouSelecionado_Col, MARCADO, Index)
'
'        Case TAB_CLIENTE
'
'            Call Marca_Desmarca(objGridCliente, iGrid_CliSelecionado_Col, MARCADO, Index)
'
'        Case TAB_FATURA
'
'            Call Marca_Desmarca(objGridFatura, iGrid_FatSelecionado_Col, MARCADO, Index)

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoMarcarTodos_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192150)

    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)

Dim iLinha As Integer

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    GL_objMDIForm.MousePointer = vbHourglass

    Select Case Index

        Case TAB_Selecao
            
            For iLinha = 0 To FiliaisEmpresa.ListCount - 1
                FiliaisEmpresa.Selected(iLinha) = False
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO
            
        Case TAB_VOUCHER, TAB_CLIENTE, TAB_FATURA
            Call Marca_Desmarca(DESMARCADO)
            
        Case 11
            
            For iLinha = 0 To TipoDoc.ListCount - 1
                TipoDoc.Selected(iLinha) = False
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO
            
        Case 12
            
            For iLinha = 0 To TipoFaturamento.ListCount - 1
                TipoFaturamento.Selected(iLinha) = False
            Next
            
            iFrameSelecaoAlterado = REGISTRO_ALTERADO

'        Case TAB_VOUCHER
'
'            Call Marca_Desmarca(objGridVoucher, iGrid_VouSelecionado_Col, DESMARCADO, Index)
'
'        Case TAB_CLIENTE
'
'            Call Marca_Desmarca(objGridCliente, iGrid_CliSelecionado_Col, DESMARCADO, Index)
'
'        Case TAB_FATURA
'
'            Call Marca_Desmarca(objGridFatura, iGrid_FatSelecionado_Col, DESMARCADO, Index)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192151)

    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objCliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoExcCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    If Me.ActiveControl Is ExcCliente Then
        ExcCliente.Text = objCliente.sNomeReduzido
    Else
        GridExcCliente.TextMatrix(GridExcCliente.Row, iGrid_ExcCliente_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub BotaoClientes_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoClientes_Click

    If GridExcCliente.Row = 0 Then gError 192152

    If Me.ActiveControl Is ExcCliente Then
        objCliente.sNomeReduzido = ExcCliente.Text
    Else
        objCliente.sNomeReduzido = GridExcCliente.TextMatrix(GridExcCliente.Row, iGrid_ExcCliente_Col)
    End If

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoExcCliente)

    Exit Sub
    
Erro_BotaoClientes_Click:
    
    Select Case gErr
    
        Case 192152
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192153)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoItemFat_Click(Index As Integer)

Dim sFiltro As String
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoItemFat_Click

    Select Case Index

        Case TAB_FilialEmpresa
        
            If GridFilialEmpresa.Row = 0 Then gError 192154
        
            sFiltro = "FilialEmpresa = ?"
            colSelecao.Add Codigo_Extrai(GridFilialEmpresa.TextMatrix(GridFilialEmpresa.Row, iGrid_FilialEmpresa_Col))
        

        Case TAB_CLIENTE
            
            If GridCliente.Row = 0 Then gError 192155
            
            sFiltro = "Cliente = ?"
            colSelecao.Add LCodigo_Extrai(GridCliente.TextMatrix(GridCliente.Row, iGrid_CliCliente_Col))

        Case Else


    End Select
    
    Call Chama_Tela("TRPDocsParaFatLista", colSelecao, Nothing, Nothing, sFiltro)

    Exit Sub
    
Erro_BotaoItemFat_Click:
    
    Select Case gErr
    
        Case 192154, 192155
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192156)
    
    End Select
    
    Exit Sub
    
End Sub

Private Function Carrega_Combo_Modelo() As Long

Dim lErro As Long
Dim colCobranca As New Collection
Dim objCobranca As ClassCobrancaEmailPadrao

On Error GoTo Erro_Carrega_Combo_Modelo

    Modelo.Clear

    'Le os modelos válidos para o atraso em questão
    lErro = CF("CobrancaEmailPadrao_Le_ComAtraso", colCobranca, -1)
    If lErro <> SUCESSO Then gError 192157
        
    'Carrega a Combo com os Dados da Colecao
    For Each objCobranca In colCobranca
        Modelo.AddItem objCobranca.lCodigo & SEPARADOR & objCobranca.sDescricao
        Modelo.ItemData(Modelo.NewIndex) = objCobranca.lCodigo
    Next
    
    Carrega_Combo_Modelo = SUCESSO

    Exit Function
    
Erro_Carrega_Combo_Modelo:

    Carrega_Combo_Modelo = gErr
    
    Select Case gErr
    
        Case 192157

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192158)
    
    End Select
    
    Exit Function
    
End Function

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192159)

    End Select

    Exit Sub
  
End Sub

Private Sub NomeDiretorio_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iPOS As Integer

On Error GoTo Erro_NomeDiretorio_Validate

    If Len(Trim(NomeDiretorio.Text)) = 0 Then Exit Sub
    
    If right(NomeDiretorio.Text, 1) <> "\" And right(NomeDiretorio.Text, 1) <> "/" Then
        iPOS = InStr(1, NomeDiretorio.Text, "/")
        If iPOS = 0 Then
            NomeDiretorio.Text = NomeDiretorio.Text & "\"
        Else
            NomeDiretorio.Text = NomeDiretorio.Text & "/"
        End If
    End If

    If Len(Trim(Dir(NomeDiretorio.Text, vbDirectory))) = 0 Then gError 192160

    Exit Sub

Erro_NomeDiretorio_Validate:

    Cancel = True

    Select Case gErr

        Case 192160, 76
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_INVALIDO", gErr, NomeDiretorio.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192161)

    End Select

    Exit Sub

End Sub

Private Sub OptSoGerar_Click()

    If OptSoGerar.Value Then
        'LabelModelo.ForeColor = COR_CAMPO_NAO_OBRIGATORIO
        Modelo.Enabled = False
        Modelo.ListIndex = -1
    Else
        Modelo.Enabled = True
        'LabelModelo.ForeColor = COR_CAMPO_OBRIGATORIO
    End If

End Sub

Private Sub OptGerarEnviar_Click()

    If OptSoGerar.Value Then
        'LabelModelo.ForeColor = COR_CAMPO_NAO_OBRIGATORIO
        Modelo.Enabled = False
        Modelo.ListIndex = -1
    Else
        'LabelModelo.ForeColor = COR_CAMPO_OBRIGATORIO
        Modelo.Enabled = True
    End If

End Sub

Sub Default_Tela()

Dim lErro As Long
Dim sConteudo As String
Dim iIndice As Integer

On Error GoTo Erro_Default_Tela

    lErro = Carrega_Grid_CondPagto
    If lErro <> SUCESSO Then gError 192092

    Modelo.ListIndex = -1
    
    OptSoGerar.Value = True
    Modelo.Enabled = True
    'LabelModelo.ForeColor = COR_CAMPO_OBRIGATORIO
    
    DataEmiPadrao.PromptInclude = False
    DataEmiPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmiPadrao.PromptInclude = True
    
    DataEmiVouPadrao.PromptInclude = False
    DataEmiVouPadrao.Text = Format(DateAdd("d", -1, gdtDataAtual), "dd/mm/yy")
    DataEmiVouPadrao.PromptInclude = True
    
    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192192
    
    NomeDiretorio.Text = sConteudo
    Call NomeDiretorio_Validate(bSGECancelDummy)
    
    lErro = CF("TRPConfig_Le", TRPCONFIG_DIRETORIO_MODELO_FAT_HTML, EMPRESA_TODA, sConteudo)
    If lErro <> SUCESSO Then gError 192192
    
    ModeloFat.Text = sConteudo
    
    For iIndice = 0 To FiliaisEmpresa.ListCount - 1
        FiliaisEmpresa.Selected(iIndice) = True
    Next
    
    For iIndice = 0 To TipoDoc.ListCount - 1
        TipoDoc.Selected(iIndice) = True
    Next
    
    For iIndice = 0 To TipoFaturamento.ListCount - 1
        TipoFaturamento.Selected(iIndice) = True
    Next
    
    optIndividual.Value = vbUnchecked
    
    AbrirFatHtml.Value = vbUnchecked
    
    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Default_Tela:

    Select Case gErr
    
        Case 192092, 192192

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192162)

    End Select

    Exit Sub

End Sub

Sub BotaoGerar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGerar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 192163

    'Limpa Tela
    Call Limpa_Tela_Faturamento

    Exit Sub

Erro_BotaoGerar_Click:

    Select Case gErr

        Case 192163

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192164)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colArqHtml As New Collection
Dim vValor As Variant

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(NomeDiretorio.Text)) = 0 Then gError 192165
    If Len(Trim(ModeloFat.Text)) = 0 Then gError 192201
    
'    If OptGerarEnviar.Value Then
'        If LCodigo_Extrai(Modelo.Text) = 0 Then gError 192166
'    End If

    'Preenche o objFaturamento
    lErro = Move_Tela_Memoria(gobjFaturamento)
    If lErro <> SUCESSO Then gError 192167

    lErro = CF("TRPFaturamento_Gera", gobjFaturamento, colArqHtml)
    If lErro <> SUCESSO Then gError 192168
    
    If AbrirFatHtml.Value = vbChecked Then
    
        For Each vValor In colArqHtml
            Call Shell("explorer.exe " & CStr(vValor), vbMaximizedFocus)
        Next
        
    End If

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 192165
            Call Rotina_Erro(vbOKOnly, "ERRO_DIRETORIO_NAO_PREENCHIDO", gErr)
            NomeDiretorio.SetFocus
        
        Case 192166
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
            Modelo.SetFocus
    
        Case 192167, 192168
        
        Case 192201
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELO_NAO_PREENCHIDO", gErr)
            ModeloFat.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192169)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objFaturamento As ClassTRPFaturamento) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoFatura As ClassTRPFATInfoFatura

On Error GoTo Erro_Move_Tela_Memoria

    objFaturamento.sDiretorio = NomeDiretorio.Text
    objFaturamento.sModelo = ModeloFat.Text
    
    If OptGerarEnviar Then
        objFaturamento.iEnviarPorEmail = MARCADO
    Else
        objFaturamento.iEnviarPorEmail = DESMARCADO
    End If

    objFaturamento.lModelo = LCodigo_Extrai(Modelo)
    
    If iTelaDesatualizada = MARCADO Then gError 192170 'Tem que atualizar as informações
    
    iLinha = 0
    For Each objInfoFatura In objFaturamento.colInfoFaturas

        iLinha = iLinha + 1
        
        If objInfoFatura.iMarcado = MARCADO Then

            objInfoFatura.dtDataEmissao = StrParaDate(GridFatura.TextMatrix(iLinha, iGrid_FatEmissao_Col))
            objInfoFatura.dtDataVencimento = StrParaDate(GridFatura.TextMatrix(iLinha, iGrid_FatDataVenc_Col))
    
            If objInfoFatura.dtDataEmissao = DATA_NULA Then gError 192171
            If objInfoFatura.dtDataVencimento = DATA_NULA Then gError 192172

        End If

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 192170
            Call Rotina_Erro(vbOKOnly, "ERRO_TELA_DESATUALIZADA", gErr)
        
        Case 192171
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_NAO_PREENCHIDA_GRID", gErr, iLinha)
        
        Case 192172
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VENC_NAO_PREENCHIDA_GRID", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192173)

    End Select

End Function

Private Sub BotaoEmiAplicar_Click()

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_BotaoEmiAplicar_Click

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    If StrParaDate(DataEmiPadrao.Text) <> DATA_NULA Then
    
        For iLinha = 1 To objGridCondPagto.iLinhasExistentes
            
            GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissao_Col) = Format(StrParaDate(DataEmiPadrao.Text), "dd/mm/yyyy")
    
        Next
    
    End If

    Exit Sub
    
Erro_BotaoEmiAplicar_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoEmiVouAplicar_Click()

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_BotaoEmiVouAplicar_Click

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    If StrParaDate(DataEmiVouPadrao.Text) <> DATA_NULA Then
    
        For iLinha = 1 To objGridCondPagto.iLinhasExistentes
            
            GridCondPagto.TextMatrix(iLinha, iGrid_CPDataEmissaoAte_Col) = Format(StrParaDate(DataEmiVouPadrao.Text), "dd/mm/yyyy")
    
        Next
    
    End If

    Exit Sub
    
Erro_BotaoEmiVouAplicar_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoVencAplicar_Click()

Dim lErro As Long
Dim iLinha As Integer

On Error GoTo Erro_BotaoVencAplicar_Click

    iFrameSelecaoAlterado = REGISTRO_ALTERADO

    If StrParaDate(DataVencPadrao.Text) <> DATA_NULA Then
    
        For iLinha = 1 To objGridCondPagto.iLinhasExistentes
            
            GridCondPagto.TextMatrix(iLinha, iGrid_CPDataVencimento_Col) = Format(StrParaDate(DataVencPadrao.Text), "dd/mm/yyyy")
    
        Next
    
    End If

    Exit Sub
    
Erro_BotaoVencAplicar_Click:
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192174)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        If Me.ActiveControl Is ExcCliente Then Call BotaoClientes_Click
    
    End If
    
End Sub

Private Function Marca_Desmarca_Cliente(ByVal iFlag As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoCliente As ClassTRPFATInfoCliente

On Error GoTo Erro_Marca_Desmarca_Cliente

    If iLinha > 0 And iLinha <= objGridCliente.iLinhasExistentes Then
        
        Set objInfoCliente = gobjFaturamento.colInfoClientes.Item(iLinha)
        
        objInfoCliente.iMarcado = iFlag

        For Each objInfoFatura In objInfoCliente.colInfoFaturas
                        
            If objInfoFatura.lCliente = objInfoCliente.lCliente Then
            
                GridFatura.TextMatrix(objInfoFatura.iLinha, iGrid_FatSelecionado_Col) = CStr(iFlag)
                
            End If
            
            Call Marca_Desmarca_Fatura(iFlag, objInfoFatura.iLinha)
        
        Next
    
    End If
    
    Call Grid_Refresh_Checkbox(objGridFatura)
    
    Call Calcula_TotalVou
    
    Marca_Desmarca_Cliente = SUCESSO
    
    Exit Function
    
Erro_Marca_Desmarca_Cliente:

    Marca_Desmarca_Cliente = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192175)
    
    End Select
    
    Exit Function
    
End Function

Private Function Marca_Desmarca_Fatura(ByVal iFlag As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoCliente As ClassTRPFATInfoCliente

On Error GoTo Erro_Marca_Desmarca_Fatura

    If iLinha > 0 And iLinha <= objGridFatura.iLinhasExistentes Then
   
        Set objInfoFatura = gobjFaturamento.colInfoFaturas.Item(iLinha)

        objInfoFatura.iMarcado = iFlag
        
        For Each objInfoVoucher In objInfoFatura.colInfoVouchers
            
            GridVoucher.TextMatrix(objInfoVoucher.iLinha, iGrid_VouSelecionado_Col) = CStr(iFlag)
            
            Call Marca_Desmarca_Voucher(iFlag, objInfoVoucher.iLinha)
        
        Next
        
        If iFlag = MARCADO Then
        
            For Each objInfoCliente In gobjFaturamento.colInfoClientes
            
                If objInfoCliente.lCliente = objInfoFatura.lCliente Then
                
                    If objInfoCliente.iMarcado <> MARCADO Then
                    
                        objInfoCliente.iMarcado = MARCADO
                        GridCliente.TextMatrix(objInfoCliente.iLinha, iGrid_CliSelecionado_Col) = CStr(MARCADO)
                    
                        Call Grid_Refresh_Checkbox(objGridCliente)
                    End If
                
                    Exit For
                End If
            
            Next
            
        End If
        
    End If
    
    Call Grid_Refresh_Checkbox(objGridVoucher)
    
    Call Calcula_TotalVou
    
    Marca_Desmarca_Fatura = SUCESSO
    
    Exit Function
    
Erro_Marca_Desmarca_Fatura:

    Marca_Desmarca_Fatura = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192176)
    
    End Select
    
    Exit Function
    
End Function

Private Function Marca_Desmarca_Voucher(ByVal iFlag As Integer, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim objInfoFatura As ClassTRPFATInfoFatura
Dim objInfoCliente As ClassTRPFATInfoCliente

On Error GoTo Erro_Marca_Desmarca_Voucher

    If iLinha > 0 And iLinha <= objGridVoucher.iLinhasExistentes Then
   
        Set objInfoVoucher = gobjFaturamento.colInfoVouchers.Item(iLinha)
            
        objInfoVoucher.iMarcado = iFlag
        
        'Se está marcando um voucher tem que deixar marcado a fatura e o cliente
        If iFlag = MARCADO Then
        
            For Each objInfoFatura In gobjFaturamento.colInfoFaturas
            
                If objInfoFatura.lFatura = objInfoVoucher.lFatura Then
                
                    If objInfoFatura.iMarcado <> MARCADO Then
                    
                        objInfoFatura.iMarcado = MARCADO
                        GridFatura.TextMatrix(objInfoFatura.iLinha, iGrid_FatSelecionado_Col) = CStr(MARCADO)
                        
                        For Each objInfoCliente In gobjFaturamento.colInfoClientes
                        
                            If objInfoCliente.lCliente = objInfoFatura.lCliente Then
                            
                                If objInfoCliente.iMarcado <> MARCADO Then
                                
                                    objInfoCliente.iMarcado = MARCADO
                                    GridCliente.TextMatrix(objInfoCliente.iLinha, iGrid_CliSelecionado_Col) = CStr(MARCADO)
                                
                                    Call Grid_Refresh_Checkbox(objGridCliente)
                                End If
                            
                                Exit For
                            End If
                        
                        Next
                        
                        Call Grid_Refresh_Checkbox(objGridFatura)

                    End If
                                        
                    Exit For
                End If
            Next
        
        End If
        
    End If
    
    Call Calcula_TotalVou
        
    Marca_Desmarca_Voucher = SUCESSO
    
    Exit Function
    
Erro_Marca_Desmarca_Voucher:

    Marca_Desmarca_Voucher = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192177)
    
    End Select
    
    Exit Function
    
End Function

Private Sub UpDownDataEmiPadrao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmiPadrao_DownClick

    DataEmiPadrao.SetFocus

    If Len(DataEmiPadrao.ClipText) > 0 Then

        sData = DataEmiPadrao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 192178

        DataEmiPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmiPadrao_DownClick:

    Select Case gErr

        Case 192178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192179)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmiPadrao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmiPadrao_UpClick

    DataEmiPadrao.SetFocus

    If Len(Trim(DataEmiPadrao.ClipText)) > 0 Then

        sData = DataEmiPadrao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 192180

        DataEmiPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmiPadrao_UpClick:

    Select Case gErr

        Case 192180

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192181)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiPadrao, iAlterado)
    
End Sub

Private Sub DataEmiPadrao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiPadrao_Validate

    If Len(Trim(DataEmiPadrao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataEmiPadrao.Text)
        If lErro <> SUCESSO Then gError 192182

    End If

    Exit Sub

Erro_DataEmiPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 192182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192183)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiPadrao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataVencPadrao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataVencPadrao_DownClick

    DataVencPadrao.SetFocus

    If Len(DataVencPadrao.ClipText) > 0 Then

        sData = DataVencPadrao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 192178

        DataVencPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataVencPadrao_DownClick:

    Select Case gErr

        Case 192178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192179)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVencPadrao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataVencPadrao_UpClick

    DataVencPadrao.SetFocus

    If Len(Trim(DataVencPadrao.ClipText)) > 0 Then

        sData = DataVencPadrao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 192180

        DataVencPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataVencPadrao_UpClick:

    Select Case gErr

        Case 192180

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192181)

    End Select

    Exit Sub

End Sub

Private Sub DataVencPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencPadrao, iAlterado)
    
End Sub

Private Sub DataVencPadrao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencPadrao_Validate

    If Len(Trim(DataVencPadrao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataVencPadrao.Text)
        If lErro <> SUCESSO Then gError 192182

    End If

    Exit Sub

Erro_DataVencPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 192182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192183)

    End Select

    Exit Sub

End Sub

Private Sub DataVencPadrao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataEmiVouPadrao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmiVouPadrao_DownClick

    DataEmiVouPadrao.SetFocus

    If Len(DataEmiVouPadrao.ClipText) > 0 Then

        sData = DataEmiVouPadrao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 192178

        DataEmiVouPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmiVouPadrao_DownClick:

    Select Case gErr

        Case 192178

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192179)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEmiVouPadrao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataEmiVouPadrao_UpClick

    DataEmiVouPadrao.SetFocus

    If Len(Trim(DataEmiVouPadrao.ClipText)) > 0 Then

        sData = DataEmiVouPadrao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 192180

        DataEmiVouPadrao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataEmiVouPadrao_UpClick:

    Select Case gErr

        Case 192180

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192181)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiVouPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmiVouPadrao, iAlterado)
    
End Sub

Private Sub DataEmiVouPadrao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmiVouPadrao_Validate

    If Len(Trim(DataEmiVouPadrao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataEmiVouPadrao.Text)
        If lErro <> SUCESSO Then gError 192182

    End If

    Exit Sub

Erro_DataEmiVouPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 192182

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192183)

    End Select

    Exit Sub

End Sub

Private Sub DataEmiVouPadrao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Sub BotaoCliente_Click()

Dim objCliente As New ClassCliente
Dim objFornecedores As New ClassFornecedor

On Error GoTo Erro_BotaoCliente_Click

    If GridCliente.Row = 0 Then gError 192189
       
    If gobjFaturamento.colInfoClientes.Item(GridCliente.Row).iTipo = TRP_CLIENTEINFO_TIPO_CLIENTE Then
        
        objCliente.lCodigo = LCodigo_Extrai(GridCliente.TextMatrix(GridCliente.Row, iGrid_CliCliente_Col))
    
        Call Chama_Tela("Clientes", objCliente)
        
    Else
 
        objFornecedores.lCodigo = LCodigo_Extrai(GridCliente.TextMatrix(GridCliente.Row, iGrid_CliCliente_Col))
    
        Call Chama_Tela("Fornecedores", objFornecedores)
        
    End If


    Exit Sub

Erro_BotaoCliente_Click:

    Select Case gErr
    
        Case 192189
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192190)

    End Select

    Exit Sub
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub
Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Private Sub BotaoVoucher_Click()

Dim objVoucher As New ClassTRPVouchers

On Error GoTo Erro_BotaoVoucher_Click

    If GridVoucher.Row = 0 Then gError 192875

    objVoucher.lNumVou = StrParaLong(GridVoucher.TextMatrix(GridVoucher.Row, iGrid_VouNumero_Col))
    objVoucher.sSerie = GridVoucher.TextMatrix(GridVoucher.Row, iGrid_VouSerie_Col)
    objVoucher.sTipVou = GridVoucher.TextMatrix(GridVoucher.Row, iGrid_VouTipoV_Col)

    Call Chama_Tela("TRPVoucher", objVoucher)

    Exit Sub

Erro_BotaoVoucher_Click:

    Select Case gErr
    
        Case 192875
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192876)

    End Select

    Exit Sub
    
End Sub

Private Function Carrega_CategoriaClienteItem() As Long
'Carrega a Combo CategoriaClienteItem

Dim lErro As Long
Dim colItensCategoria As New Collection
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente

On Error GoTo Erro_Carrega_CategoriaClienteItem

    TipoFaturamento.Clear

    objCategoriaCliente.sCategoria = TRP_CATEGORIA_CONDFAT

    'Lê a tabela CategoriaProdutoItem a partir da Categoria
    lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colItensCategoria)
    If lErro <> SUCESSO Then gError 194197

    'Insere na combo CategoriaClienteItem
    For Each objCategoriaClienteItem In colItensCategoria

        'Insere na combo CategoriaCliente
        TipoFaturamento.AddItem objCategoriaClienteItem.sItem

    Next

    Carrega_CategoriaClienteItem = SUCESSO

    Exit Function

Erro_Carrega_CategoriaClienteItem:

    Carrega_CategoriaClienteItem = gErr

    Select Case gErr

        Case 194197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194198)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_ExcVou(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Produtos1

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_ExcVou

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data")

    'campos de edição do grid
    objGridInt.colCampo.Add (ExcVouTipo.Name)
    objGridInt.colCampo.Add (ExcVouSerie.Name)
    objGridInt.colCampo.Add (ExcVouNum.Name)
    objGridInt.colCampo.Add (ExcVouData.Name)
    
    'indica onde estao situadas as colunas do grid
    iGrid_ExcVouTipo_Col = 1
    iGrid_ExcVouSerie_Col = 2
    iGrid_ExcVouNum_Col = 3
    iGrid_ExcVouData_Col = 4
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridExcVou

    'Linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 4
    
    'Largura da primeira coluna
    GridExcVou.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExcVou = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ExcVou:

    Inicializa_Grid_ExcVou = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 192100)

    End Select

    Exit Function

End Function

Private Sub GridExcVou_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridExcVou, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcVou, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridExcVou)

End Sub

Private Sub GridExcVou_GotFocus()
    Call Grid_Recebe_Foco(objGridExcVou)
End Sub

Private Sub GridExcVou_EnterCell()
    Call Grid_Entrada_Celula(objGridExcVou, iAlterado)
End Sub

Private Sub GridExcVou_LeaveCell()
    Call Saida_Celula(objGridExcVou)
End Sub

Private Sub GridExcVou_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridExcVou, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridExcVou, iAlterado)
    End If

End Sub

Private Sub GridExcVou_RowColChange()
    Call Grid_RowColChange(objGridExcVou)
End Sub

Private Sub GridExcVou_Scroll()
    Call Grid_Scroll(objGridExcVou)
End Sub

Private Sub GridExcVou_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridExcVou)

End Sub

Private Sub GridExcVou_LostFocus()
    Call Grid_Libera_Foco(objGridExcVou)
End Sub

Public Sub ExcVouTipo_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcVouTipo_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcVou)
End Sub

Public Sub ExcVouTipo_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcVou)
End Sub

Public Sub ExcVouTipo_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridExcVou.objControle = ExcVouTipo
    lErro = Grid_Campo_Libera_Foco(objGridExcVou)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ExcVouSerie_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcVouSerie_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcVou)
End Sub

Public Sub ExcVouSerie_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcVou)
End Sub

Public Sub ExcVouSerie_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridExcVou.objControle = ExcVouSerie
    lErro = Grid_Campo_Libera_Foco(objGridExcVou)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ExcVouNum_Change()
    iAlterado = REGISTRO_ALTERADO
    iFrameSelecaoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcVouNum_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridExcVou)
End Sub

Public Sub ExcVouNum_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridExcVou)
End Sub

Public Sub ExcVouNum_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridExcVou.objControle = ExcVouNum
    lErro = Grid_Campo_Libera_Foco(objGridExcVou)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Function Saida_Celula_ExcVouNum(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVoucher As New ClassTRPVouchers

On Error GoTo Erro_Saida_Celula_ExcVouNum

    Set objGridInt.objControle = ExcVouNum

    If Len(Trim(ExcVouNum.Text)) > 0 Then
        
        objVoucher.lNumVou = StrParaLong(ExcVouNum.Text)
        objVoucher.sSerie = GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouSerie_Col)
        objVoucher.sTipVou = GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouTipo_Col)
        
        lErro = CF("TRPVouchers_Le", objVoucher)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 194427
        
        If lErro <> SUCESSO Then gError 194428
        
        If objVoucher.iStatus = STATUS_TRP_VOU_CANCELADO Then gError 194429
        
        GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouData_Col) = Format(objVoucher.dtData, "dd/mm/yyyy")
    
        'verifica se precisa preencher o grid com uma nova linha
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else
        GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouData_Col) = ""
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 194430

    Saida_Celula_ExcVouNum = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcVouNum:

    Saida_Celula_ExcVouNum = gErr

    Select Case gErr
    
        Case 194427, 194430
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
           
        Case 194428
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_NAO_CADASTRADO", gErr, objVoucher.lNumVou, objVoucher.sSerie, objVoucher.sTipVou)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 194429
            Call Rotina_Erro(vbOKOnly, "ERRO_VOUCHER_JA_CANCELADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194431)

    End Select

    Exit Function

End Function

Public Sub BotaoExcVou_Click()

Dim colSelecao As New Collection
Dim objVoucher As New ClassTRPVouchers

On Error GoTo Erro_BotaoExcVou_Click

    If GridExcVou.Row = 0 Then gError 194427

    If Me.ActiveControl Is ExcVouNum Then
        objVoucher.lNumVou = StrParaLong(ExcVouNum.Text)
    Else
        objVoucher.lNumVou = StrParaLong(GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouNum_Col))
    End If

    Call Chama_Tela("TRPVoucherRapidoLista", colSelecao, objVoucher, objEventoVoucher)

    Exit Sub
    
Erro_BotaoExcVou_Click:
    
    Select Case gErr
    
        Case 194427
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 194428)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoVoucher_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objVoucher As ClassTRPVouchers

On Error GoTo Erro_objEventoVoucher_evSelecao

    Set objVoucher = obj1

    If Me.ActiveControl Is ExcVouTipo Then
        ExcVouTipo.Text = objVoucher.sTipVou
    Else
        GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouTipo_Col) = objVoucher.sTipVou
    End If

    If Me.ActiveControl Is ExcVouSerie Then
        ExcVouSerie.Text = objVoucher.sSerie
    Else
        GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouSerie_Col) = objVoucher.sSerie
    End If

    If Me.ActiveControl Is ExcVouNum Then
        ExcVouNum.Text = CStr(objVoucher.lNumVou)
    Else
        GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouNum_Col) = CStr(objVoucher.lNumVou)
    End If
    
    GridExcVou.TextMatrix(GridExcVou.Row, iGrid_ExcVouData_Col) = Format(objVoucher.dtData, "dd/mm/yyyy")

    'verifica se precisa preencher o grid com uma nova linha
    If GridExcVou.Row - GridExcVou.FixedRows = objGridExcVou.iLinhasExistentes Then
        objGridExcVou.iLinhasExistentes = objGridExcVou.iLinhasExistentes + 1
    End If
        
    Me.Show

    Exit Sub

Erro_objEventoVoucher_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194429)

    End Select

    Exit Sub

End Sub

Private Sub Calcula_TotalVou()

Dim lErro As Long
Dim objInfoVoucher As ClassTRPFATInfoVoucher
Dim dTotal As Double

On Error GoTo Erro_Calcula_TotalVou

    For Each objInfoVoucher In gobjFaturamento.colInfoVouchers
            
        If objInfoVoucher.iMarcado = MARCADO Then
        
            dTotal = dTotal + objInfoVoucher.dValor
        
        End If

    Next
    
    TotalVou.Caption = Format(dTotal, "STANDARD")

    Exit Sub

Erro_Calcula_TotalVou:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 194429)

    End Select

    Exit Sub

End Sub
