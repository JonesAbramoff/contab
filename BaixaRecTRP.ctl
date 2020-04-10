VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl BaixaRec 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9405
   Begin VB.Frame Frame11 
      Caption         =   "Outros"
      Height          =   390
      Left            =   120
      TabIndex        =   166
      Top             =   0
      Visible         =   0   'False
      Width           =   7410
      Begin VB.Frame Frame12 
         Caption         =   "Vendedor"
         Height          =   720
         Left            =   0
         TabIndex        =   171
         Top             =   375
         Width           =   8520
         Begin MSMask.MaskEdBox Vendedor 
            Height          =   315
            Left            =   1200
            TabIndex        =   172
            Top             =   240
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label VendedorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
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
            Left            =   210
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   173
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.ComboBox FormaPagto 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   167
         Top             =   240
         Width           =   1905
      End
      Begin MSMask.MaskEdBox CodigoViagem 
         Height          =   315
         Left            =   1200
         TabIndex        =   168
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   6
         Mask            =   "999999"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFormaPagto 
         Caption         =   "Forma de Pagamento: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   3210
         TabIndex        =   170
         Top             =   285
         Width           =   1995
      End
      Begin VB.Label LabelViagem 
         Caption         =   "Viagem:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   169
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5100
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   780
      Width           =   9135
      Begin VB.Frame Frame9 
         Caption         =   "Filtros"
         Height          =   3810
         Left            =   120
         TabIndex        =   83
         Top             =   1140
         Width           =   8895
         Begin VB.Frame Frame3 
            Caption         =   "Tipo de Documento"
            Height          =   1575
            Left            =   5715
            TabIndex        =   155
            Top             =   2040
            Width           =   3030
            Begin VB.OptionButton TipoDocApenas 
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
               Left            =   90
               TabIndex        =   23
               Top             =   960
               Width           =   1050
            End
            Begin VB.OptionButton TipoDocTodos 
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
               Left            =   75
               TabIndex        =   22
               Top             =   360
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.ComboBox TipoDocSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "BaixaRecTRP.ctx":0000
               Left            =   1155
               List            =   "BaixaRecTRP.ctx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   24
               Top             =   930
               Width           =   1755
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "Conta Corrente do Adiantamento"
            Height          =   1575
            Left            =   2400
            TabIndex        =   154
            Top             =   2040
            Width           =   3195
            Begin VB.ComboBox ContaCorrenteSeleciona 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "BaixaRecTRP.ctx":0004
               Left            =   1170
               List            =   "BaixaRecTRP.ctx":0006
               TabIndex        =   21
               Top             =   930
               Width           =   1935
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
               Left            =   150
               TabIndex        =   19
               Top             =   375
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
               Left            =   150
               TabIndex        =   20
               Top             =   960
               Width           =   1095
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Adiantamento"
            Height          =   1575
            Left            =   225
            TabIndex        =   151
            Top             =   2040
            Width           =   2055
            Begin MSComCtl2.UpDown UpDownRADe 
               Height          =   300
               Left            =   1680
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   450
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               OrigLeft        =   1680
               OrigTop         =   457
               OrigRight       =   1920
               OrigBottom      =   757
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataRADe 
               Height          =   300
               Left            =   615
               TabIndex        =   15
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
            Begin MSComCtl2.UpDown UpDownRAAte 
               Height          =   300
               Left            =   1680
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   960
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataRAAte 
               Height          =   300
               Left            =   615
               TabIndex        =   17
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label8 
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
               TabIndex        =   153
               Top             =   990
               Width           =   375
            End
            Begin VB.Label Label9 
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
               TabIndex        =   152
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Valor da Parcela"
            Height          =   1575
            Left            =   6570
            TabIndex        =   148
            Top             =   255
            Width           =   2175
            Begin MSMask.MaskEdBox ValorDe 
               Height          =   300
               Left            =   735
               TabIndex        =   13
               Top             =   435
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox ValorAte 
               Height          =   300
               Left            =   735
               TabIndex        =   14
               Top             =   960
               Width           =   1140
               _ExtentX        =   2011
               _ExtentY        =   529
               _Version        =   393216
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label7 
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
               Left            =   360
               TabIndex        =   150
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label2 
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
               Left            =   315
               TabIndex        =   149
               Top             =   990
               Width           =   375
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Nº do Título"
            Height          =   1575
            Left            =   4650
            TabIndex        =   90
            Top             =   255
            Width           =   1770
            Begin MSMask.MaskEdBox TituloInic 
               Height          =   300
               Left            =   585
               TabIndex        =   11
               Top             =   435
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "999999999"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox TituloFim 
               Height          =   300
               Left            =   585
               TabIndex        =   12
               Top             =   960
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   9
               Mask            =   "999999999"
               PromptChar      =   " "
            End
            Begin VB.Label Label22 
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
               Left            =   180
               TabIndex        =   101
               Top             =   990
               Width           =   375
            End
            Begin VB.Label Label21 
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
               Left            =   225
               TabIndex        =   102
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Data de Vencimento"
            Height          =   1575
            Left            =   2415
            TabIndex        =   87
            Top             =   255
            Width           =   2100
            Begin MSComCtl2.UpDown UpDownVencInic 
               Height          =   300
               Left            =   1680
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   480
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencInic 
               Height          =   300
               Left            =   615
               TabIndex        =   7
               Top             =   465
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownVencFim 
               Height          =   300
               Left            =   1680
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   990
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox VencFim 
               Height          =   300
               Left            =   600
               TabIndex        =   9
               Top             =   990
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label20 
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
               TabIndex        =   103
               Top             =   1020
               Width           =   375
            End
            Begin VB.Label Label17 
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
               TabIndex        =   104
               Top             =   510
               Width           =   375
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de Emissão"
            Height          =   1575
            Left            =   210
            TabIndex        =   84
            Top             =   255
            Width           =   2070
            Begin MSComCtl2.UpDown UpDownEmissaoInic 
               Height          =   300
               Left            =   1680
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   457
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox EmissaoInic 
               Height          =   300
               Left            =   615
               TabIndex        =   3
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
            Begin MSComCtl2.UpDown UpDownEmissaoFim 
               Height          =   300
               Left            =   1680
               TabIndex        =   6
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
               Left            =   615
               TabIndex        =   5
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin VB.Label Label16 
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
               TabIndex        =   105
               Top             =   990
               Width           =   375
            End
            Begin VB.Label Label11 
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
               TabIndex        =   106
               Top             =   480
               Width           =   375
            End
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Cliente"
         Height          =   900
         Left            =   90
         TabIndex        =   80
         Top             =   120
         Width           =   8895
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5490
            TabIndex        =   2
            Top             =   315
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1605
            TabIndex        =   1
            Top             =   315
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   20
            PromptChar      =   "_"
         End
         Begin VB.Label LabelCli 
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
            Height          =   255
            Left            =   810
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   107
            Top             =   360
            Width           =   675
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4875
            TabIndex        =   108
            Top             =   360
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5160
      Index           =   2
      Left            =   105
      TabIndex        =   25
      Top             =   735
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Débitos"
         Height          =   1620
         Index           =   2
         Left            =   75
         TabIndex        =   86
         Top             =   3495
         Visible         =   0   'False
         Width           =   8910
         Begin VB.CheckBox SelecionarDB 
            Height          =   225
            Left            =   6555
            TabIndex        =   160
            Top             =   180
            Width           =   390
         End
         Begin VB.TextBox OBSDB 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   225
            Left            =   4200
            TabIndex        =   158
            Top             =   480
            Width           =   2070
         End
         Begin MSMask.MaskEdBox SaldoDB 
            Height          =   225
            Left            =   4170
            TabIndex        =   157
            Top             =   300
            Width           =   825
            _ExtentX        =   1455
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
         Begin MSMask.MaskEdBox DataEmissaoDB 
            Height          =   240
            Left            =   0
            TabIndex        =   159
            Top             =   30
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   423
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox ValorDB 
            Height          =   225
            Left            =   3000
            TabIndex        =   161
            Top             =   105
            Width           =   825
            _ExtentX        =   1455
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
         Begin MSMask.MaskEdBox NumTituloDB 
            Height          =   225
            Left            =   2250
            TabIndex        =   162
            Top             =   90
            Width           =   675
            _ExtentX        =   1191
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
         Begin MSMask.MaskEdBox TipoDocDB 
            Height          =   225
            Left            =   1560
            TabIndex        =   163
            Top             =   90
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   50
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox FilialDB 
            Height          =   225
            Left            =   5265
            TabIndex        =   164
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
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
         Begin MSFlexGridLib.MSFlexGrid GridDebitos 
            Height          =   1320
            Left            =   15
            TabIndex        =   165
            Top             =   225
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   2328
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Adiantamentos de Clientes"
         Height          =   1620
         Index           =   1
         Left            =   75
         TabIndex        =   88
         Top             =   3495
         Visible         =   0   'False
         Width           =   8910
         Begin MSMask.MaskEdBox FilialClienteRA 
            Height          =   225
            Left            =   5025
            TabIndex        =   146
            Top             =   960
            Width           =   1245
            _ExtentX        =   2196
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteRA 
            Height          =   225
            Left            =   2640
            TabIndex        =   147
            Top             =   870
            Width           =   2805
            _ExtentX        =   4948
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Hist 
            Height          =   225
            Left            =   4245
            TabIndex        =   145
            Top             =   1035
            Width           =   2925
            _ExtentX        =   5159
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Seq 
            Height          =   225
            Left            =   5865
            TabIndex        =   144
            Top             =   795
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Doc 
            Height          =   225
            Left            =   4515
            TabIndex        =   143
            Top             =   705
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
            PromptChar      =   " "
         End
         Begin VB.CheckBox SelecionarRA 
            Height          =   225
            Left            =   7890
            TabIndex        =   51
            Top             =   345
            Width           =   510
         End
         Begin MSMask.MaskEdBox ContaCorrenteRA 
            Height          =   225
            Left            =   1635
            TabIndex        =   46
            Top             =   285
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MeioPagtoRA 
            Height          =   225
            Left            =   2820
            TabIndex        =   47
            Top             =   330
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   2
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DataMovimentoRA 
            Height          =   225
            Left            =   285
            TabIndex        =   45
            Top             =   300
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
         Begin MSMask.MaskEdBox ValorRA 
            Height          =   225
            Left            =   4350
            TabIndex        =   48
            Top             =   315
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
         Begin MSMask.MaskEdBox SaldoRA 
            Height          =   225
            Left            =   5415
            TabIndex        =   49
            Top             =   300
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
         Begin MSMask.MaskEdBox FilialRA 
            Height          =   225
            Left            =   6450
            TabIndex        =   50
            Top             =   315
            Width           =   1890
            _ExtentX        =   3334
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
         Begin MSFlexGridLib.MSFlexGrid GridRecebAntecipados 
            Height          =   1275
            Left            =   120
            TabIndex        =   52
            Top             =   210
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   2249
            _Version        =   393216
            Rows            =   5
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Dados da Perda"
         Height          =   1590
         Index           =   3
         Left            =   75
         TabIndex        =   85
         Top             =   3525
         Visible         =   0   'False
         Width           =   8910
         Begin MSMask.MaskEdBox HistoricoPerda 
            Height          =   300
            Left            =   1620
            TabIndex        =   53
            Top             =   645
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
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
            Left            =   705
            TabIndex        =   113
            Top             =   690
            Width           =   810
         End
      End
      Begin VB.Frame FrameRecebimento 
         Caption         =   "Dados do Recebimento"
         Height          =   1620
         Index           =   0
         Left            =   75
         TabIndex        =   100
         Top             =   3495
         Width           =   8910
         Begin VB.Frame FrameTipoMeioPagto 
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   0
            Left            =   6750
            TabIndex        =   81
            Top             =   240
            Width           =   1770
         End
         Begin VB.Frame FrameTipoMeioPagto 
            BorderStyle     =   0  'None
            Height          =   585
            Index           =   1
            Left            =   6780
            TabIndex        =   82
            Top             =   240
            Width           =   1770
         End
         Begin VB.ComboBox ContaCorrente 
            Height          =   315
            Left            =   1125
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   315
            Width           =   1695
         End
         Begin MSMask.MaskEdBox Historico 
            Height          =   300
            Left            =   1125
            TabIndex        =   44
            Top             =   1200
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   465
            TabIndex        =   109
            Top             =   345
            Width           =   570
         End
         Begin VB.Label Label13 
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
            Left            =   210
            TabIndex        =   110
            Top             =   1230
            Width           =   825
         End
         Begin VB.Label Label6 
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
            Left            =   525
            TabIndex        =   111
            Top             =   795
            Width           =   510
         End
         Begin VB.Label ValorReceber 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1125
            TabIndex        =   112
            Top             =   765
            Width           =   1680
         End
      End
      Begin VB.Frame Cobranca 
         Caption         =   "Parcelas em Aberto"
         Height          =   2550
         Left            =   75
         TabIndex        =   99
         Top             =   405
         Width           =   8910
         Begin MSMask.MaskEdBox NumVou 
            Height          =   225
            Left            =   3480
            TabIndex        =   141
            Top             =   960
            Width           =   780
            _ExtentX        =   1376
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
         Begin MSMask.MaskEdBox NumCCred 
            Height          =   225
            Left            =   2955
            TabIndex        =   140
            Top             =   780
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Height          =   225
            Left            =   6090
            TabIndex        =   92
            Top             =   1035
            Width           =   1770
            _ExtentX        =   3122
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
         Begin MSMask.MaskEdBox FilialClienteItem 
            Height          =   225
            Left            =   5505
            TabIndex        =   138
            Top             =   675
            Width           =   1245
            _ExtentX        =   2196
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ClienteItem 
            Height          =   225
            Left            =   3120
            TabIndex        =   139
            Top             =   585
            Width           =   2805
            _ExtentX        =   4948
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
            PromptChar      =   " "
         End
         Begin VB.CheckBox Selecionar 
            Height          =   225
            Left            =   7365
            TabIndex        =   35
            Top             =   375
            Width           =   525
         End
         Begin MSMask.MaskEdBox DataVencimento 
            Height          =   225
            Left            =   870
            TabIndex        =   30
            Top             =   705
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Saldo 
            Height          =   225
            Left            =   3210
            TabIndex        =   32
            Top             =   1005
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
         Begin MSMask.MaskEdBox ValorBaixar 
            Height          =   225
            Left            =   4410
            TabIndex        =   33
            Top             =   1035
            Width           =   1005
            _ExtentX        =   1773
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
         Begin MSMask.MaskEdBox Numero 
            Height          =   225
            Left            =   2010
            TabIndex        =   29
            Top             =   450
            Width           =   780
            _ExtentX        =   1376
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
         Begin MSMask.MaskEdBox Tipo 
            Height          =   225
            Left            =   1230
            TabIndex        =   89
            Top             =   465
            Width           =   615
            _ExtentX        =   1085
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Parcela 
            Height          =   225
            Left            =   2130
            TabIndex        =   31
            Top             =   915
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "99"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorDesconto 
            Height          =   225
            Left            =   4725
            TabIndex        =   37
            Top             =   345
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorMulta 
            Height          =   225
            Left            =   6195
            TabIndex        =   34
            Top             =   405
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox ValorJuros 
            Height          =   225
            Left            =   7365
            TabIndex        =   36
            Top             =   720
            Width           =   855
            _ExtentX        =   1508
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
         Begin MSMask.MaskEdBox Cobrador 
            Height          =   225
            Left            =   495
            TabIndex        =   28
            Top             =   465
            Width           =   2055
            _ExtentX        =   3625
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
         Begin MSFlexGridLib.MSFlexGrid GridParcelas 
            Height          =   1230
            Left            =   120
            TabIndex        =   38
            Top             =   195
            Width           =   8625
            _ExtentX        =   15214
            _ExtentY        =   2170
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            FocusRect       =   2
         End
         Begin MSMask.MaskEdBox ValorParcela 
            Height          =   225
            Left            =   4500
            TabIndex        =   91
            Top             =   675
            Width           =   1260
            _ExtentX        =   2223
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
         Begin MSMask.MaskEdBox ValorAReceber 
            Height          =   225
            Left            =   3120
            TabIndex        =   93
            Top             =   705
            Width           =   1275
            _ExtentX        =   2249
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
      End
      Begin VB.Frame Frame10 
         Caption         =   "Tipo de Baixa"
         Height          =   495
         Left            =   75
         TabIndex        =   79
         Top             =   2985
         Width           =   8910
         Begin VB.OptionButton Recebimento 
            Caption         =   "Débito / Devolução"
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
            Left            =   4795
            TabIndex        =   41
            Top             =   210
            Width           =   2055
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Adiantamento"
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
            Left            =   2660
            TabIndex        =   40
            Top             =   210
            Width           =   1575
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Recebimento"
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
            Left            =   630
            TabIndex        =   39
            Top             =   210
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton Recebimento 
            Caption         =   "Perda"
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
            Index           =   3
            Left            =   7410
            TabIndex        =   42
            Top             =   210
            Width           =   825
         End
      End
      Begin MSComCtl2.UpDown UpDownDataBaixa 
         Height          =   300
         Left            =   2550
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataBaixa 
         Height          =   300
         Left            =   1485
         TabIndex        =   26
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataCredito 
         Height          =   300
         Left            =   5415
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   90
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataCredito 
         Height          =   300
         Left            =   4335
         TabIndex        =   27
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label18 
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
         Left            =   135
         TabIndex        =   114
         Top             =   150
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data Crédito:"
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
         Left            =   3120
         TabIndex        =   115
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label TotalBaixar 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6765
         TabIndex        =   116
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label5 
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
         Left            =   6180
         TabIndex        =   117
         Top             =   150
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4605
      Index           =   3
      Left            =   195
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox CTBGerencial 
         Height          =   210
         Left            =   4560
         TabIndex        =   156
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
         Left            =   6300
         TabIndex        =   60
         Top             =   315
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
         Left            =   6300
         TabIndex        =   58
         Top             =   0
         Width           =   1245
      End
      Begin VB.ComboBox CTBModelo 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   840
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
         Left            =   7755
         TabIndex        =   59
         Top             =   0
         Width           =   1245
      End
      Begin MSMask.MaskEdBox CTBSeqContraPartida 
         Height          =   225
         Left            =   4800
         TabIndex        =   68
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
         TabIndex        =   70
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox CTBHistorico 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4245
         MaxLength       =   150
         TabIndex        =   69
         Top             =   2175
         Width           =   1770
      End
      Begin VB.ListBox CTBListHistoricos 
         Height          =   2985
         Left            =   6330
         TabIndex        =   72
         Top             =   1515
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Frame CTBFrame7 
         Caption         =   "Descrição do Elemento Selecionado"
         Height          =   1050
         Left            =   195
         TabIndex        =   95
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
            TabIndex        =   118
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
            TabIndex        =   119
            Top             =   300
            Width           =   570
         End
         Begin VB.Label CTBContaDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   120
            Top             =   285
            Width           =   3720
         End
         Begin VB.Label CTBCclDescricao 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1845
            TabIndex        =   121
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
         Left            =   3450
         TabIndex        =   63
         Top             =   960
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin MSMask.MaskEdBox CTBConta 
         Height          =   225
         Left            =   525
         TabIndex        =   64
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox CTBDataContabil 
         Height          =   300
         Left            =   570
         TabIndex        =   57
         Top             =   525
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
         TabIndex        =   56
         Top             =   135
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
         Left            =   3810
         TabIndex        =   55
         Top             =   120
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
         TabIndex        =   71
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         TabIndex        =   61
         Top             =   630
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
         TabIndex        =   122
         Top             =   165
         Width           =   720
      End
      Begin VB.Label CTBOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   750
         TabIndex        =   123
         Top             =   120
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
         TabIndex        =   124
         Top             =   600
         Width           =   735
      End
      Begin VB.Label CTBPeriodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5010
         TabIndex        =   125
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label CTBExercicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2910
         TabIndex        =   126
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
         TabIndex        =   127
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
         TabIndex        =   128
         Top             =   945
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
         TabIndex        =   129
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
         TabIndex        =   130
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
         TabIndex        =   131
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
         TabIndex        =   132
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label CTBTotalDebito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3705
         TabIndex        =   133
         Top             =   3030
         Width           =   1155
      End
      Begin VB.Label CTBTotalCredito 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2460
         TabIndex        =   134
         Top             =   3030
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
         TabIndex        =   135
         Top             =   555
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
         TabIndex        =   136
         Top             =   165
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
         TabIndex        =   137
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.CheckBox ImprimirRecibo 
      Caption         =   "Imprimir Recibo ao Gravar"
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
      Left            =   4560
      TabIndex        =   142
      Top             =   165
      Width           =   2880
   End
   Begin VB.PictureBox Picture4 
      Height          =   555
      Left            =   7650
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   75
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "BaixaRecTRP.ctx":0008
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "BaixaRecTRP.ctx":0186
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "BaixaRecTRP.ctx":06B8
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5520
      Left            =   90
      TabIndex        =   98
      Top             =   420
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9737
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Títulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
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
Attribute VB_Name = "BaixaRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTBaixaRec
Attribute objCT.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
    Set objCT = New CTBaixaRec
    Set objCT.objUserControl = Me
    
    Set objCT.gobjInfoUsu = New CTBaixaRecVGTRP
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTBaixaRecTRP

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

Private Sub DataBaixa_GotFocus()
     Call objCT.DataBaixa_GotFocus
End Sub

Private Sub DataCredito_GotFocus()
     Call objCT.DataCredito_GotFocus
End Sub

Private Sub DataEmissaoDB_GotFocus()
     Call objCT.DataEmissaoDB_GotFocus
End Sub

Private Sub DataEmissaoDB_KeyPress(KeyAscii As Integer)
     Call objCT.DataEmissaoDB_KeyPress(KeyAscii)
End Sub

Private Sub DataEmissaoDB_Validate(Cancel As Boolean)
     Call objCT.DataEmissaoDB_Validate(Cancel)
End Sub

Private Sub EmissaoFim_GotFocus()
     Call objCT.EmissaoFim_GotFocus
End Sub

Private Sub EmissaoInic_GotFocus()
     Call objCT.EmissaoInic_GotFocus
End Sub

Private Sub HistoricoPerda_Change()
     Call objCT.HistoricoPerda_Change
End Sub

Private Sub LabelCli_Click()
     Call objCT.LabelCli_Click
End Sub

Private Sub TipoDocDB_GotFocus()
     Call objCT.TipoDocDB_GotFocus
End Sub

Private Sub TipoDocDB_KeyPress(KeyAscii As Integer)
     Call objCT.TipoDocDB_KeyPress(KeyAscii)
End Sub

Private Sub TipoDocDB_Validate(Cancel As Boolean)
     Call objCT.TipoDocDB_Validate(Cancel)
End Sub

Private Sub NumTituloDB_GotFocus()
     Call objCT.NumTituloDB_GotFocus
End Sub

Private Sub NumTituloDB_KeyPress(KeyAscii As Integer)
     Call objCT.NumTituloDB_KeyPress(KeyAscii)
End Sub

Private Sub NumTituloDB_Validate(Cancel As Boolean)
     Call objCT.NumTituloDB_Validate(Cancel)
End Sub

Private Sub TituloFim_GotFocus()
     Call objCT.TituloFim_GotFocus
End Sub

Private Sub TituloInic_GotFocus()
     Call objCT.TituloInic_GotFocus
End Sub

Private Sub ValorDB_GotFocus()
     Call objCT.ValorDB_GotFocus
End Sub

Private Sub ValorDB_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDB_KeyPress(KeyAscii)
End Sub

Private Sub ValorDB_Validate(Cancel As Boolean)
     Call objCT.ValorDB_Validate(Cancel)
End Sub

Private Sub SaldoDB_GotFocus()
     Call objCT.SaldoDB_GotFocus
End Sub

Private Sub SaldoDB_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoDB_KeyPress(KeyAscii)
End Sub

Private Sub SaldoDB_Validate(Cancel As Boolean)
     Call objCT.SaldoDB_Validate(Cancel)
End Sub

Private Sub SelecionarDB_GotFocus()
     Call objCT.SelecionarDB_GotFocus
End Sub

Private Sub SelecionarDB_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionarDB_KeyPress(KeyAscii)
End Sub

Private Sub SelecionarDB_Validate(Cancel As Boolean)
     Call objCT.SelecionarDB_Validate(Cancel)
End Sub

Private Sub FilialDB_GotFocus()
     Call objCT.FilialDB_GotFocus
End Sub

Private Sub FilialDB_KeyPress(KeyAscii As Integer)
     Call objCT.FilialDB_KeyPress(KeyAscii)
End Sub

Private Sub FilialDB_Validate(Cancel As Boolean)
     Call objCT.FilialDB_Validate(Cancel)
End Sub

Private Sub ContaCorrenteRA_GotFocus()
     Call objCT.ContaCorrenteRA_GotFocus
End Sub

Private Sub ContaCorrenteRA_KeyPress(KeyAscii As Integer)
     Call objCT.ContaCorrenteRA_KeyPress(KeyAscii)
End Sub

Private Sub ContaCorrenteRA_Validate(Cancel As Boolean)
     Call objCT.ContaCorrenteRA_Validate(Cancel)
End Sub

Private Sub ContaCorrente_Click()
     Call objCT.ContaCorrente_Click
End Sub

Private Sub ContaCorrente_Validate(Cancel As Boolean)
     Call objCT.ContaCorrente_Validate(Cancel)
End Sub

Private Sub DataBaixa_Change()
     Call objCT.DataBaixa_Change
End Sub

Private Sub DataBaixa_Validate(Cancel As Boolean)
     Call objCT.DataBaixa_Validate(Cancel)
End Sub

Private Sub DataMovimentoRA_GotFocus()
     Call objCT.DataMovimentoRA_GotFocus
End Sub

Private Sub DataMovimentoRA_KeyPress(KeyAscii As Integer)
     Call objCT.DataMovimentoRA_KeyPress(KeyAscii)
End Sub

Private Sub DataMovimentoRA_Validate(Cancel As Boolean)
     Call objCT.DataMovimentoRA_Validate(Cancel)
End Sub

Private Sub DataVencimento_GotFocus()
     Call objCT.DataVencimento_GotFocus
End Sub

Private Sub DataVencimento_KeyPress(KeyAscii As Integer)
     Call objCT.DataVencimento_KeyPress(KeyAscii)
End Sub

Private Sub DataVencimento_Validate(Cancel As Boolean)
     Call objCT.DataVencimento_Validate(Cancel)
End Sub

Private Sub EmissaoFim_Change()
     Call objCT.EmissaoFim_Change
End Sub

Private Sub EmissaoFim_Validate(Cancel As Boolean)
     Call objCT.EmissaoFim_Validate(Cancel)
End Sub

Private Sub EmissaoInic_Change()
     Call objCT.EmissaoInic_Change
End Sub

Private Sub EmissaoInic_Validate(Cancel As Boolean)
     Call objCT.EmissaoInic_Validate(Cancel)
End Sub

Private Sub Filial_Click()
     Call objCT.Filial_Click
End Sub

Private Sub Filial_Validate(Cancel As Boolean)
     Call objCT.Filial_Validate(Cancel)
End Sub

Private Sub FilialEmpresa_GotFocus()
     Call objCT.FilialEmpresa_GotFocus
End Sub

Private Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
     Call objCT.FilialEmpresa_KeyPress(KeyAscii)
End Sub

Private Sub FilialEmpresa_Validate(Cancel As Boolean)
     Call objCT.FilialEmpresa_Validate(Cancel)
End Sub

Private Sub FilialRA_GotFocus()
     Call objCT.FilialRA_GotFocus
End Sub

Private Sub FilialRA_KeyPress(KeyAscii As Integer)
     Call objCT.FilialRA_KeyPress(KeyAscii)
End Sub

Private Sub FilialRA_Validate(Cancel As Boolean)
     Call objCT.FilialRA_Validate(Cancel)
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Cliente_Change()
     Call objCT.Cliente_Change
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)
     Call objCT.Cliente_Validate(Cancel)
End Sub

Private Sub GridParcelas_Click()
     Call objCT.GridParcelas_Click
End Sub

Private Sub GridParcelas_GotFocus()
     Call objCT.GridParcelas_GotFocus
End Sub

Private Sub GridParcelas_EnterCell()
     Call objCT.GridParcelas_EnterCell
End Sub

Private Sub GridParcelas_LeaveCell()
     Call objCT.GridParcelas_LeaveCell
End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridParcelas_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)
     Call objCT.GridParcelas_KeyPress(KeyAscii)
End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)
     Call objCT.GridParcelas_Validate(Cancel)
End Sub

Private Sub GridParcelas_RowColChange()
     Call objCT.GridParcelas_RowColChange
End Sub

Private Sub GridParcelas_Scroll()
     Call objCT.GridParcelas_Scroll
End Sub

Private Sub GridRecebAntecipados_Click()
     Call objCT.GridRecebAntecipados_Click
End Sub

Private Sub GridRecebAntecipados_GotFocus()
     Call objCT.GridRecebAntecipados_GotFocus
End Sub

Private Sub GridRecebAntecipados_EnterCell()
     Call objCT.GridRecebAntecipados_EnterCell
End Sub

Private Sub GridRecebAntecipados_LeaveCell()
     Call objCT.GridRecebAntecipados_LeaveCell
End Sub

Private Sub GridRecebAntecipados_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridRecebAntecipados_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridRecebAntecipados_KeyPress(KeyAscii As Integer)
     Call objCT.GridRecebAntecipados_KeyPress(KeyAscii)
End Sub

Private Sub GridRecebAntecipados_Validate(Cancel As Boolean)
     Call objCT.GridRecebAntecipados_Validate(Cancel)
End Sub

Private Sub GridRecebAntecipados_RowColChange()
     Call objCT.GridRecebAntecipados_RowColChange
End Sub

Private Sub GridRecebAntecipados_Scroll()
     Call objCT.GridRecebAntecipados_Scroll
End Sub

Private Sub GridDebitos_Click()
     Call objCT.GridDebitos_Click
End Sub

Private Sub GridDebitos_GotFocus()
     Call objCT.GridDebitos_GotFocus
End Sub

Private Sub GridDebitos_EnterCell()
     Call objCT.GridDebitos_EnterCell
End Sub

Private Sub GridDebitos_LeaveCell()
     Call objCT.GridDebitos_LeaveCell
End Sub

Private Sub GridDebitos_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.GridDebitos_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridDebitos_KeyPress(KeyAscii As Integer)
     Call objCT.GridDebitos_KeyPress(KeyAscii)
End Sub

Private Sub GridDebitos_Validate(Cancel As Boolean)
     Call objCT.GridDebitos_Validate(Cancel)
End Sub

Private Sub GridDebitos_RowColChange()
     Call objCT.GridDebitos_RowColChange
End Sub

Private Sub GridDebitos_Scroll()
     Call objCT.GridDebitos_Scroll
End Sub

Private Sub Historico_Change()
     Call objCT.Historico_Change
End Sub

Private Sub MeioPagtoRA_GotFocus()
     Call objCT.MeioPagtoRA_GotFocus
End Sub

Private Sub MeioPagtoRA_KeyPress(KeyAscii As Integer)
     Call objCT.MeioPagtoRA_KeyPress(KeyAscii)
End Sub

Private Sub MeioPagtoRA_Validate(Cancel As Boolean)
     Call objCT.MeioPagtoRA_Validate(Cancel)
End Sub

Private Sub Numero_GotFocus()
     Call objCT.Numero_GotFocus
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
     Call objCT.Numero_KeyPress(KeyAscii)
End Sub

Private Sub Numero_Validate(Cancel As Boolean)
     Call objCT.Numero_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Recebimento_Click(Index As Integer)
     Call objCT.Recebimento_Click(Index)
End Sub

Private Sub Parcela_GotFocus()
     Call objCT.Parcela_GotFocus
End Sub

Private Sub Parcela_KeyPress(KeyAscii As Integer)
     Call objCT.Parcela_KeyPress(KeyAscii)
End Sub

Private Sub Parcela_Validate(Cancel As Boolean)
     Call objCT.Parcela_Validate(Cancel)
End Sub

Private Sub Saldo_GotFocus()
     Call objCT.Saldo_GotFocus
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
     Call objCT.Saldo_KeyPress(KeyAscii)
End Sub

Private Sub Saldo_Validate(Cancel As Boolean)
     Call objCT.Saldo_Validate(Cancel)
End Sub

Private Sub SaldoRA_GotFocus()
     Call objCT.SaldoRA_GotFocus
End Sub

Private Sub SaldoRA_KeyPress(KeyAscii As Integer)
     Call objCT.SaldoRA_KeyPress(KeyAscii)
End Sub

Private Sub SaldoRA_Validate(Cancel As Boolean)
     Call objCT.SaldoRA_Validate(Cancel)
End Sub

Private Sub Selecionar_Click()
     Call objCT.Selecionar_Click
End Sub

Private Sub SelecionarRA_Click()
     Call objCT.SelecionarRA_Click
End Sub

Private Sub SelecionarDB_Click()
     Call objCT.SelecionarDB_Click
End Sub

Private Sub Tipo_GotFocus()
     Call objCT.Tipo_GotFocus
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
     Call objCT.Tipo_KeyPress(KeyAscii)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub TituloFim_Change()
     Call objCT.TituloFim_Change
End Sub

Private Sub TituloFim_Validate(Cancel As Boolean)
     Call objCT.TituloFim_Validate(Cancel)
End Sub

Private Sub TituloInic_Change()
     Call objCT.TituloInic_Change
End Sub

Private Sub UpDownDataBaixa_DownClick()
     Call objCT.UpDownDataBaixa_DownClick
End Sub

Private Sub UpDownDataBaixa_UpClick()
     Call objCT.UpDownDataBaixa_UpClick
End Sub

Private Sub UpDownEmissaoFim_DownClick()
     Call objCT.UpDownEmissaoFim_DownClick
End Sub

Private Sub UpDownEmissaoFim_UpClick()
     Call objCT.UpDownEmissaoFim_UpClick
End Sub

Private Sub UpDownEmissaoInic_DownClick()
     Call objCT.UpDownEmissaoInic_DownClick
End Sub

Private Sub UpDownEmissaoInic_UpClick()
     Call objCT.UpDownEmissaoInic_UpClick
End Sub

Private Sub UpDownVencFim_DownClick()
     Call objCT.UpDownVencFim_DownClick
End Sub

Private Sub UpDownVencFim_UpClick()
     Call objCT.UpDownVencFim_UpClick
End Sub

Private Sub UpDownVencInic_DownClick()
     Call objCT.UpDownVencInic_DownClick
End Sub

Private Sub UpDownVencInic_UpClick()
     Call objCT.UpDownVencInic_UpClick
End Sub

Private Sub ValorAReceber_GotFocus()
     Call objCT.ValorAReceber_GotFocus
End Sub

Private Sub ValorAReceber_KeyPress(KeyAscii As Integer)
     Call objCT.ValorAReceber_KeyPress(KeyAscii)
End Sub

Private Sub ValorAReceber_Validate(Cancel As Boolean)
     Call objCT.ValorAReceber_Validate(Cancel)
End Sub

Private Sub ValorBaixar_Change()
     Call objCT.ValorBaixar_Change
End Sub

Private Sub ValorBaixar_GotFocus()
     Call objCT.ValorBaixar_GotFocus
End Sub

Private Sub ValorBaixar_KeyPress(KeyAscii As Integer)
     Call objCT.ValorBaixar_KeyPress(KeyAscii)
End Sub

Private Sub ValorBaixar_Validate(Cancel As Boolean)
     Call objCT.ValorBaixar_Validate(Cancel)
End Sub

Private Sub ValorDesconto_Change()
     Call objCT.ValorDesconto_Change
End Sub

Private Sub ValorDesconto_GotFocus()
     Call objCT.ValorDesconto_GotFocus
End Sub

Private Sub ValorDesconto_KeyPress(KeyAscii As Integer)
     Call objCT.ValorDesconto_KeyPress(KeyAscii)
End Sub

Private Sub ValorDesconto_Validate(Cancel As Boolean)
     Call objCT.ValorDesconto_Validate(Cancel)
End Sub

Private Sub ValorJuros_Change()
     Call objCT.ValorJuros_Change
End Sub

Private Sub ValorMulta_Change()
     Call objCT.ValorMulta_Change
End Sub

Private Sub ValorMulta_GotFocus()
     Call objCT.ValorMulta_GotFocus
End Sub

Private Sub ValorMulta_KeyPress(KeyAscii As Integer)
     Call objCT.ValorMulta_KeyPress(KeyAscii)
End Sub

Private Sub ValorMulta_Validate(Cancel As Boolean)
     Call objCT.ValorMulta_Validate(Cancel)
End Sub

Private Sub ValorJuros_GotFocus()
     Call objCT.ValorJuros_GotFocus
End Sub

Private Sub ValorJuros_KeyPress(KeyAscii As Integer)
     Call objCT.ValorJuros_KeyPress(KeyAscii)
End Sub

Private Sub ValorJuros_Validate(Cancel As Boolean)
     Call objCT.ValorJuros_Validate(Cancel)
End Sub

Private Sub Selecionar_GotFocus()
     Call objCT.Selecionar_GotFocus
End Sub

Private Sub Selecionar_KeyPress(KeyAscii As Integer)
     Call objCT.Selecionar_KeyPress(KeyAscii)
End Sub

Private Sub Selecionar_Validate(Cancel As Boolean)
     Call objCT.Selecionar_Validate(Cancel)
End Sub

Private Sub SelecionarRA_GotFocus()
     Call objCT.SelecionarRA_GotFocus
End Sub

Private Sub SelecionarRA_KeyPress(KeyAscii As Integer)
     Call objCT.SelecionarRA_KeyPress(KeyAscii)
End Sub

Private Sub SelecionarRA_Validate(Cancel As Boolean)
     Call objCT.SelecionarRA_Validate(Cancel)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Function Trata_Parametros(Optional objBaixaReceber As ClassBaixaReceber) As Long
     Trata_Parametros = objCT.Trata_Parametros(objBaixaReceber)
End Function

Private Sub ValorRA_GotFocus()
     Call objCT.ValorRA_GotFocus
End Sub

Private Sub ValorRA_KeyPress(KeyAscii As Integer)
     Call objCT.ValorRA_KeyPress(KeyAscii)
End Sub

Private Sub ValorRA_Validate(Cancel As Boolean)
     Call objCT.ValorRA_Validate(Cancel)
End Sub

Private Sub ValorParcela_GotFocus()
     Call objCT.ValorParcela_GotFocus
End Sub

Private Sub ValorParcela_KeyPress(KeyAscii As Integer)
     Call objCT.ValorParcela_KeyPress(KeyAscii)
End Sub

Private Sub ValorParcela_Validate(Cancel As Boolean)
     Call objCT.ValorParcela_Validate(Cancel)
End Sub

Private Sub VencFim_Change()
     Call objCT.VencFim_Change
End Sub

Private Sub VencFim_GotFocus()
     Call objCT.VencFim_GotFocus
End Sub

Private Sub VencFim_Validate(Cancel As Boolean)
     Call objCT.VencFim_Validate(Cancel)
End Sub

Private Sub VencInic_Change()
     Call objCT.VencInic_Change
End Sub

Private Sub VencInic_GotFocus()
     Call objCT.VencInic_GotFocus
End Sub

Private Sub VencInic_Validate(Cancel As Boolean)
     Call objCT.VencInic_Validate(Cancel)
End Sub

Private Sub DataCredito_Change()
     Call objCT.DataCredito_Change
End Sub

Private Sub DataCredito_Validate(Cancel As Boolean)
     Call objCT.DataCredito_Validate(Cancel)
End Sub

Private Sub UpDownDataCredito_DownClick()
     Call objCT.UpDownDataCredito_DownClick
End Sub

Private Sub UpDownDataCredito_UpClick()
     Call objCT.UpDownDataCredito_UpClick
End Sub

Private Sub Cobrador_GotFocus()
     Call objCT.Cobrador_GotFocus
End Sub

Private Sub Cobrador_KeyPress(KeyAscii As Integer)
     Call objCT.Cobrador_KeyPress(KeyAscii)
End Sub

Private Sub Cobrador_Validate(Cancel As Boolean)
     Call objCT.Cobrador_Validate(Cancel)
End Sub

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

Private Sub CTBLancAutomatico_Click()
     Call objCT.CTBLancAutomatico_Click
End Sub

Private Sub CTBAglutina_Click()
     Call objCT.CTBAglutina_Click
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

Private Sub CTBDataContabil_Validate(Cancel As Boolean)
     Call objCT.CTBDataContabil_Validate(Cancel)
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

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub
Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub
Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label21, Source, X, Y)
End Sub
Private Sub Label21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label21, Button, Shift, X, Y)
End Sub
Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub
Private Sub Label17_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label17, Source, X, Y)
End Sub
Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label17, Button, Shift, X, Y)
End Sub
Private Sub Label16_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label16, Source, X, Y)
End Sub
Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label16, Button, Shift, X, Y)
End Sub
Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub
Private Sub LabelCli_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCli, Source, X, Y)
End Sub
Private Sub LabelCli_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCli, Button, Shift, X, Y)
End Sub
Private Sub LabelFilial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFilial, Source, X, Y)
End Sub
Private Sub LabelFilial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFilial, Button, Shift, X, Y)
End Sub
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub
Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub
Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub
Private Sub ValorReceber_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorReceber, Source, X, Y)
End Sub
Private Sub ValorReceber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorReceber, Button, Shift, X, Y)
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub
Private Sub Label18_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label18, Source, X, Y)
End Sub
Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label18, Button, Shift, X, Y)
End Sub
Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub
Private Sub TotalBaixar_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalBaixar, Source, X, Y)
End Sub
Private Sub TotalBaixar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalBaixar, Button, Shift, X, Y)
End Sub
Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
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
     Call objCT.Opcao_BeforeClick(Cancel)
End Sub

Private Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridParcelas_MouseDown(Button, Shift, X, Y)

End Sub
Private Sub GridRecebAntecipados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridRecebAntecipados_MouseDown(Button, Shift, X, Y)

End Sub
Private Sub GridDebitos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    Call objCT.GridDebitos_MouseDown(Button, Shift, X, Y)

End Sub

Private Sub Cliente_Preenche()
     Call objCT.Cliente_Preenche
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

'##################################################################
'Inserido por Wagner 10/11/2006
Private Sub DataRADe_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRADe_Change(objCT)
End Sub

Private Sub DataRADe_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRADe_Validate(objCT, Cancel)
End Sub

Private Sub DataRADe_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRADe_GotFocus(objCT)
End Sub

Private Sub DataRAAte_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRAAte_Change(objCT)
End Sub

Private Sub DataRAAte_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRAAte_Validate(objCT, Cancel)
End Sub

Private Sub DataRAAte_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataRAAte_GotFocus(objCT)
End Sub

Private Sub UpDownRADe_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownRADe_DownClick(objCT)
End Sub

Private Sub UpDownRADe_UpClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownRADe_UpClick(objCT)
End Sub

Private Sub UpDownRAAte_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownRAAte_DownClick(objCT)
End Sub

Private Sub UpDownRAAte_UpClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDownRAAte_UpClick(objCT)
End Sub

Private Sub CtaCorrenteTodas_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CtaCorrenteTodas_Click(objCT)
End Sub

Private Sub CtaCorrenteApenas_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CtaCorrenteApenas_Click(objCT)
End Sub

Private Sub ContaCorrenteSeleciona_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ContaCorrenteSeleciona_Change(objCT)
End Sub

Private Sub ContaCorrenteSeleciona_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ContaCorrenteSeleciona_Validate(objCT, Cancel)
End Sub

Private Sub ContaCorrenteSeleciona_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ContaCorrenteSeleciona_Change(objCT)
End Sub

Private Sub ValorDe_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ValorDe_Change(objCT)
End Sub

Private Sub ValorDe_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ValorDe_Validate(objCT, Cancel)
End Sub

Private Sub ValorAte_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ValorAte_Change(objCT)
End Sub

Private Sub ValorAte_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ValorAte_Validate(objCT, Cancel)
End Sub

Private Sub TipoDocTodos_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoDocTodos_Click(objCT)
End Sub

Private Sub TipoDocApenas_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoDocApenas_Click(objCT)
End Sub

Private Sub TipoDocSeleciona_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoDocSeleciona_Change(objCT)
End Sub

Private Sub TipoDocSeleciona_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.TipoDocSeleciona_Change(objCT)
End Sub
'################################################################


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


