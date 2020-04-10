VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CondicoesPagtoOcx 
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9405
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5505
      Index           =   1
      Left            =   150
      TabIndex        =   53
      Top             =   1035
      Width           =   9000
      Begin VB.ComboBox FormaPagto 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3465
         Width           =   1905
      End
      Begin VB.ComboBox CargoMinimo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1665
         TabIndex        =   11
         Top             =   3975
         Width           =   1935
      End
      Begin VB.ListBox Condicoes 
         Height          =   5175
         IntegralHeight  =   0   'False
         Left            =   5730
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   255
         Width           =   3165
      End
      Begin VB.Frame Frame5 
         Caption         =   "Utilização"
         Height          =   660
         Left            =   390
         TabIndex        =   45
         Top             =   2085
         Width           =   5205
         Begin VB.CheckBox EmRecebimento 
            Caption         =   "Contas a Receber"
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
            Left            =   2790
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Checked
            Width           =   1875
         End
         Begin VB.CheckBox EmPagamento 
            Caption         =   "Contas a Pagar"
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
            Left            =   600
            TabIndex        =   7
            Top             =   240
            Value           =   1  'Checked
            Width           =   1770
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Identificação"
         Height          =   1755
         Left            =   390
         TabIndex        =   42
         Top             =   165
         Width           =   5205
         Begin VB.CheckBox PreCadastrada 
            Caption         =   "Pré-cadastrada"
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
            Height          =   345
            Left            =   3450
            TabIndex        =   4
            Top             =   255
            Width           =   1680
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2160
            Picture         =   "CondicoesPagtoOcx.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Numeração Automática"
            Top             =   315
            Width           =   300
         End
         Begin MSMask.MaskEdBox DescReduzida 
            Height          =   315
            Left            =   1605
            TabIndex        =   5
            Top             =   735
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   30
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Descricao 
            Height          =   315
            Left            =   1605
            TabIndex        =   6
            Top             =   1185
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1605
            TabIndex        =   2
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desc. Reduzida:"
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
            TabIndex        =   70
            Top             =   810
            Width           =   1425
         End
         Begin VB.Label Label2 
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
            Height          =   195
            Left            =   630
            TabIndex        =   44
            Top             =   1215
            Width           =   930
         End
         Begin VB.Label Label8 
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
            Left            =   900
            TabIndex        =   43
            Top             =   345
            Width           =   660
         End
      End
      Begin MSMask.MaskEdBox AcrescimoFinanceiro 
         Height          =   315
         Left            =   2340
         TabIndex        =   9
         Top             =   2985
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TetoParcela 
         Height          =   300
         Left            =   1665
         TabIndex        =   12
         Top             =   4455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Teto por Parcela:"
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
         Left            =   120
         TabIndex        =   73
         Top             =   4485
         Width           =   1500
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
         Left            =   390
         TabIndex        =   72
         Top             =   3510
         Width           =   1995
      End
      Begin VB.Label LabelCargoMinimo 
         AutoSize        =   -1  'True
         Caption         =   "Cargo Mínimo:"
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
         Left            =   375
         TabIndex        =   71
         Top             =   4050
         Width           =   1230
      End
      Begin VB.Label Label13 
         Caption         =   "Condições de Pagamento"
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
         Left            =   5760
         TabIndex        =   57
         Top             =   15
         Width           =   2190
      End
      Begin VB.Label Label7 
         Caption         =   "Acréscimo Financeiro:"
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
         Left            =   405
         TabIndex        =   46
         Top             =   3030
         Width           =   1920
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5595
      Index           =   2
      Left            =   210
      TabIndex        =   41
      Top             =   1050
      Visible         =   0   'False
      Width           =   8940
      Begin MSMask.MaskEdBox ParcDias 
         Height          =   315
         Left            =   4155
         TabIndex        =   26
         Top             =   2925
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.ComboBox ParcIntervalo 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "CondicoesPagtoOcx.ctx":00EA
         Left            =   2760
         List            =   "CondicoesPagtoOcx.ctx":00F7
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2955
         Width           =   1335
      End
      Begin VB.ComboBox ParcDataBase 
         Height          =   315
         ItemData        =   "CondicoesPagtoOcx.ctx":0122
         Left            =   390
         List            =   "CondicoesPagtoOcx.ctx":0135
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2955
         Width           =   2310
      End
      Begin VB.ComboBox ParcModif 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "CondicoesPagtoOcx.ctx":018C
         Left            =   5445
         List            =   "CondicoesPagtoOcx.ctx":019C
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2940
         Width           =   1860
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parcelas"
         Height          =   2475
         Left            =   105
         TabIndex        =   47
         Top             =   -15
         Width           =   8700
         Begin VB.CommandButton BotaoGerarParcelas 
            Caption         =   "Gerar Parcelas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3165
            TabIndex        =   21
            Top             =   1935
            Width           =   2685
         End
         Begin VB.Frame FIntervalo 
            Caption         =   "Intervalo para outras parcelas"
            Height          =   1155
            Left            =   4620
            TabIndex        =   49
            Top             =   675
            Width           =   3945
            Begin VB.Frame FrameIntervalo 
               BorderStyle     =   0  'None
               Height          =   660
               Index           =   1
               Left            =   1680
               TabIndex        =   54
               Top             =   315
               Width           =   2130
               Begin MSMask.MaskEdBox IntervaloParcelas 
                  Height          =   315
                  Left            =   1575
                  TabIndex        =   20
                  Top             =   165
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   3
                  Mask            =   "###"
                  PromptChar      =   " "
               End
               Begin VB.Label Label6 
                  Caption         =   "Intervalo em Dias:"
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
                  Height          =   210
                  Left            =   -15
                  TabIndex        =   55
                  Top             =   210
                  Width           =   1560
               End
            End
            Begin VB.OptionButton Intervalo 
               Caption         =   "Em Dias"
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
               Index           =   1
               Left            =   165
               TabIndex        =   24
               Top             =   300
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton Intervalo 
               Caption         =   "Mensal"
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
               Left            =   165
               TabIndex        =   19
               Top             =   795
               Width           =   990
            End
            Begin VB.Frame FrameIntervalo 
               BorderStyle     =   0  'None
               Height          =   795
               Index           =   0
               Left            =   1575
               TabIndex        =   50
               Top             =   270
               Visible         =   0   'False
               Width           =   2295
               Begin MSMask.MaskEdBox DiaDoMes 
                  Height          =   315
                  Left            =   1395
                  TabIndex        =   51
                  Top             =   195
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   556
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  MaxLength       =   2
                  Mask            =   "##"
                  PromptChar      =   " "
               End
               Begin VB.Label LabelDiaDoMes 
                  Caption         =   "Dia do Mês:"
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
                  Height          =   210
                  Left            =   315
                  TabIndex        =   52
                  Top             =   240
                  Width           =   1035
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Primeira Parcela"
            Height          =   1155
            Left            =   570
            TabIndex        =   48
            Top             =   675
            Width           =   3930
            Begin VB.ComboBox CmbModificador 
               Height          =   315
               ItemData        =   "CondicoesPagtoOcx.ctx":01CE
               Left            =   1875
               List            =   "CondicoesPagtoOcx.ctx":01DE
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   300
               Width           =   1860
            End
            Begin VB.OptionButton OptDiasPrimeiraParcela 
               Caption         =   "Dias"
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
               Left            =   390
               TabIndex        =   16
               Top             =   345
               Value           =   -1  'True
               Width           =   750
            End
            Begin VB.OptionButton OptDataFixa 
               Caption         =   "Data Fixa"
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
               TabIndex        =   23
               Top             =   795
               Width           =   1245
            End
            Begin MSMask.MaskEdBox DiasParaPrimeiraParcela 
               Height          =   315
               Left            =   1215
               TabIndex        =   17
               Top             =   300
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   3
               Mask            =   "###"
               PromptChar      =   " "
            End
         End
         Begin MSMask.MaskEdBox NumeroParcelas 
            Height          =   315
            Left            =   1410
            TabIndex        =   15
            Top             =   270
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   56
            Top             =   315
            Width           =   750
         End
      End
      Begin MSMask.MaskEdBox ParcPercReceb 
         Height          =   270
         Left            =   7320
         TabIndex        =   28
         Top             =   2955
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
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
         Format          =   "0.0#####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   2550
         Left            =   90
         TabIndex        =   22
         Top             =   2535
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   4498
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
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
         Left            =   6480
         TabIndex        =   69
         Top             =   5280
         Width           =   570
      End
      Begin VB.Label ParcPercTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7185
         TabIndex        =   68
         Top             =   5235
         Width           =   1245
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5610
      Index           =   3
      Left            =   195
      TabIndex        =   58
      Top             =   1005
      Visible         =   0   'False
      Width           =   8955
      Begin VB.Frame Frame7 
         Caption         =   "Resultado"
         Height          =   3855
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   8715
         Begin VB.TextBox ResultadoVcto 
            Enabled         =   0   'False
            Height          =   300
            Left            =   585
            TabIndex        =   35
            Top             =   1110
            Width           =   1530
         End
         Begin VB.TextBox ResultadoValor 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2190
            TabIndex        =   36
            Top             =   1110
            Width           =   1530
         End
         Begin MSFlexGridLib.MSFlexGrid GridResultado 
            Height          =   3435
            Left            =   2070
            TabIndex        =   34
            Top             =   300
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   6059
            _Version        =   393216
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Parâmetros"
         Height          =   1335
         Left            =   135
         TabIndex        =   60
         Top             =   105
         Width           =   8700
         Begin VB.CommandButton BotaoExemplo 
            Caption         =   "Gerar Resultado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6690
            TabIndex        =   33
            Top             =   435
            Width           =   1470
         End
         Begin MSMask.MaskEdBox ExemploDataRef 
            Height          =   300
            Left            =   4545
            TabIndex        =   30
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataRef 
            Height          =   300
            Left            =   5730
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   225
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ExemploValor 
            Height          =   300
            Left            =   780
            TabIndex        =   29
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExemploDataEmissao 
            Height          =   300
            Left            =   4560
            TabIndex        =   31
            Top             =   615
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEmissao 
            Height          =   300
            Left            =   5730
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox ExemploDataEntrega 
            Height          =   300
            Left            =   4575
            TabIndex        =   32
            Top             =   1005
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataEntrega 
            Height          =   300
            Left            =   5730
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   990
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   67
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3975
            TabIndex        =   66
            Top             =   255
            Width           =   480
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2970
            TabIndex        =   65
            Top             =   630
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Data de Entrega:"
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
            Left            =   3000
            TabIndex        =   64
            Top             =   1020
            Width           =   1470
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7035
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "CondicoesPagtoOcx.ctx":0210
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "CondicoesPagtoOcx.ctx":036A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "CondicoesPagtoOcx.ctx":04F4
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "CondicoesPagtoOcx.ctx":0A26
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   6045
      Left            =   105
      TabIndex        =   1
      Top             =   660
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10663
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Parcelas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Exemplo"
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
Attribute VB_Name = "CondicoesPagtoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iFrameAtual As Integer

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iAlterado As Integer
Dim iFrameIntervaloAtual As Integer

'DECLARAÇÃO DAS CONSTANTES GLOBAIS DA TELA
Const UMA_PARCELA = 1
Const CONDPAGTO_NORMAL = 0
Const CONDPAGTO_DATAFIXA = 1

Const NUM_MAX_PARCELAS_CONDPAGTO = 200

Dim iGrid_ParcDataBase_Col As Integer
Dim iGrid_ParcModif_Col As Integer
Dim iGrid_ParcIntervalo_Col As Integer
Dim iGrid_ParcDias_Col As Integer
Dim iGrid_ParcPercReceb_Col As Integer

Dim objGridParcelas As AdmGrid

Dim iGrid_ResultadoVcto_Col As Integer
Dim iGrid_ResultadoValor_Col As Integer

Dim objGridResultado As AdmGrid

Private Sub BotaoExemplo_Click()
    
Dim lErro As Long, objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_BotaoExemplo_Click

    lErro = Move_Tela_Memoria(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125175
        
    objCondicaoPagto.dValorTotal = StrParaDbl(ExemploValor.Text)
    objCondicaoPagto.dtDataEmissao = MaskedParaDate(ExemploDataEmissao)
    objCondicaoPagto.dtDataEntrega = MaskedParaDate(ExemploDataEntrega)
    objCondicaoPagto.dtDataRef = MaskedParaDate(ExemploDataRef)
    
    lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, True, False, True)
    If lErro <> SUCESSO Then gError 125197
    
    lErro = Exibe_Dados_Exemplo(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125198
    
    Exit Sub
     
Erro_BotaoExemplo_Click:

    Select Case gErr
          
        Case 125175, 125197, 125198
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154637)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoGerarParcelas_Click()

Dim lErro As Long, objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_BotaoGerarParcelas_Click

    'Verifica preenchimento do número de parcelas
    If Len(Trim(NumeroParcelas.Text)) = 0 Then gError 182746

    lErro = Move_Tela_Memoria(objCondicaoPagto, False)
    If lErro <> SUCESSO Then gError 125176
        
    lErro = CondicaoPagto_Gera_Parcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125177
    
    lErro = Exibe_Dados_Parcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125178
    
    Call BotaoExemplo_Click
    
    Exit Sub
     
Erro_BotaoGerarParcelas_Click:

    Select Case gErr
          
        Case 125176 To 125178
        
        Case 182746
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_NAO_PREENCHIDA", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154638)
     
    End Select
     
    Exit Sub

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático da próxima CondicaoPagto
    lErro = CF("CondicaoPagto_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57546

    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57546
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154639)
    
    End Select

    Exit Sub

End Sub

Private Sub AcrescimoFinanceiro_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub AcrescimoFinanceiro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_AcrescimoFinanceiro_Validate

    If Len(Trim(AcrescimoFinanceiro.Text)) = 0 Then Exit Sub

    lErro = Porcentagem_Critica_Negativa(AcrescimoFinanceiro.Text)
    
    If lErro = SUCESSO Then
        
        AcrescimoFinanceiro.Text = Format(AcrescimoFinanceiro.Text, "Fixed")
    
    Else
        
        Error 16414
    
    End If
    
    Exit Sub
    
Erro_AcrescimoFinanceiro_Validate:

    Cancel = True

    
    Select Case Err

        Case 16414
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154640)
    
    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim iIndice As Integer
Dim iEncontrou As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then Error 16436
    
    'se é a condição pagamento à vista --> Erro
    If CInt(Codigo.Text) = COD_A_VISTA Then Error 58100
    
    objCondicaoPagto.iCodigo = CInt(Codigo.Text)

    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 16437
    
    'Não achou a Condição de Pagamento --> erro
    If lErro = 19205 Then Error 16438
    
    'Pedido de confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CONDICAOPAGTO", objCondicaoPagto.iCodigo)

    If vbMsgRes = vbYes Then

        'Exclui CondicaoPagto
        lErro = CF("CondicaoPagto_Exclui", objCondicaoPagto)
        If lErro <> SUCESSO Then Error 16439

        'Procura índice da CondicaoPagto na ListBox
        iEncontrou = 0
        For iIndice = 0 To Condicoes.ListCount - 1
            
            If Condicoes.ItemData(iIndice) = objCondicaoPagto.iCodigo Then
                iEncontrou = 1
                Exit For
            End If
            
        Next

        'Remove CondicaoPagto da ListBox
        If iEncontrou = 1 Then Condicoes.RemoveItem (iIndice)

        Call Limpa_Tela_CondicaoPagto

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 16436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 16437, 16439

        Case 16438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)
        
        Case 58100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_EXCLUSAO_COND_PAGAMENTO_A_VISTA", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154641)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 16415
    
    Call Limpa_Tela_CondicaoPagto
        
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 16415
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154642)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80458

    Call Limpa_Tela_CondicaoPagto

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154643)

    End Select

    Exit Sub

End Sub

Function Limpa_Tela_CondicaoPagto()
'limpa todos os campos de input da tela CondicoesPagto

Dim iIndice As Integer
Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Call Limpa_Tela(Me)
                
    'Desmarca ListBox Condicoes
    Condicoes.ListIndex = -1
        
    'Limpa os campos da tela que não foram limpos pela rotina Limpa_Tela
    EmPagamento.Value = vbChecked
    EmRecebimento.Value = vbChecked
    
    CmbModificador.ListIndex = -1
    OptDiasPrimeiraParcela.Value = True
    
    CargoMinimo.ListIndex = -1
    FormaPagto.ListIndex = -1
    
    Call HabilitaFrameIntervalo(True)
    
    Intervalo(0).Value = True
    
    'Limpa os Grids
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridResultado)
    
    ParcPercTotal.Caption = ""

    iAlterado = 0

End Function


Private Sub CmbModificador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CmbModificador_GotFocus()
    
    OptDiasPrimeiraParcela.Value = True

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) > 0 Then
    
        lErro = Inteiro_Critica(Codigo.Text)
        If lErro <> SUCESSO Then Error 16408
        
    End If
    
    Exit Sub
    
Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 16408
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154644)
    
    End Select

    Exit Sub

End Sub
Private Sub Condicoes_DblClick()

Dim lErro As Long
Dim iIndice As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Condicoes_DblClick

    objCondicaoPagto.iCodigo = Condicoes.ItemData(Condicoes.ListIndex)

    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then gError 16405

    lErro = CF("CondicaoPagto_Le_Parcelas", objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125179

    If lErro = SUCESSO Then

        'Mostra os dados na tela
        lErro = Exibe_Dados_CondicaoPagto(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 16406
        
    Else
        'CondicaoPagto não está cadastrada
        gError 16407

    End If
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

    Exit Sub

Erro_Condicoes_DblClick:

    Select Case gErr

        Case 16405, 16406

        Case 16407
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA1", gErr, Condicoes.Text)

        Case 125179

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154645)

    End Select

    Exit Sub

End Sub

Private Sub DescReduzida_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Descricao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiaDoMes_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiaDoMes_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DiaDoMes, iAlterado)

End Sub

Private Sub DiaDoMes_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DiaDoMes_Validate

    If Len(Trim(DiaDoMes.Text)) > 0 Then
        
        lErro = Inteiro_Critica(DiaDoMes.Text)
        If lErro <> SUCESSO Then Error 16411
        
        If (DiaDoMes < 1) Or (DiaDoMes > 30) Then Error 16412
        
    End If

    Exit Sub

Erro_DiaDoMes_Validate:

    Cancel = True

    
    Select Case Err

        Case 16411

        Case 16412
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIA_DO_MES_INVALIDO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154646)

    End Select

    Exit Sub
    
End Sub
Private Sub DiasParaPrimeiraParcela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DiasParaPrimeiraParcela_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DiasParaPrimeiraParcela, iAlterado)
    OptDiasPrimeiraParcela.Value = True

End Sub

Private Sub DiasParaPrimeiraParcela_Validate(Cancel As Boolean)
    
    If Len(Trim(IntervaloParcelas.Text)) = 0 Then
        IntervaloParcelas.Text = DiasParaPrimeiraParcela.Text
    End If

End Sub

Private Sub EmPagamento_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub EmRecebimento_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

 
On Error GoTo Erro_Form_Load
     
    Set colCodigoDescricao = New AdmColCodigoNome

    'Leitura dos códigos e descrições das condições de pagamento
    lErro = CF("Cod_Nomes_Le", "CondicoesPagto", "Codigo", "DescReduzida", STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 16391

    'Preenche ListBox Condições com DescReduzidas de CondicoesPagto
    For Each objCodigoDescricao In colCodigoDescricao
        
        Condicoes.AddItem objCodigoDescricao.sNome
        Condicoes.ItemData(Condicoes.NewIndex) = objCodigoDescricao.iCodigo
        
    Next
    
    'Carrega a combo
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_CARGO_VENDEDOR, CargoMinimo)
    If lErro <> SUCESSO Then gError 124021
    
    FormaPagto.AddItem ""
    FormaPagto.ItemData(FormaPagto.NewIndex) = 0
    
    Set colCodigoDescricao = New AdmColCodigoNome
    lErro = CF("FormasPagamento_Le_CodNome", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 16391
    
    'Preenche ListBox Condições com DescReduzidas de CondicoesPagto
    For Each objCodigoDescricao In colCodigoDescricao
        
        FormaPagto.AddItem objCodigoDescricao.sNome
        FormaPagto.ItemData(FormaPagto.NewIndex) = objCodigoDescricao.iCodigo
        
    Next
    
    lErro = Inicializa_GridParcelas
    If lErro <> SUCESSO Then gError 124021
    
    lErro = Inicializa_GridResultado
    If lErro <> SUCESSO Then gError 124021
    
    Call Exemplo_ValoresDefault
    
    iFrameAtual = 1
    iFrameIntervaloAtual = 1
    iAlterado = 0
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 16391

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154647)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

 Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

    Set objGridParcelas = Nothing
    Set objGridResultado = Nothing

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
        
On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CondicoesPagto"
    
    'Move dados da tela para memória,
    'realizando conversões necessárias de campos da tela para campos do BD
    'A tipagem dos valores DEVE SER A MESMA DO BD
    lErro = Move_Tela_Memoria(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 33985
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
      
    With objCondicaoPagto
    
        colCampoValor.Add "Codigo", .iCodigo, 0, "Codigo"
        colCampoValor.Add "DescReduzida", .sDescReduzida, STRING_CONDICAO_PAGTO_DESCRICAO_REDUZIDA, "DescReduzida"
        colCampoValor.Add "Descricao", .sDescricao, STRING_CONDICAO_PAGTO_DESCRICAO, "Descricao"
        colCampoValor.Add "EmPagamento", .iEmPagamento, 0, "EmPagamento"
        colCampoValor.Add "EmRecebimento", .iEmRecebimento, 0, "EmRecebimento"
        colCampoValor.Add "NumeroParcelas", .iNumeroParcelas, 0, "NumeroParcelas"
        colCampoValor.Add "DiasParaPrimeiraParcela", .iDiasParaPrimeiraParcela, 0, "DiasParaPrimeiraParcela"
        colCampoValor.Add "IntervaloParcelas", .iIntervaloParcelas, 0, "IntervaloParcelas"
        colCampoValor.Add "Mensal", .iMensal, 0, "Mensal"
        colCampoValor.Add "DiaDoMes", .iDiaDoMes, 0, "DiaDoMes"
        colCampoValor.Add "AcrescimoFinanceiro", .dAcrescimoFinanceiro, 0, "AcrescimoFinanceiro"
        colCampoValor.Add "Modificador", .iModificador, 0, "Modificador"
        colCampoValor.Add "DataFixa", .iDataFixa, 0, "DataFixa"

    End With
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 33985

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154648)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Tela_Preenche

    objCondicaoPagto.iCodigo = colCampoValor.Item("Codigo").vValor

    If objCondicaoPagto.iCodigo <> 0 Then

        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 16392
        
        lErro = CF("CondicaoPagto_Le_Parcelas", objCondicaoPagto)
        If lErro <> SUCESSO Then Error 16392

        If lErro = SUCESSO Then
            lErro = Exibe_Dados_CondicaoPagto(objCondicaoPagto)
            If lErro <> SUCESSO Then Error 16478
        End If
        
    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 16392, 16478

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154649)

    End Select

    Exit Sub
        
End Sub

Function Trata_Parametros(Optional objCondicaoPagto As ClassCondicaoPagto) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se existir uma CondicaoPagto selecionada, exibir seus dados
    If Not (objCondicaoPagto Is Nothing) Then

        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 16392

        If lErro = SUCESSO Then

            lErro = CF("CondicaoPagto_Le_Parcelas", objCondicaoPagto)
            If lErro <> SUCESSO Then Error 16392
            
            'Mostra dados na tela
            lErro = Exibe_Dados_CondicaoPagto(objCondicaoPagto)
            If lErro <> SUCESSO Then Error 16393

        Else
            
            'Limpa a tela
            Call Limpa_Tela_CondicaoPagto
            
            'Exibe apenas o código
            Codigo.Text = CStr(objCondicaoPagto.iCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 16392, 16393

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154650)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Exibe_Dados_CondicaoPagto(objCondicaoPagto As ClassCondicaoPagto) As Long
'Mostra os dados da CondicaoPagto especificada em objCondicaoPagto

Dim iIndice As Integer
Dim lErro As Long, iCargo As Integer

On Error GoTo Erro_Exibe_Dados_CondicaoPagto

    Codigo.Text = CStr(objCondicaoPagto.iCodigo)
    Descricao.Text = objCondicaoPagto.sDescricao
    DescReduzida.Text = objCondicaoPagto.sDescReduzida
    EmPagamento.Value = objCondicaoPagto.iEmPagamento
    EmRecebimento.Value = objCondicaoPagto.iEmRecebimento
    NumeroParcelas.Text = CStr(objCondicaoPagto.iNumeroParcelas)
    TetoParcela.Text = Format(objCondicaoPagto.dTetoParcela, "STANDARD")
    
    Call NumeroParcelas_Validate(bSGECancelDummy)
    
    If objCondicaoPagto.iDataFixa = CONDPAGTO_DATAFIXA Then
        
        OptDataFixa.Value = True
        DiasParaPrimeiraParcela.Text = ""
        CmbModificador.ListIndex = CONDPAGTO_MODIFICADOR_VAZIO
        
    Else
        
        OptDiasPrimeiraParcela.Value = True
    
        DiasParaPrimeiraParcela.Text = CStr(objCondicaoPagto.iDiasParaPrimeiraParcela)
        
        If objCondicaoPagto.iModificador = CONDPAGTO_MODIFICADOR_FORAMES Then
           CmbModificador.ListIndex = CONDPAGTO_MODIFICADOR_FORAMES
        ElseIf objCondicaoPagto.iModificador = CONDPAGTO_MODIFICADOR_FORAQUINZENA Then
           CmbModificador.ListIndex = CONDPAGTO_MODIFICADOR_FORAQUINZENA

            CmbModificador.ListIndex = CONDPAGTO_MODIFICADOR_VAZIO
        End If
        
    End If
    
    'Verifica se existe só uma parcela
    If objCondicaoPagto.iNumeroParcelas = UMA_PARCELA Then
        
        Intervalo(0).Value = True
        IntervaloParcelas.Text = ""
        DiaDoMes.Text = ""
    
    Else
        
        Intervalo(0).Value = CBool(objCondicaoPagto.iMensal)
        Intervalo(1).Value = Not CBool(objCondicaoPagto.iMensal)
        IntervaloParcelas.Text = CStr(objCondicaoPagto.iIntervaloParcelas)
        DiaDoMes.Text = CStr(objCondicaoPagto.iDiaDoMes)
    
    End If
    
    AcrescimoFinanceiro.Text = CStr(objCondicaoPagto.dAcrescimoFinanceiro * 100)
    
    'Coloca Cargo no Text
    If objCondicaoPagto.iCargoMinimo <> 0 Then
        CargoMinimo.Text = CStr(objCondicaoPagto.iCargoMinimo)
    
        'Tenta selecionar
        lErro = Combo_Seleciona(CargoMinimo, iCargo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 125180
    Else
        CargoMinimo.ListIndex = -1
    End If
    
    'Tenta selecionar
    Call Combo_Seleciona_ItemData(FormaPagto, objCondicaoPagto.iFormaPagamento)
    
    lErro = Exibe_Dados_Parcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 125180
    
    'Desseleciona Condicao de Pagamento na ListBox
    Condicoes.ListIndex = -1
    
    'colocar exemplo
    Call Exemplo_ValoresDefault
    Call BotaoExemplo_Click
    
    iAlterado = 0
    
    Exibe_Dados_CondicaoPagto = SUCESSO
    
    Exit Function
    
Erro_Exibe_Dados_CondicaoPagto:
    
    Exibe_Dados_CondicaoPagto = gErr
    
    Select Case gErr

        Case 125180

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154651)

    End Select

    Exit Function
    
End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

 Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub


Private Sub FormaPagto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Intervalo_Click(Index As Integer)

    If iFrameIntervaloAtual <> Index Then
    
        FrameIntervalo(Index).Visible = True
        FrameIntervalo(iFrameIntervaloAtual).Visible = False
        
        'Armazena o novo valor de iFrameIntervaloAtual
        iFrameIntervaloAtual = Index
        
        iAlterado = REGISTRO_ALTERADO
    
    End If

End Sub

Private Sub IntervaloParcelas_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub IntervaloParcelas_GotFocus()

    Call MaskEdBox_TrataGotFocus(IntervaloParcelas, iAlterado)

End Sub

Private Sub IntervaloParcelas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IntervaloParcelas_Validate

    'Verifica preenchimento do Intervalo de Parcelas
    If Len(Trim(IntervaloParcelas.Text)) > 0 Then
    
        lErro = Inteiro_Critica(IntervaloParcelas.Text)
        If lErro <> SUCESSO Then Error 16413
        
    End If
    
    Exit Sub
    
Erro_IntervaloParcelas_Validate:

    Cancel = True


    Select Case Err

        Case 16413
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154652)
    
    End Select

    Exit Sub

End Sub

Private Sub NumeroParcelas_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub NumeroParcelas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumeroParcelas, iAlterado)

End Sub

Private Sub NumeroParcelas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroParcelas_Validate

    'Verifica preenchimento do Número de Parcelas
    If Len(Trim(NumeroParcelas.Text)) = 0 Then Exit Sub
    
    'Critica se é do tipo Inteiro
    lErro = Inteiro_Critica(NumeroParcelas.Text)
    If lErro <> SUCESSO Then Error 16409
    
    'Verifica se é menor ou igual ao limite máximo
    If CInt(NumeroParcelas.Text) > NUM_MAXIMO_PARCELAS Then Error 6904
    
    'Verifica se existe só uma parcela
    If CInt(NumeroParcelas.Text) = UMA_PARCELA Then
        
        'Limpa o Frame e Desabilita
        Intervalo(0).Value = True
        DiaDoMes.Text = ""
        IntervaloParcelas = ""
        
        Call HabilitaFrameIntervalo(False)

    Else
        
        Call HabilitaFrameIntervalo(True)
        
    End If
    
    Exit Sub
    
Erro_NumeroParcelas_Validate:

    Cancel = True


    Select Case Err
        
        Case 6904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_PARCELAS_EXCESSIVO", Err, NumeroParcelas.Text, NUM_MAXIMO_PARCELAS)
        
        Case 16409
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 154653)
    
    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long, iParcelasComPerc As Integer, objParc As ClassCondicaoPagtoParc
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento do código
    If Len(Trim(Codigo.Text)) = 0 Then gError 16417
    
    If CInt(Codigo.Text) = COD_A_VISTA Then gError 58099
    
    'Verifica preenchimento da descrição reduzida
    If Len(Trim(DescReduzida.Text)) = 0 Then gError 16418

    'Verifica preenchimento do número de parcelas
    If Len(Trim(NumeroParcelas.Text)) = 0 Then gError 16419
    
    If OptDiasPrimeiraParcela.Value = True Then
        
        'Verifica preenchimento de Dias Para Primeira Parcela
        If Len(Trim(DiasParaPrimeiraParcela.Text)) = 0 Then gError 16420
    
    End If
    
    If CInt(NumeroParcelas.Text) <> UMA_PARCELA Then
    
        'Se intervalo Mensal foi selecionado,
        If Intervalo(0).Value = True Then
            'Verifica preenchimento do Dia do Mês
            If Len(Trim(DiaDoMes.Text)) = 0 Then gError 16421
        End If
    
        'Se intervalo Em dias foi selecionado,
        If Intervalo(1).Value = True Then
            'Verifica preenchimento do Intervalo entre parcelas
            If Len(Trim(IntervaloParcelas.Text)) = 0 Then gError 16422
        End If
    
    End If
    
    If EmPagamento.Value = vbUnchecked And EmRecebimento = vbUnchecked Then gError 47893
    
    'Move dados da tela para memória
    lErro = Move_Tela_Memoria(objCondicaoPagto)
    If lErro <> SUCESSO Then gError 16396
        
    'Verifica se o número de parcelas informado é igual ao número de linhas com % <> 0
    For Each objParc In objCondicaoPagto.colParcelas
        
        If objParc.dPercReceb <> 0 Then iParcelasComPerc = iParcelasComPerc + 1
        
        If objParc.iTipoDataBase = 0 Then gError 9999
        If objParc.iTipoIntervalo = 0 Then gError 9999
        
    Next
    
    If objCondicaoPagto.iNumeroParcelas <> iParcelasComPerc Then gError 125181
    
    lErro = Trata_Alteracao(objCondicaoPagto, objCondicaoPagto.iCodigo)
    If lErro <> SUCESSO Then gError 16474
    
    'Chama função de gravação
    lErro = CF("CondicaoPagto_Grava", objCondicaoPagto)
    If lErro <> SUCESSO Then Error 16423

    'Remove a CondicaoPagto da ListBox Condições
    Call Condicoes_Exclui(objCondicaoPagto.iCodigo)
    
    'Insere a CondicaoPagto na ListBox Condições
    Call Condicoes_Adiciona(objCondicaoPagto)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 16417
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 16418
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESCRICAO_REDUZIDA_NAO_PREENCHIDA", gErr)

        Case 16419
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PARCELAS_NAO_PREENCHIDA", gErr)
        
        Case 16420
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIAS_PARA_PRIMEIRA_PARCELA_NAO_PREENCHIDA", gErr)
        
        Case 16421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DIA_DO_MES_NAO_PREENCHIDO", gErr)
        
        Case 16422
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INTERVALO_ENTRE_PARCELAS_NAO_PREENCHIDO", gErr)
        
        Case 16396, 16423, 16474
            
        Case 47893
            lErro = Rotina_Erro(vbOKOnly, "ERRO_USADA_EM_NAO_PREENCHIDA", gErr)
            EmPagamento.SetFocus
        
        Case 58099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALTERACAO_COND_PAGAMENTO_A_VISTA", gErr)
        
        Case 125181
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMPARCELA_DIFERENTE_LINHASPREECHIDA", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154654)

    End Select
    
    Exit Function
        
End Function

Function Move_GridParcelas_Memoria(objCondicaoPagto As ClassCondicaoPagto) As Long

Dim lErro As Long, iIndice As Integer, objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_Move_GridParcelas_Memoria

    If objGridParcelas.iLinhasExistentes <> 0 Then
    
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
            Set objParc = New ClassCondicaoPagtoParc
        
            With objParc
            
                .iCodigo = objCondicaoPagto.iCodigo
                .iSeq = iIndice
                
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ParcDataBase_Col))) <> 0 Then
                    ParcDataBase.Text = GridParcelas.TextMatrix(iIndice, iGrid_ParcDataBase_Col)
                    .iTipoDataBase = ParcDataBase.ItemData(ParcDataBase.ListIndex)
                End If
                
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ParcIntervalo_Col))) <> 0 Then
                    ParcIntervalo.Text = GridParcelas.TextMatrix(iIndice, iGrid_ParcIntervalo_Col)
                    .iTipoIntervalo = ParcIntervalo.ItemData(ParcIntervalo.ListIndex)
                End If
                    
                If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ParcModif_Col))) = 0 Then
                    ParcModif.ListIndex = 0
                Else
                    ParcModif.Text = GridParcelas.TextMatrix(iIndice, iGrid_ParcModif_Col)
                End If
                
                If ParcModif.ListIndex = -1 Then
                    .iModificador = 0
                Else
                    .iModificador = ParcModif.ItemData(ParcModif.ListIndex)
                End If
                
                .iDias = StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_ParcDias_Col))
                .dPercReceb = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ParcPercReceb_Col)) / 100
                
            End With
            
            'Verifica se o quantidades de dias é maior que 31
            If objParc.iTipoIntervalo = 3 And objParc.iDias > 31 Then gError 125182
            
            objCondicaoPagto.colParcelas.Add objParc
            
        Next
        
        'Verifica a Soma dos Percentuais das Parcelas é Maior ou menor que 1
        If PercentParaDbl(ParcPercTotal.Caption) <> 1 Then gError 125183
    
    End If
    
    Move_GridParcelas_Memoria = SUCESSO
     
    Exit Function
    
Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = gErr
     
    Select Case gErr
          
        Case 125182
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMDIAS", gErr)
          
        Case 125183
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PERCENTUAL_NAO_VALIDA", gErr)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154655)
     
    End Select
     
    Exit Function

End Function

Function Move_Tela_Memoria(objCondicaoPagto As ClassCondicaoPagto, Optional ByVal bComParcelas As Boolean = True) As Long
'Move os dados da tela para memória.

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    'Preenche objCondicaoPagto
    objCondicaoPagto.iCodigo = StrParaInt(Codigo.Text)
    objCondicaoPagto.sDescReduzida = DescReduzida.Text
    objCondicaoPagto.sDescricao = Descricao.Text
    
    If CmbModificador.ListIndex = -1 Then
        objCondicaoPagto.iModificador = 0
    Else
        objCondicaoPagto.iModificador = CmbModificador.ItemData(CmbModificador.ListIndex)
    End If
    
    If OptDataFixa.Value = True Then
        objCondicaoPagto.iDataFixa = CONDPAGTO_DATAFIXA
    Else
        objCondicaoPagto.iDataFixa = CONDPAGTO_NORMAL
    End If
    
    If NumeroParcelas.Text <> "" Then
        objCondicaoPagto.iNumeroParcelas = CInt(NumeroParcelas.Text)
    Else
        objCondicaoPagto.iNumeroParcelas = 0
    End If
    
    objCondicaoPagto.iDiasParaPrimeiraParcela = StrParaInt(DiasParaPrimeiraParcela.Text)
    
    objCondicaoPagto.iEmPagamento = EmPagamento.Value
    objCondicaoPagto.iEmRecebimento = EmRecebimento.Value
    objCondicaoPagto.iMensal = IIf(Intervalo(0).Value, 1, 0)
    
    'Preenche Dia do Mês ou Intervalo Parcelas
    If objCondicaoPagto.iMensal Then
        If DiaDoMes.Text <> "" Then objCondicaoPagto.iDiaDoMes = CInt(DiaDoMes.Text)
    Else 'intervalo em dias
        objCondicaoPagto.iIntervaloParcelas = StrParaInt(IntervaloParcelas.Text)
    End If

    If Len(Trim(AcrescimoFinanceiro.Text)) > 0 Then
        objCondicaoPagto.dAcrescimoFinanceiro = CDbl(AcrescimoFinanceiro.Text) / 100
    End If

    objCondicaoPagto.dTetoParcela = StrParaDbl(TetoParcela.Text)

    objCondicaoPagto.iCargoMinimo = Codigo_Extrai(CargoMinimo.Text)
    
    If FormaPagto.ListIndex = -1 Then
        objCondicaoPagto.iFormaPagamento = 0
    Else
        objCondicaoPagto.iFormaPagamento = FormaPagto.ItemData(FormaPagto.ListIndex)
    End If
    
    If bComParcelas Then
        lErro = Move_GridParcelas_Memoria(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 125184
    End If
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 125184
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154656)

    End Select

    Exit Function

End Function

Private Sub Condicoes_Adiciona(objCondicaoPagto As ClassCondicaoPagto)

    Condicoes.AddItem objCondicaoPagto.sDescReduzida
    Condicoes.ItemData(Condicoes.NewIndex) = objCondicaoPagto.iCodigo

End Sub

Private Sub Condicoes_Exclui(iCodigo As Integer)

Dim iIndice As Integer

    For iIndice = 0 To Condicoes.ListCount - 1

        If Condicoes.ItemData(iIndice) = iCodigo Then

            Condicoes.RemoveItem iIndice
            Exit For

        End If

    Next

End Sub

Private Sub HabilitaFrameIntervalo(bHabilita As Boolean)

    FIntervalo.Enabled = bHabilita
    Intervalo(0).Enabled = bHabilita
    Intervalo(1).Enabled = bHabilita
    LabelDiaDoMes.Enabled = bHabilita
    DiaDoMes.Enabled = bHabilita
    IntervaloParcelas.Enabled = bHabilita
        
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONDICAO_PAGAMENTO
    Set Form_Load_Ocx = Me
    Caption = "Condições de Pagamento"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CondicoesPagto"
    
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

Private Sub OptDataFixa_Click()

    DiasParaPrimeiraParcela.Text = ""
    CmbModificador.ListIndex = -1
    
End Sub

Private Sub UpDownDataEmissao_DownClick()
'Diminui DataEnissao em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125185

    Exit Sub

Erro_UpDownDataEmissao_DownClick:

    Select Case gErr

        Case 125185
            ExemploDataEmissao.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154657)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownDataEmissao_UpClick()
'Aumenta DataEmissao em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataEmissao_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125186

    Exit Sub

Erro_UpDownDataEmissao_UpClick:

    Select Case gErr

        Case 125186
            ExemploDataEmissao.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154658)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_DownClick()
'Diminui DataEntrega em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataEntrega, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125187

    Exit Sub

Erro_UpDownDataEntrega_DownClick:

    Select Case gErr

        Case 125187
            ExemploDataEntrega.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154659)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataEntrega_UpClick()
'Aumenta DataEntrega em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataEntrega_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataEntrega, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125188

    Exit Sub

Erro_UpDownDataEntrega_UpClick:

    Select Case gErr

        Case 125188
            ExemploDataEntrega.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154660)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataRef_DownClick()
'Diminui DataRef em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataRef_DownClick

    'Aciona rotina que diminui data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataRef, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 125189

    Exit Sub

Erro_UpDownDataRef_DownClick:

    Select Case gErr

        Case 125189
            ExemploDataRef.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154661)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataRef_UpClick()
'Aumenta DataRef em UM dia

Dim lErro As Long

On Error GoTo Erro_UpDownDataRef_UpClick

    'Aciona rotina que aumenta data em UM dia
    lErro = Data_Up_Down_Click(ExemploDataRef, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 125190

    Exit Sub

Erro_UpDownDataRef_UpClick:

    Select Case gErr

        Case 125190
            ExemploDataRef.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154662)

    End Select

    Exit Sub

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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If
End Sub


Private Sub LabelDiaDoMes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDiaDoMes, Source, X, Y)
End Sub

Private Sub LabelDiaDoMes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDiaDoMes, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame4(Opcao.SelectedItem.Index).Visible = True
        Frame4(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
'        Select Case iFrameAtual
'
'            Case TAB_Identificacao
'                Parent.HelpContextID = IDH_TIPOS_FORN_IDENT
'
'            Case TAB_DadosFinanceiros
'                Parent.HelpContextID = IDH_TIPOS_FORN_DADOS_FIN
'
'        End Select
        
    End If

End Sub

Private Function Inicializa_GridParcelas() As Long

Dim iIndice As Integer

    Set objGridParcelas = New AdmGrid

    'tela em questão
    Set objGridParcelas.objForm = Me

    'titulos do grid
    objGridParcelas.colColuna.Add ("")
    objGridParcelas.colColuna.Add ("Data Base")
    objGridParcelas.colColuna.Add ("Modificador")
    objGridParcelas.colColuna.Add ("Intervalo")
    objGridParcelas.colColuna.Add ("Dia / Dias")
    objGridParcelas.colColuna.Add ("% Recebto.")

   'campos de edição do grid
    objGridParcelas.colCampo.Add (ParcDataBase.Name)
    objGridParcelas.colCampo.Add (ParcModif.Name)
    objGridParcelas.colCampo.Add (ParcIntervalo.Name)
    objGridParcelas.colCampo.Add (ParcDias.Name)
    objGridParcelas.colCampo.Add (ParcPercReceb.Name)
    
    'Colunas do Grid
    iGrid_ParcDataBase_Col = 1
    iGrid_ParcModif_Col = 2
    iGrid_ParcIntervalo_Col = 3
    iGrid_ParcDias_Col = 4
    iGrid_ParcPercReceb_Col = 5

    objGridParcelas.objGrid = GridParcelas
    
    'tulio 9/5/02
    GridParcelas.Rows = NUM_MAX_PARCELAS_CONDPAGTO
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGridParcelas.iLinhasVisiveis = 6

    GridParcelas.ColWidth(0) = 300

    objGridParcelas.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridParcelas.iIncluirHScroll = GRID_INCLUIR_HSCROLL
    
    objGridParcelas.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    Call Grid_Inicializa(objGridParcelas)
    
    Inicializa_GridParcelas = SUCESSO

End Function

Private Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim iLinhasExistentes As Integer
    
    iLinhasExistentes = objGridParcelas.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

    If iLinhasExistentes > objGridParcelas.iLinhasExistentes Then Call Calcular_Total_Parcelas

End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)
    
End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        Select Case GridParcelas.Col
    
            Case iGrid_ParcDataBase_Col

                lErro = Saida_Celula_ParcDataBase(objGridInt)
                If lErro <> SUCESSO Then gError 125191

            Case iGrid_ParcIntervalo_Col

                lErro = Saida_Celula_ParcIntervalo(objGridInt)
                If lErro <> SUCESSO Then gError 125192

            Case iGrid_ParcDias_Col

                lErro = Saida_Celula_ParcDias(objGridInt)
                If lErro <> SUCESSO Then gError 125193

            Case iGrid_ParcModif_Col

                lErro = Saida_Celula_ParcModif(objGridInt)
                If lErro <> SUCESSO Then gError 125194

            Case iGrid_ParcPercReceb_Col

                lErro = Saida_Celula_ParcPercReceb(objGridInt)
                If lErro <> SUCESSO Then gError 125195

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 125196

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 125191 To 125195
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 125196

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154663)

    End Select

    Exit Function

End Function

Private Sub ParcDataBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcDataBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcDataBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcDataBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcDataBase
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParcIntervalo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcIntervalo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcIntervalo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcIntervalo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcIntervalo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParcDias_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcDias_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcDias_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcDias_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcDias
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParcModif_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcModif_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcModif_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcModif_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcModif
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ParcPercReceb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ParcPercReceb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ParcPercReceb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ParcPercReceb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcPercReceb
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_ParcDataBase(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcDataBase

    Set objGridInt.objControle = ParcDataBase

    If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
        
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ParcDataBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcDataBase:

    Saida_Celula_ParcDataBase = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154664)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcIntervalo(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcIntervalo

    Set objGridInt.objControle = ParcIntervalo

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ParcIntervalo = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcIntervalo:

    Saida_Celula_ParcIntervalo = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154665)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcDias(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcDias

    Set objGridInt.objControle = ParcDias

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 22348
    
    Saida_Celula_ParcDias = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcDias:

    Saida_Celula_ParcDias = gErr

    Select Case gErr

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154666)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcModif(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcModif

    Set objGridInt.objControle = ParcModif
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 22348
    
    Saida_Celula_ParcModif = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcModif:

    Saida_Celula_ParcModif = Err

    Select Case Err

        Case 22348
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154667)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ParcPercReceb(objGridInt As AdmGrid) As Long
'faz a critica da celula conta do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ParcPercReceb

    Set objGridInt.objControle = ParcPercReceb

    If Len(Trim(ParcPercReceb.ClipText)) <> 0 Then
    
        lErro = Porcentagem_Critica(ParcPercReceb.ClipText)
        If lErro <> SUCESSO Then gError 124055
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 22348
    
    Call Calcular_Total_Parcelas
    
    Saida_Celula_ParcPercReceb = SUCESSO

    Exit Function

Erro_Saida_Celula_ParcPercReceb:

    Saida_Celula_ParcPercReceb = gErr

    Select Case gErr

        Case 22348, 124055
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 154668)

    End Select

    Exit Function

End Function

Function CondicaoPagto_Gera_Parcelas(ByVal objCondicaoPagto As ClassCondicaoPagto) As Long
'Gera objCondicaoPagto.colParcelas a partir de dados de objCondicaoPagto

Dim iIndice As Integer, iResto As Integer, dValorParcela As Double, iSeq As Integer
Dim dValor As Double, iNumParcelas As Integer, objParc As ClassCondicaoPagtoParc, iDiasAcum As Integer

On Error GoTo Erro_CondicaoPagto_Gera_Parcelas

    dValor = 1000000
    iNumParcelas = objCondicaoPagto.iNumeroParcelas
    
    If iNumParcelas <> 1 Then
    
        'Calcula o resto da divisão inteira
        iResto = Resto(dValor * 100, iNumParcelas)
    
        If (iResto <> 0) Then
            dValorParcela = (dValor * 100) / iNumParcelas
            dValorParcela = Int(dValorParcela) / 100
        Else
            dValorParcela = dValor / iNumParcelas
        End If
        
    Else
    
        dValorParcela = dValor
        iResto = 0
        
    End If

    'Acrescentar valores das parcelas na coleção
    For iIndice = 1 To iNumParcelas - iResto

        Set objParc = New ClassCondicaoPagtoParc
        
        iSeq = iSeq + 1
        
        objParc.iSeq = iSeq
        objParc.dPercReceb = dValorParcela / 1000000
        
        objCondicaoPagto.colParcelas.Add objParc

    Next

    'Soma 0.01 ao Valor da Parcela
    dValorParcela = dValorParcela + 0.01

    'Se a divisão não foi exata acrescentar as "iResto" últimas parcelas adicionadas de 0.01
    For iIndice = 1 To iResto

        Set objParc = New ClassCondicaoPagtoParc
        
        iSeq = iSeq + 1
        
        objParc.iSeq = iSeq
        objParc.dPercReceb = dValorParcela / 1000000
        
        objCondicaoPagto.colParcelas.Add objParc

    Next

    'completa os outros atributos
    For Each objParc In objCondicaoPagto.colParcelas
        
        If objParc.iSeq = 1 Then
            objParc.iTipoDataBase = IIf(objCondicaoPagto.iDataFixa = 0, CONDPAGTO_TIPODATABASE_EMISSAO, CONDPAGTO_TIPODATABASE_DATAFIXA)
            objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAS
            objParc.iDias = objCondicaoPagto.iDiasParaPrimeiraParcela
        Else
            If objCondicaoPagto.iMensal Then
                objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_PARCANTERIOR
                objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAMES
                objParc.iDias = objCondicaoPagto.iDiaDoMes
            Else
                objParc.iTipoIntervalo = CONDPAGTO_TIPOINTERVALO_DIAS
                If objCondicaoPagto.iDataFixa = 0 Then
                    objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_EMISSAO
                    objParc.iDias = iDiasAcum + objCondicaoPagto.iIntervaloParcelas
                Else
                    objParc.iTipoDataBase = CONDPAGTO_TIPODATABASE_PARCANTERIOR
                    objParc.iDias = objCondicaoPagto.iIntervaloParcelas
                End If
            End If
        End If
        objParc.iModificador = objCondicaoPagto.iModificador
        iDiasAcum = objParc.iDias
        
    Next
    
    CondicaoPagto_Gera_Parcelas = SUCESSO
     
    Exit Function
    
Erro_CondicaoPagto_Gera_Parcelas:

    CondicaoPagto_Gera_Parcelas = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154669)
     
    End Select
     
    Exit Function

End Function

Sub Exemplo_ValoresDefault()

    ExemploValor.Text = Format(1000, "Standard")
    Call DateParaMasked(ExemploDataRef, gdtDataAtual)
    Call DateParaMasked(ExemploDataEmissao, gdtDataAtual)
    Call DateParaMasked(ExemploDataEntrega, gdtDataAtual)

End Sub

Private Function Inicializa_GridResultado() As Long

Dim iIndice As Integer

    Set objGridResultado = New AdmGrid

    'tela em questão
    Set objGridResultado.objForm = Me

    'titulos do grid
    objGridResultado.colColuna.Add ("")
    objGridResultado.colColuna.Add ("Vencimento")
    objGridResultado.colColuna.Add ("Valor")

   'campos de edição do grid
    objGridResultado.colCampo.Add (ResultadoVcto.Name)
    objGridResultado.colCampo.Add (ResultadoValor.Name)
    
    'Colunas do Grid
    iGrid_ResultadoVcto_Col = 1
    iGrid_ResultadoValor_Col = 2

    objGridResultado.objGrid = GridResultado
    
    'tulio 9/5/02
    GridResultado.Rows = NUM_MAX_PARCELAS_CONDPAGTO
    
    'linhas visiveis do grid sem contar com as linhas fixas
    objGridResultado.iLinhasVisiveis = 6

    GridResultado.ColWidth(0) = 300

    objGridResultado.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridResultado.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGridResultado)
    
    Inicializa_GridResultado = SUCESSO

End Function

Private Sub GridResultado_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridResultado, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridResultado, iAlterado)
    End If

End Sub

Private Sub GridResultado_GotFocus()

    Call Grid_Recebe_Foco(objGridResultado)

End Sub

Private Sub GridResultado_EnterCell()

    Call Grid_Entrada_Celula(objGridResultado, iAlterado)

End Sub

Private Sub GridResultado_LeaveCell()

    Call Saida_Celula(objGridResultado)

End Sub

Private Sub GridResultado_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridResultado)

End Sub

Private Sub GridResultado_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridResultado, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridResultado, iAlterado)
    End If

End Sub

Private Sub GridResultado_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridResultado)
    
End Sub

Private Sub GridResultado_RowColChange()

    Call Grid_RowColChange(objGridResultado)

End Sub

Private Sub GridResultado_Scroll()

    Call Grid_Scroll(objGridResultado)

End Sub

Sub Calcular_Total_Parcelas()
'Calcula o somatório dos percentuais de parcelas à receber

Dim dPercReceb As Double
Dim iIndice As Integer
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        dPercReceb = dPercReceb + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ParcPercReceb_Col))
    Next
    
    'Imprime na Tela a Porcentagem contida no grid
    ParcPercTotal.Caption = Format(dPercReceb / 100, "Percent")
   
    Exit Sub

End Sub

Function Exibe_Dados_Parcelas(objCondicaoPagto As ClassCondicaoPagto) As Long

Dim iIndice As Integer
Dim lErro As Long, objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_Exibe_Dados_Parcelas

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridParcelas)

    'trazer dados do grid de parcelas
    For Each objParc In objCondicaoPagto.colParcelas
    
        iIndice = iIndice + 1
        
        GridParcelas.TextMatrix(iIndice, iGrid_ParcDataBase_Col) = ParcDataBase.List(objParc.iTipoDataBase - 1)
        GridParcelas.TextMatrix(iIndice, iGrid_ParcIntervalo_Col) = ParcIntervalo.List(objParc.iTipoIntervalo - 1)
        GridParcelas.TextMatrix(iIndice, iGrid_ParcDias_Col) = objParc.iDias
        GridParcelas.TextMatrix(iIndice, iGrid_ParcModif_Col) = ParcModif.List(objParc.iModificador)
        GridParcelas.TextMatrix(iIndice, iGrid_ParcPercReceb_Col) = Format(objParc.dPercReceb * 100, "0.0#####")
        
    Next
    
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.colParcelas.Count
    
    Call Calcular_Total_Parcelas
    
    Exibe_Dados_Parcelas = SUCESSO
     
    Exit Function
    
Erro_Exibe_Dados_Parcelas:

    Exibe_Dados_Parcelas = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154670)
     
    End Select
     
    Exit Function

End Function

Function Exibe_Dados_Exemplo(objCondicaoPagto As ClassCondicaoPagto) As Long

Dim iIndice As Integer
Dim lErro As Long, objParc As ClassCondicaoPagtoParc

On Error GoTo Erro_Exibe_Dados_Exemplo

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridResultado)

    'trazer dados do grid de parcelas
    For Each objParc In objCondicaoPagto.colParcelas
    
        iIndice = iIndice + 1
        
        GridResultado.TextMatrix(iIndice, iGrid_ResultadoVcto_Col) = Format(objParc.dtVencimento, "dd/mm/yyyy")
        GridResultado.TextMatrix(iIndice, iGrid_ResultadoValor_Col) = Format(objParc.dValor, "standard")
        
    Next
    
    objGridResultado.iLinhasExistentes = iIndice
    
    Exibe_Dados_Exemplo = SUCESSO
     
    Exit Function
    
Erro_Exibe_Dados_Exemplo:

    Exibe_Dados_Exemplo = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 154671)
     
    End Select
     
    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case ParcModif.Name

            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ParcDataBase_Col))) <> 0 Then
                ParcModif.Enabled = True
            Else
                ParcModif.Enabled = False
            End If

        Case ParcIntervalo.Name
            
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ParcDataBase_Col))) <> 0 Then
                ParcIntervalo.Enabled = True
            Else
                ParcIntervalo.Enabled = False
            End If

        Case ParcDias.Name
        
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ParcDataBase_Col))) <> 0 Then
                ParcDias.Enabled = True
            Else
                ParcDias.Enabled = False
            End If
        
        Case ParcPercReceb.Name
        
            If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ParcDataBase_Col))) <> 0 Then
                ParcPercReceb.Enabled = True
            Else
                ParcPercReceb.Enabled = False
            End If

    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 154672)

    End Select

    Exit Sub

End Sub

Private Sub CargoMinimo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CargoMinimo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CargoMinimo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CargoMinimo_Validate

    If CargoMinimo.Text <> "" Then
    
        'Valida o tipo de relacionamento selecionado pelo cliente
        lErro = CF("CamposGenericos_Validate", CAMPOSGENERICOS_CARGO_VENDEDOR, CargoMinimo, "AVISO_CRIAR_CARGO_VENDEDOR")
        If lErro <> SUCESSO Then gError 195867
    
    End If
    
    Exit Sub

Erro_CargoMinimo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 195867
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195868)

    End Select

End Sub

Private Sub TetoParcela_Change()
    iAlterado = MARCADO
End Sub

Private Sub TetoParcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TetoParcela_Validate

    If Len(TetoParcela.Text) > 0 Then
    
        'Faza critica do valor do saldo Inicial
        lErro = Valor_Critica(TetoParcela.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
    End If

    Exit Sub

Erro_TetoParcela_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155331)
            
    End Select
        
    Exit Sub

End Sub

