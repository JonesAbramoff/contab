VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ExcecoesICMSOcx 
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9510
   Begin VB.ComboBox ICMSMotivo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "ExcecoesICMSOcx.ctx":0000
      Left            =   2625
      List            =   "ExcecoesICMSOcx.ctx":0002
      TabIndex        =   69
      Text            =   "ICMSMotivo"
      Top             =   6660
      Width           =   4770
   End
   Begin VB.TextBox cBenef 
      Height          =   288
      Left            =   2625
      MaxLength       =   10
      TabIndex        =   68
      Top             =   6285
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   15
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ExcecoesICMSOcx.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ExcecoesICMSOcx.ctx":015E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ExcecoesICMSOcx.ctx":02E8
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ExcecoesICMSOcx.ctx":081A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério"
      Height          =   2505
      Left            =   45
      TabIndex        =   37
      Top             =   525
      Width           =   9330
      Begin VB.Frame Frame8 
         Caption         =   "Aplicação"
         Height          =   450
         Left            =   75
         TabIndex        =   64
         Top             =   465
         Width           =   9090
         Begin VB.OptionButton TipoAplicacao 
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
            Height          =   195
            Index           =   0
            Left            =   7830
            TabIndex        =   5
            Top             =   195
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton TipoAplicacao 
            Caption         =   "Não aplicável para consumidor final"
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
            Left            =   3825
            TabIndex        =   4
            Top             =   195
            Width           =   3450
         End
         Begin VB.OptionButton TipoAplicacao 
            Caption         =   "Somente para consumidor final"
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
            Left            =   480
            TabIndex        =   3
            Top             =   195
            Width           =   3000
         End
      End
      Begin VB.ComboBox GrupoOrigemMercadoria 
         Height          =   315
         ItemData        =   "ExcecoesICMSOcx.ctx":0998
         Left            =   2685
         List            =   "ExcecoesICMSOcx.ctx":09A5
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2145
         Width           =   6465
      End
      Begin VB.OptionButton OptCliForn 
         Caption         =   "Fornecedor"
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
         Index           =   1
         Left            =   4620
         TabIndex        =   10
         Top             =   1425
         Width           =   1920
      End
      Begin VB.OptionButton OptCliForn 
         Caption         =   "Cliente"
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
         Index           =   0
         Left            =   2715
         TabIndex        =   9
         Top             =   1440
         Value           =   -1  'True
         Width           =   1920
      End
      Begin VB.ComboBox Origem 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Top             =   165
         Width           =   888
      End
      Begin VB.ComboBox Destino 
         Height          =   315
         Left            =   5550
         TabIndex        =   2
         Text            =   "Destino"
         Top             =   165
         Width           =   888
      End
      Begin VB.Frame Frame4 
         Caption         =   "Produtos"
         Height          =   510
         Left            =   90
         TabIndex        =   38
         Top             =   900
         Width           =   9075
         Begin VB.CheckBox TodosProdutos 
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
            Height          =   252
            Left            =   480
            TabIndex        =   6
            Top             =   195
            Width           =   1110
         End
         Begin VB.ComboBox ItemCategoriaProduto 
            Height          =   315
            Left            =   6615
            TabIndex        =   8
            Top             =   150
            Width           =   2436
         End
         Begin VB.ComboBox CategoriaProduto 
            Height          =   315
            Left            =   2640
            TabIndex        =   7
            Top             =   150
            Width           =   2820
         End
         Begin VB.Label Label5 
            Caption         =   "Categoria:"
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
            Left            =   1710
            TabIndex        =   40
            Top             =   180
            Width           =   930
         End
         Begin VB.Label Label2 
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
            Left            =   6045
            TabIndex        =   39
            Top             =   180
            Width           =   510
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Clientes"
         Height          =   510
         Index           =   0
         Left            =   105
         TabIndex        =   41
         Top             =   1620
         Width           =   9075
         Begin VB.CheckBox TodosClientes 
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
            Height          =   252
            Left            =   465
            TabIndex        =   11
            Top             =   195
            Width           =   1020
         End
         Begin VB.ComboBox CategoriaCliente 
            Height          =   315
            Left            =   2595
            TabIndex        =   12
            Top             =   135
            Width           =   2820
         End
         Begin VB.ComboBox ItemCategoriaCliente 
            Height          =   315
            Left            =   6570
            TabIndex        =   13
            Top             =   135
            Width           =   2475
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5985
            TabIndex        =   43
            Top             =   180
            Width           =   510
         End
         Begin VB.Label Label4 
            Caption         =   "Categoria:"
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
            Left            =   1635
            TabIndex        =   42
            Top             =   165
            Width           =   930
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Fornecedores"
         Height          =   495
         Index           =   1
         Left            =   105
         TabIndex        =   49
         Top             =   1620
         Visible         =   0   'False
         Width           =   9075
         Begin VB.ComboBox ItemCategoriaFornecedor 
            Height          =   315
            Left            =   6570
            TabIndex        =   16
            Top             =   135
            Width           =   2475
         End
         Begin VB.ComboBox CategoriaFornecedor 
            Height          =   315
            Left            =   2595
            TabIndex        =   15
            Top             =   135
            Width           =   2820
         End
         Begin VB.CheckBox TodosFornecedores 
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
            Height          =   252
            Left            =   465
            TabIndex        =   14
            Top             =   195
            Width           =   1020
         End
         Begin VB.Label Label9 
            Caption         =   "Categoria:"
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
            Left            =   1635
            TabIndex        =   51
            Top             =   165
            Width           =   930
         End
         Begin VB.Label Label8 
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
            Left            =   5985
            TabIndex        =   50
            Top             =   180
            Width           =   510
         End
      End
      Begin VB.Label LabelGrupoOrigemMercadoria 
         Caption         =   "Origem da Mercadoria:"
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
         Left            =   630
         TabIndex        =   61
         Top             =   2175
         Width           =   1980
      End
      Begin VB.Label DescrEstDest 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1935
         TabIndex        =   48
         Top             =   180
         Width           =   2685
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   47
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "Destino:"
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
         Left            =   4710
         TabIndex        =   45
         Top             =   225
         Width           =   720
      End
      Begin VB.Label DescrEstado 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6435
         TabIndex        =   44
         Top             =   180
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tratamento"
      Height          =   3195
      Index           =   1
      Left            =   45
      TabIndex        =   36
      Top             =   3015
      Width           =   9330
      Begin VB.Frame Frame7 
         Caption         =   "ICMS relativo ao Fundo de Combate à Pobreza (FCP) na UF de destino"
         Height          =   570
         Left            =   195
         TabIndex        =   62
         Top             =   2565
         Width           =   9015
         Begin MSMask.MaskEdBox ICMSPercFCP 
            Height          =   285
            Left            =   1650
            TabIndex        =   30
            Top             =   225
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label11 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   600
            TabIndex        =   63
            Top             =   270
            Width           =   990
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Situação tributária"
         Height          =   990
         Left            =   180
         TabIndex        =   60
         Top             =   195
         Width           =   9015
         Begin VB.ComboBox TipoTributacaoSimples 
            Height          =   315
            Left            =   1740
            TabIndex        =   19
            Top             =   570
            Width           =   7215
         End
         Begin VB.ComboBox TipoTributacao 
            Height          =   315
            Left            =   1740
            TabIndex        =   18
            Top             =   225
            Width           =   7215
         End
         Begin VB.Label Label7 
            Caption         =   "Simples (CSOSN):"
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
            Height          =   225
            Left            =   180
            TabIndex        =   66
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Normal (CST):"
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
            Height          =   225
            Left            =   510
            TabIndex        =   65
            Top             =   255
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ICMS Próprio"
         Height          =   525
         Left            =   4560
         TabIndex        =   57
         Top             =   1215
         Width           =   4650
         Begin MSMask.MaskEdBox RedBaseCalculo 
            Height          =   285
            Left            =   2160
            TabIndex        =   21
            Top             =   165
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label LabelRedBase 
            AutoSize        =   -1  'True
            Caption         =   "Red. Base Cálculo:"
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
            Left            =   480
            TabIndex        =   58
            Top             =   210
            Width           =   1605
         End
      End
      Begin VB.Frame FrameICMSProprio 
         Caption         =   "ICMS do Destino"
         Height          =   525
         Left            =   195
         TabIndex        =   54
         Top             =   1215
         Width           =   4320
         Begin MSMask.MaskEdBox Aliquota 
            Height          =   285
            Left            =   1710
            TabIndex        =   20
            Top             =   180
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label LabelAliquota 
            AutoSize        =   -1  'True
            Caption         =   "Alíquota:"
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
            Left            =   870
            TabIndex        =   55
            Top             =   225
            Width           =   750
         End
      End
      Begin VB.Frame FrameICMSST 
         Caption         =   "Substituição Tributária"
         Height          =   810
         Left            =   195
         TabIndex        =   52
         Top             =   1755
         Width           =   9015
         Begin VB.CheckBox ICMSSTBaseDupla 
            Caption         =   "Cálculo por ""Base Dupla"" a partir de"
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
            Left            =   4095
            TabIndex        =   27
            ToolTipText     =   $"ExcecoesICMSOcx.ctx":09F6
            Top             =   510
            Width           =   3540
         End
         Begin VB.OptionButton PautaOuMargem 
            Caption         =   "Option1"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   5205
            TabIndex        =   24
            Top             =   210
            Width           =   225
         End
         Begin VB.OptionButton PautaOuMargem 
            Caption         =   "Option1"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   495
            TabIndex        =   22
            Top             =   195
            Value           =   -1  'True
            Width           =   240
         End
         Begin MSMask.MaskEdBox MargemLucroSubst 
            Height          =   285
            Left            =   2640
            TabIndex        =   23
            Top             =   180
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ValorPauta 
            Height          =   300
            Left            =   6510
            TabIndex        =   25
            Top             =   180
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   529
            _Version        =   393216
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
         Begin MSMask.MaskEdBox RedBaseCalculoSubst 
            Height          =   285
            Left            =   2640
            TabIndex        =   26
            Top             =   480
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   15
            Format          =   "##0.#0\%"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownICMSSTBaseDuplaIni 
            Height          =   300
            Left            =   8670
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   465
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
         End
         Begin MSMask.MaskEdBox ICMSSTBaseDuplaIni 
            Height          =   300
            Left            =   7650
            TabIndex        =   28
            Top             =   465
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin VB.Label LabelRedBaseSubst 
            AutoSize        =   -1  'True
            Caption         =   "Red. Base Cálculo:"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   960
            TabIndex        =   59
            Top             =   525
            Width           =   1605
         End
         Begin VB.Label LabelPauta 
            AutoSize        =   -1  'True
            Caption         =   "Pauta (R$):"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   5445
            TabIndex        =   56
            Top             =   210
            Width           =   990
         End
         Begin VB.Label LabelMarg 
            AutoSize        =   -1  'True
            Caption         =   "Margem de Lucro (%):"
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   735
            TabIndex        =   53
            Top             =   210
            Width           =   1845
         End
      End
   End
   Begin VB.TextBox Fundamentacao 
      Height          =   288
      Left            =   1695
      TabIndex        =   0
      Top             =   135
      Width           =   5460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Motivo de Desoneração:"
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
      Index           =   68
      Left            =   195
      TabIndex        =   70
      Top             =   6690
      Width           =   2085
   End
   Begin VB.Label Label12 
      Caption         =   "Código do Benefício Fiscal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   67
      Top             =   6300
      Width           =   2475
   End
   Begin VB.Label LabelFundamentacao 
      Caption         =   "Fundamentação:"
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
      Height          =   240
      Left            =   255
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   46
      Top             =   165
      Width           =   1440
   End
End
Attribute VB_Name = "ExcecoesICMSOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Const TIPO_PAUTA = 1
Const TIPO_MARGEM = 0

Dim iTipoAnt As Integer
Dim sUFDestAnt As String

Private WithEvents objEventoExcecoesICMS As AdmEvento
Attribute objEventoExcecoesICMS.VB_VarHelpID = -1

Private Sub Traz_Excecao_Tela(objExcecoesICMS As ClassICMSExcecao)
'Traz a Exceção a Tela

Dim lErro As Long
Dim iIndice As Integer, iCodigo As Integer
Dim objEstado As New ClassEstado
Dim objTipoTribICMS As New ClassTipoTribICMS
Dim bCancel As Boolean

On Error GoTo Erro_Traz_Excecao_Tela

    lErro = CF("ICMSExcecao_Le", objExcecoesICMS)
    If lErro <> SUCESSO Then gError 198582
    
    'Preenche a Fundamentação na Tela
    Fundamentacao.Text = objExcecoesICMS.sFundamentacao

    TodosProdutos.Value = vbUnchecked
    TodosClientes.Value = vbUnchecked
    TodosFornecedores.Value = vbUnchecked
    
    'Se a Categoria do Produto estiver Preenchida
    If objExcecoesICMS.sCategoriaProduto <> "" Then
        
        'Coloca a categoria na Tela e chama o Validate da CategoriaProduto
        CategoriaProduto.Text = objExcecoesICMS.sCategoriaProduto
        Call CategoriaProduto_Validate(bCancel)
        
        'Coloca a item da categoria na Tela e chama o Validate do ItemCategoriaProduto
        ItemCategoriaProduto.Text = objExcecoesICMS.sCategoriaProdutoItem
        Call ItemCategoriaProduto_Validate(bSGECancelDummy)
    
    Else
        'Senão marca todos
        TodosProdutos.Value = 1
    End If
    
    TipoAplicacao(objExcecoesICMS.iTipoAplicacao).Value = True
    
    '##############################################
    'Alterado por Wagner 29/09/05
    OptCliForn(objExcecoesICMS.iTipoCliForn).Value = True
    Call OptCliForn_Click(objExcecoesICMS.iTipoCliForn)
    
    If objExcecoesICMS.iTipoCliForn = ICMSEXCECOES_TIPOCLIFORN_CLIENTE Then

        'Se a Categoria do Cliente estiver Preenchida
        If objExcecoesICMS.sCategoriaCliente <> "" Then
            
            'Coloca na Tela e Chama o Validate de CategoriaCliente
            CategoriaCliente.Text = objExcecoesICMS.sCategoriaCliente
            Call CategoriaCliente_Validate(bCancel)
            
            'Coloca o Item da Categoria na tela e chama o Validate
            ItemCategoriaCliente.Text = objExcecoesICMS.sCategoriaClienteItem
            Call ItemCategoriaCliente_Validate(bSGECancelDummy)
            
        Else
            'Senão marca todos
            TodosClientes.Value = 1
        End If
        
    Else
        'Se a Categoria do Fornecedor estiver Preenchida
        If objExcecoesICMS.sCategoriaFornecedor <> "" Then
            
            'Coloca na Tela e Chama o Validate de CategoriaFornecedor
            CategoriaFornecedor.Text = objExcecoesICMS.sCategoriaFornecedor
            Call CategoriaFornecedor_Validate(bCancel)
            
            'Coloca o Item da Categoria na tela e chama o Validate
            ItemCategoriaFornecedor.Text = objExcecoesICMS.sCategoriaFornecedorItem
            Call ItemCategoriaFornecedor_Validate(bSGECancelDummy)
            
        Else
            'Senão marca todos
            TodosFornecedores.Value = 1
        End If
    End If
    '##############################################
    
    'Preenche destino e chama lostFocus
    Destino.Text = objExcecoesICMS.sEstadoDestino
    Call Destino_Validate(bSGECancelDummy)

    'William 30/04/01
    'Preenche origem e chama validate
    Origem.Text = objExcecoesICMS.sEstadoOrigem
    Call Origem_Validate(bSGECancelDummy)

    Call Combo_Seleciona_ItemData(GrupoOrigemMercadoria, objExcecoesICMS.iGrupoOrigemMercadoria)

    'pesquisa o tipo na lista e seleciona-o
    TipoTributacao.Text = CStr(objExcecoesICMS.iTipo)
    Call TipoTributacao_Validate(bSGECancelDummy)

    'pesquisa o tipo na lista e seleciona-o
    TipoTributacaoSimples.Text = CStr(objExcecoesICMS.iTipoSimples)
    Call TipoTributacaoSimples_Validate(bSGECancelDummy)

    If Aliquota.Enabled = True Then Aliquota.Text = CStr(objExcecoesICMS.dAliquota * 100)
    If RedBaseCalculo.Enabled = True Then RedBaseCalculo.Text = CStr(objExcecoesICMS.dPercRedBaseCalculo * 100)
    
    If objExcecoesICMS.iUsaPauta = MARCADO Then
        PautaOuMargem(TIPO_PAUTA).Value = True
        Call PautaOuMargem_Click(TIPO_PAUTA)
    Else
        PautaOuMargem(TIPO_MARGEM).Value = True
        Call PautaOuMargem_Click(TIPO_MARGEM)
    End If
    If RedBaseCalculoSubst.Enabled = True Then RedBaseCalculoSubst.Text = CStr(objExcecoesICMS.dPercRedBaseCalculoSubst * 100)
    
    If objExcecoesICMS.dPercMargemLucro <> 0 Then MargemLucroSubst.Text = CStr(objExcecoesICMS.dPercMargemLucro * 100)
    If objExcecoesICMS.dValorPauta <> 0 Then ValorPauta.Text = Format(objExcecoesICMS.dValorPauta, "STANDARD")

    ICMSPercFCP.Text = CStr(objExcecoesICMS.dICMSPercFCP * 100)
    
    If objExcecoesICMS.iICMSSTBaseDupla = MARCADO Then
        ICMSSTBaseDupla.Value = vbChecked
    Else
        ICMSSTBaseDupla.Value = vbUnchecked
    End If
    Call DateParaMasked(ICMSSTBaseDuplaIni, objExcecoesICMS.dtICMSSTBaseDuplaIni)
    
    cBenef.Text = objExcecoesICMS.scBenef
    
    If objExcecoesICMS.iICMSMotivo <> 0 Then
        Call Combo_Seleciona_ItemData(ICMSMotivo, objExcecoesICMS.iICMSMotivo)
    Else
        ICMSMotivo.ListIndex = -1
    End If
    
    iAlterado = 0

    Exit Sub
    
Erro_Traz_Excecao_Tela:

    Select Case gErr
    
        Case 198582

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159696)

    End Select

    Exit Sub
    
End Sub

Private Sub Aliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Aliquota_Validate

    If Len(Aliquota.Text) > 0 Then

        'Testa o valor
        lErro = Porcentagem_Critica2(Aliquota.Text)
        If lErro <> SUCESSO Then Error 21459

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_Aliquota_Validate:

    Cancel = True


    Select Case Err

        Case 21459

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159697)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaClienteItem As ClassCategoriaClienteItem
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaCliente_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se a CategoriaCliente foi preenchida
    If CategoriaCliente.ListIndex <> -1 Then

        objCategoriaCliente.sCategoria = CategoriaCliente.Text

        'Lê os dados de Itens da Categoria do Cliente
        lErro = CF("CategoriaCliente_Le_Itens", objCategoriaCliente, colCategoria)
        If lErro <> SUCESSO Then Error 33440

        ItemCategoriaCliente.Enabled = True

        'Limpa os dados de ItemCategoriaCliente
        ItemCategoriaCliente.Clear

        'Preenche ItemCategoriaCliente
        For Each objCategoriaClienteItem In colCategoria

            ItemCategoriaCliente.AddItem objCategoriaClienteItem.sItem

        Next
        TodosClientes.Value = 0
    
    Else
        
        'Senão Desablita ItemCategoriaCliente
        ItemCategoriaCliente.ListIndex = -1
        ItemCategoriaCliente.Enabled = False
    
    End If

    Exit Sub

Erro_CategoriaCliente_Click:

    Select Case Err

        Case 33440 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159698)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaCliente_Validate

    If Len(CategoriaCliente.Text) <> 0 And CategoriaCliente.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22991
        
        If lErro <> SUCESSO Then Error 22992
    
    End If
    
    'Se a CategoriaCliente estiver em branco desabilita e limpa a combo
    If Len(CategoriaCliente.Text) = 0 Then
        ItemCategoriaCliente.Enabled = False
        ItemCategoriaCliente.Clear
    End If
    
    Exit Sub

Erro_CategoriaCliente_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 22991
         
        Case 22992
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaCliente.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159699)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CategoriaProduto_Click()

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaProdutoItem As New ClassCategoriaProdutoItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaProduto_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se tiver selecionado alguma CategoriaProduto
    If CategoriaProduto.ListIndex <> -1 Then

        'Preenche o objeto com a Categoria
         objCategoriaProduto.sCategoria = CategoriaProduto.Text

        'Lê os dados de itens de categorias de produto
        lErro = CF("CategoriaProduto_Le_Itens", objCategoriaProduto, colCategoria)
        If lErro <> SUCESSO Then Error 33458

        ItemCategoriaProduto.Enabled = True
        ItemCategoriaProduto.Clear

        'Preenche ItemCategoriaProduto
        For Each objCategoriaProdutoItem In colCategoria

            ItemCategoriaProduto.AddItem (objCategoriaProdutoItem.sItem)

        Next
        
        TodosProdutos.Value = 0
    Else
    
        'Senão desabilita a ItemCategoriaProduto
        ItemCategoriaProduto.ListIndex = -1
        ItemCategoriaProduto.Enabled = False
    End If

    Exit Sub

Erro_CategoriaProduto_Click:

    Select Case Err

        Case 33458

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159700)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaProduto_Validate

    If Len(CategoriaProduto) <> 0 And CategoriaProduto.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22993
        
        If lErro <> SUCESSO Then Error 22994
        
    End If
    
    'Se estiver em Branco desabilita e limpa a combo
    If Len(CategoriaProduto) = 0 Then
        ItemCategoriaProduto.Enabled = False
        ItemCategoriaProduto.Clear
    End If
    
    Exit Sub

Erro_CategoriaProduto_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 22993 'Tratado na Rotina chamadora
        
        Case 22994
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", Err, CategoriaProduto.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159701)

    End Select

    Exit Sub

End Sub

Private Sub cBenef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Destino_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub GrupoOrigemMercadoria_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub


'William 30/04/01
Private Sub Origem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Destino_Click()

Dim lErro As Long
Dim objEstado As New ClassEstado
On Error GoTo Erro_Destino_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se tiver selecionado algum Destino
    If Destino.ListIndex <> -1 Then
        
        'Lê no BD o Estado selecionado
        objEstado.sSigla = Destino.Text

        lErro = CF("Estado_Le", objEstado)
        If lErro <> SUCESSO And lErro <> 28485 Then Error 19301

        'Se não encontrou ----> ERRO
        If lErro = 28485 Then Error 19302
        
        'Preenche com a Descricao
        DescrEstado.Caption = objEstado.sNome
        
        If sUFDestAnt <> Destino.Text Then
            If objEstado.dICMSPercFCP <> 0 Then
                ICMSPercFCP.Text = CStr(objEstado.dICMSPercFCP * 100)
            Else
                ICMSPercFCP.Text = ""
            End If
        End If
    Else
        
        'Senão limpa a combo
        DescrEstado.Caption = ""
    
    End If
    
    sUFDestAnt = Destino.Text
        
    Exit Sub
    
Erro_Destino_Click:

    Select Case Err

        Case 19301 'Tratado na rotina chamada
        
        Case 19302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Destino.Text)
            Destino.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159702)

    End Select

    Exit Sub

End Sub

'William 30/04/01
Private Sub Origem_Click()

Dim lErro As Long
Dim objEstado As New ClassEstado

On Error GoTo Erro_Origem_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Se tiver selecionado alguma Origem
    If Origem.ListIndex <> -1 Then
        
        'Lê no BD o Estado selecionado
        objEstado.sSigla = Origem.Text

        lErro = CF("Estado_Le", objEstado)
        If lErro <> SUCESSO And lErro <> 28485 Then Error 19301

        'Se não encontrou ----> ERRO
        If lErro = 28485 Then Error 19302
        
        'Preenche com a Descricao
        DescrEstDest.Caption = objEstado.sNome
    Else
        
        'Senão limpa a combo
        DescrEstDest.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Origem_Click:

    Select Case Err

        Case 19301 'Tratado na rotina chamada
        
        Case 19302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Origem.Text)
            Origem.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159703)

    End Select

    Exit Sub

End Sub


Public Sub Form_Unload(Cancel As Integer)

    Set objEventoExcecoesICMS = Nothing

End Sub

Private Sub ItemCategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaProduto_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub LabelFundamentacao_Click()

Dim lErro As Long
Dim objExcecoesICMS As New ClassICMSExcecao
Dim colSelecao As New Collection

On Error GoTo Erro_LabelFundamentacao_Click
    
    'Chama a lista de ICMS
    Call Chama_Tela("ExcecoesICMSLista", colSelecao, objExcecoesICMS, objEventoExcecoesICMS)

    Exit Sub

Erro_LabelFundamentacao_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159704)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcecoesICMS_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objExcecoesICMS As ClassICMSExcecao

On Error GoTo Erro_objEventoExcecoesICMS_evSelecao

    Set objExcecoesICMS = obj1
    
    'Preenche a Tela com a Exceção Icms
    Call Traz_Excecao_Tela(objExcecoesICMS)

    Me.Show

    Exit Sub

Erro_objEventoExcecoesICMS_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159705)

    End Select

    Exit Sub

End Sub

Function Verifica_Identificacao_Preenchida() As Long
'verifica se todos os dados necessarios p/identificacao de uma excecao foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Verifica_Identificacao_Preenchida
    
    'Testa se a Fundamentação está preenchida
    If Len(Fundamentacao.Text) = 0 Then gError 21702

    'Testa se destino foi preenchido
    If Len(Destino.Text) = 0 Then gError 21460
    
    'Testa se TodosProdutos está marcado
    If TodosProdutos.Value = 0 Then

        'Testa se Categoria do produto está preenchida
        If Len(CategoriaProduto.Text) = 0 Then gError 21461

        'Testa se Valor da Categoria do produto está preenchida
        If Len(ItemCategoriaProduto.Text) = 0 Then gError 21462

    End If

    '########################################
    'Alterado por Wagner
    If OptCliForn(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then

        'Testa se TodosClientes está marcado
        If TodosClientes.Value = vbUnchecked Then
    
            'Testa se Categoria do cliente está preenchida
            If Len(CategoriaCliente.Text) = 0 Then gError 21463
    
            'Testa se Valor da Categoria do cliente está preenchida
            If Len(ItemCategoriaCliente.Text) = 0 Then gError 21464
    
        End If
        
    Else
    
        'Testa se TodosClientes está marcado
        If TodosFornecedores.Value = vbUnchecked Then
    
            'Testa se Categoria do cliente está preenchida
            If Len(CategoriaFornecedor.Text) = 0 Then gError 140405
    
            'Testa se Valor da Categoria do cliente está preenchida
            If Len(ItemCategoriaFornecedor.Text) = 0 Then gError 140406
    
        End If
    
    End If
    '########################################
    
    Verifica_Identificacao_Preenchida = SUCESSO
     
    Exit Function
    
Erro_Verifica_Identificacao_Preenchida:

     Verifica_Identificacao_Preenchida = gErr
     
     Select Case gErr
          
        Case 21702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FUNDAMENTACAO_NAO_PREENCHIDA", gErr)

        Case 21460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_ESTADO_NAO_PREENCHIDA", gErr)

        Case 21461
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTO_NAO_INFORMADA", gErr)

        Case 21462
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_NAO_INFORMADO1", gErr)

        Case 21463
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_NAO_PREENCHIDA", gErr)

        Case 21464
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_CLIENTE_ITEM_NAO_PREENCHIDA", gErr)

        '#########################################
        'Inserido por Wagner 29/09/05
        Case 140405
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_FORNECEDOR_NAO_PREENCHIDA", gErr)

        Case 140406
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_FORNECEDOR_ITEM_NAO_PREENCHIDA", gErr)
        '#########################################

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159706)
     
     End Select
     
     Exit Function

End Function

Private Function Move_Identificacao_Memoria(objICMSExcecoes As ClassICMSExcecao) As Long
'Move a tela para objICMSExcecoes
Dim iIndice As Integer

    objICMSExcecoes.sFundamentacao = Fundamentacao.Text
    objICMSExcecoes.sEstadoDestino = Destino.Text
    objICMSExcecoes.sCategoriaProduto = CategoriaProduto.Text
    objICMSExcecoes.sCategoriaProdutoItem = ItemCategoriaProduto.Text
    objICMSExcecoes.sCategoriaCliente = CategoriaCliente.Text
    objICMSExcecoes.sCategoriaClienteItem = ItemCategoriaCliente.Text
    
    '######################################
    'Inserido por Wagner 29/09/05
    objICMSExcecoes.sCategoriaFornecedor = CategoriaFornecedor.Text
    objICMSExcecoes.sCategoriaFornecedorItem = ItemCategoriaFornecedor.Text
    
    If OptCliForn(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then
    
        objICMSExcecoes.iTipoCliForn = ICMSEXCECOES_TIPOCLIFORN_CLIENTE
    
    Else
    
        objICMSExcecoes.iTipoCliForn = ICMSEXCECOES_TIPOCLIFORN_FORNECEDOR
    
    End If
    '######################################
    
    'William 30/04/01
    If Len(Trim(Origem.Text)) = 0 Then
        objICMSExcecoes.sEstadoOrigem = "  "
    Else
        objICMSExcecoes.sEstadoOrigem = Origem.Text
    End If

    objICMSExcecoes.iGrupoOrigemMercadoria = GrupoOrigemMercadoria.ItemData(GrupoOrigemMercadoria.ListIndex)
    
    For iIndice = 0 To 2
        If TipoAplicacao(iIndice).Value Then
            objICMSExcecoes.iTipoAplicacao = iIndice
            Exit For
        End If
    Next
    
End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsgRet As VbMsgBoxResult
Dim objICMSExcecoes As New ClassICMSExcecao

On Error GoTo Erro_BotaoExcluir_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os dados foram preenchidos corretamente
    If Verifica_Identificacao_Preenchida <> SUCESSO Then Error 22987
    
    'Pede Confirmação para exclusão ao usuário
    vbMsgRet = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_EXCECAO")

    If vbMsgRet = vbYes Then
        
        'Preenche o objICMSExcecoes
        If Move_Identificacao_Memoria(objICMSExcecoes) <> SUCESSO Then Error 22989
        
        'Exclui a Execeção ICMS
        lErro = CF("ICMSExcecao_Exclui", objICMSExcecoes)
        If lErro <> SUCESSO Then Error 21466
        
        'Limpa a tela
        Call Limpa_Tela_ExcecoesICMS
        
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21466, 22987, 22989 'tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159707)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
     'Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 21470
    
    Call Rotina_Aviso(vbOKOnly, "AVISO_ALTER_NECESS_REINICIO_CORPORATOR")
    
    'LImpa a tela
    Call Limpa_Tela_ExcecoesICMS

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 21470 'tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159708)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 21471
    
    'Limpa a tela
    Call Limpa_Tela_ExcecoesICMS

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 21471 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159709)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_ExcecoesICMS()
'Limpa a Tela de Exceções de ICMS

    Fundamentacao.Text = ""
    TodosProdutos.Value = 0
    CategoriaProduto.ListIndex = -1
    TodosClientes.Value = 0
    CategoriaCliente.ListIndex = -1
    TipoTributacao.ListIndex = -1
    TipoTributacaoSimples.ListIndex = -1
    Origem.ListIndex = -1
    
    GrupoOrigemMercadoria.ListIndex = 0
    
    '####################################
    'Inserido por Wagner 29/09/05
    TodosFornecedores.Value = 0
    CategoriaFornecedor.ListIndex = -1
    '####################################
    
    TipoAplicacao(0).Value = True
    
    PautaOuMargem(TIPO_MARGEM).Value = True
    PautaOuMargem(TIPO_MARGEM).Enabled = False
    PautaOuMargem(TIPO_PAUTA).Enabled = False
    MargemLucroSubst.Enabled = False
    ValorPauta.Enabled = False
    LabelPauta.Enabled = False
    LabelMarg.Enabled = False
    LabelRedBaseSubst.Enabled = False
    RedBaseCalculoSubst.Enabled = False
    
    ICMSSTBaseDupla.Value = vbUnchecked
    ICMSPercFCP.Text = ""
            
    cBenef.Text = ""
    ICMSMotivo.ListIndex = -1
    
    iAlterado = 0

    Exit Sub

End Sub

Private Sub CategoriaCliente_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Origem_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
Dim objEstado As New ClassEstado

On Error GoTo Erro_Origem_Validate

    If Len(Trim(Origem.Text)) > 0 And Origem.ListIndex = -1 Then
            
        'pesquisa o item na lista
        lErro = Combo_Seleciona(Origem, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19303
            
        'Se não encontrou ----> ERRO
        If lErro <> SUCESSO Then Error 19304
        
    End If

    If Len(Trim(Origem.Text)) = 0 Then DescrEstDest.Caption = ""

    Exit Sub

Erro_Origem_Validate:

    Cancel = True


    Select Case Err

        Case 19303 'tratado na Rotina chamada
        
        Case 19304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Origem.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159710)

    End Select

    Exit Sub

End Sub

Private Sub Destino_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
Dim objEstado As New ClassEstado

On Error GoTo Erro_Destino_Validate

    If Len(Destino.Text) > 0 And Destino.ListIndex = -1 Then
        
        'pesquisa o item na lista
        lErro = Combo_Seleciona(Destino, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 19303
            
        'Se não encontrou ----> ERRO
        If lErro <> SUCESSO Then Error 19304
        
    End If

    Exit Sub

Erro_Destino_Validate:

    Cancel = True


    Select Case Err

        Case 19303 'tratado na Rotina chamada
        
        Case 19304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTADO_NAO_CADASTRADO", Err, Destino.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159711)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sCodigo As String
Dim iIndice As Integer
'Dim colTiposTribICMS As New AdmColCodigoNome
'Dim objTiposTribICMS As New AdmCodigoNome
Dim colCategoriaProduto As New Collection
Dim colCategoriaCliente As New Collection
Dim objEstado As New ClassEstado
Dim objCategoriaProduto As New ClassCategoriaProduto
Dim objCategoriaCliente As New ClassCategoriaCliente
Dim colEstado As New Collection
Dim colTiposTribICMS As New Collection, colTiposTribICMSSimples As New Collection
Dim objTipoTribICMS As ClassTipoTribICMS, objTipoTribICMSSimples As ClassTipoTribICMSSimples

'###########################################
'Inserido por Wagner
Dim colCategoriaFornecedor As New Collection
Dim objCategoriaFornecedor As New ClassCategoriaFornecedor
'###########################################

On Error GoTo Erro_Form_Load

    iTipoAnt = -1

    'Le os estados
    lErro = CF("Estados_Le_Todos", colEstado)
    If lErro <> SUCESSO Then gError 21478
    
    'Preenche a combo de Estados
    For Each objEstado In colEstado
        Destino.AddItem objEstado.sSigla
        Origem.AddItem objEstado.sSigla 'William 30/04/01
    Next

    'Le as categorias de produto
    lErro = CF("CategoriasProduto_Le_Todas", colCategoriaProduto)
    If lErro <> SUCESSO And lErro <> 22542 Then gError 21566

    'Preenche CategoriaProduto
    For Each objCategoriaProduto In colCategoriaProduto

        CategoriaProduto.AddItem objCategoriaProduto.sCategoria

    Next

    'Le as categorias de cliente
    lErro = CF("CategoriaCliente_Le_Todos", colCategoriaCliente)
    If lErro <> SUCESSO Then gError 21479

    'Preenche CategoriaCliente
    For Each objCategoriaCliente In colCategoriaCliente

        CategoriaCliente.AddItem objCategoriaCliente.sCategoria

    Next
    
    '##########################################################
    'Inserido por Wagner
    'Le as categorias de Fornecedor
    lErro = CF("CategoriaFornecedor_Le_Todos", colCategoriaFornecedor)
    If lErro <> SUCESSO And lErro <> 68486 Then gError 140407

    'Preenche CategoriaFornecedor
    For Each objCategoriaFornecedor In colCategoriaFornecedor

        CategoriaFornecedor.AddItem objCategoriaFornecedor.sCategoria

    Next
    '##########################################################

'    'Le cada Codigo e Descrição da tabela TiposTribICMS e poe na colecao
'    lErro = CF("Cod_Nomes_Le", "TiposTribICMS", "Tipo", "Descricao", STRING_TIPO_ICMS_DESCRICAO, colTiposTribICMS)
'    If lErro <> SUCESSO Then gError 21480
'
'    iIndice = 0
'
'    'Preenche TipoTributacao
'    For Each objTiposTribICMS In colTiposTribICMS
'
'        sCodigo = CStr(objTiposTribICMS.iCodigo) & SEPARADOR & objTiposTribICMS.sNome
'        TipoTributacao.AddItem (sCodigo)
'        TipoTributacao.ItemData(iIndice) = objTiposTribICMS.iCodigo
'        iIndice = iIndice + 1
'
'    Next

    lErro = CF("TiposTribICMS_Le_Todos", colTiposTribICMS)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = CF("TiposTribICMSSimples_Le_Todos", colTiposTribICMSSimples)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objTipoTribICMSSimples In colTiposTribICMSSimples
        TipoTributacaoSimples.AddItem Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao
        TipoTributacaoSimples.ItemData(TipoTributacaoSimples.NewIndex) = objTipoTribICMSSimples.iTipo
    Next

    For Each objTipoTribICMS In colTiposTribICMS
        TipoTributacao.AddItem Format(objTipoTribICMS.iTipoTribCST, "00") & SEPARADOR & objTipoTribICMS.sDescricao
        TipoTributacao.ItemData(TipoTributacao.NewIndex) = objTipoTribICMS.iTipo
    Next

    Set objEventoExcecoesICMS = New AdmEvento

    'poderia pegar a UF da filial corrente
    Destino.Text = "RJ"
    Call Destino_Validate(bSGECancelDummy)
        
    Call CategoriaProduto_Click
    Call CategoriaCliente_Click
    Call TipoTributacao_Click
    Call CategoriaFornecedor_Click 'Inserido por Wagner
    
    GrupoOrigemMercadoria.ListIndex = 0
    
    lErro = CF("Carrega_Combo", ICMSMotivo, "TribICMSMotivos", "Codigo", TIPO_INT, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError 21566
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 21478, 21479, 21480, 21566, 140407 'tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159712)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objExcecoesICMS As ClassICMSExcecao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se objICMSExcecao estiver preenchido
    If Not (objExcecoesICMS Is Nothing) Then
        
        'Preenche a tela com o objExcecoesICMS passado
        Call Traz_Excecao_Tela(objExcecoesICMS)

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159713)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Gravar_Registro()
'Gravação

Dim lErro As Long
Dim objICMSExcecoes As New ClassICMSExcecao

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se todos os campos foram gravados corretamentes
    If Verifica_Identificacao_Preenchida <> SUCESSO Then Error 22988
    
    'Preenche o objICMSExcecoes
    If Move_Identificacao_Memoria(objICMSExcecoes) <> SUCESSO Then Error 22990
    
    lErro = Trata_Alteracao(objICMSExcecoes, objICMSExcecoes.iGrupoOrigemMercadoria, objICMSExcecoes.sEstadoDestino, objICMSExcecoes.sCategoriaProduto, objICMSExcecoes.sCategoriaProdutoItem, objICMSExcecoes.sCategoriaCliente, objICMSExcecoes.sCategoriaClienteItem, objICMSExcecoes.sEstadoOrigem, objICMSExcecoes.sCategoriaFornecedor, objICMSExcecoes.sCategoriaFornecedorItem) 'Alterado por Wagner
    If lErro <> SUCESSO Then Error 32320
        
    'identifica o tipo da prioridade
    If TodosClientes.Value = 1 Then
        objICMSExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_PRODUTO
    Else
        If TodosProdutos.Value = 1 Then
            objICMSExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE
        Else
            objICMSExcecoes.iPrioridade = TIPOTRIB_PRIORIDADE_CLIENTE_PRODUTO
        End If
    End If

    If TipoTributacao.ListIndex = -1 Then Error 21725
    If TipoTributacaoSimples.ListIndex = -1 Then Error 21725
    objICMSExcecoes.iTipo = TipoTributacao.ItemData(TipoTributacao.ListIndex)
    objICMSExcecoes.iTipoSimples = TipoTributacaoSimples.ItemData(TipoTributacaoSimples.ListIndex)

    'Se campos habilitados, move seus dados
    If Aliquota.Enabled = True Then
        If Len(Aliquota.Text) > 0 Then objICMSExcecoes.dAliquota = CDbl(Aliquota.Text / 100)
    End If
    If RedBaseCalculo.Enabled = True Then
        If Len(RedBaseCalculo.Text) > 0 Then objICMSExcecoes.dPercRedBaseCalculo = CDbl(RedBaseCalculo.Text / 100)
    End If
    If RedBaseCalculoSubst.Enabled = True Then
        If Len(RedBaseCalculoSubst.Text) > 0 Then objICMSExcecoes.dPercRedBaseCalculoSubst = CDbl(RedBaseCalculoSubst.Text / 100)
    End If
    If MargemLucroSubst.Enabled = True Then
        If Len(MargemLucroSubst.Text) > 0 Then objICMSExcecoes.dPercMargemLucro = CDbl(MargemLucroSubst.Text / 100)
    End If
    If ValorPauta.Enabled = True Then
        If Len(ValorPauta.Text) > 0 Then objICMSExcecoes.dValorPauta = StrParaDbl(ValorPauta.Text)
    End If
    If PautaOuMargem(TIPO_PAUTA).Value Then
        objICMSExcecoes.iUsaPauta = MARCADO
    Else
        objICMSExcecoes.iUsaPauta = DESMARCADO
    End If
    If Len(ICMSPercFCP.Text) > 0 Then objICMSExcecoes.dICMSPercFCP = CDbl(ICMSPercFCP.Text / 100)
    
    If ICMSSTBaseDupla.Value = vbChecked Then
        objICMSExcecoes.iICMSSTBaseDupla = MARCADO
    Else
        objICMSExcecoes.iICMSSTBaseDupla = DESMARCADO
    End If
    objICMSExcecoes.dtICMSSTBaseDuplaIni = MaskedParaDate(ICMSSTBaseDuplaIni)
    
    objICMSExcecoes.scBenef = Trim(cBenef.Text)
    
    objICMSExcecoes.iICMSMotivo = Codigo_Extrai(ICMSMotivo.Text)
    
    'Grava a Execeção no BD
    lErro = CF("ICMSExcecao_Grava", objICMSExcecoes)
    If lErro <> SUCESSO Then Error 21489
    
    'Limpa a Tela
    Call Limpa_Tela_ExcecoesICMS

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 21489, 22988, 22990, 32320 'tratado na Rotina chamada
        
        Case 21725
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ICMS_NAO_PREENCHIDO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159714)

    End Select

    Exit Function

End Function

Private Sub Fundamentacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaCliente_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ItemCategoriaCliente_Validate

    If Len(ItemCategoriaCliente.Text) <> 0 And ItemCategoriaCliente.ListIndex = -1 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaCliente)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22997
        
        If lErro <> SUCESSO Then Error 22998
    
    End If

    Exit Sub

Erro_ItemCategoriaCliente_Validate:

    Cancel = True


    Select Case Err

        Case 22997 'tratado na Rotina chamada
        
        Case 22998
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIACLIENTEITEM_INEXISTENTE", Err, ItemCategoriaCliente.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159715)

    End Select

    Exit Sub

End Sub

Private Sub ItemCategoriaProduto_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer
On Error GoTo Erro_ItemCategoriaProduto_Validate

    If Len(ItemCategoriaProduto.Text) <> 0 And ItemCategoriaProduto.ListIndex = -1 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaProduto)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 22995
        
        If lErro <> SUCESSO Then Error 22996
    
    End If

    Exit Sub

Erro_ItemCategoriaProduto_Validate:

    Cancel = True


    Select Case Err

        Case 22995 'tratado na Rotina chamada
        
        Case 22996
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAPRODUTOITEM_INEXISTENTE", Err, ItemCategoriaProduto.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 159716)

    End Select

    Exit Sub

End Sub

Private Sub MargemLucroSubst_Validate(Cancel As Boolean)

Dim lErro As Long, dValor As Double

On Error GoTo Erro_MargemLucroSubst_Validate

    If Len(Trim(MargemLucroSubst.Text)) > 0 Then

        'Critica valor
        lErro = Valor_NaoNegativo_Critica(MargemLucroSubst.Text)
        If lErro <> SUCESSO Then Error 21499

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_MargemLucroSubst_Validate:

    Cancel = True


    Select Case Err

        Case 21499

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159717)

    End Select

    Exit Sub

End Sub

Private Sub RedBaseCalculo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RedBaseCalculo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RedBaseCalculo_Validate

    If Len(RedBaseCalculo.Text) > 0 Then

        'Critica valor
        lErro = Porcentagem_Critica2(RedBaseCalculo.Text)
        If lErro <> SUCESSO Then Error 21500

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_RedBaseCalculo_Validate:

    Cancel = True


    Select Case Err

        Case 21500

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159718)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacao_Click()

Dim lErro As Long
Dim objTiposTribICMS As New ClassTipoTribICMS
On Error GoTo Erro_TipoTributacao_Click

    iAlterado = REGISTRO_ALTERADO
    
    If TipoTributacao.ListIndex <> -1 Then
    
        objTiposTribICMS.iTipo = TipoTributacao.ItemData(TipoTributacao.ListIndex)
        
        'Lê os dados sobre o tipo de tributação
        lErro = CF("TipoTribICMS_Le", objTiposTribICMS)
        If lErro <> SUCESSO And lErro <> 21534 Then Error 21501

        If lErro = 21534 Then Error 21536

        'De acordo com os dados lidos, se for permitido, abilita campos. Caso contrário, desabilita.
        If objTiposTribICMS.iPermiteAliquota <> TIPOTRIB_PERMITE_ALIQUOTA Then
            LabelAliquota.Enabled = False
            Aliquota.Text = ""
            Aliquota.Enabled = False
        Else
            LabelAliquota.Enabled = True
            Aliquota.Enabled = True
        End If

        If objTiposTribICMS.iPermiteMargLucro <> TIPOTRIB_PERMITE_MARGLUCRO Then
            LabelMarg.Enabled = False
            MargemLucroSubst.Text = ""
            MargemLucroSubst.Enabled = False
            LabelPauta.Enabled = False
            ValorPauta.Text = ""
            ValorPauta.Enabled = False
            PautaOuMargem(TIPO_PAUTA).Enabled = False
            PautaOuMargem(TIPO_MARGEM).Enabled = False
            LabelRedBaseSubst.Enabled = False
            RedBaseCalculoSubst.Enabled = False
            RedBaseCalculoSubst.Text = ""
        Else
            PautaOuMargem(TIPO_PAUTA).Enabled = True
            PautaOuMargem(TIPO_MARGEM).Enabled = True
            If PautaOuMargem(TIPO_PAUTA).Value Then
                LabelPauta.Enabled = True
                ValorPauta.Enabled = True
            Else
                LabelMarg.Enabled = True
                MargemLucroSubst.Enabled = True
            End If
            LabelRedBaseSubst.Enabled = True
            RedBaseCalculoSubst.Enabled = True
        End If

        If objTiposTribICMS.iPermiteReducaoBase <> TIPOTRIB_PERMITE_REDUCAOBASE Then
            LabelRedBase.Enabled = False
            RedBaseCalculo.Text = ""
            RedBaseCalculo.Enabled = False
        Else
            LabelRedBase.Enabled = True
            RedBaseCalculo.Enabled = True
        End If
        
        If iTipoAnt <> objTiposTribICMS.iTipo Then
            Select Case objTiposTribICMS.iTipo
                Case 0
                    TipoTributacaoSimples.Text = "8"
                Case 8
                    TipoTributacaoSimples.Text = "9"
                Case 99
                    TipoTributacaoSimples.Text = "10"
                Case 4, 6, 9, 10
                    TipoTributacaoSimples.Text = "4"
                Case Else
                    TipoTributacaoSimples.Text = "1"
            End Select
            Call TipoTributacaoSimples_Validate(bSGECancelDummy)
        End If
        
        iTipoAnt = objTiposTribICMS.iTipo
    Else
        iTipoAnt = -1
        
        'limpa e desabilita os campos
        Aliquota.Text = ""
        Aliquota.Enabled = False
        MargemLucroSubst.Text = ""
        MargemLucroSubst.Enabled = False
        RedBaseCalculo.Text = ""
        RedBaseCalculo.Enabled = False
        LabelPauta.Enabled = False
        ValorPauta.Text = ""
        ValorPauta.Enabled = False
        PautaOuMargem(TIPO_PAUTA).Enabled = False
        PautaOuMargem(TIPO_MARGEM).Enabled = False
        LabelRedBaseSubst.Enabled = False
        RedBaseCalculoSubst.Enabled = False
        RedBaseCalculoSubst.Text = ""
        LabelMarg.Enabled = False
        
    End If
    
    Exit Sub
    
Erro_TipoTributacao_Click:

    Select Case Err

        Case 21501

        Case 21536
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS", Err, objTiposTribICMS.iTipo)
            TipoTributacao.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159719)

    End Select

    Exit Sub
    
End Sub

Private Sub TipoTributacao_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
On Error GoTo Erro_TipoTributacao_Validate

    If Len(Trim(TipoTributacao.Text)) <> 0 Then

         'Verifica se está preenchida com o ítem selecionado na ComboBox TipoTributacao
        If TipoTributacao.ListIndex = -1 Then

            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(TipoTributacao, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 21712

            'Não existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then Error 21713

            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then Error 21714

        End If

    End If

    Exit Sub

Erro_TipoTributacao_Validate:

    Cancel = True


    Select Case Err

        Case 21712

        Case 21713, 21714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS", Err, TipoTributacao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159720)

    End Select

    Exit Sub

End Sub

Private Sub TodosClientes_Click()

Dim lErro As Long
On Error GoTo Erro_TodosClientes_Click

    'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
    If TodosProdutos.Value = 1 And TodosClientes.Value = 1 And OptCliForn(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then Error 21503

    If TodosClientes.Value = 1 Then CategoriaCliente.ListIndex = -1

    Exit Sub

Erro_TodosClientes_Click:

    Select Case Err

        Case 21503
            lErro = Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", Err)
            TodosClientes.Value = 0

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159721)

    End Select

    Exit Sub

End Sub

Private Sub TodosProdutos_Click()

Dim lErro As Long
On Error GoTo Erro_TodosProdutos_Click

    '##########################################
    'Alterado por Wagner 29/09/05
    
    If OptCliForn(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then

        'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
        If TodosClientes.Value = vbChecked And TodosProdutos.Value = vbChecked Then gError 21504
    
    Else
    
        'TodosCLientes e todos Produto não podem ser marcados ao mesmo tempo
        If TodosFornecedores.Value = vbChecked And TodosProdutos.Value = vbChecked Then gError 140408
    
    End If
    '##########################################
    
    If TodosProdutos.Value = vbChecked Then CategoriaProduto.ListIndex = -1

    Exit Sub

Erro_TodosProdutos_Click:

    Select Case gErr

        Case 21504
            Call Rotina_Erro(vbOKOnly, "AVISO_NAO_E_POSSIVEL_SELECIONAR_TODOS", gErr)
            TodosProdutos.Value = vbUnchecked
            
        Case 140408
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_E_POSSIVEL_SELECIONAR_TODOS_FORN", gErr)
            TodosProdutos.Value = vbUnchecked

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159722)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EXCECOES_ICMS
    Set Form_Load_Ocx = Me
    Caption = "Exceções ICMS"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ExcecoesICMS"
    
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

'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************
'********************************************************************************


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fundamentacao Then
            Call LabelFundamentacao_Click
        End If
    
    End If

End Sub


Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
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

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub DescrEstado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescrEstado, Source, X, Y)
End Sub

Private Sub DescrEstado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescrEstado, Button, Shift, X, Y)
End Sub

Private Sub LabelRedBase_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelRedBase, Source, X, Y)
End Sub

Private Sub LabelRedBase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelRedBase, Button, Shift, X, Y)
End Sub

Private Sub LabelAliquota_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelAliquota, Source, X, Y)
End Sub

Private Sub LabelAliquota_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelAliquota, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub LabelMarg_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelMarg, Source, X, Y)
End Sub

Private Sub LabelMarg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelMarg, Button, Shift, X, Y)
End Sub

Private Sub LabelFundamentacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFundamentacao, Source, X, Y)
End Sub

Private Sub LabelFundamentacao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFundamentacao, Button, Shift, X, Y)
End Sub

'#######################################################
'Inserido por Wagner 29/09/05
Private Sub OptCliForn_Click(Index As Integer)

    If OptCliForn(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Value = True Then
    
        ItemCategoriaFornecedor.Text = ""
        CategoriaFornecedor.Text = ""
        
        Frame(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Visible = True
        Frame(ICMSEXCECOES_TIPOCLIFORN_FORNECEDOR).Visible = False
    
    Else
        ItemCategoriaCliente.Text = ""
        CategoriaCliente.Text = ""
        
        Frame(ICMSEXCECOES_TIPOCLIFORN_CLIENTE).Visible = False
        Frame(ICMSEXCECOES_TIPOCLIFORN_FORNECEDOR).Visible = True
    
    End If

End Sub

Private Sub ItemCategoriaFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaFornecedor_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ItemCategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ItemCategoriaFornecedor_Validate

    If Len(ItemCategoriaFornecedor.Text) <> 0 And ItemCategoriaFornecedor.ListIndex = -1 Then
    
        'pesquisa o item na lista
        lErro = Combo_Item_Igual(ItemCategoriaFornecedor)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 140400
        
        If lErro <> SUCESSO Then gError 140401
    
    End If

    Exit Sub

Erro_ItemCategoriaFornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 140400 'tratado na Rotina chamada
        
        Case 140401
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIAFORNECEDORITEM_INEXISTENTE", gErr, ItemCategoriaFornecedor.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159723)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaFornecedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CategoriaFornecedor_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objCategoriaFornecedorItem As New ClassCategoriaFornItem
Dim colCategoria As New Collection

On Error GoTo Erro_CategoriaFornecedor_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Verifica se a CategoriaFornecedor foi preenchida
    If CategoriaFornecedor.ListIndex <> -1 Then

        objCategoriaFornecedorItem.sCategoria = CategoriaFornecedor.Text

        'Lê os dados de Itens da Categoria do Fornecedor
        lErro = CF("CategoriaFornecedor_Le_Itens", objCategoriaFornecedorItem, colCategoria)
        If lErro <> SUCESSO Then gError 140402

        ItemCategoriaFornecedor.Enabled = True

        'Limpa os dados de ItemCategoriaFornecedor
        ItemCategoriaFornecedor.Clear

        'Preenche ItemCategoriaFornecedor
        For Each objCategoriaFornecedorItem In colCategoria

            ItemCategoriaFornecedor.AddItem objCategoriaFornecedorItem.sItem

        Next
        TodosFornecedores.Value = 0
    
    Else
        
        'Senão Desablita ItemCategoriaFornecedor
        ItemCategoriaFornecedor.ListIndex = -1
        ItemCategoriaFornecedor.Enabled = False
    
    End If

    Exit Sub

Erro_CategoriaFornecedor_Click:

    Select Case gErr

        Case 140402 'Tratado na Rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159724)

    End Select

    Exit Sub

End Sub

Private Sub CategoriaFornecedor_Validate(Cancel As Boolean)

Dim lErro As Long, iCodigo As Integer

On Error GoTo Erro_CategoriaFornecedor_Validate

    If Len(CategoriaFornecedor.Text) <> 0 And CategoriaFornecedor.ListIndex = -1 Then
    
        'pesquisa a categoria na lista
        lErro = Combo_Item_Igual(CategoriaFornecedor)
        If lErro <> SUCESSO And lErro <> 12253 Then gError 140403
        
        If lErro <> SUCESSO Then gError 140404
    
    End If
    
    'Se a CategoriaFornecedor estiver em branco desabilita e limpa a combo
    If Len(CategoriaFornecedor.Text) = 0 Then
        ItemCategoriaFornecedor.Enabled = False
        ItemCategoriaFornecedor.Clear
    End If
    
    Exit Sub

Erro_CategoriaFornecedor_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 140403
         
        Case 140404
            Call Rotina_Erro(vbOKOnly, "ERRO_CATEGORIA_NAO_CADASTRADA", gErr, CategoriaFornecedor.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159725)

    End Select

    Exit Sub

End Sub

Private Sub TodosFornecedores_Click()

Dim lErro As Long
On Error GoTo Erro_TodosFornecedores_Click

    'TodosFornecedors e todos Produto não podem ser marcados ao mesmo tempo
    If TodosProdutos.Value = vbChecked And TodosFornecedores.Value = vbChecked And OptCliForn(ICMSEXCECOES_TIPOCLIFORN_FORNECEDOR).Value = True Then gError 140407

    If TodosFornecedores.Value = vbChecked Then CategoriaFornecedor.ListIndex = -1

    Exit Sub

Erro_TodosFornecedores_Click:

    Select Case gErr

        Case 140407
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_E_POSSIVEL_SELECIONAR_TODOS_FORN", gErr)
            TodosFornecedores.Value = vbUnchecked

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159726)

    End Select

    Exit Sub

End Sub
'#######################################################

Private Sub PautaOuMargem_Click(Index As Integer)

    If PautaOuMargem(TIPO_PAUTA).Value Then
        If PautaOuMargem(TIPO_PAUTA).Enabled Then
            MargemLucroSubst.Enabled = False
            ValorPauta.Enabled = True
            MargemLucroSubst.Text = ""
            LabelPauta.Enabled = True
            LabelMarg.Enabled = False
        End If
    Else
        If PautaOuMargem(TIPO_MARGEM).Enabled Then
            ValorPauta.Enabled = False
            MargemLucroSubst.Enabled = True
            ValorPauta.Text = ""
            LabelPauta.Enabled = False
            LabelMarg.Enabled = True
        End If
    End If
    
End Sub

Public Sub ValorPauta_Validate(Cancel As Boolean)

Dim lErro As Long
Dim dValorPauta As Double

On Error GoTo Erro_ValorPauta_Validate

    If Len(Trim(ValorPauta.Text)) > 0 Then

        'Critica se valor é não negativo
        lErro = Valor_NaoNegativo_Critica(ValorPauta.Text)
        If lErro <> SUCESSO Then gError 198580

        dValorPauta = CDbl(ValorPauta.Text)

        ValorPauta.Text = Format(dValorPauta, "Standard")

    End If

    Exit Sub

Erro_ValorPauta_Validate:

    Cancel = True

    Select Case gErr

        Case 198580

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 198581)

    End Select

    Exit Sub

End Sub

Private Sub MargemLucroSubst_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ValorPauta_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RedBaseCalculoSubst_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub RedBaseCalculoSubst_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_RedBaseCalculoSubst_Validate

    If Len(RedBaseCalculoSubst.Text) > 0 Then

        'Critica valor
        lErro = Porcentagem_Critica2(RedBaseCalculoSubst.Text)
        If lErro <> SUCESSO Then gError 21500

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_RedBaseCalculoSubst_Validate:

    Cancel = True

    Select Case gErr

        Case 21500

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159718)

    End Select

    Exit Sub

End Sub

Private Sub TipoTributacaoSimples_Validate(Cancel As Boolean)

Dim iCodigo As Integer
Dim lErro As Long
On Error GoTo Erro_TipoTributacaoSimples_Validate

    If Len(Trim(TipoTributacaoSimples.Text)) <> 0 Then

         'Verifica se está preenchida com o ítem selecionado na ComboBox TipoTributacao
        If TipoTributacaoSimples.ListIndex = -1 Then

            'Verifica se existe o ítem na List da Combo. Se existir seleciona.
            lErro = Combo_Seleciona(TipoTributacaoSimples, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 21712

            'Não existe o ítem com o CÓDIGO na List da ComboBox
            If lErro = 6730 Then Error 21713

            'Não existe o ítem com a STRING na List da ComboBox
            If lErro = 6731 Then Error 21714

        End If

    End If

    Exit Sub

Erro_TipoTributacaoSimples_Validate:

    Cancel = True

    Select Case Err

        Case 21712

        Case 21713, 21714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_CADASTRADO_TIPOSTRIBICMS", Err, TipoTributacao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159720)

    End Select

    Exit Sub

End Sub

Private Sub ICMSPercFCP_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSPercFCP_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSPercFCP_Validate

    If Len(ICMSPercFCP.Text) > 0 Then

        'Testa o valor
        lErro = Porcentagem_Critica2(ICMSPercFCP.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_ICMSPercFCP_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159697)

    End Select

    Exit Sub

End Sub

Private Sub ICMSSTBaseDupla_Click()

Dim lErro As Long

On Error GoTo Erro_ICMSSTBaseDupla_Click

    If ICMSSTBaseDupla.Value = vbChecked Then
        ICMSSTBaseDuplaIni.Enabled = True
        UpDownICMSSTBaseDuplaIni.Enabled = True
        Call DateParaMasked(ICMSSTBaseDuplaIni, gdtDataHoje)
    Else
        ICMSSTBaseDuplaIni.Enabled = False
        UpDownICMSSTBaseDuplaIni.Enabled = False
        Call DateParaMasked(ICMSSTBaseDuplaIni, DATA_NULA)
    End If

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_ICMSSTBaseDupla_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 159697)

    End Select

    Exit Sub
    
End Sub

Private Sub ICMSSTBaseDuplaIni_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub ICMSSTBaseDuplaIni_GotFocus()
    Call MaskEdBox_TrataGotFocus(ICMSSTBaseDuplaIni, iAlterado)
End Sub

Private Sub ICMSSTBaseDuplaIni_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSTBaseDuplaIni_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(ICMSSTBaseDuplaIni.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(ICMSSTBaseDuplaIni.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_ICMSSTBaseDuplaIni_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160796)

    End Select

    Exit Sub

End Sub

Private Sub UpDownICMSSTBaseDuplaIni_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownICMSSTBaseDuplaIni_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(ICMSSTBaseDuplaIni, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownICMSSTBaseDuplaIni_DownClick:

    Select Case gErr

        Case 31202

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160804)

    End Select

    Exit Sub

End Sub

Private Sub UpDownICMSSTBaseDuplaIni_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownICMSSTBaseDuplaIni_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(ICMSSTBaseDuplaIni, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownICMSSTBaseDuplaIni_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160805)

    End Select

    Exit Sub

End Sub

Public Sub ICMSMotivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ICMSMotivo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
