VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl Vendedores 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   KeyPreview      =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   9045
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5010
      Index           =   2
      Left            =   75
      TabIndex        =   15
      Top             =   660
      Visible         =   0   'False
      Width           =   8865
      Begin VB.Frame Frame1 
         Caption         =   "Incide sobre"
         Height          =   1230
         Index           =   0
         Left            =   6120
         TabIndex        =   52
         Top             =   1035
         Width           =   2730
         Begin VB.CheckBox ComissaoSobreTotal 
            Caption         =   "Tudo"
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
            Left            =   165
            TabIndex        =   18
            Top             =   225
            Width           =   780
         End
         Begin VB.CheckBox ComissaoVenda 
            Caption         =   "Venda"
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
            Left            =   1725
            TabIndex        =   19
            Top             =   225
            Value           =   1  'Checked
            Width           =   870
         End
         Begin VB.CheckBox ComissaoFrete 
            Caption         =   "Frete"
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
            TabIndex        =   20
            Top             =   555
            Width           =   780
         End
         Begin VB.CheckBox ComissaoSeguro 
            Caption         =   "Seguro"
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
            Left            =   1725
            TabIndex        =   21
            Top             =   555
            Width           =   960
         End
         Begin VB.CheckBox ComissaoICM 
            Caption         =   "Outras Desp."
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
            TabIndex        =   22
            Top             =   870
            Width           =   1455
         End
         Begin VB.CheckBox ComissaoIPI 
            Caption         =   "IPI"
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
            Left            =   1725
            TabIndex        =   23
            Top             =   870
            Width           =   600
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Exce��es"
         Height          =   2715
         Left            =   45
         TabIndex        =   84
         Top             =   2265
         Width           =   8805
         Begin MSMask.MaskEdBox ExcMetaP 
            Height          =   210
            Left            =   4080
            TabIndex        =   91
            Top             =   1275
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExcPercComissP 
            Height          =   210
            Left            =   3120
            TabIndex        =   92
            Top             =   1275
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin VB.CommandButton BotaoProdutos 
            Caption         =   "Produtos"
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
            Left            =   45
            TabIndex        =   90
            Top             =   2370
            Width           =   1815
         End
         Begin MSMask.MaskEdBox ExcMeta 
            Height          =   210
            Left            =   5745
            TabIndex        =   89
            Top             =   375
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
            PromptChar      =   " "
         End
         Begin VB.TextBox ExcProdDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   210
            Left            =   1365
            MaxLength       =   250
            TabIndex        =   88
            Top             =   360
            Width           =   3090
         End
         Begin MSMask.MaskEdBox ExcProd 
            Height          =   210
            Left            =   330
            TabIndex        =   86
            Top             =   360
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox ExcPercComiss 
            Height          =   210
            Left            =   4770
            TabIndex        =   87
            Top             =   375
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   370
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
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
         Begin MSFlexGridLib.MSFlexGrid GridExc 
            Height          =   2115
            Left            =   45
            TabIndex        =   85
            Top             =   210
            Width           =   8715
            _ExtentX        =   15372
            _ExtentY        =   3731
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
      Begin VB.Frame Frame2 
         Caption         =   "Estat�sticas"
         Height          =   555
         Left            =   45
         TabIndex        =   42
         Top             =   480
         Width           =   8805
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Saldo a Receber:"
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
            TabIndex        =   43
            Top             =   210
            Width           =   1500
         End
         Begin VB.Label SaldoComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1965
            TabIndex        =   44
            Top             =   165
            Width           =   1575
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data da �ltima Venda:"
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
            Left            =   4365
            TabIndex        =   45
            Top             =   195
            Width           =   1935
         End
         Begin VB.Label DataUltVenda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6405
            TabIndex        =   46
            Top             =   165
            Width           =   1365
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   645
         Left            =   45
         TabIndex        =   53
         Top             =   1620
         Width           =   5985
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   855
            TabIndex        =   24
            Top             =   240
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   315
            Left            =   3330
            TabIndex        =   25
            Top             =   225
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Banco 
            Height          =   315
            Left            =   5130
            TabIndex        =   26
            Top             =   210
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin VB.Label LabelBanco 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4500
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   56
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Ag�ncia:"
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
            Left            =   2505
            TabIndex        =   55
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "N�mero:"
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
            Left            =   90
            TabIndex        =   54
            Top             =   315
            Width           =   705
         End
      End
      Begin VB.Frame SSFrame3 
         Height          =   510
         Left            =   45
         TabIndex        =   39
         Top             =   -30
         Width           =   8805
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   0
            Left            =   1140
            TabIndex        =   41
            Top             =   165
            Width           =   7590
         End
         Begin VB.Label Label4 
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
            Height          =   210
            Left            =   180
            TabIndex        =   40
            Top             =   165
            Width           =   870
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Porcentagens"
         Height          =   600
         Left            =   45
         TabIndex        =   47
         Top             =   1035
         Width           =   5985
         Begin MSMask.MaskEdBox PercComissao 
            Height          =   315
            Left            =   870
            TabIndex        =   16
            Top             =   210
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercComissaoEmissao 
            Height          =   315
            Left            =   3330
            TabIndex        =   17
            Top             =   210
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label PercComissaoBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5145
            TabIndex        =   51
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label19 
            Caption         =   "Na Baixa:"
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
            Left            =   4275
            TabIndex        =   50
            Top             =   255
            Width           =   840
         End
         Begin VB.Label Label10 
            Caption         =   "Total:"
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
            Left            =   315
            TabIndex        =   48
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Na Emiss�o:"
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
            Height          =   210
            Left            =   2235
            TabIndex        =   49
            Top             =   270
            Width           =   1065
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4950
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   720
      Width           =   8820
      Begin VB.TextBox Obs 
         Height          =   1530
         Left            =   1275
         MaxLength       =   5000
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1905
         Width           =   4560
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Aut�nomo"
         Height          =   975
         Index           =   0
         Left            =   0
         TabIndex        =   70
         Top             =   3870
         Visible         =   0   'False
         Width           =   5835
         Begin MSMask.MaskEdBox CGC0 
            Height          =   315
            Left            =   1275
            TabIndex        =   12
            Top             =   210
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox RG 
            Height          =   315
            Left            =   1275
            TabIndex        =   13
            Top             =   585
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "RG:"
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
            Left            =   885
            TabIndex        =   79
            Top             =   645
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "CPF:"
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
            Left            =   825
            TabIndex        =   78
            Top             =   285
            Width           =   420
         End
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Empregado"
         Height          =   990
         Index           =   1
         Left            =   0
         TabIndex        =   71
         Top             =   3855
         Visible         =   0   'False
         Width           =   5850
         Begin MSMask.MaskEdBox Matricula 
            Height          =   315
            Left            =   1275
            TabIndex        =   72
            Top             =   255
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Matr�cula:"
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
            TabIndex        =   73
            Top             =   300
            Width           =   885
         End
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados da Empresa"
         Height          =   990
         Index           =   2
         Left            =   0
         TabIndex        =   63
         Top             =   3855
         Visible         =   0   'False
         Width           =   5850
         Begin VB.TextBox RazaoSocial 
            Height          =   315
            Left            =   1275
            MaxLength       =   40
            TabIndex        =   69
            Top             =   240
            Width           =   4500
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   4455
            TabIndex        =   64
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGC 
            Height          =   315
            Left            =   1275
            TabIndex        =   65
            Top             =   600
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Raz�o Social:"
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
            TabIndex        =   68
            Top             =   285
            Width           =   1200
         End
         Begin VB.Label Label35 
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
            Left            =   255
            TabIndex        =   67
            Top             =   675
            Width           =   975
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "I.E.:"
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
            Left            =   4035
            TabIndex        =   66
            Top             =   645
            Width           =   375
         End
      End
      Begin VB.ComboBox CodUsuario 
         Height          =   315
         Left            =   3915
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   45
         Width           =   1935
      End
      Begin VB.ComboBox Cargo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1290
         TabIndex        =   8
         Top             =   1530
         Width           =   1800
      End
      Begin VB.ComboBox Vinculo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "VendedoresPur.ctx":0000
         Left            =   1290
         List            =   "VendedoresPur.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3495
         Width           =   2790
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   2220
         TabIndex        =   61
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1830
         Picture         =   "VendedoresPur.ctx":0035
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numera��o Autom�tica"
         Top             =   60
         Width           =   300
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   1155
         Width           =   1800
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   3915
         TabIndex        =   7
         Top             =   1125
         Width           =   1965
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   45
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Nome 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Top             =   405
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1290
         TabIndex        =   4
         Top             =   780
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.ListBox VendedoresList 
         Height          =   4545
         Left            =   5985
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   315
         Width           =   2685
      End
      Begin MSMask.MaskEdBox Superior 
         Height          =   315
         Left            =   3915
         TabIndex        =   9
         Top             =   1530
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   300
         Left            =   5040
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   765
         Width           =   225
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataInicio 
         Height          =   300
         Left            =   3915
         TabIndex        =   5
         Top             =   765
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelAssunto 
         AutoSize        =   -1  'True
         Caption         =   "Observa��o:"
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
         Left            =   135
         TabIndex        =   83
         Top             =   1965
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "In�cio:"
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
         Left            =   3315
         TabIndex        =   82
         Top             =   795
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Usu�rio:"
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
         Left            =   3120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   77
         Top             =   105
         Width           =   705
      End
      Begin VB.Label LabelSuperior 
         AutoSize        =   -1  'True
         Caption         =   "Superior:"
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
         Left            =   3135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   75
         Top             =   1575
         Width           =   780
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cargo:"
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
         Left            =   690
         TabIndex        =   74
         Top             =   1605
         Width           =   585
      End
      Begin VB.Label LabelVinculo 
         AutoSize        =   -1  'True
         Caption         =   "V�nculo:"
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
         Left            =   510
         TabIndex        =   62
         Top             =   3570
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C�digo:"
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
         Left            =   615
         TabIndex        =   33
         Top             =   90
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
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
         Left            =   720
         TabIndex        =   34
         Top             =   465
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome Red.:"
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
         Left            =   255
         TabIndex        =   35
         Top             =   855
         Width           =   1020
      End
      Begin VB.Label TipoLabel 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   825
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   36
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label13 
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
         Height          =   210
         Left            =   5985
         TabIndex        =   38
         Top             =   15
         Width           =   1440
      End
      Begin VB.Label RegiaoVendaLabel 
         AutoSize        =   -1  'True
         Caption         =   "Regi�o:"
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
         Left            =   3210
         TabIndex        =   37
         Top             =   1170
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5010
      Index           =   3
      Left            =   75
      TabIndex        =   27
      Top             =   660
      Visible         =   0   'False
      Width           =   8865
      Begin TelasCprPur.TabEndereco TabEnd 
         Height          =   3510
         Index           =   0
         Left            =   75
         TabIndex        =   80
         Top             =   975
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   6191
      End
      Begin VB.Frame SSFrame2 
         Height          =   525
         Left            =   75
         TabIndex        =   57
         Top             =   -15
         Width           =   8760
         Begin VB.Label Label6 
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
            Height          =   210
            Left            =   180
            TabIndex        =   58
            Top             =   180
            Width           =   870
         End
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   1
            Left            =   1140
            TabIndex        =   59
            Top             =   180
            Width           =   7095
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6840
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   45
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "VendedoresPur.ctx":011F
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "VendedoresPur.ctx":0279
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "VendedoresPur.ctx":0403
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "VendedoresPur.ctx":0935
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5385
      Left            =   45
      TabIndex        =   60
      Top             =   330
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   9499
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identifica��o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comiss�o"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endere�o"
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
Attribute VB_Name = "Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Private WithEvents objCT As CTVendedores
Attribute objCT.VB_VarHelpID = -1

Private Sub Ativo_Click()
     Call objCT.Ativo_Click
End Sub

Private Sub CGC_Change()
    Call objCT.CGC_Change
End Sub

Private Sub InscricaoEstadual_Change()
    Call objCT.InscricaoEstadual_Change
End Sub

Private Sub RazaoSocial_Change()
    Call objCT.RazaoSocial_Change
End Sub

'Alterado por Maur�cio Maciel em 11/04/03
Private Sub Vinculo_Click()
    Call objCT.Vinculo_Click
End Sub

Private Sub Banco_GotFocus()
     Call objCT.Banco_GotFocus
End Sub

Private Sub BotaoProxNum_Click()
     Call objCT.BotaoProxNum_Click
End Sub

Private Sub Agencia_Change()
     Call objCT.Agencia_Change
End Sub

Private Sub Banco_Change()
     Call objCT.Banco_Change
End Sub

Private Sub BotaoExcluir_Click()
     Call objCT.BotaoExcluir_Click
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

Private Sub Codigo_Change()
     Call objCT.Codigo_Change
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub ComissaoFrete_Click()
     Call objCT.ComissaoFrete_Click
End Sub

Private Sub ComissaoICM_Click()
     Call objCT.ComissaoICM_Click
End Sub

Private Sub ComissaoIPI_Click()
     Call objCT.ComissaoIPI_Click
End Sub

Private Sub ComissaoSeguro_Click()
     Call objCT.ComissaoSeguro_Click
End Sub

Private Sub ComissaoSobreTotal_Click()
     Call objCT.ComissaoSobreTotal_Click
End Sub

Private Sub ContaCorrente_Change()
     Call objCT.ContaCorrente_Change
End Sub

Private Sub VendedoresList_Adiciona(objVendedor As ClassVendedor)
     Call objCT.VendedoresList_Adiciona(objVendedor)
End Sub

Private Sub VendedoresList_Exclui(iCodigo As Integer)
     Call objCT.VendedoresList_Exclui(iCodigo)
End Sub

Public Sub Form_Load()
     Call objCT.Form_Load
End Sub

Public Sub Form_Activate()
     Call objCT.Form_Activate
End Sub

Public Sub Form_Deactivate()
     Call objCT.Form_Deactivate
End Sub

Private Sub LabelBanco_Click()
     Call objCT.LabelBanco_Click
End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)
     Call objCT.NomeReduzido_Validate(Cancel)
End Sub

Private Sub PercComissao_Validate(Cancel As Boolean)
     Call objCT.PercComissao_Validate(Cancel)
End Sub

Private Sub Regiao_Validate(Cancel As Boolean)
     Call objCT.Regiao_Validate(Cancel)
End Sub

Private Sub Tipo_Validate(Cancel As Boolean)
     Call objCT.Tipo_Validate(Cancel)
End Sub

Private Sub TipoLabel_Click()
     Call objCT.TipoLabel_Click
End Sub

Function Trata_Parametros(Optional objVendedor As ClassVendedor) As Long
     Trata_Parametros = objCT.Trata_Parametros(objVendedor)
End Function

Private Sub Matricula_Change()
     Call objCT.Matricula_Change
End Sub

Private Sub Nome_Change()
     Call objCT.Nome_Change
End Sub

Private Sub NomeReduzido_Change()
     Call objCT.NomeReduzido_Change
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub PercComissao_Change()
     Call objCT.PercComissao_Change
End Sub

Private Sub Regiao_Change()
     Call objCT.Regiao_Change
End Sub

Private Sub Regiao_Click()
     Call objCT.Regiao_Click
End Sub

Private Sub Tipo_Change()
     Call objCT.Tipo_Change
End Sub

Private Sub Tipo_Click()
     Call objCT.Tipo_Click
End Sub

Private Sub UserControl_Initialize()

    Set objCT = New CTVendedores
    Set objCT.objUserControl = Me
    
    Set objCT.gobjInfoUsu = New CTVendedoresVGPur
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTVendedoresPur

End Sub

Private Sub VendedoresList_DblClick()
     Call objCT.VendedoresList_DblClick
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
     Call objCT.Form_QueryUnload(Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub PercComissaoEmissao_Change()
     Call objCT.PercComissaoEmissao_Change
End Sub

Private Sub PercComissaoEmissao_Validate(Cancel As Boolean)
     Call objCT.PercComissaoEmissao_Validate(Cancel)
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

Private Sub VendedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(VendedorLabel(Index), Source, X, Y)
End Sub

Private Sub VendedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VendedorLabel(Index), Button, Shift, X, Y)
End Sub


Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
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

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub RegiaoVendaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(RegiaoVendaLabel, Source, X, Y)
End Sub

Private Sub RegiaoVendaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(RegiaoVendaLabel, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label20_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label20, Source, X, Y)
End Sub

Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label20, Button, Shift, X, Y)
End Sub

Private Sub SaldoComissao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SaldoComissao, Source, X, Y)
End Sub

Private Sub SaldoComissao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SaldoComissao, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub DataUltVenda_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataUltVenda, Source, X, Y)
End Sub

Private Sub DataUltVenda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataUltVenda, Button, Shift, X, Y)
End Sub

Private Sub LabelBanco_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelBanco, Source, X, Y)
End Sub

Private Sub LabelBanco_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelBanco, Button, Shift, X, Y)
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label25, Source, X, Y)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label22, Source, X, Y)
End Sub

Private Sub Label22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label22, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub PercComissaoBaixa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(PercComissaoBaixa, Source, X, Y)
End Sub

Private Sub PercComissaoBaixa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(PercComissaoBaixa, Button, Shift, X, Y)
End Sub

Private Sub Label19_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label19, Source, X, Y)
End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label19, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label14_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label14, Source, X, Y)
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label14, Button, Shift, X, Y)
End Sub

Private Sub LabelVinculo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelVinculo, Source, X, Y)
End Sub

Private Sub LabelVinculo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelVinculo, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub Cargo_Change()
    Call objCT.Cargo_Change
End Sub

Private Sub Cargo_Click()
    Call objCT.Cargo_Click
End Sub

Private Sub Cargo_Validate(Cancel As Boolean)
     Call objCT.Cargo_Validate(Cancel)
End Sub

Private Sub Superior_Change()
    Call objCT.Superior_Change
End Sub

Private Sub Superior_Validate(Cancel As Boolean)
     Call objCT.Superior_Validate(Cancel)
End Sub

Private Sub LabelSuperior_Click()
    Call objCT.LabelSuperior_Click
End Sub

Public Sub CodUsuario_Change()
    Call objCT.CodUsuario_Change
End Sub

Public Sub CodUsuario_Click()
    Call objCT.CodUsuario_Click
End Sub

Private Sub CGC_GotFocus()
     Call objCT.CGC_GotFocus
End Sub

Private Sub CGC0_GotFocus()
     Call objCT.CGC0_GotFocus
End Sub

Private Sub RG_GotFocus()
     Call objCT.RG_GotFocus
End Sub

Private Sub CGC0_Change()
     Call objCT.CGC0_Change
End Sub

Private Sub RG_Change()
     Call objCT.RG_Change
End Sub

Private Sub CGC_Validate(Cancel As Boolean)
     Call objCT.CGC_Validate(Cancel)
End Sub

Private Sub CGC0_Validate(Cancel As Boolean)
     Call objCT.CGC0_Validate(Cancel)
End Sub


Private Sub GridExc_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_Click(objCT)
End Sub

Private Sub GridExc_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_EnterCell(objCT)
End Sub

Private Sub GridExc_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_GotFocus(objCT)
End Sub

Private Sub GridExc_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridExc_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridExc_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_LeaveCell(objCT)
End Sub

Private Sub GridExc_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_Validate(objCT, Cancel)
End Sub

Private Sub GridExc_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_RowColChange(objCT)
End Sub

Private Sub GridExc_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridExc_Scroll(objCT)
End Sub

Private Sub ExcProd_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProd_Change(objCT)
End Sub

Private Sub ExcProd_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProd_GotFocus(objCT)
End Sub

Private Sub ExcProd_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProd_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcProd_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProd_Validate(objCT, Cancel)
End Sub

Private Sub ExcProdDesc_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProdDesc_Change(objCT)
End Sub

Private Sub ExcProdDesc_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProdDesc_GotFocus(objCT)
End Sub

Private Sub ExcProdDesc_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProdDesc_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcProdDesc_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcProdDesc_Validate(objCT, Cancel)
End Sub

Private Sub ExcMeta_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcMeta_Change(objCT)
End Sub

Private Sub ExcMeta_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcMeta_GotFocus(objCT)
End Sub

Private Sub ExcMeta_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcMeta_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcMeta_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcMeta_Validate(objCT, Cancel)
End Sub

Private Sub ExcPercComiss_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcPercComiss_Change(objCT)
End Sub

Private Sub ExcPercComiss_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcPercComiss_GotFocus(objCT)
End Sub

Private Sub ExcPercComiss_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcPercComiss_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ExcPercComiss_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ExcPercComiss_Validate(objCT, Cancel)
End Sub

Private Sub Obs_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.Obs_Change(objCT)
End Sub

Private Sub Obs_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Obs_Validate(objCT, Cancel)
End Sub

Private Sub DataInicio_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInicio_Change(objCT)
End Sub

Private Sub DataInicio_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInicio_GotFocus(objCT)
End Sub

Private Sub DataInicio_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.DataInicio_Validate(objCT, Cancel)
End Sub

Private Sub BotaoProdutos_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoProdutos_Click(objCT)
End Sub

Private Sub UpDown_DownClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDown_DownClick(objCT)
End Sub

Private Sub UpDown_UpClick()
     Call objCT.gobjInfoUsu.gobjTelaUsu.UpDown_UpClick(objCT)
End Sub
