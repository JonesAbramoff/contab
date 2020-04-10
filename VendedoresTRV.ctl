VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Vendedores 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "j"
      Height          =   4815
      Index           =   1
      Left            =   135
      TabIndex        =   57
      Top             =   1005
      Width           =   9195
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Autônomo"
         Height          =   1695
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   3000
         Visible         =   0   'False
         Width           =   6315
         Begin MSMask.MaskEdBox CGC0 
            Height          =   315
            Left            =   1635
            TabIndex        =   10
            Top             =   300
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
            Left            =   1635
            TabIndex        =   11
            Top             =   750
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            PromptChar      =   " "
         End
         Begin VB.Label Label16 
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
            Left            =   1200
            TabIndex        =   103
            Top             =   810
            Width           =   345
         End
         Begin VB.Label Label15 
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
            Left            =   1140
            TabIndex        =   102
            Top             =   375
            Width           =   420
         End
      End
      Begin VB.ComboBox CodUsuario 
         Height          =   315
         Left            =   4680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   1725
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Empregado"
         Height          =   1710
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   2985
         Visible         =   0   'False
         Width           =   6315
         Begin MSMask.MaskEdBox Matricula 
            Height          =   315
            Left            =   1590
            TabIndex        =   59
            Top             =   375
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
            Caption         =   "Matrícula:"
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
            TabIndex        =   60
            Top             =   420
            Width           =   885
         End
      End
      Begin VB.ListBox VendedoresList 
         Height          =   3765
         Left            =   6645
         Sorted          =   -1  'True
         TabIndex        =   70
         Top             =   330
         Width           =   2460
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   6915
         TabIndex        =   69
         Top             =   4410
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   1605
         Width           =   1935
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   2265
         Picture         =   "VendedoresTRV.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Numeração Automática"
         Top             =   60
         Width           =   300
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   2775
         TabIndex        =   2
         Top             =   105
         Width           =   855
      End
      Begin VB.ComboBox Vinculo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "VendedoresTRV.ctx":00EA
         Left            =   1710
         List            =   "VendedoresTRV.ctx":00F7
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2550
         Width           =   2790
      End
      Begin VB.ComboBox Cargo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   2070
         Width           =   1935
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados da Empresa"
         Height          =   1710
         Index           =   2
         Left            =   120
         TabIndex        =   62
         Top             =   2985
         Visible         =   0   'False
         Width           =   6315
         Begin VB.TextBox RazaoSocial 
            Height          =   315
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   63
            Top             =   240
            Width           =   3630
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   1560
            TabIndex        =   64
            Top             =   1230
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
            Left            =   1575
            TabIndex        =   65
            Top             =   720
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin VB.Label Label34 
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
            Left            =   180
            TabIndex        =   68
            Top             =   1305
            Width           =   1290
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
            Left            =   525
            TabIndex        =   67
            Top             =   795
            Width           =   975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Razão Social:"
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
            Left            =   270
            TabIndex        =   66
            Top             =   300
            Width           =   1200
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   360
         Left            =   6480
         TabIndex        =   61
         Top             =   4395
         Width           =   2535
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
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
         Left            =   1710
         TabIndex        =   4
         Top             =   540
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1710
         TabIndex        =   5
         Top             =   1065
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Superior 
         Height          =   315
         Left            =   4575
         TabIndex        =   8
         Top             =   2070
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
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
         Left            =   3870
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   100
         Top             =   90
         Width           =   705
      End
      Begin VB.Label RegiaoVendaLabel 
         AutoSize        =   -1  'True
         Caption         =   "Região:"
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
         Left            =   6135
         TabIndex        =   79
         Top             =   4455
         Visible         =   0   'False
         Width           =   675
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
         Left            =   6645
         TabIndex        =   78
         Top             =   105
         Width           =   1440
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
         Left            =   1125
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   77
         Top             =   1650
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nome Reduzido:"
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
         TabIndex        =   76
         Top             =   1110
         Width           =   1410
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
         Left            =   1050
         TabIndex        =   75
         Top             =   585
         Width           =   555
      End
      Begin VB.Label Label1 
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
         Left            =   945
         TabIndex        =   74
         Top             =   90
         Width           =   660
      End
      Begin VB.Label LabelVinculo 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo:"
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
         Left            =   855
         TabIndex        =   73
         Top             =   2625
         Width           =   735
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
         Left            =   1005
         TabIndex        =   72
         Top             =   2115
         Width           =   570
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
         Left            =   3795
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   71
         Top             =   2115
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4920
      Index           =   4
      Left            =   135
      TabIndex        =   53
      Top             =   990
      Visible         =   0   'False
      Width           =   9135
      Begin TelasCprTRV.TabEndereco TabEnd 
         Height          =   3600
         Index           =   0
         Left            =   150
         TabIndex        =   99
         Top             =   825
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   6350
      End
      Begin VB.Frame SSFrame2 
         Height          =   525
         Left            =   225
         TabIndex        =   54
         Top             =   30
         Width           =   8460
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   1
            Left            =   1140
            TabIndex        =   56
            Top             =   180
            Width           =   7095
         End
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
            TabIndex        =   55
            Top             =   180
            Width           =   870
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4890
      Index           =   2
      Left            =   135
      TabIndex        =   23
      Top             =   990
      Visible         =   0   'False
      Width           =   9105
      Begin VB.Frame SSFrame4 
         Caption         =   "Porcentagens"
         Height          =   1350
         Left            =   210
         TabIndex        =   46
         Top             =   1470
         Width           =   3750
         Begin MSMask.MaskEdBox PercComissao 
            Height          =   315
            Left            =   1215
            TabIndex        =   47
            Top             =   300
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
            Left            =   1200
            TabIndex        =   48
            Top             =   788
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            Caption         =   "Na Emissão:"
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
            Left            =   105
            TabIndex        =   52
            Top             =   840
            Width           =   1065
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
            Left            =   675
            TabIndex        =   51
            Top             =   360
            Width           =   495
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
            Left            =   2010
            TabIndex        =   50
            Top             =   840
            Width           =   840
         End
         Begin VB.Label PercComissaoBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2895
            TabIndex        =   49
            Top             =   788
            Width           =   765
         End
      End
      Begin VB.Frame SSFrame3 
         Height          =   510
         Left            =   225
         TabIndex        =   43
         Top             =   30
         Width           =   8220
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
            TabIndex        =   45
            Top             =   165
            Width           =   870
         End
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   0
            Left            =   1140
            TabIndex        =   44
            Top             =   165
            Width           =   6750
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   915
         Left            =   195
         TabIndex        =   36
         Top             =   2940
         Width           =   8205
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   1830
            TabIndex        =   37
            Top             =   420
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Agencia 
            Height          =   315
            Left            =   4860
            TabIndex        =   38
            Top             =   420
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
            Left            =   7020
            TabIndex        =   39
            Top             =   420
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
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
            Left            =   1005
            TabIndex        =   42
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   4020
            TabIndex        =   41
            Top             =   480
            Width           =   765
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
            Left            =   6255
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Estatísticas"
         Height          =   705
         Left            =   210
         TabIndex        =   31
         Top             =   645
         Width           =   8220
         Begin VB.Label DataUltVenda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6405
            TabIndex        =   35
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Data da Última Venda:"
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
            TabIndex        =   34
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label SaldoComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1965
            TabIndex        =   33
            Top             =   270
            Width           =   1575
         End
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
            TabIndex        =   32
            Top             =   315
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Incide sobre"
         Height          =   1350
         Index           =   0
         Left            =   3975
         TabIndex        =   24
         Top             =   1470
         Width           =   4455
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
            Left            =   3765
            TabIndex        =   30
            Top             =   780
            Width           =   600
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
            Left            =   2190
            TabIndex        =   29
            Top             =   780
            Width           =   1455
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
            Left            =   1110
            TabIndex        =   28
            Top             =   780
            Width           =   990
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
            TabIndex        =   27
            Top             =   780
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
            Left            =   1110
            TabIndex        =   26
            Top             =   390
            Value           =   1  'Checked
            Width           =   870
         End
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
            TabIndex        =   25
            Top             =   375
            Width           =   780
         End
      End
      Begin MSMask.MaskEdBox PercCallCenter 
         Height          =   312
         Left            =   1308
         TabIndex        =   97
         Top             =   4092
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Format          =   "#0.#0\%"
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
         Caption         =   "Call Center:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   216
         Left            =   276
         TabIndex        =   98
         Top             =   4128
         Width           =   1068
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   4875
      Index           =   3
      Left            =   135
      TabIndex        =   12
      Top             =   990
      Visible         =   0   'False
      Width           =   9120
      Begin VB.Frame Frame4 
         Caption         =   "Regiões de venda"
         Height          =   4215
         Left            =   5490
         TabIndex        =   20
         Top             =   630
         Width           =   3615
         Begin MSMask.MaskEdBox PercComissReg 
            Height          =   225
            Left            =   1410
            TabIndex        =   96
            Top             =   885
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   397
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
         Begin MSMask.MaskEdBox DescRegiao 
            Height          =   225
            Left            =   1200
            TabIndex        =   95
            Top             =   570
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CodRegiao 
            Height          =   225
            Left            =   420
            TabIndex        =   94
            Top             =   555
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   397
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99999"
            PromptChar      =   " "
         End
         Begin VB.CommandButton BotaoRegiaoVenda 
            Caption         =   "Regiões de Venda"
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
            Left            =   105
            TabIndex        =   21
            Top             =   3765
            Width           =   2085
         End
         Begin MSFlexGridLib.MSFlexGrid GridRegiaoVenda 
            Height          =   1110
            Left            =   30
            TabIndex        =   22
            Top             =   255
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   1958
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Comissão"
         Height          =   2115
         Left            =   105
         TabIndex        =   18
         Top             =   630
         Width           =   5370
         Begin MSMask.MaskEdBox PercComiss 
            Height          =   225
            Left            =   4080
            TabIndex        =   89
            Top             =   660
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
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
         Begin VB.ComboBox ComissMoeda 
            Height          =   315
            Left            =   3030
            TabIndex        =   88
            Text            =   "ComissMoeda"
            Top             =   630
            Width           =   975
         End
         Begin MSMask.MaskEdBox ComissValorAte 
            Height          =   225
            Left            =   1725
            TabIndex        =   87
            Top             =   675
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSMask.MaskEdBox ComissValorDe 
            Height          =   225
            Left            =   255
            TabIndex        =   86
            Top             =   645
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSFlexGridLib.MSFlexGrid GridComissao 
            Height          =   1110
            Left            =   30
            TabIndex        =   19
            Top             =   255
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   1958
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Redutores de comissão por investimento"
         Height          =   2040
         Left            =   105
         TabIndex        =   16
         Top             =   2805
         Width           =   5355
         Begin VB.ComboBox RedMoeda 
            Height          =   315
            Left            =   2850
            TabIndex        =   92
            Text            =   "RedMoeda"
            Top             =   705
            Width           =   975
         End
         Begin MSMask.MaskEdBox RedValorDe 
            Height          =   225
            Left            =   330
            TabIndex        =   90
            Top             =   690
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSMask.MaskEdBox RedValorAte 
            Height          =   225
            Left            =   1680
            TabIndex        =   91
            Top             =   750
            Width           =   1230
            _ExtentX        =   2170
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
         Begin MSMask.MaskEdBox PercRed 
            Height          =   225
            Left            =   3900
            TabIndex        =   93
            Top             =   735
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   397
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
         Begin MSFlexGridLib.MSFlexGrid GridRedComissao 
            Height          =   1110
            Left            =   30
            TabIndex        =   17
            Top             =   240
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   1958
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
      Begin VB.Frame Frame7 
         Height          =   525
         Left            =   225
         TabIndex        =   13
         Top             =   30
         Width           =   8790
         Begin VB.Label Label17 
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
            TabIndex        =   15
            Top             =   180
            Width           =   870
         End
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   2
            Left            =   1140
            TabIndex        =   14
            Top             =   180
            Width           =   7095
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7215
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "VendedoresTRV.ctx":011F
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "VendedoresTRV.ctx":029D
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "VendedoresTRV.ctx":07CF
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "VendedoresTRV.ctx":0959
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5430
      Left            =   120
      TabIndex        =   85
      Top             =   510
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9578
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão\Área de atuação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
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

'Alterado por Maurício Maciel em 11/04/03
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
    
    Set objCT.gobjInfoUsu = New CTVendedoresVGTRV
    Set objCT.gobjInfoUsu.gobjTelaUsu = New CTVendedoresTRV
    
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
    Call objCT.gobjInfoUsu.gobjTelaUsu.UserControl_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub VendedorLabel_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(VendedorLabel(Index), Source, X, Y)
End Sub

Private Sub VendedorLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(VendedorLabel(Index), Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
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
    Call objCT.gobjInfoUsu.gobjTelaUsu.Cargo_Change(objCT)
End Sub

Private Sub Cargo_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Cargo_Click(objCT)
End Sub

Private Sub Cargo_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Cargo_Validate(objCT, Cancel)
End Sub

Private Sub Superior_Change()
    Call objCT.gobjInfoUsu.gobjTelaUsu.Superior_Change(objCT)
End Sub

Private Sub Superior_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.Superior_Validate(objCT, Cancel)
End Sub

Private Sub LabelSuperior_Click()
    Call objCT.gobjInfoUsu.gobjTelaUsu.LabelSuperior_Click(objCT)
End Sub

Private Sub GridComissao_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Click(objCT)
End Sub

Private Sub GridComissao_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_EnterCell(objCT)
End Sub

Private Sub GridComissao_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_GotFocus(objCT)
End Sub

Private Sub GridComissao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridComissao_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridComissao_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_LeaveCell(objCT)
End Sub

Private Sub GridComissao_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Validate(objCT, Cancel)
End Sub

Private Sub GridComissao_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_RowColChange(objCT)
End Sub

Private Sub GridComissao_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridComissao_Scroll(objCT)
End Sub

Private Sub GridRedComissao_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_Click(objCT)
End Sub

Private Sub GridRedComissao_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_EnterCell(objCT)
End Sub

Private Sub GridRedComissao_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_GotFocus(objCT)
End Sub

Private Sub GridRedComissao_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridRedComissao_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridRedComissao_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_LeaveCell(objCT)
End Sub

Private Sub GridRedComissao_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_Validate(objCT, Cancel)
End Sub

Private Sub GridRedComissao_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_RowColChange(objCT)
End Sub

Private Sub GridRedComissao_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRedComissao_Scroll(objCT)
End Sub

Private Sub GridRegiaoVenda_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_Click(objCT)
End Sub

Private Sub GridRegiaoVenda_EnterCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_EnterCell(objCT)
End Sub

Private Sub GridRegiaoVenda_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_GotFocus(objCT)
End Sub

Private Sub GridRegiaoVenda_KeyDown(KeyCode As Integer, Shift As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_KeyDown(objCT, KeyCode, Shift)
End Sub

Private Sub GridRegiaoVenda_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_KeyPress(objCT, KeyAscii)
End Sub

Private Sub GridRegiaoVenda_LeaveCell()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_LeaveCell(objCT)
End Sub

Private Sub GridRegiaoVenda_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_Validate(objCT, Cancel)
End Sub

Private Sub GridRegiaoVenda_RowColChange()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_RowColChange(objCT)
End Sub

Private Sub GridRegiaoVenda_Scroll()
     Call objCT.gobjInfoUsu.gobjTelaUsu.GridRegiaoVenda_Scroll(objCT)
End Sub

Private Sub ComissValorDe_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorDe_Change(objCT)
End Sub

Private Sub ComissValorDe_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorDe_GotFocus(objCT)
End Sub

Private Sub ComissValorDe_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorDe_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ComissValorDe_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorDe_Validate(objCT, Cancel)
End Sub

Private Sub ComissValorAte_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorAte_Change(objCT)
End Sub

Private Sub ComissValorAte_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorAte_GotFocus(objCT)
End Sub

Private Sub ComissValorAte_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorAte_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ComissValorAte_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissValorAte_Validate(objCT, Cancel)
End Sub

Private Sub ComissMoeda_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissMoeda_Change(objCT)
End Sub

Private Sub ComissMoeda_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissMoeda_GotFocus(objCT)
End Sub

Private Sub ComissMoeda_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissMoeda_KeyPress(objCT, KeyAscii)
End Sub

Private Sub ComissMoeda_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.ComissMoeda_Validate(objCT, Cancel)
End Sub

Private Sub PercComiss_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_Change(objCT)
End Sub

Private Sub PercComiss_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_GotFocus(objCT)
End Sub

Private Sub PercComiss_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_KeyPress(objCT, KeyAscii)
End Sub

Private Sub PercComiss_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComiss_Validate(objCT, Cancel)
End Sub

Private Sub RedValorDe_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorDe_Change(objCT)
End Sub

Private Sub RedValorDe_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorDe_GotFocus(objCT)
End Sub

Private Sub RedValorDe_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorDe_KeyPress(objCT, KeyAscii)
End Sub

Private Sub RedValorDe_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorDe_Validate(objCT, Cancel)
End Sub

Private Sub RedValorAte_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorAte_Change(objCT)
End Sub

Private Sub RedValorAte_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorAte_GotFocus(objCT)
End Sub

Private Sub RedValorAte_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorAte_KeyPress(objCT, KeyAscii)
End Sub

Private Sub RedValorAte_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedValorAte_Validate(objCT, Cancel)
End Sub

Private Sub RedMoeda_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedMoeda_Change(objCT)
End Sub

Private Sub RedMoeda_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedMoeda_GotFocus(objCT)
End Sub

Private Sub RedMoeda_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedMoeda_KeyPress(objCT, KeyAscii)
End Sub

Private Sub RedMoeda_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.RedMoeda_Validate(objCT, Cancel)
End Sub

Private Sub PercRed_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercRed_Change(objCT)
End Sub

Private Sub PercRed_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercRed_GotFocus(objCT)
End Sub

Private Sub PercRed_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercRed_KeyPress(objCT, KeyAscii)
End Sub

Private Sub PercRed_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercRed_Validate(objCT, Cancel)
End Sub

Private Sub CodRegiao_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiao_Change(objCT)
End Sub

Private Sub CodRegiao_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiao_GotFocus(objCT)
End Sub

Private Sub CodRegiao_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiao_KeyPress(objCT, KeyAscii)
End Sub

Private Sub CodRegiao_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.CodRegiao_Validate(objCT, Cancel)
End Sub

Private Sub BotaoRegiaoVenda_Click()
     Call objCT.gobjInfoUsu.gobjTelaUsu.BotaoRegiaoVenda_Click(objCT)
End Sub

Private Sub PercComissReg_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissReg_Change(objCT)
End Sub

Private Sub PercComissReg_GotFocus()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissReg_GotFocus(objCT)
End Sub

Private Sub PercComissReg_KeyPress(KeyAscii As Integer)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissReg_KeyPress(objCT, KeyAscii)
End Sub

Private Sub PercComissReg_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercComissReg_Validate(objCT, Cancel)
End Sub

Private Sub PercCallCenter_Change()
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercCallCenter_Change(objCT)
End Sub

Private Sub PercCallCenter_Validate(Cancel As Boolean)
     Call objCT.gobjInfoUsu.gobjTelaUsu.PercCallCenter_Validate(objCT, Cancel)
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

