VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl VendedoresOcx 
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
      Height          =   4410
      Index           =   1
      Left            =   130
      TabIndex        =   0
      Top             =   1020
      Width           =   8610
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Autônomo"
         Height          =   1650
         Index           =   0
         Left            =   75
         TabIndex        =   66
         Top             =   2670
         Visible         =   0   'False
         Width           =   5685
         Begin MSMask.MaskEdBox CGC0 
            Height          =   315
            Left            =   1635
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   750
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
            Left            =   1200
            TabIndex        =   78
            Top             =   810
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
            Left            =   1140
            TabIndex        =   77
            Top             =   375
            Width           =   420
         End
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados do Empregado"
         Height          =   1650
         Index           =   1
         Left            =   90
         TabIndex        =   67
         Top             =   2655
         Visible         =   0   'False
         Width           =   5685
         Begin MSMask.MaskEdBox Matricula 
            Height          =   315
            Left            =   1230
            TabIndex        =   68
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
            Left            =   240
            TabIndex        =   69
            Top             =   420
            Width           =   885
         End
      End
      Begin VB.Frame FrameInfoVinculo 
         Caption         =   "Dados da Empresa"
         Height          =   1710
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   2595
         Visible         =   0   'False
         Width           =   5625
         Begin VB.TextBox RazaoSocial 
            Height          =   315
            Left            =   1575
            MaxLength       =   40
            TabIndex        =   65
            Top             =   240
            Width           =   3630
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   1560
            TabIndex        =   60
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
            TabIndex        =   61
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
            TabIndex        =   64
            Top             =   300
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
            Left            =   525
            TabIndex        =   63
            Top             =   795
            Width           =   975
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
            TabIndex        =   62
            Top             =   1305
            Width           =   1290
         End
      End
      Begin VB.ComboBox CodUsuario 
         Height          =   315
         Left            =   3975
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   45
         Width           =   1725
      End
      Begin VB.ComboBox Cargo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         TabIndex        =   71
         Top             =   1725
         Width           =   1935
      End
      Begin VB.ComboBox Vinculo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "VendedoresOcx.ctx":0000
         Left            =   1710
         List            =   "VendedoresOcx.ctx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2145
         Width           =   2790
      End
      Begin VB.CheckBox Ativo 
         Caption         =   "Ativo"
         Height          =   195
         Left            =   2085
         TabIndex        =   57
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton BotaoProxNum 
         Height          =   285
         Left            =   1560
         Picture         =   "VendedoresOcx.ctx":0035
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Numeração Automática"
         Top             =   60
         Width           =   300
      End
      Begin VB.ComboBox Tipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         TabIndex        =   5
         Top             =   1305
         Width           =   1935
      End
      Begin VB.ComboBox Regiao 
         Height          =   315
         Left            =   3690
         TabIndex        =   6
         Top             =   1275
         Width           =   1965
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   315
         Left            =   1020
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
         Left            =   1005
         TabIndex        =   3
         Top             =   465
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NomeReduzido 
         Height          =   315
         Left            =   1665
         TabIndex        =   4
         Top             =   855
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.ListBox VendedoresList 
         Height          =   3960
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   330
         Width           =   2685
      End
      Begin MSMask.MaskEdBox Superior 
         Height          =   315
         Left            =   3690
         TabIndex        =   72
         Top             =   1725
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label11 
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
         Left            =   3165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   76
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
         Left            =   2910
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   74
         Top             =   1770
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
         Left            =   120
         TabIndex        =   73
         Top             =   1800
         Width           =   585
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
         TabIndex        =   58
         Top             =   2220
         Width           =   735
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
         Left            =   285
         TabIndex        =   29
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
         Left            =   375
         TabIndex        =   30
         Top             =   525
         Width           =   555
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
         Left            =   210
         TabIndex        =   31
         Top             =   945
         Width           =   1410
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
         Left            =   195
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Top             =   1335
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
         Left            =   5880
         TabIndex        =   34
         Top             =   105
         Width           =   1440
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
         Left            =   2985
         TabIndex        =   33
         Top             =   1320
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4485
      Index           =   3
      Left            =   165
      TabIndex        =   23
      Top             =   930
      Visible         =   0   'False
      Width           =   8640
      Begin TelasCpr.TabEndereco TabEnd 
         Height          =   3540
         Index           =   0
         Left            =   30
         TabIndex        =   70
         Top             =   840
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   6244
      End
      Begin VB.Frame SSFrame2 
         Height          =   525
         Left            =   60
         TabIndex        =   53
         Top             =   75
         Width           =   8460
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
            TabIndex        =   54
            Top             =   180
            Width           =   870
         End
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   1
            Left            =   1140
            TabIndex        =   55
            Top             =   180
            Width           =   7095
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4050
      Index           =   2
      Left            =   195
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   8625
      Begin VB.Frame Frame1 
         Caption         =   "Incide sobre"
         Height          =   1350
         Index           =   0
         Left            =   3975
         TabIndex        =   48
         Top             =   1470
         Width           =   4455
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
            TabIndex        =   14
            Top             =   375
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
            TabIndex        =   15
            Top             =   390
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
            TabIndex        =   16
            Top             =   780
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
            Left            =   1110
            TabIndex        =   17
            Top             =   780
            Width           =   990
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
            TabIndex        =   18
            Top             =   780
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
            Left            =   3765
            TabIndex        =   19
            Top             =   780
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Estatísticas"
         Height          =   705
         Left            =   210
         TabIndex        =   38
         Top             =   645
         Width           =   8220
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
            TabIndex        =   39
            Top             =   315
            Width           =   1500
         End
         Begin VB.Label SaldoComissao 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1965
            TabIndex        =   40
            Top             =   270
            Width           =   1575
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
            TabIndex        =   41
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label DataUltVenda 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6405
            TabIndex        =   42
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.Frame SSFrame6 
         Caption         =   "Conta Corrente"
         Height          =   915
         Left            =   195
         TabIndex        =   49
         Top             =   2940
         Width           =   8205
         Begin MSMask.MaskEdBox ContaCorrente 
            Height          =   315
            Left            =   1830
            TabIndex        =   20
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
            TabIndex        =   21
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
            TabIndex        =   22
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
            TabIndex        =   52
            Top             =   480
            Width           =   615
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
            TabIndex        =   51
            Top             =   480
            Width           =   765
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
            TabIndex        =   50
            Top             =   480
            Width           =   720
         End
      End
      Begin VB.Frame SSFrame3 
         Height          =   510
         Left            =   225
         TabIndex        =   35
         Top             =   30
         Width           =   8220
         Begin VB.Label VendedorLabel 
            Height          =   210
            Index           =   0
            Left            =   1140
            TabIndex        =   37
            Top             =   165
            Width           =   6750
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
            TabIndex        =   36
            Top             =   165
            Width           =   870
         End
      End
      Begin VB.Frame SSFrame4 
         Caption         =   "Porcentagens"
         Height          =   1350
         Left            =   210
         TabIndex        =   43
         Top             =   1470
         Width           =   3750
         Begin MSMask.MaskEdBox PercComissao 
            Height          =   315
            Left            =   1215
            TabIndex        =   12
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
            TabIndex        =   13
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
         Begin VB.Label PercComissaoBaixa 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2895
            TabIndex        =   47
            Top             =   788
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
            Left            =   2010
            TabIndex        =   46
            Top             =   840
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
            Left            =   675
            TabIndex        =   44
            Top             =   360
            Width           =   495
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
            TabIndex        =   45
            Top             =   840
            Width           =   1065
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6735
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "VendedoresOcx.ctx":011F
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "VendedoresOcx.ctx":0279
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "VendedoresOcx.ctx":0403
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "VendedoresOcx.ctx":0935
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4980
      Left            =   120
      TabIndex        =   56
      Top             =   555
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   8784
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "VendedoresOcx"
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
