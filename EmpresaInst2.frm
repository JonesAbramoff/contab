VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpresaInst2 
   Caption         =   "Instalação de Empresa - 2ª Fase"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "EmpresaInst2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7650
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   106
      Top             =   45
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1110
         Picture         =   "EmpresaInst2.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "EmpresaInst2.frx":02C8
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "EmpresaInst2.frx":07FA
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Index           =   3
      Left            =   225
      TabIndex        =   72
      Top             =   765
      Width           =   9045
      Begin VB.Frame Frame6 
         Caption         =   "Contador"
         Height          =   1515
         Left            =   4395
         TabIndex        =   88
         Top             =   675
         Width           =   4530
         Begin VB.TextBox ContadorNome 
            Height          =   285
            Left            =   795
            MaxLength       =   50
            TabIndex        =   90
            Top             =   300
            Width           =   3510
         End
         Begin VB.TextBox ContadorCRC 
            Height          =   285
            Left            =   795
            MaxLength       =   20
            TabIndex        =   89
            Top             =   1095
            Width           =   2265
         End
         Begin MSMask.MaskEdBox ContadorCPF 
            Height          =   315
            Left            =   795
            TabIndex        =   91
            Top             =   690
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   11
            Mask            =   "99999999999"
            PromptChar      =   " "
         End
         Begin VB.Label Label8 
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
            Left            =   300
            TabIndex        =   94
            Top             =   750
            Width           =   420
         End
         Begin VB.Label Label5 
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
            Height          =   195
            Left            =   165
            TabIndex        =   93
            Top             =   330
            Width           =   555
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CRC:"
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
            TabIndex        =   92
            Top             =   1140
            Width           =   450
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Inscrições"
         Height          =   1515
         Left            =   150
         TabIndex        =   82
         Top             =   675
         Width           =   3990
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   1950
            TabIndex        =   83
            Top             =   675
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            Format          =   "#,#; ; ; "
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscricaoMunicipal 
            Height          =   315
            Left            =   1950
            TabIndex        =   84
            Top             =   1095
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            Format          =   "#,#; ; ; "
            PromptChar      =   " "
         End
         Begin VB.Label CGC 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   0
            Left            =   1965
            TabIndex        =   105
            Top             =   270
            Width           =   1860
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Municipal:"
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
            Left            =   165
            TabIndex        =   87
            Top             =   1155
            Width           =   1725
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual:"
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
            TabIndex        =   86
            Top             =   705
            Width           =   1650
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CGC:"
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
            Left            =   1440
            TabIndex        =   85
            Top             =   315
            Width           =   450
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Junta Comercial - Registro"
         Height          =   1110
         Left            =   195
         TabIndex        =   76
         Top             =   2310
         Width           =   3990
         Begin MSMask.MaskEdBox DataJucerja 
            Height          =   300
            Left            =   1950
            TabIndex        =   77
            Top             =   690
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownDataJucerja 
            Height          =   300
            Left            =   3075
            TabIndex        =   78
            Top             =   690
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox Jucerja 
            Height          =   315
            Left            =   1950
            TabIndex        =   79
            Top             =   270
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            Format          =   "#,#; ; ; "
            PromptChar      =   " "
         End
         Begin VB.Label Label9 
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
            Left            =   1410
            TabIndex        =   81
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label39 
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
            Left            =   1185
            TabIndex        =   80
            Top             =   315
            Width           =   720
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Ramo"
         Height          =   1110
         Left            =   4395
         TabIndex        =   73
         Top             =   2310
         Width           =   4530
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   750
            MaxLength       =   80
            TabIndex        =   74
            Top             =   480
            Width           =   3540
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Ramo:"
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
            TabIndex        =   75
            Top             =   495
            Width           =   555
         End
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   1050
         TabIndex        =   98
         Top             =   165
         Width           =   3405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         TabIndex        =   97
         Top             =   240
         Width           =   795
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   5340
         TabIndex        =   96
         Top             =   165
         Width           =   3405
      End
      Begin VB.Label Label11 
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
         Height          =   195
         Left            =   4815
         TabIndex        =   95
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Index           =   2
      Left            =   225
      TabIndex        =   100
      Top             =   765
      Width           =   9045
      Begin VB.ListBox Modulos 
         Height          =   1860
         Left            =   2220
         Style           =   1  'Checkbox
         TabIndex        =   101
         Top             =   1320
         Width           =   4245
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         Left            =   2160
         TabIndex        =   104
         Top             =   390
         Width           =   795
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   3015
         TabIndex        =   103
         Top             =   360
         Width           =   3405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Módulos Ativos"
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
         TabIndex        =   102
         Top             =   1065
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   3585
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   9045
      Begin VB.Frame SSFrame1 
         Caption         =   "Empresa"
         Height          =   1425
         Left            =   420
         TabIndex        =   115
         Top             =   300
         Width           =   6855
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1095
            TabIndex        =   116
            Top             =   330
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label Nome 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1080
            TabIndex        =   119
            Top             =   810
            Width           =   5370
         End
         Begin VB.Label Label6 
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
            Left            =   390
            TabIndex        =   118
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label4 
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
            Left            =   450
            TabIndex        =   117
            Top             =   855
            Width           =   585
         End
      End
      Begin VB.Frame SSFrame2 
         Caption         =   "Filial"
         Height          =   1455
         Left            =   420
         TabIndex        =   110
         Top             =   1920
         Width           =   6855
         Begin MSMask.MaskEdBox CodFilial 
            Height          =   345
            Left            =   1050
            TabIndex        =   111
            Top             =   360
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   609
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
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
            Height          =   225
            Left            =   360
            TabIndex        =   114
            Top             =   870
            Width           =   585
         End
         Begin VB.Label Label3 
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
            Height          =   225
            Left            =   270
            TabIndex        =   113
            Top             =   405
            Width           =   675
         End
         Begin VB.Label FilialNome 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   1035
            TabIndex        =   112
            Top             =   825
            Width           =   5370
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Index           =   5
      Left            =   240
      TabIndex        =   1
      Top             =   765
      Width           =   9045
      Begin VB.Frame Frame2 
         Caption         =   "ISS"
         Height          =   1260
         Left            =   1500
         TabIndex        =   11
         Top             =   780
         Width           =   2865
         Begin VB.CheckBox ISSIncluso 
            Caption         =   "Incluso"
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
            Left            =   180
            TabIndex        =   12
            Top             =   795
            Width           =   1020
         End
         Begin MSMask.MaskEdBox ISSPercPadrao 
            Height          =   315
            Left            =   1710
            TabIndex        =   13
            Top             =   330
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "Alíquota Padrão:"
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
            Left            =   180
            TabIndex        =   14
            Top             =   375
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "IPI"
         Height          =   1290
         Left            =   4530
         TabIndex        =   7
         Top             =   1725
         Width           =   2355
         Begin VB.OptionButton IPIContrib50 
            Caption         =   "Contribuinte 50%"
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
            Left            =   165
            TabIndex        =   10
            Top             =   945
            Width           =   1830
         End
         Begin VB.OptionButton IPIContrib 
            Caption         =   "Contribuinte Normal"
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
            Left            =   165
            TabIndex        =   9
            Top             =   630
            Width           =   2160
         End
         Begin VB.OptionButton IPINaoContrib 
            Caption         =   "Não Contribuinte"
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
            Left            =   165
            TabIndex        =   8
            Top             =   300
            Width           =   2160
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "ICMS"
         Height          =   870
         Left            =   4560
         TabIndex        =   5
         Top             =   795
         Width           =   2355
         Begin VB.CheckBox ICMSPorEstimativa 
            Caption         =   "Por Estimativa"
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
            Left            =   210
            TabIndex        =   6
            Top             =   390
            Width           =   1665
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "IR"
         Height          =   900
         Left            =   1500
         TabIndex        =   2
         Top             =   2130
         Width           =   2850
         Begin MSMask.MaskEdBox IRPercPadrao 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   315
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Caption         =   "Alíquota Padrão:"
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
            Left            =   105
            TabIndex        =   4
            Top             =   390
            Width           =   1485
         End
      End
      Begin VB.Label Label17 
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
         Height          =   195
         Left            =   4830
         TabIndex        =   18
         Top             =   210
         Width           =   465
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   3
         Left            =   5355
         TabIndex        =   17
         Top             =   142
         Width           =   3405
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         Left            =   225
         TabIndex        =   16
         Top             =   210
         Width           =   795
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   3
         Left            =   1080
         TabIndex        =   15
         Top             =   142
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Index           =   4
      Left            =   225
      TabIndex        =   19
      Top             =   765
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame FrameEndEntrega 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   2250
         Left            =   195
         TabIndex        =   45
         Top             =   1170
         Width           =   8775
         Begin VB.ComboBox Estado 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1215
            TabIndex        =   48
            Top             =   960
            Width           =   660
         End
         Begin VB.TextBox Endereco 
            Height          =   315
            Index           =   1
            Left            =   1230
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   120
            Width           =   6345
         End
         Begin VB.ComboBox Pais 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   4035
            TabIndex        =   46
            Top             =   975
            Width           =   1995
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   315
            Index           =   1
            Left            =   3975
            TabIndex        =   49
            Top             =   525
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   315
            Index           =   1
            Left            =   1215
            TabIndex        =   50
            Top             =   525
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEP 
            Height          =   315
            Index           =   1
            Left            =   6645
            TabIndex        =   51
            Top             =   555
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone1 
            Height          =   315
            Index           =   1
            Left            =   1215
            TabIndex        =   52
            Top             =   1380
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone2 
            Height          =   315
            Index           =   1
            Left            =   1215
            TabIndex        =   53
            Top             =   1800
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Email 
            Height          =   315
            Index           =   1
            Left            =   3975
            TabIndex        =   54
            Top             =   1755
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   315
            Index           =   1
            Left            =   6630
            TabIndex        =   55
            Top             =   1800
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fax 
            Height          =   315
            Index           =   1
            Left            =   3975
            TabIndex        =   56
            Top             =   1380
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label38 
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
            Left            =   240
            TabIndex        =   67
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label37 
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
            Left            =   3240
            TabIndex        =   66
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label33 
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
            Left            =   465
            TabIndex        =   65
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   555
            TabIndex        =   64
            Top             =   615
            Width           =   585
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 1:"
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
            TabIndex        =   63
            Top             =   1425
            Width           =   1005
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 2:"
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
            TabIndex        =   62
            Top             =   1845
            Width           =   1005
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Internet:"
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
            Left            =   3150
            TabIndex        =   61
            Top             =   1845
            Width           =   765
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
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
            Left            =   3510
            TabIndex        =   60
            Top             =   1425
            Width           =   405
         End
         Begin VB.Label Label27 
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
            Left            =   6105
            TabIndex        =   59
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Left            =   5820
            TabIndex        =   58
            Top             =   1845
            Width           =   750
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   3420
            TabIndex        =   57
            Top             =   1005
            Width           =   495
         End
      End
      Begin VB.Frame FrameEndPrincipal 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   2250
         Left            =   195
         TabIndex        =   22
         Top             =   1170
         Width           =   8775
         Begin VB.ComboBox Pais 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   4035
            TabIndex        =   25
            Top             =   975
            Width           =   1995
         End
         Begin VB.TextBox Endereco 
            Height          =   315
            Index           =   0
            Left            =   1215
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   120
            Width           =   6345
         End
         Begin VB.ComboBox Estado 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   1215
            TabIndex        =   23
            Top             =   960
            Width           =   660
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   315
            Index           =   0
            Left            =   3975
            TabIndex        =   26
            Top             =   525
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Bairro 
            Height          =   315
            Index           =   0
            Left            =   1215
            TabIndex        =   27
            Top             =   525
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CEP 
            Height          =   315
            Index           =   0
            Left            =   6630
            TabIndex        =   28
            Top             =   555
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   9
            Mask            =   "#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone1 
            Height          =   315
            Index           =   0
            Left            =   1215
            TabIndex        =   29
            Top             =   1380
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Telefone2 
            Height          =   315
            Index           =   0
            Left            =   1215
            TabIndex        =   30
            Top             =   1800
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Email 
            Height          =   315
            Index           =   0
            Left            =   3975
            TabIndex        =   31
            Top             =   1755
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contato 
            Height          =   315
            Index           =   0
            Left            =   6630
            TabIndex        =   32
            Top             =   1800
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Fax 
            Height          =   315
            Index           =   0
            Left            =   3975
            TabIndex        =   33
            Top             =   1380
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Format          =   "###-####"
            PromptChar      =   " "
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "País:"
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
            Left            =   3420
            TabIndex        =   44
            Top             =   1005
            Width           =   495
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "Contato:"
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
            Left            =   5820
            TabIndex        =   43
            Top             =   1845
            Width           =   750
         End
         Begin VB.Label Label60 
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
            Left            =   6105
            TabIndex        =   42
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "Fax:"
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
            Left            =   3510
            TabIndex        =   41
            Top             =   1425
            Width           =   405
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Internet:"
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
            Left            =   3150
            TabIndex        =   40
            Top             =   1845
            Width           =   765
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 2:"
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
            TabIndex        =   39
            Top             =   1845
            Width           =   1005
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Telefone 1:"
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
            TabIndex        =   38
            Top             =   1425
            Width           =   1005
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            Caption         =   "Bairro:"
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
            Left            =   555
            TabIndex        =   37
            Top             =   615
            Width           =   585
         End
         Begin VB.Label Label52 
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
            Left            =   465
            TabIndex        =   36
            Top             =   1005
            Width           =   675
         End
         Begin VB.Label Label51 
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
            Left            =   3240
            TabIndex        =   35
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label50 
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
            Left            =   240
            TabIndex        =   34
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.OptionButton EndEntrega 
         Caption         =   "Endereço de Entrega"
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
         Left            =   2670
         TabIndex        =   21
         Top             =   750
         Width           =   2220
      End
      Begin VB.OptionButton EndPrincipal 
         Caption         =   "Endereço Principal"
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
         Left            =   300
         TabIndex        =   20
         Top             =   750
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   1110
         TabIndex        =   71
         Top             =   150
         Width           =   3405
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         TabIndex        =   70
         Top             =   225
         Width           =   795
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   5385
         TabIndex        =   69
         Top             =   150
         Width           =   3405
      End
      Begin VB.Label Label21 
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
         Height          =   195
         Left            =   4860
         TabIndex        =   68
         Top             =   225
         Width           =   465
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   4065
      Left            =   150
      TabIndex        =   99
      Top             =   405
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   7170
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Módulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Complementos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereços"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tributação"
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
Attribute VB_Name = "EmpresaInst2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iFrameAtual As Integer

Private Sub CEP_GotFocus(Index As Integer)
    
    If Me.ActiveControl Is CEP(0) Then
        Call MaskEdBox_TrataGotFocus(CEP(0))
    Else
        Call MaskEdBox_TrataGotFocus(CEP(1))
    End If
    
End Sub

Private Sub CodFilial_GotFocus()

    Call MaskEdBox_TrataGotFocus(CodFilial)
    
End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo)
    
End Sub

Private Sub ContadorCPF_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(ContadorCPF)

End Sub

Private Sub DataJucerja_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataJucerja)

End Sub

Private Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 159467)

    End Select

    Exit Sub

End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

