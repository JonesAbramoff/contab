VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FilialEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filial-Empresa"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "FilialEmpresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4800
      Index           =   5
      Left            =   180
      TabIndex        =   65
      Top             =   720
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame Frame20 
         Caption         =   "Nota Fiscal de Consumidor Eletrônica NFC-e"
         Height          =   1050
         Left            =   15
         TabIndex        =   146
         Top             =   3630
         Width           =   4905
         Begin MSMask.MaskEdBox NFCECSC 
            Height          =   315
            Left            =   930
            TabIndex        =   40
            Top             =   645
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   36
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox idNFCECSC 
            Height          =   315
            Left            =   930
            TabIndex        =   39
            Top             =   270
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            PromptChar      =   "_"
         End
         Begin VB.Label Label29 
            Caption         =   "ID:"
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
            Left            =   600
            TabIndex        =   148
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label28 
            Caption         =   "CSC:"
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
            Left            =   435
            TabIndex        =   147
            Top             =   705
            Width           =   435
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Regime Tributário"
         Height          =   675
         Left            =   0
         TabIndex        =   143
         Top             =   1365
         Width           =   9030
         Begin VB.ComboBox RegimeTrib 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":014A
            Left            =   1740
            List            =   "FilialEmpresa.frx":015A
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   240
            Width           =   7215
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Regime Tributário:"
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
            TabIndex        =   144
            Top             =   300
            Width           =   1560
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Envio de Recibo Provisório de Serviço Eletrônico"
         Height          =   1050
         Left            =   4995
         TabIndex        =   141
         Top             =   3630
         Width           =   4020
         Begin VB.ComboBox RPSAmbiente 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":01F0
            Left            =   1710
            List            =   "FilialEmpresa.frx":01FA
            TabIndex        =   41
            Top             =   420
            Width           =   2205
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Ambiente:"
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
            Left            =   705
            TabIndex        =   142
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.CheckBox LucroPresumido 
         Caption         =   "Lucro Presumido"
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
         Left            =   2220
         TabIndex        =   46
         Top             =   60
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.CheckBox PISNaoCumulativo 
         Caption         =   "PIS/COFINS não cumulativos"
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
         Left            =   3795
         TabIndex        =   47
         Top             =   90
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.CheckBox SuperSimples 
         Caption         =   "Simples Nacional"
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
         TabIndex        =   42
         Top             =   90
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Frame Frame11 
         Caption         =   "Simples Federal"
         Enabled         =   0   'False
         Height          =   1005
         Left            =   5595
         TabIndex        =   101
         Top             =   -165
         Visible         =   0   'False
         Width           =   3240
         Begin VB.CheckBox SimplesFederal 
            Caption         =   "Inscrito"
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
            TabIndex        =   43
            Top             =   285
            Width           =   1035
         End
         Begin MSMask.MaskEdBox SimplesFederalAliq 
            Height          =   315
            Left            =   1635
            TabIndex        =   44
            Top             =   540
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox SimplesFederalTeto 
            Height          =   300
            Left            =   1650
            TabIndex        =   45
            Top             =   945
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Faturamento Até:"
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
            Left            =   150
            TabIndex        =   103
            Top             =   1005
            Width           =   1470
         End
         Begin VB.Label Label333 
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
            Height          =   240
            Left            =   810
            TabIndex        =   102
            Top             =   600
            Width           =   825
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Atividade"
         Height          =   600
         Left            =   0
         TabIndex        =   92
         Top             =   705
         Width           =   9030
         Begin VB.OptionButton AtivTrib 
            Caption         =   "Prestação de Serviços"
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
            Left            =   5940
            TabIndex        =   29
            Top             =   240
            Width           =   2280
         End
         Begin VB.OptionButton AtivTrib 
            Caption         =   "Industrial"
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
            Left            =   3495
            TabIndex        =   28
            Top             =   240
            Width           =   1170
         End
         Begin VB.OptionButton AtivTrib 
            Caption         =   "Comercial"
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
            Left            =   960
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "IR"
         Height          =   630
         Left            =   15
         TabIndex        =   70
         Top             =   2850
         Width           =   2730
         Begin MSMask.MaskEdBox IRPercPadrao 
            Height          =   315
            Left            =   1695
            TabIndex        =   32
            Top             =   210
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
            TabIndex        =   71
            Top             =   285
            Width           =   1485
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "ICMS"
         Height          =   600
         Left            =   0
         TabIndex        =   69
         Top             =   2175
         Width           =   2760
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
            Left            =   975
            TabIndex        =   31
            Top             =   255
            Width           =   1665
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "IPI"
         Height          =   1305
         Left            =   2805
         TabIndex        =   68
         Top             =   2175
         Width           =   2130
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
            Left            =   45
            TabIndex        =   33
            Top             =   255
            Value           =   -1  'True
            Width           =   2040
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
            Left            =   45
            TabIndex        =   34
            Top             =   570
            Width           =   2040
         End
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
            Left            =   45
            TabIndex        =   35
            Top             =   900
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ISS"
         Height          =   1305
         Left            =   4995
         TabIndex        =   66
         Top             =   2175
         Width           =   4035
         Begin VB.ComboBox RegimeEspecialTrib 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":0215
            Left            =   165
            List            =   "FilialEmpresa.frx":022B
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   885
            Width           =   3810
         End
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
            Left            =   2715
            TabIndex        =   37
            Top             =   225
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin MSMask.MaskEdBox ISSPercPadrao 
            Height          =   315
            Left            =   1695
            TabIndex        =   36
            Top             =   210
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Format          =   "#0.#0\%"
            PromptChar      =   "_"
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Regime Especial de Tributação:"
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
            TabIndex        =   145
            Top             =   570
            Width           =   2715
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
            TabIndex        =   67
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   63
         Top             =   390
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
         TabIndex        =   74
         Top             =   420
         Width           =   795
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   5355
         TabIndex        =   64
         Top             =   390
         Width           =   3405
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
         TabIndex        =   73
         Top             =   420
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4785
      Index           =   6
      Left            =   150
      TabIndex        =   104
      Top             =   750
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame Frame16 
         Caption         =   "SPED Fiscal"
         Height          =   585
         Left            =   5700
         TabIndex        =   125
         Top             =   4125
         Width           =   3225
         Begin VB.ComboBox SpedFiscalPerfil 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":02DF
            Left            =   1050
            List            =   "FilialEmpresa.frx":02EC
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   210
            Width           =   1065
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            Caption         =   "Perfil:"
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
            TabIndex        =   126
            Top             =   270
            Width           =   510
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Contribuintes substitutos"
         Height          =   1875
         Left            =   5700
         TabIndex        =   122
         Top             =   2205
         Width           =   3225
         Begin VB.ComboBox ContribUF 
            Height          =   315
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   525
            Width           =   765
         End
         Begin MSMask.MaskEdBox ContribIE 
            Height          =   285
            Left            =   1410
            TabIndex        =   123
            Top             =   540
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridContrib 
            Height          =   600
            Left            =   30
            TabIndex        =   56
            Top             =   270
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   1058
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Instituições Responsáveis pela Administração do Cadastro"
         Height          =   2505
         Left            =   75
         TabIndex        =   114
         Top             =   2205
         Width           =   5550
         Begin MSMask.MaskEdBox AdmCadInscricao 
            Height          =   285
            Left            =   3480
            TabIndex        =   120
            Top             =   540
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "0000-0/00"
            PromptChar      =   " "
         End
         Begin VB.ComboBox AdmCadCodigo 
            Height          =   315
            Left            =   555
            Style           =   2  'Dropdown List
            TabIndex        =   119
            Top             =   525
            Width           =   2910
         End
         Begin MSFlexGridLib.MSFlexGrid GridAdmCad 
            Height          =   1020
            Left            =   45
            TabIndex        =   55
            Top             =   270
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   1799
            _Version        =   393216
            Rows            =   7
            Cols            =   4
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Signatário da Escrituração"
         Height          =   975
         Left            =   75
         TabIndex        =   110
         Top             =   1200
         Width           =   8850
         Begin VB.ComboBox SignatQualif 
            Height          =   315
            Left            =   4215
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   585
            Width           =   4545
         End
         Begin VB.TextBox SignatNome 
            Height          =   285
            Left            =   900
            MaxLength       =   50
            TabIndex        =   52
            Top             =   255
            Width           =   7830
         End
         Begin MSMask.MaskEdBox SignatCPF 
            Height          =   315
            Left            =   900
            TabIndex        =   53
            Top             =   585
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   11
            Format          =   "###\.###\.###-##; ; ; "
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin VB.Label Label41 
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
            Left            =   405
            TabIndex        =   113
            Top             =   645
            Width           =   420
         End
         Begin VB.Label Label40 
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
            Left            =   270
            TabIndex        =   112
            Top             =   285
            Width           =   555
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Qualificação:"
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
            Left            =   3030
            TabIndex        =   111
            Top             =   630
            Width           =   1125
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "NFe"
         Height          =   900
         Left            =   75
         TabIndex        =   105
         Top             =   300
         Width           =   8850
         Begin VB.ComboBox indSincPadrao 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":02F9
            Left            =   6945
            List            =   "FilialEmpresa.frx":0307
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   165
            Width           =   1800
         End
         Begin VB.TextBox CertificadoA1A3 
            Height          =   315
            Left            =   3750
            MaxLength       =   80
            TabIndex        =   49
            Top             =   165
            Width           =   1650
         End
         Begin VB.ComboBox NFeAmbiente 
            Height          =   315
            ItemData        =   "FilialEmpresa.frx":0326
            Left            =   885
            List            =   "FilialEmpresa.frx":0330
            TabIndex        =   48
            Top             =   165
            Width           =   1800
         End
         Begin MSMask.MaskEdBox CNAE 
            Height          =   315
            Left            =   885
            TabIndex        =   51
            Top             =   510
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   7
            Format          =   "0000-0/00"
            Mask            =   "#######"
            PromptChar      =   " "
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Ind.Sincronismo:"
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
            Left            =   5490
            TabIndex        =   149
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Certificado:"
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
            Left            =   2700
            TabIndex        =   109
            Top             =   225
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "NFe:"
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
            TabIndex        =   108
            Top             =   225
            Width           =   420
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CNAE:"
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
            TabIndex        =   107
            Top             =   585
            Width           =   570
         End
         Begin VB.Label DescCNAE 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   2205
            TabIndex        =   106
            Top             =   525
            Width           =   6525
         End
      End
      Begin VB.Label Label45 
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
         Left            =   4710
         TabIndex        =   118
         Top             =   45
         Width           =   465
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   4
         Left            =   5235
         TabIndex        =   117
         Top             =   0
         Width           =   3405
      End
      Begin VB.Label Label42 
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
         Left            =   105
         TabIndex        =   116
         Top             =   45
         Width           =   795
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   5
         Left            =   945
         TabIndex        =   115
         Top             =   0
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Index           =   4
      Left            =   120
      TabIndex        =   82
      Top             =   720
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame FrameEndEntrega 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3420
         Left            =   195
         TabIndex        =   84
         Top             =   1110
         Visible         =   0   'False
         Width           =   8775
         Begin DicPrincipal.TabEndereco TabEnd 
            Height          =   3375
            Index           =   1
            Left            =   120
            TabIndex        =   131
            Top             =   105
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   5953
         End
      End
      Begin VB.Frame FrameEndPrincipal 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   3420
         Left            =   195
         TabIndex        =   83
         Top             =   1110
         Width           =   8775
         Begin DicPrincipal.TabEndereco TabEnd 
            Height          =   3375
            Index           =   0
            Left            =   120
            TabIndex        =   130
            Top             =   105
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   5953
         End
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
         Left            =   2400
         TabIndex        =   25
         Top             =   840
         Value           =   -1  'True
         Width           =   1965
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
         Left            =   4800
         TabIndex        =   26
         Top             =   840
         Width           =   2220
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
         TabIndex        =   86
         Top             =   405
         Width           =   465
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   5385
         TabIndex        =   62
         Top             =   360
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
         TabIndex        =   85
         Top             =   405
         Width           =   795
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   3
         Left            =   1110
         TabIndex        =   61
         Top             =   360
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4725
      Index           =   3
      Left            =   195
      TabIndex        =   75
      Top             =   750
      Visible         =   0   'False
      Width           =   9045
      Begin VB.Frame Frame6 
         Caption         =   "Contador"
         Height          =   4410
         Left            =   15
         TabIndex        =   76
         Top             =   330
         Width           =   8910
         Begin VB.Frame Frame17 
            Caption         =   "Endereço do Escritório"
            Height          =   3480
            Left            =   45
            TabIndex        =   128
            Top             =   885
            Width           =   8775
            Begin VB.Frame Frame18 
               BorderStyle     =   0  'None
               Caption         =   "Frame7"
               Height          =   3225
               Left            =   60
               TabIndex        =   129
               Top             =   210
               Width           =   8580
               Begin DicPrincipal.TabEndereco TabEnd 
                  Height          =   3225
                  Index           =   2
                  Left            =   0
                  TabIndex        =   140
                  Top             =   0
                  Width           =   8505
                  _ExtentX        =   15002
                  _ExtentY        =   5689
               End
            End
         End
         Begin VB.TextBox ContadorCRC 
            Height          =   315
            Left            =   1575
            MaxLength       =   20
            TabIndex        =   23
            Top             =   540
            Width           =   2355
         End
         Begin VB.TextBox ContadorNome 
            Height          =   315
            Left            =   1575
            MaxLength       =   50
            TabIndex        =   21
            Top             =   180
            Width           =   3960
         End
         Begin MSMask.MaskEdBox ContadorCPF 
            Height          =   315
            Left            =   6510
            TabIndex        =   22
            Top             =   180
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   11
            Format          =   "###\.###\.###-##; ; ; "
            Mask            =   "###########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CNPJCTB 
            Height          =   315
            Left            =   6510
            TabIndex        =   24
            Top             =   585
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Format          =   "00\.000\.000\/0000-00; ; ; "
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ do Escritório:"
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
            Left            =   4770
            TabIndex        =   127
            Top             =   630
            Width           =   1650
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
            Left            =   1080
            TabIndex        =   79
            Top             =   585
            Width           =   450
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
            Left            =   975
            TabIndex        =   78
            Top             =   210
            Width           =   555
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
            Left            =   6015
            TabIndex        =   77
            Top             =   225
            Width           =   420
         End
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
         TabIndex        =   81
         Top             =   75
         Width           =   465
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   5340
         TabIndex        =   60
         Top             =   30
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
         Left            =   180
         TabIndex        =   80
         Top             =   75
         Width           =   795
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   59
         Top             =   30
         Width           =   3405
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4725
      Index           =   2
      Left            =   195
      TabIndex        =   88
      Top             =   735
      Visible         =   0   'False
      Width           =   9045
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   7335
         Picture         =   "FilialEmpresa.frx":034B
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3210
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   7335
         Picture         =   "FilialEmpresa.frx":1365
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4035
         Width           =   1425
      End
      Begin VB.ListBox Modulos 
         Height          =   3660
         Left            =   1065
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   975
         Width           =   6195
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
         Left            =   1080
         TabIndex        =   91
         Top             =   735
         Width           =   1305
      End
      Begin VB.Label EmpresaLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   20
         Top             =   180
         Width           =   3405
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
         Left            =   210
         TabIndex        =   90
         Top             =   225
         Width           =   795
      End
      Begin VB.Label FilialLabel 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   5340
         TabIndex        =   58
         Top             =   180
         Width           =   3405
      End
      Begin VB.Label Label43 
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
         TabIndex        =   89
         Top             =   225
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   4725
      Index           =   1
      Left            =   105
      TabIndex        =   72
      Top             =   720
      Width           =   9045
      Begin VB.Frame SSFrame2 
         Caption         =   "Filial"
         Height          =   3615
         Left            =   210
         TabIndex        =   98
         Top             =   1125
         Width           =   5220
         Begin VB.TextBox Ramo 
            Height          =   315
            Left            =   1455
            MaxLength       =   80
            TabIndex        =   5
            Top             =   1260
            Width           =   3705
         End
         Begin VB.Frame Frame8 
            Caption         =   "Junta Comercial - Registro"
            Height          =   660
            Left            =   210
            TabIndex        =   136
            Top             =   2865
            Width           =   4785
            Begin MSMask.MaskEdBox DataJucerja 
               Height          =   315
               Left            =   3345
               TabIndex        =   10
               Top             =   255
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownDataJucerja 
               Height          =   300
               Left            =   4470
               TabIndex        =   11
               Top             =   240
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox Jucerja 
               Height          =   315
               Left            =   1230
               TabIndex        =   9
               Top             =   255
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
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
               Left            =   2805
               TabIndex        =   138
               Top             =   315
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
               Left            =   465
               TabIndex        =   137
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Inscrições"
            Height          =   1320
            Left            =   210
            TabIndex        =   132
            Top             =   1545
            Width           =   4785
            Begin MSMask.MaskEdBox CGC 
               Height          =   315
               Left            =   1230
               TabIndex        =   6
               Top             =   195
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   14
               Format          =   "00\.000\.000\/0000-00; ; ; "
               Mask            =   "##############"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox InscricaoEstadual 
               Height          =   315
               Left            =   1230
               TabIndex        =   7
               Top             =   555
               Width           =   1860
               _ExtentX        =   3281
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox InscricaoMunicipal 
               Height          =   315
               Left            =   1230
               TabIndex        =   8
               Top             =   915
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   18
               PromptChar      =   " "
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Municipal:"
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
               Left            =   315
               TabIndex        =   135
               Top             =   975
               Width           =   885
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Estadual:"
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
               TabIndex        =   134
               Top             =   600
               Width           =   810
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "CNPJ:"
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
               Left            =   660
               TabIndex        =   133
               Top             =   240
               Width           =   540
            End
         End
         Begin MSMask.MaskEdBox Nome 
            Height          =   315
            Left            =   1455
            TabIndex        =   3
            Top             =   540
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1455
            TabIndex        =   2
            Top             =   180
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1455
            TabIndex        =   4
            Top             =   900
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
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
            Left            =   870
            TabIndex        =   139
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Nome Fantasia:"
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
            Left            =   75
            TabIndex        =   121
            Top             =   945
            Width           =   1335
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
            Height          =   195
            Left            =   750
            TabIndex        =   100
            Top             =   225
            Width           =   675
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
            Left            =   855
            TabIndex        =   99
            Top             =   585
            Width           =   585
         End
      End
      Begin VB.Frame SSFrame1 
         Caption         =   "Empresa"
         Height          =   930
         Left            =   210
         TabIndex        =   94
         Top             =   195
         Width           =   5220
         Begin MSMask.MaskEdBox CodigoEmpresa 
            Height          =   315
            Left            =   1455
            TabIndex        =   1
            Top             =   180
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   2
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label EmpresaLabel 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Index           =   0
            Left            =   1440
            TabIndex        =   97
            Top             =   540
            Width           =   3705
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
            Left            =   705
            TabIndex        =   96
            Top             =   225
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
            Left            =   780
            TabIndex        =   95
            Top             =   555
            Width           =   585
         End
      End
      Begin MSComctlLib.TreeView Filiais 
         Height          =   4440
         Left            =   5640
         TabIndex        =   0
         Top             =   285
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   7832
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresas-Filiais"
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
         Left            =   5670
         TabIndex        =   87
         Top             =   15
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7230
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   93
      Top             =   15
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "FilialEmpresa.frx":2547
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "FilialEmpresa.frx":26A1
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "FilialEmpresa.frx":282B
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "FilialEmpresa.frx":2D5D
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5205
      Left            =   90
      TabIndex        =   12
      Top             =   360
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Módulos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contador"
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
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "SPED"
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
Attribute VB_Name = "FilialEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents objCT As CTFilialEmpresa
Attribute objCT.VB_VarHelpID = -1

Private Sub BotaoDesmarcarTodos_Click()
     Call objCT.BotaoDesmarcarTodos_Click
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

Private Sub BotaoMarcarTodos_Click()
     Call objCT.BotaoMarcarTodos_Click
End Sub

Private Sub CGC_GotFocus()
     Call objCT.CGC_GotFocus
End Sub

Private Sub CGC_Validate(Cancel As Boolean)
     Call objCT.CGC_Validate(Cancel)
End Sub

Private Sub Codigo_GotFocus()
     Call objCT.Codigo_GotFocus
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)
     Call objCT.Codigo_Validate(Cancel)
End Sub

Private Sub CodigoEmpresa_GotFocus()
     Call objCT.CodigoEmpresa_GotFocus
End Sub

Private Sub CodigoEmpresa_Validate(Cancel As Boolean)
     Call objCT.CodigoEmpresa_Validate(Cancel)
End Sub

Private Sub ContadorCPF_GotFocus()
     Call objCT.ContadorCPF_GotFocus
End Sub

Private Sub ContadorCPF_Validate(Cancel As Boolean)
     Call objCT.ContadorCPF_Validate(Cancel)
End Sub

Private Sub ContribIE_Change()
    Call objCT.ContribIE_Change
End Sub

Private Sub ContribIE_Validate(Cancel As Boolean)
    Call objCT.ContribIE_Validate(Cancel)
End Sub

Private Sub ContribUF_Change()
    Call objCT.ContribUF_Change
End Sub

Private Sub ContribUF_Click()
    Call objCT.ContribUF_Click
End Sub

Private Sub ContribUF_Validate(Cancel As Boolean)
    Call objCT.ContribUF_Validate(Cancel)
End Sub

Private Sub DataJucerja_GotFocus()
     Call objCT.DataJucerja_GotFocus
End Sub

Private Sub DataJucerja_Validate(Cancel As Boolean)
     Call objCT.DataJucerja_Validate(Cancel)
End Sub

Private Sub EndEntrega_Click()
     Call objCT.EndEntrega_Click
End Sub

Private Sub EndPrincipal_Click()
     Call objCT.EndPrincipal_Click
End Sub

Private Sub Filiais_NodeClick(ByVal Node As MSComctlLib.Node)
     Call objCT.Filiais_NodeClick(Node)
End Sub

Private Sub IRPercPadrao_Validate(Cancel As Boolean)
     Call objCT.IRPercPadrao_Validate(Cancel)
End Sub

Private Sub ISSPercPadrao_Validate(Cancel As Boolean)
     Call objCT.ISSPercPadrao_Validate(Cancel)
End Sub

Private Sub Nome_Validate(Cancel As Boolean)
     Call objCT.Nome_Validate(Cancel)
End Sub

Private Sub Opcao_Click()
     Call objCT.Opcao_Click
End Sub

Private Sub Form_Load()

    Set objCT = New CTFilialEmpresa
    Set objCT.objTela = Me

    Call objCT.Form_Load

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (objCT Is Nothing) Then
        Call objCT.Form_Unload(Cancel)
        If Cancel = False Then
             Set objCT.objTela = Nothing
             Set objCT = Nothing
        End If
    End If
End Sub

Private Sub objCT_Unload()
   Unload Me
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub SimplesFederal_Click()
    objCT.SimplesFederal_Click
End Sub

Private Sub SimplesFederalAliq_Validate(Cancel As Boolean)
     Call objCT.SimplesFederalAliq_Validate(Cancel)
End Sub

Private Sub SimplesFederalTeto_Validate(Cancel As Boolean)
     Call objCT.SimplesFederalTeto_Validate(Cancel)
End Sub

Private Sub SuperSimples_Click()
    objCT.SuperSimples_Click
End Sub

Private Sub PISNaoCumulativo_Click()
    objCT.PISNaoCumulativo_Click
End Sub

Private Sub LucroPresumido_Click()
    objCT.LucroPresumido_Click
End Sub

Private Sub ICMSPorEstimativa_Click()
    objCT.ICMSPorEstimativa_Click
End Sub

Private Sub CNAE_GotFocus()
    objCT.CNAE_GotFocus
End Sub

Private Sub CNAE_Validate(Cancel As Boolean)
     Call objCT.CNAE_Validate(Cancel)
End Sub

Private Sub GridAdmCad_Click()
    Call objCT.GridAdmCad_Click
End Sub

Private Sub GridAdmCad_GotFocus()
    Call objCT.GridAdmCad_GotFocus
End Sub

Private Sub GridAdmCad_EnterCell()
    Call objCT.GridAdmCad_EnterCell
End Sub

Private Sub GridAdmCad_LeaveCell()
    Call objCT.GridAdmCad_LeaveCell
End Sub

Private Sub GridAdmCad_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.GridAdmCad_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridAdmCad_KeyPress(KeyAscii As Integer)
    Call objCT.GridAdmCad_KeyPress(KeyAscii)
End Sub

Private Sub GridAdmCad_LostFocus()
    Call objCT.GridAdmCad_LostFocus
End Sub

Private Sub GridAdmCad_RowColChange()
    Call objCT.GridAdmCad_RowColChange
End Sub

Private Sub GridAdmCad_Scroll()
    Call objCT.GridAdmCad_Scroll
End Sub

Private Sub AdmCadCodigo_GotFocus()
    Call objCT.AdmCadCodigo_GotFocus
End Sub

Private Sub AdmCadCodigo_KeyPress(KeyAscii As Integer)
    Call objCT.AdmCadCodigo_KeyPress(KeyAscii)
End Sub

Private Sub AdmCadCodigo_LostFocus()
    Call objCT.AdmCadCodigo_LostFocus
End Sub

Private Sub AdmCadInscricao_GotFocus()
    Call objCT.AdmCadInscricao_GotFocus
End Sub

Private Sub AdmCadInscricao_KeyPress(KeyAscii As Integer)
    Call objCT.AdmCadInscricao_KeyPress(KeyAscii)
End Sub

Private Sub AdmCadInscricao_LostFocus()
    Call objCT.AdmCadInscricao_LostFocus
End Sub

Private Sub SignatCPF_GotFocus()
     Call objCT.SignatCPF_GotFocus
End Sub

Private Sub SignatCPF_Validate(Cancel As Boolean)
     Call objCT.SignatCPF_Validate(Cancel)
End Sub

Private Sub GridContrib_Click()
    Call objCT.GridContrib_Click
End Sub

Private Sub GridContrib_GotFocus()
    Call objCT.GridContrib_GotFocus
End Sub

Private Sub GridContrib_EnterCell()
    Call objCT.GridContrib_EnterCell
End Sub

Private Sub GridContrib_LeaveCell()
    Call objCT.GridContrib_LeaveCell
End Sub

Private Sub GridContrib_KeyDown(KeyCode As Integer, Shift As Integer)
    Call objCT.GridContrib_KeyDown(KeyCode, Shift)
End Sub

Private Sub GridContrib_KeyPress(KeyAscii As Integer)
    Call objCT.GridContrib_KeyPress(KeyAscii)
End Sub

Private Sub GridContrib_Validate(Cancel As Boolean)
    Call objCT.GridContrib_Validate(Cancel)
End Sub

Private Sub GridContrib_RowColChange()
    Call objCT.GridContrib_RowColChange
End Sub

Private Sub GridContrib_Scroll()
    Call objCT.GridContrib_Scroll
End Sub

Private Sub ContribIE_GotFocus()
    Call objCT.ContribIE_GotFocus
End Sub

Private Sub ContribIE_KeyPress(KeyAscii As Integer)
    Call objCT.ContribIE_KeyPress(KeyAscii)
End Sub

Private Sub ContribUF_GotFocus()
    Call objCT.ContribUF_GotFocus
End Sub

Private Sub ContribUF_KeyPress(KeyAscii As Integer)
    Call objCT.ContribUF_KeyPress(KeyAscii)
End Sub

Private Sub CNPJCTB_GotFocus()
     Call objCT.CNPJCTB_GotFocus
End Sub

Private Sub CNPJCTB_Validate(Cancel As Boolean)
     Call objCT.CNPJCTB_Validate(Cancel)
End Sub

Private Sub RegimeEspecialTrib_Click()
     Call objCT.RegimeEspecialTrib_Click
End Sub

Private Sub RegimeTrib_Click()
     Call objCT.RegimeTrib_Click
End Sub

Private Sub idNFCECSC_Change()
     Call objCT.idNFCECSC_Change
End Sub

Private Sub idNFCECSC_Validate(Cancel As Boolean)
     Call objCT.idNFCECSC_Validate(Cancel)
End Sub

Private Sub indSincPadrao_Change()
     Call objCT.indSincPadrao_Change
End Sub

Private Sub indSincPadrao_Click()
     Call objCT.indSincPadrao_Click
End Sub

Private Sub NFCECSC_Change()
     Call objCT.NFCECSC_Change
End Sub

Private Sub NFCECSC_Validate(Cancel As Boolean)
     Call objCT.NFCECSC_Validate(Cancel)
End Sub

